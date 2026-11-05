"""Deterministic offline discovery, reviewable plans and reusable reconciliation profiles."""
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal
import hashlib
import json
import re
from itertools import combinations
from uuid import uuid4

from openpyxl.utils import get_column_letter
import spreadsheet_compare as compare

VERSION = 3
RAW_FIELDS = {'date': 'audit date', 'form': 'audit form', 'factory': 'factory',
              'line': 'line/machine', 'rate': 'mqaa%', 'group': 'internal factory'}


def normalized(value):
    return re.sub(r'\s+', ' ', str(value if value is not None else '').strip()).casefold()


def factory(value):
    value = str(value or '').strip().upper()
    return 'JVB' if value == 'JV3' else value


def headers(sheet, row):
    return {normalized(compare.cell(sheet, row, col)['value']): col for col in range(1, sheet['columns']+1)
            if normalized(compare.cell(sheet, row, col)['value'])}


def raw_tables(book):
    tables = []
    for name, sheet in book['sheets'].items():
        for row in range(1, min(sheet['rows'], 12)+1):
            mapping = headers(sheet, row)
            if set(RAW_FIELDS.values()) <= set(mapping):
                tables.append({'sheet': name, 'header': row, 'first': row+1, 'last': sheet['rows'],
                               'fields': {key: mapping[label] for key, label in RAW_FIELDS.items()}})
                break
    return tables


def matrices(book):
    result = []
    for name, sheet in book['sheets'].items():
        for row in range(1, min(sheet['rows'], 12)+1):
            mapping = headers(sheet, row)
            if 'line' not in mapping or 'average' not in mapping:
                continue
            line, average = mapping['line'], mapping['average']
            total = [r for r in range(row+1, sheet['rows']+1) if any(
                normalized(compare.cell(sheet, r, col)['value']) == 'grand total' for col in range(1, line+1))]
            if len(total) != 1 or average <= line+1:
                continue
            last = total[0]-1
            value_columns = list(range(line+1, average))
            if any(not normalized(compare.cell(sheet, row, col)['value']) for col in value_columns):
                continue
            result.append(dict(sheet=name, header=row, first=row+1, last=last,
                               total=total[0], line=line, average=average, value_columns=value_columns))
            break
    return result


def day(value):
    if isinstance(value, (date, datetime)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    match = re.match(r'(\d{4}-\d{2}-\d{2})', str(value or '').strip())
    if not match:
        return None
    try:
        return date.fromisoformat(match[1]).isoformat()
    except ValueError:
        return None


def periods(book):
    found = set()
    for sheet in book['sheets'].values():
        for row in range(1, min(sheet['rows'], 4)+1):
            for col in range(1, min(sheet['columns'], 4)+1):
                match = re.search(r'From:\s*(\d{4}-\d{2}-\d{2})\s+To:\s*(\d{4}-\d{2}-\d{2})', str(compare.cell(sheet, row, col)['value'] or ''), re.I)
                if match:
                    found.add(match.groups())
    return sorted(found)


def table_rows(book, table):
    sheet = book['sheets'][table['sheet']]
    for row in range(table['first'], table['last']+1):
        values = {key: compare.cell(sheet, row, col)['value'] for key, col in table['fields'].items()}
        if not any(value not in (None, '') for value in values.values()):
            continue
        yield row, values


def range_text(left, top, right, bottom):
    return f'{get_column_letter(left)}{top}:{get_column_letter(right)}{bottom}'


def mqaa_plan(source, report, scope=None):
    source_tables, target_matrices = raw_tables(source), matrices(report)
    if not source_tables or not target_matrices:
        raise ValueError('Chưa nhận diện được bảng dữ liệu MQAA và ma trận Line / Average / GRAND TOTAL. Hãy dùng quy tắc thủ công cho mẫu này.')
    combined = [table for table in source_tables if normalized(table['sheet']) == 'combined']
    if len(source_tables) > 1 and len(combined) != 1:
        raise ValueError('Nguồn có nhiều bảng MQAA nhưng chưa có một sheet Combined duy nhất. Hãy gộp nguồn trước để tránh đếm lặp.')
    table = combined[0] if combined else source_tables[0]
    records = list(table_rows(source, table))
    source_factories = sorted({factory(values['factory']) for _, values in records if factory(values['factory'])})
    qc_factories = sorted({factory(values['factory']) for raw in raw_tables(report)
                           for _, values in table_rows(report, raw) if factory(values['factory'])})
    selected = [factory(scope)] if scope else qc_factories
    if len(selected) != 1 or selected[0] not in source_factories:
        raise ValueError('Ma trận MQAA không ghi nhà máy trên từng ô. Hãy chỉ định đúng một nhà máy thuộc nguồn để đối chiếu; không tự trộn các nhà máy.')
    if qc_factories and selected != qc_factories:
        raise ValueError('Nhà máy đã chọn không khớp dữ liệu chi tiết trong file báo cáo QC.')
    source_periods, report_periods = periods(source), periods(report)
    if len(source_periods) > 1 or len(report_periods) != 1:
        raise ValueError('Chưa xác định được một kỳ báo cáo MQAA duy nhất từ nội dung file.')
    period = report_periods[0]
    if source_periods and source_periods[0] != period:
        raise ValueError('Khoảng ngày trong file nguồn và báo cáo QC khác nhau. Chọn hai file cùng kỳ trước khi chạy.')
    scoped = [(row, values) for row, values in records if factory(values['factory']) in selected]
    if any(not day(values['date']) or not period[0] <= day(values['date']) <= period[1] for _, values in scoped):
        raise ValueError('Nguồn chứa ngày không đọc được hoặc ngoài kỳ của báo cáo QC. Chưa tự cắt bớt dữ liệu để chạy.')
    rules, descriptions = [], []
    for matrix in target_matrices:
        sheet = report['sheets'][matrix['sheet']]
        rows_range = range_text(matrix['line']+1, matrix['first'], matrix['average']-1, matrix['last'])
        common = dict(kind='calculation', left={'workbook': report['id'], 'sheet': matrix['sheet']},
                      operation='average', source_range=rows_range, number_mode='percent', tolerance='0',
                      rounding_digits='', missing_policy='skip', ignore_zero=False)
        rules.append(dict(common, name=f'QC nội bộ · Average từng dòng · {matrix["sheet"]}', direction='rows',
                          target_range=range_text(matrix['average'], matrix['first'], matrix['average'], matrix['last']),
                          label_range=range_text(matrix['line'], matrix['first'], matrix['line'], matrix['last'])))
        rules.append(dict(common, name=f'QC nội bộ · Grand Total từng cột · {matrix["sheet"]}', direction='columns',
                          target_range=range_text(matrix['line']+1, matrix['total'], matrix['average']-1, matrix['total']),
                          label_range=range_text(matrix['line']+1, matrix['header'], matrix['average']-1, matrix['header'])))
        rules.append(dict(kind='summary_comparison', name=f'Nguồn → QC · Line × Audit form · {matrix["sheet"]}',
                          left={'workbook': source['id'], 'sheet': table['sheet']},
                          right={'workbook': report['id'], 'sheet': matrix['sheet']}, source_table=table, matrix=matrix,
                          scope_factories=selected, period=list(period), number_mode='percent', tolerance='0',
                          aggregate_method='mean_reported_rows', rounding_digits='',
                          source_range=range_text(table['fields']['rate'], table['first'], table['fields']['rate'], table['last']),
                          target_range=rows_range))
        descriptions.append(f'{matrix["sheet"]}: Average từng dòng, Grand Total từng cột, và đối chiếu từng Line × Audit form với nguồn.')
    formula_errors = []
    for name, sheet in report['sheets'].items():
        errors = [compare.address(row, col) for (row, col), item in sheet['cells'].items()
                  if (item.get('formula') and '#REF!' in item['formula']) or str(item['value']).strip().upper() in compare.ERRORS]
        if errors:
            rules.append(dict(kind='formula_audit', name=f'Lỗi Excel / công thức · {name}', left={'workbook': report['id'], 'sheet': name}))
            formula_errors += [f'{name}!{cell}' for cell in errors]
    signatures = Counter(tuple(normalized(values[key]) for key in ('date', 'factory', 'group', 'line', 'form', 'rate')) for _, values in scoped)
    repeated = sum(count-1 for count in signatures.values() if count>1)
    warnings = [f'Chỉ đối chiếu phạm vi {selected[0]}. Các nhà máy còn lại trong nguồn không có ma trận QC tương ứng trong lần chạy này.',
                'Sai số 0: mọi chênh lệch đều được ghi nhận, kể cả chênh lệch do làm tròn. Trung bình bỏ trống / N/A và giữ số 0.',
                'Đối chiếu nguồn dùng trung bình MQAA% theo các dòng export, giữ nguyên dòng lặp. Chưa thay bằng trung bình theo lượt audit hoặc tỷ lệ tổng điểm.']
    if repeated:
        warnings.append(f'Có {repeated} dòng trùng chữ ký ngày / nhà máy / nhóm / line / form / tỷ lệ. Có thể là nhiều lỗi của cùng một audit; không tự xóa khi chưa có Audit ID.')
    if formula_errors:
        warnings.append(f'Phát hiện {len(formula_errors)} ô lỗi Excel / tham chiếu cần xem lại trong báo cáo QC.')
    return dict(profile='mqaa', profile_label='MQAA · kiểm nội bộ và đối chiếu nguồn', rules=rules,
                scope=dict(factories=selected, source_factories=source_factories, period=list(period),
                           source_sheet=table['sheet'], source_rows=len(records), scoped_rows=len(scoped),
                           repeated_signature_rows=repeated, matrices=[m['sheet'] for m in target_matrices]),
                warnings=warnings, descriptions=descriptions,
                exclusions=['Chưa kiểm tra ý nghĩa nội dung lỗi, phản hồi hoặc remark.',
                            'Không đánh đồng F1 / F2 / F3 / F4 trong VH với các nhà máy VH / VH2 / VH3 / VH4.',
                            'Không xác minh nguồn ngoài hoặc tính lại công thức Excel / liên kết ngoài.'])


def plain_tables(book):
    tables = []
    for name, sheet in book['sheets'].items():
        for row in range(1, min(sheet['rows'], 12)+1):
            populated = [(col, normalized(compare.cell(sheet, row, col)['value'])) for col in range(1, sheet['columns']+1)
                         if compare.cell(sheet, row, col)['value'] not in (None, '')]
            if len(populated)<2 or any(not isinstance(compare.cell(sheet,row,col)['value'],str) for col,_ in populated):
                continue
            mapping = dict((label,col) for col,label in populated)
            if len(mapping)!=len(populated) or row>=sheet['rows']:
                continue
            tables.append(dict(sheet=name,header=row,first=row+1,last=sheet['rows'],fields=mapping))
            break
    return tables


def generic_plan(source, report):
    left_tables, right_tables = plain_tables(source), plain_tables(report)
    rules, descriptions, used = [], [], set()
    for right in right_tables:
        candidates = [left for left in left_tables if set(left['fields'])==set(right['fields'])]
        named = [left for left in candidates if normalized(left['sheet'])==normalized(right['sheet'])]
        candidates = named or candidates
        if len(candidates)!=1 or candidates[0]['sheet'] in used:
            continue
        left = candidates[0]
        fields = list(right['fields'])
        preferred = sorted(fields, key=lambda key: (key not in {'id','code','mã','ma','key','sku','audit id'}, fields.index(key)))
        candidates = []
        for label in preferred:
            values = [compare.cell(book['sheets'][table['sheet']],row,table['fields'][label])['value']
                      for book,table in [(source,left),(report,right)] for row in range(table['first'],table['last']+1)]
            if label in {'id','code','mã','ma','key','sku','audit id'} or all(isinstance(value,str) for value in values):
                candidates.append(label)
        selected_keys = None
        for size in range(1,min(3,len(candidates))+1):
            for labels in combinations(candidates[:12],size):
                valid = True
                for book, table in [(source,left),(report,right)]:
                    keys = [tuple(normalized(compare.cell(book['sheets'][table['sheet']],row,table['fields'][label])['value']) for label in labels)
                            for row in range(table['first'],table['last']+1)]
                    if not keys or not all(all(part and part not in compare.MISSING for part in key) for key in keys) or len(set(keys))!=len(keys): valid=False;break
                if valid:
                    selected_keys=labels
                    break
            if selected_keys: break
        if selected_keys is None:
            continue
        values = [label for label in fields if label not in selected_keys]
        if not values: continue
        used.add(left['sheet'])
        rules.append(dict(kind='comparison',name=f'Bảng · {left["sheet"]} → {right["sheet"]}',
                          left={'workbook':source['id'],'sheet':left['sheet']}, right={'workbook':report['id'],'sheet':right['sheet']},
                          left_range=range_text(min(left['fields'].values()),left['first'],max(left['fields'].values()),left['last']),
                          right_range=range_text(min(right['fields'].values()),right['first'],max(right['fields'].values()),right['last']),
                          matching='keys',left_keys=','.join(get_column_letter(left['fields'][key]) for key in selected_keys),right_keys=','.join(get_column_letter(right['fields'][key]) for key in selected_keys),
                          left_columns=','.join(get_column_letter(left['fields'][label]) for label in values),
                          right_columns=','.join(get_column_letter(right['fields'][label]) for label in values),
                          number_mode='number',tolerance='0',rounding_digits='',value_type='auto',ignore_case=True))
        descriptions.append(f'{left["sheet"]} ↔ {right["sheet"]}: khóa đề xuất “{", ".join(selected_keys)}”; đối chiếu {len(values)} cột theo tên tiêu đề, không theo vị trí.')
    if not rules:
        raise ValueError('Chưa tìm được cặp bảng cùng tiêu đề và cột khóa duy nhất. Hãy thiết lập một quy tắc thủ công và lưu preset cho mẫu này.')
    skipped = [name for name in report['sheets'] if name not in {rule['right']['sheet'] for rule in rules}]
    return dict(profile='tables',profile_label='Bảng cùng cấu trúc · ghép theo tên cột',rules=rules,
                scope=dict(paired_tables=len(rules),skipped_sheets=skipped),descriptions=descriptions,
                warnings=['Khóa được đề xuất từ dữ liệu duy nhất, vẫn cần xác nhận đúng ý nghĩa nghiệp vụ trước khi chạy.']+
                         ([f'Chưa tự ghép được các sheet: {", ".join(skipped)}.'] if skipped else []),
                exclusions=['Chưa suy luận các phần có bố cục khác nhau hoặc ý nghĩa nghiệp vụ khác nhau.',
                            'Không tự đổi đơn vị, tính tổng, hoặc lấy trung bình cho bảng thông thường.'])


def preview(payload):
    if not isinstance(payload,dict): raise ValueError('Invalid workflow input.')
    source, report = compare.workbook(payload.get('source_id')), compare.workbook(payload.get('report_id'))
    if source['id']==report['id']: raise ValueError('Chọn hai file khác nhau cho nguồn và báo cáo.')
    requested = payload.get('profile','auto')
    if requested not in {'auto','mqaa','tables'}: raise ValueError('Unknown workflow profile.')
    is_mqaa = bool(raw_tables(source) and matrices(report))
    if requested=='auto' and not is_mqaa and matrices(source) and raw_tables(report):
        raise ValueError('Hai vai trò MQAA đang bị đảo: chọn file Combined làm nguồn và file có SUMMARY làm báo cáo cần kiểm tra.')
    plan = mqaa_plan(source,report,payload.get('factory_scope')) if requested=='mqaa' or (requested=='auto' and is_mqaa) else generic_plan(source,report)
    _,_,digits=compare.settings({'comparison_rounding_digits':payload.get('comparison_rounding_digits',1)})
    for rule in plan['rules']: rule['comparison_rounding_digits']=digits
    plan['comparison_rounding_digits']=digits
    if digits is not None:
        plan['warnings']=[warning for warning in plan['warnings'] if not warning.startswith('Sai số 0:')]
        plan['warnings'].insert(0,f'Làm tròn kết quả và giá trị báo cáo đến {digits} chữ số thập phân theo ROUND_HALF_UP trước khi so sánh; sai số sau làm tròn bằng 0. Không làm tròn các đầu vào trước khi tính trung bình.')
    if not 1<=len(plan['rules'])<=20: raise ValueError('Workflow vượt 20 quy tắc. Hãy chia nhỏ các sheet thành từng lượt đối chiếu.')
    plan.update(id=uuid4().hex,version=VERSION,inputs=[{'id':book['id'],'name':book['name'],'sha256':book['sha256']} for book in (source,report)],
                state='ready_for_review', stages=['Nhận diện cấu trúc','Kiểm tra phạm vi và phương pháp','Xác nhận và chạy','Rà soát chênh lệch','Xuất HTML'])
    compare.write_json(compare.ROOT/'workflows'/f'{plan["id"]}.json',plan)
    return plan


def execute(plan_id, acknowledged=False, reviewer='operator'):
    if acknowledged is not True: raise ValueError('Xác nhận phạm vi và phương pháp trên kế hoạch trước khi chạy.')
    try:
        plan=json.loads((compare.ROOT/'workflows'/f'{compare.identifier(plan_id)}.json').read_text(encoding='utf-8'))
    except FileNotFoundError as error:
        raise ValueError('Kế hoạch không còn tồn tại. Hãy nhận diện lại hai file.') from error
    if plan.get('version')!=VERSION:
        raise ValueError('Phiên bản workflow đã thay đổi. Hãy nhận diện lại để dùng quy tắc mới.')
    for item in plan['inputs']:
        book=compare.workbook(item['id'])
        path=compare.ROOT/'inputs'/book['id']/('source'+book['extension'])
        if hashlib.sha256(path.read_bytes()).hexdigest()!=item['sha256']:
            raise ValueError('Bản sao file đã thay đổi sau bước nhận diện. Hãy import lại và lập kế hoạch mới.')
    result=compare.run(plan['rules'])
    result['workflow']={key:value for key,value in plan.items() if key!='rules'}|{'state':'completed','review_acknowledged':True,'reviewer':reviewer}
    compare.write_json(compare.ROOT/'results'/f'{result["id"]}.json',result)
    return result


def summary_comparison(rule):
    left_book,left_sheet=compare.source(rule['left'])
    right_book,right_sheet=compare.source(rule['right'])
    _,tolerance,digits=compare.settings(rule)
    table,matrix=rule['source_table'],rule['matrix']
    if rule.get('aggregate_method')!='mean_reported_rows': raise ValueError('Unknown aggregation method.')
    grouped=defaultdict(list)
    for row,values in table_rows(left_book,table):
        if factory(values['factory']) in rule['scope_factories']:
            grouped[(normalized(values['line']),normalized(values['form']))].append((row,values))
    output=[]
    for row in range(matrix['first'],matrix['last']+1):
        line=compare.cell(right_sheet,row,matrix['line'])['value']
        for col in matrix['value_columns']:
            form=compare.cell(right_sheet,matrix['header'],col)['value']
            selected=grouped[(normalized(line),normalized(form))]
            inputs=[{'cell':compare.address(index,table['fields']['rate']),
                     **compare.cell(left_sheet,index,table['fields']['rate'])} for index,_ in selected]
            numeric=[compare.numeric(item,'percent') for item in inputs]
            numbers=[number for kind,number in numeric if kind=='number']
            invalid=any(kind not in {'number','missing'} for kind,_ in numeric)
            current=compare.cell(right_sheet,row,col)
            kind,actual=compare.numeric(current,'percent')
            expected=sum(numbers)/len(numbers) if numbers else None
            groups={normalized(values['group']) for _,values in selected}
            delta=None
            if len(groups)>1:
                status,reason='ambiguous_key','Một Line × Audit form xuất hiện ở nhiều Internal factory; chưa tự trộn nhóm.'
            elif invalid:
                status,reason='invalid_input','Nguồn có tỷ lệ không hợp lệ hoặc công thức chưa có giá trị.'
            elif kind in {'error','uncached'}:
                status,reason=('formula_error' if kind=='error' else 'formula_uncached'),'Ô báo cáo chứa lỗi / công thức chưa lưu kết quả.'
            elif expected is None:
                status,reason=('no_data' if kind=='missing' else 'missing_input'),'Không có tỷ lệ nguồn hợp lệ cho Line × Audit form trong phạm vi đã xác nhận.'
            elif kind=='missing':
                status,reason='missing_value',f'Có {len(numbers)} tỷ lệ nguồn hợp lệ nhưng ô QC trống / N/A.'
            elif kind!='number':
                status,reason='invalid_value','Tỷ lệ QC không được hiểu là phần trăm hợp lệ.'
            else:
                status,delta,reason=compare.compare_numbers(expected,actual,current,'percent',tolerance,digits)
            rounded_expected=compare.round_for_comparison(expected,'percent',digits)
            rounded_actual=compare.round_for_comparison(actual,'percent',digits) if kind=='number' else None
            output.append(compare.finding(rule,status,left_book,rule['left']['sheet'],inputs[0]['cell'] if inputs else '',
                other_file=right_book['name'],other_sheet=rule['right']['sheet'],other_cell=compare.address(row,col),
                label=f'{line} / {form}',expected=float(rounded_expected) if rounded_expected is not None else None,
                actual=float(rounded_actual) if digits is not None and rounded_actual is not None else current['value'],
                delta=float(delta) if delta is not None else None,inputs=inputs,operation='average',valid_count=len(numbers),
                ignored_count=len(inputs)-len(numbers),reason=reason+' Trung bình MQAA% theo dòng export; giữ nguyên dòng lặp.',
                source_role='aggregation_inputs', expected_exact=str(rounded_expected) if rounded_expected is not None else None,
                actual_exact=str(rounded_actual) if rounded_actual is not None else None, delta_exact=str(delta) if delta is not None else None,
                expected_unrounded=str(expected) if expected is not None else None,actual_unrounded=str(actual) if kind=='number' else None))
    return output


def formula_audit(rule):
    book,sheet=compare.source(rule['left'])
    return [compare.finding(rule,'formula_error',book,rule['left']['sheet'],compare.address(row,col),
                           label='Lỗi tham chiếu / Excel',expected=None,actual=item['value'],formula=item.get('formula'),delta=None,
                           reason='Kiểm tra liên kết / tham chiếu và tính lại file. Không sửa tự động.')
            for (row,col),item in sheet['cells'].items() if (item.get('formula') and '#REF!' in item['formula'])
            or str(item['value']).strip().upper() in compare.ERRORS]
