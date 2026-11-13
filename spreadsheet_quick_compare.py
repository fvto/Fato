"""Compare occupied row counts and saved percentages at the same cell positions."""
from collections import Counter
from datetime import datetime
import re
from uuid import uuid4

import spreadsheet_compare as compare


def occupied_rows(sheet):
    # Count content, not formatting or the worksheet's last-used-row metadata.
    return len({row for (row, _), item in sheet['cells'].items()
                if item.get('formula') or item['value'] not in (None, '')})


def percentage(item):
    # Quoted/escaped percent signs are literal text in an Excel number format.
    number_format = re.sub(r'"[^"]*"|\\.', '', item.get('format', ''))
    value = item['value']
    return ('%' in number_format or
            (isinstance(value, str) and value.strip().endswith('%')
             and compare.numeric(item, 'percent')[0] == 'number'))


def percent_text(number):
    value = format(number * 100, 'f')
    return (value.rstrip('0').rstrip('.') if '.' in value else value) + '%'


def with_row_deltas(result):
    """Include explicit row differences, including for results saved before this fix."""
    summary = result.get('quick_summary')
    if not summary:
        return result
    summary['row_delta'] = summary['left_rows'] - summary['right_rows']
    for sheet in summary['sheets']:
        a, b = sheet['left_rows'], sheet['right_rows']
        sheet['row_delta'] = a - b if a is not None and b is not None else None
    for row in result['results']:
        if result['rules'][row['rule_index'] - 1]['kind'] != 'quick_rows':
            continue
        a, b = row.get('expected'), row.get('actual')
        if a is None or b is None:
            continue
        delta = a - b
        direction = (f'File A nhiều hơn {delta} dòng.' if delta > 0 else
                     f'File B nhiều hơn {-delta} dòng.' if delta < 0 else 'Số dòng giống nhau.')
        row.update(delta=delta, delta_exact=str(delta), delta_display=f'{abs(delta)} dòng',
                   reason=f'File A: {a} dòng; file B: {b} dòng. {direction} '
                          'Đếm dòng có nội dung, gồm tiêu đề và dòng tổng; bỏ dòng trống.')
    return result


def run(payload):
    if not isinstance(payload, dict):
        raise ValueError('Chọn hai file Excel cần kiểm tra.')
    left = compare.workbook(payload.get('source_id'))
    right = compare.workbook(payload.get('report_id'))
    if left['id'] == right['id']:
        raise ValueError('Chọn hai file khác nhau.')
    pairs = [(name, name) for name in left['sheets'] if name in right['sheets']]
    if not pairs and len(left['sheets']) == len(right['sheets']) == 1:
        pairs = [(next(iter(left['sheets'])), next(iter(right['sheets'])))]
    used_left, used_right = {a for a, _ in pairs}, {b for _, b in pairs}
    pairs += [(name, None) for name in left['sheets'] if name not in used_left]
    pairs += [(None, name) for name in right['sheets'] if name not in used_right]
    rules, findings, sheets = [], [], []
    percent_checked = percent_issues = unmatched_sheets = 0

    def record(rule, status, location='', **values):
        findings.append(compare.finding(rule, status, left if rule['left']['workbook'] == left['id'] else right,
            rule['left']['sheet'], location, rule_index=len(rules), number_mode=rule['number_mode'], **values))
        if len(findings) > compare.MAX_RESULTS:
            raise ValueError('Có hơn 50.000 mục kiểm tra. Hãy chia nhỏ file rồi kiểm tra lại; chưa lưu kết quả thiếu.')

    for left_name, right_name in pairs:
        left_sheet = left['sheets'].get(left_name)
        right_sheet = right['sheets'].get(right_name)
        rule = {'kind': 'quick_rows', 'name': 'Số dòng có nội dung', 'number_mode': 'number', 'tolerance': '0',
                'left': {'workbook': left['id'] if left_name else right['id'], 'sheet': left_name or right_name}}
        if left_name and right_name:
            rule['right'] = {'workbook': right['id'], 'sheet': right_name}
        rules.append(rule)
        count_left = occupied_rows(left_sheet) if left_sheet is not None else None
        count_right = occupied_rows(right_sheet) if right_sheet is not None else None
        matched = left_name is not None and right_name is not None
        status = ('match' if count_left == count_right else 'difference') if matched else 'unmatched'
        provenance = ({'other_file': right['name'], 'other_sheet': right_name, 'other_cell': ''} if matched else {})
        record(rule, status, expected=count_left, actual=count_right, label='Số dòng',
               reason=('Đếm dòng có nội dung, gồm tiêu đề và dòng tổng; bỏ dòng trống. '
                       + ('Số dòng giống nhau.' if status == 'match' else 'Số dòng khác nhau.')) if matched
                      else 'Trang tính chỉ có trong một file; chưa đối chiếu được các ô % ở trang này.', **provenance)
        sheets.append({'left_sheet': left_name, 'right_sheet': right_name, 'left_rows': count_left,
                       'right_rows': count_right, 'status': status})
        if not matched:
            unmatched_sheets += 1
            continue
        positions = {position for sheet in (left_sheet, right_sheet)
                     for position, item in sheet['cells'].items() if percentage(item)}
        if not positions:
            continue
        if len(findings) + len(positions) > compare.MAX_RESULTS:
            raise ValueError('Có hơn 50.000 mục kiểm tra. Hãy chia nhỏ file rồi kiểm tra lại; chưa lưu kết quả thiếu.')
        rule = {**rule, 'kind': 'quick_percent', 'name': 'Giá trị % cùng vị trí', 'number_mode': 'percent'}
        rules.append(rule)
        for row, col in sorted(positions):
            a, b = compare.cell(left_sheet, row, col), compare.cell(right_sheet, row, col)
            kind_a, number_a = compare.numeric(a, 'percent')
            kind_b, number_b = compare.numeric(b, 'percent')
            numbers = kind_a == kind_b == 'number'
            if 'uncached' in (kind_a, kind_b):
                status, reason = 'formula_uncached', 'Ô công thức chưa có kết quả đã lưu. Tính lại và lưu file trong Excel / WPS.'
            elif 'error' in (kind_a, kind_b):
                status, reason = 'formula_error', 'Một bên có lỗi Excel; chưa xác minh được giá trị %.'
            elif 'missing' in (kind_a, kind_b):
                status, reason = 'missing_value', ('Cả hai ô thiếu giá trị %; chưa xác minh được.' if kind_a == kind_b
                                                   else 'Một bên thiếu giá trị % ở cùng vị trí.')
            elif not numbers:
                status, reason = 'invalid_value', 'Một bên không phải giá trị % hợp lệ; số thường lớn hơn 1 không tự được hiểu thành %.'
            else:
                status = 'match' if number_a == number_b else 'difference'
                reason = 'Giá trị % giống nhau.' if status == 'match' else 'Giá trị % khác nhau; đối chiếu chính xác, không làm tròn.'
            extra = {}
            for key, kind, number, item in [('expected', kind_a, number_a, a), ('actual', kind_b, number_b, b)]:
                extra[key] = percent_text(number) if kind == 'number' else item['value']
                extra[key + '_display'] = extra[key]
                if kind == 'number':
                    extra[key + '_exact'] = str(number)
            delta = number_a - number_b if numbers else None
            if delta is not None:
                extra.update(delta_exact=str(delta), delta_display=percent_text(delta))
            record(rule, status, compare.address(row, col), label=f'Dòng {row} · cột {compare.get_column_letter(col)}',
                   reason=reason, delta=float(delta) if delta is not None else None,
                   formula=a.get('formula'), other_formula=b.get('formula'),
                   other_file=right['name'], other_sheet=right_name, other_cell=compare.address(row, col), **extra)
            percent_checked += 1
            percent_issues += status != 'match'
    summary = {'left_name': left['name'], 'right_name': right['name'],
               'source_id': left['id'], 'report_id': right['id'],
               'left_rows': sum(occupied_rows(sheet) for sheet in left['sheets'].values()),
               'right_rows': sum(occupied_rows(sheet) for sheet in right['sheets'].values()),
               'rows_equal': bool(sheets) and all(sheet['status'] == 'match' for sheet in sheets),
               'percent_checked': percent_checked, 'percent_issues': percent_issues,
               'percent_equal': False if percent_issues else True if percent_checked and not unmatched_sheets else None,
               'unmatched_sheets': unmatched_sheets, 'sheets': sheets,
               'method': 'Ghép trang tính cùng tên; nếu mỗi file chỉ có một trang tính thì ghép hai trang đó. '
                         'Đếm dòng có nội dung, gồm tiêu đề và dòng tổng, bỏ dòng trống. '
                         'So sánh ô có ký hiệu % hoặc định dạng phần trăm theo đúng vị trí dòng/cột, không làm tròn. '
                         'Không so sánh cột văn bản hoặc số thường, không tính lại trung bình hay tổng.'}
    result = {'id': uuid4().hex, 'created_at': datetime.now().isoformat(timespec='seconds'),
              'comparison_policy': 'exact', 'comparison_rounding_digits': [], 'rules': rules,
              'counts': dict(Counter(item['status'] for item in findings)), 'total': len(findings),
              'results': findings, 'quick_summary': summary}
    with_row_deltas(result)
    compare.write_json(compare.ROOT/'results'/f'{result["id"]}.json', result)
    return result
