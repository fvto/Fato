"""Compact exact-comparison reports shared by HTML and PowerPoint exports."""
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from html import escape
from collections import Counter
from openpyxl.utils import coordinate_to_tuple

import spreadsheet_compare as compare
from spreadsheet_compare_html import REVIEW, LABELS, capture_plan, capture_images, json_scope, finding_value

CSS = '''*{box-sizing:border-box}body{margin:0;color:#172b42;background:#f1f4f7;font:13px/1.5 Arial,sans-serif}
main{max-width:1200px;margin:24px auto;padding:28px;background:white}h1{font-size:25px;margin:0 0 8px}p{margin:8px 0}small,.muted{color:#526274}
table{border-collapse:collapse;width:100%;margin:12px 0}td,th{padding:8px;border:1px solid #ccd6df;text-align:left;vertical-align:top;overflow-wrap:anywhere}th{background:#edf2f7}
details{margin:10px 0}summary{cursor:pointer;color:#00667c}a{color:#00667c}figure{margin:18px 0;break-inside:avoid}img{max-width:100%;height:auto;display:block}figcaption{overflow-wrap:anywhere}
.group-table>tbody>tr>td:last-child{width:55%}.detail-table{font-size:12px}.detail-table td{padding:6px}.method{border-top:1px solid #ccd6df;padding-top:10px}.note{font-size:12px;color:#526274}
@media(max-width:700px){main{margin:0;padding:16px}.group-table>tbody>tr>td:last-child{width:auto}td,th{padding:5px}.detail-table{font-size:11px}}
@media print{main{margin:0;padding:0}body{background:white}details:not([open])>*:not(summary){display:none}tr{break-inside:avoid}}'''


def text(value):
    return escape('—' if value is None else str(value),quote=True)


def exact_value(row,side):
    key='actual' if side=='report' else 'expected'
    value=row.get(key+'_exact')
    raw=row.get(key)
    if value is None:
        return '—' if raw is None else str(raw)
    number=Decimal(value)
    if row.get('number_mode','number')!='number':
        return decimal_text(number*100)+'%'
    return decimal_text(number)


def decimal_text(number):
    number=number.normalize()
    return str(number) if number and abs(number)<Decimal('0.000001') else format(number,'f')


def delta_value(row):
    value=row.get('delta_exact')
    if value is None: value=row.get('delta')
    if value is None: return '—'
    number=Decimal(str(value))
    if row.get('number_mode','number')!='number':
        return decimal_text(number*100)+'%'
    return decimal_text(number)


def display_value(value,is_delta=False):
    raw=str(value)
    suffix='%' if raw.endswith('%') else ''
    try:
        number=Decimal(raw[:-len(suffix)] if suffix else raw)
    except InvalidOperation:
        return raw
    if not number.is_finite(): return raw
    if is_delta and suffix and number and abs(number)<Decimal('0.0005'):
        return f'{number:.2E}%'
    formatted=f'{number:.3f}'.rstrip('0').rstrip('.')
    if formatted=='-0': formatted='0'
    if suffix: return formatted+suffix
    return format(number,'f') if number==number.to_integral_value() else formatted


def prepare(result,images=True):
    if result.get('comparison_policy') not in {'exact','rounded'}:
        raise ValueError('Kết quả cũ có thể dùng sai số / làm tròn. Hãy chạy lại với sai số 0 trước khi xuất báo cáo.')
    issues=[row for row in result['results'] if row['status'] in REVIEW]
    issues=compare.qc_ordered_findings(issues,result['rules'])
    tiles,references=capture_plan(result,issues)
    groups={}
    for index,(row,links) in enumerate(zip(issues,references),1):
        rule=result['rules'][row['rule_index']-1]
        sheet=row.get('other_sheet') or row['sheet']
        spec=rule.get('right') or rule['left']
        key=(spec['workbook'],sheet)
        if key not in groups:
            groups[key]={'name':sheet,'sheet':sheet,'status':'','rows':[]}
        location=row.get('other_cell') or row.get('cell') or ', '.join(map(str,row.get('right_rows') or row.get('left_rows',[])))
        reported=finding_value(row,rule,'right') if row.get('right_rows') else exact_value(row,'report')
        expected=finding_value(row,rule,'left') if row.get('left_rows') else exact_value(row,'source')
        groups[key]['rows'].append({'index':index,'label':str(row.get('label') or ''),'location':f'{sheet}!{location}',
                                   'reported':display_value(reported),'expected':display_value(expected),
                                   'delta':display_value(delta_value(row),is_delta=True),
                                   'reported_exact':reported,'expected_exact':expected,'delta_exact':delta_value(row),
                                   'reported_unrounded':row.get('actual_unrounded'),'expected_unrounded':row.get('expected_unrounded'),
                                   'reason':row.get('reason',''),'status':LABELS[row['status']],
                                   'images':[image_id for _,image_id in links]})
    for group in groups.values():
        group['status']='; '.join(f'{label}: {count}' for label,count in Counter(row['status'] for row in group['rows']).items())
    workflow=result.get('workflow')
    scope=json_scope(workflow['scope']) if workflow else 'Theo các sheet / vùng dữ liệu đã chọn.'
    return dict(result=result,issues=issues,groups=list(groups.values()),tiles=tiles,
                images=capture_images(tiles) if images else {},scope=scope)


def policy_label(result):
    digits=result.get('comparison_rounding_digits',[])
    return f'Làm tròn {", ".join(map(str,digits))} số lẻ trước khi so sánh · sai số 0' if digits else 'So sánh giá trị đầy đủ · sai số 0'


def compact_tiles(tiles):
    """Merge worksheet evidence into larger compact regions for a shorter slide appendix."""
    sheets={}
    for tile in tiles:
        key=(tile['book']['id'],tile['sheet_name'])
        if key not in sheets:
            sheets[key]={**tile,'highlights':set(),'inputs':set()}
        sheets[key]['highlights'].update(tile['highlights'])
        sheets[key]['inputs'].update(tile.get('inputs',set()))
    result=[]
    for sheet in sheets.values():
        if not sheet['highlights']: continue
        points=sheet['highlights']|sheet['inputs']
        min_col,max_col=min(col for _,col in points),max(col for _,col in points)
        left=1 if min_col<=18 else min_col
        right=max_col
        for block in sorted({(row-1)//30 for row,_ in sheet['highlights']}):
            for col_block in range(left,right+1,18):
                highlights={(row,col) for row,col in sheet['highlights'] if block*30+1<=row<=(block+1)*30 and col_block<=col<col_block+18}
                if not highlights: continue
                result.append({**sheet,'id':f'image-{len(result)+1}','compact':True,'row_start':block*30+1,
                               'row_end':min((block+1)*30,sheet['sheet']['rows']),
                               'col_start':col_block,'col_end':min(col_block+17,right)})
    return result


def render_compact_grid(tile):
    from openpyxl.utils import get_column_letter
    from spreadsheet_compare_html import worksheet_value
    top,bottom,left,right=tile['row_start'],tile['row_end'],tile['col_start'],tile['col_end']
    region=f'{get_column_letter(left)}{top}:{get_column_letter(right)}{bottom}'
    rows=['<tr><th></th>'+''.join(f'<th>{get_column_letter(col)}</th>' for col in range(left,right+1))+'</tr>']
    for row in range(top,bottom+1):
        cells=[]
        for col in range(left,right+1):
            item=compare.cell(tile['sheet'],row,col)
            value=worksheet_value(item)
            # Neighbouring long comments are context; full values remain in the report data/notes.
            if (row,col) not in tile['highlights'] and len(str(value))>80: value=str(value)[:77]+'…'
            style='highlight' if (row,col) in tile['highlights'] else 'input' if (row,col) in tile['inputs'] else ''
            cells.append(f'<td class="{style}">{text(value)}</td>')
        rows.append(f'<tr><th>{row}</th>'+''.join(cells)+'</tr>')
    title=f'{tile["sheet_name"]} · {region}'
    return f'''<!doctype html><html lang="vi"><meta charset="utf-8"><style>body{{margin:0;font:13px/1.25 Arial,sans-serif;color:#172b42}}#capture{{width:1472px;padding:12px}}h2{{font-size:18px;margin:0 0 10px}}table{{border-collapse:collapse;table-layout:fixed;width:100%}}th,td{{border:1px solid #b6c3d0;padding:5px;overflow-wrap:anywhere}}th{{background:#edf2f7}}th:first-child{{width:36px}}.highlight{{background:#ffe49a;font-weight:bold}}.input{{background:#e2f0fd}}</style><div id="capture"><h2>{text(title)}</h2><table>{''.join(rows)}</table></div></html>''',region


def presentation_model(result):
    model=prepare(result,images=False)
    tiles=compact_tiles(model['tiles'])
    images=capture_images(tiles)
    return {'title':'Báo cáo chênh lệch','scope':model['scope'],'result_id':result['id'],'policy':policy_label(result),
            'count':len(model['issues']),'groups':model['groups'],'comparison_policy':policy_label(result),
            'findings':sorted([row for group in model['groups'] for row in group['rows']],key=lambda row:row['index']),
            'sources':result.get('workflow',{}).get('inputs',[]),
            'evidence':[{'id':tile['id'],'caption':f'{tile["sheet_name"]} · {images[tile["id"]][1]}',
                         'png':images[tile['id']][0]} for tile in tiles]}


def render_html(result):
    model=prepare(result)
    grouped=[]
    for group in model['groups']:
        rows=[]
        for row in group['rows']:
            links=' '.join(f'<a href="#{image_id}">{image_id.removeprefix("image-")}</a>' for image_id in dict.fromkeys(row['images']))
            rows.append(f'<tr id="finding-{row["index"]}"><td>{row["index"]}. {text(row["label"])}<br><small>{text(row["location"])}</small></td><td title="{text(row["reported_exact"])}">{text(row["reported"])}</td><td title="{text(row["expected_exact"])}">{text(row["expected"])}</td><td title="{text(row["delta_exact"])}">{text(row["delta"])}</td><td>{links}</td></tr>')
        detail='<table class="detail-table"><thead><tr><th>Chỉ tiêu / vị trí</th><th>Báo cáo</th><th>Tính lại / nguồn</th><th>Chênh lệch</th><th>Ảnh</th></tr></thead><tbody>'+''.join(rows)+'</tbody></table>'
        grouped.append(f'<tr class="issue-group"><td>{text(group["name"])}</td><td>{text(group["status"])}</td><td>{len(rows)}</td><td><details class="finding-group"><summary>Xem {len(rows)} vị trí</summary>{detail}</details></td></tr>')
    if not grouped: grouped=['<tr><td colspan="4">Không có điểm khác nhau theo quy tắc đã chọn.</td></tr>']
    figures=[]
    for tile in model['tiles']:
        png,region=model['images'][tile['id']]
        caption=f'{tile["book"]["name"]} · {tile["sheet_name"]} · {region}'
        figures.append(f'<figure id="{tile["id"]}"><figcaption>Ảnh {tile["id"].removeprefix("image-")}: {text(caption)}</figcaption><img alt="{text(caption)}" src="data:image/png;base64,{png}"></figure>')
    books={}
    for rule in result['rules']:
        for side in ('left','right'):
            if rule.get(side):
                book,_=compare.source(rule[side]);books[book['id']]=book
    sources=''.join(f'<li>{text(book["name"])}<br><small>SHA-256: {text(book["sha256"])}</small></li>' for book in books.values())
    workflow=result.get('workflow') or {}
    notes=workflow.get('warnings',[])+workflow.get('exclusions',[])
    extra=''.join(f'<li>{text(note)}</li>' for note in notes)
    exported=datetime.now(timezone(timedelta(hours=7))).strftime('%d/%m/%Y %H:%M UTC+7')
    return f'''<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'"><title>Báo cáo chênh lệch</title><style>{CSS}</style></head><body><main>
<h1>Báo cáo chênh lệch</h1><p><b>{len(model['issues'])} mục khác nhau</b> · {len(model['groups'])} nhóm · <b>Sai số 0</b> · {exported}</p>
<p class="muted">{text(model['scope'])}</p><p class="note">Hiển thị tối đa 3 số lẻ, bỏ số 0 thừa. {text(policy_label(result))}.</p>
<table class="group-table"><thead><tr><th>Phần kiểm tra</th><th>Vấn đề</th><th>Số mục</th><th>Chi tiết</th></tr></thead><tbody>{''.join(grouped)}</tbody></table>
<details id="evidence"><summary>Ảnh highlight ({len(figures)})</summary><p class="note">Vàng: ô cần kiểm tra. Xanh: đầu vào. Ảnh dựng từ dữ liệu Excel, đã nhúng để xem offline.</p>{''.join(figures)}</details>
<details class="method"><summary>Nguồn và phương pháp</summary><p>Mã kết quả: {text(result['id'])}. {text(policy_label(result))}.</p><ul>{sources}</ul><ul>{extra}</ul></details>
</main></body></html>'''
