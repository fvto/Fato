"""Self-contained, local HTML evidence reports with captured worksheet regions."""
import base64
from collections import Counter
from datetime import datetime, timedelta, timezone
from html import escape

from openpyxl.utils import coordinate_to_tuple, get_column_letter
import spreadsheet_compare as compare

REVIEW = {'difference', 'missing_value', 'missing_input', 'formula_error', 'formula_uncached',
          'invalid_input', 'invalid_value', 'ambiguous_key', 'invalid_key', 'unmatched'}
LABELS = {'difference': 'Khác giá trị', 'missing_value': 'Thiếu giá trị', 'missing_input': 'Thiếu dữ liệu đầu vào',
          'formula_error': 'Lỗi Excel', 'formula_uncached': 'Công thức chưa lưu kết quả',
          'invalid_input': 'Đầu vào không hợp lệ', 'invalid_value': 'Giá trị không hợp lệ',
          'ambiguous_key': 'Trùng khóa', 'invalid_key': 'Khóa không hợp lệ', 'unmatched': 'Dòng chỉ có ở một file'}
MAX_CAPTURES = 100
GUIDANCE = {
    'difference': 'Hai giá trị khác nhau theo quy tắc đã chọn. Đối chiếu dữ liệu gốc và xác nhận giá trị đúng trước khi cập nhật.',
    'missing_value': 'Một bên thiếu giá trị. Kiểm tra ô trống hoặc N/A và bổ sung theo nguồn đã xác nhận.',
    'missing_input': 'Có kết quả nhưng thiếu đầu vào để kiểm chứng. Kiểm tra vùng dữ liệu và cách tính.',
    'formula_error': 'Ô chứa lỗi Excel. Kiểm tra công thức và dữ liệu đầu vào, tính lại rồi lưu file.',
    'formula_uncached': 'Công thức chưa có kết quả được lưu. Mở Excel / WPS, tính lại, lưu và import lại file.',
    'invalid_input': 'Dữ liệu đầu vào không phù hợp quy tắc. Kiểm tra nội dung ô, kiểu số và các giá trị thiếu.',
    'invalid_value': 'Giá trị không được hiểu nhất quán ở hai bên. Kiểm tra kiểu dữ liệu và cách lưu phần trăm.',
    'ambiguous_key': 'Khóa ghép bị trùng nên chưa ghép dòng. Kiểm tra dòng trùng hoặc bổ sung cột khóa để phân biệt.',
    'invalid_key': 'Khóa ghép bị thiếu hoặc không hợp lệ. Kiểm tra mã nhận diện và công thức ở cột khóa.',
    'unmatched': 'Chỉ tìm thấy dòng ở một file. Kiểm tra dòng bị thiếu, mã khóa khác nhau hoặc phạm vi dữ liệu được chọn.',
}
CSS = """
*{box-sizing:border-box}body{margin:0;background:#f1f4f7;color:#172b42;font:14px/1.6 Arial,sans-serif}
main{max-width:1180px;margin:32px auto;padding:36px;background:white;border-radius:12px}
h1{font-size:30px;line-height:1.25;margin:10px 0}h2{margin-top:32px}h3{margin:0 0 12px}
.muted,small{color:#516276}.eyebrow{color:#156880;font-weight:bold;letter-spacing:2px}
.summary,.pair{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.pair{grid-template-columns:1fr 1fr}
.metric,.value,.scope{padding:16px;border:1px solid #ccd6df;border-radius:8px;overflow-wrap:anywhere}
.metric strong{display:block;font-size:28px}.note{padding:16px;border-left:4px solid #188497;background:#eef7f8}
article,figure{border:1px solid #ccd6df;border-radius:8px;padding:20px;margin:20px 0;break-inside:avoid}
.badge{display:inline-block;background:#fff0bb;color:#723e00;padding:3px 10px;border-radius:4px}
.value strong{display:block;font-size:22px;white-space:pre-wrap;overflow-wrap:anywhere}
.value.changed{background:#fff5cf;border-color:#e6b835}a{color:#00667c}img{display:block;max-width:100%;height:auto;border:1px solid #cbd5df;margin-top:14px}
table{border-collapse:collapse;width:100%;font-size:13px}td,th{padding:9px;border:1px solid #ccd6df;text-align:left;overflow-wrap:anywhere}
th{background:#edf2f7}code{overflow-wrap:anywhere}.scopes{display:grid;gap:12px}figcaption{overflow-wrap:anywhere}
@media(max-width:700px){main{margin:0;padding:20px}.summary,.pair{grid-template-columns:1fr}h1{font-size:25px}}
@media print{body{background:white}main{margin:0;max-width:none;padding:0}a{color:inherit}figure img{max-height:750px;object-fit:contain}h2{break-after:avoid}}
"""


def text(value):
    return escape('—' if value is None else str(value), quote=True)


def value_text(value, mode='number'):
    if isinstance(value, (float, int)) and not isinstance(value, bool) and mode != 'number':
        return f'{value * 100:.2f}%'
    return '—' if value is None else str(value)


def finding_value(finding, rule, side):
    key = 'expected' if side == 'left' else 'actual'
    mode = finding.get('number_mode', 'number')
    row_numbers = finding.get('left_rows' if side == 'left' else 'right_rows', [])
    if row_numbers and rule['kind'] == 'comparison':
        _, sheet = compare.source(rule[side])
        selected = compare.columns(rule[side + '_columns'])
        return '\n'.join(', '.join(f'{compare.address(row, col)} = {value_text(compare.cell(sheet, row, col)["value"], mode)}'
                                  for col in selected) for row in row_numbers)
    return value_text(finding.get(key), mode)


def capture_plan(result, issues):
    """Group highlighted cells into bounded regions, sharing images across findings."""
    tiles, references = {}, []
    for finding in issues:
        rule = result['rules'][finding['rule_index'] - 1]
        links = []
        for side, address_key, rows_key in [('left', 'cell', 'left_rows'), ('right', 'other_cell', 'right_rows')]:
            spec = rule.get(side)
            if not spec:
                continue
            points = []
            if finding.get(address_key):
                points.append((*coordinate_to_tuple(finding[address_key]), not (side=='left' and finding.get('source_role')=='aggregation_inputs')))
            if side=='left' and finding.get('inputs'):
                points += [(*coordinate_to_tuple(item['cell']),False) for item in finding['inputs']
                           if compare.numeric(item,finding.get('number_mode','number'))[0]!='missing']
            if finding.get(rows_key):
                selected = compare.columns(rule.get(side + '_columns'))
                if rule.get('matching', 'keys') == 'keys':
                    selected += compare.columns(rule.get(side + '_keys'))
                points += [(row, col, True) for row in finding[rows_key] for col in selected]
            for row, col, highlighted in points:
                key = (spec['workbook'], spec['sheet'], (row - 1)//8, (col - 1)//6)
                if key not in tiles:
                    if len(tiles) >= MAX_CAPTURES:
                        raise ValueError('Báo cáo vượt 100 vùng ảnh. Hãy chia nhỏ các quy tắc / vùng dữ liệu rồi xuất từng báo cáo HTML.')
                    book, sheet = compare.source(spec)
                    tiles[key] = dict(id=f'image-{len(tiles)+1}', book=book, sheet=sheet,
                                      sheet_name=spec['sheet'], row_start=key[2]*8+1,
                                      col_start=key[3]*6+1, highlights=set(), inputs=set())
                tile = tiles[key]
                tile['highlights' if highlighted else 'inputs'].add((row, col))
                link = (side, tile['id'])
                if link not in links:
                    links.append(link)
        references.append(links)
    return list(tiles.values()), references


def worksheet_value(item):
    value=item['value']
    if value is None and item.get('formula'): return '[Công thức chưa lưu kết quả]'
    if value is None: return ''
    if isinstance(value,(int,float)) and not isinstance(value,bool) and '%' in item.get('format',''):
        return f'{value*100:.3f}'.rstrip('0').rstrip('.')+'%'
    if isinstance(value,str) and value.strip().endswith('%'):
        kind,number=compare.numeric(item,'percent')
        if kind=='number': return f'{number*100:.3f}'.rstrip('0').rstrip('.')+'%'
    return value


def worksheet_html(tile):
    if tile.get('compact'):
        from spreadsheet_report import render_compact_grid
        return render_compact_grid(tile)
    sheet = tile['sheet']
    top, left = tile['row_start'], tile['col_start']
    bottom, right = min(top + 7, max(top, sheet['rows'])), min(left + 5, max(left, sheet['columns']))
    region = f'{get_column_letter(left)}{top}:{get_column_letter(right)}{bottom}'
    heading = f"{tile['book']['name']} · {tile['sheet_name']} · {region}"
    rows = ['<tr><th></th>' + ''.join(f'<th>{get_column_letter(col)}</th>' for col in range(left, right+1)) + '</tr>']
    for row in range(top, bottom+1):
        cells = []
        for col in range(left, right+1):
            item = compare.cell(sheet, row, col)
            value = worksheet_value(item)
            highlighted = (row, col) in tile['highlights']
            address = compare.address(row, col)
            cell_class = 'highlight' if highlighted else 'input' if (row,col) in tile.get('inputs',set()) else ''
            cells.append(f'<td class="{cell_class}"><small>{address}</small><span>{text(value)}</span></td>')
        rows.append(f'<tr><th>{row}</th>' + ''.join(cells) + '</tr>')
    return f'''<!doctype html><html lang="vi"><meta charset="utf-8"><style>
body{{margin:0;font:14px/1.45 Arial,sans-serif;color:#172b42;background:white}}
#capture{{width:1080px;padding:24px}}h2{{font-size:18px;margin:0 0 8px;overflow-wrap:anywhere}}
p{{color:#526274;font-size:12px;margin:0 0 16px}}table{{border-collapse:collapse;width:100%;table-layout:fixed}}
th,td{{border:1px solid #b6c3d0;padding:9px;vertical-align:top;overflow-wrap:anywhere;white-space:pre-wrap}}
th{{background:#edf2f7;font-size:12px}}th:first-child{{width:48px}}small{{display:block;color:#526274;font-size:10px;margin-bottom:5px}}
.highlight{{background:#ffe49a;box-shadow:inset 0 0 0 2px #c67b00}}.highlight span{{font-weight:bold}}
.input{{background:#e2f0fd;box-shadow:inset 0 0 0 1px #8bb5d8}}
</style><div id="capture"><h2>{text(heading)}</h2><p>Vùng dữ liệu dựng lại từ file Excel · vàng: điểm cần kiểm tra · xanh: đầu vào để tính / đối chiếu · hiển thị giá trị đã lưu</p><table>{''.join(rows)}</table></div></html>''', region


def capture_images(tiles):
    if not tiles:
        return {}
    from playwright.sync_api import sync_playwright, Error
    try:
        with sync_playwright() as playwright:
            try:
                browser = playwright.chromium.launch(channel='chrome', headless=True)
            except Error:
                browser = playwright.chromium.launch(headless=True)
            try:
                context = browser.new_context(viewport={'width': 1120, 'height': 900}, java_script_enabled=False)
                context.route('**/*', lambda route: route.abort())
                page = context.new_page()
                images = {}
                for tile in tiles:
                    source, region = worksheet_html(tile)
                    page.set_content(source, wait_until='load')
                    png = page.locator('#capture').screenshot(type='png', animations='disabled')
                    images[tile['id']] = (base64.b64encode(png).decode('ascii'), region)
                return images
            finally:
                browser.close()
    except Error as error:
        raise ValueError('Không tạo được ảnh báo cáo. Hãy kiểm tra Chrome trên máy rồi thử xuất HTML lại.') from error


def build_report(result):
    from spreadsheet_report import render_html
    return render_html(result)


def json_scope(scope):
    if 'factories' in scope:
        return f'Nhà máy {", ".join(scope["factories"])}; kỳ {" đến ".join(scope["period"])}; {scope["scoped_rows"]}/{scope["source_rows"]} dòng nguồn trong phạm vi; ma trận {", ".join(scope["matrices"])}.'
    return f'{scope.get("paired_tables", 0)} cặp bảng; sheet chưa ghép: {", ".join(scope.get("skipped_sheets", [])) or "không"}.'
