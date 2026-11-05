"""Append complete monthly COPQ source exports to the configured Dashboard workbook."""

from calendar import monthrange
from datetime import date, datetime
from hashlib import sha256
import os
from pathlib import Path
import re
from tempfile import NamedTemporaryFile, gettempdir
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter
from openpyxl.utils.datetime import to_excel

from dashboard_paths import APP_ROOT
from report_validation import parse_date, validate_file
from sfc_excel import read_export
from sfc_report_tools import plant_of, read_json


PLANTS = {'VH', 'VH2', 'VH3', 'VH4', 'JV', 'JV2', 'JVB'}
SHEET = 'CoPQ Dashboard database'
SHEET_XML = 'xl/worksheets/sheet1.xml'
WORKBOOK_XML = 'xl/workbook.xml'
DATE_VALUE = re.compile(rb'<c r="B\d+"(?: s="\d+")?(?: t="n")?><v>(\d+(?:\.\d+)?)</v></c>')


def configured_workbook():
    config = APP_ROOT / 'report_tools_config.json'
    value = read_json(config).get('copq_dashboard_workbook') if config.is_file() else None
    return APP_ROOT / value if value else None


def xml_cell(column, row, style, value):
    start = f'<c r="{column}{row}" s="{style}"'
    if value is None:
        return start + '/>'
    if isinstance(value, str):
        if re.search(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', value):
            raise ValueError('Dữ liệu COPQ có ký tự Excel không hỗ trợ.')
        return start + f' t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'
    if isinstance(value, bool):
        return start + f' t="b"><v>{int(value)}</v></c>'
    if isinstance(value, (date, datetime)):
        value = to_excel(value)
    return start + f'><v>{value}</v></c>'


def read_month(paths, headers, period):
    rows = []
    for path in paths:
        report = validate_file(path, 'copq', 'copq', period)
        if not report['valid']:
            raise ValueError('Có file nguồn COPQ chưa đạt kiểm tra dữ liệu tháng.')
        header, records = read_export(path)
        records = [header, *records]
        header_index = next((i for i, values in enumerate(records[:10])
                             if set(headers[1:]) <= {str(value).strip() for value in values}), None)
        if header_index is None:
            raise ValueError('File nguồn thiếu cột của Dashboard Trúc.')
        columns = {str(value).strip(): i for i, value in enumerate(records[header_index]) if value is not None}
        source_rows = records[header_index+1:header_index+1+report['rows']]
        if len(records) - header_index - 1 != report['rows'] + len(report['excluded_summary_rows']):
            raise ValueError('Số dòng nguồn COPQ khác kết quả kiểm tra.')
        site = plant_of(path.name)
        for source_row in source_rows:
            values = [source_row[columns[h]] if columns[h] < len(source_row) else None for h in headers[1:]]
            if parse_date(values[0]).strftime('%Y-%m') != period:
                raise ValueError('Dòng COPQ nằm ngoài tháng đã chọn.')
            if not isinstance(values[0], (date, datetime)):
                values[0] = parse_date(values[0])
            rows.append((site, *values))
    return sorted(rows, key=lambda row: parse_date(row[1]))


def append_month(target, paths, period):
    if len(paths) != len(PLANTS) or {plant_of(path.name) for path in paths} != PLANTS:
        return 'File tổng COPQ chờ đủ dữ liệu 7 nhà máy.'
    if not target.is_file():
        raise ValueError('Không tìm thấy file tổng Dashboard Trúc đã cấu hình.')
    stamp = target.stat().st_mtime_ns
    with target.open('rb') as stream:
        book = load_workbook(stream, read_only=True)
        try:
            if book.sheetnames != [SHEET] or book.epoch != datetime(1899, 12, 30):
                raise ValueError('File tổng không đúng mẫu Dashboard Trúc.')
            ws = book[SHEET]
            old_end = ws.max_row
            headers = [str(value or '').strip() for value in next(ws.iter_rows(max_row=1, max_col=21, values_only=True))]
        finally:
            book.close()
    if headers[:2] != ['Site', 'Date'] or not all(headers):
        raise ValueError('Tiêu đề file tổng không đúng mẫu Dashboard Trúc.')

    first = date.fromisoformat(period + '-01')
    last = first.replace(day=monthrange(first.year, first.month)[1])
    begin, end = to_excel(first), to_excel(last) + 1
    with ZipFile(target) as archive:
        prefix, rest = archive.read(SHEET_XML).split(b'<sheetData>', 1)
        body, footer = rest.rsplit(b'</sheetData>', 1)
        existing = seen = 0
        for match in DATE_VALUE.finditer(body):
            seen += 1
            existing += begin <= float(match[1]) < end
        if seen != old_end - 1:
            raise ValueError('Không xác minh được toàn bộ cột Date trong file tổng.')
        if existing:
            return f'File tổng COPQ đã có {existing:,} dòng tháng {period}.'

        rows = read_month(paths, headers, period)
        if not rows:
            return f'Tháng {period} không có dòng COPQ để thêm.'
        new_end = old_end + len(rows)
        last_row = body[body.rindex(f'<row r="{old_end}"'.encode()):]
        styles = dict(re.findall(rb'<c r="([A-Z]+)\d+" s="(\d+)"', last_row))
        columns = [(col, styles.get(col.encode(), b'0').decode()) for col in map(get_column_letter, range(1,22))]
        if columns[1][1] == '0':
            raise ValueError('Không xác định được định dạng ngày của file tổng.')

        def update_range(data, pattern):
            result, count = re.subn(pattern, lambda m: m[1] + str(new_end).encode(), data)
            if count != 1:
                raise ValueError('Không xác định được phạm vi dữ liệu trong file tổng.')
            return result

        prefix = update_range(prefix, rb'(<dimension ref="A1:[A-Z]+)\d+(?=")')
        footer = update_range(footer, rb'(<autoFilter ref="A1:U)\d+(?=")')
        workbook_xml = update_range(archive.read(WORKBOOK_XML), rb'(\$A\$1:\$U\$)\d+')
        with NamedTemporaryFile(dir=target.parent, prefix='.copq-history-', suffix='.xlsx', delete=False) as tmp:
            temporary = Path(tmp.name)
        try:
            with ZipFile(temporary, 'w', ZIP_DEFLATED, compresslevel=4) as result:
                for name in archive.namelist():
                    if name == SHEET_XML:
                        with result.open(name, 'w') as stream:
                            for block in (prefix, b'<sheetData>', body):
                                stream.write(block)
                            for index, row in enumerate(rows, old_end + 1):
                                cells = ''.join(xml_cell(col, index, style, value) for (col, style), value in zip(columns, row))
                                stream.write(f'<row r="{index}" spans="1:21">{cells}</row>'.encode())
                            stream.write(b'</sheetData>' + footer)
                    else:
                        result.writestr(name, workbook_xml if name == WORKBOOK_XML else archive.read(name))
            archive.close()
            if target.stat().st_mtime_ns != stamp:
                raise PermissionError('File tổng vừa thay đổi.')
            temporary.replace(target)
        finally:
            temporary.unlink(missing_ok=True)
    return f'Đã thêm {len(rows):,} dòng tháng {period} vào file tổng COPQ.'


def update_history(paths, period, *, verify_export_names=False):
    try:
        target = configured_workbook()
        if target is None:
            return ''
        if verify_export_names and len(paths) == len(PLANTS):
            first = date.fromisoformat(period + '-01')
            last = first.replace(day=monthrange(first.year, first.month)[1])
            name = re.compile(rf'Cost Of Poor Quality\s+{first}_{last}-(?:VH[234]?|JV[23B]?)\.(?:xls|xlsx|xlsm)', re.I)
            if any(not name.fullmatch(path.name) for path in paths):
                return 'File tổng COPQ chờ bộ file xuất đầy đủ đúng tháng từ SFC.'
        key = sha256(os.fsencode(target.resolve())).hexdigest()[:20]
        lock_path = Path(gettempdir()) / f'chingluh-copq-{key}.lock'
        with lock_path.open('ab') as lock:
            if os.name == 'nt':
                import msvcrt
                lock.seek(0)
                msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
            return append_month(target, paths, period)
    except PermissionError:
        return 'Chưa cập nhật được file tổng COPQ. Đóng file Excel rồi xuất lại tháng này.'
    except Exception as error:
        return f'Chưa cập nhật file tổng COPQ: {error}'


def update_direct_export(files, filters):
    dates = {key.rsplit('_', 1)[-1].lower(): value['value'] for key, value in filters.items() if value['kind'] == 'date'}
    first = dates.get('dtdatefrom') or dates.get('dtfromdate')
    last = dates.get('dtdateto') or dates.get('dttodate')
    if not first or not last:
        return ''
    start = date.fromisoformat(first)
    if start.day != 1 or last != start.replace(day=monthrange(start.year, start.month)[1]).isoformat():
        return ''
    scope = {key.rsplit('_', 1)[-1].lower(): value['value'] for key, value in filters.items() if value['kind'] == 'combo'}
    if any(key not in scope or scope[key] not in (None, '', 'ALL') for key in ('cbocategory', 'cbotype')):
        return ''
    return update_history([Path(file['path']) for file in files], first[:7])
