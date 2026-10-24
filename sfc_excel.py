"""Read and combine SFC exports in the application runtime without losing rows."""
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote
import hashlib
import io
import json
import openpyxl
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill

EXPORT_ROOT = Path(__file__).resolve().parent / 'Report' / 'HFPAExports'


def read_export(path):
    path = Path(path)
    with path.open('rb') as stream:
        signature = stream.read(8)
    sheets = []
    if signature.startswith(b'PK'):
        with path.open('rb') as stream:
            workbook = openpyxl.load_workbook(stream, read_only=True, data_only=False)
            try:
                sheets = [(sheet.title, list(sheet.iter_rows(values_only=True))) for sheet in workbook]
            finally:
                workbook.close()
    elif signature == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        try:
            import xlrd
        except ImportError as error:
            raise RuntimeError('Cần thư viện xlrd để đọc file .xls từ SFC: python -m pip install xlrd') from error
        workbook = xlrd.open_workbook(str(path), logfile=io.StringIO())
        for sheet in workbook.sheets():
            rows = []
            for r in range(sheet.nrows):
                row = []
                for cell in sheet.row(r):
                    row.append(xlrd.xldate_as_datetime(cell.value, workbook.datemode)
                               if cell.ctype == xlrd.XL_CELL_DATE else cell.value)
                rows.append(tuple(row))
            sheets.append((sheet.name, rows))
        workbook.release_resources()
    else:
        raise ValueError(f'{path.name}: máy chủ không trả về file Excel hợp lệ (có thể là trang lỗi/đăng nhập).')
    tables = [(name, rows) for name, rows in sheets if any(any(v is not None and v != '' for v in row) for row in rows)]
    if len(tables) != 1:
        raise ValueError(f'{path.name}: cần một bảng dữ liệu để gộp, nhận được {len(tables)} bảng; giữ nguyên file nguồn để kiểm tra.')
    _, rows = tables[0]
    rows = [tuple(row) for row in rows if any(v is not None and v != '' for v in row)]
    header = rows[0]
    data = rows[1:]
    # DevExpress appends a totals row (blank group/ID, count under Audit Date).
    # Keep that row in the original download, never append it as an audit record.
    labels = [str(value or '').lower() for value in header]
    date_col = next((i for i, label in enumerate(labels) if 'audit date' in label), None)
    id_col = next((i for i, label in enumerate(labels) if 'audit id' in label), None)
    if data and date_col is not None and id_col is not None:
        tail = data[-1]
        if not tail[0] and not tail[id_col] and re_summary_count(tail[date_col]):
            data = data[:-1]
    return header, data


def re_summary_count(value):
    return str(value).replace(',', '').replace('.', '').strip().isdigit()


def validate_export(path, date_from=None, date_to=None, nike=None):
    header, rows = read_export(path)
    labels = [str(v or '').strip() for v in header]
    if len(labels) < 5 or len(set(labels) - {''}) < 5:
        raise ValueError('File xuất không có tiêu đề dữ liệu HFPA hợp lệ.')
    internal_columns = {'SO#', 'Country', 'Pass/Fail'}
    if nike is True and (not any('Factory Group' in label for label in labels)
                         or internal_columns.intersection(labels)):
        raise ValueError('File tải về không đúng mẫu NIKE (External).')
    if nike is False and not internal_columns.issubset(labels):
        raise ValueError('File tải về không đúng mẫu nội bộ (thiếu SO#, Country hoặc Pass/Fail).')
    date_column = next((i for i, label in enumerate(labels) if 'audit date' in label.lower()), None)
    if date_from and date_column is None:
        raise ValueError('Không tìm thấy cột Audit Date để kiểm tra kỳ xuất.')
    first = date.fromisoformat(date_from) if date_from else None
    last = date.fromisoformat(date_to) if date_to else None
    for row in rows:
        if date_column is None:
            break
        value = row[date_column]
        if isinstance(value, datetime):
            value = value.date()
        elif not isinstance(value, date):
            raw = str(value).strip()
            parsed = None
            for fmt in ('%d-%b-%Y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y'):
                try:
                    parsed = datetime.strptime(raw, fmt).date()
                    break
                except ValueError:
                    pass
            if parsed is None:
                raise ValueError(f'Không đọc được Audit Date: {raw!r}')
            value = parsed
        if first and not first <= value <= last:
            raise ValueError(f'File có ngày {value} nằm ngoài kỳ yêu cầu {first} → {last}.')
    return len(rows)


def file_info(path, rows=None):
    path = Path(path).resolve()
    relative = path.relative_to(EXPORT_ROOT.resolve()).as_posix()
    result = dict(name=path.name, path=str(path), url='/api/exports/' + quote(relative, safe='/'),
                  size=path.stat().st_size, sha256=hashlib.sha256(path.read_bytes()).hexdigest())
    if rows is not None:
        result['rows'] = rows
    return result


def safe_cell(sheet, value):
    cell = WriteOnlyCell(sheet, value=value)
    # Exported strings are data, including any text beginning with '='.
    if isinstance(value, str):
        cell.data_type = 's'
    return cell


def combine_exports(files, path, plant, nike, include_sources=True):
    if not files:
        raise ValueError('Không có file nguồn để gộp.')
    for item in files:
        if item.get('plant') != plant or item.get('nike') is not nike:
            raise ValueError('Không được gộp lẫn nhà máy hoặc chế độ NIKE/nội bộ.')
    # Preflight before opening write streams; a schema error must leave no
    # half-written workbook and must never shift data into the wrong columns.
    source_schemas = []
    expected_count = 0
    for item in files:
        header, rows = read_export(item['path'])
        source_schemas.append((item['name'], header, bool(rows)))
        expected_count += len(rows)
    populated = [header for _, header, has_rows in source_schemas if has_rows]
    expected_header = populated[0] if populated else max((h for _, h, _ in source_schemas), key=len)
    for name, header, has_rows in source_schemas:
        # SFC omits the defect columns on days with zero records. An empty
        # prefix schema is safe; any change affecting actual rows still stops.
        if header != expected_header and (has_rows or header != expected_header[:len(header)]):
            raise ValueError(f'{name}: cấu trúc cột thay đổi; dừng gộp để tránh lệch dữ liệu.')
    if expected_count > 1048575:
        raise ValueError('Vượt giới hạn dòng Excel; các file tháng vẫn được giữ nguyên.')
    workbook = openpyxl.Workbook(write_only=True)
    data_sheet = workbook.create_sheet('HFPA')
    data_sheet.freeze_panes = 'A2'
    schema = None
    count = 0
    audit_rows = []
    for item in files:
        header, rows = read_export(item['path'])
        if schema is None:
            schema = expected_header
            cells = []
            for i, label in enumerate(schema, 1):
                cell = safe_cell(data_sheet, label)
                cell.font = Font(name='Calibri', bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor='17365D')
                cell.alignment = Alignment(wrap_text=True, vertical='center')
                cells.append(cell)
                data_sheet.column_dimensions[openpyxl.utils.get_column_letter(i)].width = min(36, max(16, len(str(label)) * .8))
            data_sheet.row_dimensions[1].height = 45
            data_sheet.append(cells)
        elif header != schema and rows:
            raise ValueError(f'{item["name"]}: cấu trúc cột thay đổi; dừng gộp để tránh lệch dữ liệu.')
        for row in rows:
            if count >= 1048575:
                raise ValueError('Vượt giới hạn dòng Excel; các file tháng vẫn được giữ nguyên.')
            data_sheet.append([safe_cell(data_sheet, value) for value in row])
            count += 1
        audit_rows.append([plant, 'NIKE (External)' if nike else 'Nội bộ', item['date_from'], item['date_to'], item['name'], len(rows), item['sha256']])
    data_sheet.auto_filter.ref = f'A1:{openpyxl.utils.get_column_letter(len(schema))}{count + 1}'
    if include_sources:
        audit = workbook.create_sheet('NguonTheoThang')
        audit.freeze_panes = 'A2'
        for col, width in [('A',12),('B',23),('C',15),('D',15),('E',62),('F',14),('G',68)]:
            audit.column_dimensions[col].width = width
        audit.append(['Nhà máy', 'Chế độ', 'Từ ngày', 'Đến ngày', 'File nguồn', 'Số dòng', 'SHA256'])
        for row in audit_rows:
            audit.append([safe_cell(audit, v) for v in row])
    path = Path(path)
    pending = path.with_suffix('.partial.xlsx')
    workbook.save(pending)
    # Reopen and reconcile the saved artifact before publishing its final name.
    check = openpyxl.load_workbook(pending, read_only=True, data_only=True)
    try:
        actual = sum(1 for _ in check['HFPA'].iter_rows(min_row=2, values_only=True))
        expected = sum(row[5] for row in audit_rows)
        if actual != expected or actual != count:
            raise ValueError(f'Gộp chưa đủ dòng: {actual}/{expected}')
    finally:
        check.close()
    pending.replace(path)
    return file_info(path, count)
