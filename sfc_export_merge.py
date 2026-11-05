"""Combine completed factory exports without changing or discarding source rows."""
from datetime import datetime
from pathlib import Path
import hashlib

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill

from sfc_excel import file_info, safe_cell

METADATA_SHEETS = {'nguontheongay', 'nguontheothang', 'sources'}
PLANTS = {'VH', 'VH2', 'VH3', 'VH4', 'JV', 'JV2', 'JVB'}


def read_sheets(path):
    path = Path(path)
    with path.open('rb') as stream:
        signature = stream.read(8)
    if signature.startswith(b'PK'):
        with path.open('rb') as stream:
            book = openpyxl.load_workbook(stream, read_only=True, data_only=False)
            try:
                return [(sheet.title, [tuple(row) for row in sheet.iter_rows(values_only=True)]) for sheet in book]
            finally:
                book.close()
    if signature == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        import xlrd
        book = xlrd.open_workbook(str(path))
        try:
            return [(sheet.name, [tuple(
                xlrd.xldate_as_datetime(cell.value, book.datemode)
                if cell.ctype == xlrd.XL_CELL_DATE else cell.value
                for cell in sheet.row(index)) for index in range(sheet.nrows)])
                for sheet in book.sheets()]
        finally:
            book.release_resources()
    raise ValueError(f'{path.name}: chưa hỗ trợ gộp định dạng này; file nguồn vẫn được giữ nguyên.')


def trim(row):
    row = list(row)
    while row and row[-1] in (None, ''):
        row.pop()
    return tuple(row)


def table(rows, plant):
    """Recognize MQAA's preamble explicitly; never append its summary as an audit."""
    nonempty = [trim(row) for row in rows if any(value not in (None, '') for value in row)]
    if not nonempty:
        return None
    if str(nonempty[0][0]).strip().lower() == 'mqaa report - factory':
        period = next((str(row[0]) for row in nonempty[1:] if str(row[0]).lower().startswith('from:')), '')
        fields = ('Factory', 'Avg Rate', 'Line', 'Issue')
        metrics = {}
        for row in nonempty[1:]:
            key = str(row[0] or '').strip()
            if key in fields:
                if key in metrics:
                    raise ValueError('Báo cáo Factory lặp nhãn chỉ tiêu; chưa gộp để tránh mất dữ liệu.')
                metrics[key] = row[1:]
            elif not key.lower().startswith('from:'):
                raise ValueError('Báo cáo Factory có chỉ tiêu chưa nhận diện: ' + key)
        header = ('Report Period',) + fields + ('Status',)
        width = max((len(values) for values in metrics.values()), default=0)
        records = []
        for index in range(width):
            values = tuple(metrics.get(key, ())[index] if index < len(metrics.get(key, ())) else None for key in fields)
            if any(value not in (None, '') for value in values):
                records.append((period,) + values + ('Data',))
        return header, records or [(period, None, None, None, None, 'No data')]
    for index, row in enumerate(nonempty[:12]):
        labels = [str(value or '').strip().lower() for value in row]
        if {'audit date', 'audit form', 'factory'} <= set(labels):
            records = []
            factory_col = labels.index('factory')
            percent_col = labels.index('mqaa%') if 'mqaa%' in labels else None
            for record in nonempty[index + 1:]:
                populated = {i for i, value in enumerate(record) if value not in (None, '')}
                if percent_col is not None and populated == {percent_col}:
                    continue
                actual = str(record[factory_col] or '').strip().upper() if len(record) > factory_col else ''
                if actual == 'JV3':
                    actual = 'JVB'
                if actual != plant:
                    raise ValueError(f'File {plant} chứa dữ liệu nhà máy {actual or "không xác định"}; chưa gộp.')
                if len(record) > len(row):
                    raise ValueError('Dữ liệu MQAA vượt số cột tiêu đề; chưa gộp.')
                records.append(record + (None,) * (len(row) - len(record)))
            return row, records
    # Plain single-row headers are only appended with matching source schemas.
    header = nonempty[0]
    if (len(header) >= 2 and all(isinstance(value, str) and value.strip() for value in header)
            and len(set(header)) == len(header)
            and all(len(row) <= len(header) for row in nonempty[1:])):
        return header, [row + (None,) * (len(header) - len(row)) for row in nonempty[1:]]
    return None


def combine_factory_exports(files, target, expected_plants):
    expected_plants = list(expected_plants)
    actual = [item.get('plant') for item in files]
    if (not files or len(set(expected_plants)) != len(expected_plants)
            or not set(expected_plants) <= PLANTS or actual != expected_plants):
        raise ValueError('Chưa đủ file của các nhà máy đã chọn; chưa tạo file gộp.')
    sources = []
    tables = []
    for item in files:
        if item.get('sha256') and hashlib.sha256(Path(item['path']).read_bytes()).hexdigest() != item['sha256']:
            raise ValueError(f'{item["name"]}: file nguồn đã thay đổi; chưa gộp.')
        sheets = read_sheets(item['path'])
        data = [(name, rows) for name, rows in sheets if name.lower() not in METADATA_SHEETS
                and any(any(value not in (None, '') for value in row) for row in rows)]
        if not data:
            raise ValueError(f'{item["name"]}: không có bảng Excel để gộp.')
        sources.append((item, sheets))
        tables.append(table(data[0][1], item['plant']) if len(data) == 1 else None)
    tabular = all(part is not None for part in tables)
    if tabular and any(part[0] != tables[0][0] for part in tables):
        raise ValueError('Cấu trúc cột giữa các nhà máy khác nhau; giữ file nguồn, chưa gộp để tránh lệch dữ liệu.')
    if tabular:
        header = ('Source Factory', 'Source File') + tables[0][0]
        records = [(item['plant'], item['name']) + row for item, part in zip(files, tables) for row in part[1]]
    else:
        width = max(len(row) for _, sheets in sources for name, rows in sheets
                    if name.lower() not in METADATA_SHEETS for row in rows)
        header = ('Source Factory', 'Source File', 'Source Sheet', 'Source Row') + tuple(f'Column {i + 1}' for i in range(width))
        records = [(item['plant'], item['name'], name, index) + tuple(row) + (None,) * (width - len(row))
                   for item, sheets in sources for name, rows in sheets if name.lower() not in METADATA_SHEETS
                   for index, row in enumerate(rows, 1)]
    count = len(records)
    if count > 1048575 or len(header) > 16384:
        raise ValueError('Dữ liệu vượt giới hạn Excel; giữ các file nguồn để kiểm tra.')
    book = openpyxl.Workbook(write_only=True)
    combined = book.create_sheet('Combined')
    combined.freeze_panes = 'C2'
    combined.row_dimensions[1].height = 32
    widths = {'Source File': 38, 'Report Period': 38, 'Factory': 22, 'Line': 22, 'Issue': 74}
    for index, label in enumerate(header, 1):
        combined.column_dimensions[openpyxl.utils.get_column_letter(index)].width = (
            widths.get(label, min(42, max(16, len(str(label)) * .9))))
    cells = [safe_cell(combined, value) for value in header]
    for cell in cells:
        cell.font = Font(name='Calibri', bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='17365D')
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    cells[0].comment = Comment('\n'.join(f'{item["plant"]}: {item["name"]}\nSHA256: {item.get("sha256", "")}' for item in files), 'SFC Dashboard')
    combined.append(cells)
    for row_index, row in enumerate(records, 2):
        if 'Issue' in header:
            issue = row[header.index('Issue')]
            lines = sum(max(1, (len(line) + 66) // 67) for line in str(issue or '').split('\n'))
            combined.row_dimensions[row_index].height = min(409, max(18, lines * 15))
        cells = [safe_cell(combined, value) for value in row]
        for index, value in enumerate(row):
            if isinstance(value, datetime):
                cells[index].number_format = 'yyyy-mm-dd'
            elif str(header[index]).lower() in ('mqaa%', 'avg rate') and isinstance(value, (int, float)):
                cells[index].number_format = '0.0%'
            if header[index] == 'Issue':
                cells[index].alignment = Alignment(wrap_text=True, vertical='top')
        combined.append(cells)
    combined.auto_filter.ref = f'A1:{openpyxl.utils.get_column_letter(len(header))}{count + 1}'
    target = Path(target)
    pending = target.with_suffix('.partial.xlsx')
    try:
        book.save(pending)
        check = openpyxl.load_workbook(pending, read_only=True)
        try:
            if check.sheetnames != ['Combined'] or sum(1 for _ in check['Combined'].iter_rows(min_row=2)) != count:
                raise ValueError('Số dòng sau gộp không khớp file nguồn; chưa công bố file.')
            if {row[0] for row in records} != {item['plant'] for item, part in zip(files, tables) if not tabular or part[1]}:
                raise ValueError('Thiếu nhà máy trong file gộp; chưa công bố file.')
        finally:
            check.close()
        pending.replace(target)
    finally:
        pending.unlink(missing_ok=True)
    return dict(file_info(target, count), plants=expected_plants,
                merge_layout='combined_table' if tabular else 'stacked_rows')
