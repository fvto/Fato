"""Local Excel table append with explicit headers and row-level provenance."""
from datetime import date, datetime
from pathlib import Path
from uuid import uuid4
import hashlib
import io
import json
import logging
import math
import os
import re
import shutil
import subprocess
import unicodedata
import zipfile
from itertools import zip_longest
from xml.etree.ElementTree import ParseError

import openpyxl
import xlrd
from spreadsheet_combine_names import naming_details
from spreadsheet_combine_profiles import remove_known_footer
import spreadsheet_combine_csv as csv_input

ROOT = Path(__file__).resolve().parent / 'Report' / 'CombinedExcel'
MAX_BYTES = 200 * 1024 * 1024
MAX_TOTAL_BYTES = 5 * 1024 * 1024 * 1024
MAX_FILES = 150
XLSX_ROW_THRESHOLD = 1_000_000
MAX_CELLS = 500_000
STREAMING_EXPORT_CELLS = 50_000
SUPPORTED_EXTENSIONS = {'.xls', '.xlsx', '.xlsm', '.csv'}
SOURCE_COLUMNS = ['Source File', 'Source Sheet', 'Source Row']


def output_filename(value):
    """One Windows-safe XLSX name for storage, metadata and download headers."""
    if not isinstance(value, str) or not value.strip():
        raise ValueError('Nhập tên file xuất.')
    name = value.strip()
    if name.lower().endswith('.xlsx'):
        name = name[:-5]
    elif name.lower().endswith(('.xls', '.xlsm', '.csv', '.tsv')):
        raise ValueError('File gộp được xuất ở định dạng .xlsx. Hãy dùng đuôi .xlsx hoặc bỏ đuôi file.')
    if (not name or name.endswith(('.', ' ')) or len(name) > 160
            or re.search(r'[<>:"/\\|?*\x00-\x1f\x7f]', name)
            or re.fullmatch(r'(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?', name, re.I)):
        raise ValueError('Tên file không hợp lệ. Dùng tối đa 160 ký tự, không chứa \\ / : * ? " < > | hoặc tên dành riêng của Windows.')
    return name + '.xlsx'


def suggested_filename(sources):
    return output_filename(naming_details(sources)['output_name'])


def key(value):
    return ' '.join(unicodedata.normalize('NFKC', str(value or '')).split()).casefold()


def detect_header(sheet):
    """Rank single-row headers by width and support from rows underneath."""
    best, best_score = 1, -1
    rows = sheet['rows']
    for index, row in enumerate(rows[:50]):
        populated = [(i, v) for i, v in enumerate(row) if v not in (None, '')]
        if not populated or any(not isinstance(v, str) for _, v in populated):
            continue
        labels = [key(v) for _, v in populated]
        if len(labels) != len(set(labels)):
            continue
        supported = 0
        for following in rows[index + 1:index + 6]:
            filled = sum(i < len(following) and following[i] not in (None, '') for i, _ in populated)
            supported += filled / len(populated)
        score = len(labels) * 3 + supported - index * .01
        if score > best_score:
            best, best_score = index + 1, score
    return best


def folder_for(job_id):
    if not re.fullmatch(r'[a-f0-9]{32}', str(job_id or '')):
        raise ValueError('Không tìm thấy lượt gộp file.')
    folder = ROOT / job_id
    if not (folder / 'inputs.json').is_file():
        raise ValueError('Lượt gộp không còn tồn tại. Hãy chọn lại file.')
    return folder


def json_write(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, allow_nan=False), encoding='utf-8')


def read_sheets(path, csv_options=None):
    try:
        if Path(path).suffix.lower() == '.csv':
            sheet = csv_input.read_csv(Path(path), MAX_CELLS, csv_options)
            return [sheet] if any(any(v not in (None, '') for v in row) for row in sheet['rows']) else []
        return _read_sheets(path)
    except (xlrd.XLRDError, xlrd.compdoc.CompDocError, zipfile.BadZipFile, KeyError, IndexError, ParseError,
            openpyxl.utils.exceptions.InvalidFileException) as error:
        raise ValueError('Không đọc được cấu trúc Excel. Hãy mở và lưu lại bằng Excel trước khi gộp.') from error


def _read_sheets(path, cell_limit=MAX_CELLS):
    """Read values and formats only; never execute formulas or source instructions."""
    path = Path(path)
    with path.open('rb') as stream:
        signature = stream.read(8)
    sheets, total = [], 0
    if signature == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        import xlrd
        book = xlrd.open_workbook(str(path), formatting_info=True, logfile=io.StringIO())
        try:
            for sheet in book.sheets():
                total += sheet.nrows * sheet.ncols
                if cell_limit is not None and total > cell_limit:
                    raise ValueError('File vượt 500.000 ô. Hãy chia nhỏ dữ liệu.')
                rows, formats = [], []
                for row in sheet.get_rows():
                    values, styles = [], []
                    for cell in row:
                        value = cell.value
                        fmt = book.format_map[book.xf_list[cell.xf_index].format_key].format_str
                        if cell.ctype == xlrd.XL_CELL_DATE:
                            value = xlrd.xldate_as_datetime(value, book.datemode)
                        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                            value = bool(value)
                        elif cell.ctype == xlrd.XL_CELL_ERROR:
                            raise ValueError(f'{sheet.name}: có ô lỗi Excel. Hãy sửa file nguồn trước khi gộp.')
                        values.append(value)
                        styles.append(fmt)
                    rows.append(values)
                    formats.append(styles)
                sheets.append(dict(name=sheet.name, rows=rows, formats=formats))
        finally:
            book.release_resources()
    elif signature.startswith(b'PK'):
        with zipfile.ZipFile(path) as archive:
            if sum(item.file_size for item in archive.infolist()) > 160 * 1024 * 1024:
                raise ValueError('File sau giải nén quá lớn. Hãy chia nhỏ dữ liệu.')
        with path.open('rb') as stream, path.open('rb') as cached_stream:
            book = openpyxl.load_workbook(stream, read_only=True, data_only=False, keep_links=False)
            cached = openpyxl.load_workbook(cached_stream, read_only=True, data_only=True, keep_links=False)
            try:
                for sheet in book:
                    if cell_limit is not None and (sheet.max_row or 0) * (sheet.max_column or 0) > cell_limit:
                        raise ValueError('File vượt 500.000 ô. Hãy chia nhỏ dữ liệu.')
                    rows, formats = [], []
                    for row, saved in zip(sheet.iter_rows(), cached[sheet.title].iter_rows()):
                        total += len(row)
                        if cell_limit is not None and total > cell_limit:
                            raise ValueError('File vượt 500.000 ô. Hãy chia nhỏ dữ liệu.')
                        values = []
                        for cell, stored in zip(row, saved):
                            if cell.data_type == 'e' or stored.data_type == 'e':
                                raise ValueError(f'{sheet.title}!{cell.coordinate}: ô lỗi Excel.')
                            if cell.data_type == 'f' and stored.value is None:
                                raise ValueError(f'{sheet.title}!{cell.coordinate}: công thức chưa có kết quả đã lưu. Mở Excel, tính lại và lưu file trước khi gộp.')
                            values.append(stored.value if cell.data_type == 'f' else cell.value)
                        rows.append(values)
                        formats.append([cell.number_format for cell in row])
                    sheets.append(dict(name=sheet.title, rows=rows, formats=formats))
            finally:
                book.close()
                cached.close()
    else:
        raise ValueError('Nội dung file không phải Excel .xls/.xlsx hợp lệ.')
    return [sheet for sheet in sheets if any(any(v not in (None, '') for v in row) for row in sheet['rows'])]


class UploadValidationError(ValueError):
    def __init__(self, errors):
        self.errors = errors
        self.message = 'Có file không hợp lệ. Sửa hoặc bỏ các file được liệt kê rồi tải lại; chưa tạo lượt gộp.'
        super().__init__(self.message + '\n' + '\n'.join(f'{e["file"]}: {e["error"]}' for e in errors))


def upload(files, csv_options=None, progress=None):
    if not 1 <= len(files) <= MAX_FILES:
        raise ValueError(f'Chọn từ 1 đến {MAX_FILES} file Excel/CSV mỗi lần.')
    csv_options = csv_input.options(csv_options)
    job_id = uuid4().hex
    folder = ROOT / job_id
    folder.mkdir(parents=True)
    items, hashes, cells, errors = [], set(), 0, []
    total_bytes = 0
    try:
        for index, storage in enumerate(files):
            name = str(storage.filename or '').replace('\\', '/').rsplit('/', 1)[-1]
            extension = Path(name).suffix.lower()
            if progress:
                progress(f'Đọc file {index+1}/{len(files)}: {name}')
            try:
                if extension not in SUPPORTED_EXTENSIONS:
                    raise ValueError('Chỉ hỗ trợ .xls, .xlsx, .xlsm và .csv.')
                target = folder / f'input_{index}{extension}'
                digest, size = hashlib.sha256(), 0
                with target.open('wb') as stream:
                    while chunk := storage.stream.read(64 * 1024):
                        size += len(chunk)
                        total_bytes += len(chunk)
                        if total_bytes > MAX_TOTAL_BYTES:
                            raise ValueError('Tổng file vượt giới hạn 5 GB. Hãy chia nhỏ lượt gộp.')
                        if size > MAX_BYTES:
                            raise ValueError(f'Vượt giới hạn {MAX_BYTES // (1024 * 1024)} MB mỗi file.')
                        digest.update(chunk)
                        stream.write(chunk)
                digest = digest.hexdigest()
                if digest in hashes:
                    raise ValueError('File trùng nội dung đã chọn. Hãy bỏ file trùng rồi thử lại.')
                hashes.add(digest)
                if extension == '.csv':
                    from spreadsheet_combine_large import inspect_csv
                    metadata = inspect_csv(target, csv_options, progress, cache_scan=True)
                    cells += metadata['cells']
                    items.append(dict(id=index, name=name, sha256=digest, **metadata))
                    continue
                source_path = target
                if extension in {'.xls','.xlsx','.xlsm'} and size > 2 * 1024**2:
                    from spreadsheet_combine_large import inspect_xlsx
                    metadata = inspect_xlsx(target)
                    if not metadata:
                        raise ValueError('Không có dữ liệu.')
                    items.append(dict(id=index,name=name,stored=target.name,sha256=digest,
                                      sheets=metadata,large=True))
                    continue
                try:
                    sheets = read_sheets(source_path)
                except ValueError as error:
                    if extension not in {'.xls','.xlsx','.xlsm'} or not any(text in str(error) for text in ('500.000','sau giải nén')):
                        raise
                    from spreadsheet_combine_large import inspect_xlsx
                    metadata = inspect_xlsx(target)
                    if not metadata:
                        raise ValueError('Không có dữ liệu.')
                    items.append(dict(id=index,name=name,stored=target.name,sha256=digest,sheets=metadata,large=True))
                    continue
                try:
                    if not sheets:
                        raise ValueError('Không có dữ liệu.')
                    cells += sum(sum(max(1, len(row)) for row in s['rows']) for s in sheets)
                    item = dict(id=index, name=name, stored=target.name, sha256=digest,
                                sheets=[dict(name=s['name'], rows=len(s['rows']),
                                             columns=max(map(len, s['rows']), default=0),
                                             header_row=detect_header(s)) for s in sheets])
                    items.append(item)
                finally:
                    del sheets
            except (ValueError, OSError, zipfile.BadZipFile) as error:
                errors.append(dict(file=name, error=str(error)))
                if total_bytes > MAX_TOTAL_BYTES:
                    break
        if errors:
            raise UploadValidationError(errors)
        if cells > MAX_CELLS:
            for item in items:
                item['large'] = True
        json_write(folder / 'inputs.json', items)
        estimated=sum(max(0,s['rows']-s['header_row']) for item in items for s in item['sheets'] if s.get('default_selected',s is item['sheets'][0]))
        return dict(id=job_id, estimated_rows=estimated, xlsx_row_threshold=XLSX_ROW_THRESHOLD,
                    files=[{k: v for k, v in item.items() if k not in {'stored','original_stored'}} for item in items])
    except Exception:
        # This newly-created job folder contains only this failed upload.
        target=folder.resolve()
        if target.parent != ROOT.resolve() or not re.fullmatch('[a-f0-9]{32}',target.name):
            raise ValueError('Không dọn thư mục ngoài vùng gộp file.')
        shutil.rmtree(target)
        raise


def encode_cell(value, fmt):
    if isinstance(value, (datetime, date)):
        dt = datetime.combine(value, datetime.min.time()) if not isinstance(value, datetime) else value
        return dict(value=openpyxl.utils.datetime.to_excel(dt), format='yyyy-mm-dd' if dt.time() == datetime.min.time() else 'yyyy-mm-dd hh:mm:ss')
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError('Có giá trị số không hợp lệ trong file nguồn.')
    return dict(value=None if value == '' else value, format=fmt or 'General')


def clean_table(sheet, header_row, end_row=None, footer_mode='auto'):
    rows = sheet['rows']
    if not isinstance(header_row, int) or isinstance(header_row, bool) or not 1 <= header_row <= len(rows):
        raise ValueError('Dòng tiêu đề không hợp lệ.')
    end_row = len(rows) if end_row is None else end_row
    if type(end_row) is not int or not header_row <= end_row <= len(rows):
        raise ValueError('Dòng cuối phải nằm từ dòng tiêu đề đến dòng cuối trang tính.')
    if footer_mode not in {'auto', 'keep'}:
        raise ValueError('Cách xử lý dòng tổng không hợp lệ.')
    header = list(rows[header_row - 1])
    while header and header[-1] in (None, ''):
        header.pop()
    # Empty leading/spacer columns are removable only when the selected data
    # underneath is empty too. Never discard values under an unnamed column.
    positions = []
    for index, value in enumerate(header):
        if value in (None, ''):
            if any(index < len(row) and row[index] not in (None, '') for row in rows[header_row:end_row]):
                raise ValueError(f'{sheet["name"]}: có dữ liệu dưới cột chưa có tên.')
        else:
            positions.append(index)
    if not positions or any(not isinstance(header[i], str) or not header[i].strip() for i in positions):
        raise ValueError(f'{sheet["name"]}: tiêu đề phải là tên cột đầy đủ, không có ô trống hoặc ô gộp.')
    labels = [' '.join(header[i].split()) for i in positions]
    keys = [key(v) for v in labels]
    if len(set(keys)) != len(keys):
        raise ValueError(f'{sheet["name"]}: tên cột bị trùng.')
    records, blanks, repeated, footers = [], 0, 0, 0
    for index in range(header_row, end_row):
        row = rows[index]
        if not any(v not in (None, '') for v in row):
            blanks += 1
            continue
        if any(v not in (None, '') for v in row[len(header):]):
            raise ValueError(f'{sheet["name"]}, dòng {index + 1}: dữ liệu vượt số cột tiêu đề.')
        values = [row[i] if i < len(row) else None for i in positions]
        if [key(v) for v in values] == keys:
            repeated += 1
            continue
        formats = [sheet['formats'][index][i] if i < len(sheet['formats'][index]) else 'General' for i in positions]
        records.append((index + 1, values, formats))
    profile = ''
    if footer_mode == 'auto':
        records, footers, profile = remove_known_footer(keys, records)
    result = []
    for source_row, values, formats in records:
        cells = []
        for label, value, fmt in zip(keys, values, formats):
            if (re.search(r'(?:%|\brate\b|\bpercent(?:age)?\b|tỷ lệ)', label)
                    and isinstance(value, str) and re.fullmatch(r'[+-]?\d+(?:\.\d+)?%', value.strip())):
                raw = value.strip()[:-1]
                precision = len(raw.split('.')[1]) if '.' in raw else 0
                value, fmt = float(raw) / 100, '0' + ('.' + '0' * precision if precision else '') + '%'
            cells.append(encode_cell(value, fmt))
        result.append((source_row, cells))
    return dict(header=labels, keys=keys, records=result, blank_rows=blanks,
                repeated_headers=repeated, footer_rows=footers, profile=profile,
                header_row=header_row, end_row=end_row)


def prepare(job_id, selections, options=None, progress=None):
    options = {} if options is None else options
    if not isinstance(options, dict) or options.get('schema_mode', 'union') not in {'union', 'strict'} or options.get('footer_mode', 'auto') not in {'auto', 'keep'}:
        raise ValueError('Tùy chọn gộp không hợp lệ.')
    folder = folder_for(job_id)
    items = json.loads((folder / 'inputs.json').read_text(encoding='utf-8'))
    if any(item.get('large') for item in items):
        from spreadsheet_combine_large import prepare as prepare_large
        return prepare_large(job_id, selections, options, progress)
    if not isinstance(selections, list) or not selections:
        raise ValueError('Chọn ít nhất một bảng để gộp.')
    columns, column_keys, parts, seen = [], [], [], set()
    for selection in selections:
        if not isinstance(selection, dict):
            raise ValueError('Lựa chọn bảng không hợp lệ.')
        file_id = selection.get('file_id')
        if type(file_id) is not int or not 0 <= file_id < len(items):
            raise ValueError('File được chọn không hợp lệ.')
        item = items[file_id]
        path = folder / item['stored']
        try:
            with path.open('rb') as stream:
                digest = hashlib.file_digest(stream, 'sha256').hexdigest()
            if digest != item.get('stored_sha256', item['sha256']):
                raise ValueError('File nguồn đã thay đổi, hãy tải lại.')
            if item.get('converted'):
                with (folder / item['original_stored']).open('rb') as stream:
                    if hashlib.file_digest(stream,'sha256').hexdigest() != item['sha256']:
                        raise ValueError('File nguồn đã thay đổi, hãy tải lại.')
            sheets = read_sheets(path, item.get('csv_options')) if path.suffix == '.csv' else read_sheets(path)
            try:
                selected = next((s for s in sheets if s['name'] == selection.get('sheet')), None)
                if selected is None or (file_id, selected['name']) in seen:
                    raise ValueError('Trang tính bị trùng hoặc không còn tồn tại.')
                seen.add((file_id, selected['name']))
                table = clean_table(selected, selection.get('header_row', detect_header(selected)), selection.get('end_row'), options.get('footer_mode', 'auto'))
                sheet_name = selected['name']
            finally:
                selected = None
                del sheets
        except (ValueError, OSError) as error:
            raise ValueError(f'{item["name"]}: {error}') from error
        if options.get('schema_mode') == 'strict' and parts and set(table['keys']) != set(parts[0][2]['keys']):
            raise ValueError(f'{item["name"]} / {sheet_name}: bộ cột khác bảng đầu tiên. Chọn Giữ tất cả cột nếu muốn ghép các bảng khác cấu trúc.')
        for label, normalized in zip(table['header'], table['keys']):
            if normalized not in column_keys:
                column_keys.append(normalized)
                columns.append(label)
        parts.append((item, sheet_name, table))
    # Check the expanded union before allocating rows: 150 narrow sources can
    # otherwise create a huge sparse table despite a small input cell count.
    row_count = sum(len(part[2]['records']) for part in parts)
    if row_count * (len(columns) + 3) > MAX_CELLS or len(columns) + 3 > 16384:
        from spreadsheet_combine_large import prepare as prepare_large
        return prepare_large(job_id, selections, options, progress)
    provenance = []
    for label in SOURCE_COLUMNS:
        candidate, number = label, 1
        while key(candidate) in column_keys or candidate in provenance:
            candidate = f'{label} (combine {number})'
            number += 1
        provenance.append(candidate)
    rows, sources = [], []
    for item, sheet_name, table in parts:
        indexes = [column_keys.index(label) for label in table['keys']]
        for source_row, cells in table['records']:
            row = [dict(value=None, format='General') for _ in columns]
            for target, cell in zip(indexes, cells):
                row[target] = cell
            row += [dict(value=v, format='General') for v in (item['name'], sheet_name, source_row)]
            rows.append(row)
        sources.append(dict(file=item['name'], sheet=sheet_name, rows=len(table['records']),
                            columns=len(table['header']), blank_rows=table['blank_rows'],
                            repeated_headers=table['repeated_headers'], footer_rows=table['footer_rows'],
                            sha256=item['sha256'], header_row=table['header_row'], end_row=table['end_row'],
                            profile=table['profile'],
                            missing_columns=[columns[i] for i, label in enumerate(column_keys) if label not in table['keys']]))
    if not rows:
        raise ValueError('Không có dòng dữ liệu sau tiêu đề để gộp.')
    if len(rows) * (len(columns) + 3) > MAX_CELLS or len(columns) + 3 > 16384:
        raise ValueError('Bảng gộp vượt giới hạn 500.000 ô. Hãy chia nhỏ dữ liệu.')
    shared = set.intersection(*(set(part[2]['keys']) for part in parts))
    warnings = ['Các bảng không có cột chung. Kiểm tra lại trang tính và dòng tiêu đề trước khi xuất.'] if not shared and len(parts) > 1 else []
    for file_id in dict.fromkeys(selection['file_id'] for selection in selections):
        warnings.extend(f'{items[file_id]["name"]}: {warning}' for warning in items[file_id].get('warnings', []))
    columns += provenance
    return dict(columns=columns, rows=rows, sources=sources, source_columns=provenance,
                **naming_details(sources), warnings=warnings, options=options,
                row_count=len(rows), column_count=len(columns),
                file_count=len({item['file_id'] for item in selections}),
                footer_rows=sum(s['footer_rows'] for s in sources),
                repeated_headers=sum(s['repeated_headers'] for s in sources))


def summary(payload):
    result={k: v for k, v in payload.items() if k != 'rows'} | dict(
        export_mode='power_query' if payload['row_count']>XLSX_ROW_THRESHOLD else 'xlsx',
        xlsx_row_threshold=XLSX_ROW_THRESHOLD,
        preview=[[cell['value'] for cell in row] for row in payload['rows'][:12]],
        preview_formats=[[cell['format'] for cell in row] for row in payload['rows'][:12]])
    if payload['row_count']>XLSX_ROW_THRESHOLD:
        warnings=[message for message in payload.get('warnings',[]) if 'Kết quả vượt giới hạn dòng một sheet Excel' not in message]
        message='Kết quả trên 1 triệu dòng: chỉ xuất gói Power Query, không xuất toàn bộ dữ liệu thành sheet XLSX.'
        if message not in warnings:warnings.append(message)
        result['warnings']=warnings
    return result


def runtime():
    base = Path.home() / '.cache/codex-runtimes/codex-primary-runtime/dependencies'
    node = Path(os.environ.get('CHINGLUH_ARTIFACT_NODE', str(base / 'node/bin/node.exe')))
    modules = Path(os.environ.get('CHINGLUH_ARTIFACT_MODULES', str(base / 'node/node_modules')))
    if not node.is_file() or not (modules / '@oai/artifact-tool').exists():
        raise ValueError('Chưa tìm thấy bộ xuất Excel. Hãy cài runtime Codex hoặc cấu hình CHINGLUH_ARTIFACT_NODE và CHINGLUH_ARTIFACT_MODULES.')
    return node, modules


def export(job_id, selections, output_name=None, options=None, progress=None):
    payload = prepare(job_id, selections, options, progress)
    payload['output_name'] = output_filename(payload['output_name'] if output_name is None else output_name)
    if payload['row_count']>XLSX_ROW_THRESHOLD:
        from spreadsheet_combine_powerquery import export as export_powerquery
        return export_powerquery(job_id,payload,progress)
    folder = folder_for(job_id)
    # Each immutable result has its own directory, so concurrent requests cannot
    # overwrite another user's preview, payload or finished download.
    result_id = uuid4().hex
    result_folder = folder / result_id
    result_folder.mkdir()
    json_write(result_folder / 'payload.json', summary(payload) if payload.get('large') else payload)
    output = result_folder / payload['output_name']
    builder = Path(__file__).resolve().parent / 'scripts/spreadsheet_combine_export.mjs'
    try:
        if payload.get('large'):
            from spreadsheet_combine_large import write_verify
            payload['combined_sheets'] = write_verify(output, payload, progress)
        elif len(payload['rows']) * len(payload['columns']) > STREAMING_EXPORT_CELLS:
            from spreadsheet_combine_stream import write
            write(output, payload)
        else:
            node, modules = runtime()
            subprocess.run([str(node), str(builder), str(result_folder), str(modules)],
                           check=True, capture_output=True, text=True, encoding='utf-8',
                           timeout=180, creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
        if not payload.get('large'):
            verify(output, payload)
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as error:
        output.unlink(missing_ok=True)
        logging.getLogger(__name__).exception('Excel export failed; stderr=%s', getattr(error, 'stderr', ''))
        raise ValueError('Chưa xuất được file Excel. Hãy thử lại hoặc xem nhật ký Dashboard.') from error
    result = summary(payload) | dict(id=job_id, result_id=result_id,
                                    download_url=f'/api/combine/{job_id}/{result_id}/download')
    json_write(result_folder / 'result.json', result)
    return result


def verify(path, payload):
    book = openpyxl.load_workbook(path, read_only=True, data_only=False)
    try:
        if book.sheetnames != ['Combined', 'Sources']:
            raise ValueError('File gộp thiếu bảng dữ liệu hoặc thông tin nguồn.')
        sheet = book['Combined']
        header = next(sheet.iter_rows(values_only=True), ())
        if list(header) != payload['columns']:
            raise ValueError('Số dòng hoặc cột sau gộp không khớp nguồn.')
        for actual, expected in zip_longest(sheet.iter_rows(min_row=2), payload['rows']):
            if actual is None or expected is None or len(actual) != len(expected):
                raise ValueError('Số dòng hoặc cột sau gộp không khớp nguồn.')
            for cell, wanted in zip(actual, expected):
                value = cell.value
                if cell.data_type == 'f':
                    raise ValueError('File gộp có công thức ngoài dự kiến.')
                if isinstance(value, (datetime, date)):
                    value = openpyxl.utils.datetime.to_excel(value)
                target = wanted['value']
                same = math.isclose(value, target, rel_tol=1e-12, abs_tol=1e-9) if type(value) in (int, float) and type(target) in (int, float) else value == target
                if not same:
                    raise ValueError(f'Giá trị sau gộp không khớp tại {cell.coordinate}.')
    finally:
        book.close()
