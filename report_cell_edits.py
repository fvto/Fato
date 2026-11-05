"""Edit one validated source cell, with a revision check and a recoverable backup."""

import json
import math
import os
from pathlib import Path
from secrets import token_bytes
import shutil
import subprocess
import tempfile
from decimal import InvalidOperation
from uuid import uuid4

from itsdangerous import BadData, URLSafeTimedSerializer
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from report_validation import OLE, CONTRACTS, blank, number, tables, validate_file


TOKENS = URLSafeTimedSerializer(token_bytes(32), salt='report-cell')
STALE = 'File đã thay đổi hoặc lượt kiểm tra đã hết hạn. Bấm Kiểm tra dữ liệu rồi thử lại.'

# Open only the disposable copy in a separate Excel instance. Retain the BIFF
# format and workbook features instead of reconstructing it from extracted rows.
EDIT_XLS = r"""
$ErrorActionPreference = 'Stop'
$edit = [Console]::In.ReadToEnd() | ConvertFrom-Json
$excel = New-Object -ComObject Excel.Application
try {
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Open($edit.path, 0, $false)
    try {
        if ($book.ReadOnly) { throw 'Workbook is read only' }
        $sheet = $book.Worksheets.Item([string]$edit.sheet)
        if ($edit.delete_rows.Count -gt 0) {
            foreach ($row in $edit.delete_rows) {
                $sheet.Rows.Item([int]$row).Delete() | Out-Null
            }
            foreach ($update in $edit.footer.PSObject.Properties) {
                $sheet.Range($update.Name).Value2 = [double]$update.Value
            }
        } elseif ($edit.value -is [string]) {
            $cell = $sheet.Range([string]$edit.address)
            $cell.Value2 = "'" + $edit.value
        } else {
            $sheet.Range([string]$edit.address).Value2 = [double]$edit.value
        }
        $book.Save()
    } finally {
        $book.Close($false)
    }
} finally {
    $excel.Quit()
}
"""


def edit_token(tool, group, period, path, revision, issue):
    return TOKENS.dumps(dict(tool=tool, group=group, period=period, path=str(path),
                             revision=revision, sheet=issue['sheet'], row=issue['row'],
                             column=issue['column'], cell_type=issue['cell_type']))


def write_cell(path, sheet, address, value, *, delete_rows=(), footer=None):
    with path.open('rb') as stream:
        legacy = stream.read(8) == OLE
    if legacy or delete_rows:
        if os.name != 'nt':
            raise ValueError('Thao tác này cần Microsoft Excel trên Windows.')
        excel_path = path
        if delete_rows and not legacy and path.suffix.lower() == '.xls':
            from zipfile import ZipFile

            with ZipFile(path) as archive:
                extension = '.xlsm' if 'xl/vbaProject.bin' in archive.namelist() else '.xlsx'
            excel_path = path.with_suffix(extension)
            path.rename(excel_path)
        payload = json.dumps(dict(path=str(excel_path), sheet=sheet, address=address, value=value,
                                  delete_rows=delete_rows, footer=footer or {}))
        try:
            subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', EDIT_XLS],
                           input=payload.encode('ascii'), capture_output=True, check=True, timeout=90,
                           creationflags=subprocess.CREATE_NO_WINDOW)
        except (subprocess.SubprocessError, OSError) as error:
            raise ValueError('Excel không lưu được thay đổi. Kiểm tra Microsoft Excel và quyền sửa sheet.') from error
        if excel_path != path:
            excel_path.replace(path)
        return
    try:
        from zipfile import ZipFile

        with ZipFile(path) as archive:
            has_macros = 'xl/vbaProject.bin' in archive.namelist()
        with path.open('rb') as stream:
            book = openpyxl.load_workbook(stream, keep_vba=has_macros, rich_text=True)
        try:
            cell = book[sheet][address]
            if book[sheet].protection.sheet and cell.protection.locked:
                raise ValueError('Ô Excel được bảo vệ.')
            cell.value = value
            if isinstance(value, str):
                cell.data_type = 's'
            book.save(path)
        finally:
            book.close()
    except Exception as error:
        raise ValueError('Không ghi được ô Excel. Kiểm tra định dạng và quyền sửa sheet.') from error


def footer_after_delete(path, report, sheet, deleted_rows):
    footer = next((item for item in report['excluded_summary_rows'] if item['sheet'] == sheet), None)
    if footer is None:
        return {}
    rows = dict(tables(path))[sheet]
    header = next(item for item in report['sheets'] if item['name'] == sheet)
    footer_row = footer['row']
    count_column = footer['count_column']
    new_footer_row = footer_row - len(deleted_rows)
    updates = {f'{count_column}{new_footer_row}': header['rows'] - len(deleted_rows)}
    for index, total in enumerate(rows[footer_row - 1]):
        column = get_column_letter(index + 1)
        if column == count_column or blank(total):
            continue
        removed_values = [rows[row - 1][index] for row in deleted_rows
                          if index < len(rows[row - 1]) and not blank(rows[row - 1][index])]
        if not removed_values:
            continue
        try:
            removed = sum(number(value) for value in removed_values)
            if not removed:
                continue
            original_total = sum(number(record[index]) for record in rows[header['header_row']:footer_row - 1]
                                 if index < len(record) and not blank(record[index]))
            if original_total != number(total):
                raise ValueError('unverified total')
        except (InvalidOperation, ValueError) as error:
            raise ValueError('Không xác minh được dòng tổng của Excel. Hãy xóa dòng và cập nhật tổng trực tiếp trong Excel.') from error
        updates[f'{column}{new_footer_row}'] = float(original_total - removed)
    return updates


def edit_cell(tool, token, value, backups, file_hash, write_json, now, *, delete_row=False, row_tokens=()):
    try:
        target = TOKENS.loads(token, max_age=3600)
        additional = [TOKENS.loads(item, max_age=3600) for item in row_tokens]
    except BadData as error:
        raise RuntimeError(STALE) from error
    if tool != target['tool']:
        raise ValueError('Ô không thuộc tool đang chọn.')
    scope = ('tool', 'group', 'period', 'path', 'revision', 'sheet')
    if any(any(item[key] != target[key] for key in scope) for item in additional):
        raise ValueError('Chọn các dòng trong cùng file và sheet từ một lượt kiểm tra.')
    deleted_rows = sorted({item['row'] for item in [target, *additional]}, reverse=True) if delete_row else []
    path = Path(target['path'])
    revision = target['revision']

    def check_revision():
        if (not path.is_file() or path.resolve() != path or
                path.stat().st_size > CONTRACTS['max_file_bytes'] or file_hash(path) != revision):
            raise RuntimeError(STALE)

    check_revision()
    sheet, row, column = target['sheet'], target['row'], target['column']
    if value is None and not delete_row:
        for title, rows in tables(path):
            if title == sheet:
                index = column_index_from_string(column) - 1
                current = rows[row - 1][index] if index < len(rows[row - 1]) else None
                check_revision()
                return dict(current='' if current is None else str(current))
        raise RuntimeError(STALE)

    if not delete_row:
        if not value.strip() or len(value) > 32767 or ILLEGAL_CHARACTERS_RE.search(value):
            raise ValueError('Nhập giá trị hợp lệ, tối đa 32.767 ký tự.')
        if value.lstrip().startswith('='):
            raise ValueError('Nhập giá trị trực tiếp, không nhập công thức.')
        if target['cell_type'] == 'number':
            try:
                value = float(number(value))
                if not math.isfinite(value):
                    raise ValueError('number out of range')
            except (ValueError, InvalidOperation, OverflowError) as error:
                raise ValueError('Nhập số hợp lệ theo yêu cầu của ô.') from error

    def validate(candidate):
        return validate_file(candidate, tool, target['group'], target['period'])

    with tempfile.TemporaryDirectory(prefix='.report-edit-', dir=path.parent) as folder:
        original = Path(folder) / 'original' / path.name
        original.parent.mkdir()
        shutil.copy2(path, original)
        if file_hash(original) != revision:
            raise RuntimeError(STALE)
        before = validate(original)
        edited = Path(folder) / path.name
        shutil.copy2(original, edited)
        footer = footer_after_delete(original, before, sheet, deleted_rows) if delete_row else None
        write_cell(edited, sheet, f'{column}{row}', value, delete_rows=deleted_rows, footer=footer)
        after = validate(edited)
        if delete_row:
            if (after['rows'] != before['rows'] - len(deleted_rows) or
                    after['error_count'] >= before['error_count'] or
                    len(after['excluded_summary_rows']) != len(before['excluded_summary_rows'])):
                raise ValueError('Không xác minh được kết quả xóa dòng. File nguồn chưa thay đổi.')
        elif (after['error_count'] >= before['error_count'] or
                any((issue['sheet'], issue['row'], issue['column']) == (sheet, row, column)
                    for issue in after['issues'])):
            raise ValueError('Giá trị mới chưa khắc phục được lỗi tại ô này. Kiểm tra lại dữ liệu cần nhập.')

        check_revision()
        entry = backups / uuid4().hex
        entry.mkdir(parents=True)
        try:
            shutil.copy2(original, entry / path.name)
            write_json(entry / 'edit.json', dict(time=now(), source=str(path), sheet=sheet,
                       action='delete_row' if delete_row else 'edit_cell', row=row,
                       rows=sorted(deleted_rows), cell=None if delete_row else f'{column}{row}',
                       footer_cells=list(footer or {}), before_sha256=revision, after_sha256=file_hash(edited)))
            check_revision()
            if path.with_name('~$' + path.name).exists():
                raise RuntimeError('File đang mở trong Excel. Đóng file rồi thử lưu lại.')
            edited.replace(path)
        except Exception:
            shutil.rmtree(entry)
            raise
        return dict(saved=True, backup_id=entry.name, deleted_count=len(deleted_rows))
