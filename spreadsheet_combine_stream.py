"""Write large, already-combined payloads without a second in-memory workbook."""
import warnings

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn


SOURCE_HEADERS = ['File', 'Sheet', 'Data Rows', 'Source Columns', 'Header Row', 'End Row',
                  'Blank Rows Skipped', 'Repeated Headers Skipped', 'Verified Totals Skipped',
                  'Columns Absent in Source', 'Recognized Layout', 'SHA256']


def source_values(source):
    return [source['file'], source['sheet'], source['rows'], source['columns'],
            source['header_row'], source['end_row'], source['blank_rows'],
            source['repeated_headers'], source['footer_rows'],
            ', '.join(source.get('missing_columns', [])), source.get('profile') or '', source['sha256']]


def write(path, payload):
    book = Workbook(write_only=True)
    font = Font(name='Arial', size=10, color='202020')
    header_font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    alignment = Alignment(vertical='center')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    fill = PatternFill('solid', fgColor='263D4A')

    def table(name, headers, rows, count, sample):
        sheet = book.create_sheet(name)
        sheet.sheet_view.showGridLines = False
        sheet.sheet_format.defaultRowHeight = 20
        sheet.row_dimensions[1].height = 46
        sheet.freeze_panes = 'A2'
        if name == 'Combined':
            sheet.sheet_properties.tabColor = '263D4A'
        for index, label in enumerate(headers):
            sample_length = max((len(str(row[index]['value'])) for row in sample
                                 if isinstance(row[index]['value'], str)), default=0)
            width = max(14, min(52, max(len(label) * .8, sample_length * .9)))
            if label.casefold() == 'file' or 'source file' in label.casefold():
                width = 46
            if 'sha256' in label.casefold():
                width = 72
            sheet.column_dimensions[get_column_letter(index + 1)].width = width
        heading = []
        for label in headers:
            cell = WriteOnlyCell(sheet, value=label)
            cell.data_type = 's'
            cell.font, cell.alignment, cell.fill = header_font, header_alignment, fill
            heading.append(cell)
        sheet.append(heading)
        for row in rows:
            output = []
            for source in row:
                cell = WriteOnlyCell(sheet, value=source['value'])
                if isinstance(source['value'], str):
                    # Preserve formula-looking strings and apostrophes literally.
                    cell.data_type = 's'
                cell.number_format = source['format'] or 'General'
                cell.font, cell.alignment = font, alignment
                output.append(cell)
            sheet.append(output)
        ref = f'A1:{get_column_letter(len(headers))}{count + 1}'
        excel_table = Table(displayName=name + 'Table', ref=ref,
                            tableColumns=[TableColumn(id=i + 1, name=label) for i, label in enumerate(headers)],
                            autoFilter=AutoFilter(ref=ref))
        # openpyxl warns for all write-only tables; their columns are explicitly
        # supplied above, which is the action requested by that warning.
        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', message='In write-only mode you must add table columns manually')
            sheet.add_table(excel_table)

    try:
        table('Combined', payload['columns'], iter(payload['rows']), len(payload['rows']), payload['rows'][:100])
        sources = [[dict(value=value, format='General') for value in source_values(source)] for source in payload['sources']]
        table('Sources', SOURCE_HEADERS, iter(sources), len(sources), sources[:100])
        book.save(path)
    finally:
        book.close()
