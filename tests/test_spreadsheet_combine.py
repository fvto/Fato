"""Exercise data-loss boundaries of the generic combine workspace."""
from contextlib import ExitStack
from datetime import datetime
from pathlib import Path
import io
import json
import sys
import tempfile
import unittest
from unittest.mock import patch
from werkzeug.datastructures import FileStorage

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import spreadsheet_combine as combine
import sfc_dashboard_server as server


def sheet(rows, name='Data'):
    return dict(name=name, rows=rows, formats=[['General'] * len(row) for row in rows])


class CombineTests(unittest.TestCase):
    def test_filename_uses_selected_source_periods_and_keeps_report_context(self):
        sources = [dict(file='FGWH 2026-10-27_2026-10-27.xls', sheet='FGWH'),
                   dict(file='FGWH 2026-10-02_2026-10-02.xls', sheet='FGWH')]
        self.assertEqual(combine.suggested_filename(sources), 'FGWH_2026-10-02_2026-10-27_Combined.xlsx')
        self.assertEqual(combine.suggested_filename(sources[:1]), 'FGWH_2026-10-27_2026-10-27_Combined.xlsx')
        context = [dict(file='HFPA_VH_Internal_2026-09-01_2026-09-30.xlsx', sheet='HFPA'),
                   dict(file='HFPA_VH_Internal_2026-10-01_2026-10-31.xlsx', sheet='HFPA')]
        self.assertEqual(combine.suggested_filename(context), 'HFPA_VH_Internal_2026-09-01_2026-10-31_Combined.xlsx')

    def test_filename_does_not_invent_period_or_report_type(self):
        for names in [('A.xlsx', 'B.xlsx'),
                      ('FGWH 2026-02-30_2026-03-01.xls', 'FGWH 2026-10-02_2026-10-02.xls'),
                      ('FGWH 2026-10-02_2026-10-02.xls', 'HFPA 2026-10-27_2026-10-27.xls')]:
            self.assertEqual(combine.suggested_filename([dict(file=n, sheet='Data') for n in names]), 'Data_2files_Combined.xlsx')

    def test_output_name_validates_windows_names_and_xlsx_extension(self):
        self.assertEqual(combine.output_filename('Báo cáo tháng 10'), 'Báo cáo tháng 10.xlsx')
        self.assertEqual(combine.output_filename('FGWH.XLSX'), 'FGWH.xlsx')
        for name in ['', '../outside', 'A/B', 'A\\B', 'C:report', 'CON', 'NUL.xlsx', 'COM1.data',
                     'A.xls', 'A.csv', 'A\nB', 'a' * 161, '..', 'A.']:
            with self.subTest(name=name), self.assertRaises(ValueError):
                combine.output_filename(name)

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root_patch = patch.object(combine, 'ROOT', Path(self.temp.name))
        self.root_patch.start()
        self.addCleanup(self.root_patch.stop)

    def import_tables(self, tables):
        files = [FileStorage(io.BytesIO(str(index).encode()), filename=f'{index}.xlsx') for index in range(len(tables))]
        with patch.object(combine, 'read_sheets', side_effect=[[t] for t in tables]):
            job = combine.upload(files)
        choices = [dict(file_id=i, sheet=t['name'], header_row=1) for i, t in enumerate(tables)]
        return job, choices

    def test_aligns_names_not_positions_and_preserves_missing_zero_duplicates(self):
        tables = [sheet([['ID', 'Score'], ['001', 0], ['001', 0]]),
                  sheet([[' score ', 'ID', 'New'], [5, '002', '=literal']])]
        job, choices = self.import_tables(tables)
        with patch.object(combine, 'read_sheets', side_effect=[[t] for t in tables]):
            result = combine.prepare(job['id'], choices)
        self.assertEqual(result['columns'], ['ID', 'Score', 'New'] + combine.SOURCE_COLUMNS)
        rows = [[c['value'] for c in row] for row in result['rows']]
        self.assertEqual(rows[0][:3], ['001', 0, None])
        self.assertEqual(rows[1][:3], ['001', 0, None])
        self.assertEqual(rows[2][:3], ['002', 5, '=literal'])
        self.assertEqual([r[-1] for r in rows], [2, 3, 2])

    def test_fgwh_footer_requires_both_count_and_defective_pairs(self):
        headers = ['Factory Code', 'SO#', 'Audit Date', "Defective Q'ty(pair)", "Total Defect Q'ty", 'Defective Rate']
        rows = [headers, ['JV', '001', datetime(2026, 10, 2), 2, 3, '6.3%'], ['', '', '1', '2', '', '']]
        result = combine.clean_table(sheet(rows), 1)
        self.assertEqual((len(result['records']), result['footer_rows']), (1, 1))
        self.assertEqual(result['records'][0][1][5], dict(value=.063, format='0.0%'))
        rows[-1][3] = '3'
        with self.assertRaisesRegex(ValueError, 'tổng FGWH'):
            combine.clean_table(sheet(rows), 1)
        rows[-1][3], rows[-1][2] = '2', '9'
        with self.assertRaisesRegex(ValueError, 'tổng FGWH'):
            combine.clean_table(sheet(rows), 1)

    def test_generic_numeric_tail_is_kept_and_selected_header_respected(self):
        data = sheet([['Report'], ['ID', 'Count'], ['A', 0], ['', 9]])
        result = combine.clean_table(data, 2)
        self.assertEqual(len(result['records']), 2)
        self.assertEqual(result['footer_rows'], 0)
        self.assertEqual(result['records'][1][0], 4)

    def test_rejects_ambiguous_headers_and_data_beyond_headers(self):
        for rows in [[['A', ' a '], [1, 2]], [['A', ''], [1, 2]], [['A'], [1, 2]]]:
            with self.subTest(rows=rows), self.assertRaises(ValueError):
                combine.clean_table(sheet(rows), 1)

    def test_generic_naming_supports_daily_compact_monthly_and_shared_report_names(self):
        cases = [
            (['Sales_2026-10-01.xlsx', 'Sales_2026-10-12.xlsx'], 'Sales_2026-10-01_2026-10-12_Combined.xlsx'),
            (['Stock_20261002.xlsx', 'Stock_20261008.xlsx'], 'Stock_2026-10-02_2026-10-08_Combined.xlsx'),
            (['COPQ_2026-09.xlsx', 'COPQ_2026-10.xlsx'], 'COPQ_2026-09_2026-10_Combined.xlsx'),
            (['Payroll_Hanoi.xlsx', 'Payroll_HCM.xlsx'], 'Payroll_2files_Combined.xlsx'),
            (['Sales_2026-10.xlsx', 'Sales_2026-10-02.xlsx'], 'Sales_2files_Combined.xlsx'),
            (['Báo cáo kho_2026.10.01.xlsx', 'Báo cáo kho_2026.10.09.xlsx'], 'Báo_cáo_kho_2026-10-01_2026-10-09_Combined.xlsx'),
        ]
        for names, expected in cases:
            with self.subTest(names=names):
                self.assertEqual(combine.suggested_filename([dict(file=n, sheet='Data') for n in names]), expected)

    def test_detects_header_after_report_title_and_respects_end_row(self):
        data = sheet([['Monthly sales'], ['October 2026'], ['Item', 'Amount'],
                      ['001', 100], ['002', 200], ['Total', 300]])
        self.assertEqual(combine.detect_header(data), 3)
        result = combine.clean_table(data, 3, 5)
        self.assertEqual([r[0] for r in result['records']], [4, 5])
        self.assertEqual(result['footer_rows'], 0)
        self.assertEqual(len(combine.clean_table(data, 3)['records']), 3)

    def test_preserves_existing_source_columns_when_recombining(self):
        tables = [sheet([['ID', 'Source File', 'Source File (combine 1)'], ['001', 'original.xlsx', 'previous.xlsx']])]
        job, choices = self.import_tables(tables)
        with patch.object(combine, 'read_sheets', return_value=tables):
            result = combine.prepare(job['id'], choices)
        self.assertEqual(result['source_columns'][0], 'Source File (combine 2)')
        self.assertEqual([c['value'] for c in result['rows'][0]][:3], ['001', 'original.xlsx', 'previous.xlsx'])

    def test_strict_schema_accepts_reordered_columns_but_blocks_missing_columns(self):
        tables = [sheet([['ID', 'Amount'], ['001', 10]]), sheet([['Amount', 'ID'], [20, '002']])]
        job, choices = self.import_tables(tables)
        with patch.object(combine, 'read_sheets', side_effect=[[t] for t in tables]):
            result = combine.prepare(job['id'], choices, dict(schema_mode='strict'))
        self.assertEqual([c['value'] for c in result['rows'][1]][:2], ['002', 20])
        tables[1] = sheet([['ID', 'Extra'], ['002', 20]])
        with patch.object(combine, 'read_sheets', side_effect=[[t] for t in tables]):
            with self.assertRaisesRegex(ValueError, 'bộ cột khác'):
                combine.prepare(job['id'], choices, dict(schema_mode='strict'))
        with patch.object(combine, 'read_sheets', side_effect=[[t] for t in tables]):
            result = combine.prepare(job['id'], choices)
        self.assertEqual(result['sources'][0]['missing_columns'], ['Extra'])
        self.assertEqual(result['sources'][1]['missing_columns'], ['Amount'])

    def test_generic_spacers_whitespace_percentages_and_unknown_footer(self):
        result = combine.clean_table(sheet([
            ['', 'Item\nCode', '', 'Completion %', 'SKU'],
            ['', '001', '', '25.0%', '5%'],
            ['', 'Total', '', 3, '']]), 1)
        self.assertEqual(result['header'], ['Item Code', 'Completion %', 'SKU'])
        self.assertEqual(result['records'][0][1][1], dict(value=.25, format='0.0%'))
        self.assertEqual(result['records'][0][1][2]['value'], '5%')
        self.assertEqual(len(result['records']), 2)
        self.assertEqual(combine.key('ORDER\n QTY'), combine.key('order\u00a0qty'))
        with self.assertRaisesRegex(ValueError, 'chưa có tên'):
            combine.clean_table(sheet([['', 'ID'], ['lost data', '001']]), 1)

    def test_footer_keep_is_explicit_and_options_are_validated(self):
        headers = ['Factory Code', 'SO#', 'Audit Date', "Defective Q'ty(pair)", "Total Defect Q'ty"]
        result = combine.clean_table(sheet([headers, ['JV', '001', datetime(2026, 10, 2), 2, 3], ['', '', '1', '2', '']]), 1, footer_mode='keep')
        self.assertEqual((len(result['records']), result['footer_rows']), (2, 0))
        for options in [[], {'schema_mode': 'guess'}, {'footer_mode': 'guess'}]:
            with self.subTest(options=options), self.assertRaises(ValueError):
                combine.prepare('a' * 32, [], options)

    def test_removes_only_exact_repeated_header_and_blank_rows(self):
        result = combine.clean_table(sheet([['A', 'B'], [0, ''], ['', ''], ['A', 'B'], ['A', 'other']]), 1)
        self.assertEqual((result['blank_rows'], result['repeated_headers'], len(result['records'])), (1, 1, 2))

    def test_duplicate_upload_does_not_leave_partial_job(self):
        files = [FileStorage(io.BytesIO(b'same'), filename=name) for name in ['A.xls', 'B.xls']]
        with patch.object(combine, 'read_sheets', return_value=[sheet([['A'], [1]])]):
            with self.assertRaisesRegex(ValueError, 'trùng nội dung'):
                combine.upload(files)
        self.assertEqual(list(combine.ROOT.iterdir()), [])

    def test_maximum_150_files_and_reject_151_before_saving(self):
        files = [FileStorage(io.BytesIO(str(i).encode()), filename=f'{i}.xlsx') for i in range(151)]
        with self.assertRaisesRegex(ValueError, '1 đến 150'):
            combine.upload(files)
        self.assertEqual(list(combine.ROOT.iterdir()), [])
        with patch.object(combine, 'read_sheets', return_value=[sheet([['ID'], ['001']])]):
            job = combine.upload(files[:150])
        self.assertEqual(len(job['files']), 150)
        response = server.app.test_client().get('/combine')
        self.assertIn('data-max-files="150"', response.text)
        self.assertIn('Tối đa 150 file', response.text)

    def test_changed_source_and_duplicate_sheet_selection_block(self):
        tables = [sheet([['A'], [1]])]
        job, choices = self.import_tables(tables)
        with patch.object(combine, 'read_sheets', return_value=tables):
            with self.assertRaisesRegex(ValueError, 'Trang tính bị trùng'):
                combine.prepare(job['id'], choices * 2)
        (combine.folder_for(job['id']) / 'input_0.xlsx').write_bytes(b'changed')
        with self.assertRaisesRegex(ValueError, 'đã thay đổi'):
            combine.prepare(job['id'], choices)

    def test_http_upload_preview_guards_and_private_headers(self):
        client = server.app.test_client()
        self.assertEqual(client.get('/combine').status_code, 200)
        self.assertEqual(client.get('/combine').headers['Cache-Control'], 'no-store')
        self.assertEqual(client.post('/api/combine/uploads', headers={'Origin': 'https://elsewhere.example'}).status_code, 403)
        self.assertEqual(client.post('/api/combine/uploads').status_code, 400)
        self.assertEqual(client.post('/api/combine/bad/preview', json={}).status_code, 400)
        self.assertEqual(client.get('/api/combine/example').json, dict(example=None))
        with patch.object(combine, 'read_sheets', return_value=[sheet([['ID', 'Count'], ['001', 0]])]):
            response = client.post('/api/combine/uploads', data={'files': (io.BytesIO(b'test'), 'A.xlsx')})
            self.assertEqual(response.status_code, 200)
            job = response.json
            choices = [dict(file_id=0, sheet='Data', header_row=1)]
            result = client.post(f'/api/combine/{job["id"]}/preview', json={'selections': choices})
            self.assertEqual(result.status_code, 200)
            self.assertEqual(result.json['row_count'], 1)
            self.assertEqual(result.json['preview'][0][:2], ['001', 0])

    def test_corrupt_excel_and_non_object_request_are_actionable(self):
        client = server.app.test_client()
        result = client.post('/api/combine/uploads', data={
            'files': (io.BytesIO(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1broken'), 'broken.xls')})
        self.assertEqual(result.status_code, 400)
        self.assertIn('error', result.json)
        self.assertEqual(client.post('/api/combine/bad/preview', json=[]).status_code, 400)

    @unittest.skipUnless((Path.home() / '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node.exe').is_file(), 'Requires bundled Artifact Tool runtime')
    def test_real_export_preserves_literal_strings_dates_rates_and_row_sources(self):
        tables = [sheet([['ID', 'Date', 'Defective Rate', 'Text'],
                         ['001', datetime(2026, 10, 2), '6.3%', '=2+3'],
                         ['002', datetime(2026, 10, 9), '0.0%', '<script>data</script>']])]
        job, choices = self.import_tables(tables)
        with patch.object(combine, 'read_sheets', return_value=tables):
            result = combine.export(job['id'], choices, 'Báo cáo FGWH')
        response = server.app.test_client().get(result['download_url'])
        self.assertEqual(response.status_code, 200)
        self.assertEqual(result['output_name'], 'Báo cáo FGWH.xlsx')
        from urllib.parse import unquote
        self.assertIn('Báo cáo FGWH.xlsx', unquote(response.headers['Content-Disposition']))
        import openpyxl
        book = openpyxl.load_workbook(io.BytesIO(response.data))
        try:
            data = book['Combined']
            self.assertEqual(data['A2'].value, '001')
            self.assertEqual(data['B2'].value, datetime(2026, 10, 2))
            self.assertEqual(data['C2'].value, .063)
            self.assertEqual((data['D2'].value, data['D2'].data_type), ('=2+3', 's'))
            self.assertEqual(data['D3'].value, '<script>data</script>')
            self.assertEqual(data.freeze_panes, 'A2')
            self.assertEqual(book['Sources']['C2'].value, 2)
            # Exercise XLSX ingestion and cached-formula validation on a copy.
            path = combine.folder_for(job['id']) / result['result_id'] / result['output_name']
            self.assertEqual(combine.read_sheets(path)[0]['rows'][1][0], '001')
            # Author a formula test fixture by changing the XML in a disposable
            # ZIP copy, leaving out its cached result deliberately.
            import zipfile
            broken = path.with_name('uncached.xlsx')
            with zipfile.ZipFile(path) as src, zipfile.ZipFile(broken, 'w') as dest:
                for info in src.infolist():
                    content = src.read(info.filename)
                    if info.filename == 'xl/worksheets/sheet1.xml':
                        import xml.etree.ElementTree as ET
                        tree = ET.fromstring(content)
                        ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
                        cell = next(c for c in tree.iter(ns + 'c') if c.get('r') == 'C2')
                        cell.clear()
                        cell.set('r', 'C2')
                        ET.SubElement(cell, ns + 'f').text = '1+1'
                        content = ET.tostring(tree, encoding='utf-8')
                    dest.writestr(info, content)
            with self.assertRaisesRegex(ValueError, 'chưa có kết quả đã lưu'):
                combine.read_sheets(broken)
        finally:
            book.close()
            response.close()

    def test_legacy_combined_download_gets_meaningful_name(self):
        tables = [sheet([['A'], [1]])]
        job, _ = self.import_tables(tables)
        result_id = 'a' * 32
        result_folder = combine.folder_for(job['id']) / result_id
        result_folder.mkdir()
        (result_folder / 'Combined.xlsx').write_bytes(b'legacy workbook')
        combine.json_write(result_folder / 'result.json', dict(
            id=job['id'], result_id=result_id,
            sources=[dict(file='FGWH 2026-10-02_2026-10-27.xls', sheet='FGWH')]))
        response = server.app.test_client().get(f'/api/combine/{job["id"]}/{result_id}/download')
        try:
            self.assertEqual(response.status_code, 200)
            self.assertIn('FGWH_2026-10-02_2026-10-27_Combined.xlsx', response.headers['Content-Disposition'])
            self.assertEqual(response.data, b'legacy workbook')
        finally:
            response.close()


if __name__ == '__main__':
    unittest.main()
