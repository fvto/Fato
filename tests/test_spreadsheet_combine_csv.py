"""CSV and mixed-format regressions, including real XLSX export and HTTP limits."""
from pathlib import Path
from datetime import datetime
import codecs
import csv
import io
import json
import sys
import tempfile
import unittest
from unittest.mock import patch

import openpyxl
from werkzeug.datastructures import FileStorage

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import spreadsheet_combine as combine
import sfc_dashboard_server as server


def workbook_bytes(rows):
    book = openpyxl.Workbook()
    for row in rows:
        book.active.append(row)
    stream = io.BytesIO()
    book.save(stream)
    book.close()
    return stream.getvalue()


def csv_bytes(rows, delimiter=',', encoding='utf-8-sig'):
    stream = io.StringIO(newline='')
    csv.writer(stream, delimiter=delimiter).writerows(rows)
    return stream.getvalue().encode(encoding)


def storage(name, content):
    return FileStorage(io.BytesIO(content), filename=name)


def selections(job):
    return [dict(file_id=file['id'], sheet=file['sheets'][0]['name'],
                 header_row=file['sheets'][0]['header_row']) for file in job['files']]


class CsvCombineTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root_patch = patch.object(combine, 'ROOT', Path(self.temp.name) / 'jobs')
        self.root_patch.start()
        self.addCleanup(self.root_patch.stop)
        self.client = server.app.test_client()

    def read(self, content, config=None):
        path = Path(self.temp.name) / 'sample.csv'
        path.write_bytes(content)
        return combine.read_sheets(path, config)[0]

    def test_unicode_boms_and_encodings_roundtrip_without_replacement(self):
        rows = [['Mã', 'Tên'], ['001', 'Nguyễn Thị Hà']]
        for encoding in ['utf-8', 'utf-8-sig', 'utf-16', 'utf-32']:
            with self.subTest(encoding=encoding):
                self.assertEqual(self.read(csv_bytes(rows, encoding=encoding))['rows'], rows)
        for bom, encoding in [(codecs.BOM_UTF16_BE, 'utf-16-be'), (codecs.BOM_UTF32_BE, 'utf-32-be')]:
            with self.subTest(encoding=encoding):
                self.assertEqual(self.read(bom + 'Mã,Tên\n001,Nguyễn Thị Hà'.encode(encoding))['rows'], rows)
        samples = [('cp1258', [['Tên', 'Ma\u0303'], ['Nguyê\u0303n', '001']]),
                   ('cp1252', [['Name', 'City'], ['André €', 'Paris']]),
                   ('cp1250', [['Name', 'City'], ['Łódź', 'Warszawa']]),
                   ('iso8859-1', [['Name', 'City'], ['André', 'Paris']]),
                   ('shift_jis', [['氏名', '番号'], ['田中', '001']]),
                   ('gb18030', [['姓名', '编号'], ['张三', '001']]),
                   ('big5', [['姓名', '編號'], ['張三', '001']])]
        for encoding, sample in samples:
            with self.subTest(encoding=encoding):
                self.assertEqual(self.read(csv_bytes(sample, encoding=encoding), {'encoding': encoding})['rows'], sample)
        # Legacy detection is visibly qualified; a short byte string can have
        # multiple valid decodings, so explicit overrides remain available.
        detected = self.read(csv_bytes(samples[0][1], encoding='cp1258'))
        self.assertEqual(detected['rows'], samples[0][1])
        self.assertTrue(detected['warnings'])

    def test_delimiters_quoted_fields_multiline_and_repeated_header(self):
        for delimiter in [',', ';', '\t', '|']:
            rows = [['ID', 'Note', 'Rate %'], ['001', f'a{delimiter}b "quoted"\nnext line', '25.0%'],
                    [], ['ID', 'Note', 'Rate %'], ['002', '=2+3', '0.0%']]
            with self.subTest(delimiter=repr(delimiter)):
                source = self.read(csv_bytes(rows, delimiter))
                self.assertEqual(source['csv']['delimiter'], delimiter)
                self.assertEqual(source['rows'], rows)
                table = combine.clean_table(source, 1)
                self.assertEqual((table['blank_rows'], table['repeated_headers']), (1, 1))
                self.assertEqual([row[0] for row in table['records']], [2, 5])
                self.assertEqual(table['records'][0][1][0]['value'], '001')
                self.assertEqual(table['records'][0][1][2]['value'], .25)

    def test_preamble_sep_directive_single_column_and_explicit_delimiter(self):
        source = self.read(b'Report October\nID;Amount\n001;1,25\n002;2,50\n')
        self.assertEqual(source['csv']['delimiter'], ';')
        self.assertEqual(combine.detect_header(source), 2)
        self.assertEqual(combine.clean_table(source, 2, 3)['records'][0][1][1]['value'], '1,25')
        source = self.read(codecs.BOM_UTF8 + b'sep=;\r\nID;Name\r\n001;A\r\n')
        self.assertEqual((source['csv']['delimiter'], combine.detect_header(source)), (';', 2))
        self.assertEqual(combine.clean_table(source, 2)['records'][0][0], 3)
        source = self.read(b'ID\n001\n002\n')
        self.assertEqual(source['rows'], [['ID'], ['001'], ['002']])
        self.assertEqual(self.read(b'ID;Note\n001;a,b\n', {'delimiter': ';'})['rows'][1], ['001', 'a,b'])

    def test_csv_structure_errors_name_each_file_and_leave_no_partial_job(self):
        result = self.client.post('/api/combine/uploads', data={'files': [
            (io.BytesIO(b'ID,Name\n001,"unterminated'), 'broken.csv'),
            (io.BytesIO(b''), 'empty.csv'),
            (io.BytesIO(b'not excel'), 'corrupt.xlsx'),
            (io.BytesIO(b'data'), 'bad.txt'),
        ]})
        self.assertEqual(result.status_code, 400)
        self.assertEqual([error['file'] for error in result.json['errors']],
                         ['broken.csv', 'empty.csv', 'corrupt.xlsx', 'bad.txt'])
        self.assertEqual(list(combine.ROOT.iterdir()), [])

    def test_wrong_encoding_controls_overlong_fields_and_cell_budget(self):
        for content, config in [(b'ID,Name\n001,\xff', {'encoding':'utf-8-sig'}),
                                (b'ID\n\x00\n', None), (b'ID\n\x01\n', None),
                                (b'ID\n' + b'x' * 32768, None)]:
            with self.subTest(content=content[:20]), self.assertRaises(ValueError):
                self.read(content, config)
        with patch.object(combine, 'MAX_CELLS', 4):
            for content in [b'A,B\n1,2\n3,4', b'A\n\n\n\n\n']:
                with self.subTest(content=content), self.assertRaisesRegex(ValueError, '500.000'):
                    self.read(content)
        for config in [[], {'encoding':'unknown'}, {'delimiter':'xx'}, {'encoding':[]}, {'extra':True}]:
            with self.subTest(config=config), self.assertRaises(ValueError):
                self.read(b'ID\n001', config)
        result = self.client.post('/api/combine/uploads', data={
            'files': (io.BytesIO(b'ID\n001'), 'ok.csv'), 'csv_options': 'bad-json'})
        self.assertEqual(result.status_code, 400)

    def test_csv_uses_existing_header_schema_footer_and_duplicate_rules(self):
        job = combine.upload([storage('a.csv', b'ID,Count\n001,1\n'),
                              storage('b.xlsx', workbook_bytes([['Count', 'ID', 'Other'], [2, '002', 'yes']]))])
        choices = selections(job)
        with self.assertRaisesRegex(ValueError, 'b.xlsx.*bộ cột khác'):
            combine.prepare(job['id'], choices, {'schema_mode':'strict'})
        result = combine.prepare(job['id'], choices)
        self.assertEqual(result['columns'][:3], ['ID', 'Count', 'Other'])
        self.assertEqual([c['value'] for c in result['rows'][0]][:3], ['001', '1', None])
        self.assertEqual([c['value'] for c in result['rows'][1]][:3], ['002', 2, 'yes'])
        with self.assertRaisesRegex(ValueError, 'trùng nội dung'):
            combine.upload([storage('a.csv', b'ID\n001'), storage('copy.CSV', b'ID\n001')])
        for content, message in [(b'ID, id \n001,002', 'trùng'), (b'ID,\n001,2', 'vượt số cột'),
                                 (b'ID\n001,2', 'vượt số cột')]:
            uploaded = combine.upload([storage('invalid-header.csv', content)])
            with self.assertRaisesRegex(ValueError, f'invalid-header.csv.*{message}'):
                combine.prepare(uploaded['id'], [dict(file_id=0, sheet='CSV', header_row=1)])

    def test_expanded_union_uses_streaming_above_memory_threshold(self):
        with patch.object(combine, 'MAX_CELLS', 30):
            job = combine.upload([storage(f'{i}.csv', csv_bytes([[f'A{i}', f'B{i}'], ['a', 'b'], ['c', 'd']])) for i in range(3)])
            result=combine.prepare(job['id'], selections(job))
            self.assertTrue(result['large'])
            self.assertEqual((result['row_count'],result['column_count']),(6,9))
            self.assertEqual(len(list(result['rows'])),6)

    def test_large_150_file_export_streams_with_formats_and_source_tables(self):
        files = [storage('Excel.xlsx', workbook_bytes([
            ['ID', 'Text', 'Date', 'Rate %'], ['0001', "'prefix", datetime(2026, 10, 2), .25]]))]
        files += [storage(f'{i}.csv', csv_bytes([
            ['ID', 'Text', 'Date', 'Rate %'],
            *[[f'{i:03}-{row:03}', '=2+3', '2026-10-02', '25.0%'] for row in range(100)]]))
            for i in range(1, 150)]
        job = combine.upload(files)
        with patch.object(combine, 'runtime', side_effect=AssertionError('Large exports must stream')):
            result = combine.export(job['id'], selections(job), 'Large mixed')
        self.assertEqual((result['file_count'], result['row_count']), (150, 14901))
        path = combine.folder_for(job['id']) / result['result_id'] / result['output_name']
        book = openpyxl.load_workbook(path)
        try:
            data, sources = book['Combined'], book['Sources']
            self.assertEqual(data.max_row, 14902)
            self.assertEqual(data['A2'].value, '0001')
            self.assertEqual(data['B2'].value, "'prefix")
            self.assertEqual(data['C2'].value, datetime(2026, 10, 2))
            self.assertEqual((data['B3'].value, data['B3'].data_type), ('=2+3', 's'))
            self.assertEqual((data['D3'].value, data['D3'].number_format), (.25, '0.0%'))
            self.assertEqual(data['A2'].font.name, 'Arial')
            self.assertEqual(data['A2'].alignment.vertical, 'center')
            self.assertEqual(data.freeze_panes, 'A2')
            self.assertEqual(data.sheet_format.defaultRowHeight, 20)
            self.assertEqual(data.row_dimensions[1].height, 46)
            self.assertEqual(data.tables['CombinedTable'].ref, 'A1:G14902')
            self.assertEqual(data.tables['CombinedTable'].autoFilter.ref, 'A1:G14902')
            self.assertEqual(sources.tables['SourcesTable'].ref, 'A1:L151')
            self.assertEqual(sources['A151'].value, '149.csv')
            self.assertEqual(data['A14902'].value, '149-099')
        finally:
            book.close()

    def test_150_real_mixed_files_http_preview_and_151_rejection(self):
        files = [('new.xlsx', workbook_bytes([['ID', 'Count'], ['001', 1]])),
                 ('macro.xlsm', workbook_bytes([['ID', 'Count'], ['002', 2]])),
                 ('legacy.xls', (Path(__file__).parent / 'fixtures/combine_legacy.xls').read_bytes())]
        files += [(f'{i}.CSV', csv_bytes([['Count', 'ID'], [str(i), f'{i:03}']],
                                       delimiter=[';', ',', '\t', '|'][i % 4],
                                       encoding=['utf-8', 'utf-8-sig', 'utf-16'][i % 3])) for i in range(4, 151)]
        self.assertEqual(len(files), 150)
        response = self.client.post('/api/combine/uploads', data={'files': [(io.BytesIO(b), n) for n, b in files]})
        self.assertEqual(response.status_code, 200, response.json)
        job = response.json
        result = self.client.post(f'/api/combine/{job["id"]}/preview', json={
            'selections': selections(job), 'options': {'schema_mode':'strict'}})
        self.assertEqual(result.status_code, 200, result.json)
        self.assertEqual((result.json['file_count'], result.json['row_count']), (150, 150))
        self.assertEqual(result.json['preview'][2][:2], ['003', 3])
        self.assertEqual(result.json['preview'][3][:2], ['004', '4'])
        response = self.client.post('/api/combine/uploads', data={
            'files': [(io.BytesIO(f'ID\n{i}'.encode()), f'{i}.csv') for i in range(151)]})
        self.assertEqual(response.status_code, 400)
        self.assertIn('150', response.json['error'])
        self.assertEqual(len(list(combine.ROOT.iterdir())), 1)
        self.assertIn('.csv', self.client.get('/combine').text)
        self.mixed_job = job

    @unittest.skipUnless((Path.home() / '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node.exe').is_file(), 'Requires bundled Artifact Tool runtime')
    def test_real_mixed_export_150_files_keeps_literals_formats_and_sources(self):
        self.test_150_real_mixed_files_http_preview_and_151_rejection()
        job = self.mixed_job
        result = combine.export(job['id'], selections(job))
        response = self.client.get(result['download_url'])
        self.addCleanup(response.close)
        self.assertEqual(response.status_code, 200)
        book = openpyxl.load_workbook(io.BytesIO(response.data), read_only=True)
        try:
            rows = list(book['Combined'].iter_rows(values_only=True))
            self.assertEqual(len(rows), 151)
            self.assertEqual(rows[1][:2], ('001', 1))
            self.assertEqual(rows[-1][:2], ('150', '150'))
            self.assertEqual(rows[-1][-3:], ('150.CSV', 'CSV', 2))
            self.assertEqual(len(list(book['Sources'].iter_rows())), 151)
        finally:
            book.close()
        literal_job = combine.upload([storage('text.csv', csv_bytes([
            ['ID', 'Name', 'Completion %'], ['0001', 'Nguyễn, Hà\n"quoted"', '25.0%'],
            ['0002', '=2+3', '0%'], ['0003', "'prefix", '100%']]))])
        literal_result = combine.export(literal_job['id'], selections(literal_job))
        output = combine.folder_for(literal_job['id']) / literal_result['result_id'] / literal_result['output_name']
        book = openpyxl.load_workbook(output)
        try:
            self.assertEqual(book['Combined']['A2'].value, '0001')
            self.assertEqual(book['Combined']['B2'].value, 'Nguyễn, Hà\n"quoted"')
            self.assertEqual((book['Combined']['B3'].value, book['Combined']['B3'].data_type), ('=2+3', 's'))
            self.assertEqual(book['Combined']['B4'].value, "'prefix")
            self.assertEqual(book['Combined']['C2'].value, .25)
        finally:
            book.close()


if __name__ == '__main__':
    unittest.main()
