import io
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import openpyxl

import spreadsheet_compare as compare
import spreadsheet_quick_compare as quick
import sfc_dashboard_server as server


class QuickCompareTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.patch = patch.object(compare, 'ROOT', self.root)
        self.patch.start()

    def tearDown(self):
        compare.load_workbook.cache_clear()
        self.patch.stop()
        self.temp.cleanup()

    def book(self, name, sheets, formats=None):
        book = openpyxl.Workbook()
        book.remove(book.active)
        for title, rows in sheets.items():
            sheet = book.create_sheet(title)
            for row in rows:
                sheet.append(row)
            for address, number_format in (formats or {}).get(title, {}).items():
                sheet[address].number_format = number_format
        data = io.BytesIO()
        book.save(data)
        book.close()
        data.seek(0)
        return compare.import_file(data, name)['id']

    def test_only_percentages_and_occupied_rows_are_checked(self):
        a = self.book('A.xlsx', {'Data': [['ID', 'Rate', 'Count'], ['A', '95%', 10], [], ['B', '90%', 20]]})
        b = self.book('B.xlsx', {'Data': [['ID', 'Rate', 'Count'], ['Renamed', '95%', 999], [], ['B', '91%', 30]]})
        result = quick.run({'source_id': a, 'report_id': b})
        summary = result['quick_summary']
        self.assertEqual((summary['left_rows'], summary['right_rows']), (3, 3))
        self.assertTrue(summary['rows_equal'])
        self.assertFalse(summary['percent_equal'])
        self.assertEqual(summary['percent_checked'], 2)
        self.assertEqual(result['counts'], {'match': 2, 'difference': 1})
        issue = next(row for row in result['results'] if row['status'] == 'difference')
        self.assertEqual((issue['cell'], issue['other_cell']), ('B4', 'B4'))
        self.assertEqual((issue['expected'], issue['actual']), ('90%', '91%'))
        self.assertEqual(compare.read_result(result['id'])['quick_summary'], summary)

    def test_formats_percent_text_and_exact_differences(self):
        a = self.book('A.xlsx', {'A': [['Metric', 'Rate'], ['one', .95], ['two', '94.40001%']]}, {'A': {'B2': '0.0%'}})
        b = self.book('B.xlsx', {'Renamed': [['Metric', 'Rate'], ['one', '95%'], ['two', '94.40002%']]})
        result = quick.run({'source_id': a, 'report_id': b})
        self.assertEqual(result['comparison_policy'], 'exact')
        self.assertEqual(result['quick_summary']['percent_issues'], 1)
        issue = next(row for row in result['results'] if row['status'] == 'difference')
        self.assertEqual(issue['expected_display'], '94.40001%')
        self.assertEqual(issue['actual_display'], '94.40002%')
        self.assertEqual(issue['delta_display'], '-0.00001%')
        from spreadsheet_report import prepare
        model = prepare(result, images=False)
        self.assertEqual(model['groups'][0]['rows'][0]['reported'], '94.40002%')
        self.assertEqual(model['groups'][0]['rows'][0]['expected'], '94.40001%')

    def test_missing_rows_values_errors_and_uncached_formulas_are_reported(self):
        a = self.book('A.xlsx', {'Data': [['Rate'], ['95%'], ['0%'], ['20%'], ['10%']]})
        b = self.book('B.xlsx', {'Data': [['Rate'], [None], ['=1/2'], ['#DIV/0!']]}, {'Data': {'A3': '0%'}})
        result = quick.run({'source_id': a, 'report_id': b})
        self.assertFalse(result['quick_summary']['rows_equal'])
        self.assertEqual(result['quick_summary']['row_delta'], 2)
        self.assertEqual(result['results'][0]['delta'], 2)
        self.assertIn('File A nhiều hơn 2 dòng', result['results'][0]['reason'])
        self.assertEqual(result['quick_summary']['percent_issues'], 4)
        self.assertEqual(result['counts'], {'difference': 1, 'missing_value': 2, 'formula_uncached': 1, 'formula_error': 1})

    def test_no_percentages_never_reports_percentage_agreement(self):
        a = self.book('A.xlsx', {'Data': [['ID', 'Rate%'], ['A', 95]]})
        b = self.book('B.xlsx', {'Data': [['ID', 'Rate%'], ['A', 90]]})
        result = quick.run({'source_id': a, 'report_id': b})
        self.assertEqual(result['quick_summary']['percent_checked'], 0)
        self.assertIsNone(result['quick_summary']['percent_equal'])
        self.assertTrue(result['quick_summary']['rows_equal'])
        from spreadsheet_report import render_html
        self.assertIn('Không tìm thấy ô % để kiểm tra', render_html(result))

    def test_per_sheet_counts_and_unmatched_sheets_are_not_hidden_by_equal_totals(self):
        a = self.book('A.xlsx', {'Data': [['95%'], ['90%']], 'Extra': [['80%']]})
        b = self.book('B.xlsx', {'Data': [['95%']], 'Other': [['80%'], ['90%']]})
        result = quick.run({'source_id': a, 'report_id': b})
        summary = result['quick_summary']
        self.assertEqual(summary['left_rows'], summary['right_rows'])
        self.assertFalse(summary['rows_equal'])
        self.assertEqual(summary['row_delta'], 0)
        self.assertEqual(summary['sheets'][0]['row_delta'], 1)
        self.assertIsNone(summary['sheets'][1]['row_delta'])
        self.assertEqual(summary['unmatched_sheets'], 2)
        self.assertEqual(result['counts']['unmatched'], 2)

    def test_saved_results_get_explicit_row_differences_without_rerunning(self):
        a = self.book('A.xlsx', {'Data': [['Rate']] + [['95%']] * 84})
        b = self.book('B.xlsx', {'Data': [['Rate']] + [['95%']] * 13})
        result = quick.run({'source_id': a, 'report_id': b})
        # Simulate a result saved by the previous version, which omitted the delta.
        result['quick_summary'].pop('row_delta')
        for sheet in result['quick_summary']['sheets']:
            sheet.pop('row_delta')
        row = result['results'][0]
        for key in ('delta', 'delta_exact', 'delta_display'):
            row.pop(key, None)
        compare.write_json(self.root/'results'/f'{result["id"]}.json', result)
        client = server.app.test_client()
        with patch.object(quick, 'run', side_effect=AssertionError('Must restore without reanalysis')):
            restored = client.get('/api/compare/results/'+result['id']).get_json()
            self.assertEqual(restored['quick_summary']['row_delta'], 71)
            self.assertEqual(restored['quick_summary']['sheets'][0]['row_delta'], 71)
            self.assertEqual(restored['results'][0]['delta'], 71)
            self.assertEqual(restored['results'][0]['delta_display'], '71 dòng')
            self.assertIn('File A nhiều hơn 71 dòng', restored['results'][0]['reason'])
            downloaded = client.get(restored['json_url']).get_json()
            self.assertEqual(downloaded['quick_summary']['row_delta'], 71)
        self.assertNotIn('row_delta', compare.read_result(result['id'])['quick_summary'])

    def test_api_origin_validation_downloads_and_result_limit(self):
        a = self.book('A.xlsx', {'Data': [['95%']]})
        b = self.book('B.xlsx', {'Data': [['95%']]})
        client = server.app.test_client()
        payload = {'source_id': a, 'report_id': b}
        self.assertEqual(client.post('/api/compare/quick', json=payload, headers={'Origin': 'https://other.example'}).status_code, 403)
        self.assertEqual(client.post('/api/compare/quick', json={'source_id': a, 'report_id': a}).status_code, 400)
        self.assertEqual(client.post('/api/compare/quick', json={}).status_code, 400)
        result = client.post('/api/compare/quick', json=payload).get_json()
        self.assertEqual(client.get(result['csv_url']).status_code, 200)
        self.assertTrue(client.get(result['json_url']).get_json()['quick_summary']['percent_equal'])
        existing = set((self.root/'results').glob('*.json'))
        with patch.object(compare, 'MAX_RESULTS', 1), self.assertRaisesRegex(ValueError, '50.000'):
            quick.run(payload)
        self.assertEqual(set((self.root/'results').glob('*.json')), existing)


if __name__ == '__main__':
    unittest.main()
