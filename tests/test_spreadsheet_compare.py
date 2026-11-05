import io
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import spreadsheet_compare as compare
import spreadsheet_compare_html as html_report
import sfc_dashboard_server as server


class ComparisonTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.patch = patch.object(compare, 'ROOT', self.root)
        self.patch.start()

    def tearDown(self):
        compare.load_workbook.cache_clear()
        self.patch.stop()
        self.temp.cleanup()

    def import_rows(self, rows, name='Source.xlsx', sheet='Data'):
        book = openpyxl.Workbook()
        book.active.title = sheet
        for row in rows:
            book.active.append(row)
        stream = io.BytesIO()
        book.save(stream)
        book.close()
        stream.seek(0)
        return compare.import_file(stream, name)

    def rule(self, book, **overrides):
        return dict(kind='calculation', name='Independent calculation', left={'workbook': book['id'], 'sheet': 'Data'},
                    source_range='B2:C2', target_range='D2', operation='average', direction='rows',
                    number_mode='number', tolerance='0', **overrides)

    def test_average_does_not_drop_zero_or_convert_na_to_zero(self):
        book = self.import_rows([['Line', 'A', 'B', 'Average'], ['Line1', 0, 100, 50], ['Line2', 'N/A', 94.4, 94.4], ['Line3', 'N/A', None, '#DIV/0!']])
        rule = self.rule(book)
        rule.update(source_range='B2:C4', target_range='D2:D4', label_range='A2:A4')
        result = compare.run([rule])
        self.assertEqual([r['status'] for r in result['results']], ['match', 'match', 'formula_error'])
        self.assertEqual(result['results'][0]['valid_count'], 2)
        self.assertEqual(result['results'][1]['ignored_count'], 1)

    def test_percent_comparison_never_accepts_tolerance_or_rounding_differences(self):
        self.assertEqual(compare.numeric({'value':'94.4%','format':'General'})[1], compare.Decimal('.944'))
        self.assertEqual(compare.numeric({'value':'30','format':'General'}, 'percent')[0], 'invalid')
        self.assertEqual(compare.numeric({'value':'30','format':'General'}, 'percentage_points')[1], compare.Decimal('.3'))
        self.assertEqual(compare.numeric({'value':0.3,'format':'0.0%'}, 'percentage_points')[1], compare.Decimal('.3'))
        book = self.import_rows([['Line','A','B','Average'],['X','95.2%','95.4%','95.0%']])
        rule = self.rule(book)
        rule.update(number_mode='percent', tolerance='0', rounding_digits=0)
        self.assertEqual(compare.run([rule])['results'][0]['status'], 'difference')
        rule['rounding_digits'] = 2
        self.assertEqual(compare.run([rule])['results'][0]['status'], 'difference')
        rule['tolerance'] = '0.4'
        with self.assertRaisesRegex(ValueError,'Sai số phải bằng 0'):
            compare.run([rule])
        status,delta,_=compare.compare_numbers(compare.Decimal('1'),compare.Decimal('1.00000000000000001'),{},'number',compare.Decimal(0),None)
        self.assertEqual(status,'difference')
        self.assertNotEqual(delta,0)

    def test_formula_without_cache_is_not_claimed_to_be_an_error_or_zero(self):
        book = self.import_rows([['Line','A','B','Average'],['X',10,20,'=AVERAGE(B2:C2)']])
        result = compare.run([self.rule(book)])['results'][0]
        self.assertEqual(result['status'], 'formula_uncached')
        self.assertEqual(result['expected'], 15)
        self.assertEqual(result['formula'], '=AVERAGE(B2:C2)')

    def test_rounds_calculated_result_before_comparison_and_preserves_original(self):
        book=self.import_rows([['Line','A','B','Average'],['X','95.2%','95.3%','95.3%']])
        rule=self.rule(book)
        rule.update(number_mode='percent',comparison_rounding_digits=1)
        result=compare.run([rule])
        row=result['results'][0]
        self.assertEqual(row['status'],'match')
        self.assertEqual(row['expected_exact'],'0.953')
        self.assertEqual(compare.Decimal(row['expected_unrounded']),compare.Decimal('.9525'))
        self.assertEqual(result['comparison_policy'],'rounded')
        self.assertEqual(result['comparison_rounding_digits'],[1])
        rule['comparison_rounding_digits']=3
        row=compare.run([rule])['results'][0]
        self.assertEqual(row['status'],'difference')
        self.assertEqual(compare.Decimal(row['expected_exact']),compare.Decimal('.9525'))

    def test_percentage_rounding_uses_half_up_in_percentage_units(self):
        self.assertEqual(compare.round_for_comparison(compare.Decimal('.97625'),'percent',1),compare.Decimal('.976'))
        self.assertEqual(compare.round_for_comparison(compare.Decimal('.9765'),'percent',1),compare.Decimal('.977'))
        self.assertEqual(compare.round_for_comparison(compare.Decimal('-.9765'),'percent',1),compare.Decimal('-.977'))

    def test_weighted_average_and_ratio_are_different_from_unweighted_average(self):
        book = self.import_rows([['Line','A','B','Result','W1','W2'],['X',80,100,95,1,3]])
        rule = self.rule(book)
        rule.update(operation='weighted_average', weights_range='E2:F2')
        self.assertEqual(compare.run([rule])['results'][0]['status'], 'match')
        rule['operation'] = 'ratio_of_sums'
        self.assertEqual(compare.run([rule])['results'][0]['expected'], 45)
        rule['operation'] = 'average'
        self.assertEqual(compare.run([rule])['results'][0]['expected'], 90)

    def test_column_checks_and_arbitrary_ranges_do_not_depend_on_a_template(self):
        book = self.import_rows([['Name','Input X','Input Y'],['A',2,10],['B',4,20],['Totals',6,30]])
        rule = self.rule(book)
        rule.update(source_range='B2:C3', target_range='B4:C4', direction='columns', operation='sum', label_range='B1:C1')
        result = compare.run([rule])
        self.assertEqual(result['counts'], {'match': 2})
        self.assertEqual([r['cell'] for r in result['results']], ['B4','C4'])

    def test_ranges_are_validated_and_result_cells_cannot_be_inputs(self):
        book = self.import_rows([['Line','A','B','Average'],['X',1,3,2]])
        for changes in [{'source_range':'D2:D2'}, {'source_range':'A1:A1048576'}, {'target_range':'D2:D9'}, {'source_range':'__import__(1)'}, {'tolerance':'NaN'}]:
            rule = self.rule(book)
            rule.update(changes)
            with self.assertRaises(ValueError):
                compare.run([rule])

    def test_text_input_is_reviewed_instead_of_silently_removed(self):
        book = self.import_rows([['Line','A','B','Average'],['X','unexpected',3,3]])
        result = compare.run([self.rule(book)])['results'][0]
        self.assertEqual(result['status'], 'invalid_input')
        self.assertIn('B2', result['reason'])

    def comparison(self, left, right, **overrides):
        return dict(kind='comparison', name='Any column comparison', left={'workbook':left['id'],'sheet':'Data'},
            right={'workbook':right['id'],'sheet':'Data'},left_range='A2:C4',right_range='A2:C4',
            left_keys='A,B',right_keys='A,B',left_columns='C',right_columns='C',matching='keys',number_mode='number',tolerance=0,**overrides)

    def test_composite_keys_match_reordered_rows_and_report_missing_rows(self):
        left = self.import_rows([['Factory','Line','Score'],['VH','L1',10],['VH2','L1',20],['VH','L2',30]])
        right = self.import_rows([['Factory','Line','Score'],['vh2','L1',20],['VH','L1',11],['JV','L3',40]],name='Comparison.xlsx')
        result = compare.run([self.comparison(left,right)])
        self.assertEqual(result['counts'], {'difference':1,'match':1,'unmatched':2})
        mismatch = next(row for row in result['results'] if row['status']=='difference')
        self.assertEqual((mismatch['cell'],mismatch['other_cell']),('C2','C3'))

    def test_duplicate_keys_are_ambiguous_not_first_match(self):
        left = self.import_rows([['Factory','Line','Score'],['VH','L1',10],['VH','L1',20]])
        right = self.import_rows([['Factory','Line','Score'],['VH','L1',10]])
        result = compare.run([self.comparison(left,right)])
        self.assertEqual(result['counts'], {'ambiguous_key':1})
        self.assertEqual(result['results'][0]['left_rows'], [2,3])

    def test_auto_text_and_numeric_differences_and_unmatched_blank_keys(self):
        left = self.import_rows([['Factory','Line','Score'],['VH','L1','Pass'],['VH','L2',10],[None,'L3',5]])
        right = self.import_rows([['Factory','Line','Score'],['VH','L1','pass'],['VH','L2','Ten'],[None,'L3',5]])
        result = compare.run([self.comparison(left,right)])
        self.assertEqual(result['counts'], {'match':1,'invalid_value':1,'invalid_key':1})

    def test_results_persist_and_csv_contains_all_rows_and_safe_formula_text(self):
        book = self.import_rows([['Line','A','B','Average'],['@untrusted',1,3,0]])
        rule = self.rule(book)
        rule['label_range'] = 'A2'
        result = compare.run([rule])
        self.assertEqual(compare.read_result(result['id'])['counts'], {'difference':1})
        csv = compare.result_csv(result)
        self.assertTrue(csv.startswith('\ufeff'))
        self.assertIn("'@untrusted", csv)
        self.assertIn('D2', csv)

    def test_import_and_api_errors_are_local_and_downloads_work(self):
        with server.app.test_client() as client:
            self.assertEqual(client.get('/compare').status_code,200)
            rejected = client.post('/api/compare/run',json={'rules':[]},headers={'Origin':'https://other.example'})
            self.assertEqual(rejected.status_code,403)
            invalid = client.post('/api/compare/workbooks',data={'files':(io.BytesIO(b'not excel'),'wrong.xlsx')})
            self.assertEqual(invalid.status_code,400)
            self.assertEqual(client.get('/api/compare/workbooks/not-an-id').status_code,400)
            book = self.import_rows([['Line','A','B','Average'],['X',1,3,2]])
            response = client.post('/api/compare/run',json={'rules':[self.rule(book)]})
            self.assertEqual(response.status_code,200,response.get_json())
            result = response.get_json()
            self.assertEqual(result['total'],1)
            self.assertEqual(result['filtered_total'],0)
            all_rows=client.get('/api/compare/results/'+result['id']+'?status=all').get_json()
            self.assertEqual(len(all_rows['results']),1)
            self.assertEqual(client.get(result['csv_url']).status_code,200)
            self.assertEqual(client.get(result['json_url']).get_json()['total'],1)

    def test_upload_is_bounded_and_does_not_leave_an_incomplete_input(self):
        with patch.object(compare, 'MAX_BYTES', 10):
            with self.assertRaisesRegex(ValueError,'40 MB'):
                compare.import_file(io.BytesIO(b'X'*20),'oversize.xlsx')
        self.assertEqual(list((self.root/'inputs').iterdir()), [])

    def test_filtered_results_are_not_hidden_behind_pages_of_matches(self):
        book = self.import_rows([['Line','A','B','Average']] + [[f'L{i}',1,3,2 if i<250 else 9] for i in range(251)])
        rule = self.rule(book)
        rule.update(source_range='B2:C252',target_range='D2:D252')
        with server.app.test_client() as client:
            result=client.post('/api/compare/run',json={'rules':[rule]}).get_json()
            self.assertEqual(result['total'],251)
            self.assertEqual(result['filtered_total'],1)
            self.assertEqual(result['results'][0]['cell'],'D252')
            page=client.get('/api/compare/results/'+result['id']+'?status=all&offset=200').get_json()
            self.assertEqual(len(page['results']),51)

    def test_saved_exports_can_be_imported_without_an_arbitrary_path(self):
        book=self.import_rows([['A','B'],[1,2]])
        export_root=self.root/'exports'
        folder=export_root/'batch'
        folder.mkdir(parents=True)
        (folder/'Saved.xlsx').write_bytes((self.root/'inputs'/book['id']/'source.xlsx').read_bytes())
        with patch.object(compare,'EXPORTS',export_root), server.app.test_client() as client:
            catalog=client.get('/api/compare/catalog').get_json()['files']
            self.assertEqual(len(catalog),1)
            self.assertNotIn('relative',catalog[0])
            imported=client.post('/api/compare/workbooks',json={'catalog_id':catalog[0]['id']})
            self.assertEqual(imported.status_code,200)
            self.assertEqual(imported.get_json()['workbooks'][0]['name'],'Saved.xlsx')
            self.assertEqual(client.post('/api/compare/workbooks',json={'catalog_id':'../../not-a-file'}).status_code,400)

    def test_delete_export_removes_only_selected_file_and_preserves_imported_copy(self):
        book = self.import_rows([['A', 'B'], [1, 2]])
        source = self.root/'inputs'/book['id']/'source.xlsx'
        export_root = self.root/'exports'
        for batch in ('first', 'second'):
            folder = export_root/batch
            folder.mkdir(parents=True)
            (folder/'Saved.xlsx').write_bytes(source.read_bytes())
        with patch.object(compare, 'EXPORTS', export_root), server.app.test_client() as client:
            selected = next(item for item in compare.catalog() if item['relative'] == 'first/Saved.xlsx')
            url = '/api/compare/catalog/'+selected['id']+'/delete'
            denied = client.post(url, json={}, headers={'Origin': 'https://other.example'})
            self.assertEqual(denied.status_code, 403)
            self.assertTrue((export_root/'first/Saved.xlsx').exists())
            response = client.post(url, json={})
            self.assertEqual(response.status_code, 200, response.get_json())
            self.assertEqual(response.get_json()['deleted']['id'], selected['id'])
            self.assertFalse((export_root/'first/Saved.xlsx').exists())
            self.assertTrue((export_root/'second/Saved.xlsx').exists())
            self.assertTrue(source.exists())
            with client.get('/api/compare/workbooks/'+book['id']+'/download') as download:
                self.assertEqual(download.status_code, 200)
            self.assertEqual(len(client.get('/api/compare/catalog').get_json()['files']), 1)
            self.assertEqual(client.post(url, json={}).status_code, 400)
            self.assertEqual(client.post('/api/compare/catalog/unknown/delete', json={}).status_code, 400)

    def test_delete_export_rejects_paths_outside_exports(self):
        outside = self.root/'Do not delete.xlsx'
        outside.write_bytes(b'keep')
        selection = dict(id='a'*32, name=outside.name, relative='../Do not delete.xlsx')
        with patch.object(compare, 'EXPORTS', self.root/'exports'), patch.object(compare, 'catalog', return_value=[selection]):
            with self.assertRaisesRegex(ValueError, 'outside'):
                compare.delete_export(selection['id'])
        self.assertEqual(outside.read_bytes(), b'keep')

    def test_delete_export_reports_locked_file_without_removing_catalog_entry(self):
        export_root = self.root/'exports'
        folder = export_root/'batch'
        folder.mkdir(parents=True)
        saved = folder/'Saved.xlsx'
        saved.write_bytes(b'keep')
        with patch.object(compare, 'EXPORTS', export_root), server.app.test_client() as client:
            selected = compare.catalog()[0]
            with patch.object(Path, 'unlink', side_effect=PermissionError('locked')):
                response = client.post('/api/compare/catalog/'+selected['id']+'/delete', json={})
            self.assertEqual(response.status_code, 400)
            self.assertIn('Excel / WPS', response.get_json()['error'])
            self.assertTrue(saved.exists())
            self.assertEqual(len(compare.catalog()), 1)

    def test_bulk_delete_covers_more_than_80_exports_and_only_confirmed_files(self):
        export_root = self.root/'exports'
        folder = export_root/'batch'
        folder.mkdir(parents=True)
        for index in range(85):
            (folder/f'Saved{index}.xlsx').write_bytes(b'export')
        for name in ('metadata.json', 'unfinished.partial.xlsx'):
            (folder/name).write_bytes(b'keep')
        with patch.object(compare, 'EXPORTS', export_root), server.app.test_client() as client:
            catalog = client.get('/api/compare/catalog').get_json()
            self.assertEqual(len(catalog['files']), 80)
            self.assertEqual(catalog['total'], 85)
            self.assertEqual(len(catalog['delete_ids']), 85)
            # New exports arriving after confirmation are outside the snapshot.
            (folder/'New.xlsx').write_bytes(b'new')
            url = '/api/compare/catalog/delete-all'
            payload = {'catalog_ids': catalog['delete_ids']}
            self.assertEqual(client.post(url, json=payload, headers={'Origin': 'https://other.example'}).status_code, 403)
            self.assertEqual(client.post(url, json={'catalog_ids': ['../../outside.xlsx']}).status_code, 400)
            self.assertEqual(client.post(url, json=[]).status_code, 400)
            response = client.post(url, json=payload)
            self.assertEqual(response.status_code, 200, response.get_json())
            self.assertEqual(response.get_json()['deleted_count'], 85)
            self.assertEqual(response.get_json()['failed'], [])
            self.assertTrue((folder/'New.xlsx').exists())
            self.assertTrue((folder/'metadata.json').exists())
            self.assertTrue((folder/'unfinished.partial.xlsx').exists())
            self.assertEqual(client.get('/api/compare/catalog').get_json()['total'], 1)
            repeated = client.post(url, json=payload).get_json()
            self.assertEqual(repeated['deleted_count'], 0)
            self.assertEqual(repeated['missing_count'], 85)

    def test_bulk_delete_reports_partial_failure_and_can_retry(self):
        export_root = self.root/'exports'
        folder = export_root/'batch'
        folder.mkdir(parents=True)
        for name in ('Locked.xlsx', 'Saved.xls'):
            (folder/name).write_bytes(b'export')
        unlink = Path.unlink
        def reject_locked(path, *args, **kwargs):
            if path.name == 'Locked.xlsx':
                raise PermissionError('locked')
            return unlink(path, *args, **kwargs)
        with patch.object(compare, 'EXPORTS', export_root), server.app.test_client() as client:
            ids = client.get('/api/compare/catalog').get_json()['delete_ids']
            with patch.object(Path, 'unlink', reject_locked):
                result = client.post('/api/compare/catalog/delete-all', json={'catalog_ids': ids}).get_json()
            self.assertEqual(result['deleted_count'], 1)
            self.assertEqual([item['name'] for item in result['failed']], ['Locked.xlsx'])
            self.assertTrue((folder/'Locked.xlsx').exists())
            self.assertFalse((folder/'Saved.xls').exists())
            remaining = client.get('/api/compare/catalog').get_json()['delete_ids']
            retry = client.post('/api/compare/catalog/delete-all', json={'catalog_ids': remaining}).get_json()
            self.assertEqual(retry['deleted_count'], 1)
            self.assertEqual(retry['failed'], [])

    def test_bulk_delete_validates_every_target_before_deleting(self):
        export_root = self.root/'exports'
        folder = export_root/'batch'
        folder.mkdir(parents=True)
        inside = folder/'Saved.xlsx'
        inside.write_bytes(b'keep')
        outside = self.root/'Outside.xlsx'
        outside.write_bytes(b'keep')
        items = [dict(id='a'*32, relative='batch/Saved.xlsx'), dict(id='b'*32, relative='../Outside.xlsx')]
        with patch.object(compare, 'EXPORTS', export_root), patch.object(compare, 'catalog', return_value=items):
            with self.assertRaisesRegex(ValueError, 'outside'):
                compare.delete_exports([item['id'] for item in items])
        self.assertTrue(inside.exists())
        self.assertTrue(outside.exists())

    def test_html_report_contains_all_issues_embedded_images_and_escaped_data(self):
        left = self.import_rows([['Key', 'Value'], ['<script>alert(1)</script>', 10], ['Only A', 20]], name='A<img>.xlsx')
        right = self.import_rows([['Key', 'Value'], ['<script>alert(1)</script>', 11], ['Only B', 30]], name='B.xlsx')
        rule = self.comparison(left, right)
        rule.update(left_range='A2:B3', right_range='A2:B3', left_keys='A', right_keys='A', left_columns='B', right_columns='B')
        result = compare.run([rule])
        tiles, links = html_report.capture_plan(result, result['results'])
        self.assertEqual(len(tiles), 2)
        self.assertTrue(all(links))
        self.assertIn((2, 2), tiles[0]['highlights'])
        self.assertIn((3, 1), tiles[0]['highlights'])
        self.assertIn('class="highlight"', html_report.worksheet_html(tiles[0])[0])
        with patch('spreadsheet_report.capture_images', return_value={tile['id']: ('UE5H', 'A1:B3') for tile in tiles}), server.app.test_client() as client:
            response = client.get('/api/compare/results/'+result['id']+'/download.html?status=match&offset=200')
        self.assertEqual(response.status_code, 200)
        self.assertIn('attachment;', response.headers['Content-Disposition'])
        document = response.get_data(as_text=True)
        self.assertEqual(document.count('<tr id="finding-'), 3)
        self.assertIn('<details class="finding-group">',document)
        self.assertEqual(document.count('src="data:image/png;base64,'), 2)
        self.assertIn('&lt;script&gt;', document)
        self.assertIn('A&lt;img&gt;.xlsx', document)
        self.assertNotIn('<script', document)
        self.assertNotIn('src="http', document)
        self.assertIn('SHA-256:', document)

    def test_html_report_refuses_to_silently_omit_screenshots(self):
        book = self.import_rows([['Line', 'A', 'B', 'Average'], ['X', 1, 3, 9]])
        result = compare.run([self.rule(book)])
        with patch.object(html_report, 'MAX_CAPTURES', 0):
            with self.assertRaisesRegex(ValueError, 'chia nhỏ'):
                html_report.build_report(result)

    def test_html_report_without_issues_needs_no_browser(self):
        book = self.import_rows([['Line', 'A', 'B', 'Average'], ['X', 1, 3, 2]])
        result = compare.run([self.rule(book)])
        with patch('spreadsheet_report.capture_images', wraps=html_report.capture_images) as captures:
            document = html_report.build_report(result)
        captures.assert_called_once_with([])
        self.assertIn('Không có điểm khác nhau', document)
        self.assertNotIn('<article id="finding-', document)


if __name__ == '__main__':
    unittest.main()
