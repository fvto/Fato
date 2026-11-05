from contextlib import contextmanager
from datetime import datetime
import json
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import sfc_controller as controller
import sfc_downloads as downloads
import sfc_export_merge as merge
from sfc_excel import file_info


class FactoryMergeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.patches = [patch('sfc_excel.EXPORT_ROOT', self.root), patch.object(downloads, 'EXPORT_ROOT', self.root)]
        for item in self.patches:
            item.start()

    def tearDown(self):
        for item in reversed(self.patches):
            item.stop()
        self.temp.cleanup()

    def source(self, plant, extra=False, preamble=True, metadata=False, actual=None):
        path = self.root / f'MQAA-{plant}.xlsx'
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = 'MQAA'
        if preamble:
            sheet.append(['MQAA REPORT - All'])
            sheet.append(['From: 2026-10-19 To: 2026-10-24'])
        header = ['Audit Date', 'Audit form', 'Factory', 'MQAA%', 'Remark']
        sheet.append(header + (['Extra'] if extra else []))
        for _ in range(2):
            sheet.append([datetime(2026, 10, 19), 'Cutting', actual or plant, 0.8, '=literal text'])
        sheet.append([None, None, None, '=AVERAGE(D5:D6)'])
        sheet.append([None, None, None, 0.8])
        if metadata:
            book.create_sheet('NguonTheoNgay').append(['File sources', 'Count'])
        book.save(path)
        book.close()
        return dict(file_info(path), plant=plant)

    def test_seven_factories_merge_without_deduplicating_audits_or_counting_summaries(self):
        scope = controller.SEVEN_FACTORIES.copy()
        files = [self.source(plant, metadata=(plant == 'VH2'), preamble=(plant != 'VH2')) for plant in scope]
        target = self.root / 'Combined.xlsx'
        result = merge.combine_factory_exports(files, target, scope)
        self.assertEqual(result['rows'], 14)
        self.assertEqual(result['merge_layout'], 'combined_table')
        book = openpyxl.load_workbook(target)
        try:
            rows = list(book['Combined'].values)
            self.assertEqual(rows[0][:2], ('Source Factory', 'Source File'))
            self.assertEqual([row[0] for row in rows[1:]], [plant for plant in scope for _ in range(2)])
            self.assertEqual(book['Combined']['G2'].value, '=literal text')
            self.assertEqual(book['Combined']['G2'].data_type, 's')
            self.assertEqual(book['Combined']['F2'].number_format, '0.0%')
            self.assertEqual(book.sheetnames, ['Combined'])
            self.assertIn('VH2: MQAA-VH2.xlsx', book['Combined']['A1'].comment.text)
        finally:
            book.close()
        self.assertTrue(all(Path(item['path']).is_file() for item in files))

    def test_wrong_schema_or_missing_factory_never_publishes_a_combined_file(self):
        files = [self.source('VH'), self.source('VH2', extra=True)]
        target = self.root / 'Combined.xlsx'
        with self.assertRaisesRegex(ValueError, 'Cấu trúc cột'):
            merge.combine_factory_exports(files, target, ['VH', 'VH2'])
        self.assertFalse(target.exists())
        with self.assertRaisesRegex(ValueError, 'Chưa đủ file'):
            merge.combine_factory_exports(files[:1], target, ['VH', 'VH2'])
        self.assertFalse(list(self.root.glob('*.partial.xlsx')))

    def test_wrong_factory_and_changed_file_are_rejected(self):
        wrong = self.source('VH', actual='VH2')
        with self.assertRaisesRegex(ValueError, 'chứa dữ liệu nhà máy VH2'):
            merge.combine_factory_exports([wrong], self.root / 'wrong.xlsx', ['VH'])
        original = self.source('VH')
        Path(original['path']).write_bytes(b'changed')
        with self.assertRaisesRegex(ValueError, 'đã thay đổi'):
            merge.combine_factory_exports([original], self.root / 'changed.xlsx', ['VH'])

    def test_complex_reports_stack_all_source_sheets_into_one_sheet(self):
        files = []
        for plant in ['VH', 'VH2']:
            path = self.root / f'Aggregate-{plant}.xlsx'
            book = openpyxl.Workbook()
            book.active.append(['Factory analysis'])
            book.active.append([10, 20])
            book.create_sheet('Details').append(['All original content'])
            book.save(path)
            book.close()
            files.append(dict(file_info(path), plant=plant))
        result = merge.combine_factory_exports(files, self.root / 'Aggregate.xlsx', ['VH', 'VH2'])
        self.assertEqual(result['merge_layout'], 'stacked_rows')
        book = openpyxl.load_workbook(result['path'], read_only=True)
        try:
            self.assertEqual(book.sheetnames, ['Combined'])
            rows = list(book['Combined'].iter_rows(min_row=2, max_col=6, values_only=True))
            self.assertIn(('VH', 'Aggregate-VH.xlsx', 'Sheet', 2, 10, 20), rows)
            self.assertIn(('VH2', 'Aggregate-VH2.xlsx', 'Details', 1, 'All original content', None), rows)
        finally:
            book.close()

    def test_factory_matrix_is_transposed_with_empty_factories_on_one_sheet(self):
        files = []
        scope = ['VH', 'VH2', 'JV']
        for plant in scope:
            path = self.root / f'Factory-{plant}.xlsx'
            book = openpyxl.Workbook()
            sheet = book.active
            sheet.append(['MQAA REPORT - Factory'])
            sheet.append(['From: 2026-10-19  To: 2026-10-24'])
            if plant != 'JV':
                sheet.append([])
                sheet.append(['Factory', 'Factory 1', 'RUBBER'])
                sheet.append(['Avg Rate', ' 97.8%', '100.0%'])
                sheet.append(['Line', 'Line 1', 'B - 1'])
                sheet.append(['Issue', '=literal issue', 'Issue 1\nIssue 2'])
            book.save(path)
            book.close()
            files.append(dict(file_info(path), plant=plant))
        result = merge.combine_factory_exports(files, self.root / 'Factory_Combined.xlsx', scope)
        self.assertEqual(result['merge_layout'], 'combined_table')
        self.assertEqual(result['rows'], 5)
        book = openpyxl.load_workbook(result['path'])
        try:
            self.assertEqual(book.sheetnames, ['Combined'])
            rows = list(book['Combined'].values)
            self.assertEqual(rows[0][3:7], ('Factory', 'Avg Rate', 'Line', 'Issue'))
            self.assertEqual(rows[1][3:], ('Factory 1', ' 97.8%', 'Line 1', '=literal issue', 'Data'))
            self.assertEqual(rows[2][6], 'Issue 1\nIssue 2')
            self.assertEqual(rows[-1][0], 'JV')
            self.assertEqual(rows[-1][-1], 'No data')
            self.assertEqual(book['Combined']['G2'].data_type, 's')
        finally:
            book.close()

    def test_batch_combines_only_after_all_exports_and_restores_scope(self):
        scope = ['VH', 'VH2']
        files = {plant: self.source(plant) for plant in scope}
        page = MagicMock()
        page.url = 'https://sfc.chingluh.com/vhmes/QA/MQAAReport.aspx'
        page.title.return_value = 'MQAA'
        page.context.cookies.return_value = [{'name': 'MyFactoryName', 'value': 'VH'}]
        browser = SimpleNamespace(close=lambda: None)
        @contextmanager
        def playwright():
            yield SimpleNamespace(chromium=SimpleNamespace(connect_over_cdp=lambda *a, **k: browser))
        fake = SimpleNamespace(SELECTED_FACTORIES=scope.copy(), SEVEN_FACTORIES=scope,
            CDP_URL='', SFC_CONNECT_WAIT_MS=1000, SFC_WAIT_MS=1000,
            sync_playwright=playwright, get_sfc_page=lambda b: page)
        def switch(plant):
            fake.SELECTED_FACTORIES = [plant]
            return {'success': True}
        def execute(action, target_id=None, params=None):
            return {'success': True} if action == 'restore_filters' else {
                'success': True, 'downloaded_file': files[params['_expected_plant']]}
        fake.switch_factory = switch
        fake._execute_dynamic_action = execute
        with patch.object(downloads, 'wait_idle'), patch.object(downloads, 'capture_filters', return_value={}):
            result = downloads.export_selected(fake, 'click_button', 'btnExcel', {'combine': True}, scope)
        self.assertTrue(result['success'], result)
        self.assertEqual(len(result['downloaded_files']), 1)
        self.assertEqual(len(result['source_files']), 2)
        self.assertEqual(result['downloaded_file']['rows'], 4)
        self.assertEqual(fake.SELECTED_FACTORIES, scope)
        manifest = json.loads(next(self.root.glob('batch_*/manifest.json')).read_text(encoding='utf-8'))
        self.assertEqual(len(manifest['combined_files']), 1)
        def fail(action, target_id=None, params=None):
            if action != 'restore_filters' and params['_expected_plant'] == 'VH2':
                return {'success': False, 'error': 'Missing export'}
            return execute(action, target_id, params)
        fake._execute_dynamic_action = fail
        with patch.object(downloads, 'wait_idle'), patch.object(downloads, 'capture_filters', return_value={}), \
                patch.object(merge, 'combine_factory_exports') as combine:
            result = downloads.export_selected(fake, 'click_button', 'btnExcel', {'combine': True}, scope)
        self.assertFalse(result['success'])
        combine.assert_not_called()
        self.assertEqual(len(result['downloaded_files']), 1)
        self.assertEqual(fake.SELECTED_FACTORIES, scope)

    def test_one_factory_combine_routes_through_batch(self):
        with patch.object(controller, 'SELECTED_FACTORIES', ['VH2']), \
                patch.object(controller, 'browser_reserved', return_value=False), \
                patch.object(downloads, 'export_selected', return_value={'success': True}) as batch:
            controller.execute_dynamic_action('click_button', 'btnExcel', {'combine': True})
        self.assertEqual(batch.call_args.args[-1], ['VH2'])


if __name__ == '__main__':
    unittest.main()
