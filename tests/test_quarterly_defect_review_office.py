import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from report_tools.quarterly_defect_review import office_output as office


class ComError(Exception):
    def __init__(self, code):
        super().__init__(code, 'COM failure')
        self.hresult = code


class WpsReopenTests(unittest.TestCase):
    def test_reconnects_with_a_new_proxy_when_wps_is_still_shutting_down(self):
        client, stale, fresh = MagicMock(), MagicMock(), MagicMock()
        stale.Presentations.Open.side_effect = ComError(-2147023174)
        stale.Quit.side_effect = ComError(-2147417848)
        client.DispatchEx.side_effect = [stale, fresh]
        with patch.object(office.time, 'sleep'):
            app, doc = office._open_verification_session(client, Path('saved.pptx'))
        self.assertIs(app, fresh)
        self.assertIs(doc, fresh.Presentations.Open.return_value)
        self.assertEqual(client.DispatchEx.call_count, 2)
        self.assertEqual(stale.Presentations.Open.call_count, 1)
        fresh.Presentations.Open.assert_called_once_with('saved.pptx', True, False, False)

    def test_no_retry_for_permissions_and_bounded_retry_for_rpc(self):
        client = MagicMock()
        client.DispatchEx.side_effect = ComError(-2147024891)
        with self.assertRaises(ComError):
            office._open_verification_session(client, Path('saved.pptx'))
        self.assertEqual(client.DispatchEx.call_count, 1)
        client.reset_mock()
        client.DispatchEx.side_effect = ComError(-2147023174)
        with patch.object(office.time, 'sleep'), self.assertRaises(ComError):
            office._open_verification_session(client, Path('saved.pptx'))
        self.assertEqual(client.DispatchEx.call_count, 3)

    def test_shutdown_only_ignores_already_disconnected_apps(self):
        app = MagicMock()
        app.Quit.side_effect = ComError(-2147023174)
        office._quit_after_close(app)
        app.Quit.side_effect = AttributeError('PowerPoint.Application.Quit')
        office._quit_after_close(app)
        app.Quit.side_effect = ComError(-2147024891)
        with self.assertRaises(ComError):
            office._quit_after_close(app)

    def test_saved_document_close_tolerates_only_disconnection(self):
        document = MagicMock()
        document.Close.side_effect = ComError(-2147417848)
        office._close_saved_document(document, False)
        document.Close.assert_called_once_with(False)
        document.Close.side_effect = ComError(-2147024891)
        with self.assertRaises(ComError):
            office._close_saved_document(document)

    def _build(self, directory, reopened_slides=7):
        source = Path(directory) / 'source.xlsx'
        source.write_bytes(b'workbook fixture')
        target = Path(directory) / 'saved.pptx'
        client, pythoncom = MagicMock(), MagicMock()
        xl, old_ppt, fresh_ppt, workbook, presentation, reopened = [MagicMock() for _ in range(6)]
        events = []
        apps = iter([xl, old_ppt, fresh_ppt])

        def dispatch(name):
            app = next(apps)
            events.append('new ' + ('excel' if app is xl else 'old ppt' if app is old_ppt else 'fresh ppt'))
            return app

        client.DispatchEx.side_effect = dispatch
        xl.Path = old_ppt.Path = fresh_ppt.Path = 'C:/Kingsoft/WPS'
        xl.Workbooks.Open.return_value = workbook
        old_ppt.Presentations.Add.return_value = presentation
        fresh_ppt.Presentations.Open.return_value = reopened
        presentation.Slides.Count = 7
        reopened.Slides.Count = reopened_slides
        reopened.Slides.Item.return_value.Shapes.Count = 0
        workbook.Worksheets.return_value.ChartObjects.return_value.Count = 0
        presentation.SaveAs.side_effect = lambda path: (events.append('save'), Path(path).write_bytes(b'Native WPS bytes'))
        presentation.Close.side_effect = lambda: events.append('close ppt')
        workbook.Close.side_effect = lambda *a: events.append('close workbook')
        xl.Quit.side_effect = lambda: events.append('quit excel')

        def quit_dead_ppt():
            events.append('quit old ppt')
            raise ComError(-2147023174)

        old_ppt.Quit.side_effect = quit_dead_ppt
        # Accessing a dead proxy, as the old implementation did, must fail.
        old_ppt.Presentations.Open.side_effect = AssertionError('Reused dead WPS connection')
        ranges = {key: ('Sheet', 'A1', 'B2') for key in (
            'hfpa_scope_table', 'ftt_scope_table', 'hfpa_3month_table',
            'hfpa_group_table', 'hfpa_detail_matrix', 'hfpa_model_matrix',
            'ftt_group_table', 'ftt_detail_matrix', 'ftt_model_matrix')}
        result = {'catalog_version': 'test', 'mapping_version': 'test',
                  'detail_support_level': 'group_only', 'manifests': [],
                  'dr_config': {'status': 'unconfirmed', 'note': 'Test'},
                  'review': {'status_label': 'Review', 'factual_notes': [], 'missing_scopes': [], 'items': []}}
        win32com = types.ModuleType('win32com')
        win32com.client = client
        with patch.dict(sys.modules, {'pythoncom': pythoncom, 'win32com': win32com, 'win32com.client': client}), \
                patch.object(office, '_copy_range_image', return_value=(100, 100)), \
                patch.object(office, '_add_fitted_picture'), \
                patch.object(office, '_add_slide_header'), \
                patch.object(office, '_add_note_box'), \
                patch.object(office, '_inspect_pptx_package', return_value={'chart_xml_count': 0, 'external_links': []}):
            try:
                meta = office.build_office_presentation(target, final_excel_path=source, excel_meta={'ranges': ranges},
                    period='FY27-Q1', date_from='2026-06-01', date_to='2026-08-31',
                    hfpa_result=result, ftt_result=result, preview_dir=Path(directory) / 'previews')
            finally:
                pythoncom.CoUninitialize.assert_called_once()
        return meta, events, target, old_ppt, reopened

    def test_save_close_excel_then_fresh_reopen_and_preserve_native_bytes(self):
        with tempfile.TemporaryDirectory() as directory:
            meta, events, target, old, reopened = self._build(directory)
            self.assertTrue(meta['reopen_verified'])
            self.assertEqual(meta['verification_engine'], 'wps')
            self.assertEqual(meta['slide_count'], 7)
            self.assertEqual(len(meta['slide_previews']), 7)
            self.assertEqual(target.read_bytes(), b'Native WPS bytes')
            self.assertEqual(events, ['new excel', 'new old ppt', 'save', 'close ppt',
                                     'close workbook', 'quit excel', 'quit old ppt', 'new fresh ppt'])
            old.Presentations.Open.assert_not_called()
            reopened.Close.assert_called_once()

    def test_reopen_still_rejects_incomplete_slides(self):
        with tempfile.TemporaryDirectory() as directory, self.assertRaises(office.OfficeExportError):
            self._build(directory, reopened_slides=6)


if __name__ == '__main__':
    unittest.main()
