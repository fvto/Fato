"""Regression coverage for the NIKE review's real controller contract."""
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase, mock

from openpyxl import Workbook

import sfc_quarterly_defect_review_acquisition as acquisition
from report_runtime import Cancelled


def workbook(path, metric, factory='VH', day='2026-06-01', internal=False):
    path.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    if metric == 'HFPA':
        sheet.append(['Factory Group Cd_UccId', 'Audit Date (DD-MMM-YYYY)',
                      'Internal Audit ID (YYMMDDno+Plant+Line)', 'Plant', 'Line', 'Model',
                      'Audit Sample Size', "Defective Q'ty(pair)",
                      "Defect Q'ty (Total Defect Q'ty)", 'Bond gap / Bonding'])
        sheet.append(['CL_' + factory, day, 'audit-1', 'N14', 'Line 1', 'Model A', 100, 4, 4, 4])
    else:
        sheet.append(['Factory Group', 'Factory Code', 'Inspection Date (DD-MMM-YYYY)',
                      'Plant', 'Line', 'Model', 'Total A Grade Pairs',
                      'Total B Grade Pairs (Reject)', 'Total C Grade Pairs (Reject)',
                      'Total Rework', '1st Time Rework Pairs', 'Bond gap / Rat hole_Rework'])
        sheet.append(['CLG', factory, day, 'N14', 'Line 1', 'Model A', 100, 0, 0, 4, 4, 4])
    if internal:
        sheet.cell(1, sheet.max_column + 1, 'SO#' if metric == 'HFPA' else 'Metric')
    book.save(path)
    book.close()
    return path


class Controller:
    def __init__(self, root, hfpa_result=None, fail_hfpa=False):
        self.root = root
        self.calls = []
        self.hfpa_result = hfpa_result
        self.fail_hfpa = fail_hfpa

    def switch_factory(self, plant):
        self.calls.append(('switch', plant))
        self.plant = plant
        return {'success': True, 'current_plant': plant}

    def apply_sfc_filters(self, **params):
        self.calls.append(('HFPA', params))
        if self.hfpa_result is not None:
            return self.hfpa_result
        if self.fail_hfpa:
            return {'success': False, 'error': 'HFPA export failed'}
        path = workbook(Path(str(params['export_path']) + '.xlsx'), 'HFPA', self.plant,
                        day=params['date_from'])
        return {'success': True, 'file': {'path': str(path)}}

    def navigate_to(self, url):
        self.calls.append(('navigate', url))
        return {'success': True}

    def execute_dynamic_action(self, action, target, params):
        self.calls.append(('FTT', params))
        self.action = action
        path = workbook(self.root / 'download.xlsx', 'FTT', self.plant)
        return {'success': True, 'downloaded_file': {'path': str(path)}}


class QuarterlyDefectReviewAcquisitionTests(TestCase):
    def run_acquisition(self, root, ctrl, stage=lambda *_: None):
        run = {'period': 'FY27-Q1', 'date_from': '2026-06-01', 'date_to': '2026-06-30'}
        with mock.patch.object(acquisition, 'PLANTS', ('VH',)), \
             mock.patch.object(acquisition, 'acquire_factory_sandals_catalog', return_value=None):
            result = acquisition.acquire(root / 'workspace', run, stage, controller=ctrl,
                                         cache_root=root / 'cache')
        return run, result

    def test_both_nike_sources_use_dom_and_stage_controller_downloads(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            ctrl = Controller(root)
            run, result = self.run_acquisition(root, ctrl)
            self.assertEqual(result['export_transport'], 'DOM')
            self.assertEqual(result['missing'], [])
            self.assertEqual(result['errors'], [])
            self.assertEqual(len(result['downloaded']), 2)
            hfpa = next(p for name, p in ctrl.calls if name == 'HFPA')
            self.assertIs(hfpa['nike'], True)
            self.assertEqual(hfpa['factory'], '')
            self.assertEqual(hfpa['action'], 'excel')
            self.assertEqual(hfpa['expected_plant'], 'VH')
            self.assertEqual(hfpa['date_to'], '2026-06-07')
            ranges = [p for name, p in ctrl.calls if name == 'HFPA']
            self.assertEqual([(p['date_from'], p['date_to']) for p in ranges], [
                ('2026-06-01', '2026-06-07'), ('2026-06-08', '2026-06-14'),
                ('2026-06-15', '2026-06-21'), ('2026-06-22', '2026-06-28'),
                ('2026-06-29', '2026-06-30'),
            ])
            ftt = next(p for name, p in ctrl.calls if name == 'FTT')
            self.assertEqual(ftt['report_type'], 'nike')
            self.assertEqual(ftt['_expected_plant'], 'VH')
            self.assertEqual(ctrl.action, 'ftt_station_export')
            self.assertEqual([name for name, _ in ctrl.calls if name in ('HFPA', 'FTT')], ['HFPA'] * 5 + ['FTT'])
            self.assertEqual(len(list((root / 'workspace/Inputs').rglob('*.xlsx'))), 2)

    def test_timeout_stops_before_ftt_and_does_not_retry_or_switch_source(self):
        with TemporaryDirectory() as tmp:
            ctrl = Controller(Path(tmp), hfpa_result={'success': False, 'timed_out': True, 'error': 'still busy'})
            with self.assertRaisesRegex(acquisition.DomExportPending, 'still busy'):
                self.run_acquisition(Path(tmp), ctrl)
            self.assertEqual(sum(name == 'HFPA' for name, _ in ctrl.calls), 1)
            self.assertNotIn('FTT', [name for name, _ in ctrl.calls])

    def test_busy_controller_stops_without_another_export(self):
        with TemporaryDirectory() as tmp:
            ctrl = Controller(Path(tmp), hfpa_result={'success': False, 'busy': True, 'error': 'reserved'})
            with self.assertRaises(acquisition.DomExportPending):
                self.run_acquisition(Path(tmp), ctrl)

    def test_timeout_records_the_blocked_slot_and_confirmed_count(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            ctrl = Controller(root, hfpa_result={'success': False, 'timed_out': True, 'error': 'timeout'})
            run = {'period': 'FY27-Q1', 'date_from': '2026-06-01', 'date_to': '2026-06-30'}
            with mock.patch.object(acquisition, 'PLANTS', ('VH',)), \
                 mock.patch.object(acquisition, 'acquire_factory_sandals_catalog', return_value=None):
                with self.assertRaises(acquisition.DomExportPending) as caught:
                    acquisition.acquire(root, run, lambda *_: None, controller=ctrl, cache_root=root / 'cache')
            state = run['acquisition']
            self.assertEqual(state['blocked']['plant'], 'VH')
            self.assertEqual(state['blocked']['metric'], 'HFPA')
            self.assertEqual(state['blocked']['reason'], 'timeout')
            self.assertEqual(len(state['errors']), 1)
            self.assertIn('HFPA NIKE VH', caught.exception.user_message)
            self.assertIn('0/2', caught.exception.user_message)
            from report_runtime import friendly_error
            self.assertEqual(friendly_error(caught.exception), caught.exception.user_message)
            self.assertNotIn('FTT', [name for name, _ in ctrl.calls])

    def test_jv1_export_is_jv_but_jv2_stays_a_different_factory(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            slot = {'plant': 'JV', 'metric': 'FTT', 'date_from': '2026-06-01', 'date_to': '2026-06-30'}
            acquisition._validate_nike_download(workbook(root / 'JV1.xlsx', 'FTT', 'JV1'), slot)
            with self.assertRaisesRegex(ValueError, 'sai nha may'):
                acquisition._validate_nike_download(workbook(root / 'JV2.xlsx', 'FTT', 'JV2'), slot)

    def test_wrapped_playwright_timeout_is_not_retried(self):
        with TemporaryDirectory() as tmp:
            ctrl = Controller(Path(tmp), hfpa_result={
                'success': False, 'error': 'Page.expect_download: Timeout 900000ms exceeded',
            })
            with self.assertRaises(acquisition.DomExportPending):
                self.run_acquisition(Path(tmp), ctrl)
            self.assertNotIn('FTT', [name for name, _ in ctrl.calls])

    def test_completed_error_does_not_prevent_ftt_from_running(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            _, result = self.run_acquisition(root, Controller(root, fail_hfpa=True))
            self.assertEqual(len(result['errors']), 1)
            self.assertEqual(result['downloaded'][0]['group'], 'ftt_nike')

    def test_wrong_factory_and_out_of_period_downloads_are_rejected(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            slot = dict(metric='FTT', plant='VH', date_from='2026-06-01', date_to='2026-06-30')
            with self.assertRaisesRegex(ValueError, 'sai nha may'):
                acquisition._validate_nike_download(workbook(root / 'JV.xlsx', 'FTT', 'JV'), slot)
            with self.assertRaisesRegex(ValueError, 'ngoai ky'):
                acquisition._validate_nike_download(workbook(root / 'July.xlsx', 'FTT', day='2026-07-01'), slot)

    def test_internal_cache_is_not_reused_as_nike(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            workbook(root / 'cache/HFPA_NIKE/VH/VH_hfpa_nike_2026-06-01_2026-06-30.xlsx',
                     'HFPA', internal=True)
            ctrl = Controller(root)
            _, result = self.run_acquisition(root, ctrl)
            self.assertTrue(any(name == 'HFPA' for name, _ in ctrl.calls))
            self.assertEqual(len(result['downloaded']), 2)

    def test_verified_cache_skips_exports(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            for metric in ('HFPA', 'FTT'):
                group = metric.lower() + '_nike'
                workbook(root / f'cache/{group.upper()}/VH/VH_{group}_2026-06-01_2026-06-30.xlsx', metric)
            ctrl = Controller(root)
            _, result = self.run_acquisition(root, ctrl)
            self.assertEqual(result['staged_files'], 2)
            self.assertEqual(ctrl.calls, [])

    def test_completed_month_is_reused_after_safe_cancellation(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            first = Controller(root)
            def cancel_after_month(message, *_):
                if message.startswith('[DOM') and 'Da luu va kiem tra' in message:
                    raise Cancelled('stop after completed monthly HFPA export')
            with self.assertRaises(Cancelled):
                self.run_acquisition(root, first, cancel_after_month)
            second = Controller(root)
            _, result = self.run_acquisition(root, second)
            calls = [p for name, p in second.calls if name == 'HFPA']
            self.assertEqual(calls, [])
            self.assertEqual(len(result['downloaded']), 1)
            self.assertEqual(result['downloaded'][0]['group'], 'ftt_nike')
            self.assertEqual(result['staged_files'], 2)

    def test_completed_dom_parts_resume_without_exporting_the_first_week_again(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            first = Controller(root)
            def cancel_next_part(message, *_):
                if message == 'DOM HFPA NIKE VH: 2026-06-08 - 2026-06-14':
                    raise Cancelled('stop after first part')
            with self.assertRaises(Cancelled):
                self.run_acquisition(root, first, cancel_next_part)
            second = Controller(root)
            _, result = self.run_acquisition(root, second)
            calls = [p for name, p in second.calls if name == 'HFPA']
            self.assertEqual(len(calls), 4)
            self.assertEqual(calls[0]['date_from'], '2026-06-08')
            self.assertEqual(len(result['downloaded']), 2)

    def test_cancel_is_not_swallowed_or_followed_by_an_export(self):
        with TemporaryDirectory() as tmp:
            ctrl = Controller(Path(tmp))
            def stage(message, *_):
                if message.startswith('[DOM'):
                    raise Cancelled('stop')
            with self.assertRaises(Cancelled):
                self.run_acquisition(Path(tmp), ctrl, stage)
            self.assertEqual(ctrl.calls, [])
