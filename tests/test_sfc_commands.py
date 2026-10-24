import json
from pathlib import Path
import sys
import tempfile
import threading
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import openpyxl
import sfc_commands as commands
import sfc_excel as excel
import sfc_command_jobs as jobs

PROMPT = ('Tải excel HFPA từ tháng 1 - tháng 10 năm 2026 4 nhà máy VH VH2 JV JV2 giúp tôi, '
          'chia nhỏ file từng tháng sau đó combine lại thành 1 file, áp dụng tương tự cho 4 nhà máy, '
          'phải xuất Nike (external)')
HEADER = ['Factory Group Cd_UccId', 'Audit Date (dd-mmm-yyyy)', 'Internal AUDIT ID', 'Plant', 'Line']


def fixture(path, day, plant, extra_header=None):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(HEADER + ([extra_header] if extra_header else []))
    sheet.append([plant, day, plant + day, '=literal', '001'] + (['x'] if extra_header else []))
    sheet.append(['', '1', '', '', ''])  # DevExpress summary, not an audit record.
    workbook.save(path)
    workbook.close()


class Controller:
    def __init__(self, fail_at=None):
        self.plant = None
        self.calls = []
        self.fail_at = fail_at

    def create_hfpa_session(self, plant):
        self.plant = plant
        owner = self
        class FakeSession:
            def prepare(self):
                pass

            def export_period(self, first, last, nike, path, factory=''):
                result = owner.apply_sfc_filters(date_from=first, date_to=last, nike=nike,
                    export_path=path, factory=factory, expected_plant=plant, action='excel')
                commands.require_success(result)
                yield {'file': result['file']}
        return FakeSession()

    def switch_factory(self, plant):
        self.plant = plant
        return dict(success=True, message='Switched ' + plant)

    def navigate_to(self, page):
        return dict(success=True, title=page)

    def apply_sfc_filters(self, **params):
        self.calls.append((self.plant, params))
        if len(self.calls) == self.fail_at:
            return dict(success=False, error='Download failed')
        assert params['expected_plant'] == self.plant
        assert params['nike'] is True
        assert params['factory'] == ''
        path = Path(str(params['export_path']) + '.xlsx')
        fixture(path, params['date_from'], self.plant)
        rows = excel.validate_export(path, params['date_from'], params['date_to'], True)
        return dict(success=True, file=excel.file_info(path, rows))


class CommandTests(unittest.TestCase):
    def test_exact_user_prompt(self):
        intent = commands.parse_intent(PROMPT)
        self.assertEqual(intent['target_plants'], ['VH', 'VH2', 'JV', 'JV2'])
        self.assertEqual((intent['date_from'], intent['date_to']), ('2026-01-01', '2026-10-31'))
        self.assertIsNone(intent['target_line'])
        self.assertTrue(intent['nike'] and intent['monthly'] and intent['combine'])
        self.assertEqual(len(commands.periods(intent)), 10)
        self.assertEqual(commands.periods(intent)[1], ('2026-02-01', '2026-02-28'))

    def test_existing_single_commands(self):
        self.assertEqual(commands.parse_intent('Đổi sang nhà máy VH3')['target_plant'], 'VH3')
        self.assertEqual(commands.parse_intent('Chọn chuyền VL - D3 ngày hôm nay rồi tìm kiếm')['target_line'], 'VL - D3')
        intent = commands.parse_intent('Xuất excel HFPA nội bộ VH2 từ 20/10/2026 đến 24/10/2026')
        self.assertIs(intent['nike'], False)
        self.assertEqual(intent['date_from'], '2026-10-20')

    def test_invalid_dates_fail_before_actions(self):
        for prompt in ['Xuất excel HFPA từ tháng 10 - tháng 1 năm 2026', 'HFPA từ 31/02/2026 đến 01/03/2026']:
            with self.assertRaises(ValueError):
                commands.parse_intent(prompt)

    def test_40_downloads_four_verified_merges(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with patch.object(commands, 'EXPORT_ROOT', root), patch.object(excel, 'EXPORT_ROOT', root):
                controller = Controller()
                events = list(commands.execute(controller, PROMPT))
                self.assertEqual(events[-1]['status'], 'done', events[-1])
                self.assertEqual(len(controller.calls), 40)
                self.assertEqual(len(events[-1]['files']), 4)
                for result in events[-1]['files']:
                    self.assertEqual(result['rows'], 10)
                    workbook = openpyxl.load_workbook(result['path'], read_only=True, data_only=False)
                    try:
                        self.assertEqual(len(list(workbook['HFPA'].values)), 11)
                        self.assertEqual(len(list(workbook['NguonTheoThang'].values)), 11)
                        self.assertEqual(workbook['HFPA']['D2'].data_type, 's')
                    finally:
                        workbook.close()
                manifest = json.loads(next(root.glob('*/manifest.json')).read_text(encoding='utf-8'))
                self.assertEqual(manifest['status'], 'done')
                self.assertEqual(len(manifest['monthly_files']), 40)

    def test_failure_never_reports_done(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with patch.object(commands, 'EXPORT_ROOT', root), patch.object(excel, 'EXPORT_ROOT', root):
                controller = Controller(fail_at=6)
                events = list(commands.execute(controller, PROMPT))
                self.assertEqual(events[-1]['status'], 'error')
                self.assertFalse(any(event['status'] == 'done' for event in events))
                manifest = json.loads(next(root.glob('*/manifest.json')).read_text(encoding='utf-8'))
                self.assertEqual(len(manifest['monthly_files']), 5)
                self.assertEqual(manifest['status'], 'error')

    def test_narrow_to_vh_reuses_verified_months_and_pauses_other_plants(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with patch.object(commands, 'EXPORT_ROOT', root), patch.object(excel, 'EXPORT_ROOT', root):
                previous = list(commands.execute(Controller(fail_at=3), PROMPT))
                self.assertEqual(previous[-1]['status'], 'error')
                manifest_path = next(root.glob('*/manifest.json'))
                (manifest_path.parent / '_parts').mkdir()
                prompt = ('Tải excel HFPA từ tháng 1 - tháng 10 năm 2026 nhà máy VH, '
                          'chia nhỏ file từng tháng sau đó combine lại thành 1 file, phải xuất Nike (external)')
                controller = Controller()
                events = list(commands.execute(controller, prompt))
                self.assertEqual(events[-1]['status'], 'done', events[-1])
                self.assertEqual(len(controller.calls), 8)  # first two months already verified
                self.assertEqual({plant for plant, _ in controller.calls}, {'VH'})
                manifest = json.loads(manifest_path.read_text(encoding='utf-8'))
                self.assertEqual(manifest['paused_plants'], ['JV', 'JV2', 'VH2'])
                self.assertEqual(len(manifest['monthly_files']), 10)
                self.assertEqual(len(manifest['combined_files']), 1)

    def test_reject_wrong_month_and_schema(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with patch.object(excel, 'EXPORT_ROOT', root):
                a, b = root / 'a.xlsx', root / 'b.xlsx'
                fixture(a, '2026-01-01', 'VH')
                fixture(b, '2026-02-01', 'VH', 'Different column')
                with self.assertRaises(ValueError):
                    excel.validate_export(a, '2026-02-01', '2026-02-28', True)
                info = [dict(excel.file_info(p), plant='VH', nike=True, date_from='2026-01-01', date_to='2026-02-28') for p in (a,b)]
                with self.assertRaises(ValueError):
                    excel.combine_exports(info, root / 'combined.xlsx', 'VH', True)
                self.assertFalse((root / 'combined.xlsx').exists())
                with self.assertRaises(ValueError):
                    excel.combine_exports(info, root / 'combined.xlsx', 'JV', True)

    def test_failed_switch_stops_before_export(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(commands, 'EXPORT_ROOT', Path(directory)):
            controller = Controller()
            class FailedSession:
                def prepare(self):
                    raise RuntimeError('Wrong plant')
            controller.create_hfpa_session = lambda _: FailedSession()
            events = list(commands.execute(controller, 'Xuất excel HFPA VH2 ngày 2026-01-01 NIKE'))
            self.assertEqual(events[-1]['status'], 'error')
            self.assertEqual(controller.calls, [])

    def test_single_hfpa_export_uses_verified_exporter_without_dom(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with patch.object(commands, 'EXPORT_ROOT', root), patch.object(excel, 'EXPORT_ROOT', root):
                controller = Controller()
                controller.switch_factory = lambda _: self.fail('Excel must not wait for DOM')
                events = list(commands.execute(controller, 'Xuất Excel HFPA VH từ ngày 2026-10-24 đến ngày 2026-10-24 NIKE (external)'))
                self.assertEqual(events[-1]['status'], 'done', events[-1])
                self.assertEqual(len(controller.calls), 1)
                self.assertEqual(len(events[-1]['files']), 1)

    def test_disconnect_does_not_stop_job(self):
        release, finished = threading.Event(), threading.Event()
        class Worker:
            def execute_command_stream(self, prompt):
                yield commands.event(1, 'running', 'Started')
                release.wait(3)
                yield commands.event(5, 'done', 'Finished')
                finished.set()
        with tempfile.TemporaryDirectory() as directory, patch.object(commands, 'EXPORT_ROOT', Path(directory)):
            job = jobs.start(Worker(), 'test')
            stream = jobs.stream(job['id'])
            self.assertIn('Started', next(stream))
            stream.close()
            release.set()
            self.assertTrue(finished.wait(3))
            self.assertEqual(jobs.snapshot(job['id'])['status'], 'done')
            self.assertIn('Finished', ''.join(jobs.stream(job['id'])))


if __name__ == '__main__':
    unittest.main()
