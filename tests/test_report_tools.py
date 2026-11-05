import io
import json
from datetime import datetime
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch, MagicMock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import sfc_report_tools as reports
import sfc_dashboard_server as server
import openpyxl


class ReportToolsTests(unittest.TestCase):
    def test_read_json_retries_a_temporary_windows_file_lock(self):
        with (patch.object(Path, "read_text", side_effect=[PermissionError("sharing violation"), '{"status":"running"}']) as read,
              patch.object(reports.time, "sleep")):
            self.assertEqual(reports.read_json("run.json"), {"status": "running"})
        self.assertEqual(read.call_count, 2)

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.original = self.root / 'original'
        self.original.mkdir()
        self.engines = self.root / 'engines'
        (self.engines / 'copq').mkdir(parents=True)
        (self.engines / 'copq' / 'engine.py').write_text('# fixture only', encoding='utf8')
        self.exports = self.root / 'exports'
        self.exports.mkdir()
        self.run_root = self.root / 'runs'
        self.patches = [patch.object(reports, 'TOOLS_ROOT', self.engines),
                        patch.object(reports, 'RUN_ROOT', self.run_root),
                        patch.object(reports, 'EXPORT_ROOT', self.exports),
                        patch.object(reports, 'original_root', return_value=self.original)]
        for item in self.patches: item.start()

    def tearDown(self):
        for item in reversed(self.patches): item.stop()
        self.temp.cleanup()

    def source(self, month='09'):
        source = self.original / 'COPQ_Input' / 'September' / f'Cost Of Poor Quality 2026-{month}-01_2026-{month}-30-VH.xls'
        source.parent.mkdir(parents=True, exist_ok=True)
        workbook = openpyxl.Workbook()
        workbook.active.append(['Date','Category','Type','Factory','Model','Ttl cost ($)'])
        workbook.active.append([f'2026-{month}-01','QC','Rework','VH','Fixture',12])
        stream = io.BytesIO()
        workbook.save(stream)
        workbook.close()
        source.write_bytes(stream.getvalue())
        return source

    def test_period_and_site_alias(self):
        self.assertEqual(reports.period_of('Report_2026-09-01_2026-09-30-VH.xls'), '2026-09')
        self.assertEqual(reports.period_of('Report_2026-08-01_2026-09-30-VH.xls'), 'mixed')
        self.assertIsNone(reports.period_of('HFPA 2026-10-03 11_29_50-VH.xlsx'))
        self.assertEqual(reports.plant_of('Cost Of Poor Quality-JV3.xls'), 'JVB')

    def test_mes410_uses_audit_month_instead_of_download_timestamp(self):
        folder = self.original / 'Input' / 'HFPA'
        folder.mkdir(parents=True)
        path = folder / 'HFPA 2026-10-03 11_29_50-VH.xlsx'
        workbook = openpyxl.Workbook()
        workbook.active.append(['Audit Date (dd-mmm-yyyy)', 'Plant'])
        workbook.active.append([datetime(2026, 9, 1), 'VH'])
        workbook.save(path)
        workbook.close()
        item = next(i for i in reports.input_catalog('hfpa') if i['source'] == 'local')
        self.assertEqual(item['period'], '2026-09')

    def test_nested_original_inputs_are_copied_and_never_modified(self):
        original = self.source()
        original_bytes = original.read_bytes()
        items = reports.input_catalog('copq')
        self.assertEqual(len(items), 1)
        self.assertIn('September/', items[0]['id'])
        with patch.object(reports.subprocess, 'Popen') as spawn:
            run = reports.start_run('copq', '2026-09', 'combine', [{'id':items[0]['id']}])
        workspace = reports.run_folder(run['id']) / 'workspace'
        staged = workspace / 'COPQ_Input' / original.with_suffix('.xlsx').name
        book = openpyxl.load_workbook(staged)
        self.assertEqual(book.active['F2'].value,12)
        book.close()
        staged.write_bytes(b'changed isolated copy')
        self.assertEqual(original.read_bytes(), original_bytes)
        self.assertEqual(spawn.call_args.kwargs['cwd'], workspace)
        self.assertEqual(run['status'], 'queued')

    def test_multipart_upload_starts_once_with_isolated_input(self):
        with patch.object(reports.subprocess, 'Popen') as spawn, server.app.test_client() as client:
            response = client.post('/api/report-tools/copq/runs', data={
                'period':'2026-09', 'mode':'combine', 'selections':'[]',
                'copq':(io.BytesIO(b'Date,Category,Type,Factory,Model,Ttl cost ($)\n2026-09-01,QC,Rework,VH,Fixture,12\n'), 'COPQ-VH.csv')},
                content_type='multipart/form-data')
            self.assertEqual(response.status_code, 202, response.json)
            run = response.json
            self.assertEqual(run['inputs'][0]['name'], 'COPQ-VH.csv')
            self.assertTrue((reports.run_folder(run['id'])/'workspace'/'COPQ_Input'/'COPQ-VH.xlsx').is_file())
            spawn.assert_called_once()
            second = client.post('/api/report-tools/copq/runs', json={
                'period':'2026-09', 'mode':'combine', 'selections':[]})
            self.assertEqual(second.status_code, 409)
            spawn.assert_called_once()

    def test_reject_wrong_month_unknown_paths_and_duplicate_submission(self):
        self.source()
        item = reports.input_catalog('copq')[0]
        with self.assertRaisesRegex(ValueError, 'tháng'):
            reports.start_run('copq', '2026-08', 'combine', [{'id':item['id']}])
        with self.assertRaises(ValueError):
            reports.start_run('copq', '2026-09', 'combine', [{'id':'sfc:../../outside.xls'}])
        with patch.object(reports.subprocess, 'Popen'):
            reports.start_run('copq', '2026-09', 'combine', [{'id':item['id']}])
            with self.assertRaises(RuntimeError):
                reports.start_run('copq', '2026-09', 'combine', [{'id':item['id']}])

    def test_saved_report_does_not_restart_on_read_and_download_is_scoped(self):
        run_id = 'a' * 32
        folder = reports.run_folder(run_id)
        output = folder / 'workspace' / 'Output'
        output.mkdir(parents=True)
        (output / 'report.xlsx').write_bytes(b'completed fixture')
        reports.write_json(folder / 'run.json', dict(id=run_id,tool='copq',status='done',
            started_at=reports.now(),files=[dict(name='report.xlsx')]))
        with patch.object(reports.subprocess, 'Popen') as spawn, server.app.test_client() as client:
            self.assertEqual(client.get('/reports').status_code, 200)
            for _ in range(2):
                self.assertEqual(client.get('/api/report-tools/runs/' + run_id).json['status'], 'done')
            response = client.get(f'/api/report-tools/runs/{run_id}/files/report.xlsx')
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.data, b'completed fixture')
            response.close()
            self.assertEqual(client.get(f'/api/report-tools/runs/{run_id}/files/../../run.json').status_code, 404)
            self.assertEqual(client.get(f'/api/report-tools/runs/{run_id}/files/unknown.xlsx').status_code, 404)
            spawn.assert_not_called()

    def test_dead_worker_is_marked_failed_without_rerunning(self):
        run_id = 'b' * 32
        folder = reports.run_folder(run_id)
        folder.mkdir(parents=True)
        reports.write_json(folder / 'run.json', dict(id=run_id,status='running',pid=123,started_at=reports.now(),files=[]))
        with patch.object(reports, 'process_alive', return_value=False), patch.object(reports.subprocess, 'Popen') as spawn:
            self.assertEqual(reports.snapshot(run_id)['status'], 'error')
            spawn.assert_not_called()

    def test_sfc_raw_copq_catalog_excludes_summary_and_partial_files(self):
        for name in ('Cost Of Poor Quality 2026-09-01_2026-09-30-VH.xls',
                     'Cost Of Poor Quality Report 2026-09-01_2026-09-30-VH.xls',
                     'Cost Of Poor Quality 2026-09-01_2026-09-30-VH.partial.xls'):
            (self.exports / name).write_bytes(b'fixture')
        items = reports.input_catalog('copq')
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]['source'], 'sfc')


if __name__ == '__main__':
    unittest.main()
