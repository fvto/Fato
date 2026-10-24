from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import dashboard_paths as paths
import dashboard_launcher as launcher
import sfc_controller as controller


class DashboardIsolationTests(unittest.TestCase):
    def test_legacy_path_is_relative_to_app_folder_not_current_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            app = root / 'Chingluh Dashboard'
            legacy = root / 'MQAA-Record'
            app.mkdir(); legacy.mkdir()
            (app / 'dashboard_config.json').write_text('{"legacy_mqaa_root":"../MQAA-Record"}', encoding='utf8')
            master = legacy / 'MQAA_Master_DoiChieu_21_Forms_MoiNhat.xlsx'
            master.write_bytes(b'path test only')
            with patch.object(paths, 'APP_ROOT', app), patch.dict(paths.os.environ, {'CHINGLUH_MQAA_ROOT': ''}):
                self.assertEqual(paths.mqaa_workspace(), legacy.resolve())
                self.assertEqual(paths.mqaa_master_file(), master)
                with patch.object(controller.os, 'startfile') as open_file:
                    self.assertTrue(controller.open_master_excel()['success'])
                    open_file.assert_called_once_with(str(master))

    def test_mqaa_scan_uses_legacy_script_and_working_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            legacy = Path(directory)
            script = legacy / 'auto_scan_all_21_models.py'
            script.write_text('raise RuntimeError("must not execute in a unit test")', encoding='utf8')
            process = SimpleNamespace(stdout=iter(['scan output\n']), wait=lambda: 0)
            def immediate_thread(**kwargs):
                return SimpleNamespace(start=kwargs['target'])
            with patch.object(controller, 'mqaa_workspace', return_value=legacy), \
                    patch.object(controller, 'MQAA_STATUS', {'running': False, 'progress': '', 'last_run': None}), \
                    patch.object(controller.threading, 'Thread', side_effect=immediate_thread), \
                    patch.object(controller.subprocess, 'Popen', return_value=process) as spawn:
                self.assertEqual(controller.start_mqaa_scan_background()['status'], 'started')
                self.assertEqual(spawn.call_args.args[0][-1], str(script))
                self.assertEqual(spawn.call_args.kwargs['cwd'], str(legacy))

    def test_launcher_rejects_a_dashboard_from_another_folder(self):
        health = {'service': 'sfc-dashboard', 'ready': True, 'app_root': str(launcher.ROOT)}
        with patch.object(launcher, 'read_json', return_value=health):
            self.assertTrue(launcher.dashboard_ready())
            health['app_root'] = str(launcher.ROOT.parent / 'Other project')
            self.assertFalse(launcher.dashboard_ready())


if __name__ == '__main__':
    unittest.main()
