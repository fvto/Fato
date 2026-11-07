from pathlib import Path
import json
import subprocess
import unittest

from report_runtime import AcquisitionError, friendly_error


class StationSelectionTests(unittest.TestCase):
    def test_missing_station_control_allows_only_explicit_no_selection(self):
        source = (Path(__file__).resolve().parents[1] / 'sfc_controller.py').read_text(encoding='utf-8')
        script = source.split('page.evaluate("""({stations, selection}) => {', 1)[1].split('}""",', 1)[0]
        function = '({stations, selection}) => {' + script + '}'
        for selection, stations, allowed in [('none', [], True), ('explicit', ['HFPA'], False), (None, [], False)]:
            with self.subTest(selection=selection):
                program = 'global.aspxGetControlCollection=()=>({Get:()=>null}); const select=' + function
                program += '; select(' + json.dumps(dict(selection=selection, stations=stations)) + ');'
                result = subprocess.run(['node', '-e', program], capture_output=True, text=True)
                self.assertEqual(result.returncode == 0, allowed)

    def test_acquisition_summary_keeps_safe_slot_detail(self):
        message = 'qastation/JV: ' + friendly_error(RuntimeError('Page.evaluate: Không tìm thấy danh sách trạm SFC'))
        self.assertIn('HFPA', message)
        self.assertNotIn('Page.evaluate', message)
        self.assertEqual(friendly_error(AcquisitionError(message)), message)
