"""HFPA reference colors stay independent of other report palettes."""
from pathlib import Path
import tempfile
import shutil
import unittest

import openpyxl
import pandas as pd

from report_tools.hfpa.agents.color_manager import ColorManager
from report_tools.hfpa.hfpa_common import compute_top3_defects

REFERENCE = Path(__file__).resolve().parents[1] / 'report_tools/hfpa/reference/HFPA_Defect_Colors_06042026.xlsx'


class HFPAPaletteTests(unittest.TestCase):
    def test_reference_colors_and_aliases(self):
        colors = ColorManager(str(REFERENCE)).load_color_map()
        book = openpyxl.load_workbook(REFERENCE, read_only=True, data_only=True)
        try:
            rows = [row for row in list(book['Sheet1'].values)[4:] if row[2]]
            self.assertEqual(len(rows), 55)
            for row in rows:
                self.assertEqual(colors[ColorManager.normalize_name(row[2])], str(row[7]).lstrip('#').upper())
        finally:
            book.close()
        self.assertEqual(colors['lacing'], '316CE5')
        self.assertEqual(colors['hairy edge'], 'C4BD97')
        self.assertEqual(colors['colormigration'], '57A67B')
        self.assertEqual(colors['bmaterialfailure'], 'B46AC8')

    def test_new_defects_leave_original_reference_intact(self):
        original = REFERENCE.read_bytes()
        with tempfile.TemporaryDirectory() as directory:
            active = Path(directory) / 'Color_template.xlsx'
            shutil.copyfile(REFERENCE, active)
            manager = ColorManager(str(active))
            manager.sync_unmapped_defects(['Lacing', 'Brand new test defect'])
            self.assertEqual(manager.load_color_map()['lacing'], '316CE5')
            self.assertIn('brandnewtestdefect', manager.load_color_map())
        self.assertEqual(REFERENCE.read_bytes(), original)

    def test_descending_top3_preserves_source_ties_and_reference_colors(self):
        qa = pd.DataFrame({'FTY': ['VH'] * 5, 'ShoeName_Std': ['Model'] * 5,
                           'Issues_Std': ['Lacing', 'Yellowing', 'Hairy edge', 'Cleanness', 'Yellowing'],
                           'IssueQty_Validated': [2, 4, 4, 1, 1]})
        top = {'VH': pd.DataFrame({'Shoename': ['Model']})}
        result = compute_top3_defects(qa, top, ColorManager(str(REFERENCE)).load_color_map())
        self.assertEqual(result['Defect Type'].tolist(), ['Yellowing', 'Hairy edge', 'Lacing'])
        self.assertEqual(result['Defect Quantity'].tolist(), [5, 4, 2])
        self.assertEqual(result['Color HEX'].tolist(), ['EEDCB4', 'C4BD97', '316CE5'])
        qa.loc[0, 'IssueQty_Validated'] = 4
        result = compute_top3_defects(qa, top, ColorManager(str(REFERENCE)).load_color_map())
        self.assertEqual(result['Defect Type'].tolist(), ['Yellowing', 'Lacing', 'Hairy edge'])


if __name__ == '__main__':
    unittest.main()
