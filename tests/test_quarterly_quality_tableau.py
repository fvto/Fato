from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook

from report_tools.quarterly_quality import tableau_source as source
from report_tools.quarterly_quality.quarterly_quality import Counts
from report_tools.quarterly_quality.sfc_inputs import FTT_LOT
import sfc_quarterly_quality_acquisition as acquisition


class TableauSourceTests(unittest.TestCase):
    def fixture(self, root):
        path = root / 'source.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.title = 'VH FTT'
        sheet.append(source.HEADER)
        sheet.append([datetime(2026, 9, 1), 'VH', 'A Line 1', 'Old', 100, 10, 'HFPA', 'Cleanness', 6])
        sheet.append([datetime(2026, 9, 1), 'VH', 'A Line 1', 'Old', 100, 10, 'HFPA', 'Over cement', 4])
        sheet.append([datetime(2026, 9, 1), 'VH', 'A Line 1', 'Chart', 100000, 99999, 'Assembly', 'B/', 0])
        sheet.append(['2026/10/01', 'VH', 'A Line 1', 'Existing', 200, 20, 'HFPA', 'Cleanness', 20])
        sheet.append([datetime(2026, 9, 1), 'VH', 'A Line 1', 'Chart', 100000, 99999, 'Assembly', 'C/', 0])
        book.create_sheet('Bond test - Production')['A1'] = 'Preserve this sheet'
        book.create_sheet('Dates')['A1'] = '=TODAY()'
        book.save(path)
        return path

    def test_inventory_handles_text_dates_repeated_pair_counts_and_chart_rows(self):
        with TemporaryDirectory() as temporary:
            path = self.fixture(Path(temporary))
            data = source.read_source(path, ('2026-09', '2026-10', '2026-11'), '2026-11-04')
            self.assertEqual(data.values['HFPA', '2026-09', 'VH'], Counts(100, 10))
            self.assertEqual(data.values['HFPA', '2026-10', 'VH'], Counts(200, 20))
            self.assertNotIn(('FTT', '2026-09', 'VH'), data.values)
            self.assertEqual(data.covered_to('HFPA', '2026-09', 'VH', '2026-09-30'), '2026-09-30')

    def test_append_preserves_other_parts_and_existing_values_without_moving_chart_rows_to_end(self):
        with TemporaryDirectory() as temporary, patch.object(source, 'APP_ROOT', Path(temporary)):
            path = self.fixture(Path(temporary))
            data = source.read_source(path, ('2026-09', '2026-10', '2026-11'), '2026-11-04')
            with ZipFile(path) as package:
                untouched = {name: package.read(name) for name in package.namelist() if name != data.part}
            added = [['2026-11-04', 'VH', 'A Line 2', 'New', 50, 2, 'HFPA', 'Cleanness', 2]]
            windows = [dict(group='hfpa', plant='VH', date_from='2026-11-01', date_to='2026-11-04')]
            result = source.append_records(data, added, windows, lambda *_: None)
            self.assertEqual(result['added_rows'], 1)
            with ZipFile(path) as package:
                self.assertTrue(all(package.read(name) == value for name, value in untouched.items()))
            book = load_workbook(path, read_only=True, data_only=True)
            rows = list(book['VH FTT'].values)
            self.assertEqual([row[3] for row in rows[1:]], ['Old', 'Old', 'Chart', 'Existing', 'Chart', 'New'])
            self.assertEqual(rows[-1][3], 'New')
            self.assertEqual(rows[1][4:6], (100, 10))
            self.assertEqual(rows[6][4:6], (50, 2))
            book.close()
            cached = source.read_source(path, ('2026-11',), '2026-11-05')
            self.assertEqual(cached.covered_to('HFPA', '2026-11', 'VH', '2026-11-05'), '2026-11-04')
            with patch.object(source, 'default_path', return_value=path), \
                    patch.object(acquisition, 'local_today', return_value=datetime(2026, 11, 5).date()):
                next_plan = acquisition.plan('FY27-Q2', '2026-11-05')
                same_plan = acquisition.plan('FY27-Q2', '2026-11-04')
            self.assertEqual(next(slot['date_from'] for slot in next_plan['missing']
                                  if slot['plant'] == 'VH' and slot['group'] == 'hfpa'
                                  and slot['date_to'] == '2026-11-05'), '2026-11-05')
            self.assertFalse(any(slot['plant'] == 'VH' and slot['group'] == 'hfpa'
                                 and slot['date_to'] == '2026-11-04' for slot in same_plan['missing']))
            before = source.fingerprint(path)
            source.append_records(cached, [], [], lambda *_: None)
            self.assertEqual(source.fingerprint(path), before)

    def test_source_changed_after_preflight_is_not_overwritten(self):
        with TemporaryDirectory() as temporary:
            path = self.fixture(Path(temporary))
            data = source.read_source(path, ('2026-11',), '2026-11-04')
            changed = path.read_bytes() + b' '
            path.write_bytes(changed)
            with self.assertRaisesRegex(ValueError, 'đã thay đổi'):
                source.append_records(data, [['2026-11-04', 'VH', 'A Line 1', 'New', 1, 0, 'HFPA', '', 0]], [], lambda *_: None)
            self.assertEqual(path.read_bytes(), changed)

    def test_saved_ftt_grade_and_rework_survive_reusing_the_nine_column_data(self):
        with TemporaryDirectory() as temporary, patch.object(source, 'APP_ROOT', Path(temporary)):
            path = self.fixture(Path(temporary))
            data = source.read_source(path, ('2026-11',), '2026-11-04')
            counts = Counts(50, 10, grade=48, rework=8)
            rows = [['2026-11-04', 'VH', 'A Line 1', 'Model', 50, 10, 'Assembly', 'Cleanness', 10]]
            windows = [dict(group='ftt', plant='VH', date_from='2026-11-01', date_to='2026-11-04')]
            source.append_records(data, rows, windows, lambda *_: None, metrics={('FTT', '2026-11', 'VH'): counts})
            cached = source.read_source(path, ('2026-11',), '2026-11-04')
            self.assertEqual(cached.values['FTT', '2026-11', 'VH'], counts)

    def test_plan_skips_saved_data_and_fetches_only_the_uncovered_window(self):
        with TemporaryDirectory() as temporary:
            path = self.fixture(Path(temporary))
            with patch.object(source, 'default_path', return_value=path), \
                    patch.object(acquisition, 'local_today', return_value=datetime(2026, 11, 5).date()):
                plan = acquisition.plan('FY27-Q2', '2026-11-04')
            slots = plan['missing']
            self.assertFalse(any(slot['plant'] == 'VH' and slot['group'] in ('hfpa', 'qastation')
                                 and slot['date_from'][:7] in ('2026-09', '2026-10') for slot in slots))
            self.assertTrue(any(slot['group'] == 'ftt' and slot['date_from'] == '2026-09-01' for slot in slots))

    def test_ftt_normalization_counts_lots_and_metrics_once_across_issue_rows(self):
        header = (*FTT_LOT, 'Total Pair Produced', 'Station', 'Metric', 'Number of pairs', 'Defect Issues', "Issues Q'ty")
        lot = ('2026-11-04', 'CLG', 'VH', 'A', '1', 'Original', 'Model', 'SO1', 'SP', 'M', 'S', 'C')
        rows = [(*lot, 100, 'Assembly', 'AGrade', 100, '', 0),
                (*lot, 100, 'Assembly', 'Rework1', 10, 'Cleanness', 6),
                (*lot, 100, 'Assembly', 'Rework1', 10, 'Over cement', 4)]
        result = source.normalize_ftt((header, rows), dict(plant='VH'))
        self.assertEqual(len(result), 2)
        self.assertTrue(all(row[4:7] == [100, 10, 'Assembly'] for row in result))
        self.assertEqual(sum(row[8] for row in result), 10)
        self.assertTrue(all(row[2] == 'A Line 1' for row in result))
        named_line = list(lot)
        named_line[4] = 'Line 1'
        repeated = [(*named_line, *row[len(FTT_LOT):]) for row in rows]
        self.assertTrue(all(row[2] == 'A Line 1' for row in source.normalize_ftt((header, repeated), dict(plant='VH'))))


if __name__ == '__main__':
    unittest.main()
