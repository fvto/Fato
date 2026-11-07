from collections import Counter, defaultdict
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from report_tools.quarterly_quality import sfc_inputs
from report_tools.quarterly_quality.quarterly_quality import Counts, PLANTS
from report_tools.quarterly_quality.template_output import WorksheetPatch, X


class QuarterlyQualitySourcesTests(unittest.TestCase):
    def test_unavailable_qastation_uses_mes_counts_and_keeps_other_factories(self):
        inputs = []
        for plant in PLANTS:
            for group in ('hfpa', 'ftt', 'qastation'):
                if (plant, group) != ('JV', 'qastation'):
                    inputs.append(dict(group=group, plant=plant, date_from='2026-09-01',
                                       date_to='2026-09-30', name=f'{group}-{plant}.xlsx'))
        missing = dict(group='qastation', plant='JV', date_from='2026-09-01', date_to='2026-09-30')
        run = dict(period='FY27-Q2', inputs=inputs, acquisition=dict(
            date_from='2026-09-01', date_to='2026-09-30', unavailable=[missing]))
        with patch.object(sfc_inputs, 'source_counts', return_value=(Counts(100, 10), 1)), \
                patch.object(sfc_inputs, 'corrected_hfpa_counts', return_value=(Counts(80, 8), 1, 1, [])) as reconcile:
            quarter = sfc_inputs.load_sfc_quarter(Path('.'), run, lambda *_: None)
        self.assertEqual(reconcile.call_count, 3)
        self.assertEqual(quarter.values['HFPA', '2026-09', 'JV'], Counts(100, 10))
        self.assertEqual(quarter.totals('HFPA', 'CLG'), Counts(340, 34))
        # An unexplained missing file must still fail rather than silently bypass validation.
        run['acquisition']['unavailable'] = []
        with self.assertRaises(sfc_inputs.SourceRowError):
            sfc_inputs.load_sfc_quarter(Path('.'), run, lambda *_: None)
        run['inputs'].append(dict(missing, name='empty-qa-JV.xlsx'))
        def reconcile_empty(path, item, *args, **kwargs):
            if item['plant'] == 'JV':
                raise sfc_inputs.EmptyQAStation('QAStation trống')
            return Counts(80, 8), 1, 1, []
        with patch.object(sfc_inputs, 'source_counts', return_value=(Counts(100, 10), 1)), \
                patch.object(sfc_inputs, 'corrected_hfpa_counts', side_effect=reconcile_empty):
            quarter = sfc_inputs.load_sfc_quarter(Path('.'), run, lambda *_: None)
        self.assertEqual(quarter.totals('HFPA', 'CLG'), Counts(340, 34))
        self.assertEqual(run['acquisition']['unavailable'][0]['plant'], 'JV')

    def test_cached_ftt_refresh_replaces_ranking_and_totals_without_touching_hfpa(self):
        from openpyxl import Workbook
        from report_tools.quarterly_quality import tableau_source
        with TemporaryDirectory() as temporary:
            path = Path(temporary)/'tableau.xlsx'
            book = Workbook()
            sheet = book.active
            sheet.title = tableau_source.SHEET
            sheet.append(tableau_source.HEADER)
            for plant in PLANTS:
                sheet.append(['2026-09-01', plant, 'A Line 1', 'Old FTT Model 2', 100, 10, 'Assembly', 'Cleanness', 10])
                sheet.append(['2026-09-01', plant, 'A Line 1', 'HFPA Model 2 keep', 50, 5, 'HFPA', 'Rocking', 5])
            book.save(path)
            book.close()
            original = path.read_bytes()
            inputs = [dict(group='ftt', plant=plant, date_from='2026-09-01', date_to='2026-09-30',
                           replace_ftt=True, name=f'{plant}.xlsx') for plant in PLANTS]
            run = dict(period='FY27-Q2', inputs=inputs, acquisition=dict(date_from='2026-09-01', date_to='2026-09-30',
                       tableau=dict(path=str(path), sha256=tableau_source.fingerprint(path))))
            header = (*sfc_inputs.FTT_LOT, 'Total Pair Produced', 'Station', 'Metric', 'Number of pairs',
                      'Defect Issues', "Issues Q'ty")
            def export(source):
                lot = ('2026-09-01', 'CLG', source.stem, 'A', '1', 'Correct Model', 'Wrong Model 2',
                       'SO', 'SP', 'M', 'S', 'C')
                return header, [(*lot, 100, 'Assembly', 'AGrade', 100, '', 0),
                                (*lot, 100, 'Assembly', 'Rework1', 10, 'Cleanness', 10)]
            with patch.object(sfc_inputs, 'read_export', side_effect=export), \
                    patch.object(tableau_source, 'append_records') as save:
                quarter = sfc_inputs.load_sfc_quarter(Path(temporary), run, lambda *_: None)
            self.assertEqual(quarter.values['FTT', '2026-09', 'VH'], Counts(100, 10, 100, 10))
            self.assertEqual(quarter.rankings['FTT', 'VN'][0], [('Correct Model', 20, 200)])
            self.assertEqual(quarter.rankings['FTT', 'VN'][1], [('Cleanness', 20, 20)])
            self.assertEqual(quarter.rankings['HFPA', 'VN'][0], [('HFPA Model 2 keep', 10, 100)])
            self.assertEqual(quarter.rankings['HFPA', 'VN'][1], [('Rocking', 10, 10)])
            self.assertTrue(all(row[3] == 'Correct Model' for row in save.call_args.args[1]))
            self.assertEqual(path.read_bytes(), original)

    def test_ftt_model_totals_and_issues_use_model_column_without_model2_fallback(self):
        header = (*sfc_inputs.FTT_LOT, "Total Pair Produced", "Station", "Metric", "Number of pairs",
                  "Defect Issues", "Issues Q'ty")
        rows = []
        for model, model2, produced, failed, issues in (
            ("Model A", "Model 2 X", 100, 10, (("Cleanness", 6), ("Over cement", 4))),
            ("Model A", "Model 2 Y", 200, 20, (("Over cement", 20),)),
            ("Model B", "Model 2 X", 100, 5, (("Rocking", 5),)),
            (None, "Do not substitute", 50, 2, (("Rocking", 2),)),
        ):
            lot = ("2026-11-04", "CLG", "VH", "A", "1", model, model2,
                   f"SO-{model}-{model2}", "SP", "M", "S", "C")
            rows.append((*lot, produced, "Assembly", "AGrade", produced, "", 0))
            for issue, qty in issues:
                rows.append((*lot, produced, "Assembly", "Rework1", failed, issue, qty))
        details = (Counter(), defaultdict(Counter))
        item = dict(group="ftt", plant="VH", date_from="2026-11-01", date_to="2026-11-05")
        counts, count = sfc_inputs.source_counts(Path("unused.xlsx"), item, details, export=(header, rows))
        self.assertEqual((counts, count), (Counts(450, 37, 450, 37), len(rows)))
        self.assertEqual(details[0], {"Model A": 30, "Model B": 5})
        self.assertEqual(details[1]["Model A"], {"Over cement": 24, "Cleanness": 6})
        self.assertEqual(details[1]["Model B"], {"Rocking": 5})
        self.assertNotIn("Do not substitute", details[1])

    def test_updating_shared_formula_anchor_preserves_the_other_cell_formula(self):
        xml = f'''<worksheet xmlns="{X[1:-1]}"><sheetData><row r="23">
        <c r="Y23"><f t="shared" si="1" ref="Y23:Z23">1-Y14/Y5</f><v>0.9</v></c>
        <c r="Z23" t="e"><f t="shared" si="1"/><v>#DIV/0!</v></c>
        </row></sheetData></worksheet>'''.encode()
        sheet = WorksheetPatch(xml)
        sheet.set(23, 25, 0.9, 'IFERROR(1-Y14/Y5,"")')
        self.assertEqual(sheet.formula_at("Z23"), "1-Z14/Z5")
        self.assertFalse(sheet.cells["Z23"].find(X + "f").attrib)
        sheet.set(23, 26, None, f'IFERROR({sheet.formula_at("Z23")},"")')
        self.assertEqual(sheet.cells["Z23"].get("t"), "str")

    def test_hfpa_details_use_only_the_days_matching_the_approved_workbook(self):
        header = (*sfc_inputs.HFPA_FIELDS, "SO#", "Country", "Pass/Fail", "Model 2",
                  "Defect Q'ty (Total Defect Q'ty)", "Cleanness")
        rows = [
            ("01-Oct-2026", "CL_VH", "1", 100, 8, "SO1", "VN", "Pass", "First", 8, 8),
            ("02-Oct-2026", "CL_VH", "2", 50, 5, "SO2", "VN", "Pass", "Later", 5, 5),
        ]
        slot = dict(group="hfpa", plant="VH", date_from="2026-10-01", date_to="2026-10-31")
        details, matched = (Counter(), defaultdict(Counter)), []
        with patch.object(sfc_inputs, "read_export", return_value=(header, rows)):
            total, count = sfc_inputs.source_counts(Path("unused.xlsx"), slot, details,
                                                   approved=Counts(100, 8), matched=matched)
        self.assertEqual((total, count), (Counts(150, 13), 2))
        self.assertEqual(matched, ["2026-10-01"])
        self.assertEqual(details[0], {"First": 8})
        self.assertEqual(details[1], {"First": {"Cleanness": 8}})

    def test_sfc_uses_current_quarter_counts_and_ignores_prior_quarter_inputs(self):
        items = [
            dict(group=group, plant=plant, date_from=f"2026-{month}-01",
                 date_to=f"2026-{month}-30", name=f"{group}-{plant}-{month}.xlsx")
            for month in ("06", "09")
            for group in ("hfpa", "qastation", "ftt")
            for plant in PLANTS
        ]
        refreshed = ("FTT", "2026-09", "VH")
        corrected = ("HFPA", "2026-09", "VH")
        missing = ("HFPA", "2026-09", "VH2")
        run = dict(period="FY27-Q2", inputs=items, acquisition=dict(
            date_from="2026-09-01", date_to="2026-09-30"))
        with (patch.object(sfc_inputs, "source_counts", return_value=(Counts(150, 13), 2)) as ftt,
              patch.object(sfc_inputs, "corrected_hfpa_counts", return_value=(Counts(151, 14), 2, 2, [])) as hfpa):
            quarter = sfc_inputs.load_sfc_quarter(Path("unused"), run, lambda *_: None,
                                                  baseline=SimpleNamespace(values={
                                                      refreshed: Counts(100, 8), corrected: Counts(100, 8)}))
        self.assertEqual(quarter.protected, set())
        self.assertEqual(quarter.values[refreshed], Counts(150, 13))
        self.assertEqual(quarter.values[missing], Counts(151, 14))
        self.assertEqual(quarter.values[corrected], Counts(151, 14))
        self.assertEqual(quarter.totals("HFPA", "VN"), Counts(302, 28))
        self.assertEqual(quarter.unmatched_rankings, set())
        self.assertEqual(run["hfpa_baseline_differences"][0]["corrected_insp"], 151)
        self.assertEqual(run["ftt_baseline_differences"][0]["corrected_insp"], 150)
        self.assertEqual(set(key[1] for key in quarter.values), {"2026-09"})
        self.assertEqual((ftt.call_count, hfpa.call_count), (4, 4))
        self.assertTrue(all(call.args[1]["date_from"] == "2026-09-01" for call in (*ftt.call_args_list, *hfpa.call_args_list)))
        self.assertIsNone(quarter.prior_date_to)

    def test_quarter_corrects_hfpa_with_the_monthly_rules(self):
        mes = ((*sfc_inputs.HFPA_FIELDS, "SO#", "Country", "Pass/Fail", "Plant", "Line", "Model 2",
                "Defect Q'ty (Total Defect Q'ty)", "Rocking", "Wrinkle bottom"),
               [("24-Oct-2026", "CL_VH", "A", 248, 38, "SO", "VN", "Pass", "S2", 4, "Model", 38, 30, 8)])
        qa = (("InspDate", "Line", "ShoeName", "InspQty", "FailQty", "Station", "Issues", "IssueQty"),
              [("2026-10-24", "S2 Line 4", "Model", 260, 40, "HFPA", "Rocking", 40)])
        slot = dict(group="hfpa", plant="VH", date_from="2026-10-01", date_to="2026-10-31")
        details = (Counter(), defaultdict(Counter))
        with patch.object(sfc_inputs, "read_export", side_effect=[mes, qa]):
            total, mes_rows, qa_rows, audit = sfc_inputs.corrected_hfpa_counts(
                Path("mes.xlsx"), slot, Path("qa.xlsx"), details
            )
        self.assertEqual((total, mes_rows, qa_rows), (Counts(248, 38), 1, 1))
        self.assertEqual(details[0], {"Model": 38})
        self.assertEqual(details[1], {"Model": {"Rocking": 30, "Wrinkle bottom": 8}})
        self.assertEqual(audit[0]["Status"], "MES410_Corrected")

    def test_unnamed_all_pass_inspection_requires_matching_mes410_evidence(self):
        qa = (("InspDate", "Line", "ShoeName", "InspQty", "FailQty", "Station", "Issues", "IssueQty"),
              [("2026-11-04", "A Line2", "", 16, 0, "HFPA", "", 0)])
        header = (*sfc_inputs.HFPA_FIELDS, "SO#", "Country", "Pass/Fail", "Plant", "Line", "Model 2",
                  "Defect Q'ty (Total Defect Q'ty)", "Rocking")
        slot = dict(group="hfpa", plant="JV2", date_from="2026-11-01", date_to="2026-11-04")
        for model, sample, issues, valid in (("", 16, 0, True), ("", 17, 0, False),
                                           ("Other model", 16, 0, False), ("", 16, 1, False)):
            with self.subTest(model=model, sample=sample, issues=issues):
                mes = (header, [("04-Nov-2026", "CL_JV2", "audit-1", sample, 0,
                                 "SO", "ID", "Pass", "A", "2", model, issues, issues)])
                with patch.object(sfc_inputs, "read_export", side_effect=[mes, qa]):
                    if valid:
                        total, _, _, audit = sfc_inputs.corrected_hfpa_counts(Path("mes.xlsx"), slot, Path("qa.xlsx"), None)
                        self.assertEqual(total, Counts(16, 0))
                        self.assertEqual(audit[0]["Reason"], "MES410_Blank_Model_Matched")
                    else:
                        with self.assertRaisesRegex(sfc_inputs.SourceRowError, "model \\(trống\\)"):
                            sfc_inputs.corrected_hfpa_counts(Path("mes.xlsx"), slot, Path("qa.xlsx"), None)


if __name__ == "__main__":
    unittest.main()
