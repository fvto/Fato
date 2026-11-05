"""Briefs explain existing evidence without guessing causes or changing the checks."""

from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import openpyxl
import sfc_report_tools as reports
from report_brief import copq_issues, hfpa_issues, summarize_checks
from report_exports import check_outputs


HEADERS = ["Site", "Type", "Missing Fields", "Validation Errors", "Data Quality Status",
           "Ttl cost ($)", "Working hours", "Defective Qty(Pair)"]


def fixture_brief():
    records = [(i + 2, ["JV", "B/C", "Factory", None, "Missing Data", 1, 0, 1]) for i in range(234)]
    records += [
        (722, ["JV", "Inbound&PI claim", None, "Ttl cost ($) > 1000", "Validation Error", 2552, 0, 332]),
        (766, ["JV", "Inbound&PI claim", None, "Ttl cost ($) > 1000", "Validation Error", 11537.15, 0, 639]),
        (1795, ["VH", "Reinspection", "Model", "Invalid style nbr => cannot link model", "Missing Data & Validation Error", 6.98, .5, 126]),
        (2160, ["VH2", "Reinspection", None, "Working hours > 5 & Defective Qty(Pair) < 300", "Validation Error", 54.31, 8, 95]),
        (2162, ["JV", "B/C", None, None, "OK", 1, 0, 1]),
        (2163, ["JV", "B/C", None, "Auto-resolved Model from Style Nbr: Fixture", "Auto-Resolved", 1, 0, 1]),
    ]
    issues = copq_issues(HEADERS, records, "COPQ_Clean.xlsx")
    checks = [dict(rule="copq.quality", status="warning", message="COPQ", mismatches=238, issues=issues)]
    checks += [dict(rule=rule, status="passed", message="Matched") for rule in ("copq.row_count", "copq.cost")]
    return summarize_checks(checks)


class ReportBriefTests(unittest.TestCase):
    def test_groups_exact_counts_sites_rows_and_numeric_examples(self):
        brief = fixture_brief()
        self.assertEqual(brief["flagged_records"], 238)
        self.assertEqual([item["count"] for item in brief["issues"]], [234, 2, 1, 1])
        self.assertEqual(sum(item["count"] for item in brief["issues"]), 238)
        missing, costs, model, hours = brief["issues"]
        self.assertEqual(missing["sites"], {"JV": 234})
        self.assertEqual(missing["types"], {"B/C": 234})
        self.assertTrue(missing["rows_truncated"])
        self.assertEqual(costs["rows"], [722, 766])
        self.assertEqual(costs["examples"][1]["values"]["Ttl cost ($)"], "11537.15")
        self.assertEqual(costs["severity"], "verify")
        self.assertIn("có thể hợp lệ", costs["action"])
        self.assertIn("chưa đủ căn cứ", model["action"])
        self.assertEqual(model["rows"], [1795])
        self.assertEqual(hours["examples"][0]["values"]["Working hours"], "8")
        self.assertEqual(hours["examples"][0]["values"]["Defective Qty(Pair)"], "95")
        self.assertEqual(len(brief["verified"]), 1)

    def test_hfpa_explains_unmatched_sources_and_different_count(self):
        header = ["FTY", "Validation_Status", "QA_IssueQty_Sum", "HFPA_Defect_Qty"]
        issues = hfpa_issues(header, [
            (2, ["VH", "HFPA_Only", None, 3]),
            (5, ["JV", "QAStation_Only", 5, None]),
            (9, ["VH2", "Different_Count", 5, 3]),
            (10, ["VH2", "Matched_Exact", 2, 2]),
        ], "master.xlsx")
        self.assertEqual(len(issues), 3)
        self.assertEqual([item["rows"] for item in issues], [[2], [5], [9]])
        self.assertIn("MES410", issues[0]["title"])
        self.assertIn("QAStation", issues[1]["title"])
        self.assertEqual(issues[2]["sites"], {"VH2": 1})
        self.assertEqual(issues[2]["examples"][0]["values"]["HFPA_Defect_Qty"], "3")

    def test_unknown_rule_does_not_expose_untranslated_notes_or_invent_cause(self):
        issues = copq_issues(HEADERS, [(2, ["JV", "B/C", None, "SECRET unexpected detail", "Validation Error"])], "result.xlsx")
        self.assertNotIn("SECRET", str(issues))
        self.assertIn("chưa có diễn giải riêng", issues[0]["reason"])
        self.assertIn("Validation Errors", issues[0]["action"])

    def test_reconciliation_warning_prevents_claim_that_totals_match(self):
        checks = [
            dict(rule="copq.row_count", status="passed", message="Count matched"),
            dict(rule="copq.cost", status="warning", message="Chi phí chưa khớp", expected="12", actual="13"),
            dict(rule="engine.warnings", status="warning", message="Database cần xem lại"),
        ]
        brief = summarize_checks(checks)
        self.assertEqual(brief["verified"], [])
        self.assertEqual(len(brief["issues"]), 2)
        self.assertIn("12", brief["issues"][0]["reason"])
        self.assertEqual(brief["issues"][1]["title"], "Database cần xem lại")
        self.assertEqual(summarize_checks([])["status"], "passed")

    def test_export_uses_actual_excel_rows_and_adds_brief_without_mutating_file(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            output = folder / "workspace/Output"
            output.mkdir(parents=True)
            path = output / "result.xlsx"
            book = openpyxl.Workbook()
            sheet = book.active
            sheet.title = "COPQ_Clean"
            sheet.append(HEADERS)
            sheet.append(["JV", "B/C", "Factory", None, "Missing Data", 12, 0, 1])
            sheet.append([None] * len(HEADERS))
            sheet.append(["JV", "Claim", None, "Ttl cost ($) > 1000", "Validation Error", 2000, 0, 1])
            book.save(path)
            book.close()
            before = path.read_bytes()
            files = reports.collect_files(folder)
            review = check_outputs(folder, dict(tool="copq", mode="combine", validation=[dict(group="copq", rows=2, totals={"Ttl cost ($)": 2012})]), files)
            self.assertEqual(review["brief"]["flagged_records"], 2)
            self.assertEqual([item["rows"] for item in review["brief"]["issues"]], [[2], [4]])
            self.assertEqual(path.read_bytes(), before)

    def test_legacy_snapshot_brief_does_not_rescan_output_workbooks(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            folder = root / ("a" * 32)
            folder.mkdir()
            run = dict(status="done", started_at=reports.now(), export_review=dict(checks=[dict(rule="copq.quality", status="warning", message="COPQ", mismatches=238)]))
            reports.write_json(folder / "run.json", run)
            with patch.object(reports, "RUN_ROOT", root), patch.object(openpyxl, "load_workbook", side_effect=AssertionError("unneeded read")):
                snapshot = reports.snapshot(folder.name)
            self.assertEqual(snapshot["export_review"]["brief"]["flagged_records"], 238)
            self.assertIn("Lượt cũ", snapshot["export_review"]["brief"]["issues"][0]["reason"])


if __name__ == "__main__":
    unittest.main()
