"""Re-Inspection must cover failed HFPA lots, and MA must cover QC."""

import io
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

import openpyxl
from werkzeug.datastructures import FileStorage

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import report_validation as validation
import report_tool_worker as worker
import sfc_report_tools as reports


def source(group, plant, *, qc=1, ma=1, fail=1, passed=0):
    labels = validation.CONTRACTS["schemas"][group]["required"]
    defaults = {
        "Date": "2026-09-01", "Audit Date (dd-mmm-yyyy)": "2026-09-01",
        "Plant": "A", "Line": "A01", "Model": "Fixture", "Model 2": "Fixture",
        "SO#": "SO", "Country": "VN", "Pass/Fail": "Fail",
        "Audit Sample Size": 100, "Defect Q'ty (Total Defect Q'ty)": 50,
        "Category": "QC", "Reinspection Q'ty": 10000, "2nd Audit Sample Size": 10,
        "2nd Pass/Fail": "Pass", "Remark": "", "Metric": "FTT",
        "Defect Issues": "Defect", "Issues Q'ty": 2,
    }
    if group == "reinspection":
        variants = [dict(Category="QC")] * qc + [dict(Category="MA")] * ma
    elif group == "hfpa":
        variants = [{"Pass/Fail": "Fail"}] * fail + [{"Pass/Fail": "Pass"}] * passed
    else:
        variants = [{}]
    book = openpyxl.Workbook()
    book.active.append(labels)
    for variant in variants:
        row = {**defaults, **variant}
        book.active.append([row[label] for label in labels])
    stream = io.BytesIO()
    book.save(stream)
    book.close()
    stream.seek(0)
    return FileStorage(stream=stream, filename=f"{group}_2026-09-{plant}.xlsx",
                       content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


class ReinspectionCoverageTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        engines = self.root / "engines"
        (engines / "reinspection").mkdir(parents=True)
        for item in [patch.object(reports, "TOOLS_ROOT", engines),
                     patch.object(reports, "RUN_ROOT", self.root / "runs")]:
            item.start()
            self.addCleanup(item.stop)

    def uploads(self, **counts):
        return [(group, source(group, plant, **counts))
                for group in ("reinspection", "ftt", "hfpa")
                for plant in reports.COMMON_PLANTS]

    def test_equal_greater_and_zero_counts_pass(self):
        for counts in [dict(qc=2, ma=2, fail=2), dict(qc=3, ma=4, fail=2),
                       dict(qc=0, ma=0, fail=0, passed=2)]:
            with self.subTest(counts=counts):
                result = reports.start_run("reinspection", "2026-09", "database", [],
                                           self.uploads(**counts), validate_only=True)
                self.assertTrue(result["valid"])

    def test_each_inequality_blocks_preflight_with_counts(self):
        for counts, rules in [
            (dict(qc=1, ma=2, fail=2), {"qc_hfpa"}),
            (dict(qc=2, ma=1, fail=2), {"ma_hfpa", "ma_qc"}),
            (dict(qc=3, ma=2, fail=1), {"ma_qc"}),
            (dict(qc=0, ma=1, fail=1), {"qc_hfpa"}),
            (dict(qc=1, ma=0, fail=1), {"ma_hfpa", "ma_qc"}),
        ]:
            with self.subTest(counts=counts), self.assertRaises(validation.PreflightError) as caught:
                reports.start_run("reinspection", "2026-09", "database", [],
                                  self.uploads(**counts), validate_only=True)
            errors = [r for r in caught.exception.reports if not r["valid"]]
            self.assertEqual(len(errors), 4)
            for report in errors:
                self.assertEqual({i["rule"].rsplit(".", 1)[1] for i in report["issues"]}, rules)
                self.assertTrue(all("thiếu" in i["message"] for i in report["issues"]))

    def test_failed_audits_count_rows_excluding_passes_and_footer(self):
        upload = source("hfpa", "VH", fail=2, passed=3)
        path = self.root / upload.filename
        path.write_bytes(upload.stream.getvalue())
        book = openpyxl.load_workbook(path)
        labels = [cell.value for cell in book.active[1]]
        footer = [None] * len(labels)
        footer[labels.index("Audit Date (dd-mmm-yyyy)")] = 5
        book.active.append(footer)
        book.save(path)
        book.close()
        result = validation.validate_file(path, "reinspection", "hfpa", "2026-09")
        self.assertTrue(result["valid"], result)
        self.assertEqual(result["rows"], 5)
        self.assertEqual(result["hfpa_fail_count"], 2)
        self.assertEqual(len(result["excluded_summary_rows"]), 1)

    def test_other_plants_cannot_compensate_for_shortfall(self):
        uploads = self.uploads(qc=5, ma=5, fail=1)
        uploads = [(group, source(group, "VH", qc=1, ma=1, fail=2)
                    if reports.plant_of(upload.filename) == "VH" else upload)
                   for group, upload in uploads]
        with self.assertRaises(validation.PreflightError) as caught:
            reports.start_run("reinspection", "2026-09", "database", [], uploads, validate_only=True)
        errors = [r for r in caught.exception.reports if not r["valid"]]
        self.assertEqual(len(errors), 1)
        self.assertEqual(reports.plant_of(errors[0]["name"]), "VH")

    def test_cached_file_summary_does_not_cache_cross_file_result(self):
        with self.assertRaises(validation.PreflightError):
            reports.start_run("reinspection", "2026-09", "database", [],
                              self.uploads(qc=1, ma=1, fail=2), validate_only=True)
        result = reports.start_run("reinspection", "2026-09", "database", [],
                                   self.uploads(qc=1, ma=1, fail=1), validate_only=True)
        self.assertTrue(result["valid"])
        self.assertGreater(result["cached_files"], 0)

    def test_direct_run_blocks_before_starting_worker(self):
        with patch.object(reports.subprocess, "Popen") as spawn:
            with self.assertRaises(validation.PreflightError):
                reports.start_run("reinspection", "2026-09", "database", [],
                                  self.uploads(qc=2, ma=1, fail=1))
            spawn.assert_not_called()

    def test_worker_checks_again_after_automatic_acquisition(self):
        with patch.object(reports.subprocess, "Popen"):
            run = reports.start_run("reinspection", "2026-09", "database", [], self.uploads())
        folder = reports.run_folder(run["id"])
        run["acquire_missing"] = True
        reports.write_json(folder / "run.json", run)

        def acquire(folder, run, stage):
            item = next(i for i in run["inputs"] if i["group"] == "hfpa" and
                        reports.plant_of(i["name"]) == "VH")
            path = folder / "workspace" / "HFPA" / item["staged_name"]
            path.write_bytes(source("hfpa", "VH", fail=2).stream.getvalue())

        previous_cwd, previous_path = Path.cwd(), sys.path[:]
        try:
            with patch("sfc_reinspection_acquisition.acquire", side_effect=acquire):
                self.assertEqual(worker.execute(folder), 1)
        finally:
            os.chdir(previous_cwd)
            sys.path[:] = previous_path
        result = reports.read_json(folder / "run.json")
        self.assertEqual(result["status"], "error")
        self.assertIn("HFPA Fail = 2", result["error"])
        self.assertEqual(len([c for c in result["business_review"] if c["status"] == "error"]), 2)
        self.assertFalse(result["files"])


if __name__ == "__main__":
    unittest.main()
