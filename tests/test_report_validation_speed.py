"""Content-based preflight reuse must preserve exhaustive validation semantics."""

import io
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import report_validation as validation
import sfc_report_tools as reports
import sfc_dashboard_server as server


CSV = (
    b"Date,Category,Type,Factory,Model,Ttl cost ($)\n"
    b"2026-09-01,QC,Touch-up,VH,Fixture,12\n"
)


class PreflightSpeedTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.path = self.root / "COPQ-VH.csv"
        self.path.write_bytes(CSV)
        validation._PREFLIGHT_CACHE.clear()
        self.addCleanup(validation._PREFLIGHT_CACHE.clear)

    def check(self, path=None, **kwargs):
        return validation.validate_preflight(
            path or self.path, "copq", "copq", kwargs.pop("period", "2026-09"), **kwargs
        )

    def test_identical_content_reuses_summary_without_sharing_mutable_state(self):
        with patch.object(validation, "validate_file", wraps=validation.validate_file) as scan:
            first, reused = self.check()
            self.assertFalse(reused)
            first["totals"].clear()
            renamed = self.root / "COPQ-VH2.csv"
            renamed.write_bytes(CSV)
            second, reused = self.check(renamed)
            self.assertTrue(reused)
            self.assertEqual(scan.call_count, 1)
            self.assertEqual(second["name"], renamed.name)
            self.assertEqual(second["totals"]["Ttl cost ($)"], 12)
            self.assertEqual(self.path.read_bytes(), CSV)
            self.assertFalse(self.path.with_suffix(".xlsx").exists())

    def test_reinspection_ftt_reuses_only_engine_ready_bytes(self):
        from openpyxl import Workbook, load_workbook

        path = self.root / "FTT-VH.xlsx"

        def write(quantity):
            book = Workbook()
            book.active.append(["Date", "Model", "Metric", "Defect Issues", "Issues Q'ty"])
            book.active.append(["2026-09-01", "Fixture", "B/Reject", "Test", quantity])
            book.save(path)

        def check(**kwargs):
            return validation.validate_preflight(path, "reinspection", "ftt", "2026-09", **kwargs)

        write(12)
        original = path.read_bytes()
        self.assertTrue(check()[0]["engine_ready"])
        result, reused = check(normalize=True)
        self.assertTrue(reused)
        self.assertEqual(result["staged_name"], path.name)
        self.assertEqual(path.read_bytes(), original)

        write("1,200")
        result, reused = check(normalize=True)
        self.assertFalse(reused)
        self.assertTrue(result["valid"])
        self.assertNotIn("engine_ready", result)
        with path.open("rb") as stream:
            book = load_workbook(stream, read_only=True)
            try:
                self.assertEqual(book.active["E2"].value, 1200)
            finally:
                book.close()

        write(-12)
        result, reused = check(normalize=True)
        self.assertFalse(reused)
        self.assertFalse(result["valid"])

    def test_same_size_and_timestamp_change_is_still_checked(self):
        self.check()
        before = self.path.stat()
        self.path.write_bytes(CSV.replace(b",12\n", b",-1\n"))
        os.utime(self.path, ns=(before.st_atime_ns, before.st_mtime_ns))
        self.assertEqual(self.path.stat().st_size, before.st_size)
        result, reused = self.check()
        self.assertFalse(reused)
        self.assertFalse(result["valid"])
        self.assertEqual(result["issues"][0]["row"], 2)
        self.assertEqual(result["issues"][0]["column"], "F")

    def test_month_mime_schema_and_contract_changes_invalidate_summary(self):
        self.assertTrue(self.check()[0]["valid"])
        result, reused = self.check(period="2026-10")
        self.assertFalse(reused)
        self.assertFalse(result["valid"])
        result, reused = self.check(mime="text/html")
        self.assertFalse(reused)
        self.assertFalse(result["valid"])
        with patch.object(validation, "schema_for", return_value="ftt"):
            result, reused = self.check()
            self.assertFalse(reused)
            self.assertFalse(result["valid"])
        with patch.dict(validation.CONTRACTS, max_file_bytes=1):
            result, reused = self.check()
            self.assertFalse(reused)
            self.assertFalse(result["valid"])
        # A changed rule must invalidate even when it does not change the file limit.
        with patch.dict(validation.CONTRACTS, version="new-rules"):
            self.assertFalse(self.check()[1])

    def test_cache_is_bounded_and_invalid_results_are_reused(self):
        with patch.object(validation, "_PREFLIGHT_CACHE_SIZE", 2):
            self.check()
            self.path.write_bytes(CSV.replace(b",12\n", b",-1\n"))
            self.assertFalse(self.check()[0]["valid"])
            self.assertTrue(self.check()[1])
            self.path.write_bytes(CSV.replace(b",12\n", b",-2\n"))
            self.check()
            self.assertEqual(len(validation._PREFLIGHT_CACHE), 2)
            self.path.write_bytes(CSV)
            self.assertFalse(self.check()[1])

    def test_upload_preflight_skips_engine_copy_catalog_and_normalization(self):
        with (
            patch.object(reports, "RUN_ROOT", self.root / "runs"),
            patch.object(reports, "input_catalog", side_effect=AssertionError("unneeded scan")),
            patch.object(reports.shutil, "copytree", side_effect=AssertionError("unneeded copy")),
            patch.object(reports.subprocess, "Popen") as spawn,
            patch.object(validation.openpyxl, "Workbook", side_effect=AssertionError("unneeded write")),
            server.app.test_client() as client,
        ):
            for index in range(2):
                response = client.post(
                    "/api/report-tools/copq/validate",
                    data={"period": "2026-09", "mode": "combine", "selections": "[]",
                          "copq": (io.BytesIO(CSV), self.path.name)},
                    content_type="multipart/form-data",
                )
                self.assertEqual(response.status_code, 200, response.json)
                self.assertEqual(response.json["cached_files"], index)
                self.assertEqual(response.json["checked_files"], 1)
                self.assertGreaterEqual(response.json["validation_time_ms"], 0)
            spawn.assert_not_called()
            self.assertEqual(list((self.root / "runs").glob("*/run.json")), [])


if __name__ == "__main__":
    unittest.main()
