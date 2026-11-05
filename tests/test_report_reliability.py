"""P0 regression tests use local synthetic data only; no SFC or AI connection."""

from datetime import datetime, timedelta
import io
import json
import subprocess
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import openpyxl
import sfc_report_tools as reports
import sfc_dashboard_server as server
from report_validation import validate_file, CONTRACTS
from report_runtime import Monitor, Cancelled, diagnostics, friendly_error
from report_exports import check_outputs
from report_retention import cleanup


def workbook(path, schema="copq", rows=None, extra_sheet=False):
    path.parent.mkdir(parents=True, exist_ok=True)
    book = openpyxl.Workbook()
    labels = CONTRACTS["schemas"][schema]["required"]
    book.active.append(labels)
    if rows is None:
        defaults = {
            "Date": "2026-09-01",
            "Category": "QC",
            "Type": "Touch-up",
            "Factory": "VH",
            "Model": "Fixture",
            "Ttl cost ($)": 12,
            "InspDate": "2026-09-01",
            "Line": "A01",
            "ShoeName": "Fixture",
            "InspQty": 100,
            "FailQty": 2,
            "Station": "HFPA",
            "Issues": "Defect",
            "IssueQty": 2,
            "Audit Date (dd-mmm-yyyy)": "2026-09-01",
            "Plant": "A",
            "Model 2": "Fixture",
            "SO#": "SO",
            "Country": "VN",
            "Pass/Fail": "Fail",
            "Audit Sample Size": 100,
            "Defect Q'ty (Total Defect Q'ty)": 2,
            "Remark": "",
            "Reinspection Q'ty": 12,
            "2nd Audit Sample Size": 10,
            "2nd Pass/Fail": "Fail",
            "Metric": "B/Reject",
            "Defect Issues": "B/Defect",
            "Issues Q'ty": 2,
        }
        rows = [[defaults[label] for label in labels]]
    for row in rows:
        book.active.append(row)
    if extra_sheet:
        book.create_sheet("Other").append(["Unreviewed data"])
    stream = io.BytesIO()
    book.save(stream)
    book.close()
    path.write_bytes(stream.getvalue())
    return path


class ReliabilityTests(unittest.TestCase):
    def test_preflight_failure_exposes_the_validated_input_issue(self):
        from report_validation import PreflightError

        error = PreflightError([dict(error_count=1, issues=[dict(message="QAStation JV2 thiếu model ngày 04/11.")])])
        self.assertIn("QAStation JV2 thiếu model ngày 04/11.", friendly_error(error))

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.addCleanup(self.temp.cleanup)

    def test_all_source_contracts_accept_valid_data(self):
        for tool, group, schema in [
            ("copq", "copq", "copq"),
            ("copq", "ftt", "ftt"),
            ("hfpa", "ftt", "qastation"),
            ("hfpa", "hfpa", "hfpa"),
            ("reinspection", "reinspection", "reinspection"),
        ]:
            with self.subTest(schema=schema):
                path = workbook(self.root / (schema + "-VH.xlsx"), schema)
                result = validate_file(path, tool, group, "2026-09")
                self.assertTrue(result["valid"], result)
                self.assertEqual(result["rows"], 1)

    def test_scans_late_rows_and_returns_position_without_cell_contents(self):
        rows = [
            ["2026-09-01", "QC", "Touch-up", "VH", "Fixture", 12] for _ in range(150)
        ]
        rows[-1][-1] = "SECRET-PAYROLL"
        result = validate_file(
            workbook(self.root / "COPQ-VH.xlsx", rows=rows), "copq", "copq", "2026-09"
        )
        self.assertEqual(result["issues"][0]["row"], 151)
        self.assertEqual(result["issues"][0]["column"], "F")
        self.assertNotIn("SECRET", json.dumps(result))
        self.assertEqual(result["rows"], 150)

    def test_wrong_type_mime_size_formula_and_wrong_month_are_blocked(self):
        path = self.root / "COPQ-VH.xlsx"
        path.write_text("<html>Login SECRET</html>")
        self.assertFalse(validate_file(path, "copq", "copq", "2026-09")["valid"])
        workbook(path)
        self.assertFalse(
            validate_file(path, "copq", "copq", "2026-09", mime="text/html")["valid"]
        )
        with patch.dict(CONTRACTS, {"max_file_bytes": 1}):
            self.assertFalse(validate_file(path, "copq", "copq", "2026-09")["valid"])
        workbook(path, rows=[["2026-08-31", "QC", "Touch-up", "VH", "Fixture", "=1+2"]])
        result = validate_file(path, "copq", "copq", "2026-09")
        self.assertGreaterEqual(result["error_count"], 2)

    def test_hidden_additional_sheet_cannot_be_silently_ignored(self):
        path = workbook(self.root / "QAStation-VH.xlsx", "qastation", extra_sheet=True)
        result = validate_file(path, "hfpa", "ftt", "2026-09")
        self.assertFalse(result["valid"])
        self.assertEqual(len(result["sheets"]), 1)

    def test_hfpa_station_filter_is_specific_to_tool(self):
        path = workbook(self.root / "QAStation-VH.xlsx", "qastation")
        book = openpyxl.load_workbook(path)
        book.active["F2"] = "FTT"
        book.save(path)
        book.close()
        self.assertFalse(validate_file(path, "hfpa", "ftt", "2026-09")["valid"])

    def test_reinspection_keeps_earlier_and_missing_original_audit_date(self):
        path = workbook(self.root / "Re-Inspection-VH.xlsx", "reinspection")
        book = openpyxl.load_workbook(path)
        book.active["A2"] = "2026-08-31"
        book.active.append([None, "MA", "SO", "VN", "Fixture", 12, 10, "Fail", ""])
        book.save(path)
        book.close()
        result = validate_file(path, "reinspection", "reinspection", "2026-09")
        self.assertTrue(result["valid"], result)
        self.assertEqual(result["rows"], 2)

    def test_only_verified_footer_is_excluded_and_source_unchanged(self):
        path = workbook(self.root / "HFPA-VH.xls", "hfpa")
        book = openpyxl.load_workbook(io.BytesIO(path.read_bytes()))
        labels = [c.value for c in book.active[1]]
        row = [None] * len(labels)
        row[0] = 1
        row[labels.index("Defect Q'ty (Total Defect Q'ty)")] = 2
        book.active.append(row)
        stream = io.BytesIO()
        book.save(stream)
        book.close()
        path.write_bytes(stream.getvalue())
        original = path.read_bytes()
        result = validate_file(path, "hfpa", "hfpa", "2026-09")
        self.assertTrue(result["valid"], result)
        self.assertEqual(result["rows"], 1)
        self.assertEqual(result["excluded_summary_rows"][0]["row"], 3)
        self.assertEqual(path.read_bytes(), original)
        row[0] = 999
        book = openpyxl.load_workbook(io.BytesIO(path.read_bytes()))
        book.active["A3"] = 999
        stream = io.BytesIO()
        book.save(stream)
        book.close()
        path.write_bytes(stream.getvalue())
        self.assertFalse(validate_file(path, "hfpa", "hfpa", "2026-09")["valid"])

    def test_csv_normalization_keeps_source_hash_and_rework_rule_count(self):
        path = self.root / "COPQ-VH.csv"
        path.write_text(
            "Date,Category,Type,Factory,Model,Ttl cost ($)\n2026-09-01,QC,Rework,VH,Fixture,12\n",
            encoding="utf8",
        )
        result = validate_file(path, "copq", "copq", "2026-09", normalize=True)
        self.assertTrue(result["valid"])
        self.assertTrue(path.with_suffix(".xlsx").exists())
        self.assertEqual(result["business_exclusions"]["rework_rows"], 1)

    def test_copq_count_footer_and_missing_type_keep_legacy_semantics(self):
        path = workbook(
            self.root / "COPQ-VH.xlsx",
            rows=[
                ["2026-09-01", "QC", None, "VH", "Fixture", 12],
                ["Count=1", None, None, None, None, None],
            ],
        )
        result = validate_file(path, "copq", "copq", "2026-09")
        self.assertTrue(result["valid"], result)
        self.assertEqual(result["rows"], 1)
        self.assertEqual(result["excluded_summary_rows"][0]["row"], 3)

    def test_cancel_marker_never_interrupts_a_write_and_is_seen_at_next_stage(self):
        run = dict(steps=[])
        monitor = Monitor(self.root, run)
        monitor.stage("Writing output", "writing")
        (self.root / "cancel.request").touch()
        target = self.root / "completed.xlsx"
        target.write_bytes(b"completed atomic write")
        with self.assertRaises(Cancelled):
            monitor.stage("Next step", "next")
        self.assertEqual(target.read_bytes(), b"completed atomic write")
        self.assertEqual(run["steps"][0]["status"], "running")

    def test_diagnostics_has_code_and_frames_without_values_or_absolute_paths(self):
        try:
            raise RuntimeError("SECRET cookie=abc password=def")
        except RuntimeError as error:
            result = diagnostics(error)
        encoded = json.dumps(result)
        self.assertNotIn("SECRET", encoded)
        self.assertNotIn("cookie", encoded)
        self.assertNotIn(str(self.root), encoded)
        self.assertEqual(result["code"], "RuntimeError")

    def test_api_preflight_does_not_spawn_and_start_revalidates(self):
        source = workbook(self.root / "COPQ-VH.xlsx")
        engines = self.root / "engines"
        (engines / "copq").mkdir(parents=True)
        with (
            patch.object(reports, "TOOLS_ROOT", engines),
            patch.object(reports, "RUN_ROOT", self.root / "runs"),
            patch.object(reports, "input_catalog", return_value=[]),
            patch.object(reports.subprocess, "Popen") as spawn,
            server.app.test_client() as client,
        ):

            def request(endpoint, content):
                return client.post(
                    endpoint,
                    data={
                        "period": "2026-09",
                        "mode": "combine",
                        "selections": "[]",
                        "copq": (io.BytesIO(content), "COPQ-VH.xlsx"),
                    },
                    content_type="multipart/form-data",
                )

            response = request("/api/report-tools/copq/validate", source.read_bytes())
            self.assertEqual(response.status_code, 200, response.json)
            spawn.assert_not_called()
            self.assertFalse(
                (self.root / "runs").exists()
                and list((self.root / "runs").glob("*/run.json"))
            )
            response = request(
                "/api/report-tools/copq/runs", b"<html>Login SECRET</html>"
            )
            self.assertEqual(response.status_code, 422)
            spawn.assert_not_called()
            self.assertNotIn("SECRET", response.get_data(as_text=True))

    def test_cross_origin_mutation_is_rejected(self):
        with server.app.test_client() as client:
            response = client.post(
                "/api/report-tools/copq/validate",
                json={},
                headers={"Origin": "https://unrelated.example"},
            )
            self.assertEqual(response.status_code, 403)

    def test_invalid_request_shapes_return_friendly_errors(self):
        with server.app.test_client() as client:
            for body in (
                [1],
                {"mode": []},
                {"selections": [{"id": []}]},
                {"selections": [{"id": "x", "group": []}]},
            ):
                self.assertEqual(
                    client.post(
                        "/api/report-tools/copq/validate", json=body
                    ).status_code,
                    400,
                )

    def test_windows_reader_lock_retries_atomic_status_update(self):
        target = self.root / "run.json"
        original = Path.replace
        attempts = []

        def briefly_locked(source, destination):
            attempts.append(1)
            if len(attempts) < 3:
                raise PermissionError("reader briefly holds file")
            return original(source, destination)

        with patch.object(Path, "replace", briefly_locked):
            reports.write_json(target, {"status": "done"})
        self.assertEqual(reports.read_json(target)["status"], "done")
        self.assertEqual(len(attempts), 3)
        self.assertEqual(list(self.root.glob("*.tmp")), [])

    def test_final_worker_commit_is_not_overwritten_by_stale_poll(self):
        folder = self.root / ("d" * 32)
        folder.mkdir()
        stale = dict(
            id=folder.name,
            status="running",
            pid=123,
            started_at=reports.now(),
            files=[],
        )
        final = {**stale, "status": "done"}
        reports.write_json(folder / "run.json", final)
        with (
            patch.object(reports, "RUN_ROOT", self.root),
            patch.object(reports, "read_json", side_effect=[stale, final]),
            patch.object(reports, "process_alive", return_value=False),
        ):
            self.assertEqual(reports.snapshot(folder.name)["status"], "done")
        self.assertEqual(reports.read_json(folder / "run.json")["status"], "done")

    def test_queued_cancel_is_honored_by_actual_worker(self):
        folder = self.root / ("b" * 32)
        (folder / "workspace").mkdir(parents=True)
        run = dict(
            id=folder.name,
            tool="copq",
            title="Fixture",
            period="2026-09",
            mode="combine",
            status="queued",
            started_at=reports.now(),
            monitor_version=1,
            inputs=[],
            files=[],
            steps=[],
        )
        reports.write_json(folder / "run.json", run)
        with patch.object(reports, "RUN_ROOT", self.root):
            self.assertTrue(reports.request_cancel(folder.name)["cancel_requested"])
            worker = subprocess.run(
                [
                    sys.executable,
                    str(reports.APP_ROOT / "report_tool_worker.py"),
                    str(folder),
                ],
                capture_output=True,
                timeout=10,
            )
            self.assertEqual(reports.snapshot(folder.name)["status"], "cancelled")
            self.assertEqual(worker.stdout, b"")
            self.assertEqual(worker.stderr, b"")

    def test_download_finishes_before_cancel_and_browser_filters_are_restored(self):
        from test_hfpa_acquisition import HFPAcquisitionTests
        import sfc_hfpa_acquisition as acquisition

        fixture = HFPAcquisitionTests()
        fixture.setUp()
        self.addCleanup(fixture.tearDown)
        folder = fixture.runs / ("c" * 32)
        (folder / "workspace").mkdir(parents=True)
        fake = fixture.fake_controller()
        original = fake._execute_dynamic_action.side_effect

        def export_then_cancel(*args, **kwargs):
            result = original(*args, **kwargs)
            if args[0] != "restore_filters":
                (folder / "cancel.request").touch()
            return result

        fake._execute_dynamic_action.side_effect = export_then_cancel
        run = dict(period="2026-09", inputs=[], acquisition=acquisition.plan("2026-09"))
        with self.assertRaises(Cancelled):
            acquisition.acquire(folder, run, lambda message: None, fake)
        self.assertEqual(len(run["inputs"]), 1)
        self.assertTrue(
            any(
                call.args[0] == "restore_filters"
                for call in fake._execute_dynamic_action.call_args_list
            )
        )
        self.assertEqual(len(acquisition.plan("2026-09")["available"]), 1)

    def test_export_warns_on_row_cost_and_business_mismatch_and_gates_download(self):
        folder = self.root / ("a" * 32)
        output = folder / "workspace/Output"
        output.mkdir(parents=True)
        book = openpyxl.Workbook()
        book.active.title = "COPQ_Clean"
        book.active.append(["Ttl cost ($)", "Data Quality Status"])
        book.active.append([99, "Missing Data"])
        book.save(output / "result.xlsx")
        book.close()
        files = reports.collect_files(folder)
        run = dict(
            id=folder.name,
            tool="copq",
            mode="combine",
            status="done",
            started_at=reports.now(),
            files=files,
            validation=[dict(group="copq", rows=2, totals={"Ttl cost ($)": 12})],
        )
        review = check_outputs(folder, run, files)
        self.assertEqual(review["warning_count"], 3)
        self.assertTrue(files[0]["sha256"])
        self.assertEqual(files[0]["sheets"][0]["name"], "COPQ_Clean")
        run.update(export_review=review, requires_ack=True)
        reports.write_json(folder / "run.json", run)
        with (
            patch.object(reports, "RUN_ROOT", self.root),
            server.app.test_client() as client,
        ):
            url = f"/api/report-tools/runs/{folder.name}/files/result.xlsx"
            self.assertEqual(client.get(url).status_code, 409)
            response = client.get(url + "?ack=1")
            self.assertEqual(response.status_code, 200)
            response.close()
            with (output / "result.xlsx").open("ab") as stream:
                stream.write(b"changed after review")
            self.assertEqual(client.get(url + "?ack=1").status_code, 409)

    def test_retention_defaults_to_preview_and_preserves_old_or_active_runs(self):
        old = (datetime.fromisoformat(reports.now()) - timedelta(days=40)).isoformat()
        for index, monitor, status in [
            (1, 1, "done"),
            (2, 1, "running"),
            (3, None, "done"),
        ]:
            folder = self.root / (str(index) * 32)
            folder.mkdir()
            reports.write_json(
                folder / "run.json",
                dict(
                    id=folder.name,
                    monitor_version=monitor,
                    status=status,
                    finished_at=old,
                ),
            )
        with patch.object(reports, "RUN_ROOT", self.root):
            self.assertEqual(cleanup(), ["1" * 32])
            self.assertTrue((self.root / ("1" * 32)).exists())
            cleanup(apply=True)
        self.assertFalse((self.root / ("1" * 32)).exists())
        self.assertTrue((self.root / ("2" * 32)).exists())
        self.assertTrue((self.root / ("3" * 32)).exists())


if __name__ == "__main__":
    unittest.main()
