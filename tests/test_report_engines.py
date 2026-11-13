"""Offline integration: run all three original engines against synthetic inputs."""

import io
import json
import subprocess
from pathlib import Path
import sys
import tempfile
import time
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import openpyxl
from werkzeug.datastructures import FileStorage
import sfc_report_tools as reports


def fixture(group, plant):
    values = {
        "copq": {
            "Date": "2026-09-01",
            "Category": "Production",
            "Type": "B/C",
            "Factory": plant,
            "Fty/Line": "A01",
            "Style Nbr": "TEST-001",
            "Mold": "M01",
            "Model": "Fixture model",
            "Defective Qty(Pair)": 2,
            "FOB ($)": 5,
            "Manpower": 1,
            "Labor cost ($)": 1,
            "Working hours": 1,
            "SGS Retest Qty (Pic)": 0,
            "SGS Test unit price/pic ($)": 0,
            "SGS Retest failure C grade": 0,
            "Ttl cost ($)": 10,
            "PO": "TEST-PO",
            "Remark": "",
            "Remark (Re-inspection Station)": "",
        },
        "ftt": {
            "Date": "2026-09-01",
            "FactoryGroup": "CL",
            "Factory": plant,
            "Plant": "A",
            "Line": "A01",
            "Model": "Fixture model",
            "Model 2": "Fixture model",
            "SO#": "TEST",
            "Season": "TEST",
            "Cutting Die": "TEST",
            "Style Nbr": "TEST-001",
            "Colorway Cd": "TEST",
            "Total Pair Produced": 100,
            "Station": "FTT",
            "Metric": "B/Reject",
            "Number of pairs": 2,
            "Defect Issues": "B/Over cement",
            "Issues Q'ty": 2,
            "Defective Area": "",
        },
        "qastation": {
            "InspDate": "2026-09-01",
            "Line": "A01",
            "ShoeName": "Fixture model",
            "InspQty": 100,
            "FailQty": 2,
            "Station": "HFPA",
            "Issues": "Over cement",
            "IssueQty": 2,
        },
        "hfpa": {
            "Factory Group Cd_UccId": plant,
            "SO#": "TEST",
            "Country": "VN",
            "Season": "TEST",
            "Audit Date (dd-mmm-yyyy)": "2026-09-01",
            "Internal AUDIT ID (YYMMDDNO+Plant+Line)": "TEST",
            "Plant": "A",
            "Line": "A01",
            "Auditor": "TEST",
            "PO Nbr (header+line item nbr)": "TEST",
            "Model": "Fixture model",
            "Model 2": "Fixture model",
            "Style Nbr": "TEST-001",
            "Colorway Cd": "TEST",
            "Time table": "TEST",
            "Produced Q'ty": 100,
            "Audit Sample Size": 100,
            "Defective Q'ty(pair)": 2,
            "Defective Rate": 0.02,
            "Pass/Fail": "Fail",
            "Defect Q'ty (Total Defect Q'ty)": 2,
            "Over cement": 2,
        },
        "reinspection": {
            "Category": "MA" if plant == "JV" else "QC",
            "SO#": "TEST",
            "Country": "VN",
            "Audit Date (dd-mmm-yyyy)": "2026-08-31",
            "Model 2": "Fixture model",
            "Reinspection Q'ty": 100,
            "2nd Audit Sample Size": 100,
            "2nd Pass/Fail": "Fail",
            "Defect Q'ty (Total Defect Q'ty)": 2,
            "Remark": "",
            "Over cement": 2,
        },
    }
    data = values[group]
    book = openpyxl.Workbook()
    book.active.append(list(data))
    book.active.append(list(data.values()))
    if group == "reinspection":
        other = dict(data, Category="QC" if data["Category"] == "MA" else "MA")
        book.active.append(list(other.values()))
    stream = io.BytesIO()
    book.save(stream)
    book.close()
    stream.seek(0)
    return FileStorage(
        stream=stream,
        filename=f"{group}_2026-09-01_2026-09-30-{plant}.xlsx",
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


class OriginalEngineTests(unittest.TestCase):
    def test_hfpa_keeps_model_inspections_and_ranks_by_mes410_failed_pairs(self):
        import pandas as pd
        from report_tools.hfpa.agents.reconciliation_engine_agent import ReconciliationEngineAgent

        qa = pd.DataFrame({
            "Date_Raw": ["2026-10-07"] * 3,
            "FTY": ["VH2"] * 3,
            "Plant_Parsed": ["A10"] * 3,
            "Line_Parsed": ["3"] * 3,
            "Station_Raw": ["HFPA"] * 3,
            "Model_Raw": ["Model A", "Model A", "Model B"],
            "InspQty_Raw": [33, 33, 373],
            "FailQty_Raw": [4, 4, 9],
            "Issues_Raw": ["Thread end", "Cleanness", "Cleanness"],
            "IssueQty_Raw": [7, 5, 9],
        })
        mes = pd.DataFrame({
            "Audit Date (dd-mmm-yyyy)": ["2026-10-07"] * 2,
            "FTY": ["VH2"] * 2,
            "Plant": ["A10"] * 2,
            "Line": [3, 3],
            "Model 2": ["Model A", "Model B"],
            "Audit Sample Size": [33, 373],
            "Defective Q'ty(pair)": [4, 8],
            "Defect Q'ty (Total Defect Q'ty)": [7, 8],
            "Thread end": [3, 0],
            "Cleanness": [4, 6],
            "Rocking": [0, 2],
        })

        with tempfile.TemporaryDirectory() as temporary:
            engine = ReconciliationEngineAgent(temporary)
            validated, lots, audit = engine.reconcile_and_validate(qa, mes)
            pivot, _, top5, defects = engine.generate_aggregations(validated, mes, {})

        self.assertEqual(lots["QA_InspQty"].tolist(), [406])
        self.assertEqual(lots["QA_FailQty"].tolist(), [13])
        self.assertEqual(pivot["InspQty"].tolist(), [33, 373])
        self.assertEqual(pivot["FailQty_Original"].tolist(), [4, 9])
        self.assertEqual(pivot["FailQty"].tolist(), [4, 8])
        self.assertAlmostEqual(pivot.loc[pivot["Shoename"] == "Model B", "DR%_Validated"].iloc[0], 8 / 373)
        self.assertEqual(top5["VH2"]["Shoename"].tolist(), ["Model B", "Model A"])
        self.assertEqual(top5["VH2"]["Sum of InspQty"].tolist(), [406, 406])
        self.assertEqual(top5["VH2"]["Sum of FailQty (Original)"].tolist(), [9, 4])
        self.assertEqual(top5["VH2"]["Sum of FailQty (Validated)"].tolist(), [8, 4])
        self.assertAlmostEqual(top5["VH2"]["DR%"].iloc[0], 8 / 406)
        model_a = defects[(defects["FTY"] == "VH2") & (defects["Shoename"] == "Model A")]
        self.assertAlmostEqual(model_a["Defect Share %"].iloc[0], 7 / 12)
        model_b = defects[(defects["FTY"] == "VH2") & (defects["Shoename"] == "Model B")]
        self.assertEqual(model_b["Defect Type"].tolist(), ["Cleanness", "Rocking"])
        self.assertEqual(model_b["Defect Quantity"].tolist(), [6, 2])
        self.assertEqual([row["Status"] for row in audit], ["MES410_Corrected"])

    def test_copq_combine_includes_valid_files_without_month_in_name(self):
        with tempfile.TemporaryDirectory() as temporary, patch.object(
            reports, "RUN_ROOT", Path(temporary) / "runs"
        ):
            jv = fixture("copq", "JV")
            jv.filename = "COPQ-JV.xlsx"
            run = reports.start_run("copq", "2026-09", "combine", [], [
                ("copq", fixture("copq", "VH")), ("copq", jv)
            ])
            self.assertEqual(reports.WORKERS[run["id"]].wait(timeout=60), 0)
            run = reports.snapshot(run["id"])
            checks = {check["rule"]: check for check in run["export_review"]["checks"]}
            self.assertEqual(checks["copq.row_count"]["actual"], 2)
            self.assertEqual(checks["copq.cost"]["status"], "passed")

    def test_all_three_tools_keep_excel_pptx_and_bundle_exports(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            for tool in ("copq", "reinspection", "hfpa"):
                with (
                    self.subTest(tool=tool),
                    patch.object(reports, "RUN_ROOT", root / "runs"),
                    patch.object(reports, "input_catalog", return_value=[]),
                ):
                    uploads = []
                    for group in reports.TOOLS[tool]["groups"]:
                        plants = (
                            reports.TOOLS[tool]["plants"]
                            if group == "copq"
                            else reports.COMMON_PLANTS
                        )
                        for plant in plants:
                            upload = fixture(
                                "qastation" if tool == "hfpa" and group == "ftt" else group,
                                plant,
                            )
                            if tool == "reinspection" and group == "reinspection" and plant == "JV":
                                upload.filename = "00exportVH_" + upload.filename
                            uploads.append((group, upload))
                    run = reports.start_run(tool, "2026-09", "all", [], uploads)
                    deadline = time.monotonic() + 120
                    while time.monotonic() < deadline:
                        time.sleep(0.1)
                        run = reports.snapshot(run["id"])
                        if run["status"] not in ("queued", "running"):
                            break
                    # The final JSON commit precedes interpreter shutdown. Windows
                    # cannot remove a temporary workspace while it is a child's cwd.
                    worker = reports.WORKERS.get(run["id"])
                    if worker:
                        self.assertEqual(worker.wait(timeout=30), 0)
                    self.assertEqual(
                        run["status"],
                        "done",
                        json.dumps(
                            {
                                k: run.get(k)
                                for k in ("status", "stage", "error", "diagnostics")
                            },
                            ensure_ascii=True,
                        ),
                    )
                    self.assertTrue(
                        any(f["name"].endswith(".xlsx") for f in run["files"])
                    )
                    self.assertTrue(
                        any(f["name"].endswith(".pptx") for f in run["files"])
                    )
                    self.assertTrue(run["bundle"])
                    workspace = reports.run_folder(run["id"]) / "workspace"
                    if tool == "reinspection":
                        self.assertTrue(all(
                            check["status"] == "passed" for check in run["business_review"]
                        ), run["business_review"])
                    if tool == "hfpa":
                        from pptx import Presentation
                        from report_tools.hfpa.hfpa_common import verify_hfpa_deliverables_consistency

                        master, = (workspace / "Output").glob("HFPA_FTT_Combined_Master_*.xlsx")
                        book = openpyxl.load_workbook(master, read_only=True)
                        try:
                            header, *lots = book["Lot_Defect_Variance"].values
                            status = header.index("Validation_Status")
                            self.assertEqual([row[status] for row in lots], ["Matched_Exact"] * 4)
                        finally:
                            book.close()
                        palette = openpyxl.load_workbook(workspace / "Color_template.xlsx", read_only=True)
                        try:
                            colors = {row[2]: row[7] for row in palette["Sheet1"].values if row[2]}
                            self.assertEqual(colors["Lace"], "316CE5")
                        finally:
                            palette.close()
                        output = master.parent
                        review = verify_hfpa_deliverables_consistency(output, "Sep", "2026")
                        self.assertEqual(review["status"], "passed", review)
                        excel_review = verify_hfpa_deliverables_consistency(output, "Sep", "2026", include_pptx=False)
                        self.assertEqual(excel_review["status"], "passed", excel_review)
                        self.assertFalse(any(c["rule"].startswith("hfpa.pptx.") for c in excel_review["checks"]))
                        monthly, = output.glob("HFPA_Template_Updated_Sep_2026.xlsx")
                        book = openpyxl.load_workbook(monthly)
                        sheet = book["HFPA"]
                        sheet["D22"] = None
                        sheet["M22"] = None
                        sheet["G28"] = "=1"
                        sheet["M21"].border = openpyxl.styles.Border()
                        book.save(monthly)
                        book.close()
                        pptx, = output.glob("HFPA_Performance_Report_Sep_2026.pptx")
                        deck = Presentation(pptx)
                        chart = next(shape for shape in deck.slides[0].shapes if shape.name == "Chart 194")
                        chart._element.getparent().remove(chart._element)
                        deck.save(pptx)
                        book = openpyxl.load_workbook(master)
                        pivot = book["Pivot1_Validation"]
                        header = [c.value for c in pivot[1]]
                        pivot.cell(2, header.index("DR%_Validated") + 1, 0.99)
                        pending = book.create_sheet("Model_Adjustments")
                        pending.append(["Status"])
                        pending.append(["Missing_QAStation_Model"])
                        book.save(master)
                        book.close()
                        review = verify_hfpa_deliverables_consistency(output, "Sep", "2026")
                        warnings = {c["rule"] for c in review["checks"] if c["status"] == "warning"}
                        self.assertTrue({"hfpa.pptx.charts", "hfpa.top5_ranks", "hfpa.defect_shares",
                                         "hfpa.excel.formulas", "hfpa.excel.layout", "hfpa.pivot_rates", "hfpa.model_review"} <= warnings, review)
                    if tool == "copq":
                        audit = subprocess.run(
                            [
                                sys.executable,
                                str(workspace / "agents/run_agent.py"),
                                "audit",
                                "--period",
                                "2026-09",
                            ],
                            cwd=workspace,
                            capture_output=True,
                            text=True,
                            encoding="utf8",
                            timeout=30,
                        )
                        self.assertEqual(
                            audit.returncode,
                            0,
                            audit.stdout[-2000:] + audit.stderr[-1000:],
                        )
                    self.assertTrue(
                        all(s["status"] == "done" for s in run["steps"]), run["steps"]
                    )
                    self.assertEqual(run.get("log", "").count("TEST-PO"), 0)


class HFPABusinessRulesTests(unittest.TestCase):
    """Specific unit tests for HFPA business rules, tie-breaking, missing data, and unit reconciliation."""

    def test_missing_model_is_zero_only_when_mes410_has_the_same_line_and_day(self):
        import pandas as pd
        from report_tools.hfpa.agents.reconciliation_engine_agent import ReconciliationEngineAgent

        qa = pd.DataFrame({
            "InspDate": ["2026-10-21", "2026-10-22"],
            "FTY": ["JV2", "JV2"], "Line": ["F Line 6", "F Line 6"],
            "ShoeName": ["Model absent", "Model unverified"],
            "InspQty": [1, 10], "FailQty": [0, 1], "Station": ["HFPA", "HFPA"],
            "Issues": ["", "Cleanness"], "IssueQty": [0, 1],
        })
        mes = pd.DataFrame({
            "FTY": ["JV2"], "Audit Date": ["2026-10-21"], "Plant": ["F"], "Line": [6],
            "Model 2": ["Other model"], "Audit Sample Size": [106],
            "Defective Q'ty(pair)": [4], "Defect Q'ty (Total Defect Q'ty)": [4],
            "Cleanness": [4],
        })
        adjusted, _, audit = ReconciliationEngineAgent.reconcile_and_validate(qa, mes)
        first = adjusted[adjusted["ShoeName_Std"] == "Model absent"].iloc[0]
        second = adjusted[adjusted["ShoeName_Std"] == "Model unverified"].iloc[0]
        self.assertEqual((first["InspQty_Validated"], first["FailQty_Validated"], first["IssueQty_Validated"]), (0, 0, 0))
        self.assertEqual(second["InspQty_Validated"], 10)
        self.assertEqual([row["Status"] for row in audit],
                         ["MES410_Corrected", "No_MES410_Model", "Missing_QAStation_Model"])

    def test_sop_pivot_counts_distinct_quantity_pairs_without_repeating_issue_rows(self):
        import pandas as pd
        from report_tools.hfpa.agents.reconciliation_engine_agent import ReconciliationEngineAgent as Engine
        qa = pd.DataFrame({
            "InspDate": ["2026-10-07"] * 3, "FTY": ["VH"] * 3, "Line": ["A Line 1"] * 3,
            "ShoeName": ["Model"] * 3, "InspQty": [50, 50, 50], "FailQty": [5, 5, 4],
            "Station": ["HFPA"] * 3, "Issues": ["Cleanness", "Rocking", "Cleanness"], "IssueQty": [3, 2, 4],
        })
        mes = pd.DataFrame({
            "Audit Date": ["2026-10-07"], "FTY": ["VH"], "Plant": ["A"], "Line": [1], "Model 2": ["Model"],
            "Audit Sample Size": [100], "Defective Q'ty(pair)": [9],
            "Defect Q'ty (Total Defect Q'ty)": [9], "Cleanness": [7], "Rocking": [2],
        })
        adjusted, lots, audit = Engine.reconcile_and_validate(qa, mes)
        pivot = Engine.inspection_pivot(adjusted)
        self.assertEqual((pivot["InspQty"].sum(), pivot["FailQty"].sum()), (100, 9))
        self.assertEqual(lots["Validation_Status"].tolist(), ["Matched_Exact"])
        self.assertEqual(audit, [])
        engine_path = str(Path(__file__).resolve().parents[1] / "report_tools" / "hfpa")
        with patch.object(sys, "path", [engine_path, *sys.path]):
            from process_qastation import validate_and_correct_data, generate_pivot1_and_top5
            legacy, legacy_mes, _ = validate_and_correct_data(qa.copy(), mes.copy())
            legacy_pivot, _, _ = generate_pivot1_and_top5(legacy, legacy_mes)
        pd.testing.assert_frame_equal(pivot, legacy_pivot)

    def test_alias_match_is_used_for_model_corrections(self):
        import pandas as pd
        from report_tools.hfpa.agents.reconciliation_engine_agent import ReconciliationEngineAgent as Engine
        qa = pd.DataFrame({
            "InspDate": ["2026-10-07"] * 2, "FTY": ["VH"] * 2, "Line": ["Line 1"] * 2,
            "ShoeName": ["A", "B"], "InspQty": [30, 70], "FailQty": [3, 7],
            "Station": ["HFPA"] * 2, "Issues": ["Cleanness"] * 2, "IssueQty": [3, 7],
        })
        mes = pd.DataFrame({
            "Audit Date": ["2026-10-07"] * 2, "FTY": ["VH"] * 2, "Plant": ["PlantA"] * 2,
            "Line": [1, 1], "Model 2": ["A", "B"], "Audit Sample Size": [40, 60],
            "Defective Q'ty(pair)": [4, 5], "Defect Q'ty (Total Defect Q'ty)": [4, 5], "Cleanness": [4, 5],
        })
        adjusted, _, audit = Engine.reconcile_and_validate(qa, mes)
        pivot = Engine.inspection_pivot(adjusted)
        self.assertEqual(pivot["InspQty"].tolist(), [40, 60])
        self.assertEqual(pivot["FailQty"].tolist(), [4, 5])
        self.assertEqual([row["Status"] for row in audit], ["MES410_Corrected"] * 2)

    def test_zero_issues_are_distinguished_from_missing_data(self):
        import pandas as pd
        from report_tools.hfpa.hfpa_common import compute_top3_defects
        qa = pd.DataFrame({"FTY": ["VH"], "ShoeName_Std": ["Model"], "Issues_Std": [""],
                           "IssueQty_Validated": [0], "FailQty_Validated": [0]})
        top = {"VH": pd.DataFrame({"Shoename": ["Model"]})}
        result = compute_top3_defects(qa, top)
        self.assertEqual(result["Data Status"].tolist(), ["No_Defects"])

    def test_defect_table_borders_and_legend_preserve_columns_beyond_t(self):
        engine_path = str(Path(__file__).resolve().parents[1] / "report_tools" / "hfpa")
        with patch.object(sys, "path", [engine_path, *sys.path]):
            from populate_database import apply_elegant_borders, populate_custom_defect_color_legend
        from report_tools.hfpa.hfpa_common import hfpa_defect_layout_issues
        book = openpyxl.Workbook()
        sheet = book.active
        for header, start in ((21, 22), (27, 28), (33, 34), (39, 40)):
            sheet.cell(start, 4, "Model")
            sheet.cell(start, 12, "Model")
            for col in range(13, 31):
                sheet.cell(header, col, f"Defect {col}")
                sheet.cell(start, col, 0.01)
        apply_elegant_borders(sheet, "Oct")
        with patch.object(sys, "path", [engine_path, *sys.path]):
            populate_custom_defect_color_legend(sheet, "Oct", "2026")
        for header, start in ((21, 22), (27, 28), (33, 34), (39, 40)):
            for col in range(13, 31):
                self.assertEqual(sheet.cell(header, col).value, f"Defect {col}")
                self.assertEqual(sheet.cell(start, col).value, 0.01)
                self.assertEqual(sheet.cell(header, col).border.top.style, "medium")
                self.assertTrue(sheet.cell(start, col).border.right.style)
        self.assertEqual(sheet["AD26"].border.right.style, "medium")
        self.assertEqual(sheet["AD26"].border.bottom.style, "medium")
        self.assertEqual(sheet["L66"].value, "Defect 30")
        self.assertIn("$AD$", str(sheet.print_area))
        self.assertFalse(hfpa_defect_layout_issues(sheet))
        # Repeated generation and the Excel save/reload retain the expanded table and legend.
        apply_elegant_borders(sheet, "Oct")
        with patch.object(sys, "path", [engine_path, *sys.path]):
            populate_custom_defect_color_legend(sheet, "Oct", "2026")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "expanded.xlsx"
            book.save(path)
            saved = openpyxl.load_workbook(path)
            self.assertFalse(hfpa_defect_layout_issues(saved.active))
            self.assertEqual(saved.active["AD22"].value, 0.01)
            saved.close()
        # Independent review catches damage even after the renderer succeeds.
        sheet["AD26"].border = openpyxl.styles.Border()
        self.assertTrue(any(issue["cell"] == "AD26" for issue in hfpa_defect_layout_issues(sheet)))

    def test_defect_layout_handles_long_headers_and_rejects_orphan_values(self):
        engine_path = str(Path(__file__).resolve().parents[1] / "report_tools" / "hfpa")
        with patch.object(sys, "path", [engine_path, *sys.path]):
            from populate_database import apply_elegant_borders
        from report_tools.hfpa.hfpa_common import hfpa_defect_layout_issues
        book = openpyxl.Workbook()
        sheet = book.active
        sheet["M21"] = "Midsole/Outsole to upper bond gap with an unusually long defect description"
        sheet["M22"] = 0.12
        apply_elegant_borders(sheet, "Oct", "2027")
        self.assertGreater(sheet.row_dimensions[21].height, 48)
        self.assertIn("2027", sheet["J21"].value)
        self.assertFalse(hfpa_defect_layout_issues(sheet))
        sheet.print_area = "A1:M44"
        sheet["N21"] = "Another defect"
        sheet["N22"] = 0.01
        apply_elegant_borders(sheet, "Oct")
        self.assertTrue(any("vùng in" in issue["reason"] for issue in hfpa_defect_layout_issues(sheet)))
        sheet["AE22"] = 0.25
        issues = hfpa_defect_layout_issues(sheet)
        self.assertTrue(any(issue["cell"] == "AE22" for issue in issues))
        with self.assertRaisesRegex(ValueError, "AE22"):
            apply_elegant_borders(sheet, "Oct")
        self.assertEqual(sheet["AE22"].value, 0.25)

    def test_two_models_same_line(self):
        """Test two models manufactured on the same line: verify Pivot1_Validation preserves model-level quantities."""
        import pandas as pd
        from report_tools.hfpa.agents.reconciliation_engine_agent import ReconciliationEngineAgent

        qa_data = pd.DataFrame({
            "InspDate": ["2026-10-05", "2026-10-05"],
            "FTY": ["VH", "VH"],
            "Station": ["HFPA", "HFPA"],
            "Line": ["Line 1", "Line 1"],
            "ShoeName": ["Model Alpha", "Model Beta"],
            "InspQty": [80, 20],
            "FailQty": [4, 2],
            "Issues": ["Thread end", "Cleanness"],
            "IssueQty": [4, 2],
        })
        mes_data = pd.DataFrame({
            "FTY": ["VH", "VH"],
            "Audit Date": ["2026-10-05", "2026-10-05"],
            "Plant": ["", ""],
            "Line": ["Line 1", "Line 1"],
            "Model": ["Model Alpha", "Model Beta"],
            "Audit Sample Size": [80, 20],
            "Defective Q'ty(pair)": [4, 2],
            "Defect Q'ty (Total Defect Q'ty)": [4, 2],
        })

        with tempfile.TemporaryDirectory() as temp_dir:
            agent = ReconciliationEngineAgent(temp_dir)
            qa_val, lots, _ = agent.reconcile_and_validate(qa_data, mes_data)
            p1, _, top5_dict, _ = agent.generate_aggregations(qa_val, mes_data, {})

        # Pivot1 must contain 2 separate rows for Model Alpha and Model Beta
        p1_models = p1[p1["Line"] == "1"].sort_values("Shoename")
        self.assertEqual(len(p1_models), 2)
        alpha_row = p1_models[p1_models["Shoename"] == "Model Alpha"].iloc[0]
        beta_row = p1_models[p1_models["Shoename"] == "Model Beta"].iloc[0]
        self.assertEqual(alpha_row["InspQty"], 80)
        self.assertEqual(alpha_row["FailQty"], 4)
        self.assertEqual(beta_row["InspQty"], 20)
        self.assertEqual(beta_row["FailQty"], 2)

    def test_tie_breaking_ranks(self):
        """Test deterministic tie-breaking for both models and defects."""
        import pandas as pd
        from report_tools.hfpa.hfpa_common import compute_model_ranks_mes410, compute_top3_defects

        # Model tie-breaking:
        # Model 1 & Model 2: both have 10 fails, but Model 2 has larger sample (200 vs 100) -> Model 2 ranks higher
        # Model 3 & Model 4: identical fails (5) and samples (50) -> Alphabetical: 'Model A' ranks higher than 'Model Z'
        mes_tie = pd.DataFrame({
            "FTY": ["VH", "VH", "VH", "VH"],
            "Model": ["Model 1", "Model 2", "Model Z", "Model A"],
            "Audit Sample Size": [100, 200, 50, 50],
            "Defective Q'ty(pair)": [10, 10, 5, 5],
        })
        tot_s, tot_f, ranked_models = compute_model_ranks_mes410(mes_tie)
        ranked_names = ranked_models["Model"].tolist()
        self.assertEqual(ranked_names[0], "Model 2")
        self.assertEqual(ranked_names[1], "Model 1")
        self.assertEqual(ranked_names[2], "Model A")
        self.assertEqual(ranked_names[3], "Model Z")

        # Defect tie-breaking:
        # Equal issue counts retain Quality Tracking order.
        qa_def_tie = pd.DataFrame({
            "FTY": ["VH", "VH"],
            "ShoeName_Std": ["Model 1", "Model 1"],
            "Issues_Std": ["Defect B", "Defect A"],
            "IssueQty_Validated": [5, 5],
        })
        top5_dict = {"VH": pd.DataFrame({"Shoename": ["Model 1"], "Sum of InspQty": [300], "DR%": [0.03]})}
        p2_df = compute_top3_defects(qa_def_tie, top5_dict)
        def_ranks = p2_df["Defect Type"].tolist()
        self.assertEqual(def_ranks[0], "Defect B")
        self.assertEqual(def_ranks[1], "Defect A")

    def test_missing_data_handling(self):
        """Test handling of Top 5 model missing QAStation defects, with and without MES410 fallback."""
        import pandas as pd
        from report_tools.hfpa.hfpa_common import compute_top3_defects

        # QAStation has NO defect rows for 'Model Ghost'
        qa_empty = pd.DataFrame(columns=["FTY", "ShoeName_Std", "Issues_Std", "IssueQty_Validated"])
        top5_dict = {"VH": pd.DataFrame({"Shoename": ["Model Ghost"], "Sum of InspQty": [500], "DR%": [0.02]})}

        # Case 1: No MES410 fallback
        p2_missing = compute_top3_defects(qa_empty, top5_dict)
        self.assertEqual(len(p2_missing), 1)
        self.assertEqual(p2_missing.iloc[0]["Defect Rank"], "Thiếu dữ liệu (QAStation)")
        self.assertEqual(p2_missing.iloc[0]["Data Status"], "Missing_QAStation")

        # Case 2: With MES410 fallback providing defect breakdown
        mes_fallback = pd.DataFrame({
            "FTY": ["VH"],
            "Model": ["Model Ghost"],
            "Audit Sample Size": [500],
            "Defective Q'ty(pair)": [10],
            "Sole attachment": [6],
            "Produced Q'ty": [1000],
            "PO Nbr (header+line item nbr)": [123456],
            "Cleanness": [4],
        })
        p2_fallback = compute_top3_defects(qa_empty, top5_dict, mes410_fallback_df=mes_fallback)
        self.assertEqual(len(p2_fallback), 2)
        self.assertIn("MES410 Thay thế", p2_fallback.iloc[0]["Defect Rank"])
        self.assertEqual(p2_fallback.iloc[0]["Defect Type"], "Sole attachment")
        self.assertEqual(p2_fallback.iloc[0]["Defect Quantity"], 6)

    def test_source_discrepancy_and_separated_variance_units(self):
        """Test separated variance units (Insp_Variance, Fail_Pairs_Variance, Issue_Variance) and exact match rule."""
        import pandas as pd
        from report_tools.hfpa.agents.reconciliation_engine_agent import ReconciliationEngineAgent
        from report_tools.hfpa.hfpa_common import compare_defect_shares

        # Lot 1: InspQty mismatch (100 vs 95)
        # Lot 2: FailQty mismatch (5 vs 4)
        # Lot 3: IssueQty mismatch (7 vs 6)
        # Lot 4: Exact match on all 3
        qa_data = pd.DataFrame({
            "InspDate": ["2026-10-01", "2026-10-02", "2026-10-03", "2026-10-04"],
            "FTY": ["VH", "VH", "VH", "VH"],
            "Station": ["HFPA"] * 4,
            "Line": ["Line 1", "Line 2", "Line 3", "Line 4"],
            "ShoeName": ["M1", "M2", "M3", "M4"],
            "InspQty": [100, 100, 100, 100],
            "FailQty": [5, 5, 5, 5],
            "Issues": ["D1", "D2", "D3", "D4"],
            "IssueQty": [5, 5, 7, 5],
        })
        mes_data = pd.DataFrame({
            "FTY": ["VH", "VH", "VH", "VH"],
            "Audit Date": ["2026-10-01", "2026-10-02", "2026-10-03", "2026-10-04"],
            "Plant": ["", "", "", ""],
            "Line": ["Line 1", "Line 2", "Line 3", "Line 4"],
            "Model": ["M1", "M2", "M3", "M4"],
            "Audit Sample Size": [95, 100, 100, 100],          # Lot 1 sample diff (100 vs 95)
            "Defective Q'ty(pair)": [5, 4, 5, 5],              # Lot 2 fail pairs diff (5 vs 4)
            "Defect Q'ty (Total Defect Q'ty)": [5, 5, 6, 5],  # Lot 3 issue qty diff (7 vs 6)
        })

        with tempfile.TemporaryDirectory() as temp_dir:
            agent = ReconciliationEngineAgent(temp_dir)
            _, lots, _ = agent.reconcile_and_validate(qa_data, mes_data)

        # Check separated variances exist
        self.assertIn("Insp_Variance", lots.columns)
        self.assertIn("Fail_Pairs_Variance", lots.columns)
        self.assertIn("Issue_Variance", lots.columns)

        # Lot 1 (Line 1): Insp_Variance != 0 -> Different_Count
        lot1 = lots[lots["Line_Std"] == "1"].iloc[0]
        self.assertEqual(lot1["Insp_Variance"], 5)
        self.assertEqual(lot1["Validation_Status"], "Different_Count")

        # Lot 2 (Line 2): Fail_Pairs_Variance != 0 -> Different_Count
        lot2 = lots[lots["Line_Std"] == "2"].iloc[0]
        self.assertEqual(lot2["Fail_Pairs_Variance"], 1)
        self.assertEqual(lot2["Validation_Status"], "Different_Count")

        # Lot 3 (Line 3): Issue_Variance != 0 -> Different_Count
        lot3 = lots[lots["Line_Std"] == "3"].iloc[0]
        self.assertEqual(lot3["Issue_Variance"], 1)
        self.assertEqual(lot3["Validation_Status"], "Different_Count")

        # Lot 4 (Line 4): all 3 == 0 -> Matched_Exact
        lot4 = lots[lots["Line_Std"] == "4"].iloc[0]
        self.assertEqual(lot4["Insp_Variance"], 0)
        self.assertEqual(lot4["Fail_Pairs_Variance"], 0)
        self.assertEqual(lot4["Issue_Variance"], 0)
        self.assertEqual(lot4["Validation_Status"], "Matched_Exact")

        # Test compare_defect_shares precision
        self.assertTrue(compare_defect_shares(0.164210, 0.1642, display_decimals=4))
        self.assertFalse(compare_defect_shares(0.164210, 0.1650, display_decimals=4))


if __name__ == "__main__":
    unittest.main()
