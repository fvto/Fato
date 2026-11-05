"""Inspect deliverables and reconcile tool-specific invariants before downloads."""

from decimal import Decimal
from pathlib import Path

import openpyxl
from report_brief import copq_issues, hfpa_issues, summarize_checks


def check_outputs(folder, run, files):
    output = Path(folder) / "workspace" / "Output"
    checks = list(run.get("business_review", []))

    def check(rule, passed, message, **metadata):
        checks.append(
            dict(
                rule=rule,
                status="passed" if passed else "warning",
                message=message,
                **metadata,
            )
        )

    names = [item["name"] for item in files]
    from sfc_report_tools import DELIVERABLES, TOOLS

    for path in output.rglob("*"):
        if (
            path.is_file()
            and path.suffix.lower() in DELIVERABLES
            and path.relative_to(output).as_posix() not in names
            and not path.name.startswith(("~$", "."))
        ):
            check(
                "file.integrity",
                False,
                "File chưa hoàn chỉnh/không đọc được; không cho tải: " + path.name,
            )
    mode, tool = run["mode"], run["tool"]
    for extension in TOOLS[tool].get("expected_extensions", {}).get(mode, []):
        check(
            "deliverable" + extension,
            any(n.endswith(extension) for n in names),
            "Có file kết quả đúng định dạng " + extension + ".",
        )
    inputs = run.get("validation", [])
    for item in files:
        if not item["name"].endswith(".xlsx"):
            continue
        book = openpyxl.load_workbook(
            output / item["name"], read_only=True, data_only=False
        )
        try:
            item["sheets"] = [
                dict(name=ws.title, rows=ws.max_row, columns=ws.max_column)
                for ws in book
            ]
            if tool == "copq" and "COPQ_Clean" in book.sheetnames:
                ws = book["COPQ_Clean"]
                rows = ws.iter_rows(values_only=True)
                header = list(next(rows, ()))
                indexed_records = [(index, row) for index, row in enumerate(rows, 2)
                                   if any(v is not None for v in row)]
                records = [row for _, row in indexed_records]
                expected = sum(
                    v["rows"] - v.get("business_exclusions", {}).get("rework_rows", 0)
                    for v in inputs
                    if v["group"] == "copq"
                )
                check(
                    "copq.row_count",
                    len(records) == expected,
                    "Đối chiếu số bản ghi COPQ → COPQ_Clean (loại Rework theo quy tắc cũ).",
                    expected=expected,
                    actual=len(records),
                )
                if "Ttl cost ($)" in header:
                    column = header.index("Ttl cost ($)")
                    actual = sum(
                        (Decimal(str(r[column] or 0)) for r in records), Decimal(0)
                    )
                    expected_cost = sum(
                        Decimal(str(v.get("totals", {}).get("Ttl cost ($)", 0)))
                        - Decimal(
                            v.get("business_exclusions", {}).get("rework_cost", "0")
                        )
                        for v in inputs
                        if v["group"] == "copq"
                    )
                    check(
                        "copq.cost",
                        abs(actual - expected_cost) <= Decimal("0.01"),
                        "Đối chiếu tổng chi phí COPQ.",
                        expected=str(expected_cost),
                        actual=str(actual),
                    )
                if "Data Quality Status" in header:
                    column = header.index("Data Quality Status")
                    mismatches = sum(
                        str(r[column]) not in ("OK", "Auto-Resolved") for r in records
                    )
                    check(
                        "copq.quality",
                        mismatches == 0,
                        "Kiểm tra dữ liệu thiếu hoặc vi phạm quy tắc COPQ.",
                        mismatches=mismatches,
                        issues=copq_issues(header, indexed_records, item["name"]),
                    )
            if tool == "copq" and "Combined_All" in book.sheetnames:
                ws = book["Combined_All"]
                records = list(ws.iter_rows(values_only=True))
                actual = max(0, len(records) - 1)
                expected = sum(
                    v.get("bc_rows", 0) for v in inputs if v["group"] == "ftt"
                )
                check(
                    "ftt.bc_rows",
                    actual == expected,
                    "Đối chiếu số dòng B/C được lọc theo quy tắc FTT.",
                    expected=expected,
                    actual=actual,
                )
                if records and "FTT_Qty" in records[0]:
                    column = list(records[0]).index("FTT_Qty")
                    actual_quantity = sum(
                        Decimal(str(row[column] or 0)) for row in records[1:]
                    )
                    expected_quantity = sum(
                        Decimal(v.get("bc_quantity", "0"))
                        for v in inputs
                        if v["group"] == "ftt"
                    )
                    check(
                        "ftt.bc_quantity",
                        actual_quantity == expected_quantity,
                        "Đối chiếu tổng số lượng lỗi B/C.",
                        expected=str(expected_quantity),
                        actual=str(actual_quantity),
                    )
            if tool == "hfpa" and "Lot_Defect_Variance" in book.sheetnames:
                rows = book["Lot_Defect_Variance"].iter_rows(values_only=True)
                header = list(next(rows, ()))
                if "Validation_Status" in header:
                    column = header.index("Validation_Status")
                    indexed_records = list(enumerate(rows, 2))
                    count = sum(row[column] != "Matched_Exact" for _, row in indexed_records)
                    check(
                        "hfpa.lot_variance",
                        count == 0,
                        "Đối chiếu lot QAStation và MES410; xem sheet Lot_Defect_Variance.",
                        mismatches=count,
                        issues=hfpa_issues(header, indexed_records, item["name"]),
                    )
                for group, sheet in [
                    ("ftt", "HFPA_Quality_Tracking"),
                    ("hfpa", "HFPA_Mes410"),
                ]:
                    expected = sum(v["rows"] for v in inputs if v["group"] == group)
                    actual = sum(
                        1
                        for row in book[sheet].iter_rows(min_row=2, values_only=True)
                        if any(v is not None for v in row)
                    )
                    check(
                        "hfpa." + group + ".rows",
                        expected == actual,
                        "Đối chiếu số dòng nguồn HFPA trong master.",
                        expected=expected,
                        actual=actual,
                    )
            if tool == "reinspection":
                required = (
                    {"VH", "VH2", "JV", "JV2"}
                    if "Recycle" in item["name"]
                    else {"re-vh", "RE-VH2", "RE-JV", "RE-JV2", "QC", "MA"}
                )
                check(
                    "reinspection.sheets",
                    {v.lower() for v in required}
                    <= {v.lower() for v in book.sheetnames},
                    "Kiểm tra đủ sheet nhà máy/phân loại trong Re-Inspection.",
                )
                if "Recycle" not in item["name"]:
                    from sfc_report_tools import plant_of

                    for source in inputs:
                        if source["group"] != "reinspection":
                            continue
                        plant = plant_of(source["name"])
                        sheet = "RE-" + plant
                        if sheet in book:
                            expected = source.get("category_counts", {}).get("QC", 0)
                            actual = book[sheet]["H6"].value
                            check(
                                "reinspection." + plant + ".qc_count",
                                actual == expected
                                if expected
                                else actual in (None, 0, "=H6"),
                                "Đối chiếu số lượt QC Re-Inspection · " + plant + ".",
                                expected=expected,
                                actual=actual,
                            )
        finally:
            book.close()
    # A missing expected business table must be visible, never a silent pass.
    required_rules = []
    if tool == "reinspection" and mode != "pptx":
        required_rules.append("reinspection.sheets")
    if tool == "copq":
        if mode in ("all", "combine"):
            required_rules.append("copq.row_count")
        if mode in ("all", "ftt"):
            required_rules.append("ftt.bc_rows")
    if tool == "hfpa" and mode != "check":
        required_rules.append("hfpa.lot_variance")
        required_rules.append("hfpa.model_review")
        if mode in ("all", "excel"):
            try:
                import sys
                import calendar
                tools_root = Path(__file__).resolve().parent / "report_tools" / "hfpa"
                if str(tools_root) not in sys.path:
                    sys.path.insert(0, str(tools_root))
                from hfpa_common import verify_hfpa_deliverables_consistency
                target_period = run.get("period", "")
                p_month, p_year = None, None
                if target_period and "-" in target_period:
                    y_str, m_str = target_period.split("-")
                    p_year = y_str
                    p_month = calendar.month_abbr[int(m_str)]
                hfpa_res = verify_hfpa_deliverables_consistency(
                    output, target_month=p_month, target_year=p_year, include_pptx=mode == "all"
                )
                for chk in hfpa_res.get("checks", []):
                    checks.append(chk)
                required_rules.extend([
                    "hfpa.master.tables",
                    "hfpa.pivot_rates",
                    "hfpa.excel.formulas",
                    "hfpa.excel.layout",
                    "hfpa.total_sample",
                    "hfpa.total_defects",
                    "hfpa.top5_ranks",
                    "hfpa.top5_dr",
                    "hfpa.defect_shares",
                ])
                if mode == "all":
                    required_rules.extend(["hfpa.pptx.charts", "hfpa.pptx.links"])
            except Exception as e:
                check("hfpa.cross_deliverables", False, f"Lỗi đối chiếu deliverable HFPA: {e}")
    for rule in required_rules:
        if not any(c["rule"] == rule for c in checks):
            check(rule, False, "Thiếu bảng kết quả để đối chiếu quy tắc " + rule + ".")
    warnings = sum(c["status"] == "warning" for c in checks)
    return dict(
        status="warning" if warnings else "passed",
        checks=checks,
        warning_count=warnings,
        brief=summarize_checks(checks),
        scope="Kiểm tra tính toàn vẹn, số dòng/tổng theo tool; không xác nhận hình thức hay ý nghĩa mọi số liệu trên slide.",
    )
