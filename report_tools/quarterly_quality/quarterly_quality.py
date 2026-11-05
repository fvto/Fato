"""Build quarterly quality metrics from the workbook's monthly counts."""

from calendar import month_abbr
from dataclasses import dataclass, field
from decimal import InvalidOperation
import math
import re

from openpyxl.utils import get_column_letter

from report_validation import PreflightError, number, tables


PLANTS = ("VH", "VH2", "JV", "JV2")
REGIONS = {"VN": PLANTS[:2], "Indo": PLANTS[2:], "CLG": PLANTS}
SCOPES = (*PLANTS, *REGIONS)
BLOCKS = {
    "HFPA": ("Month HFPA", 4, 5, 14, "Inspection", "Defect"),
    "FTT": ("FTT month-quarter", 1, 2, 24, "Total pair produced", "Total Defect"),
}
QUARTER_BLOCKS = {
    "HFPA": ("Quarter HFPA", 1, 3, 3, 12),
    "FTT": ("FTT month-quarter", 55, 4, 57, 66),
}


@dataclass(frozen=True)
class Counts:
    inspected: float
    defects: float
    grade: float | None = None
    rework: float | None = None

    @property
    def rate(self):
        return 1 - self.defects / self.inspected if self.inspected else None


def quarter_months(period):
    fiscal_year, quarter = int(period[2:4]), int(period[-1])
    start = (1999 + fiscal_year) * 12 + 5 + (quarter - 1) * 3
    return tuple(f"{index // 12:04d}-{index % 12 + 1:02d}"
                 for index in range(start - 3, start + 3))


def cell_at(rows, row, col):
    return rows[row - 1][col - 1] if row <= len(rows) and col <= len(rows[row - 1]) else None


@dataclass
class Quarter:
    period: str
    prior_months: tuple
    months: tuple
    covered: tuple
    values: dict
    locations: dict
    controls: dict
    date_from: str | None = None
    date_to: str | None = None
    prior_date_from: str | None = None
    prior_date_to: str | None = None
    rankings: dict = field(default_factory=dict)
    protected: set = field(default_factory=set)
    unmatched_rankings: set = field(default_factory=set)
    source_differences: list = field(default_factory=list)

    def totals(self, metric, scope, prior=False):
        months = self.prior_months[:len(self.covered)] if prior else self.covered
        records = [self.values.get((metric, month, scope)) for month in months]
        if records and all(record is not None for record in records):
            return Counts(sum(record.inspected for record in records),
                          sum(record.defects for record in records))
        return None


def read_quarter(source, period, sheet_summaries):
    """Allow a quarter in progress with a complete prefix of one or two months."""
    months = quarter_months(period)
    prior, current = months[:3], months[3:]
    values, locations, errors, seen = {}, {}, [], set()
    source = {name.strip(): rows for name, rows in source.items()}
    for metric, (name, header, inspected_row, defects_row, inspected_label, defects_label) in BLOCKS.items():
        rows = source.get(name, [])

        def cell(row, col):
            return cell_at(rows, row, col)

        def issue(message, row=None, col=None):
            errors.append((message, name, row, get_column_letter(col) if col else None))

        if not rows:
            issue("Thiếu sheet dữ liệu tháng theo mẫu.")
            continue
        previous_errors = len(errors)
        for offset, plant in enumerate(PLANTS):
            for start, label in ((inspected_row, inspected_label), (defects_row, defects_label)):
                row = start + offset
                if (cell(row, 1), cell(row, 2)) != (plant, label):
                    issue(f"Cần dòng {plant} / {label} theo mẫu.", row, 1)
        if len(errors) != previous_errors:
            continue

        columns = {}
        year = None
        if metric == "FTT":
            # In the reference, the first June belongs to Q1 of the first FY (D55).
            fiscal_year = re.fullmatch(r"FY(\d{2})", str(cell(55, 4)).strip())
            if not fiscal_year or cell(1, 3) != "Jun":
                issue("Cần tháng Jun ở C1 và năm tài chính FYxx ở D55 để xác định năm FTT.", 55, 4)
                continue
            year = 1999 + int(fiscal_year[1])
        previous_month = None
        for col in range(3, len(rows[header - 1]) + 1):
            label = cell(header, col)
            if label not in tuple(month_abbr)[1:]:
                break
            month = tuple(month_abbr).index(label)
            if metric == "HFPA" and cell(1, col) is not None:
                anchor = re.fullmatch(r"20\d{2}", str(cell(1, col)))
                if not anchor or month != 1 or (year is not None and int(anchor[0]) != year + 1):
                    issue("Năm HFPA cần nằm trên cột Jan tương ứng.", 1, col)
                    break
                year = int(anchor[0])
            elif month == 1 and previous_month == 12:
                if metric == "FTT":
                    year += 1
                else:
                    issue("Thiếu năm dương lịch phía trên cột Jan.", 1, col)
                    break
            if year is None or (previous_month is not None and month != previous_month % 12 + 1):
                issue("Các cột tháng phải liên tiếp và có năm xác định.", header, col)
                break
            columns[f"{year:04d}-{month:02d}"] = col
            previous_month = month
        count = 0
        for month in months:
            if month not in columns:
                continue
            col = columns[month]
            for offset, plant in enumerate(PLANTS):
                base_row, loss_row = inspected_row + offset, defects_row + offset
                base, loss = cell(base_row, col), cell(loss_row, col)
                if all(value in (None, "") for value in (base, loss)):
                    continue
                seen.add(month)
                try:
                    inspected = number(base)
                    if inspected <= 0:
                        raise ValueError
                except (InvalidOperation, ValueError):
                    if month in current:
                        issue(f"{plant} cần số lượng kiểm tra/sản xuất lớn hơn 0.", base_row, col)
                    continue
                grade = rework = None
                if metric == "FTT":
                    grade_value, rework_value = cell(9 + offset, col), cell(17 + offset, col)
                    if grade_value is not None or rework_value is not None:
                        try:
                            grade, rework = number(grade_value), number(rework_value)
                            if grade > inspected or rework > grade:
                                raise ValueError
                        except (InvalidOperation, ValueError):
                            if month in current:
                                issue(f"{plant}: A Grade và Total Rework cần hợp lệ và không vượt số đôi sản xuất.",
                                      9 + offset, col)
                            continue
                try:
                    defects = inspected - grade + rework if grade is not None else number(loss)
                    if defects > inspected:
                        raise ValueError
                except (InvalidOperation, ValueError):
                    if month in current:
                        issue(f"{plant} cần số lỗi từ 0 đến số lượng kiểm tra/sản xuất.", loss_row, col)
                    continue
                key = (metric, month, plant)
                values[key] = Counts(
                    float(inspected), float(defects),
                    float(grade) if grade is not None else None,
                    float(rework) if rework is not None else None,
                )
                locations[key] = f"{get_column_letter(col)}{base_row}, {get_column_letter(col)}{loss_row}"
                count += 1
            for scope, plants in REGIONS.items():
                records = [values.get((metric, month, plant)) for plant in plants]
                if all(record is not None for record in records):
                    values[metric, month, scope] = Counts(
                        sum(record.inspected for record in records),
                        sum(record.defects for record in records),
                    )
        sheet_summaries.append(dict(name=name, header_row=header, rows=count))

    covered = tuple(month for month in current if month in seen and
                    all((metric, month, "CLG") in values for metric in BLOCKS))
    if not covered:
        errors.append((f"{period} chưa có tháng đủ dữ liệu HFPA và FTT cho 4 nhà máy.", None, None, None))
    for month in current:
        if month in seen and month not in covered:
            errors.append((f"Tháng {month} thiếu dữ liệu hợp lệ của một chỉ số hoặc nhà máy.", None, None, None))
    if covered != current[:len(covered)]:
        errors.append(("Dữ liệu quý phải có từ tháng đầu và liên tiếp đến tháng gần nhất.", None, None, None))
    quarter = Quarter(period, prior, current, covered, values, locations, {})
    if not errors:
        quarter.controls = quarter_controls(source, quarter)
    return quarter, errors


def quarter_controls(source, quarter):
    """Compare recomputed counts with any existing fiscal quarter column."""
    controls = {}
    for metric, (sheet, year_row, first_col, base_row, loss_row) in QUARTER_BLOCKS.items():
        rows = source.get(sheet, [])
        if len(rows) < loss_row + len(PLANTS) - 1:
            continue
        year = None
        for col, label in enumerate(rows[year_row][first_col - 1:], first_col):
            year = cell_at(rows, year_row, col) or year
            if f"{year}-{label}" != quarter.period:
                continue
            for offset, plant in enumerate(PLANTS):
                source_values = tuple(cell_at(rows, row, col) for row in (base_row + offset, loss_row + offset))
                actual = quarter.totals(metric, plant)
                for name, row, expected, observed in (
                    ("Kiểm tra / sản xuất", base_row + offset, source_values[0], actual.inspected),
                    ("Lỗi", loss_row + offset, source_values[1], actual.defects),
                ):
                    expected = expected if type(expected) in (int, float) else None
                    controls[metric, plant, name] = (
                        f"{sheet}!{get_column_letter(col)}{row}", expected, observed,
                    )
            break
    return controls


def control_status(expected, observed):
    if expected is None:
        return "Thiếu số tổng"
    return "Khớp" if math.isclose(expected, observed, rel_tol=0, abs_tol=0.000001) else "Chênh lệch"


def read_excel_rankings(source, quarter):
    source = {name.strip(): rows for name, rows in source.items()}
    for metric, title, first, layout in (
        ("HFPA", "Quarter HFPA", 14, (("Indo", 4, 14), ("VN", 7, 17))),
        ("FTT", "FTT month-quarter", 4, (("Indo", 38, 48), ("VN", 41, 51))),
    ):
        rows = source.get(title, [])
        for region, model_start, defect_start in layout:
            groups = []
            for start in (model_start, defect_start):
                ranks = []
                for row in range(start, start + 3):
                    name, base, bad, rate = (cell_at(rows, row, col) for col in range(first, first + 4))
                    if not isinstance(name, str) or not name.strip():
                        continue
                    if type(base) in (float, int) and base > 0 and type(bad) in (float, int) and 0 <= bad <= base:
                        ranks.append((name.strip(), bad, base))
                    elif type(rate) in (float, int) and 0 <= rate <= 1:
                        ranks.append((name.strip(), rate, None))
                groups.append(ranks)
            quarter.rankings[metric, region] = tuple(groups)
    note = str(cell_at(source.get("Quarter HFPA", []), 28, 1) or "")
    cutoff = re.search(rf"\b{re.escape(quarter.period)} (?:KPI through|SFC fills through) (\d{{4}}-\d{{2}}-\d{{2}})\b", note)
    if cutoff and cutoff[1][:7] == quarter.covered[-1]:
        from datetime import date
        try:
            quarter.date_to = date.fromisoformat(cutoff[1]).isoformat()
            quarter.date_from = quarter.months[0] + "-01"
        except ValueError:
            pass


def run_report(workspace, run, stage):
    from .template_output import template_paths, write_excel, write_powerpoint

    excel_template, ppt_template = template_paths()
    if run.get("acquire_missing"):
        from .sfc_inputs import load_sfc_quarter

        baseline, baseline_errors = read_quarter(dict(tables(excel_template, data_only=True)), run["period"], [])
        issues = [dict(message=message, sheet=sheet, row=row, column=column)
                  for message, sheet, row, column in baseline_errors if sheet]
        if issues:
            report = dict(name=excel_template.name, group="workbook", sheets=[], rows=0,
                          valid=False, error_count=len(issues), issues=issues)
            run["validation"] = [report]
            raise PreflightError([report])
        quarter = load_sfc_quarter(workspace, run, stage, baseline=baseline)
        stage("Đang tổng hợp KPI quý từ dữ liệu SFC...", "processing")
    else:
        item, = run["inputs"]
        source = item["name"]
        path = workspace / "Input" / item.get("staged_name", source)
        source_tables = dict(tables(path, data_only=True))
        quarter, errors = read_quarter(source_tables, run["period"], [])
        if errors:
            raise ValueError("Workbook không đủ dữ liệu quý đã chọn.")
        read_excel_rankings(source_tables, quarter)
    output = workspace / "Output"
    basename = f"HFPA-FTT-{run['period']}"
    write_excel(excel_template, output / f"{basename}.xlsx", quarter)
    run["business_review"] = []
    for metric in BLOCKS:
        differences = run.get(f"{metric.lower()}_baseline_differences", [])
        if not differences:
            continue
        quantity_label = "kiểm" if metric == "HFPA" else "sản xuất"
        run["business_review"].append(dict(
            rule=f"quarter.{metric.lower()}.template_differences", status="warning",
            message=f"Dùng số {metric} từ dữ liệu SFC đã kiểm tra cho {len(differences)} nhóm khác số trong mẫu Excel.",
            issues=[dict(title=f"{metric} · {r['month']} · {r['plant']}",
                         reason=(f"Mẫu Excel: {r['template_insp']:g} đôi {quantity_label}, {r['template_fail']:g} đôi lỗi. "
                                 f"SFC đã kiểm tra: {r['corrected_insp']:g} đôi {quantity_label}, {r['corrected_fail']:g} đôi lỗi."),
                         severity="reconcile") for r in differences],
        ))
    if run.get("hfpa_model_review"):
        review = run["hfpa_model_review"]
        run["business_review"].append(dict(
            rule="quarter.hfpa.model_review", status="warning",
            message=f"Có {len(review)} nhóm HFPA cần đối chiếu thêm. Báo cáo dùng số QAStation của các nhóm này.",
            issues=[dict(
                title=f"{r['FTY']} · {r['Date_Std']} · {r['Plant_Std']} Line {r['Line_Std']} · {r['ShoeName_Std']}",
                reason=("Thiếu dòng MES410 cùng khóa." if r["Status"] == "No_MES410_Model" else
                        "Thông tin chi tiết MES410 chưa đủ để hiệu chỉnh tự động."),
                severity="reconcile",
            ) for r in review],
        ))
    if quarter.source_differences:
        count = len(quarter.source_differences)
        run["business_review"].append(dict(
            rule="quarter.protected_input", status="warning",
            message=f"Giữ số Excel gốc ở {count} nhóm chỉ số/tháng/nhà máy có chênh lệch với SFC.",
            issues=[dict(title=f"{metric} · {month} · {plant}", reason=(
                f"Excel: {original.inspected:g} kiểm tra/sản xuất, {original.defects:g} lỗi. "
                f"SFC: {sfc.inspected:g} kiểm tra/sản xuất, {sfc.defects:g} lỗi."),
                action=(f"Giữ số Excel gốc. Chi tiết SFC khớp số Excel đến {matched}." if matched else
                        "Giữ số Excel gốc. Chưa đối chiếu khớp chi tiết SFC với số này."),
                severity="reconcile")
                for metric, month, plant, original, sfc, matched in quarter.source_differences]))
    if quarter.unmatched_rankings:
        areas = ", ".join(f"{metric} {region}" for metric, region in sorted(quarter.unmatched_rankings))
        run["business_review"].append(dict(
            rule="quarter.ranking_source", status="warning",
            message=(f"Top model / Top lỗi chưa đối chiếu khớp cho {areas}. "
                     "Giữ bảng có sẵn trong Excel và để trống biểu đồ tương ứng trong PowerPoint.")))
    for metric in BLOCKS:
        if run.get("acquire_missing"):
            current = quarter.totals(metric, "CLG")
            has_quantity = current.rate is not None
            run["business_review"].append(dict(
                rule=f"quarter.{metric.lower()}.sfc", status="passed" if has_quantity else "warning",
                message=(f"{metric}: đã kiểm tra file đúng nhà máy và khoảng ngày SFC. "
                         f"CLG {current.rate:.2%} từ {quarter.date_from} đến {quarter.date_to}."
                         if has_quantity else f"{metric}: các file SFC có mẫu số bằng 0 trong kỳ báo cáo.")))
            continue
        results = [control_status(expected, observed) for key, (_, expected, observed) in quarter.controls.items()
                   if key[0] == metric]
        passed = bool(results) and all(result == "Khớp" for result in results)
        message = f"{metric}: tổng theo tháng khớp bảng tổng quý nguồn." if passed else (
            f"{metric}: bảng tổng quý nguồn thiếu số hoặc có chênh lệch. Báo cáo dùng số được tính lại từ từng tháng.")
        run["business_review"].append(dict(rule=f"quarter.{metric.lower()}", status="passed" if passed else "warning",
                                           message=message))
    stage("Đang cập nhật biểu đồ trên mẫu PowerPoint...", "presentation")
    write_powerpoint(ppt_template, output / f"{basename}.pptx", quarter)
