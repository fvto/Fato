"""Update the boss's original OOXML templates without rebuilding their layouts."""

from calendar import month_abbr
from copy import deepcopy
from datetime import date
from io import BytesIO
import json
import math
import os
from pathlib import Path
import re
import textwrap
from zipfile import ZipFile

from lxml import etree
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.label import DataLabelList

from dashboard_paths import APP_ROOT
from report_tools.hfpa.agents.color_manager import ColorManager
from .quarterly_quality import PLANTS, REGIONS, Counts


X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
KEEP_FORMULA = object()
EXCEL_NAMES = ("Month HFPA", "Quarter HFPA ", "FTT month-quarter")
EXCEL_SCOPES = (*PLANTS, "CLG", "VN", "Indo")
SHAPES = {
    "FTT": dict(title=4, period=14, VN=(2, 6, 7), Indo=(5, 3, 13)),
    "HFPA": dict(title=5, period=6, VN=(2, 8, 10), Indo=(3, 11, 12)),
}


def template_paths():
    config = APP_ROOT / "report_tools_config.json"
    options = json.loads(config.read_text(encoding="utf-8")).get("quarterly_templates", {}) if config.exists() else {}
    roots = (APP_ROOT / "Report" / "Quarterly", Path.home() / "Downloads")
    paths = []
    for kind, name in (("excel", "Quarterly Report HFPA-FTT.xlsx"),
                       ("powerpoint", "Partner QBR Format Q2 FY27-CLG.pptx")):
        default = next((root / name for root in roots if (root / name).is_file()), roots[0] / name)
        path = Path(options.get(kind, default)).expanduser()
        paths.append(path if path.is_absolute() else APP_ROOT / path)
    for path in paths:
        if not path.is_file():
            raise ValueError(f"Thiếu mẫu báo cáo quý {path.name}. Đặt mẫu trong Report/Quarterly hoặc cấu hình quarterly_templates.")
    return tuple(paths)


def verify_template_quarter(period):
    excel, powerpoint = template_paths()
    book = load_workbook(excel, read_only=True, data_only=True)
    found = False
    try:
        if tuple(book.sheetnames) != EXCEL_NAMES:
            raise ValueError("Mẫu Excel cần đúng ba sheet theo format QBR.")
        sheet = book[EXCEL_NAMES[1]]
        years, labels = sheet.iter_rows(min_row=1, max_row=2, min_col=3, values_only=True)
        year = None
        for anchor, label in zip(years, labels):
            year = anchor or year
            if f"{year}-{label}" == period:
                found = True
                break
    finally:
        book.close()
    if not found:
        raise ValueError(f"Mẫu Excel hiện không có cột {period}. Cần mẫu có sẵn quý này.")
    quality_slides(Presentation(powerpoint))


def copy_package(source, output, parts):
    with ZipFile(source) as original, ZipFile(output, "w") as result:
        existing = set(original.namelist())
        for entry in original.infolist():
            result.writestr(entry, parts.get(entry.filename, original.read(entry.filename)))
        for name, data in parts.items():
            if name not in existing:
                result.writestr(name, data)


def month_columns(sheet, first, header, anchor):
    year, previous = None, None
    columns = {}
    for col in range(first, sheet.max_column + 1):
        label = sheet.cell(header, col).value
        if label not in tuple(month_abbr)[1:]:
            break
        month = tuple(month_abbr).index(label)
        if anchor is not None:
            year = anchor
            if month == 1 and previous == 12:
                year += 1
                anchor = year
        elif sheet.cell(1, col).value:
            year = int(sheet.cell(1, col).value)
        columns[f"{year:04d}-{month:02d}"] = col
        previous = month
    return columns


class WorksheetPatch:
    def __init__(self, xml):
        self.root = etree.fromstring(xml)
        self.rows = {int(row.get("r")): row for row in self.root.find(X + "sheetData")}
        self.cells = {cell.get("r"): cell for row in self.rows.values() for cell in row.findall(X + "c")}
        self.shared_anchors = {}
        self.shared_members = {}
        for address, cell in self.cells.items():
            formula = cell.find(X + "f")
            if formula is not None and formula.get("t") == "shared":
                index = formula.get("si")
                self.shared_members.setdefault(index, []).append(address)
                if formula.text:
                    self.shared_anchors[index] = (address, formula.text)

    def formula_at(self, address):
        formula = self.cells[address].find(X + "f")
        if formula is None or formula.text or formula.get("t") != "shared":
            return formula.text if formula is not None else None
        anchor, expression = self.shared_anchors[formula.get("si")]
        return Translator("=" + expression, origin=anchor).translate_formula(address)[1:]

    def set(self, row, col, value, formula=KEEP_FORMULA):
        address = f"{get_column_letter(col)}{row}"
        cell = self.cells.get(address)
        if cell is None:
            if value is None:
                return
            parent = self.rows[row]
            cell = etree.Element(X + "c", r=address)
            next_cell = next((candidate for candidate in parent.findall(X + "c")
                              if column_index_from_string(re.match(r"[A-Z]+", candidate.get("r"))[0]) > col), None)
            if next_cell is None:
                parent.append(cell)
            else:
                next_cell.addprevious(cell)
            self.cells[address] = cell
        old_formula = cell.find(X + "f")
        if formula is not KEEP_FORMULA and old_formula is not None and old_formula.get("ref"):
            # A replaced shared-formula anchor must not orphan untouched cells.
            for member in self.shared_members.get(old_formula.get("si"), ()):
                if member == address:
                    continue
                member_formula = self.cells[member].find(X + "f")
                if member_formula is not None and member_formula.get("t") == "shared":
                    member_formula.text = self.formula_at(member)
                    member_formula.attrib.clear()
        for child in tuple(cell):
            if child.tag in (X + "v", X + "is") or child.tag == X + "f" and formula is not KEEP_FORMULA:
                cell.remove(child)
        cell.attrib.pop("t", None)
        if formula is not KEEP_FORMULA and formula:
            etree.SubElement(cell, X + "f").text = formula
        if isinstance(value, str):
            cell.set("t", "inlineStr")
            etree.SubElement(etree.SubElement(cell, X + "is"), X + "t").text = value
        elif value is not None:
            etree.SubElement(cell, X + "v").text = format(value, ".15g")
        elif cell.find(X + "f") is not None:
            cell.set("t", "str")
            etree.SubElement(cell, X + "v")


def aggregate_block(patch, sheet, quarter, metric, columns, source_columns):
    """Update the selected quarter and the existing rightmost Sum formulas."""
    hfpa = metric == "HFPA"
    label_row, first, base, bad, rate = (2, 3, 3, 12, 21) if hfpa else (56, 4, 57, 66, 75)
    year = None
    periods = {}
    for col in range(first, sheet.max_column + 1):
        year = sheet.cell(label_row - 1, col).value or year
        label = sheet.cell(label_row, col).value
        if label == "Sum":
            sum_col = col
            break
        periods[f"{year}-{label}"] = col
    selected = periods.get(quarter.period)
    if selected is None:
        raise ValueError(f"Mẫu Excel hiện không có cột {quarter.period}. Cần mẫu có sẵn quý này.")
    start, end = (get_column_letter(columns[month]) for month in (quarter.covered[0], quarter.covered[-1]))
    prefix = "'Month HFPA'!" if hfpa else ""
    current_col = get_column_letter(selected)
    sum_letter = get_column_letter(sum_col)
    for i, scope in enumerate(EXCEL_SCOPES):
        record = quarter.totals(metric, scope)
        for offset, source_row, value in ((base, source_columns[0], record.inspected), (bad, source_columns[1], record.defects)):
            row = offset + i
            if hfpa or i < len(PLANTS):
                patch.set(row, selected, value, f"SUM({prefix}{start}{source_row+i}:{end}{source_row+i})")
            else:
                patch.set(row, selected, value)
            aggregate = value + sum(sheet.cell(row, c).value or 0 for c in periods.values() if c != selected)
            patch.set(row, sum_col, aggregate)
        for col, letter, value in (
            (selected, current_col, record.rate),
            (sum_col, sum_letter, None),
        ):
            if col == sum_col:
                counts = [record.inspected, record.defects]
                for index, row in enumerate((base+i, bad+i)):
                    counts[index] += sum(sheet.cell(row, c).value or 0 for c in periods.values() if c != selected)
                value = 1 - counts[1] / counts[0] if counts[0] else None
            formula = f'IFERROR(1-{letter}{bad+i}/{letter}{base+i},"")'
            patch.set(rate+i, col, value, formula)


def write_excel(template, output, quarter):
    book = load_workbook(template, data_only=True)
    if tuple(book.sheetnames) != EXCEL_NAMES:
        raise ValueError("Mẫu Excel cần đúng ba sheet theo format QBR.")
    hfpa, hfpa_quarter, ftt = (book[name] for name in EXCEL_NAMES)
    mappings = {
        "HFPA": month_columns(hfpa, 3, 4, None),
        "FTT": month_columns(ftt, 3, 1, 1999 + int(ftt["D55"].value[2:])),
    }
    with ZipFile(template) as package:
        parts = ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml"]
        month_hfpa, summary_hfpa, month_ftt = (WorksheetPatch(package.read(name)) for name in parts)
        chart_name = "xl/charts/chart1.xml"
        chart = etree.fromstring(package.read(chart_name))
    for metric, monthly, monthly_sheet, summary, summary_sheet, rows in (
        ("HFPA", month_hfpa, hfpa, summary_hfpa, hfpa_quarter, (5, 14, 23)),
        ("FTT", month_ftt, ftt, month_ftt, ftt, (2, 24, 31)),
    ):
        columns = mappings[metric]
        if any(month not in columns for month in quarter.months):
            raise ValueError(f"Mẫu Excel thiếu cột tháng của {quarter.period}.")
        base, bad, rate = rows
        for month in quarter.covered:
            col = columns[month]
            letter = get_column_letter(col)
            for i, scope in enumerate(EXCEL_SCOPES if metric == "HFPA" else (*PLANTS, "CLG")):
                record = quarter.values.get((metric, month, scope))
                locked = (metric, month, scope) in quarter.protected
                if not locked:
                    monthly.set(base+i, col, record.inspected if record else None)
                    monthly.set(bad+i, col, record.defects if record else None)
                if metric == "FTT":
                    grades = [quarter.values.get((metric, month, p)) for p in (PLANTS if scope == "CLG" else (scope,))]
                    for start, attribute in ((9, "grade"), (17, "rework")):
                        amounts = [getattr(value, attribute) if value and record else None for value in grades]
                        if not locked:
                            monthly.set(start+i, col, sum(amounts) if all(v is not None for v in amounts) else None)
                    if scope == "CLG":
                        continue
                expression = (f"({letter}{9+i}-{letter}{17+i})/{letter}{base+i}"
                              if metric == "FTT" and record and record.grade is not None else
                              f"1-{letter}{bad+i}/{letter}{base+i}")
                monthly.set(rate+i, col, record.rate if record else None, f'IFERROR({expression},"")')
        # Leave approved input cells in other months intact, including any
        # month after an earlier selected cutoff. Blank zero-sample errors.
        for month, col in columns.items():
            if month in quarter.covered:
                continue
            letter = get_column_letter(col)
            for i in range(len(EXCEL_SCOPES) if metric == "HFPA" else len(PLANTS)):
                if monthly_sheet.cell(rate+i, col).value == "#DIV/0!" and not monthly_sheet.cell(base+i, col).value:
                    formula = monthly.formula_at(f"{letter}{rate+i}")
                    if formula:
                        monthly.set(rate+i, col, None, f'IFERROR({formula},"")')
        aggregate_block(summary, summary_sheet, quarter, metric, columns, rows)
        for region, model_start, defect_start in (("Indo", 4, 14), ("VN", 7, 17)) if metric == "HFPA" else (
            ("Indo", 38, 48), ("VN", 41, 51)
        ):
            if (metric, region) in quarter.unmatched_rankings:
                continue
            model, defects = quarter.rankings.get((metric, region), ([], []))
            first = 14 if metric == "HFPA" else 4
            for start, ranks in ((model_start, model), (defect_start, defects)):
                for i in range(3):
                    label, numerator, denominator = ranks[i] if i < len(ranks) else (None, None, None)
                    rate = numerator / denominator if denominator else numerator
                    for j, value in enumerate((label, denominator, numerator if denominator else None, rate)):
                        formula = (f'IFERROR(F{start+i}/E{start+i},"")'
                                   if metric == "FTT" and j == 3 and (denominator or rate is None) else None)
                        summary.set(start+i, first+j, value, formula)

    if quarter.date_to:
        historic_note = hfpa_quarter["A28"].value
        summary_hfpa.set(28, 1, f"{historic_note}  |  {quarter.period} SFC fills through {quarter.date_to}")

    # Refresh the chart's cached HFPA points so previews match the updated cells.
    extend_hfpa_chart(chart, quarter, mappings["HFPA"])
    book.close()
    updated = {name: etree.tostring(patch.root) for name, patch in zip(parts, (month_hfpa, summary_hfpa, month_ftt))}
    updated[chart_name] = etree.tostring(chart)
    copy_package(template, output, updated)
    add_excel_charts(output, quarter, mappings)


def add_excel_charts(output, quarter, mappings):
    """Add neatly arranged native Excel charts below the data tables (Style B)."""
    with ZipFile(output) as package:
        cached_sheets = {name: package.read(name) for name in package.namelist()
                         if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)}
    wb = load_workbook(output)
    
    # 1. Sheet: FTT month-quarter
    ftt_sheet = next((s for s in wb.sheetnames if s.strip() == "FTT month-quarter"), None)
    if ftt_sheet and quarter.months:
        ws = wb[ftt_sheet]
        ftt_cols = mappings.get("FTT", {})
        covered_months = [m for m in quarter.months if m in ftt_cols]
        if covered_months:
            start_col = ftt_cols[covered_months[0]]
            end_col = ftt_cols[covered_months[-1]]
            # The region series need real source cells for every selected
            # quarter, rather than references to empty rows 36 and 37.
            for col in range(start_col, end_col + 1):
                letter = get_column_letter(col)
                for row, offset in ((36, 0), (37, 2)):
                    ws.cell(row, col, (
                        f'=IFERROR(({letter}{9+offset}+{letter}{10+offset}'
                        f'-{letter}{17+offset}-{letter}{18+offset})/'
                        f'({letter}{2+offset}+{letter}{3+offset}),"")'))
            cats = Reference(ws, min_col=start_col, min_row=1, max_col=end_col, max_row=1)
            
            # Row 85: Line charts (Trend tracking)
            c_vn = LineChart()
            c_vn.title = "VN FTT% Trend Tracking"
            c_vn.style = 13
            c_vn.y_axis.title = "FTT%"
            c_vn.width = 15
            c_vn.height = 9.5
            c_vn.legend.legendPos = "t"
            c_vn.y_axis.number_format = "0.0%"
            for r in (36, 31, 32):
                lbl = ws.cell(r, 1).value or ("VN" if r == 36 else "VH" if r == 31 else "VH2")
                c_vn.series.append(Series(Reference(ws, min_col=start_col, min_row=r, max_col=end_col, max_row=r), title=lbl))
            c_vn.set_categories(cats)
            ws.add_chart(c_vn, "B85")

            c_indo = LineChart()
            c_indo.title = "INDO FTT% Trend Tracking"
            c_indo.style = 13
            c_indo.y_axis.title = "FTT%"
            c_indo.width = 15
            c_indo.height = 9.5
            c_indo.legend.legendPos = "t"
            c_indo.y_axis.number_format = "0.0%"
            for r in (37, 33, 34):
                lbl = ws.cell(r, 1).value or ("Indo" if r == 37 else "JV" if r == 33 else "JV2")
                c_indo.series.append(Series(Reference(ws, min_col=start_col, min_row=r, max_col=end_col, max_row=r), title=lbl))
            c_indo.set_categories(cats)
            ws.add_chart(c_indo, "J85")

            # Row 105: Top Models (VN & Indo)
            if ws.cell(41, 4).value:
                c_mod_vn = BarChart()
                c_mod_vn.type = "col"
                c_mod_vn.title = "Top FTT Models (VN)"
                c_mod_vn.style = 10
                c_mod_vn.width = 15
                c_mod_vn.height = 9.5
                c_mod_vn.legend = None
                c_mod_vn.y_axis.number_format = "0.0%"
                c_mod_vn.dataLabels = DataLabelList()
                c_mod_vn.dataLabels.showVal = True
                c_mod_vn.dataLabels.showSerName = False
                c_mod_vn.dataLabels.showCatName = False
                c_mod_vn.add_data(Reference(ws, min_col=7, min_row=41, max_col=7, max_row=43), titles_from_data=False)
                c_mod_vn.set_categories(Reference(ws, min_col=4, min_row=41, max_col=4, max_row=43))
                ws.add_chart(c_mod_vn, "B105")

            if ws.cell(38, 4).value:
                c_mod_id = BarChart()
                c_mod_id.type = "col"
                c_mod_id.title = "Top FTT Models (Indo)"
                c_mod_id.style = 12
                c_mod_id.width = 15
                c_mod_id.height = 9.5
                c_mod_id.legend = None
                c_mod_id.y_axis.number_format = "0.0%"
                c_mod_id.dataLabels = DataLabelList()
                c_mod_id.dataLabels.showVal = True
                c_mod_id.dataLabels.showSerName = False
                c_mod_id.dataLabels.showCatName = False
                c_mod_id.add_data(Reference(ws, min_col=7, min_row=38, max_col=7, max_row=40), titles_from_data=False)
                c_mod_id.set_categories(Reference(ws, min_col=4, min_row=38, max_col=4, max_row=40))
                ws.add_chart(c_mod_id, "J105")

            # Row 125: Top Defects of Top 1 Model (VN & Indo)
            if ws.cell(51, 4).value:
                c_def_vn = BarChart()
                c_def_vn.type = "col"
                c_def_vn.title = "Top Defects of Top 1 Model (VN)"
                c_def_vn.style = 11
                c_def_vn.width = 15
                c_def_vn.height = 9.5
                c_def_vn.legend = None
                c_def_vn.y_axis.number_format = "0.0%"
                c_def_vn.dataLabels = DataLabelList()
                c_def_vn.dataLabels.showVal = True
                c_def_vn.dataLabels.showSerName = False
                c_def_vn.dataLabels.showCatName = False
                c_def_vn.add_data(Reference(ws, min_col=7, min_row=51, max_col=7, max_row=53), titles_from_data=False)
                c_def_vn.set_categories(Reference(ws, min_col=4, min_row=51, max_col=4, max_row=53))
                ws.add_chart(c_def_vn, "B125")

            if ws.cell(48, 4).value:
                c_def_id = BarChart()
                c_def_id.type = "col"
                c_def_id.title = "Top Defects of Top 1 Model (Indo)"
                c_def_id.style = 11
                c_def_id.width = 15
                c_def_id.height = 9.5
                c_def_id.legend = None
                c_def_id.y_axis.number_format = "0.0%"
                c_def_id.dataLabels = DataLabelList()
                c_def_id.dataLabels.showVal = True
                c_def_id.dataLabels.showSerName = False
                c_def_id.dataLabels.showCatName = False
                c_def_id.add_data(Reference(ws, min_col=7, min_row=48, max_col=7, max_row=50), titles_from_data=False)
                c_def_id.set_categories(Reference(ws, min_col=4, min_row=48, max_col=4, max_row=50))
                ws.add_chart(c_def_id, "J125")

    # 2. Sheet: Quarter HFPA 
    hfpa_q_sheet = next((s for s in wb.sheetnames if s.strip() == "Quarter HFPA"), None)
    hfpa_m_sheet = next((s for s in wb.sheetnames if s.strip() == "Month HFPA"), None)
    if hfpa_q_sheet and hfpa_m_sheet and quarter.months:
        ws_q = wb[hfpa_q_sheet]
        ws_m = wb[hfpa_m_sheet]
        hfpa_cols = mappings.get("HFPA", {})
        covered_months = [m for m in quarter.months if m in hfpa_cols]
        if covered_months:
            start_col = hfpa_cols[covered_months[0]]
            end_col = hfpa_cols[covered_months[-1]]
            cats = Reference(ws_m, min_col=start_col, min_row=4, max_col=end_col, max_row=4)

            # Row 65: Line charts (Trend tracking)
            c_vn_h = LineChart()
            c_vn_h.title = "VN HFPA% Trend Tracking"
            c_vn_h.style = 13
            c_vn_h.y_axis.title = "HFPA%"
            c_vn_h.width = 15
            c_vn_h.height = 9.5
            c_vn_h.legend.legendPos = "t"
            c_vn_h.y_axis.number_format = "0.0%"
            for r in (28, 23, 24):
                lbl = ws_m.cell(r, 1).value or ("VN" if r == 28 else "VH" if r == 23 else "VH2")
                c_vn_h.series.append(Series(Reference(ws_m, min_col=start_col, min_row=r, max_col=end_col, max_row=r), title=lbl))
            c_vn_h.set_categories(cats)
            ws_q.add_chart(c_vn_h, "B65")

            c_id_h = LineChart()
            c_id_h.title = "INDO HFPA% Trend Tracking"
            c_id_h.style = 13
            c_id_h.y_axis.title = "HFPA%"
            c_id_h.width = 15
            c_id_h.height = 9.5
            c_id_h.legend.legendPos = "t"
            c_id_h.y_axis.number_format = "0.0%"
            for r in (29, 25, 26):
                lbl = ws_m.cell(r, 1).value or ("Indo" if r == 29 else "JV" if r == 25 else "JV2")
                c_id_h.series.append(Series(Reference(ws_m, min_col=start_col, min_row=r, max_col=end_col, max_row=r), title=lbl))
            c_id_h.set_categories(cats)
            ws_q.add_chart(c_id_h, "J65")

            # Row 85: Top Models (VN & Indo)
            if ws_q.cell(7, 14).value:
                c_mod_vn_h = BarChart()
                c_mod_vn_h.type = "col"
                c_mod_vn_h.title = "Top HFPA Models (VN)"
                c_mod_vn_h.style = 10
                c_mod_vn_h.width = 15
                c_mod_vn_h.height = 9.5
                c_mod_vn_h.legend = None
                c_mod_vn_h.y_axis.number_format = "0.0%"
                c_mod_vn_h.dataLabels = DataLabelList()
                c_mod_vn_h.dataLabels.showVal = True
                c_mod_vn_h.dataLabels.showSerName = False
                c_mod_vn_h.dataLabels.showCatName = False
                c_mod_vn_h.add_data(Reference(ws_q, min_col=17, min_row=7, max_col=17, max_row=9), titles_from_data=False)
                c_mod_vn_h.set_categories(Reference(ws_q, min_col=14, min_row=7, max_col=14, max_row=9))
                ws_q.add_chart(c_mod_vn_h, "B85")

            if ws_q.cell(4, 14).value:
                c_mod_id_h = BarChart()
                c_mod_id_h.type = "col"
                c_mod_id_h.title = "Top HFPA Models (Indo)"
                c_mod_id_h.style = 12
                c_mod_id_h.width = 15
                c_mod_id_h.height = 9.5
                c_mod_id_h.legend = None
                c_mod_id_h.y_axis.number_format = "0.0%"
                c_mod_id_h.dataLabels = DataLabelList()
                c_mod_id_h.dataLabels.showVal = True
                c_mod_id_h.dataLabels.showSerName = False
                c_mod_id_h.dataLabels.showCatName = False
                c_mod_id_h.add_data(Reference(ws_q, min_col=17, min_row=4, max_col=17, max_row=6), titles_from_data=False)
                c_mod_id_h.set_categories(Reference(ws_q, min_col=14, min_row=4, max_col=14, max_row=6))
                ws_q.add_chart(c_mod_id_h, "J85")

            # Row 105: Top Defects (VN & Indo)
            if ws_q.cell(17, 14).value:
                c_def_vn_h = BarChart()
                c_def_vn_h.type = "col"
                c_def_vn_h.title = "Top HFPA Defects (VN)"
                c_def_vn_h.style = 11
                c_def_vn_h.width = 15
                c_def_vn_h.height = 9.5
                c_def_vn_h.legend = None
                c_def_vn_h.y_axis.number_format = "0.0%"
                c_def_vn_h.dataLabels = DataLabelList()
                c_def_vn_h.dataLabels.showVal = True
                c_def_vn_h.dataLabels.showSerName = False
                c_def_vn_h.dataLabels.showCatName = False
                c_def_vn_h.add_data(Reference(ws_q, min_col=17, min_row=17, max_col=17, max_row=19), titles_from_data=False)
                c_def_vn_h.set_categories(Reference(ws_q, min_col=14, min_row=17, max_col=14, max_row=19))
                ws_q.add_chart(c_def_vn_h, "B105")

            if ws_q.cell(14, 14).value:
                c_def_id_h = BarChart()
                c_def_id_h.type = "col"
                c_def_id_h.title = "Top HFPA Defects (Indo)"
                c_def_id_h.style = 11
                c_def_id_h.width = 15
                c_def_id_h.height = 9.5
                c_def_id_h.legend = None
                c_def_id_h.y_axis.number_format = "0.0%"
                c_def_id_h.dataLabels = DataLabelList()
                c_def_id_h.dataLabels.showVal = True
                c_def_id_h.dataLabels.showSerName = False
                c_def_id_h.dataLabels.showCatName = False
                c_def_id_h.add_data(Reference(ws_q, min_col=17, min_row=14, max_col=17, max_row=16), titles_from_data=False)
                c_def_id_h.set_categories(Reference(ws_q, min_col=14, min_row=14, max_col=14, max_row=16))
                ws_q.add_chart(c_def_id_h, "J105")

    wb.save(output)
    wb.close()
    repair_excel_chart_references(output, cached_sheets=cached_sheets)


def repair_excel_chart_references(output, *, cached_sheets=None):
    """Keep multilevel chart labels readable by WPS and preserve formula caches.

    openpyxl serializes multiLvlStrRef with the cache before its formula and
    omits ptCount. WPS rejects that category source and displays an empty plot.
    Patch only the affected OOXML nodes; keep all series, styles and anchors.
    """
    original = Path(output).read_bytes()
    updated = {}
    with ZipFile(BytesIO(original)) as package:
        for name in package.namelist():
            if not re.fullmatch(r"xl/charts/chart\d+\.xml", name):
                continue
            root = etree.fromstring(package.read(name))
            changed = False
            for reference in root.iter(C + "multiLvlStrRef"):
                formula = reference.find(C + "f")
                cache = reference.find(C + "multiLvlStrCache")
                if formula is None or cache is None:
                    continue
                if reference.index(formula) != 0:
                    reference.remove(formula)
                    reference.insert(0, formula)
                    changed = True
                count = cache.find(C + "ptCount")
                if count is None:
                    # All levels share the same category positions. Quarter/year
                    # levels are sparse, so use the highest point index.
                    points = [int(point.get("idx")) for point in cache.iter(C + "pt")]
                    count = etree.Element(C + "ptCount", val=str(max(points, default=-1) + 1))
                    cache.insert(0, count)
                    changed = True
            if changed:
                updated[name] = etree.tostring(root)
        for name, xml in (cached_sheets or {}).items():
            source = etree.fromstring(xml)
            cached = {cell.get("r"): cell for cell in source.iter(X + "c")
                      if cell.find(X + "f") is not None and cell.find(X + "v") is not None}
            if not cached:
                continue
            target = etree.fromstring(package.read(name))
            for cell in target.iter(X + "c"):
                saved = cached.get(cell.get("r"))
                if saved is None or cell.find(X + "f") is None:
                    continue
                value = cell.find(X + "v")
                if value is not None:
                    cell.remove(value)
                cell.insert(cell.index(cell.find(X + "f")) + 1, deepcopy(saved.find(X + "v")))
                if saved.get("t") in ("str", "b", "e"):
                    cell.set("t", saved.get("t"))
            updated[name] = etree.tostring(target)
    if updated:
        copy_package(BytesIO(original), output, updated)


def extend_hfpa_chart(chart, quarter, columns):
    scopes = {23: "VH", 24: "VH2", 25: "JV", 26: "JV2", 27: "CLG"}
    active = {col: month for month, col in columns.items() if month in quarter.covered}
    last_covered = columns[quarter.covered[-1]]
    for reference in chart.iter(C + "numRef"):
        formula = reference.find(C + "f")
        match = re.fullmatch(r"('Month HFPA'!\$H\$(\d+):\$)([A-Z]+)(\$\d+)", formula.text or "")
        if not match or int(match[2]) not in scopes:
            continue
        cache = reference.find(C + "numCache")
        if cache is None:
            continue
        limit = column_index_from_string(match[3])
        if last_covered > limit:
            limit = last_covered
            formula.text = f"{match[1]}{get_column_letter(limit)}{match[4]}"
            cache.find(C + "ptCount").set("val", str(limit - 7))
        points = {int(point.get("idx")): point for point in cache.findall(C + "pt")}
        for col, month in active.items():
            if col > limit:
                continue
            index = col - 8
            point = points.get(index)
            record = quarter.values.get(("HFPA", month, scopes[int(match[2])])) if month in quarter.covered else None
            if record and record.rate is not None:
                if point is None:
                    point = etree.SubElement(cache, C + "pt", idx=str(index))
                    etree.SubElement(point, C + "v")
                point.find(C + "v").text = format(record.rate, ".15g")
            elif point is not None:
                cache.remove(point)
    for reference in chart.iter(C + "multiLvlStrRef"):
        formula, cache = reference.find(C + "f"), reference.find(C + "multiLvlStrCache")
        match = re.fullmatch(r"('Month HFPA'!\$H\$2:\$)([A-Z]+)(\$4)", formula.text or "")
        if not match or cache is None:
            continue
        limit = column_index_from_string(match[2])
        if last_covered <= limit:
            continue
        formula.text = f"{match[1]}{get_column_letter(last_covered)}{match[3]}"
        cache.find(C + "ptCount").set("val", str(last_covered - 7))
        level = cache.find(C + "lvl")
        for col in range(limit + 1, last_covered + 1):
            month = next(month for month, index in columns.items() if index == col)
            point = etree.SubElement(level, C + "pt", idx=str(col - 8))
            etree.SubElement(point, C + "v").text = month_abbr[int(month[5:])]


def font(size, bold=False, italic=False):
    family = "arialbd.ttf" if bold else "ariali.ttf" if italic else "arial.ttf"
    path = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts" / family
    try:
        return ImageFont.truetype(str(path), size)
    except OSError:
        return ImageFont.load_default(size=size)


def png(image):
    result = BytesIO()
    image.save(result, "PNG")
    return result.getvalue()


# Pixel coordinates of the embedded pictures in the supplied PowerPoint.
# Keep the original picture as the background and replace only its data marks.
TREND_FORMAT = {
    ("FTT", "VN"): ((304, 699, 1093), 90, 392, 50, 85, 5, 423,
                     ("#002060", "#C9A779", "#8C9BAF"), "#66FFFF"),
    ("FTT", "Indo"): ((308, 700, 1091), 89, 392, 50, 85, 5, 422,
                       ("#6453AD", "#9A79C5", "#89CDC5"), "#FFFF66"),
    ("HFPA", "VN"): ((264, 673, 1082), 74, 431, 80, 90, 2, 461,
                      ("#4B375F", "#DBC9FF", "#7C5C78"), "#00FFFF"),
    ("HFPA", "Indo"): ((270, 677, 1085), 95, 430, 70, 98, 4, 460,
                        ("#4B375F", "#E0D4DF", "#DCF3FF"), "#FFD966"),
}
RANK_FORMAT = {
    ("FTT", "VN", False): (38, 146, 52, ((85, 147), (284, 345), (483, 544)), "#5B9BD5"),
    ("FTT", "VN", True): (29, 142, 68, ((75, 139), (281, 344), (486, 549)), None),
    ("FTT", "Indo", False): (36, 145, 69, ((91, 154), (294, 357), (497, 560)), "#9C7BC7"),
    ("FTT", "Indo", True): (30, 141, 61, ((80, 144), (287, 351), (494, 558)), None),
    ("HFPA", "VN", False): (47, 160, 52, ((60, 138), (257, 334), (453, 531)), "#5B9BD5"),
    ("HFPA", "VN", True): (30, 172, 110, ((70, 157), (290, 376), (509, 595)), None),
    ("HFPA", "Indo", False): (43, 157, 68, ((70, 157), (290, 377), (510, 597)), "#8C7CA2"),
    ("HFPA", "Indo", True): (30, 185, 105, ((70, 157), (289, 376), (508, 595)), None),
}


def erase_trend_marks(image, left, top, right, bottom, sample_x):
    """Copy a clean column of grid/background pixels across the data region."""
    strip = image.crop((sample_x, top, sample_x + 1, bottom))
    image.paste(strip.resize((right - left, bottom - top), Image.Resampling.NEAREST), (left, top))


def trend_image(image, quarter, metric, region):
    xs, top, bottom, low, high, step, month_y, colors, highlight = TREND_FORMAT[metric, region]
    original_low, original_high = low, high
    scopes = (region, *REGIONS[region])
    series = [[record.rate if record else None for record in (
        quarter.values.get((metric, month, scope)) if month in quarter.covered else None
        for month in quarter.months)] for scope in scopes]
    rates = [100 * rate for line in series for rate in line if rate is not None]
    # Keep the original axis unless a new value needs a wider scale.
    if rates and (min(rates) < low or max(rates) > high):
        divisions = (high - low) // step
        first = max(0, math.floor(min(rates)))
        last = min(100, math.ceil(max(rates)))
        step = max(step, math.ceil((last - first) / divisions))
        low = min(first, max(0, 100 - divisions * step))
        high = low + divisions * step
        if high > 100:
            low, high = 0, 100
    erase_trend_marks(image, 150, 64 if metric == "FTT" else 54, 1200, bottom, 128)
    if (low, high) != (original_low, original_high):
        right = 100 if metric == "FTT" else 55
        erase_trend_marks(image, 13, top - 13, right, bottom + 12, right + 1)

    def y(rate):
        return bottom - (bottom - top) * (rate * 100 - low) / (high - low)

    draw = ImageDraw.Draw(image)
    if (low, high) != (original_low, original_high):
        divisions = (original_high - original_low) // TREND_FORMAT[metric, region][5]
        for i in range(divisions + 1):
            rate = (high - i * (high - low) / divisions) / 100
            draw.text((right - 3, y(rate)), f"{rate:.0%}", font=font(17, italic=True),
                      fill="#B3B3B3", anchor="rm")
    for x, month in zip(xs, quarter.months):
        erase_trend_marks(image, x - 40, month_y - 13, x + 40, month_y + 15, x - 50)
        draw.text((x, month_y), month_abbr[int(month[5:])], font=font(21, bold=True), fill="black", anchor="mm")
    for j, (line, color) in enumerate(zip(series, colors)):
        marker = "#4472C4" if (metric, region, j) == ("FTT", "Indo", 1) else color
        for i, (x, value) in enumerate(zip(xs, line)):
            if value is None:
                continue
            level = y(value)
            if i and line[i - 1] is not None:
                draw.line((xs[i - 1], y(line[i - 1]), x, level), fill=color,
                          width=5 if metric == "HFPA" and j == 0 else 4 if metric == "FTT" else 3)
            draw.ellipse((x - 5, level - 5, x + 5, level + 5), fill="white", outline=marker, width=2)
    for i, x in enumerate(xs):
        regional = series[0][i]
        plant_marks = [(y(line[i]), line[i]) for line in series[1:] if line[i] is not None]
        for j, (level, value) in enumerate(sorted(plant_marks)):
            position = level - 22 if j == 0 else level + 22
            draw.text((x, position), f"{value:.1%}", font=font(17, bold=metric == "FTT", italic=metric == "HFPA"),
                      fill="black", anchor="mm")
        if regional is not None:
            label = f"{regional:.1%}"
            level = y(regional)
            center = x if metric == "FTT" else x + (44 if i < 2 else -50)
            f = font(18, bold=True)
            width = math.ceil(draw.textlength(label, font=f))
            draw.rectangle((center - width/2 - 5, level - 14, center + width/2 + 5, level + 14), fill=highlight)
            draw.text((center, level), label, font=f, fill="black", anchor="mm")
    return png(image)


def rank_image(image, metric, region, ranks, defect_colors=None):
    header, baseline, max_height, bars, model_color = RANK_FORMAT[metric, region, defect_colors is not None]
    w, h = image.size
    image.paste((0, 0, 0, 0), (0, header, w, baseline))
    image.paste((0, 0, 0, 0), (0, baseline + 1, w, h))
    draw = ImageDraw.Draw(image)
    rates = [numerator / denominator if denominator else numerator for _, numerator, denominator in ranks]
    highest = max(rates, default=1) or 1
    for i, ((label, _, _), rate, (left, right)) in enumerate(zip(ranks, rates, bars)):
        center = (left + right) / 2
        top = baseline - max_height * rate / highest
        color = model_color if defect_colors is None else defect_colors[i]
        if rate > 0:
            draw.rectangle((left, top, right, baseline - 1), fill=color)
        draw.text((center, top - 19), f"{rate:.1%}", font=font(17, bold=True), fill="black", anchor="mm")
        lines = textwrap.wrap(" ".join(label.split()), width=max(16, w // 29))[:3]
        label_height = 16 if len(lines) < 3 else 13
        for j, line in enumerate(lines):
            draw.text((center, baseline + 12 + j * label_height), line, font=font(label_height, bold=True),
                      fill="black", anchor="mt")
    return png(image)


def defect_chart_colors(rankings):
    names = {ColorManager.normalize_name(label) for _, defects in rankings.values() for label, _, _ in defects}
    if not names:
        return {}
    tools = APP_ROOT / "report_tools"
    manager = ColorManager(str(tools / "hfpa" / "Color_template.xlsx"))
    _, colors, _, _ = manager.load_existing_template()
    _, reference, _, _ = ColorManager(str(tools / "reinspection" / "Color_Template.xlsx")).load_existing_template()
    if not colors or not reference:
        raise ValueError("Không đọc được bảng màu đang dùng cho báo cáo HFPA.")
    # Match the HFPA worker's reference sync without writing to either workbook.
    colors.update({name: reference[name] for name in colors.keys() & reference.keys()})
    missing = sorted(names - colors.keys())
    colors.update(zip(missing, manager.get_curated_palette(len(missing), set(colors.values())), strict=True))
    return {name: f"#{colors[name]}" for name in names}


def quality_slides(deck):
    """Locate the QMS slides in both the compact deck and the full QBR."""
    found = {}
    for slide in deck.slides:
        for shape in slide.shapes:
            match = re.search(r"\bQMS\s+(FTT|HFPA)\s+ANALYSIS\b", shape.text, re.I) if shape.has_text_frame else None
            if match is None:
                continue
            metric = match[1].upper()
            if metric in found:
                raise ValueError(f"Mẫu PowerPoint có nhiều slide QMS {metric} ANALYSIS.")
            spec = SHAPES[metric]
            shapes = {item.shape_id: item for item in slide.shapes}
            pictures = (spec["period"], *spec["VN"], *spec["Indo"])
            if any(index not in shapes or not hasattr(shapes[index], "image") for index in pictures):
                raise ValueError(f"Slide QMS {metric} thiếu hình biểu đồ theo mẫu QBR.")
            found[metric] = slide, shape
    if found.keys() != SHAPES.keys():
        raise ValueError("Mẫu PowerPoint cần một slide QMS FTT ANALYSIS và một slide QMS HFPA ANALYSIS.")
    return found


def set_table_text(cell, value, style_cell):
    """Retain the template's paragraph and run formatting, including blank cells."""
    paragraph = cell.text_frame.paragraphs[0]
    if not paragraph.runs:
        run = paragraph.add_run()
        reference = style_cell.text_frame.paragraphs[0].runs
        if reference and reference[0]._r.rPr is not None:
            run._r.insert(0, deepcopy(reference[0]._r.rPr))
    paragraph.runs[0].text = value
    for run in paragraph.runs[1:]:
        run.text = ""


def consumer_summary_table_image(quarter):
    """Render a crisp, Excel-styled screenshot of the current quarter HFPA summary table."""
    w, h = 547, 486
    image = Image.new("RGB", (w, h), color=(255, 255, 255))
    draw = ImageDraw.Draw(image)

    def get_font(size, bold=False):
        for family in ("calibri.ttf", "arial.ttf"):
            path = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts" / family
            try:
                return ImageFont.truetype(str(path), size)
            except OSError:
                continue
        return ImageFont.load_default()

    font_regular = get_font(14)

    col_widths = [85, 80, 80, 80, 80, 80, 62]
    cols = [0]
    for cw in col_widths:
        cols.append(cols[-1] + cw)
    cols[-1] = w - 1

    row_h = 24
    start_y = 27
    num_rows = 19
    grid_color = (212, 212, 212)
    yellow_color = (255, 255, 0)

    if getattr(quarter, "date_to", None):
        start = date.fromisoformat(quarter.date_from)
        end = date.fromisoformat(quarter.date_to)
        date_hdr = f"{start:%b}-{end:%b}{end.day}"
    elif getattr(quarter, "covered", None):
        start = date.fromisoformat(quarter.covered[0] + "-01")
        end = date.fromisoformat(quarter.covered[-1] + "-01")
        date_hdr = f"{start:%b}-{end:%b}"
    else:
        date_hdr = getattr(quarter, "period", "")

    headers = ["FTY", "Insp", "Def", "HFPA%", date_hdr]

    def get_counts(plant):
        if hasattr(quarter, "totals"):
            tot = quarter.totals("HFPA", plant)
            if tot is not None:
                return float(tot.inspected), float(tot.defects)
        return 0.0, 0.0

    vh_insp, vh_def = get_counts("VH")
    vh2_insp, vh2_def = get_counts("VH2")
    vn_insp, vn_def = vh_insp + vh2_insp, vh_def + vh2_def
    jv_insp, jv_def = get_counts("JV")
    jv2_insp, jv2_def = get_counts("JV2")
    indo_insp, indo_def = jv_insp + jv2_insp, jv_def + jv2_def
    clg_insp, clg_def = vn_insp + indo_insp, vn_def + indo_def

    def fmt_num(val):
        return f"{int(round(val)):,d}" if val > 0 else "0"

    def fmt_rate(insp, defs):
        return f"{(1.0 - defs / insp):.1%}" if insp > 0 else "—"

    rows_data = [
        ("VH", fmt_num(vh_insp), fmt_num(vh_def), fmt_rate(vh_insp, vh_def)),
        ("VH2", fmt_num(vh2_insp), fmt_num(vh2_def), fmt_rate(vh2_insp, vh2_def)),
        ("VN", fmt_num(vn_insp), fmt_num(vn_def), fmt_rate(vn_insp, vn_def)),
        ("JV", fmt_num(jv_insp), fmt_num(jv_def), fmt_rate(jv_insp, jv_def)),
        ("JV2", fmt_num(jv2_insp), fmt_num(jv2_def), fmt_rate(jv2_insp, jv2_def)),
        ("Indo", fmt_num(indo_insp), fmt_num(indo_def), fmt_rate(indo_insp, indo_def)),
        ("CLG", fmt_num(clg_insp), fmt_num(clg_def), fmt_rate(clg_insp, clg_def)),
    ]
    yellow_row_indices = {3, 6, 7}

    for r_idx in yellow_row_indices:
        y1 = start_y + r_idx * row_h
        y2 = y1 + row_h
        draw.rectangle((0, y1, cols[4], y2), fill=yellow_color)

    for r in range(num_rows + 1):
        y = start_y + r * row_h
        if y < h:
            draw.line([(0, y), (w, y)], fill=grid_color, width=1)

    for x in cols:
        draw.line([(x, start_y), (x, min(h, start_y + num_rows * row_h))], fill=grid_color, width=1)

    y_hdr = start_y + row_h / 2
    for c_idx, text in enumerate(headers):
        x1, x2 = cols[c_idx], cols[c_idx + 1]
        if c_idx == 0 or c_idx == 4:
            draw.text((x1 + 8, y_hdr), text, font=font_regular, fill=(0, 0, 0), anchor="lm")
        else:
            draw.text((x2 - 8, y_hdr), text, font=font_regular, fill=(0, 0, 0), anchor="rm")

    for r_idx, row in enumerate(rows_data, start=1):
        y_mid = start_y + r_idx * row_h + row_h / 2
        for c_idx, val in enumerate(row):
            x1, x2 = cols[c_idx], cols[c_idx + 1]
            if c_idx == 0:
                draw.text((x1 + 8, y_mid), val, font=font_regular, fill=(0, 0, 0), anchor="lm")
            else:
                draw.text((x2 - 8, y_mid), val, font=font_regular, fill=(0, 0, 0), anchor="rm")

    out = BytesIO()
    image.save(out, format="PNG")
    return out.getvalue()


def update_consumer_corner(slide, quarter, parts, package_names, relationships):
    """Update executive corner sticky notes on Slide 1: date box and HFPA summary table picture."""
    slide_name = str(slide.part.partname).lstrip("/")
    rels_name = f"ppt/slides/_rels/{Path(slide_name).name}.rels"
    rels = etree.fromstring(relationships[slide_name]) if slide_name in relationships else None

    for shape in slide.shapes:
        if shape.shape_id == 5 or (shape.has_text_frame and shape.name == "Rectangle 4"):
            if getattr(quarter, "date_to", None):
                start = date.fromisoformat(quarter.date_from)
                end = date.fromisoformat(quarter.date_to)
                box_text = f"{start:%b} to\n{end:%b} {end.day}"
            elif getattr(quarter, "covered", None):
                start = date.fromisoformat(quarter.covered[0] + "-01")
                end = date.fromisoformat(quarter.covered[-1] + "-01")
                box_text = f"{start:%b} to\n{end:%b}"
            else:
                box_text = getattr(quarter, "period", "").replace("-", " ")
            if shape.has_text_frame and shape.text_frame.paragraphs:
                p = shape.text_frame.paragraphs[0]
                if p.runs:
                    p.runs[0].text = box_text
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.text = box_text
                for p_extra in shape.text_frame.paragraphs[1:]:
                    for r in p_extra.runs:
                        r.text = ""

        if (shape.shape_id == 6 or (hasattr(shape, "image") and shape.name == "Picture 5")) and rels is not None:
            table_png = consumer_summary_table_image(quarter)
            index = 1
            stem = f"quarterly_{Path(slide_name).stem}_corner_summary"
            while f"ppt/media/{stem}_{index}.png" in package_names:
                index += 1
            media_name = f"ppt/media/{stem}_{index}.png"
            package_names.add(media_name)
            parts[media_name] = table_png
            relationship_id = f"rIdQuarterlyCorner{shape.shape_id}"
            while any(item.get("Id") == relationship_id for item in rels):
                relationship_id += "x"
            etree.SubElement(rels, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                             Id=relationship_id,
                             Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                             Target="../media/" + Path(media_name).name)
            shape._pic.blipFill.blip.rEmbed = relationship_id

    if rels is not None:
        parts[rels_name] = etree.tostring(rels)


def update_consumer_promise(slide, quarter):
    """The QBR labels VH and JV represent their respective two-factory groups."""
    tables = [(shape.table, row, col) for shape in slide.shapes if shape.has_table
              for row in range(len(shape.table.rows)) for col in range(len(shape.table.columns))
              if " ".join(shape.table.cell(row, col).text.upper().split()) == "QUALITY (HFPA)"]
    if len(tables) != 1:
        raise ValueError("Consumer Promise Review cần một bảng Quality (HFPA) theo mẫu.")
    table, header, column = tables[0]
    groups = {"VH": REGIONS["VN"], "JV": REGIONS["Indo"], "CLG": PLANTS}
    rows = {}
    for row in range(header + 2, len(table.rows)):
        label = next((cell.text.strip().upper() for cell in reversed(tuple(table.rows[row].cells)[:column])
                      if cell.text.strip().upper() in groups), None)
        if label is not None:
            if label in rows:
                raise ValueError(f"Consumer Promise Review có nhiều dòng {label}.")
            rows[label] = row
    if rows.keys() != groups.keys() or column + 1 >= len(table.columns):
        raise ValueError("Consumer Promise Review thiếu dòng VH, JV, CLG hoặc cột quý trước.")
    fiscal_year, period = quarter.period.split("-")
    set_table_text(table.cell(header + 1, column), period + fiscal_year,
                   table.cell(header + 1, column))
    periods = (False, True) if quarter.prior_date_to else (False,)
    for label, plants in groups.items():
        row = rows[label]
        for prior in periods:
            records = [quarter.totals("HFPA", plant, prior=prior) for plant in plants]
            rate = None
            if all(record is not None for record in records):
                total = Counts(sum(record.inspected for record in records), sum(record.defects for record in records))
                rate = total.rate
            set_table_text(table.cell(row, column + int(prior)), f"{rate:.1%}" if rate is not None else "—",
                           table.cell(row, column + 1))


def write_powerpoint(template, output, quarter):
    deck = Presentation(template)
    slides = quality_slides(deck)
    parts = {}
    defect_colors = defect_chart_colors(quarter.rankings)
    period_label = quarter.period.replace("-", " ")
    if quarter.date_to:
        start, end = (date.fromisoformat(day) for day in (quarter.date_from, quarter.date_to))
        date_label = f"Data from {start:%b} {start.day} to {end:%b} {end.day}"
    else:
        start, end = (date.fromisoformat(month + "-01") for month in (quarter.covered[0], quarter.covered[-1]))
        date_label = f"Data from {start:%B} to {end:%B}"
    consumer_slides = [slide for slide in deck.slides
                       if any(shape.has_text_frame and "CONSUMER PROMISE REVIEW" in shape.text.upper() for shape in slide.shapes)]
    with ZipFile(template) as package:
        package_names = set(package.namelist())
        relevant_slides = [*slides.values(), *((s, None) for s in consumer_slides)]
        relationships = {}
        for s, _ in relevant_slides:
            partname = str(s.part.partname).lstrip("/")
            rel_path = f"ppt/slides/_rels/{Path(partname).name}.rels"
            if rel_path in package_names:
                relationships[partname] = package.read(rel_path)
    for metric, (slide, title_shape) in slides.items():
        spec = SHAPES[metric]
        shapes = {shape.shape_id: shape for shape in slide.shapes}
        title = title_shape.text_frame.paragraphs[0].runs[0]
        title.text = f"{period_label} QMS {metric} ANALYSIS"
        slide_name = str(slide.part.partname).lstrip("/")
        rels = etree.fromstring(relationships[slide_name])

        def replace_image(shape_id, make_image):
            shape = shapes[shape_id]
            with Image.open(BytesIO(shape.image.blob)) as original:
                data = make_image(original.convert("RGBA"))
            # Period pictures are shared with non-QMS slides in the full QBR.
            # Give each replacement its own relationship and media part.
            index = 1
            stem = f"quarterly_{Path(slide_name).stem}_{shape_id}"
            while f"ppt/media/{stem}_{index}.png" in package_names:
                index += 1
            media_name = f"ppt/media/{stem}_{index}.png"
            package_names.add(media_name)
            parts[media_name] = data
            relationship_id = f"rIdQuarterly{shape_id}"
            while any(item.get("Id") == relationship_id for item in rels):
                relationship_id += "x"
            etree.SubElement(rels, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                             Id=relationship_id,
                             Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                             Target="../media/" + Path(media_name).name)
            shape._pic.blipFill.blip.rEmbed = relationship_id

        def period_image(image):
            bounds = image.getbbox()
            image.paste((0, 0, 0, 0), bounds)
            draw = ImageDraw.Draw(image)
            points = 23
            while draw.textlength(date_label, font=font(points, italic=True)) > image.width - 2 * bounds[0]:
                points -= 1
            draw.text((bounds[0], (bounds[1] + bounds[3]) / 2), date_label,
                      font=font(points, italic=True), fill="#333333", anchor="lm")
            return png(image)

        excel_filename = Path(output).with_suffix(".xlsx").name if output else f"HFPA-FTT-{quarter.period}.xlsx"
        def add_hyperlink(shape):
            hlink_id = f"rIdHlink{shape.shape_id}"
            while any(item.get("Id") == hlink_id for item in rels):
                hlink_id += "x"
            etree.SubElement(rels, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                             Id=hlink_id,
                             Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                             Target=excel_filename)
            A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
            R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
            c_nv_pr = shape._element.xpath(".//p:cNvPr")[0]
            for h in c_nv_pr.findall(A_NS + "hlinkClick"):
                c_nv_pr.remove(h)
            etree.SubElement(c_nv_pr, A_NS + "hlinkClick", attrib={R_NS + "id": hlink_id})

        replace_image(spec["period"], period_image)
        for region in ("VN", "Indo"):
            trend, top_model, top_defect = spec[region]
            models, defects = quarter.rankings.get((metric, region), ([], []))
            replace_image(trend, lambda image: trend_image(image, quarter, metric, region))
            add_hyperlink(shapes[trend])
            replace_image(top_model, lambda image: rank_image(image, metric, region, models))
            add_hyperlink(shapes[top_model])
            palette = [defect_colors[ColorManager.normalize_name(label)] for label, _, _ in defects]
            replace_image(top_defect, lambda image: rank_image(image, metric, region, defects, palette))
            add_hyperlink(shapes[top_defect])
        parts[slide_name] = slide.part.blob
        parts[f"ppt/slides/_rels/{Path(slide_name).name}.rels"] = etree.tostring(rels)
    for slide in consumer_slides:
        update_consumer_promise(slide, quarter)
        update_consumer_corner(slide, quarter, parts, package_names, relationships)
        parts[str(slide.part.partname).lstrip("/")] = slide.part.blob
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame
            prior_note = (f"Prior comparison through {quarter.prior_date_to}. " if quarter.prior_date_to else
                          "Prior-period figures use the saved template. ")
            note = (f"Quarterly HFPA summary {quarter.period}, through {quarter.date_to or quarter.covered[-1]}. "
                    f"{prior_note}"
                    "VH = VH + VH2. JV = JV + JV2. CLG = VH + VH2 + JV + JV2. "
                    "HFPA = 1 - total failed pairs / total inspected pairs. The vs PQ column shows the prior-period rate.")
            notes.text = "\n".join(line for line in notes.text.splitlines()
                                   if not line.startswith("Quarterly HFPA summary ")) + "\n" + note
            parts[str(slide.notes_slide.part.partname).lstrip("/")] = slide.notes_slide.part.blob
    copy_package(template, output, parts)
    excel = Path(output).with_suffix('.xlsx')
    if excel.is_file():
        from .editable_charts import convert
        convert(BytesIO(Path(output).read_bytes()), excel, output, excel)
