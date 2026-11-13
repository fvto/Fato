"""Generate finalized, formula-driven Excel workbook with styled native charts."""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from report_tools.quarterly_defect_review.analysis import pct
from report_tools.quarterly_defect_review.contracts import (
    AREAS,
    SCOPES,
    BaseRecord,
    DefectFact,
    LineCatalogEntry,
)
from report_tools.quarterly_defect_review.defect_mapping import load_defect_mapping

# Styles (corporate non-red, non-green palette matching project guidelines)
FONT_FAMILY = 'Calibri'
TITLE_FILL = PatternFill('solid', fgColor='1B365D')
SUBTITLE_FILL = PatternFill('solid', fgColor='2B547E')
HEADER_FILL = PatternFill('solid', fgColor='244062')
SOFT_HEADER_FILL = PatternFill('solid', fgColor='EBF1F8')
ZEBRA_FILL = PatternFill('solid', fgColor='F8FAFC')
TOTAL_FILL = PatternFill('solid', fgColor='E2E8F0')
WARN_FILL = PatternFill('solid', fgColor='FEF3C7')
ERROR_FILL = PatternFill('solid', fgColor='FEE2E2')
INFO_FILL = PatternFill('solid', fgColor='E0F2FE')

TITLE_FONT = Font(name=FONT_FAMILY, size=14, bold=True, color='FFFFFF')
SECTION_FONT = Font(name=FONT_FAMILY, size=11, bold=True, color='FFFFFF')
HDR_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color='FFFFFF')
SUBHDR_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color='1B365D')
BODY_FONT = Font(name=FONT_FAMILY, size=10, color='1E293B')
BOLD_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color='1E293B')
NOTE_FONT = Font(name=FONT_FAMILY, size=9, italic=True, color='475569')

THIN_SIDE = Side(style='thin', color='CBD5E1')
MED_NAVY = Side(style='medium', color='1B365D')
CELL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
TOTAL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=MED_NAVY, bottom=MED_NAVY)

SCOPE_SERIES_COLORS = {
    'VH': '1F4E78',   # Deep Navy
    'JV': 'D97706',   # Warm Amber
    'VB': '007A87',   # Rich Teal
}
RANK_SERIES_COLORS = ['1F77B4', '007A87', 'D97706']


def _quarter_month_keys(date_from: str, date_to: str) -> list[str]:
    """Return ordered YYYY-MM month strings spanning date_from..date_to (max 3-6)."""
    if not date_from or not date_to:
        return []
    start = datetime.strptime(date_from[:10], '%Y-%m-%d').date()
    end = datetime.strptime(date_to[:10], '%Y-%m-%d').date()
    months: list[str] = []
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month) and len(months) < 6:
        months.append(f'{y:04d}-{m:02d}')
        m += 1
        if m > 12:
            m = 1
            y += 1
    return months


def _month_label(ym: str) -> str:
    return datetime.strptime(ym + '-01', '%Y-%m-%d').strftime('%b %Y')


def _style_range(
    ws: Any,
    min_col: int,
    min_row: int,
    max_col: int,
    max_row: int,
    *,
    font: Font | None = None,
    fill: PatternFill | None = None,
    border: Border | None = CELL_BORDER,
    alignment: Alignment | None = None,
    num_format: str | None = None,
) -> None:
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if font:
                cell.font = font
            if fill:
                cell.fill = fill
            if border:
                cell.border = border
            if alignment:
                cell.alignment = alignment
            if num_format:
                cell.number_format = num_format


def _write_banner(ws: Any, row: int, col1: int, col2: int, text: str, fill: PatternFill = SUBTITLE_FILL, font: Font = SECTION_FONT) -> None:
    ws.cell(row=row, column=col1, value=text)
    if col2 > col1:
        ws.merge_cells(start_row=row, start_column=col1, end_row=row, end_column=col2)
    _style_range(
        ws,
        col1,
        row,
        col2,
        row,
        font=font,
        fill=fill,
        border=CELL_BORDER,
        alignment=Alignment(horizontal='left', vertical='center', wrap_text=True),
    )
    ws.row_dimensions[row].height = 24


def _make_clean_datalabels(*, hide_zeros: bool = True) -> DataLabelList:
    dl = DataLabelList()
    dl.showVal = True
    dl.showCatName = False
    dl.showSerName = False
    dl.showLegendKey = False
    dl.showPercent = False
    dl.showBubbleSize = False
    if hide_zeros:
        dl.numFmt = '#,##0;;""'
    return dl


def _build_stacked_bar_chart(
    title: str,
    ws: Any,
    data_min_col: int,
    data_max_col: int,
    hdr_row: int,
    cat_col: int,
    min_row: int,
    max_row: int,
    *,
    y_title: str = 'Defect Quantity (Issues)',
    width: float = 18.5,
    height: float = 10.5,
    active_series: tuple[bool, ...] | None = None,
) -> BarChart:
    chart = BarChart()
    chart.type = 'col'
    chart.grouping = 'stacked'
    chart.overlap = 100
    chart.style = 10
    chart.title = title
    chart.y_axis.title = y_title
    chart.y_axis.number_format = '#,##0'
    chart.y_axis.majorGridlines = openpyxl.chart.axis.ChartLines()
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend.position = 'b'
    chart.width = width
    chart.height = height

    data_ref = Reference(ws, min_col=data_min_col, max_col=data_max_col, min_row=hdr_row, max_row=max_row)
    cats_ref = Reference(ws, min_col=cat_col, min_row=min_row, max_row=max_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    colors = [SCOPE_SERIES_COLORS['VH'], SCOPE_SERIES_COLORS['JV'], SCOPE_SERIES_COLORS['VB']]
    for idx, series in enumerate(chart.series):
        color = colors[idx % len(colors)]
        series.graphicalProperties.solidFill = color
        series.graphicalProperties.line.solidFill = color
        if active_series is None or (idx < len(active_series) and active_series[idx]):
            series.dLbls = _make_clean_datalabels(hide_zeros=True)

    return chart


def _build_horizontal_bar_chart(
    title: str,
    ws: Any,
    val_col: int,
    hdr_row: int,
    cat_col: int,
    min_row: int,
    max_row: int,
    *,
    color: str = '1F77B4',
    width: float = 19.5,
    height: float = 11.0,
    max_val: float | None = None,
) -> BarChart:
    chart = BarChart()
    chart.type = 'bar'
    chart.grouping = 'clustered'
    chart.style = 10
    chart.title = title
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.scaling.orientation = 'maxMin'
    chart.y_axis.number_format = '#,##0'
    if max_val and max_val > 0:
        import math
        step = max(1, 10 ** max(0, int(math.log10(max_val)) - 1))
        raw_unit = max_val / 4.5
        chart.y_axis.majorUnit = int(math.ceil(raw_unit / step) * step)
    chart.legend = None
    chart.width = width
    chart.height = height

    data_ref = Reference(ws, min_col=val_col, max_col=val_col, min_row=hdr_row, max_row=max_row)
    cats_ref = Reference(ws, min_col=cat_col, min_row=min_row, max_row=max_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    if chart.series:
        s = chart.series[0]
        s.graphicalProperties.solidFill = color
        s.graphicalProperties.line.solidFill = color
        s.dLbls = _make_clean_datalabels(hide_zeros=True)
    return chart


def _write_drilldown_rollup(
    wb: openpyxl.Workbook,
    hfpa_base: list[BaseRecord],
    hfpa_facts: list[DefectFact],
    ftt_base: list[BaseRecord],
    ftt_facts: list[DefectFact],
) -> tuple[int, dict[tuple[str, str], dict[str, float]]]:
    """Write normalized rollup fact table used by Excel SUMIFS formulas & verification."""
    ws = wb.create_sheet('Drilldown_Normalized')
    ws.views.sheetView[0].showGridLines = True

    headers = [
        'Metric',
        'Report_Area',
        'Factory',
        'Plant',
        'Line_Std',
        'Model',
        'Standard_Group_ID',
        'Standard_Group_Label',
        'NIKE_Subtype_Label',
        'Detail_Label',
        'Detail_Status',
        'Issue_Qty',
        'Provenance_Sample',
    ]
    ws.append(headers)
    _style_range(ws, 1, 1, len(headers), 1, font=HDR_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal='center', vertical='center', wrap_text=True))
    ws.row_dimensions[1].height = 24

    rollup: dict[tuple[str, ...], dict[str, Any]] = {}
    for fact in list(hfpa_facts) + list(ftt_facts):
        if fact.report_area not in AREAS:
            continue
        key = (
            fact.metric,
            fact.report_area,
            fact.factory,
            fact.plant,
            fact.line_std,
            fact.model_label,
            fact.standard_group_id or '',
            fact.standard_group_label or '(Chưa ánh xạ)',
            fact.nike_subtype_label or '',
            fact.detail_label or ('Chưa có chi tiết' if fact.is_synthetic_gap else '(Cấp nhóm cha)'),
            fact.detail_status,
        )
        item = rollup.setdefault(key, {'qty': 0.0, 'prov': f'{fact.source_file}#{fact.source_sheet}:R{fact.source_row}'})
        item['qty'] += float(fact.issue_quantity or 0.0)

    row_num = 2
    for key in sorted(rollup):
        val = rollup[key]
        ws.append(list(key) + [val['qty'], val['prov']])
        if row_num % 2 == 1:
            _style_range(ws, 1, row_num, len(headers), row_num, font=BODY_FONT, fill=ZEBRA_FILL)
        else:
            _style_range(ws, 1, row_num, len(headers), row_num, font=BODY_FONT)
        ws.cell(row=row_num, column=12).number_format = '#,##0'
        row_num += 1

    last_row = max(2, row_num - 1)

    # Monthly HFPA aggregation helper for Slide 1 dynamic 3-month table
    monthly_hfpa: dict[tuple[str, str], dict[str, float]] = defaultdict(lambda: {'sample': 0.0, 'def_pairs': 0.0, 'issues': 0.0})
    for rec in hfpa_base:
        if rec.report_area not in AREAS and rec.factory not in {'VH', 'VH2', 'JV', 'JV2'}:
            continue
        ym = (rec.date or '')[:7]
        if not ym:
            continue
        bucket = monthly_hfpa[(ym, rec.factory)]
        bucket['sample'] += float(rec.inspected_or_produced_qty or 0.0)
        bucket['def_pairs'] += float(rec.defective_pairs_qty or 0.0)
        bucket['issues'] += float(rec.defect_qty_control or 0.0)

    widths = [9, 12, 9, 9, 10, 28, 22, 26, 24, 28, 18, 14, 28]
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    return last_row, monthly_hfpa


def _write_metric_overview_block(
    ws: Any,
    *,
    start_row: int,
    metric: str,
    result: dict[str, Any],
    rollup_last_row: int,
) -> dict[str, tuple[str, str]]:
    """Write Scope Volume table + Top 3 Groups table with Excel SUM/SUMIFS formulas + native chart."""
    named_ranges: dict[str, tuple[str, str]] = {}
    scope_totals = result['scope_totals']

    _write_banner(
        ws,
        start_row,
        1,
        10,
        f'{metric} NIKE — TỔNG HỢP THEO KHU (VH / JV / VB / CLG) & TOP 3 NHÓM LỖI CHÍNH',
        fill=SUBTITLE_FILL,
    )

    # 1) Scope Volume block (row start+1 .. start+6)
    hdr_row = start_row + 1
    if metric == 'HFPA':
        hdrs = ['Khu vực', 'Số bản ghi', 'Audit Sample Size', 'Produced Qty', 'Defective Pairs', 'Tổng Lỗi (Issues)', 'Coverage chi tiết', 'Trạng thái nguồn', 'Ghi chú', '']
    else:
        hdrs = ['Khu vực', 'Số bản ghi', 'Produced (A+B+C)', 'A-Grade Pairs', 'Rework Pairs', 'Tổng Lỗi (Issues)', 'B+C Reject Pairs', 'Trạng thái nguồn', 'Ghi chú', '']
    for c, h in enumerate(hdrs[:9], 1):
        ws.cell(row=hdr_row, column=c, value=h)
    _style_range(ws, 1, hdr_row, 9, hdr_row, font=HDR_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal='center', vertical='center', wrap_text=True))
    ws.row_dimensions[hdr_row].height = 26

    scope_labels = {'VH': 'VH (VH+VH2 trừ VB)', 'JV': 'JV (JV+JV2)', 'VB': 'VB (Factory Sandals)', 'CLG': 'CLG (VH+JV+VB)'}
    r_idx = hdr_row + 1
    first_area_row = r_idx
    for area in AREAS:
        st = scope_totals[area]
        ws.cell(row=r_idx, column=1, value=scope_labels[area])
        ws.cell(row=r_idx, column=2, value=int(st['record_count']) if st['has_source_data'] else 'Thiếu dữ liệu')
        if metric == 'HFPA':
            ws.cell(row=r_idx, column=3, value=round(st['inspected_or_produced_qty'], 2) if st['has_source_data'] else 0)
            ws.cell(row=r_idx, column=4, value=round(st['produced_qty'], 2) if st['has_source_data'] else 0)
            ws.cell(row=r_idx, column=5, value=round(st['defective_pairs_qty'], 2) if st['has_source_data'] else 0)
        else:
            ws.cell(row=r_idx, column=3, value=round(st['inspected_or_produced_qty'], 2) if st['has_source_data'] else 0)
            ws.cell(row=r_idx, column=4, value=round(st['a_grade_pairs'], 2) if st['has_source_data'] else 0)
            ws.cell(row=r_idx, column=5, value=round(st['rework_pairs_qty'], 2) if st['has_source_data'] else 0)

        # Total defect issues via SUMIFS referencing Drilldown_Normalized when data exists cleanly
        ws.cell(
            row=r_idx,
            column=6,
            value=f'=SUMIFS(Drilldown_Normalized!$L$2:$L${rollup_last_row},Drilldown_Normalized!$A$2:$A${rollup_last_row},"{metric}",Drilldown_Normalized!$B$2:$B${rollup_last_row},"{area}")',
        )
        if metric == 'HFPA':
            cov = pct(st['mapped_defect_issues'], st['total_defect_issues']) if st['total_defect_issues'] > 0 else None
            ws.cell(row=r_idx, column=7, value=(cov / 100.0) if cov is not None else 'N/A')
        else:
            ws.cell(row=r_idx, column=7, value=round(st['b_grade_pairs'] + st['c_grade_pairs'], 2) if st['has_source_data'] else 0)
        status_txt = {'ok': 'Đủ nguồn', 'no_defect': 'Không có lỗi', 'missing_data': 'Thiếu dữ liệu'}.get(st['scope_status'], st['scope_status'])
        ws.cell(row=r_idx, column=8, value=status_txt)
        ws.cell(row=r_idx, column=9, value='DR chưa chốt công thức' if st['has_source_data'] else 'Thiếu nguồn ' + area)
        _style_range(ws, 1, r_idx, 9, r_idx, font=BODY_FONT, fill=WARN_FILL if not st['has_source_data'] else (ZEBRA_FILL if r_idx % 2 == 1 else None))
        for num_c in (2, 3, 4, 5, 6):
            if isinstance(ws.cell(row=r_idx, column=num_c).value, (int, float)) or str(ws.cell(row=r_idx, column=num_c).value).startswith('='):
                ws.cell(row=r_idx, column=num_c).number_format = '#,##0'
        ws.cell(row=r_idx, column=7).number_format = '0.0%' if metric == 'HFPA' else '#,##0.#0'
        r_idx += 1

    # CLG total row = SUM of VH..VB
    clg_row = r_idx
    ws.cell(row=clg_row, column=1, value=scope_labels['CLG'])
    for c in (2, 3, 4, 5, 6):
        col_let = get_column_letter(c)
        ws.cell(row=clg_row, column=c, value=f'=SUM({col_let}{first_area_row}:{col_let}{clg_row - 1})')
        ws.cell(row=clg_row, column=c).number_format = '#,##0'
    ws.cell(row=clg_row, column=7, value=(pct(scope_totals['CLG']['mapped_defect_issues'], scope_totals['CLG']['total_defect_issues']) / 100.0) if (metric == 'HFPA' and scope_totals['CLG']['total_defect_issues'] > 0) else (f'=SUM(G{first_area_row}:G{clg_row - 1})' if metric == 'FTT' else 'N/A'))
    if metric == 'HFPA':
        ws.cell(row=clg_row, column=7).number_format = '0.0%'
    ws.cell(row=clg_row, column=8, value=result['review']['status_label'])
    ws.cell(row=clg_row, column=9, value=f"DR: {result['dr_config']['status']} | Level: {result['detail_support_level']}")
    _style_range(ws, 1, clg_row, 9, clg_row, font=BOLD_FONT, fill=TOTAL_FILL, border=TOTAL_BORDER)
    named_ranges[f'{metric}_scope_table'] = (f'A{hdr_row}', f'I{clg_row}')

    # 2) Top 3 Parent Groups Table (starts at clg_row + 2)
    grp_title_row = clg_row + 2
    ws.cell(row=grp_title_row, column=1, value=f'{metric} NIKE — TOP 3 NHÓM LỖI CLG & ĐÓNG GÓP THEO KHU')
    ws.merge_cells(start_row=grp_title_row, start_column=1, end_row=grp_title_row, end_column=9)
    _style_range(ws, 1, grp_title_row, 9, grp_title_row, font=SUBHDR_FONT, fill=SOFT_HEADER_FILL)

    g_hdr = grp_title_row + 1
    g_headers = ['Hạng', 'Nhóm chuẩn HFPA', 'VH', 'JV', 'VB', 'CLG (=VH+JV+VB)', 'Đóng góp VH (%)', 'Đóng góp JV (%)', 'Đóng góp VB (%)']
    for c, h in enumerate(g_headers, 1):
        ws.cell(row=g_hdr, column=c, value=h)
    _style_range(ws, 1, g_hdr, 9, g_hdr, font=HDR_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal='center', vertical='center', wrap_text=True))
    ws.row_dimensions[g_hdr].height = 26

    top_groups = result['top_groups']
    g_first_row = g_hdr + 1
    cur = g_first_row
    for g in top_groups:
        grp_lbl = g['group_label']
        ws.cell(row=cur, column=1, value=g['rank'])
        ws.cell(row=cur, column=2, value=grp_lbl)
        for col_offset, area in enumerate(AREAS, 3):
            ws.cell(
                row=cur,
                column=col_offset,
                value=f'=SUMIFS(Drilldown_Normalized!$L$2:$L${rollup_last_row},Drilldown_Normalized!$A$2:$A${rollup_last_row},"{metric}",Drilldown_Normalized!$B$2:$B${rollup_last_row},"{area}",Drilldown_Normalized!$H$2:$H${rollup_last_row},B{cur})',
            )
        ws.cell(row=cur, column=6, value=f'=SUM(C{cur}:E{cur})')
        ws.cell(row=cur, column=7, value=f'=IF(F{cur}>0,C{cur}/F{cur},"N/A")')
        ws.cell(row=cur, column=8, value=f'=IF(F{cur}>0,D{cur}/F{cur},"N/A")')
        ws.cell(row=cur, column=9, value=f'=IF(F{cur}>0,E{cur}/F{cur},"N/A")')
        _style_range(ws, 1, cur, 9, cur, font=BODY_FONT, fill=ZEBRA_FILL if cur % 2 == 1 else None)
        ws.cell(row=cur, column=1).alignment = Alignment(horizontal='center')
        for nc in (3, 4, 5, 6):
            ws.cell(row=cur, column=nc).number_format = '#,##0'
        for pc in (7, 8, 9):
            ws.cell(row=cur, column=pc).number_format = '0.0%'
        cur += 1

    # Add remainder / Other groups row below Top 3 for full reconciliation (NOT part of chart Top 3)
    if cur == g_first_row:
        ws.cell(row=cur, column=1, value='-')
        ws.cell(row=cur, column=2, value='(Không có nhóm lỗi > 0)')
        _style_range(ws, 1, cur, 9, cur, font=NOTE_FONT)
        cur += 1
    g_last_chart_row = max(g_first_row, cur - 1)

    rem_row = cur
    ws.cell(row=rem_row, column=1, value='Khác')
    ws.cell(row=rem_row, column=2, value='Nhóm khác / chưa ánh xạ')
    for col_offset, col_let in ((3, 'C'), (4, 'D'), (5, 'E')):
        scope_row_num = first_area_row + (col_offset - 3)
        ws.cell(row=rem_row, column=col_offset, value=f'=MAX(0, F{scope_row_num}-SUM({col_let}{g_first_row}:{col_let}{g_last_chart_row}))')
    ws.cell(row=rem_row, column=6, value=f'=SUM(C{rem_row}:E{rem_row})')
    ws.cell(row=rem_row, column=7, value=f'=IF(F{rem_row}>0,C{rem_row}/F{rem_row},"N/A")')
    ws.cell(row=rem_row, column=8, value=f'=IF(F{rem_row}>0,D{rem_row}/F{rem_row},"N/A")')
    ws.cell(row=rem_row, column=9, value=f'=IF(F{rem_row}>0,E{rem_row}/F{rem_row},"N/A")')
    _style_range(ws, 1, rem_row, 9, rem_row, font=BODY_FONT)
    for nc in (3, 4, 5, 6):
        ws.cell(row=rem_row, column=nc).number_format = '#,##0'
    for pc in (7, 8, 9):
        ws.cell(row=rem_row, column=pc).number_format = '0.0%'

    tot_row = rem_row + 1
    ws.cell(row=tot_row, column=1, value='Σ')
    ws.cell(row=tot_row, column=2, value=f'Tổng lỗi đã ánh xạ nhóm ({metric} CLG)')
    for nc in (3, 4, 5, 6):
        col_let = get_column_letter(nc)
        ws.cell(row=tot_row, column=nc, value=f'=SUM({col_let}{g_first_row}:{col_let}{rem_row})')
        ws.cell(row=tot_row, column=nc).number_format = '#,##0'
    for idx, src_c in enumerate(('C', 'D', 'E'), 7):
        ws.cell(row=tot_row, column=idx, value=f'=IF(F{tot_row}>0,{src_c}{tot_row}/F{tot_row},"N/A")')
        ws.cell(row=tot_row, column=idx).number_format = '0.0%'
    _style_range(ws, 1, tot_row, 9, tot_row, font=BOLD_FONT, fill=TOTAL_FILL, border=TOTAL_BORDER)
    named_ranges[f'{metric}_group_table'] = (f'A{g_hdr}', f'I{tot_row}')

    # Native Excel chart for Top 3 Groups by Area
    if top_groups:
        active_flags = tuple(
            any(float(g.get(f'{area.lower()}_qty') or 0.0) > 0 for g in top_groups)
            for area in AREAS
        )
        chart = _build_stacked_bar_chart(
            f'{metric} NIKE — Top 3 Nhóm Lỗi CLG Theo Khu (VH / JV / VB)',
            ws,
            data_min_col=3,
            data_max_col=5,
            hdr_row=g_hdr,
            cat_col=2,
            min_row=g_first_row,
            max_row=g_last_chart_row,
            y_title='Số lượng lỗi (Issues)',
            active_series=active_flags,
        )
        ws.add_chart(chart, f'K{start_row}')

    return named_ranges


def _write_hfpa_3month_table(
    ws: Any,
    *,
    start_row: int,
    period: str,
    date_from: str,
    date_to: str,
    monthly_hfpa: dict[tuple[str, str], dict[str, float]],
) -> tuple[str, str]:
    """Write dynamic 3-month HFPA VN/Indo/CLG table per AGENTS.md (no hardcoded screenshot data)."""
    months = _quarter_month_keys(date_from, date_to)[:3]
    if not months:
        months = ['N/A']
    _write_banner(
        ws,
        start_row,
        1,
        3 + len(months) + 1,
        f'HFPA NIKE — TỔNG HỢP {len(months)} THÁNG ĐÚNG KỲ {period} ({", ".join(_month_label(m) if "-" in m else m for m in months)}) [VN = VH+VH2 | Indo = JV+JV2]',
        fill=SUBTITLE_FILL,
    )
    hdr_row = start_row + 1
    headers = ['Nhà máy / Khu', 'Chỉ tiêu'] + [_month_label(m) if '-' in m else m for m in months] + [f'Tổng {len(months)} tháng', 'TB / tháng']
    for c, h in enumerate(headers, 1):
        ws.cell(row=hdr_row, column=c, value=h)
    _style_range(ws, 1, hdr_row, len(headers), hdr_row, font=HDR_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal='center', vertical='center', wrap_text=True))

    # Build 2 clean blocks: Audit Sample Size & Defect Issues (HFPA pairs vs issues clearly labeled)
    r = hdr_row + 1
    sum_col = 3 + len(months)
    avg_col = sum_col + 1
    first_m_col = 3
    last_m_col = 2 + len(months)

    for metric_label, field_key in (('Audit Sample Size (Đôi)', 'sample'), ('Defective Pairs (Đôi lỗi)', 'def_pairs'), ('Total Defect Issues (Lỗi)', 'issues')):
        sub_start = r
        factory_rows: dict[str, int] = {}
        for fac in ('VH', 'VH2', 'JV', 'JV2'):
            factory_rows[fac] = r
            ws.cell(row=r, column=1, value=fac)
            ws.cell(row=r, column=2, value=metric_label)
            for idx, ym in enumerate(months):
                val = monthly_hfpa.get((ym, fac), {}).get(field_key, 0.0)
                cell = ws.cell(row=r, column=first_m_col + idx, value=round(val, 2))
                cell.number_format = '#,##0'
            ws.cell(row=r, column=sum_col, value=f'=SUM({get_column_letter(first_m_col)}{r}:{get_column_letter(last_m_col)}{r})')
            ws.cell(row=r, column=avg_col, value=f'=AVERAGE({get_column_letter(first_m_col)}{r}:{get_column_letter(last_m_col)}{r})')
            ws.cell(row=r, column=sum_col).number_format = '#,##0'
            ws.cell(row=r, column=avg_col).number_format = '#,##0'
            _style_range(ws, 1, r, avg_col, r, font=BODY_FONT, fill=ZEBRA_FILL if r % 2 == 1 else None)
            r += 1

        # VN = VH + VH2; Indo = JV + JV2; CLG = VN + Indo
        vn_row = r
        ws.cell(row=vn_row, column=1, value='VN (VH + VH2)')
        ws.cell(row=vn_row, column=2, value=metric_label)
        for c in range(first_m_col, avg_col + 1):
            col_let = get_column_letter(c)
            if c <= last_m_col:
                ws.cell(row=vn_row, column=c, value=f'={col_let}{factory_rows["VH"]}+{col_let}{factory_rows["VH2"]}')
            elif c == sum_col:
                ws.cell(row=vn_row, column=c, value=f'=SUM({get_column_letter(first_m_col)}{vn_row}:{get_column_letter(last_m_col)}{vn_row})')
            else:
                ws.cell(row=vn_row, column=c, value=f'=AVERAGE({get_column_letter(first_m_col)}{vn_row}:{get_column_letter(last_m_col)}{vn_row})')
            ws.cell(row=vn_row, column=c).number_format = '#,##0'
        _style_range(ws, 1, vn_row, avg_col, vn_row, font=BOLD_FONT, fill=SOFT_HEADER_FILL)
        r += 1

        indo_row = r
        ws.cell(row=indo_row, column=1, value='Indo (JV + JV2)')
        ws.cell(row=indo_row, column=2, value=metric_label)
        for c in range(first_m_col, avg_col + 1):
            col_let = get_column_letter(c)
            if c <= last_m_col:
                ws.cell(row=indo_row, column=c, value=f'={col_let}{factory_rows["JV"]}+{col_let}{factory_rows["JV2"]}')
            elif c == sum_col:
                ws.cell(row=indo_row, column=c, value=f'=SUM({get_column_letter(first_m_col)}{indo_row}:{get_column_letter(last_m_col)}{indo_row})')
            else:
                ws.cell(row=indo_row, column=c, value=f'=AVERAGE({get_column_letter(first_m_col)}{indo_row}:{get_column_letter(last_m_col)}{indo_row})')
            ws.cell(row=indo_row, column=c).number_format = '#,##0'
        _style_range(ws, 1, indo_row, avg_col, indo_row, font=BOLD_FONT, fill=SOFT_HEADER_FILL)
        r += 1

        clg_row = r
        ws.cell(row=clg_row, column=1, value='CLG (VN + Indo)')
        ws.cell(row=clg_row, column=2, value=metric_label)
        for c in range(first_m_col, avg_col + 1):
            col_let = get_column_letter(c)
            if c <= last_m_col:
                ws.cell(row=clg_row, column=c, value=f'={col_let}{vn_row}+{col_let}{indo_row}')
            elif c == sum_col:
                ws.cell(row=clg_row, column=c, value=f'=SUM({get_column_letter(first_m_col)}{clg_row}:{get_column_letter(last_m_col)}{clg_row})')
            else:
                ws.cell(row=clg_row, column=c, value=f'=AVERAGE({get_column_letter(first_m_col)}{clg_row}:{get_column_letter(last_m_col)}{clg_row})')
            ws.cell(row=clg_row, column=c).number_format = '#,##0'
        _style_range(ws, 1, clg_row, avg_col, clg_row, font=BOLD_FONT, fill=TOTAL_FILL, border=TOTAL_BORDER)
        r += 1

    return f'A{start_row}', f'{get_column_letter(avg_col)}{r - 1}'


def _write_metric_detail_sheet(
    wb: openpyxl.Workbook,
    *,
    sheet_name: str,
    metric: str,
    result: dict[str, Any],
) -> dict[str, tuple[str, str]]:
    """Create dedicated per-metric Detail Matrix + Top Models + Drill-Down sheet."""
    ws = wb.create_sheet(sheet_name)
    ws.views.sheetView[0].showGridLines = True
    ranges: dict[str, tuple[str, str]] = {}

    _write_banner(
        ws,
        1,
        1,
        9,
        f'{metric} NIKE — PHÂN TÍCH LỖI CHI TIẾT THỰC TẾ & TOP MODEL SONG SONG ({result["period"]}: {result["date_from"]}..{result["date_to"]})',
        fill=TITLE_FILL,
        font=TITLE_FONT,
    )

    # Compact Scannable Visual Matrix (Group rows x CLG/VH/JV/VB cols -> ideal for PPT Slide CopyPicture)
    _write_banner(
        ws,
        3,
        1,
        5,
        f'{metric} — TOP 3 LỖI CHI TIẾT TRONG TỪNG NHÓM (MẪU SỐ = TỔNG NHÓM TẠI KHU ĐÓ)',
        fill=SUBTITLE_FILL,
    )
    m_hdr = 4
    matrix_cols = ['Nhóm chuẩn HFPA (Top 3 CLG)', 'CLG (Toàn CLG)', 'VH (Trừ Sandals)', 'JV (Indo)', 'VB (Factory Sandals)']
    for c, h in enumerate(matrix_cols, 1):
        ws.cell(row=m_hdr, column=c, value=h)
    _style_range(ws, 1, m_hdr, 5, m_hdr, font=HDR_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal='center', vertical='center', wrap_text=True))
    ws.row_dimensions[m_hdr].height = 24

    r = m_hdr + 1
    top_groups = result['top_groups']
    details_by_group = result['details_by_group']
    for g in top_groups:
        gid = g['group_id']
        ws.cell(row=r, column=1, value=f"#{g['rank']} {g['group_label']}\n(CLG: {g['clg_qty']:,.0f})")
        for idx, scope in enumerate(SCOPES, 2):
            cell = details_by_group.get(gid, {}).get(scope, {})
            top_d = cell.get('top_details', [])
            if cell.get('cell_status') == 'missing_scope_data':
                txt = f'{scope}: Thiếu dữ liệu nguồn'
            elif cell.get('cell_status') == 'no_defect':
                txt = 'Không có lỗi (0)'
            elif not top_d:
                txt = (
                    f"Tổng nhóm: {cell.get('group_total_qty', 0):,.0f}\n"
                    'Thiếu dữ liệu chi tiết (chỉ có cấp nhóm HFPA)'
                )
            else:
                lines = [
                    f"{d['rank']}. {d['detail_label']}: {d['qty']:,.0f} ({d['share_of_group_in_scope_pct']:.1f}%)"
                    if d['share_of_group_in_scope_pct'] is not None
                    else f"{d['rank']}. {d['detail_label']}: {d['qty']:,.0f} (N/A)"
                    for d in top_d
                ]
                if cell.get('unreconciled_gap_qty', 0) > 0:
                    lines.append(f"Chưa có chi tiết: {cell['unreconciled_gap_qty']:,.0f} (Coverage {cell['detail_coverage_pct']:.1f}%)")
                elif cell.get('other_mapped_details_qty', 0) > 0:
                    lines.append(f"Lỗi còn lại: {cell['other_mapped_details_qty']:,.0f} ({cell['other_mapped_details_pct']:.1f}%)")
                txt = '\n'.join(lines)
            ws.cell(row=r, column=idx, value=txt)
        _style_range(ws, 1, r, 5, r, font=BODY_FONT, fill=ZEBRA_FILL if r % 2 == 1 else None, alignment=Alignment(horizontal='left', vertical='top', wrap_text=True))
        ws.cell(row=r, column=1).font = BOLD_FONT
        ws.row_dimensions[r].height = 68
        r += 1
    if not top_groups:
        ws.cell(row=r, column=1, value='(Không có nhóm lỗi trong kỳ)')
        _style_range(ws, 1, r, 5, r, font=NOTE_FONT)
        r += 1
    ranges['detail_matrix'] = (f'A3', f'E{r - 1}')

    # Parallel Matrix 2: Top 3 Models per Group & Scope (never replacing Top details!)
    r += 1
    mod_start = r
    _write_banner(
        ws,
        r,
        1,
        5,
        f'{metric} — TOP 3 MODEL TRONG TỪNG NHÓM LỖI (CHIỀU PHÂN TÍCH SONG SONG)',
        fill=SUBTITLE_FILL,
    )
    r += 1
    for c, h in enumerate(matrix_cols, 1):
        ws.cell(row=r, column=c, value=h)
    _style_range(ws, 1, r, 5, r, font=HDR_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal='center', vertical='center', wrap_text=True))
    ws.row_dimensions[r].height = 24
    r += 1

    models_by_group = result['models_by_group']
    for g in top_groups:
        gid = g['group_id']
        ws.cell(row=r, column=1, value=f"#{g['rank']} {g['group_label']}")
        for idx, scope in enumerate(SCOPES, 2):
            mcell = models_by_group.get(gid, {}).get(scope, {})
            if mcell.get('cell_status') == 'missing_scope_data':
                txt = f'{scope}: Thiếu dữ liệu'
            elif mcell.get('cell_status') == 'no_defect':
                txt = 'Không có lỗi (0)'
            elif not mcell.get('top_models'):
                txt = '(Không có model phát sinh lỗi)'
            else:
                txt = '\n'.join(
                    f"{m['rank']}. {m['model_label']}: {m['qty']:,.0f} ({m['share_of_group_in_scope_pct']:.1f}%)"
                    if m['share_of_group_in_scope_pct'] is not None
                    else f"{m['rank']}. {m['model_label']}: {m['qty']:,.0f}"
                    for m in mcell['top_models']
                )
            ws.cell(row=r, column=idx, value=txt)
        _style_range(ws, 1, r, 5, r, font=BODY_FONT, fill=ZEBRA_FILL if r % 2 == 1 else None, alignment=Alignment(horizontal='left', vertical='top', wrap_text=True))
        ws.cell(row=r, column=1).font = BOLD_FONT
        ws.row_dimensions[r].height = 58
        r += 1
    ranges['model_matrix'] = (f'A{mod_start}', f'E{r - 1}')

    # Chartable Flat Detail Table & Drill-down (cols 1..10): strictly separates Denominator 1 vs Denominator 2!
    r += 2
    dd_start = r
    _write_banner(
        ws,
        r,
        1,
        11,
        f'{metric} — BẢNG DRILL-DOWN ĐẦY ĐỦ: TÁCH BIỆT TỶ TRỌNG TRONG NHÓM CLG (%) VÀ ĐÓNG GÓP KHU VÀO CÙNG LỖI CLG (%)',
        fill=SUBTITLE_FILL,
    )
    r += 1
    dd_hdrs = [
        'Nhóm chuẩn HFPA',
        'Hạng CLG',
        'Loại FTT NIKE',
        'Tên lỗi chi tiết thực tế',
        'CLG (D_CLG)',
        'VH (D_VH)',
        'JV (D_JV)',
        'VB (D_VB)',
        'Tỷ trọng trong nhóm CLG (% của G_CLG)',
        'Đóng góp VH vào lỗi (%) [% của D_CLG]',
        'Đóng góp JV/VB vào lỗi (% của D_CLG)',
    ]
    for c, h in enumerate(dd_hdrs, 1):
        ws.cell(row=r, column=c, value=h)
    _style_range(ws, 1, r, 11, r, font=HDR_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal='center', vertical='center', wrap_text=True))
    ws.row_dimensions[r].height = 30
    dd_hdr_row = r
    r += 1

    dd_first_data = r
    for g in top_groups:
        gid = g['group_id']
        dd = result['drilldown_by_group'].get(gid, {})
        for row_item in dd.get('details', []):
            ws.cell(row=r, column=1, value=g['group_label'])
            ws.cell(row=r, column=2, value=row_item['rank_clg'])
            ws.cell(row=r, column=3, value=row_item['nike_subtype_label'] or '-')
            ws.cell(row=r, column=4, value=row_item['detail_label'])
            ws.cell(row=r, column=5, value=f'=SUM(F{r}:H{r})')
            ws.cell(row=r, column=6, value=row_item['vh_qty'])
            ws.cell(row=r, column=7, value=row_item['jv_qty'])
            ws.cell(row=r, column=8, value=row_item['vb_qty'])
            g_tot = dd['group_total_clg']
            ws.cell(row=r, column=9, value=(row_item['share_of_clg_group_pct'] / 100.0) if row_item['share_of_clg_group_pct'] is not None else 'N/A')
            ws.cell(row=r, column=10, value=f'=IF(E{r}>0,F{r}/E{r},"N/A")')
            jv_p = row_item['jv_contrib_to_detail_pct']
            vb_p = row_item['vb_contrib_to_detail_pct']
            ws.cell(row=r, column=11, value=f"JV: {jv_p:.1f}% | VB: {vb_p:.1f}%" if jv_p is not None and vb_p is not None else 'N/A')
            _style_range(ws, 1, r, 11, r, font=BODY_FONT, fill=ZEBRA_FILL if r % 2 == 1 else None)
            for nc in (5, 6, 7, 8):
                ws.cell(row=r, column=nc).number_format = '#,##0'
            for pc in (9, 10):
                if ws.cell(row=r, column=pc).value != 'N/A':
                    ws.cell(row=r, column=pc).number_format = '0.0%'
            r += 1
        if dd.get('gap_clg_qty', 0) > 0:
            ws.cell(row=r, column=1, value=g['group_label'])
            ws.cell(row=r, column=2, value='Gap')
            ws.cell(row=r, column=3, value='-')
            ws.cell(row=r, column=4, value='Chưa có chi tiết (phần chênh control/nhóm)')
            ws.cell(row=r, column=5, value=dd['gap_clg_qty'])
            ws.cell(row=r, column=5).number_format = '#,##0'
            ws.cell(row=r, column=9, value=(pct(dd['gap_clg_qty'], dd['group_total_clg']) / 100.0) if dd['group_total_clg'] > 0 else 'N/A')
            ws.cell(row=r, column=9).number_format = '0.0%'
            _style_range(ws, 1, r, 11, r, font=NOTE_FONT, fill=WARN_FILL)
            r += 1

    if r > dd_first_data:
        flat_sorted = sorted(
            (
                (item['detail_label'], float(item['clg_qty']))
                for g in top_groups
                for item in result['drilldown_by_group'].get(g['group_id'], {}).get('details', [])
                if float(item['clg_qty']) > 0
            ),
            key=lambda kv: (-kv[1], kv[0]),
        )[:8]
        if flat_sorted:
            c_hdr = dd_hdr_row
            ws.cell(row=c_hdr, column=13, value='Lỗi chi tiết (Top CLG)')
            ws.cell(row=c_hdr, column=14, value='Số lỗi (CLG)')
            _style_range(ws, 13, c_hdr, 14, c_hdr, font=HDR_FONT, fill=HEADER_FILL)
            for idx_item, (lbl, qty_val) in enumerate(flat_sorted, 1):
                rr_c = c_hdr + idx_item
                ws.cell(row=rr_c, column=13, value=lbl)
                ws.cell(row=rr_c, column=14, value=qty_val)
                ws.cell(row=rr_c, column=14).number_format = '#,##0'
                _style_range(ws, 13, rr_c, 14, rr_c, font=BODY_FONT)
            ws.column_dimensions[get_column_letter(13)].width = 32
            ws.column_dimensions[get_column_letter(14)].width = 16
            chart = _build_horizontal_bar_chart(
                f'{metric} NIKE — Top Lỗi Chi Tiết Trong Top 3 Nhóm (CLG)',
                ws,
                val_col=14,
                hdr_row=c_hdr,
                cat_col=13,
                min_row=c_hdr + 1,
                max_row=c_hdr + len(flat_sorted),
                color='007A87',
                width=20.0,
                height=11.5,
                max_val=flat_sorted[0][1],
            )
            ws.add_chart(chart, f'G{m_hdr}')
    ranges['drilldown_table'] = (f'A{dd_start}', f'K{max(dd_start + 2, r - 1)}')

    widths = [30, 34, 34, 34, 34, 13, 13, 13, 20, 20, 24]
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    return ranges


def _write_catalog_mapping_review_sheet(
    wb: openpyxl.Workbook,
    *,
    catalog_entries: list[LineCatalogEntry],
    hfpa_result: dict[str, Any],
    ftt_result: dict[str, Any],
) -> None:
    ws = wb.create_sheet('Catalog_Mapping_&_Review')
    ws.views.sheetView[0].showGridLines = True

    _write_banner(ws, 1, 1, 9, 'KIỂM TRA DỮ LIỆU, DANH MỤC VB, MAPPING NIKE 3 TẦNG & CẤU HÌNH DR', fill=TITLE_FILL, font=TITLE_FONT)

    # 1. DR Unconfirmed Box
    dr = hfpa_result['dr_config']
    ws.cell(row=3, column=1, value='Trạng thái công thức DR:')
    ws.cell(row=3, column=2, value=dr['status'].upper())
    ws.cell(row=3, column=3, value=dr['note'])
    ws.merge_cells('C3:I3')
    _style_range(ws, 1, 3, 9, 3, font=BOLD_FONT, fill=INFO_FILL)

    # 2. Source Manifests
    r = 5
    _write_banner(ws, r, 1, 9, 'BẢNG MANIFEST NGUỒN ĐẦU VÀO (CHỈ HFPA NIKE & FTT NIKE)', fill=SUBTITLE_FILL)
    r += 1
    m_hdrs = ['Metric', 'Mode', 'File nguồn', 'SHA256 (prefix)', 'Sheet / HeaderRow', 'Tổng dòng', 'Giữ lại', 'Dedup bỏ', 'Mức chi tiết & Ghi chú']
    for c, h in enumerate(m_hdrs, 1):
        ws.cell(row=r, column=c, value=h)
    _style_range(ws, 1, r, 9, r, font=HDR_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal='center', wrap_text=True))
    r += 1
    for m in hfpa_result['manifests'] + ftt_result['manifests']:
        ws.append([
            m['metric'],
            m['report_mode'],
            m['file_name'],
            str(m['file_hash'])[:12],
            f"{m['sheet_name']} (H{m['header_row']})",
            m['total_rows'],
            m['kept_rows'],
            m['dedup_dropped_rows'],
            f"{m['detail_support_level']} | {'; '.join(m.get('coverage_notes', []))}",
        ])
        _style_range(ws, 1, r, 9, r, font=BODY_FONT, fill=ZEBRA_FILL if r % 2 == 1 else None)
        r += 1

    # 3. Review Items (Unmapped / Conflicts / Mismatches / Notes)
    r += 1
    _write_banner(ws, r, 1, 9, 'DANH SÁCH CẦN REVIEW (UNMAPPED / CHUYỀN XUNG ĐỘT / CONTROL MISMATCH / LƯU Ý)', fill=SUBTITLE_FILL)
    r += 1
    rv_hdrs = ['Loại review', 'Metric', 'Nhà máy/Xưởng/Chuyền', 'Nhãn gốc', 'Trạng thái', 'Chặn xuất?', 'Số lượng ảnh hưởng', 'Gợi ý / Lý do', 'Nguồn']
    for c, h in enumerate(rv_hdrs, 1):
        ws.cell(row=r, column=c, value=h)
    _style_range(ws, 1, r, 9, r, font=HDR_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal='center', wrap_text=True))
    r += 1
    all_reviews = hfpa_result['review']['items'] + ftt_result['review']['items']
    if not all_reviews:
        ws.cell(row=r, column=1, value='(Không có nhãn chưa ánh xạ hoặc xung đột chuyền)')
        _style_range(ws, 1, r, 9, r, font=NOTE_FONT)
        r += 1
    else:
        for item in all_reviews:
            ws.append([
                item['review_type'],
                item['metric'],
                f"{item['factory']} / {item['plant']} / {item['line_std']}",
                item['raw_label'],
                item['status'],
                'CÓ (Chặn)' if item['blocking'] else 'Không',
                item['affected_quantity'],
                f"{item['suggested_or_confirmed_mapping']} — {item['reason']}",
                item['provenance'],
            ])
            _style_range(ws, 1, r, 9, r, font=BODY_FONT, fill=ERROR_FILL if item['blocking'] else WARN_FILL)
            r += 1

    # 4. Line Catalog Factory Sandals
    r += 1
    _write_banner(ws, r, 1, 9, f'DANH MỤC CHUYỀN VB (GROUP FACTORY SANDALS TỪ FTT) — {len(catalog_entries)} CHUYỀN', fill=SUBTITLE_FILL)
    r += 1
    c_hdrs = ['Factory', 'Plant', 'Line Raw', 'Line Std', 'Source Group', 'Report Area', 'Effective From', 'Effective To', 'Provenance']
    for c, h in enumerate(c_hdrs, 1):
        ws.cell(row=r, column=c, value=h)
    _style_range(ws, 1, r, 9, r, font=HDR_FONT, fill=HEADER_FILL)
    r += 1
    for entry in catalog_entries:
        ws.append([
            entry.factory,
            entry.plant,
            entry.line_raw,
            entry.line_std,
            entry.source_group,
            entry.report_area,
            entry.effective_from,
            entry.effective_to,
            entry.provenance,
        ])
        _style_range(ws, 1, r, 9, r, font=BODY_FONT)
        r += 1

    # 5. Mapping Reference 11 <-> 25
    r += 1
    cfg = load_defect_mapping()
    _write_banner(ws, r, 1, 9, f"BẢNG QUY ĐỔI CHUẨN NIKE ({cfg['mapping_version']})", fill=SUBTITLE_FILL)
    r += 1
    for h_c, h in enumerate(('HFPA Standard Group ID', 'HFPA Standard Group Label', 'FTT Subtype ID', 'FTT Subtype Label', 'Aliases', 'Basis', '', '', ''), 1):
        if h:
            ws.cell(row=r, column=h_c, value=h)
    _style_range(ws, 1, r, 6, r, font=HDR_FONT, fill=HEADER_FILL)
    r += 1
    groups_by_id = cfg['groups_by_id']
    for sub in cfg['subtypes_by_id'].values():
        g_meta = groups_by_id[sub['standard_group_id']]
        ws.append([
            g_meta['id'],
            g_meta['label'],
            sub['id'],
            sub['label'],
            ', '.join(sub.get('aliases', [])),
            cfg.get('confirmed_by', ''),
        ])
        _style_range(ws, 1, r, 6, r, font=BODY_FONT)
        r += 1

    for idx, w in enumerate((18, 22, 30, 22, 24, 16, 16, 40, 28), 1):
        ws.column_dimensions[get_column_letter(idx)].width = w


def build_excel_workbook(
    output_path: Path | str,
    *,
    period: str,
    date_from: str,
    date_to: str,
    hfpa_result: dict[str, Any],
    ftt_result: dict[str, Any],
    hfpa_base: list[BaseRecord],
    hfpa_facts: list[DefectFact],
    ftt_base: list[BaseRecord],
    ftt_facts: list[DefectFact],
    catalog_entries: list[LineCatalogEntry],
) -> dict[str, Any]:
    """Create the complete Excel workbook with formulas & native Excel charts finalized."""
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws_over = wb.active
    ws_over.title = 'Overview_and_Charts'
    ws_over.views.sheetView[0].showGridLines = True

    rollup_last_row, monthly_hfpa = _write_drilldown_rollup(wb, hfpa_base, hfpa_facts, ftt_base, ftt_facts)

    _write_banner(
        ws_over,
        1,
        1,
        18,
        f'QUARTERLY NIKE DEFECT REVIEW ({period}: {date_from} đến {date_to}) — HFPA NIKE & FTT NIKE ĐỘC LẬP',
        fill=TITLE_FILL,
        font=TITLE_FONT,
    )
    ws_over.row_dimensions[1].height = 30
    ws_over.cell(row=2, column=1, value=f"Danh mục VB: {hfpa_result['catalog_version']} | Mapping: {hfpa_result['mapping_version']} | {hfpa_result['dr_config']['note']}")
    ws_over.merge_cells('A2:R2')
    _style_range(ws_over, 1, 2, 18, 2, font=NOTE_FONT, fill=INFO_FILL)

    hfpa_ranges = _write_metric_overview_block(
        ws_over,
        start_row=4,
        metric='HFPA',
        result=hfpa_result,
        rollup_last_row=rollup_last_row,
    )

    m3_range = _write_hfpa_3month_table(
        ws_over,
        start_row=22,
        period=period,
        date_from=date_from,
        date_to=date_to,
        monthly_hfpa=monthly_hfpa,
    )

    ftt_ranges = _write_metric_overview_block(
        ws_over,
        start_row=47,
        metric='FTT',
        result=ftt_result,
        rollup_last_row=rollup_last_row,
    )

    # Top Models Chart source block (per metric independently) at row 65
    m_chart_row = 65
    _write_banner(ws_over, m_chart_row, 1, 9, 'TOP 3 MODEL TRONG NHÓM LỖI #1 CLG (TÁCH RIÊNG HFPA NIKE VÀ FTT NIKE)', fill=SUBTITLE_FILL)
    m_hdr = m_chart_row + 1
    ws_over.cell(row=m_hdr, column=1, value='HFPA #1 Group Model')
    ws_over.cell(row=m_hdr, column=2, value='HFPA Lỗi (CLG)')
    ws_over.cell(row=m_hdr, column=4, value='FTT #1 Group Model')
    ws_over.cell(row=m_hdr, column=5, value='FTT Lỗi (CLG)')
    _style_range(ws_over, 1, m_hdr, 2, m_hdr, font=HDR_FONT, fill=HEADER_FILL)
    _style_range(ws_over, 4, m_hdr, 5, m_hdr, font=HDR_FONT, fill=HEADER_FILL)

    hfpa_top1_models = []
    if hfpa_result['top_groups']:
        g1h = hfpa_result['top_groups'][0]['group_id']
        hfpa_top1_models = hfpa_result['models_by_group'].get(g1h, {}).get('CLG', {}).get('top_models', [])
    ftt_top1_models = []
    if ftt_result['top_groups']:
        g1f = ftt_result['top_groups'][0]['group_id']
        ftt_top1_models = ftt_result['models_by_group'].get(g1f, {}).get('CLG', {}).get('top_models', [])

    max_m_rows = max(len(hfpa_top1_models), len(ftt_top1_models), 1)
    for idx in range(max_m_rows):
        rr = m_hdr + 1 + idx
        if idx < len(hfpa_top1_models):
            ws_over.cell(row=rr, column=1, value=hfpa_top1_models[idx]['model_label'])
            ws_over.cell(row=rr, column=2, value=hfpa_top1_models[idx]['qty'])
            ws_over.cell(row=rr, column=2).number_format = '#,##0'
        if idx < len(ftt_top1_models):
            ws_over.cell(row=rr, column=4, value=ftt_top1_models[idx]['model_label'])
            ws_over.cell(row=rr, column=5, value=ftt_top1_models[idx]['qty'])
            ws_over.cell(row=rr, column=5).number_format = '#,##0'
        _style_range(ws_over, 1, rr, 2, rr, font=BODY_FONT)
        _style_range(ws_over, 4, rr, 5, rr, font=BODY_FONT)

    if hfpa_top1_models:
        c_hm = _build_horizontal_bar_chart(
            f"HFPA NIKE — Top Model ({hfpa_result['top_groups'][0]['group_label']}, CLG)",
            ws_over,
            val_col=2,
            hdr_row=m_hdr,
            cat_col=1,
            min_row=m_hdr + 1,
            max_row=m_hdr + len(hfpa_top1_models),
            color='1F4E78',
            width=18.5,
            height=10.0,
            max_val=float(hfpa_top1_models[0]['qty']),
        )
        ws_over.add_chart(c_hm, 'K22')

    if ftt_top1_models:
        c_fm = _build_horizontal_bar_chart(
            f"FTT NIKE — Top Model ({ftt_result['top_groups'][0]['group_label']}, CLG)",
            ws_over,
            val_col=5,
            hdr_row=m_hdr,
            cat_col=4,
            min_row=m_hdr + 1,
            max_row=m_hdr + len(ftt_top1_models),
            color='D97706',
            width=18.5,
            height=10.0,
            max_val=float(ftt_top1_models[0]['qty']),
        )
        ws_over.add_chart(c_fm, 'K65')

    for col_idx, width in enumerate((25, 34, 14, 14, 14, 17, 16, 16, 30), 1):
        ws_over.column_dimensions[get_column_letter(col_idx)].width = width

    hfpa_det_ranges = _write_metric_detail_sheet(wb, sheet_name='Detail_Breakdown_HFPA', metric='HFPA', result=hfpa_result)
    ftt_det_ranges = _write_metric_detail_sheet(wb, sheet_name='Detail_Breakdown_FTT', metric='FTT', result=ftt_result)
    _write_catalog_mapping_review_sheet(wb, catalog_entries=catalog_entries, hfpa_result=hfpa_result, ftt_result=ftt_result)

    total_charts = (
        len(ws_over._charts)
        + len(wb['Detail_Breakdown_HFPA']._charts)
        + len(wb['Detail_Breakdown_FTT']._charts)
    )
    wb.save(out)
    wb.close()

    return {
        'excel_path': str(out.resolve()),
        'chart_sheet': 'Overview_and_Charts',
        'ranges': {
            'hfpa_scope_table': ('Overview_and_Charts', *hfpa_ranges['HFPA_scope_table']),
            'hfpa_group_table': ('Overview_and_Charts', *hfpa_ranges['HFPA_group_table']),
            'hfpa_3month_table': ('Overview_and_Charts', *m3_range),
            'ftt_scope_table': ('Overview_and_Charts', *ftt_ranges['FTT_scope_table']),
            'ftt_group_table': ('Overview_and_Charts', *ftt_ranges['FTT_group_table']),
            'hfpa_detail_matrix': ('Detail_Breakdown_HFPA', *hfpa_det_ranges['detail_matrix']),
            'hfpa_model_matrix': ('Detail_Breakdown_HFPA', *hfpa_det_ranges['model_matrix']),
            'ftt_detail_matrix': ('Detail_Breakdown_FTT', *ftt_det_ranges['detail_matrix']),
            'ftt_model_matrix': ('Detail_Breakdown_FTT', *ftt_det_ranges['model_matrix']),
        },
        'native_charts_count': total_charts,
    }
