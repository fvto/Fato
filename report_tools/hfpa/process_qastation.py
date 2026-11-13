"""
HFPA & FTT Data Processing Tool
Standard Operating Procedure (SOP): Database/2026 FTT&HFPA EN.pdf

Fully automates the 17-page official SOP workflow:
1. Pages 1-5: Clean HFPA (Quality Tracking) data & build 1st Pivot Table with Plant between Station and Line.
2. Pages 6-15: Reconcile HFPA (Quality Tracking) vs HFPA (Mes410) by Factory, Line, Date, Model, and correct mismatched counts.
3. Page 16: Calculate Top 5 Models sorted descending by Fail Q'ty (DEFECT RATE = Fail Q'ty / Sum Insp Q'ty).
4. Page 17: Calculate Top Defect Types per Top Model (Defect Rate% = Issues Qty / Total Issue Q'ty) transposed to summary table.
5. Export separate combined workbooks and auto-populate Database/HFPA_Template.xlsx.
"""

import os
import sys
import glob
import re
import datetime
from typing import Optional

if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from file_utils import ensure_file_writable, auto_archive_previous_months, detect_file_month, detect_file_month_and_year

try:
    from hfpa_common import (
        normalize_model_name,
        normalize_defect_name,
        compute_model_ranks_mes410,
        compute_top5_models,
        compute_top3_defects,
        compare_defect_shares,
    )
except ImportError:
    try:
        from .hfpa_common import (
            normalize_model_name,
            normalize_defect_name,
            compute_model_ranks_mes410,
            compute_top5_models,
            compute_top3_defects,
            compare_defect_shares,
        )
    except Exception:
        import sys
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from hfpa_common import (
            normalize_model_name,
            normalize_defect_name,
            compute_model_ranks_mes410,
            compute_top5_models,
            compute_top3_defects,
            compare_defect_shares,
        )

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FTT_DIR = os.path.join(BASE_DIR, "Input", "FTT")
INPUT_HFPA_DIR = os.path.join(BASE_DIR, "Input", "HFPA")
OUTPUT_DIR = os.path.join(BASE_DIR, "Output")
COLOR_TEMPLATE_PATH = os.path.join(BASE_DIR, "Color_template.xlsx")

# Expected Factories
FACTORIES = ["JV", "JV2", "VH", "VH2"]


def extract_fty_from_filename(filename: str) -> str:
    """Extract factory identifier (JV, JV2, VH, VH2) from filename."""
    base = os.path.splitext(os.path.basename(filename))[0]
    for fty in ["JV2", "VH2", "JV", "VH"]:
        if re.search(rf"[-_\s]{fty}$", base, re.IGNORECASE) or re.search(rf"[-_\s]{fty}[-_\s]", base, re.IGNORECASE):
            return fty
    for fty in ["JV2", "VH2", "JV", "VH"]:
        if fty in base.upper():
            return fty
    return "UNKNOWN"


def load_color_template():
    """Load defect color mappings from Color_template.xlsx."""
    try:
        from .agents.color_manager import ColorManager
    except (ImportError, ValueError):
        from agents.color_manager import ColorManager
    return ColorManager(COLOR_TEMPLATE_PATH).load_color_map()


def purge_colormap_sheets(wb):
    """Purge any redundant ColorMap_ sheets from workbook."""
    purged = []
    for s_name in list(wb.sheetnames):
        if s_name.startswith("ColorMap_"):
            del wb[s_name]
            purged.append(s_name)
    return purged


def embed_color_template_sheets(wb, color_file_path=COLOR_TEMPLATE_PATH):
    """Deprecated: Color_template.xlsx is read directly; do not duplicate sheets."""
    return purge_colormap_sheets(wb)


def parse_plant_and_line(val):
    """
    Parse Plant and Line from strings like 'N1 Line 1', 'Line 1', 'A Line1', 'A10 Line 1'
    Returns: (Plant, Line)
    """
    if pd.isna(val):
        return "", ""
    s = str(val).strip()
    m = re.match(r"^([A-Za-z0-9]+)\s*Line\s*([0-9]+)", s, re.IGNORECASE)
    if m:
        return m.group(1).upper(), m.group(2)
    m2 = re.match(r"^Line\s*([0-9]+)", s, re.IGNORECASE)
    if m2:
        return "", m2.group(1)
    # numeric line
    cleaned = re.sub(r"[^0-9]", "", s)
    return "", cleaned if cleaned else s


def load_and_combine_ftt():
    """Discover, load and combine all QAStation/FTT files into a single DataFrame."""
    ftt_files = glob.glob(os.path.join(INPUT_FTT_DIR, "*.xlsx")) + glob.glob(os.path.join(INPUT_FTT_DIR, "*.xls"))
    if not ftt_files:
        raise FileNotFoundError(f"No QAStation/FTT files found in {INPUT_FTT_DIR}")

    print(f"\n[1/5] Loading & combining {len(ftt_files)} QAStation/FTT files...")
    df_list = []
    for filepath in sorted(ftt_files):
        filename = os.path.basename(filepath)
        fty = extract_fty_from_filename(filename)
        print(f"   -> Reading: {filename} (Detected FTY: {fty})")
        df = pd.read_excel(filepath)
        df["FTY"] = fty
        df["Source_File"] = filename

        # Standardize column names between QAStation format and FTT-Internal format
        # If QAStation format: 'InspDate', 'Line', 'ShoeName', 'InspQty', 'FailQty', 'Station', 'Issues', 'IssueQty'
        if "InspDate" in df.columns:
            df["Date_Raw"] = df["InspDate"]
            df["Model_Raw"] = df["ShoeName"] if "ShoeName" in df.columns else df.get("Shoename", "")
            df["InspQty_Raw"] = pd.to_numeric(df["InspQty"], errors="coerce").fillna(0)
            df["FailQty_Raw"] = pd.to_numeric(df["FailQty"], errors="coerce").fillna(0)
            df["Issues_Raw"] = df["Issues"] if "Issues" in df.columns else ""
            df["IssueQty_Raw"] = pd.to_numeric(df["IssueQty"], errors="coerce").fillna(0)
            df["Station_Raw"] = df["Station"] if "Station" in df.columns else "HFPA"
            # Parse Plant & Line from 'Line' col
            plants, lines = zip(*df["Line"].apply(parse_plant_and_line))
            df["Plant_Parsed"] = plants
            df["Line_Parsed"] = lines
        else:
            # FTT-Internal format
            df["Date_Raw"] = df["Date"]
            df["Model_Raw"] = df["Model"] if "Model" in df.columns else df.get("ShoeName", "")
            df["InspQty_Raw"] = pd.to_numeric(df["Total Pair Produced"], errors="coerce").fillna(0)
            df["FailQty_Raw"] = pd.to_numeric(df["Issues Q'ty"], errors="coerce").fillna(0)
            df["Issues_Raw"] = df["Defect Issues"] if "Defect Issues" in df.columns else ""
            df["IssueQty_Raw"] = pd.to_numeric(df["Issues Q'ty"], errors="coerce").fillna(0)
            df["Station_Raw"] = df["Station"] if "Station" in df.columns else "Assembly"
            df["Plant_Parsed"] = df["Plant"].astype(str).str.strip().str.upper() if "Plant" in df.columns else ""
            _, lines = zip(*df["Line"].apply(parse_plant_and_line))
            df["Line_Parsed"] = lines

        df_list.append(df)

    combined_ftt = pd.concat(df_list, ignore_index=True)
    print(f"   [+] QAStation/FTT Combined rows: {len(combined_ftt):,}")
    return combined_ftt


def load_and_combine_hfpa():
    """Discover, load and combine all HFPA files into a single DataFrame."""
    hfpa_files = glob.glob(os.path.join(INPUT_HFPA_DIR, "*.xlsx")) + glob.glob(os.path.join(INPUT_HFPA_DIR, "*.xls"))
    if not hfpa_files:
        raise FileNotFoundError(f"No HFPA files found in {INPUT_HFPA_DIR}")

    print(f"\n[2/5] Loading & combining {len(hfpa_files)} HFPA files...")
    df_list = []
    for filepath in sorted(hfpa_files):
        filename = os.path.basename(filepath)
        fty = extract_fty_from_filename(filename)
        print(f"   -> Reading: {filename} (Detected FTY: {fty})")
        df = pd.read_excel(filepath)
        df["FTY"] = fty
        df["Source_File"] = filename
        df_list.append(df)

    combined_hfpa = pd.concat(df_list, ignore_index=True)
    print(f"   [+] HFPA Combined rows: {len(combined_hfpa):,}")
    return combined_hfpa


def validate_and_correct_data(qa_df: pd.DataFrame, hfpa_df: pd.DataFrame):
    """Reconcile with the same model-level rules as the dashboard."""
    from agents.reconciliation_engine_agent import ReconciliationEngineAgent

    corrected, lots, audit = ReconciliationEngineAgent.reconcile_and_validate(qa_df, hfpa_df)
    corrected.attrs["hfpa_source_rows"] = qa_df.to_dict("records")
    corrected.attrs["hfpa_audit_records"] = audit
    return corrected, hfpa_df, lots

def generate_pivot1_and_top5(qa_validated: pd.DataFrame, hfpa_df: Optional[pd.DataFrame] = None):
    """Use the dashboard inspection pivot and corrected Quality Tracking rankings."""
    from agents.reconciliation_engine_agent import ReconciliationEngineAgent

    p1 = ReconciliationEngineAgent.inspection_pivot(qa_validated)
    source = hfpa_df if hfpa_df is not None else pd.DataFrame(columns=["FTY"])
    all_top5, top5 = compute_top5_models(source, qa_p1_df=p1)
    return p1, all_top5, top5


def generate_pivot2_top_defects(qa_validated: pd.DataFrame, top5_dict: dict, hfpa_df: Optional[pd.DataFrame] = None):
    """
    Pivot Table #2 — defect types
    Fields:
    - Defect type, Defect quantity
    - Top 3 defect types per selected Top 5 shoe
    - Defect % calculations
    """
    print("\n[5/5] Generating Pivot #2 (Defect Types & Top 3 per shoe)...")
    color_map = load_color_template()
    pivot2_df = compute_top3_defects(qa_validated, top5_dict, color_map=color_map)
    return pivot2_df


def style_excel_sheet(ws, title: str):
    """Apply modern corporate styling to openpyxl worksheet."""
    ws.views.sheetView[0].showGridLines = True

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    border_thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = border_thin

    max_format_row = min(ws.max_row, 500) if ws.max_row > 1000 else ws.max_row
    for row_idx in range(2, max_format_row + 1):
        fill_color = "F2F5F9" if row_idx % 2 == 0 else "FFFFFF"
        row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = border_thin
            if not cell.fill.start_color.rgb or cell.fill.start_color.rgb == "00000000":
                cell.fill = row_fill

            val = cell.value
            if isinstance(val, (int, float)):
                header_name = str(ws.cell(row=1, column=col_idx).value or "").lower()
                if "%" in header_name or "rate" in header_name or "share" in header_name:
                    cell.number_format = "0.00%"
                    cell.alignment = align_right
                elif "qty" in header_name or "defects" in header_name or "fail" in header_name or "insp" in header_name or "sample" in header_name:
                    cell.number_format = "#,##0"
                    cell.alignment = align_right
                elif isinstance(val, float):
                    cell.number_format = "#,##0.0"
                    cell.alignment = align_right
                else:
                    cell.number_format = "#,##0"
                    cell.alignment = align_right
            elif isinstance(val, (datetime.date, datetime.datetime)):
                cell.number_format = "YYYY-MM-DD"
                cell.alignment = align_center

    for col_idx in range(1, ws.max_column + 1):
        sample_vals = [str(ws.cell(r, col_idx).value or "") for r in range(1, min(ws.max_row, 50) + 1)]
        max_len = max(len(v) for v in sample_vals) if sample_vals else 10
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)


def export_deliverables(qa_validated: pd.DataFrame, hfpa_combined: pd.DataFrame,
                        lot_comparison: pd.DataFrame, p1_df: pd.DataFrame,
                        all_top5: pd.DataFrame, p2_df: pd.DataFrame,
                        active_month: str = "Jun", active_year: str = "2026"):
    """Export with the same source preservation and audit trail as the dashboard."""
    from agents.reconciliation_engine_agent import DataBus, ReconciliationEngineAgent

    source_rows = qa_validated.attrs.get("hfpa_source_rows")
    if source_rows is None:
        raise ValueError("Cần gọi validate_and_correct_data trước khi xuất để giữ đúng dữ liệu nguồn và lịch sử hiệu chỉnh.")
    bus = DataBus(
        active_month=active_month, active_year=active_year,
        qa_raw=pd.DataFrame(source_rows), qa_validated=qa_validated, hfpa_raw=hfpa_combined,
        hfpa_clean=hfpa_combined, lot_comparison=lot_comparison, pivot1_df=p1_df,
        all_top5_df=all_top5, top5_dict=dict(tuple(all_top5.groupby("FTY"))),
        pivot2_df=p2_df, audit_records=qa_validated.attrs.get("hfpa_audit_records", []),
    )
    engine = ReconciliationEngineAgent(BASE_DIR)
    engine.export_master_workbook(bus)
    engine.export_analysis_and_audit(bus)


def main():
    from agents.hfpa_agent_orchestrator import HFPAOrchestrator

    if not HFPAOrchestrator(BASE_DIR).run_pipeline():
        raise SystemExit(1)


if __name__ == "__main__":
    main()
