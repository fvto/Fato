"""
hfpa_common.py
Single Source of Truth for HFPA normalization, ranking, tie-breaking, and cross-deliverable calculations.
"""

from __future__ import annotations

import re
import textwrap
from typing import Any, Dict, List, Optional, Tuple
import numpy as np
import pandas as pd

FACTORIES = ["VH", "VH2", "JV", "JV2"]
HFPA_DEFECT_BLOCKS = {"VH": (21, 22, 26), "VH2": (27, 28, 32),
                      "JV": (33, 34, 38), "JV2": (39, 40, 44)}


def hfpa_defect_columns(sheet, header_row):
    """Read every populated defect header, including columns beyond the old template."""
    return [column for column in range(13, sheet.max_column + 1)
            if str(sheet.cell(header_row, column).value or "").strip()]


def hfpa_defect_header_height(names):
    # Conservative allowance for 10pt wrapped headings in 16-character-wide columns.
    lines = max((sum(max(1, len(textwrap.wrap(line, width=12)))
                     for line in str(name).splitlines()) for name in names), default=1)
    height = max(48, 14 * lines + 12)
    if height > 409:
        raise ValueError("Tên lỗi quá dài để hiển thị đầy đủ trong một hàng Excel; cần kiểm tra tên lỗi nguồn.")
    return height


def hfpa_defect_layout_issues(sheet):
    """Check framing independently of the renderer before delivering the workbook."""
    from openpyxl.utils.cell import range_boundaries

    issues = []
    print_ranges = [range_boundaries(area) for area in
                    re.findall(r"\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+", str(sheet.print_area))]

    def record(factory, cell, reason):
        issues.append(dict(factory=factory, cell=cell.coordinate, reason=reason))

    def border_style(cell, side):
        return getattr(getattr(cell.border, side), "style", None)

    for factory, (header, start, end) in HFPA_DEFECT_BLOCKS.items():
        columns = hfpa_defect_columns(sheet, header)
        last = max(columns, default=15)
        if print_ranges and not any(min_col <= 11 and max_col >= last
                                    and min_row <= header and max_row >= end
                                    for min_col, min_row, max_col, max_row in print_ranges):
            record(factory, sheet.cell(header, last), "Bảng lỗi vượt vùng in đã khai báo.")
        if columns and columns != list(range(13, last + 1)):
            record(factory, sheet.cell(header, 13), "Có cột trống xen giữa các tiêu đề lỗi.")
        expected_height = hfpa_defect_header_height(sheet.cell(header, c).value for c in columns)
        if (sheet.row_dimensions[header].height or 15) < expected_height:
            record(factory, sheet.cell(header, 13), "Chiều cao hàng chưa đủ cho tên lỗi xuống dòng.")
        for column in range(11, last + 1):
            cell = sheet.cell(header, column)
            if border_style(cell, "top") != "medium" or border_style(cell, "bottom") != "medium":
                record(factory, cell, "Tiêu đề lỗi nằm ngoài khung bảng.")
            if column >= 13 and not cell.alignment.wrap_text:
                record(factory, cell, "Tên lỗi chưa được tự xuống dòng.")
            if column == last and border_style(cell, "right") != "medium":
                record(factory, cell, "Thiếu viền phải tại cột lỗi cuối.")
        for row in range(start, end + 1):
            if sheet.cell(row, 4).value != sheet.cell(row, 12).value:
                record(factory, sheet.cell(row, 12), "Model bảng lỗi không khớp thứ tự Top 5.")
            for column in range(13, sheet.max_column + 1):
                cell = sheet.cell(row, column)
                value = cell.value
                placeholder = column == 13 and value in ("Không có lỗi", "Thiếu dữ liệu (QAStation)")
                if value is not None and column not in columns and not placeholder:
                    record(factory, cell, "Có số liệu lỗi nhưng không có tiêu đề trong khung bảng.")
                if column <= last:
                    if not all(border_style(cell, side) for side in ("left", "right", "top", "bottom")):
                        record(factory, cell, "Ô lỗi thiếu đường viền bảng.")
                    if column == last and border_style(cell, "right") != "medium":
                        record(factory, cell, "Số liệu vượt viền phải của bảng.")
                    if row == end and border_style(cell, "bottom") != "medium":
                        record(factory, cell, "Thiếu viền dưới bảng lỗi.")
    return issues


FACTORY_ALIASES = {
    "VH": "VH", "CL_VH": "VH", "CL_VB": "VH", "VB": "VH",
    "VH2": "VH2", "CL_VH2": "VH2",
    "JV": "JV", "CL_JV": "JV", "CL_JV1": "JV", "JV1": "JV",
    "JV2": "JV2", "CL_JV2": "JV2",
}
MES_ISSUE_TOTAL = "Defect Q'ty (Total Defect Q'ty)"
MES_METADATA = {
    "Factory Group Cd_UccId", "SO#", "Country", "Season", "Audit Date (dd-mmm-yyyy)", "Audit Date",
    "Internal AUDIT ID (YYMMDDNO+Plant+Line)", "Plant", "Line", "Auditor",
    "PO Nbr (header+line item nbr)", "Model", "Model 2", "Style Nbr", "Colorway Cd", "Time table",
    "Produced Q'ty", "Audit Sample Size", "Audit Sample", "Defective Q'ty(pair)", "Defective Rate",
    "Pass/Fail", MES_ISSUE_TOTAL, "FTY", "Source_File", "Date", "Date_Std", "Plant_Std",
    "Line_Std", "Model_Std", "ShoeName_Std",
}


def mes410_defect_columns(frame):
    columns = list(frame.columns)
    if MES_ISSUE_TOTAL in columns:
        columns = columns[columns.index(MES_ISSUE_TOTAL) + 1:]
    return [name for name in columns if name not in MES_METADATA]


def normalize_model_name(val: Any) -> str:
    """Normalize shoe model name: strips whitespace, collapses internal whitespace and newlines."""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def normalize_defect_name(val: Any) -> str:
    """Normalize defect issue name: strips whitespace, collapses internal whitespace."""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def compute_model_ranks_mes410(
    hfpa_df: pd.DataFrame,
) -> Tuple[int, int, pd.DataFrame]:
    """
    Summarize audited sample pairs, defective pairs, and ranking for models in MES410 rows.

    Tie-breaking rule for Top 5 Models:
    1. Defective pairs descending (số đôi lỗi của model giảm dần)
    2. Audit sample size descending (số đôi kiểm của model giảm dần)
    3. Normalized model name ascending (A-Z)

    Returns:
    (total_sample_size, total_defective_pairs, ranked_models_df)
    """
    if hfpa_df.empty:
        return 0, 0, pd.DataFrame(columns=["Model", "Failures", "Samples"])

    sample_col = "Audit Sample Size" if "Audit Sample Size" in hfpa_df.columns else "Audit Sample"
    fail_col = next(
        (col for col in hfpa_df.columns if "Defective" in col and "pair" in col),
        "Defective Q'ty(pair)",
    )
    samples = pd.to_numeric(hfpa_df[sample_col], errors="coerce").fillna(0)
    failures = pd.to_numeric(hfpa_df[fail_col], errors="coerce").fillna(0)

    tot_samples = int(samples.sum())
    tot_failures = int(failures.sum())

    model_col = "Model 2" if "Model 2" in hfpa_df.columns else "Model"
    temp_df = pd.DataFrame({
        "Model": hfpa_df[model_col].apply(normalize_model_name),
        "Samples": samples,
        "Failures": failures,
    })

    temp_df = temp_df[temp_df["Model"] != ""]
    if temp_df.empty:
        return tot_samples, tot_failures, pd.DataFrame(columns=["Model", "Failures", "Samples"])

    grp = temp_df.groupby("Model", as_index=False).agg(
        Failures=("Failures", "sum"),
        Samples=("Samples", "sum"),
    )

    grp["Failures"] = grp["Failures"].astype(int)
    grp["Samples"] = grp["Samples"].astype(int)

    # Apply consistent tie-breaking:
    grp = grp.sort_values(
        by=["Failures", "Samples", "Model"],
        ascending=[False, False, True],
        kind="stable"
    ).reset_index(drop=True)

    return tot_samples, tot_failures, grp


def compute_top5_models(
    hfpa_df: pd.DataFrame,
    qa_p1_df: Optional[pd.DataFrame] = None,
) -> Tuple[pd.DataFrame, Dict[str, pd.DataFrame]]:
    """
    Generate Top 5 Models per factory and for Grand Total.
    Uses corrected Quality Tracking when provided, with MES410 for standalone callers.
    DR% = model defective pairs / factory total inspection sample size.
    """
    top5_dict: Dict[str, pd.DataFrame] = {}

    for fty in FACTORIES + ["Grand Total"]:
        mes_rows = hfpa_df if fty == "Grand Total" else hfpa_df[hfpa_df["FTY"] == fty]
        if qa_p1_df is not None:
            qa_rows = qa_p1_df if fty == "Grand Total" else qa_p1_df[qa_p1_df["FTY"] == fty]
            tot_sample = int(qa_rows["InspQty"].sum())
            ranked_df = qa_rows.groupby("Shoename", as_index=False).agg(
                Failures=("FailQty", "sum"),
                Samples=("InspQty", "sum"),
            ).rename(columns={"Shoename": "Model"})
            ranked_df = ranked_df[ranked_df["Model"] != ""].sort_values(
                ["Failures", "Samples", "Model"], ascending=[False, False, True]
            )
        else:
            tot_sample, _, ranked_df = compute_model_ranks_mes410(mes_rows)

        top5_sub = ranked_df.head(5).copy()
        models = top5_sub["Model"].tolist()
        validated_fails = top5_sub["Failures"].to_numpy()

        if qa_p1_df is not None:
            qa_fails_map = qa_rows.groupby("Shoename")["FailQty_Original"].sum().to_dict()
            orig_fails = np.array([qa_fails_map.get(m, 0) for m in models], dtype=int)
        else:
            orig_fails = validated_fails.copy()

        t5 = pd.DataFrame({
            "Shoename": models,
            "Sum of InspQty": tot_sample,
            "Sum of FailQty (Original)": orig_fails,
            "Sum of FailQty (Validated)": validated_fails,
            "DR%": validated_fails / tot_sample if tot_sample > 0 else 0.0,
            "Rank": range(1, len(models) + 1),
            "FTY": fty,
        })
        top5_dict[fty] = t5

    all_top5 = pd.concat([top5_dict[f] for f in FACTORIES + ["Grand Total"]], ignore_index=True)
    return all_top5, top5_dict


def compute_top3_defects(
    qa_val_df: pd.DataFrame,
    top5_dict: Dict[str, pd.DataFrame],
    color_map: Optional[Dict[str, str]] = None,
    mes410_fallback_df: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """
    Generate Top 3 Defect Types per Top 5 Model.
    Top 3 uses IssueQty divided by total model issues.
    If QAStation has no defects for a model, attempts MES410 fallback or flags missing data.

    Tie-breaking rule for Top 3 Defects:
    1. IssueQty descending (số lượt lỗi giảm dần)
    2. Original Quality Tracking order when issue counts are equal.
    """
    color_map = color_map or {}
    p2_records = []

    qa_df = qa_val_df.copy()
    if "ShoeName_Std" not in qa_df.columns:
        model_src = qa_df["Model_Raw"] if "Model_Raw" in qa_df.columns else qa_df.get("ShoeName", "")
        qa_df["ShoeName_Std"] = model_src.apply(normalize_model_name)
    if "Issues_Std" not in qa_df.columns:
        issues_src = qa_df["Issues_Raw"] if "Issues_Raw" in qa_df.columns else qa_df.get("Issues", "")
        qa_df["Issues_Std"] = issues_src.apply(normalize_defect_name)

    issue_qty_col = "IssueQty_Validated" if "IssueQty_Validated" in qa_df.columns else "IssueQty_Raw"
    if issue_qty_col not in qa_df.columns:
        issue_qty_col = "IssueQty" if "IssueQty" in qa_df.columns else "Issues Q'ty"
    qa_df["IssueQty_Val"] = pd.to_numeric(qa_df.get(issue_qty_col, 0), errors="coerce").fillna(0).astype(int)

    for fty in FACTORIES + ["Grand Total"]:
        if fty not in top5_dict or top5_dict[fty].empty:
            continue
        top5_models = top5_dict[fty]["Shoename"].tolist()
        sub_df = (
            qa_df[qa_df["ShoeName_Std"].isin(top5_models)]
            if fty == "Grand Total"
            else qa_df[(qa_df["FTY"] == fty) & (qa_df["ShoeName_Std"].isin(top5_models))]
        )
        defect_df = sub_df[(sub_df["IssueQty_Val"] > 0) & (sub_df["Issues_Std"] != "")]

        for model in top5_models:
            m_df = defect_df[defect_df["ShoeName_Std"] == model]
            tot_defects = int(round(m_df["IssueQty_Val"].sum()))

            if tot_defects == 0:
                model_rows = sub_df[sub_df["ShoeName_Std"] == model]
                failed_col = next((col for col in ("FailQty_Validated", "FailQty_Raw", "FailQty")
                                   if col in model_rows), None)
                no_defects = (not model_rows.empty and failed_col is not None
                              and pd.to_numeric(model_rows[failed_col], errors="coerce").eq(0).all())
                if no_defects:
                    p2_records.append({
                        "FTY": fty, "Shoename": model, "Defect Rank": "Không có lỗi",
                        "Defect Type": "N/A", "Defect Quantity": 0, "Total Shoe Defects": 0,
                        "Defect Share %": 0.0, "Spec Ratio (Total/Defects)": 0.0,
                        "Color HEX": "", "Data Status": "No_Defects",
                    })
                    continue
                # Check for MES410 fallback if provided
                mes_handled = False
                if mes410_fallback_df is not None and not mes410_fallback_df.empty:
                    m_col = "ShoeName_Std" if "ShoeName_Std" in mes410_fallback_df.columns else ("Model 2" if "Model 2" in mes410_fallback_df.columns else "Model")
                    sub_mes = (
                        mes410_fallback_df[mes410_fallback_df[m_col].apply(normalize_model_name) == model]
                        if fty == "Grand Total"
                        else mes410_fallback_df[(mes410_fallback_df["FTY"] == fty) & (mes410_fallback_df[m_col].apply(normalize_model_name) == model)]
                    )
                    # Check if MES410 has defect issue breakdown
                    defect_cols = mes410_defect_columns(sub_mes)
                    mes_def_sums = {}
                    for dc in defect_cols:
                        s_val = pd.to_numeric(sub_mes[dc], errors="coerce").fillna(0).sum()
                        if s_val > 0:
                            mes_def_sums[normalize_defect_name(dc)] = int(s_val)

                    if mes_def_sums:
                        sorted_mes = sorted(mes_def_sums.items(), key=lambda x: -x[1])
                        tot_mes_def = sum(mes_def_sums.values())
                        for rank, (d_name, d_qty) in enumerate(sorted_mes[:3], start=1):
                            share = d_qty / tot_mes_def if tot_mes_def > 0 else 0.0
                            spec_ratio = tot_mes_def / d_qty if d_qty > 0 else 0.0
                            p2_records.append({
                                "FTY": fty,
                                "Shoename": model,
                                "Defect Rank": f"Top {rank} (MES410 Thay thế)",
                                "Defect Type": d_name,
                                "Defect Quantity": int(d_qty),
                                "Total Shoe Defects": int(tot_mes_def),
                                "Defect Share %": share,
                                "Spec Ratio (Total/Defects)": spec_ratio,
                                "Color HEX": color_map.get(str(d_name).lower(), ""),
                                "Data Status": "MES410_Fallback",
                            })
                        mes_handled = True

                if not mes_handled:
                    p2_records.append({
                        "FTY": fty,
                        "Shoename": model,
                        "Defect Rank": "Thiếu dữ liệu (QAStation)",
                        "Defect Type": "N/A",
                        "Defect Quantity": 0,
                        "Total Shoe Defects": 0,
                        "Defect Share %": 0.0,
                        "Spec Ratio (Total/Defects)": 0.0,
                        "Color HEX": "",
                        "Data Status": "Missing_QAStation",
                    })
                continue

            dgrp = m_df.groupby("Issues_Std", sort=False)["IssueQty_Val"].sum().reset_index()
            dgrp.rename(columns={"Issues_Std": "Defect Type", "IssueQty_Val": "Defect Quantity"}, inplace=True)

            # Sort by quantity, preserving the source order for ties.
            dgrp = dgrp.sort_values(
                by="Defect Quantity",
                ascending=False,
                kind="stable"
            ).reset_index(drop=True)

            for rank, row in enumerate(dgrp.head(3).itertuples(), start=1):
                def_name = row._1
                def_qty = int(round(row._2))
                share = (def_qty / tot_defects) if tot_defects > 0 else 0.0
                spec_ratio = (tot_defects / def_qty) if def_qty > 0 else 0.0
                hex_color = color_map.get(str(def_name).lower(), "")

                p2_records.append({
                    "FTY": fty,
                    "Shoename": model,
                    "Defect Rank": f"Top {rank}",
                    "Defect Type": def_name,
                    "Defect Quantity": int(def_qty),
                    "Total Shoe Defects": int(tot_defects),
                    "Defect Share %": share,
                    "Spec Ratio (Total/Defects)": spec_ratio,
                    "Color HEX": hex_color,
                    "Data Status": "OK",
                })

    return pd.DataFrame(p2_records)


def compare_defect_shares(
    computed_share: float,
    reference_share: float,
    display_decimals: int = 4,
) -> bool:
    """
    Compare defect share rate at the given display precision.
    E.g. 0.164210 vs 0.1642 at 4 decimals (or 2 decimals in percent: 16.42%).
    """
    return round(float(computed_share), display_decimals) == round(float(reference_share), display_decimals)


def verify_hfpa_deliverables_consistency(
    output_dir: Any,
    target_month: Optional[str] = None,
    target_year: Optional[str] = None,
    *,
    include_pptx: bool = True,
) -> Dict[str, Any]:
    """
    Cross-checks invariants across Master Workbook, Updated Template Excel, and PowerPoint:
    1. hfpa.total_sample: Factory inspection samples are identical across all 3 outputs.
    2. hfpa.total_defects: Factory defective pairs are identical across all 3 outputs.
    3. hfpa.top5_ranks: Top 5 models and ranking (1-5) are identical across all 3 outputs.
    4. hfpa.top5_dr: Top 5 DR% values are consistent across all 3 outputs (< 1e-4 tolerance).
    5. hfpa.defect_shares: Top 3 defect shares match between Master, Excel and PPTX.
    """
    import os
    from pathlib import Path
    import openpyxl

    out_path = Path(output_dir)
    checks: List[Dict[str, Any]] = []

    def add_check(rule: str, passed: bool, message: str, **kwargs):
        checks.append({
            "rule": rule,
            "status": "passed" if passed else "warning",
            "message": message,
            **kwargs,
        })

    # Locate files
    def candidates(stem, extension):
        if target_month and target_year:
            stem = f"{stem}_{target_month}_{target_year}"
            paths = [out_path / f"{stem}{suffix}.{extension}" for suffix in ("", "_New")]
            return sorted((p for p in paths if p.is_file()), key=lambda p: p.stat().st_mtime_ns, reverse=True)
        return sorted(out_path.glob(f"{stem}*.{extension}"), reverse=True)

    master_cands = candidates("HFPA_FTT_Combined_Master", "xlsx")
    updated_cands = candidates("HFPA_Template_Updated", "xlsx")
    pptx_cands = candidates("HFPA_Performance_Report", "pptx")

    if not master_cands or not updated_cands or (include_pptx and not pptx_cands):
        add_check("hfpa.cross_check.files", False, "Thiếu một trong ba file kết quả chính (Master, Template Updated, PPTX) để đối chiếu trước khi xuất.")
        return {"status": "warning", "checks": checks}

    master_file = master_cands[0]
    updated_file = updated_cands[0]
    pptx_file = pptx_cands[0] if include_pptx else None

    # 1. Read Master Workbook
    try:
        wb_m = openpyxl.load_workbook(master_file, read_only=True, data_only=True)
    except Exception as e:
        add_check("hfpa.master.readable", False, f"Không thể đọc Master workbook: {e}")
        return {"status": "warning", "checks": checks}

    top5_sheet_name = "Top5_Models" if "Top5_Models" in wb_m.sheetnames else ("Top5_Shoes" if "Top5_Shoes" in wb_m.sheetnames else None)
    top3_sheet_name = "Top3_Defects" if "Top3_Defects" in wb_m.sheetnames else ("Pivot2_Top3_Defects" if "Pivot2_Top3_Defects" in wb_m.sheetnames else None)
    required_tables = {"Pivot1_Validation", "HFPA_Quality_Tracking_Adjusted", "Lot_Defect_Variance",
                       "HFPA_Quality_Tracking", "HFPA_Mes410"}
    add_check("hfpa.master.tables", bool(top5_sheet_name and top3_sheet_name and required_tables <= set(wb_m.sheetnames)),
              "Kiểm tra đủ bảng kiểm tra, Top 5 và Top 3 trong master HFPA.")
    pending_models = []
    if "Model_Adjustments" in wb_m.sheetnames:
        rows = iter(wb_m["Model_Adjustments"].values)
        header = list(next(rows, ()))
        if "Status" not in header:
            pending_models.append("Model_Adjustments thiếu cột Status.")
        else:
            index = header.index("Status")
            pending_models = [row for row in rows if row[index] != "MES410_Corrected"]
    add_check("hfpa.model_review", not pending_models,
              "Kiểm tra các nhóm HFPA đã đủ dữ liệu để hiệu chỉnh theo SOP.",
              mismatches=len(pending_models))

    master_top5: Dict[str, List[Dict[str, Any]]] = {f: [] for f in FACTORIES}
    master_samples: Dict[str, int] = {}
    master_defects: Dict[str, int] = {}
    master_top3: Dict[Tuple[str, str, str], float] = {}

    if top5_sheet_name:
        ws_t5 = wb_m[top5_sheet_name]
        rows = list(ws_t5.iter_rows(values_only=True))
        if rows:
            hdr = [str(c).strip() if c else "" for c in rows[0]]
            col_map = {name: idx for idx, name in enumerate(hdr)}
            for r in rows[1:]:
                if not any(r):
                    continue
                fty = str(r[col_map.get("FTY", 6)] or "").strip()
                if fty not in FACTORIES:
                    continue
                model = normalize_model_name(r[col_map.get("Shoename", 0)])
                insp = int(r[col_map.get("Sum of InspQty", 1)] or 0)
                val_fail = int(r[col_map.get("Sum of FailQty (Validated)", 3)] or 0)
                dr_val = float(r[col_map.get("DR%", 4)] or 0.0)
                rank = int(r[col_map.get("Rank", 5)] or len(master_top5[fty]) + 1)
                master_top5[fty].append({"model": model, "insp": insp, "def": val_fail, "dr": dr_val, "rank": rank})
                master_samples[fty] = insp

    if "Pivot1_Validation" in wb_m.sheetnames and "HFPA_Quality_Tracking_Adjusted" in wb_m.sheetnames:
        rows_p1 = wb_m["Pivot1_Validation"].values
        p1_header = list(next(rows_p1))
        f_col = p1_header.index("FTY")
        insp_col = p1_header.index("InspQty")
        fail_col = p1_header.index("FailQty")
        rate_col = p1_header.index("DR%_Validated")
        invalid_rates = []
        master_samples, master_defects = dict.fromkeys(FACTORIES, 0), dict.fromkeys(FACTORIES, 0)
        for index, row in enumerate(rows_p1, 2):
            if row[f_col] in master_defects:
                master_samples[row[f_col]] += int(row[insp_col] or 0)
                master_defects[row[f_col]] += int(row[fail_col] or 0)
                expected_rate = row[fail_col] / row[insp_col] if row[insp_col] else 0
                if not isinstance(row[rate_col], (int, float)) or not np.isfinite(row[rate_col]) or abs(row[rate_col] - expected_rate) > 1e-12:
                    invalid_rates.append(index)
        add_check("hfpa.pivot_rates", not invalid_rates, "Kiểm tra DR% của từng dòng sau hiệu chỉnh.",
                  invalid_rows=invalid_rates[:20])
    elif "Lot_Defect_Variance" in wb_m.sheetnames:
        ws_lot = wb_m["Lot_Defect_Variance"]
        rows_lot = list(ws_lot.iter_rows(values_only=True))
        if rows_lot:
            l_hdr = [str(c).strip() if c else "" for c in rows_lot[0]]
            f_col = l_hdr.index("FTY") if "FTY" in l_hdr else -1
            fail_col_idx = -1
            if "HFPA_Defective_Pairs" in l_hdr:
                fail_col_idx = l_hdr.index("HFPA_Defective_Pairs")
            elif "HFPA_Defect_Qty" in l_hdr:
                # Note: in legacy sheets HFPA_Defect_Qty held issues or pairs
                fail_col_idx = l_hdr.index("HFPA_Defect_Qty")
            if f_col != -1 and fail_col_idx != -1:
                lot_def_totals: Dict[str, int] = {f: 0 for f in FACTORIES}
                for r in rows_lot[1:]:
                    rf = str(r[f_col] or "").strip()
                    if rf in lot_def_totals:
                        try:
                            lot_def_totals[rf] += int(r[fail_col_idx] or 0)
                        except Exception:
                            pass
                master_defects = lot_def_totals

    if top3_sheet_name:
        ws_t3 = wb_m[top3_sheet_name]
        rows_t3 = list(ws_t3.iter_rows(values_only=True))
        if rows_t3:
            t3_hdr = [str(c).strip() if c else "" for c in rows_t3[0]]
            f_idx = t3_hdr.index("FTY") if "FTY" in t3_hdr else 0
            m_idx = t3_hdr.index("Shoename") if "Shoename" in t3_hdr else 1
            d_idx = t3_hdr.index("Defect Type") if "Defect Type" in t3_hdr else 3
            s_idx = t3_hdr.index("Defect Share %") if "Defect Share %" in t3_hdr else 6
            for r in rows_t3[1:]:
                if not any(r):
                    continue
                r_fty = str(r[f_idx] or "").strip()
                r_mod = normalize_model_name(r[m_idx])
                r_def = normalize_defect_name(r[d_idx])
                try:
                    r_share = float(r[s_idx] or 0.0)
                    if r_fty in FACTORIES and r_share > 0:
                        master_top3[(r_fty, r_mod, r_def)] = r_share
                except Exception:
                    pass
    wb_m.close()

    # 2. Read Updated Excel Template
    wb_u = openpyxl.load_workbook(updated_file, read_only=True, data_only=True)
    if "HFPA" not in wb_u.sheetnames:
        add_check("hfpa.template_updated.sheet", False, f"Sheet 'HFPA' không có trong {os.path.basename(updated_file)}")
        wb_u.close()
        return {"status": "warning", "checks": checks}

    ws_u = wb_u["HFPA"]
    month_cols: Dict[str, int] = {}
    for c in range(3, 16):
        v = ws_u.cell(1, c).value
        if v:
            month_cols[str(v).strip()] = c

    if target_month and target_month in month_cols:
        m_col = month_cols[target_month]
    elif target_month:
        wb_u.close()
        add_check("hfpa.month", False, f"Excel tổng hợp thiếu cột tháng {target_month}.")
        return {"status": "warning", "checks": checks}
    else:
        # Pick latest non-empty month
        m_col = None
        for m_name in ["Dec", "Nov", "Oct", "Sep", "Aug", "Jul", "Jun", "May", "Apr", "Mar", "Feb", "Jan"]:
            if m_name in month_cols:
                c = month_cols[m_name]
                if any((ws_u.cell(r, c).value or 0) > 0 for r in [2, 3, 4, 5]):
                    m_col = c
                    break
        if m_col is None:
            wb_u.close()
            add_check("hfpa.month", False, "Cần chỉ định tháng khi Excel tổng hợp chưa có số đôi kiểm.")
            return {"status": "warning", "checks": checks}

    u_samples = {
        "VH": int(ws_u.cell(2, m_col).value or 0),
        "VH2": int(ws_u.cell(3, m_col).value or 0),
        "JV": int(ws_u.cell(4, m_col).value or 0),
        "JV2": int(ws_u.cell(5, m_col).value or 0),
    }
    u_defects = {
        "VH": int(ws_u.cell(9, m_col).value or 0),
        "VH2": int(ws_u.cell(10, m_col).value or 0),
        "JV": int(ws_u.cell(11, m_col).value or 0),
        "JV2": int(ws_u.cell(12, m_col).value or 0),
    }

    site_rows = {"VH": (22, 26), "VH2": (28, 32), "JV": (34, 38), "JV2": (40, 44)}
    site_hdr_rows = {"VH": 21, "VH2": 27, "JV": 33, "JV2": 39}
    u_top5: Dict[str, List[Dict[str, Any]]] = {f: [] for f in FACTORIES}
    u_top3: Dict[Tuple[str, str, str], float] = {}

    for fty, (s_row, e_row) in site_rows.items():
        hdr_row = site_hdr_rows[fty]
        def_cols: Dict[str, int] = {}
        for c in range(13, ws_u.max_column + 1):
            val = ws_u.cell(hdr_row, c).value
            if not val or not str(val).strip() or str(val).strip().startswith("Site"):
                break
            def_cols[normalize_defect_name(str(val).strip())] = c

        for r in range(s_row, e_row + 1):
            m_val = ws_u.cell(r, 4).value
            if not m_val:
                continue
            m_name = normalize_model_name(m_val)
            insp = int(ws_u.cell(r, 5).value or 0)
            fail = int(ws_u.cell(r, 6).value or 0)
            dr_cell = ws_u.cell(r, 7).value
            dr_val = float(dr_cell) if isinstance(dr_cell, (int, float)) else ((fail / insp) if insp > 0 else 0.0)
            u_top5[fty].append({"model": m_name, "insp": insp, "def": fail, "dr": dr_val})

            for d_name, c_idx in def_cols.items():
                s_val = ws_u.cell(r, c_idx).value
                if s_val is not None and isinstance(s_val, (int, float)) and s_val > 0:
                    u_top3[(fty, m_name, d_name)] = float(s_val)

    wb_u.close()
    formula_book = openpyxl.load_workbook(updated_file, data_only=False)
    formula_sheet = formula_book["HFPA"]
    layout_issues = hfpa_defect_layout_issues(formula_sheet)
    add_check("hfpa.excel.layout", not layout_issues,
              "Kiểm tra toàn bộ lỗi nằm trong khung bảng, tên lỗi xuống dòng và model khớp Top 5.",
              issues=layout_issues)
    invalid_formulas = []
    for start, end in site_rows.values():
        for row in range(start, end + 1):
            if not formula_sheet.cell(row, 4).value:
                continue
            value = formula_sheet.cell(row, 7).value
            guarded = f'=IF(E{row}>0,F{row}/E{row},"")'
            allowed = {guarded}
            if formula_sheet.cell(row, 5).value:
                allowed.add(f"=F{row}/E{row}")
            if not isinstance(value, str) or value.replace(" ", "") not in allowed:
                invalid_formulas.append(f"G{row}")
    formula_book.close()
    add_check("hfpa.excel.formulas", not invalid_formulas, "Kiểm tra công thức DR% Top 5 và mẫu số bằng 0.",
              invalid_cells=invalid_formulas)

    p_samples, p_defects = {}, {}
    p_top5 = {factory: [] for factory in FACTORIES}
    p_top3 = {}
    if include_pptx:
        # 3. Read PowerPoint presentation
        try:
            import pptx
            from pptx.oxml.ns import qn
            prs = pptx.Presentation(pptx_file)
            slide = prs.slides[0]
            chart_top5_map = {"VH": "Chart 194", "VH2": "Chart 195", "JV": "Chart 196", "JV2": "Chart 197"}
            chart_analysis_map = {"VH": "Chart 51", "VH2": "Chart 52", "JV": "Chart 53", "JV2": "Chart 54"}
            chart_defects_map = {"VH": "Chart 63", "VH2": "Chart 64", "JV": "Chart 65", "JV2": "Chart 66"}

            p_samples: Dict[str, int] = {}
            p_defects: Dict[str, int] = {}
            p_top5: Dict[str, List[Dict[str, Any]]] = {f: [] for f in FACTORIES}
            p_top3: Dict[Tuple[str, str, str], float] = {}

            shape_dict = {s.name: s for s in slide.shapes if s.has_chart}
            missing_charts = [name for group in (chart_analysis_map, chart_top5_map, chart_defects_map)
                              for name in group.values() if name not in shape_dict]
            add_check("hfpa.pptx.charts", not missing_charts, "Kiểm tra đủ biểu đồ HFPA cho từng nhà máy.",
                      missing_charts=missing_charts)
            bad_links = []
            for name in (*chart_analysis_map.values(), *chart_top5_map.values(), *chart_defects_map.values()):
                if name not in shape_dict:
                    continue
                chart = shape_dict[name].chart
                external = chart._chartSpace.find(qn("c:externalData"))
                rel_id = external.get(qn("r:id")) if external is not None else None
                rel = chart.part.rels[rel_id] if rel_id in chart.part.rels else None
                if (rel is None or not rel.is_external or rel.target_ref != updated_file.name
                        or any(not (f.text or "").startswith("'HFPA'!") for f in chart._chartSpace.iter(qn("c:f")))):
                    bad_links.append(name)
            add_check("hfpa.pptx.links", not bad_links, "Kiểm tra liên kết biểu đồ PowerPoint đến đúng Excel và sheet HFPA.",
                      invalid_links=bad_links)

            for fty, cname in chart_analysis_map.items():
                if cname in shape_dict:
                    cs = shape_dict[cname].chart._chartSpace
                    bar_series = cs.findall('.//' + qn('c:barChart') + '/' + qn('c:ser'))
                    if len(bar_series) >= 2:
                        p_vals = [float(pt.text) for pt in bar_series[0].findall('.//' + qn('c:val') + '//' + qn('c:v'))]
                        d_vals = [float(pt.text) for pt in bar_series[1].findall('.//' + qn('c:val') + '//' + qn('c:v'))]
                        if p_vals:
                            p_samples[fty] = int(round(p_vals[-1]))
                        if d_vals:
                            p_defects[fty] = int(round(d_vals[-1]))

            for fty, cname in chart_top5_map.items():
                if cname in shape_dict:
                    cs = shape_dict[cname].chart._chartSpace
                    cats = [normalize_model_name(pt.text) for pt in cs.findall('.//' + qn('c:cat') + '//' + qn('c:v'))]
                    vals = [float(pt.text) for pt in cs.findall('.//' + qn('c:val') + '//' + qn('c:v'))]
                    for idx, mod in enumerate(cats):
                        val = vals[idx] if idx < len(vals) else 0.0
                        p_top5[fty].append({"model": mod, "dr": val})

            for fty, cname in chart_defects_map.items():
                if cname in shape_dict:
                    cs = shape_dict[cname].chart._chartSpace
                    ser_elems = cs.findall('.//' + qn('c:barChart') + '/' + qn('c:ser'))
                    for ser in ser_elems:
                        tx = ser.find('.//' + qn('c:tx') + '//' + qn('c:v'))
                        sname = normalize_defect_name(tx.text) if tx is not None else ""
                        models = {
                            int(pt.get('idx', 0)): normalize_model_name(pt.findtext(qn('c:v')))
                            for pt in ser.findall('.//' + qn('c:cat') + '//' + qn('c:pt'))
                        }
                        val_pts = ser.findall('.//' + qn('c:val') + '//' + qn('c:pt'))
                        for pt in val_pts:
                            m_idx = int(pt.get('idx', 0))
                            v_elem = pt.find(qn('c:v'))
                            if v_elem is not None and m_idx in models:
                                m_name = models[m_idx]
                                fval = float(v_elem.text)
                                if fval > 0:
                                    p_top3[(fty, m_name, sname)] = fval
        except Exception as e:
            add_check("hfpa.pptx.readable", False, f"Không thể đọc PowerPoint presentation: {e}")
            return {"status": "warning", "checks": checks}

    # 4. CROSS-CHECK RULES
    # RULE 1: Total Audited Sample Consistency
    tot_samp_match = master_samples == u_samples and (not include_pptx or u_samples == p_samples)
    add_check(
        "hfpa.total_sample",
        tot_samp_match,
        "Đối chiếu tổng số đôi kiểm tra giữa Master, Excel tổng hợp và PowerPoint.",
        expected=str(u_samples),
        actual=f"Master={master_samples}; PowerPoint={p_samples}",
    )

    # RULE 2: Total Defective Pairs Consistency
    tot_def_match = master_defects == u_defects and (not include_pptx or u_defects == p_defects)
    add_check(
        "hfpa.total_defects",
        tot_def_match,
        "Đối chiếu tổng số đôi lỗi giữa Master, Excel tổng hợp và PowerPoint.",
        expected=str(u_defects),
        actual=f"Master={master_defects}; PowerPoint={p_defects}",
    )

    # RULE 3: Top 5 Ranks & Model Names Consistency
    ranks_matched = True
    mismatched_ranks: List[str] = []
    for fty in FACTORIES:
        m_models = [x["model"] for x in master_top5.get(fty, [])]
        u_models = [x["model"] for x in u_top5.get(fty, [])]
        p_models = [x["model"] for x in p_top5.get(fty, [])]
        master_ranks = [x["rank"] for x in master_top5[fty]]
        if not (m_models == u_models and (not include_pptx or u_models == p_models) and master_ranks == list(range(1, len(m_models) + 1))):
            ranks_matched = False
            mismatched_ranks.append(f"{fty} (Master: {m_models}, Excel: {u_models}, PPTX: {p_models})")
    add_check(
        "hfpa.top5_ranks",
        ranks_matched,
        "Đối chiếu thứ hạng và danh sách Top 5 Model giữa Master, Excel tổng hợp và PowerPoint.",
        mismatched_factories=mismatched_ranks,
    )

    # RULE 4: Top 5 DR% Consistency
    dr_matched = True
    dr_discrepancies: List[str] = []
    for fty in FACTORIES:
        counts = [len(master_top5[fty]), len(u_top5[fty])] + ([len(p_top5[fty])] if include_pptx else [])
        if len(set(counts)) != 1:
            dr_matched = False
            dr_discrepancies.append(f"{fty}: thiếu hoặc thừa dòng Top 5.")
            continue
        for idx, (master, excel) in enumerate(zip(master_top5[fty], u_top5[fty])):
            m_dr, u_dr = master["dr"], excel["dr"]
            p_dr = p_top5[fty][idx]["dr"] if include_pptx else u_dr
            if ((master["insp"], master["def"]) != (excel["insp"], excel["def"])
                    or not compare_defect_shares(m_dr, u_dr) or not compare_defect_shares(u_dr, p_dr)):
                dr_matched = False
                dr_discrepancies.append(f"{fty} #{idx+1} ({master_top5[fty][idx]['model']}): M={m_dr:.4f}, E={u_dr:.4f}, P={p_dr:.4f}")
    add_check(
        "hfpa.top5_dr",
        dr_matched,
        "Đối chiếu tỷ lệ lỗi DR% của Top 5 Models giữa Master, Excel tổng hợp và PowerPoint.",
        discrepancies=dr_discrepancies,
    )

    # RULE 5: Defect Shares Consistency
    def_share_discrepancies: List[str] = []
    comparisons = [("Excel", u_top3)] + ([("PPTX", p_top3)] if include_pptx else [])
    for label, other in comparisons:
        for key in sorted(master_top3.keys() - other.keys()):
            def_share_discrepancies.append(f"{key}: thiếu trong {label}")
        for key in sorted(other.keys() - master_top3.keys()):
            def_share_discrepancies.append(f"{key}: thừa trong {label}")
        for key in sorted(master_top3.keys() & other.keys()):
            if not compare_defect_shares(master_top3[key], other[key]):
                def_share_discrepancies.append(f"{key}: Master={master_top3[key]:.4f}, {label}={other[key]:.4f}")
    add_check(
        "hfpa.defect_shares",
        not def_share_discrepancies,
        "Đối chiếu tỷ trọng lỗi Top 3 giữa Master, Excel tổng hợp và PowerPoint.",
        discrepancies=def_share_discrepancies,
    )

    warnings = sum(c["status"] == "warning" for c in checks)
    return {
        "status": "warning" if warnings else "passed",
        "checks": checks,
        "warning_count": warnings,
    }
