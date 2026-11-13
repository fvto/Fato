"""
Agent 2: ReconciliationEngineAgent
High-Performance In-Memory Data Pipeline, Smart Lot Matching, and Fast Excel Serialization.

Addresses Inefficiency 2:
- Eliminates 3x redundant disk deserialization/serialization cycles.
- Introduces dual-pass smart lot reconciliation (resolves Plant prefix alias mismatches like N1 vs N2).
- Vectorizes Hamilton integer defect redistribution for exact MES410 parity.
- Optimizes heavy 25MB Excel exports with batched styling and memory management.
"""

from __future__ import annotations

import glob
import os
import re
import time
import sys
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

MONTH_MAP = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}
FACTORIES = ["VH", "VH2", "JV", "JV2"]
LOT_KEYS = ["Date_Std", "FTY", "Plant_Std", "Line_Std"]
QA_INSPECTION_KEYS = LOT_KEYS + ["Station_Raw", "ShoeName_Std"]
MODEL_KEYS = LOT_KEYS + ["ShoeName_Std"]
MES_ISSUE_TOTAL = "Defect Q'ty (Total Defect Q'ty)"


def quality_tracking_pairs(rows, keys, quantity_columns):
    """Match the SOP pivot grain, including distinct inspection and failed-pair values."""
    return rows.drop_duplicates(keys + quantity_columns).groupby(keys)[quantity_columns].sum()


def inspection_dates(values: pd.Series) -> pd.Series:
    # Parse year-first export dates separately from the day-first display format.
    iso = values.astype("string").str.match(r"^\d{4}-", na=False)
    dates = pd.to_datetime(values.where(~iso), format="mixed", dayfirst=True, errors="coerce")
    return dates.fillna(pd.to_datetime(values.where(iso), format="mixed", errors="coerce"))

try:
    from ..hfpa_common import (
        normalize_model_name,
        normalize_defect_name,
        compute_model_ranks_mes410,
        compute_top5_models,
        compute_top3_defects,
        mes410_defect_columns,
        FACTORY_ALIASES,
    )
except (ImportError, ValueError):
    try:
        from hfpa_common import (
            normalize_model_name,
            normalize_defect_name,
            compute_model_ranks_mes410,
            compute_top5_models,
            compute_top3_defects,
            mes410_defect_columns,
            FACTORY_ALIASES,
        )
    except (ImportError, ValueError):
        import sys
        sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from hfpa_common import (
            normalize_model_name,
            normalize_defect_name,
            compute_model_ranks_mes410,
            compute_top5_models,
            compute_top3_defects,
            mes410_defect_columns,
            FACTORY_ALIASES,
        )


def summarize_mes410_pairs(hfpa_df: pd.DataFrame) -> Tuple[int, int, pd.Series]:
    """Return audited pairs, failed pairs and model rankings for the selected MES410 rows."""
    tot_samples, tot_failures, ranked_df = compute_model_ranks_mes410(hfpa_df)
    s = pd.Series(ranked_df["Failures"].values, index=ranked_df["Model"].values)
    return tot_samples, tot_failures, s


@dataclass
class DataBus:
    """In-memory communication bus holding all pipeline datasets."""
    active_month: str
    active_year: str
    qa_raw: pd.DataFrame = field(default_factory=pd.DataFrame)
    hfpa_raw: pd.DataFrame = field(default_factory=pd.DataFrame)
    qa_validated: pd.DataFrame = field(default_factory=pd.DataFrame)
    hfpa_clean: pd.DataFrame = field(default_factory=pd.DataFrame)
    lot_comparison: pd.DataFrame = field(default_factory=pd.DataFrame)
    pivot1_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    all_top5_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    top5_dict: Dict[str, pd.DataFrame] = field(default_factory=dict)
    pivot2_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    audit_records: List[Dict[str, Any]] = field(default_factory=list)
    color_map: Dict[str, str] = field(default_factory=dict)
    elapsed_times: Dict[str, float] = field(default_factory=dict)
    output_excel_path: Optional[str] = None


class ReconciliationEngineAgent:
    """
    Proactive agent executing in-memory reconciliation, Hamilton integer scaling,
    and optimized data exports.
    """

    def __init__(self, base_dir: Optional[str] = None):
        self.base_dir = base_dir or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.input_ftt_dir = os.path.join(self.base_dir, "Input", "FTT")
        self.input_hfpa_dir = os.path.join(self.base_dir, "Input", "HFPA")
        self.output_dir = os.path.join(self.base_dir, "Output")
        self.database_dir = os.path.join(self.base_dir, "Database")
        self.color_template_path = os.path.join(self.base_dir, "Color_template.xlsx")
        self.logs_dir = os.path.join(self.base_dir, "logs")
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.logs_dir, exist_ok=True)

    @staticmethod
    def parse_plant_and_line(val: Any) -> Tuple[str, str]:
        """Parse Plant and Line from strings like 'N1 Line 1', 'Line 1', 'A10 Line 1'."""
        if pd.isna(val):
            return "", ""
        s = str(val).strip()
        m = re.match(r"^([A-Za-z0-9]+)\s*Line\s*([0-9]+)", s, re.IGNORECASE)
        if m:
            return m.group(1).upper(), m.group(2)
        m2 = re.match(r"^Line\s*([0-9]+)", s, re.IGNORECASE)
        if m2:
            return "", m2.group(1)
        cleaned = re.sub(r"[^0-9]", "", s)
        return "", cleaned if cleaned else s

    @staticmethod
    def extract_factory(filename: str) -> str:
        base = os.path.splitext(os.path.basename(filename))[0]
        for fty in ["JV2", "VH2", "JV", "VH"]:
            if re.search(rf"[-_\s]{fty}$", base, re.IGNORECASE) or re.search(rf"[-_\s]{fty}[-_\s]", base, re.IGNORECASE):
                return fty
        for fty in ["JV2", "VH2", "JV", "VH"]:
            if fty in base.upper():
                return fty
        return "UNKNOWN"

    def load_color_template(self) -> Dict[str, str]:
        """Load defect color mappings from Color_template.xlsx."""
        try:
            from .color_manager import ColorManager
        except (ImportError, ValueError):
            from agents.color_manager import ColorManager
        return ColorManager(self.color_template_path).load_color_map()


    def load_raw_datasets(
        self, target_month: str, target_year: str, *, input_files=None
    ) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Load FTT/QAStation and HFPA/MES410 raw files into unified dataframes."""
        t0 = time.time()
        print(f"\n[AGENT 2] Loading raw datasets for {target_month} {target_year}...")

        def assign_factory(frame, factory, filename):
            for field in ("FTY", "Factory Group Cd_UccId"):
                if field in frame:
                    normalized = frame[field].astype("string").str.strip().str.upper().map(FACTORY_ALIASES)
                    if not normalized.eq(factory).fillna(False).all():
                        raise ValueError(f"{filename}: mã nhà máy trong dữ liệu khác tên file.")
            frame["FTY"] = factory

        # 1. Load FTT
        ftt_files = (input_files["FTT"] if input_files is not None else
                     glob.glob(os.path.join(self.input_ftt_dir, "*.xlsx")) + glob.glob(os.path.join(self.input_ftt_dir, "*.xls")))
        ftt_dfs = []
        for fp in sorted(ftt_files):
            fname = os.path.basename(fp)
            if fname.startswith("~$"):
                continue
            fty = self.extract_factory(fname)
            df = pd.read_excel(fp)
            assign_factory(df, fty, fname)
            df["Source_File"] = fname

            if "InspDate" in df.columns:
                df["Date_Raw"] = df["InspDate"]
                df["Model_Raw"] = df["ShoeName"] if "ShoeName" in df.columns else df.get("Shoename", "")
                df["InspQty_Raw"] = pd.to_numeric(df["InspQty"], errors="coerce").fillna(0)
                df["FailQty_Raw"] = pd.to_numeric(df["FailQty"], errors="coerce").fillna(0)
                df["Issues_Raw"] = df["Issues"] if "Issues" in df.columns else ""
                df["IssueQty_Raw"] = pd.to_numeric(df["IssueQty"], errors="coerce").fillna(0)
                df["Station_Raw"] = df["Station"] if "Station" in df.columns else "HFPA"
                plants, lines = zip(*df["Line"].apply(self.parse_plant_and_line)) if len(df) else ((), ())
                df["Plant_Parsed"] = plants
                df["Line_Parsed"] = lines
            else:
                df["Date_Raw"] = df["Date"]
                df["Model_Raw"] = df["Model"] if "Model" in df.columns else df.get("ShoeName", "")
                df["InspQty_Raw"] = pd.to_numeric(df["Total Pair Produced"], errors="coerce").fillna(0)
                df["FailQty_Raw"] = pd.to_numeric(df["Issues Q'ty"], errors="coerce").fillna(0)
                df["Issues_Raw"] = df["Defect Issues"] if "Defect Issues" in df.columns else ""
                df["IssueQty_Raw"] = pd.to_numeric(df["Issues Q'ty"], errors="coerce").fillna(0)
                df["Station_Raw"] = df["Station"] if "Station" in df.columns else "Assembly"
                df["Plant_Parsed"] = df["Plant"].astype(str).str.strip().str.upper() if "Plant" in df.columns else ""
                _, lines = zip(*df["Line"].apply(self.parse_plant_and_line)) if len(df) else ((), ())
                df["Line_Parsed"] = lines

            ftt_dfs.append(df)

        combined_qa = pd.concat(ftt_dfs, ignore_index=True) if ftt_dfs else pd.DataFrame()

        # 2. Load HFPA
        hfpa_files = (input_files["HFPA"] if input_files is not None else
                      glob.glob(os.path.join(self.input_hfpa_dir, "*.xlsx")) + glob.glob(os.path.join(self.input_hfpa_dir, "*.xls")))
        hfpa_dfs = []
        for fp in sorted(hfpa_files):
            fname = os.path.basename(fp)
            if fname.startswith("~$"):
                continue
            fty = self.extract_factory(fname)
            df = pd.read_excel(fp)
            assign_factory(df, fty, fname)
            df["Source_File"] = fname
            hfpa_dfs.append(df)

        combined_hfpa = pd.concat(hfpa_dfs, ignore_index=True) if hfpa_dfs else pd.DataFrame()
        for files in (ftt_files, hfpa_files):
            factories = [self.extract_factory(path) for path in files if not os.path.basename(path).startswith("~$")]
            if sorted(factories) != sorted(FACTORIES):
                raise ValueError("HFPA cần đúng một file cho mỗi nhà máy VH, VH2, JV và JV2 trong từng nhóm nguồn.")
        period = f"{target_year}-{next(i for i, label in MONTH_MAP.items() if label == target_month):02d}"
        for rows, date_column in (
            (combined_qa, "Date_Raw"),
            (combined_hfpa, "Audit Date (dd-mmm-yyyy)" if "Audit Date (dd-mmm-yyyy)" in combined_hfpa else "Audit Date"),
        ):
            dates = inspection_dates(rows[date_column])
            outside_period = dates.dt.strftime("%Y-%m").ne(period)
            if outside_period.any():
                name = rows.loc[outside_period, "Source_File"].iloc[0]
                raise ValueError(f"{name}: có ngày sai hoặc nằm ngoài kỳ HFPA {period}.")
        print(f"   [+] Loaded QAStation rows: {len(combined_qa):,} | MES410 rows: {len(combined_hfpa):,} ({time.time()-t0:.2f}s)")
        return combined_qa, combined_hfpa

    @staticmethod
    def distribute_integer_defects(quantities: List[int], target_total: int) -> List[int]:
        """
        Largest Remainder Method (Hamilton/Hare-Niemeyer).
        Guarantees exact integer sum == target_total with zero roundoff drift.
        """
        total_q = sum(quantities)
        if total_q == 0 or target_total == 0:
            return [0] * len(quantities)
        quotas = [q * target_total / total_q for q in quantities]
        base = [int(x) for x in quotas]
        remainder = int(target_total - sum(base))
        if remainder > 0:
            fracs = [(quotas[i] - base[i], quantities[i], i) for i in range(len(quantities))]
            fracs.sort(key=lambda x: (x[0], x[1]), reverse=True)
            for i in range(remainder):
                idx = fracs[i % len(fracs)][2]
                base[idx] += 1
        elif remainder < 0:
            fracs = [(quotas[i] - base[i], quantities[i], i) for i in range(len(quantities)) if base[i] > 0]
            fracs.sort(key=lambda x: (x[0], -x[1]))
            for i in range(abs(remainder)):
                idx = fracs[i % len(fracs)][2]
                base[idx] -= 1
        return base

    @classmethod
    def reconcile_and_validate(
        cls, qa_df: pd.DataFrame, hfpa_df: pd.DataFrame
    ) -> Tuple[pd.DataFrame, pd.DataFrame, List[Dict[str, Any]]]:
        """
        Dual-Pass Smart Lot Reconciliation:
        Pass 1: Exact match on (Date, FTY, Plant, Line)
        Pass 2: Smart alias match on (Date, FTY, Line) when quantities match or plant alias exists.
        """
        t0 = time.time()
        print("   [*] Executing Dual-Pass Smart Lot Reconciliation...")

        # Standardize fields
        if "Date_Raw" not in qa_df.columns:
            qa_df["Date_Raw"] = qa_df["InspDate"] if "InspDate" in qa_df.columns else qa_df.get("Date", "")
        if "Plant_Parsed" not in qa_df.columns or "Line_Parsed" not in qa_df.columns:
            line_src = qa_df["Line"] if "Line" in qa_df.columns else qa_df.get("Fty/Line", "")
            plants, lines = zip(*line_src.map(cls.parse_plant_and_line)) if len(qa_df) else ((), ())
            qa_df["Plant_Parsed"] = plants
            qa_df["Line_Parsed"] = lines
        if "Model_Raw" not in qa_df.columns:
            qa_df["Model_Raw"] = qa_df["ShoeName"] if "ShoeName" in qa_df.columns else qa_df.get("Model", "")
        if "Issues_Raw" not in qa_df.columns:
            qa_df["Issues_Raw"] = qa_df["Issues"] if "Issues" in qa_df.columns else qa_df.get("Defect Issues", "")
        if "InspQty_Raw" not in qa_df.columns:
            qa_df["InspQty_Raw"] = qa_df["InspQty"] if "InspQty" in qa_df.columns else qa_df.get("Total Pair Produced", 0)
        if "FailQty_Raw" not in qa_df.columns:
            qa_df["FailQty_Raw"] = qa_df["FailQty"] if "FailQty" in qa_df.columns else qa_df.get("Issues Q'ty", 0)
        if "IssueQty_Raw" not in qa_df.columns:
            qa_df["IssueQty_Raw"] = qa_df["IssueQty"] if "IssueQty" in qa_df.columns else qa_df.get("Issues Q'ty", 0)
        if "Station_Raw" not in qa_df.columns:
            qa_df["Station_Raw"] = qa_df["Station"] if "Station" in qa_df.columns else "HFPA"

        qa_df["Date_Std"] = inspection_dates(qa_df["Date_Raw"]).dt.strftime("%Y-%m-%d")
        qa_df["Plant_Std"] = qa_df["Plant_Parsed"].astype(str).str.strip().str.upper()
        qa_df["Line_Std"] = qa_df["Line_Parsed"].astype(str).str.strip()
        qa_df["ShoeName_Std"] = qa_df["Model_Raw"].apply(normalize_model_name)
        qa_df["Issues_Std"] = qa_df["Issues_Raw"].apply(normalize_defect_name)
        for name in ("InspQty_Raw", "FailQty_Raw", "IssueQty_Raw"):
            values = pd.to_numeric(qa_df[name], errors="coerce")
            if not (values.notna() & np.isfinite(values) & values.ge(0) & values.mod(1).eq(0)).all():
                raise ValueError(f"Quality Tracking: {name} cần là số nguyên không âm.")
            qa_df[name] = values
        if ((qa_df["IssueQty_Raw"] > 0) & (qa_df["Issues_Std"] == "")).any():
            raise ValueError("Quality Tracking có lượt lỗi dương nhưng thiếu tên loại lỗi.")
        if qa_df["FailQty_Raw"].gt(qa_df["InspQty_Raw"]).any():
            raise ValueError("Quality Tracking có số đôi lỗi lớn hơn số đôi kiểm.")

        hfpa_date_col = "Audit Date (dd-mmm-yyyy)" if "Audit Date (dd-mmm-yyyy)" in hfpa_df.columns else "Audit Date"
        hfpa_df["Date_Std"] = inspection_dates(hfpa_df[hfpa_date_col]).dt.strftime("%Y-%m-%d")
        hfpa_df["Plant_Std"] = hfpa_df["Plant"].astype(str).str.strip().str.upper()
        hfpa_df["Line_Std"] = hfpa_df["Line"].map(
            lambda line: str(cls.parse_plant_and_line(line)[1]).strip()
        )
        model_col = "Model 2" if "Model 2" in hfpa_df.columns else "Model"
        hfpa_df["Model_Std"] = hfpa_df[model_col].apply(normalize_model_name)

        defect_tot_col = MES_ISSUE_TOTAL if MES_ISSUE_TOTAL in hfpa_df.columns else "Defective Q'ty(pair)"
        sample_col = "Audit Sample Size" if "Audit Sample Size" in hfpa_df.columns else "Audit Sample"
        fail_pair_col = next(
            (col for col in hfpa_df.columns if "Defective" in col and "pair" in col),
            "Defective Q'ty(pair)",
        )
        for name in dict.fromkeys((sample_col, fail_pair_col, defect_tot_col)):
            values = pd.to_numeric(hfpa_df[name], errors="coerce")
            if not (values.notna() & np.isfinite(values) & values.ge(0) & values.mod(1).eq(0)).all():
                raise ValueError(f"MES410: {name} cần là số nguyên không âm.")
            hfpa_df[name] = values
        if hfpa_df[fail_pair_col].gt(hfpa_df[sample_col]).any():
            raise ValueError("MES410 có số đôi lỗi lớn hơn số đôi kiểm.")

        # Lot-level aggregation
        hfpa_lot = hfpa_df.groupby(LOT_KEYS).agg(
            HFPA_Audit_Sample=(sample_col, "sum"),
            HFPA_Defective_Pairs=(fail_pair_col, "sum"),
            HFPA_Defect_Qty=(defect_tot_col, "sum")
        ).reset_index()

        # Each issue row repeats the model's inspection and failed-pair counts.
        qa_inspections = quality_tracking_pairs(
            qa_df, QA_INSPECTION_KEYS, ["InspQty_Raw", "FailQty_Raw"]
        ).rename(columns={"InspQty_Raw": "QA_InspQty", "FailQty_Raw": "QA_FailQty"})
        qa_inspections["QA_IssueQty_Sum"] = qa_df.groupby(QA_INSPECTION_KEYS)["IssueQty_Raw"].sum()
        qa_lot = qa_inspections.groupby(level=LOT_KEYS).agg(
            QA_InspQty=("QA_InspQty", "sum"),
            QA_FailQty=("QA_FailQty", "sum"),
            QA_IssueQty_Sum=("QA_IssueQty_Sum", "sum")
        ).reset_index()

        # Pass 1: Exact 4-tuple merge
        lot_comp = pd.merge(
            qa_lot,
            hfpa_lot,
            on=LOT_KEYS,
            how="outer"
        )
        lot_comp["MES_Plant_Std"] = lot_comp["Plant_Std"].where(lot_comp["HFPA_Audit_Sample"].notna())

        # Pass 2: Smart Plant Alias Resolver for unmatched lots
        unmatched_qa = lot_comp[lot_comp["HFPA_Audit_Sample"].isna()].copy()
        unmatched_hfpa = lot_comp[lot_comp["QA_InspQty"].isna()].copy()

        if not unmatched_qa.empty and not unmatched_hfpa.empty:
            # Only resolve aliases with one candidate on each side.
            alias_key = ["Date_Std", "FTY", "Line_Std"]
            cand_merge = pd.merge(
                unmatched_qa[alias_key + ["QA_InspQty"]].rename_axis("QA_Row").reset_index(),
                unmatched_hfpa[alias_key + ["HFPA_Audit_Sample"]].rename_axis("HFPA_Row").reset_index(),
                on=alias_key,
            )
            valid_alias = cand_merge[cand_merge["QA_InspQty"] == cand_merge["HFPA_Audit_Sample"]]
            unique_alias = valid_alias[
                ~valid_alias["QA_Row"].duplicated(keep=False)
                & ~valid_alias["HFPA_Row"].duplicated(keep=False)
            ]
            if not unique_alias.empty:
                hfpa_columns = ["HFPA_Audit_Sample", "HFPA_Defective_Pairs", "HFPA_Defect_Qty"]
                lot_comp.loc[unique_alias["QA_Row"], hfpa_columns] = lot_comp.loc[
                    unique_alias["HFPA_Row"], hfpa_columns
                ].to_numpy()
                lot_comp.loc[unique_alias["QA_Row"], "MES_Plant_Std"] = lot_comp.loc[
                    unique_alias["HFPA_Row"], "Plant_Std"
                ].to_numpy()
                lot_comp = lot_comp.drop(index=unique_alias["HFPA_Row"])
                print(f"   [+] Dual-Pass Smart Matcher resolved {len(unique_alias)} lot plant-alias discrepancies!")

        # Separate variances by exact unit: Insp pairs, Defective pairs, Issue occurrences
        lot_comp["Insp_Variance"] = lot_comp["QA_InspQty"].fillna(0) - lot_comp["HFPA_Audit_Sample"].fillna(0)
        lot_comp["Fail_Pairs_Variance"] = lot_comp["QA_FailQty"].fillna(0) - lot_comp["HFPA_Defective_Pairs"].fillna(0)
        lot_comp["Issue_Variance"] = lot_comp["QA_IssueQty_Sum"].fillna(0) - lot_comp["HFPA_Defect_Qty"].fillna(0)
        lot_comp["Defect_Variance"] = lot_comp["Issue_Variance"]

        def assign_status(row):
            qa_missing = pd.isna(row.get("QA_InspQty")) and pd.isna(row.get("QA_FailQty")) and pd.isna(row.get("QA_IssueQty_Sum"))
            hfpa_missing = pd.isna(row.get("HFPA_Audit_Sample")) and pd.isna(row.get("HFPA_Defective_Pairs")) and pd.isna(row.get("HFPA_Defect_Qty"))
            if qa_missing:
                return "HFPA_Only"
            elif hfpa_missing:
                return "QAStation_Only"
            # Strict multi-metric consistency: must match inspection pairs, defective pairs, and issue occurrences
            elif row["Insp_Variance"] == 0 and row["Fail_Pairs_Variance"] == 0 and row["Issue_Variance"] == 0:
                return "Matched_Exact"
            else:
                return "Different_Count"

        lot_comp["Validation_Status"] = lot_comp.apply(assign_status, axis=1)

        lot_cols_order = [
            "Date_Std", "FTY", "Plant_Std", "Line_Std",
            "QA_InspQty", "QA_FailQty", "QA_IssueQty_Sum",
            "HFPA_Audit_Sample", "HFPA_Defective_Pairs", "HFPA_Defect_Qty",
            "Insp_Variance", "Fail_Pairs_Variance", "Issue_Variance", "Defect_Variance",
            "Validation_Status", "MES_Plant_Std"
        ]
        lot_comp = lot_comp[[c for c in lot_cols_order if c in lot_comp.columns]]

        # Merge validation results back into QAStation rows
        qa_val = pd.merge(
            qa_df,
            lot_comp[LOT_KEYS + ["HFPA_Audit_Sample", "HFPA_Defective_Pairs", "HFPA_Defect_Qty", "Insp_Variance", "Fail_Pairs_Variance", "Issue_Variance", "Defect_Variance", "Validation_Status", "MES_Plant_Std"]],
            on=LOT_KEYS,
            how="left"
        )

        qa_val, audit_records = cls.correct_model_inspections(qa_val, hfpa_df, sample_col, fail_pair_col)

        corrections = sum(row["Status"] == "MES410_Corrected" for row in audit_records)
        print(f"   [+] Completed lot reconciliation ({len(lot_comp):,} lots, {corrections} corrections, {len(audit_records) - corrections} to review) in {time.time()-t0:.2f}s")
        return qa_val, lot_comp, audit_records

    @staticmethod
    def correct_model_inspections(qa, mes, sample_col, fail_col):
        """Apply MES410 counts only to matching QAStation models whose pair counts differ."""
        qa["InspQty_Validated"] = qa["InspQty_Raw"]
        qa["FailQty_Validated"] = qa["FailQty_Raw"]
        qa["IssueQty_Validated"] = qa["IssueQty_Raw"].fillna(0).astype(int)
        qa["Adjustment_Status"] = "QAStation"
        qa["_HFPA_Source_Order"] = np.arange(len(qa))
        mes_groups = mes.groupby(LOT_KEYS + ["Model_Std"]).indices
        mes_lots = mes.groupby(LOT_KEYS).indices
        issue_cols = mes410_defect_columns(mes) if MES_ISSUE_TOTAL in mes else []

        audit_records = []
        replacement_rows = []
        replaced = []
        referenced_models = set()
        for key, idx in qa.groupby(MODEL_KEYS).indices.items():
            source = qa.iloc[idx]
            mes_plant = source["MES_Plant_Std"].iloc[0] if "MES_Plant_Std" in source else key[2]
            mes_plant = mes_plant if pd.notna(mes_plant) else key[2]
            reference_key = (*key[:2], mes_plant, *key[3:])
            referenced_models.add(reference_key)
            qa_sample, qa_failed = source[["InspQty_Raw", "FailQty_Raw"]].drop_duplicates().sum()
            before = {
                **dict(zip(MODEL_KEYS, key)),
                "QA_InspQty": qa_sample,
                "QA_FailQty": qa_failed,
                "QA_IssueQty": source["IssueQty_Validated"].sum(),
            }
            if not key[-1] and reference_key in mes_groups and qa_failed == 0 and before["QA_IssueQty"] == 0:
                reference = mes.iloc[mes_groups[reference_key]]
                samples = pd.to_numeric(reference[sample_col], errors="coerce")
                failures = pd.to_numeric(reference[fail_col], errors="coerce")
                totals = (pd.to_numeric(reference[MES_ISSUE_TOTAL], errors="coerce")
                          if MES_ISSUE_TOTAL in reference else pd.Series(np.nan, index=reference.index))
                counts = reference[issue_cols].fillna(0).apply(pd.to_numeric, errors="coerce")
                if (samples.notna().all() and samples.sum() == qa_sample and failures.eq(0).all()
                        and totals.eq(0).all() and counts.eq(0).all().all()):
                    # Both exports identify the same unnamed, all-pass inspection.
                    qa.loc[source.index, "Adjustment_Status"] = "MES410_Corrected"
                    audit_records.append({
                        **before, "Status": "MES410_Corrected", "Reason": "MES410_Blank_Model_Matched",
                        "MES_InspQty": samples.sum(), "MES_FailQty": 0, "MES_IssueQty": 0,
                    })
                    continue
            if not key[-1] or reference_key not in mes_groups:
                if qa_sample > 0 or qa_failed > 0 or before["QA_IssueQty"] > 0:
                    # The supplied MES410 window contains this day/line, but records no audit for this model.
                    if key[-1] and reference_key[:-1] in mes_lots:
                        qa.loc[source.index, ["InspQty_Validated", "FailQty_Validated", "IssueQty_Validated"]] = 0
                        qa.loc[source.index, "Adjustment_Status"] = "MES410_Corrected"
                        audit_records.append({
                            **before, "Status": "MES410_Corrected", "Reason": "MES410_No_Model",
                            "MES_InspQty": 0, "MES_FailQty": 0, "MES_IssueQty": 0,
                        })
                        continue
                    qa.loc[source.index, "Adjustment_Status"] = "No_MES410_Model"
                    audit_records.append({
                        **before, "Status": "No_MES410_Model",
                        "MES_InspQty": None, "MES_FailQty": None, "MES_IssueQty": None,
                    })
                continue

            reference = mes.iloc[mes_groups[reference_key]]
            mes_sample, mes_failed = reference[[sample_col, fail_col]].sum(min_count=1)
            if (qa_sample, qa_failed) == (mes_sample, mes_failed):
                continue

            counts = reference[issue_cols].fillna(0).apply(pd.to_numeric, errors="coerce")
            totals = pd.to_numeric(reference[MES_ISSUE_TOTAL], errors="coerce") if MES_ISSUE_TOTAL in reference else pd.Series(np.nan, index=reference.index)
            can_correct = (
                source["Station_Raw"].nunique() == 1
                and reference[[sample_col, fail_col]].notna().all().all()
                and pd.notna(mes_sample) and pd.notna(mes_failed)
                and 0 <= mes_failed <= mes_sample
                and mes_sample % 1 == 0 and mes_failed % 1 == 0
                and (bool(issue_cols) or totals.eq(0).all())
                and counts.notna().all().all()
                and counts.ge(0).all().all() and counts.mod(1).eq(0).all().all()
                and counts.sum(axis=1).eq(totals).all()
            )
            status = "MES410_Corrected" if can_correct else "Review_MES410"
            audit_records.append({
                **before,
                "Status": status,
                "MES_InspQty": mes_sample,
                "MES_FailQty": mes_failed,
                "MES_IssueQty": totals.sum(min_count=1),
            })
            if not can_correct:
                qa.loc[source.index, "Adjustment_Status"] = status
                continue

            by_issue = counts.sum().groupby(normalize_defect_name, sort=False).sum()
            issue_order = source["Issues_Std"].drop_duplicates().tolist()
            issue_order.extend(name for name in by_issue.index if name not in issue_order)
            by_issue = by_issue.reindex(issue_order, fill_value=0)
            issue_rows = [(name, int(qty)) for name, qty in by_issue.items() if qty > 0] or [("", 0)]
            original = source.iloc[0].to_dict()
            original.update(InspQty_Raw=qa_sample, FailQty_Raw=qa_failed)
            for issue, quantity in issue_rows:
                replacement_rows.append({
                    **original,
                    "InspQty_Validated": int(mes_sample),
                    "FailQty_Validated": int(mes_failed),
                    "Issues_Std": issue,
                    "IssueQty_Validated": quantity,
                    "Adjustment_Status": status,
                })
            replaced.extend(source.index)

        for key in sorted(mes_groups.keys() - referenced_models):
            reference = mes.iloc[mes_groups[key]]
            sample, failed = reference[[sample_col, fail_col]].sum(min_count=1)
            if sample > 0 or failed > 0:
                audit_records.append({
                    **dict(zip(MODEL_KEYS, key)), "Status": "Missing_QAStation_Model",
                    "QA_InspQty": 0, "QA_FailQty": 0, "QA_IssueQty": 0,
                    "MES_InspQty": sample, "MES_FailQty": failed,
                    "MES_IssueQty": reference[MES_ISSUE_TOTAL].sum(min_count=1) if MES_ISSUE_TOTAL in reference else None,
                })
        if replacement_rows:
            qa = pd.concat([qa.drop(index=replaced), pd.DataFrame(replacement_rows)], ignore_index=True)
        qa = qa.sort_values("_HFPA_Source_Order", kind="stable").drop(columns="_HFPA_Source_Order").reset_index(drop=True)
        return qa, audit_records

    @staticmethod
    def inspection_pivot(qa_val: pd.DataFrame) -> pd.DataFrame:
        """Count each corrected Quality Tracking inspection once across its issue rows."""
        original = quality_tracking_pairs(
            qa_val, QA_INSPECTION_KEYS, ["InspQty_Raw", "FailQty_Raw"]
        ).rename(columns={"InspQty_Raw": "InspQty_Original", "FailQty_Raw": "FailQty_Original"})
        adjusted = quality_tracking_pairs(
            qa_val, QA_INSPECTION_KEYS, ["InspQty_Validated", "FailQty_Validated"]
        ).rename(columns={"InspQty_Validated": "InspQty", "FailQty_Validated": "FailQty"})
        p1 = original.join(adjusted).reset_index()

        p1.rename(columns={
            "Date_Std": "InspDate",
            "Station_Raw": "Station",
            "Plant_Std": "Plant",
            "Line_Std": "Line",
            "ShoeName_Std": "Shoename",
        }, inplace=True)

        p1["DR%_Original"] = np.where(p1["InspQty_Original"] > 0, p1["FailQty_Original"] / p1["InspQty_Original"], 0.0)
        p1["DR%_Validated"] = np.where(p1["InspQty"] > 0, p1["FailQty"] / p1["InspQty"], 0.0)

        cols_order = [
            "InspDate", "FTY", "Station", "Plant", "Line", "Shoename",
            "InspQty_Original", "InspQty", "FailQty_Original", "FailQty", "DR%_Original", "DR%_Validated"
        ]
        return p1[cols_order]

    def generate_aggregations(
        self, qa_val: pd.DataFrame, hfpa_df: pd.DataFrame, color_map: Dict[str, str]
    ) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, pd.DataFrame], pd.DataFrame]:
        """Generate Pivot #1 and Pivot #2 aggregations in memory."""
        t0 = time.time()
        print("   [*] Generating Pivot #1 and Pivot #2 aggregations...")
        p1 = self.inspection_pivot(qa_val)

        # SOP pages 16–17 use the corrected Quality Tracking inspection and issue counts.
        all_top5, top5_dict = compute_top5_models(hfpa_df, p1)

        # Pivot 2: Top Defects per Top 5 Model (via canonical logic)
        p2_df = compute_top3_defects(qa_val, top5_dict, color_map)

        print(f"   [+] Aggregations generated ({len(all_top5)} Top5 rows, {len(p2_df)} Pivot2 rows) in {time.time()-t0:.2f}s")
        return p1, all_top5, top5_dict, p2_df

    def export_master_workbook(self, bus: DataBus) -> str:
        """
        Export the source data, adjusted Quality Tracking and analysis together.
        """
        t0 = time.time()
        master_path = os.path.join(self.output_dir, f"HFPA_FTT_Combined_Master_{bus.active_month}_{bus.active_year}.xlsx")
        print(f"   [*] Streaming Master Combined Workbook to: {master_path}...")

        qa_cols = [
            "InspDate", "FTY", "Line", "ShoeName", "InspQty",
            "FailQty", "Station", "Issues", "IssueQty"
        ]
        fty_order = {f: i for i, f in enumerate(["VH", "VH2", "JV", "JV2"])}
        qa_raw = bus.qa_raw.sort_values("FTY", key=lambda fty: fty.map(fty_order), kind="stable")
        qa_export = qa_raw[[c for c in qa_cols if c in qa_raw.columns]]
        adjusted_cols = {
            "InspQty": "InspQty_Validated", "FailQty": "FailQty_Validated",
            "Issues": "Issues_Std", "IssueQty": "IssueQty_Validated",
        }
        qa_adjusted = bus.qa_validated.sort_values("FTY", key=lambda fty: fty.map(fty_order), kind="stable")
        adjusted_export = qa_adjusted[
            [adjusted_cols.get(col, col) for col in qa_cols] + ["Adjustment_Status"]
        ].rename(columns={source: label for label, source in adjusted_cols.items()})

        with pd.ExcelWriter(master_path, engine="openpyxl") as writer:
            qa_export.to_excel(writer, sheet_name="HFPA_Quality_Tracking", index=False)
            adjusted_export.to_excel(writer, sheet_name="HFPA_Quality_Tracking_Adjusted", index=False)
            bus.hfpa_raw.to_excel(writer, sheet_name="HFPA_Mes410", index=False)
            bus.pivot1_df.to_excel(writer, sheet_name="Pivot1_Validation", index=False)
            bus.all_top5_df.to_excel(writer, sheet_name="Top5_Models", index=False)
            bus.pivot2_df.to_excel(writer, sheet_name="Top3_Defects", index=False)
            bus.lot_comparison.to_excel(writer, sheet_name="Lot_Defect_Variance", index=False)
            if bus.audit_records:
                pd.DataFrame(bus.audit_records).to_excel(writer, sheet_name="Model_Adjustments", index=False)

        print(f"   [+] Master Combined Workbook created in {time.time()-t0:.2f}s")
        return master_path

    def export_analysis_and_audit(self, bus: DataBus) -> Tuple[str, str]:
        """Export Executive Analysis Report and Audit Trail."""
        t0 = time.time()
        report_path = os.path.join(self.output_dir, f"QAStation_HFPA_Analysis_Report_{bus.active_month}_{bus.active_year}.xlsx")
        audit_path = os.path.join(self.logs_dir, f"reconciliation_audit_trail_{bus.active_month}_{bus.active_year}.csv")

        # Executive summary
        exec_summary = []
        for fty in FACTORIES + ["Grand Total"]:
            sub_qa = bus.pivot1_df if fty == "Grand Total" else bus.pivot1_df[bus.pivot1_df["FTY"] == fty]
            sub_hfpa = bus.hfpa_raw if fty == "Grand Total" else bus.hfpa_raw[bus.hfpa_raw["FTY"] == fty]
            sub_lot = bus.lot_comparison if fty == "Grand Total" else bus.lot_comparison[bus.lot_comparison["FTY"] == fty]

            qa_sample = sub_qa["InspQty_Original"].sum()
            qa_def = sub_qa["FailQty_Original"].sum()
            tot_sample = sub_qa["InspQty"].sum()
            corrected_failed_pairs = sub_qa["FailQty"].sum()
            hfpa_col = "Defect Q'ty (Total Defect Q'ty)" if "Defect Q'ty (Total Defect Q'ty)" in sub_hfpa.columns else "Defective Q'ty(pair)"
            hfpa_def = sub_hfpa[hfpa_col].sum() if hfpa_col in sub_hfpa.columns else 0

            matched_exact = len(sub_lot[sub_lot["Validation_Status"] == "Matched_Exact"])
            total_lots = len(sub_lot)

            exec_summary.append({
                "Factory": fty,
                "Total Audited Sample (InspQty)": tot_sample,
                "QAStation Defects (Original)": qa_def,
                "QAStation Defects (Validated/HFPA)": corrected_failed_pairs,
                "HFPA Total Defects": hfpa_def,
                "DR% (Original)": (qa_def / qa_sample) if qa_sample > 0 else 0.0,
                "DR% (Validated)": (corrected_failed_pairs / tot_sample) if tot_sample > 0 else 0.0,
                "Reconciled Lots": f"{matched_exact}/{total_lots} ({(matched_exact/total_lots*100):.1f}%)" if total_lots > 0 else "N/A",
                "QAStation InspQty (Original)": qa_sample,
            })

        summary_df = pd.DataFrame(exec_summary)

        with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="Executive_Summary", index=False)
            bus.all_top5_df.to_excel(writer, sheet_name="Top5_Shoes", index=False)
            bus.pivot2_df.to_excel(writer, sheet_name="Pivot2_Top3_Defects", index=False)

        # Audit trail
        if bus.audit_records:
            audit_df = pd.DataFrame(bus.audit_records)
            audit_df.to_csv(audit_path, index=False, encoding="utf-8-sig")

        print(f"   [+] Executive Analysis & Audit Trail exported in {time.time()-t0:.2f}s")
        return report_path, audit_path

    def run_reconciliation(self, target_month: str, target_year: str, *, input_files=None) -> DataBus:
        """Run full in-memory reconciliation pipeline and return populated DataBus."""
        t_start = time.time()
        qa_raw, hfpa_raw = self.load_raw_datasets(target_month, target_year, input_files=input_files)

        # Proactively provision colors for any new defect types into Color_template.xlsx (non-red, non-green)
        try:
            try:
                from .color_manager import ColorManager
            except (ImportError, ValueError):
                from agents.color_manager import ColorManager
            cm = ColorManager(self.color_template_path)
            raw_defects = qa_raw["Issues_Raw"].dropna().unique().tolist()
            cm.sync_unmapped_defects(raw_defects)
        except Exception as e:
            print(f"[-] ColorManager auto-sync warning: {e}")

        color_map = self.load_color_template()
        qa_val, lot_comp, audit_records = self.reconcile_and_validate(qa_raw, hfpa_raw)
        p1, all_top5, top5_dict, p2 = self.generate_aggregations(qa_val, hfpa_raw, color_map)

        bus = DataBus(
            active_month=target_month,
            active_year=target_year,
            qa_raw=qa_raw,
            hfpa_raw=hfpa_raw,
            qa_validated=qa_val,
            hfpa_clean=hfpa_raw,
            lot_comparison=lot_comp,
            pivot1_df=p1,
            all_top5_df=all_top5,
            top5_dict=top5_dict,
            pivot2_df=p2,
            audit_records=audit_records,
            color_map=color_map,
        )

        # Export deliverables
        self.export_master_workbook(bus)
        self.export_analysis_and_audit(bus)

        total_elapsed = time.time() - t_start
        bus.elapsed_times["reconciliation_total"] = total_elapsed
        print(f"[AGENT 2] In-memory reconciliation completed successfully in {total_elapsed:.2f}s!")
        return bus
