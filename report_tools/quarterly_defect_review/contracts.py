"""Typed data contracts and validation helpers for quarterly NIKE defect review."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any

METRICS: tuple[str, ...] = ('HFPA', 'FTT')
AREAS: tuple[str, ...] = ('VH', 'JV', 'VB')
SCOPES: tuple[str, ...] = ('CLG', 'VH', 'JV', 'VB')

DEFAULT_PERIOD = 'FY27-Q1'
DEFAULT_DATE_RANGE: tuple[str, str] = ('2026-06-01', '2026-08-31')

STATUS_LABELS: dict[str, str] = {
    'checking': 'Đang kiểm tra',
    'missing_source': 'Thiếu nguồn',
    'partial': 'Có kết quả một phần',
    'needs_mapping_review': 'Cần review mapping',
    'ready_to_export': 'Sẵn sàng xuất',
    'exported': 'Xuất hoàn tất',
    'export_error': 'Lỗi xuất',
}

DETAIL_STATUS_LABELS: dict[str, str] = {
    'mapped': 'Đã ánh xạ đầy đủ',
    'subtype_only': 'Chỉ có cấp loại FTT NIKE',
    'group_only': 'Chưa có dữ liệu chi tiết',
    'no_defect': 'Không có lỗi',
    'missing_detail': 'Thiếu dữ liệu chi tiết',
    'unmapped': 'Chưa ánh xạ',
    'unreconciled_gap': 'Chưa có chi tiết',
}

COVERAGE_WARNING_THRESHOLD_PCT = 80.0


@dataclass(frozen=True)
class DRConfig:
    """Explicit placeholder for unconfirmed DR rate configuration."""

    status: str = 'unconfirmed'
    source_metrics: tuple[str, ...] = ('HFPA', 'FTT')
    requires_external_dr_file: bool = False
    formula_name: str | None = None
    numerator: str | None = None
    denominator: str | None = None
    note: str = (
        'Công thức DR (tên gọi, tử số, mẫu số) chưa được chốt chính thức. '
        'Khi chốt sẽ lấy từ HFPA NIKE và FTT NIKE; không dùng số/tỷ lệ từ ảnh Q2 và không tự đặt công thức.'
    )

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class SourceManifest:
    """Provenance record for a normalized NIKE source file."""

    run_id: str
    metric: str
    report_mode: str
    factory: str
    date_from: str
    date_to: str
    file_path: str
    file_name: str
    file_hash: str
    schema_version: str
    sheet_name: str
    header_row: int
    total_rows: int = 0
    kept_rows: int = 0
    excluded_summary_rows: int = 0
    excluded_empty_rows: int = 0
    dedup_dropped_rows: int = 0
    duplicate_file_skipped: bool = False
    detail_support_level: str = 'tier1_group_only'
    coverage_notes: list[str] = field(default_factory=list)
    filter_snapshot: dict[str, Any] = field(default_factory=dict)
    raw_columns: list[str] = field(default_factory=list)
    schema_fields: list[dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class LineCatalogEntry:
    """Factory Sandals line catalog entry sourced from FTT Group catalog."""

    factory: str
    plant: str
    line_raw: str
    line_std: str
    source_group: str
    report_area: str
    effective_from: str = '1900-01-01'
    effective_to: str = '2999-12-31'
    provenance: str = 'ftt_factory_sandals_catalog'

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class BaseRecord:
    """Single deduplicated inspection/production grain record (denominator-safe)."""

    source_record_id: str
    metric: str
    date: str
    factory: str
    plant: str
    line_raw: str
    line_std: str
    model_key: str
    model_label: str
    inspected_or_produced_qty: float
    produced_qty: float = 0.0
    defective_pairs_qty: float | None = None
    defect_qty_control: float = 0.0
    rework_pairs_qty: float = 0.0
    rework1_pairs_qty: float = 0.0
    a_grade_pairs: float = 0.0
    b_grade_pairs: float = 0.0
    c_grade_pairs: float = 0.0
    quantity_unit: str = 'pairs'
    control_unit: str = 'issues'
    report_area: str = 'UNMATCHED'
    area_status: str = 'pending'
    area_provenance: str = ''
    ucc_id: str = ''
    style_code: str = ''
    colorway_code: str = ''
    season: str = ''
    source_file: str = ''
    source_sheet: str = ''
    source_row: int = 0
    schema_level: str = 'tier1_group_only'
    has_zero_defects: bool = False
    metadata: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class DefectFact:
    """Single atomic defect observation linked to one BaseRecord."""

    base_record_id: str
    source_file: str
    source_sheet: str
    source_row: int
    source_column: str
    metric: str
    date: str
    factory: str
    plant: str
    line_std: str
    model_key: str
    model_label: str
    standard_group_raw: str
    standard_group_id: str | None
    standard_group_label: str | None
    nike_subtype_raw: str
    nike_subtype_id: str | None
    nike_subtype_label: str | None
    detail_raw: str
    detail_id: str | None
    detail_label: str | None
    detail_tier: str
    severity: str
    issue_quantity: float
    quantity_unit: str = 'issues'
    report_area: str = 'UNMATCHED'
    detail_status: str = 'mapped'
    mapping_version: str = ''
    mapping_basis: str = ''
    is_synthetic_gap: bool = False
    location_raw: str = ''
    location_split_status: str = 'single_or_none'
    source_control_total: float | None = None
    source_control_unit: str = 'issues'

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class MappingReviewItem:
    """Auditable review entry for unmapped labels, line conflicts, gaps, or blocking mismatches."""

    review_type: str
    metric: str
    factory: str
    plant: str
    line_std: str
    raw_label: str
    raw_key: str
    suggested_or_confirmed_mapping: str
    affected_quantity: float
    status: str
    reason: str
    provenance: str
    blocking: bool = False

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def fiscal_quarter_dates(period: str) -> tuple[str, str]:
    """Return inclusive ISO (date_from, date_to) for FYxx-Qy where Q1 starts June 1."""
    if not isinstance(period, str) or len(period) != 7 or not period.startswith('FY') or '-Q' not in period:
        raise ValueError(f'Kỳ quý không hợp lệ: {period!r}. Dùng định dạng FY27-Q1.')
    fy_short = int(period[2:4])
    quarter = int(period[6])
    if quarter not in (1, 2, 3, 4):
        raise ValueError(f'Quý không hợp lệ: {period!r}.')
    fy_end_year = 2000 + fy_short
    # Ching Luh fiscal year: Q1=Jun-Aug (prev year), Q2=Sep-Nov (prev year),
    # Q3=Dec (prev)-Feb (fy_end_year), Q4=Mar-May (fy_end_year).
    import calendar
    from datetime import date

    if quarter == 1:
        start = date(fy_end_year - 1, 6, 1)
        end = date(fy_end_year - 1, 8, 31)
    elif quarter == 2:
        start = date(fy_end_year - 1, 9, 1)
        end = date(fy_end_year - 1, 11, 30)
    elif quarter == 3:
        feb_days = calendar.monthrange(fy_end_year, 2)[1]
        start = date(fy_end_year - 1, 12, 1)
        end = date(fy_end_year, 2, feb_days)
    else:
        start = date(fy_end_year, 3, 1)
        end = date(fy_end_year, 5, 31)
    return start.isoformat(), end.isoformat()

