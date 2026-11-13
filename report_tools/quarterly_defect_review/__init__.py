"""Isolated quarterly NIKE defect review branch (HFPA NIKE & FTT NIKE)."""

from report_tools.quarterly_defect_review.analysis import (
    aggregate_defects,
    analyze_metric,
    build_review,
    rank_details,
    rank_groups,
    rank_models,
)
from report_tools.quarterly_defect_review.area_mapping import (
    build_area_catalog,
    classify_areas,
    normalize_line_code,
)
from report_tools.quarterly_defect_review.contracts import (
    AREAS,
    DEFAULT_DATE_RANGE,
    DEFAULT_PERIOD,
    DRConfig,
    METRICS,
    SCOPES,
)
from report_tools.quarterly_defect_review.defect_mapping import (
    load_defect_mapping,
    map_defects,
)
from report_tools.quarterly_defect_review.nike_inputs import normalize_nike_source


def run_report(*args, **kwargs):
    from report_tools.quarterly_defect_review.report import run_report as _run_report
    return _run_report(*args, **kwargs)

__all__ = [
    'AREAS',
    'DEFAULT_DATE_RANGE',
    'DEFAULT_PERIOD',
    'DRConfig',
    'METRICS',
    'SCOPES',
    'aggregate_defects',
    'analyze_metric',
    'build_area_catalog',
    'build_review',
    'classify_areas',
    'load_defect_mapping',
    'map_defects',
    'normalize_line_code',
    'normalize_nike_source',
    'rank_details',
    'rank_groups',
    'rank_models',
    'run_report',
]
