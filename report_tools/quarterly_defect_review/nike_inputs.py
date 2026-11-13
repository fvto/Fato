"""Strict HFPA NIKE and FTT NIKE input normalization, grain dedup, and schema checks."""

from __future__ import annotations

import hashlib
import os
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import openpyxl
import xlrd

from report_tools.quarterly_defect_review.area_mapping import (
    normalize_factory_code,
    normalize_line_code,
    normalize_plant_code,
    split_combined_plant_line,
)
from report_tools.quarterly_defect_review.contracts import (
    BaseRecord,
    DefectFact,
    MappingReviewItem,
    SourceManifest,
)
from report_tools.quarterly_defect_review.defect_mapping import (
    clean_label_key,
    load_defect_mapping,
    map_defects,
)

HFPA_NIKE_REQUIRED_COLUMNS = {
    'factory group cd_uccid',
    'audit date (dd-mmm-yyyy)',
    'plant',
    'line',
    'model',
    'audit sample size',
    "defect q'ty (total defect q'ty)",
}
HFPA_INTERNAL_FORBIDDEN = {'so#', 'country', 'pass/fail', 'model 2'}
HFPA_BASE_COLUMNS = {
    'factory group cd_uccid',
    'season',
    'audit date (dd-mmm-yyyy)',
    'internal audit id (yymmddno+plant+line)',
    'plant',
    'line',
    'auditor',
    'po nbr (header+line item nbr)',
    'model',
    'style nbr',
    'colorway cd',
    'time table',
    "produced q'ty",
    'audit sample size',
    "defective q'ty(pair)",
    'defective rate',
    "defect q'ty (total defect q'ty)",
}

FTT_NIKE_REQUIRED_COLUMNS = {
    'factory code',
    'plant',
    'line',
    'inspection date (dd-mmm-yyyy)',
    'model',
    'total a grade pairs',
    'total rework',
    'total b grade pairs (reject)',
    'total c grade pairs (reject)',
}
FTT_BASE_COLUMNS = {
    'factory group',
    'factory code',
    'plant',
    'line',
    'inspection date (dd-mmm-yyyy)',
    'season',
    'model',
    'style code',
    'colorway code',
    'total a grade pairs',
    '1st time rework pairs',
    'total rework',
    'total b grade pairs (reject)',
    'total c grade pairs (reject)',
}
FTT_INTERNAL_FORBIDDEN = {'defect issues', "issues q'ty", 'defective area', 'total pair produced', 'metric'}


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open('rb') as handle:
        for chunk in iter(lambda: handle.read(65536), b''):
            digest.update(chunk)
    return digest.hexdigest()


def infer_factory_from_filename(filename: str) -> str:
    match = re.search(r'\b(VH2|VH|JV2|JV)\b', str(filename).upper())
    if match:
        return match.group(1)
    match2 = re.search(r'[-_](VH2|VH|JV2|JV)(?:[-_.]|\b)', str(filename).upper())
    return match2.group(1) if match2 else ''


def parse_excel_date(value: Any, datemode: int = 0) -> str:
    """Parse ISO/Excel serial/dd-mmm-yyyy date strings safely into YYYY-MM-DD."""
    if value is None or value == '':
        return ''
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (int, float)) and value > 20000:
        if datemode == 1:
            base = datetime(1904, 1, 1)
        else:
            base = datetime(1899, 12, 30)
        return (base + timedelta(days=float(value))).date().isoformat()
    text = str(value).strip()
    if not text:
        return ''
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}.*', text):
        return text[:10]
    for fmt in ('%d-%b-%Y', '%d-%B-%Y', '%d/%m/%Y', '%Y/%m/%d', '%m/%d/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(text[:11], fmt).date().isoformat()
        except ValueError:
            continue
    return text[:10]


def to_float(value: Any) -> float:
    if value is None or value == '':
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(',', '').replace('%', '')
    if not text:
        return 0.0
    return float(text)


def _is_blank_row(row: list[Any] | tuple[Any, ...]) -> bool:
    return all(v is None or str(v).strip() == '' for v in row)


def _is_hfpa_footer_row(row: list[Any], col_idx: dict[str, int], data_count: int) -> bool:
    if not row:
        return False
    ucc = str(row[col_idx['factory group cd_uccid']] if col_idx.get('factory group cd_uccid', -1) < len(row) else '').strip()
    audit_id = ''
    if 'internal audit id (yymmddno+plant+line)' in col_idx:
        idx = col_idx['internal audit id (yymmddno+plant+line)']
        audit_id = str(row[idx] if idx < len(row) else '').strip()
    date_val = str(row[col_idx['audit date (dd-mmm-yyyy)']] if 'audit date (dd-mmm-yyyy)' in col_idx else '').strip()
    if ucc or audit_id:
        return False
    clean_count = date_val.replace(',', '').strip()
    if clean_count.isdigit() and (int(clean_count) == data_count or data_count == 0):
        return True
    return clean_count == '0' and data_count == 0


def _read_tabular_rows(path: Path, preferred_sheet: str | None = None) -> tuple[str, list[list[Any]], int]:
    """Read raw rows cleanly from .xls, .xlsx, or .xlsm without invoking Excel/WPS."""
    suffix = path.suffix.lower()
    if suffix == '.xls':
        with open(os.devnull, 'w', encoding='utf-8') as devnull:
            book = xlrd.open_workbook(str(path), logfile=devnull)
        sheet_name = preferred_sheet if preferred_sheet in book.sheet_names() else book.sheet_names()[0]
        sheet = book.sheet_by_name(sheet_name)
        rows = []
        for rx in range(sheet.nrows):
            row_vals = []
            for cx in range(sheet.ncols):
                cell = sheet.cell(rx, cx)
                val = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    val = xlrd.xldate_as_datetime(val, book.datemode)
                row_vals.append(val)
            rows.append(row_vals)
        return sheet_name, rows, book.datemode

    if suffix in ('.xlsx', '.xlsm'):
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        try:
            sheet_name = preferred_sheet if preferred_sheet and preferred_sheet in wb.sheetnames else wb.sheetnames[0]
            ws = wb[sheet_name]
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            return sheet_name, rows, 0
        finally:
            wb.close()

    raise ValueError(f'Định dạng file không được hỗ trợ cho {path.name!r}. Chỉ nhận .xls, .xlsx, .xlsm.')


def _detect_header_and_metric(rows: list[list[Any]], path: Path) -> tuple[int, list[str], dict[str, int]]:
    for idx, raw in enumerate(rows[:25]):
        norm_cols = [clean_label_key(c) for c in raw]
        non_empty = [c for c in norm_cols if c]
        if len(non_empty) < 5:
            continue
        norm_set = set(non_empty)
        if HFPA_NIKE_REQUIRED_COLUMNS <= norm_set or FTT_NIKE_REQUIRED_COLUMNS <= norm_set or (
            'metric' in norm_set and ('defect group' in norm_set or 'standard group' in norm_set or 'detail defect' in norm_set)
        ):
            col_idx: dict[str, int] = {}
            for pos, cname in enumerate(norm_cols):
                if cname and cname not in col_idx:
                    col_idx[cname] = pos
            return idx, [str(c or '').strip() for c in raw], col_idx
    raise ValueError(f'Không nhận diện được hàng tiêu đề hợp lệ trong {path.name!r}.')


def _parse_ftt_defect_column(raw_col: str) -> tuple[str, str] | None:
    """Extract (subtype_raw, severity) from 'Bond gap / Rat hole_Rework' etc."""
    text = re.sub(r'\s+', ' ', str(raw_col or '').replace('\u200b', '').replace('\xa0', ' ')).strip()
    for sev_suffix, sev_name in (('_Rework', 'Rework'), ('_B-Grade', 'B-Grade'), ('_C-Grade', 'C-Grade')):
        if text.lower().endswith(sev_suffix.lower()):
            subtype = text[: -len(sev_suffix)].strip()
            return subtype, sev_name
    return None


def _parse_multi_locations(location_raw: str) -> tuple[str, str]:
    text = str(location_raw or '').strip()
    if not text:
        return '', 'single_or_none'
    if any(sep in text for sep in (';', ',', '/', '&', '+')):
        return text, 'multi_location_unsplit'
    return text, 'single'


def normalize_nike_source(
    source_inputs: list[Path | str | dict[str, Any]],
    metric: str,
    run_id: str = 'manual',
    date_from: str | None = None,
    date_to: str | None = None,
    mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Normalize NIKE source files or in-memory normalized rows cleanly into BaseRecord + DefectFact."""
    metric = metric.upper()
    if metric not in ('HFPA', 'FTT'):
        raise ValueError(f'Metric không hợp lệ: {metric!r}')

    mapping_cfg = mapping or load_defect_mapping()
    manifests: list[SourceManifest] = []
    base_records: list[BaseRecord] = []
    raw_facts: list[DefectFact] = []
    review_items: list[MappingReviewItem] = []

    seen_file_hashes: set[str] = set()
    seen_grain_occurrences: set[tuple[Any, ...]] = set()
    control_mismatch_blocking = False

    # Support either file paths / dicts describing files OR in-memory synthetic rows (for pure unit tests)
    for item in source_inputs:
        if isinstance(item, dict) and 'rows' in item and 'file_path' not in item:
            _ingest_memory_rows(
                payload=item,
                metric=metric,
                run_id=run_id,
                date_from=date_from,
                date_to=date_to,
                mapping_cfg=mapping_cfg,
                manifests=manifests,
                base_records=base_records,
                raw_facts=raw_facts,
                review_items=review_items,
                seen_grain_occurrences=seen_grain_occurrences,
            )
            continue

        file_meta = dict(item) if isinstance(item, dict) else {'file_path': item}
        path = Path(file_meta['file_path'])
        if not path.is_file():
            raise FileNotFoundError(f'Không tìm thấy file nguồn NIKE: {path}')

        sha = file_sha256(path)
        factory_hint = normalize_factory_code(file_meta.get('factory') or infer_factory_from_filename(path.name))
        if sha in seen_file_hashes:
            manifests.append(
                SourceManifest(
                    run_id=run_id,
                    metric=metric,
                    report_mode='NIKE',
                    factory=factory_hint,
                    date_from=date_from or '',
                    date_to=date_to or '',
                    file_path=str(path),
                    file_name=path.name,
                    file_hash=sha,
                    schema_version='duplicate',
                    sheet_name='',
                    header_row=0,
                    duplicate_file_skipped=True,
                    coverage_notes=['File trùng SHA256 hoàn toàn đã được bỏ qua để tránh cộng đôi.'],
                )
            )
            continue
        seen_file_hashes.add(sha)

        preferred = 'FTT Template' if metric == 'FTT' else 'HFPA'
        sheet_name, rows, datemode = _read_tabular_rows(path, preferred_sheet=preferred)
        if not rows:
            raise ValueError(f'File {path.name!r} không có dữ liệu.')

        hdr_idx, raw_headers, col_idx = _detect_header_and_metric(rows, path)
        norm_set = set(col_idx.keys())

        if metric == 'HFPA':
            if norm_set & HFPA_INTERNAL_FORBIDDEN:
                raise ValueError(
                    f'File {path.name!r} là HFPA Internal (chứa {sorted(norm_set & HFPA_INTERNAL_FORBIDDEN)}), '
                    'tuyệt đối không dùng thay thế HFPA NIKE.'
                )
            if not (HFPA_NIKE_REQUIRED_COLUMNS <= norm_set):
                raise ValueError(f'File {path.name!r} không đúng schema HFPA NIKE.')
            _ingest_hfpa_table(
                path=path,
                sha=sha,
                sheet_name=sheet_name,
                hdr_idx=hdr_idx,
                raw_headers=raw_headers,
                col_idx=col_idx,
                rows=rows,
                datemode=datemode,
                factory_hint=factory_hint,
                run_id=run_id,
                date_from=date_from,
                date_to=date_to,
                mapping_cfg=mapping_cfg,
                manifests=manifests,
                base_records=base_records,
                raw_facts=raw_facts,
                review_items=review_items,
                seen_grain_occurrences=seen_grain_occurrences,
            )
        else:
            if norm_set & FTT_INTERNAL_FORBIDDEN:
                raise ValueError(
                    f'File {path.name!r} là FTT Internal (chứa {sorted(norm_set & FTT_INTERNAL_FORBIDDEN)}), '
                    'tuyệt đối không dùng thay thế FTT NIKE.'
                )
            if not (FTT_NIKE_REQUIRED_COLUMNS <= norm_set):
                raise ValueError(f'File {path.name!r} không đúng schema FTT NIKE.')
            _ingest_ftt_table(
                path=path,
                sha=sha,
                sheet_name=sheet_name,
                hdr_idx=hdr_idx,
                raw_headers=raw_headers,
                col_idx=col_idx,
                rows=rows,
                datemode=datemode,
                factory_hint=factory_hint,
                run_id=run_id,
                date_from=date_from,
                date_to=date_to,
                mapping_cfg=mapping_cfg,
                manifests=manifests,
                base_records=base_records,
                raw_facts=raw_facts,
                review_items=review_items,
                seen_grain_occurrences=seen_grain_occurrences,
            )

    mapped_facts, unmapped_reviews = map_defects(raw_facts, mapping=mapping_cfg)
    review_items.extend(unmapped_reviews)

    # Check if any detail sum > control in SAME unit & population
    if any(r.review_type == 'control_mismatch' and r.blocking for r in review_items):
        control_mismatch_blocking = True

    levels = {m.detail_support_level for m in manifests if not m.duplicate_file_skipped}
    if 'tier3_detail' in levels:
        overall_level = 'tier3_detail'
    elif 'tier2_subtype_only' in levels:
        overall_level = 'tier2_subtype_only'
    else:
        overall_level = 'tier1_group_only'

    return {
        'metric': metric,
        'manifests': manifests,
        'base_records': base_records,
        'defect_facts': mapped_facts,
        'review_items': review_items,
        'detail_support_level': overall_level,
        'has_blocking_mismatch': control_mismatch_blocking,
        'mapping_version': mapping_cfg['mapping_version'],
    }


def _ingest_hfpa_table(
    *,
    path: Path,
    sha: str,
    sheet_name: str,
    hdr_idx: int,
    raw_headers: list[str],
    col_idx: dict[str, int],
    rows: list[list[Any]],
    datemode: int,
    factory_hint: str,
    run_id: str,
    date_from: str | None,
    date_to: str | None,
    mapping_cfg: dict[str, Any],
    manifests: list[SourceManifest],
    base_records: list[BaseRecord],
    raw_facts: list[DefectFact],
    review_items: list[MappingReviewItem],
    seen_grain_occurrences: set[tuple[Any, ...]],
) -> None:
    # Identify group columns beyond HFPA_BASE_COLUMNS (plus optional observed detail column if an enriched NIKE export has one)
    defect_cols: list[tuple[int, str, str]] = []
    detail_col_idx = col_idx.get('detail defect') or col_idx.get('defect detail') or col_idx.get('detail_raw')
    detail_qty_idx = col_idx.get("detail q'ty") or col_idx.get('detail qty') or col_idx.get('detail_quantity')
    for pos, raw_h in enumerate(raw_headers):
        nk = clean_label_key(raw_h)
        if not nk or nk in HFPA_BASE_COLUMNS or pos in (detail_col_idx, detail_qty_idx):
            continue
        defect_cols.append((pos, raw_h.strip(), nk))

    data_rows = rows[hdr_idx + 1 :]
    raw_data_count = 0
    kept_rows = 0
    excluded_summary = 0
    excluded_empty = 0
    dedup_dropped = 0
    within_file_counts: dict[tuple[Any, ...], int] = {}
    dates_seen: list[str] = []
    factories_seen: set[str] = set()

    # Pre-scan index of non-empty trailing rows for DevExpress footer detection
    last_nonempty_rel = -1
    for i, r in enumerate(data_rows):
        if not _is_blank_row(r):
            last_nonempty_rel = i

    for rel_idx, raw in enumerate(data_rows):
        excel_row = hdr_idx + 2 + rel_idx
        padded = list(raw) + [None] * max(0, len(raw_headers) - len(raw))
        if _is_blank_row(padded):
            excluded_empty += 1
            continue
        if rel_idx == last_nonempty_rel and _is_hfpa_footer_row(padded, col_idx, raw_data_count):
            excluded_summary += 1
            continue

        raw_data_count += 1
        rec_date = parse_excel_date(padded[col_idx['audit date (dd-mmm-yyyy)']], datemode)
        if date_from and rec_date and rec_date < date_from:
            continue
        if date_to and rec_date and rec_date > date_to:
            continue

        ucc = str(padded[col_idx['factory group cd_uccid']] or '').strip()
        factory = normalize_factory_code(factory_hint or '', ucc_id=ucc)
        plant = normalize_plant_code(padded[col_idx['plant']])
        line_raw = str(padded[col_idx['line']] or '').strip()
        if not plant and line_raw:
            plant, line_raw = split_combined_plant_line(line_raw)
        line_std = normalize_line_code(line_raw)
        audit_id = str(padded[col_idx.get('internal audit id (yymmddno+plant+line)', -1)] if 'internal audit id (yymmddno+plant+line)' in col_idx else '').strip()
        po_nbr = str(padded[col_idx.get('po nbr (header+line item nbr)', -1)] if 'po nbr (header+line item nbr)' in col_idx else '').strip()
        model_label = str(padded[col_idx['model']] or '').strip()
        style_code = str(padded[col_idx.get('style nbr', -1)] if 'style nbr' in col_idx else '').strip()
        colorway = str(padded[col_idx.get('colorway cd', -1)] if 'colorway cd' in col_idx else '').strip()
        time_table = str(padded[col_idx.get('time table', -1)] if 'time table' in col_idx else '').strip()
        season = str(padded[col_idx.get('season', -1)] if 'season' in col_idx else '').strip()

        produced_qty = to_float(padded[col_idx.get("produced q'ty", -1)] if "produced q'ty" in col_idx else 0)
        sample_size = to_float(padded[col_idx['audit sample size']])
        defective_pairs = to_float(padded[col_idx.get("defective q'ty(pair)", -1)] if "defective q'ty(pair)" in col_idx else 0)
        defect_ctrl = to_float(padded[col_idx["defect q'ty (total defect q'ty)"]])

        grain_base = (
            'HFPA',
            rec_date,
            factory,
            plant,
            line_std,
            audit_id,
            po_nbr,
            model_label,
            style_code,
            colorway,
            time_table,
            round(sample_size, 4),
            round(defect_ctrl, 4),
        )
        occ_idx = within_file_counts.get(grain_base, 0)
        within_file_counts[grain_base] = occ_idx + 1
        full_grain = grain_base + (occ_idx,)
        if full_grain in seen_grain_occurrences:
            dedup_dropped += 1
            continue
        seen_grain_occurrences.add(full_grain)

        kept_rows += 1
        if rec_date:
            dates_seen.append(rec_date)
        if factory:
            factories_seen.add(factory)

        rec_id = f"HFPA:{path.name}:{excel_row}:{audit_id or kept_rows}"
        base_rec = BaseRecord(
            source_record_id=rec_id,
            metric='HFPA',
            date=rec_date,
            factory=factory,
            plant=plant,
            line_raw=line_raw,
            line_std=line_std,
            model_key=model_label,
            model_label=model_label or '(Không rõ model)',
            inspected_or_produced_qty=sample_size,
            produced_qty=produced_qty,
            defective_pairs_qty=defective_pairs,
            defect_qty_control=defect_ctrl,
            quantity_unit='pairs',
            control_unit='issues',
            ucc_id=ucc,
            style_code=style_code,
            colorway_code=colorway,
            season=season,
            source_file=path.name,
            source_sheet=sheet_name,
            source_row=excel_row,
            schema_level='tier3_detail' if detail_col_idx is not None else 'tier1_group_only',
            has_zero_defects=(defect_ctrl == 0),
        )
        base_records.append(base_rec)

        row_group_sum = 0.0
        for pos, col_label, _nk in defect_cols:
            qty = to_float(padded[pos])
            if qty == 0.0:
                continue
            row_group_sum += qty
            raw_facts.append(
                DefectFact(
                    base_record_id=rec_id,
                    source_file=path.name,
                    source_sheet=sheet_name,
                    source_row=excel_row,
                    source_column=col_label,
                    metric='HFPA',
                    date=rec_date,
                    factory=factory,
                    plant=plant,
                    line_std=line_std,
                    model_key=base_rec.model_key,
                    model_label=base_rec.model_label,
                    standard_group_raw=col_label,
                    standard_group_id=None,
                    standard_group_label=None,
                    nike_subtype_raw='',
                    nike_subtype_id=None,
                    nike_subtype_label=None,
                    detail_raw='',
                    detail_id=None,
                    detail_label=None,
                    detail_tier='tier1_group_only',
                    severity='HFPA',
                    issue_quantity=qty,
                    quantity_unit='issues',
                    source_control_total=defect_ctrl,
                    source_control_unit='issues',
                )
            )

        if defect_ctrl > 0 and row_group_sum > defect_ctrl + 1e-6:
            review_items.append(
                MappingReviewItem(
                    review_type='control_mismatch',
                    metric='HFPA',
                    factory=factory,
                    plant=plant,
                    line_std=line_std,
                    raw_label=f'Dòng {excel_row}',
                    raw_key=rec_id,
                    suggested_or_confirmed_mapping='Tổng 11 nhóm vượt tổng Defect Q\'ty cùng đơn vị',
                    affected_quantity=round(row_group_sum - defect_ctrl, 4),
                    status='Không khớp control',
                    reason=f'Tổng 11 nhóm = {row_group_sum}, vượt quá Defect Q\'ty control = {defect_ctrl}.',
                    provenance=f'{path.name}:{sheet_name}:{excel_row}',
                    blocking=True,
                )
            )

    support_level = 'tier1_group_only' if defect_cols else 'empty_day_or_no_defect_cols'
    coverage_notes = []
    if kept_rows == 0 and raw_data_count == 0:
        coverage_notes.append('File ngày rỗng (không có bản ghi audit HFPA trong ngày).')
    if defect_cols:
        coverage_notes.append('Schema SFC HFPA NIKE chỉ có 11 nhóm chuẩn HFPA; chưa có cột chi tiết cấp 3.')

    manifests.append(
        SourceManifest(
            run_id=run_id,
            metric='HFPA',
            report_mode='NIKE',
            factory=','.join(sorted(factories_seen)) or factory_hint,
            date_from=date_from or (min(dates_seen) if dates_seen else ''),
            date_to=date_to or (max(dates_seen) if dates_seen else ''),
            file_path=str(path),
            file_name=path.name,
            file_hash=sha,
            schema_version='hfpa_nike_28col' if len(raw_headers) >= 28 else f'hfpa_nike_{len(raw_headers)}col',
            sheet_name=sheet_name,
            header_row=hdr_idx + 1,
            total_rows=raw_data_count,
            kept_rows=kept_rows,
            excluded_summary_rows=excluded_summary,
            excluded_empty_rows=excluded_empty,
            dedup_dropped_rows=dedup_dropped,
            detail_support_level=support_level,
            coverage_notes=coverage_notes,
            filter_snapshot={'report_mode': 'NIKE', 'factory_hint': factory_hint},
            raw_columns=raw_headers,
        )
    )


def _ingest_ftt_table(
    *,
    path: Path,
    sha: str,
    sheet_name: str,
    hdr_idx: int,
    raw_headers: list[str],
    col_idx: dict[str, int],
    rows: list[list[Any]],
    datemode: int,
    factory_hint: str,
    run_id: str,
    date_from: str | None,
    date_to: str | None,
    mapping_cfg: dict[str, Any],
    manifests: list[SourceManifest],
    base_records: list[BaseRecord],
    raw_facts: list[DefectFact],
    review_items: list[MappingReviewItem],
    seen_grain_occurrences: set[tuple[Any, ...]],
) -> None:
    defect_cols: list[tuple[int, str, str, str]] = []
    for pos, raw_h in enumerate(raw_headers):
        nk = clean_label_key(raw_h)
        if not nk or nk in FTT_BASE_COLUMNS:
            continue
        parsed = _parse_ftt_defect_column(raw_h)
        if parsed:
            subtype_raw, severity = parsed
        else:
            subtype_raw, severity = raw_h.strip(), 'Unknown'
        defect_cols.append((pos, raw_h.strip(), subtype_raw, severity))

    data_rows = rows[hdr_idx + 1 :]
    raw_data_count = 0
    kept_rows = 0
    excluded_empty = 0
    dedup_dropped = 0
    within_file_counts: dict[tuple[Any, ...], int] = {}
    dates_seen: list[str] = []
    factories_seen: set[str] = set()

    for rel_idx, raw in enumerate(data_rows):
        excel_row = hdr_idx + 2 + rel_idx
        padded = list(raw) + [None] * max(0, len(raw_headers) - len(raw))
        if _is_blank_row(padded[:14]):
            excluded_empty += 1
            continue
        # Ignore all-zero template blank tail rows if date & plant & line are empty
        if not str(padded[col_idx['plant']] or '').strip() and not str(padded[col_idx['line']] or '').strip():
            excluded_empty += 1
            continue

        raw_data_count += 1
        rec_date = parse_excel_date(padded[col_idx['inspection date (dd-mmm-yyyy)']], datemode)
        if date_from and rec_date and rec_date < date_from:
            continue
        if date_to and rec_date and rec_date > date_to:
            continue

        factory = normalize_factory_code(padded[col_idx['factory code']] or factory_hint)
        plant = normalize_plant_code(padded[col_idx['plant']])
        line_raw = str(padded[col_idx['line']] or '').strip()
        if not plant and line_raw:
            plant, line_raw = split_combined_plant_line(line_raw)
        line_std = normalize_line_code(line_raw)
        season = str(padded[col_idx.get('season', -1)] if 'season' in col_idx else '').strip()
        model_label = str(padded[col_idx['model']] or '').strip()
        style_code = str(padded[col_idx.get('style code', -1)] if 'style code' in col_idx else '').strip()
        colorway = str(padded[col_idx.get('colorway code', -1)] if 'colorway code' in col_idx else '').strip()

        a_pairs = to_float(padded[col_idx['total a grade pairs']])
        rew1_pairs = to_float(padded[col_idx.get('1st time rework pairs', -1)] if '1st time rework pairs' in col_idx else 0)
        rew_pairs = to_float(padded[col_idx['total rework']])
        b_pairs = to_float(padded[col_idx['total b grade pairs (reject)']])
        c_pairs = to_float(padded[col_idx['total c grade pairs (reject)']])
        produced_pairs = a_pairs + b_pairs + c_pairs

        # Defect columns sum = exact total Defect Issues across Rework + B-Grade + C-Grade
        row_defect_sum = sum(to_float(padded[pos]) for pos, _, _, _ in defect_cols)

        grain_base = (
            'FTT',
            rec_date,
            factory,
            plant,
            line_std,
            season,
            model_label,
            style_code,
            colorway,
            round(a_pairs, 4),
            round(rew_pairs, 4),
            round(b_pairs, 4),
            round(c_pairs, 4),
        )
        occ_idx = within_file_counts.get(grain_base, 0)
        within_file_counts[grain_base] = occ_idx + 1
        full_grain = grain_base + (occ_idx,)
        if full_grain in seen_grain_occurrences:
            dedup_dropped += 1
            continue
        seen_grain_occurrences.add(full_grain)

        kept_rows += 1
        if rec_date:
            dates_seen.append(rec_date)
        if factory:
            factories_seen.add(factory)

        rec_id = f"FTT:{path.name}:{excel_row}:{factory}:{plant}:{line_std}"
        base_rec = BaseRecord(
            source_record_id=rec_id,
            metric='FTT',
            date=rec_date,
            factory=factory,
            plant=plant,
            line_raw=line_raw,
            line_std=line_std,
            model_key=model_label,
            model_label=model_label or '(Không rõ model)',
            inspected_or_produced_qty=produced_pairs,
            produced_qty=produced_pairs,
            defective_pairs_qty=rew_pairs + b_pairs + c_pairs,
            defect_qty_control=row_defect_sum,
            rework_pairs_qty=rew_pairs,
            rework1_pairs_qty=rew1_pairs,
            a_grade_pairs=a_pairs,
            b_grade_pairs=b_pairs,
            c_grade_pairs=c_pairs,
            quantity_unit='pairs',
            control_unit='issues',
            ucc_id=str(padded[col_idx.get('factory group', -1)] or '').strip() if 'factory group' in col_idx else '',
            style_code=style_code,
            colorway_code=colorway,
            season=season,
            source_file=path.name,
            source_sheet=sheet_name,
            source_row=excel_row,
            schema_level='tier2_subtype_only',
            has_zero_defects=(row_defect_sum == 0),
        )
        base_records.append(base_rec)

        for pos, col_label, subtype_raw, severity in defect_cols:
            qty = to_float(padded[pos])
            if qty == 0.0:
                continue
            raw_facts.append(
                DefectFact(
                    base_record_id=rec_id,
                    source_file=path.name,
                    source_sheet=sheet_name,
                    source_row=excel_row,
                    source_column=col_label,
                    metric='FTT',
                    date=rec_date,
                    factory=factory,
                    plant=plant,
                    line_std=line_std,
                    model_key=base_rec.model_key,
                    model_label=base_rec.model_label,
                    standard_group_raw='',
                    standard_group_id=None,
                    standard_group_label=None,
                    nike_subtype_raw=subtype_raw,
                    nike_subtype_id=None,
                    nike_subtype_label=None,
                    detail_raw=subtype_raw,
                    detail_id=None,
                    detail_label=None,
                    detail_tier='tier2_subtype_only',
                    severity=severity,
                    issue_quantity=qty,
                    quantity_unit='issues',
                    source_control_total=row_defect_sum,
                    source_control_unit='issues',
                )
            )

    manifests.append(
        SourceManifest(
            run_id=run_id,
            metric='FTT',
            report_mode='NIKE',
            factory=','.join(sorted(factories_seen)) or factory_hint,
            date_from=date_from or (min(dates_seen) if dates_seen else ''),
            date_to=date_to or (max(dates_seen) if dates_seen else ''),
            file_path=str(path),
            file_name=path.name,
            file_hash=sha,
            schema_version='ftt_nike_89col' if len(raw_headers) >= 89 else f'ftt_nike_{len(raw_headers)}col',
            sheet_name=sheet_name,
            header_row=hdr_idx + 1,
            total_rows=raw_data_count,
            kept_rows=kept_rows,
            excluded_empty_rows=excluded_empty,
            dedup_dropped_rows=dedup_dropped,
            detail_support_level='tier2_subtype_only',
            coverage_notes=['Schema FTT NIKE có 25 cột phân loại trung gian FTT x 3 cấp (Rework/B-Grade/C-Grade); chưa có cấp chi tiết thứ 3 ngoài 25 loại này.'],
            filter_snapshot={'report_mode': 'NIKE', 'factory_hint': factory_hint},
            raw_columns=raw_headers,
        )
    )


def _ingest_memory_rows(
    *,
    payload: dict[str, Any],
    metric: str,
    run_id: str,
    date_from: str | None,
    date_to: str | None,
    mapping_cfg: dict[str, Any],
    manifests: list[SourceManifest],
    base_records: list[BaseRecord],
    raw_facts: list[DefectFact],
    review_items: list[MappingReviewItem],
    seen_grain_occurrences: set[tuple[Any, ...]],
) -> None:
    """Ingest normalized fixture rows supporting base-vs-detail dedup, multi-tier control, and locations."""
    rows = payload.get('rows', [])
    source_name = str(payload.get('source_name') or 'memory_fixture')
    report_mode = str(payload.get('report_mode') or 'NIKE').upper()
    if report_mode != 'NIKE':
        raise ValueError(f'Nguồn {source_name!r} không ở chế độ NIKE (report_mode={report_mode!r}).')

    # Group rows per audit/lot ID when present, ensuring inspected/produced denominator is counted ONCE.
    grouped_rows: dict[str, list[dict[str, Any]]] = {}
    ordered_keys: list[str] = []
    for idx, row in enumerate(rows, 1):
        rec_date = str(row.get('date') or '')[:10]
        if date_from and rec_date and rec_date < date_from:
            continue
        if date_to and rec_date and rec_date > date_to:
            continue
        factory = normalize_factory_code(row.get('factory') or '')
        plant = normalize_plant_code(row.get('plant') or '')
        line_raw = str(row.get('line') or '').strip()
        line_std = normalize_line_code(line_raw)
        model_label = str(row.get('model') or row.get('model_label') or '').strip()
        audit_key = str(row.get('audit_id') or row.get('source_record_id') or '').strip()
        if audit_key:
            base_key = f"{metric}:{factory}:{plant}:{line_std}:{audit_key}"
        else:
            # Distinct row ordinal when no audit/lot ID supplied
            base_key = f"{metric}:{source_name}:{idx}"
        if base_key not in grouped_rows:
            grouped_rows[base_key] = []
            ordered_keys.append(base_key)
        grouped_rows[base_key].append(dict(row, _row_num=idx))

    has_tier3 = False
    has_tier2 = False
    kept_records = 0

    for base_key in ordered_keys:
        grp = grouped_rows[base_key]
        head = grp[0]
        factory = normalize_factory_code(head.get('factory') or '')
        plant = normalize_plant_code(head.get('plant') or '')
        line_raw = str(head.get('line') or '').strip()
        line_std = normalize_line_code(line_raw)
        rec_date = str(head.get('date') or '')[:10]
        model_label = str(head.get('model') or head.get('model_label') or '').strip()
        inspected_qty = to_float(head.get('inspected_qty') if head.get('inspected_qty') is not None else head.get('sample_size') or head.get('produced_qty') or 0)
        produced_qty = to_float(head.get('produced_qty') or inspected_qty)
        defective_pairs = head.get('defective_pairs')
        def_pairs_val = to_float(defective_pairs) if defective_pairs is not None else None
        quantity_unit = str(head.get('quantity_unit') or 'pairs')
        control_unit = str(head.get('control_unit') or 'issues')

        # Separate explicit hierarchy levels so parent/subtype subtotals are NEVER added to detail facts!
        detail_rows = [r for r in grp if str(r.get('row_level') or 'detail').lower() == 'detail']
        parent_controls = [r for r in grp if str(r.get('row_level') or '').lower() in ('parent', 'group', 'subtotal')]

        # Explicit parent_control_total either on parent_controls rows or attribute 'group_control_total' on rows
        group_controls: dict[str, tuple[float, str]] = {}
        for r in parent_controls:
            g_raw = str(r.get('standard_group') or r.get('defect_group') or '').strip()
            g_key = clean_label_key(g_raw)
            g_unit = str(r.get('issue_unit') or control_unit)
            group_controls[g_key] = (group_controls.get(g_key, (0.0, g_unit))[0] + to_float(r.get('issue_quantity')), g_unit)
        for r in detail_rows:
            if r.get('group_control_total') is not None:
                g_raw = str(r.get('standard_group') or r.get('defect_group') or '').strip()
                g_key = clean_label_key(g_raw)
                g_unit = str(r.get('control_unit') or control_unit)
                group_controls[g_key] = (to_float(r.get('group_control_total')), g_unit)

        detail_sum = 0.0
        detail_sum_by_group: dict[str, float] = {}
        record_facts: list[DefectFact] = []

        for r in detail_rows:
            detail_raw = str(r.get('detail_defect') or r.get('detail_raw') or '').strip()
            subtype_raw = str(r.get('ftt_subtype') or r.get('nike_subtype') or '').strip()
            group_raw = str(r.get('standard_group') or r.get('defect_group') or '').strip()
            qty = to_float(r.get('issue_quantity'))
            loc_raw, loc_status = _parse_multi_locations(str(r.get('location') or r.get('locations') or ''))
            issue_unit = str(r.get('issue_unit') or 'issues')

            if detail_raw:
                tier = 'tier3_detail'
                has_tier3 = True
            elif subtype_raw:
                tier = 'tier2_subtype_only'
                has_tier2 = True
            else:
                tier = 'tier1_group_only'

            # Multi-location in a single row NEVER multiplies issue_quantity!
            g_key = clean_label_key(group_raw)
            if qty == 0.0 and not detail_raw and not subtype_raw and not group_raw:
                continue

            detail_sum += qty
            detail_sum_by_group[g_key] = detail_sum_by_group.get(g_key, 0.0) + qty

            record_facts.append(
                DefectFact(
                    base_record_id=base_key,
                    source_file=source_name,
                    source_sheet='memory',
                    source_row=int(r['_row_num']),
                    source_column='detail_defect' if detail_raw else ('ftt_subtype' if subtype_raw else 'standard_group'),
                    metric=metric,
                    date=rec_date,
                    factory=factory,
                    plant=plant,
                    line_std=line_std,
                    model_key=model_label,
                    model_label=model_label or '(Không rõ model)',
                    standard_group_raw=group_raw,
                    standard_group_id=None,
                    standard_group_label=None,
                    nike_subtype_raw=subtype_raw,
                    nike_subtype_id=None,
                    nike_subtype_label=None,
                    detail_raw=detail_raw,
                    detail_id=str(r.get('detail_id') or '') or None,
                    detail_label=str(r.get('detail_label') or detail_raw or '') or None,
                    detail_tier=tier,
                    severity=str(r.get('severity') or metric),
                    issue_quantity=qty,
                    quantity_unit=issue_unit,
                    location_raw=loc_raw,
                    location_split_status=loc_status,
                )
            )

        # Reconcile control vs detail per group when explicit control totals exist
        total_control_for_rec = 0.0
        for g_key, (ctrl_val, ctrl_unit) in group_controls.items():
            d_sum = detail_sum_by_group.get(g_key, 0.0)
            total_control_for_rec += ctrl_val
            if ctrl_unit == control_unit:
                if d_sum > ctrl_val + 1e-6:
                    review_items.append(
                        MappingReviewItem(
                            review_type='control_mismatch',
                            metric=metric,
                            factory=factory,
                            plant=plant,
                            line_std=line_std,
                            raw_label=g_key,
                            raw_key=f'{base_key}:{g_key}',
                            suggested_or_confirmed_mapping='Kiểm tra tổng nhóm vs tổng chi tiết cùng đơn vị',
                            affected_quantity=round(d_sum - ctrl_val, 4),
                            status='Không khớp control',
                            reason=f'Tổng chi tiết ({d_sum}) > tổng nhóm control ({ctrl_val}) trong cùng đơn vị {ctrl_unit}.',
                            provenance=source_name,
                            blocking=True,
                        )
                    )
                elif ctrl_val > d_sum + 1e-6:
                    gap = ctrl_val - d_sum
                    record_facts.append(
                        DefectFact(
                            base_record_id=base_key,
                            source_file=source_name,
                            source_sheet='memory',
                            source_row=int(head['_row_num']),
                            source_column='unreconciled_gap',
                            metric=metric,
                            date=rec_date,
                            factory=factory,
                            plant=plant,
                            line_std=line_std,
                            model_key=model_label,
                            model_label=model_label or '(Không rõ model)',
                            standard_group_raw=g_key,
                            standard_group_id=None,
                            standard_group_label=None,
                            nike_subtype_raw='',
                            nike_subtype_id=None,
                            nike_subtype_label=None,
                            detail_raw='Chưa có chi tiết',
                            detail_id='__unreconciled_gap__',
                            detail_label='Chưa có chi tiết',
                            detail_tier='tier1_group_only',
                            severity=metric,
                            issue_quantity=gap,
                            quantity_unit=ctrl_unit,
                            detail_status='unreconciled_gap',
                            is_synthetic_gap=True,
                            mapping_basis='Chênh lệch giữa tổng nhóm control và tổng lỗi chi tiết đối chiếu được',
                        )
                    )
        # Enforce control checks cleanly even when a detail-only batch omits some parent totals
        if group_controls:
            # Start with explicit control totals, adding any mapped groups that lacked an explicit control
            uncontrolled_sum = sum(
                qty for gk, qty in detail_sum_by_group.items() if gk not in group_controls
            )
            total_control_for_rec = sum(v[0] for v in group_controls.values()) + uncontrolled_sum
        else:
            total_control_for_rec = detail_sum

        kept_records += 1
        base_records.append(
            BaseRecord(
                source_record_id=base_key,
                metric=metric,
                date=rec_date,
                factory=factory,
                plant=plant,
                line_raw=line_raw,
                line_std=line_std,
                model_key=model_label,
                model_label=model_label or '(Không rõ model)',
                inspected_or_produced_qty=inspected_qty,
                produced_qty=produced_qty,
                defective_pairs_qty=def_pairs_val,
                defect_qty_control=total_control_for_rec,
                quantity_unit=quantity_unit,
                control_unit=control_unit,
                source_file=source_name,
                source_sheet='memory',
                source_row=int(head['_row_num']),
                schema_level='tier3_detail' if has_tier3 else ('tier2_subtype_only' if has_tier2 else 'tier1_group_only'),
                has_zero_defects=(total_control_for_rec == 0 and detail_sum == 0),
            )
        )
        raw_facts.extend(record_facts)

    manifests.append(
        SourceManifest(
            run_id=run_id,
            metric=metric,
            report_mode='NIKE',
            factory=','.join(sorted({r.factory for r in base_records if r.source_file == source_name})),
            date_from=date_from or '',
            date_to=date_to or '',
            file_path=source_name,
            file_name=source_name,
            file_hash=hashlib.sha256(source_name.encode('utf-8')).hexdigest()[:16],
            schema_version='memory_normalized_nike',
            sheet_name='memory',
            header_row=1,
            total_rows=len(rows),
            kept_rows=kept_records,
            detail_support_level='tier3_detail' if has_tier3 else ('tier2_subtype_only' if has_tier2 else 'tier1_group_only'),
        )
    )
