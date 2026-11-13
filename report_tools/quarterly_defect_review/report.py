"""End-to-end orchestrator for the isolated Quarterly NIKE Defect Review branch."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from dashboard_paths import APP_ROOT
from report_tools.quarterly_defect_review.analysis import analyze_metric
from report_tools.quarterly_defect_review.area_mapping import build_area_catalog, classify_areas
from report_tools.quarterly_defect_review.contracts import (
    DEFAULT_DATE_RANGE,
    DEFAULT_PERIOD,
    STATUS_LABELS,
    DRConfig,
    LineCatalogEntry,
    MappingReviewItem,
    fiscal_quarter_dates,
)
from report_tools.quarterly_defect_review.defect_mapping import load_defect_mapping
from report_tools.quarterly_defect_review.excel_output import build_excel_workbook
from report_tools.quarterly_defect_review.nike_inputs import normalize_nike_source
from report_tools.quarterly_defect_review.office_output import OfficeExportError, build_office_presentation

DEFAULT_ISOLATED_CACHE_DIR = APP_ROOT / 'Cache' / 'QuarterlyDefectReview'
DEFAULT_CATALOG_CACHE_PATH = DEFAULT_ISOLATED_CACHE_DIR / 'line_catalog_factory_sandals.json'
DEFAULT_MAPPING_OVERRIDES_PATH = DEFAULT_ISOLATED_CACHE_DIR / 'mapping_overrides.json'

SUPPORTED_TABULAR_EXTS = {'.xls', '.xlsx', '.xlsm', '.csv'}
SUPPORTED_CATALOG_EXTS = {'.json', '.xls', '.xlsx', '.xlsm', '.csv'}


def resolve_period_and_dates(
    period: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    cutoff: str | None = None,
) -> tuple[str, str, str]:
    """Resolve effective quarter and inclusive ISO date bounds (with optional cutoff clamp)."""
    eff_period = (period or DEFAULT_PERIOD).strip() or DEFAULT_PERIOD
    q_from, q_to = fiscal_quarter_dates(eff_period) if '-Q' in eff_period else DEFAULT_DATE_RANGE
    eff_from = str(date_from or q_from).strip()[:10]
    eff_to = str(date_to or q_to).strip()[:10]
    if cutoff and str(cutoff).strip():
        eff_cutoff = str(cutoff).strip()[:10]
        if eff_cutoff < eff_to:
            eff_to = eff_cutoff
    if eff_from > eff_to:
        raise ValueError(f'Khoảng ngày không hợp lệ: Từ ngày {eff_from} lớn hơn Đến ngày {eff_to}.')
    return eff_period, eff_from, eff_to


def _list_group_files(folder: Path, allowed_exts: set[str]) -> list[Path]:
    if not folder.is_dir():
        return []
    return sorted(
        (
            p for p in folder.rglob('*')
            if p.is_file() and not p.name.startswith('~$') and p.suffix.lower() in allowed_exts
        ),
        key=lambda p: str(p).lower(),
    )


def _discover_workspace_inputs(workspace: Path) -> dict[str, list[Path]]:
    """Collect grouped NIKE inputs without scanning legacy Tableau or Internal folders."""
    inputs_root = workspace / 'Inputs'
    hfpa_files = _list_group_files(inputs_root / 'hfpa_nike', SUPPORTED_TABULAR_EXTS)
    ftt_files = _list_group_files(inputs_root / 'ftt_nike', SUPPORTED_TABULAR_EXTS)
    catalog_files = _list_group_files(inputs_root / 'sandals_catalog', SUPPORTED_CATALOG_EXTS)

    # Also support flat files placed directly in workspace/Inputs/ cleanly by unambiguous name
    if inputs_root.is_dir():
        for item in sorted(inputs_root.iterdir(), key=lambda p: p.name.lower()):
            if not item.is_file() or item.name.startswith('~$'):
                continue
            low = item.name.lower()
            ext = item.suffix.lower()
            if ext == '.json' and ('sandal' in low or 'catalog' in low):
                catalog_files.append(item)
            elif ext in SUPPORTED_TABULAR_EXTS:
                if 'hfpa' in low:
                    hfpa_files.append(item)
                elif 'ftt' in low:
                    ftt_files.append(item)

    return {
        'hfpa_nike': hfpa_files,
        'ftt_nike': ftt_files,
        'sandals_catalog': catalog_files,
    }


def _ftt_group_of(rec: Any) -> str:
    meta = getattr(rec, 'metadata', None) or {}
    return str(getattr(rec, 'ucc_id', '') or meta.get('source_group') or meta.get('group') or '').strip()


def _load_catalog_inputs(
    catalog_inputs: list[Path | str | LineCatalogEntry | dict[str, Any]] | None,
    *,
    ftt_base_records: list[Any],
    fallback_cache_path: Path | None = None,
) -> dict[str, Any]:
    """Assemble Factory Sandals catalog from explicit inputs, live FTT Group rows, and isolated SFC cache."""
    entries: list[LineCatalogEntry | dict[str, Any]] = []

    for item in catalog_inputs or []:
        if isinstance(item, (LineCatalogEntry, dict)):
            entries.append(item)
            continue
        p = Path(item)
        if not p.is_file():
            continue
        if p.suffix.lower() == '.json':
            payload = json.loads(p.read_text(encoding='utf-8'))
            raw_list = payload.get('entries', payload) if isinstance(payload, dict) else payload
            if isinstance(raw_list, list):
                for entry in raw_list:
                    if isinstance(entry, dict):
                        entry.setdefault('provenance', f'catalog_json:{p.name}')
                        entries.append(entry)
        elif p.suffix.lower() in SUPPORTED_TABULAR_EXTS:
            parsed = normalize_nike_source([p], metric='FTT', run_id='catalog')
            for rec in parsed['base_records']:
                grp = _ftt_group_of(rec)
                if 'sandal' in grp.lower():
                    d = rec.to_dict()
                    d['source_group'] = grp
                    d['provenance'] = f'catalog_file:{p.name}'
                    entries.append(d)

    if not entries and fallback_cache_path and fallback_cache_path.is_file():
        payload = json.loads(fallback_cache_path.read_text(encoding='utf-8'))
        raw_list = payload.get('entries', payload) if isinstance(payload, dict) else payload
        if isinstance(raw_list, list):
            for entry in raw_list:
                if isinstance(entry, dict):
                    entry.setdefault('provenance', f'isolated_cache:{fallback_cache_path.name}')
                    entries.append(entry)

    ftt_rows = []
    for rec in ftt_base_records:
        grp = _ftt_group_of(rec)
        if 'sandal' in grp.lower():
            d = rec.to_dict()
            d['source_group'] = grp
            ftt_rows.append(d)
    return build_area_catalog(entries=entries, ftt_rows=ftt_rows)


def _overall_status(hfpa_result: dict[str, Any], ftt_result: dict[str, Any], export_ok: bool) -> str:
    codes = {hfpa_result['review']['status_code'], ftt_result['review']['status_code']}
    if not export_ok or 'export_error' in codes:
        return 'export_error'
    if codes == {'missing_source'}:
        return 'missing_source'
    if 'needs_mapping_review' in codes:
        return 'needs_mapping_review'
    if 'partial' in codes or 'missing_source' in codes:
        return 'partial'
    return 'exported' if hfpa_result['review']['status_code'] == 'ready_to_export' and ftt_result['review']['status_code'] == 'ready_to_export' else 'partial'


def _build_business_review(hfpa_result: dict[str, Any], ftt_result: dict[str, Any], catalog_count: int, office_meta: dict[str, Any] | None) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    for res in (hfpa_result, ftt_result):
        metric = res['metric']
        rev = res['review']
        status_code = rev['status_code']
        ui_status = 'error' if status_code in ('export_error', 'missing_source') else ('warning' if status_code != 'ready_to_export' else 'passed')
        if res['top_groups']:
            g_str = '; '.join(f"#{g['rank']} {g['group_label']} ({g['clg_qty']:,.0f})" for g in res['top_groups'])
            msg = f"{metric} NIKE [{rev['status_label']}]: CLG={res['scope_totals']['CLG']['total_defect_issues']:,.0f} lỗi. Top 3 nhóm: {g_str}."
        else:
            msg = f"{metric} NIKE [{rev['status_label']}]: chưa có dữ liệu hoặc chưa có nhóm lỗi dương."
        items.append({'rule': f'defect_review.{metric.lower()}', 'status': ui_status, 'message': msg})

    items.append({
        'rule': 'defect_review.sandals_catalog',
        'status': 'passed' if catalog_count > 0 else 'warning',
        'message': (
            f'Danh mục Factory Sandals (VB): {catalog_count} line chuẩn hóa (dùng phân bổ chung cho cả HFPA & FTT).'
            if catalog_count > 0
            else 'Chưa có danh mục Factory Sandals (VB); các line Factory Sandals chưa thể tách riêng khỏi VH.'
        ),
    })
    items.append({
        'rule': 'defect_review.dr_unconfirmed',
        'status': 'warning',
        'message': hfpa_result['dr_config']['note'],
    })
    if office_meta:
        items.append({
            'rule': 'defect_review.office_reopen',
            'status': 'passed' if office_meta.get('reopen_verified') else 'error',
            'message': (
                f"Đã kiểm tra mở lại PPTX qua {office_meta.get('verification_app')}: "
                f"{office_meta.get('slide_count')} slide, {office_meta.get('native_charts_pasted')} biểu đồ native, "
                f"external links hợp lệ ({len(office_meta.get('external_links', []))})."
            ),
        })
    return items


def run_quarterly_defect_review(
    *,
    hfpa_inputs: list[Path | str | dict[str, Any]] | None = None,
    ftt_inputs: list[Path | str | dict[str, Any]] | None = None,
    catalog_inputs: list[Path | str | LineCatalogEntry | dict[str, Any]] | None = None,
    output_dir: Path | str,
    period: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    cutoff: str | None = None,
    run_id: str = 'manual',
    validate_only: bool = False,
    skip_office: bool = False,
    mapping_overrides: dict[str, Any] | None = None,
    fallback_catalog_path: Path | None = DEFAULT_CATALOG_CACHE_PATH,
    dr_config: DRConfig | None = None,
    preview_dir: Path | str | None = None,
    stage: Callable[[str, str | None], None] | None = None,
) -> dict[str, Any]:
    """Execute independent HFPA NIKE & FTT NIKE pipelines and generate audited Excel + PPTX deliverables."""
    notify = stage or (lambda _msg, _phase=None: None)
    eff_period, eff_from, eff_to = resolve_period_and_dates(period, date_from, date_to, cutoff)
    out_dir = Path(output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    eff_preview_dir = Path(preview_dir).resolve() if preview_dir else (out_dir / 'previews')

    mapping = load_defect_mapping(
        overrides_path=DEFAULT_MAPPING_OVERRIDES_PATH if DEFAULT_MAPPING_OVERRIDES_PATH.is_file() else None,
        extra_overrides=mapping_overrides,
    )
    eff_dr = dr_config or DRConfig()

    notify('Đang chuẩn hóa dữ liệu HFPA NIKE (độc lập)...', 'hfpa_parse')
    hfpa_norm = normalize_nike_source(
        list(hfpa_inputs or []),
        metric='HFPA',
        run_id=run_id,
        date_from=eff_from,
        date_to=eff_to,
        mapping=mapping,
    )

    notify('Đang chuẩn hóa dữ liệu FTT NIKE (độc lập)...', 'ftt_parse')
    ftt_norm = normalize_nike_source(
        list(ftt_inputs or []),
        metric='FTT',
        run_id=run_id,
        date_from=eff_from,
        date_to=eff_to,
        mapping=mapping,
    )

    notify('Đang xây dựng danh mục Factory Sandals (VB) & phân khu VH/JV/VB...', 'area_mapping')
    catalog = _load_catalog_inputs(
        catalog_inputs,
        ftt_base_records=ftt_norm['base_records'],
        fallback_cache_path=fallback_catalog_path,
    )
    extra_catalog_reviews: list[MappingReviewItem] = []
    if catalog['count'] == 0 and (hfpa_norm['base_records'] or ftt_norm['base_records']):
        extra_catalog_reviews.append(
            MappingReviewItem(
                review_type='missing_sandals_catalog',
                metric='ALL',
                factory='',
                plant='',
                line_std='',
                raw_label='Factory Sandals',
                raw_key='factory_sandals',
                suggested_or_confirmed_mapping='Cần nạp danh mục Factory Sandals từ FTT hoặc SFC',
                affected_quantity=0.0,
                status='needs_review',
                reason='Chưa có danh mục line Factory Sandals nên chưa thể tách riêng VB khỏi VH.',
                provenance='area_mapping',
                blocking=False,
            )
        )

    hfpa_base, hfpa_facts, hfpa_area_reviews = classify_areas(hfpa_norm['base_records'], hfpa_norm['defect_facts'], catalog)
    ftt_base, ftt_facts, ftt_area_reviews = classify_areas(ftt_norm['base_records'], ftt_norm['defect_facts'], catalog)

    notify('Đang phân tích 3 tầng lỗi & top model riêng cho HFPA NIKE và FTT NIKE...', 'analysis')
    hfpa_result = analyze_metric(
        metric='HFPA',
        base_records=hfpa_base,
        defect_facts=hfpa_facts,
        review_items=hfpa_norm['review_items'] + hfpa_area_reviews + extra_catalog_reviews,
        manifests=hfpa_norm['manifests'],
        detail_support_level=hfpa_norm['detail_support_level'],
        catalog_version=catalog['catalog_version'],
        mapping_version=hfpa_norm['mapping_version'],
        period=eff_period,
        date_from=eff_from,
        date_to=eff_to,
        dr_config=eff_dr,
        mapping=mapping,
    )
    ftt_result = analyze_metric(
        metric='FTT',
        base_records=ftt_base,
        defect_facts=ftt_facts,
        review_items=ftt_norm['review_items'] + ftt_area_reviews + extra_catalog_reviews,
        manifests=ftt_norm['manifests'],
        detail_support_level=ftt_norm['detail_support_level'],
        catalog_version=catalog['catalog_version'],
        mapping_version=ftt_norm['mapping_version'],
        period=eff_period,
        date_from=eff_from,
        date_to=eff_to,
        dr_config=eff_dr,
        mapping=mapping,
    )

    has_blocking = (
        hfpa_result['review']['has_blocking_issues']
        or ftt_result['review']['has_blocking_issues']
    )
    has_any_data = bool(hfpa_base or ftt_base)

    excel_meta: dict[str, Any] | None = None
    office_meta: dict[str, Any] | None = None
    export_ok = True
    export_error: str | None = None

    safe_period = eff_period.replace('/', '-').replace('\\', '-')
    excel_path = out_dir / f'Quarterly_NIKE_Defect_Review_{safe_period}.xlsx'
    pptx_path = out_dir / f'Quarterly_NIKE_Defect_Review_{safe_period}.pptx'

    if not validate_only and not has_blocking and has_any_data:
        try:
            notify('Đang tạo workbook Excel chuẩn hóa & biểu đồ nguồn...', 'excel')
            excel_meta = build_excel_workbook(
                excel_path,
                period=eff_period,
                date_from=eff_from,
                date_to=eff_to,
                hfpa_result=hfpa_result,
                ftt_result=ftt_result,
                hfpa_base=hfpa_base,
                hfpa_facts=hfpa_facts,
                ftt_base=ftt_base,
                ftt_facts=ftt_facts,
                catalog_entries=catalog['entries'],
            )
            if not skip_office:
                notify('Đang copy/paste biểu đồ sang PowerPoint & kiểm tra mở lại...', 'presentation')
                office_meta = build_office_presentation(
                    pptx_path,
                    final_excel_path=excel_path,
                    excel_meta=excel_meta,
                    period=eff_period,
                    date_from=eff_from,
                    date_to=eff_to,
                    hfpa_result=hfpa_result,
                    ftt_result=ftt_result,
                    preview_dir=eff_preview_dir,
                )
        except OfficeExportError as exc:
            export_ok = False
            export_error = str(exc)
    elif has_blocking:
        export_ok = False
        export_error = 'Phát hiện lỗi kiểm tra chặn xuất (sai đối chiếu control hoặc sai file nguồn NIKE). Không xuất Excel/PPTX.'

    status_code = _overall_status(hfpa_result, ftt_result, export_ok)
    business_review = _build_business_review(hfpa_result, ftt_result, catalog['count'], office_meta)
    if export_error:
        business_review.append({'rule': 'defect_review.export', 'status': 'error', 'message': export_error})

    summary_payload = {
        'tool': 'quarterly_defect_review',
        'run_id': run_id,
        'period': eff_period,
        'date_from': eff_from,
        'date_to': eff_to,
        'cutoff': cutoff or '',
        'generated_at': datetime.now().isoformat(timespec='seconds'),
        'status_code': status_code,
        'status_label': STATUS_LABELS.get(status_code, status_code),
        'export_ok': export_ok,
        'export_error': export_error,
        'catalog': {
            'catalog_version': catalog['catalog_version'],
            'count': catalog['count'],
            'entries': [e.to_dict() for e in catalog['entries']],
        },
        'mapping_version': mapping['mapping_version'],
        'dr_config': eff_dr.to_dict(),
        'hfpa': hfpa_result,
        'ftt': ftt_result,
        'excel_meta': excel_meta,
        'office_meta': office_meta,
        'files': {
            'excel': str(excel_path) if excel_meta else None,
            'pptx': str(pptx_path) if office_meta else None,
            'slide_previews': office_meta.get('slide_previews', []) if office_meta else [],
        },
        'business_review': business_review,
    }

    summary_path = out_dir / 'review_summary.json'
    summary_path.write_text(json.dumps(summary_payload, indent=2, ensure_ascii=False), encoding='utf-8')
    summary_payload['files']['review_summary'] = str(summary_path)

    return summary_payload


def run_report(
    workspace: Path | str,
    run: dict[str, Any] | None = None,
    stage: Callable[[str, str | None], None] | None = None,
) -> dict[str, Any]:
    """Worker adapter executed by report_tool_worker.py for quarterly_defect_review."""
    ws = Path(workspace).resolve()
    run_data = run if isinstance(run, dict) else {}
    fields = run_data.get('fields') if isinstance(run_data.get('fields'), dict) else {}

    period = run_data.get('period') or fields.get('period') or DEFAULT_PERIOD
    date_from = run_data.get('date_from') or fields.get('date_from')
    date_to = run_data.get('date_to') or fields.get('date_to')
    cutoff = run_data.get('cutoff') or fields.get('cutoff')
    mode = str(run_data.get('mode') or fields.get('mode') or 'all').strip().lower()
    mapping_overrides = run_data.get('mapping_overrides') or fields.get('mapping_overrides')

    discovered = _discover_workspace_inputs(ws)
    out_dir = ws / 'Output'
    result = run_quarterly_defect_review(
        hfpa_inputs=discovered['hfpa_nike'],
        ftt_inputs=discovered['ftt_nike'],
        catalog_inputs=discovered['sandals_catalog'],
        output_dir=out_dir,
        period=period,
        date_from=date_from,
        date_to=date_to,
        cutoff=cutoff,
        run_id=str(run_data.get('run_id') or ws.name),
        validate_only=(mode == 'check'),
        skip_office=(mode in ('check', 'excel')),
        mapping_overrides=mapping_overrides,
        stage=stage,
    )

    run_data['business_review'] = result['business_review']
    run_data['summary'] = {
        'period': result['period'],
        'date_from': result['date_from'],
        'date_to': result['date_to'],
        'status_code': result['status_code'],
        'status_label': result['status_label'],
        'hfpa_defects_clg': result['hfpa']['scope_totals']['CLG']['total_defect_issues'],
        'ftt_defects_clg': result['ftt']['scope_totals']['CLG']['total_defect_issues'],
        'verification_app': (result.get('office_meta') or {}).get('verification_app'),
        'review_summary_file': 'review_summary.json',
    }
    if not result['export_ok'] and result.get('export_error'):
        raise RuntimeError(result['export_error'])

    return result
