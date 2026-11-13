"""Pure aggregation, ranking, coverage, and drill-down analysis engine (HFPA & FTT isolated)."""

from __future__ import annotations

from typing import Any

from report_tools.quarterly_defect_review.contracts import (
    AREAS,
    COVERAGE_WARNING_THRESHOLD_PCT,
    DEFAULT_DATE_RANGE,
    DEFAULT_PERIOD,
    SCOPES,
    STATUS_LABELS,
    BaseRecord,
    DefectFact,
    DRConfig,
    MappingReviewItem,
    SourceManifest,
)
from report_tools.quarterly_defect_review.defect_mapping import load_defect_mapping


def pct(numerator: float, denominator: float) -> float | None:
    """Compute exact percentage with zero-denominator returning None (rendered as N/A)."""
    if denominator is None or denominator <= 0:
        return None
    return (float(numerator) / float(denominator)) * 100.0


def fmt_pct(value: float | None, decimals: int = 1) -> str:
    """Format percentage at full internal precision rounded only on display."""
    if value is None:
        return 'N/A'
    return f'{value:.{decimals}f}%'


def _empty_scope_summary() -> dict[str, Any]:
    return {
        'record_count': 0,
        'inspected_or_produced_qty': 0.0,
        'produced_qty': 0.0,
        'defective_pairs_qty': 0.0,
        'a_grade_pairs': 0.0,
        'rework_pairs_qty': 0.0,
        'rework1_pairs_qty': 0.0,
        'b_grade_pairs': 0.0,
        'c_grade_pairs': 0.0,
        'total_defect_issues': 0.0,
        'mapped_defect_issues': 0.0,
        'unmapped_defect_issues': 0.0,
        'unreconciled_gap_issues': 0.0,
        'zero_defect_records': 0,
        'has_source_data': False,
        'scope_status': 'missing_data',
    }


def aggregate_defects(
    base_records: list[BaseRecord],
    defect_facts: list[DefectFact],
    review_items: list[MappingReviewItem] | None = None,
    *,
    metric: str,
    manifests: list[SourceManifest] | None = None,
    detail_support_level: str | None = None,
    catalog_version: str = '',
    mapping_version: str = '',
    period: str = DEFAULT_PERIOD,
    date_from: str = DEFAULT_DATE_RANGE[0],
    date_to: str = DEFAULT_DATE_RANGE[1],
    dr_config: DRConfig | None = None,
    mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Aggregate normalized BaseRecords and DefectFacts separately for one metric across VH/JV/VB/CLG."""
    metric = metric.upper()
    cfg = mapping or load_defect_mapping()
    groups_meta = cfg['groups_by_id']
    dr = dr_config or DRConfig()
    reviews = list(review_items or [])

    scope_totals: dict[str, dict[str, Any]] = {s: _empty_scope_summary() for s in SCOPES}
    excluded_totals = {'UNMATCHED': 0.0, 'CONFLICT': 0.0}

    # Base denominator aggregation (counted strictly ONCE per BaseRecord, never per DefectFact!)
    model_base_vol: dict[tuple[str, str], dict[str, float]] = {}
    for rec in base_records:
        area = rec.report_area
        if area not in AREAS:
            excluded_totals[area if area in excluded_totals else 'UNMATCHED'] += float(rec.defect_qty_control or 0.0)
            continue
        for s in (area, 'CLG'):
            st = scope_totals[s]
            st['record_count'] += 1
            st['inspected_or_produced_qty'] += float(rec.inspected_or_produced_qty or 0.0)
            st['produced_qty'] += float(rec.produced_qty or 0.0)
            st['defective_pairs_qty'] += float(rec.defective_pairs_qty or 0.0)
            st['a_grade_pairs'] += float(rec.a_grade_pairs or 0.0)
            st['rework_pairs_qty'] += float(rec.rework_pairs_qty or 0.0)
            st['rework1_pairs_qty'] += float(rec.rework1_pairs_qty or 0.0)
            st['b_grade_pairs'] += float(rec.b_grade_pairs or 0.0)
            st['c_grade_pairs'] += float(rec.c_grade_pairs or 0.0)
            st['total_defect_issues'] += float(rec.defect_qty_control or 0.0)
            if rec.has_zero_defects:
                st['zero_defect_records'] += 1
            st['has_source_data'] = True

            mv = model_base_vol.setdefault(
                (s, rec.model_label),
                {'inspected_or_produced_qty': 0.0, 'defective_pairs_qty': 0.0, 'record_count': 0.0},
            )
            mv['inspected_or_produced_qty'] += float(rec.inspected_or_produced_qty or 0.0)
            mv['defective_pairs_qty'] += float(rec.defective_pairs_qty or 0.0)
            mv['record_count'] += 1.0

    # Seed all 11 official HFPA groups cleanly
    group_totals: dict[tuple[str, str], float] = {(s, gid): 0.0 for s in SCOPES for gid in groups_meta}
    group_mapped_detail_totals: dict[tuple[str, str], float] = {(s, gid): 0.0 for s in SCOPES for gid in groups_meta}
    group_unreconciled_gap: dict[tuple[str, str], float] = {(s, gid): 0.0 for s in SCOPES for gid in groups_meta}
    group_unmapped_qty: dict[tuple[str, str], float] = {(s, gid): 0.0 for s in SCOPES for gid in groups_meta}

    # (scope, group_id, detail_id) -> payload
    detail_buckets: dict[tuple[str, str, str], dict[str, Any]] = {}
    # (scope, group_id, model_label) -> payload
    model_group_buckets: dict[tuple[str, str, str], dict[str, Any]] = {}

    observed_tiers: set[str] = set()
    multi_location_notes: list[str] = []

    for fact in defect_facts:
        if fact.metric.upper() != metric:
            continue
        area = fact.report_area
        if area not in AREAS:
            continue
        qty = float(fact.issue_quantity or 0.0)
        if fact.location_split_status == 'multi_location_unsplit' and fact.location_raw:
            note = f'{fact.detail_label or fact.detail_raw} ({fact.location_raw}): giữ nguyên {qty:g} không nhân theo vị trí'
            if note not in multi_location_notes:
                multi_location_notes.append(note)

        gid = fact.standard_group_id
        for s in (area, 'CLG'):
            st = scope_totals[s]
            if fact.detail_status == 'unmapped' or not gid:
                st['unmapped_defect_issues'] += qty
                if gid:
                    group_unmapped_qty[(s, gid)] += qty
                continue

            if fact.is_synthetic_gap or fact.detail_status == 'unreconciled_gap':
                group_totals[(s, gid)] += qty
                group_unreconciled_gap[(s, gid)] += qty
                st['unreconciled_gap_issues'] += qty
                mb = model_group_buckets.setdefault(
                    (s, gid, fact.model_label),
                    {
                        'model_key': fact.model_key,
                        'model_label': fact.model_label,
                        'qty': 0.0,
                        'lines': set(),
                        'fact_count': 0,
                    },
                )
                mb['qty'] += qty
                mb['lines'].add(f'{fact.plant} {fact.line_std}'.strip())
                mb['fact_count'] += 1
                continue

            group_totals[(s, gid)] += qty
            st['mapped_defect_issues'] += qty
            observed_tiers.add(fact.detail_tier)

            # Accumulate model within group (all defects in group, regardless of detail tier!)
            mb = model_group_buckets.setdefault(
                (s, gid, fact.model_label),
                {
                    'model_key': fact.model_key,
                    'model_label': fact.model_label,
                    'qty': 0.0,
                    'lines': set(),
                    'fact_count': 0,
                },
            )
            mb['qty'] += qty
            mb['lines'].add(f'{fact.plant} {fact.line_std}'.strip())
            mb['fact_count'] += 1

            # Accumulate detail defect ONLY when genuinely observed below parent group (tier3_detail or tier2_subtype_only)
            if fact.detail_status in ('mapped', 'subtype_only') and fact.detail_id and fact.detail_label:
                group_mapped_detail_totals[(s, gid)] += qty
                db = detail_buckets.setdefault(
                    (s, gid, fact.detail_id),
                    {
                        'detail_id': fact.detail_id,
                        'detail_label': fact.detail_label,
                        'detail_raw_examples': [],
                        'detail_status': fact.detail_status,
                        'detail_tier': fact.detail_tier,
                        'nike_subtype_id': fact.nike_subtype_id,
                        'nike_subtype_label': fact.nike_subtype_label,
                        'qty': 0.0,
                        'models': {},
                        'lines': set(),
                        'provenance_rows': [],
                    },
                )
                db['qty'] += qty
                if fact.detail_raw and fact.detail_raw not in db['detail_raw_examples']:
                    db['detail_raw_examples'].append(fact.detail_raw)
                db['models'][fact.model_label] = db['models'].get(fact.model_label, 0.0) + qty
                db['lines'].add(f'{fact.factory} {fact.plant} Line {fact.line_std}')
                if len(db['provenance_rows']) < 10:
                    db['provenance_rows'].append(
                        f'{fact.source_file}#{fact.source_sheet}:R{fact.source_row}({fact.source_column})'
                    )

    # Compute derived scope status and ensure scope total_defect_issues reflects fact totals
    # when base_records had 0 control total (or keep exact control).
    for s in SCOPES:
        st = scope_totals[s]
        if not st['has_source_data']:
            st['scope_status'] = 'missing_data'
        elif st['total_defect_issues'] == 0 and sum(group_totals[(s, g)] for g in groups_meta) == 0:
            st['scope_status'] = 'no_defect'
        else:
            st['scope_status'] = 'ok'
            fact_sum = sum(group_totals[(s, g)] for g in groups_meta) + st['unmapped_defect_issues']
            if st['total_defect_issues'] == 0 and fact_sum > 0:
                st['total_defect_issues'] = fact_sum

    if not detail_support_level:
        if 'tier3_detail' in observed_tiers:
            detail_support_level = 'tier3_detail'
        elif 'tier2_subtype_only' in observed_tiers:
            detail_support_level = 'tier2_subtype_only'
        else:
            detail_support_level = 'tier1_group_only'

    # Ensure unexplained gaps are accounted if group_totals exceeds mapped details + explicit gap facts
    for s in SCOPES:
        for gid in groups_meta:
            g_total = group_totals[(s, gid)]
            mapped_d = group_mapped_detail_totals[(s, gid)]
            explicit_gap = group_unreconciled_gap[(s, gid)]
            if g_total > mapped_d + explicit_gap + 1e-6:
                group_unreconciled_gap[(s, gid)] = g_total - mapped_d

    return {
        'metric': metric,
        'period': period,
        'date_from': date_from,
        'date_to': date_to,
        'catalog_version': catalog_version,
        'mapping_version': mapping_version or cfg['mapping_version'],
        'dr_config': dr.to_dict(),
        'detail_support_level': detail_support_level,
        'manifests': [m.to_dict() if hasattr(m, 'to_dict') else dict(m) for m in (manifests or [])],
        'groups_meta': groups_meta,
        'scope_totals': scope_totals,
        'excluded_totals': excluded_totals,
        'group_totals': group_totals,
        'group_mapped_detail_totals': group_mapped_detail_totals,
        'group_unreconciled_gap': group_unreconciled_gap,
        'group_unmapped_qty': group_unmapped_qty,
        'detail_buckets': detail_buckets,
        'model_group_buckets': model_group_buckets,
        'model_base_vol': model_base_vol,
        'review_items': reviews,
        'multi_location_notes': multi_location_notes,
    }


def rank_groups(aggregated: dict[str, Any], top_n: int = 3) -> dict[str, Any]:
    """Rank Top N parent defect groups by CLG total and compute VH/JV/VB contributions."""
    groups_meta = aggregated['groups_meta']
    group_totals = aggregated['group_totals']

    ranked_all: list[dict[str, Any]] = []
    for gid, meta in groups_meta.items():
        vh = group_totals[('VH', gid)]
        jv = group_totals[('JV', gid)]
        vb = group_totals[('VB', gid)]
        clg = vh + jv + vb
        if abs(clg - group_totals[('CLG', gid)]) > 1e-6:
            raise AssertionError(f'CLG != VH+JV+VB cho nhóm {gid}: {group_totals[("CLG", gid)]} vs {clg}')
        ranked_all.append(
            {
                'group_id': gid,
                'group_label': meta['label'],
                'short_label': meta.get('short_label', meta['label']),
                'color': meta.get('color', '1F77B4'),
                'clg_qty': clg,
                'vh_qty': vh,
                'jv_qty': jv,
                'vb_qty': vb,
                'vh_contrib_pct': pct(vh, clg),
                'jv_contrib_pct': pct(jv, clg),
                'vb_contrib_pct': pct(vb, clg),
            }
        )

    ranked_all.sort(key=lambda item: (-item['clg_qty'], item['group_id']))
    non_zero = [g for g in ranked_all if g['clg_qty'] > 0]
    top_groups = non_zero[:top_n]
    for pos, g in enumerate(top_groups, 1):
        g['rank'] = pos

    tie_at_cutoff = len(non_zero) > top_n and top_groups and abs(non_zero[top_n]['clg_qty'] - top_groups[-1]['clg_qty']) < 1e-9

    # Track per-area top 1 group to flag when an area's #1 differs from CLG's #1
    area_top1_notes: list[str] = []
    clg_top1_id = top_groups[0]['group_id'] if top_groups else None
    for area in AREAS:
        if not aggregated['scope_totals'][area]['has_source_data']:
            continue
        area_sorted = sorted(
            (
                {'group_id': gid, 'group_label': groups_meta[gid]['label'], 'qty': group_totals[(area, gid)]}
                for gid in groups_meta
            ),
            key=lambda x: (-x['qty'], x['group_id']),
        )
        if area_sorted and area_sorted[0]['qty'] > 0 and clg_top1_id and area_sorted[0]['group_id'] != clg_top1_id:
            area_top1_notes.append(
                f"Nhóm lỗi lớn nhất tại {area} là {area_sorted[0]['group_label']} ({area_sorted[0]['qty']:,.0f}), "
                f"khác với nhóm đứng đầu CLG ({top_groups[0]['group_label']})."
            )

    return {
        'all_groups': ranked_all,
        'top_groups': top_groups,
        'tie_at_cutoff': tie_at_cutoff,
        'area_top1_notes': area_top1_notes,
    }


def rank_details(aggregated: dict[str, Any], top_groups: list[dict[str, Any]], top_n: int = 3) -> dict[str, Any]:
    """Within each Top group and each scope independently, rank Top <=N actual detail defects."""
    detail_buckets = aggregated['detail_buckets']
    group_totals = aggregated['group_totals']
    group_mapped_totals = aggregated['group_mapped_detail_totals']
    group_gaps = aggregated['group_unreconciled_gap']
    group_unmapped = aggregated['group_unmapped_qty']
    scope_totals = aggregated['scope_totals']
    support_level = aggregated['detail_support_level']

    by_group: dict[str, dict[str, Any]] = {}
    for g in top_groups:
        gid = g['group_id']
        scope_cells: dict[str, Any] = {}
        clg_detail_map: dict[str, float] = {
            key[2]: val['qty']
            for key, val in detail_buckets.items()
            if key[0] == 'CLG' and key[1] == gid
        }

        for scope in SCOPES:
            g_total = group_totals[(scope, gid)]
            mapped_total = group_mapped_totals[(scope, gid)]
            gap_qty = group_gaps[(scope, gid)]
            unmapped_qty = group_unmapped[(scope, gid)]
            has_data = scope_totals[scope]['has_source_data']

            candidates = [
                {
                    'detail_id': val['detail_id'],
                    'detail_label': val['detail_label'],
                    'detail_tier': val['detail_tier'],
                    'detail_status': val['detail_status'],
                    'nike_subtype_label': val['nike_subtype_label'],
                    'qty': val['qty'],
                    'share_of_group_in_scope_pct': pct(val['qty'], g_total),
                    'area_contribution_to_detail_pct': pct(val['qty'], clg_detail_map.get(val['detail_id'], 0.0))
                    if scope in AREAS
                    else (100.0 if val['qty'] > 0 else None),
                    'models': sorted(val['models'].items(), key=lambda kv: (-kv[1], kv[0]))[:3],
                }
                for key, val in detail_buckets.items()
                if key[0] == scope and key[1] == gid and val['qty'] > 0
            ]
            candidates.sort(key=lambda d: (-d['qty'], d['detail_id']))
            top_items = candidates[:top_n]
            for pos, item in enumerate(top_items, 1):
                item['rank'] = pos

            top_sum = sum(d['qty'] for d in top_items)
            other_mapped = max(0.0, mapped_total - top_sum)
            coverage_pct = pct(mapped_total, g_total) if g_total > 0 else (0.0 if not candidates else 100.0)
            tie_at_cutoff = (
                len(candidates) > top_n
                and top_items
                and abs(candidates[top_n]['qty'] - top_items[-1]['qty']) < 1e-9
            )

            if not has_data:
                cell_status = 'missing_scope_data'
                status_note = f'{scope} thiếu dữ liệu nguồn'
            elif g_total == 0:
                cell_status = 'no_defect'
                status_note = 'Không có lỗi'
            elif not candidates:
                cell_status = 'missing_detail'
                status_note = (
                    f'Thiếu dữ liệu chi tiết: nguồn chỉ có nhóm chuẩn {g["group_label"]} '
                    f'({g_total:,.0f}), chưa có tên lỗi chi tiết ở cấp sâu hơn'
                )
            elif gap_qty > 1e-6 and coverage_pct is not None and coverage_pct < 99.9:
                cell_status = 'partial_coverage'
                status_note = (
                    f'Top trong phần đã có chi tiết (độ bao phủ {fmt_pct(coverage_pct)}; '
                    f'chưa có chi tiết: {gap_qty:,.0f})'
                )
            elif support_level == 'tier2_subtype_only':
                cell_status = 'subtype_level_only'
                status_note = (
                    f'Xếp hạng theo phân loại FTT NIKE thực tế (độ bao phủ {fmt_pct(coverage_pct)}; '
                    'chưa có cấp chi tiết thứ 3)'
                )
            else:
                cell_status = 'ok'
                status_note = f'Đủ chi tiết thực tế (độ bao phủ {fmt_pct(coverage_pct)})'

            scope_cells[scope] = {
                'scope': scope,
                'group_id': gid,
                'group_label': g['group_label'],
                'group_total_qty': g_total,
                'mapped_detail_qty': mapped_total,
                'top_details': top_items,
                'all_details_count': len(candidates),
                'other_mapped_details_qty': other_mapped,
                'other_mapped_details_pct': pct(other_mapped, g_total),
                'unreconciled_gap_qty': gap_qty,
                'unreconciled_gap_pct': pct(gap_qty, g_total),
                'unmapped_in_group_qty': unmapped_qty,
                'detail_coverage_pct': coverage_pct,
                'tie_at_cutoff': tie_at_cutoff,
                'cell_status': cell_status,
                'status_note': status_note,
            }

        by_group[gid] = scope_cells
    return by_group


def rank_models(aggregated: dict[str, Any], top_groups: list[dict[str, Any]], top_n: int = 3) -> dict[str, Any]:
    """Within each Top group and scope independently, rank Top N models (never replacing details)."""
    model_buckets = aggregated['model_group_buckets']
    model_base_vol = aggregated['model_base_vol']
    group_totals = aggregated['group_totals']
    scope_totals = aggregated['scope_totals']

    by_group: dict[str, dict[str, Any]] = {}
    for g in top_groups:
        gid = g['group_id']
        scope_cells: dict[str, Any] = {}
        for scope in SCOPES:
            g_total = group_totals[(scope, gid)]
            has_data = scope_totals[scope]['has_source_data']
            candidates = []
            for (s, m_gid, _m_label), val in model_buckets.items():
                if s != scope or m_gid != gid or val['qty'] <= 0:
                    continue
                vol = model_base_vol.get((scope, val['model_label']), {})
                candidates.append(
                    {
                        'model_key': val['model_key'],
                        'model_label': val['model_label'],
                        'qty': val['qty'],
                        'share_of_group_in_scope_pct': pct(val['qty'], g_total),
                        'ref_inspected_or_produced_qty': vol.get('inspected_or_produced_qty', 0.0),
                        'ref_defective_pairs_qty': vol.get('defective_pairs_qty', 0.0),
                    }
                )
            candidates.sort(key=lambda m: (-m['qty'], m['model_label']))
            top_items = candidates[:top_n]
            for pos, item in enumerate(top_items, 1):
                item['rank'] = pos
            tie_at_cutoff = (
                len(candidates) > top_n
                and top_items
                and abs(candidates[top_n]['qty'] - top_items[-1]['qty']) < 1e-9
            )
            scope_cells[scope] = {
                'scope': scope,
                'group_id': gid,
                'group_total_qty': g_total,
                'has_source_data': has_data,
                'top_models': top_items,
                'all_models_count': len(candidates),
                'tie_at_cutoff': tie_at_cutoff,
                'cell_status': 'missing_scope_data' if not has_data else ('no_defect' if g_total == 0 else 'ok'),
            }
        by_group[gid] = scope_cells
    return by_group


def _build_drilldown(
    aggregated: dict[str, Any],
    top_groups: list[dict[str, Any]],
) -> dict[str, Any]:
    """Build explicit drill-down tables separating group-share % from area-contribution %."""
    detail_buckets = aggregated['detail_buckets']
    model_buckets = aggregated['model_group_buckets']
    group_totals = aggregated['group_totals']
    drilldown: dict[str, Any] = {}

    for g in top_groups:
        gid = g['group_id']
        clg_details = [
            val for (s, mgid, _did), val in detail_buckets.items() if s == 'CLG' and mgid == gid and val['qty'] > 0
        ]
        clg_details.sort(key=lambda d: (-d['qty'], d['detail_id']))

        detail_rows = []
        for idx, clg_d in enumerate(clg_details, 1):
            did = clg_d['detail_id']
            d_clg = clg_d['qty']
            d_vh = detail_buckets.get(('VH', gid, did), {}).get('qty', 0.0)
            d_jv = detail_buckets.get(('JV', gid, did), {}).get('qty', 0.0)
            d_vb = detail_buckets.get(('VB', gid, did), {}).get('qty', 0.0)
            detail_rows.append(
                {
                    'rank_clg': idx,
                    'detail_id': did,
                    'detail_label': clg_d['detail_label'],
                    'nike_subtype_label': clg_d['nike_subtype_label'] or '',
                    'detail_tier': clg_d['detail_tier'],
                    'clg_qty': d_clg,
                    'vh_qty': d_vh,
                    'jv_qty': d_jv,
                    'vb_qty': d_vb,
                    # Denominator 1: G(metric, CLG, group)
                    'share_of_clg_group_pct': pct(d_clg, group_totals[('CLG', gid)]),
                    # Denominator 2: D(metric, CLG, group, detail) -> NEVER mixed with Denominator 1!
                    'vh_contrib_to_detail_pct': pct(d_vh, d_clg),
                    'jv_contrib_to_detail_pct': pct(d_jv, d_clg),
                    'vb_contrib_to_detail_pct': pct(d_vb, d_clg),
                    'top_models_for_detail': sorted(clg_d['models'].items(), key=lambda kv: (-kv[1], kv[0]))[:5],
                    'lines_count': len(clg_d['lines']),
                    'provenance_samples': clg_d['provenance_rows'],
                }
            )

        all_models_clg = sorted(
            (val for (s, mgid, _ml), val in model_buckets.items() if s == 'CLG' and mgid == gid and val['qty'] > 0),
            key=lambda m: (-m['qty'], m['model_label']),
        )
        drilldown[gid] = {
            'group_id': gid,
            'group_label': g['group_label'],
            'group_total_clg': group_totals[('CLG', gid)],
            'gap_clg_qty': aggregated['group_unreconciled_gap'][('CLG', gid)],
            'details': detail_rows,
            'models_clg': [
                {
                    'model_label': m['model_label'],
                    'qty': m['qty'],
                    'share_of_clg_group_pct': pct(m['qty'], group_totals[('CLG', gid)]),
                    'lines': sorted(m['lines'])[:10],
                }
                for m in all_models_clg[:15]
            ],
        }
    return drilldown


def build_review(
    aggregated: dict[str, Any],
    group_ranking: dict[str, Any],
    detail_ranking: dict[str, Any],
) -> dict[str, Any]:
    """Compile explicit run status, completeness, coverage, DR unconfirmed notice, and review items."""
    reviews: list[MappingReviewItem] = list(aggregated['review_items'])
    scope_totals = aggregated['scope_totals']
    missing_scopes = [s for s in AREAS if not scope_totals[s]['has_source_data']]
    blocking_reviews = [r for r in reviews if r.blocking]
    unmapped_reviews = [r for r in reviews if r.review_type == 'unmapped_defect']

    support_level = aggregated['detail_support_level']
    low_coverage_cells = [
        f"{g['group_label']} ({scope}: {fmt_pct(cell['detail_coverage_pct'])})"
        for g in group_ranking['top_groups']
        for scope, cell in detail_ranking.get(g['group_id'], {}).items()
        if cell['cell_status'] in ('missing_detail', 'partial_coverage')
    ]

    if not scope_totals['CLG']['has_source_data']:
        status_code = 'missing_source'
    elif blocking_reviews:
        status_code = 'export_error'
    elif unmapped_reviews:
        status_code = 'needs_mapping_review'
    elif missing_scopes or support_level != 'tier3_detail' or low_coverage_cells:
        status_code = 'partial'
    else:
        status_code = 'ready_to_export'

    factual_notes: list[str] = []
    for g in group_ranking['top_groups']:
        gid = g['group_id']
        contribs = [(a, g[f'{a.lower()}_contrib_pct'] or 0.0, g[f'{a.lower()}_qty']) for a in AREAS]
        top_area = max(contribs, key=lambda item: item[1])
        clg_cell = detail_ranking.get(gid, {}).get('CLG', {})
        top_d = clg_cell.get('top_details', [])
        top_d_text = f"lỗi dẫn đầu là {top_d[0]['detail_label']} ({top_d[0]['qty']:,.0f})" if top_d else 'chưa có dữ liệu lỗi chi tiết cấp sâu hơn'
        factual_notes.append(
            f"Nhóm #{g['rank']} {g['group_label']}: tổng CLG = {g['clg_qty']:,.0f}, "
            f"khu đóng góp cao nhất là {top_area[0]} ({top_area[2]:,.0f}, {top_area[1]:.1f}%), {top_d_text}."
        )
    factual_notes.extend(group_ranking.get('area_top1_notes', []))
    if low_coverage_cells:
        factual_notes.append('Lưu ý độ chi tiết / bao phủ: ' + '; '.join(low_coverage_cells[:6]) + '.')

    return {
        'status_code': status_code,
        'status_label': STATUS_LABELS[status_code],
        'can_export_full_details': status_code == 'ready_to_export',
        'can_export_with_limitations': status_code in ('ready_to_export', 'partial', 'needs_mapping_review'),
        'has_blocking_issues': bool(blocking_reviews),
        'missing_scopes': missing_scopes,
        'excluded_unclassified_qty': aggregated['excluded_totals'],
        'detail_support_level': support_level,
        'low_coverage_cells': low_coverage_cells,
        'tie_at_group_cutoff': group_ranking['tie_at_cutoff'],
        'dr_config': aggregated['dr_config'],
        'multi_location_notes': aggregated.get('multi_location_notes', []),
        'factual_notes': factual_notes,
        'items': [r.to_dict() if hasattr(r, 'to_dict') else dict(r) for r in reviews],
    }


def analyze_metric(
    *,
    metric: str,
    base_records: list[BaseRecord],
    defect_facts: list[DefectFact],
    review_items: list[MappingReviewItem] | None = None,
    manifests: list[SourceManifest] | None = None,
    detail_support_level: str | None = None,
    catalog_version: str = '',
    mapping_version: str = '',
    period: str = DEFAULT_PERIOD,
    date_from: str = DEFAULT_DATE_RANGE[0],
    date_to: str = DEFAULT_DATE_RANGE[1],
    dr_config: DRConfig | None = None,
    mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Run complete independent analysis pipeline for one metric (HFPA or FTT)."""
    aggregated = aggregate_defects(
        base_records,
        defect_facts,
        review_items,
        metric=metric,
        manifests=manifests,
        detail_support_level=detail_support_level,
        catalog_version=catalog_version,
        mapping_version=mapping_version,
        period=period,
        date_from=date_from,
        date_to=date_to,
        dr_config=dr_config,
        mapping=mapping,
    )
    groups = rank_groups(aggregated, top_n=3)
    details = rank_details(aggregated, groups['top_groups'], top_n=3)
    models = rank_models(aggregated, groups['top_groups'], top_n=3)
    drilldown = _build_drilldown(aggregated, groups['top_groups'])
    review = build_review(aggregated, groups, details)

    return {
        'metric': aggregated['metric'],
        'period': aggregated['period'],
        'date_from': aggregated['date_from'],
        'date_to': aggregated['date_to'],
        'catalog_version': aggregated['catalog_version'],
        'mapping_version': aggregated['mapping_version'],
        'detail_support_level': aggregated['detail_support_level'],
        'dr_config': aggregated['dr_config'],
        'manifests': aggregated['manifests'],
        'scope_totals': aggregated['scope_totals'],
        'all_groups': groups['all_groups'],
        'top_groups': groups['top_groups'],
        'tie_at_group_cutoff': groups['tie_at_cutoff'],
        'area_top1_notes': groups['area_top1_notes'],
        'details_by_group': details,
        'models_by_group': models,
        'drilldown_by_group': drilldown,
        'review': review,
    }
