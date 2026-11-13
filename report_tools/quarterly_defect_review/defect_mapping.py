"""Three-tier NIKE defect taxonomy mapping and unmapped label review."""

from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path
from typing import Any

from report_tools.quarterly_defect_review.contracts import DefectFact, MappingReviewItem

DEFAULT_MAPPING_PATH = Path(__file__).with_name('defect_mapping.json')


def clean_label_key(text: Any) -> str:
    """Normalize invisible unicode characters and whitespace strictly without fuzzy matching."""
    cleaned = str(text or '')
    cleaned = cleaned.replace('\u200b', '').replace('\ufeff', '').replace('\xa0', ' ')
    cleaned = cleaned.replace('\r', ' ').replace('\n', ' ')
    cleaned = re.sub(r'\s*/\s*', '/', cleaned)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip().lower()
    return cleaned


def _detail_slug(text: str) -> str:
    slug = re.sub(r'[^a-z0-9]+', '_', clean_label_key(text)).strip('_')
    if not slug:
        return 'item'
    return slug


def load_defect_mapping(
    config_path: Path | str | None = None,
    overrides_path: Path | str | None = None,
    extra_overrides: dict[str, Any] | list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Load the official 11-group / 25-subtype NIKE defect mapping plus optional user overrides."""
    path = Path(config_path) if config_path else DEFAULT_MAPPING_PATH
    raw = json.loads(path.read_text(encoding='utf-8'))

    groups_by_id: dict[str, dict[str, Any]] = {}
    group_lookup: dict[str, dict[str, Any]] = {}
    for entry in raw['standard_groups']:
        groups_by_id[entry['id']] = dict(entry)
        candidates = [entry['id'], entry['label'], entry.get('short_label', '')] + list(entry.get('hfpa_aliases', []))
        for candidate in candidates:
            key = clean_label_key(candidate)
            if key:
                group_lookup[key] = groups_by_id[entry['id']]

    subtypes_by_id: dict[str, dict[str, Any]] = {}
    subtype_lookup: dict[str, dict[str, Any]] = {}
    for entry in raw['ftt_subtypes']:
        subtypes_by_id[entry['id']] = dict(entry)
        candidates = [entry['id'], entry['label']] + list(entry.get('aliases', []))
        for candidate in candidates:
            key = clean_label_key(candidate)
            if key:
                subtype_lookup[key] = subtypes_by_id[entry['id']]

    # Confirmed detail/group overrides (never fuzzy, keyed strictly by (source_metric, raw_key))
    custom_rules: dict[tuple[str, str], dict[str, Any]] = {}
    override_payloads: list[Any] = []
    if overrides_path and Path(overrides_path).is_file():
        override_payloads.append(json.loads(Path(overrides_path).read_text(encoding='utf-8')))
    if extra_overrides:
        override_payloads.append(extra_overrides)

    for payload in override_payloads:
        rules = payload.get('rules', payload) if isinstance(payload, dict) else payload
        if isinstance(rules, dict):
            rules = list(rules.values())
        for rule in rules or []:
            if not isinstance(rule, dict):
                continue
            metric = str(rule.get('source_metric') or '*').strip().upper()
            raw_key = clean_label_key(rule.get('raw_key') or rule.get('raw_label'))
            if not raw_key:
                continue
            custom_rules[(metric, raw_key)] = dict(rule)

    version_hash = hashlib.sha256(
        json.dumps({'base': raw, 'overrides': sorted(custom_rules.items())}, sort_keys=True, ensure_ascii=False).encode('utf-8')
    ).hexdigest()[:10]
    mapping_version = f"{raw['mapping_version']}+{version_hash}" if custom_rules else raw['mapping_version']

    return {
        'mapping_version': mapping_version,
        'confirmed_by': raw.get('confirmed_by'),
        'confirmed_at': raw.get('confirmed_at'),
        'groups_by_id': groups_by_id,
        'group_lookup': group_lookup,
        'subtypes_by_id': subtypes_by_id,
        'subtype_lookup': subtype_lookup,
        'custom_rules': custom_rules,
    }


def _resolve_group(raw_group: str | None, mapping: dict[str, Any]) -> dict[str, Any] | None:
    if not raw_group:
        return None
    key = clean_label_key(raw_group)
    return mapping['group_lookup'].get(key)


def _resolve_subtype(raw_subtype: str | None, mapping: dict[str, Any]) -> dict[str, Any] | None:
    if not raw_subtype:
        return None
    key = clean_label_key(raw_subtype)
    return mapping['subtype_lookup'].get(key)


def _resolve_custom(metric: str, raw_label: str, mapping: dict[str, Any]) -> dict[str, Any] | None:
    key = clean_label_key(raw_label)
    if not key:
        return None
    return mapping['custom_rules'].get((metric.upper(), key)) or mapping['custom_rules'].get(('*', key))


def map_defects(
    raw_facts: list[DefectFact | dict[str, Any]],
    mapping: dict[str, Any] | None = None,
) -> tuple[list[DefectFact], list[MappingReviewItem]]:
    """Map raw extracted defect facts onto the 3-tier taxonomy without silent fallback to Other."""
    cfg = mapping or load_defect_mapping()
    version = cfg['mapping_version']
    groups_by_id = cfg['groups_by_id']
    mapped_facts: list[DefectFact] = []
    unmapped_agg: dict[tuple[str, str, str], dict[str, Any]] = {}

    for item in raw_facts:
        fact = item if isinstance(item, DefectFact) else DefectFact(**item)
        fact.mapping_version = version

        # Already marked missing_detail or unreconciled_gap upstream retains explicit status.
        if fact.detail_status in ('no_defect', 'missing_detail', 'unreconciled_gap') or fact.is_synthetic_gap:
            group_info = groups_by_id.get(fact.standard_group_id or '') or _resolve_group(fact.standard_group_raw, cfg)
            if group_info:
                fact.standard_group_id = group_info['id']
                fact.standard_group_label = group_info['label']
            mapped_facts.append(fact)
            continue

        group_info = (
            groups_by_id.get(fact.standard_group_id or '')
            or _resolve_group(fact.standard_group_id, cfg)
            or _resolve_group(fact.standard_group_raw, cfg)
        )
        subtype_info = (
            cfg['subtypes_by_id'].get(fact.nike_subtype_id or '')
            or _resolve_subtype(fact.nike_subtype_id, cfg)
            or _resolve_subtype(fact.nike_subtype_raw, cfg)
        )
        custom_rule = (
            _resolve_custom(fact.metric, fact.detail_raw, cfg)
            or _resolve_custom(fact.metric, fact.nike_subtype_raw, cfg)
            or _resolve_custom(fact.metric, fact.standard_group_raw, cfg)
        )

        # Custom rule applies only if confirmed explicitly
        if custom_rule:
            target_group_id = custom_rule.get('standard_group_id')
            if target_group_id and target_group_id in groups_by_id:
                group_info = groups_by_id[target_group_id]
            target_sub_id = custom_rule.get('nike_subtype_id')
            if target_sub_id and target_sub_id in cfg['subtypes_by_id']:
                subtype_info = cfg['subtypes_by_id'][target_sub_id]
            if custom_rule.get('detail_id'):
                fact.detail_id = str(custom_rule['detail_id'])
            if custom_rule.get('detail_label'):
                fact.detail_label = str(custom_rule['detail_label'])

        if subtype_info and not group_info:
            group_info = groups_by_id.get(subtype_info['standard_group_id'])

        if subtype_info:
            fact.nike_subtype_id = subtype_info['id']
            fact.nike_subtype_label = subtype_info['label']

        # Case 1: Could not resolve even a standard parent group -> UNMAPPED (never guess Other!)
        if not group_info:
            raw_label = fact.detail_raw or fact.nike_subtype_raw or fact.standard_group_raw or '(trống)'
            raw_key = clean_label_key(raw_label) or '__empty__'
            fact.standard_group_id = None
            fact.standard_group_label = None
            fact.detail_status = 'unmapped'
            fact.detail_id = f"{fact.metric.lower()}__unmapped__{_detail_slug(raw_key)}"
            fact.detail_label = raw_label
            fact.mapping_basis = 'Chưa ánh xạ - không gộp ngầm vào Other/Miscellaneous'
            bucket = unmapped_agg.setdefault(
                (fact.metric, raw_key, fact.source_column),
                {
                    'metric': fact.metric,
                    'factory': fact.factory,
                    'plant': fact.plant,
                    'line_std': fact.line_std,
                    'raw_label': raw_label,
                    'raw_key': f"{fact.metric}:{raw_key}",
                    'qty': 0.0,
                    'provenance': f"{fact.source_file}:{fact.source_sheet}:{fact.source_column}",
                },
            )
            bucket['qty'] += float(fact.issue_quantity or 0.0)
            mapped_facts.append(fact)
            continue

        fact.standard_group_id = group_info['id']
        fact.standard_group_label = group_info['label']

        # Case 2: Observed genuine detail label (Tier 3)
        if fact.detail_tier == 'tier3_detail' and (fact.detail_raw or fact.detail_label):
            raw_text = (fact.detail_raw or fact.detail_label or '').strip()
            fact.detail_raw = raw_text
            if not fact.detail_label:
                fact.detail_label = raw_text
            detail_key_part = clean_label_key(fact.detail_label)
            # Key ALWAYS includes source metric + parent group + detail key
            fact.detail_id = fact.detail_id or f"{fact.metric.lower()}__{group_info['id']}__{_detail_slug(detail_key_part)}"
            fact.detail_status = 'mapped'
            fact.mapping_basis = custom_rule.get('basis', 'Quan sát trực tiếp ở tầng lỗi chi tiết nguồn NIKE') if custom_rule else 'Quan sát trực tiếp ở tầng lỗi chi tiết nguồn NIKE'
        # Case 3: Tier 2 subtype only (e.g. real FTT NIKE 25 subtypes)
        elif fact.detail_tier == 'tier2_subtype_only':
            if subtype_info:
                fact.detail_raw = fact.detail_raw or fact.nike_subtype_raw or subtype_info['label']
                fact.detail_label = subtype_info['label']
                fact.detail_id = f"{fact.metric.lower()}__{group_info['id']}__{subtype_info['id']}"
                fact.detail_status = 'subtype_only'
                fact.mapping_basis = 'Phân loại theo cột FTT NIKE thực tế (nguồn không có cấp chi tiết thứ 3)'
            else:
                raw_label = fact.nike_subtype_raw or fact.detail_raw or fact.source_column or '(không xác định)'
                raw_key = clean_label_key(raw_label) or '__empty__'
                fact.detail_status = 'unmapped'
                fact.detail_id = f"{fact.metric.lower()}__{group_info['id']}__unmapped__{_detail_slug(raw_key)}"
                fact.detail_label = raw_label
                fact.mapping_basis = 'Chưa ánh xạ phân loại FTT NIKE'
                bucket = unmapped_agg.setdefault(
                    (fact.metric, f"{group_info['id']}:{raw_key}", fact.source_column),
                    {
                        'metric': fact.metric,
                        'factory': fact.factory,
                        'plant': fact.plant,
                        'line_std': fact.line_std,
                        'raw_label': raw_label,
                        'raw_key': f"{fact.metric}:{group_info['id']}:{raw_key}",
                        'qty': 0.0,
                        'provenance': f"{fact.source_file}:{fact.source_sheet}:{fact.source_column}",
                    },
                )
                bucket['qty'] += float(fact.issue_quantity or 0.0)
        # Case 4: Tier 1 HFPA parent group only (real SFC HFPA NIKE 28-col export)
        else:
            fact.detail_tier = 'tier1_group_only'
            fact.detail_id = None
            fact.detail_label = None
            fact.detail_status = 'group_only'
            fact.mapping_basis = 'Nguồn HFPA NIKE chỉ cung cấp tổng theo 11 nhóm chuẩn HFPA; thiếu chi tiết cấp sâu hơn'

        mapped_facts.append(fact)

    reviews: list[MappingReviewItem] = [
        MappingReviewItem(
            review_type='unmapped_defect',
            metric=entry['metric'],
            factory=entry['factory'],
            plant=entry['plant'],
            line_std=entry['line_std'],
            raw_label=entry['raw_label'],
            raw_key=entry['raw_key'],
            suggested_or_confirmed_mapping='Chưa ánh xạ (cần xác nhận nhóm chuẩn NIKE)',
            affected_quantity=round(entry['qty'], 4),
            status='Chưa ánh xạ',
            reason='Nhãn lỗi không khớp bảng quy đổi NIKE; giữ nguyên không gộp ngầm vào Other/Miscellaneous.',
            provenance=entry['provenance'],
            blocking=False,
        )
        for entry in sorted(unmapped_agg.values(), key=lambda item: (-item['qty'], item['raw_key']))
    ]
    return mapped_facts, reviews

