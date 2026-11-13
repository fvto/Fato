"""Factory Sandals catalog resolution and exclusive VH / JV / VB area classification."""

from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path
from typing import Any

from report_tools.quarterly_defect_review.contracts import (
    BaseRecord,
    DefectFact,
    LineCatalogEntry,
    MappingReviewItem,
)

VN_FACTORIES = {'VH', 'VH2'}
INDO_FACTORIES = {'JV', 'JV2'}


def normalize_factory_code(factory: Any, ucc_id: Any = '') -> str:
    """Normalize factory identifier without fabricating unknown codes."""
    raw = str(factory or '').strip().upper()
    if raw in {'VH', 'VH2', 'JV', 'JV2'}:
        return raw
    if raw in {'CL_VH', 'CL_VH1'}:
        return 'VH'
    if raw in {'CL_VH2'}:
        return 'VH2'
    if raw in {'JV1', 'CL_JV', 'CL_JV1'}:
        return 'JV'
    if raw in {'CL_JV2'}:
        return 'JV2'
    ucc = str(ucc_id or '').strip().upper()
    if raw in {'', 'CLG', 'CL_VB', 'VB'}:
        if ucc in {'CL_VH', 'CL_VB'}:
            return 'VH'
        if ucc == 'CL_VH2':
            return 'VH2'
        if ucc in {'CL_JV', 'CL_JV1'}:
            return 'JV'
        if ucc == 'CL_JV2':
            return 'JV2'
    return raw


def normalize_plant_code(plant: Any) -> str:
    """Canonical uppercase plant code with whitespace normalized."""
    cleaned = str(plant or '').replace('\u200b', '').replace('\xa0', ' ')
    return re.sub(r'\s+', ' ', cleaned).strip().upper()


def normalize_line_code(line_raw: Any) -> str:
    """Normalize Line1/Line 1/1 consistently while keeping suffixes (1B, 13, 01, PCC) intact."""
    cleaned = str(line_raw or '').replace('\u200b', '').replace('\xa0', ' ')
    cleaned = re.sub(r'\s+', ' ', cleaned).strip().upper()
    if not cleaned:
        return ''
    # Strip a single leading 'LINE' prefix (e.g. 'Line 1' -> '1', 'Line1B' -> '1B'),
    # but keep compound names like 'SR LINE 01' intact!
    match = re.fullmatch(r'LINE\s*[-_:]?\s*([0-9]+[A-Z]*|[A-Z0-9-]+)', cleaned)
    if match:
        token = match.group(1).strip()
        # Strip leading zeroes only on pure numeric tokens? Better: '01' -> '1' ONLY when pure digits
        # Wait: keep '01' vs '1' only if safe; '1' == '01' in numeric line numbers:
        if token.isdigit():
            return str(int(token))
        num_suffix = re.fullmatch(r'0*([1-9][0-9]*)([A-Z]+)', token)
        if num_suffix:
            return f"{num_suffix.group(1)}{num_suffix.group(2)}"
        return token
    if cleaned.isdigit():
        return str(int(cleaned))
    num_suffix = re.fullmatch(r'0*([1-9][0-9]*)([A-Z]+)', cleaned)
    if num_suffix:
        return f"{num_suffix.group(1)}{num_suffix.group(2)}"
    return cleaned


def split_combined_plant_line(text: Any) -> tuple[str, str]:
    """Split strings like 'N14 Line 1', 'N15 Line 10', or 'A Line 1B' into (plant, line_raw)."""
    cleaned = re.sub(r'\s+', ' ', str(text or '').replace('\u200b', '').replace('\xa0', ' ')).strip()
    if not cleaned:
        return '', ''
    match = re.match(r'^([A-Za-z0-9_-]+?)\s+(Line\s*.*)$', cleaned, flags=re.IGNORECASE)
    if match:
        return normalize_plant_code(match.group(1)), match.group(2).strip()
    parts = cleaned.split(' ', 1)
    if len(parts) == 2 and parts[0].upper() not in {'LINE', 'SR'}:
        return normalize_plant_code(parts[0]), parts[1].strip()
    return '', cleaned


def build_area_catalog(
    entries: list[LineCatalogEntry | dict[str, Any]] | None = None,
    ftt_rows: list[dict[str, Any]] | None = None,
    catalog_path: Path | str | None = None,
) -> dict[str, Any]:
    """Build verified Factory Sandals line catalog + deterministic catalog_version hash."""
    normalized: list[LineCatalogEntry] = []
    seen_keys: set[tuple[str, str, str, str, str, str]] = set()

    def _add(entry_obj: LineCatalogEntry) -> None:
        key = (
            entry_obj.factory,
            entry_obj.plant,
            entry_obj.line_std,
            entry_obj.source_group.lower(),
            entry_obj.effective_from,
            entry_obj.effective_to,
        )
        if key not in seen_keys and entry_obj.plant and entry_obj.line_std:
            seen_keys.add(key)
            normalized.append(entry_obj)

    sources: list[Any] = list(entries or [])
    if catalog_path and Path(catalog_path).is_file():
        path = Path(catalog_path)
        if path.suffix.lower() == '.json':
            payload = json.loads(path.read_text(encoding='utf-8'))
            sources.extend(payload.get('entries', payload) if isinstance(payload, dict) else payload)

    for raw in sources:
        if isinstance(raw, LineCatalogEntry):
            _add(raw)
            continue
        if not isinstance(raw, dict):
            continue
        group = str(raw.get('source_group') or raw.get('group') or 'Factory Sandals').strip()
        if 'sandal' not in group.lower():
            continue
        factory = normalize_factory_code(raw.get('factory') or 'VH')
        plant = normalize_plant_code(raw.get('plant'))
        line_raw = str(raw.get('line_raw') or raw.get('line') or '').strip()
        if (not plant or not line_raw) and raw.get('combined_line'):
            c_plant, c_line = split_combined_plant_line(raw['combined_line'])
            plant = plant or c_plant
            line_raw = line_raw or c_line
        elif not plant and line_raw:
            c_plant, c_line = split_combined_plant_line(line_raw)
            if c_plant:
                plant, line_raw = c_plant, c_line
        line_std = str(raw.get('line_std') or normalize_line_code(line_raw)).strip().upper()
        _add(
            LineCatalogEntry(
                factory=factory,
                plant=plant,
                line_raw=line_raw or line_std,
                line_std=line_std,
                source_group=group,
                report_area=str(raw.get('report_area') or 'VB').upper(),
                effective_from=str(raw.get('effective_from') or '1900-01-01')[:10],
                effective_to=str(raw.get('effective_to') or '2999-12-31')[:10],
                provenance=str(raw.get('provenance') or 'explicit_catalog'),
            )
        )

    for row in ftt_rows or []:
        group = str(row.get('source_group') or row.get('factory_group') or row.get('group') or '').strip()
        if 'sandal' not in group.lower():
            continue
        factory = normalize_factory_code(row.get('factory') or row.get('factory_code') or '')
        plant = normalize_plant_code(row.get('plant'))
        line_raw = str(row.get('line_raw') or row.get('line') or '').strip()
        if not plant and line_raw:
            plant, line_raw = split_combined_plant_line(line_raw)
        line_std = normalize_line_code(line_raw)
        if not factory or not plant or not line_std:
            continue
        _add(
            LineCatalogEntry(
                factory=factory,
                plant=plant,
                line_raw=line_raw,
                line_std=line_std,
                source_group=group,
                report_area='VB',
                effective_from=str(row.get('effective_from') or '1900-01-01')[:10],
                effective_to=str(row.get('effective_to') or '2999-12-31')[:10],
                provenance=str(row.get('provenance') or 'ftt_factory_sandals_export'),
            )
        )

    serialized = [entry.to_dict() for entry in sorted(normalized, key=lambda e: (e.factory, e.plant, e.line_std, e.effective_from))]
    version_hash = hashlib.sha256(json.dumps(serialized, sort_keys=True, ensure_ascii=False).encode('utf-8')).hexdigest()[:10]
    return {
        'catalog_version': f'sandals-{len(normalized)}-{version_hash}',
        'entries': normalized,
        'count': len(normalized),
    }


def _match_catalog(
    record: BaseRecord,
    entries: list[LineCatalogEntry],
) -> list[LineCatalogEntry]:
    """Match on (factory, plant, line_std) + effective date without cross-plant collision."""
    rec_factory = normalize_factory_code(record.factory, record.ucc_id)
    rec_plant = normalize_plant_code(record.plant)
    rec_line = normalize_line_code(record.line_std or record.line_raw)
    rec_date = (record.date or '')[:10]
    if not rec_factory or not rec_plant or not rec_line:
        return []

    matches: list[LineCatalogEntry] = []
    for entry in entries:
        if entry.plant != rec_plant or entry.line_std != rec_line:
            continue
        # Factory check: JV must never match a VH entry unless the entry itself claims JV;
        # VH/VH2 entries match strict factory if specified, otherwise allowed across VH* if identical plant+line?
        # Requirement says match by factory + plant + line! So require exact factory equality
        # when catalog entry declares factory cleanly.
        if entry.factory and entry.factory != rec_factory:
            continue
        if rec_date and (rec_date < entry.effective_from or rec_date > entry.effective_to):
            continue
        matches.append(entry)
    return matches


def classify_areas(
    base_records: list[BaseRecord],
    defect_facts: list[DefectFact],
    catalog: dict[str, Any],
) -> tuple[list[BaseRecord], list[DefectFact], list[MappingReviewItem]]:
    """Assign each BaseRecord and its DefectFacts exclusively to VH, JV, VB, or review bucket."""
    entries: list[LineCatalogEntry] = list(catalog.get('entries') or [])
    review_buckets: dict[tuple[str, str, str, str, str], dict[str, Any]] = {}
    record_area_map: dict[str, tuple[str, str]] = {}

    def _add_review(
        review_type: str,
        rec: BaseRecord,
        reason: str,
        status: str,
        suggested: str,
        blocking: bool = False,
    ) -> None:
        key = (review_type, rec.metric, rec.factory, rec.plant, rec.line_std or rec.line_raw)
        bucket = review_buckets.setdefault(
            key,
            {
                'review_type': review_type,
                'metric': rec.metric,
                'factory': rec.factory,
                'plant': rec.plant,
                'line_std': rec.line_std or rec.line_raw,
                'raw_label': f"{rec.factory}/{rec.plant}/{rec.line_raw}",
                'raw_key': f"{rec.factory}|{rec.plant}|{rec.line_std}",
                'suggested': suggested,
                'qty': 0.0,
                'status': status,
                'reason': reason,
                'provenance': rec.source_file,
                'blocking': blocking,
            },
        )
        bucket['qty'] += float(rec.defect_qty_control or 0.0)

    for rec in base_records:
        rec.factory = normalize_factory_code(rec.factory, rec.ucc_id)
        rec.plant = normalize_plant_code(rec.plant)
        rec.line_std = normalize_line_code(rec.line_std or rec.line_raw)

        if not rec.factory or not rec.plant or not rec.line_std:
            rec.report_area = 'UNMATCHED'
            rec.area_status = 'missing_key'
            rec.area_provenance = 'Thiếu factory/plant/line để phân khu an toàn'
            _add_review(
                'missing_line_key',
                rec,
                'Bản ghi thiếu khóa nhà máy/xưởng/chuyền; không tự suy đoán khu.',
                'Thiếu khóa chuyền',
                'Kiểm tra nguồn xuất SFC',
                blocking=True,
            )
            record_area_map[rec.source_record_id] = (rec.report_area, rec.area_status)
            continue

        matches = _match_catalog(rec, entries)
        distinct_areas = {m.report_area for m in matches}
        if len(distinct_areas) > 1 or len(matches) > 1:
            rec.report_area = 'CONFLICT'
            rec.area_status = 'multi_match'
            rec.area_provenance = f'Khớp {len(matches)} dòng danh mục hiệu lực chồng chéo'
            _add_review(
                'multi_match_line',
                rec,
                'Một bản ghi khớp nhiều quy tắc danh mục hiệu lực cùng lúc; đưa vào review để tránh đếm đôi.',
                'Khớp nhiều khu',
                'Kiểm tra khoảng ngày hiệu lực catalog',
                blocking=True,
            )
            record_area_map[rec.source_record_id] = (rec.report_area, rec.area_status)
            continue

        if rec.factory in INDO_FACTORIES:
            if matches:
                rec.report_area = 'CONFLICT'
                rec.area_status = 'jv_sandals_conflict'
                rec.area_provenance = f'JV/JV2 trùng danh mục Factory Sandals ({matches[0].provenance})'
                _add_review(
                    'jv_sandals_conflict',
                    rec,
                    'Chuyền thuộc JV/JV2 nhưng đồng thời nằm trong Group Factory Sandals của FTT; cần xác nhận trước khi tính.',
                    'Xung đột JV vs VB',
                    'Xác nhận giữ JV hay chuyển VB',
                    blocking=True,
                )
            else:
                rec.report_area = 'JV'
                rec.area_status = 'ok'
                rec.area_provenance = f'Nhà máy {rec.factory} -> JV (JV + JV2)'
        elif rec.factory in VN_FACTORIES:
            if matches:
                rec.report_area = 'VB'
                rec.area_status = 'ok'
                rec.area_provenance = f'Khớp Factory Sandals ({matches[0].provenance}) -> VB'
            else:
                rec.report_area = 'VH'
                rec.area_status = 'ok'
                rec.area_provenance = f'Nhà máy {rec.factory} ngoài Factory Sandals -> VH'
                if rec.ucc_id.strip().upper() == 'CL_VB':
                    _add_review(
                        'hfpa_uccid_catalog_mismatch',
                        rec,
                        'HFPA ghi UccId CL_VB nhưng chuyền chưa có trong danh mục Group Factory Sandals của FTT; tạm giữ ở VH theo quy tắc lấy danh mục FTT làm chuẩn.',
                        'Chờ đối chiếu danh mục VB',
                        'Bổ sung chuyền vào danh mục Factory Sandals nếu hợp lệ',
                        blocking=False,
                    )
        else:
            rec.report_area = 'UNMATCHED'
            rec.area_status = 'unknown_factory'
            rec.area_provenance = f'Nhà máy không thuộc VH/VH2/JV/JV2: {rec.factory!r}'
            _add_review(
                'unknown_factory',
                rec,
                f'Mã nhà máy {rec.factory!r} không thuộc VH/VH2/JV/JV2.',
                'Chưa khớp nhà máy',
                'Kiểm tra lại file nguồn',
                blocking=True,
            )

        record_area_map[rec.source_record_id] = (rec.report_area, rec.area_status)

    for fact in defect_facts:
        area, _ = record_area_map.get(fact.base_record_id, ('UNMATCHED', 'missing'))
        fact.report_area = area

    review_items = [
        MappingReviewItem(
            review_type=v['review_type'],
            metric=v['metric'],
            factory=v['factory'],
            plant=v['plant'],
            line_std=v['line_std'],
            raw_label=v['raw_label'],
            raw_key=v['raw_key'],
            suggested_or_confirmed_mapping=v['suggested'],
            affected_quantity=round(v['qty'], 4),
            status=v['status'],
            reason=v['reason'],
            provenance=v['provenance'],
            blocking=v['blocking'],
        )
        for v in sorted(review_buckets.values(), key=lambda x: (-x['qty'], x['metric'], x['raw_key']))
    ]
    return base_records, defect_facts, review_items

