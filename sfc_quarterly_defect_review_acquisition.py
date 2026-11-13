"""Isolated SFC acquisition for Quarterly NIKE Defect Review (HFPA NIKE, FTT NIKE & Factory Sandals VB catalog)."""

from __future__ import annotations

import calendar
from datetime import date, datetime, timedelta
from html.parser import HTMLParser
import json
from pathlib import Path
import re
import shutil
from typing import Any, Callable

from dashboard_paths import APP_ROOT
from report_tools.quarterly_defect_review.area_mapping import (
    normalize_factory_code,
    normalize_line_code,
    normalize_plant_code,
    split_combined_plant_line,
)
from report_tools.quarterly_defect_review.contracts import LineCatalogEntry
from report_tools.quarterly_defect_review.report import resolve_period_and_dates

ISOLATED_CACHE_ROOT = APP_ROOT / 'Cache' / 'QuarterlyDefectReview'
CATALOG_CACHE_PATH = ISOLATED_CACHE_ROOT / 'line_catalog_factory_sandals.json'

QA_URL = 'https://sfc.chingluh.com/vhmes/QA/QA.aspx?tree_id=Q602'
QA_DETAIL_BASE = 'https://sfc.chingluh.com/vhmes/QA/'
PLANTS: tuple[str, ...] = ('VH', 'VH2', 'JV', 'JV2')


class _TableAndLinkParser(HTMLParser):
    """Extract links and visible table cells from SFC ASP.NET pages safely."""

    def __init__(self) -> None:
        super().__init__()
        self.links: list[tuple[str, str]] = []
        self.rows: list[list[str]] = []
        self._in_a = False
        self._a_href = ''
        self._a_text: list[str] = []
        self._in_tr = False
        self._in_cell = False
        self._cur_row: list[str] = []
        self._cur_cell: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr_map = dict(attrs)
        t = tag.lower()
        if t == 'a':
            self._in_a = True
            self._a_href = str(attr_map.get('href') or '')
            self._a_text = []
        elif t == 'tr':
            self._in_tr = True
            self._cur_row = []
        elif t in ('td', 'th') and self._in_tr:
            self._in_cell = True
            self._cur_cell = []

    def handle_data(self, data: str) -> None:
        if self._in_a:
            self._a_text.append(data)
        if self._in_cell:
            self._cur_cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t == 'a' and self._in_a:
            text = re.sub(r'\s+', ' ', ''.join(self._a_text)).strip()
            self.links.append((text, self._a_href))
            self._in_a = False
        elif t in ('td', 'th') and self._in_cell:
            cell_txt = re.sub(r'\s+', ' ', ''.join(self._cur_cell)).strip()
            self._cur_row.append(cell_txt)
            self._in_cell = False
        elif t == 'tr' and self._in_tr:
            if any(self._cur_row):
                self.rows.append(self._cur_row)
            self._in_tr = False


def parse_factory_sandals_href(qa_html: str) -> str | None:
    """Locate the dynamic Factory Sandals detail href from QA.aspx HTML (e.g., QADetail.aspx?factory=VT)."""
    parser = _TableAndLinkParser()
    parser.feed(qa_html)
    for text, href in parser.links:
        if 'sandal' in text.lower() and href:
            return href
    match = re.search(r'QADetail\.aspx\?[^"\']*factory=VT[^"\']*', qa_html, flags=re.IGNORECASE)
    return match.group(0) if match else None


def parse_factory_sandals_html(
    html_text: str,
    *,
    factory: str = 'VH',
    provenance: str = 'sfc_qadetail_vt',
) -> list[dict[str, Any]]:
    """Parse live QADetail.aspx?factory=VT page into normalized Factory Sandals LineCatalogEntry dicts."""
    parser = _TableAndLinkParser()
    parser.feed(html_text)
    fac = normalize_factory_code(factory) or 'VH'
    seen: set[tuple[str, str, str]] = set()
    entries: list[dict[str, Any]] = []

    def _push(plant_raw: str, line_raw: str) -> None:
        plant = normalize_plant_code(plant_raw)
        line_clean = re.sub(r'\s+', ' ', str(line_raw or '')).strip()
        if not plant and line_clean:
            plant, line_clean = split_combined_plant_line(line_clean)
        line_std = normalize_line_code(line_clean)
        if not plant or not line_std:
            return
        # Guard against non-line headers/summary strings like 'Total', 'Factory Sandals'
        if plant in {'TOTAL', 'LINE', 'PLANT', 'FACTORY', 'SANDALS', 'ALL'} or line_std in {'TOTAL', 'LINE', 'ALL'}:
            return
        key = (fac, plant, line_std)
        if key in seen:
            return
        seen.add(key)
        entries.append(
            LineCatalogEntry(
                factory=fac,
                plant=plant,
                line_raw=line_clean or line_std,
                line_std=line_std,
                source_group='Factory Sandals',
                report_area='VB',
                provenance=provenance,
            ).to_dict()
        )

    # First inspect explicit line links like QADetail.aspx?...&line=N14%20Line%201 or cell text like "N14 Line 1" / "VB PCC"
    for text, href in parser.links:
        low_href = href.lower()
        if any(tok in low_href for tok in ('line=', 'qadata', 'ftt', 'qa')) and text:
            p_cand, l_cand = split_combined_plant_line(text)
            if p_cand and l_cand:
                _push(p_cand, l_cand)

    # Also inspect explicit two-column or combined table rows from dg/gv tables
    for row in parser.rows:
        if len(row) >= 2 and re.fullmatch(r'(?:N\d+|S\d+|VB|PCC|LINE|[A-Z]\d+)', row[0].strip(), flags=re.IGNORECASE):
            if re.search(r'(?:line|pcc|\d+[a-z]?)', row[1], flags=re.IGNORECASE):
                _push(row[0], row[1])
                continue
        for cell in row:
            if re.match(r'^(?:N\d+|S\d+|VB)\s+(?:Line\s*\d+[A-Za-z]*|PCC)\b', cell, flags=re.IGNORECASE):
                p_cand, l_cand = split_combined_plant_line(cell)
                if p_cand and l_cand:
                    _push(p_cand, l_cand)

    return entries


def _month_windows(date_from: str, date_to: str) -> list[dict[str, str]]:
    start = date.fromisoformat(date_from)
    end = date.fromisoformat(date_to)
    windows: list[dict[str, str]] = []
    cur = date(start.year, start.month, 1)
    while cur <= end:
        last_day = calendar.monthrange(cur.year, cur.month)[1]
        m_start = max(start, cur)
        m_end = min(end, date(cur.year, cur.month, last_day))
        windows.append({
            'month': f'{cur.year:04d}-{cur.month:02d}',
            'date_from': m_start.isoformat(),
            'date_to': m_end.isoformat(),
        })
        if cur.month == 12:
            cur = date(cur.year + 1, 1, 1)
        else:
            cur = date(cur.year, cur.month + 1, 1)
    return windows


def _cached_slot_file(group: str, plant: str, d_from: str, d_to: str, cache_root: Path) -> Path | None:
    folder = cache_root / group.upper() / plant
    if not folder.is_dir():
        return None
    prefix = f'{plant}_{group}_{d_from}_{d_to}'
    for ext in ('.xlsx', '.xlsm', '.xls'):
        cand = folder / f'{prefix}{ext}'
        if cand.is_file() and cand.stat().st_size > 0:
            return cand
    return None


def plan(
    period: str = 'FY27-Q1',
    cutoff: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    *,
    cache_root: Path | str | None = None,
) -> dict[str, Any]:
    """Plan isolated HFPA NIKE, FTT NIKE & Factory Sandals catalog acquisition for the selected window."""
    eff_period, eff_from, eff_to = resolve_period_and_dates(period, date_from, date_to, cutoff)
    root = Path(cache_root) if cache_root else ISOLATED_CACHE_ROOT
    windows = _month_windows(eff_from, eff_to)

    reused: list[dict[str, Any]] = []
    missing: list[dict[str, Any]] = []

    for group in ('hfpa_nike', 'ftt_nike'):
        for plant in PLANTS:
            for win in windows:
                cached = _cached_slot_file(group, plant, win['date_from'], win['date_to'], root)
                slot = {
                    'group': group,
                    'metric': 'HFPA' if group == 'hfpa_nike' else 'FTT',
                    'report_mode': 'NIKE',
                    'plant': plant,
                    'month': win['month'],
                    'date_from': win['date_from'],
                    'date_to': win['date_to'],
                    'file_path': str(cached) if cached else None,
                }
                if cached:
                    reused.append(slot)
                else:
                    missing.append(slot)

    return {
        'period': eff_period,
        'date_from': eff_from,
        'date_to': eff_to,
        'cutoff': cutoff or '',
        'mode': 'NIKE',
        'workflow': 'isolated_nike_review',
        'available': [],
        'cache_root': str(root),
        'catalog_cache_path': str(root / 'line_catalog_factory_sandals.json'),
        'total': len(reused) + len(missing),
        'reused': reused,
        'missing': missing,
    }


def acquire_factory_sandals_catalog(
    *,
    plants: tuple[str, ...] = ('VH',),
    cache_root: Path | str | None = None,
    session_factory: Callable[[str], Any] | None = None,
) -> Path | None:
    """Fetch live Factory Sandals (QADetail.aspx?factory=VT) lines via authenticated SFC HTTP session."""
    root = Path(cache_root) if cache_root else ISOLATED_CACHE_ROOT
    root.mkdir(parents=True, exist_ok=True)
    catalog_out = root / 'line_catalog_factory_sandals.json'

    if session_factory is None:
        import sfc_http
        session_factory = sfc_http.Session

    collected: list[dict[str, Any]] = []
    for plant in plants:
        sess = session_factory(plant)
        # Ensure ASP.NET factory cookie matches plant
        sess.request('https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx')
        raw_qa = sess.request(QA_URL)
        qa_html = raw_qa.decode('utf-8', errors='replace') if isinstance(raw_qa, (bytes, bytearray)) else str(raw_qa or '')
        href = parse_factory_sandals_href(qa_html) or 'QADetail.aspx?factory=VT'
        target_url = href if href.startswith('http') else (QA_DETAIL_BASE + href.lstrip('/'))
        raw_vt = sess.request(target_url)
        vt_html = raw_vt.decode('utf-8', errors='replace') if isinstance(raw_vt, (bytes, bytearray)) else str(raw_vt or '')
        entries = parse_factory_sandals_html(vt_html, factory=plant, provenance=f'sfc:{target_url}')
        collected.extend(entries)

    if not collected:
        return catalog_out if catalog_out.is_file() else None

    payload = {
        'source': 'SFC QA.aspx -> Factory Sandals (VT)',
        'captured_at': datetime.now().isoformat(timespec='seconds'),
        'count': len(collected),
        'entries': collected,
    }
    catalog_out.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding='utf-8')
    return catalog_out


def _stage_file(src: Path, dest_dir: Path) -> Path:
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / src.name
    shutil.copy2(src, dest)
    return dest


class DomExportPending(RuntimeError):
    """The browser is busy or timed out; do not send another export."""

    code = 'sfc.dom_export_pending'

    def __init__(self, message: str, *, busy: bool = False):
        super().__init__(message)
        self.reason = 'busy' if busy else 'timeout'
        self.user_message = 'SFC đang bận hoặc chưa hoàn tất xuất file. Các file đã xác nhận được giữ lại.'


def _dom_result(result: Any, operation: str) -> dict[str, Any]:
    if not isinstance(result, dict) or result.get('success') is not True:
        result = result if isinstance(result, dict) else {}
        message = result.get('error') or f'{operation}: chua nhan duoc ket qua hop le.'
        if (result.get('timed_out') or result.get('busy')
                or re.search(r'\b(?:timeout|timed\s*out)\b', str(message), re.IGNORECASE)):
            raise DomExportPending(message, busy=bool(result.get('busy')))
        raise RuntimeError(message)
    return result


def _validate_nike_download(path: Path, slot: dict[str, Any]) -> None:
    from report_tools.quarterly_defect_review.nike_inputs import normalize_nike_source

    # Do not filter dates here: that would hide records outside the requested slot.
    data = normalize_nike_source(
        [dict(file_path=str(path), factory=slot['plant'])], metric=slot['metric'], run_id='acquisition',
    )
    for record in data['base_records']:
        if record.factory != slot['plant']:
            raise ValueError(f"File NIKE sai nha may: {record.factory}, can {slot['plant']}.")
        if not slot['date_from'] <= record.date <= slot['date_to']:
            raise ValueError(f"File NIKE co ngay ngoai ky: {record.date}.")


def _export_hfpa_dom(ctrl: Any, slot: dict[str, Any], stem: Path,
                     stage: Callable[[str, str | None], None]) -> Path:
    """Export at most seven days through DOM, then publish one checked month."""
    from sfc_excel import combine_exports, file_info

    first = date.fromisoformat(slot['date_from'])
    last = date.fromisoformat(slot['date_to'])
    parts = []
    folder = stem.parent / '_dom_parts' / stem.name
    folder.mkdir(parents=True, exist_ok=True)
    current = first
    while current <= last:
        end = min(current + timedelta(days=6), last)
        part_slot = dict(slot, date_from=current.isoformat(), date_to=end.isoformat())
        part_stem = folder / f'{current.isoformat()}_{end.isoformat()}'
        stage(f"DOM HFPA NIKE {slot['plant']}: {current} - {end}", 'acquisition')
        exported = None
        for suffix in ('.xls', '.xlsx'):
            candidate = Path(str(part_stem) + suffix)
            if candidate.is_file():
                try:
                    _validate_nike_download(candidate, part_slot)
                except (OSError, ValueError):
                    continue
                exported = candidate
                break
        if exported is None:
            result = _dom_result(ctrl.apply_sfc_filters(
                date_from=part_slot['date_from'], date_to=part_slot['date_to'], factory='',
                action='excel', nike=True, export_path=part_stem, expected_plant=slot['plant'],
            ), 'Xuat HFPA NIKE qua DOM')
            info = result.get('file') or result.get('downloaded_file') or {}
            exported = Path(info.get('path') or info.get('file_path') or '')
            if not exported.is_file():
                raise RuntimeError('HFPA NIKE: DOM chua tra ve file Excel da tai xong.')
            _validate_nike_download(exported, part_slot)
            cached_part = Path(str(part_stem) + exported.suffix)
            if exported.resolve() != cached_part.resolve():
                shutil.copy2(exported, cached_part)
            exported = cached_part
        parts.append(dict(file_info(exported), date_from=part_slot['date_from'],
                          date_to=part_slot['date_to'], plant=slot['plant'], nike=True))
        current = end + timedelta(days=1)
    # The disjoint contiguous parts exactly partition the requested monthly slot.
    # The writer validates column schemas and saved row counts before publishing.
    stage(f"Dang ghep {len(parts)} phan DOM HFPA NIKE {slot['plant']}", 'acquisition')
    output = Path(str(stem) + '.xlsx')
    combine_exports(parts, output, slot['plant'], True, include_sources=False)
    return output


def acquire(
    folder: Path | str,
    run: dict[str, Any],
    stage: Callable[[str, str | None], None],
    *,
    controller: Any = None,
    session_factory: Callable[[str], Any] | None = None,
    cache_root: Path | str | None = None,
) -> dict[str, Any]:
    """Export HFPA/FTT NIKE through DOM; HTTP is used only for the line catalog."""
    from report_runtime import Cancelled

    base = Path(folder).resolve()
    ws = base / 'workspace' if (base / 'workspace').is_dir() else base
    inputs_root = ws / 'Inputs'
    root = Path(cache_root) if cache_root else ISOLATED_CACHE_ROOT
    root.mkdir(parents=True, exist_ok=True)
    fields = run.get('fields') if isinstance(run.get('fields'), dict) else {}
    acq_plan = plan(
        period=run.get('period') or fields.get('period') or 'FY27-Q1',
        cutoff=run.get('cutoff') or fields.get('cutoff'),
        date_from=run.get('date_from') or fields.get('date_from'),
        date_to=run.get('date_to') or fields.get('date_to'), cache_root=root,
    )
    errors: list[str] = []
    downloaded: list[dict[str, Any]] = []
    state = {**acq_plan, 'downloaded': downloaded, 'errors': errors, 'staged_files': 0,
             'export_transport': 'DOM'}
    run['acquisition'] = state

    # Cached files must be NIKE workbooks for the correct factory and dates.
    pending = list(acq_plan['missing'])
    reused = []
    for slot in acq_plan['reused']:
        try:
            src = Path(slot['file_path'])
            _validate_nike_download(src, slot)
        except (OSError, ValueError) as exc:
            pending.append(dict(slot, file_path=None))
            stage(f"Cache NIKE can tai lai: {slot['plant']} {slot['month']} ({exc})", 'acquisition')
        else:
            _stage_file(src, inputs_root / slot['group'])
            state['staged_files'] += 1
            reused.append(slot)
    state['reused'] = reused
    state['missing'] = pending

    stage('Dang lay danh muc chuyen Factory Sandals (VB) tu FTT...', 'acquisition')
    try:
        cat_path = acquire_factory_sandals_catalog(
            plants=('VH',), cache_root=root, session_factory=session_factory,
        )
        if cat_path and cat_path.is_file():
            _stage_file(cat_path, inputs_root / 'sandals_catalog')
    except Cancelled:
        raise
    except Exception as exc:
        cached = root / 'line_catalog_factory_sandals.json'
        if cached.is_file():
            _stage_file(cached, inputs_root / 'sandals_catalog')
            errors.append(f'Catalog VB dung cache, chua refresh duoc: {exc}')
        else:
            errors.append(f'Chua lay duoc danh muc Factory Sandals: {exc}')

    ctrl = controller
    if pending and ctrl is None:
        import sfc_controller
        ctrl = sfc_controller

    # Complete a pair of HFPA/FTT exports before moving to the next factory/month.
    # A slow HFPA quarter must not hide whether FTT export works.
    pending.sort(key=lambda s: (s['month'], PLANTS.index(s['plant']), s['metric'] != 'HFPA'))
    for index, slot in enumerate(pending, 1):
        metric, plant = slot['metric'], slot['plant']
        stage(f"[DOM {index}/{len(pending)}] Xuat {metric} NIKE {plant}: "
              f"{slot['date_from']} - {slot['date_to']}", 'acquisition')
        target_dir = root / slot['group'].upper() / plant
        target_dir.mkdir(parents=True, exist_ok=True)
        stem = target_dir / f"{plant}_{slot['group']}_{slot['date_from']}_{slot['date_to']}"
        try:
            _dom_result(ctrl.switch_factory(plant), 'Chuyen nha may SFC')
            if metric == 'HFPA':
                exported_hfpa = _export_hfpa_dom(ctrl, slot, stem, stage)
                info = {'path': str(exported_hfpa)}
            else:
                _dom_result(ctrl.navigate_to(QA_URL), 'Mo QA FTT')
                result = _dom_result(ctrl.execute_dynamic_action('ftt_station_export', None, {
                    '_expected_plant': plant, 'date_from': slot['date_from'],
                    'date_to': slot['date_to'], 'report_type': 'nike',
                    'stations': [], 'station_selection': 'none', 'file_name': stem.name,
                }), 'Xuat FTT NIKE qua DOM')
                info = result.get('downloaded_file') or {}
            exported = Path(info.get('path') or info.get('file_path') or '')
            if not exported.is_file():
                raise RuntimeError(f'{metric} NIKE: DOM chua tra ve file Excel da tai xong.')
            _validate_nike_download(exported, slot)
            cached = Path(str(stem) + exported.suffix)
            if exported.resolve() != cached.resolve():
                shutil.copy2(exported, cached)
            staged = _stage_file(cached, inputs_root / slot['group'])
            state['staged_files'] += 1
            downloaded.append({key: slot[key] for key in ('group', 'plant', 'month', 'date_from', 'date_to')}
                              | dict(name=staged.name, report_mode='NIKE', transport='DOM'))
            # Removing completed slots from missing also makes a cancelled run
            # accurately describe the files already kept for the next attempt.
            state['missing'] = [s for s in state['missing'] if s is not slot]
            stage(f"[DOM {index}/{len(pending)}] Da luu va kiem tra {staged.name}", 'acquisition')
        except DomExportPending as exc:
            state['blocked'] = {key: slot[key] for key in
                                ('group', 'metric', 'plant', 'month', 'date_from', 'date_to')}
            state['blocked']['reason'] = exc.reason
            errors.append(f"{slot['group']}/{plant}/{slot['month']}: {exc}")
            condition = 'SFC đang bận' if exc.reason == 'busy' else 'SFC chưa hoàn tất sau thời gian chờ'
            exc.user_message = (
                f"{condition}: {metric} NIKE {plant}, {slot['date_from']} đến {slot['date_to']}. "
                f"Đã xác nhận {state['staged_files']}/{state['total']} file. "
                'Kiểm tra trang SFC trước khi chạy lại; các file đã xác nhận sẽ được dùng lại.'
            )
            stage(exc.user_message, 'acquisition')
            raise
        except Cancelled:
            raise
        except Exception as exc:
            message = f"{slot['group']}/{plant}/{slot['month']}: {exc}"
            errors.append(message)
            stage(message, 'acquisition')

    if not state['staged_files'] and errors:
        raise RuntimeError('Chua xuat duoc du lieu NIKE qua DOM: ' + '; '.join(errors[:3]))
    return state
