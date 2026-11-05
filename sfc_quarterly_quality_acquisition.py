"""Acquire QAStation HFPA, MES410 and FTT Internal in quarter month windows."""

from datetime import date, datetime, timedelta, timezone
import shutil
import zipfile

import sfc_report_tools as reports
from report_tools.quarterly_quality.quarterly_quality import PLANTS, quarter_months
from sfc_copq_acquisition import FTT_URL, download as download_ftt, month_window
from sfc_hfpa_acquisition import HFPA_URL, QA_URL, download as download_hfpa, validate_file as validate_hfpa_file


def local_today():
    return datetime.now(timezone(timedelta(hours=7))).date()


def plan(period, cutoff):
    from report_tools.quarterly_quality.tableau_source import default_path, read_source

    current = quarter_months(period)[3:]
    first, end = month_window(current[0])[0], month_window(current[-1])[1]
    try:
        last = date.fromisoformat(cutoff) if isinstance(cutoff, str) else None
    except ValueError:
        last = None
    if last is None or last.isoformat() != cutoff or not first <= cutoff <= end:
        raise ValueError(f'Chọn ngày chốt trong quý, từ {first} đến {end}.')
    if last > local_today():
        raise ValueError('Ngày chốt SFC chưa đến. Chọn ngày hôm nay hoặc ngày trước đó.')
    index = current.index(cutoff[:7])
    path = default_path()
    tableau = read_source(path, current, cutoff) if path else None
    slots = []
    for month in current[:index + 1]:
        start, end = month_window(month)
        last = min(end, cutoff)
        for metric, groups in (('HFPA', ('hfpa', 'qastation')), ('FTT', ('ftt',))):
            for plant in PLANTS:
                covered = tableau.covered_to(metric, month, plant, last) if tableau else None
                if covered == last:
                    continue
                missing_from = (date.fromisoformat(covered) + timedelta(days=1)).isoformat() if covered else start
                for group in groups:
                    slots.append(dict(group=group, plant=plant, date_from=missing_from, date_to=last))
    result = dict(date_from=first, date_to=cutoff, total=len(slots), available=[], missing=slots,
                workflow=dict(hfpa=dict(url=HFPA_URL, plants=list(PLANTS), nike=False, report_type='internal'),
                              qastation=dict(url=QA_URL, plants=list(PLANTS), stations=['HFPA'], report_type='station'),
                              ftt=dict(url=FTT_URL, plants=list(PLANTS), stations=[], report_type='internal')))
    if tableau:
        result['tableau'] = dict(path=str(tableau.path), sha256=tableau.revision, sheet='VH FTT',
                                 cached_groups=len(tableau.values))
    return result


def download(controller, group, plant, window):
    if group == 'qastation':
        return download_hfpa(controller, 'ftt', plant, window)
    return (download_hfpa if group == 'hfpa' else download_ftt)(controller, group, plant, window)


def filename(group, plant, window, suffix):
    prefix = {'hfpa': f'HFPA_{plant}_Internal_', 'qastation': 'QAStation_', 'ftt': 'FTT-Internal_'}[group]
    return f'{prefix}{window["date_from"]}_{window["date_to"]}-{plant}{suffix}'


def validate_file(path, group, first, last):
    """Check the container before staging; the worker validates every source row."""
    if path.stat().st_size > reports.CONTRACTS['max_file_bytes']:
        return False
    with path.open('rb') as stream:
        signature = stream.read(8)
    if signature == reports.OLE:
        return group != 'qastation' or validate_hfpa_file(path, 'ftt', first, last)
    if not signature.startswith(b'PK') or not zipfile.is_zipfile(path):
        return False
    with zipfile.ZipFile(path) as package:
        valid = ('xl/workbook.xml' in package.namelist() and
                 sum(entry.file_size for entry in package.infolist()) <= reports.CONTRACTS['max_expanded_bytes'])
    return valid and (group != 'qastation' or validate_hfpa_file(path, 'ftt', first, last))


def acquire(folder, run, stage, controller=None):
    from sfc_report_acquisition import acquire_month

    window = run['acquisition']
    # Older saved runs can still contain download slots for the preceding quarter.
    window['missing'] = [slot for slot in window['missing']
                         if window['date_from'] <= slot['date_from'] <= window['date_to']]
    covered = [month for month in quarter_months(run['period'])[3:] if month <= window['date_to'][:7]]
    if not window.get('tableau'):
        window['total'] = len(covered) * len(PLANTS) * 3
    elif not window['missing']:
        stage('Dữ liệu trong sheet VH FTT đã đủ đến ngày chốt.')
        return
    stage('Đang tìm file SFC đã xuất đúng nhà máy, mẫu và khoảng ngày...')
    root = reports.EXPORT_ROOT.resolve()
    candidates = {}
    for path in root.rglob('*'):
        if path.suffix.lower() in ('.xls', '.xlsx', '.xlsm') and path.is_file() and path.resolve().is_relative_to(root):
            candidates.setdefault(path.name.lower(), []).append(path)
    missing = []
    today = local_today().isoformat()
    for slot in window['missing']:
        group, plant = slot['group'], slot['plant']
        paths = [p for suffix in ('.xlsx', '.xlsm', '.xls')
                 for p in candidates.get(filename(group, plant, slot, suffix).lower(), [])]
        # A file exported earlier today can be stale for a cutoff that includes today.
        source = next((p for p in sorted(paths, key=lambda item: item.stat().st_mtime_ns, reverse=True)
                       if slot['date_to'] < today and validate_file(p, group, slot['date_from'], slot['date_to'])), None)
        if source is None:
            missing.append(slot)
            continue
        stage(f'Dùng lại {group.upper()} · {plant} · {slot["date_from"]} đến {slot["date_to"]}')
        name = filename(group, plant, slot, source.suffix.lower())
        destination = folder / 'workspace' / reports.TOOLS['quarterly_quality']['groups'][group][1] / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        before = source.stat()
        shutil.copyfile(source, destination)
        after = source.stat()
        if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
            missing.append(slot)
            continue
        record = dict(slot, name=name)
        run['inputs'].append(dict(record, size=destination.stat().st_size, sha256=reports.file_hash(destination)))
        window['reused'].append(record)
    window['missing'] = missing
    acquire_month(folder, run, stage, controller, tool='quarterly_quality')
