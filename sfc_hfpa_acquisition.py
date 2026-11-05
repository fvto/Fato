"""Monthly HFPA sources: QAStation HFPA/OK and MES410 with NIKE disabled."""
from datetime import date, datetime
from functools import lru_cache
from pathlib import Path
import re

import sfc_report_tools as reports
from sfc_copq_acquisition import month_window
from sfc_excel import read_export, validate_export

PLANTS = ('VH', 'VH2', 'JV', 'JV2')
QA_URL = 'https://sfc.chingluh.com/vhmes/QA/QA.aspx'
HFPA_URL = 'https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504'


def source_path(item):
    if item['source'] == 'sfc':
        return reports.EXPORT_ROOT / item['id'][4:]
    return reports.original_root('hfpa') / reports.TOOLS['hfpa']['groups'][item['group']][1] / item['relative']


@lru_cache(maxsize=512)
def valid_month(filename, size, modified, group, first, last):
    try:
        if group == 'hfpa':
            validate_export(filename, first, last, nike=False)
            return True
        header, rows = read_export(filename)
        labels = [str(value or '').strip().lower() for value in header]
        if not {'inspdate','line','shoename','inspqty','failqty','station','issues','issueqty'} <= set(labels):
            return False  # FTT Internal is a different export from QAStation/OK.
        day_col, station_col = labels.index('inspdate'), labels.index('station')
        for row in rows:
            if str(row[station_col] or '').strip().upper() != 'HFPA':
                return False
            value = row[day_col]
            if isinstance(value, datetime): value = value.date()
            elif not isinstance(value, date):
                raw = str(value).strip()
                for fmt in ('%Y-%m-%d','%d-%b-%Y','%d/%m/%Y','%m/%d/%Y'):
                    try:
                        value = datetime.strptime(raw, fmt).date()
                        break
                    except ValueError:
                        continue
                else:
                    return False
            if not date.fromisoformat(first) <= value <= date.fromisoformat(last):
                return False
        return True  # An empty valid monthly export also counts as acquired.
    except Exception:
        return False


def validate_file(path, group, first, last):
    path = Path(path)
    stat = path.stat()
    return valid_month(str(path), stat.st_size, stat.st_mtime_ns, group, first, last)


def cache_matches(item, first, last, path=None):
    name = item['name']
    full_range = first + '_' + last in name
    if item['group'] == 'ftt':
        if not full_range or not name.lower().startswith('qastation_'):
            return False
    else:
        # Legacy local MES410 names contain the download timestamp; use the
        # actual audit period. Unproven SFC timestamps may be partial exports.
        legacy = item['source'] == 'local' and item.get('period') == first[:7] and re.match(r'^HFPA\s+20\d{2}-', name, re.I)
        if not (full_range and '_internal_' in name.lower() or legacy):
            return False
    return validate_file(source_path(item) if path is None else path, item['group'], first, last)


def plan(period, mode='all'):
    if mode not in reports.TOOLS['hfpa']['modes']:
        raise ValueError('Chế độ HFPA không hợp lệ.')
    first, last = month_window(period)
    available = {}
    for item in reports.input_catalog('hfpa'):
        key = (item['group'], item.get('plant'))
        if key[1] not in PLANTS or not cache_matches(item, first, last):
            continue
        if key not in available or item['source'] == 'sfc' and available[key]['source'] != 'sfc':
            available[key] = item
    reused, missing = [], []
    for group in ('ftt','hfpa'):
        for plant in PLANTS:
            if (group,plant) in available:
                reused.append(available[(group,plant)])
            else:
                missing.append(dict(group=group,plant=plant))
    return dict(period=period,date_from=first,date_to=last,total=8,available=reused,missing=missing,
        workflow=dict(ftt=dict(url=QA_URL,plants=list(PLANTS),stations=['HFPA'],report_type='station'),
                      hfpa=dict(url=HFPA_URL,plants=list(PLANTS),nike=False,factory='',report_type='internal')))


def download(controller, group, plant, window):
    result = controller.navigate_to(QA_URL if group=='ftt' else HFPA_URL)
    if not result.get('success'): return result
    if group == 'ftt':
        return controller._execute_dynamic_action('ftt_station_export',params=dict(
            date_from=window['date_from'],date_to=window['date_to'],_expected_plant=plant,
            stations=['HFPA'],station_selection='explicit',report_type='station',
            file_name=f'QAStation_{window["date_from"]}_{window["date_to"]}'))
    result = controller.apply_sfc_filters(date_from=window['date_from'],date_to=window['date_to'],
        factory='',nike=False,action='excel',expected_plant=plant)
    return dict(result,downloaded_file=result.get('file'))


def filename(group, plant, window, suffix):
    prefix = 'QAStation_' if group=='ftt' else f'HFPA_{plant}_Internal_'
    return f'{prefix}{window["date_from"]}_{window["date_to"]}-{plant}{suffix}'


def acquire(folder, run, stage, controller=None):
    from sfc_report_acquisition import acquire_month
    return acquire_month(folder, run, stage, controller, tool='hfpa')
