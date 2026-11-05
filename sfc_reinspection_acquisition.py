"""User-confirmed Re-Inspection sources for the same month and four plants."""
from datetime import date, datetime
from decimal import InvalidOperation
from functools import lru_cache
from pathlib import Path
import re

import sfc_copq_acquisition as copq
import sfc_hfpa_acquisition as hfpa
import sfc_report_tools as reports
from sfc_excel import read_export
from report_validation import number

PLANTS = ('VH','VH2','JV','JV2')
REINSPECTION_URL = 'https://sfc.chingluh.com/vhmes/QA/QCRECHK.aspx'


def source_path(item):
    if item['source'] == 'sfc':
        return reports.EXPORT_ROOT / item['id'][4:]
    return reports.original_root('reinspection') / reports.TOOLS['reinspection']['groups'][item['group']][1] / item['relative']


@lru_cache(maxsize=512)
def valid_month(filename, size, modified, group, first, last):
    try:
        if group == 'hfpa':
            return hfpa.validate_file(Path(filename),group,first,last)
        if group == 'ftt' and not copq.valid_schema(filename,size,modified,'ftt'):
            return False
        header, rows = read_export(filename)
        labels = [str(value or '').strip().lower() for value in header]
        if group == 'reinspection':
            if not {'category','so#','country','model 2','reinspection q\'ty',
                    '2nd audit sample size','2nd pass/fail','remark'} <= set(labels):
                return False
            column = next(i for i,label in enumerate(labels) if 'audit date' in label)
        else:
            column = labels.index('date')
        if group == 'reinspection' and rows:
            # QCRECHK has no Audit ID column. Its trailing DevExpress totals
            # row stores record count in Audit Date and total defects beside it.
            tail = rows[-1]
            populated = {i for i,value in enumerate(tail) if value not in ('',None)}
            totals = {column} | {i for i,label in enumerate(labels) if label.startswith("defect q'ty")}
            if populated <= totals:
                try:
                    if number(tail[column]) == len(rows) - 1:
                        rows = rows[:-1]
                except (ValueError, InvalidOperation):
                    pass  # A regular record still needs the normal audit-date check.
        first_day, last_day = date.fromisoformat(first), date.fromisoformat(last)
        for row in rows:
            value = row[column]
            # QCRECHK can include a real re-inspection whose original audit
            # was not recorded. Keep it unchanged when Category and SO# identify
            # the record; the verified export period is checked by cache_matches.
            if (group == 'reinspection' and (value is None or isinstance(value,str) and not value.strip())
                    and row[labels.index('category')] in ('QC','MA')
                    and str(row[labels.index('so#')] or '').strip()):
                continue
            if isinstance(value,datetime): value = value.date()
            elif not isinstance(value,date):
                for fmt in ('%Y-%m-%d','%d-%b-%Y','%d/%m/%Y','%m/%d/%Y'):
                    try:
                        value = datetime.strptime(str(value).strip(),fmt).date()
                        break
                    except ValueError:
                        continue
                else:
                    return False
            # QCRECHK's Audit Date is the original audit date. Existing monthly
            # SFC exports include Aug 31 audits in September re-inspections;
            # the requested period belongs to the verified export filters/name.
            if group != 'reinspection' and not first_day <= value <= last_day:
                return False
        return True
    except Exception:
        return False


def validate_file(path, group, first, last):
    path = Path(path)
    stat = path.stat()
    return valid_month(str(path),stat.st_size,stat.st_mtime_ns,group,first,last)


def cache_matches(item, first, last):
    group,name = item['group'],item['name']
    if group == 'hfpa':
        return hfpa.cache_matches(item,first,last,path=source_path(item))
    prefix = r'^Re-Inspection[ _]+' if group=='reinspection' else r'^FTT-Internal_'
    if not re.match(prefix + re.escape(first+'_'+last) + r'-',name,re.I):
        return False  # Exclude station exports, filtered types, and partial months.
    return validate_file(source_path(item),group,first,last)


def plan(period, mode='all'):
    if mode not in reports.TOOLS['reinspection']['modes']:
        raise ValueError('Chế độ Re-Inspection không hợp lệ.')
    first,last = copq.month_window(period)
    available = {}
    for item in reports.input_catalog('reinspection'):
        key = (item['group'],item.get('plant'))
        existing = available.get(key)
        if key[1] not in PLANTS or (existing and (existing['source'] == 'sfc' or item['source'] != 'sfc')):
            continue
        if cache_matches(item,first,last):
            available[key] = item
    reused,missing = [],[]
    for group in ('reinspection','ftt','hfpa'):
        for plant in PLANTS:
            if (group,plant) in available: reused.append(available[(group,plant)])
            else: missing.append(dict(group=group,plant=plant))
    return dict(period=period,date_from=first,date_to=last,total=12,available=reused,missing=missing,
        workflow=dict(reinspection=dict(url=REINSPECTION_URL,plants=list(PLANTS),line=None,type=None),
                      ftt=dict(url=copq.FTT_URL,plants=list(PLANTS),stations=[],report_type='internal'),
                      hfpa=dict(url=hfpa.HFPA_URL,plants=list(PLANTS),nike=False,factory='',report_type='internal')))


def download(controller, group, plant, window):
    if group == 'ftt': return copq.download(controller,group,plant,window)
    if group == 'hfpa': return hfpa.download(controller,group,plant,window)
    result = controller.navigate_to(REINSPECTION_URL)
    if not result.get('success'): return result
    return controller._execute_dynamic_action('reinspection_export',params=dict(
        date_from=window['date_from'],date_to=window['date_to'],_expected_plant=plant))


def filename(group, plant, window, suffix):
    if group == 'ftt': return copq.filename(group,plant,window,suffix)
    if group == 'hfpa': return hfpa.filename(group,plant,window,suffix)
    return f'Re-Inspection {window["date_from"]}_{window["date_to"]}-{plant}{suffix}'


def acquire(folder, run, stage, controller=None):
    from sfc_report_acquisition import acquire_month
    return acquire_month(folder,run,stage,controller,tool='reinspection')
