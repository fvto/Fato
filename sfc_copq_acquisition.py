"""The user-confirmed monthly acquisition contract for COPQ reports."""
import calendar
import csv
from functools import lru_cache
import io
from pathlib import Path
import re
import zipfile

import sfc_report_tools as reports

COPQ_PLANTS = ('VH', 'VH2', 'VH3', 'VH4', 'JV', 'JV2', 'JVB')
FTT_PLANTS = ('VH', 'VH2', 'JV', 'JV2')
COPQ_URL = 'https://sfc.chingluh.com/vhmes/QA/QAReportCOPQ.aspx'
FTT_URL = 'https://sfc.chingluh.com/vhmes/QA/QA.aspx'


def month_window(period):
    if not re.fullmatch(r'20\d{2}-(0[1-9]|1[0-2])', period or ''):
        raise ValueError('Chọn tháng báo cáo theo định dạng YYYY-MM.')
    year, month = map(int, period.split('-'))
    return period + '-01', f'{period}-{calendar.monthrange(year, month)[1]:02d}'


def required_inputs(mode):
    if mode == 'combine': return {'copq':COPQ_PLANTS}
    if mode == 'ftt': return {'ftt':FTT_PLANTS}
    if mode == 'all': return {'copq':COPQ_PLANTS, 'ftt':FTT_PLANTS}
    raise ValueError('Chế độ COPQ không hợp lệ.')


def source_path(item):
    if item['source'] == 'sfc':
        return reports.EXPORT_ROOT / item['id'][4:]
    return reports.original_root('copq') / reports.TOOLS['copq']['groups'][item['group']][1] / item['relative']


@lru_cache(maxsize=512)
def valid_schema(filename, size, modified, group):
    path = Path(filename)
    try:
        if path.suffix.lower() == '.csv':
            with path.open(encoding='utf-8-sig') as stream:
                rows = [row for _,row in zip(range(12), csv.reader(stream))]
        elif zipfile.is_zipfile(path):
            import openpyxl
            with path.open('rb') as stream:
                book = openpyxl.load_workbook(stream, read_only=True, data_only=True)
                try:
                    rows = list(book.worksheets[0].iter_rows(min_row=1,max_row=12,values_only=True))
                finally:
                    book.close()
        else:
            import xlrd
            book = xlrd.open_workbook(str(path), on_demand=True, logfile=io.StringIO())
            try:
                sheet = book.sheet_by_index(0)
                rows = [sheet.row_values(i) for i in range(min(12,sheet.nrows))]
            finally:
                book.release_resources()
        for row in rows:
            labels = {re.sub(r'[^a-z0-9]', '', str(value).lower()) for value in row}
            required = {'date','category','type','factory','ttlcost'} if group == 'copq' else {'date','factory','metric','numberofpairs','defectissues','so'}
            if required <= labels:
                return True
    except Exception:
        pass
    return False


def cache_matches(item, first, last):
    name = item['name']
    if first + '_' + last not in name:
        return False
    if item['group'] == 'copq' and not re.match(r'^Cost Of Poor Quality\s+20\d{2}-', name, re.I):
        return False  # Type-specific and summary files cannot stand in for cleared COPQ filters.
    if item['group'] == 'ftt' and not name.lower().startswith('ftt-internal_'):
        return False
    path = source_path(item)
    stat = path.stat()
    return valid_schema(str(path), stat.st_size, stat.st_mtime_ns, item['group'])


def plan(period, mode='all'):
    first, last = month_window(period)
    required = required_inputs(mode)
    available = {}
    for item in reports.input_catalog('copq'):
        group, plant = item['group'], item.get('plant')
        if group not in required or plant not in required[group] or not cache_matches(item, first, last):
            continue
        key = (group,plant)
        if key not in available or item['source'] == 'sfc' and available[key]['source'] != 'sfc':
            available[key] = item
    reused, missing = [], []
    for group, plants in required.items():
        for plant in plants:
            if (group,plant) in available:
                reused.append(available[(group,plant)])
            else:
                missing.append(dict(group=group,plant=plant))
    return dict(period=period, date_from=first, date_to=last, total=sum(map(len,required.values())),
                available=reused, missing=missing,
                workflow=dict(copq=dict(url=COPQ_URL,plants=list(COPQ_PLANTS),category=None,type=None),
                              ftt=dict(url=FTT_URL,plants=list(FTT_PLANTS),stations=[],report_type='internal')))


def download(controller, group, plant, window):
    result = controller.navigate_to(COPQ_URL if group=='copq' else FTT_URL)
    if not result.get('success'): return result
    params = dict(date_from=window['date_from'],date_to=window['date_to'],_expected_plant=plant)
    if group == 'copq':
        params['selections'] = {
            'rpnlMain_ContentPlaceHolder1_cboCategory':dict(value='',text=''),
            'rpnlMain_ContentPlaceHolder1_cboType':dict(value='',text='')}
        return controller._execute_dynamic_action('click_button','rpnlMain_ContentPlaceHolder1_btnExcel',params)
    params.update(stations=[],station_selection='none',report_type='internal',
                  file_name=f'FTT-Internal_{window["date_from"]}_{window["date_to"]}')
    return controller._execute_dynamic_action('ftt_station_export',params=params)


def validate_file(path, group, first, last):
    stat = path.stat()
    return valid_schema(str(path),stat.st_size,stat.st_mtime_ns,group)


def filename(group, plant, window, suffix):
    prefix = 'Cost Of Poor Quality ' if group=='copq' else 'FTT-Internal_'
    return f'{prefix}{window["date_from"]}_{window["date_to"]}-{plant}{suffix}'


def acquire(folder, run, stage, controller=None):
    from sfc_report_acquisition import acquire_month
    return acquire_month(folder, run, stage, controller, tool='copq')
