"""Recover completed MQAA SQL timeouts with sequential, verified daily exports."""
from copy import deepcopy
from datetime import date, datetime, timedelta
import logging
from pathlib import Path
from uuid import uuid4

import openpyxl
from openpyxl.styles import Font, PatternFill

import sfc_downloads as downloads
from sfc_excel import file_info, read_export, safe_cell
from sfc_report_tools import write_json

logger=logging.getLogger(__name__)


def recover_timeout_page(page,action_type,target_id,params,timeout):
    if (action_type=='click_button' and 'mqaaReport.aspx'.lower() in page.url.lower()
            and any(k in (target_id or '').lower() for k in ('excel','export','download'))
            and 'execution timeout expired' in page.title().lower()):
        if not params.get('date_from') or not params.get('date_to'):
            raise ValueError('Chọn lại khoảng ngày MQAA để tải từng ngày sau lỗi timeout.')
        page.goto(page.url,wait_until='domcontentloaded',timeout=timeout)
        return True
    return False


def date_controls(filters):
    first = [key for key,item in filters.items() if item['kind']=='date' and 'from' in key.lower()]
    last = [key for key,item in filters.items() if item['kind']=='date' and 'to' in key.lower()]
    if len(first)!=1 or len(last)!=1:
        raise ValueError('Không xác định được hai bộ lọc ngày MQAA; chưa tải từng ngày.')
    return first[0],last[0]


def verify_scope(page, filters, plant):
    actual_filters=downloads.capture_filters(page)
    if any(actual_filters.get(key)!=value for key,value in filters.items()):
        raise ValueError('Bộ lọc MQAA thay đổi trong khi xuất; dừng để tránh trộn dữ liệu.')
    actual = next((c['value'] for c in page.context.cookies() if c['name']=='MyFactoryName'),None)
    if downloads.factory_code(actual)!=plant:
        raise ValueError('Nhà máy MQAA thay đổi trong khi xuất; chưa gộp dữ liệu.')


def export_once(page, target_id, timeout, plant, format_name, filters):
    """Observe the finished POST, including SQL errors, before awaiting a download."""
    downloads.apply_filters(page,filters,timeout)
    verify_scope(page,filters,plant)
    url=page.url
    received=[]
    handler=lambda download:received.append(download)
    page.on('download',handler)
    try:
        with page.expect_response(lambda r:r.request.method=='POST' and r.url==url,timeout=timeout) as pending:
            page.locator('#'+target_id).click(no_wait_after=True)
        downloads.require_attachment(pending.value)
        verify_scope(page,filters,plant)
        download=received[0] if received else page.wait_for_event('download',timeout=timeout)
        return downloads.save_download(download,plant,format_name)
    finally:
        page.remove_listener('download',handler)


def audit_day(value):
    if isinstance(value,datetime): return value.date().isoformat()
    if isinstance(value,date): return value.isoformat()
    for fmt in ('%Y-%m-%d','%d-%b-%Y','%d/%m/%Y','%m/%d/%Y'):
        try: return datetime.strptime(str(value).strip(),fmt).date().isoformat()
        except ValueError: pass
    raise ValueError('Không đọc được Audit Date MQAA: '+repr(value))


def daily_table(path, day, plant):
    """MQAA All has a title and date preamble before its actual table header."""
    title,body=read_export(path)
    rows=[title]+body
    report_title=str(title[0] or '').strip().lower()
    if not report_title.startswith('mqaa'):
        raise ValueError('File tải về không có tiêu đề báo cáo MQAA; chưa gộp.')
    for index,header in enumerate(rows[:12]):
        labels=[str(v or '').strip().lower() for v in header]
        if {'audit date','audit form','factory'}<=set(labels):
            break
    else:
        if report_title.endswith('all'):
            raise ValueError('MQAA All report thiếu cột Audit Date, Audit form hoặc Factory; chưa gộp.')
        # Aggregate report layouts are retained on individual daily sheets.
        return None,rows
    date_column=labels.index('audit date')
    factory_column=labels.index('factory')
    percent_column=labels.index('mqaa%') if 'mqaa%' in labels else None
    records=[]
    for row in rows[index+1:]:
        populated={i for i,v in enumerate(row) if v not in ('',None)}
        if not populated: continue
        # The native export can end with AVERAGE(MQAA%) and a cached percentage.
        # These are daily summaries, never audit records to append to the range.
        if percent_column is not None and populated=={percent_column}:
            continue
        if audit_day(row[date_column])!=day:
            raise ValueError('File MQAA chứa ngày ngoài ngày yêu cầu: '+day)
        if downloads.factory_code(str(row[factory_column]).strip().upper())!=plant:
            raise ValueError('File MQAA chứa nhà máy khác '+plant+'; chưa gộp.')
        records.append(row)
    return tuple(header),records


def merge_days(parts, target, plant, first, last, filters):
    days=[]
    current,end=date.fromisoformat(first),date.fromisoformat(last)
    while current<=end:
        days.append(current.isoformat());current+=timedelta(days=1)
    if ([p['date_from'] for p in parts]!=days or
            any(p['date_from']!=p['date_to'] or p['plant']!=plant for p in parts)):
        raise ValueError('Chưa đủ ngày hoặc bị trùng ngày/nhà máy; chưa tạo file gộp MQAA.')
    tables=[daily_table(p['path'],p['date_from'],plant) for p in parts]
    tabular=all(header is not None for header,_ in tables)
    if any(header is not None for header,_ in tables) and not tabular:
        raise ValueError('Mẫu báo cáo MQAA thay đổi giữa các ngày; chưa gộp.')
    book=openpyxl.Workbook(write_only=True)
    count=0
    if tabular:
        header=tables[0][0]
        if any(h!=header for h,_ in tables):
            raise ValueError('Cấu trúc cột MQAA thay đổi giữa các ngày; chưa gộp.')
        count=sum(len(rows) for _,rows in tables)
        if count>1048575: raise ValueError('Vượt giới hạn dòng Excel; giữ các file ngày để kiểm tra.')
        sheet=book.create_sheet('MQAA')
        sheet.freeze_panes='A2'
        cells=[]
        for label in header:
            cell=safe_cell(sheet,label)
            cell.font=Font(name='Calibri',bold=True,color='FFFFFF')
            cell.fill=PatternFill('solid',fgColor='17365D')
            cells.append(cell)
        sheet.append(cells)
        labels=[str(v or '').strip().lower() for v in header]
        for i,label in enumerate(header,1):
            sheet.column_dimensions[openpyxl.utils.get_column_letter(i)].width=min(48,max(16,len(str(label))+2))
        for _,rows in tables:
            for row in rows:
                cells=[]
                for i,value in enumerate(row):
                    cell=safe_cell(sheet,value)
                    if labels[i]=='mqaa%' and isinstance(value,(int,float)):cell.number_format='0.0%'
                    if labels[i]=='audit date' and isinstance(value,(date,datetime)):cell.number_format='yyyy-mm-dd'
                    cells.append(cell)
                sheet.append(cells)
        sheet.auto_filter.ref=f'A1:{openpyxl.utils.get_column_letter(len(header))}{count+1}'
    else:
        for part,(_,rows) in zip(parts,tables):
            sheet=book.create_sheet(part['date_from'])
            for row in rows: sheet.append([safe_cell(sheet,v) for v in row])
    sources=book.create_sheet('NguonTheoNgay')
    sources.append(['Nhà máy','Từ ngày','Đến ngày','File nguồn','Số dòng dữ liệu','SHA256'])
    for part,(header,rows) in zip(parts,tables):
        sources.append([plant,part['date_from'],part['date_to'],part['name'],len(rows) if header else None,part['sha256']])
    sources.append(['Bộ lọc MQAA',str(filters)])
    target=Path(target)
    pending=target.with_suffix('.partial.xlsx')
    try:
        book.save(pending)
        check=openpyxl.load_workbook(pending,read_only=True,data_only=False)
        try:
            if tabular and sum(1 for _ in check['MQAA'].iter_rows(min_row=2,values_only=True))!=count:
                raise ValueError('Số dòng MQAA sau khi gộp không khớp; chưa công bố file.')
        finally: check.close()
        pending.replace(target)
    finally: pending.unlink(missing_ok=True)
    return dict(file_info(target,count if tabular else None),plant=plant,date_from=first,date_to=last,
                daily_fallback=True,parts=parts)


def export_mqaa_data(page,target_id,timeout,plant,format_name,confirmed_timeout=False):
    url=page.url
    filters=downloads.capture_filters(page)
    if not confirmed_timeout:
        try:
            return export_once(page,target_id,timeout,plant,format_name,filters)
        except downloads.SFCServerTimeout:
            pass  # The server has finished; it is now safe to issue smaller requests.
    from_key,to_key=date_controls(filters)
    first,last=filters[from_key]['value'],filters[to_key]['value']
    current,end=date.fromisoformat(first),date.fromisoformat(last)
    if current>=end:
        raise downloads.SFCServerTimeout('SFC vẫn báo timeout khi xuất một ngày; chưa có file gộp MQAA.')
    logger.info('MQAA %s: SQL timeout confirmed; exporting daily %s to %s',plant,first,last)
    folder=downloads.EXPORT_ROOT/('mqaa_daily_'+uuid4().hex)
    folder.mkdir(parents=True)
    parts=[]
    manifest=dict(status='running',prompt=f'MQAA {plant} · {first} → {last}',
        intent=dict(target_plants=[plant],date_from=first,date_to=last),filters=filters,
        monthly_files=parts,combined_files=[],export_method='sql_timeout_daily_v1')
    failure=None
    try:
        write_json(folder/'manifest.json',manifest)
        # The confirmed SQL error replaced the WebForms page; load a fresh form once.
        page.goto(url,wait_until='domcontentloaded',timeout=timeout)
        while current<=end:
            day=current.isoformat()
            logger.info('MQAA %s: downloading %s',plant,day)
            daily=deepcopy(filters)
            daily[from_key]['value']=daily[to_key]['value']=day
            part=export_once(page,target_id,timeout,plant,format_name,daily)
            daily_table(part['path'],day,plant)
            parts.append(dict(part,date_from=day,date_to=day))
            logger.info('MQAA %s: verified daily file %s (%s days complete)',plant,day,len(parts))
            write_json(folder/'manifest.json',manifest)
            current+=timedelta(days=1)
        result=merge_days(parts,folder/f'MQAA_{first}_{last}-{plant}.xlsx',plant,first,last,filters)
        manifest.update(status='done',combined_files=[{k:v for k,v in result.items() if k!='parts'}])
    except Exception as error:
        failure=error
        context=f'MQAA ngày {current.isoformat()}' if current<=end else f'Gộp MQAA {first} → {last}'
        manifest.update(status='error',error=f'{context}: {error}',failed_date=current.isoformat() if current<=end else None,
                        failed_plants=[dict(plant=plant,error=str(error))])
    finally:
        try:
            # A failed daily SQL response has also replaced the form.
            if isinstance(failure,downloads.SFCServerTimeout):
                page.goto(url,wait_until='domcontentloaded',timeout=timeout)
            # A client/network timeout can still be pending: do not replay or restore it.
            if failure is None or isinstance(failure,(downloads.SFCServerTimeout,ValueError)):
                downloads.apply_filters(page,filters,timeout)
                verify_scope(page,filters,plant)
        except Exception as error:
            manifest['restore_error']=str(error)
        write_json(folder/'manifest.json',manifest)
    if failure:
        error=RuntimeError(manifest['error']+f'. Đã giữ {len(parts)} file ngày; chưa tạo file gộp.')
        error.downloaded_files=parts
        raise error from failure
    result['message']=f'SFC bị timeout; đã tải {len(parts)} ngày và gộp thành một file MQAA · {plant}.'
    logger.info('MQAA %s: merged %s daily files into %s',plant,len(parts),result['name'])
    return result
