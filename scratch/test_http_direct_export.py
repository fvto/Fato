from pathlib import Path
source=(Path(__file__).parent/'test_http_export_correct.py').read_text(encoding='utf8')
source=source[:source.index("form.fields['__EVENTTARGET']='ctl00$rpnlMain$ContentPlaceHolder1$btnSearch'")]+"""
form.fields['__EVENTTARGET']='ctl00$rpnlMain$ContentPlaceHolder1$btnExcel';form.fields['__EVENTARGUMENT']=''
started=time.monotonic();print('Direct export Jan VH NIKE',flush=True)
with opener.open(urllib.request.Request(url,data=urllib.parse.urlencode(form.fields).encode()),timeout=900) as response:
 data=response.read();print('Export response',response.status,len(data),round(time.monotonic()-started,1),flush=True)
 folder=root/'Report/HFPAExports/http_pilot';folder.mkdir(exist_ok=True)
 path=folder/'HFPA_VH_NIKE_Jan_direct.xls';path.write_bytes(data)
 from sfc_excel import validate_export,read_export
 print('Validated',validate_export(path,'2026-01-01','2026-01-31',True),flush=True)
 h,rows=read_export(path);print('Groups',sorted(set(r[0] for r in rows)),flush=True)
"""
exec(compile(source,'scratch/test_http_direct_export.py','exec'))
