from pathlib import Path
import sys,urllib.parse,time
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
exec((Path(__file__).parent/'inspect_sfc_http_form.py').read_text(encoding='utf-8').split('for k,v in form.fields.items():')[0])
from datetime import datetime,timezone
for k in list(form.fields):
 if 'DateFrom' in k or 'DateTo' in k:
  day='2026-01-01' if 'DateFrom' in k else '2026-01-31'
  if k.endswith('_Raw'):form.fields[k]=str(int(datetime.fromisoformat(day).replace(tzinfo=timezone.utc).timestamp()*1000))
  elif k.endswith('$dtDateFrom') or k.endswith('$dtDateTo'):form.fields[k]=day
form.fields['ctl00$rpnlMain$ContentPlaceHolder1$ckbNIKE']='C'
form.fields['__EVENTTARGET']='ctl00$rpnlMain$ContentPlaceHolder1$btnExcel'
form.fields['__EVENTARGUMENT']=''
started=time.monotonic()
request=urllib.request.Request('https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504',data=urllib.parse.urlencode(form.fields).encode(),headers={'Referer':'https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504'})
with opener.open(request,timeout=900) as response:
 data=response.read();print('Export:',response.status, 'seconds',round(time.monotonic()-started,1),'type',response.headers.get('Content-Type'),'disposition',response.headers.get('Content-Disposition'),'bytes',len(data),flush=True)
 folder=root/'Report/HFPAExports/http_pilot';folder.mkdir(exist_ok=True)
 if data[:8]==b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' or data[:2]==b'PK':
  path=folder/('HFPA_VH_NIKE_2026-01.xls' if data[:2]!=b'PK' else 'HFPA_VH_NIKE_2026-01.xlsx');path.write_bytes(data)
  from sfc_excel import validate_export,read_export
  print('VALIDATED rows:',validate_export(path,'2026-01-01','2026-01-31',True))
  h,rows=read_export(path);print('Factory groups:',sorted(set(str(r[0]) for r in rows)));print('Headers:',h)
 else:
  # Only save application markup locally, never session cookies.
  (root/'scratch/http_pilot_response.html').write_bytes(data)
  print('Response is HTML, saved for form inspection')
