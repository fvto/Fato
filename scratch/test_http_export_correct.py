from pathlib import Path
import sys,time,urllib.parse,re
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
exec((Path(__file__).parent/'inspect_sfc_http_form.py').read_text(encoding='utf-8').split('with opener.open')[0])
class CleanForm(Form):
 def handle_starttag(self,tag,attrs):
  a=dict(attrs)
  if tag=='input' and a.get('type','').lower() in ('submit','button','image'):return
  super().handle_starttag(tag,attrs)
url='https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504'
with opener.open(url,timeout=20) as response:body=response.read().decode('utf8')
form=CleanForm();form.feed(body)
from datetime import datetime,timezone
for k in list(form.fields):
 if 'DateFrom' in k or 'DateTo' in k:
  day='2026-01-01' if 'DateFrom' in k else '2026-01-31'
  if k.endswith('_Raw'):form.fields[k]=str(int(datetime.fromisoformat(day).replace(tzinfo=timezone.utc).timestamp()*1000))
  elif k.endswith('$dtDateFrom') or k.endswith('$dtDateTo'):form.fields[k]=day
form.fields['ctl00$rpnlMain$ContentPlaceHolder1$ckbNIKE']='C'
form.fields['__EVENTTARGET']='ctl00$rpnlMain$ContentPlaceHolder1$btnSearch'
form.fields['__EVENTARGUMENT']=''
started=time.monotonic();print('Searching Jan VH NIKE',flush=True)
with opener.open(urllib.request.Request(url,data=urllib.parse.urlencode(form.fields).encode()),timeout=900) as response:
 body=response.read().decode('utf8');print('Search returned',round(time.monotonic()-started,1),'seconds',len(body),'bytes',flush=True)
 form=CleanForm();form.feed(body)
 print('Returned dates/NIKE:',{k:v for k,v in form.fields.items() if k.endswith(('$dtDateFrom','$dtDateTo','$ckbNIKE'))},flush=True)
 form.fields['__EVENTTARGET']='ctl00$rpnlMain$ContentPlaceHolder1$btnExcel';form.fields['__EVENTARGUMENT']=''
 started=time.monotonic();print('Exporting searched data',flush=True)
 with opener.open(urllib.request.Request(url,data=urllib.parse.urlencode(form.fields).encode()),timeout=900) as response:
  data=response.read();print('Export returned',round(time.monotonic()-started,1),response.headers.get('Content-Type'),len(data),flush=True)
  folder=root/'Report/HFPAExports/http_pilot';folder.mkdir(exist_ok=True)
  if data[:8]==b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' or data[:2]==b'PK':
   path=folder/('HFPA_VH_NIKE_2026-01_correct.xls' if data[:2]!=b'PK' else 'HFPA_VH_NIKE_2026-01_correct.xlsx');path.write_bytes(data)
   from sfc_excel import validate_export,read_export
   print('VALIDATED rows:',validate_export(path,'2026-01-01','2026-01-31',True),flush=True)
   h,rows=read_export(path);print('Factory groups:',sorted(set(str(r[0]) for r in rows)),flush=True)
  else:
   (root/'scratch/http_pilot_response.html').write_bytes(data);print('HTML returned',flush=True)
