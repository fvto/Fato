from pathlib import Path
import sys,json,time,hashlib,urllib.request,calendar
from datetime import date,datetime
from collections import Counter
import openpyxl
sys.stdout.reconfigure(encoding='utf8')
root=Path(__file__).resolve().parents[1]
folder=root/'Report/HFPAExports/20261024_105912_10c9a3'
while True:
 manifest=json.loads((folder/'manifest.json').read_text(encoding='utf8'))
 if manifest['status']!='running':break
 time.sleep(3)
assert manifest['status']=='done',manifest.get('error')
assert manifest['intent']['target_plants']==['VH']
assert set(manifest['paused_plants'])=={'VH2','JV','JV2'}
months=manifest['monthly_files'];combined=manifest['combined_files']
assert len(months)==10 and len(combined)==1
for month,item in enumerate(months,1):
 assert item['plant']=='VH' and item['nike'] is True
 assert item['date_from']==f'2026-{month:02d}-01'
 assert item['date_to']==f'2026-{month:02d}-{calendar.monthrange(2026,month)[1]}'
 assert hashlib.sha256(Path(item['path']).read_bytes()).hexdigest()==item['sha256']
 parts=item['parts'];assert len(parts)==calendar.monthrange(2026,month)[1]
 assert [p['date_from'] for p in parts]==[f'2026-{month:02d}-{d:02d}' for d in range(1,len(parts)+1)]
 assert all(p['date_from']==p['date_to'] and p['plant']=='VH' and p['nike'] is True for p in parts)
 assert sum(p['rows'] for p in parts)==item['rows']
 for p in parts:assert hashlib.sha256(Path(p['path']).read_bytes()).hexdigest()==p['sha256']
 print('Verified month',month,'rows',item['rows'],flush=True)
item=combined[0];expected=sum(p['rows'] for p in months)
assert item['rows']==expected
book=openpyxl.load_workbook(item['path'],read_only=True,data_only=False)
counts=Counter();groups=set();lo=None;hi=None;total=0
try:
 iterator=book['HFPA'].iter_rows(values_only=True);header=next(iterator)
 col=next(i for i,v in enumerate(header) if 'audit date' in str(v).lower())
 assert 'Factory Group' in header[0]
 for row in iterator:
  raw=row[col]
  day=raw.date() if isinstance(raw,datetime) else raw if isinstance(raw,date) else datetime.strptime(raw,'%d-%b-%Y').date()
  assert date(2026,1,1)<=day<=date(2026,10,31)
  counts[day.month]+=1;groups.add(row[0]);lo=day if lo is None else min(lo,day);hi=day if hi is None else max(hi,day);total+=1
 assert total==expected and groups<={'CL_VH','CL_VB'}
 assert counts==Counter({i+1:m['rows'] for i,m in enumerate(months)})
 audit=list(book['NguonTheoThang'].iter_rows(min_row=2,values_only=True))
 assert len(audit)==10 and sum(r[5] for r in audit)==total
 assert [r[4] for r in audit]==[m['name'] for m in months]
finally:book.close()
hash_download=hashlib.sha256()
with urllib.request.urlopen('http://localhost:5050'+item['url'],timeout=60) as response:
 while chunk:=response.read(1024*1024):hash_download.update(chunk)
assert hash_download.hexdigest()==item['sha256']
jobid=(root/'scratch/sfc_live_http_job_id.txt').read_text()
with urllib.request.urlopen('http://localhost:5050/api/command/jobs/'+jobid,timeout=10) as response:job=json.load(response)
assert job['status']=='done' and job['events'][-1]['completed']==10
report=dict(status='verified',plant='VH',nike=True,monthly_files=10,combined_files=1,rows=total,actual_date_min=str(lo),actual_date_max=str(hi),factory_groups=sorted(groups),paused_plants=manifest['paused_plants'],rows_by_month=dict(sorted(counts.items())),download_sha256_verified=True,combined_path=item['path'])
(folder/'validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps(report,ensure_ascii=False,indent=2),flush=True)
