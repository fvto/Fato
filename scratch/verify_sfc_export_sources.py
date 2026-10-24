from pathlib import Path
import sys
sys.stdout.reconfigure(encoding='utf8')
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from collections import Counter
import json,urllib.request,hashlib
from sfc_excel import read_export
root=Path('Report/HFPAExports/20261024_105912_10c9a3')
a_header,a_rows=read_export('Report/HFPAExports/http_pilot/HFPA_VH_NIKE_Jan_week_direct.xls')
parts=root/'_parts/HFPA_VH_NIKE_2026-01-01_2026-01-31'
b=[]
for n in range(1,8):
 h,r=read_export(parts/f'2026-01-{n:02d}.xls');b.extend(r)
assert Counter(a_rows)==Counter(b)
print('Full week vs daily sources: identical',len(b),'records and values')
manifest=json.loads((root/'manifest.json').read_text(encoding='utf8'))
item=manifest['monthly_files'][0]
with urllib.request.urlopen('http://localhost:5050'+item['url'],timeout=20) as response:data=response.read()
assert hashlib.sha256(data).hexdigest()==item['sha256']
print('Dashboard download: HTTP 200, file SHA256 matches, rows',item['rows'])
job_id=Path('scratch/sfc_live_http_job_id.txt').read_text()
with urllib.request.urlopen('http://localhost:5050/api/command/jobs/'+job_id,timeout=10) as response:job=json.load(response)
print('Current job:',job['status'],job['events'][-1]['msg'])
