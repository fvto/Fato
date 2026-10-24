from pathlib import Path
import urllib.request,json,time,sys
sys.stdout.reconfigure(encoding='utf8')
prompt=Path('scratch/sfc_prompt_test.txt').read_text(encoding='utf8').strip()
request=urllib.request.Request('http://localhost:5050/api/command/stream',data=json.dumps({'prompt':prompt}).encode(),headers={'Content-Type':'application/json'})
items=[];started=time.monotonic()
with urllib.request.urlopen(request,timeout=3600) as response:
 for raw in response:
  line=raw.decode('utf8').strip()
  if not line.startswith('data:'):continue
  item=json.loads(line[5:]);items.append(item)
  Path('scratch/sfc_live_http_command_test.json').write_text(json.dumps(items,ensure_ascii=False,indent=2),encoding='utf8')
  if item.get('job_id'):Path('scratch/sfc_live_http_job_id.txt').write_text(item['job_id'])
  print(round(time.monotonic()-started,1),item.get('status'),item.get('msg'),flush=True)
  if item.get('status') in ('done','error'):break
