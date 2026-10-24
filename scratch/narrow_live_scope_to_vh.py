from pathlib import Path
import json,time,urllib.request,subprocess,sys
sys.stdout.reconfigure(encoding='utf8')
root=Path(__file__).resolve().parents[1]
job=(root/'scratch/sfc_live_http_job_id.txt').read_text()
while True:
 with urllib.request.urlopen('http://localhost:5050/api/command/jobs/'+job,timeout=10) as response:state=json.load(response)
 msg=state['events'][-1]['msg']
 if state['status']!='running':raise RuntimeError('Job ended before scope update: '+msg)
 if 'Đang ghép và đối chiếu file tháng VH:' in msg:
  print('Safe scope-change point: daily requests finished; '+msg,flush=True)
  command="$sfcPid = (Get-NetTCPConnection -LocalPort 5050 -State Listen | Select-Object -First 1).OwningProcess; $sfcService = Get-CimInstance Win32_Process -Filter \"ProcessId=$sfcPid\"; if ($sfcService.CommandLine -notmatch 'sfc_dashboard_server.py') { throw 'Unexpected process' }; Stop-Process -Id $sfcPid"
  subprocess.run(['powershell','-NoProfile','-Command',command],check=True,capture_output=True)
  folder=root/'Report/HFPAExports/20261024_105912_10c9a3'
  manifest=json.loads((folder/'manifest.json').read_text(encoding='utf8'))
  manifest.update(status='interrupted',error='Phạm vi đổi sang chỉ VH theo yêu cầu người dùng.',paused_plants=['VH2','JV','JV2'])
  (folder/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf8')
  prompt='Tải excel HFPA từ tháng 1 - tháng 10 năm 2026 nhà máy VH, chia nhỏ file từng tháng sau đó combine lại thành 1 file, phải xuất Nike (external)'
  (root/'scratch/sfc_vh_prompt_test.txt').write_text(prompt,encoding='utf8')
  print('VH sources preserved. VH2/JV/JV2 paused before any factory switch.',flush=True)
  break
 time.sleep(.25)
