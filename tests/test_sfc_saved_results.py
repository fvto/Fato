import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import sfc_command_jobs as jobs
import sfc_commands as commands
import sfc_excel as excel
from test_sfc_commands import fixture

class SavedDashboardResults(unittest.TestCase):
 def test_status_poll_does_not_wait_for_page_or_request_screenshots(self):
  import sfc_dashboard_server as server
  with patch.object(jobs,'running',return_value=False),patch.object(server.sfc_controller,'get_connection_status',return_value={'connected':True,'current_plant':'VH'}) as connection,patch.object(server.sfc_controller,'get_live_status',return_value={'connected':True,'screenshot':'explicit-image'}) as screen:
   with server.app.test_client() as client:
    self.assertEqual(client.get('/api/status').json['current_plant'],'VH')
    connection.assert_called_once();screen.assert_not_called()
    self.assertEqual(client.get('/api/status?screen=1').json['screenshot'],'explicit-image')
    screen.assert_called_once()
 def test_restart_restores_completed_result_without_exporting_again(self):
  with tempfile.TemporaryDirectory() as directory,patch.object(commands,'EXPORT_ROOT',Path(directory)),patch.object(excel,'EXPORT_ROOT',Path(directory)),patch.dict(jobs.JOBS,{},clear=True):
   root=Path(directory);folder=root/'completed_vh';folder.mkdir();path=folder/'VH.xlsx';fixture(path,'2026-10-24','VH')
   item=excel.file_info(path,1)
   intent=commands.parse_intent('Xuất excel HFPA VH ngày 2026-10-24 NIKE')
   manifest=dict(status='done',prompt='Completed VH request',intent=intent,monthly_files=[item],combined_files=[item])
   (folder/'manifest.json').write_text(json.dumps(manifest),encoding='utf8')
   (root/'last_prompt.txt').write_text('Stale prompt',encoding='utf8')
   with patch.object(jobs,'start',side_effect=AssertionError('Must not rerun exports')):
    job_id=jobs.latest_job_id();self.assertEqual(job_id,'saved_completed_vh')
    self.assertEqual(jobs.latest_prompt(),'Completed VH request')
    job=jobs.snapshot(job_id);self.assertEqual(job['status'],'done')
    self.assertEqual(job['events'][-1]['files'][0]['url'],item['url'])
    self.assertIn('done',''.join(jobs.stream(job_id)))
 def test_results_exclude_partial_failed_and_outside_files(self):
  with tempfile.TemporaryDirectory() as directory,patch.object(commands,'EXPORT_ROOT',Path(directory)/'exports'):
   root=commands.EXPORT_ROOT;root.mkdir();outside=Path(directory)/'outside.xlsx';outside.write_bytes(b'no')
   for index,status in enumerate(['done','error','running']):
    folder=root/str(index);folder.mkdir();path=folder/'report.partial.xlsx';path.write_bytes(b'no')
    items=[dict(path=str(path)),dict(path=str(outside))]
    (folder/'manifest.json').write_text(json.dumps(dict(status=status,prompt='x',intent={},monthly_files=items)),encoding='utf8')
   self.assertEqual(jobs.saved_results(),[])
 def test_dashboard_shows_saved_downloads_and_resume_id(self):
  import sfc_dashboard_server as server
  with tempfile.TemporaryDirectory() as directory,patch.object(commands,'EXPORT_ROOT',Path(directory)),patch.object(excel,'EXPORT_ROOT',Path(directory)),patch.dict(jobs.JOBS,{},clear=True):
   folder=Path(directory)/'vh';folder.mkdir();path=folder/'VH.xlsx';fixture(path,'2026-10-24','VH');item=excel.file_info(path,1)
   (folder/'manifest.json').write_text(json.dumps(dict(status='done',prompt='VH only',intent={'target_plants':['VH']},monthly_files=[item],combined_files=[item])),encoding='utf8')
   with server.app.test_client() as client:
    response=client.get('/');self.assertEqual(response.status_code,200)
    html=response.get_data(as_text=True)
    self.assertIn('Kết quả đã lưu',html);self.assertIn(item['url'],html);self.assertIn('data-last-job="saved_vh"',html)
    self.assertIn('class="prompt-input" value=""',html)
    self.assertNotIn('id="saved-command-results" style="margin-top:16px;" open',html)
    response=client.get('/api/command/jobs/saved_vh/stream')
    self.assertIn('"status": "done"',response.get_data(as_text=True))

if __name__=='__main__':unittest.main()
