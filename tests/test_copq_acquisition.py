from contextlib import contextmanager
import os
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

import openpyxl

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import sfc_copq_acquisition as acquisition
import sfc_report_tools as reports
import sfc_dashboard_server as server


def raw_file(path, group):
    path.parent.mkdir(parents=True,exist_ok=True)
    book=openpyxl.Workbook()
    book.active.append(['Date','Category','Type','Factory','Ttl cost ($)'] if group=='copq' else
                       ['Date','Factory','Metric','Number of pairs','Defect Issues','SO#'])
    book.save(path)
    book.close()


class AcquisitionTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.root=Path(self.temp.name)
        self.exports=self.root/'exports';self.exports.mkdir()
        self.runs=self.root/'runs';self.runs.mkdir()
        self.engines=self.root/'engines';(self.engines/'copq').mkdir(parents=True)
        self.patches=[patch.object(reports,'EXPORT_ROOT',self.exports),patch.object(reports,'RUN_ROOT',self.runs),
                      patch.object(reports,'TOOLS_ROOT',self.engines),patch.object(reports,'original_root',return_value=self.root/'original')]
        for p in self.patches:p.start()

    def tearDown(self):
        for p in reversed(self.patches):p.stop()
        self.temp.cleanup()

    def test_scope_month_and_cache_require_full_unfiltered_month(self):
        self.assertEqual(acquisition.month_window('2024-02'),('2024-02-01','2024-02-29'))
        raw_file(self.exports/'Cost Of Poor Quality 2026-09-01_2026-09-30-VH.xlsx','copq')
        raw_file(self.exports/'Cost Of Poor Quality Defective Return 2026-09-01_2026-09-30-JV.xlsx','copq')
        raw_file(self.exports/'Cost Of Poor Quality 2026-09-01_2026-09-20-VH2.xlsx','copq')
        raw_file(self.exports/'FTT-Internal_2026-09-01_2026-09-30-VH.xlsx','ftt')
        plan=acquisition.plan('2026-09')
        self.assertEqual(plan['total'],11)
        self.assertEqual(len(plan['available']),2)
        self.assertEqual(len(plan['missing']),9)
        self.assertEqual(len(acquisition.plan('2026-09','combine')['missing']),6)
        self.assertEqual(len(acquisition.plan('2026-09','ftt')['missing']),3)

    def test_empty_data_can_be_reused_and_never_triggers_download_loop(self):
        for group,plants in acquisition.required_inputs('all').items():
            for plant in plants:
                prefix='Cost Of Poor Quality ' if group=='copq' else 'FTT-Internal_'
                raw_file(self.exports/f'{prefix}2026-09-01_2026-09-30-{plant}.xlsx',group)
        plan=acquisition.plan('2026-09')
        self.assertEqual(len(plan['available']),11)
        self.assertEqual(plan['missing'],[])
        fake=MagicMock()
        acquisition.acquire(self.root,{'acquisition':dict(total=11,missing=[])},MagicMock(),fake)
        fake.sync_playwright.assert_not_called()
        fake._execute_dynamic_action.assert_not_called()

    def test_auto_start_needs_no_manual_input_and_reserves_browser(self):
        with patch.object(reports.subprocess,'Popen'):
            run=reports.start_run('copq','2026-09','all',[],acquire_missing=True)
        self.assertEqual(run['phase'],'acquisition')
        self.assertEqual(len(run['acquisition']['missing']),11)
        self.assertTrue(reports.browser_reserved())
        with server.app.test_client() as client:
            response=client.post('/api/page/execute',json={'action_type':'set_dates'})
            self.assertEqual(response.status_code,409)
        run.update(status='running',pid=os.getpid())
        reports.write_json(reports.run_folder(run['id'])/'run.json',run)
        self.assertFalse(reports.browser_reserved(exclude_pid=os.getpid()))
        run['phase']='processing'
        reports.write_json(reports.run_folder(run['id'])/'run.json',run)
        self.assertFalse(reports.browser_reserved())

    def test_download_contract_is_seven_copq_and_four_internal_with_zero_stations(self):
        folder=self.runs/('c'*32);(folder/'workspace').mkdir(parents=True)
        page=SimpleNamespace(url=acquisition.COPQ_URL,context=SimpleNamespace(cookies=lambda:[{'name':'MyFactoryName','value':'VH'}]))
        browser=SimpleNamespace(close=lambda:None)
        @contextmanager
        def client():
            yield SimpleNamespace(chromium=SimpleNamespace(connect_over_cdp=lambda *a,**k:browser))
        calls=[]
        def export(action,target_id=None,params=None):
            if action=='restore_filters':return dict(success=True)
            group='copq' if action=='click_button' else 'ftt'
            plant=params['_expected_plant']
            calls.append((group,plant,params))
            prefix='Cost Of Poor Quality ' if group=='copq' else 'FTT-Internal_'
            path=self.exports/f'{prefix}2026-09-01_2026-09-30-{plant}.xlsx'
            raw_file(path,group)
            return dict(success=True,downloaded_file={'path':str(path)})
        fake=SimpleNamespace(sync_playwright=client,CDP_URL='test',SFC_CONNECT_WAIT_MS=1000,SFC_WAIT_MS=1000,
            get_sfc_page=lambda b:page,sfc_downloads=SimpleNamespace(wait_idle=lambda *a:None,
            factory_code=lambda x:x,capture_filters=lambda p:{}),SEVEN_FACTORIES=list(acquisition.COPQ_PLANTS),
            switch_factory=lambda p:dict(success=True),navigate_to=lambda url:dict(success=True),_execute_dynamic_action=export)
        plan=acquisition.plan('2026-09')
        run=dict(period='2026-09',inputs=[],acquisition=plan)
        acquisition.acquire(folder,run,lambda message:None,fake)
        self.assertEqual(len(calls),11)
        self.assertEqual([plant for group,plant,_ in calls if group=='ftt'],['VH','VH2','JV','JV2'])
        for group,plant,params in calls:
            self.assertEqual(params['date_from'],'2026-09-01')
            self.assertEqual(params['date_to'],'2026-09-30')
            if group=='ftt':
                self.assertEqual(params['stations'],[])
                self.assertEqual(params['station_selection'],'none')
                self.assertEqual(params['report_type'],'internal')
            else:
                self.assertTrue(all(item['value']=='' for item in params['selections'].values()))
        self.assertEqual(len(acquisition.plan('2026-09')['available']),11)
        victim=self.exports/'FTT-Internal_2026-09-01_2026-09-30-VH.xlsx'
        victim.unlink()
        retry=acquisition.plan('2026-09')
        self.assertEqual(len(retry['missing']),1)
        fake._execute_dynamic_action=MagicMock(return_value=dict(success=False,error='Timeout waiting for download'))
        run=dict(period='2026-09',inputs=[],acquisition=retry)
        with self.assertRaisesRegex(RuntimeError,'Timeout'):
            acquisition.acquire(folder,run,lambda message:None,fake)
        fake._execute_dynamic_action.assert_called_once()
        self.assertEqual(len(acquisition.plan('2026-09')['available']),10)


if __name__=='__main__':unittest.main()
