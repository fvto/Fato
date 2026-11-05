from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

import openpyxl

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import sfc_hfpa_acquisition as acquisition
import sfc_report_tools as reports


def raw_file(path, group, day='2026-09-10', station='HFPA', nike=False):
    path.parent.mkdir(parents=True,exist_ok=True)
    book=openpyxl.Workbook()
    if group=='ftt':
        book.active.append(['InspDate','Line','ShoeName','InspQty','FailQty','Station','Issues','IssueQty'])
        if day: book.active.append([datetime.fromisoformat(day),'Line 1','Model',100,2,station,'Defect',2])
    else:
        book.active.append(['Factory Group','Audit Date (dd-mmm-yyyy)','Internal AUDIT ID','Model','Line'] if nike else
                           ['Factory Group','Audit Date (dd-mmm-yyyy)','Internal AUDIT ID','SO#','Country','Pass/Fail'])
        if day: book.active.append(['CL_VH',datetime.fromisoformat(day),'audit-1','SO-1','VN','Pass'])
    book.save(path)
    book.close()


class HFPAcquisitionTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.root=Path(self.temp.name)
        self.exports=self.root/'exports';self.exports.mkdir()
        self.runs=self.root/'runs';self.runs.mkdir()
        self.engines=self.root/'engines';(self.engines/'hfpa').mkdir(parents=True)
        self.patches=[patch.object(reports,'EXPORT_ROOT',self.exports),patch.object(reports,'RUN_ROOT',self.runs),
                      patch.object(reports,'TOOLS_ROOT',self.engines),patch.object(reports,'original_root',return_value=self.root/'original')]
        for p in self.patches:p.start()

    def tearDown(self):
        for p in reversed(self.patches):p.stop()
        self.temp.cleanup()

    def test_reuses_only_hfpa_station_and_internal_mes410_in_month(self):
        raw_file(self.exports/'QAStation_2026-09-01_2026-09-30-VH.xlsx','ftt')
        raw_file(self.exports/'QAStation_2026-09-01_2026-09-30-JV.xlsx','ftt',station='FTT')
        raw_file(self.exports/'QAStation_2026-09-01_2026-09-20-VH2.xlsx','ftt')
        raw_file(self.exports/'FTT-Internal_2026-09-01_2026-09-30-JV2.xlsx','ftt')
        raw_file(self.exports/'HFPA_VH_Internal_2026-09-01_2026-09-30-VH.xlsx','hfpa',nike=True)
        raw_file(self.exports/'HFPA_VH2_Internal_2026-09-01_2026-09-30-VH2.xlsx','hfpa',day='2026-10-01')
        raw_file(self.root/'original/Input/HFPA/HFPA 2026-10-03 11_29_50-JV.xlsx','hfpa')
        plan=acquisition.plan('2026-09')
        self.assertEqual(plan['total'],8)
        self.assertEqual({(i['group'],i['plant']) for i in plan['available']},{('ftt','VH'),('hfpa','JV')})
        self.assertEqual(len(plan['missing']),6)
        self.assertFalse(plan['workflow']['hfpa']['nike'])
        self.assertEqual(plan['workflow']['ftt']['stations'],['HFPA'])

    def test_valid_empty_month_is_cached_without_browser_or_download_loop(self):
        window=dict(date_from='2026-09-01',date_to='2026-09-30')
        for group in ('ftt','hfpa'):
            for plant in acquisition.PLANTS:
                raw_file(self.exports/acquisition.filename(group,plant,window,'.xlsx'),group,day=None)
        plan=acquisition.plan('2026-09')
        self.assertEqual(len(plan['available']),8)
        self.assertEqual(plan['missing'],[])
        fake=MagicMock()
        acquisition.acquire(self.root,{'acquisition':plan},MagicMock(),fake)
        fake.sync_playwright.assert_not_called()
        fake.apply_sfc_filters.assert_not_called()

    def test_auto_start_accepts_missing_input_and_reserves_browser(self):
        with patch.object(reports.subprocess,'Popen'):
            run=reports.start_run('hfpa','2026-09','all',[],acquire_missing=True)
        self.assertEqual(len(run['acquisition']['missing']),8)
        self.assertEqual(run['phase'],'acquisition')
        self.assertTrue(reports.browser_reserved())

    def fake_controller(self):
        page=SimpleNamespace(url=acquisition.HFPA_URL,context=SimpleNamespace(cookies=lambda:[{'name':'MyFactoryName','value':'VH'}]))
        browser=SimpleNamespace(close=lambda:None)
        @contextmanager
        def client():
            yield SimpleNamespace(chromium=SimpleNamespace(connect_over_cdp=lambda *a,**k:browser))
        fake=SimpleNamespace(sync_playwright=client,CDP_URL='test',SFC_CONNECT_WAIT_MS=1000,SFC_WAIT_MS=1000,
            get_sfc_page=lambda b:page,sfc_downloads=SimpleNamespace(wait_idle=lambda *a:None,
            factory_code=lambda x:x,capture_filters=lambda p:{}),SEVEN_FACTORIES=list(acquisition.PLANTS),
            switch_factory=MagicMock(return_value=dict(success=True)),navigate_to=MagicMock(return_value=dict(success=True)))
        def export(action,target_id=None,params=None):
            if action=='restore_filters':return dict(success=True)
            path=self.exports/f'QAStation_2026-09-01_2026-09-30-{params["_expected_plant"]}.xlsx'
            raw_file(path,'ftt')
            return dict(success=True,downloaded_file={'path':str(path)})
        def mes(**params):
            path=self.exports/f'raw-{params["expected_plant"]}.xlsx'
            raw_file(path,'hfpa')
            return dict(success=True,file={'path':str(path)})
        fake._execute_dynamic_action=MagicMock(side_effect=export)
        fake.apply_sfc_filters=MagicMock(side_effect=mes)
        return fake

    def test_downloads_four_station_ok_and_four_mes410_nike_off_and_reuses_them(self):
        folder=self.runs/('d'*32);(folder/'workspace').mkdir(parents=True)
        fake=self.fake_controller()
        run=dict(period='2026-09',inputs=[],acquisition=acquisition.plan('2026-09'))
        acquisition.acquire(folder,run,lambda message:None,fake)
        calls=[c for c in fake._execute_dynamic_action.call_args_list if c.args[0]!='restore_filters']
        self.assertEqual(len(calls),4)
        for call in calls:
            params=call.kwargs['params']
            self.assertEqual(params['stations'],['HFPA'])
            self.assertEqual(params['report_type'],'station')
            self.assertEqual(params['date_to'],'2026-09-30')
        self.assertEqual(fake.apply_sfc_filters.call_count,4)
        self.assertEqual([c.kwargs['expected_plant'] for c in fake.apply_sfc_filters.call_args_list],list(acquisition.PLANTS))
        for call in fake.apply_sfc_filters.call_args_list:
            self.assertIs(call.kwargs['nike'],False)
            self.assertEqual(call.kwargs['factory'],'')
            self.assertEqual(call.kwargs['action'],'excel')
        self.assertEqual(len(run['inputs']),8)
        self.assertEqual(len(acquisition.plan('2026-09')['available']),8)
        self.assertIn(acquisition.HFPA_URL,[c.args[0] for c in fake.navigate_to.call_args_list])

    def test_timeout_keeps_completed_downloads_and_does_not_reconnect_or_restore(self):
        folder=self.runs/('e'*32);(folder/'workspace').mkdir(parents=True)
        fake=self.fake_controller()
        fake.apply_sfc_filters=MagicMock(return_value=dict(success=False,timed_out=True,error='SFC vẫn đang xử lý sau 15 phút.'))
        run=dict(period='2026-09',inputs=[],acquisition=acquisition.plan('2026-09'))
        with self.assertRaisesRegex(RuntimeError,'15 phút'):
            acquisition.acquire(folder,run,lambda message:None,fake)
        fake.apply_sfc_filters.assert_called_once()
        self.assertEqual(len(run['inputs']),4)
        self.assertEqual(len(run['acquisition']['not_attempted']),3)
        self.assertEqual(len(acquisition.plan('2026-09')['available']),4)
        self.assertFalse(any(c.args[0]=='restore_filters' for c in fake._execute_dynamic_action.call_args_list))


if __name__=='__main__':unittest.main()
