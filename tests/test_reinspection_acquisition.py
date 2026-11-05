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
import sfc_reinspection_acquisition as acquisition
import sfc_report_tools as reports


def raw_file(path, group, day='2026-09-10', nike=False):
    path.parent.mkdir(parents=True,exist_ok=True)
    book=openpyxl.Workbook()
    if group=='reinspection':
        book.active.append(['Category','SO#','Country','Model 2',"Reinspection Q'ty",'2nd Audit Sample Size',
                            '2nd Pass/Fail','Remark','Audit Date (dd-mmm-yyyy)'])
        if day: book.active.append(['QC','SO-1','VN','Model',100,12,'Pass','',datetime.fromisoformat(day)])
    elif group=='ftt':
        book.active.append(['Date','Factory','Metric','Number of pairs','Defect Issues','SO#'])
        if day: book.active.append([datetime.fromisoformat(day),'VH','FTT',100,'Defect','SO-1'])
    else:
        book.active.append(['Factory Group','Audit Date (dd-mmm-yyyy)','Internal AUDIT ID','Model','Line'] if nike else
                           ['Factory Group','Audit Date (dd-mmm-yyyy)','Internal AUDIT ID','SO#','Country','Pass/Fail'])
        if day: book.active.append(['CL_VH',datetime.fromisoformat(day),'audit-1','SO-1','VN','Pass'])
    book.save(path);book.close()


class ReinspectionAcquisitionTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name)
        self.exports=self.root/'exports';self.exports.mkdir()
        self.runs=self.root/'runs';self.runs.mkdir()
        self.engines=self.root/'engines';(self.engines/'reinspection').mkdir(parents=True)
        self.patches=[patch.object(reports,'EXPORT_ROOT',self.exports),patch.object(reports,'RUN_ROOT',self.runs),
                      patch.object(reports,'TOOLS_ROOT',self.engines),patch.object(reports,'original_root',return_value=self.root/'original')]
        for p in self.patches:p.start()

    def tearDown(self):
        for p in reversed(self.patches):p.stop()
        self.temp.cleanup()

    def test_contract_is_four_plants_three_sources_and_same_full_month(self):
        plan=acquisition.plan('2024-02')
        self.assertEqual(plan['date_to'],'2024-02-29')
        self.assertEqual(plan['total'],12)
        self.assertEqual(len(plan['missing']),12)
        for source in plan['workflow'].values():self.assertEqual(source['plants'],['VH','VH2','JV','JV2'])
        self.assertEqual(plan['workflow']['reinspection']['url'],acquisition.REINSPECTION_URL)
        self.assertIsNone(plan['workflow']['reinspection']['line'])
        self.assertIsNone(plan['workflow']['reinspection']['type'])
        self.assertEqual(plan['workflow']['ftt']['stations'],[])
        self.assertEqual(plan['workflow']['ftt']['report_type'],'internal')
        self.assertIs(plan['workflow']['hfpa']['nike'],False)

    def test_cache_rejects_filtered_wrong_schema_and_rows_outside_month(self):
        raw_file(self.exports/'Re-Inspection 2026-09-01_2026-09-30-VH.xlsx','reinspection')
        raw_file(self.exports/'Re-Inspection QC 2026-09-01_2026-09-30-JV.xlsx','reinspection')
        raw_file(self.exports/'Re-Inspection 2026-09-01_2026-09-20-VH2.xlsx','reinspection')
        raw_file(self.exports/'Re-Inspection 2026-09-01_2026-09-30-JV2.xlsx','ftt')
        raw_file(self.exports/'FTT-Internal_2026-09-01_2026-09-30-VH.xlsx','reinspection')
        raw_file(self.exports/'FTT-Internal_2026-09-01_2026-09-30-JV.xlsx','ftt',day='2026-08-31')
        raw_file(self.exports/'HFPA_VH_Internal_2026-09-01_2026-09-30-VH.xlsx','hfpa',nike=True)
        raw_file(self.root/'original/HFPA/HFPA 2026-10-03 11_29_50-JV.xlsx','hfpa')
        plan=acquisition.plan('2026-09')
        self.assertEqual({(i['group'],i['plant']) for i in plan['available']},{('reinspection','VH'),('hfpa','JV')})
        self.assertEqual(len(plan['missing']),10)

    def test_empty_valid_files_are_reused_without_browser(self):
        window=dict(date_from='2026-09-01',date_to='2026-09-30')
        for group in ('reinspection','ftt','hfpa'):
            for plant in acquisition.PLANTS:
                raw_file(self.exports/acquisition.filename(group,plant,window,'.xlsx'),group,day=None)
        plan=acquisition.plan('2026-09')
        self.assertEqual(len(plan['available']),12)
        self.assertEqual(plan['missing'],[])
        fake=MagicMock()
        acquisition.acquire(self.root,{'acquisition':plan},MagicMock(),fake)
        fake.sync_playwright.assert_not_called()
        fake._execute_dynamic_action.assert_not_called()

    def test_devexpress_trailing_count_is_not_an_audit_date(self):
        path=self.exports/'Re-Inspection 2026-09-01_2026-09-30-VH.xlsx'
        raw_file(path,'reinspection')
        book=openpyxl.load_workbook(path)
        book.active.append(['','','','','','','','','1'])
        book.save(path);book.close()
        self.assertEqual(len(acquisition.plan('2026-09')['available']),1)
        # A numeric value in the middle is invalid data, not a footer to skip.
        book=openpyxl.load_workbook(path)
        book.active.append(['QC','SO-2','VN','Model',10,2,'Fail','','2026-09-11'])
        book.save(path);book.close()
        self.assertEqual(acquisition.plan('2026-09')['available'],[])

    def test_legacy_numeric_footer_after_excel_save_keeps_file_selected(self):
        path = self.exports / 'Re-Inspection 2026-09-01_2026-09-30-JV2.xls'
        path.touch()
        header = ['Category', 'SO#', 'Country', 'Model 2', "Reinspection Q'ty",
                  '2nd Audit Sample Size', '2nd Pass/Fail', 'Remark',
                  'Audit Date (dd-mmm-yyyy)', "Defect Q'ty (Total Defect Q'ty)"]
        rows = [
            ['QC', 'SO-1', 'VN', '', 100, 12, 'Pass', '', datetime(2026, 9, 10), 2.0],
            ['', '', '', '', '', '', '', '', 1.0, 2.0],
        ]
        with patch.object(acquisition, 'read_export', return_value=(header, rows)):
            plan = acquisition.plan('2026-09')
        self.assertEqual([item['name'] for item in plan['available']], [path.name])

    def test_reinspection_keeps_original_audit_dates_before_export_period(self):
        raw_file(self.exports/'Re-Inspection 2026-09-01_2026-09-30-VH.xlsx','reinspection',day='2026-08-31')
        self.assertEqual(len(acquisition.plan('2026-09')['available']),1)
        self.assertEqual(acquisition.plan('2026-08')['available'],[])

    def test_reinspection_preserves_identified_record_with_missing_original_audit_date(self):
        path=self.exports/'Re-Inspection 2026-09-01_2026-09-30-VH.xlsx'
        raw_file(path,'reinspection')
        book=openpyxl.load_workbook(path)
        book.active.append(['QC','SO-2','USA','',None,72,'','',''])
        book.active.append(['','','','','','','','','2'])
        book.save(path);book.close()
        original=path.read_bytes()
        self.assertEqual(len(acquisition.plan('2026-09')['available']),1)
        self.assertEqual(acquisition.plan('2026-08')['available'],[])
        self.assertEqual(path.read_bytes(),original)
        _,rows=acquisition.read_export(path)
        self.assertEqual(rows[1][1],'SO-2')
        self.assertIsNone(rows[1][8])

    def test_missing_audit_date_requires_reinspection_record_identity(self):
        for category,so in (('QC',''),('','SO-2')):
            with self.subTest(category=category,so=so):
                path=self.exports/f'identity-{category}-{so}.xlsx'
                raw_file(path,'reinspection')
                book=openpyxl.load_workbook(path)
                book.active.append([category,so,'USA','Model',100,12,'Pass','',''])
                book.save(path);book.close()
                self.assertFalse(acquisition.validate_file(path,'reinspection','2026-09-01','2026-09-30'))

    def test_missing_or_malformed_dates_still_reject_other_data(self):
        for group,value in (('ftt',None),('reinspection','invalid-date'),('reinspection',42)):
            with self.subTest(group=group,value=value):
                path=self.exports/f'{group}-{value}.xlsx'
                raw_file(path,group)
                book=openpyxl.load_workbook(path)
                book.active.cell(2,9 if group=='reinspection' else 1).value=value
                book.active.cell(2,9 if group=='reinspection' else 1).number_format='General'
                book.save(path);book.close()
                self.assertFalse(acquisition.validate_file(path,group,'2026-09-01','2026-09-30'))

    def test_automatic_job_accepts_missing_files_and_reserves_sfc(self):
        with patch.object(reports.subprocess,'Popen'):
            run=reports.start_run('reinspection','2026-09','all',[],acquire_missing=True)
        self.assertEqual(run['phase'],'acquisition')
        self.assertEqual(len(run['acquisition']['missing']),12)
        self.assertTrue(reports.browser_reserved())

    def test_downloads_correct_sources_and_keeps_completed_files_for_reuse(self):
        folder=self.runs/('f'*32);(folder/'workspace').mkdir(parents=True)
        page=SimpleNamespace(url=acquisition.REINSPECTION_URL,
            context=SimpleNamespace(cookies=lambda:[{'name':'MyFactoryName','value':'VH'}]))
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
            group='reinspection' if action=='reinspection_export' else 'ftt'
            if group=='ftt':
                self.assertEqual(params['stations'],[])
                self.assertEqual(params['station_selection'],'none')
                self.assertEqual(params['report_type'],'internal')
            self.assertEqual(params['date_from'],'2026-09-01')
            self.assertEqual(params['date_to'],'2026-09-30')
            path=self.exports/acquisition.filename(group,params['_expected_plant'],params,'.xlsx')
            raw_file(path,group)
            return dict(success=True,downloaded_file={'path':str(path)})
        def mes(**params):
            self.assertIs(params['nike'],False)
            self.assertEqual(params['factory'],'')
            self.assertEqual(params['date_to'],'2026-09-30')
            path=self.exports/f'raw-{params["expected_plant"]}.xlsx'
            raw_file(path,'hfpa')
            return dict(success=True,file={'path':str(path)})
        fake._execute_dynamic_action=MagicMock(side_effect=export)
        fake.apply_sfc_filters=MagicMock(side_effect=mes)
        run=dict(period='2026-09',inputs=[],acquisition=acquisition.plan('2026-09'))
        acquisition.acquire(folder,run,lambda message:None,fake)
        self.assertEqual([c.kwargs['params']['_expected_plant'] for c in fake._execute_dynamic_action.call_args_list
                          if c.args[0]=='reinspection_export'],list(acquisition.PLANTS))
        self.assertEqual(fake.apply_sfc_filters.call_count,4)
        self.assertEqual(len(run['inputs']),12)
        self.assertEqual(len(acquisition.plan('2026-09')['available']),12)
        victim=self.exports/'Re-Inspection 2026-09-01_2026-09-30-VH.xlsx'
        victim.unlink()
        retry=acquisition.plan('2026-09')
        self.assertEqual(retry['missing'],[dict(group='reinspection',plant='VH')])
        fake._execute_dynamic_action=MagicMock(return_value=dict(success=False,error='Timeout waiting for download'))
        with self.assertRaisesRegex(RuntimeError,'Timeout'):
            acquisition.acquire(folder,dict(period='2026-09',inputs=[],acquisition=retry),lambda message:None,fake)
        fake._execute_dynamic_action.assert_called_once()
        self.assertEqual(len(acquisition.plan('2026-09')['available']),11)


if __name__=='__main__':unittest.main()
