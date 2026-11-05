from contextlib import contextmanager
from copy import deepcopy
from datetime import datetime
import json
from pathlib import Path
import sys
import subprocess
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

import openpyxl

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import sfc_controller as controller
import sfc_downloads as downloads
import sfc_mqaa_exports as mqaa
from sfc_excel import file_info,read_export


class MQAADailyTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.root=Path(self.temp.name)
        self.patches=[patch.object(downloads,'EXPORT_ROOT',self.root),patch('sfc_excel.EXPORT_ROOT',self.root)]
        for p in self.patches:p.start()
        self.filters={'dtDateFrom':dict(kind='date',value='2026-10-19'),
                      'dtDateTo':dict(kind='date',value='2026-10-21'),
                      'cboType':dict(kind='combo',value='UPPER'),
                      'cboReport':dict(kind='combo',value='ALL'),
                      'ckbNIKE':dict(kind='check',value=False)}
        self.page=MagicMock()
        self.page.url=controller.URL_MAPPING['mqaa_report']
        self.page.context.cookies.return_value=[dict(name='MyFactoryName',value='VH2')]

    def tearDown(self):
        for p in reversed(self.patches):p.stop()
        self.temp.cleanup()

    def raw(self,day,plant='VH2',empty=False,extra_header=False):
        path=self.root/(day+'-'+plant+'.xlsx')
        b=openpyxl.Workbook()
        s=b.active
        s.append(['MQAA REPORT - All'])
        s.append([f'From: {day}  To: {day}'])
        s.append([])
        header=['Audit Date','Audit form','Factory','MQAA%','Remark']
        if extra_header:header.append('New column')
        s.append(header)
        if not empty:
            # Two identical records must both survive; row values are data.
            for _ in range(2):s.append([datetime.fromisoformat(day),'Tooling',plant,0.8,'=literal text'])
        s.append([None,None,None,'=AVERAGE(D5:D6)'])
        s.append([None,None,None,0.8])
        b.save(path);b.close()
        return dict(file_info(path),plant=plant,date_from=day,date_to=day)

    def test_only_completed_sql_response_triggers_fallback(self):
        for status in (200,500):
            response=SimpleNamespace(status=status,headers={'content-type':'text/html'},
                text=lambda:'<h1>Server Error</h1>Execution Timeout Expired. The timeout period elapsed.')
            with self.assertRaises(downloads.SFCServerTimeout): downloads.require_attachment(response)
        response=SimpleNamespace(status=500,headers={'content-type':'text/html'},text=lambda:'Permission denied')
        with self.assertRaises(RuntimeError) as raised:downloads.require_attachment(response)
        self.assertNotIsInstance(raised.exception,downloads.SFCServerTimeout)

    def test_native_mqaa_control_names_are_captured_with_dates_and_dropdowns(self):
        page=MagicMock()
        downloads.capture_filters(page)
        script=page.evaluate.call_args.args[0]
        code=r'''
const vm=require('node:vm'),fs=require('node:fs'),assert=require('node:assert/strict');
const script=JSON.parse(fs.readFileSync(0,'utf8'));
const controls=[{name:'dtDateFrom',GetDate:()=>new Date(2026,9,19)},
{name:'dtDateTo',GetDate:()=>new Date(2026,9,24)},
...['cboType','cboGroup','cboReportType'].map(name=>({name,GetValue:()=>null,GetItemCount:()=>3})),
{name:'wucHeaderMainTop',GetValue:()=> 'ignore',GetItemCount:()=>7}];
const result=vm.runInNewContext('('+script+')()',{
location:{pathname:'/vhmes/QA/MQAAReport.aspx'},
aspxGetControlCollection:()=>({ForEachControl:f=>controls.forEach(f)})});
assert.equal(Object.keys(result).length,5);
assert.equal(result.dtDateFrom.value,'2026-10-19');
assert.equal(result.dtDateTo.value,'2026-10-24');
assert.equal(result.cboType.value,null);
'''
        result=subprocess.run(['node','-e',code],input=json.dumps(script),text=True,capture_output=True,timeout=5)
        self.assertEqual(result.returncode,0,result.stderr)

    def test_timeout_downloads_every_day_once_and_merges_verified_rows(self):
        calls=[]
        def export(page,target,timeout,plant,format_name,filters):
            calls.append(deepcopy(filters))
            if len(calls)==1:raise downloads.SFCServerTimeout('Execution Timeout Expired')
            return self.raw(filters['dtDateFrom']['value'],empty=len(calls)==3)
        with patch.object(downloads,'capture_filters',return_value=self.filters), \
                patch.object(downloads,'apply_filters') as restore, \
                patch.object(mqaa,'export_once',side_effect=export):
            result=mqaa.export_mqaa_data(self.page,'btnExcel',1000,'VH2',controller.format_download_filename)
        self.assertEqual(len(calls),4)
        self.assertEqual([f['dtDateFrom']['value'] for f in calls[1:]],['2026-10-19','2026-10-20','2026-10-21'])
        for f in calls[1:]:
            self.assertEqual(f['dtDateFrom'],f['dtDateTo'])
            for key in ('cboType','cboReport','ckbNIKE'):self.assertEqual(f[key],self.filters[key])
        self.assertEqual(result['rows'],4)
        self.assertEqual(len(result['parts']),3)
        self.assertTrue(result['daily_fallback'])
        self.page.goto.assert_called_once()
        restore.assert_called_once_with(self.page,self.filters,1000)
        book=openpyxl.load_workbook(result['path'],read_only=True,data_only=False)
        self.assertEqual(book.sheetnames,['MQAA','NguonTheoNgay'])
        rows=list(book['MQAA'].iter_rows(values_only=True))
        self.assertEqual(len(rows),5)
        self.assertEqual(rows[1],rows[2])
        self.assertEqual(book['MQAA']['E2'].data_type,'s')
        self.assertEqual(book['MQAA']['D2'].number_format,'0.0%')
        book.close()
        manifest=json.loads(next(self.root.glob('mqaa_daily_*/manifest.json')).read_text(encoding='utf8'))
        self.assertEqual(manifest['status'],'done')
        self.assertEqual(len(manifest['monthly_files']),3)
        self.assertEqual(len(manifest['combined_files']),1)

    def test_daily_failure_keeps_parts_and_never_publishes_incomplete_range(self):
        with patch.object(downloads,'capture_filters',return_value=self.filters), \
                patch.object(downloads,'apply_filters'), \
                patch.object(mqaa,'export_once',side_effect=[downloads.SFCServerTimeout('SQL timeout'),
                    self.raw('2026-10-19'),downloads.SFCServerTimeout('SQL timeout')]) as export:
            with self.assertRaisesRegex(RuntimeError,'2026-10-20') as raised:
                mqaa.export_mqaa_data(self.page,'btnExcel',1000,'VH2',controller.format_download_filename)
        self.assertEqual(export.call_count,3)
        self.assertEqual(len(raised.exception.downloaded_files),1)
        self.assertTrue(Path(raised.exception.downloaded_files[0]['path']).is_file())
        self.assertFalse(list(self.root.glob('mqaa_daily_*/MQAA_*.xlsx')))
        manifest=json.loads(next(self.root.glob('mqaa_daily_*/manifest.json')).read_text(encoding='utf8'))
        self.assertEqual(manifest['status'],'error')
        self.assertEqual(manifest['failed_date'],'2026-10-20')

    def test_client_timeout_and_permission_error_never_start_daily_requests(self):
        for error in (RuntimeError('download timed out'),RuntimeError('No permission')):
            with patch.object(downloads,'capture_filters',return_value=self.filters), \
                    patch.object(mqaa,'export_once',side_effect=error) as export:
                with self.assertRaises(RuntimeError):
                    mqaa.export_mqaa_data(self.page,'btnExcel',1000,'VH2',controller.format_download_filename)
            export.assert_called_once()
        self.page.goto.assert_not_called()

    def test_one_day_timeout_is_not_retried(self):
        self.filters['dtDateTo']['value']='2026-10-19'
        with patch.object(downloads,'capture_filters',return_value=self.filters), \
                patch.object(mqaa,'export_once',side_effect=downloads.SFCServerTimeout('SQL timeout')) as export:
            with self.assertRaises(downloads.SFCServerTimeout):
                mqaa.export_mqaa_data(self.page,'btnExcel',1000,'VH2',controller.format_download_filename)
        export.assert_called_once()
        self.page.goto.assert_not_called()

    def test_merge_rejects_missing_duplicate_wrong_factory_date_and_changed_schema(self):
        a=self.raw('2026-10-19');b=self.raw('2026-10-20')
        for parts in ([a],[a,a],[a,dict(b,plant='VH')]):
            with self.subTest(parts=parts):
                with self.assertRaises(ValueError):
                    mqaa.merge_days(parts,self.root/'combined.xlsx','VH2','2026-10-19','2026-10-20',self.filters)
        wrong=self.raw('2026-10-21')
        with self.assertRaisesRegex(ValueError,'ngoài ngày'):
            mqaa.daily_table(wrong['path'],'2026-10-20','VH2')
        wrong=self.raw('2026-10-20',plant='VH')
        with self.assertRaisesRegex(ValueError,'nhà máy'):
            mqaa.daily_table(wrong['path'],'2026-10-20','VH2')
        b=self.raw('2026-10-20',extra_header=True)
        with self.assertRaisesRegex(ValueError,'Cấu trúc cột'):
            mqaa.merge_days([a,b],self.root/'combined.xlsx','VH2','2026-10-19','2026-10-20',self.filters)
        self.assertFalse((self.root/'combined.xlsx').exists())

    def test_unrelated_export_is_rejected_and_summary_reports_keep_daily_sheets(self):
        parts=[]
        for day in ('2026-10-19','2026-10-20'):
            path=self.root/(day+'.xlsx')
            b=openpyxl.Workbook()
            b.active.append(['MQAA REPORT - Factory'])
            b.active.append(['Factory','Score'])
            b.active.append(['VH2',0.8])
            b.save(path);b.close()
            parts.append(dict(file_info(path),plant='VH2',date_from=day,date_to=day))
        result=mqaa.merge_days(parts,self.root/'summary.xlsx','VH2','2026-10-19','2026-10-20',self.filters)
        b=openpyxl.load_workbook(result['path'],read_only=True)
        self.assertEqual(b.sheetnames,['2026-10-19','2026-10-20','NguonTheoNgay'])
        self.assertEqual(b['2026-10-19']['B3'].value,0.8)
        b.close()
        for title in ('COPQ Report','MQAA REPORT - All'):
            b=openpyxl.Workbook();b.active.append([title]);b.active.append(['Wrong','Columns'])
            b.save(self.root/'wrong.xlsx');b.close()
            with self.assertRaises(ValueError):mqaa.daily_table(self.root/'wrong.xlsx','2026-10-19','VH2')

    def test_scope_change_prevents_the_post(self):
        with patch.object(downloads,'apply_filters'), \
                patch.object(downloads,'capture_filters',return_value=self.filters):
            with self.assertRaisesRegex(ValueError,'Nhà máy'):
                mqaa.export_once(self.page,'btnExcel',1000,'VH',controller.format_download_filename,self.filters)
        self.page.expect_response.assert_not_called()

    def test_sql_error_response_is_handled_before_waiting_for_download(self):
        response=SimpleNamespace(status=500,headers={'content-type':'text/html'},
                                 text=lambda:'Execution Timeout Expired')
        self.page.expect_response.return_value.__enter__.return_value=SimpleNamespace(value=response)
        with patch.object(downloads,'apply_filters'), \
                patch.object(downloads,'capture_filters',return_value=self.filters):
            with self.assertRaises(downloads.SFCServerTimeout):
                mqaa.export_once(self.page,'btnExcel',1000,'VH2',controller.format_download_filename,self.filters)
        self.page.wait_for_event.assert_not_called()
        self.page.remove_listener.assert_called_once()

    def test_existing_timeout_proceeds_directly_to_days_and_success_never_splits(self):
        with patch.object(downloads,'capture_filters',return_value=self.filters), \
                patch.object(downloads,'apply_filters'), \
                patch.object(mqaa,'export_once',side_effect=lambda *args:self.raw(args[-1]['dtDateFrom']['value'])) as export:
            result=mqaa.export_mqaa_data(self.page,'btnExcel',1000,'VH2',controller.format_download_filename,confirmed_timeout=True)
        self.assertEqual(export.call_count,3)
        self.assertEqual(len(result['parts']),3)
        with patch.object(downloads,'capture_filters',return_value=self.filters), \
                patch.object(mqaa,'export_once',return_value={'name':'original.xlsx'}) as export:
            self.assertEqual(mqaa.export_mqaa_data(self.page,'btnExcel',1000,'VH2',controller.format_download_filename),{'name':'original.xlsx'})
        export.assert_called_once()

    def test_mqaa_controller_recovers_existing_sql_error_without_replaying_range(self):
        @contextmanager
        def playwright():
            yield SimpleNamespace(chromium=SimpleNamespace(connect_over_cdp=lambda *a,**k:browser))
        browser=SimpleNamespace(close=MagicMock())
        self.page.title.return_value='Execution Timeout Expired. The timeout period elapsed'
        self.page.evaluate.return_value={'page':{'url':self.page.url}}
        with patch.object(controller,'sync_playwright',playwright), \
                patch.object(controller,'get_sfc_page',return_value=self.page), \
                patch.object(downloads,'wait_idle'),patch.object(downloads,'apply_filters'), \
                patch.object(downloads,'capture_filters',return_value=self.filters), \
                patch.object(controller,'take_b64_screenshot',return_value=''),patch.object(controller.time,'sleep'), \
                patch.object(mqaa,'export_mqaa_data',return_value=dict(name='MQAA-VH2.xlsx',daily_fallback=True)) as export:
            result=controller._execute_dynamic_action('click_button','btnExcel',
                {'date_from':'2026-10-19','date_to':'2026-10-21'})
        self.assertTrue(result['success'],result)
        self.assertTrue(export.call_args.kwargs['confirmed_timeout'])
        self.assertEqual(len(result['downloaded_files']),1)
        self.page.goto.assert_called_once()


if __name__=='__main__':unittest.main()
