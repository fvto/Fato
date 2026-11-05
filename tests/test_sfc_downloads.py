from contextlib import contextmanager
import io
import json
from pathlib import Path
import sys
import tempfile
import subprocess
import threading
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch
import zipfile

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import sfc_controller as controller
import sfc_downloads as downloads
import sfc_dashboard_server as server
import sfc_commands as commands


def workbook_bytes():
    stream = io.BytesIO()
    workbook = openpyxl.Workbook()
    workbook.active.append(['COPQ', 'September 2026'])
    workbook.save(stream)
    workbook.close()
    return stream.getvalue()


def download(name='COPQ.xlsx', content=None, failure=None):
    payload = workbook_bytes() if content is None else content
    return SimpleNamespace(suggested_filename=name, failure=lambda: failure,
        save_as=lambda target: Path(target).write_bytes(payload))


@contextmanager
def client(browser):
    yield SimpleNamespace(chromium=SimpleNamespace(connect_over_cdp=lambda *a, **k: browser))


class DownloadTests(unittest.TestCase):
    def test_second_export_is_rejected_instead_of_queued(self):
        entered, release = threading.Event(), threading.Event()
        result = []
        def first_action(*args):
            entered.set()
            release.wait(3)
            return dict(success=True)
        with patch.object(controller, '_execute_dynamic_action_locked', side_effect=first_action) as action:
            worker = threading.Thread(target=lambda: result.append(controller.execute_dynamic_action('click_button', 'btnExcel')))
            worker.start()
            try:
                self.assertTrue(entered.wait(2))
                duplicate = controller.execute_dynamic_action('click_button', 'btnExcel')
                self.assertFalse(duplicate['success'])
                self.assertTrue(duplicate['busy'])
                self.assertEqual(action.call_count, 1)
                with server.app.test_client() as client:
                    response = client.post('/api/page/execute', json={'action_type':'click_button','target_id':'btnExcel'})
                    self.assertEqual(response.status_code, 409)
            finally:
                release.set()
                worker.join(3)
        self.assertTrue(result[0]['success'])
        self.assertFalse(controller.DYNAMIC_ACTION_LOCK.locked())

    def test_scanner_preserves_native_copq_options_and_invalid_type_state(self):
        script = r'''
const vm=require('node:vm'), fs=require('node:fs'), assert=require('node:assert/strict');
const inspect=JSON.parse(fs.readFileSync(0,'utf8'));
const category={name:'rpnlMain_ContentPlaceHolder1_cboCategory',GetText:()=> 'Internal',GetValue:()=> 'INT',
  GetItemCount:()=>2,GetItem:i=>[{value:'EXT',text:'External'},{value:'INT',text:'Internal'}][i]};
const type={name:'rpnlMain_ContentPlaceHolder1_cboType',GetText:()=> 'DR',GetValue:()=> 'DR',
  GetItemCount:()=>3,GetItem:i=>[{value:'BC',text:'B/C'},{value:'REI',text:'Reinspection'},{value:'REP',text:'Rework'}][i]};
const element={getBoundingClientRect:()=>({width:10,height:10}),closest:()=>null};
const context={location:{href:'https://sfc.chingluh.com/vhmes/QA/QAReportCOPQ.aspx',pathname:'/vhmes/QA/QAReportCOPQ.aspx'},
  document:{title:'Report COPQ',getElementById:()=>element},window:{getComputedStyle:()=>({display:'block',visibility:'visible'})},
  aspxGetControlCollection:()=>({elements:{category,type}})};
const result=vm.runInNewContext('('+inspect+')()',context);
assert.equal(result.dropdowns[0].options.length,4);
assert.equal(result.dropdowns[0].options[1].value,'ALL');
assert.equal(result.dropdowns[0].options[0].is_clear,true);
assert.equal(result.dropdowns[0].selected_value,'INT');
assert.equal(result.dropdowns[0].supports_all,false);
assert.equal(result.dropdowns[1].options.length,5);
assert.equal(result.dropdowns[1].requires_selection,true);
assert.equal(result.dropdowns[1].selected_value,null);
'''
        result = subprocess.run(['node', '-e', script], input=json.dumps(controller.INSPECT_JS),
                                text=True, capture_output=True, timeout=10)
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_permission_response_is_reported_instead_of_waiting_for_download(self):
        response = SimpleNamespace(status=200, headers={'content-type':'text/html'},
            text=lambda:"<script>OnShowMyMessage('No permission to manipulate data!','Warning','True');</script>")
        with self.assertRaisesRegex(RuntimeError, 'No permission'):
            downloads.require_attachment(response)

    def test_empty_copq_option_clears_actual_control_and_posts_null(self):
        page = MagicMock()
        page.evaluate.side_effect = ['DR', None, None]
        page.expect_response.return_value.__enter__.return_value = SimpleNamespace(value=SimpleNamespace(status=200))
        downloads.set_copq_combobox(page, 'rpnlMain_ContentPlaceHolder1_cboType', '', 1000)
        mutation = page.evaluate.call_args_list[1]
        self.assertIsNone(mutation.args[1]['value'])
        self.assertIn("c.SetText('')", mutation.args[0])
        self.assertIn('c.SetSelectedIndex(-1)', mutation.args[0])
        self.assertIn('__doPostBack', mutation.args[0])
        page.expect_response.assert_called_once()

    def test_all_copq_option_posts_supported_null_value(self):
        page = MagicMock()
        page.evaluate.side_effect = ['DR', None, None]
        page.expect_response.return_value.__enter__.return_value = SimpleNamespace(value=SimpleNamespace(status=200))
        downloads.set_copq_combobox(page, 'rpnlMain_ContentPlaceHolder1_cboType', 'ALL', 1000)
        self.assertIsNone(page.evaluate.call_args_list[1].args[1]['value'])

    def test_copq_seven_factories_last_month_command(self):
        from datetime import datetime, timedelta, timezone
        today = datetime.now(timezone(timedelta(hours=7))).date()
        last = today.replace(day=1) - timedelta(days=1)
        intent = commands.parse_intent('Xuất Excel COPQ 7 nhà máy last month')
        self.assertEqual(intent['target_page'], 'copq')
        self.assertEqual(intent['target_plants'], list(commands.PLANTS))
        self.assertEqual(intent['date_from'], last.replace(day=1).isoformat())
        self.assertEqual(intent['date_to'], last.isoformat())
        fake = SimpleNamespace(SELECTED_FACTORIES=list(commands.PLANTS),
            set_selected_factories=MagicMock(return_value={'success':True}),
            navigate_to=MagicMock(return_value={'success':True}),
            execute_dynamic_action=MagicMock(return_value={'success':True,'message':'7 Excel files',
                'downloaded_files':[{'name':p+'.xls'} for p in commands.PLANTS]}))
        events = list(commands.execute(fake, 'Xuất Excel COPQ 7 nhà máy last month'))
        self.assertEqual(events[-1]['status'], 'done', events[-1])
        self.assertEqual(len(events[-1]['files']), 7)
        self.assertEqual(fake.execute_dynamic_action.call_args.args[2]['date_from'], intent['date_from'])
        fake.navigate_to.assert_called_once_with('copq_data')
        self.assertEqual(fake.execute_dynamic_action.call_args.args[:2], ('click_button', 'rpnlMain_ContentPlaceHolder1_btnExcel'))

    def test_uuid_download_is_saved_as_real_excel_without_changing_bytes(self):
        payload = workbook_bytes()
        with tempfile.TemporaryDirectory() as directory, patch.object(downloads, 'EXPORT_ROOT', Path(directory)), \
                patch('sfc_excel.EXPORT_ROOT', Path(directory)):
            result = downloads.save_download(download('83ec9ad5-e4dd-46e1-a436-b6117e3cc1d9', payload),
                                             'VH2', controller.format_download_filename)
            self.assertTrue(result['name'].endswith('-VH2.xlsx'))
            self.assertEqual(Path(result['path']).read_bytes(), payload)
            self.assertFalse(list(Path(directory).rglob('*.partial')))
            self.assertIn('/api/exports/download_', result['url'])

    def test_macro_extension_and_windows_filename_are_preserved(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(downloads, 'EXPORT_ROOT', Path(directory)), \
                patch('sfc_excel.EXPORT_ROOT', Path(directory)):
            result = downloads.save_download(download('FTT-NIKE 14:10:09.xlsm'), 'JV', controller.format_download_filename)
            self.assertEqual(result['name'], 'FTT-NIKE 14_10_09-JV.xlsm')
            with patch.object(server, 'EXPORT_ROOT', Path(directory)), server.app.test_client() as client:
                response = client.get(result['url'])
                self.assertEqual(response.status_code, 200)
                self.assertIn('.xlsm', response.headers['Content-Disposition'])
                self.assertEqual(response.data, Path(result['path']).read_bytes())
                response.close()

    def test_html_login_or_failed_download_never_becomes_excel(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(downloads, 'EXPORT_ROOT', Path(directory)):
            for item in (download(content=b'<html>Login</html>'), download(failure='canceled')):
                with self.assertRaises((ValueError, RuntimeError)):
                    downloads.save_download(item, 'VH', controller.format_download_filename)
            self.assertFalse(list(Path(directory).rglob('*.xlsx')))
            self.assertFalse(list(Path(directory).rglob('*.partial')))

    def test_non_workbook_zip_is_rejected(self):
        content = io.BytesIO()
        with zipfile.ZipFile(content, 'w') as archive:
            archive.writestr('word/document.xml', '<document/>')
        with tempfile.TemporaryDirectory() as directory, patch.object(downloads, 'EXPORT_ROOT', Path(directory)):
            with self.assertRaises(ValueError):
                downloads.save_download(download(content=content.getvalue()), 'VH', controller.format_download_filename)

    def page(self):
        page = MagicMock()
        page.url = controller.URL_MAPPING['copq']
        page.title.return_value = 'COPQ'
        page.context.cookies.return_value = [{'name':'MyFactoryName', 'value':'VH'}]
        page.wait_for_event.return_value = download()
        page.expect_response.return_value.__enter__.return_value = SimpleNamespace(value=SimpleNamespace(
            status=200, headers={'content-disposition':'attachment; filename=COPQ.xls'},finished=lambda:None))
        return page

    def test_menu_capture_is_armed_before_click_and_uses_actual_factory(self):
        page = self.page()
        events = []
        page.on.side_effect = lambda event, handler: events.append('armed') if event == 'download' else None
        page.query_selector.return_value.click.side_effect = lambda **kwargs: events.append('clicked')
        browser = SimpleNamespace(contexts=[SimpleNamespace(pages=[page])], close=MagicMock())
        with patch.object(controller, 'sync_playwright', lambda: client(browser)), \
                patch.object(controller, 'take_b64_screenshot', return_value=''), patch.object(controller.time, 'sleep'), \
                patch.object(downloads, 'save_download', return_value={'name':'COPQ-VH.xlsx'}) as save:
            result = controller._execute_dynamic_action('select_menu', 'btnExportExcel', {'index':0})
        self.assertTrue(result['success'], result)
        self.assertEqual(events, ['armed', 'clicked'])
        self.assertEqual(save.call_args.args[1], 'VH')
        self.assertEqual(result['downloaded_files'], [{'name':'COPQ-VH.xlsx'}])

    def test_jv3_cookie_matches_dashboard_jvb_factory(self):
        page = self.page()
        page.context.cookies.return_value = [{'name':'MyFactoryName', 'value':'JV3'}]
        browser = SimpleNamespace(contexts=[SimpleNamespace(pages=[page])], close=MagicMock())
        with patch.object(controller, 'sync_playwright', lambda: client(browser)):
            result = controller.switch_factory('JVB')
        self.assertTrue(result['success'], result)
        page.goto.assert_not_called()
        page.click.assert_not_called()

    def test_copq_data_opens_and_scans_without_exporting_selected_factories(self):
        page = self.page()
        date_filters = {'dtDateFrom':{'kind':'date', 'value':'2026-09-01'}}
        browser = SimpleNamespace(contexts=[SimpleNamespace(pages=[page])], close=MagicMock())
        with patch.object(controller, 'sync_playwright', lambda: client(browser)), \
                patch.object(controller, 'SELECTED_FACTORIES', controller.SEVEN_FACTORIES.copy()), \
                patch.object(controller, 'take_b64_screenshot', return_value=''), patch.object(controller.time, 'sleep'), \
                patch.object(downloads, 'capture_filters', return_value=date_filters), \
                patch.object(downloads, 'apply_filters') as apply, \
                patch.object(downloads, 'save_download') as save, \
                patch.object(downloads, 'export_selected') as batch:
            result = controller.execute_dynamic_action('select_menu', 'btnExportExcel', {'index':1})
        self.assertTrue(result['success'], result)
        self.assertTrue(result['requires_filter_selection'])
        self.assertEqual(result['downloaded_files'], [])
        page.goto.assert_called_once_with(controller.URL_MAPPING['copq_data'],
            wait_until='domcontentloaded', timeout=controller.SFC_WAIT_MS)
        apply.assert_called_once_with(page, date_filters, controller.SFC_WAIT_MS)
        page.locator.assert_not_called()
        save.assert_not_called()
        batch.assert_not_called()

    def test_excel_on_copq_data_searches_detail_page_then_captures_excel(self):
        page = self.page()
        page.url = controller.URL_MAPPING['copq_data']
        browser = SimpleNamespace(contexts=[SimpleNamespace(pages=[page])], close=MagicMock())
        with patch.object(controller, 'sync_playwright', lambda: client(browser)), \
                patch.object(controller, 'take_b64_screenshot', return_value=''), patch.object(controller.time, 'sleep'), \
                patch.object(downloads, 'save_download', return_value={'name':'COPQ-Detail-VH.xls'}):
            result = controller._execute_dynamic_action('click_button', 'rpnlMain_ContentPlaceHolder1_btnExcel')
        self.assertTrue(result['success'], result)
        self.assertEqual([call.args[0] for call in page.locator.call_args_list],
                         ['#rpnlMain_ContentPlaceHolder1_btnSearch', '#rpnlMain_ContentPlaceHolder1_btnExcel'])
        self.assertEqual(result['downloaded_file']['name'], 'COPQ-Detail-VH.xls')

    def test_reinspection_clears_line_and_type_then_searches_before_excel(self):
        page=self.page()
        page.url='https://sfc.chingluh.com/vhmes/QA/QCRECHK.aspx'
        page.evaluate.return_value=True
        prefix='rpnlMain_ContentPlaceHolder1_'
        filters={prefix+'dtDateFrom':dict(kind='date',value='2026-09-01'),
                 prefix+'dtDateTo':dict(kind='date',value='2026-09-30'),
                 prefix+'cboDeptFrom':dict(kind='combo',value=None),
                 prefix+'cboType':dict(kind='combo',value=None)}
        with patch.object(downloads,'apply_filters') as apply, \
                patch.object(downloads,'capture_filters',return_value=filters), \
                patch.object(downloads,'save_download',return_value={'name':'Re-Inspection-VH.xls'}) as save:
            result=downloads.export_reinspection_data(page,1000,'VH',controller.format_download_filename,
                dict(date_from='2026-09-01',date_to='2026-09-30'))
        self.assertEqual(result['name'],'Re-Inspection-VH.xls')
        apply.assert_called_once_with(page,filters,1000)
        self.assertEqual([c.args[0] for c in page.locator.call_args_list],['#'+prefix+'btnSearch','#'+prefix+'btnExcel'])
        self.assertEqual(save.call_args.args[1],'VH')
        clear=next(c.args[0] for c in page.evaluate.call_args_list if 'c.SetSelectedIndex(-1)' in c.args[0])
        code=r'''
const vm=require('node:vm'),assert=require('node:assert/strict'),fs=require('node:fs');
const script=JSON.parse(fs.readFileSync(0,'utf8'));
const controls=Object.fromEntries(['cboDeptFrom','cboType'].map(key=>['rpnlMain_ContentPlaceHolder1_'+key,{
 index:0,value:'OLD',text:'OLD',SetSelectedIndex(x){this.index=x},SetValue(x){this.value=x},SetText(x){this.text=x}}]));
vm.runInNewContext('('+script+')()',{aspxGetControlCollection:()=>({Get:id=>controls[id]})});
for(const c of Object.values(controls)){assert.equal(c.value,null);assert.equal(c.text,'');assert.equal(c.index,-1)}
'''
        outcome=subprocess.run(['node','-e',code],input=json.dumps(clear),text=True,capture_output=True,timeout=5)
        self.assertEqual(outcome.returncode,0,outcome.stderr)

    def test_reinspection_does_not_export_if_postback_restores_a_filter(self):
        page=self.page();page.url='https://sfc.chingluh.com/vhmes/QA/QCRECHK.aspx'
        page.evaluate.return_value=True
        prefix='rpnlMain_ContentPlaceHolder1_'
        expected={prefix+'dtDateFrom':dict(kind='date',value='2026-09-01'),
                  prefix+'dtDateTo':dict(kind='date',value='2026-09-30'),
                  prefix+'cboDeptFrom':dict(kind='combo',value=None),prefix+'cboType':dict(kind='combo',value=None)}
        wrong=dict(expected);wrong[prefix+'cboType']=dict(kind='combo',value='QC')
        with patch.object(downloads,'apply_filters'),patch.object(downloads,'capture_filters',side_effect=[expected,wrong]), \
                patch.object(downloads,'save_download') as save:
            with self.assertRaisesRegex(ValueError,'Line/Type'):
                downloads.export_reinspection_data(page,1000,'VH',controller.format_download_filename,
                    dict(date_from='2026-09-01',date_to='2026-09-30'))
        save.assert_not_called()
        self.assertEqual([c.args[0] for c in page.locator.call_args_list],['#'+prefix+'btnSearch'])

    def test_internal_ftt_none_does_not_turn_into_select_all(self):
        page=self.page()
        page.url=controller.URL_MAPPING['ftt']
        page.expect_download.return_value.__enter__.return_value=SimpleNamespace(value=download())
        browser=SimpleNamespace(contexts=[SimpleNamespace(pages=[page])],close=MagicMock())
        with patch.object(controller,'sync_playwright',lambda:client(browser)), \
                patch.object(controller,'take_b64_screenshot',return_value=''),patch.object(controller.time,'sleep'), \
                patch.object(downloads,'save_download',return_value={'name':'FTT-Internal-VH.xlsx'}):
            result=controller._execute_dynamic_action('ftt_station_export',params={
                'stations':[],'station_selection':'none','report_type':'internal'})
        self.assertTrue(result['success'],result)
        call=next(call for call in page.evaluate.call_args_list if '({stations, selection})' in call.args[0])
        code=r'''
const vm=require('node:vm'),assert=require('node:assert/strict'),fs=require('node:fs');
const script=JSON.parse(fs.readFileSync(0,'utf8'));let selected=[0,1],selectAll=0;
const control={UnselectAll(){selected=[]},SelectAll(){selectAll++},GetSelectedIndices(){return selected}};
vm.runInNewContext('('+script+')({stations:[],selection:"none"})',{
 aspxGetControlCollection:()=>({Get:()=>control})});
assert.equal(selected.length,0);assert.equal(selectAll,0);
'''
        outcome=subprocess.run(['node','-e',code],input=json.dumps(call.args[0]),text=True,capture_output=True,timeout=5)
        self.assertEqual(outcome.returncode,0,outcome.stderr)

    def test_hfpa_station_is_selected_exclusively_and_export_uses_ok(self):
        page=self.page()
        page.url=controller.URL_MAPPING['ftt']
        page.expect_download.return_value.__enter__.return_value=SimpleNamespace(value=download())
        browser=SimpleNamespace(contexts=[SimpleNamespace(pages=[page])],close=MagicMock())
        with patch.object(controller,'sync_playwright',lambda:client(browser)), \
                patch.object(controller,'take_b64_screenshot',return_value=''),patch.object(controller.time,'sleep'), \
                patch.object(downloads,'save_download',return_value={'name':'QAStation-VH.xlsx'}):
            result=controller._execute_dynamic_action('ftt_station_export',params={
                'stations':['HFPA'],'station_selection':'explicit','report_type':'station'})
        self.assertTrue(result['success'],result)
        self.assertTrue(any(call.args[-1]=='rpnlMain_ContentPlaceHolder1_popupSetArea_btnSave'
                            for call in page.evaluate.call_args_list))
        call=next(call for call in page.evaluate.call_args_list if '({stations, selection})' in call.args[0])
        code=r'''
const vm=require('node:vm'),assert=require('node:assert/strict'),fs=require('node:fs');
const script=JSON.parse(fs.readFileSync(0,'utf8'));let selected=[0,1],selectAll=0;
const items=[{value:'T01',text:'FTT'},{value:'T22',text:'HFPA'}];
const control={UnselectAll(){selected=[]},SelectAll(){selectAll++},GetSelectedIndices(){return selected},
 GetItemCount(){return items.length},GetItem(i){return items[i]},SelectIndices(indices){selected.push(...indices)}};
const env={aspxGetControlCollection:()=>({Get:()=>control})};
vm.runInNewContext('('+script+')({stations:["HFPA"],selection:"explicit"})',env);
assert.deepEqual(selected,[1]);assert.equal(selectAll,0);
assert.throws(()=>vm.runInNewContext('('+script+')({stations:["MISSING"],selection:"explicit"})',env));
'''
        outcome=subprocess.run(['node','-e',code],input=json.dumps(call.args[0]),text=True,capture_output=True,timeout=5)
        self.assertEqual(outcome.returncode,0,outcome.stderr)

    def test_menu_and_button_failures_are_not_reported_as_success(self):
        for action, target in [('select_menu', 'btnExportExcel'), ('click_button', 'btnExcel')]:
            page = self.page()
            page.expect_download.side_effect = RuntimeError('download timed out')
            browser = SimpleNamespace(contexts=[SimpleNamespace(pages=[page])], close=MagicMock())
            with patch.object(controller, 'sync_playwright', lambda: client(browser)), \
                    patch.object(downloads, 'require_attachment', side_effect=RuntimeError('download timed out')):
                result = controller._execute_dynamic_action(action, target)
            self.assertFalse(result['success'])
            self.assertIn('download timed out', result['error'])

    def test_seven_factories_replay_same_filters_and_preserve_selection(self):
        page = self.page()
        browser = SimpleNamespace(close=MagicMock())
        scope = controller.SEVEN_FACTORIES.copy()
        fake = SimpleNamespace(SELECTED_FACTORIES=scope.copy(), SEVEN_FACTORIES=scope,
            CDP_URL=controller.CDP_URL, SFC_CONNECT_WAIT_MS=300000, SFC_WAIT_MS=900000,
            sync_playwright=lambda: client(browser), get_sfc_page=lambda b:page)
        seen = []
        def switch(plant):
            fake.SELECTED_FACTORIES = [plant]
            return dict(success=True)
        def execute(action, target_id=None, params=None):
            if action == 'restore_filters':
                return dict(success=True)
            seen.append(params)
            return dict(success=True, downloaded_file={'name':params['_expected_plant']+'.xlsx'})
        fake.switch_factory = switch
        fake._execute_dynamic_action = execute
        filters = {'dtDateFrom':{'kind':'date', 'value':'2026-09-01'}}
        with tempfile.TemporaryDirectory() as directory, patch.object(downloads, 'EXPORT_ROOT', Path(directory)), \
                patch.object(downloads, 'wait_idle'), patch.object(downloads, 'capture_filters', return_value=filters):
            result = downloads.export_selected(fake, 'select_menu', 'btnExportExcel', {'index':1}, scope)
            manifest = json.loads(next(Path(directory).glob('*/manifest.json')).read_text(encoding='utf8'))
        self.assertTrue(result['success'], result)
        self.assertEqual([item['_expected_plant'] for item in seen], scope)
        self.assertTrue(all(item['_filters'] == filters and item['index'] == 1 for item in seen))
        self.assertEqual(len(result['downloaded_files']), 7)
        self.assertEqual(fake.SELECTED_FACTORIES, scope)
        self.assertEqual(manifest['status'], 'done')

        def partial(action, target_id=None, params=None):
            if action != 'restore_filters' and params['_expected_plant'] == 'JVB':
                return dict(success=False, error='No permission to manipulate data!')
            return execute(action, target_id, params)
        fake._execute_dynamic_action = partial
        with tempfile.TemporaryDirectory() as directory, patch.object(downloads, 'EXPORT_ROOT', Path(directory)), \
                patch.object(downloads, 'wait_idle'), patch.object(downloads, 'capture_filters', return_value=filters):
            result = downloads.export_selected(fake, 'select_menu', 'btnExportExcel', {'index':0}, scope)
        self.assertFalse(result['success'])
        self.assertEqual(len(result['downloaded_files']), 6)
        self.assertEqual(result['failed_plants'][0]['plant'], 'JVB')
        self.assertEqual(fake.SELECTED_FACTORIES, scope)


if __name__ == '__main__':
    unittest.main()
