from pathlib import Path
import sys
import json
import subprocess
import unittest
from contextlib import contextmanager
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import sfc_controller as controller
import sfc_hfpa as hfpa


@contextmanager
def client(browser):
    yield SimpleNamespace(chromium=SimpleNamespace(connect_over_cdp=lambda *a, **k: browser))


class WaitingTests(unittest.TestCase):
    def setUp(self):
        self.page = MagicMock()
        self.page.url = controller.URL_MAPPING['hfpa']
        self.page.title.return_value = 'HFPA'
        self.page.screenshot.return_value = b''
        self.page.context.cookies.return_value = [{'name':'MyFactoryName', 'value':'VH'}]
        self.browser = SimpleNamespace(contexts=[SimpleNamespace(pages=[self.page])], close=MagicMock())

    def test_same_factory_preserves_current_request_and_page(self):
        with patch.object(controller, 'sync_playwright', lambda: client(self.browser)):
            result = controller.switch_factory('VH')
        self.assertTrue(result['success'])
        self.page.goto.assert_not_called()
        self.page.click.assert_not_called()

    def test_same_destination_never_reloads(self):
        with patch.object(controller, 'sync_playwright', lambda: client(self.browser)), patch.object(controller.time, 'sleep'):
            result = controller.navigate_to('hfpa')
        self.assertTrue(result['success'])
        self.page.goto.assert_not_called()

    def test_postback_waits_for_callback_response_without_navigation(self):
        response = SimpleNamespace(request=SimpleNamespace(method='POST'),
                                   url=controller.URL_MAPPING['hfpa'], status=200, finished=MagicMock())
        @contextmanager
        def expect_response(predicate, timeout):
            self.assertTrue(predicate(response))
            self.assertGreaterEqual(timeout, 15 * 60 * 1000)
            yield SimpleNamespace(value=response)
        page = SimpleNamespace(expect_response=expect_response, locator=MagicMock(), wait_for_load_state=MagicMock())
        with patch.object(hfpa, 'wait_idle') as idle:
            hfpa.postback(page, 'btnSearch')
        response.finished.assert_called_once()
        idle.assert_called_once_with(page)
        page.wait_for_load_state.assert_called_once_with('domcontentloaded', timeout=hfpa.WAIT_MS)

    def test_waits_until_all_callbacks_end_and_removes_handlers(self):
        page = MagicMock()
        hfpa.wait_idle(page)
        script = page.evaluate.call_args.args[0]
        test = r'''
const vm = require('node:vm');
const fs = require('node:fs');
const assert = require('node:assert/strict');
const script = JSON.parse(fs.readFileSync(0,'utf8'));
function event() {
  const handlers = new Set();
  return {AddHandler:f=>handlers.add(f),RemoveHandler:f=>handlers.delete(f),fire:()=>[...handlers].forEach(f=>f()),size:()=>handlers.size};
}
let gridBusy=true, comboBusy=true, finished=false;
const gridEnd=event(), comboEnd=event();
const controls=[{InCallback:()=>gridBusy,EndCallback:gridEnd},{InCallback:()=>comboBusy,EndCallback:comboEnd}];
const context={setTimeout,clearTimeout,setInterval,clearInterval,aspxGetControlCollection:()=>({ForEachControl:f=>controls.forEach(f),Get:()=>({})})};
const run=vm.runInNewContext('('+script+')',context);
(async()=>{
  const pending=run(900000).then(()=>{finished=true});
  gridBusy=false;gridEnd.fire();
  await new Promise(r=>setTimeout(r,20));
  assert.equal(finished,false,'One finished callback must not release a still-busy form');
  comboBusy=false;comboEnd.fire();
  await pending;
  assert.equal(finished,true);
  assert.equal(gridEnd.size()+comboEnd.size(),0,'EndCallback handlers must be removed after completion');
  console.log('Callback completion verified');
})().catch(e=>{console.error(e);process.exit(1)});
'''
        result = subprocess.run(['node','-e',test],input=json.dumps(script),text=True,capture_output=True,timeout=5)
        self.assertEqual(result.returncode,0,result.stderr)
        self.assertIn('Callback completion verified',result.stdout)


if __name__ == '__main__':
    unittest.main()
