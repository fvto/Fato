// Dashboard synchronization tests with SFC responses mocked; no live browser actions.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const elements = new Map();
function element() {
    return {innerHTML:'', innerText:'', options:[], style:{}, addEventListener(){},
        classList:{add(){}, remove(){}, toggle(){}}};
}
const context = vm.createContext({console, setTimeout(){}, setInterval(){},
    localStorage:{getItem(){return null;}}, window:{addEventListener(){}},
    document:{getElementById(id){
        if (!elements.has(id)) elements.set(id, element());
        return elements.get(id);
    }}});
const html = fs.readFileSync(path.join(__dirname, '../templates/index.html'), 'utf8');
vm.runInContext([...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)]
    .map(match => match[1]).join('\n'), context);
context.updatePlantButtonsUI = context.showToast = context.addLog = () => {};
let renders = 0;
context.renderAdaptiveControls = () => { renders++; };
let status = {connected:true, url:'https://example.test/copq', title:'COPQ'};
let scans = 0;
let failScan = false;
let delayScan = null;
context.fetch = async url => {
    if (url.startsWith('/api/status')) return {json:async()=>({...status})};
    if (url === '/api/navigate') return {json:async()=>({success:true, title:status.title})};
    assert.equal(url, '/api/page/inspect');
    scans++;
    const inspection = {page:{url:status.url, title:status.title}, dropdowns:[], buttons:[], date_inputs:[]};
    if (delayScan) await delayScan;
    return {json:async()=>failScan ? {success:false} : {success:true, inspection}};
};
const run = code => vm.runInContext(code, context);
(async () => {
    await run('refreshStatus()');
    assert.equal(scans, 1, 'Initial connected status must automatically scan');
    run("selectedDropdownOptions.saved={value:'chosen'}");
    await run('refreshStatus()');
    assert.equal(scans, 1, 'Unchanged status must preserve edits and avoid repeated scans');
    assert.equal(run('selectedDropdownOptions.saved.value'), 'chosen');

    status = {...status, url:'https://example.test/hfpa', title:'HFPA'};
    await run('refreshStatus()');
    assert.equal(scans, 2, 'External page changes must automatically scan');
    assert.equal(run('selectedDropdownOptions.saved'), undefined, 'Old module selections must reset');
    assert.equal(elements.get('current-page-name').innerText, 'HFPA');

    status.command_running = true;
    status.title = 'Updated HFPA';
    await run('refreshStatus()');
    assert.equal(scans, 2, 'Polling must not inspect while an export is running');
    status.command_running = false;
    await run('refreshStatus()');
    assert.equal(scans, 3, 'Finishing a command must refresh the controls');
    status.processing = true;
    await run('refreshStatus()');
    status.processing = false;
    await run('refreshStatus()');
    assert.equal(scans, 4, 'Finishing work on the same URL must refresh too');

    status = {...status, url:'https://example.test/mqaa', title:'MQAA'};
    failScan = true;
    await run('refreshStatus()');
    failScan = false;
    await run('refreshStatus()');
    assert.equal(elements.get('current-page-name').innerText, 'MQAA', 'Failed scans must retry on the next poll');

    const beforeNavigation = scans;
    await run("jumpToQmsModule('https://example.test/mqaa', 'MQAA')");
    await run("navTo('mqaa')");
    assert.equal(scans, beforeNavigation + 2, 'Both navigation methods must refresh even when reopening the same page');

    let release;
    delayScan = new Promise(resolve => { release = resolve; });
    const previousRenders = renders;
    const pending = run('scanCurrentPage(false)');
    status = {...status, url:'https://example.test/new', title:'New module'};
    run('resetAdaptiveControls()');
    delayScan = null;
    release();
    await pending;
    assert.equal(renders, previousRenders + 1, 'An obsolete scan must be discarded and replaced');
    assert.equal(elements.get('current-page-name').innerText, 'New module');

    status.connected = false;
    await run('refreshStatus()');
    assert.equal(run('currentInspectionData'), null, 'Disconnect must clear stale controls');
    status.connected = true;
    await run('refreshStatus()');
    assert.equal(elements.get('current-page-name').innerText, 'New module', 'Reconnect must automatically scan');
    console.log('Auto scan: navigation, external changes, busy completion, retry, stale responses and reconnect passed');
})().catch(error => { console.error(error); process.exitCode = 1; });
