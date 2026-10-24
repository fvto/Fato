// Pure JavaScript tests for streamed progress; no browser or user UI is controlled.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const elements = new Map();
function node() {
    const attrs = new Map();
    return {children: [], innerHTML: '', innerText: '', textContent: '', value: '', style: {},
        classList: {add() {}, remove() {}, toggle() {}}, addEventListener() {},
        setAttribute(k,v) {attrs.set(k,v);}, getAttribute(k) {return attrs.has(k) ? attrs.get(k) : this[k];},
        appendChild(v) {this.children.push(v);}, prepend(v) {this.children.unshift(v);},
        replaceChildren() {this.children=[];}, querySelectorAll() {return this.children;}};
}
const storage = new Map();
const context = vm.createContext({console, TextDecoder, TextEncoder, ReadableStream,
    setTimeout() {}, setInterval() {},
    localStorage: {getItem:k=>storage.get(k), setItem:(k,v)=>storage.set(k,v), removeItem:k=>storage.delete(k)},
    document: {body:node(),createElement:node,querySelectorAll:()=>[],getElementById(id) {
        if (!elements.has(id)) elements.set(id,node());
        return elements.get(id);
    }}, window:{addEventListener() {}}});
const html = fs.readFileSync(path.join(__dirname,'../templates/index.html'),'utf8');
const script = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)].map(x=>x[1]).join('\n');
vm.runInContext(script,context);
function response(events) {
    const bytes = new TextEncoder().encode(events.map(x=>'data: '+JSON.stringify(x)+'\n\n').join(''));
    return {ok:true,body:new ReadableStream({start(c) {
        // Include splits inside Vietnamese UTF-8 text and JSON boundaries.
        for(let i=0;i<bytes.length;i+=7)c.enqueue(bytes.slice(i,i+7));
        c.close();
    }})};
}
(async()=>{
    context.response=response([
        {step:1,status:'success',msg:'Đã nhận diện',job_id:'example',intent:{target_plant:'VH',target_plants:['VH','VH2','JV','JV2'],nike:true,monthly:true,combine:true}},
        {step:5,status:'error',msg:'Tải file thất bại'}]);
    await vm.runInContext('consumeCommandStream(response)',context);
    assert.match(elements.get('intent-chips-bar').innerHTML,/VH, VH2, JV, JV2/);
    assert.match(elements.get('intent-chips-bar').innerHTML,/NIKE \(External\)/);
    assert.equal(elements.get('exec-status-badge').textContent,'CHƯA HOÀN TẤT');
    assert.equal(storage.get('sfc-last-job'),'example');
    vm.runInContext('commandTerminal=false',context);
    const output={name:'VH combined.xlsx',url:'/api/exports/example/VH.xlsx',rows:100};
    context.response=response([{step:4,status:'success',files:[output]},{step:5,status:'done',files:[output]}]);
    await vm.runInContext('consumeCommandStream(response)',context);
    assert.equal(elements.get('command-files').children.length,1,'Repeated final events must not duplicate download links');
    vm.runInContext('commandTerminal=false',context);
    context.response=response([{step:1,status:'running',msg:'Started'}]);
    await assert.rejects(vm.runInContext('consumeCommandStream(response)',context),/bị ngắt/);
    vm.runInContext('commandBusy=false;commandTerminal=false',context);
    context.fetch=()=>{throw new Error('Mode selection must not wait on Chrome');};
    elements.get('nike-toggle').checked=true;
    vm.runInContext('toggleNike()',context);
    assert.equal(storage.get('sfc-nike-mode'),'true');
    assert.equal(elements.get('nike-mode-text').innerText,'External (NIKE)');
    vm.runInContext('commandBusy=true',context);
    elements.get('nike-toggle').checked=false;
    vm.runInContext('toggleNike()',context);
    assert.equal(elements.get('nike-toggle').checked,true,'Cannot change mode during an export');
    vm.runInContext('commandBusy=false',context);
    context.document.getElementById('date-from').value='2026-10-24';
    context.document.getElementById('date-to').value='2026-10-24';
    context.document.getElementById('factory-select').value='';
    const requests=[];
    context.fetch=async(url,options)=>{
        requests.push({url,options});
        return response([{step:1,status:'success',intent:{target_plant:'VH',nike:true}},
                         {step:5,status:'done',files:[output]}]);
    };
    await vm.runInContext("triggerAction('excel')",context);
    assert.equal(requests[0].url,'/api/command/stream');
    assert.match(JSON.parse(requests[0].options.body).prompt,/HFPA VH.*2026-10-24.*NIKE/);
    await vm.runInContext("executeCommand('existing-job')",context);
    assert.equal(requests[1].url,'/api/command/jobs/existing-job/stream');
    assert.equal(requests[1].options,undefined,'Reconnect must not submit another export');
    context.fetch=async()=>({json:async()=>({connected:false,form:{nike:false}})});
    await vm.runInContext('refreshStatus()',context);
    assert.equal(elements.get('nike-toggle').checked,true,'An old browser filter must not overwrite dashboard export mode');
    console.log('Dashboard: stream decoding, result links, NIKE mode, Excel button, reconnect and stale browser state passed');
})().catch(error=>{console.error(error);process.exitCode=1;});
