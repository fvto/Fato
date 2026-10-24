// Execute the served dashboard's JavaScript against its real HTTP API.
// This is a JavaScript/API integration test; no browser or computer UI is controlled.
const fs=require('node:fs');
const vm=require('node:vm');
const assert=require('node:assert/strict');
const base='http://localhost:5050';
function element(){
 const attrs=new Map();
 return {children:[],innerHTML:'',innerText:'',textContent:'',value:'',style:{},classList:{add(){},remove(){},toggle(){}},
 addEventListener(){},setAttribute(k,v){attrs.set(k,v);},getAttribute(k){return attrs.has(k)?attrs.get(k):this[k];},
 appendChild(v){this.children.push(v);},prepend(v){this.children.unshift(v);},replaceChildren(){this.children=[];},querySelectorAll(){return this.children;}};
}
(async()=>{
 const html=await(await fetch(base)).text();
 const original='HFPA_VH_NIKE_2026-01-01_2026-10-31_Combined.xlsx';
 assert(html.includes(original),'Existing VH combined result must remain downloadable after restart');
 assert(html.includes('Kết quả đã lưu'));
 const elements=new Map(),storage=new Map(),calls=[];
 const document={body:element(),createElement:element,querySelectorAll:()=>[],getElementById(id){if(!elements.has(id))elements.set(id,element());return elements.get(id);}};
 const context=vm.createContext({console,TextDecoder,TextEncoder,setTimeout(){},setInterval(){},document,window:{addEventListener(){}},
 localStorage:{getItem:k=>storage.get(k),setItem:(k,v)=>storage.set(k,v),removeItem:k=>storage.delete(k)},
 fetch:(url,options)=>{calls.push({url,method:options?.method||'GET',body:options?.body});return fetch(base+url,options);}});
 const script=[...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)].map(x=>x[1]).join('\n');
 vm.runInContext(script,context);
 if(process.argv.includes('--reconnect-only')){
  const id=html.match(/data-last-job="([^"]+)"/)[1];
  document.getElementById('ai-prompt-input').value='Khôi phục kết quả đã lưu';
  await vm.runInContext('executeCommand('+JSON.stringify(id)+')',context);
  assert.equal(calls.filter(x=>x.method==='POST').length,0);
  assert.equal(document.getElementById('command-files').children.length,1);
  const started=performance.now();
  const status=await(await fetch(base+'/api/status')).json();
  const elapsed=performance.now()-started;
  assert.equal(status.connected,true);assert.equal(status.current_plant,'VH');
  assert(elapsed<3000,'Status must not wait for the frozen page');
  const report=JSON.parse(fs.readFileSync('scratch/dashboard_live_integration.json','utf8'));
  Object.assign(report,{restored_after_server_restart:true,status_response_ms:Math.round(elapsed),reconnect_after_restart_did_not_rerun:true});
  fs.writeFileSync('scratch/dashboard_live_integration.json',JSON.stringify(report,null,2));
  console.log('After server restart: saved files restored, zero export POSTs; status '+Math.round(elapsed)+'ms');
  return;
 }
 document.getElementById('date-from').value='2026-10-24';
 document.getElementById('date-to').value='2026-10-24';
 document.getElementById('factory-select').value='';
 const results=[];
 for(const nike of [true,false]){
  document.getElementById('nike-toggle').checked=nike;
  vm.runInContext('toggleNike()',context);
  await vm.runInContext("triggerAction('excel')",context);
  const id=storage.get('sfc-last-job');
  const job=await(await fetch(base+'/api/command/jobs/'+id)).json();
  assert.equal(job.status,'done',JSON.stringify(job.events.at(-1)));
  assert.equal(job.events.find(x=>x.intent).intent.nike,nike);
  assert.deepEqual(job.events.find(x=>x.intent).intent.target_plants,['VH']);
  const links=document.getElementById('command-files').children;
  assert.equal(links.length,1,'Dashboard must expose a download link');
  assert.equal(document.getElementById('btn-prompt-exec').disabled,false);
  const item=job.events.at(-1).files[0];
  assert.equal(links[0].getAttribute('href'),item.url);
  const download=await fetch(base+item.url);assert.equal(download.status,200);
  assert.equal((await download.arrayBuffer()).byteLength,item.size);
  results.push({nike,job_id:id,file:item});
  console.log('Dashboard Excel '+(nike?'NIKE':'Internal')+': done, '+item.rows+' rows, download HTTP 200');
 }
 const posts=calls.filter(x=>x.method==='POST').length;
 await vm.runInContext('executeCommand('+JSON.stringify(results[1].job_id)+')',context);
 assert.equal(calls.filter(x=>x.method==='POST').length,posts,'Reopening results must not rerun the export');
 assert(calls.filter(x=>x.method==='POST').every(x=>x.url==='/api/command/stream'));
 const refreshed=await(await fetch(base)).text();assert(refreshed.includes(original));
 fs.writeFileSync('scratch/dashboard_live_integration.json',JSON.stringify({status:'passed',results,original_vh_result_still_visible:true,reconnect_did_not_rerun:true},null,2));
 console.log('Saved VH result and reconnect without another export: passed');
})().catch(error=>{console.error(error);process.exitCode=1;});
