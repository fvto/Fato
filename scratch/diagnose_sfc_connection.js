// Read-only diagnostics for the application's existing local DOM connection.
const pending = new Map();
async function main() {
  const targets = await (await fetch('http://localhost:9222/json')).json();
  const target = targets.find(x => x.type === 'page' && x.url.includes('sfc.chingluh.com/vhmes/QA/QCHFPA.aspx'));
  if (!target) throw new Error('No SFC HFPA target');
  const socket = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve,reject)=>{socket.onopen=resolve;socket.onerror=reject;});
  socket.onmessage = event => {
    const message=JSON.parse(event.data);
    if (pending.has(message.id)) {pending.get(message.id)(message);pending.delete(message.id);}
  };
  function read(id,method,params={}) {
    return new Promise((resolve,reject)=>{
      const timer=setTimeout(()=>reject(new Error('Diagnostic timeout: '+method)),8000);
      pending.set(id,message=>{clearTimeout(timer);resolve(message.result ?? message.error);});
      socket.send(JSON.stringify({id,method,params}));
    });
  }
  try {
    const tree=await read(1,'Page.getFrameTree');
    const frame=tree.frameTree?.frame;
    console.log('Root frame:',JSON.stringify({url:frame?.url,unreachableUrl:frame?.unreachableUrl,mimeType:frame?.mimeType}));
    const expression=`(() => {const g=aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_grdExcel');return {ready:document.readyState,inCallback:g.InCallback(),inCallbackCode:String(g.InCallback),counts:Object.fromEntries(Object.keys(g).filter(k=>/callback|request/i.test(k)&&['string','number','boolean'].includes(typeof g[k])).map(k=>[k,g[k]])),resources:performance.getEntriesByType('resource').slice(-8).map(x=>({path:new URL(x.name).pathname,duration:Math.round(x.duration),size:x.transferSize}))};})()`;
    const result=await read(2,'Runtime.evaluate',{expression,returnByValue:true});
    console.log('SFC diagnostics:',JSON.stringify(result.result?.value ?? result.exceptionDetails));
  } finally {socket.close();}
}
main().catch(error=>{console.error(error.message);process.exitCode=1;});
