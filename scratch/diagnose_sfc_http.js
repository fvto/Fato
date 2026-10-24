// Read-only network checks using the dedicated SFC session; never print cookies.
const pending=new Map();
async function main() {
  const info=await (await fetch('http://localhost:9222/json/version')).json();
  const socket=new WebSocket(info.webSocketDebuggerUrl);
  await new Promise((resolve,reject)=>{socket.onopen=resolve;socket.onerror=reject;});
  socket.onmessage=event=>{const message=JSON.parse(event.data);if(pending.has(message.id)){pending.get(message.id)(message);pending.delete(message.id);}};
  function read(id,method,params={}) {
    return new Promise((resolve,reject)=>{
      const timer=setTimeout(()=>reject(new Error('Browser diagnostic timeout: '+method)),8000);
      pending.set(id,message=>{clearTimeout(timer);message.error?reject(new Error(message.error.message)):resolve(message.result);});
      socket.send(JSON.stringify({id,method,params}));
    });
  }
  try {
    const version=await read(1,'Browser.getVersion');
    console.log('Browser transport responds:',version.product);
    const all=await read(2,'Storage.getCookies');
    const cookies=all.cookies.filter(c=>c.domain.replace(/^\./,'')==='sfc.chingluh.com');
    console.log('SFC cookie names:',cookies.map(c=>c.name).join(', '));
    async function check(label,path,selected) {
      const start=Date.now();
      try {
        const response=await fetch('https://sfc.chingluh.com'+path,{
          headers:{Cookie:selected.map(c=>`${c.name}=${c.value}`).join('; ')},signal:AbortSignal.timeout(8000)
        });
        const body=await response.text();
        console.log(label,JSON.stringify({status:response.status,path:new URL(response.url).pathname,ms:Date.now()-start,bytes:body.length,hasHfpaControls:body.includes('ContentPlaceHolder1_ckbNIKE')}));
      } catch(error) {console.log(label,error.name,Date.now()-start);}
    }
    await Promise.all([
      check('Existing-session login route','/vhmes/Login.aspx',cookies),
      check('Isolated HTTP session HFPA','/vhmes/QA/QCHFPA.aspx?tree_id=Q504',cookies.filter(c=>!/^asp\.net_sessionid$/i.test(c.name)))
    ]);
  } finally {socket.close();}
}
main().catch(error=>{console.error(error.message);process.exitCode=1;});
