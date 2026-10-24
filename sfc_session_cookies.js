// Private IPC helper. Cookie values are consumed in memory by the Python client.
async function main() {
    const statusOnly = process.argv.includes('--status');
    const version = await (await fetch('http://localhost:9222/json/version', {signal:AbortSignal.timeout(5000)})).json();
    const socket = new WebSocket(version.webSocketDebuggerUrl);
    const result = await new Promise((resolve,reject) => {
        let cookies = [];
        const timer = setTimeout(()=>reject(new Error('SFC session connection timed out')),5000);
        socket.onopen=()=>socket.send(JSON.stringify({id:1,method:'Storage.getCookies'}));
        socket.onerror=()=>{clearTimeout(timer);reject(new Error('Cannot read the dedicated SFC session'));};
        socket.onmessage=event=>{
            const message=JSON.parse(event.data);
            if(message.id!==1 && message.id!==2) return;
            if(message.error) { clearTimeout(timer); reject(new Error('SFC session read failed')); return; }
            if(message.id===1) {
                cookies=message.result.cookies.filter(c=>c.domain.replace(/^\./,'')==='sfc.chingluh.com' && !/^asp\.net_sessionid$/i.test(c.name));
                if(statusOnly) { socket.send(JSON.stringify({id:2,method:'Target.getTargets'})); return; }
                clearTimeout(timer);
                resolve(cookies);
            } else {
                clearTimeout(timer);
                const page=message.result.targetInfos.find(t=>{
                    try { return t.type==='page' && new URL(t.url).hostname==='sfc.chingluh.com'; }
                    catch { return false; }
                });
                resolve({connected:!!page,url:page?.url||'',title:page?.title||'SFC',
                    requires_login:!!page && /login\.aspx/i.test(page.url),
                    current_plant:cookies.find(c=>c.name==='MyFactoryName')?.value||null});
            }
        };
    }).finally(()=>socket.close());
    process.stdout.write(JSON.stringify(result));
}
main().catch(error=>{process.stderr.write(error.message);process.exitCode=1;});
