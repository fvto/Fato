(async () => {
  const targets = await (await fetch('http://localhost:9222/json')).json();
  const target = targets.find(x => x.type === 'page' && x.url.includes('sfc.chingluh.com'));
  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise(r => ws.onopen = r);
  
  const evalJs = (expr) => new Promise((resolve, reject) => {
    ws.onmessage = e => {
      const msg = JSON.parse(e.data);
      if (msg.id === 1) {
        if (msg.result?.exceptionDetails) reject(msg.result.exceptionDetails);
        else resolve(msg.result?.result?.value);
      }
    };
    ws.send(JSON.stringify({id: 1, method: 'Runtime.evaluate', params: {expression: expr, returnByValue: true}}));
  });

  const res = await evalJs(`(() => {
    const el = document.getElementById('rpnlMain_ContentPlaceHolder1_btnExportExcel_DXI0i0_');
    const searchEl = document.getElementById('rpnlMain_ContentPlaceHolder1_btnSearch');
    return {
      sub0Exists: !!el,
      sub0Id: el ? el.id : null,
      sub0Tag: el ? el.tagName : null,
      searchExists: !!searchEl,
      searchTag: searchEl ? searchEl.tagName : null
    };
  })()`);

  console.log(JSON.stringify(res, null, 2));
  ws.close();
})().catch(console.error);
