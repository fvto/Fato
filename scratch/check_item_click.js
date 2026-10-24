(async () => {
  const targets = await (await fetch('http://localhost:9222/json')).json();
  const target = targets.find(x => x.type === 'page' && x.url.includes('sfc.chingluh.com'));
  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise(r => ws.onopen = r);

  const evalJs = (expr) => new Promise((resolve, reject) => {
    ws.onmessage = e => resolve(JSON.parse(e.data).result?.result?.value);
    ws.send(JSON.stringify({id: 1, method: 'Runtime.evaluate', params: {expression: expr, returnByValue: true}}));
  });

  const code = `
  (() => {
    const el = document.getElementById('rpnlMain_ContentPlaceHolder1_btnExportExcel_DXI0i1_');
    const menu = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_btnExportExcel');
    return {
      elOuter: el ? el.outerHTML.substring(0, 200) : null,
      parentOuter: el && el.parentElement ? el.parentElement.outerHTML.substring(0, 200) : null,
      menuItemClick: typeof menu.ItemClick,
      menuOnItemClick: typeof menu.onItemClick
    };
  })()
  `;

  const res = await evalJs(code);
  console.log(JSON.stringify(res, null, 2));
  ws.close();
})().catch(console.error);
