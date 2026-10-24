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
    const menu = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_btnExportExcel');
    // DevExpress menu event firing
    return {
      menuType: menu ? menu.constructor.name : null,
      keys: menu ? Object.keys(menu).filter(k => /click|item|do/i.test(k)) : []
    };
  })()
  `;

  const res = await evalJs(code);
  console.log(JSON.stringify(res, null, 2));
  ws.close();
})().catch(console.error);
