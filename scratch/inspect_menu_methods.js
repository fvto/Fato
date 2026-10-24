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
    const coll = aspxGetControlCollection();
    const menu = coll.Get('rpnlMain_ContentPlaceHolder1_btnExportExcel');
    const search = coll.Get('rpnlMain_ContentPlaceHolder1_btnSearch');
    const item0 = menu ? menu.GetItem(0) : null;
    const sub0 = item0 && item0.GetItem(0) ? item0.GetItem(0) : null;
    
    // Also check DOM element of menu item
    const domItems = Array.from(document.querySelectorAll('.dxm-item, [id*="btnExportExcel"]')).map(el => ({
      id: el.id,
      text: el.innerText.trim(),
      onclick: el.getAttribute('onclick')
    }));

    return {
      menuMethods: menu ? Object.keys(menu).filter(k => typeof menu[k] === 'function') : [],
      searchMethods: search ? Object.keys(search).filter(k => typeof search[k] === 'function') : [],
      sub0Methods: sub0 ? Object.keys(sub0).filter(k => typeof sub0[k] === 'function') : [],
      domItems: domItems.slice(0, 10)
    };
  })()`);

  console.log(JSON.stringify(res, null, 2));
  ws.close();
})().catch(console.error);
