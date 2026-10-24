const fs = require('fs');

async function testInspection() {
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
    const coll = aspxGetControlCollection();
    const ctrl = coll.Get('rpnlMain_ContentPlaceHolder1_btnExportExcel');
    const hasGetItemCount = typeof ctrl.GetItemCount === 'function';
    const count = hasGetItemCount ? ctrl.GetItemCount() : -1;
    const item0 = count > 0 ? ctrl.GetItem(0) : null;
    const subCount = item0 && typeof item0.GetItemCount === 'function' ? item0.GetItemCount() : -1;
    const subs = [];
    if (subCount > 0) {
      for (let i = 0; i < subCount; i++) {
        const s = item0.GetItem(i);
        subs.push({
          index: i,
          text: s.GetText ? s.GetText() : s.text,
          name: s.name,
          domId: 'rpnlMain_ContentPlaceHolder1_btnExportExcel_DXI0i' + i + '_'
        });
      }
    }
    return { hasGetItemCount, count, subCount, subs };
  })()
  `;

  const res = await evalJs(code);
  console.log(JSON.stringify(res, null, 2));
  ws.close();
}

testInspection().catch(console.error);
