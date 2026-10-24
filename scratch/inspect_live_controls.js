const fs = require('fs');

(async () => {
  const targets = await (await fetch('http://localhost:9222/json')).json();
  const target = targets.find(x => x.type === 'page' && x.url.includes('sfc.chingluh.com'));
  if (!target) {
    console.log('No sfc page found');
    return;
  }
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

  const script = `
  (() => {
    const res = {
      title: document.title,
      url: location.href,
      controls: [],
      buttons: [],
      inputs: [],
      menus: [],
      dropdowns: []
    };

    // 1. DevExpress controls
    if (typeof aspxGetControlCollection === 'function') {
      const coll = aspxGetControlCollection();
      for (let k in coll.elements) {
        const c = coll.elements[k];
        if (!c || !c.name) continue;
        
        const info = {
          name: c.name,
          type: c.constructor ? c.constructor.name : typeof c,
          text: (typeof c.GetText === 'function') ? c.GetText() : null
        };

        // If Menu / PopupMenu
        if (typeof c.GetItemCount === 'function') {
          info.itemCount = c.GetItemCount();
          info.items = [];
          for (let i = 0; i < c.GetItemCount(); i++) {
            const item = c.GetItem(i);
            if (item) {
              const subItems = [];
              if (typeof item.GetItemCount === 'function') {
                for (let j = 0; j < item.GetItemCount(); j++) {
                  const s = item.GetItem(j);
                  if (s) subItems.push({ index: j, text: s.GetText ? s.GetText() : s.text, name: s.name });
                }
              }
              info.items.push({
                index: i,
                text: item.GetText ? item.GetText() : item.text,
                name: item.name,
                subItems
              });
            }
          }
        }
        res.controls.push(info);
      }
    }

    // 2. DOM buttons and clickable actions
    document.querySelectorAll('input[type="button"], input[type="submit"], button, .dxbButton, a[onclick]').forEach((el, idx) => {
      const text = el.innerText || el.value || el.title || el.getAttribute('aria-label') || '';
      if (text.trim() || el.id) {
        res.buttons.push({
          id: el.id,
          tag: el.tagName,
          type: el.type || '',
          text: text.trim().replace(/\\s+/g, ' '),
          className: el.className
        });
      }
    });

    // 3. Select / Dropdowns in DOM
    document.querySelectorAll('select').forEach(sel => {
      res.dropdowns.push({
        id: sel.id,
        name: sel.name,
        options: Array.from(sel.options).map(o => ({ value: o.value, text: o.text.trim() }))
      });
    });

    // 4. Date inputs
    document.querySelectorAll('input[type="text"], input:not([type])').forEach(inp => {
      if (inp.id && (inp.id.toLowerCase().includes('date') || inp.className.includes('DateEdit') || inp.name?.toLowerCase().includes('date'))) {
        res.inputs.push({
          id: inp.id,
          name: inp.name,
          value: inp.value,
          className: inp.className
        });
      }
    });

    return res;
  })()
  `;

  const data = await evalJs(script);
  console.log(JSON.stringify(data, null, 2));
  ws.close();
})().catch(console.error);
