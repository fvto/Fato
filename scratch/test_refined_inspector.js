const fs = require('fs');

async function inspectSfcPage() {
  const targets = await (await fetch('http://localhost:9222/json')).json();
  const target = targets.find(x => x.type === 'page' && x.url.includes('sfc.chingluh.com'));
  if (!target) return { connected: false, error: 'Chưa tìm thấy tab SFC trong Chrome.' };

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

  const inspectionCode = `
  (() => {
    const coll = (typeof aspxGetControlCollection === 'function') ? aspxGetControlCollection() : null;
    const elements = coll ? coll.elements : {};
    
    const pageInfo = {
      title: document.title,
      url: location.href,
      pathname: location.pathname
    };

    // 1. Phân loại Ô ngày (Date Edits)
    const dateInputs = [];
    // 2. Phân loại Menu xổ xuống & ComboBox (Dropdowns)
    const dropdowns = [];
    // 3. Phân loại Nút hành động (Buttons)
    const buttons = [];
    // 4. Phân loại Checkboxes
    const checkboxes = [];

    // Helper: kiểm tra phần tử có hiển thị trên màn hình không
    function isVisible(el) {
      if (!el) return false;
      const style = window.getComputedStyle(el);
      return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetWidth > 0 && el.offsetHeight > 0;
    }

    // A. Quét qua DevExpress Control Collection
    if (coll) {
      for (let key in elements) {
        const ctrl = elements[key];
        if (!ctrl || !ctrl.name) continue;
        const name = ctrl.name;

        // Bỏ qua các control nội bộ DDD, FNP, popup phụ
        if (name.includes('_DDD_') || name.includes('_FNP') || name.includes('DXFREditor') || name.includes('DXHFP')) continue;

        // Date Edit
        if (name.toLowerCase().includes('date') || name.includes('dtDate')) {
          if (typeof ctrl.GetText === 'function') {
            const domInput = document.getElementById(name + '_I') || document.getElementById(name);
            dateInputs.push({
              id: name,
              input_id: name + '_I',
              name: name,
              label: name.toLowerCase().includes('from') ? 'Từ ngày (Date From)' : (name.toLowerCase().includes('to') ? 'Đến ngày (Date To)' : 'Ngày'),
              value: ctrl.GetText(),
              type: 'date'
            });
          }
        }

        // CheckBox (ví dụ NIKE)
        else if (typeof ctrl.GetChecked === 'function') {
          const domEl = document.getElementById(name);
          const label = domEl ? (domEl.innerText || domEl.getAttribute('title') || name.split('_').pop()) : name;
          checkboxes.push({
            id: name,
            name: name,
            label: label.trim().replace(/^\\s+/, ''),
            checked: ctrl.GetChecked(),
            type: 'checkbox'
          });
        }

        // ComboBox (Ví dụ Chuyền/Xưởng cboDeptFrom)
        else if (typeof ctrl.GetItemCount === 'function' && !ctrl.items && ctrl.name.toLowerCase().includes('cbo')) {
          const items = [];
          for (let i = 0; i < ctrl.GetItemCount(); i++) {
            const it = ctrl.GetItem(i);
            if (it && it.text) {
              items.push({ value: it.value || it.text, text: it.text.trim() });
            }
          }
          if (items.length > 0) {
            dropdowns.push({
              id: name,
              name: name,
              type: 'combobox',
              label: 'Chọn Chuyền / Xưởng hoặc Tiêu chí',
              selected: ctrl.GetText ? ctrl.GetText() : (items[0] ? items[0].text : ''),
              item_count: items.length,
              options: items.slice(0, 200) // giới hạn tránh quá tải
            });
          }
        }

        // Menu xổ xuống (Ví dụ Export Excel Menu với 3 sub-items trên COPQ Report)
        else if (ctrl.items && ctrl.items.length > 0) {
          const mainItem = ctrl.GetItem(0);
          if (mainItem && typeof mainItem.GetItemCount === 'function' && mainItem.GetItemCount() > 0) {
            const subOptions = [];
            for (let j = 0; j < mainItem.GetItemCount(); j++) {
              const sub = mainItem.GetItem(j);
              const subDomId = name + '_DXI0i' + j + '_';
              subOptions.push({
                index: j,
                id: subDomId,
                name: sub.name || ('SUB_' + j),
                text: (sub.GetText ? sub.GetText() : sub.text || '').trim(),
                value: sub.name || (sub.GetText ? sub.GetText() : '')
              });
            }
            if (subOptions.length > 0) {
              dropdowns.push({
                id: name,
                name: name,
                type: 'menu_dropdown',
                label: 'Tùy chọn Xuất file / Dạng Báo cáo',
                selected: subOptions[0].text,
                item_count: subOptions.length,
                options: subOptions
              });
            }
          }
        }

        // Buttons DevExpress
        else if (name.toLowerCase().includes('btn') || (typeof ctrl.DoClick === 'function' && !name.includes('RoundPanel'))) {
          const domEl = document.getElementById(name);
          let label = ctrl.GetText ? ctrl.GetText() : '';
          if (!label && domEl) {
            label = domEl.innerText || domEl.getAttribute('title') || '';
          }
          // Gán nhãn thông minh nếu icon không có chữ
          if (!label) {
            if (name.toLowerCase().includes('search')) label = 'Tìm kiếm dữ liệu';
            else if (name.toLowerCase().includes('export')) label = 'Xuất file';
            else if (name.toLowerCase().includes('flowchart') || name.toLowerCase().includes('home')) label = 'Về Trang chủ';
            else label = name.split('_').pop();
          }

          let icon = 'bolt';
          let btnType = 'primary';
          const lower = (name + ' ' + label).toLowerCase();
          if (lower.includes('search') || lower.includes('tim')) {
            icon = 'search';
            btnType = 'search';
          } else if (lower.includes('excel') || lower.includes('export') || lower.includes('xuat')) {
            icon = 'file-excel';
            btnType = 'export';
          } else if (lower.includes('flowchart') || lower.includes('home')) {
            icon = 'home';
            btnType = 'nav';
          }

          buttons.push({
            id: name,
            name: name,
            label: label.trim(),
            icon: icon,
            btn_type: btnType
          });
        }
      }
    }

    // B. Quét bổ sung các thẻ HTML thuần nếu trang không dùng DevExpress
    document.querySelectorAll('select').forEach((sel, idx) => {
      if (!dropdowns.some(d => d.id === sel.id)) {
        const options = Array.from(sel.options).map(o => ({ value: o.value, text: o.text.trim() }));
        if (options.length > 0) {
          dropdowns.push({
            id: sel.id || ('select_' + idx),
            name: sel.name || sel.id,
            type: 'html_select',
            label: sel.getAttribute('title') || 'Danh mục lựa chọn',
            selected: sel.options[sel.selectedIndex]?.text || '',
            item_count: options.length,
            options: options.slice(0, 150)
          });
        }
      }
    });

    return {
      success: true,
      page: pageInfo,
      date_inputs: dateInputs,
      dropdowns: dropdowns,
      buttons: buttons,
      checkboxes: checkboxes
    };
  })()
  `;

  try {
    const data = await evalJs(inspectionCode);
    ws.close();
    return data;
  } catch (err) {
    ws.close();
    return { success: false, error: String(err) };
  }
}

inspectSfcPage().then(res => console.log(JSON.stringify(res, null, 2))).catch(console.error);
