import json
import time
import sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

CDP_URL = "http://localhost:9222"

def inspect_active_page():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP_URL, timeout=10000)
        page = next((pg for ctx in browser.contexts for pg in ctx.pages
                     if 'sfc.chingluh.com' in pg.url and '/vhmes/' in pg.url), None)
        if not page:
            browser.close()
            return {"success": False, "error": "Chưa tìm thấy tab SFC đang mở."}

        eval_script = """() => {
            const coll = (typeof aspxGetControlCollection === 'function') ? aspxGetControlCollection() : null;
            const elements = coll ? coll.elements : {};

            const pageInfo = {
                title: document.title,
                url: location.href,
                pathname: location.pathname
            };

            const dateInputs = [];
            const dropdowns = [];
            const buttons = [];
            const checkboxes = [];

            if (coll) {
                for (let key in elements) {
                    const ctrl = elements[key];
                    if (!ctrl || !ctrl.name) continue;
                    const name = ctrl.name;

                    // Bỏ qua các control nội bộ DDD, FNP, popup phụ
                    if (name.includes('_DDD_') || name.includes('_FNP') || name.includes('DXFREditor') || name.includes('DXHFP')) continue;

                    // 1. Date Edits
                    if (name.toLowerCase().includes('date') || name.includes('dtDate')) {
                        if (typeof ctrl.GetText === 'function') {
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

                    // 2. Menu Dropdowns (DevExpress ASPxMenu with sub-items)
                    else if (typeof ctrl.GetItemCount === 'function' && ctrl.GetItemCount() > 0) {
                        const item0 = ctrl.GetItem(0);
                        const subCount = (item0 && typeof item0.GetItemCount === 'function') ? item0.GetItemCount() : 0;
                        if (subCount > 0) {
                            const subOptions = [];
                            for (let j = 0; j < subCount; j++) {
                                const sub = item0.GetItem(j);
                                const subDomId = name + '_DXI0i' + j + '_';
                                const subText = (sub.GetText ? sub.GetText() : sub.text || '').trim();
                                subOptions.push({
                                    index: j,
                                    id: subDomId,
                                    name: sub.name || ('SUB_' + j),
                                    text: subText || ('Tùy chọn ' + (j + 1)),
                                    value: sub.name || subText
                                });
                            }
                            dropdowns.push({
                                id: name,
                                name: name,
                                type: 'menu_dropdown',
                                label: name.toLowerCase().includes('export') ? 'Tùy chọn Xuất báo cáo / File' : 'Menu thả xuống',
                                selected: subOptions[0] ? subOptions[0].text : '',
                                item_count: subOptions.length,
                                options: subOptions
                            });
                        }
                        // Nếu là ComboBox (cboDeptFrom...)
                        else if (name.toLowerCase().includes('cbo') || name.toLowerCase().includes('dept') || name.toLowerCase().includes('combo')) {
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
                                    options: items.slice(0, 200)
                                });
                            }
                        }
                    }

                    // 3. Checkboxes (Loại trừ các nút bấm có method GetChecked)
                    else if (typeof ctrl.GetChecked === 'function' && (name.includes('ckb') || name.toLowerCase().includes('check'))) {
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

                    // 4. Buttons (DevExpress ASPxButton, RoundPanel buttons, search, export...)
                    else if (name.toLowerCase().includes('btn') || (typeof ctrl.DoClick === 'function' && !name.includes('RoundPanel'))) {
                        const domEl = document.getElementById(name);
                        let label = ctrl.GetText ? ctrl.GetText() : '';
                        if (!label && domEl) {
                            label = domEl.innerText || domEl.getAttribute('title') || '';
                        }
                        if (!label) {
                            if (name.toLowerCase().includes('search')) label = 'Tìm kiếm dữ liệu';
                            else if (name.toLowerCase().includes('export') || name.toLowerCase().includes('excel')) label = 'Xuất file Excel';
                            else if (name.toLowerCase().includes('flowchart') || name.toLowerCase().includes('home')) label = 'Trang chủ FlowChart';
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

                        // Nếu nút này đã là menu dropdown thì không cần lặp lại vào danh sách nút đơn
                        const isAlreadyMenu = dropdowns.some(d => d.id === name);
                        if (!isAlreadyMenu) {
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
            }

            // Quét bổ sung các thẻ HTML thuần
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
                page: pageInfo,
                date_inputs: dateInputs,
                dropdowns: dropdowns,
                buttons: buttons,
                checkboxes: checkboxes
            };
        }"""

        data = page.evaluate(eval_script)
        browser.close()
        return {"success": True, "inspection": data}

if __name__ == "__main__":
    res = inspect_active_page()
    print(json.dumps(res, indent=2, ensure_ascii=False))
