import sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    # Inspect menuRightTop control
    res = page.evaluate("""() => {
        const out = {};
        try {
            // Find menu in aspxGetControlCollection
            const coll = aspxGetControlCollection();
            for (let k in coll.elements) {
                if (k.includes('menuRightTop') || k.includes('cboFactory')) {
                    const c = coll.elements[k];
                    out[k] = {
                        type: c ? c.constructor.name : typeof c,
                        itemCount: c && typeof c.GetItemCount === 'function' ? c.GetItemCount() : null
                    };
                }
            }
        } catch(e) {
            out.error = e.toString();
        }
        
        // Also check if there's an ItemClick handler or URL on the items
        const items = [];
        for (let i = 0; i <= 6; i++) {
            const el = document.getElementById(`wucHeaderMainTop_TopPanel_menuRightTop_DXI0i${i}_`);
            if (el) {
                items.push({
                    idx: i,
                    id: el.id,
                    text: el.innerText.trim(),
                    onclick: el.getAttribute('onclick'),
                    html: el.outerHTML
                });
            }
        }
        out.items = items;
        return out;
    }""")
    print("DevExpress menu control info:", res)
