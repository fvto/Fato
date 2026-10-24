import sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    # Ensure on QCHFPA
    if "QCHFPA" not in page.url:
        page.goto("https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504")
        page.wait_for_load_state("load")
        
    res = page.evaluate("""() => {
        const cbo = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_cboDeptFrom');
        if (!cbo) return { error: 'cbo not found' };
        
        // Select VL - D3
        let matched = null;
        for (let i = 0; i < cbo.GetItemCount(); i++) {
            const it = cbo.GetItem(i);
            if (it.text.includes('VL - D3')) {
                cbo.SetSelectedIndex(i);
                matched = { index: i, text: it.text, val: it.value, currText: cbo.GetText() };
                break;
            }
        }
        return matched;
    }""")
    print("DevExpress cbo set test:", res)
    page.screenshot(path="scratch/cbo_test.png")
