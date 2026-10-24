import sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    page.goto("https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx")
    page.wait_for_load_state("networkidle")
    
    info = page.evaluate("""() => {
        const menu = aspxGetControlCollection().Get('wucHeaderMainTop_TopPanel_menuRightTop');
        // Check event handler for ItemClick
        return {
            name: menu ? menu.name : null,
            uniqueID: menu ? menu.uniqueID : null,
            itemClickHandlers: menu && menu.ItemClick ? menu.ItemClick.fireList.map(f => f.toString()) : []
        };
    }""")
    print("ItemClick handlers:", info)
