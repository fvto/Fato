import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    print("Current URL before click:", page.url)
    
    # Check how menuRightTop handles ItemClick
    click_script = """() => {
        const menu = aspxGetControlCollection().Get('wucHeaderMainTop_TopPanel_menuRightTop');
        if (menu) {
            // Find item VH3
            const root = menu.GetItem(0); // Root "VH"
            console.log('Root item:', root.text);
            for (let i = 0; i < root.GetItemCount(); i++) {
                const sub = root.GetItem(i);
                console.log('Sub item', i, sub.text);
            }
            return {
                rootText: root.text,
                subCount: root.GetItemCount(),
                hasItemClick: typeof menu.ItemClick !== 'undefined'
            };
        }
        return { error: 'menu not found' };
    }"""
    print("Menu structure:", page.evaluate(click_script))
    
    # Try clicking the VH3 item in DOM
    # First hover or click the root item to open dropdown if needed, or click item directly
    print("\nClicking root menu to open...")
    page.click("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0_")
    time.sleep(1)
    
    print("Checking if sub items are visible...")
    vh3_el = page.query_selector("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0i2_")
    if vh3_el:
        print("Found VH3 element. InnerText:", vh3_el.inner_text().strip())
        print("Clicking VH3...")
        vh3_el.click()
        time.sleep(3)
        print("URL after clicking VH3:", page.url)
    else:
        print("VH3 element not found!")
