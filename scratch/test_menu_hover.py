import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = next(pg for pg in browser.contexts[0].pages if 'sfc.chingluh.com' in pg.url and '/vhmes/' in pg.url)
    
    # Try clicking the menu root then sub-item
    root_menu = page.query_selector("#rpnlMain_ContentPlaceHolder1_btnExportExcel_DXI0_")
    if root_menu:
        print("Root menu found, hovering / clicking...")
        root_menu.hover()
        time.sleep(0.3)
        sub_item = page.query_selector("#rpnlMain_ContentPlaceHolder1_btnExportExcel_DXI0i0_")
        print("Sub-item 0 visible after hover:", sub_item.is_visible() if sub_item else False)
    browser.close()
