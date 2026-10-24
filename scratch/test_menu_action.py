import sys
import json
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

def test_menu_action():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        page = next(pg for pg in browser.contexts[0].pages if 'sfc.chingluh.com' in pg.url and '/vhmes/' in pg.url)
        print("Page:", page.title())

        # Hover root
        root = page.query_selector("#rpnlMain_ContentPlaceHolder1_btnExportExcel_DXI0_")
        if root:
            root.hover()
            time.sleep(0.3)
            sub = page.query_selector("#rpnlMain_ContentPlaceHolder1_btnExportExcel_DXI0i0_")
            if sub:
                print("Clicking 1.Report COPQ...")
                sub.click()
                time.sleep(1.5)
                print("Clicked successfully!")

        browser.close()

if __name__ == "__main__":
    test_menu_action()
