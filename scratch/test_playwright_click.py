import sys
import json
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

def test_execute():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        page = next(pg for pg in browser.contexts[0].pages if 'sfc.chingluh.com' in pg.url and '/vhmes/' in pg.url)
        print("Page:", page.title(), page.url)

        # 1. Test clicking search
        search_btn = page.query_selector("#rpnlMain_ContentPlaceHolder1_btnSearch")
        if search_btn:
            print("Clicking search button...")
            search_btn.click()
            time.sleep(1)
            print("Search button clicked successfully!")

        browser.close()

if __name__ == "__main__":
    test_execute()
