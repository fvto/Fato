import sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = next(pg for pg in browser.contexts[0].pages if 'sfc.chingluh.com' in pg.url)
    print("Found page:", page.title(), page.url)

    # Check search button
    search_btn = page.query_selector("#rpnlMain_ContentPlaceHolder1_btnSearch")
    print("Search button exists:", bool(search_btn))

    # Check export menu
    export_menu = page.query_selector("#rpnlMain_ContentPlaceHolder1_btnExportExcel_DXI0_")
    print("Export menu button exists:", bool(export_menu))

    sub_items = page.query_selector_all("[id^='rpnlMain_ContentPlaceHolder1_btnExportExcel_DXI0i']")
    print("Sub items found:", len(sub_items))
    for s in sub_items:
        print(" - Sub item id:", s.get_attribute("id"), "text:", s.inner_text().strip())

    browser.close()
