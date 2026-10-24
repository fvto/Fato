import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

t0 = time.time()
with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    if "QMSMain.aspx" not in page.url:
        page.goto("https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx", wait_until="domcontentloaded", timeout=15000)
    print(f"Goto finished in {time.time() - t0:.2f}s")
    
    # Click root menu
    t1 = time.time()
    page.click("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0_")
    time.sleep(0.3)
    
    # Click VH (idx 0)
    page.click("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0i0_")
    time.sleep(1.5)
    print(f"Switch completed in {time.time() - t1:.2f}s")
    print(f"Total time: {time.time() - t0:.2f}s")
    print("Current factory:", page.inner_text("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0_").strip())
