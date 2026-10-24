import sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

test_urls = [
    "https://sfc.chingluh.com/vhmes/Default.aspx",
    "https://sfc.chingluh.com/vhmes/Main.aspx",
    "https://sfc.chingluh.com/vhmes/Index.aspx",
    "https://sfc.chingluh.com/vhmes/Home.aspx",
    "https://sfc.chingluh.com/vhmes/QA/Default.aspx",
    "https://sfc.chingluh.com/vhmes/QA/MQAACheck.aspx",
    "https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504"
]

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    for u in test_urls:
        try:
            resp = page.goto(u, timeout=7000)
            page.wait_for_timeout(1000)
            print(f"[{resp.status if resp else 'NoResp'}] {u} -> Title: '{page.title()}' | URL: {page.url}")
        except Exception as e:
            print(f"[ERR] {u} -> {e}")
            
    # Go back to HFPA
    page.goto("https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504")
