import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

FACTORIES = ["VH", "VH2", "VH3", "VH4", "JV", "JV2", "JVB"]

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    # Go to QMSMain where the factory switcher exists
    if "QMSMain.aspx" not in page.url:
        page.goto("https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx")
        page.wait_for_load_state("networkidle")
        
    def do_switch(fac_name):
        print(f"\n--- Thử đổi sang Nhà máy: {fac_name} ---")
        # 1. Click root menu to open dropdown
        page.click("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0_")
        time.sleep(0.5)
        
        # 2. Map name to index
        idx_map = {"VH": 0, "VH2": 1, "VH3": 2, "VH4": 3, "JV": 4, "JV2": 5, "JVB": 6}
        idx = idx_map.get(fac_name.upper())
        if idx is None:
            print(f"Không tìm thấy index cho {fac_name}")
            return False
            
        selector = f"#wucHeaderMainTop_TopPanel_menuRightTop_DXI0i{idx}_"
        el = page.query_selector(selector)
        if el:
            el.click()
            time.sleep(2)
            try:
                page.wait_for_load_state("load", timeout=5000)
            except Exception:
                pass
            current_fac = page.inner_text("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0_").strip()
            print(f"Đã đổi thành công! Nhà máy hiện tại trên Header: '{current_fac}'")
            return current_fac == fac_name
        else:
            print(f"Không tìm thấy selector: {selector}")
            return False

    # Test switching to VH2 then back to VH
    do_switch("VH2")
    time.sleep(1)
    do_switch("VH")
    print("\nHoàn tất kiểm tra chuyển Nhà máy!")
