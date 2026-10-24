import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    print("Testing interaction on QCHFPA.aspx...")
    # Set Date From to 2026-10-20
    res = page.evaluate("""() => {
        const out = {};
        try {
            const dtFrom = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_dtDateFrom');
            if (dtFrom) {
                dtFrom.SetText('2026-10-20');
                out.dtFrom = dtFrom.GetText();
            }
            
            const dtTo = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_dtDateTo');
            if (dtTo) {
                dtTo.SetText('2026-10-24');
                out.dtTo = dtTo.GetText();
            }
            
            const cboDept = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_cboDeptFrom');
            if (cboDept) {
                // Find VL - D3
                for (let i = 0; i < cboDept.GetItemCount(); i++) {
                    const it = cboDept.GetItem(i);
                    if (it.text.includes('VL - D3')) {
                        cboDept.SetSelectedIndex(i);
                        out.selectedDept = it.text;
                        break;
                    }
                }
            }
        } catch(e) {
            out.error = e.toString();
        }
        return out;
    }""")
    print("Result:", res)
    time.sleep(1)
    
    # Test clicking search
    search_res = page.evaluate("""() => {
        try {
            const btn = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_btnSearch');
            if (btn) {
                btn.DoClick();
                return { success: true };
            }
            return { success: false, reason: 'Button not found' };
        } catch(e) {
            return { error: e.toString() };
        }
    }""")
    print("Search trigger result:", search_res)
    time.sleep(1)
    try:
        page.wait_for_load_state("load", timeout=10000)
    except Exception:
        pass
    time.sleep(2)

    
    # Check grid row count
    grid_info = page.evaluate("""() => {
        try {
            const grid = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_grdExcel');
            if (grid) {
                return {
                    visibleRowsOnPage: grid.GetVisibleRowsOnPage(),
                    pageCount: grid.GetPageCount ? grid.GetPageCount() : 1,
                    pageIndex: grid.GetPageIndex ? grid.GetPageIndex() : 0
                };
            }
            return { error: 'Grid not found' };
        } catch(e) {
            return { error: e.toString() };
        }
    }""")
    print("Grid info after search:", grid_info)
    page.screenshot(path="scratch/sfc_test_preview.png")
    print("Saved preview to scratch/sfc_test_preview.png")

