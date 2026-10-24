import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = next(pg for pg in browser.contexts[0].pages if 'sfc.chingluh.com' in pg.url and '/vhmes/' in pg.url)
    
    res = page.evaluate("""() => {
        const dtF = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_dtDateFrom');
        const dtT = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_dtDateTo');
        if (dtF) dtF.SetText('2026-10-01');
        if (dtT) dtT.SetText('2026-10-24');
        return {
            from: dtF ? dtF.GetText() : null,
            to: dtT ? dtT.GetText() : null
        };
    }""")
    print("Dates set result:", res)
    browser.close()
