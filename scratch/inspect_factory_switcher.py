import sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    # Go to QMSMain to inspect the exact dropdown shown in user's screenshot
    if 'QMSMain.aspx' not in page.url:
        page.goto("https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx")
        page.wait_for_load_state("networkidle")
        
    print("Page URL:", page.url)
    
    # Find the factory dropdown element
    res = page.evaluate("""() => {
        const out = {
            dropdowns: [],
            factoryElements: []
        };
        
        // Search for elements containing VH, VH2, VH3, VH4, JV, JV2, JVB
        document.querySelectorAll("select, ul, li, div, a, span, button").forEach(el => {
            const txt = (el.innerText || '').trim();
            const id = el.id || '';
            const cls = el.className || '';
            // Check if it's the dropdown container or button
            if (txt.includes('VH') && txt.includes('VH2') && txt.includes('JV')) {
                out.dropdowns.push({
                    tag: el.tagName,
                    id,
                    cls,
                    html: el.outerHTML.slice(0, 500)
                });
            }
        });
        
        // Also look at all links or items that have VH, VH2, etc.
        document.querySelectorAll("a, li, [onclick]").forEach(el => {
            const t = (el.innerText || '').trim();
            const href = el.href || el.getAttribute('href') || '';
            const oc = el.getAttribute('onclick') || '';
            if (['VH', 'VH2', 'VH3', 'VH4', 'JV', 'JV2', 'JVB'].includes(t)) {
                out.factoryElements.push({
                    tag: el.tagName,
                    text: t,
                    id: el.id,
                    href,
                    onclick: oc,
                    parentHtml: el.parentElement ? el.parentElement.outerHTML.slice(0, 300) : ''
                });
            }
        });
        
        return out;
    }""")
    
    print("Found dropdown containers:", len(res['dropdowns']))
    for d in res['dropdowns']:
        print("Dropdown:", d['tag'], d['id'], d['cls'])
        print(d['html'][:300])
        print("-" * 50)
        
    print("\nFactory elements found:", len(res['factoryElements']))
    for fe in res['factoryElements']:
        print(fe)
