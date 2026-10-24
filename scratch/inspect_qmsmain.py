import sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    page.goto("https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx?dept_code=&tree_id=10")
    page.wait_for_load_state("networkidle")
    print("Page Title:", page.title())
    print("Page URL:", page.url)
    
    # Extract left navigation tree or submenu
    menu_items = page.evaluate("""() => {
        const items = [];
        // Look for DevExpress TreeView or standard tree
        document.querySelectorAll("a, span, [id*='Tree'], td[onclick]").forEach(el => {
            const text = el.innerText ? el.innerText.trim() : '';
            const href = el.href || el.getAttribute('href') || '';
            const id = el.id || '';
            if (text && (href || el.getAttribute('onclick')) && items.length < 150) {
                items.push({ id, text: text.replace(/\\n/g, ' '), href });
            }
        });
        return items;
    }""")
    print("Total QMS menu items:", len(menu_items))
    for it in menu_items[:30]:
        print(it)
        
    page.screenshot(path="scratch/qmsmain_preview.png")
    print("Saved preview to scratch/qmsmain_preview.png")
