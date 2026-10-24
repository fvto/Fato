import sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    page.goto("https://sfc.chingluh.com/vhmes/FlowChart.aspx")
    page.wait_for_load_state("networkidle")
    print("Page Title:", page.title())
    print("Page URL:", page.url)
    
    # Extract all links, iframes, menus, navigation tree items
    tree_data = page.evaluate("""() => {
        const links = [];
        document.querySelectorAll("a, area, [onclick]").forEach(el => {
            const href = el.href || el.getAttribute('href') || '';
            const text = (el.innerText || el.getAttribute('title') || el.getAttribute('alt') || '').trim();
            const oc = el.getAttribute('onclick') || '';
            if (href || oc || text) {
                links.push({ text: text.slice(0, 50), href: href.slice(0, 80), onclick: oc.slice(0, 80) });
            }
        });
        
        // Also check iframes
        const iframes = Array.from(document.querySelectorAll("iframe, frame")).map(f => ({
            name: f.name || f.id,
            src: f.src
        }));
        
        // Check image maps
        const areas = Array.from(document.querySelectorAll("area")).map(a => ({
            alt: a.alt,
            href: a.href,
            title: a.title
        }));
        
        return { linksCount: links.length, links: links.slice(0, 40), iframes, areas };
    }""")
    
    print("Found links:", tree_data['linksCount'])
    print("Iframes:", tree_data['iframes'])
    print("Areas:", tree_data['areas'])
    for l in tree_data['links'][:25]:
        print("Item:", l)
        
    page.screenshot(path="scratch/flowchart_preview.png")
    print("Screenshot saved to scratch/flowchart_preview.png")
