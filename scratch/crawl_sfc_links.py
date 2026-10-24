import sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    # Check all links and elements on the page
    html = page.content()
    with open("scratch/current_page_dump.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("Dumped HTML size:", len(html))
    
    # Extract any URLs, tree items, scripts mentioning pages
    info = page.evaluate("""() => {
        const out = {
            links: [],
            menuItems: [],
            scriptsWithAspx: []
        };
        document.querySelectorAll("a").forEach(a => {
            if (a.href) out.links.push({ text: a.innerText.trim(), href: a.href });
        });
        
        // Check scripts for .aspx references
        document.querySelectorAll("script").forEach(s => {
            const matches = s.innerText.match(/['"][^'"]*?\\.aspx[^'"]*?['"]/g);
            if (matches) {
                matches.forEach(m => {
                    const clean = m.replace(/['"]/g, '');
                    if (!out.scriptsWithAspx.includes(clean)) out.scriptsWithAspx.push(clean);
                });
            }
        });
        return out;
    }""")
    print("Links found:", len(info['links']))
    for l in info['links']:
        print("Link:", l)
    print("Aspx in scripts:", info['scriptsWithAspx'])
