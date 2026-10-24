import sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = browser.contexts[0].pages[0]
    
    # Check frames
    print("Page URL:", page.url)
    print("Frames count:", len(page.frames))
    for i, frame in enumerate(page.frames):
        print(f"Frame {i}: name='{frame.name}' url='{frame.url}'")
        
    # Check if there are navigation links or tree menu
    nav_info = page.evaluate("""() => {
        const links = [];
        document.querySelectorAll("a, [id*='Tree'], [id*='Menu'], [id*='Nav']").forEach(el => {
            const text = (el.innerText || '').trim();
            const href = el.href || '';
            const id = el.id || '';
            if ((text || href) && links.length < 50) {
                links.push({ id, text: text.slice(0, 40), href: href.slice(0, 60) });
            }
        });
        return links;
    }""")
    print("Nav elements sample:", len(nav_info))
    for it in nav_info[:20]:
        print(it)
