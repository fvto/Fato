import os
import sys
import re
import json
import time
import base64
import shutil
import subprocess
import threading
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright
import sfc_downloads
from dashboard_paths import mqaa_workspace, mqaa_master_file

sys.stdout.reconfigure(encoding='utf-8')

CDP_URL = "http://localhost:9222"
SFC_WAIT_MS = 15 * 60 * 1000
SFC_CONNECT_WAIT_MS = 5 * 60 * 1000

SFC_LOCK = threading.RLock()
DYNAMIC_ACTION_LOCK = threading.Lock()
CACHED_FACTORIES = []
MQAA_SCAN_PROCESS = None
MQAA_STATUS = {"running": False, "progress": "", "last_run": None}

LAST_STATUS_TIME = 0
LAST_STATUS_CACHE = None
CONNECTION_STATUS_CACHE = None
CONNECTION_STATUS_TIME = 0
CONNECTION_STATUS_LOCK = threading.Lock()

SEVEN_FACTORIES = ["VH4", "VH3", "VH2", "VH", "JVB", "JV2", "JV"]
SELECTED_FACTORIES = ["VH"]

def format_download_filename(filename, plant=None):
    """
    Đảm bảo mọi file xuất ra từ SFC đều có hậu tố mã nhà máy,
    Ví dụ: QAStation_2026-10-24_2026-10-24.xlsx -> QAStation_2026-10-24_2026-10-24-VH.xlsx
    """
    if not filename:
        return filename
    path = Path(filename)
    stem = path.stem
    suffix = path.suffix
    if not plant:
        plant = SELECTED_FACTORIES[0] if SELECTED_FACTORIES else 'VH'
    plant = str(plant).upper().strip()
    plant_suffix = f"-{plant}"
    if stem.endswith(plant_suffix):
        return f"{stem}{suffix}"
    return f"{stem}-{plant}{suffix}"

URL_MAPPING = {
    "home": "https://sfc.chingluh.com/vhmes/FlowChart.aspx",
    "portal": "https://sfc.chingluh.com/vhmes/FlowChart.aspx",
    "qms": "https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx?dept_code=&tree_id=10",
    "hfpa": "https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504",
    "mqaa": "https://sfc.chingluh.com/vhmes/QA/MQAACheck.aspx",
    "mqaa_report": "https://sfc.chingluh.com/vhmes/QA/MQAAReport.aspx",
    "aql": "https://sfc.chingluh.com/vhmes/QA/AQLTracking.aspx",
    "copq": "https://sfc.chingluh.com/vhmes/QA/QAReportTrackingCOPQ.aspx",
    "copq_data": "https://sfc.chingluh.com/vhmes/QA/QAReportCOPQ.aspx",
    "copq_report": "https://sfc.chingluh.com/vhmes/QA/QAReportTrackingCOPQ.aspx",
    "copq_record": "https://sfc.chingluh.com/vhmes/QA/QARecordCOPQ.aspx",
    "bond_test": "https://sfc.chingluh.com/vhmes/QA/BondTest.aspx",
    "qa_analysis": "https://sfc.chingluh.com/vhmes/QA/QAAnalysis.aspx",
    "qa": "https://sfc.chingluh.com/vhmes/QA/QA.aspx",
    "ftt": "https://sfc.chingluh.com/vhmes/QA/QA.aspx"
}


def create_hfpa_session(plant):
    from sfc_http import Session
    return Session('JV3' if plant == 'JVB' else plant)


def get_connection_status():
    """Read browser-level metadata; never wait for a busy page or screenshot."""
    global CONNECTION_STATUS_CACHE, CONNECTION_STATUS_TIME
    import shutil
    from pathlib import Path
    now = time.monotonic()
    if CONNECTION_STATUS_CACHE is not None and now - CONNECTION_STATUS_TIME < 2:
        return dict(CONNECTION_STATUS_CACHE, mqaa_status=MQAA_STATUS)
    if not CONNECTION_STATUS_LOCK.acquire(blocking=False):
        return dict(CONNECTION_STATUS_CACHE or {'connected': False}, mqaa_status=MQAA_STATUS)
    try:
        node = shutil.which('node')
        if not node:
            raise RuntimeError('Chưa tìm thấy Node.js để kết nối Chrome SFC.')
        result = subprocess.run([node, str(Path(__file__).with_name('sfc_session_cookies.js')), '--status'],
                                capture_output=True, text=True, encoding='utf-8', timeout=12, check=True,
                                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        status = json.loads(result.stdout)
        status['current_plant'] = sfc_downloads.factory_code(status.get('current_plant'))
        CONNECTION_STATUS_CACHE = dict(status, form={}, available_plants=SEVEN_FACTORIES, selected_plants=SELECTED_FACTORIES)
    except Exception:
        CONNECTION_STATUS_CACHE = dict(connected=False, message='Chưa kết nối được Chrome SFC. Hãy mở Dashboard và DOM.', selected_plants=SELECTED_FACTORIES)
    finally:
        CONNECTION_STATUS_TIME = time.monotonic()
        CONNECTION_STATUS_LOCK.release()
    return dict(CONNECTION_STATUS_CACHE, mqaa_status=MQAA_STATUS, selected_plants=SELECTED_FACTORIES)

def ensure_sfc_english(page_or_context):
    try:
        ctx = page_or_context if hasattr(page_or_context, 'cookies') else getattr(page_or_context, 'context', None)
        if not ctx:
            return
        cookies_to_add = []
        existing = {c['name']: c['value'] for c in ctx.cookies() if 'chingluh' in c.get('domain', '')}
        for pl in SEVEN_FACTORIES:
            k = f"MyLanguageName{'JV3' if pl == 'JVB' else pl}"
            if existing.get(k) != 'en':
                cookies_to_add.append({'name': k, 'value': 'en', 'domain': 'sfc.chingluh.com', 'path': '/'})
        if existing.get('MyLanguageName') != 'en':
            cookies_to_add.append({'name': 'MyLanguageName', 'value': 'en', 'domain': 'sfc.chingluh.com', 'path': '/'})
        if cookies_to_add:
            ctx.add_cookies(cookies_to_add)
    except Exception:
        pass


def get_sfc_page(browser):
    from urllib.parse import urlparse
    page = next((pg for ctx in browser.contexts for pg in ctx.pages
                 if urlparse(pg.url).hostname == 'sfc.chingluh.com' and '/vhmes/' in pg.url), None)
    if page is None:
        raise RuntimeError('Không tìm thấy tab SFC. Hãy mở Dashboard và DOM rồi đăng nhập SFC.')
    if 'login.aspx' in page.url.lower():
        raise RuntimeError('Phiên SFC chưa đăng nhập. Vui lòng đăng nhập trên cửa sổ Chrome SFC rồi thử lại.')
    page.set_default_timeout(SFC_WAIT_MS)
    page.set_default_navigation_timeout(SFC_WAIT_MS)
    ensure_sfc_english(page)
    return page


def take_b64_screenshot(page):
    try:
        img_bytes = page.screenshot(type="jpeg", quality=60, timeout=5000)
        return f"data:image/jpeg;base64,{base64.b64encode(img_bytes).decode('utf-8')}"
    except Exception:
        return ""

def get_live_status():
    # Polling must not queue behind a long SFC postback and delay the command.
    if not SFC_LOCK.acquire(blocking=False):
        return dict(LAST_STATUS_CACHE or {"connected": False}, processing=True)
    try:
        return _get_live_status_impl()
    finally:
        SFC_LOCK.release()


def _get_live_status_impl():
    """Lấy trạng thái thực tế của SFC có khóa an toàn đa luồng & cache nhẹ 1.5s"""
    global LAST_STATUS_TIME, LAST_STATUS_CACHE
    now = time.time()
    if LAST_STATUS_CACHE and (now - LAST_STATUS_TIME < 1.5):
        return LAST_STATUS_CACHE

    with SFC_LOCK:
        try:
            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(CDP_URL, timeout=10000)
                page = get_sfc_page(browser)

                if not page:
                    return {"connected": False, "message": "Không tìm thấy tab SFC Ching Luh trong Chrome"}

                current_url = page.url
                title = page.title()
                img_b64 = take_b64_screenshot(page)

                current_plant = "VH"
                try:
                    cookies = page.context.cookies()
                    for c in cookies:
                        if c.get('name') == 'MyFactoryName' and sfc_downloads.factory_code(c.get('value')) in SEVEN_FACTORIES:
                            current_plant = sfc_downloads.factory_code(c.get('value'))
                            break
                except Exception:
                    pass

                form_data = {}
                try:
                    form_data = page.evaluate("""() => {
                        const res = {};
                        try {
                            const dtF = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_dtDateFrom');
                            if (dtF) res.dateFrom = dtF.GetText();
                            const dtT = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_dtDateTo');
                            if (dtT) res.dateTo = dtT.GetText();
                            const cbo = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_cboDeptFrom');
                            if (cbo) res.selectedDept = cbo.GetText();
                            const nike = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_ckbNIKE');
                            if (nike) res.nike = nike.GetChecked();
                        } catch(e) {}
                        return res;
                    }""")
                except Exception:
                    pass

                browser.close()

                result = {
                    "connected": True,
                    "url": current_url,
                    "title": title,
                    "current_plant": current_plant,
                    "selected_plants": SELECTED_FACTORIES,
                    "available_plants": ["VH", "VH2", "VH3", "VH4", "JV", "JV2", "JVB"],
                    "form": form_data,
                    "screenshot": img_b64,
                    "mqaa_status": MQAA_STATUS
                }
                LAST_STATUS_TIME = now
                LAST_STATUS_CACHE = result
                return result
        except Exception as e:
            return {"connected": False, "error": str(e)}

SELECTED_FACTORIES = ["VH"]

def set_selected_factories(factory_list):
    global SELECTED_FACTORIES, LAST_STATUS_TIME
    valid = list(dict.fromkeys(f.upper().strip() for f in factory_list if f.upper().strip() in SEVEN_FACTORIES or f.upper().strip() == 'ALL'))
    if not valid:
        return {"success": False, "error": "Danh sách nhà máy không hợp lệ"}
    if 'ALL' in valid or len(valid) == len(SEVEN_FACTORIES):
        SELECTED_FACTORIES = SEVEN_FACTORIES.copy()
        plant_label = "ALL"
    else:
        SELECTED_FACTORIES = valid
        plant_label = ", ".join(valid)

    if len(SELECTED_FACTORIES) == 1:
        try:
            switch_factory(SELECTED_FACTORIES[0])
        except Exception:
            pass

    LAST_STATUS_TIME = 0
    return {
        "success": True,
        "selected_plants": SELECTED_FACTORIES,
        "is_all": len(SELECTED_FACTORIES) == len(SEVEN_FACTORIES),
        "current_plant": plant_label,
        "message": f"Selected {len(SELECTED_FACTORIES)} factories: {plant_label}"
    }


def switch_factory(factory_name):
    """
    Chuyển đổi giữa 7 Nhà máy chính của Tập đoàn Ching Luh:
    VH, VH2, VH3, VH4, JV, JV2, JVB hoặc ALL
    """
    global LAST_STATUS_TIME, SELECTED_FACTORIES
    idx_map = {"VH": 0, "VH2": 1, "VH3": 2, "VH4": 3, "JV": 4, "JV2": 5, "JVB": 6}
    target = sfc_downloads.factory_code(factory_name.upper().strip())
    if target == 'ALL':
        SELECTED_FACTORIES = SEVEN_FACTORIES.copy()
        return {
            "success": True,
            "current_plant": "ALL",
            "selected_plants": SELECTED_FACTORIES,
            "message": "Selected: All 7 Factories (ALL - Global Multi-Plant Mode)"
        }
    idx = idx_map.get(target)
    if idx is None:
        return {"success": False, "error": f"Nhà máy '{factory_name}' không hợp lệ. Chỉ gồm: ALL, VH, VH2, VH3, VH4, JV, JV2, JVB"}
    SELECTED_FACTORIES = [target]

    with SFC_LOCK:
        if browser_reserved():
            return dict(success=False, busy=True, error='Báo cáo đang tự tải dữ liệu SFC; vui lòng chờ, không đổi trang hoặc bộ lọc.')
        try:
            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(CDP_URL, timeout=SFC_CONNECT_WAIT_MS)
                page = get_sfc_page(browser)

                if not page:
                    return {"success": False, "error": "Chưa kết nối được Chrome"}

                actual_plant = next((c['value'] for c in page.context.cookies()
                                     if c['name'] == 'MyFactoryName'), None)
                actual_plant = sfc_downloads.factory_code(actual_plant)
                if actual_plant == target:
                    browser.close()
                    return {"success": True, "current_plant": target,
                            "message": f"Đang ở Nhà máy {target}; giữ nguyên trang hiện tại."}
                if 'QCHFPA.aspx' in page.url:
                    from sfc_hfpa import wait_idle
                    wait_idle(page)

                if not page.query_selector("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0_"):
                    page.goto("https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx", wait_until="domcontentloaded", timeout=SFC_WAIT_MS)
                    time.sleep(0.5)

                page.click("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0_")
                time.sleep(0.3)

                selector = f"#wucHeaderMainTop_TopPanel_menuRightTop_DXI0i{idx}_"
                el = page.query_selector(selector)
                if el:
                    el.click()
                    time.sleep(2)
                    try:
                        page.wait_for_load_state("load", timeout=SFC_WAIT_MS)
                    except Exception:
                        pass

                    curr_text = ""
                    header_el = page.query_selector("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0_")
                    if header_el:
                        curr_text = header_el.inner_text().strip()

                    img_b64 = take_b64_screenshot(page)
                    actual_plant = next((c['value'] for c in page.context.cookies()
                                         if c['name'] == 'MyFactoryName'), None)
                    actual_plant = sfc_downloads.factory_code(actual_plant)
                    if actual_plant != target:
                        browser.close()
                        return {"success": False, "error": f"Chuyển nhà máy chưa được xác nhận: {actual_plant}, yêu cầu {target}"}
                    browser.close()
                    LAST_STATUS_TIME = 0

                    return {
                        "success": True,
                        "current_plant": curr_text or target,
                        "message": f"Đã chuyển thành công sang Nhà máy: {curr_text or target}",
                        "screenshot": img_b64
                    }
                else:
                    browser.close()
                    return {"success": False, "error": f"Không tìm thấy nút chọn nhà máy {target}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

def get_factories():
    """Lấy danh sách 138 chuyền/xưởng trên SFC"""
    global CACHED_FACTORIES
    kb_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sfc_memory_knowledge_base.json")
    if os.path.exists(kb_path):
        try:
            with open(kb_path, "r", encoding="utf-8") as f:
                kb = json.load(f)
                facs = kb.get("factories_and_lines", [])
                if facs:
                    CACHED_FACTORIES = facs
                    return CACHED_FACTORIES
        except Exception:
            pass

    if CACHED_FACTORIES:
        return CACHED_FACTORIES
        
    with SFC_LOCK:
        try:
            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(CDP_URL, timeout=10000)
                page = get_sfc_page(browser)
                items = page.evaluate("""() => {
                    const list = [];
                    try {
                        const cbo = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_cboDeptFrom');
                        if (cbo) {
                            for (let i = 0; i < cbo.GetItemCount(); i++) {
                                const it = cbo.GetItem(i);
                                if (it && it.text) {
                                    list.push({ text: it.text.trim(), value: it.value });
                                }
                            }
                        }
                    } catch(e) {}
                    return list;
                }""")
                browser.close()
                if items:
                    CACHED_FACTORIES = items
                return CACHED_FACTORIES
        except Exception:
            return []

def navigate_to(destination):
    """Điều hướng đến các trang chức năng chính trên SFC và trả về ảnh chụp ngay"""
    global LAST_STATUS_TIME
    dest_key = destination.lower().strip()
    target_url = URL_MAPPING.get(dest_key, destination)
    
    if not target_url.startswith("http"):
        target_url = f"https://sfc.chingluh.com/vhmes/{target_url.lstrip('/')}"
        
    with SFC_LOCK:
        if browser_reserved():
            return dict(success=False, busy=True, error='Báo cáo đang tự tải dữ liệu SFC; vui lòng chờ, không đổi trang hoặc bộ lọc.')
        try:
            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(CDP_URL, timeout=SFC_CONNECT_WAIT_MS)
                page = get_sfc_page(browser)
                same_page = page.url == target_url or (dest_key == 'hfpa' and 'QCHFPA.aspx' in page.url)
                if not same_page:
                    if 'QCHFPA.aspx' in page.url:
                        from sfc_hfpa import wait_idle
                        wait_idle(page)
                    page.goto(target_url, wait_until="domcontentloaded", timeout=SFC_WAIT_MS)
                if 'login.aspx' in page.url.lower():
                    raise RuntimeError('SFC yêu cầu đăng nhập lại; chưa mở được phân hệ yêu cầu.')
                time.sleep(1.0)
                img_b64 = take_b64_screenshot(page)
                res = {
                    "success": True, 
                    "url": page.url, 
                    "title": page.title(),
                    "screenshot": img_b64
                }
                browser.close()
                LAST_STATUS_TIME = 0
                return res
        except Exception as e:
            return {"success": False, "error": str(e)}

def apply_sfc_filters(date_from=None, date_to=None, factory=None, action=None, nike=None,
                      export_path=None, expected_plant=None):
    import sfc_hfpa
    with SFC_LOCK:
        if browser_reserved():
            return dict(success=False, busy=True, error='Báo cáo đang tự tải dữ liệu SFC; vui lòng chờ.')
        return sfc_hfpa.run(sys.modules[__name__], date_from, date_to, factory, action, nike,
                            export_path, expected_plant)


INSPECT_JS = """() => {
    const coll = (typeof aspxGetControlCollection === 'function') ? aspxGetControlCollection() : null;
    const elements = coll ? coll.elements : {};

    const pageInfo = {
        title: document.title,
        url: location.href,
        pathname: location.pathname
    };

    const isGarbage = (name) => {
        const n = name.toLowerCase();
        return n.includes('_ddd') || n.includes('_fnp') || n.includes('dxfreditor') ||
               n.includes('dxhfp') || n.includes('dxcbtn') || n.includes('_cell') ||
               n.includes('dxpager') || n.includes('dxp-') || n.includes('pcupdatecost') ||
               n.includes('titlename') || (n.includes('grd') && !n.includes('btn') && !n.includes('cbo')) ||
               n.includes('txttmp') || n.includes('lbl_msg') || n.includes('popupmsg') ||
               n.includes('wucheadermaintop') || n.includes('menurighttop') ||
               n.endsWith('_i') || n.endsWith('_b') || n.endsWith('_t') || n.endsWith('_d');
    };

    const isVisible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = window.getComputedStyle(el);
        return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
    };

    const dateInputs = [];
    const dropdowns = [];
    const buttons = [];
    const checkboxes = [];

    if (coll) {
        for (let key in elements) {
            const ctrl = elements[key];
            if (!ctrl || !ctrl.name) continue;
            const name = ctrl.name;

            if (isGarbage(name)) continue;

            const domEl = document.getElementById(name) || document.getElementById(name + '_I');
            if (domEl && !isVisible(domEl)) continue;
            if (domEl && (domEl.closest('.dxpc-mainDiv') || domEl.closest('.dxgvTable') || domEl.closest('table[id*="grd"]'))) continue;

            // 1. Date Edits
            if ((name.toLowerCase().includes('date') || name.includes('dtDate')) && !name.toLowerCase().includes('btn')) {
                if (typeof ctrl.GetText === 'function') {
                    dateInputs.push({
                        id: name,
                        input_id: name + '_I',
                        name: name,
                        label: name.toLowerCase().includes('from') ? 'Date From' : (name.toLowerCase().includes('to') ? 'Date To' : 'Date'),
                        value: ctrl.GetText(),
                        type: 'date'
                    });
                }
            }

            // 2. Menu Dropdowns (DevExpress ASPxMenu with sub-items)
            else if (typeof ctrl.GetItemCount === 'function' && ctrl.GetItemCount() > 0) {
                const item0 = ctrl.GetItem(0);
                const subCount = (item0 && typeof item0.GetItemCount === 'function') ? item0.GetItemCount() : 0;
                if (subCount > 0 && !name.toLowerCase().includes('header') && !name.toLowerCase().includes('menurighttop')) {
                    const subOptions = [];
                    for (let j = 0; j < subCount; j++) {
                        const sub = item0.GetItem(j);
                        const subDomId = name + '_DXI0i' + j + '_';
                        const subText = (sub.GetText ? sub.GetText() : sub.text || '').trim();
                        subOptions.push({
                            index: j,
                            id: subDomId,
                            name: sub.name || ('SUB_' + j),
                            text: subText || ('Option ' + (j + 1)),
                            value: sub.name || subText
                        });
                    }
                    dropdowns.push({
                        id: name,
                        root_dom_id: name + '_DXI0_',
                        name: name,
                        type: 'menu_dropdown',
                        label: name.toLowerCase().includes('export') ? 'COPQ Report Export Options' : 'Dropdown Menu',
                        selected: subOptions[0] ? subOptions[0].text : '',
                        item_count: subOptions.length,
                        options: subOptions
                    });
                }
                else if (name.toLowerCase().includes('cbo') || name.toLowerCase().includes('dept') || name.toLowerCase().includes('combo') || name.toLowerCase().includes('type') || name.toLowerCase().includes('kind') || name.toLowerCase().includes('category') || name.toLowerCase().includes('reason')) {
                    const items = [];
                    const isChuyen = name.toLowerCase().includes('dept') || name.toLowerCase().includes('chuyen') || name.toLowerCase().includes('line');
                    
                    // Always include ALL as default for all dropdowns
                    const allText = isChuyen ? '-- All Lines / Workshops (ALL - Default) --' : '-- All (ALL - Default) --';
                    const supportsAll = !pageInfo.pathname.endsWith('/QAReportCOPQ.aspx');
                    if (supportsAll) items.push({ value: '', text: allText, is_all: true });
                    else {
                        items.push({ value: '', text: '-- Không chọn --', is_clear: true });
                        items.push({ value: 'ALL', text: 'ALL', is_all: true });
                    }

                    for (let i = 0; i < ctrl.GetItemCount(); i++) {
                        const it = ctrl.GetItem(i);
                        if (it && it.text) {
                            const txt = it.text.trim();
                            if (txt === '' || txt.toLowerCase() === 'all' || txt.toLowerCase().includes('tất cả')) continue;
                            items.push({ value: (it.value !== undefined) ? it.value : txt, text: txt });
                        }
                    }
                    if (items.length > 0) {
                        let label = isChuyen ? 'Select Line / Workshop (Default: ALL)' : 'Filter Criteria';
                        if (name.toLowerCase().includes('category')) label = 'Category (Default: ALL)';
                        else if (name.toLowerCase().includes('type')) label = 'Defect Type (Default: ALL)';
                        else if (name.toLowerCase().includes('reason')) label = 'Reason Code (Default: ALL)';
                        else if (name.toLowerCase().includes('kind')) label = 'Inspection Kind (Default: ALL)';

                        // Mặc định luôn là ALL
                        const defaultSelected = supportsAll ? allText : ctrl.GetText();
                        if (!supportsAll) label = name.toLowerCase().includes('category') ? 'Category' : 'Type';

                        dropdowns.push({
                            id: name,
                            name: name,
                            type: 'combobox',
                            label: label,
                            selected: defaultSelected,
                            selected_value: !supportsAll && (ctrl.GetValue() == null || ctrl.GetValue() === '') ? '' :
                                (supportsAll || items.some(it => it.value === ctrl.GetValue()) ? ctrl.GetValue() : null),
                            requires_selection: !supportsAll && ctrl.GetValue() != null && ctrl.GetValue() !== '' && !items.some(it => it.value === ctrl.GetValue()),
                            supports_clear: !supportsAll,
                            supports_all: supportsAll,
                            item_count: items.length,
                            options: items.slice(0, 200)
                        });
                    }
                }
            }

            // 3. Checkboxes
            else if (typeof ctrl.GetChecked === 'function' && (name.includes('ckb') || name.toLowerCase().includes('check'))) {
                const label = domEl ? (domEl.innerText || domEl.getAttribute('title') || name.split('_').pop()) : name;
                checkboxes.push({
                    id: name,
                    name: name,
                    label: label.trim().replace(/^\\s+/, ''),
                    checked: ctrl.GetChecked(),
                    type: 'checkbox'
                });
            }

            // 4. Buttons
            else if (name.toLowerCase().includes('btn') || typeof ctrl.DoClick === 'function') {
                let label = ctrl.GetText ? ctrl.GetText() : (domEl ? (domEl.innerText || domEl.getAttribute('title') || '') : '');
                if (!label) {
                    if (name.toLowerCase().includes('search')) label = 'Search';
                    else if (name.toLowerCase().includes('btnexcelall')) label = 'Export by Station (FTT Report)';
                    else if (name.toLowerCase().includes('export') || name.toLowerCase().includes('excel')) label = 'Export Excel';
                    else if (name.toLowerCase().includes('flowchart') || name.toLowerCase().includes('home')) label = 'FlowChart Home';
                    else if (name.toLowerCase().includes('print')) label = 'Print';
                    else if (name.toLowerCase().includes('update')) label = 'Update';
                    else if (name.toLowerCase().includes('clear')) label = 'Clear Filter';
                    else label = name.split('_').pop();
                }
                if (name.toLowerCase().includes('btnexcelall')) {
                    label = 'Export by Station (FTT Report)';
                } else if (name.toLowerCase() === 'rpnlmain_contentplaceholder1_btnexcel' && pageInfo.pathname.includes('QA.aspx')) {
                    label = 'Export Overview (Grid)';
                }
                if (!label || !label.trim()) continue;

                let icon = 'bolt';
                let btnType = 'primary';
                const lower = (name + ' ' + label).toLowerCase();
                if (lower.includes('search') || lower.includes('tim')) {
                    icon = 'magnifying-glass';
                    btnType = 'search';
                } else if (name.toLowerCase().includes('btnexcelall')) {
                    icon = 'folder-tree';
                    btnType = 'ftt_station_export';
                } else if (lower.includes('excel') || lower.includes('export') || lower.includes('xuat')) {
                    icon = 'file-excel';
                    btnType = 'export';
                } else if (lower.includes('flowchart') || lower.includes('home')) {
                    icon = 'house';
                    btnType = 'nav';
                }

                const isAlreadyMenu = dropdowns.some(d => d.id === name);
                if (!isAlreadyMenu) {
                    buttons.push({
                        id: name,
                        name: name,
                        label: label.trim(),
                        icon: icon,
                        btn_type: btnType
                    });
                }
            }
        }
    }

    // Extract Station list for FTT system in official English
    const stationEnMap = {
        'T01': 'Cutting',
        'T09': 'Stitching',
        'T11': 'Assembly',
        'T22': 'HFPA',
        'T23': 'Assembly IPQC',
        'T24': 'Stockfitting',
        'Chặt': 'Cutting',
        'May': 'Stitching',
        'Thành hình': 'Assembly',
        'Thành hình IPQC': 'Assembly IPQC',
        'Gia công đế': 'Stockfitting'
    };
    const fttStations = [];
    const isQAPage = pageInfo.pathname.includes('QA.aspx');
    const chkl = elements['rpnlMain_ContentPlaceHolder1_popupSetArea_chklStation'];
    if (chkl && typeof chkl.GetItemCount === 'function' && chkl.GetItemCount() > 0) {
        for (let i = 0; i < chkl.GetItemCount(); i++) {
            const it = chkl.GetItem(i);
            if (it) {
                const enName = stationEnMap[it.value] || stationEnMap[it.text] || it.text;
                fttStations.push({ index: i, text: enName, value: it.value });
            }
        }
    } else if (isQAPage) {
        fttStations.push(
            { index: 0, text: 'Cutting', value: 'T01' },
            { index: 1, text: 'Stitching', value: 'T09' },
            { index: 2, text: 'Assembly', value: 'T11' },
            { index: 3, text: 'HFPA', value: 'T22' },
            { index: 4, text: 'Assembly IPQC', value: 'T23' },
            { index: 5, text: 'Stockfitting', value: 'T24' }
        );
    }

    return {
        page: pageInfo,
        date_inputs: dateInputs,
        dropdowns: dropdowns,
        buttons: buttons,
        checkboxes: checkboxes,
        ftt_stations: fttStations
    };
}"""


def inspect_active_page():
    """Tự động quét và nhận diện các button, dropdown, ô chọn trên trang SFC hiện tại"""
    with SFC_LOCK:
        try:
            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(CDP_URL, timeout=10000)
                page = get_sfc_page(browser)
                if not page:
                    browser.close()
                    return {"success": False, "error": "Chưa kết nối được trang SFC."}

                data = page.evaluate(INSPECT_JS)
                img_b64 = take_b64_screenshot(page)
                current_plant = "VH"
                try:
                    for c in page.context.cookies():
                        if c.get('name') == 'MyFactoryName' and sfc_downloads.factory_code(c.get('value')) in SEVEN_FACTORIES:
                            current_plant = sfc_downloads.factory_code(c.get('value'))
                            break
                except Exception:
                    pass

                browser.close()
                return {
                    "success": True,
                    "current_plant": current_plant,
                    "available_plants": SEVEN_FACTORIES,
                    "inspection": data,
                    "screenshot": img_b64
                }
        except Exception as e:
            return {"success": False, "error": str(e)}


def browser_reserved():
    import sfc_report_tools
    return sfc_report_tools.browser_reserved(exclude_pid=os.getpid())


def is_copq_data_menu(action_type, target_id, params):
    params = params or {}
    return (action_type == 'select_menu' and 'btnexportexcel' in (target_id or '').lower() and
            (params.get('value') == 'EXCEL2' or params.get('index') == 1 or
             str(params.get('sub_id', '')).endswith('_DXI0i1_')))


def execute_dynamic_action(action_type, target_id=None, params=None):
    if not DYNAMIC_ACTION_LOCK.acquire(blocking=False):
        return dict(success=False, busy=True, error='Đang thực hiện một thao tác SFC. Không gửi lại lệnh xuất.')
    try:
        return _execute_dynamic_action_locked(action_type, target_id, params)
    finally:
        DYNAMIC_ACTION_LOCK.release()


def _execute_dynamic_action_locked(action_type, target_id=None, params=None):
    if is_copq_data_menu(action_type, target_id, params):
        action_type = 'open_copq_data'
    with SFC_LOCK:
        if browser_reserved():
            return dict(success=False, busy=True, error='Báo cáo đang tự tải dữ liệu SFC; vui lòng chờ, không đổi trang hoặc bộ lọc.')
        is_export = (action_type in ('select_menu', 'ftt_station_export', 'reinspection_export') or
                     (action_type == 'click_button' and any(k in (target_id or '').lower()
                      for k in ('excel', 'export', 'xuat', 'download'))))
        if is_export and (len(SELECTED_FACTORIES) > 1 or (params or {}).get('combine') is True):
            try:
                return sfc_downloads.export_selected(sys.modules[__name__], action_type, target_id,
                                                     params or {}, list(SELECTED_FACTORIES))
            except Exception as error:
                return dict(success=False, error=str(error))
        return _execute_dynamic_action(action_type, target_id, params)


def _execute_dynamic_action(action_type, target_id=None, params=None):
    """Thực thi thao tác nhận diện động trên SFC (chọn dropdown menu, click button, đổi ngày...)"""
    global LAST_STATUS_TIME
    params = params or {}
    if is_copq_data_menu(action_type, target_id, params):
        action_type = 'open_copq_data'
    with SFC_LOCK:
        try:
            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(CDP_URL, timeout=SFC_CONNECT_WAIT_MS)
                page = get_sfc_page(browser)
                if params.get('_page_url') and page.url != params['_page_url']:
                    page.goto(params['_page_url'], wait_until='domcontentloaded', timeout=SFC_WAIT_MS)
                from sfc_mqaa_exports import recover_timeout_page
                mqaa_timeout_recovered = recover_timeout_page(page,action_type,target_id,params,SFC_WAIT_MS)
                sfc_downloads.wait_idle(page, SFC_WAIT_MS)
                current_plant = next((c['value'] for c in page.context.cookies()
                                      if c['name'] == 'MyFactoryName'), None)
                current_plant = sfc_downloads.factory_code(current_plant)
                if current_plant not in SEVEN_FACTORIES:
                    raise ValueError('Không xác định được nhà máy của file xuất.')
                if params.get('_expected_plant') and current_plant != params['_expected_plant']:
                    raise ValueError('Nhà máy SFC không khớp với nhà máy yêu cầu.')
                downloaded_file = None
                if params.get('_filters'):
                    sfc_downloads.apply_filters(page, params['_filters'], SFC_WAIT_MS)
                mqaa_check = page.url.split('?', 1)[0].lower().endswith('/mqaacheck.aspx')
                if action_type in ('select_menu', 'click_button', 'open_copq_data', 'ftt_station_export') and (params.get('date_from') or params.get('date_to')):
                    dates = sfc_downloads.capture_filters(page)
                    for cid, item in dates.items():
                        if item['kind'] != 'date':
                            continue
                        if mqaa_check and cid == 'dtDate' and target_id in ('btnSearch', 'btnExcelMQAA'):
                            first, last = params.get('date_from'), params.get('date_to')
                            if first and last and first != last:
                                raise ValueError('MQAA Check xử lý từng ngày. Mở MQAA Report để dùng cả khoảng đã chọn.')
                            item['value'] = first or last
                        elif any(marker in cid.lower() for marker in ('datefrom','dateto','fromdate','todate')):
                            value = params.get('date_from') if 'from' in cid.lower() else params.get('date_to')
                            if value:
                                item['value'] = value
                    sfc_downloads.apply_filters(page, dates, SFC_WAIT_MS)

                # 1. Đặt ngày tháng
                if action_type == 'open_copq_data':
                    dates = {k: v for k, v in sfc_downloads.capture_filters(page).items() if v['kind'] == 'date'}
                    if page.url != URL_MAPPING['copq_data']:
                        page.goto(URL_MAPPING['copq_data'], wait_until='domcontentloaded', timeout=SFC_WAIT_MS)
                    sfc_downloads.wait_idle(page, SFC_WAIT_MS)
                    sfc_downloads.apply_filters(page, dates, SFC_WAIT_MS)
                elif action_type == 'clear_copq_filters':
                    if 'QAReportCOPQ.aspx' not in page.url:
                        raise ValueError('Hãy mở COPQ Data trước khi xóa lựa chọn.')
                    for suffix in ('cboCategory', 'cboType'):
                        sfc_downloads.set_copq_combobox(page, 'rpnlMain_ContentPlaceHolder1_' + suffix, None, SFC_WAIT_MS)
                elif action_type == 'restore_filters':
                    pass
                elif action_type == 'reinspection_export':
                    downloaded_file = sfc_downloads.export_reinspection_data(
                        page,SFC_WAIT_MS,current_plant,format_download_filename,params)
                elif action_type == 'set_dates':
                    date_from = params.get('date_from')
                    date_to = params.get('date_to')
                    page.evaluate("""({f, t, single}) => {
                        const coll = (typeof aspxGetControlCollection === 'function') ? aspxGetControlCollection() : null;
                        if (coll) {
                            for (let k in coll.elements) {
                                const c = coll.elements[k];
                                if (!c || !c.name) continue;
                                const n = c.name.toLowerCase();
                                let value;
                                if (single && n === 'dtdate' && f === t) value = f;
                                else if (n.includes('date') && n.includes('from')) value = f;
                                else if (n.includes('date') && n.includes('to')) value = t;
                                if (!value) continue;
                                if (typeof c.SetDate === 'function') {
                                    const [year, month, day] = value.split('-').map(Number);
                                    c.SetDate(new Date(year, month - 1, day));
                                } else if (typeof c.SetText === 'function') c.SetText(value);
                            }
                        }
                    }""", {"f": date_from, "t": date_to, "single": mqaa_check})

                # 2. Chọn tùy chọn trong Menu xổ xuống (như Menu Xuất Excel trên COPQ)
                elif action_type == 'select_menu':
                    root_id = target_id
                    sub_id = params.get('sub_id')
                    index = params.get('index', 0)
                    root_selector = f"#{root_id}_DXI0_"
                    sub_selector = f"#{sub_id}" if sub_id else f"#{root_id}_DXI0i{index}_"
                    source_url = page.url
                    received_downloads = []
                    on_download = lambda download: received_downloads.append(download)
                    page.on('download', on_download)
                    try:
                        if page.query_selector(root_selector):
                            # Navigation can leave the pointer at the same screen
                            # position. A fresh mouseenter must open this menu.
                            page.mouse.move(0, 0)
                            page.hover(root_selector)
                            page.locator(sub_selector).wait_for(state='visible', timeout=15000)
                        item = page.query_selector(sub_selector)
                        if item is None:
                            raise ValueError('Không tìm thấy tùy chọn xuất: ' + sub_selector)
                        content = page.query_selector(sub_selector + '_T')
                        with page.expect_response(lambda r: r.request.method == 'POST' and r.url == source_url,
                                                  timeout=SFC_WAIT_MS) as menu_response:
                            (content or item).click(no_wait_after=True)
                        sfc_downloads.require_attachment(menu_response.value)
                        dl = received_downloads[0] if received_downloads else page.wait_for_event('download', timeout=SFC_WAIT_MS)
                        downloaded_file = sfc_downloads.save_download(dl, current_plant, format_download_filename)
                    finally:
                        page.remove_listener('download', on_download)

                # 3. Bấm nút hành động (Tìm kiếm, Xuất...)
                elif action_type == 'click_button':
                    selections = params.get('selections', {})
                    if 'QAReportCOPQ.aspx' in page.url:
                        for suffix in ('cboCategory', 'cboType'):
                            for cid, item in selections.items():
                                if cid.endswith(suffix):
                                    sfc_downloads.set_copq_combobox(page, cid, item.get('value'), SFC_WAIT_MS)
                    page.evaluate("""({selections, bid}) => {
                        const coll = (typeof aspxGetControlCollection === 'function') ? aspxGetControlCollection() : null;
                        if (coll && selections) {
                            for (let cid in selections) {
                                const cbo = coll.Get(cid);
                                const sel = selections[cid];
                                if (cbo && sel) {
                                    const v = sel.value;
                                    const t = sel.text || '';
                                    const isAll = (!v || v === 'ALL' || t.includes('Tất cả') || t.includes('ALL'));
                                    if (isAll) {
                                        if (typeof cbo.SetValue === 'function') cbo.SetValue(null);
                                        if (typeof cbo.SetText === 'function') cbo.SetText('');
                                        if (typeof cbo.SetSelectedIndex === 'function') cbo.SetSelectedIndex(-1);
                                        const inp = document.getElementById(cid + '_I');
                                        if (inp) inp.value = '';
                                    } else if (typeof cbo.FindItemByText === 'function') {
                                        const it = cbo.FindItemByText(t) || cbo.FindItemByValue(v);
                                        if (it) cbo.SetSelectedItem(it);
                                        else if (typeof cbo.SetText === 'function') cbo.SetText(t || v);
                                    }
                                }
                            }
                        }
                    }""", {"selections": selections, "bid": target_id})

                    is_export_click = any(k in (target_id or '').lower() for k in ['excel', 'export', 'xuat', 'download'])
                    if is_export_click and 'QAReportCOPQ.aspx' in page.url:
                        downloaded_file = sfc_downloads.export_copq_data(page, SFC_WAIT_MS, current_plant, format_download_filename)
                    elif is_export_click and 'mqaaReport.aspx'.lower() in page.url.lower():
                        from sfc_mqaa_exports import export_mqaa_data
                        downloaded_file = export_mqaa_data(page,target_id,SFC_WAIT_MS,current_plant,format_download_filename,
                            confirmed_timeout=mqaa_timeout_recovered or params.get('_mqaa_timeout_plant')==current_plant)
                    elif is_export_click:
                        try:
                            with page.expect_download(timeout=SFC_WAIT_MS) as dl_info:
                                btn_el = page.query_selector(f"#{target_id}")
                                if btn_el:
                                    btn_el.click()
                                else:
                                    page.evaluate("""(bid) => {
                                        const coll = (typeof aspxGetControlCollection === 'function') ? aspxGetControlCollection() : null;
                                        const ctrl = coll ? coll.Get(bid) : null;
                                        if (ctrl && typeof ctrl.DoClick === 'function') ctrl.DoClick();
                                        else {
                                            const el = document.getElementById(bid);
                                            if (el) el.click();
                                        }
                                    }""", target_id)
                            dl = dl_info.value
                            downloaded_file = sfc_downloads.save_download(dl, current_plant, format_download_filename)
                        except Exception as dl_err:
                            raise RuntimeError('Chưa lưu được file Excel: ' + str(dl_err)) from dl_err
                    elif 'QAReportCOPQ.aspx' in page.url and (target_id or '').endswith('btnSearch'):
                        sfc_downloads.copq_search(page, SFC_WAIT_MS)
                    else:
                        btn_el = page.query_selector(f"#{target_id}")
                        if btn_el:
                            btn_el.click()
                        else:
                            page.evaluate("""(bid) => {
                                const coll = (typeof aspxGetControlCollection === 'function') ? aspxGetControlCollection() : null;
                                const ctrl = coll ? coll.Get(bid) : null;
                                if (ctrl && typeof ctrl.DoClick === 'function') ctrl.DoClick();
                                else {
                                    const el = document.getElementById(bid);
                                    if (el) el.click();
                                }
                            }""", target_id)
                        time.sleep(1.5)

                # 4. Chọn ComboBox
                elif action_type == 'select_combobox':
                    if 'QAReportCOPQ.aspx' in page.url and (target_id or '').endswith(('cboCategory', 'cboType')):
                        sfc_downloads.set_copq_combobox(page, target_id, params.get('value'), SFC_WAIT_MS)
                    val = params.get('value', '')
                    txt = params.get('text', '')
                    page.evaluate("""({cid, v, t}) => {
                        const coll = (typeof aspxGetControlCollection === 'function') ? aspxGetControlCollection() : null;
                        const cbo = coll ? coll.Get(cid) : null;
                        const isAll = (!v || v === 'ALL' || (t && (t.includes('Tất cả') || t.includes('ALL'))));
                        if (cbo) {
                            if (isAll) {
                                if (typeof cbo.SetValue === 'function') cbo.SetValue(null);
                                if (typeof cbo.SetText === 'function') cbo.SetText('');
                                if (typeof cbo.SetSelectedIndex === 'function') cbo.SetSelectedIndex(-1);
                            } else if (typeof cbo.FindItemByText === 'function') {
                                const it = cbo.FindItemByText(t) || cbo.FindItemByValue(v);
                                if (it) cbo.SetSelectedItem(it);
                                else if (typeof cbo.SetText === 'function') cbo.SetText(t || v);
                            } else if (typeof cbo.SetValue === 'function') {
                                cbo.SetValue(v);
                            }
                        } else {
                            const sel = document.getElementById(cid);
                            if (sel) {
                                sel.value = v;
                                sel.dispatchEvent(new Event('change'));
                            }
                        }
                        const inp = document.getElementById(cid + '_I');
                        if (inp && isAll) inp.value = '';
                    }""", {"cid": target_id, "v": val, "t": txt})
                    time.sleep(1.0)

                # 5. Đổi trạng thái Checkbox
                elif action_type == 'toggle_checkbox':
                    chk = params.get('checked')
                    page.evaluate("""({cid, c}) => {
                        const coll = (typeof aspxGetControlCollection === 'function') ? aspxGetControlCollection() : null;
                        const ckb = coll ? coll.Get(cid) : null;
                        if (ckb && typeof ckb.SetChecked === 'function') ckb.SetChecked(c);
                        else {
                            const el = document.getElementById(cid);
                            if (el) {
                                el.checked = c;
                                el.dispatchEvent(new Event('change'));
                            }
                        }
                    }""", {"cid": target_id, "c": chk})

                # 6. Xuất Báo Cáo FTT theo Station (Select Station Dialog)
                elif action_type == 'ftt_station_export':
                    stations = params.get('stations', [])
                    report_type = params.get('report_type', 'station')
                    file_name = params.get('file_name', '')
                    date_from = params.get('date_from')
                    date_to = params.get('date_to')

                    # Đặt ngày nếu có
                    if date_from or date_to:
                        page.evaluate("""({f, t}) => {
                            const coll = (typeof aspxGetControlCollection === 'function') ? aspxGetControlCollection() : null;
                            if (coll) {
                                for (let k in coll.elements) {
                                    const c = coll.elements[k];
                                    if (!c || !c.name) continue;
                                    const n = c.name.toLowerCase();
                                    if (f && (n.includes('datefrom') || (n.includes('date') && n.includes('from')))) {
                                        if (typeof c.SetText === 'function') c.SetText(f);
                                    }
                                    if (t && (n.includes('dateto') || (n.includes('date') && n.includes('to')))) {
                                        if (typeof c.SetText === 'function') c.SetText(t);
                                    }
                                }
                            }
                        }""", {"f": date_from, "t": date_to})

                    # Mở popup Select Station nếu chưa mở
                    page.evaluate("""() => {
                        const coll = aspxGetControlCollection();
                        const popup = coll.Get('rpnlMain_ContentPlaceHolder1_popupSetArea');
                        if (!popup || !popup.IsVisible()) {
                            const btnALL = coll.Get('rpnlMain_ContentPlaceHolder1_btnExcelALL');
                            if (btnALL) btnALL.DoClick();
                            else {
                                const el = document.getElementById('rpnlMain_ContentPlaceHolder1_btnExcelALL');
                                if (el) el.click();
                            }
                        }
                    }""")
                    time.sleep(1.0)

                    # Chọn Stations trong CheckboxList
                    page.evaluate("""({stations, selection}) => {
                        const coll = aspxGetControlCollection();
                        const chkl = coll.Get('rpnlMain_ContentPlaceHolder1_popupSetArea_chklStation');
                        if (chkl) {
                            chkl.UnselectAll();
                            if (selection === 'none') {
                                if (chkl.GetSelectedIndices().length !== 0) throw new Error('SFC chưa bỏ chọn hết trạm.');
                            } else if (!stations || stations === 'ALL' || stations.length === 0 || (Array.isArray(stations) && stations.includes('ALL'))) {
                                chkl.SelectAll();
                            } else {
                                for (let i = 0; i < chkl.GetItemCount(); i++) {
                                    const it = chkl.GetItem(i);
                                    if (it && (stations.includes(it.value) || stations.includes(it.text))) {
                                        chkl.SelectIndices([i]);
                                    }
                                }
                                const selected = chkl.GetSelectedIndices();
                                if (selected.length !== stations.length || selected.some(i => {
                                    const it = chkl.GetItem(i);
                                    return !it || !(stations.includes(it.value) || stations.includes(it.text));
                                })) throw new Error('SFC chưa chọn đúng các trạm yêu cầu; dừng xuất.');
                            }
                        } else throw new Error('Không tìm thấy danh sách trạm SFC; dừng xuất.');
                    }""", {"stations": stations, "selection": params.get('station_selection')})
                    time.sleep(0.4)

                    if file_name:
                        page.evaluate("""({fn}) => {
                            const coll = aspxGetControlCollection();
                            const txt = coll.Get('rpnlMain_ContentPlaceHolder1_popupSetArea_ASPxRoundPanel11_txtFileName');
                            if (txt && typeof txt.SetText === 'function') txt.SetText(fn);
                        }""", {"fn": file_name})

                    target_btn_id = 'rpnlMain_ContentPlaceHolder1_popupSetArea_btnSave'
                    if report_type == 'nike':
                        target_btn_id = 'rpnlMain_ContentPlaceHolder1_popupSetArea_ASPxRoundPanel11_btnExcelNIKE'
                    elif report_type == 'internal':
                        target_btn_id = 'rpnlMain_ContentPlaceHolder1_popupSetArea_ASPxRoundPanel11_btnExcelInternal'

                    downloaded_file = None
                    try:
                        with page.expect_download(timeout=SFC_WAIT_MS) as dl_info:
                            page.evaluate("""(bid) => {
                                const coll = aspxGetControlCollection();
                                const btn = coll.Get(bid);
                                if (btn && typeof btn.DoClick === 'function') btn.DoClick();
                                else {
                                    const el = document.getElementById(bid);
                                    if (el) el.click();
                                }
                            }""", target_btn_id)
                        dl = dl_info.value
                        downloaded_file = sfc_downloads.save_download(dl, current_plant, format_download_filename)
                    except Exception as dl_err:
                        raise RuntimeError('Chưa lưu được file Excel FTT: ' + str(dl_err)) from dl_err

                try:
                    page.wait_for_load_state("load", timeout=5000)
                except Exception:
                    pass
                time.sleep(0.5)

                img_b64 = take_b64_screenshot(page)
                inspection = page.evaluate(INSPECT_JS)
                browser.close()
                LAST_STATUS_TIME = 0

                downloaded_file_info = locals().get('downloaded_file')
                msg = (downloaded_file_info.get('message') or f"Exported successfully: {downloaded_file_info['name']}") if downloaded_file_info else "Executed action on SFC successfully."
                return {
                    "success": True,
                    "action": action_type,
                    "requires_filter_selection": action_type == 'open_copq_data',
                    "url": inspection['page']['url'],
                    "target_id": target_id,
                    "screenshot": img_b64,
                    "inspection": inspection,
                    "downloaded_file": downloaded_file_info,
                    "downloaded_files": [downloaded_file_info] if downloaded_file_info else [],
                    "message": msg
                }
        except Exception as e:
            return {"success": False, "error": str(e), "downloaded_files": getattr(e,'downloaded_files',[])}


def parse_intent(text):
    from sfc_commands import parse_intent as parse
    return parse(text)


def execute_command_stream(user_prompt):
    import sfc_commands
    yield from sfc_commands.execute(sys.modules[__name__], user_prompt)


def parse_and_execute_command(user_prompt):
    """Wrapper đồng bộ cho các lệnh API thông thường"""
    last_res = None
    for step in execute_command_stream(user_prompt):
        last_res = step
    return {
        "success": bool(last_res and last_res.get('status') == 'done'),
        "command": user_prompt,
        "details": last_res.get("msg") if last_res else "Đã hoàn thành",
        "result": last_res
    }

def start_mqaa_scan_background():
    """Kích hoạt quét 21 Form MQAA ngầm"""
    global MQAA_SCAN_PROCESS, MQAA_STATUS
    if MQAA_STATUS["running"]:
        return {"status": "already_running", "message": "Quá trình quét 21 form đang chạy dở!"}

    workspace = str(mqaa_workspace())
    script_path = os.path.join(workspace, "auto_scan_all_21_models.py")
    if not os.path.isfile(script_path):
        return {"status": "error", "message": "Không tìm thấy chương trình MQAA cũ tại " + script_path}

    def run_worker():
        global MQAA_STATUS
        MQAA_STATUS["running"] = True
        MQAA_STATUS["progress"] = "Đang bắt đầu kết nối và quét 21 Form..."
        MQAA_STATUS["last_run"] = datetime.now().strftime("%H:%M:%S %d/%m/%Y")
        try:
            p = subprocess.Popen(
                [sys.executable, "-u", script_path],
                cwd=workspace,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8'
            )
            for line in p.stdout:
                line_str = line.strip()
                if line_str:
                    MQAA_STATUS["progress"] = line_str
            p.wait()
            MQAA_STATUS["running"] = False
            MQAA_STATUS["progress"] = "🎉 Đã hoàn thành quét toàn bộ 21 Form và xuất file Excel Master!"
        except Exception as e:
            MQAA_STATUS["running"] = False
            MQAA_STATUS["progress"] = f"Lỗi: {e}"

    t = threading.Thread(target=run_worker, daemon=True)
    t.start()
    return {"status": "started", "message": "Đã kích hoạt tiến trình quét 21 Form ngầm thành công."}

def open_master_excel():
    target = mqaa_master_file()
    if target:
        os.startfile(str(target))
        return {"success": True, "file": target.name}
    return {"success": False, "error": "Chưa tìm thấy file Master Excel"}
