import os
import sys
import re
import json
import time
import base64
import subprocess
import threading
from datetime import datetime
from playwright.sync_api import sync_playwright
from dashboard_paths import mqaa_workspace, mqaa_master_file

sys.stdout.reconfigure(encoding='utf-8')

CDP_URL = "http://localhost:9222"
SFC_WAIT_MS = 15 * 60 * 1000
SFC_CONNECT_WAIT_MS = 5 * 60 * 1000

SFC_LOCK = threading.RLock()
CACHED_FACTORIES = []
MQAA_SCAN_PROCESS = None
MQAA_STATUS = {"running": False, "progress": "", "last_run": None}

LAST_STATUS_TIME = 0
LAST_STATUS_CACHE = None
CONNECTION_STATUS_CACHE = None
CONNECTION_STATUS_TIME = 0
CONNECTION_STATUS_LOCK = threading.Lock()

SEVEN_FACTORIES = ["VH4", "VH3", "VH2", "VH", "JVB", "JV2", "JV"]

URL_MAPPING = {
    "home": "https://sfc.chingluh.com/vhmes/FlowChart.aspx",
    "portal": "https://sfc.chingluh.com/vhmes/FlowChart.aspx",
    "qms": "https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx?dept_code=&tree_id=10",
    "hfpa": "https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504",
    "mqaa": "https://sfc.chingluh.com/vhmes/QA/MQAACheck.aspx",
    "mqaa_report": "https://sfc.chingluh.com/vhmes/QA/MQAAReport.aspx",
    "aql": "https://sfc.chingluh.com/vhmes/QA/AQLTracking.aspx",
    "copq": "https://sfc.chingluh.com/vhmes/QA/QARecordCOPQ.aspx",
    "bond_test": "https://sfc.chingluh.com/vhmes/QA/BondTest.aspx",
    "qa_analysis": "https://sfc.chingluh.com/vhmes/QA/QAAnalysis.aspx"
}


def create_hfpa_session(plant):
    from sfc_http import Session
    return Session(plant)


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
        CONNECTION_STATUS_CACHE = dict(status, form={}, available_plants=SEVEN_FACTORIES)
    except Exception:
        CONNECTION_STATUS_CACHE = dict(connected=False, message='Chưa kết nối được Chrome SFC. Hãy mở Dashboard và DOM.')
    finally:
        CONNECTION_STATUS_TIME = time.monotonic()
        CONNECTION_STATUS_LOCK.release()
    return dict(CONNECTION_STATUS_CACHE, mqaa_status=MQAA_STATUS)

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
                        if c.get('name') == 'MyFactoryName' and c.get('value') in SEVEN_FACTORIES:
                            current_plant = c.get('value')
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

def switch_factory(factory_name):
    """
    Chuyển đổi giữa 7 Nhà máy chính của Tập đoàn Ching Luh:
    VH, VH2, VH3, VH4, JV, JV2, JVB
    """
    global LAST_STATUS_TIME
    idx_map = {"VH": 0, "VH2": 1, "VH3": 2, "VH4": 3, "JV": 4, "JV2": 5, "JVB": 6}
    target = factory_name.upper().strip()
    idx = idx_map.get(target)
    if idx is None:
        return {"success": False, "error": f"Nhà máy '{factory_name}' không hợp lệ. Chỉ gồm: VH, VH2, VH3, VH4, JV, JV2, JVB"}

    with SFC_LOCK:
        try:
            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(CDP_URL, timeout=SFC_CONNECT_WAIT_MS)
                page = get_sfc_page(browser)

                if not page:
                    return {"success": False, "error": "Chưa kết nối được Chrome"}

                actual_plant = next((c['value'] for c in page.context.cookies()
                                     if c['name'] == 'MyFactoryName'), None)
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
    return sfc_hfpa.run(sys.modules[__name__], date_from, date_to, factory, action, nike,
                        export_path, expected_plant)


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
