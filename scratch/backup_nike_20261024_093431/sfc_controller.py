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

sys.stdout.reconfigure(encoding='utf-8')

CDP_URL = "http://localhost:9222"

SFC_LOCK = threading.Lock()
CACHED_FACTORIES = []
MQAA_SCAN_PROCESS = None
MQAA_STATUS = {"running": False, "progress": "", "last_run": None}

LAST_STATUS_TIME = 0
LAST_STATUS_CACHE = None

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

def take_b64_screenshot(page):
    try:
        img_bytes = page.screenshot(type="jpeg", quality=60)
        return f"data:image/jpeg;base64,{base64.b64encode(img_bytes).decode('utf-8')}"
    except Exception:
        return ""

def get_live_status():
    """Lấy trạng thái thực tế của SFC có khóa an toàn đa luồng & cache nhẹ 1.5s"""
    global LAST_STATUS_TIME, LAST_STATUS_CACHE
    now = time.time()
    if LAST_STATUS_CACHE and (now - LAST_STATUS_TIME < 1.5):
        return LAST_STATUS_CACHE

    with SFC_LOCK:
        try:
            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(CDP_URL)
                page = None
                for ctx in browser.contexts:
                    for pg in ctx.pages:
                        if 'chingluh' in pg.url or 'vhmes' in pg.url:
                            page = pg
                            break
                    if page:
                        break
                if not page and browser.contexts and browser.contexts[0].pages:
                    page = browser.contexts[0].pages[0]

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
                browser = p.chromium.connect_over_cdp(CDP_URL)
                page = None
                for ctx in browser.contexts:
                    for pg in ctx.pages:
                        if 'chingluh' in pg.url or 'vhmes' in pg.url:
                            page = pg
                            break
                    if page:
                        break
                if not page and browser.contexts and browser.contexts[0].pages:
                    page = browser.contexts[0].pages[0]

                if not page:
                    return {"success": False, "error": "Chưa kết nối được Chrome"}

                if not page.query_selector("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0_"):
                    page.goto("https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx", wait_until="domcontentloaded", timeout=15000)
                    time.sleep(0.5)

                page.click("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0_")
                time.sleep(0.3)

                selector = f"#wucHeaderMainTop_TopPanel_menuRightTop_DXI0i{idx}_"
                el = page.query_selector(selector)
                if el:
                    el.click()
                    time.sleep(2)
                    try:
                        page.wait_for_load_state("load", timeout=5000)
                    except Exception:
                        pass

                    curr_text = ""
                    header_el = page.query_selector("#wucHeaderMainTop_TopPanel_menuRightTop_DXI0_")
                    if header_el:
                        curr_text = header_el.inner_text().strip()

                    img_b64 = take_b64_screenshot(page)
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
                browser = p.chromium.connect_over_cdp(CDP_URL)
                page = browser.contexts[0].pages[0]
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
                browser = p.chromium.connect_over_cdp(CDP_URL)
                page = browser.contexts[0].pages[0]
                page.goto(target_url, wait_until="domcontentloaded", timeout=20000)
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

def apply_sfc_filters(date_from=None, date_to=None, factory=None, action=None):
    """Thao tác trên trang QA HFPA"""
    global LAST_STATUS_TIME
    with SFC_LOCK:
        try:
            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(CDP_URL)
                page = browser.contexts[0].pages[0]

                if 'QCHFPA' not in page.url:
                    page.goto("https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504", wait_until="domcontentloaded", timeout=15000)
                    time.sleep(1.5)

                res = page.evaluate("""(params) => {
                    const out = {};
                    const coll = typeof aspxGetControlCollection !== 'undefined' ? aspxGetControlCollection() : null;
                    if (!coll) return { error: 'Không tìm thấy DevExpress control collection' };

                    if (params.dateFrom) {
                        const dtF = coll.Get('rpnlMain_ContentPlaceHolder1_dtDateFrom');
                        if (dtF) {
                            dtF.SetText(params.dateFrom);
                            out.dateFrom = dtF.GetText();
                        }
                    }

                    if (params.dateTo) {
                        const dtT = coll.Get('rpnlMain_ContentPlaceHolder1_dtDateTo');
                        if (dtT) {
                            dtT.SetText(params.dateTo);
                            out.dateTo = dtT.GetText();
                        }
                    }

                    if (params.factory) {
                        const cbo = coll.Get('rpnlMain_ContentPlaceHolder1_cboDeptFrom');
                        if (cbo) {
                            const target = params.factory.toLowerCase().trim();
                            let matched = false;
                            for (let i = 0; i < cbo.GetItemCount(); i++) {
                                const it = cbo.GetItem(i);
                                if (it && it.text && (it.text.toLowerCase() === target || it.text.toLowerCase().includes(target))) {
                                    cbo.SetSelectedIndex(i);
                                    out.factory = it.text;
                                    matched = true;
                                    break;
                                }
                            }
                            if (!matched) out.factoryWarning = 'Không tìm thấy chuyền: ' + params.factory;
                        }
                    }

                    return out;
                }""", {"dateFrom": date_from, "dateTo": date_to, "factory": factory})

                time.sleep(0.5)

                action_res = None
                if action == 'search':
                    page.evaluate("""() => {
                        const btn = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_btnSearch');
                        if (btn) btn.DoClick();
                    }""")
                    try:
                        page.wait_for_load_state("load", timeout=10000)
                    except Exception:
                        pass
                    time.sleep(2.0)

                    try:
                        row_info = page.evaluate("""() => {
                            const grid = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_grdExcel');
                            return grid ? {
                                visibleRows: grid.GetVisibleRowsOnPage(),
                                pageCount: grid.GetPageCount ? grid.GetPageCount() : 1
                            } : null;
                        }""")
                        action_res = f"Đã tìm kiếm xong. Tìm thấy {row_info.get('visibleRows', 0)} hàng trên trang!" if row_info else "Đã bấm tìm kiếm."
                    except Exception:
                        action_res = "Đã gửi lệnh tìm kiếm."

                elif action == 'excel':
                    page.evaluate("""() => {
                        const btn = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_btnExcel');
                        if (btn) btn.DoClick();
                    }""")
                    time.sleep(3)
                    action_res = "Đã kích hoạt lệnh tải file Excel từ máy chủ SFC."

                img_b64 = take_b64_screenshot(page)
                browser.close()
                LAST_STATUS_TIME = 0

                return {
                    "success": True,
                    "filter_result": res,
                    "action_result": action_res,
                    "screenshot": img_b64
                }
        except Exception as e:
            return {"success": False, "error": str(e)}

def parse_intent(text):
    """Bộ bóc tách thực thể và ý định tiếng Việt siêu thông minh"""
    t = text.lower().strip()
    out = {
        "target_plant": None,
        "target_line": None,
        "target_page": None,
        "date_from": None,
        "date_to": None,
        "action": None,
        "summary": []
    }

    # 1. Phát hiện 7 Nhà máy (Ưu tiên theo độ dài giảm dần để match VH2, VH3, VH4, JV2, JVB trước VH, JV)
    sorted_plants = sorted(SEVEN_FACTORIES, key=len, reverse=True)
    for plant in sorted_plants:
        p_low = plant.lower()
        if re.search(rf'\b{p_low}\b', t):
            out["target_plant"] = plant
            out["summary"].append(f"Nhà máy: {plant}")
            break

    # Hỗ trợ thêm "nhà máy 1" -> VH, "nhà máy 2" -> VH2, "nhà máy 3" -> VH3, "nhà máy 4" -> VH4
    if not out["target_plant"]:
        nm_match = re.search(r'nhà máy\s*([1-4])\b', t)
        if nm_match:
            num = nm_match.group(1)
            mapped = "VH" if num == "1" else f"VH{num}"
            out["target_plant"] = mapped
            out["summary"].append(f"Nhà máy: {mapped}")

    # 2. Phát hiện Phân hệ / Màn hình
    if any(k in t for k in ['hfpa', 'q504']):
        out["target_page"] = "hfpa"
        out["summary"].append("Trang: QA HFPA (Q504)")
    elif any(k in t for k in ['báo cáo mqaa', 'mqaareport']):
        out["target_page"] = "mqaa_report"
        out["summary"].append("Trang: Báo cáo MQAA")
    elif any(k in t for k in ['mqaa', 'kiểm toán']):
        out["target_page"] = "mqaa"
        out["summary"].append("Trang: MQAA Check")
    elif any(k in t for k in ['aql', 'aqltracking']):
        out["target_page"] = "aql"
        out["summary"].append("Trang: AQL Tracking")
    elif any(k in t for k in ['trang chủ', 'portal', 'flowchart']):
        out["target_page"] = "home"
        out["summary"].append("Trang: Trang chủ FlowChart")
    elif any(k in t for k in ['qms', 'quản lý chất lượng']):
        out["target_page"] = "qms"
        out["summary"].append("Trang: QMS Quality Portal")

    # 3. Phát hiện Ngày tháng
    today_str = datetime.now().strftime("%Y-%m-%d")
    if 'hôm nay' in t or 'today' in t:
        out["date_from"] = today_str
        out["date_to"] = today_str
        out["summary"].append(f"Ngày: {today_str} (Hôm nay)")
    else:
        dates_found = re.findall(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})|(\d{1,2}[-/]\d{1,2}[-/]\d{4})', t)
        if dates_found:
            parsed = []
            for d1, d2 in dates_found:
                raw = (d1 or d2).replace('/', '-')
                parts = raw.split('-')
                if len(parts[0]) == 4:
                    parsed.append(f"{parts[0]}-{int(parts[1]):02d}-{int(parts[2]):02d}")
                else:
                    parsed.append(f"{parts[2]}-{int(parts[1]):02d}-{int(parts[0]):02d}")
            if len(parsed) >= 2:
                out["date_from"], out["date_to"] = parsed[0], parsed[1]
                out["summary"].append(f"Khoảng ngày: {parsed[0]} ➔ {parsed[1]}")
            elif len(parsed) == 1:
                out["date_from"], out["date_to"] = parsed[0], parsed[0]
                out["summary"].append(f"Ngày: {parsed[0]}")

    # 4. Phát hiện Chuyền sản xuất
    line_match = re.search(r'\b(vl\s*-\s*[a-z0-9]+|va\s*-\s*[a-z0-9\s]+line\s*\d+|vb\s*-\s*[a-z0-9\s]+line\s*\d+)\b', t)
    if line_match:
        out["target_line"] = line_match.group(1).upper()
        out["summary"].append(f"Chuyền: {out['target_line']}")
    else:
        chuyen_match = re.search(r'(chuyền|xưởng|line|máy)\s+([a-zA-Z0-9\s\-]+?)(?=(\s+ngày|\s+từ|\s+và|\s+rồi|\s+tìm|$))', t)
        if chuyen_match:
            cand = chuyen_match.group(2).strip()
            if cand.upper() not in SEVEN_FACTORIES:
                out["target_line"] = cand
                out["summary"].append(f"Chuyền: {cand}")

    # 5. Phát hiện Hành động
    if any(k in t for k in ['xuất excel', 'tải excel', 'download excel', 'export', 'lấy excel']):
        out["action"] = "excel"
        out["summary"].append("Hành động: Xuất file Excel")
    elif any(k in t for k in ['tìm kiếm', 'tìm', 'search', 'lọc', 'tra cứu', 'xem']):
        out["action"] = "search"
        out["summary"].append("Hành động: Tìm kiếm dữ liệu")
    elif out["target_line"] or out["date_from"]:
        out["action"] = "search"
        out["summary"].append("Hành động: Tự động Tìm kiếm")

    return out

def execute_command_stream(user_prompt):
    """
    Generator phát trực tiếp (Realtime Streaming) từng bước xử lý cho giao diện
    """
    # Bước 1: Phân tích câu lệnh
    yield {
        "step": 1,
        "title": "Phân tích ngữ nghĩa câu lệnh",
        "status": "running",
        "msg": f"🧠 Đang bóc tách ý định từ: '{user_prompt}'..."
    }
    time.sleep(0.3)

    intent = parse_intent(user_prompt)
    yield {
        "step": 1,
        "title": "Phân tích ngữ nghĩa câu lệnh",
        "status": "success",
        "msg": f"✅ Đã nhận diện: { ' | '.join(intent['summary']) if intent['summary'] else 'Không có bộ lọc cụ thể'}",
        "intent": intent
    }

    last_screenshot = None

    # Bước 2: Đổi Nhà máy (nếu có yêu cầu)
    if intent["target_plant"]:
        yield {
            "step": 2,
            "title": f"Chuyển Nhà máy: {intent['target_plant']}",
            "status": "running",
            "msg": f"🏭 Đang thao tác chuyển menu góc phải sang Nhà máy {intent['target_plant']}..."
        }
        res_plant = switch_factory(intent["target_plant"])
        if res_plant.get("success"):
            last_screenshot = res_plant.get("screenshot")
            yield {
                "step": 2,
                "title": f"Chuyển Nhà máy: {intent['target_plant']}",
                "status": "success",
                "msg": f"✅ {res_plant.get('message')}",
                "screenshot": last_screenshot,
                "plant": intent["target_plant"]
            }
        else:
            yield {
                "step": 2,
                "title": f"Chuyển Nhà máy: {intent['target_plant']}",
                "status": "warn",
                "msg": f"⚠️ {res_plant.get('error')}"
            }
    else:
        yield {
            "step": 2,
            "title": "Chuyển Nhà máy",
            "status": "skipped",
            "msg": "⏭️ Giữ nguyên Nhà máy đang hoạt động"
        }

    # Bước 3: Điều hướng màn hình (nếu có yêu cầu)
    if intent["target_page"]:
        page_name = intent["target_page"].upper()
        yield {
            "step": 3,
            "title": f"Điều hướng phân hệ: {page_name}",
            "status": "running",
            "msg": f"🧭 Đang mở trang {page_name} trên SFC..."
        }
        res_nav = navigate_to(intent["target_page"])
        if res_nav.get("success"):
            last_screenshot = res_nav.get("screenshot")
            yield {
                "step": 3,
                "title": f"Điều hướng phân hệ: {page_name}",
                "status": "success",
                "msg": f"✅ Đã mở trang: {res_nav.get('title')}",
                "screenshot": last_screenshot
            }
        else:
            yield {
                "step": 3,
                "title": f"Điều hướng phân hệ: {page_name}",
                "status": "warn",
                "msg": f"⚠️ Lỗi điều hướng: {res_nav.get('error')}"
            }
    else:
        yield {
            "step": 3,
            "title": "Điều hướng phân hệ",
            "status": "skipped",
            "msg": "⏭️ Giữ nguyên phân hệ làm việc hiện tại"
        }

    # Bước 4: Thiết lập bộ lọc Chuyền/Ngày & Thực hiện Tìm kiếm/Excel
    if intent["target_line"] or intent["date_from"] or intent["action"]:
        act_desc = "Xuất file Excel" if intent["action"] == "excel" else "Tìm kiếm dữ liệu"
        yield {
            "step": 4,
            "title": f"Thao tác dữ liệu: {act_desc}",
            "status": "running",
            "msg": f"⚙️ Đang điền Chuyền: '{intent['target_line'] or 'Mặc định'}', Ngày: {intent['date_from'] or 'Hôm nay'} và thực hiện {act_desc}..."
        }
        res_action = apply_sfc_filters(
            date_from=intent["date_from"],
            date_to=intent["date_to"],
            factory=intent["target_line"],
            action=intent["action"]
        )
        if res_action.get("success"):
            last_screenshot = res_action.get("screenshot")
            yield {
                "step": 4,
                "title": f"Thao tác dữ liệu: {act_desc}",
                "status": "success",
                "msg": f"✅ {res_action.get('action_result')}",
                "screenshot": last_screenshot
            }
        else:
            yield {
                "step": 4,
                "title": f"Thao tác dữ liệu: {act_desc}",
                "status": "warn",
                "msg": f"⚠️ {res_action.get('error')}"
            }
    else:
        yield {
            "step": 4,
            "title": "Thao tác dữ liệu",
            "status": "skipped",
            "msg": "⏭️ Không yêu cầu bộ lọc hoặc xuất dữ liệu bổ sung"
        }

    # Bước 5: Hoàn tất tổng thể
    if not last_screenshot:
        try:
            status_now = get_live_status()
            last_screenshot = status_now.get("screenshot")
        except Exception:
            pass

    yield {
        "step": 5,
        "title": "Hoàn tất yêu cầu",
        "status": "done",
        "msg": "🎉 Toàn bộ chuỗi thao tác đã được thực hiện xong và đồng bộ thành công!",
        "screenshot": last_screenshot
    }

def parse_and_execute_command(user_prompt):
    """Wrapper đồng bộ cho các lệnh API thông thường"""
    last_res = None
    for step in execute_command_stream(user_prompt):
        last_res = step
    return {
        "command": user_prompt,
        "details": last_res.get("msg") if last_res else "Đã hoàn thành",
        "result": last_res
    }

def start_mqaa_scan_background():
    """Kích hoạt quét 21 Form MQAA ngầm"""
    global MQAA_SCAN_PROCESS, MQAA_STATUS
    if MQAA_STATUS["running"]:
        return {"status": "already_running", "message": "Quá trình quét 21 form đang chạy dở!"}

    workspace = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(workspace, "auto_scan_all_21_models.py")

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
    workspace = os.path.dirname(os.path.abspath(__file__))
    f1 = os.path.join(workspace, "MQAA_Master_DoiChieu_21_Forms_MoiNhat.xlsx")
    f2 = os.path.join(workspace, "MQAA_Master_DoiChieu_21_Forms.xlsx")
    target = f1 if os.path.exists(f1) else f2
    if os.path.exists(target):
        os.startfile(target)
        return {"success": True, "file": os.path.basename(target)}
    return {"success": False, "error": "Chưa tìm thấy file Master Excel"}
