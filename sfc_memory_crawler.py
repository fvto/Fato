import os
import sys
import json
import time
from datetime import datetime
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

MEMORY_OUTPUT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sfc_memory_knowledge_base.json")

def crawl_entire_sfc():
    print("=" * 70)
    print("🕷️ BẮT ĐẦU CÀO DỮ LIỆU & XÂY DỰNG BỘ NHỚ TOÀN DIỆN CHO SFC CHING LUH")
    print("=" * 70)
    
    memory_db = {
        "crawled_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "portal": {},
        "production_modules": [],
        "qms_modules": [],
        "factories_and_lines": [],
        "hfpa_metadata": {},
        "mqaa_metadata": {},
        "available_endpoints": {}
    }

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        page = browser.contexts[0].pages[0]
        page.on("dialog", lambda dialog: dialog.accept())

        # -------------------------------------------------------------
        # 1. CÀO DỮ LIỆU TRANG CHỦ PORTAL (FlowChart.aspx)
        # -------------------------------------------------------------
        print("\n[1/5] Đang cào dữ liệu Trang chủ Portal (FlowChart.aspx)...")
        try:
            page.goto("https://sfc.chingluh.com/vhmes/FlowChart.aspx", timeout=15000)
            page.wait_for_load_state("networkidle", timeout=10000)
            time.sleep(1.5)

            portal_data = page.evaluate("""() => {
                // Lấy thông tin user
                const userEl = document.querySelector(".user-info, [class*='user'], [id*='user'], .header-right");
                const userName = userEl ? userEl.innerText.trim().replace(/\\s+/g, ' ') : '';
                
                // Lấy danh sách nhà máy (select)
                const factorySelect = document.querySelector("select");
                const factories = factorySelect ? Array.from(factorySelect.options).map(o => ({ value: o.value, text: o.text })) : [];
                
                // Lấy các chỉ số KPI
                const kpis = [];
                document.querySelectorAll(".card, [class*='kpi'], [class*='stat'], [class*='metric']").forEach(c => {
                    const t = c.innerText.trim().replace(/\\s+/g, ' ');
                    if (t && t.length < 100) kpis.push(t);
                });

                // Lấy các Production Modules
                const modules = [];
                document.querySelectorAll("a, [onclick*='clickModule']").forEach(el => {
                    const title = (el.innerText || '').trim();
                    const href = el.href || el.getAttribute('href') || '';
                    if (title && href && !href.endsWith('#') && !modules.some(m => m.href === href)) {
                        modules.push({ title: title.replace(/\\n+/g, ' - '), href });
                    }
                });

                return { userName, factories, kpis, modules };
            }""")

            memory_db["portal"] = {
                "title": page.title(),
                "url": page.url,
                "user_name": portal_data.get("userName"),
                "portal_factories": portal_data.get("factories"),
                "kpis": portal_data.get("kpis")
            }
            memory_db["production_modules"] = portal_data.get("modules", [])
            print(f"   ✅ Đã trích xuất {len(portal_data.get('modules', []))} Module sản xuất chính.")
        except Exception as e:
            print(f"   ⚠️ Lỗi khi cào FlowChart.aspx: {e}")

        # -------------------------------------------------------------
        # 2. CÀO DỮ LIỆU TOÀN BỘ PHÂN HỆ QMS (QMSMain.aspx)
        # -------------------------------------------------------------
        print("\n[2/5] Đang cào dữ liệu Cây danh mục Quản lý Chất lượng QMS (QMSMain.aspx)...")
        try:
            page.goto("https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx?dept_code=&tree_id=10", timeout=15000)
            page.wait_for_load_state("networkidle", timeout=10000)
            time.sleep(2)

            qms_data = page.evaluate("""() => {
                const items = [];
                // Lấy toàn bộ link trong navigation tree hoặc danh sách menu
                document.querySelectorAll("a").forEach(a => {
                    const text = a.innerText.trim();
                    const href = a.href || '';
                    if (text && href && !href.endsWith('#') && href.includes('chingluh.com')) {
                        const cleanHref = href.split('#')[0];
                        if (!items.some(it => it.href === cleanHref)) {
                            items.push({ name: text.replace(/\\n+/g, ' '), href: cleanHref });
                        }
                    }
                });
                return items;
            }""")

            memory_db["qms_modules"] = qms_data
            print(f"   ✅ Đã trích xuất {len(qms_data)} tính năng / màn hình thuộc phân hệ QMS!")
            for m in qms_data[:10]:
                print(f"      - {m['name']} -> {m['href']}")
        except Exception as e:
            print(f"   ⚠️ Lỗi khi cào QMSMain.aspx: {e}")

        # -------------------------------------------------------------
        # 3. CÀO DỮ LIỆU CHI TIẾT TRANG QA HFPA (QCHFPA.aspx)
        # -------------------------------------------------------------
        print("\n[3/5] Đang cào chi tiết 138 xưởng & cấu trúc dữ liệu QA HFPA (QCHFPA.aspx)...")
        try:
            page.goto("https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504", timeout=15000)
            page.wait_for_load_state("load", timeout=10000)
            time.sleep(2)

            hfpa_data = page.evaluate("""() => {
                const res = {
                    factories: [],
                    columns: [],
                    filters: []
                };

                // Lấy 138 xưởng
                try {
                    const cbo = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_cboDeptFrom');
                    if (cbo) {
                        for (let i = 0; i < cbo.GetItemCount(); i++) {
                            const it = cbo.GetItem(i);
                            if (it && it.text) {
                                res.factories.push({ text: it.text.trim(), value: it.value });
                            }
                        }
                    }
                } catch(e) {}

                // Lấy cấu trúc các cột bảng GridView grdExcel
                try {
                    document.querySelectorAll("[id*='grdExcel'] th, [id*='grdExcel_DXHeadersRow'] td").forEach(c => {
                        const t = c.innerText.trim();
                        if (t && !res.columns.includes(t)) res.columns.push(t);
                    });
                } catch(e) {}

                return res;
            }""")

            memory_db["factories_and_lines"] = hfpa_data.get("factories", [])
            memory_db["hfpa_metadata"] = {
                "title": page.title(),
                "url": page.url,
                "total_factories": len(hfpa_data.get("factories", [])),
                "grid_columns": hfpa_data.get("columns", [])
            }
            print(f"   ✅ Đã thu thập {len(hfpa_data.get('factories', []))} Xưởng/Chuyền và {len(hfpa_data.get('columns', []))} cột thuộc tính dữ liệu.")
        except Exception as e:
            print(f"   ⚠️ Lỗi khi cào QCHFPA.aspx: {e}")

        # -------------------------------------------------------------
        # 4. CÀO DỮ LIỆU BÁO CÁO MQAA (MQAAReport.aspx)
        # -------------------------------------------------------------
        print("\n[4/5] Đang cào cấu trúc Báo cáo MQAA Report (MQAAReport.aspx)...")
        try:
            page.goto("https://sfc.chingluh.com/vhmes/QA/MQAAReport.aspx", timeout=15000)
            page.wait_for_load_state("load", timeout=10000)
            time.sleep(2)

            report_meta = page.evaluate("""() => {
                const controls = [];
                document.querySelectorAll("input, select, button, [id*='cbo'], [id*='btn']").forEach(el => {
                    const id = el.id || '';
                    const tag = el.tagName;
                    const val = (el.innerText || el.value || '').trim();
                    if (id && !id.includes('_DDD_') && !id.includes('_Raw')) {
                        controls.push({ id, tag, val: val.slice(0, 40) });
                    }
                });
                return controls;
            }""")
            memory_db["mqaa_report_metadata"] = {
                "title": page.title(),
                "url": page.url,
                "controls": report_meta[:20]
            }
            print(f"   ✅ Đã thu thập cấu trúc điều khiển của MQAAReport.aspx ({len(report_meta)} controls).")
        except Exception as e:
            print(f"   ⚠️ Lỗi khi cào MQAAReport.aspx: {e}")

        # -------------------------------------------------------------
        # 5. CÀO DỮ LIỆU AQL TRACKING (AQLTracking.aspx)
        # -------------------------------------------------------------
        print("\n[5/5] Đang cào cấu trúc Theo dõi AQL (AQLTracking.aspx)...")
        try:
            page.goto("https://sfc.chingluh.com/vhmes/QA/AQLTracking.aspx", timeout=15000)
            page.wait_for_load_state("load", timeout=10000)
            time.sleep(2)

            aql_meta = page.evaluate("""() => {
                const controls = [];
                document.querySelectorAll("input, select, button, [id*='cbo'], [id*='btn']").forEach(el => {
                    const id = el.id || '';
                    const tag = el.tagName;
                    const val = (el.innerText || el.value || '').trim();
                    if (id && !id.includes('_DDD_') && !id.includes('_Raw')) {
                        controls.push({ id, tag, val: val.slice(0, 40) });
                    }
                });
                return controls;
            }""")
            memory_db["aql_metadata"] = {
                "title": page.title(),
                "url": page.url,
                "controls": aql_meta[:20]
            }
            print(f"   ✅ Đã thu thập cấu trúc điều khiển của AQLTracking.aspx ({len(aql_meta)} controls).")
        except Exception as e:
            print(f"   ⚠️ Lỗi khi cào AQLTracking.aspx: {e}")

        # Quay lại trang HFPA mặc định
        page.goto("https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504")
        time.sleep(1)

        # Lưu bộ nhớ vào JSON
        with open(MEMORY_OUTPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(memory_db, f, ensure_ascii=False, indent=2)

        print("\n" + "=" * 70)
        print(f"🎉 HOÀN THÀNH CÀO DỮ LIỆU & LƯU VÀO BỘ NHỚ: {MEMORY_OUTPUT_FILE}")
        print("=" * 70)
        return memory_db

if __name__ == "__main__":
    crawl_entire_sfc()
