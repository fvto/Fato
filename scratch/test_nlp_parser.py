import sys
import re
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

SEVEN_FACTORIES = ["VH4", "VH3", "VH2", "VH", "JVB", "JV2", "JV"] # Check longer names first!

def parse_intent(text):
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

    # 1. Detect 7 Plants
    for plant in SEVEN_FACTORIES:
        p_low = plant.lower()
        # Look for plant as a word
        if re.search(rf'\b{p_low}\b', t):
            # Make sure it's not part of line name like "VH - D3" or "VL - D3" unless explicitly "nhà máy"
            out["target_plant"] = plant
            out["summary"].append(f"Nhà máy: {plant}")
            break

    # 2. Detect Destination Page
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

    # 3. Detect Dates
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

    # 4. Detect Line / Chuyền
    line_match = re.search(r'\b(vl\s*-\s*[a-z0-9]+|va\s*-\s*[a-z0-9\s]+line\s*\d+|vb\s*-\s*[a-z0-9\s]+line\s*\d+)\b', t)
    if line_match:
        out["target_line"] = line_match.group(1).upper()
        out["summary"].append(f"Chuyền: {out['target_line']}")
    else:
        # Search for words following chuyền/xưởng
        chuyen_match = re.search(r'(chuyền|xưởng|line|máy)\s+([a-zA-Z0-9\s\-]+?)(?=(\s+ngày|\s+từ|\s+và|\s+rồi|\s+tìm|$))', t)
        if chuyen_match:
            cand = chuyen_match.group(2).strip()
            if cand.upper() not in SEVEN_FACTORIES:
                out["target_line"] = cand
                out["summary"].append(f"Chuyền: {cand}")

    # 5. Detect Action
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

# Test phrases
tests = [
    "đổi sang nhà máy VH3",
    "chọn nhà máy VH4 và lọc chuyền VL - D3 hôm nay rồi tìm kiếm",
    "xuất excel chuyền VL-D4 từ ngày 2026-10-20 đến 2026-10-24",
    "vào mục kiểm toán MQAA",
    "chuyển qua nhà máy JV",
    "về trang chủ SFC",
    "chọn chuyền VL - D3 ngày hôm nay"
]

for s in tests:
    print(f"\nCâu lệnh: '{s}'")
    res = parse_intent(s)
    print(" -> Phân tích được:", " | ".join(res['summary']) if res['summary'] else "Không rõ")
