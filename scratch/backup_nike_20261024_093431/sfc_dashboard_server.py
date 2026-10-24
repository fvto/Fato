import os
import sys
import json
from flask import Flask, render_template, request, jsonify, Response, stream_with_context
import openpyxl
import sfc_controller

sys.stdout.reconfigure(encoding='utf-8')

app = Flask(__name__, template_folder="templates", static_folder="static")

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/status", methods=["GET"])
def api_status():
    status = sfc_controller.get_live_status()
    return jsonify(status)

@app.route("/api/factories", methods=["GET"])
def api_factories():
    factories = sfc_controller.get_factories()
    return jsonify({"factories": factories, "count": len(factories)})

@app.route("/api/switch_factory", methods=["POST"])
def api_switch_factory():
    data = request.get_json() or {}
    factory = data.get("factory", "VH")
    res = sfc_controller.switch_factory(factory)
    return jsonify(res)

@app.route("/api/memory", methods=["GET"])

def api_memory():
    kb_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sfc_memory_knowledge_base.json")
    if os.path.exists(kb_path):
        try:
            with open(kb_path, "r", encoding="utf-8") as f:
                return jsonify(json.load(f))
        except Exception as e:
            return jsonify({"error": str(e)})
    return jsonify({"modules": [], "qms_modules": []})

@app.route("/api/navigate", methods=["POST"])
def api_navigate():
    data = request.get_json() or {}
    dest = data.get("destination", "hfpa")
    res = sfc_controller.navigate_to(dest)
    return jsonify(res)

@app.route("/api/action", methods=["POST"])
def api_action():
    data = request.get_json() or {}
    date_from = data.get("dateFrom")
    date_to = data.get("dateTo")
    factory = data.get("factory")
    action = data.get("action")
    res = sfc_controller.apply_sfc_filters(date_from=date_from, date_to=date_to, factory=factory, action=action)
    return jsonify(res)

@app.route("/api/command", methods=["POST"])
def api_command():
    data = request.get_json() or {}
    prompt = data.get("prompt", "")
    if not prompt:
        return jsonify({"success": False, "error": "Lệnh không được để trống!"}), 400
    res = sfc_controller.parse_and_execute_command(prompt)
    return jsonify(res)

@app.route("/api/command/stream", methods=["POST"])
def api_command_stream():
    data = request.get_json() or {}
    prompt = data.get("prompt", "")
    if not prompt:
        def err_stream():
            yield f"data: {json.dumps({'step': 0, 'status': 'error', 'msg': 'Lệnh không được để trống!'}, ensure_ascii=False)}\n\n"
        return Response(stream_with_context(err_stream()), mimetype="text/event-stream")

    def generate():
        try:
            for step in sfc_controller.execute_command_stream(prompt):
                yield f"data: {json.dumps(step, ensure_ascii=False)}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'step': 99, 'status': 'error', 'msg': f'Lỗi hệ thống: {str(e)}'}, ensure_ascii=False)}\n\n"

    return Response(stream_with_context(generate()), mimetype="text/event-stream", headers={
        "Cache-Control": "no-cache",
        "X-Accel-Buffering": "no"
    })

@app.route("/api/mqaa/scan", methods=["POST"])
def api_mqaa_scan():
    res = sfc_controller.start_mqaa_scan_background()
    return jsonify(res)

@app.route("/api/mqaa/open_excel", methods=["POST"])
def api_mqaa_open_excel():
    res = sfc_controller.open_master_excel()
    return jsonify(res)

@app.route("/api/mqaa/summary", methods=["GET"])
def api_mqaa_summary():
    workspace = os.path.dirname(os.path.abspath(__file__))
    xlsx_path = os.path.join(workspace, "MQAA_Master_DoiChieu_21_Forms_MoiNhat.xlsx")
    if not os.path.exists(xlsx_path):
        xlsx_path = os.path.join(workspace, "MQAA_Master_DoiChieu_21_Forms.xlsx")
        
    if not os.path.exists(xlsx_path):
        return jsonify({"has_data": False, "message": "Chưa có file Master Excel."})

    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        if "TongQuan_21_Forms" not in wb.sheetnames:
            return jsonify({"has_data": False, "message": "Không tìm thấy sheet TongQuan_21_Forms"})
            
        ws = wb["TongQuan_21_Forms"]
        rows = []
        for r in range(3, 24):
            stt = ws.cell(row=r, column=1).value
            if stt is None:
                continue
            cat = ws.cell(row=r, column=2).value
            model = ws.cell(row=r, column=3).value
            q_web = ws.cell(row=r, column=5).value
            q_ex = ws.cell(row=r, column=6).value
            rmk_web = ws.cell(row=r, column=7).value
            rmk_ex = ws.cell(row=r, column=8).value
            match_q = ws.cell(row=r, column=9).value
            match_rmk = ws.cell(row=r, column=10).value
            pct = ws.cell(row=r, column=11).value
            status = ws.cell(row=r, column=12).value
            
            rows.append({
                "stt": stt,
                "cat": cat,
                "model": model,
                "q_web": q_web,
                "q_ex": q_ex,
                "rmk_web": rmk_web,
                "rmk_ex": rmk_ex,
                "match_q": match_q,
                "match_rmk": match_rmk,
                "pct": pct,
                "status": status
            })
            
        total_models = len(rows)
        perfect_count = sum(1 for r in rows if "100%" in str(r.get("status", "")))
        
        return jsonify({
            "has_data": True,
            "file": os.path.basename(xlsx_path),
            "total_models": total_models,
            "perfect_count": perfect_count,
            "rows": rows
        })
    except Exception as e:
        return jsonify({"has_data": False, "error": str(e)})

if __name__ == "__main__":
    port = 5050
    print(f"🚀 SFC CONTROL COCKPIT SERVER DANG CHAY TAI: http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)

