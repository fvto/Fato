import os
import sys
import json
import logging
from flask import Flask, render_template, request, jsonify, Response, stream_with_context, send_from_directory
import openpyxl
import sfc_controller
import sfc_command_jobs
from sfc_commands import EXPORT_ROOT, COMMAND_LOCK
from dashboard_paths import APP_ROOT, mqaa_master_file

sys.stdout.reconfigure(encoding='utf-8')

app = Flask(__name__, template_folder="templates", static_folder="static")
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.jinja_env.auto_reload = True

@app.before_request
def prevent_overlapping_actions():
    if request.method == 'POST' and request.path in {
        '/api/action', '/api/nike', '/api/navigate', '/api/switch_factory', '/api/mqaa/scan', '/api/page/execute'
    } and (sfc_command_jobs.running() or COMMAND_LOCK.locked()):
        return jsonify(success=False, error='Trợ lý đang chạy lệnh. Vui lòng chờ hoàn tất để tránh thay đổi bộ lọc.'), 409

@app.route('/api/exports/<path:filename>')
def api_export_file(filename):
    if not filename.lower().endswith(('.xls', '.xlsx')) or '.partial' in filename:
        return jsonify(error='Không phải file Excel đã hoàn tất.'), 404
    return send_from_directory(EXPORT_ROOT, filename, as_attachment=True)

@app.route("/api/health", methods=["GET"])
def api_health():
    return jsonify({"service": "sfc-dashboard", "ready": True, "app_root": str(APP_ROOT)})

@app.route("/")
def index():
    return render_template("index.html", last_command_prompt=sfc_command_jobs.latest_prompt(),
                           last_command_job_id=sfc_command_jobs.latest_job_id(),
                           saved_command_results=sfc_command_jobs.saved_results())

@app.route("/api/status", methods=["GET"])
def api_status():
    if sfc_command_jobs.running():
        return jsonify(dict(sfc_controller.CONNECTION_STATUS_CACHE or {'connected': False}, command_running=True))
    status = (sfc_controller.get_live_status() if request.args.get('screen') == '1'
              else sfc_controller.get_connection_status())
    return jsonify(status)

@app.route("/api/factories", methods=["GET"])
def api_factories():
    factories = sfc_controller.get_factories()
    return jsonify({"factories": factories, "count": len(factories)})

@app.route("/api/switch_factory", methods=["POST"])
def api_switch_factory():
    data = request.get_json() or {}
    factories = data.get("factories")
    if factories and isinstance(factories, list):
        res = sfc_controller.set_selected_factories(factories)
        return jsonify(res)
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

DEFAULT_NAV_TILES = [
    {
        "id": "hfpa",
        "title": "QA HFPA",
        "sub": "Mã Q504",
        "dest": "hfpa",
        "icon": "fa-solid fa-clipboard-list",
        "color": "#38bdf8",
        "bg": "rgba(56, 189, 248, 0.08)",
        "borderColor": "rgba(56, 189, 248, 0.35)"
    },
    {
        "id": "copq",
        "title": "COPQ Report",
        "sub": "3 Dạng Xuất",
        "dest": "copq",
        "icon": "fa-solid fa-file-excel",
        "color": "#34d399",
        "bg": "rgba(16, 185, 129, 0.08)",
        "borderColor": "rgba(16, 185, 129, 0.45)"
    },
    {
        "id": "mqaa",
        "title": "MQAA Check",
        "sub": "Kiểm toán 21 Mẫu",
        "dest": "mqaa",
        "icon": "fa-solid fa-shoe-prints",
        "color": "#fbbf24",
        "bg": "rgba(251, 191, 36, 0.08)",
        "borderColor": "rgba(251, 191, 36, 0.35)"
    },
    {
        "id": "qms",
        "title": "QMS Quality",
        "sub": "Cây Quản trị QA",
        "dest": "qms",
        "icon": "fa-solid fa-diagram-project",
        "color": "#a78bfa",
        "bg": "rgba(167, 139, 250, 0.08)",
        "borderColor": "rgba(167, 139, 250, 0.35)"
    },
    {
        "id": "home",
        "title": "SFC Portal",
        "sub": "Trang chủ FlowChart",
        "dest": "home",
        "icon": "fa-solid fa-house",
        "color": "#94a3b8",
        "bg": "rgba(148, 163, 184, 0.08)",
        "borderColor": "rgba(148, 163, 184, 0.35)"
    }
]

@app.route("/api/nav/config", methods=["GET", "POST"])
def api_nav_config():
    cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "custom_nav_config.json")
    if request.method == "POST":
        data = request.get_json() or {}
        tiles = data.get("tiles")
        if not isinstance(tiles, list):
            return jsonify({"success": False, "error": "Dữ liệu cấu hình không hợp lệ."}), 400
        try:
            with open(cfg_path, "w", encoding="utf-8") as f:
                json.dump(tiles, f, ensure_ascii=False, indent=2)
            return jsonify({"success": True, "tiles": tiles})
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    # GET
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                return jsonify({"success": True, "tiles": json.load(f)})
        except Exception:
            pass
    return jsonify({"success": True, "tiles": DEFAULT_NAV_TILES})


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
    nike = data.get("nike")
    if nike is not None and not isinstance(nike, bool):
        return jsonify({"success": False, "error": "NIKE phải là giá trị bật/tắt."}), 400
    res = sfc_controller.apply_sfc_filters(date_from=date_from, date_to=date_to, factory=factory, action=action, nike=nike)
    return jsonify(res)

@app.route("/api/nike", methods=["POST"])
def api_nike():
    data = request.get_json() or {}
    nike = data.get("nike")
    if not isinstance(nike, bool):
        return jsonify({"success": False, "error": "NIKE phải là giá trị bật/tắt."}), 400
    return jsonify(sfc_controller.apply_sfc_filters(nike=nike))

@app.route("/api/page/inspect", methods=["GET"])
def api_page_inspect():
    res = sfc_controller.inspect_active_page()
    return jsonify(res)

@app.route("/api/page/execute", methods=["POST"])
def api_page_execute():
    data = request.get_json() or {}
    action_type = data.get("action_type")
    target_id = data.get("target_id")
    params = data.get("params", {})
    if not action_type:
        return jsonify({"success": False, "error": "Thiếu loại thao tác (action_type)."}), 400
    res = sfc_controller.execute_dynamic_action(action_type, target_id, params)
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

    try:
        job = sfc_command_jobs.start(sfc_controller, prompt)
    except ValueError as error:
        return jsonify(success=False, error=str(error)), 409
    return command_stream_response(job['id'])

def command_stream_response(job_id, after=-1):
    return Response(stream_with_context(sfc_command_jobs.stream(job_id, after)), mimetype="text/event-stream", headers={
        "Cache-Control": "no-cache",
        "X-Accel-Buffering": "no"
    })

@app.route('/api/command/jobs/<job_id>')
def api_command_job(job_id):
    job = sfc_command_jobs.snapshot(job_id)
    return (jsonify(job), 200) if job else (jsonify(error='Không tìm thấy tác vụ; server có thể đã khởi động lại.'), 404)

@app.route('/api/command/jobs/<job_id>/stream')
def api_command_job_stream(job_id):
    if not sfc_command_jobs.snapshot(job_id):
        return jsonify(error='Không tìm thấy tác vụ.'), 404
    try:
        after = int(request.args.get('after', '-1'))
    except ValueError:
        return jsonify(error='Mốc sự kiện không hợp lệ.'), 400
    return command_stream_response(job_id, after)

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
    xlsx_path = mqaa_master_file()
    if not xlsx_path:
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
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(name)s %(message)s')
    port = 5050
    print(f"🚀 SFC CONTROL COCKPIT SERVER DANG CHAY TAI: http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)

