import os
import sys
import re
import json
import logging
from flask import Flask, render_template, request, jsonify, Response, stream_with_context, send_from_directory
import openpyxl
import sfc_controller
import sfc_command_jobs
import sfc_report_tools
from sfc_commands import EXPORT_ROOT, COMMAND_LOCK
from dashboard_paths import APP_ROOT, mqaa_master_file
from sfc_downloads import EXCEL_EXTENSIONS

sys.stdout.reconfigure(encoding='utf-8')

app = Flask(__name__, template_folder="templates", static_folder="static")
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.jinja_env.auto_reload = True
from spreadsheet_compare_web import blueprint as comparison_blueprint
app.register_blueprint(comparison_blueprint)
from spreadsheet_combine_web import blueprint as combine_blueprint
app.register_blueprint(combine_blueprint)


@app.before_request
def protect_local_report_mutations():
    if request.method == 'POST' and request.path.startswith(('/api/report-tools/', '/api/compare/', '/api/combine/')):
        origin = request.headers.get('Origin')
        if origin and origin.rstrip('/') != request.host_url.rstrip('/'):
            return jsonify(error='Chỉ chấp nhận thao tác từ trang Dashboard đang mở.'), 403


@app.after_request
def private_report_responses(response):
    if request.path.startswith(('/api/report-tools', '/api/compare', '/api/combine')) or request.path in ('/reports', '/compare', '/combine'):
        response.headers['Cache-Control'] = 'no-store'
        response.headers['X-Content-Type-Options'] = 'nosniff'
    return response


@app.errorhandler(413)
def upload_too_large(error):
    return jsonify(error='Tổng file vượt giới hạn 640 MB. Chọn ít file hơn hoặc chia nhỏ dữ liệu.'), 413

@app.before_request
def prevent_overlapping_actions():
    if request.method == 'POST' and request.path in {
        '/api/action', '/api/nike', '/api/navigate', '/api/switch_factory', '/api/mqaa/scan', '/api/page/execute',
        '/api/command', '/api/command/stream'
    } and (sfc_command_jobs.running() or COMMAND_LOCK.locked() or sfc_controller.DYNAMIC_ACTION_LOCK.locked() or sfc_report_tools.browser_reserved()):
        return jsonify(success=False, error='Đang xử lý một thao tác SFC. Vui lòng chờ hoàn tất; không cần bấm xuất lại.'), 409

def format_download_filename(filename, plant=None):
    if not filename:
        return filename
    stem, ext = os.path.splitext(filename)
    if not plant:
        m = re.search(r'[-_](VH\d?|JV[B2]?)\b', stem, re.IGNORECASE)
        if m:
            plant = m.group(1).upper()
        else:
            plant = getattr(sfc_controller, 'SELECTED_FACTORIES', ['VH'])[0]
    plant = str(plant).upper().strip()
    plant_suffix = f"-{plant}"
    if stem.endswith(plant_suffix):
        return f"{stem}{ext}"
    return f"{stem}-{plant}{ext}"

@app.route('/api/exports/<path:filename>')
def api_export_file(filename):
    if not filename.lower().endswith(EXCEL_EXTENSIONS) or '.partial' in filename:
        return jsonify(error='Không phải file Excel đã hoàn tất.'), 404

    target_path = (EXPORT_ROOT / filename).resolve()
    if not target_path.is_relative_to(EXPORT_ROOT.resolve()):
        return jsonify(error='Đường dẫn file không hợp lệ.'), 404
    curr_plant = getattr(sfc_controller, 'SELECTED_FACTORIES', ['VH'])[0]

    # If the file doesn't exist under the requested path, try alternative with or without plant suffix
    if not target_path.is_file():
        stem, ext = os.path.splitext(target_path.name)
        parent = target_path.parent
        # Try finding with -<plant>
        alt1 = parent / f"{stem}-{curr_plant}{ext}"
        if alt1.is_file():
            target_path = alt1
        else:
            # Try removing -<plant> if requested with suffix
            m = re.search(r'-(VH\d?|JV[B2]?)$', stem, re.IGNORECASE)
            if m:
                base_stem = stem[:m.start()]
                alt2 = parent / f"{base_stem}{ext}"
                if alt2.is_file():
                    target_path = alt2

    if not target_path.is_file():
        return jsonify(error='File không tồn tại trên hệ thống.'), 404

    # Determine factory for download_name
    m = re.search(r'[-_](VH\d?|JV[B2]?)\b', target_path.stem, re.IGNORECASE)
    plant = m.group(1).upper() if m else curr_plant

    download_name = format_download_filename(target_path.name, plant)
    rel_path = target_path.relative_to(EXPORT_ROOT)
    return send_from_directory(EXPORT_ROOT, str(rel_path).replace('\\', '/'), as_attachment=True, download_name=download_name)

@app.route("/api/health", methods=["GET"])
def api_health():
    return jsonify({"service": "sfc-dashboard", "ready": True, "app_root": str(APP_ROOT)})


@app.route('/reports')
def report_tools_page():
    return render_template('report_tools.html')


@app.route('/api/report-tools')
def api_report_tools():
    return jsonify(tools=sfc_report_tools.registry(include_inputs=request.args.get('catalog') != '0'), runs=sfc_report_tools.runs())


@app.route('/api/report-tools/<tool>/inputs')
def api_report_tool_inputs(tool):
    source = request.args.get('source')
    if source not in ('local','sfc') or tool not in sfc_report_tools.TOOLS:
        return jsonify(error='Tool hoặc nguồn dữ liệu không hợp lệ.'), 400
    return jsonify(inputs=sfc_report_tools.input_catalog(tool, source=source))


@app.route('/api/report-tools/<tool>/runs', methods=['POST'])
@app.route('/api/report-tools/<tool>/validate', methods=['POST'])
def api_start_report_tool(tool):
    try:
        request.max_content_length = 640 * 1024 * 1024
        data = request.get_json(silent=True) if request.is_json else request.form.to_dict()
        data = data or {}
        if not isinstance(data, dict):
            raise ValueError('Dữ liệu yêu cầu phải là một đối tượng cấu hình.')
        if not all(isinstance(data.get(key, ''), str) for key in ('period','mode')):
            raise ValueError('Tháng và chế độ xử lý phải là chuỗi hợp lệ.')
        selections = data.get('selections', [])
        if isinstance(selections, str): selections = json.loads(selections)
        if not isinstance(selections, list) or not all(isinstance(item, dict) for item in selections):
            raise ValueError('Danh sách file đầu vào không hợp lệ.')
        if not all(isinstance(item.get('id'), str) and isinstance(item.get('group', ''), str) for item in selections):
            raise ValueError('Mã file hoặc nhóm dữ liệu không hợp lệ.')
        uploads = [(group, storage) for group in sfc_report_tools.TOOLS.get(tool, {}).get('groups', {})
                   for storage in request.files.getlist(group) if storage.filename]
        automatic = data.get('acquire_missing') in (True, 'true', '1', 1)
        validate_only = request.path.endswith('/validate')
        if automatic:
            if not COMMAND_LOCK.acquire(blocking=False):
                return jsonify(error='SFC đang chạy lệnh khác. Chờ hoàn tất rồi chạy báo cáo.'), 409
            locked = sfc_controller.SFC_LOCK.acquire(blocking=False)
            try:
                if not locked or sfc_command_jobs.running():
                    return jsonify(error='SFC đang bận. Chờ thao tác hiện tại hoàn tất.'), 409
                run = sfc_report_tools.start_run(tool, data.get('period'), data.get('mode','all'),
                                                selections, uploads, acquire_missing=True, validate_only=validate_only,
                                                cutoff=data.get('cutoff'))
            finally:
                if locked: sfc_controller.SFC_LOCK.release()
                COMMAND_LOCK.release()
        else:
            run = sfc_report_tools.start_run(tool, data.get('period'), data.get('mode', 'all'), selections, uploads, validate_only=validate_only)
        return jsonify(run), 200 if validate_only else 202
    except sfc_report_tools.PreflightError as error:
        return jsonify(error=str(error), valid=False, validation=error.reports), 422
    except RuntimeError as error:
        return jsonify(error=str(error)), 409
    except (ValueError, OSError, KeyError) as error:
        return jsonify(error=str(error) if isinstance(error, ValueError) else 'Không đọc/ghi được file. Kiểm tra quyền và đóng file đang mở.'), 400


@app.route('/api/report-tools/<tool>/cell', methods=['POST'])
def api_report_cell(tool):
    request.max_content_length = 256 * 1024
    data = request.get_json(silent=True)
    if not isinstance(data, dict) or not isinstance(data.get('token'), str):
        return jsonify(error='Yêu cầu sửa ô không hợp lệ.'), 400
    value = data.get('value')
    delete_row = data.get('delete_row', False)
    row_tokens = data.get('row_tokens', [])
    if (type(delete_row) is not bool or (delete_row and 'value' in data)
            or not isinstance(row_tokens, list) or len(row_tokens) > 99
            or any(not isinstance(token, str) for token in row_tokens)
            or (row_tokens and not delete_row)):
        return jsonify(error='Yêu cầu xóa dòng không hợp lệ.'), 400
    if 'value' in data and not isinstance(value, str):
        return jsonify(error='Nhập giá trị cần lưu.'), 400
    try:
        return jsonify(sfc_report_tools.report_cell(tool, data['token'], value, delete_row=delete_row, row_tokens=row_tokens))
    except RuntimeError as error:
        return jsonify(error=str(error)), 409
    except ValueError as error:
        return jsonify(error=str(error)), 422
    except OSError:
        return jsonify(error='Không đọc/ghi được file. Đóng file Excel và kiểm tra quyền truy cập rồi thử lại.'), 409


@app.route('/api/report-tools/runs/<run_id>/cancel', methods=['POST'])
def api_cancel_report_tool(run_id):
    try:
        return jsonify(sfc_report_tools.request_cancel(run_id)), 202
    except (ValueError, OSError) as error:
        return jsonify(error=str(error)), 409


@app.route('/api/report-tools/runs/<run_id>')
def api_report_tool_run(run_id):
    try:
        return jsonify(sfc_report_tools.snapshot(run_id))
    except (ValueError, OSError) as error:
        return jsonify(error=str(error)), 404


@app.route('/api/report-tools/runs/<run_id>/files/<path:filename>')
def api_report_tool_file(run_id, filename):
    try:
        run = sfc_report_tools.snapshot(run_id)
        allowed = {item['name'] for item in run.get('files', [])}
        if run.get('bundle'): allowed.add(run['bundle']['name'])
        output = (sfc_report_tools.run_folder(run_id) / 'workspace' / 'Output').resolve()
        path = (output / filename).resolve()
        if filename not in allowed or not path.is_relative_to(output) or not path.is_file():
            raise FileNotFoundError('File báo cáo không tồn tại hoặc chưa hoàn tất.')
        metadata = next((item for item in run.get('files', []) if item['name']==filename), run.get('bundle') or {})
        if metadata.get('sha256') and sfc_report_tools.file_hash(path) != metadata['sha256']:
            return jsonify(error='File đã thay đổi sau khi kiểm tra. Chạy lại hoặc kiểm tra bản gốc trước khi tải.'), 409
        if run.get('requires_ack') and request.args.get('ack') != '1':
            return jsonify(error='Có sai lệch hoặc kết quả chưa hoàn tất. Xem đối chiếu và xác nhận trước khi tải.', export_review=run.get('export_review')), 409
        response = send_from_directory(output, path.relative_to(output).as_posix(), as_attachment=True,
                                       download_name=sfc_report_tools.safe_name(path.name))
        response.headers['X-Content-Type-Options'] = 'nosniff'
        response.headers['Cache-Control'] = 'no-store'
        return response
    except (ValueError, OSError) as error:
        return jsonify(error=str(error)), 404

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
    app.run(host=os.environ.get('CHINGLUH_BIND_HOST', '127.0.0.1'), port=port, debug=False, threaded=True)

