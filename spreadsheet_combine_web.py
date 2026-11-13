"""HTTP workspace for combining user-selected local Excel tables."""
import json
import re
from flask import Blueprint, jsonify, render_template, request, send_file
import spreadsheet_combine as combine
import spreadsheet_combine_tasks as tasks

blueprint = Blueprint('spreadsheet_combine', __name__)


def configuration():
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        raise ValueError('Lựa chọn bảng không hợp lệ.')
    return data


def saved_result(job_id, result_id):
    folder = combine.folder_for(job_id)
    if not re.fullmatch(r'[a-f0-9]{32}', str(result_id or '')):
        raise ValueError('Không tìm thấy file gộp.')
    result_folder = folder / result_id
    marker = result_folder / 'result.json'
    if not marker.is_file():
        raise ValueError('File gộp chưa sẵn sàng. Hãy thực hiện gộp lại.')
    result = json.loads(marker.read_text(encoding='utf-8'))
    if result.get('kind')=='power_query':
        name=result.get('output_name','')
        if not name.endswith('.zip'):raise ValueError('Tên gói Power Query không hợp lệ.')
        stored_name=combine.output_filename(name[:-4])[:-5]+'.zip'
        if stored_name!=name:raise ValueError('Tên gói Power Query không hợp lệ.')
    else:stored_name = combine.output_filename(result.get('output_name', 'Combined.xlsx'))
    output = result_folder / stored_name
    if not output.is_file():
        raise ValueError('File gộp chưa sẵn sàng. Hãy thực hiện gộp lại.')
    # Results created before named exports remain downloadable with a useful name.
    if result.get('kind')!='power_query':
        result['output_name'] = combine.output_filename(result.get('output_name') or combine.suggested_filename(result['sources']))
    return output, result


@blueprint.errorhandler(ValueError)
def invalid(error):
    if isinstance(error, combine.UploadValidationError):
        return jsonify(error=error.message, errors=error.errors), 400
    return jsonify(error=str(error)), 400


@blueprint.errorhandler(413)
def upload_too_large(error):
    return jsonify(error='Tổng file vượt giới hạn 5 GB. Hãy chia nhỏ lượt gộp.'), 413


@blueprint.route('/combine')
def workspace():
    return render_template('spreadsheet_combine.html', max_files=combine.MAX_FILES,
                           max_bytes=combine.MAX_BYTES, max_file_mb=combine.MAX_BYTES // (1024 * 1024),
                           max_total_bytes=combine.MAX_TOTAL_BYTES)


@blueprint.route('/api/combine/uploads', methods=['POST'])
def uploads():
    # Allow multipart metadata in addition to the exact aggregate file budget.
    # upload() enforces the file bytes themselves, including the 5 GiB boundary.
    request.max_content_length = combine.MAX_TOTAL_BYTES + 1024 * 1024
    try:
        csv_options = json.loads(request.form.get('csv_options', '{}'))
    except (ValueError, TypeError) as error:
        raise ValueError('Tùy chọn CSV không hợp lệ.') from error
    files=request.files.getlist('files')
    if (request.content_length or 0)>8*1024*1024:
        return jsonify(tasks.stage_upload(files,csv_options)),202
    return jsonify(combine.upload(files, csv_options))


@blueprint.route('/api/combine/tasks/<task_id>')
def task_status(task_id):
    return jsonify(tasks.snapshot(task_id))


@blueprint.route('/api/combine/<job_id>/inputs')
def saved_inputs(job_id):
    folder=combine.folder_for(job_id)
    items=json.loads((folder/'inputs.json').read_text(encoding='utf-8'))
    return jsonify(id=job_id,xlsx_row_threshold=combine.XLSX_ROW_THRESHOLD,
        files=[{key:value for key,value in item.items() if key not in {'stored','original_stored'}} for item in items])


@blueprint.route('/api/combine/<job_id>/preview', methods=['POST'])
def preview(job_id):
    data = configuration()
    if tasks.is_large(job_id):
        return jsonify(tasks.start(lambda progress: combine.summary(combine.prepare(job_id,data.get('selections'),data.get('options'),progress)))),202
    return jsonify(combine.summary(combine.prepare(job_id, data.get('selections'), data.get('options'))))


@blueprint.route('/api/combine/<job_id>/search-index', methods=['POST'])
def search_index(job_id):
    data=configuration()
    from spreadsheet_combine_search import create
    return jsonify(tasks.start(lambda progress: create(job_id,data.get('selections'),data.get('options'),progress))),202


@blueprint.route('/api/combine/<job_id>/search', methods=['POST'])
def search(job_id):
    data=configuration()
    from spreadsheet_combine_search import query
    return jsonify(tasks.start(lambda progress: query(job_id,data.get('index_id'),data.get('filters'),data.get('page',1),progress))),202


@blueprint.route('/api/combine/<job_id>/run', methods=['POST'])
def run(job_id):
    data = configuration()
    if tasks.is_large(job_id):
        return jsonify(tasks.start(lambda progress: combine.export(job_id,data.get('selections'),data.get('output_name'),data.get('options'),progress))),202
    return jsonify(combine.export(job_id, data.get('selections'), data.get('output_name'), data.get('options')))


@blueprint.route('/api/combine/<job_id>/<result_id>/download')
def download(job_id, result_id):
    output, result = saved_result(job_id, result_id)
    return send_file(output, as_attachment=True, download_name=result['output_name'])


@blueprint.route('/api/combine/example')
def example():
    marker = combine.ROOT / 'example.json'
    if not marker.is_file():
        return jsonify(example=None)
    result = json.loads(marker.read_text(encoding='utf-8'))
    try:
        _, result = saved_result(result['id'], result['result_id'])
    except ValueError:
        return jsonify(example=None)
    return jsonify(example=result)
