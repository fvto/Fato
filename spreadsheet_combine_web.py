"""HTTP workspace for combining user-selected local Excel tables."""
import json
import re
from flask import Blueprint, jsonify, render_template, request, send_file
import spreadsheet_combine as combine

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
    stored_name = combine.output_filename(result.get('output_name', 'Combined.xlsx'))
    output = result_folder / stored_name
    if not output.is_file():
        raise ValueError('File gộp chưa sẵn sàng. Hãy thực hiện gộp lại.')
    # Results created before named exports remain downloadable with a useful name.
    result['output_name'] = combine.output_filename(result.get('output_name') or combine.suggested_filename(result['sources']))
    return output, result


@blueprint.errorhandler(ValueError)
def invalid(error):
    return jsonify(error=str(error)), 400


@blueprint.route('/combine')
def workspace():
    return render_template('spreadsheet_combine.html', max_files=combine.MAX_FILES)


@blueprint.route('/api/combine/uploads', methods=['POST'])
def uploads():
    request.max_content_length = 640 * 1024 * 1024
    return jsonify(combine.upload(request.files.getlist('files')))


@blueprint.route('/api/combine/<job_id>/preview', methods=['POST'])
def preview(job_id):
    data = configuration()
    return jsonify(combine.summary(combine.prepare(job_id, data.get('selections'), data.get('options'))))


@blueprint.route('/api/combine/<job_id>/run', methods=['POST'])
def run(job_id):
    data = configuration()
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
