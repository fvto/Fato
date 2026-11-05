"""HTTP endpoints for the local spreadsheet comparison workspace."""
from flask import Blueprint, Response, jsonify, render_template, request, send_file

import spreadsheet_compare as compare
import spreadsheet_chat as chat
from spreadsheet_chat_tools import artifact_csv
from spreadsheet_compare_html import build_report
import spreadsheet_workflows as workflows

blueprint = Blueprint('spreadsheet_compare', __name__)
ISSUES = {'difference', 'missing_value', 'missing_input', 'formula_error', 'formula_uncached',
          'invalid_input', 'invalid_value', 'ambiguous_key', 'invalid_key', 'unmatched'}


@blueprint.errorhandler(ValueError)
def invalid_request(error):
    return jsonify(error=str(error)), 400


@blueprint.route('/compare')
def page():
    return render_template('spreadsheet_compare.html')


@blueprint.route('/api/compare/chat/status')
def chat_status():
    return jsonify(chat.status())


@blueprint.route('/api/compare/chat/messages', methods=['POST'])
def chat_message():
    return jsonify(chat.submit(request.get_json(silent=True) or {})), 202


@blueprint.route('/api/compare/chat/import', methods=['POST'])
def import_chat_answer():
    return jsonify(chat.import_answer(request.get_json(silent=True) or {})), 201


@blueprint.route('/api/compare/chat/conversations/<chat_id>')
def chat_conversation(chat_id):
    return jsonify(chat.conversation(chat_id))


@blueprint.route('/api/compare/chat/jobs/<job_id>')
def chat_job(job_id):
    return jsonify(chat.public_job(job_id))


@blueprint.route('/api/compare/chat/jobs/<job_id>/artifacts/<artifact_id>.<extension>')
def chat_artifact(job_id, artifact_id, extension):
    artifact = chat.read_artifact(job_id, artifact_id)
    if extension == 'csv':
        return Response(artifact_csv(artifact), content_type='text/csv; charset=utf-8',
                        headers={'Content-Disposition': f'attachment; filename=Evidence_{artifact_id[:8]}.csv'})
    if extension == 'json':
        return Response(compare.json.dumps(artifact, ensure_ascii=False), content_type='application/json; charset=utf-8',
                        headers={'Content-Disposition': f'attachment; filename=Evidence_{artifact_id[:8]}.json'})
    raise ValueError('Unknown evidence format.')


@blueprint.route('/api/compare/catalog')
def catalog():
    items = compare.catalog(limit=None)
    return jsonify(files=[{k: v for k, v in item.items() if k != 'relative'} for item in items[:80]],
                   total=len(items), delete_ids=[item['id'] for item in items])


@blueprint.route('/api/compare/catalog/delete-all', methods=['POST'])
def delete_all_exports():
    data = request.get_json(silent=True) or {}
    if not isinstance(data, dict):
        raise ValueError('Invalid export selection.')
    return jsonify(compare.delete_exports(data.get('catalog_ids')))


@blueprint.route('/api/compare/catalog/<catalog_id>/delete', methods=['POST'])
def delete_export(catalog_id):
    return jsonify(deleted=compare.delete_export(catalog_id))


@blueprint.route('/api/compare/workbooks', methods=['POST'])
def upload():
    if request.is_json:
        selection = next((item for item in compare.catalog() if item['id'] == (request.get_json() or {}).get('catalog_id')), None)
        if selection is None:
            raise ValueError('Choose an available saved export.')
        path = (compare.EXPORTS / selection['relative']).resolve()
        if not path.is_relative_to(compare.EXPORTS.resolve()):
            raise ValueError('Export path is outside the report folder.')
        with path.open('rb') as stream:
            return jsonify(workbooks=[compare.import_file(stream, selection['name'])])
    files = request.files.getlist('files')
    if not 1 <= len(files) <= 5:
        raise ValueError('Choose between 1 and 5 Excel files.')
    try:
        return jsonify(workbooks=[compare.import_file(file.stream, file.filename) for file in files])
    except ValueError:
        raise
    except Exception as error:
        raise ValueError('Cannot read this Excel file. Check that it opens in Excel / WPS.') from error


@blueprint.route('/api/compare/workbooks/<workbook_id>')
def describe(workbook_id):
    return jsonify(compare.describe(compare.workbook(workbook_id)))


@blueprint.route('/api/compare/workbooks/<workbook_id>/download')
def download_workbook(workbook_id):
    book = compare.workbook(workbook_id)
    path = compare.ROOT / 'inputs' / book['id'] / ('source' + book['extension'])
    return send_file(path, as_attachment=True, download_name=book['name'])


def result_page(result):
    status = request.args.get('status', 'issues')
    rows = result['results']
    if status == 'issues':
        rows = [row for row in rows if row['status'] in ISSUES]
    elif status != 'all':
        rows = [row for row in rows if row['status'] == status]
    try:
        offset = max(0, int(request.args.get('offset', '0')))
    except ValueError as error:
        raise ValueError('Invalid page offset.') from error
    return {k: v for k, v in result.items() if k != 'results'} | {
        'results': rows[offset:offset + 200], 'filtered_total': len(rows), 'offset': offset,
        'issue_count': sum(result['counts'].get(status, 0) for status in ISSUES),
        'csv_url': f'/api/compare/results/{result["id"]}/download.csv',
        'json_url': f'/api/compare/results/{result["id"]}/download.json',
        'html_url': f'/api/compare/results/{result["id"]}/download.html'}


@blueprint.route('/api/compare/run', methods=['POST'])
def run():
    data = request.get_json(silent=True) or {}
    try:
        result = compare.run(data.get('rules'))
    except (KeyError, TypeError) as error:
        raise ValueError('Incomplete check. Review file, sheet and range settings.') from error
    return jsonify(result_page(result))


@blueprint.route('/api/compare/workflows/preview', methods=['POST'])
def preview_workflow():
    return jsonify(workflows.preview(request.get_json(silent=True) or {}))


@blueprint.route('/api/compare/workflows/<plan_id>/run', methods=['POST'])
def run_workflow(plan_id):
    data = request.get_json(silent=True) or {}
    if not isinstance(data, dict):
        raise ValueError('Invalid workflow input.')
    return jsonify(result_page(workflows.execute(plan_id, data.get('acknowledged'))))


@blueprint.route('/api/compare/results/<result_id>')
def result(result_id):
    return jsonify(result_page(compare.read_result(result_id)))


@blueprint.route('/api/compare/results/<result_id>/download.<extension>')
def download(result_id, extension):
    result = compare.read_result(result_id)
    if extension == 'pptx':
        from spreadsheet_pptx import pptx_bytes
        return Response(pptx_bytes(result),content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                        headers={'Content-Disposition': f'attachment; filename=Comparison_Report_{result_id[:8]}.pptx'})
    if extension == 'html':
        return Response(build_report(result), content_type='text/html; charset=utf-8',
                        headers={'Content-Disposition': f'attachment; filename=Comparison_Report_{result_id[:8]}.html'})
    if extension == 'csv':
        return Response(compare.result_csv(result), content_type='text/csv; charset=utf-8',
                        headers={'Content-Disposition': f'attachment; filename=Comparison_{result_id[:8]}.csv'})
    if extension == 'json':
        return Response(compare.json.dumps(result, ensure_ascii=False), content_type='application/json; charset=utf-8',
                        headers={'Content-Disposition': f'attachment; filename=Comparison_{result_id[:8]}.json'})
    raise ValueError('Unknown download format.')
