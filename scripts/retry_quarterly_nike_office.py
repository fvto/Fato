"""Retry only the native WPS/Office export of a failed NIKE quarterly run."""
from pathlib import Path
import json
import os
import shutil
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from report_runtime import Monitor, Cancelled, diagnostics, friendly_error
from sfc_report_tools import read_json, write_json, collect_files, make_bundle, now, process_alive
from report_exports import check_outputs
from report_tools.quarterly_defect_review.office_output import build_office_presentation
from report_tools.quarterly_defect_review.report import _build_business_review, _overall_status, STATUS_LABELS


def retry_office(folder):
    folder = Path(folder).resolve()
    run_path = folder / 'run.json'
    run = read_json(run_path)
    if (run.get('tool') != 'quarterly_defect_review' or run.get('mode') != 'all'
            or run.get('status') != 'error' or process_alive(run.get('pid'))):
        raise ValueError('Chỉ thử lại bước Office của báo cáo NIKE đã dừng do lỗi.')
    output = folder / 'workspace' / 'Output'
    summary_path = output / 'review_summary.json'
    summary = read_json(summary_path)
    if (summary.get('export_ok') or not summary.get('excel_meta')
            or not summary.get('export_error') or summary.get('period') != run.get('period')):
        raise ValueError('Không có kết quả Excel/phân tích phù hợp để thử lại Office.')
    excel = output / f'Quarterly_NIKE_Defect_Review_{run["period"]}.xlsx'
    pptx = excel.with_suffix('.pptx')
    if not excel.is_file():
        raise FileNotFoundError(excel)
    backup = folder / 'workspace' / 'office_retry_backup'
    backup.mkdir(exist_ok=True)
    for path in (excel, pptx, summary_path):
        if path.is_file() and not (backup / path.name).exists():
            shutil.copy2(path, backup / path.name)
    run.setdefault('office_retry_history', []).append({
        'started_at': now(), 'previous_error': summary['export_error'],
        'previous_diagnostics': run.get('diagnostics'),
    })
    for key in ('error', 'diagnostics', 'finished_at', 'progress', 'export_review'):
        run.pop(key, None)
    for step in run.get('steps', []):
        if step['id'] in ('presentation', 'export'):
            step.update(status='pending')
            step.pop('finished_at', None)
    run.update(status='running', pid=os.getpid(), requires_ack=True)
    monitor = Monitor(folder, run)
    try:
        monitor.stage('Đang thử lại xuất và mở kiểm tra bằng WPS, dùng kết quả phân tích đã có...', 'presentation')
        office_meta = build_office_presentation(
            pptx, final_excel_path=excel, excel_meta=summary['excel_meta'],
            period=summary['period'], date_from=summary['date_from'], date_to=summary['date_to'],
            hfpa_result=summary['hfpa'], ftt_result=summary['ftt'], preview_dir=output / 'previews',
        )
        summary.update(export_ok=True, export_error=None, office_meta=office_meta, office_retried_at=now())
        summary['files'].update(excel=str(excel), pptx=str(pptx), slide_previews=office_meta['slide_previews'])
        summary['status_code'] = _overall_status(summary['hfpa'], summary['ftt'], True)
        summary['status_label'] = STATUS_LABELS.get(summary['status_code'], summary['status_code'])
        summary['business_review'] = _build_business_review(
            summary['hfpa'], summary['ftt'], summary['catalog']['count'], office_meta)
        write_json(summary_path, summary)
        run['business_review'] = summary['business_review']
        run['summary'].update(status_code=summary['status_code'], status_label=summary['status_label'],
                              verification_app=office_meta['verification_app'])
        monitor.stage('Đang đối chiếu và đóng gói kết quả đã kiểm tra bằng WPS...', 'export')
        files = collect_files(folder)
        review = check_outputs(folder, run, files)
        run.update(files=files, export_review=review, requires_ack=review['status'] == 'warning',
                   bundle=make_bundle(folder, files))
        monitor.check_cancel()
        monitor.finish_step('done')
        run.update(status='done', stage='Đã hoàn tất; đã mở kiểm tra bằng WPS/Office.', finished_at=now())
    except BaseException as error:
        cancelled = isinstance(error, Cancelled)
        monitor.finish_step('cancelled' if cancelled else 'error')
        run.update(status='cancelled' if cancelled else 'error', error=friendly_error(error),
                   stage='Chưa hoàn tất bước kiểm tra Office.', diagnostics=diagnostics(error),
                   finished_at=now(), requires_ack=True, files=collect_files(folder))
        summary.update(export_ok=False, export_error=str(error))
        write_json(summary_path, summary)
    write_json(run_path, run)
    return run


if __name__ == '__main__':
    if sys.stdout is not None:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    if sys.stderr is not None:
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    result = retry_office(sys.argv[1])
    if sys.stdout is not None:
        print(json.dumps({'status': result['status'], 'stage': result['stage']}, ensure_ascii=False))
    raise SystemExit(0 if result['status'] == 'done' else 1)
