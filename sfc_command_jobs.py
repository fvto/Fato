"""Keep long exports running when a dashboard tab disconnects or reloads."""
from datetime import datetime, timedelta, timezone
import json
import threading
from uuid import uuid4
from pathlib import Path

JOBS = {}
LOCK = threading.Lock()


def saved_results(limit=8):
    """Only completed, existing exports are offered as downloadable results."""
    from sfc_commands import EXPORT_ROOT
    from sfc_downloads import EXCEL_EXTENSIONS
    from urllib.parse import quote
    root = EXPORT_ROOT.resolve()
    results = []
    paths = sorted(EXPORT_ROOT.glob('*/manifest.json'), key=lambda p: p.stat().st_mtime, reverse=True)
    for path in paths:
        try:
            manifest = json.loads(path.read_text(encoding='utf-8'))
            partial = manifest.get('status') == 'error' and (bool(manifest.get('failed_plants')) or
                (manifest.get('intent', {}).get('combine') is True and bool(manifest.get('monthly_files'))))
            if manifest.get('status') != 'done' and not partial:
                continue
            files = []
            for item in manifest.get('combined_files') or manifest.get('monthly_files', []):
                source = Path(item['path']).resolve()
                if (not source.is_relative_to(root) or not source.is_file()
                        or source.suffix.lower() not in EXCEL_EXTENSIONS or '.partial' in source.name):
                    continue
                files.append(dict(item, name=source.name,
                    url='/api/exports/' + quote(source.relative_to(root).as_posix(), safe='/')))
            if files:
                results.append(dict(id='saved_' + path.parent.name, prompt=manifest['prompt'],
                    intent=manifest['intent'], files=files, monthly_count=len(manifest.get('monthly_files', [])),
                    status='error' if partial else 'done', error=manifest.get('error'),
                    saved_at=datetime.fromtimestamp(path.stat().st_mtime, timezone(timedelta(hours=7))).isoformat()))
                if len(results) >= limit:
                    break
        except (OSError, ValueError, KeyError, TypeError):
            continue
    return results


def restore_saved_result():
    with LOCK:
        if JOBS:
            return
        results = saved_results(limit=1)
        if not results:
            return
        result = results[0]
        job_id = result['id']
        events = [dict(step=1, status='success', msg='Đã khôi phục kết quả đã lưu.', intent=result['intent']),
                  dict(step=5, status=result['status'], msg=(result.get('error') or 'Kết quả đã hoàn tất; có thể tải file bên dưới.'),
                       files=result['files'], completed=result['monthly_count'], total=result['monthly_count'])]
        for index, item in enumerate(events):
            item.update(job_id=job_id, event_index=index)
        JOBS[job_id] = dict(id=job_id, prompt=result['prompt'], status=result['status'], events=events,
                           condition=threading.Condition(), started_at=result['saved_at'])


def latest_prompt():
    restore_saved_result()
    with LOCK:
        if JOBS:
            return JOBS[next(reversed(JOBS))]['prompt']
    from sfc_commands import EXPORT_ROOT
    path = EXPORT_ROOT / 'last_prompt.txt'
    return path.read_text(encoding='utf-8') if path.exists() else ''


def latest_job_id():
    restore_saved_result()
    with LOCK:
        return next(reversed(JOBS), '')


def running():
    with LOCK:
        return any(job['status'] == 'running' for job in JOBS.values())


def start(controller, prompt):
    restore_saved_result()
    with LOCK:
        if any(job['status'] == 'running' for job in JOBS.values()):
            raise ValueError('Đang xử lý một lệnh khác. Vui lòng chờ hoàn tất.')
        job_id = uuid4().hex
        job = dict(id=job_id, prompt=prompt, status='running', events=[], condition=threading.Condition(),
                   started_at=datetime.now(timezone(timedelta(hours=7))).isoformat())
        JOBS[job_id] = job
    from sfc_commands import EXPORT_ROOT
    EXPORT_ROOT.mkdir(parents=True, exist_ok=True)
    (EXPORT_ROOT / 'last_prompt.txt').write_text(prompt, encoding='utf-8')

    def work():
        try:
            for item in controller.execute_command_stream(prompt):
                with job['condition']:
                    item = dict(item, job_id=job_id, event_index=len(job['events']))
                    job['events'].append(item)
                    if item['status'] in ('done', 'error'):
                        job['status'] = item['status']
                    job['condition'].notify_all()
        except Exception as error:
            with job['condition']:
                job['status'] = 'error'
                job['events'].append(dict(step=5, status='error', msg=str(error), job_id=job_id, event_index=len(job['events'])))
                job['condition'].notify_all()
        finally:
            with job['condition']:
                if job['status'] == 'running':
                    job['status'] = 'error'
                    job['events'].append(dict(step=5, status='error', msg='Tác vụ dừng trước khi xác nhận kết quả.', job_id=job_id, event_index=len(job['events'])))
                job['condition'].notify_all()

    threading.Thread(target=work, daemon=True, name='sfc-command-' + job_id[:8]).start()
    return job


def snapshot(job_id):
    restore_saved_result()
    job = JOBS.get(job_id)
    if not job:
        return None
    with job['condition']:
        return {key: job[key] for key in ('id', 'prompt', 'status', 'started_at')} | {'events': list(job['events'])}


def stream(job_id, after=-1):
    job = JOBS[job_id]
    cursor = max(0, after + 1)
    while True:
        with job['condition']:
            if cursor >= len(job['events']) and job['status'] == 'running':
                job['condition'].wait(timeout=10)
            items = list(job['events'][cursor:])
            finished = job['status'] != 'running'
        if not items and not finished:
            yield ': keepalive\n\n'
        for item in items:
            yield 'data: ' + json.dumps(item, ensure_ascii=False) + '\n\n'
            cursor += 1
        if finished:
            return
