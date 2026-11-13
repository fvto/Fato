"""Local report engines, immutable input staging, and restart-safe report jobs."""
from datetime import datetime, timedelta, timezone
from pathlib import Path
import ctypes
import hashlib
from functools import lru_cache
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import tempfile
import time
from urllib.parse import quote
from uuid import uuid4
import zipfile

from dashboard_paths import APP_ROOT
from sfc_excel import EXPORT_ROOT
from report_validation import validate_file, validate_preflight, validate_reinspection_coverage, PreflightError, CONTRACTS, OLE
from report_runtime import friendly_error
from report_cell_edits import edit_token, edit_cell

TOOLS_ROOT = APP_ROOT / 'report_tools'
RUN_ROOT = APP_ROOT / 'Report' / 'ToolRuns'
LOCK = threading.Lock()
WORKERS = {}
EXCEL = ('.xls', '.xlsx', '.xlsm', '.csv')
DELIVERABLES = ('.xlsx', '.xls', '.pptx', '.csv', '.json', '.md', '.html', '.txt')
COMMON_PLANTS = ('VH', 'VH2', 'JV', 'JV2')
TOOLS = json.loads((APP_ROOT / 'report_registry.json').read_text(encoding='utf8'))


def now():
    return datetime.now(timezone(timedelta(hours=7))).isoformat()


def write_json(path, data):
    path = Path(path)
    tmp = path.with_name(path.name + '.' + uuid4().hex + '.tmp')
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf8')
    try:
        for attempt in range(12):
            try:
                tmp.replace(path)
                return
            except PermissionError:
                if attempt == 11:
                    raise
                # Windows readers/antivirus may briefly deny replacement of an open file.
                time.sleep(.02 * (attempt + 1))
    finally:
        tmp.unlink(missing_ok=True)


def read_json(path):
    for attempt in range(12):
        try:
            return json.loads(Path(path).read_text(encoding='utf8'))
        except PermissionError:
            if attempt == 11:
                raise
            time.sleep(.02 * (attempt + 1))


def original_root(tool):
    config_path = APP_ROOT / 'report_tools_config.json'
    config = read_json(config_path) if config_path.exists() else {}
    configured = config.get('source_roots', {}).get(tool)
    return Path(configured).expanduser() if configured else Path.home() / 'Desktop' / TOOLS[tool]['source']


def input_formats(tool, group):
    return tuple(TOOLS[tool]['formats'][group])


def safe_name(name):
    name = Path(str(name).replace('\\', '/')).name
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f\x7f]', '_', name).strip(' .')
    if not name or len(name) > 180 or name.split('.')[0].upper() in {
            'CON','PRN','AUX','NUL',*(f'COM{i}' for i in range(1,10)),*(f'LPT{i}' for i in range(1,10))}:
        raise ValueError('Tên file không an toàn. Đổi tên ngắn hơn và thêm mã nhà máy.')
    return name


def file_hash(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def plant_of(name):
    matches = re.findall(r'(?<![A-Z0-9])(VH[234]?|JV[23B]?)(?![A-Z0-9])', name.upper())
    return ('JVB' if matches[-1] == 'JV3' else matches[-1]) if matches else None


def period_of(name):
    match = re.search(r'(20\d{2})-(\d{2})-\d{2}[_ ](20\d{2})-(\d{2})-\d{2}', name)
    if match:
        if match.groups()[:2] != match.groups()[2:]:
            return 'mixed'
        return match[1] + '-' + match[2]
    return None  # MES410 names contain the download date, not necessarily the audit period.


def group_hint(name):
    name = name.lower()
    if 'cost of poor quality' in name or 'copq' in name:
        return 'copq'
    if 're-inspection' in name or 'reinspection' in name:
        return 'reinspection'
    if 'qastation' in name or 'ftt' in name:
        return 'ftt'
    if 'hfpa' in name:
        return 'hfpa'
    return None


@lru_cache(maxsize=512)
def workbook_period(filename, size, modified):
    """MES410 filenames carry export timestamps; read the actual audit dates."""
    import openpyxl
    try:
        with Path(filename).open('rb') as stream:
            workbook = openpyxl.load_workbook(stream, read_only=True, data_only=True)
            try:
                rows = workbook.worksheets[0].iter_rows(values_only=True)
                column, periods = None, set()
                for index, row in enumerate(rows):
                    if index >= 70:
                        break
                    if column is None:
                        column = next((i for i,v in enumerate(row) if isinstance(v,str) and
                                       ('audit date' in v.lower() or v.lower().strip() in ('inspdate', 'inspection date', 'date'))), None)
                        continue
                    value = row[column] if column < len(row) else None
                    if isinstance(value, datetime):
                        periods.add(value.strftime('%Y-%m'))
                    elif value:
                        for fmt in ('%Y-%m-%d', '%d-%b-%Y', '%d-%b-%y', '%d/%m/%Y', '%m/%d/%Y'):
                            try:
                                periods.add(datetime.strptime(str(value).strip(), fmt).strftime('%Y-%m'))
                                break
                            except ValueError:
                                continue
                return next(iter(periods)) if len(periods) == 1 else ('mixed' if periods else None)
            finally:
                workbook.close()
    except Exception:
        return None


def input_catalog(tool, source=None, *, include_duplicates=False):
    if tool not in TOOLS:
        raise ValueError('Tool không hợp lệ.')
    if TOOLS[tool].get('upload_only'):
        return []
    result = []
    for group, (_, relative) in TOOLS[tool]['groups'].items():
        if source == 'sfc':
            continue
        root = original_root(tool) / relative
        for path in sorted(root.rglob('*')):
            if not path.is_file() or path.suffix.lower() not in input_formats(tool, group) or path.name.startswith('~$') or any(part.startswith('.report-edit-') for part in path.parts):
                continue
            relative = path.relative_to(root).as_posix()
            stat = path.stat()
            period = period_of(path.name)
            if period is None and group == 'hfpa':
                period = workbook_period(str(path), stat.st_size, stat.st_mtime_ns)
            result.append(dict(id='local:' + group + ':' + relative, relative=relative, group=group, name=path.name,
                size=stat.st_size, period=period, plant=plant_of(path.name), source='local'))
    if source == 'local':
        return result
    if tool == 'quarterly_defect_review':
        qdr_cache = APP_ROOT / 'Cache' / 'QuarterlyDefectReview'
        if qdr_cache.is_dir():
            for path in sorted(qdr_cache.rglob('*'), key=lambda p: p.stat().st_mtime if p.is_file() else 0, reverse=True):
                if not path.is_file() or path.name.startswith('~$') or '.partial' in path.name:
                    continue
                rel_posix = path.relative_to(qdr_cache).as_posix()
                low = rel_posix.lower()
                if path.name == 'line_catalog_factory_sandals.json':
                    group = 'sandals_catalog'
                elif low.startswith('hfpa_nike/') and path.suffix.lower() in input_formats(tool, 'hfpa_nike'):
                    group = 'hfpa_nike'
                elif low.startswith('ftt_nike/') and path.suffix.lower() in input_formats(tool, 'ftt_nike'):
                    group = 'ftt_nike'
                else:
                    continue
                result.append(dict(
                    id='qdr:' + rel_posix,
                    group=group,
                    name=path.name,
                    size=path.stat().st_size,
                    period=None,
                    plant=plant_of(path.name),
                    source='sfc',
                ))
        return result
    seen_exports = set()
    for path in sorted(EXPORT_ROOT.rglob('*'), key=lambda p:p.stat().st_mtime if p.is_file() else 0, reverse=True):
        if not path.is_file() or path.suffix.lower() not in EXCEL or '.partial' in path.name or path.name.startswith('~$') or any(part.startswith('.report-edit-') for part in path.parts):
            continue
        group = group_hint(path.name)
        if group not in TOOLS[tool]['groups'] or plant_of(path.name) not in TOOLS[tool]['plants']:
            continue
        if path.suffix.lower() not in input_formats(tool, group):
            continue
        # Only raw COPQ Data belongs in the COPQ processor; summary exports are different schemas.
        if group == 'copq' and ('quality report' in path.name.lower() or 'combined' in path.name.lower()):
            continue
        if not include_duplicates and (group, path.name) in seen_exports:
            continue
        seen_exports.add((group, path.name))
        result.append(dict(id='sfc:' + path.relative_to(EXPORT_ROOT).as_posix(), group=group,
            name=path.name, size=path.stat().st_size, period=period_of(path.name), plant=plant_of(path.name), source='sfc'))
    return result


def registry(include_inputs=True):
    missing = [name for name in ('pandas', 'numpy', 'openpyxl', 'xlrd', 'pptx', 'PIL')
               if importlib.util.find_spec(name) is None]
    result = []
    for key, definition in TOOLS.items():
        tool_missing = missing + [name for name in definition.get('dependencies', [])
                                  if importlib.util.find_spec(name) is None]
        items = input_catalog(key) if include_inputs else []
        primary = next(iter(definition['groups']))
        periods = [item['period'] for item in items if item['source'] == 'local' and item['group'] == primary and item['period'] not in (None, 'mixed')]
        current = datetime.now(timezone(timedelta(hours=7)))
        default_period = max(periods) if periods else current.strftime('%Y-%m')
        period_options = []
        if definition.get('period_kind') == 'fiscal_quarter':
            fiscal_year = current.year % 100 + int(current.month >= 6)
            quarter = (current.month - 6) % 12 // 3 + 1
            index = fiscal_year * 4 + quarter - 1
            period_options = [f'FY{i // 4:02d}-Q{i % 4 + 1}' for i in range(index, index - 24, -1)]
            configured_default = definition.get('default_period')
            default_period = configured_default if configured_default in period_options else period_options[0]
        result.append(dict(id=key, title=definition['title'], description=definition['description'],
            plants=list(definition['plants']), groups=[dict(id=k, title=v[0], formats=list(input_formats(key,k))) for k,v in definition['groups'].items()],
            modes=[dict(id=k, title=v) for k,v in definition['modes'].items()], default_period=default_period,
            default_date_from=definition.get('default_date_from'), default_date_to=definition.get('default_date_to'),
            contract_version=CONTRACTS['version'], max_file_bytes=CONTRACTS['max_file_bytes'],
            flow=['upload','validate','execute','review'], automatic=definition.get('automatic', False), required_groups=definition['required_groups'],
            upload_only=definition.get('upload_only', False), period_options=period_options,
            **({'today': current.date().isoformat()} if period_options else {}),
            ready=(TOOLS_ROOT / key).is_dir() and not tool_missing, missing_dependencies=tool_missing, inputs=items))
    return result


def process_alive(pid):
    if not pid:
        return False
    if os.name == 'nt':
        kernel = ctypes.windll.kernel32
        kernel.OpenProcess.restype = ctypes.c_void_p
        handle = kernel.OpenProcess(0x1000, False, int(pid))
        if not handle:
            return False
        try:
            code = ctypes.c_ulong()
            return bool(kernel.GetExitCodeProcess(ctypes.c_void_p(handle), ctypes.byref(code))) and code.value == 259
        finally:
            kernel.CloseHandle(ctypes.c_void_p(handle))
    try:
        os.kill(int(pid), 0)
        return True
    except (OSError, ValueError):
        return False


def run_folder(run_id):
    if not re.fullmatch(r'[0-9a-f]{32}', run_id):
        raise ValueError('Mã lượt chạy không hợp lệ.')
    return RUN_ROOT / run_id


def snapshot(run_id):
    folder = run_folder(run_id)
    path = folder / 'run.json'
    if not path.is_file():
        raise FileNotFoundError('Không tìm thấy lượt chạy báo cáo.')
    run = read_json(path)
    stale_queue = run['status'] == 'queued' and (datetime.fromisoformat(now()) - datetime.fromisoformat(run['started_at'])).total_seconds() >= 60
    if stale_queue or run['status'] == 'running' and not process_alive(run.get('pid')):
        # A worker may have committed its final status after our first read, then exited.
        run = read_json(path)
        if run['status'] in ('queued','running') and not process_alive(run.get('pid')):
            cancelled = (folder / 'cancel.request').exists()
            run.update(status='cancelled' if cancelled else 'error', error='Tiến trình xử lý đã dừng. Các kết quả đã lưu vẫn được giữ lại.', finished_at=now(),requires_ack=True)
            for step in run.get('steps', []):
                if step['status'] == 'running':
                    step.update(status=run['status'], finished_at=now())
            run['files'] = collect_files(folder)
            write_json(path, run)
    # Do not expose raw legacy logs (some engines print company data).
    run['log'] = '\n'.join(f"{step.get('started_at','')} · {step['title']} · {step['status']}" for step in run.get('steps', []))
    if not run.get('monitor_version') and run.get('error'):
        run['error'] = 'Lượt chạy cũ chưa hoàn tất. Kiểm tra dữ liệu và chạy lại khi cần.'
    run['cancel_requested'] = (folder / 'cancel.request').exists()
    run['can_cancel'] = bool(run.get('monitor_version') and run['status'] in ('queued','running') and not run['cancel_requested'])
    acquisition = run.get('acquisition')
    if acquisition and run.get('phase') == 'acquisition':
        run['progress'] = dict(completed=len(acquisition.get('reused', []))+len(acquisition.get('downloaded', [])), total=acquisition['total'], unit='file đầu vào đã xác nhận')
    if run.get('export_review') and 'brief' not in run['export_review']:
        from report_brief import summarize_checks
        run['export_review']['brief'] = summarize_checks(run['export_review'].get('checks', []))
    return run


def request_cancel(run_id):
    with LOCK:
        run = snapshot(run_id)
        if not run.get('can_cancel'):
            raise ValueError('Lượt chạy đã kết thúc hoặc không hỗ trợ hủy an toàn.')
        (run_folder(run_id) / 'cancel.request').touch(exist_ok=True)
        return snapshot(run_id)


def runs(limit=12):
    for run_id, worker in list(WORKERS.items()):
        if worker.poll() is not None:
            del WORKERS[run_id]
    RUN_ROOT.mkdir(parents=True, exist_ok=True)
    result = []
    for path in sorted(RUN_ROOT.glob('*/run.json'), key=lambda p:p.stat().st_mtime, reverse=True)[:limit]:
        try:
            result.append(snapshot(path.parent.name))
        except (OSError, ValueError, KeyError):
            continue
    return result


def browser_reserved(exclude_pid=None):
    for path in RUN_ROOT.glob('*/run.json'):
        try:
            run = read_json(path)
            if run.get('phase') != 'acquisition' or run.get('status') not in ('queued','running'):
                continue
            if exclude_pid and run.get('pid') == exclude_pid:
                continue
            if run['status'] == 'queued':
                if (datetime.fromisoformat(now()) - datetime.fromisoformat(run['started_at'])).total_seconds() < 60:
                    return True
            elif process_alive(run.get('pid')):
                return True
        except (OSError, ValueError, KeyError):
            continue
    return False


def input_path(tool, item):
    if item['source'] == 'local':
        root = original_root(tool) / TOOLS[tool]['groups'][item['group']][1]
        path = root / item['relative']
    elif str(item.get('id', '')).startswith('qdr:'):
        root = APP_ROOT / 'Cache' / 'QuarterlyDefectReview'
        path = root / item['id'][4:]
    else:
        root = EXPORT_ROOT
        path = root / item['id'][4:]
    resolved = path.resolve()
    if not resolved.is_relative_to(root.resolve()):
        raise ValueError('Đường dẫn đầu vào không hợp lệ.')
    return resolved


def report_cell(tool, token, value=None, *, delete_row=False, row_tokens=()):
    with LOCK:
        if (value is not None or delete_row) and any(run['status'] in ('queued', 'running') for run in runs(100)):
            raise RuntimeError('Đang chạy báo cáo. Chờ hoàn tất rồi lưu ô.')
        return edit_cell(tool, token, value, APP_ROOT / 'Report' / 'InputBackups',
                         file_hash, write_json, now, delete_row=delete_row, row_tokens=row_tokens)


def _validate_qdr_staged_file(destination, group, date_from=None, date_to=None):
    """Validate NIKE source files and JSON catalog cleanly without altering original .xls/.xlsm formats."""
    name = destination.name
    if group == 'sandals_catalog':
        try:
            if destination.suffix.lower() == '.json':
                payload = json.loads(destination.read_text(encoding='utf-8'))
                entries = payload.get('entries', payload) if isinstance(payload, dict) else payload
                rows_cnt = len(entries) if isinstance(entries, list) else 0
            else:
                from report_tools.quarterly_defect_review.nike_inputs import normalize_nike_source
                res = normalize_nike_source([destination], metric='FTT', run_id='preflight')
                rows_cnt = len(res['base_records'])
            return dict(name=name, group=group, valid=True, error_count=0, issues=[], rows=rows_cnt, staged_name=name)
        except Exception as err:
            return dict(
                name=name,
                group=group,
                valid=False,
                error_count=1,
                issues=[dict(sheet='Catalog', row=1, column='', rule='qdr.catalog', message=str(err))],
                rows=0,
                staged_name=name,
            )

    metric = 'HFPA' if group == 'hfpa_nike' else 'FTT'
    try:
        from report_tools.quarterly_defect_review.nike_inputs import normalize_nike_source
        res = normalize_nike_source([destination], metric=metric, run_id='preflight', date_from=date_from, date_to=date_to)
        issues = []
        if res.get('has_blocking_mismatch'):
            for r in res.get('review_items', []):
                if r.blocking:
                    issues.append(dict(sheet=r.factory or metric, row=1, column='', rule=r.review_type, message=r.reason))
        return dict(
            name=name,
            group=group,
            valid=len(issues) == 0,
            error_count=len(issues),
            issues=issues,
            rows=sum(m.kept_rows for m in res['manifests']),
            staged_name=name,
        )
    except Exception as err:
        return dict(
            name=name,
            group=group,
            valid=False,
            error_count=1,
            issues=[dict(sheet=metric, row=1, column='', rule=f'qdr.{group}', message=str(err))],
            rows=0,
            staged_name=name,
        )


def start_run(tool, period, mode, selections, uploads=(), acquire_missing=False, validate_only=False, cutoff=None, date_from=None, date_to=None, mapping_overrides=None):
    started = time.perf_counter()
    if tool not in TOOLS or mode not in TOOLS[tool]['modes']:
        raise ValueError('Tool hoặc chế độ xử lý không hợp lệ.')
    is_qdr = (tool == 'quarterly_defect_review')
    eff_from, eff_to = date_from, date_to
    if TOOLS[tool].get('period_kind') == 'fiscal_quarter':
        if not re.fullmatch(r'FY\d{2}-Q[1-4]', period or ''):
            raise ValueError('Chọn quý tài chính theo định dạng FY27-Q2.')
        if tool == 'quarterly_quality':
            from report_tools.quarterly_quality.template_output import verify_template_quarter
            verify_template_quarter(period)
        elif is_qdr:
            from report_tools.quarterly_defect_review.report import resolve_period_and_dates
            _, eff_from, eff_to = resolve_period_and_dates(period, date_from, date_to, cutoff)
    elif not re.fullmatch(r'20\d{2}-(0[1-9]|1[0-2])', period or ''):
        raise ValueError('Chọn tháng báo cáo theo định dạng YYYY-MM.')
    upload_only = TOOLS[tool].get('upload_only', False)
    if upload_only and (selections or (not acquire_missing and len(uploads) != 1)):
        raise ValueError('Chọn đúng một workbook Excel tổng hợp từ máy.')
    with LOCK:
        active = [r for r in runs(100) if r['status'] == 'running' or
                  (r['status'] == 'queued' and (datetime.fromisoformat(now()) - datetime.fromisoformat(r['started_at'])).total_seconds() < 60)]
        if active:
            raise RuntimeError('Một báo cáo đang xử lý. Vui lòng chờ hoàn tất, không cần gửi lại.')
        acquisition = None
        if acquire_missing:
            if uploads:
                raise ValueError('Luồng tự tải SFC không trộn với file tải lên.')
            if tool == 'quarterly_quality':
                from sfc_quarterly_quality_acquisition import plan
                acquisition = plan(period, cutoff)
            elif is_qdr:
                from sfc_quarterly_defect_review_acquisition import plan
                acquisition = plan(period, cutoff, date_from=eff_from, date_to=eff_to)
            elif tool == 'reinspection':
                from sfc_reinspection_acquisition import plan
            elif tool == 'hfpa':
                from sfc_hfpa_acquisition import plan
            else:
                from sfc_copq_acquisition import plan
            if acquisition is None:
                acquisition = plan(period, mode)
            selections = [dict(id=item['id']) for item in acquisition['available']]
        candidates = acquisition['available'] if acquisition is not None else (input_catalog(tool) if selections else [])
        catalog = {item['id']:item for item in candidates}
        chosen, identities = [], set()
        for selection in selections:
            key = selection.get('id')
            if key not in catalog:
                raise ValueError('File đầu vào không còn tồn tại. Hãy tải lại danh sách.')
            item = catalog[key]
            group = selection.get('group', item['group'])
            if group not in TOOLS[tool]['groups']:
                raise ValueError('Nhóm dữ liệu đầu vào không hợp lệ.')
            identity = (group, item['name'])
            if identity in identities:
                raise ValueError('Trùng tên file trong nhóm đầu vào: ' + item['name'])
            identities.add(identity)
            if not is_qdr and item['period'] not in (None, period):
                raise ValueError('File không thuộc tháng đã chọn: ' + item['name'])
            source = input_path(tool, item)
            chosen.append((group, item['name'], source))
        for group, storage in uploads:
            name = safe_name(storage.filename)
            if group not in TOOLS[tool]['groups'] or Path(name).suffix.lower() not in input_formats(tool, group):
                raise ValueError('File tải lên phải thuộc đúng định dạng và nhóm dữ liệu.')
            if (group, name) in identities:
                raise ValueError('Trùng tên file tải lên: ' + name)
            identities.add((group, name))
            if not is_qdr and not upload_only and period_of(name) not in (None, period):
                raise ValueError('File không thuộc tháng đã chọn: ' + name)
            chosen.append((group, name, storage))
        required = TOOLS[tool]['required_groups'][mode]
        for group in required:
            if not acquire_missing and not any(item[0] == group for item in chosen):
                raise ValueError('Thiếu file đầu vào: ' + TOOLS[tool]['groups'][group][0])
        if is_qdr and not acquire_missing and not any(item[0] in ('hfpa_nike', 'ftt_nike') for item in chosen):
            raise ValueError('Chọn ít nhất một nguồn HFPA NIKE hoặc FTT NIKE để rà soát.')
        if not is_qdr:
            scope = set()
            for group, name, _ in chosen:
                identity = (group, plant_of(name))
                if identity in scope:
                    raise ValueError('Chỉ chọn một file cho mỗi nhà máy trong cùng nhóm và tháng, tránh cộng trùng dữ liệu: ' + name)
                scope.add(identity)
        run_id = uuid4().hex
        temporary = tempfile.TemporaryDirectory(prefix='validation-') if validate_only else None
        folder = Path(temporary.name) if temporary else run_folder(run_id)
        workspace = folder / 'workspace'
        if validate_only:
            workspace.mkdir(parents=True)
        else:
            shutil.copytree(TOOLS_ROOT / tool, workspace)
        staged, validation = [], []
        cached_files = 0
        try:
            for group, name, source in chosen:
                name = safe_name(name)
                if not is_qdr and not upload_only and plant_of(name) not in TOOLS[tool]['plants']:
                    raise ValueError('Tên file cần có mã nhà máy được tool hỗ trợ: ' + name)
                destination = workspace / TOOLS[tool]['groups'][group][1] / name
                if not is_qdr and any(item.get('staged_name') == destination.with_suffix('.xlsx').name and item['group'] == group for item in validation):
                    raise ValueError('Các file có cùng tên sau chuẩn hóa. Đổi tên file đầu vào.')
                destination.parent.mkdir(parents=True, exist_ok=True)
                if isinstance(source, Path):
                    if source.stat().st_size > CONTRACTS['max_file_bytes']:
                        raise ValueError('File vượt giới hạn 50 MB: ' + name)
                    before = source.stat()
                    shutil.copyfile(source, destination)
                    after = source.stat()
                    if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
                        raise ValueError('File đầu vào thay đổi khi đang sao chép; hãy lưu/đóng file và thử lại: ' + name)
                else:
                    with destination.open('wb') as stream:
                        count = 0
                        while block := source.stream.read(1024 * 1024):
                            count += len(block)
                            if count > CONTRACTS['max_file_bytes']:
                                raise ValueError('File vượt giới hạn 50 MB: ' + name)
                            stream.write(block)
                mime = None if isinstance(source, Path) else source.mimetype
                if is_qdr:
                    report = _validate_qdr_staged_file(destination, group, date_from=eff_from, date_to=eff_to)
                    source_hash = None if validate_only else file_hash(destination)
                elif validate_only:
                    report, cached = validate_preflight(destination, tool, group, period, mime=mime)
                    cached_files += int(cached)
                    # Detect the same canonical-name collisions as execution without writing Excel.
                    report['staged_name'] = destination.with_suffix('.xlsx').name
                    source_hash = None
                    if isinstance(source, Path) and not report['valid']:
                        with destination.open('rb') as stream:
                            signature = stream.read(8)
                        if signature.startswith(b'PK') or signature == OLE:
                            revision = file_hash(destination)
                            for issue in report['issues']:
                                if issue.get('cell_type'):
                                    issue['edit_token'] = edit_token(
                                        tool, group, period, source, revision, issue)
                else:
                    source_hash = file_hash(destination)
                    if tool == 'reinspection' and group == 'ftt':
                        report, _ = validate_preflight(destination, tool, group, period, mime=mime, normalize=True)
                    else:
                        report = validate_file(destination, tool, group, period, mime=mime, normalize=True)
                validation.append(report)
                staged_path = destination if destination.exists() else destination.with_suffix('.xlsx')
                staged.append(dict(group=group, name=name, size=staged_path.stat().st_size, sha256=source_hash))
                staged[-1]['staged_name'] = report.get('staged_name', destination.name)
                if tool == 'reinspection' and not validate_only and report['valid']:
                    staged[-1]['validated_sha256'] = file_hash(staged_path)
            if tool == 'reinspection':
                validate_reinspection_coverage(validation, plant_of=plant_of)
            if any(not report['valid'] for report in validation):
                raise PreflightError(validation)
            if tool in ('hfpa', 'reinspection') and not acquire_missing:
                for group in required:
                    present = {plant_of(item['name']) for item in staged if item['group'] == group}
                    missing = set(COMMON_PLANTS) - present
                    if missing:
                        raise ValueError(TOOLS[tool]['groups'][group][0] + ' còn thiếu nhà máy: ' + ', '.join(sorted(missing)))
            if validate_only:
                return dict(valid=True, validation=validation, will_download=acquisition['missing'] if acquisition else [],
                            cached_files=cached_files, checked_files=len(validation),
                            validation_time_ms=round((time.perf_counter() - started) * 1000))
            run = dict(id=run_id, tool=tool, title=TOOLS[tool]['title'], period=period, mode=mode,
                       status='queued', stage='Đã nhận dữ liệu, đang khởi động...', started_at=now(), inputs=staged, files=[],
                       monitor_version=1, validation=validation, steps=[dict(step,status='pending') for step in TOOLS[tool].get('steps',{}).get(mode,[])])
            if is_qdr:
                run.update(date_from=eff_from, date_to=eff_to, cutoff=cutoff or '', mapping_overrides=mapping_overrides)
            if acquisition is not None:
                run['steps'].insert(0,dict(id='acquisition',title='Thu thập dữ liệu SFC',status='pending'))
                run.update(acquire_missing=True, phase='acquisition' if acquisition['missing'] else 'processing',
                    acquisition=dict(date_from=acquisition['date_from'],date_to=acquisition['date_to'],
                        total=acquisition['total'],workflow=acquisition['workflow'],missing=acquisition['missing'],
                        reused=[dict(group=i['group'],plant=i['plant'],name=i['name']) for i in acquisition['available']],downloaded=[]))
                if acquisition.get('tableau'):
                    run['acquisition']['tableau'] = acquisition['tableau']
            write_json(folder / 'run.json', run)
            with (folder / 'run.log').open('ab') as log:
                WORKERS[run_id] = subprocess.Popen([sys.executable, '-u', str(APP_ROOT / 'report_tool_worker.py'), str(folder)],
                    cwd=workspace, stdin=subprocess.DEVNULL, stdout=log, stderr=subprocess.STDOUT,
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            return run
        except Exception as error:
            write_json(folder / 'run.json', dict(id=run_id, tool=tool, title=TOOLS[tool]['title'], period=period,
                mode=mode, status='error', stage='Chưa chạy xử lý', started_at=now(), finished_at=now(),
                monitor_version=1, inputs=staged, files=[], validation=validation, error=str(error) if isinstance(error, ValueError) else friendly_error(error)))
            raise
        finally:
            if temporary:
                temporary.cleanup()


def collect_files(folder):
    folder = Path(folder)
    output = folder / 'workspace' / 'Output'
    files = []
    for path in sorted(output.rglob('*')) if output.exists() else []:
        if not path.is_file() or path.suffix.lower() not in DELIVERABLES or path.name.startswith(('~$', '.')) or not path.resolve().is_relative_to(output.resolve()):
            continue
        if path.suffix.lower() in ('.xlsx', '.pptx'):
            try:
                if not zipfile.is_zipfile(path):
                    continue
                with zipfile.ZipFile(path) as package:
                    if package.testzip() is not None:
                        continue
                    required = 'xl/workbook.xml' if path.suffix.lower() == '.xlsx' else 'ppt/presentation.xml'
                    if required not in package.namelist():
                        continue
                if path.suffix.lower() == '.xlsx':
                    import openpyxl
                    book = openpyxl.load_workbook(path,read_only=True,data_only=False,keep_links=False)
                    book.close()
                else:
                    from pptx import Presentation
                    Presentation(path)
            except Exception:
                continue
        relative = path.relative_to(output).as_posix()
        files.append(dict(name=relative, size=path.stat().st_size, sha256=file_hash(path), modified_at=datetime.fromtimestamp(path.stat().st_mtime, timezone.utc).isoformat(),
            url=f'/api/report-tools/runs/{folder.name}/files/' + quote(relative, safe='/')))
    return files


def make_bundle(folder, files):
    output = folder / 'workspace' / 'Output'
    target = output / 'report_bundle.zip'
    with zipfile.ZipFile(target, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        for item in files:
            path = output / item['name']
            archive.write(path, 'Output/' + item['name'])
    return dict(name=target.name, size=target.stat().st_size,sha256=file_hash(target),
                url=f'/api/report-tools/runs/{folder.name}/files/{target.name}')
