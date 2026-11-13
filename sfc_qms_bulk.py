"""Persist selected PDFs, support explicit resume, and label ZIPs with missing files."""
from datetime import datetime, timedelta, timezone
import csv
from hashlib import sha256
from io import StringIO
from pathlib import Path
import re
import threading
from uuid import uuid4
from zipfile import ZipFile, ZIP_STORED

import sfc_qms as qms

DOWNLOAD_ROOT = qms.ROOT / 'Report' / 'QMSDocuments' / 'Downloads'
LOCK = threading.Lock()
JOBS = {}
MAX_ARCHIVE_BYTES = 5 * 1024 * 1024 * 1024


def timestamp():
    return datetime.now(timezone(timedelta(hours=7))).isoformat(timespec='seconds')


def job_folder(job_id):
    if not isinstance(job_id, str) or not re.fullmatch(r'[a-f0-9]{32}', job_id):
        raise FileNotFoundError('Không tìm thấy lượt tải PDF.')
    return DOWNLOAD_ROOT / job_id


def save(job):
    job['updated_at'] = timestamp()
    qms.atomic_json(job_folder(job['id']) / 'job.json', job)


def public(job):
    fields = ('id', 'plant', 'source', 'status', 'completed', 'total', 'current', 'error', 'filename', 'bytes', 'updated_at', 'failed', 'missing_count', 'cached')
    result = {key: job[key] for key in fields if key in job}
    if job['status'] in ('done', 'partial'):
        result['url'] = f'/api/qms/bulk/{job["id"]}/zip'
    result['resumable'] = job['status'] in ('partial', 'error') and bool(job.get('source'))
    return result


def snapshot(job_id):
    with LOCK:
        job = JOBS.get(job_id)
        if job is None:
            job = qms.read_json(job_folder(job_id) / 'job.json')
            if not job:
                raise FileNotFoundError('Không tìm thấy lượt tải PDF.')
            if job['status'] in ('queued', 'running'):
                job.update(status='error', error='Server đã khởi động lại trước khi tải đủ file. Bấm Tiếp tục để dùng lại các PDF đã lưu.')
                save(job)
        return public(job)


def start(plant, ids, source=None):
    if not isinstance(source, str) or source not in qms.SOURCES:
        raise ValueError('Chọn riêng QA Manual hoặc QA Document để tải tất cả.')
    if (not isinstance(ids, list) or not ids or len(ids) > 5000
            or any(not isinstance(key, str) or not re.fullmatch(r'[a-f0-9]{32}', key) for key in ids)):
        raise ValueError('Chọn từ 1 đến 5.000 PDF hợp lệ để tải.')
    catalog = qms.read_json(qms.CATALOG).get(plant, {})
    available = {item['id']: item for item in catalog.get('documents', [])}
    ids = list(dict.fromkeys(ids))
    if any(key not in available for key in ids):
        raise ValueError('Danh mục PDF đã thay đổi. Cập nhật lại danh sách trước khi tải tất cả.')
    if any(available[key]['source'] != source for key in ids):
        raise ValueError('Chỉ tải PDF thuộc phần đã chọn. QA Manual và QA Document phải tải riêng.')
    with LOCK:
        if any(job['status'] in ('queued', 'running') for job in JOBS.values()):
            raise RuntimeError('Đang tải một bộ PDF. Chờ hoàn tất rồi tạo lượt mới.')
        job_id = uuid4().hex
        source_name = 'QA_Manual' if source == 'manual' else 'QA_Document'
        job = dict(id=job_id, plant=plant, source=source, ids=ids, total=len(ids), completed=0, status='queued',
                   current='', error=None, filename=f'{source_name}_{plant}_{job_id[:8]}.zip', bytes=0)
        save(job)
        JOBS[job_id] = job
        try:
            threading.Thread(target=work, args=(job_id, available), daemon=True, name='qms-pdf-' + job_id[:8]).start()
        except Exception:
            job.update(status='error', error='Chưa khởi động được lượt tải PDF.')
            save(job)
            raise
        return public(job)


def resume(job_id):
    folder = job_folder(job_id)
    with LOCK:
        if any(job['status'] in ('queued', 'running') for job in JOBS.values()):
            raise RuntimeError('Đang tải một bộ PDF. Chờ hoàn tất rồi tiếp tục.')
        job = JOBS.get(job_id) or qms.read_json(folder / 'job.json')
        if not job or job.get('status') not in ('partial', 'error') or job.get('source') not in qms.SOURCES:
            raise ValueError('Lượt tải này không cần hoặc chưa thể tiếp tục.')
        available = {item['id']: item for item in qms.read_json(qms.CATALOG).get(job['plant'], {}).get('documents', [])}
        if any(key not in available or available[key]['source'] != job['source'] for key in job['ids']):
            raise ValueError('Danh mục đã thay đổi. Cập nhật QMS và chọn lại PDF cần tải.')
        job.update(status='queued', error=None)
        save(job)
        JOBS[job_id] = job
        threading.Thread(target=work, args=(job_id, available), daemon=True, name='qms-resume-' + job_id[:8]).start()
        return public(job)


def update(job_id, **fields):
    with LOCK:
        job = JOBS[job_id]
        job.update(fields)
        save(job)


def work(job_id, available):
    folder = job_folder(job_id)
    pending = folder / 'archive.partial'
    try:
        with LOCK:
            job = dict(JOBS[job_id])
        files = {}
        for key, info in job.get('files', {}).items():
            path = folder / 'pdfs' / (key + '.pdf')
            if key not in job['ids'] or not re.fullmatch(r'[a-f0-9]{32}', key) or not path.is_file():
                continue
            body = path.read_bytes()
            if (len(body) == info['bytes'] and body.startswith(b'%PDF-') and b'%%EOF' in body[-4096:]
                    and sha256(body).hexdigest() == info['sha256']):
                files[key] = info
        failures = []
        total_bytes = sum(info['bytes'] for info in files.values())
        update(job_id, status='running', completed=len(files), files=files, failed=[], cached=len(files),
               missing_count=job['total'] - len(files), error=None, bytes=total_bytes)
        session = qms.DocumentSession('JV3' if job['plant'] == 'JVB' else job['plant'])
        session.prepare_catalog()
        (folder / 'pdfs').mkdir(exist_ok=True)
        fatal = None
        for key in job['ids']:
            if key in files:
                continue
            entry = available[key]
            update(job_id, current=f'{entry.get("code") or entry["name"]} {entry.get("language", "")}'.strip())
            try:
                stream, filename = qms.download_pdf(job['plant'], key, session=session)
                with stream:
                    body = stream.getvalue()
                if not body.startswith(b'%PDF-') or b'%%EOF' not in body[-4096:]:
                    raise RuntimeError('SFC chưa trả về PDF hoàn chỉnh.')
                if total_bytes + len(body) > MAX_ARCHIVE_BYTES:
                    raise RuntimeError('Bộ PDF vượt giới hạn 5 GB. Chia nhỏ bằng bộ lọc rồi tải lại.')
                path = folder / 'pdfs' / (key + '.pdf')
                temporary = path.with_suffix('.partial')
                temporary.write_bytes(body)
                temporary.replace(path)
                files[key] = dict(name=filename, bytes=len(body), sha256=sha256(body).hexdigest())
                total_bytes += len(body)
                del body
                update(job_id, completed=len(files), files=files, bytes=total_bytes)
            except Exception as error:
                detail = qms.download_error(error)
                failures.append(dict(id=key, code=entry.get('code', ''), name=entry['name'], language=entry.get('language', ''), **detail))
                update(job_id, failed=failures, missing_count=job['total'] - len(files), error=detail['message'])
                if detail['kind'] not in ('missing', 'invalid_pdf'):
                    fatal = detail['message']
                    break
        missing = [key for key in job['ids'] if key not in files]
        if not files:
            update(job_id, status='error', missing_count=len(missing), error=fatal or 'SFC không trả về PDF tải được. Xem danh sách lỗi và thử lại sau khi nguồn đã được sửa.')
            return
        update(job_id, current='Đang đóng gói các PDF đã lưu')
        names = set()
        with ZipFile(pending, 'w', compression=ZIP_STORED, allowZip64=True) as archive:
            for key in job['ids']:
                if key not in files:
                    continue
                entry, info = available[key], files[key]
                filename = info['name']
                filename = Path(filename.replace('\\', '/')).name
                source_folder = qms.SOURCES[entry['source']][0]
                group_folder = {'records': 'SFC', 'hfpa': 'HFPA Guide', 'bdts': 'BDTS'}[qms.collection_of(entry)]
                path = f'{source_folder}/{group_folder}/{filename}'
                if path.casefold() in names:
                    path = f'{source_folder}/{group_folder}/{Path(filename).stem}_{key}.pdf'
                names.add(path.casefold())
                archive.write(folder / 'pdfs' / (key + '.pdf'), path)
            if missing:
                report = StringIO(newline='')
                writer = csv.writer(report)
                writer.writerow(['Mã', 'Tên tài liệu', 'Ngôn ngữ', 'HTTP', 'Lý do chưa tải được'])
                by_key = {item['id']: item for item in failures}
                for key in missing:
                    entry = available[key]
                    issue = by_key.get(key, {})
                    writer.writerow([entry.get('code', ''), entry['name'], entry.get('language', ''),
                                     issue.get('http_status', ''), issue.get('message', 'Chưa tải do lượt bị gián đoạn; bấm Tiếp tục để lấy phần còn lại.')])
                archive.writestr('PDF_CHUA_TAI_DUOC.csv', report.getvalue().encode('utf-8-sig'))
        with ZipFile(pending) as archive:
            if sum(item.filename.lower().endswith('.pdf') for item in archive.infolist()) != len(files):
                raise RuntimeError('Số PDF trong ZIP không khớp số file đã lưu.')
        source_name = 'QA_Manual' if job.get('source') == 'manual' else 'QA_Document'
        suffix = f'_THIEU_{len(missing)}_PDF' if missing else ''
        filename = f'{source_name}_{job["plant"]}_{job_id[:8]}{suffix}.zip'
        destination = folder / filename
        pending.replace(destination)
        update(job_id, status='partial' if missing else 'done', current='', filename=filename,
               missing_count=len(missing), bytes=destination.stat().st_size,
               error=(fatal or 'Một số PDF gốc không tải được. ZIP có danh sách chi tiết PDF_CHUA_TAI_DUOC.csv.') if missing else None)
    except Exception as error:
        pending.unlink(missing_ok=True)
        update(job_id, status='error', error=qms.download_error(error)['message'])


def archive_file(job_id):
    job = snapshot(job_id)
    if job['status'] not in ('done', 'partial'):
        raise RuntimeError('ZIP chưa sẵn sàng. PDF đã tải vẫn được giữ để tiếp tục.')
    path = job_folder(job_id) / job['filename']
    if not path.is_file():
        raise FileNotFoundError('File ZIP không còn trên máy. Tạo lượt tải mới.')
    return path
