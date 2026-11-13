"""Background work and progress for multi-gigabyte local combine jobs."""
from contextlib import ExitStack
from pathlib import Path
import json
import re
import shutil
import threading
import time
from uuid import uuid4

from werkzeug.datastructures import FileStorage

TASKS = {}
LOCK = threading.Lock()


def remove_incoming(folder):
    import spreadsheet_combine as combine
    target=folder.resolve()
    expected=(combine.ROOT/'_incoming').resolve()
    if target.parent!=expected or not re.fullmatch('[a-f0-9]{32}',target.name):
        raise ValueError('Không dọn thư mục ngoài vùng tải lên tạm.')
    shutil.rmtree(target)


def start(work):
    task_id = uuid4().hex
    with LOCK:
        if any(task['status']=='running' for task in TASKS.values()):
            raise ValueError('Đang xử lý một lượt gộp lớn. Chờ hoàn tất rồi thử lại.')
        TASKS[task_id] = dict(status='running',message='Đang chuẩn bị xử lý…',updated=time.time())
    def update(message):
        with LOCK:
            TASKS[task_id].update(message=message,updated=time.time())
    def run():
        try:
            result=work(update)
            with LOCK:
                TASKS[task_id].update(status='done',result=result,updated=time.time())
        except Exception as error:
            import logging
            logging.getLogger(__name__).exception('Combine background task failed')
            with LOCK:
                TASKS[task_id].update(status='error',error=str(error),errors=getattr(error,'errors',[]),updated=time.time())
    threading.Thread(target=run,daemon=True,name='combine-'+task_id[:8]).start()
    return dict(task_id=task_id,task_url=f'/api/combine/tasks/{task_id}')


def snapshot(task_id):
    with LOCK:
        if task_id not in TASKS:
            raise ValueError('Lượt xử lý không còn tồn tại. Hãy tải lại file (ứng dụng có thể đã khởi động lại).')
        return dict(TASKS[task_id])


def active():
    with LOCK:
        return any(task['status']=='running' for task in TASKS.values())


def stage_upload(files,csv_options):
    import spreadsheet_combine as combine
    if active():
        raise ValueError('Đang xử lý một lượt gộp lớn. Chờ hoàn tất rồi tải lượt mới.')
    if not 1 <= len(files) <= combine.MAX_FILES:
        raise ValueError(f'Chọn từ 1 đến {combine.MAX_FILES} file Excel/CSV mỗi lần.')
    csv_options=combine.csv_input.options(csv_options)
    folder=combine.ROOT / '_incoming' / uuid4().hex
    folder.mkdir(parents=True)
    staged=[];total=0
    try:
        for index,storage in enumerate(files):
            name=str(storage.filename or '').replace('\\','/').rsplit('/',1)[-1]
            extension=Path(name).suffix.lower()
            if extension not in combine.SUPPORTED_EXTENSIONS:
                raise combine.UploadValidationError([dict(file=name,error='Chỉ hỗ trợ Excel và CSV.')])
            target=folder / f'{index}{extension}'
            size=0
            with target.open('wb') as stream:
                while chunk:=storage.stream.read(1024*1024):
                    size+=len(chunk);total+=len(chunk)
                    if size>combine.MAX_BYTES or total>combine.MAX_TOTAL_BYTES:
                        raise combine.UploadValidationError([dict(file=name,error='Vượt giới hạn 200 MB/file hoặc tổng 5 GB/lượt.')])
                    stream.write(chunk)
            staged.append((name,target))
    except Exception:
        remove_incoming(folder)
        raise
    def work(progress):
        try:
            with ExitStack() as stack:
                sources=[FileStorage(stack.enter_context(path.open('rb')),filename=name) for name,path in staged]
                return combine.upload(sources,csv_options,progress)
        finally:
            # This folder is generated here and contains only this upload.
            remove_incoming(folder)
    try:
        return start(work)
    except Exception:
        remove_incoming(folder)
        raise


def is_large(job_id):
    import spreadsheet_combine as combine
    folder=combine.folder_for(job_id)
    return any(item.get('large') for item in json.loads((folder/'inputs.json').read_text(encoding='utf-8')))
