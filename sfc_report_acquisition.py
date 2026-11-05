"""Shared single-pass SFC acquisition, durable caching, and browser restoration."""
from pathlib import Path
import shutil
from uuid import uuid4

import sfc_report_tools as reports
from report_runtime import Cancelled, friendly_error, diagnostics


def acquire_month(folder, run, stage, controller=None, *, tool):
    """Only fetch missing slots; completed downloads are reusable after a failure."""
    import importlib
    profile = importlib.import_module('sfc_' + tool + '_acquisition')
    if controller is None:
        import sfc_controller as controller
    folder = Path(folder)
    workspace = folder / 'workspace'
    initial = None
    failures = []
    disconnected = False
    window = run['acquisition']
    def check_cancel():
        if (folder / 'cancel.request').exists():
            raise Cancelled()
    check_cancel()
    missing = window['missing']
    if not missing:
        stage(f'Đã có đủ {window["total"]} file đúng tháng; không tải lại từ SFC.')
        return
    stage(f'Đang kết nối SFC để tải {len(missing)} file còn thiếu...')
    if getattr(controller, '__name__', '') == 'sfc_controller':
        from dashboard_launcher import dom_ready, start_dom
        if not dom_ready():
            stage('Đang mở phiên Chrome SFC; đăng nhập trên cửa sổ SFC nếu được yêu cầu...')
            start_dom()
    try:
        with controller.sync_playwright() as client:
            browser = client.chromium.connect_over_cdp(controller.CDP_URL, timeout=controller.SFC_CONNECT_WAIT_MS)
            try:
                page = controller.get_sfc_page(browser)
                controller.sfc_downloads.wait_idle(page, controller.SFC_WAIT_MS)
                plant = next((c['value'] for c in page.context.cookies() if c['name']=='MyFactoryName'), None)
                initial = dict(url=page.url, plant=controller.sfc_downloads.factory_code(plant),
                               filters=controller.sfc_downloads.capture_filters(page))
                if initial['plant'] not in controller.SEVEN_FACTORIES:
                    raise ValueError('Chưa xác nhận được nhà máy của phiên SFC. Hãy đăng nhập lại rồi chạy báo cáo.')
            finally:
                browser.close()
        for index, slot in enumerate(missing):
            check_cancel()
            group, plant = slot['group'], slot['plant']
            slot_window = {**window, **slot}
            first, last = slot_window['date_from'], slot_window['date_to']
            stage(f'[{index+1}/{len(missing)}] Đang tải {group.upper()} · {plant} · {first} đến {last}')
            try:
                switched = controller.switch_factory(plant)
                if not switched.get('success'):
                    raise RuntimeError(switched.get('error'))
                result = profile.download(controller, group, plant, slot_window)
                if result.get('timed_out'):
                    raise TimeoutError(result.get('error') or 'SFC export timed out.')
                if not result.get('success') or not result.get('downloaded_file'):
                    raise RuntimeError(result.get('error') or 'SFC chưa trả về file Excel.')
                source = Path(result['downloaded_file']['path'])
                if not profile.validate_file(source, group, first, last):
                    raise ValueError('SFC trả về sai mẫu dữ liệu hoặc kỳ tháng; chưa đưa vào báo cáo.')
                name = profile.filename(group, plant, slot_window, source.suffix.lower())
                if source.name != name:
                    cache = reports.EXPORT_ROOT / ('acquired_'+uuid4().hex)
                    cache.mkdir(parents=True)
                    shutil.copyfile(source,cache/name)
                    source = cache/name
                proof = dict(group=group,plant=plant,date_from=first,date_to=last,
                             **window['workflow'][group])
                reports.write_json(source.parent/'acquisition.json',proof)
                destination = workspace / reports.TOOLS[tool]['groups'][group][1] / name
                if tool != 'quarterly_quality' and destination.suffix=='.xlsm':
                    destination=destination.with_suffix('.xlsx')
                destination.parent.mkdir(parents=True,exist_ok=True)
                shutil.copyfile(source,destination)
                record = dict(group=group,plant=plant,date_from=first,date_to=last,name=name)
                run['inputs'].append(dict(record,size=destination.stat().st_size,
                    sha256=reports.file_hash(destination),acquired_from_sfc=True))
                window.setdefault('downloaded',[]).append(record)
                reports.write_json(folder/'run.json',run)
                check_cancel()
            except Cancelled:
                raise
            except Exception as error:
                failures.append(dict(group=group,plant=plant,error=friendly_error(error),diagnostics=diagnostics(error)))
                window['errors'] = failures
                reports.write_json(folder/'run.json',run)
                if isinstance(error, TimeoutError) or any(marker in str(error).lower() for marker in ('timeout','econnrefused','closed','đăng nhập','login')):
                    disconnected = True
                    window['not_attempted'] = missing[index+1:]
                    reports.write_json(folder/'run.json',run)
                    break  # Do not replay a timed-out request or keep reconnecting to a closed browser.
    finally:
        if initial and not disconnected and initial['plant'] in controller.SEVEN_FACTORIES:
            restored = controller.switch_factory(initial['plant'])
            if restored.get('success'):
                controller._execute_dynamic_action('restore_filters',params=dict(_page_url=initial['url'],_filters=initial['filters']))
    if failures:
        raise RuntimeError('Chưa đủ dữ liệu; không chạy báo cáo thiếu nhà máy. ' + '; '.join(
            f'{item["group"]}/{item["plant"]}: {item["error"]}' for item in failures))
