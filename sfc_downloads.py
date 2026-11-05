"""Capture completed SFC exports, preserving their workbook format and scope."""
from pathlib import Path
import json
import re
from uuid import uuid4
import zipfile

from sfc_excel import EXPORT_ROOT, file_info

EXCEL_EXTENSIONS = ('.xls', '.xlsx', '.xlsm', '.xlsb')


class SFCServerTimeout(RuntimeError):
    """A completed server response confirms its SQL query has timed out."""


def factory_code(value):
    return 'JVB' if value == 'JV3' else value


def require_attachment(response):
    headers = response.headers
    if response.status < 400 and (headers.get('content-disposition') or any(
            mime in headers.get('content-type', '').lower() for mime in ('excel', 'spreadsheet', 'octet-stream'))):
        return
    content = response.text()
    if re.search(r'execution\s+timeout\s+expired|sql(?:exception)?.{0,100}timeout\s+expired', content, re.I):
        raise SFCServerTimeout('SFC báo Execution Timeout Expired; truy vấn đã kết thúc với lỗi timeout.')
    if response.status >= 400:
        raise RuntimeError(f'SFC trả lỗi HTTP {response.status}; chưa xuất được Excel.')
    messages = re.findall(r"OnShowMyMessage\('([^']*)'", content)
    raise RuntimeError(messages[-1] if messages else 'SFC trả về trang HTML thay vì file Excel; chưa xuất được dữ liệu.')


def wait_idle(page, timeout):
    page.wait_for_function("""() => {
        if (typeof aspxGetControlCollection !== 'function') return false;
        let busy = false;
        aspxGetControlCollection().ForEachControl(c => {
            if (typeof c.InCallback === 'function' && c.InCallback()) busy = true;
        });
        return !busy;
    }""", timeout=timeout)


def set_copq_combobox(page, control_id, value, timeout):
    """These two SFC controls use server postbacks, not client-only changes."""
    value = None if value in (None, '', 'ALL') else value
    wait_idle(page, timeout)
    valid = value is None or page.evaluate("({id, value}) => !!aspxGetControlCollection().Get(id).FindItemByValue(value)",
                                          dict(id=control_id, value=value))
    if not valid:
        raise ValueError('Chọn Category / Type hợp lệ trên trang COPQ Data.')
    current = page.evaluate("id => aspxGetControlCollection().Get(id).GetValue()", control_id)
    if current == value:
        return
    with page.expect_response(lambda r: r.request.method == 'POST' and 'QAReportCOPQ.aspx' in r.url,
                              timeout=timeout) as pending:
        page.evaluate("""({id, value}) => {
            const c = aspxGetControlCollection().Get(id);
            if (!c || (value !== null && !c.FindItemByValue(value))) throw new Error('Lựa chọn COPQ không hợp lệ: ' + value);
            if (value === null) {
                c.SetSelectedIndex(-1);
                c.SetValue(null);
                c.SetText('');
            } else c.SetValue(value);
            __doPostBack(c.uniqueID, '');
        }""", dict(id=control_id, value=value))
    response = pending.value
    if response.status >= 400:
        raise RuntimeError('SFC không cập nhật được bộ lọc COPQ.')
    page.wait_for_load_state('domcontentloaded', timeout=timeout)
    wait_idle(page, timeout)
    actual = page.evaluate("id => aspxGetControlCollection().Get(id).GetValue()", control_id)
    if actual != value:
        raise ValueError('SFC chưa xác nhận lựa chọn COPQ: ' + control_id)


def copq_search(page, timeout):
    wait_idle(page, timeout)
    valid = page.evaluate("""() => ['cboCategory', 'cboType'].every(suffix => {
        const c = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_' + suffix);
        return c && (c.GetValue() == null || c.GetValue() === '' || c.FindItemByValue(c.GetValue()));
    })""")
    if not valid:
        raise ValueError('Category / Type không còn hợp lệ; hãy chọn lại hoặc xóa lựa chọn.')
    with page.expect_response(lambda r: r.request.method == 'POST' and 'QAReportCOPQ.aspx' in r.url,
                              timeout=timeout) as pending:
        page.locator('#rpnlMain_ContentPlaceHolder1_btnSearch').click(no_wait_after=True)
    if pending.value.status >= 400:
        raise RuntimeError('SFC không tải được dữ liệu COPQ.')
    page.wait_for_load_state('domcontentloaded', timeout=timeout)
    wait_idle(page, timeout)


def export_copq_data(page, timeout, plant, format_name):
    copq_search(page, timeout)
    empty = page.evaluate("""() => {
        const grid = aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_grdBasicSet');
        return grid && typeof grid.GetVisibleRowsOnPage === 'function' && grid.GetVisibleRowsOnPage() === 0;
    }""")
    received = []
    handler = lambda download: received.append(download)
    page.on('download', handler)
    try:
        with page.expect_response(lambda r: r.request.method == 'POST' and 'QAReportCOPQ.aspx' in r.url,
                                  timeout=timeout) as pending:
            page.locator('#rpnlMain_ContentPlaceHolder1_btnExcel').click(no_wait_after=True)
        require_attachment(pending.value)
        download = received[0] if received else page.wait_for_event('download', timeout=timeout)
        result = save_download(download, plant, format_name)
        if empty is True:
            result['rows'] = 0
        return result
    finally:
        page.remove_listener('download', handler)


def export_reinspection_data(page, timeout, plant, format_name, params):
    """Clear Line/Type, verify postback filters, then capture the real workbook."""
    if 'qcrechk.aspx' not in page.url.lower():
        raise ValueError('Hãy mở trang Re-Inspection QCRECHK trước khi xuất.')
    prefix = 'rpnlMain_ContentPlaceHolder1_'
    first,last = params.get('date_from'),params.get('date_to')
    if not first or not last:
        raise ValueError('Cần đủ ngày bắt đầu và kết thúc Re-Inspection.')
    expected = {
        prefix+'dtDateFrom':dict(kind='date',value=first),
        prefix+'dtDateTo':dict(kind='date',value=last),
        prefix+'cboDeptFrom':dict(kind='combo',value=None),
        prefix+'cboType':dict(kind='combo',value=None)}
    apply_filters(page,expected,timeout)
    page.evaluate("""() => {
        const coll = aspxGetControlCollection();
        for (const suffix of ['cboDeptFrom','cboType']) {
            const c = coll.Get('rpnlMain_ContentPlaceHolder1_' + suffix);
            if (!c) throw new Error('Không tìm thấy bộ lọc Re-Inspection: ' + suffix);
            c.SetSelectedIndex(-1); c.SetValue(null); c.SetText('');
        }
    }""")

    def verify():
        current = capture_filters(page)
        if any(current.get(key) != value for key,value in expected.items()):
            raise ValueError('SFC chưa xác nhận đúng ngày hoặc chưa xóa Line/Type Re-Inspection.')
        if not page.evaluate("""() => ['cboDeptFrom','cboType'].every(suffix =>
            aspxGetControlCollection().Get('rpnlMain_ContentPlaceHolder1_' + suffix).GetText() === '')"""):
            raise ValueError('Line/Type Re-Inspection vẫn còn lựa chọn.')
        actual = next((c['value'] for c in page.context.cookies() if c['name']=='MyFactoryName'),None)
        if factory_code(actual) != plant:
            raise ValueError('Nhà máy SFC thay đổi khi tìm kiếm; dừng xuất Re-Inspection.')

    verify()
    with page.expect_response(lambda r: r.request.method == 'POST' and 'qcrechk.aspx' in r.url.lower(),
                              timeout=timeout) as pending:
        page.locator('#'+prefix+'btnSearch').click(no_wait_after=True)
    response = pending.value
    response.finished()
    if response.status >= 400:
        raise RuntimeError('SFC không tìm kiếm được dữ liệu Re-Inspection.')
    page.wait_for_load_state('domcontentloaded',timeout=timeout)
    # Observe EndCallback as well as InCallback before reading native filters.
    page.evaluate("""timeout => new Promise((resolve,reject) => {
        const handlers=[]; let timer,poll;
        const cleanup=()=>{clearTimeout(timer);clearInterval(poll);
            handlers.forEach(([event,handler])=>event.RemoveHandler(handler));};
        const check=()=>{let busy=false;aspxGetControlCollection().ForEachControl(c=>{
            if(typeof c.InCallback==='function' && c.InCallback()) busy=true;
        });if(!busy){cleanup();resolve(true);}};
        aspxGetControlCollection().ForEachControl(c=>{
            if(c.EndCallback?.AddHandler){const handler=()=>setTimeout(check,0);
                c.EndCallback.AddHandler(handler);handlers.push([c.EndCallback,handler]);}
        });
        timer=setTimeout(()=>{cleanup();reject(new Error('Re-Inspection callback timeout; chưa xuất lại.'));},timeout);
        poll=setInterval(check,250);check();
    })""",timeout)
    verify()
    received=[]
    handler=lambda download:received.append(download)
    page.on('download',handler)
    try:
        with page.expect_response(lambda r: r.request.method=='POST' and 'qcrechk.aspx' in r.url.lower(),
                                  timeout=timeout) as pending:
            page.locator('#'+prefix+'btnExcel').click(no_wait_after=True)
        require_attachment(pending.value)
        download = received[0] if received else page.wait_for_event('download',timeout=timeout)
        return save_download(download,plant,format_name)
    finally:
        page.remove_listener('download',handler)


def capture_filters(page):
    return page.evaluate("""() => {
        const result = {};
        const mqaa = /MQAAReport\\.aspx/i.test(location.pathname);
        const mqaaCheck = /MQAACheck\\.aspx$/i.test(location.pathname);
        const mqaaControls = ['dtDateFrom','dtDateTo','cboType','cboReportType','cboGroup'];
        aspxGetControlCollection().ForEachControl(c => {
            if (!c.name || (!c.name.startsWith('rpnlMain_ContentPlaceHolder1_') &&
                !(mqaa && mqaaControls.includes(c.name)) && !(mqaaCheck && c.name === 'dtDate')) ||
                /_DDD|popup/i.test(c.name)) return;
            if (typeof c.GetDate === 'function') {
                const d = c.GetDate();
                result[c.name] = {kind:'date', value:d ?
                    [d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-') : null};
            } else if (typeof c.GetChecked === 'function') {
                result[c.name] = {kind:'check', value:c.GetChecked()};
            } else if (typeof c.GetValue === 'function' && typeof c.GetItemCount === 'function') {
                result[c.name] = {kind:'combo', value:c.GetValue()};
            }
        });
        return result;
    }""")


def apply_filters(page, filters, timeout):
    wait_idle(page, timeout)
    if 'QAReportCOPQ.aspx' in page.url:
        for suffix in ('cboCategory', 'cboType'):
            for cid, item in filters.items():
                if cid.endswith(suffix) and item['kind'] == 'combo':
                    set_copq_combobox(page, cid, item['value'], timeout)
    page.evaluate("""filters => {
        const coll = aspxGetControlCollection();
        for (const [id, item] of Object.entries(filters)) {
            const c = coll.Get(id);
            if (!c) throw new Error('Không tìm thấy bộ lọc: ' + id);
            if (item.kind === 'date') {
                const parts = item.value ? item.value.split('-').map(Number) : null;
                c.SetDate(parts ? new Date(parts[0],parts[1]-1,parts[2]) : null);
            } else if (item.kind === 'check') c.SetChecked(item.value);
            else c.SetValue(item.value);
        }
    }""", filters)
    wait_idle(page, timeout)
    actual = capture_filters(page)
    for key, expected in filters.items():
        if actual.get(key) != expected:
            raise ValueError('SFC chưa xác nhận bộ lọc: ' + key)


def save_download(download, plant, format_name):
    """save_as waits for completion; publish only a validated workbook atomically."""
    failure = download.failure()
    if failure:
        raise RuntimeError('SFC tải file thất bại: ' + str(failure))
    # Server names can contain Windows-invalid characters (notably time colons).
    name = download.suggested_filename.replace('\\', '/').rsplit('/', 1)[-1]
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip(' .') or 'SFC_Export'
    EXPORT_ROOT.mkdir(parents=True, exist_ok=True)
    folder = EXPORT_ROOT / ('download_' + uuid4().hex)
    folder.mkdir()
    partial = folder / 'download.partial'
    try:
        download.save_as(str(partial))
        with partial.open('rb') as stream:
            signature = stream.read(8)
        suffix = Path(name).suffix.lower()
        if signature == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
            if suffix != '.xls':
                name = str(Path(name).with_suffix('.xls'))
        elif zipfile.is_zipfile(partial):
            with zipfile.ZipFile(partial) as archive:
                members = set(archive.namelist())
                if '[Content_Types].xml' not in members or not ({'xl/workbook.xml', 'xl/workbook.bin'} & members):
                    raise ValueError('SFC không trả về workbook Excel hợp lệ.')
                detected = '.xlsb' if 'xl/workbook.bin' in members else ('.xlsm' if 'xl/vbaProject.bin' in members else '.xlsx')
                if suffix not in ('.xlsx', '.xlsm', '.xlsb'):
                    name = str(Path(name).with_suffix(detected))
        else:
            raise ValueError('SFC không trả về file Excel hợp lệ (có thể là trang lỗi/đăng nhập).')
        target = folder / format_name(name, plant)
        partial.replace(target)
        return dict(file_info(target), plant=plant)
    finally:
        partial.unlink(missing_ok=True)


def export_selected(controller, action_type, target_id, params, plants):
    """Take the current filters once and apply them to each selected factory."""
    original_scope = list(controller.SELECTED_FACTORIES)
    files, error, last, failures = [], None, {}, []
    folder = EXPORT_ROOT / ('batch_' + uuid4().hex)
    folder.mkdir(parents=True)
    combine = params.get('combine') is True
    manifest = dict(status='running', prompt=('Xuất Excel + Gộp ' if combine else 'Xuất Excel ') + ', '.join(plants),
                    intent=dict(target_plants=plants, combine=combine), monthly_files=files, combined_files=[])

    def save_manifest():
        (folder / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')

    save_manifest()
    with controller.sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(controller.CDP_URL, timeout=controller.SFC_CONNECT_WAIT_MS)
        try:
            page = controller.get_sfc_page(browser)
            from sfc_mqaa_exports import recover_timeout_page
            mqaa_timeout_recovered = recover_timeout_page(page,action_type,target_id,params,controller.SFC_WAIT_MS)
            wait_idle(page, controller.SFC_WAIT_MS)
            url = page.url
            filters = capture_filters(page)
            if mqaa_timeout_recovered:
                # The old SQL error page has lost its controls. The Dashboard
                # still sends the user's chosen dropdowns and reporting dates.
                for key,item in params.get('selections',{}).items():
                    if key in filters and filters[key]['kind']=='combo':
                        value=item.get('value')
                        filters[key]['value']=None if value in ('','ALL',None) else value
            for key, item in filters.items():
                if item['kind'] == 'date':
                    value = params.get('date_from') if 'from' in key.lower() else params.get('date_to')
                    if value:
                        item['value'] = value
            manifest['prompt'] = ('Xuất Excel + Gộp ' if combine else 'Xuất Excel ') + page.title() + ' · ' + ', '.join(plants)
            manifest['filters'] = filters
            manifest['params'] = {k: v for k, v in params.items() if not k.startswith('_')}
            manifest['url'] = url
            initial_plant = factory_code(next((c['value'] for c in page.context.cookies() if c['name'] == 'MyFactoryName'), None))
            if initial_plant not in controller.SEVEN_FACTORIES:
                raise ValueError('Không xác định được nhà máy SFC hiện tại.')
        finally:
            browser.close()
    try:
        for plant in plants:
            try:
                result = controller.switch_factory(plant)
                if not result.get('success'):
                    raise RuntimeError(result.get('error', 'Không chuyển được nhà máy ' + plant))
                last = controller._execute_dynamic_action(action_type, target_id,
                    dict(params, _page_url=url, _filters=filters, _expected_plant=plant,
                         _mqaa_timeout_plant=initial_plant if mqaa_timeout_recovered else None))
                if not last.get('success') or not last.get('downloaded_file'):
                    raise RuntimeError(last.get('error') or 'Chưa lưu được Excel của ' + plant)
                files.append(last['downloaded_file'])
            except Exception as exc:
                failures.append(dict(plant=plant, error=str(exc)))
                manifest['failed_plants'] = failures
            save_manifest()
        if failures:
            error = '; '.join(item['plant'] + ': ' + item['error'] for item in failures)
    except Exception as exc:
        error = str(exc)
    finally:
        try:
            restored = controller.switch_factory(initial_plant)
            if not restored.get('success'):
                raise RuntimeError(restored.get('error', 'Không khôi phục được nhà máy.'))
            restored = controller._execute_dynamic_action('restore_filters', params=dict(_page_url=url, _filters=filters))
            if not restored.get('success'):
                raise RuntimeError(restored.get('error'))
            last.update({k: restored[k] for k in ('screenshot', 'inspection') if k in restored})
        except Exception as exc:
            error = (error + '; ' if error else '') + 'Khôi phục bộ lọc: ' + str(exc)
        controller.SELECTED_FACTORIES = original_scope
    combined = None
    if combine and not error:
        try:
            from sfc_export_merge import combine_factory_exports
            name = Path(url.split('?', 1)[0]).stem
            combined = combine_factory_exports(files, folder / f'{name}_Combined_{len(plants)}Factories.xlsx', plants)
            manifest['combined_files'] = [combined]
        except Exception as exc:
            error = 'Gộp Excel: ' + str(exc)
    manifest['status'] = 'error' if error else 'done'
    manifest['error'] = error
    history_status = ''
    if not error and url.split('?', 1)[0].lower().endswith('/qareportcopq.aspx'):
        from copq_dashboard_history import update_direct_export
        history_status = update_direct_export(files, filters)
        if history_status:
            manifest['copq_dashboard_history'] = history_status
    save_manifest()
    empty_count = sum(item.get('rows') == 0 for item in files)
    return dict(last, success=error is None, error=error, failed_plants=failures,
                downloaded_files=[combined] if combined else files, source_files=files,
                downloaded_file=combined or (files[0] if len(files) == 1 else None),
                message=(f'Đã xuất và gộp {len(files)} nhà máy thành 1 file Excel.' +
                         ' Dữ liệu được gom trong một sheet Combined.'
                         if combined else f'Đã lưu {len(files)}/{len(plants)} file Excel.') +
                        (f' {empty_count} file không có dữ liệu với bộ lọc đã chọn.' if empty_count else '') +
                        (' ' + history_status if history_status else '') +
                        (' ' + error if error else ''))
