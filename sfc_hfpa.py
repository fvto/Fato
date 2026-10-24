"""Verified HFPA filters, server postbacks, and completed Excel downloads."""
from datetime import datetime, timedelta, timezone
from pathlib import Path
import logging
from uuid import uuid4
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from sfc_excel import EXPORT_ROOT, file_info, validate_export

PREFIX = 'rpnlMain_ContentPlaceHolder1_'
logger = logging.getLogger(__name__)
WAIT_MS = 15 * 60 * 1000


def wait_idle(page):
    logger.info('Waiting for SFC callbacks to finish; keeping current page')
    page.wait_for_function("() => typeof aspxGetControlCollection === 'function'", timeout=WAIT_MS)
    page.evaluate("""timeout => new Promise((resolve, reject) => {
        const collection = aspxGetControlCollection();
        const handlers = [];
        let timer, poll;
        const cleanup = () => {
            clearTimeout(timer); clearInterval(poll);
            for (const [event, handler] of handlers) event.RemoveHandler(handler);
        };
        const check = () => {
            let busy = false;
            collection.ForEachControl(c => {
                if (typeof c.InCallback === 'function' && c.InCallback()) busy = true;
            });
            if (!busy && collection.Get('rpnlMain_ContentPlaceHolder1_ckbNIKE')) {
                cleanup(); resolve(true);
            }
        };
        collection.ForEachControl(c => {
            if (c.EndCallback && typeof c.EndCallback.AddHandler === 'function') {
                const handler = () => setTimeout(check, 0);
                c.EndCallback.AddHandler(handler);
                handlers.push([c.EndCallback, handler]);
            }
        });
        timer = setTimeout(() => { cleanup(); reject(new Error('SFC vẫn đang xử lý sau 15 phút; giữ nguyên trang, chưa gửi lại lệnh.')); }, timeout);
        poll = setInterval(check, 1000);
        check();
    })""", WAIT_MS)


def postback(page, suffix):
    # SFC can respond via a full postback or an AJAX callback. Wait for the
    # response and EndCallback instead of assuming every action navigates.
    with page.expect_response(
        lambda response: response.request.method == 'POST' and 'QCHFPA.aspx' in response.url,
        timeout=WAIT_MS,
    ) as pending:
        page.locator('#' + PREFIX + suffix).click(no_wait_after=True)
    response = pending.value
    response.finished()
    if response.status >= 400:
        raise RuntimeError(f'SFC trả lỗi HTTP {response.status}; chưa thực hiện lại thao tác.')
    page.wait_for_load_state('domcontentloaded', timeout=WAIT_MS)
    wait_idle(page)


def filters(page):
    return page.evaluate("""() => {
        const coll = aspxGetControlCollection();
        const get = suffix => coll.Get('rpnlMain_ContentPlaceHolder1_' + suffix);
        const iso = c => {
            const d = c.GetDate();
            return d ? [d.getFullYear(), String(d.getMonth()+1).padStart(2,'0'), String(d.getDate()).padStart(2,'0')].join('-') : null;
        };
        return {dateFrom: iso(get('dtDateFrom')), dateTo: iso(get('dtDateTo')),
            factory: get('cboDeptFrom').GetText(), nike: get('ckbNIKE').GetChecked()};
    }""")


def run(controller, date_from=None, date_to=None, factory=None, action=None, nike=None,
        export_path=None, expected_plant=None):
    if action not in (None, 'search', 'excel'):
        return dict(success=False, error='Thao tác HFPA không hợp lệ.')
    with controller.SFC_LOCK:
        try:
            with sync_playwright() as playwright:
                logger.info('Connecting to SFC browser')
                browser = playwright.chromium.connect_over_cdp(controller.CDP_URL, timeout=controller.SFC_CONNECT_WAIT_MS)
                try:
                    logger.info('SFC browser connected')
                    page = controller.get_sfc_page(browser)
                    dialogs = []
                    def on_dialog(dialog):
                        dialogs.append(dialog.message)
                        dialog.dismiss()
                    page.on('dialog', on_dialog)
                    if 'QCHFPA.aspx' not in page.url:
                        page.goto(controller.URL_MAPPING['hfpa'], wait_until='domcontentloaded')
                    if 'login.aspx' in page.url.lower():
                        raise RuntimeError('SFC yêu cầu đăng nhập lại; chưa xuất được Excel.')
                    wait_idle(page)
                    cookies = page.context.cookies()
                    plant = next((c['value'] for c in cookies if c['name'] == 'MyFactoryName'), None)
                    if expected_plant and plant != expected_plant:
                        raise ValueError(f'Nhà máy thực tế {plant} khác {expected_plant}; dừng xuất để tránh lẫn dữ liệu.')
                    if nike is not None and filters(page)['nike'] != nike:
                        postback(page, 'ckbNIKE_S_D')
                    result = page.evaluate("""params => {
                        const coll = aspxGetControlCollection();
                        const get = suffix => coll.Get('rpnlMain_ContentPlaceHolder1_' + suffix);
                        for (const [key, name] of [['dateFrom','dtDateFrom'],['dateTo','dtDateTo']]) {
                            if (params[key]) {
                                const [y,m,d] = params[key].split('-').map(Number);
                                get(name).SetDate(new Date(y,m-1,d));
                            }
                        }
                        if (params.factory !== null) {
                            const cbo = get('cboDeptFrom');
                            if (params.factory === '' || params.factory === 'ALL' || params.factory === undefined) { 
                                cbo.SetValue(null); 
                                cbo.SetText(''); 
                            }
                            else {
                                const normalized = x => x.toLowerCase().replace(/\\s+/g,'').replace(/-/g,'');
                                const target = normalized(params.factory);
                                const matches = [];
                                for (let i=0; i<cbo.GetItemCount(); i++) {
                                    const item = cbo.GetItem(i);
                                    if (item && normalized(item.text) === target) matches.push(i);
                                }
                                if (matches.length !== 1) return {error:'Không tìm thấy duy nhất chuyền: '+params.factory};
                                cbo.SetSelectedIndex(matches[0]);
                            }
                        }
                        return {};
                    }""", dict(dateFrom=date_from, dateTo=date_to, factory=factory))
                    if result.get('error'):
                        raise ValueError(result['error'])

                    def verify():
                        current = filters(page)
                        for key, value in [('dateFrom', date_from), ('dateTo', date_to), ('nike', nike)]:
                            if value is not None and current[key] != value:
                                raise ValueError(f'Bộ lọc {key} chưa đúng: {current[key]} (yêu cầu {value})')
                        if factory in ('', 'ALL', None) and current['factory'] and current['factory'] not in ('', 'ALL'):
                            # cboDeptFrom was cleared, but in case SFC still displays text
                            pass
                        return current

                    current = verify()
                    logger.info('HFPA filters verified: %s', current)
                    downloaded = None
                    message = None
                    # NIKE selects the external data columns in the grid. Use
                    # the grid's Excel export, rather than the separate print
                    # templates; validate the downloaded period before publishing.
                    if action in ('search', 'excel'):
                        logger.info('Waiting for HFPA search')
                        postback(page, 'btnSearch')
                        if dialogs:
                            raise ValueError('; '.join(dialogs))
                        current = verify()
                        logger.info('HFPA search completed')
                        message = 'Đã tìm kiếm với bộ lọc đã kiểm tra.'
                    if action == 'excel':
                        use_nike = current['nike']
                        button = 'btnExcel'
                        logger.info('Requesting export with %s', button)
                        with page.expect_download(timeout=WAIT_MS) as pending:
                            page.locator('#' + PREFIX + button).click(no_wait_after=True)
                        download = pending.value
                        logger.info('Download started: %s', download.suggested_filename)
                        failure = download.failure()
                        logger.info('Download finished; failure=%s', failure)
                        if failure:
                            raise RuntimeError('Tải Excel thất bại: ' + failure)
                        suffix = Path(download.suggested_filename).suffix.lower()
                        if suffix not in ('.xls', '.xlsx'):
                            raise ValueError('SFC không trả về định dạng Excel: ' + download.suggested_filename)
                        if export_path is None:
                            run_id = datetime.now(timezone(timedelta(hours=7))).strftime('%Y%m%d_%H%M%S') + '_' + uuid4().hex[:6]
                            export_path = EXPORT_ROOT / run_id / f'HFPA_{plant}_{"NIKE" if use_nike else "Internal"}_{current["dateFrom"]}_{current["dateTo"]}'
                        destination = Path(str(export_path) + suffix)
                        destination.parent.mkdir(parents=True, exist_ok=True)
                        temporary = destination.with_suffix('.partial' + suffix)
                        download.save_as(str(temporary.resolve()))
                        count = validate_export(temporary, current['dateFrom'], current['dateTo'], use_nike)
                        temporary.replace(destination)
                        downloaded = file_info(destination, count)
                        message = f'Đã lưu và kiểm tra {destination.name}: {count:,} dòng.'
                    controller.LAST_STATUS_TIME = 0
                    response = dict(success=True, filter_result=current, action_result=message)
                    if downloaded:
                        response['file'] = downloaded
                    else:
                        response['screenshot'] = controller.take_b64_screenshot(page)
                    return response
                finally:
                    browser.close()
        except PlaywrightTimeoutError as error:
            controller.LAST_STATUS_TIME = 0
            if 'connect_over_cdp' in str(error):
                return dict(success=False, error='Chưa kết nối điều khiển được Chrome SFC. Hãy kiểm tra trang đã tải xong và đã đăng nhập, rồi thử lại.')
            return dict(success=False, error='SFC vẫn đang xử lý sau 15 phút. Trang được giữ nguyên; chưa tự gửi lại lệnh và chưa xác nhận hoàn tất file.')
        except Exception as error:
            controller.LAST_STATUS_TIME = 0
            return dict(success=False, error=str(error))
