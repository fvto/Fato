"""Authenticated HFPA exports in an isolated ASP.NET session.

Only the dedicated browser's SFC login cookies are read, through private IPC.
The browser page and its potentially busy ASP.NET session are left untouched.
"""
from datetime import date, datetime, timedelta, timezone
from html.parser import HTMLParser
import http.cookiejar
import json
import logging
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import queue
import shutil
import subprocess
import threading
import time
import urllib.parse
import urllib.request

ROOT = Path(__file__).resolve().parent
HFPA = 'https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504'
QMS = 'https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx'
PREFIX = 'ctl00$rpnlMain$ContentPlaceHolder1$'
PLANTS = ('VH', 'VH2', 'VH3', 'VH4', 'JV', 'JV2', 'JV3')
GROUPS = {'VH': {'CL_VH', 'CL_VB'}, 'VH2': {'CL_VH2'}, 'JV': {'CL_JV1'}, 'JV2': {'CL_JV2'}}
logger = logging.getLogger(__name__)


class Form(HTMLParser):
    def __init__(self, markup):
        super().__init__()
        self.fields = {}
        self.feed(markup)

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag != 'input' or not attrs.get('name') or 'disabled' in attrs:
            return
        kind = attrs.get('type', 'text').lower()
        # Sending all submit buttons executes unrelated print/export handlers.
        if kind in ('submit', 'button', 'image', 'reset', 'file'):
            return
        if kind in ('checkbox', 'radio') and 'checked' not in attrs:
            return
        self.fields[attrs['name']] = attrs.get('value', '')


def login_cookies():
    node = shutil.which('node')
    if not node:
        raise RuntimeError('Không tìm thấy Node.js để kết nối phiên đăng nhập SFC.')
    try:
        result = subprocess.run([node, str(ROOT / 'sfc_session_cookies.js')],
                                capture_output=True, text=True, timeout=12, check=True,
                                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        cookies = json.loads(result.stdout)
    except Exception as error:
        raise RuntimeError('Chưa đọc được phiên đăng nhập SFC. Hãy mở Dashboard và DOM rồi đăng nhập SFC.') from error
    if not cookies:
        raise RuntimeError('Không tìm thấy phiên đăng nhập SFC trong Chrome SFC.')
    return cookies


class Session:
    def __init__(self, plant, cookies=None):
        if plant not in PLANTS:
            raise ValueError('Nhà máy chưa được hỗ trợ: ' + str(plant))
        self.plant = plant
        self.jar = http.cookiejar.CookieJar()
        self.login = login_cookies() if cookies is None else cookies
        for cookie in self.login:
            if cookie['name'].lower() == 'asp.net_sessionid':
                continue
            self.jar.set_cookie(http.cookiejar.Cookie(
                0, cookie['name'], cookie['value'], None, False,
                'sfc.chingluh.com', True, False, '/', True, True, None,
                True, None, None, {}))
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.jar))
        self.form = None
        self.worker = None

    def request(self, url, fields=None):
        data = urllib.parse.urlencode(fields).encode() if fields is not None else None
        request = urllib.request.Request(url, data=data, headers={
            'User-Agent': 'MQAA-SFC-Dashboard/1.0', 'Referer': url})
        started = time.monotonic()
        # A slow request is allowed to finish; it is never reloaded or replayed.
        with self.opener.open(request, timeout=15 * 60) as response:
            if urllib.parse.urlparse(response.url).hostname != 'sfc.chingluh.com' or 'login.aspx' in response.url.lower():
                raise RuntimeError('Phiên SFC đã hết đăng nhập. Vui lòng đăng nhập lại Chrome SFC.')
            body = response.read()
            logger.info('HFPA %s %s completed in %.1fs (%s bytes)', self.plant,
                        'POST' if fields is not None else 'GET', time.monotonic() - started, len(body))
            return body

    def prepare(self):
        markup = self.request(QMS).decode('utf-8')
        actual = next((c.value for c in self.jar if c.name == 'MyFactoryName'), None)
        if actual != self.plant:
            fields = Form(markup).fields
            fields.update(__EVENTTARGET='wucHeaderMainTop$TopPanel$menuRightTop',
                          __EVENTARGUMENT='CLICK:0i' + str(PLANTS.index(self.plant)))
            self.request(QMS, fields)
        actual = next((c.value for c in self.jar if c.name == 'MyFactoryName'), None)
        if actual != self.plant:
            raise RuntimeError(f'SFC chưa xác nhận nhà máy {self.plant}; hiện tại {actual}.')
        self.form = Form(self.request(HFPA).decode('utf-8'))
        required = ['__VIEWSTATE', PREFIX + 'dtDateFrom', PREFIX + 'dtDateTo', PREFIX + 'ckbNIKE']
        if any(name not in self.form.fields for name in required):
            raise RuntimeError('Trang SFC chưa trả về biểu mẫu HFPA đã đăng nhập.')

    def export_part(self, first, last, nike, path):
        from sfc_excel import validate_export, read_export, file_info
        def verified(source):
            rows = validate_export(source, first, last, nike)
            header, records = read_export(source)
            if records and any('Factory Group' in str(label) for label in header):
                groups = {str(row[0]).strip() for row in records}
                allowed = GROUPS.get(self.plant)
                if allowed is None or not groups <= allowed:
                    raise ValueError(f'Nhóm nhà máy trong file chưa khớp {self.plant}: {sorted(groups)}')
            return rows

        for suffix in ('.xls', '.xlsx'):
            cached = Path(str(path) + suffix)
            if cached.is_file():
                rows = verified(cached)
                logger.info('Reusing verified HFPA part %s %s', self.plant, first)
                return dict(file_info(cached, rows), date_from=first, date_to=last,
                            plant=self.plant, nike=nike)
        if self.form is None:
            self.prepare()
        fields = dict(self.form.fields)
        for control, day in [('dtDateFrom', first), ('dtDateTo', last)]:
            # ASPxClientDateEdit.SyncRawInputValue serializes local wall-clock
            # milliseconds (_aspxToLocalTime), independently of Python's TZ.
            value = datetime.combine(date.fromisoformat(day), datetime.min.time(), timezone.utc)
            fields['rpnlMain_ContentPlaceHolder1_' + control + '_Raw'] = str(int(value.timestamp() * 1000))
            fields[PREFIX + control] = day
        fields[PREFIX + 'ckbNIKE'] = 'C' if nike else 'U'
        fields[PREFIX + 'cboDeptFrom'] = ''
        fields['rpnlMain_ContentPlaceHolder1_cboDeptFrom_VI'] = ''
        fields.update(__EVENTTARGET=PREFIX + 'btnExcel', __EVENTARGUMENT='')
        logger.info('Exporting HFPA %s %s..%s NIKE=%s', self.plant, first, last, nike)
        data = self.request(HFPA, fields)
        suffix = '.xlsx' if data.startswith(b'PK') else '.xls'
        if not (data.startswith(b'PK') or data.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1')):
            raise RuntimeError('SFC không trả về file Excel; chưa lưu hoặc báo thành công.')
        destination = Path(str(path) + suffix)
        destination.parent.mkdir(parents=True, exist_ok=True)
        pending = destination.with_suffix('.partial' + suffix)
        pending.write_bytes(data)
        rows = verified(pending)
        pending.replace(destination)
        return dict(file_info(destination, rows), date_from=first, date_to=last,
                    plant=self.plant, nike=nike)

    def export_period(self, first, last, nike, path, factory=''):
        """Bound server work per request, then reconcile into one monthly file."""
        from sfc_excel import combine_exports
        if factory:
            raise ValueError('Xuất theo tháng hiện chỉ hỗ trợ tất cả chuyền; chưa thực hiện bộ lọc chuyền riêng.')
        days = []
        current, end = date.fromisoformat(first), date.fromisoformat(last)
        while current <= end:
            days.append(current.isoformat())
            current += timedelta(days=1)
        if not days:
            raise ValueError('Khoảng ngày không hợp lệ.')
        if self.worker is None:
            self.worker = Session(self.plant, cookies=self.login)
            self.worker.prepare()
        folder = Path(path).parent / '_parts' / Path(path).name
        messages = queue.Queue()
        stop = threading.Event()

        def download_group(session, dates):
            try:
                for day in dates:
                    if stop.is_set():
                        break
                    messages.put(dict(message=f'Đang tải {self.plant} ngày {day}'))
                    item = session.export_part(day, day, nike, folder / day)
                    messages.put(dict(part=item))
            except Exception as error:
                stop.set()
                messages.put(dict(error=error))
            finally:
                messages.put(dict(finished=True))

        parts, errors, finished = [], [], 0
        # Two independent ASP.NET sessions, one request at a time per session.
        # No retry, page reload, or repeated request while an export is pending.
        with ThreadPoolExecutor(max_workers=2, thread_name_prefix='sfc-export') as pool:
            for slot, session in enumerate((self, self.worker)):
                pool.submit(download_group, session, days[slot::2])
            while finished < 2:
                try:
                    message = messages.get(timeout=30)
                except queue.Empty:
                    yield dict(message=f'SFC vẫn đang xử lý {self.plant}; đã kiểm tra {len(parts)}/{len(days)} phần ngày, đang chờ phản hồi.')
                    continue
                if 'error' in message:
                    errors.append(message['error'])
                elif message.get('finished'):
                    finished += 1
                elif 'part' in message:
                    parts.append(message['part'])
                    yield dict(message=f'Đã kiểm tra {self.plant}: {len(parts)}/{len(days)} phần ngày của kỳ {first} → {last}')
                else:
                    yield message
        if errors:
            raise errors[0]
        parts.sort(key=lambda item: item['date_from'])
        if [item['date_from'] for item in parts] != days or any(item['date_from'] != item['date_to'] for item in parts):
            raise ValueError('Các phần ngày chưa đủ hoặc bị trùng; chưa tạo file tháng.')
        yield dict(message=f'Đang ghép và đối chiếu file tháng {self.plant}: {first} → {last}')
        item = combine_exports(parts, Path(str(path) + '.xlsx'), self.plant, nike, include_sources=False)
        item['parts'] = parts
        yield dict(file=item)
