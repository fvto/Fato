"""Refresh QMS navigation and the complete, authenticated QA PDF catalog."""
from datetime import datetime, timedelta, timezone
from email.message import Message
from hashlib import sha256
from io import BytesIO
import errno
import json
from pathlib import Path
import re
import threading
from urllib.parse import urljoin, urlsplit, urlunsplit, quote, unquote
from urllib.request import HTTPRedirectHandler, Request, build_opener, HTTPCookieProcessor

from lxml import html
from sfc_http import Session, Form, PLANTS, QMS

ROOT = Path(__file__).resolve().parent
MEMORY = ROOT / 'sfc_memory_knowledge_base.json'
CATALOG = ROOT / 'Report' / 'QMSDocuments' / 'catalog.json'
SOURCES = {
    'manual': ('QA Manual', 'https://sfc.chingluh.com/vhmes/QA/QAReasonBasic.aspx?tree_id=Q05'),
    'document': ('QA Document', 'https://sfc.chingluh.com/vhmes/QA/QADocument.aspx?id=QMS750'),
}
REFRESH_LOCK = threading.Lock()
GRID = 'ctl00$rpnlMain$ContentPlaceHolder1$gvData'
MAX_PDF_BYTES = 100 * 1024 * 1024


def download_error(error):
    from urllib.error import HTTPError, URLError
    if isinstance(error, HTTPError):
        status = error.code
        if status in (404, 410):
            return dict(kind='missing', http_status=status, message=f'SFC trả HTTP {status}: file PDF gốc không tồn tại trên máy chủ.')
        if status in (401, 403):
            return dict(kind='authentication', http_status=status, message=f'SFC trả HTTP {status}: cần đăng nhập lại hoặc kiểm tra quyền đọc tài liệu.')
        return dict(kind='server', http_status=status, message=f'SFC trả HTTP {status} khi tải PDF. File thành công vẫn được giữ để tiếp tục.')
    if isinstance(error, (TimeoutError, URLError, ConnectionError)):
        return dict(kind='connection', message='Kết nối tới SFC bị ngắt hoặc hết thời gian chờ. File thành công vẫn được giữ để tiếp tục.')
    if isinstance(error, OSError):
        if error.errno == errno.ENOSPC or getattr(error, 'winerror', None) == 112:
            return dict(kind='disk_full', message='Ổ đĩa lưu PDF đã hết dung lượng.')
        if isinstance(error, PermissionError):
            return dict(kind='write_permission', message='Không có quyền ghi hoặc file đang bị khóa trên máy.')
        return dict(kind='io', message='Không đọc/ghi được file PDF trên máy. Các file đã lưu vẫn được giữ.')
    if isinstance(error, RuntimeError):
        message = str(error)
        return dict(kind='invalid_pdf' if 'PDF hoàn chỉnh' in message else 'source', message=message)
    return dict(kind='unexpected', message='Không đọc được nội dung PDF trả về từ SFC. File thành công vẫn được giữ để tiếp tục.')


def allowed_url(url):
    parsed = urlsplit(url)
    return (parsed.scheme == 'https' and parsed.hostname == 'sfc.chingluh.com'
            and parsed.port in (None, 443) and not parsed.username and not parsed.password
            and parsed.path.startswith('/vhmes/'))


class SfcRedirects(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if not allowed_url(newurl):
            raise RuntimeError('SFC chuyển tới địa chỉ tải không hợp lệ.')
        return super().redirect_request(req, fp, code, msg, headers, newurl)


class DocumentSession(Session):
    def __init__(self, plant, cookies=None):
        super().__init__(plant, cookies)
        self.opener = build_opener(HTTPCookieProcessor(self.jar), SfcRedirects())

    def fetch(self, url, fields=None, limit=8 * 1024 * 1024):
        from urllib.parse import urlencode
        if not allowed_url(url):
            raise ValueError('Địa chỉ tài liệu SFC không hợp lệ.')
        data = urlencode(fields).encode() if fields is not None else None
        req = Request(url, data=data, headers={'User-Agent': 'Chingluh-QMS-Dashboard/1.0', 'Referer': QMS})
        with self.opener.open(req, timeout=90) as response:
            if 'login.aspx' in response.url.lower():
                raise RuntimeError('Phiên SFC đã hết hạn. Đăng nhập lại SFC rồi thử lại.')
            body = response.read(limit + 1)
            if len(body) > limit:
                raise RuntimeError('File hoặc trang SFC vượt giới hạn dung lượng cho phép.')
            return body, response.headers

    def prepare_catalog(self):
        body, _ = self.fetch(QMS)
        actual = next((c.value for c in self.jar if c.name == 'MyFactoryName'), None)
        if actual != self.plant:
            fields = Form(body.decode('utf-8')).fields
            fields.update(__EVENTTARGET='wucHeaderMainTop$TopPanel$menuRightTop',
                          __EVENTARGUMENT='CLICK:0i' + str(PLANTS.index(self.plant)))
            body, _ = self.fetch(QMS, fields)
        actual = next((c.value for c in self.jar if c.name == 'MyFactoryName'), None)
        if actual != self.plant:
            raise RuntimeError(f'SFC chưa xác nhận nhà máy {self.plant}.')
        return body.decode('utf-8')


def tree(markup):
    root = html.fromstring(markup)
    for element in root.xpath('//script|//style'):
        element.drop_tree()
    return root


def clean(element):
    return ' '.join(element.text_content().split())


def parse_navigation(markup):
    root = tree(markup)
    modules, seen = [], set()
    # Read categories from the live sitemap instead of guessing from old URLs.
    for column in root.xpath('//td[contains(@id,"ASPxSiteMapControl1_0_C")]'):
        category = 'Other Modules'
        for element in column.xpath('.//div[contains(@class,"dxsmLevel0") or contains(@class,"dxsmLevel1")]'):
            if 'dxsmLevel0' in element.get('class', ''):
                category = clean(element)
            else:
                for link in element.xpath('./a[@href]'):
                    add_module(modules, seen, link, category)
    if not modules or not any('QADocument.aspx' in m['href'] for m in modules):
        raise RuntimeError('Chưa đọc được cây QMS mới đã đăng nhập; giữ nguyên danh mục cũ.')
    for link in root.xpath('//div[contains(@id,"ASPxRoundPanel1_ASPxMenu1")]//a[@href]'):
        add_module(modules, seen, link, 'Other Modules')
    # Keep QMS Home itself available in the directory.
    if QMS not in seen:
        modules.append(dict(name='QMS Home', href=QMS, category='Other Modules'))
    return modules


def add_module(modules, seen, link, category):
    href = urljoin(QMS, link.get('href')).split('#')[0]
    name = clean(link)
    if name and allowed_url(href) and href not in seen:
        seen.add(href)
        modules.append(dict(name=name, href=href, category=category, code=link.get('title', '')))


def pdf_entry(link, source, plant, **metadata):
    url = urljoin(SOURCES[source][1], link.get('href', ''))
    if not allowed_url(url) or urlsplit(url).path.lower() != '/vhmes/qa/downloadfile.aspx':
        raise ValueError('Liên kết PDF SFC không hợp lệ.')
    key = sha256(f'{plant}|{source}|{url}'.encode()).hexdigest()[:32]
    return dict(id=key, source=source, plant=plant, source_url=url, **metadata)


def parse_pdf_page(markup, source, plant):
    root = tree(markup)
    entries = []
    for row in root.xpath('//tr[contains(@id,"_gvData_DXDataRow")]'):
        cells = row.xpath('./td')
        text = [clean(cell) for cell in cells]
        if source == 'document':
            if len(cells) != 12:
                raise RuntimeError('Cấu trúc bảng QA Document đã thay đổi.')
            metadata = dict(date=text[3], code=text[4], kind=text[5], project=text[6],
                            name=text[7] or text[8] or text[4], shoe=text[8], process=text[9], collection='records')
        else:
            if len(cells) != 7:
                raise RuntimeError('Cấu trúc bảng QA Manual đã thay đổi.')
            metadata = dict(date='', code=text[1], kind='QA manual', project='',
                            name=text[3] or text[2] or text[1], shoe=text[3], process='', collection='records')
        for link in row.xpath('.//a[contains(@href,"DownloadFile.aspx?")]'):
            language = 'VN' if 'hplFile_VN' in link.get('id', '') else 'EN' if source == 'document' else ''
            entries.append(pdf_entry(link, source, plant, language=language, **metadata))
    if source == 'manual':
        for link in root.xpath('//div[contains(@id,"btnHFPA_DXM") or contains(@id,"btnBDTS_DXM")]//a[contains(@href,"DownloadFile.aspx?")]'):
            collection = 'bdts' if any('btnBDTS_DXM' in el.get('id', '') for el in link.iterancestors()) else 'hfpa'
            entries.append(pdf_entry(link, source, plant, name=clean(link), code='', language='', date='',
                                     kind='BDTS' if collection == 'bdts' else 'HFPA Guide', collection=collection,
                                     project='', shoe='', process=''))
    summaries = root.xpath('//*[contains(@class,"dxp-summary")]/text()')
    match = re.search(r'Page\s+(\d+)\s+of\s+(\d+)\s+\((\d+)\s+items?\)', ' '.join(summaries))
    if not match:
        raise RuntimeError(f'Chưa đọc được phân trang {SOURCES[source][0]}; giữ nguyên bản cũ.')
    current, pages, total = map(int, match.groups())
    return entries, current, pages, total, len(root.xpath('//tr[contains(@id,"_gvData_DXDataRow")]'))


def crawl_documents(session, source, progress=None):
    markup = session.fetch(SOURCES[source][1])[0].decode('utf-8')
    entries, seen, row_count = [], set(), 0
    expected_total = None
    for page_index in range(500):
        page_entries, current, pages, total, rows = parse_pdf_page(markup, source, session.plant)
        if current != page_index + 1 or pages > 500 or (expected_total is not None and total != expected_total):
            raise RuntimeError('Phân trang SFC không nhất quán; chưa thay bản danh mục đang dùng.')
        expected_total = total
        row_count += rows
        for entry in page_entries:
            if entry['id'] not in seen:
                seen.add(entry['id'])
                entries.append(entry)
        if progress:
            progress(f'{SOURCES[source][0]}: {current}/{pages}')
        if current == pages:
            if row_count != total:
                raise RuntimeError(f'{SOURCES[source][0]} chưa đủ dòng: {row_count}/{total}.')
            return entries, dict(rows=total, pages=pages, pdfs=len(entries))
        fields = Form(markup).fields
        argument = f'PN{current}'
        fields.update(__EVENTTARGET=GRID, __EVENTARGUMENT=f'12|PAGERONCLICK{len(argument)}|{argument}')
        markup = session.fetch(SOURCES[source][1], fields)[0].decode('utf-8')
    raise RuntimeError('Danh mục SFC có quá nhiều trang; chưa thay bản danh mục đang dùng.')


def read_json(path, default=None):
    return json.loads(path.read_text(encoding='utf-8')) if path.exists() else (default or {})


def atomic_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    pending = path.with_suffix('.tmp')
    pending.write_text(json.dumps(value, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    pending.replace(path)


def refresh(plant='VH', progress=None):
    if not REFRESH_LOCK.acquire(blocking=False):
        raise RuntimeError('Đang cập nhật QMS. Chờ lượt hiện tại hoàn tất.')
    try:
        session = DocumentSession('JV3' if plant == 'JVB' else plant)
        modules = parse_navigation(session.prepare_catalog())
        documents, sources = [], {}
        for source in SOURCES:
            items, stats = crawl_documents(session, source, progress)
            documents.extend(items)
            sources[source] = stats
        timestamp = datetime.now(timezone(timedelta(hours=7))).isoformat(timespec='seconds')
        catalog = dict(plant=plant, refreshed_at=timestamp, sources=sources, documents=documents)
        catalogs = read_json(CATALOG)
        catalogs[plant] = catalog
        memory = read_json(MEMORY)
        memory.update(crawled_at=timestamp, source_url=QMS, qms_modules=modules, qms_plant=plant)
        atomic_json(CATALOG, catalogs)
        atomic_json(MEMORY, memory)
        return dict(plant=plant, refreshed_at=timestamp, modules=len(modules), sources=sources, pdfs=len(documents))
    finally:
        REFRESH_LOCK.release()


def document_catalog(plant):
    catalog = read_json(CATALOG).get(plant)
    if not catalog:
        return dict(plant=plant, refreshed_at=None, sources={}, documents=[])
    return dict(catalog, documents=[dict({k: v for k, v in item.items() if k != 'source_url'}, collection=collection_of(item),
                                       url=f'/api/qms/documents/{plant}/{item["id"]}/pdf')
                                    for item in catalog['documents']])


def collection_of(item):
    # Compatibility with catalogs saved before the library separated the SFC table and guides.
    return item.get('collection') or ('records' if item['source'] == 'document' or item.get('code')
                                      else 'bdts' if item.get('kind') == 'BDTS' else 'hfpa')


def download_pdf(plant, document_id, session=None):
    catalog = read_json(CATALOG).get(plant, {})
    entry = next((item for item in catalog.get('documents', []) if item['id'] == document_id), None)
    if not entry:
        raise FileNotFoundError('Không tìm thấy tài liệu. Cập nhật lại QMS rồi thử lại.')
    if session is None:
        session = DocumentSession('JV3' if plant == 'JVB' else plant)
        session.prepare_catalog()
    body, headers = session.fetch(entry['source_url'], limit=MAX_PDF_BYTES)
    original_name = ''
    # DownloadFile.aspx is an HTML viewer on SFC, with the original PDF in <object>.
    if not body.startswith(b'%PDF-'):
        viewer = tree(body.decode('utf-8'))
        targets = viewer.xpath('//object[@type="application/pdf"]/@data|//embed[@type="application/pdf"]/@src')
        if len(targets) == 1:
            target = urljoin(entry['source_url'], targets[0])
            parsed = urlsplit(target)
            path = unquote(parsed.path)
            if (not allowed_url(target) or not path.startswith('/vhmes/QA/UploadFile/')
                    or not path.lower().endswith('.pdf') or '..' in path.split('/')):
                raise RuntimeError('Đường dẫn PDF gốc không hợp lệ.')
            original_name = path.rsplit('/', 1)[-1]
            target = urlunsplit((parsed.scheme, parsed.netloc, quote(parsed.path, safe='/%'),
                                 quote(parsed.query, safe='=&%'), ''))
            body, headers = session.fetch(target, limit=MAX_PDF_BYTES)
    if not body.startswith(b'%PDF-') or b'%%EOF' not in body[-4096:]:
        raise RuntimeError('SFC chưa trả về PDF hoàn chỉnh. File có thể đã bị gỡ hoặc phiên đăng nhập đã hết hạn.')
    disposition = Message()
    disposition['Content-Disposition'] = headers.get('Content-Disposition', '')
    fallback = f'{entry.get("code") or entry["name"]}{"_" + entry["language"] if entry.get("language") else ""}.pdf'
    name = disposition.get_filename() or original_name or fallback
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip(' .')[:180]
    if not name.lower().endswith('.pdf'):
        name += '.pdf'
    return BytesIO(body), name
