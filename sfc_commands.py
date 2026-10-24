"""Parse SFC requests and execute verified, sequential export plans."""
import calendar
from datetime import date, datetime, timedelta, timezone
import json
from pathlib import Path
import re
import threading
import unicodedata
from uuid import uuid4

COMMAND_LOCK = threading.Lock()
EXPORT_ROOT = Path(__file__).resolve().parent / "Report" / "HFPAExports"
PLANTS = ("VH", "VH2", "VH3", "VH4", "JV", "JV2", "JVB")


def normalize(text):
    return ''.join(c for c in unicodedata.normalize('NFD', text.lower().replace('đ', 'd'))
                   if unicodedata.category(c) != 'Mn')


def parse_intent(text):
    t = normalize(text.strip())
    today = datetime.now(timezone(timedelta(hours=7))).date()
    plants = list(dict.fromkeys(p.upper() for p in re.findall(r'\b(vh[234]?|jv2|jvb|jv)\b', t)))
    if not plants:
        m = re.search(r'nha may\s*([1-4])\b', t)
        if m:
            plants = ['VH' if m[1] == '1' else 'VH' + m[1]]
    out = dict(target_plants=plants, target_plant=plants[0] if plants else None,
               target_line=None, target_page=None, date_from=None, date_to=None,
               action=None, nike=None, monthly=False, combine=False, summary=[])
    pages = [(('hfpa', 'q504'), 'hfpa'), (('bao cao mqaa', 'mqaareport'), 'mqaa_report'),
             (('mqaa', 'kiem toan'), 'mqaa'), (('aql',), 'aql'),
             (('trang chu', 'portal', 'flowchart'), 'home'), (('qms', 'quan ly chat luong'), 'qms')]
    for words, key in pages:
        if any(word in t for word in words):
            out['target_page'] = key
            break
    if re.search(r'noi bo|internal|(?:tat|bo chon|khong bat|khong chon)\s+(?:nut\s+)?nike', t):
        out['nike'] = False
    elif re.search(r'\bnike\b|\bexternal\b', t):
        out['nike'] = True

    if 'hom nay' in t or 'today' in t:
        out['date_from'] = out['date_to'] = today.isoformat()
    else:
        raw_dates = re.findall(r'\b(?:\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{4})\b', t)
        parsed = []
        for raw in raw_dates:
            a, b, c = raw.replace('/', '-').split('-')
            parsed.append(date(int(a), int(b), int(c)) if len(a) == 4 else date(int(c), int(b), int(a)))
        if parsed:
            out['date_from'], out['date_to'] = parsed[0].isoformat(), parsed[-1].isoformat()
        else:
            span = re.search(r'thang\s*(\d{1,2})\s*(?:-|–|—|den|toi)\s*(?:thang\s*)?(\d{1,2})\s*(?:nam\s*)?(\d{4})?', t)
            single = re.search(r'thang\s*(\d{1,2})(?:\s*(?:nam\s*|/)(\d{4}))?', t)
            if span or single:
                first, last = (int(span[1]), int(span[2])) if span else (int(single[1]), int(single[1]))
                year_match = re.search(r'\bnam\s*(\d{4})\b', t)
                year = int((span[3] if span else single[2]) or (year_match[1] if year_match else today.year))
                if not 1 <= first <= last <= 12:
                    raise ValueError('Khoảng tháng không hợp lệ. Hãy dùng tháng 1 đến tháng 12 theo thứ tự tăng dần.')
                out['date_from'] = date(year, first, 1).isoformat()
                out['date_to'] = date(year, last, calendar.monthrange(year, last)[1]).isoformat()
    if out['date_from'] and out['date_to'] < out['date_from']:
        raise ValueError('Ngày kết thúc phải bằng hoặc sau ngày bắt đầu.')
    match = re.search(r'\b(v[lab]\s*-\s*[a-z0-9]+(?:\s+line\s*\d+)?)\b', t)
    if match:
        out['target_line'] = re.sub(r'\s*-\s*', ' - ', match[1]).upper()
    else:
        match = re.search(r'(?:chuyen|xuong|\bline)\s+(.+?)(?=\s+(?:ngay|tu|va|roi|tim|thang)\b|$)', t)
        if match:
            out['target_line'] = match[1].strip().upper()
    if re.search(r'\bexcel\b|\bexport\b|\bxlsx?\b', t):
        out['action'] = 'excel'
    elif re.search(r'tim|search|loc|tra cuu|xem', t) or out['target_line'] or out['date_from']:
        out['action'] = 'search'
    out['monthly'] = bool(re.search(r'tung thang|moi thang|theo thang|chia nho.*thang', t))
    out['combine'] = bool(re.search(r'combine|gop|hop nhat', t))
    if plants:
        out['summary'].append('Nhà máy: ' + ', '.join(plants))
    if out['target_page']:
        out['summary'].append('Trang: ' + out['target_page'].upper())
    if out['date_from']:
        out['summary'].append(f"Ngày: {out['date_from']} → {out['date_to']}")
    if out['target_line']:
        out['summary'].append('Chuyền: ' + out['target_line'])
    if out['nike'] is not None:
        out['summary'].append('External (NIKE)' if out['nike'] else 'Nội bộ')
    if out['action']:
        out['summary'].append('Xuất Excel' if out['action'] == 'excel' else 'Tìm kiếm')
    if out['monthly']:
        out['summary'].append('Chia từng tháng')
    if out['combine']:
        out['summary'].append('Gộp riêng theo nhà máy')
    return out


def periods(intent):
    if not intent['date_from']:
        raise ValueError('Vui lòng chỉ rõ khoảng ngày/tháng để xuất nhiều file.')
    first, last = date.fromisoformat(intent['date_from']), date.fromisoformat(intent['date_to'])
    if not intent['monthly']:
        return [(first.isoformat(), last.isoformat())]
    result = []
    while first <= last:
        end = min(date(first.year, first.month, calendar.monthrange(first.year, first.month)[1]), last)
        result.append((first.isoformat(), end.isoformat()))
        first = end + timedelta(days=1)
    return result


def event(step, status, msg, **extra):
    return dict(step=step, status=status, msg=msg, **extra)


def execute(controller, prompt):
    if not COMMAND_LOCK.acquire(blocking=False):
        yield event(1, 'error', 'Đang xử lý một lệnh khác. Vui lòng chờ hoàn tất.')
        return
    try:
        yield event(1, 'running', 'Đang phân tích yêu cầu...')
        intent = parse_intent(prompt)
        batch = len(intent['target_plants']) > 1 or intent['monthly'] or intent['combine']
        if batch and (intent['action'] != 'excel' or intent['target_page'] != 'hfpa'):
            raise ValueError('Xuất nhiều nhà máy/chia tháng hiện hỗ trợ Excel HFPA. Hãy chỉ rõ HFPA và xuất Excel.')
        if batch and intent['nike'] is None:
            raise ValueError('Hãy chỉ rõ NIKE (external) hoặc nội bộ cho lệnh xuất nhiều file.')
        if not intent['summary']:
            raise ValueError('Chưa nhận diện được yêu cầu. Hãy chỉ rõ trang, nhà máy hoặc thao tác.')
        yield event(1, 'success', 'Đã nhận diện: ' + ' | '.join(intent['summary']), intent=intent)
        direct_hfpa = (intent['action'] == 'excel' and intent['target_page'] == 'hfpa'
                       and intent['date_from'] and intent['target_plants']
                       and intent['nike'] is not None and not intent['target_line'])
        if batch or direct_hfpa:
            yield from export_batch(controller, intent, prompt)
            return
        if intent['target_plant']:
            yield event(2, 'running', 'Đang chuyển nhà máy ' + intent['target_plant'])
            result = controller.switch_factory(intent['target_plant'])
            require_success(result)
            yield event(2, 'success', result['message'], plant=intent['target_plant'])
        else:
            yield event(2, 'skipped', 'Giữ nhà máy hiện tại')
        if intent['target_page']:
            yield event(3, 'running', 'Đang mở ' + intent['target_page'].upper())
            result = controller.navigate_to(intent['target_page'])
            require_success(result)
            yield event(3, 'success', 'Đã mở ' + result['title'])
        else:
            yield event(3, 'skipped', 'Giữ trang hiện tại')
        files = []
        if intent['action'] or intent['nike'] is not None:
            yield event(4, 'running', 'Đang áp dụng bộ lọc và kiểm tra kết quả...')
            result = controller.apply_sfc_filters(
                date_from=intent['date_from'], date_to=intent['date_to'], factory=intent['target_line'],
                action=intent['action'], nike=intent['nike'])
            require_success(result)
            if result.get('file'):
                files = [result['file']]
            yield event(4, 'success', result.get('action_result') or 'Đã đồng bộ chế độ NIKE', files=files)
        else:
            yield event(4, 'skipped', 'Không có thao tác dữ liệu')
        yield event(5, 'done', 'Đã hoàn tất và kiểm tra các bước được yêu cầu.', files=files)
    except Exception as error:
        yield event(5, 'error', 'Chưa hoàn tất: ' + str(error))
    finally:
        COMMAND_LOCK.release()


def require_success(result):
    if not result.get('success'):
        raise RuntimeError(result.get('error') or 'Thao tác SFC thất bại')


def export_batch(controller, intent, prompt):
    from sfc_excel import combine_exports, file_info, validate_export
    ranges = periods(intent)
    plants = intent['target_plants']
    if not plants:
        plant = controller.get_live_status().get('current_plant')
        if plant not in PLANTS:
            raise ValueError('Không xác định được nhà máy hiện tại.')
        plants = [plant]
    run_id = datetime.now(timezone(timedelta(hours=7))).strftime('%Y%m%d_%H%M%S') + '_' + uuid4().hex[:6]
    folder = EXPORT_ROOT / run_id
    previous_scope = None
    # Resume only this exact failed request, revalidating every saved daily
    # source before reuse. This avoids downloading successful parts again.
    for candidate in sorted(EXPORT_ROOT.glob('*/manifest.json'), reverse=True):
        try:
            previous = json.loads(candidate.read_text(encoding='utf-8'))
            before = previous.get('intent', {})
            same_settings = all(before.get(key) == value for key, value in intent.items()
                                if key not in ('target_plants', 'target_plant', 'summary'))
            available = set(before.get('target_plants', [])) | set(previous.get('paused_plants', []))
            compatible_scope = bool(intent['target_plants']) and set(intent['target_plants']) <= available
            if (previous.get('status') in ('error', 'interrupted') and same_settings and compatible_scope
                    and (candidate.parent / '_parts').is_dir()):
                folder = candidate.parent
                previous_scope = previous
                break
        except (OSError, ValueError):
            continue
    folder.mkdir(parents=True, exist_ok=True)
    manifest = dict(prompt=prompt, intent=intent, status='running', monthly_files=[], combined_files=[])
    manifest['export_method'] = 'isolated_daily_v1'
    if previous_scope:
        available = set(previous_scope['intent']['target_plants']) | set(previous_scope.get('paused_plants', []))
        manifest['paused_plants'] = sorted(available - set(plants))
        manifest['original_prompt'] = previous_scope.get('original_prompt', previous_scope['prompt'])
    manifest_path = folder / 'manifest.json'

    def save_manifest():
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')

    save_manifest()
    count, total = 0, len(plants) * len(ranges)
    mode = 'NIKE' if intent['nike'] else 'Internal'
    yield event(2, 'running', f'Kế hoạch: {len(plants)} nhà máy × {len(ranges)} kỳ = {total} file', total=total)
    try:
        for plant in plants:
            yield event(2, 'running', f'Đang kết nối phiên xuất riêng cho {plant}...', plant=plant)
            session = controller.create_hfpa_session(plant)
            session.prepare()
            yield event(2, 'success', f'SFC đã xác nhận nhà máy {plant}', plant=plant)
            yield event(3, 'success', f'Đã kết nối HFPA · {plant} · {mode}')
            plant_files = []
            for first, last in ranges:
                yield event(4, 'running', f'[{count + 1}/{total}] Đang tải {plant}: {first} → {last} · {mode}', completed=count, total=total)
                downloaded = None
                cached = next((item for item in (previous_scope or {}).get('monthly_files', [])
                               if item.get('plant') == plant and item.get('date_from') == first
                               and item.get('date_to') == last and item.get('nike') is intent['nike']), None)
                if cached and Path(cached['path']).is_file():
                    verified = file_info(cached['path'])
                    if verified['sha256'] != cached['sha256']:
                        raise ValueError('File tháng đã thay đổi; dừng để kiểm tra: ' + cached['name'])
                    rows = validate_export(cached['path'], first, last, intent['nike'])
                    if rows != cached['rows']:
                        raise ValueError('Số dòng file tháng đã thay đổi: ' + cached['name'])
                    downloaded = cached
                    yield event(4, 'running', 'Đã kiểm tra lại file tháng đã lưu: ' + cached['name'])
                else:
                    for update in session.export_period(
                            first, last, intent['nike'],
                            folder / f'HFPA_{plant}_{mode}_{first}_{last}-{plant}', factory=intent['target_line'] or ''):
                        if update.get('message'):
                            yield event(4, 'running', f'[{count + 1}/{total}] ' + update['message'], completed=count, total=total)
                        if update.get('file'):
                            downloaded = update['file']
                if downloaded is None:
                    raise RuntimeError('Chưa có file tháng được kiểm tra; dừng lệnh xuất.')
                item = dict(downloaded, plant=plant, date_from=first, date_to=last, nike=intent['nike'])
                manifest['monthly_files'].append(item)
                plant_files.append(item)
                count += 1
                save_manifest()
                yield event(4, 'running', f'[{count}/{total}] Đã lưu và kiểm tra {item["name"]}', completed=count, total=total, nike=intent['nike'])
            if intent['combine']:
                yield event(4, 'running', f'Đang gộp {len(plant_files)} file của {plant}...')
                path = folder / f'HFPA_{plant}_{mode}_{intent["date_from"]}_{intent["date_to"]}_Combined-{plant}.xlsx'
                combined = combine_exports(plant_files, path, plant, intent['nike'])
                manifest['combined_files'].append(combined)
                save_manifest()
                yield event(4, 'success', f'Đã gộp {plant}: {combined["rows"]:,} dòng', files=[combined])
        manifest['status'] = 'done'
        save_manifest()
        files = manifest['combined_files'] or manifest['monthly_files']
        yield event(5, 'done', f'Hoàn tất {count}/{total} file nguồn' + (f' và {len(files)} file gộp theo nhà máy.' if intent['combine'] else '.'), files=files, completed=count, total=total)
    except BaseException as error:
        manifest['status'] = 'interrupted' if isinstance(error, GeneratorExit) else 'error'
        manifest['error'] = str(error)
        save_manifest()
        raise
