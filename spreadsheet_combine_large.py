"""Direct CSV ingestion and bounded-memory combine for large source tables."""
from contextlib import contextmanager
from datetime import date, datetime
from itertools import chain, islice
from pathlib import Path
import hashlib
import json
import math
import re
import zipfile
from uuid import uuid4
import xlrd

import spreadsheet_combine_csv as csv_input
from spreadsheet_combine_xlsx import Source, Writer, canonical, MAX_ROWS
from spreadsheet_combine_stream import SOURCE_HEADERS, source_values


def inspect_csv(path, config=None, progress=None, cache_scan=False):
    import spreadsheet_combine as combine
    settings = csv_input.configuration(path, config)
    count, cells, width, found, read_error = 0, 0, 0, False, None
    def observed():
        nonlocal count, cells, width, found, read_error
        try:
            for row in csv_input.iter_records(path, settings):
                count += 1
                width = max(width, len(row))
                cells += max(1, len(row))
                found = found or any(value not in (None, '') for value in row)
                if progress and count % 50000 == 0:
                    progress(f'Đọc CSV: {path.name} · {count:,} dòng')
                yield row, ['General'] * len(row)
        except (ValueError, UnicodeError) as error:
            read_error = error
            raise
    iterator = observed()
    try:
        sample = list(islice(iterator, 55))
        first = combine.detect_header(dict(rows=[row for row, _ in sample]))
        scan, stats = None, {}
        if cache_scan and sample:
            try:
                plan = dict(sheet='CSV', header_row=first, footer_index=None,
                            **header_layout(sample[first-1][0]))
                scan = scan_records(filtered_records(chain(sample, iterator), plan, stats), plan)
            except ValueError:
                # A user can correct a header/range after upload. Actual CSV
                # parse/encoding errors must still fail the upload immediately.
                if read_error is not None:
                    raise read_error
        for _ in iterator:
            pass
    finally:
        iterator.close()
    if not found:
        raise ValueError('Không có dữ liệu.')
    if scan is not None and math.isfinite(scan['defect_sum']):
        combine.json_write(path.with_suffix(path.suffix+'.scan.json'),
            dict(version=1, header_row=first, end_row=count, scan=scan, stats=stats))
    return dict(stored=path.name, sheets=[dict(name='CSV', rows=count, columns=width,
                header_row=first, default_selected=True)],
                csv={key:settings[key] for key in ('encoding','delimiter')},
                csv_config=settings, csv_options=csv_input.options(config),
                warnings=settings['warnings'], cells=cells,
                large=cells > combine.MAX_CELLS or path.stat().st_size > 2 * 1024**2)


def inspect_xlsx(path):
    import spreadsheet_combine as combine
    with path.open('rb') as stream:
        signature=stream.read(8)
    if signature==b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        book=xlrd.open_workbook(str(path),formatting_info=True,on_demand=True)
        try:
            result=[]
            for sheet in book.sheets():
                sample=[sheet.row_values(i) for i in range(min(55,sheet.nrows))]
                if sheet.nrows and any(any(v not in (None,'') for v in row) for row in sample):
                    result.append(dict(name=sheet.name,rows=sheet.nrows,columns=sheet.ncols,
                        header_row=combine.detect_header(dict(rows=sample))))
            if result:
                largest=max(result,key=lambda sheet:sheet['rows'])
                for sheet in result:sheet['default_selected']=sheet is largest
            return result
        finally:book.release_resources()
    source = Source(path)
    result = []
    try:
        for info in source.sheets:
            iterator = source.rows(info['name'], 55)
            try:
                sample = [row for row, _ in islice(iterator, 55)]
            finally:
                iterator.close()
            if info['rows'] == 0:
                # Dimension-less workbooks are valid. Count only when needed;
                # conversion-generated workbooks already provide their counts.
                count, width = 0, 0
                for row, _ in source.rows(info['name']):
                    count += 1
                    width = max(width, len(row))
            else:
                count, width = info['rows'], info['columns']
            if count and any(any(v not in (None, '') for v in row) for row in sample):
                result.append(dict(name=info['name'], rows=count, columns=width,
                                   header_row=combine.detect_header(dict(rows=sample))))
    finally:
        source.close()
    if result:
        largest=max(result,key=lambda sheet:sheet['rows'])
        for sheet in result:
            sheet['default_selected']=sheet is largest
    return result


@contextmanager
def open_source(folder, item):
    path = folder / item['stored']
    with path.open('rb') as stream:
        signature=stream.read(8)
    if signature.startswith(b'PK'):
        source = Source(path)
        for meta in item['sheets']:
            info = next((s for s in source.sheets if s['name'] == meta['name']), None)
            if info is not None and not info['rows']:
                info.update(rows=meta['rows'], columns=meta['columns'])
        try:
            yield source.rows
        finally:
            source.close()
    else:
        import spreadsheet_combine as combine
        if path.suffix.lower() == '.csv':
            config = item.get('csv_config') or csv_input.configuration(path, item.get('csv_options'))
            def rows(name, end=None):
                for i, values in enumerate(csv_input.iter_records(path, config), 1):
                    if end is not None and i > end:
                        break
                    yield values, ['General'] * len(values)
        elif signature==b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
            book=xlrd.open_workbook(str(path),formatting_info=True,on_demand=True)
            def rows(name, end=None):
                sheet=book.sheet_by_name(name)
                for index in range(min(end,sheet.nrows) if end is not None else sheet.nrows):
                    values,formats=[],[]
                    for cell in sheet.row(index):
                        value=cell.value
                        if cell.ctype==xlrd.XL_CELL_DATE:value=xlrd.xldate_as_datetime(value,book.datemode)
                        elif cell.ctype==xlrd.XL_CELL_BOOLEAN:value=bool(value)
                        elif cell.ctype==xlrd.XL_CELL_ERROR:raise ValueError(f'{name}: có ô lỗi Excel.')
                        values.append(value)
                        formats.append(book.format_map[book.xf_list[cell.xf_index].format_key].format_str)
                    yield values,formats
            try:yield rows
            finally:book.release_resources()
            return
        else:
            raise ValueError('Không đọc được cấu trúc file nguồn.')
        yield rows


def header_plan(folder, item, selection):
    import spreadsheet_combine as combine
    meta = next((s for s in item['sheets'] if s['name'] == selection.get('sheet')), None)
    if meta is None:
        raise ValueError('Trang tính bị trùng hoặc không còn tồn tại.')
    first = selection.get('header_row', meta['header_row'])
    last = selection.get('end_row', meta['rows'])
    if type(first) is not int or not 1 <= first <= meta['rows']:
        raise ValueError('Dòng tiêu đề không hợp lệ.')
    if type(last) is not int or not first <= last <= meta['rows']:
        raise ValueError('Dòng cuối phải nằm từ dòng tiêu đề đến dòng cuối trang tính.')
    with open_source(folder, item) as rows:
        iterator = rows(meta['name'], first)
        try:
            raw = next(islice(iterator,first-1,None), ([],[]))[0]
        finally:
            iterator.close()
    return dict(item=item, sheet=meta['name'], header_row=first, end_row=last,
                offset=meta.get('source_row_offset',0), footer_index=None, **header_layout(raw))


def header_layout(raw):
    import spreadsheet_combine as combine
    # Delegate header normalization/duplicates to the original implementation.
    table = combine.clean_table(dict(name='Header',rows=[raw],formats=[['General']*len(raw)]),1,footer_mode='keep')
    header = list(raw)
    while header and header[-1] in (None,''):
        header.pop()
    return dict(header=table['header'], keys=table['keys'], width=len(header),
                positions=[i for i,value in enumerate(header) if value not in (None,'')],
                unnamed=[i for i,value in enumerate(header) if value in (None,'')])


def records(folder, plan, counters=None):
    with open_source(folder, plan['item']) as source_rows:
        yield from filtered_records(source_rows(plan['sheet'],plan['end_row']), plan, counters)


def filtered_records(source_rows, plan, counters=None):
    import spreadsheet_combine as combine
    counters = counters if counters is not None else {}
    counters.update(blank_rows=0, repeated_headers=0)
    for number, (row, formats) in enumerate(source_rows,1):
        if number <= plan['header_row']:
            continue
        if not any(value not in (None,'') for value in row):
            counters['blank_rows'] += 1
            continue
        if any(value not in (None,'') for value in row[plan['width']:]):
            raise ValueError(f'{plan["sheet"]}, dòng {number}: dữ liệu vượt số cột tiêu đề.')
        if any(i<len(row) and row[i] not in (None,'') for i in plan['unnamed']):
            raise ValueError(f'{plan["sheet"]}, dòng {number}: có dữ liệu dưới cột chưa có tên.')
        if not plan['unnamed'] and len(row)==plan['width']:
            values, styles = row, formats
        else:
            values = [row[i] if i<len(row) else None for i in plan['positions']]
            styles = [formats[i] if i<len(formats) else 'General' for i in plan['positions']]
        if (combine.key(values[0]) == plan['keys'][0]
                and [combine.key(value) for value in values] == plan['keys']):
            counters['repeated_headers'] += 1
            continue
        if number == plan['footer_index']:
            continue
        yield number, values, styles


AUDIT_KEYS = {'factory code','so#','audit date',"defective q'ty(pair)","total defect q'ty"}


def scan_records(iterator, plan, progress=None):
    count, tail, defect_sum, invalid_sum, preview = 0, None, 0, False, []
    defect = plan['keys'].index("defective q'ty(pair)") if AUDIT_KEYS.issubset(plan['keys']) else None
    for record in iterator:
        if tail and defect is not None:
            try: defect_sum += float(tail[1][defect] or 0)
            except (TypeError, ValueError): invalid_sum = True
        tail = record
        count += 1
        if len(preview) < 13:
            preview.append(record)
        if progress and count % 50000 == 0:
            progress(count)
    return dict(count=count, tail=tail, defect_sum=defect_sum,
                invalid_sum=invalid_sum, preview=preview)


def encode(values, formats, keys):
    import spreadsheet_combine as combine
    result = []
    for value, fmt, label in zip(values, formats, keys):
        if (isinstance(value,str) and value.strip().endswith('%')
                and re.search(r'(?:%|\brate\b|\bpercent(?:age)?\b|tỷ lệ)',label)
                and re.fullmatch(r'[+-]?\d+(?:\.\d+)?%',value.strip())):
            raw=value.strip()[:-1]
            precision=len(raw.split('.')[1]) if '.' in raw else 0
            value,fmt=float(raw)/100,'0'+('.'+'0'*precision if precision else '')+'%'
        result.append(combine.encode_cell(value,fmt))
    return result


class Rows:
    def __init__(self, folder, plans, keys, count, preview):
        self.folder, self.plans, self.keys, self.count, self.preview = folder,plans,keys,count,preview

    def __len__(self):
        return self.count

    def __getitem__(self, selection):
        return self.preview[selection]

    def mapped(self, plan, number, values, formats):
        columns = [dict(value=None,format='General') for _ in self.keys]
        for index, cell in zip(plan['indexes'],encode(values,formats,plan['keys'])):
            columns[index]=cell
        columns += [dict(value=value,format='General') for value in (plan['item']['name'],plan['sheet'],number+plan['offset'])]
        return columns

    def __iter__(self):
        for plan in self.plans:
            for number, values, formats in records(self.folder,plan):
                yield self.mapped(plan,number,values,formats)


def prepare(job_id, selections, options, progress=None):
    import spreadsheet_combine as combine
    folder=combine.folder_for(job_id)
    items=json.loads((folder/'inputs.json').read_text(encoding='utf-8'))
    if not isinstance(selections,list) or not selections:
        raise ValueError('Chọn ít nhất một bảng để gộp.')
    plans, columns, keys, seen, verified = [],[],[],set(),set()
    for selection in selections:
        if not isinstance(selection,dict) or type(selection.get('file_id')) is not int or not 0 <= selection['file_id'] < len(items):
            raise ValueError('File được chọn không hợp lệ.')
        item=items[selection['file_id']]
        try:
            identity=(selection['file_id'],selection.get('sheet'))
            if identity in seen:
                raise ValueError('Trang tính bị trùng hoặc không còn tồn tại.')
            seen.add(identity)
            if selection['file_id'] not in verified:
                for stored,digest in [(item['stored'],item.get('stored_sha256',item['sha256']))]+([(item['original_stored'],item['sha256'])] if item.get('converted') else []):
                    with (folder/stored).open('rb') as stream:
                        if hashlib.file_digest(stream,'sha256').hexdigest()!=digest:
                            raise ValueError('File nguồn đã thay đổi, hãy tải lại.')
                verified.add(selection['file_id'])
            plan=header_plan(folder,item,selection)
            if options.get('schema_mode')=='strict' and plans and set(plan['keys'])!=set(plans[0]['keys']):
                raise ValueError('Bộ cột khác bảng đầu tiên. Chọn Giữ tất cả cột nếu muốn ghép khác cấu trúc.')
            for label,key in zip(plan['header'],plan['keys']):
                if key not in keys:
                    keys.append(key);columns.append(label)
            plans.append(plan)
        except (ValueError,OSError,zipfile.BadZipFile) as error:
            raise ValueError(f'{item["name"]}: {error}') from error
    if len(columns)+3>16384:
        raise ValueError('Bảng gộp vượt giới hạn 16.384 cột của Excel.')
    fingerprint=hashlib.sha256(json.dumps(dict(selections=selections,options=options,
        sources=[(items[i]['sha256'],items[i].get('stored_sha256')) for i in sorted(verified)]),
        sort_keys=True,ensure_ascii=False).encode('utf-8')).hexdigest()
    cache=folder / f'large-preview-{fingerprint}.json'
    if cache.is_file():
        saved=json.loads(cache.read_text(encoding='utf-8'))
        cached_plans=saved.pop('_plans')
        cached_keys=saved.pop('_keys')
        cached_preview=saved.pop('_preview')
        saved['rows']=Rows(folder,cached_plans,cached_keys,saved['row_count'],cached_preview)
        if progress: progress('Đã đối chiếu nguồn; dùng lại bản xem trước đã kiểm tra.')
        return saved
    sources, preview_records, total = [],[],0
    for plan in plans:
        item=plan['item']
        stats, scan = {}, None
        recognized=AUDIT_KEYS.issubset(plan['keys'])
        defect=plan['keys'].index("defective q'ty(pair)") if recognized else None
        try:
            # Hashes were verified above. Reuse only a full-source scan for
            # the same header; custom ranges/headers still run the cleaner.
            if item['stored'].lower().endswith('.csv'):
                scan_path=(folder/item['stored']).with_suffix('.csv.scan.json')
                if scan_path.is_file():
                    cached=json.loads(scan_path.read_text(encoding='utf-8'))
                    if (cached.get('version')==1 and cached.get('header_row')==plan['header_row']
                            and cached.get('end_row')==plan['end_row']):
                        scan, stats=cached['scan'],cached['stats']
                        if progress:progress(f'Dùng kết quả đã kiểm tra khi tải lên: {item["name"]}')
            if scan is None:
                def update(count):
                    if progress:progress(f'Kiểm tra {item["name"]} / {plan["sheet"]}: {count:,} dòng')
                scan=scan_records(records(folder,plan,stats),plan,update)
            count,tail=scan['count'],scan['tail']
            defect_sum,invalid_sum=scan['defect_sum'],scan['invalid_sum']
            for record in scan['preview']:
                if len(preview_records)<12+len(plans):
                    preview_records.append((plan,*record))
            footer=0
            if recognized and tail and options.get('footer_mode','auto')=='auto':
                audit=plan['keys'].index('audit date')
                if all(value in (None,'') for i,value in enumerate(tail[1]) if i not in (audit,defect)):
                    try:
                        valid=not invalid_sum and float(str(tail[1][audit]).replace(',',''))==count-1 and math.isclose(float(str(tail[1][defect]).replace(',','')),defect_sum,abs_tol=1e-9)
                    except (TypeError,ValueError): valid=False
                    if not valid:
                        raise ValueError('Dòng tổng FGWH/SFC không khớp số dòng hoặc số đôi lỗi. Kiểm tra nguồn hoặc chọn giữ dòng tổng.')
                    footer=1;plan['footer_index']=tail[0]
            plan['indexes']=[keys.index(key) for key in plan['keys']]
            count-=footer;total+=count
            sources.append(dict(file=item['name'],sheet=plan['sheet'],rows=count,columns=len(plan['header']),
                **stats,footer_rows=footer,sha256=item['sha256'],header_row=plan['header_row'],end_row=plan['end_row'],
                profile='SFC audit' if recognized and tail and options.get('footer_mode','auto')=='auto' else '',
                missing_columns=[columns[i] for i,key in enumerate(keys) if key not in plan['keys']]))
        except (ValueError,OSError) as error:
            raise ValueError(f'{item["name"]}: {error}') from error
    if not total:
        raise ValueError('Không có dòng dữ liệu sau tiêu đề để gộp.')
    provenance=[]
    for label in combine.SOURCE_COLUMNS:
        candidate,number=label,1
        while combine.key(candidate) in keys or candidate in provenance:
            candidate=f'{label} (combine {number})';number+=1
        provenance.append(candidate)
    rows=Rows(folder,plans,keys,total,[])
    rows.preview=[rows.mapped(plan,number,values,formats) for plan,number,values,formats in preview_records if number!=plan['footer_index']][:12]
    warnings=[]
    if len(plans)>1 and not set.intersection(*(set(plan['keys']) for plan in plans)):
        warnings.append('Các bảng không có cột chung. Kiểm tra lại trang tính và dòng tiêu đề trước khi xuất.')
    for index in verified:
        warnings.extend(f'{items[index]["name"]}: {warning}' for warning in items[index].get('warnings',[]))
    if total>combine.XLSX_ROW_THRESHOLD:
        warnings.append('Kết quả trên 1 triệu dòng: chỉ xuất gói Power Query, không xuất toàn bộ dữ liệu thành sheet XLSX.')
    result=dict(large=True,columns=columns+provenance,rows=rows,sources=sources,source_columns=provenance,
        **combine.naming_details(sources),warnings=warnings,options=options,row_count=total,
        column_count=len(columns)+3,file_count=len(verified),footer_rows=sum(s['footer_rows'] for s in sources),
        repeated_headers=sum(s['repeated_headers'] for s in sources))
    saved={key:value for key,value in result.items() if key!='rows'}
    saved.update(_plans=plans,_keys=keys,_preview=rows.preview)
    temporary=cache.with_name(cache.name+'.'+uuid4().hex+'.tmp')
    try:
        combine.json_write(temporary,saved)
        temporary.replace(cache)
    finally:
        temporary.unlink(missing_ok=True)
    return result


def write_verify(path,payload,progress=None):
    writer=Writer(path)
    number=1
    writer.sheet('Combined',payload['columns'])
    try:
        for index,row in enumerate(payload['rows'],1):
            if writer.current['rows']>=MAX_ROWS:
                number+=1;writer.sheet(f'Combined_{number}',payload['columns'])
            writer.row(row)
            if progress and index%25000==0:
                progress(f'Xuất XLSX: {index:,}/{payload["row_count"]:,} dòng')
        writer.sheet('Sources',SOURCE_HEADERS)
        for source in payload['sources']:
            writer.row(source_values(source))
        writer.close()
    except Exception:
        writer.abort();path.unlink(missing_ok=True);raise
    # The output can contain many individually bounded source workbooks.
    # Its ZIP was authored here; verify every generated row without an input
    # archive-size limit on the aggregate output.
    source=Source(path,max_unpacked=None)
    try:
        for expected in writer.sheets:
            digest=hashlib.sha256();count=0
            if progress: progress(f'Kiểm tra file đã xuất: {expected["name"]}')
            for values,_ in source.rows(expected['name']):
                values+= [None]*(expected['columns']-len(values))
                digest.update(canonical(values));count+=1
            if count!=expected['rows'] or digest.digest()!=expected['digest'].digest():
                raise ValueError(f'Kiểm tra sau xuất không khớp tại sheet {expected["name"]}.')
    finally:
        source.close()
    return [sheet['name'] for sheet in writer.sheets if sheet['name']!='Sources']
