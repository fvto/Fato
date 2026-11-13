"""Disk-backed column filtering and exact quantity totals for combined data."""
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import json
import hashlib
import re
import sqlite3
import unicodedata
from uuid import uuid4

from openpyxl.utils.datetime import from_excel
import spreadsheet_combine as combine

PAGE_SIZE = 100
MAX_FILTERS = 30


def normalized(value):
    return unicodedata.normalize('NFKC', str(value or '')).strip().casefold()


def normalized_size(value):
    text=normalized(value).strip()
    if re.fullmatch(r'[+-]?\d+(?:\.\d+)?',text):
        number=Decimal(text)
        return format(number.normalize(),'f') if number else '0'
    return text


def normalized_code(value):
    return ''.join(normalized(value).split())


def size_range_label(value):
    return str(value or '').strip()


def normalized_range(value):
    text=''.join(normalized(size_range_label(value)).split())
    match=re.fullmatch(r'(\d+(?:\.\d+)?)(.+)',text)
    return normalized_size(match.group(1))+match.group(2) if match else text


def display(cell):
    value, fmt = cell['value'], cell.get('format', '')
    if value is None:
        return None
    if type(value) in (int, float) and 'yyyy-mm-dd' in fmt.lower():
        converted = from_excel(value)
        if isinstance(converted, (datetime,date)):
            if isinstance(converted,datetime):
                return converted.isoformat(sep=' ',timespec='seconds') if 'hh:mm' in fmt.lower() else converted.date().isoformat()
            return converted.isoformat()
    if type(value) in (int,float) and re.fullmatch(r'0(?:\.0+)?%',fmt):
        decimals=len(fmt.split('.')[1].rstrip('%')) if '.' in fmt else 0
        return f'{value*100:.{decimals}f}%'
    return str(value)


def amount(value):
    if value is None or not str(value).strip():
        return None, False
    text=str(value).strip()
    if not re.fullmatch(r'[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?',text):
        return None, True
    try:
        return str(Decimal(text.replace(',',''))),False
    except InvalidOperation:
        return None,True


class DecimalSum:
    def __init__(self):self.total=Decimal(0)
    def step(self,value):
        if value is not None:self.total+=Decimal(value)
    def finalize(self):return str(self.total)


def create(job_id,selections,options,progress=None):
    payload=combine.prepare(job_id,selections,options,progress)
    columns=payload['columns']
    if len(columns)>300:
        raise ValueError('Bảng tra cứu hỗ trợ tối đa 300 cột. Hãy chọn các bảng có ít cột hơn.')
    keys=[combine.key(column) for column in columns]
    if 'number of pairs' not in keys or 'size' not in keys:
        raise ValueError('Cần có cột Size và Number of Pairs để tính tổng số đôi theo size.')
    quantity,size=keys.index('number of pairs'),keys.index('size')
    size_range=keys.index('size range') if 'size range' in keys else None
    indexed={i for i,key in enumerate(keys) if key in {'mold code','size'}}
    folder=combine.folder_for(job_id)
    fingerprint=hashlib.sha256(json.dumps(dict(version=5,selections=selections,options=options,sources=payload['sources']),
        sort_keys=True,ensure_ascii=False).encode('utf-8')).hexdigest()
    for previous in folder.glob('search-*.json'):
        try:
            old=json.loads(previous.read_text(encoding='utf-8'))
            old_id=old.get('id','')
            if (old.get('fingerprint')==fingerprint and re.fullmatch('[a-f0-9]{32}',old_id)
                    and (folder/f'search-{old_id}.sqlite').is_file() and 'initial' in old):
                if progress:progress('Đã kiểm tra nguồn và dùng lại bảng tra cứu.')
                return old['initial']
        except (ValueError,OSError,TypeError):
            continue
    index_id=uuid4().hex
    temporary=folder/f'search-{index_id}.building.sqlite'
    target=folder/f'search-{index_id}.sqlite'
    marker=folder/f'search-{index_id}.json'
    connection=None
    try:
        connection=sqlite3.connect(temporary)
        connection.execute('PRAGMA journal_mode=OFF')
        connection.execute('PRAGMA synchronous=OFF')
        connection.execute('PRAGMA temp_store=FILE')
        connection.execute('PRAGMA cache_size=-65536')
        fields=[f'c{i} TEXT' for i in range(len(columns))]
        fields += [f'k{i} TEXT' for i in sorted(indexed)]
        fields += ['quantity TEXT','invalid INTEGER NOT NULL']
        connection.execute('CREATE TABLE entries (id INTEGER PRIMARY KEY,'+','.join(fields)+')')
        sql='INSERT INTO entries VALUES ('+','.join('?' for _ in range(1+len(fields)))+')'
        batch=[]; count=0; grand_total=Decimal(0); invalid_count=0; summary={}; omitted_sizes=False
        for count,row in enumerate(payload['rows'],1):
            values=[display(cell) for cell in row]
            numeric,invalid=amount(row[quantity]['value'])
            invalid_count+=int(invalid)
            if numeric is not None:grand_total+=Decimal(numeric)
            size_value=values[size] or '(Trống)'
            range_value=values[size_range] if size_range is not None else None
            size_key=(normalized_size(values[size]),normalized_range(range_value))
            if size_key in summary or len(summary)<1000:
                group=summary.setdefault(size_key,dict(size=size_value,size_range=size_range_label(range_value),rows=0,total=Decimal(0),invalid=0))
                group['rows']+=1;group['invalid']+=int(invalid)
                if numeric is not None:group['total']+=Decimal(numeric)
            else:omitted_sizes=True
            batch.append([count,*values,*[(normalized_size(values[i]) if i==size else normalized_code(values[i])) for i in sorted(indexed)],numeric,int(invalid)])
            if len(batch)>=2000:
                connection.executemany(sql,batch);batch.clear()
            if progress and count%50000==0:
                progress(f'Chuẩn bị tra cứu: {count:,}/{payload["row_count"]:,} dòng')
        if batch:connection.executemany(sql,batch)
        if count!=payload['row_count']:
            raise ValueError('Số dòng tra cứu không khớp dữ liệu đã chọn.')
        for index in sorted(indexed):
            if progress:progress(f'Tạo chỉ mục tra cứu cột {columns[index]}…')
            connection.execute(f'CREATE INDEX idx_k{index} ON entries(k{index})')
        connection.commit();connection.close();connection=None
        # Recheck all input hashes before publishing the searchable snapshot.
        combine.prepare(job_id,selections,options)
        temporary.replace(target)
        groups=[dict(group,total=str(group['total'])) for _,group in sorted(summary.items())]
        initial=dict(index_id=index_id,columns=columns,row_count=count,quantity_column=columns[quantity],
                    size_column=columns[size],range_column=columns[size_range] if size_range is not None else None,
                    grand_total=str(grand_total),invalid_quantities=invalid_count,
                    groups=groups,omitted_sizes=omitted_sizes)
        meta=dict(id=index_id,columns=columns,quantity=quantity,size=size,size_range=size_range,indexed=sorted(indexed),
                  unit_type=keys.index('unit type') if 'unit type' in keys else None,
                  issue=keys.index('issue') if 'issue' in keys else None,
                  row_count=count,sources=payload['sources'],warnings=payload.get('warnings',[]),
                  fingerprint=fingerprint,initial=initial)
        combine.json_write(marker,meta)
        return initial
    except Exception:
        if connection:connection.close()
        temporary.unlink(missing_ok=True)
        target.unlink(missing_ok=True)
        marker.unlink(missing_ok=True)
        raise


def query(job_id,index_id,filters,page=1,progress=None):
    if not isinstance(index_id,str) or not re.fullmatch('[a-f0-9]{32}',index_id):
        raise ValueError('Bảng tra cứu không hợp lệ. Hãy tạo lại.')
    folder=combine.folder_for(job_id)
    marker=folder/f'search-{index_id}.json'
    target=folder/f'search-{index_id}.sqlite'
    if not marker.is_file() or not target.is_file():raise ValueError('Bảng tra cứu không còn tồn tại. Hãy tạo lại.')
    meta=json.loads(marker.read_text(encoding='utf-8'))
    if not isinstance(filters,list) or len(filters)>MAX_FILTERS or type(page) is not int or not 1<=page<=1000000:
        raise ValueError('Bộ lọc hoặc trang không hợp lệ.')
    clauses=[];arguments=[]
    for item in filters:
        if not isinstance(item,dict) or type(item.get('column')) is not int or not 0<=item['column']<len(meta['columns']) or item.get('mode') not in {'contains','exact'} or not isinstance(item.get('value'),str) or not item['value'].strip() or len(item['value'])>200:
            raise ValueError('Bộ lọc không hợp lệ.')
        index=item['column']
        value=(normalized_size(item['value']) if index==meta['size'] else
               normalized_code(item['value']) if index in meta['indexed'] else normalized(item['value'].strip()))
        expression=f'k{index}' if index in meta['indexed'] else f'fold(c{index})'
        clauses.append(f'{expression} = ?' if item['mode']=='exact' else f'instr({expression}, ?) > 0')
        arguments.append(value)
    where=' WHERE '+' AND '.join(clauses) if clauses else ''
    connection=sqlite3.connect(f'{target.as_uri()}?mode=ro',uri=True)
    connection.create_function('fold',1,normalized,deterministic=True)
    connection.create_function('range_key',1,normalized_range,deterministic=True)
    connection.create_function('range_label',1,size_range_label,deterministic=True)
    connection.create_aggregate('decimal_sum',1,DecimalSum)
    try:
        if progress:progress('Đang tính tổng số đôi và tổng theo size…')
        groups=[];total=Decimal(0);matched=0;invalid=0
        size_range=meta.get('size_range')
        range_select=f'MIN(range_label(c{size_range}))' if size_range is not None else "''"
        range_group=f',range_key(c{size_range})' if size_range is not None else ''
        cursor=connection.execute(f'SELECT MIN(c{meta["size"]}),{range_select},COUNT(*),decimal_sum(quantity),SUM(invalid) FROM entries{where} GROUP BY k{meta["size"]}{range_group} ORDER BY k{meta["size"]}{range_group}',arguments)
        omitted_sizes=False
        for size,size_range,count,subtotal,bad in cursor:
            matched+=count;invalid+=bad or 0
            value=Decimal(subtotal or '0');total+=value
            if len(groups)<1000:groups.append(dict(size=size or '(Trống)',size_range=size_range or '',rows=count,total=str(value),invalid=bad or 0))
            else:omitted_sizes=True
        top_issues=[];top_issues_truncated=False
        available=meta.get('unit_type') is not None and meta.get('issue') is not None
        if available:
            if progress:progress('Đang tính top 3 Issue cho Unit Type = DRR…')
            issue=meta['issue'];unit=meta['unit_type']
            drr_where=where+(' AND ' if where else ' WHERE ')+f"fold(c{unit}) = 'drr' AND COALESCE(trim(c{issue}),'') <> '' AND quantity IS NOT NULL"
            cursor=connection.execute(f'SELECT MIN(c{meta["size"]}),{range_select},MIN(c{issue}),COUNT(*),decimal_sum(quantity) FROM entries{drr_where} GROUP BY k{meta["size"]}{range_group},fold(c{issue})',arguments)
            ranked={}
            for size,range_value,label,count,subtotal in cursor:
                key=(normalized_size(size),normalized_range(range_value))
                if key not in ranked and len(ranked)>=1000:
                    top_issues_truncated=True
                    continue
                entries=ranked.setdefault(key,[])
                entries.append(dict(size=size or '(Trống)',size_range=range_value or '',issue=label,rows=count,total=subtotal or '0'))
                entries.sort(key=lambda item:(-Decimal(item['total']),normalized(item['issue'])))
                del entries[3:]
            for key,entries in sorted(ranked.items()):
                for rank,entry in enumerate(entries,1):top_issues.append(dict(entry,rank=rank))
        if progress:progress('Đang tải các dòng phù hợp…')
        fields=','.join(f'c{i}' for i in range(len(meta['columns'])))
        rows=connection.execute(f'SELECT {fields} FROM entries{where} ORDER BY id LIMIT ? OFFSET ?',arguments+[PAGE_SIZE,(page-1)*PAGE_SIZE]).fetchall()
    finally:connection.close()
    return dict(columns=meta['columns'],rows=rows,page=page,page_size=PAGE_SIZE,matched=matched,
                grand_total=str(total),invalid_quantities=invalid,groups=groups,omitted_sizes=omitted_sizes,
                top_issues=top_issues,top_issues_available=available,top_issues_truncated=top_issues_truncated,
                quantity_column=meta['columns'][meta['quantity']],size_column=meta['columns'][meta['size']],
                range_column=meta['columns'][meta['size_range']] if meta.get('size_range') is not None else None)
