"""Portable Power Query package for results too large for worksheet export."""
from pathlib import Path
from uuid import uuid4
import csv
import hashlib
import json
import zipfile

import openpyxl
from openpyxl.workbook.defined_name import DefinedName
import spreadsheet_combine as combine
from spreadsheet_combine_search import display

CHUNK_ROWS=1_000_000


def formula(columns):
    quantity=next((name for name in columns if combine.key(name)=='number of pairs'),None)
    code='''let
    SourceFolder = Excel.CurrentWorkbook(){[Name="SourceFolder"]}[Content]{0}[Column1],
    Files = Folder.Files(SourceFolder),
    CSVFiles = Table.SelectRows(Files, each [Extension] = ".csv" and Text.StartsWith([Name], "Data_")),
    Sorted = Table.Sort(CSVFiles, {{"Name", Order.Ascending}}),
    Tables = List.Transform(Sorted[Content], each Table.PromoteHeaders(Csv.Document(_, [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), [PromoteAllScalars=true])),
    Combined = Table.Combine(Tables)'''
    if quantity:
        quoted='"'+quantity.replace('"','""')+'"'
        code+=',\n    Typed = Table.TransformColumns(Combined, {{'+quoted+', each if _ = null or _ = "" then null else Number.FromText(_, "en-US"), type number}})\nin\n    Typed\n'
    else:code+='\nin\n    Combined\n'
    return code


def export(job_id,payload,progress=None):
    folder=combine.folder_for(job_id);result_id=uuid4().hex
    result_folder=folder/result_id;result_folder.mkdir()
    output_name=payload['output_name'][:-5]+'.zip'
    output=result_folder/output_name
    manifest=dict(row_count=payload['row_count'],columns=payload['columns'],sources=payload['sources'],parts=[])
    instructions=('GOI POWER QUERY\n\n'
        '1. Giai nen toan bo goi vao mot thu muc co dinh.\n'
        '2. Mo Huong_dan.xlsx bang Microsoft Excel co Power Query (Excel 2016 tro len). WPS hien tai khong ho tro Power Query nay.\n'
        '3. Dien duong dan day du den thu muc Data da giai nen vao o B3.\n'
        '4. Data > Get Data > From Other Sources > Blank Query > Advanced Editor.\n'
        '5. Mo Combine.pq bang Notepad, copy toan bo noi dung vao Advanced Editor.\n'
        '6. Close & Load To > Only Create Connection. Neu can Pivot, chon Add this data to the Data Model. KHONG load toan bo vao worksheet.\n\n'
        'Du lieu CSV da duoc ghep theo ten cot va cac lua chon cua Dashboard; Power Query append cac phan trong thu muc Data. '
        'Grand Total va top issue co the tra cuu tren Dashboard ma khong can mo hang trieu dong trong Excel.\n')
    guide=result_folder/'Huong_dan.xlsx'
    book=openpyxl.Workbook();sheet=book.active;sheet.title='Huong dan'
    sheet['A1']='POWER QUERY · DỮ LIỆU LỚN';sheet['A3']='Đường dẫn thư mục Data';sheet['B3']=r'C:\CHON_THU_MUC_DA_GIAI_NEN\Data'
    sheet['A5']=f'Số dòng dữ liệu: {payload["row_count"]:,}'
    sheet['A7']='Dùng Microsoft Excel có Power Query. Chỉ tạo kết nối hoặc tải vào Data Model; không tải toàn bộ vào worksheet.'
    sheet['A8']='Xem HUONG_DAN.txt để nhập truy vấn Combine.pq. File này chỉ chứa hướng dẫn và ô cấu hình, chưa nhúng truy vấn.'
    sheet.column_dimensions['A'].width=40;sheet.column_dimensions['B'].width=100
    book.defined_names.add(DefinedName('SourceFolder',attr_text="'Huong dan'!$B$3"))
    book.save(guide);book.close()
    part=None;stream=None;written=0;part_rows=0;number=0
    try:
        with zipfile.ZipFile(output,'w',compression=zipfile.ZIP_DEFLATED,compresslevel=1,allowZip64=True) as archive:
            def finish():
                nonlocal stream,part
                if stream is None:return
                stream.close();stream=None
                with part.open('rb') as binary:digest=hashlib.file_digest(binary,'sha256').hexdigest()
                archive.write(part,'Data/'+part.name)
                manifest['parts'].append(dict(file=part.name,rows=part_rows,sha256=digest))
                part.unlink();part=None
            for written,row in enumerate(payload['rows'],1):
                if stream is None or part_rows>=CHUNK_ROWS:
                    finish();number+=1;part_rows=0
                    part=result_folder/f'Data_{number:04}.csv'
                    stream=part.open('w',encoding='utf-8-sig',newline='')
                    writer=csv.writer(stream);writer.writerow(payload['columns'])
                writer.writerow([display(cell) for cell in row]);part_rows+=1
                if progress and written%50000==0:progress(f'Gói Power Query: {written:,}/{payload["row_count"]:,} dòng')
            finish()
            if written!=payload['row_count']:raise ValueError('Số dòng gói Power Query không khớp nguồn.')
            archive.writestr('Combine.pq',formula(payload['columns']).encode('utf-8-sig'))
            archive.writestr('HUONG_DAN.txt',instructions.encode('utf-8-sig'))
            archive.writestr('manifest.json',json.dumps(manifest,ensure_ascii=False,indent=2).encode('utf-8'))
            archive.write(guide,guide.name)
        if progress:progress('Kiểm tra gói Power Query đã xuất…')
        with zipfile.ZipFile(output) as archive:
            if archive.testzip() is not None:raise ValueError('Gói Power Query có phần dữ liệu bị lỗi.')
    except Exception:
        if stream:stream.close()
        if part:part.unlink(missing_ok=True)
        output.unlink(missing_ok=True)
        raise
    result=combine.summary(payload)|dict(id=job_id,result_id=result_id,output_name=output_name,
        kind='power_query',download_url=f'/api/combine/{job_id}/{result_id}/download')
    combine.json_write(result_folder/'result.json',result)
    return result
