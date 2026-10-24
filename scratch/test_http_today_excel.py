from pathlib import Path
source=(Path(__file__).parent/'test_http_today_nike.py').read_text(encoding='utf8')
source=source.replace("$btnSearch'","$btnExcel'").replace("data=response.read();print(","data=response.read();(root/'scratch/today_export_response.bin').write_bytes(data);print(").replace('Search today NIKE','Export today NIKE')
exec(compile(source,'scratch/test_http_today_excel.py','exec'))
