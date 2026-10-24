from pathlib import Path
exec((Path(__file__).parent/'inspect_sfc_http_form.py').read_text(encoding='utf-8').split('with opener.open')[0])
with opener.open('https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx',timeout=30) as response:body=response.read().decode('utf8')
import re
for term in ['menuRightTop','DXCurrentFactory','VH2','JV2','SelectedFactory','function on','ItemClick']:
 print('ITEM',term)
 for m in list(re.finditer(re.escape(term),body))[:4]:print(body[max(0,m.start()-200):m.start()+1000].replace('\n',' ')[:1200])
