from pathlib import Path
exec((Path(__file__).parent/'inspect_sfc_http_form.py').read_text(encoding='utf-8').split('for k,v in form.fields.items():')[0])
import re
for term in ['menuRightTop','MyFactoryName','SelectedFactory','SFC_Factory','btnExcel','btnSearch']:
 print('\nITEM',term)
 matches=list(re.finditer(re.escape(term),body))
 for m in matches[:7]:print(body[max(0,m.start()-130):m.start()+500].replace('\n',' ')[:650])
