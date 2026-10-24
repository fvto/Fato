from pathlib import Path
exec((Path(__file__).parent/'inspect_sfc_http_form.py').read_text(encoding='utf8').split('for k,v in form.fields.items():')[0])
import re
urls=re.findall(r'<script[^>]+src="([^"]+DXR.axd[^"]+)"',body)
for url in urls:
 with opener.open('https://sfc.chingluh.com'+url.replace('&amp;','&'),timeout=15) as response:script=response.read().decode('utf8')
 for term in ['_Raw','GetRawValue','GetDateTimeInputValue','SetDateInternal']:
  matches=list(re.finditer(re.escape(term),script))
  if matches:
   print('TERM',term)
   for m in matches[:4]:print(script[max(0,m.start()-150):m.start()+400].replace('\r','').replace('\n',' ')[:550])
