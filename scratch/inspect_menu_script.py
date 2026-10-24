from pathlib import Path
exec((Path(__file__).parent/'inspect_sfc_factory.py').read_text(encoding='utf8').split('import re')[0])
import re
url=re.findall(r'<script[^>]+src="([^"]+DXR.axd[^"]+)"',body)[0].replace('&amp;','&')
with opener.open('https://sfc.chingluh.com'+url,timeout=15) as response:script=response.read().decode('utf8')
for term in ['SendPostBack', 'ProcessItemClick', 'DXCurrentFactoryCookieName']:
 print('TERM',term)
 for m in list(re.finditer(re.escape(term),script))[-8:]:print(script[max(0,m.start()-150):m.start()+300])
