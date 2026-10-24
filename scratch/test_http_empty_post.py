from pathlib import Path
exec((Path(__file__).parent/'inspect_sfc_http_form.py').read_text(encoding='utf8').split('with opener.open')[0])
import urllib.parse,time
url='https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504'
started=time.monotonic()
print('Testing empty form POST on isolated session',flush=True)
try:
 with opener.open(urllib.request.Request(url,data=b''),timeout=30) as response:
  data=response.read();print('Response',response.status,len(data),round(time.monotonic()-started,1),flush=True)
except Exception as error:print(type(error).__name__,str(error),flush=True)
