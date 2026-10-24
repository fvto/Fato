from pathlib import Path
exec((Path(__file__).parent/'inspect_sfc_http_form.py').read_text(encoding='utf8').split('with opener.open')[0])
import time,urllib.parse
class CleanForm(Form):
 def handle_starttag(self,tag,attrs):
  a=dict(attrs)
  if tag=='input' and a.get('type','').lower() in ('submit','button','image'):return
  super().handle_starttag(tag,attrs)
url='https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504'
with opener.open(url,timeout=10) as response:body=response.read().decode('utf8')
form=CleanForm();form.feed(body)
form.fields['ctl00$rpnlMain$ContentPlaceHolder1$ckbNIKE']='C'
form.fields['__EVENTTARGET']='ctl00$rpnlMain$ContentPlaceHolder1$btnSearch';form.fields['__EVENTARGUMENT']=''
started=time.monotonic();print('Search today NIKE',flush=True)
try:
 with opener.open(urllib.request.Request(url,data=urllib.parse.urlencode(form.fields).encode()),timeout=90) as response:
  data=response.read();print('Response',response.status,len(data),round(time.monotonic()-started,1),flush=True)
except Exception as error:print(type(error).__name__,str(error),flush=True)

