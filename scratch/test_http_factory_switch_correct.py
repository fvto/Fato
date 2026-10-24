from pathlib import Path
exec((Path(__file__).parent/'inspect_sfc_factory.py').read_text(encoding='utf8').split('import re')[0])
import re
with opener.open('https://sfc.chingluh.com/vhmes/Scripts/Utilities.js',timeout=10) as response:script=response.read().decode('utf8')
for term in ['Factory','Cookie','Menu']:
 print('TERM',term)
 for m in list(re.finditer(term,script))[:6]:print(script[max(0,m.start()-100):m.start()+300])
class CleanForm(Form):
 def handle_starttag(self,tag,attrs):
  a=dict(attrs)
  if tag=='input' and a.get('type','').lower() in ('submit','button','image'):return
  super().handle_starttag(tag,attrs)
form=CleanForm();form.feed(body)
form.fields['__EVENTTARGET']='wucHeaderMainTop$TopPanel$menuRightTop';form.fields['__EVENTARGUMENT']='CLICK:0i4'
import urllib.parse,time
started=time.monotonic()
with opener.open(urllib.request.Request('https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx',data=urllib.parse.urlencode(form.fields).encode()),timeout=30) as response:
 body=response.read().decode('utf8');print('Switch response',response.status,round(time.monotonic()-started,1))
 print('Factory confirmed by server cookie:',[c.value for c in jar if c.name=='MyFactoryName'])
 print('Menu root:',re.search(r'dx-vam">(VH\d?|JV\d?)</span>',body).group(1))

