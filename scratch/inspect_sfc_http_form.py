"""Inspect form structure without printing session cookies or viewstate."""
from pathlib import Path
import subprocess,json,sys,urllib.request,http.cookiejar
from html.parser import HTMLParser
sys.stdout.reconfigure(encoding='utf-8')
root=Path(__file__).resolve().parents[1]
result=subprocess.run(['node',str(root/'sfc_session_cookies.js')],capture_output=True,text=True,timeout=10,check=True)
cookies=json.loads(result.stdout)
jar=http.cookiejar.CookieJar()
for c in cookies:
 jar.set_cookie(http.cookiejar.Cookie(0,c['name'],c['value'],None,False,'sfc.chingluh.com',True,False,'/',True,True,None,True,None,None,{}))
opener=urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
class Form(HTMLParser):
 def __init__(self):super().__init__();self.fields={};self.scripts=[];self.in_script=False
 def handle_starttag(self,tag,attrs):
  a=dict(attrs)
  if tag=='input' and a.get('name'):self.fields[a['name']]=a.get('value','')
  if tag=='script':self.in_script=True
 def handle_endtag(self,tag):
  if tag=='script':self.in_script=False
 def handle_data(self,data):
  if self.in_script and any(x in data for x in ['dtDateFrom','ckbNIKE','btnSearch','btnExcel']):self.scripts.append(data)
with opener.open('https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504',timeout=20) as response:
 body=response.read().decode('utf-8');print('Ready:',response.status,response.url,'bytes',len(body))
form=Form();form.feed(body)
for k,v in form.fields.items():
 if any(x in k for x in ['DateFrom','DateTo','ckbNIKE','cboDeptFrom','EVENT','VIEWSTATE']):
  print('FIELD',k,'length',len(v),'value',v[:200] if any(x in k for x in ['DateFrom','DateTo','ckbNIKE','cboDeptFrom']) else '[state]')
for script in form.scripts:
 print('SCRIPT',script[:2100])
print('New session established:',any(c.name=='ASP.NET_SessionId' for c in jar))
