from pathlib import Path
source=(Path(__file__).parent/'test_http_export_correct.py').read_text(encoding='utf8')
needle="with opener.open(url,timeout=20) as response:body=response.read().decode('utf8')"
insert="""qms='https://sfc.chingluh.com/vhmes/QA/QMSMain.aspx'
with opener.open(qms,timeout=20) as response:page=response.read().decode('utf8')
f=CleanForm();f.feed(page);f.fields['__EVENTTARGET']='wucHeaderMainTop$TopPanel$menuRightTop';f.fields['__EVENTARGUMENT']='CLICK:0i4'
with opener.open(urllib.request.Request(qms,data=urllib.parse.urlencode(f.fields).encode()),timeout=20) as response:response.read()
print('Factory',next(c.value for c in jar if c.name=='MyFactoryName'),flush=True)
"""
source=source.replace(needle,insert+needle).replace('VH NIKE','JV NIKE').replace('HFPA_VH_NIKE','HFPA_JV_NIKE')
exec(compile(source,'scratch/test_http_jv_export.py','exec'))
