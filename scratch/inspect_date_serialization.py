from pathlib import Path
exec((Path(__file__).parent/'inspect_date_raw.py').read_text(encoding='utf8').split('for url in urls:')[0])
for url in urls:
 with opener.open('https://sfc.chingluh.com'+url.replace('&amp;','&'),timeout=15) as response:script=response.read().decode('utf8')
 for term in ['GetRaw','SetRaw','SerializeDate','Serialize','GetTimeZone','GetUTC','GetInvariantDateString']:
  matches=list(re.finditer(re.escape(term),script))
  if matches:
   print('TERM',term)
   for m in matches[:6]:print(script[max(0,m.start()-100):m.start()+400].replace('\r','').replace('\n',' ')[:500])
 print('script length',len(script),'ASPxClientDateEdit at',script.find('ASPxClientDateEdit'))
