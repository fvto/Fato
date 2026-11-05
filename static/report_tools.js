'use strict';
let tools=[], chosen=null, runItems=[], timer=null, submitting=false, validated=false, version=0;
const el=id=>document.getElementById(id);
const escapeHtml=value=>String(value??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
const size=value=>value>=1048576?(value/1048576).toFixed(1)+' MB':Math.max(1,Math.round(value/1024))+' KB';
function briefIssueHTML(issue){
 const counts=items=>Object.entries(items||{}).map(([label,count])=>`${escapeHtml(label)}: ${escapeHtml(count)}`).join(' · ');
 const labels={'Ttl cost ($)':'Chi phí (USD)','Working hours':'Giờ làm','Defective Qty(Pair)':'Số đôi','QA_IssueQty_Sum':'Số lỗi QAStation','HFPA_Defect_Qty':'Số lỗi MES410'};
 const number=value=>Number.isFinite(Number(value))?Number(value).toLocaleString('vi-VN',{maximumFractionDigits:3}):String(value);
 const samples=(issue.examples||[]).map(sample=>`Dòng ${escapeHtml(sample.row)}: ${Object.entries(sample.values||{}).map(([key,value])=>`${escapeHtml(labels[key]||key)}: ${escapeHtml(number(value))}`).join('; ')}`).join(' · ');
 const locations=issue.file?`${escapeHtml(issue.file)} · sheet ${escapeHtml(issue.sheet)}${issue.rows?.length?' · dòng '+issue.rows.map(escapeHtml).join(', '):''}${issue.rows_truncated?' (các vị trí đầu tiên)':''}`:'';
 return `<li><strong>${issue.count!=null?escapeHtml(issue.count)+' '+escapeHtml(issue.unit||'dòng')+' · ':''}${escapeHtml(issue.title)}</strong><p>${escapeHtml(issue.reason)}</p>${issue.sites?`<p class="note">Nhà máy: ${counts(issue.sites)}${Object.keys(issue.types||{}).length?' · Loại: '+counts(issue.types):''}</p>`:''}${samples?`<p class="brief-metrics">${samples}</p>`:''}<p class="brief-action">${issue.severity==='incomplete'?'Cần bổ sung':'Cần xác minh'}: ${escapeHtml(issue.action)}</p>${locations?`<p class="brief-location">Vị trí: ${locations}</p>`:''}</li>`;
}
let activeBriefId=null;
function reportBriefHTML(run,check,index){
 if(check.status!=='warning')return '';
 const brief=run.export_review?.brief;
 const reason=check.expected!==undefined?`Trước xử lý: ${check.expected}; sau xử lý: ${check.actual}.`:'Xem chi tiết đối chiếu và dữ liệu nguồn trước khi sử dụng báo cáo.';
 const issues=check.issues?.length?check.issues:[{title:check.message,reason,action:'Đối chiếu kết quả của bước này. Lượt cũ có thể chưa lưu phân nhóm nguyên nhân.',severity:'reconcile',count:check.mismatches,unit:'bản ghi'}];
 const headline=check.mismatches!==undefined?`${check.mismatches} bản ghi cần xem lại · ${issues.length} nhóm vấn đề.`:check.message;
 const shown=issues.slice(0,5),more=issues.slice(5),id=`brief-${run.id}-${index}`;
 return `<div class="brief-control" data-brief-control><button type="button" class="brief-trigger" data-brief="${escapeHtml(id)}" aria-expanded="false" aria-controls="${escapeHtml(id)}" aria-label="Giải thích cảnh báo: ${escapeHtml(check.message)}"><svg viewBox="0 0 20 20" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><path d="M10 3 18 17H2Z"/><path d="M10 7v5m0 2v1"/></svg><span>Chi tiết</span></button><section id="${escapeHtml(id)}" class="report-brief brief-popover has-issues" data-brief-panel hidden aria-label="Tóm tắt vấn đề"><button type="button" class="brief-close" data-close-brief aria-label="Đóng cảnh báo">×</button><h4>Tóm tắt vấn đề</h4><p class="brief-headline">${escapeHtml(headline)}</p>${(brief?.verified||[]).map(message=>`<p class="brief-verified">${escapeHtml(message)}</p>`).join('')}<p class="note">Cảnh báo vượt ngưỡng cần xác minh, chưa đủ để kết luận nhập sai.</p><ul class="brief-list">${shown.map(briefIssueHTML).join('')}</ul>${more.length?`<details><summary>Xem thêm ${more.length} nhóm vấn đề</summary><ul class="brief-list">${more.map(briefIssueHTML).join('')}</ul></details>`:''}</section></div>`;
}
function closeBrief(focus=false){
 const trigger=Array.from(document.querySelectorAll('[data-brief]')).find(button=>button.dataset.brief===activeBriefId);
 document.querySelectorAll('[data-brief-panel]').forEach(panel=>panel.hidden=true);
 document.querySelectorAll('[data-brief]').forEach(button=>button.setAttribute('aria-expanded','false'));
 activeBriefId=null;
 if(focus)trigger?.focus();
}
function openBrief(button){
 closeBrief();
 const panel=document.getElementById(button.dataset.brief);
 if(!panel)return;
 panel.hidden=false;
 button.setAttribute('aria-expanded','true');
 activeBriefId=panel.id;
 const rect=button.getBoundingClientRect(),width=Math.min(620,window.innerWidth-32);
 panel.style.width=width+'px';
 panel.style.left=Math.max(16,Math.min(rect.right-width,window.innerWidth-width-16))+'px';
 const below=window.innerHeight-rect.bottom-16,above=rect.top-16,down=below>=Math.min(350,above);
 panel.style.maxHeight=Math.max(120,Math.min(600,down?below:above))+'px';
 panel.style.top=(down?rect.bottom+8:Math.max(8,rect.top-panel.getBoundingClientRect().height-8))+'px';
}
function bindBriefControls(section){
 section.querySelectorAll('[data-brief]').forEach(button=>{
  button.onclick=()=>activeBriefId===button.dataset.brief?closeBrief():openBrief(button);
  if(button.dataset.brief===activeBriefId)openBrief(button);
 });
 section.querySelectorAll('[data-close-brief]').forEach(button=>button.onclick=()=>closeBrief(true));
}
document.addEventListener('click',event=>{if(!event.target.closest('[data-brief-control]'))closeBrief();});
document.addEventListener('keydown',event=>{if(event.key==='Escape'&&activeBriefId){event.preventDefault();closeBrief(true);}});
window.addEventListener('resize',()=>closeBrief());
window.addEventListener('scroll',()=>closeBrief());
function saved(key,fallback){try{return JSON.parse(localStorage.getItem(key))??fallback;}catch{return fallback;}}
function store(key,value){try{localStorage.setItem(key,JSON.stringify(value));}catch{el('notice').textContent='Trình duyệt không lưu được tùy chọn; bạn vẫn có thể chạy tool.';}}
async function jsonFetch(url,options){const response=await fetch(url,options);let data;try{data=await response.json();}catch{throw new Error('Máy chủ chưa trả về kết quả. Kiểm tra kích thước file hoặc làm mới trạng thái.');}if(!response.ok){const error=new Error(data.error||'Không thực hiện được yêu cầu.');error.data=data;throw error;}return data;}
// Older Dashboard servers expose the same acquisition profiles without the new metadata fields.
const supportsAutomatic=tool=>tool?.automatic??['copq','hfpa','reinspection'].includes(tool?.id);
const legacyBackend=()=>Boolean(chosen&&!chosen.contract_version);
function invalidate(){version++;validated=false;el('validation').replaceChildren();renderActions();}
const periodField=()=>el(chosen?.period_options?.length?'quarter-period':'period');
const originalSourceNote=el('source-note').textContent;
const quarterlySfc=()=>chosen?.id==='quarterly_quality'&&el('source').value==='auto';
function updateCutoff(){
 const field=el('cutoff');
 el('cutoff-field').hidden=!quarterlySfc();
 if(!quarterlySfc())return;
 const period=periodField().value;
 const fy=1999+Number(period.slice(2,4)), month=5+3*(Number(period.at(-1))-1);
 const start=new Date(Date.UTC(fy,month,1)).toISOString().slice(0,10);
 const end=new Date(Date.UTC(fy,month+3,0)).toISOString().slice(0,10);
 field.min=start;
 field.max=end<chosen.today?end:chosen.today;
 if(!field.value||field.value<field.min||field.value>field.max){
  const yesterday=new Date(Date.parse(chosen.today)-86400000).toISOString().slice(0,10);
  field.value=end<yesterday?end:yesterday>=start?yesterday:start;
 }
}
function selectTool(id){
 chosen=tools.find(tool=>tool.id===id);
 if(!chosen)return;
 store('sfc-report-tool',id);
 const recent=saved('sfc-tool-recent',[]);
 store('sfc-tool-recent',[id,...recent.filter(x=>x!==id)].slice(0,5));
 el('workspace').style.display='block';
 el('tool-title').textContent=chosen.title;
 el('tool-description').textContent=chosen.description;
 const quarters=chosen.period_options||[];
 el('quarter-field').hidden=!quarters.length;
 el('month-field').hidden=Boolean(quarters.length);
 el('quarter-period').replaceChildren(...quarters.map(quarter=>new Option(quarter,quarter)));
 periodField().value=chosen.default_period;
 const source=el('source');
 source.querySelector('[value="auto"]').textContent=id==='quarterly_quality'
  ?'Tự động từ SFC'
  :'Tự động: dùng dữ liệu có sẵn, tải thiếu từ SFC';
 for(const option of source.options){
  option.hidden=option.value==='auto'?!supportsAutomatic(chosen):chosen.upload_only&&option.value!=='upload';
 }
 source.value=supportsAutomatic(chosen)?'auto':chosen.upload_only?'upload':'local';
 el('cutoff').value='';
 el('mode').innerHTML=chosen.modes.map(mode=>`<option value="${escapeHtml(mode.id)}">${escapeHtml(mode.title)}</option>`).join('');
 el('mode-field').hidden=chosen.modes.length===1;
 el('notice').textContent=chosen.missing_dependencies.length?'Thiếu thư viện: '+chosen.missing_dependencies.join(', '):'';
 invalidate();
 renderTools();
 renderInputs();
 window.ChingluhUI?.transition(el('workspace'));
}
function toolVisual(id){
 const shapes={
  reinspection:'<rect x="33" y="23" width="72" height="88" rx="5" transform="rotate(-14 69 67)" fill="#d1d1c9"/><rect x="42" y="16" width="72" height="88" rx="5" transform="rotate(8 78 60)" fill="#e5e5de"/><rect x="38" y="13" width="72" height="88" rx="5" fill="#eeeee7"/><path d="M53 62h41M53 73h41M53 84h27"/><circle cx="73" cy="37" r="12"/><path d="m67 37 4 4 8-9"/>',
  copq:'<path d="M23 45 75 16l52 29-52 29z" fill="#eeeee7"/><path d="m23 59 52 29 52-29M23 73l52 29 52-29"/><path d="m23 45 0 28M127 45v28M75 74v28"/><path d="m50 31 52 29M38 38l52 29M62 24l52 29" stroke-opacity=".35"/><circle cx="75" cy="45" r="9" fill="#c5c5bd"/>',
  hfpa:'<circle cx="75" cy="61" r="44" fill="#dfdfd7"/><circle cx="75" cy="61" r="32" stroke-dasharray="1 5"/><rect x="52" y="37" width="46" height="48" rx="3" fill="#f3f3ec"/><path d="m64 59 8 8 15-17M64 75h23M75 9v11M75 102v11M23 61h11M116 61h11"/>'
 };
 return '<svg viewBox="0 0 150 126" fill="none" stroke="currentColor" stroke-width="1" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'+(shapes[id]||shapes.reinspection)+'</svg>';
}
function renderTools(){
 const favorites=saved('sfc-tool-favorites',[]),recent=saved('sfc-tool-recent',[]),query=el('search').value.toLocaleLowerCase();
 const visible=tools.filter(t=>(t.title+' '+t.description).toLocaleLowerCase().includes(query)&&(!el('favorites-only').checked||favorites.includes(t.id)));
 el('tools').innerHTML=visible.map(tool=>`<article data-tool-card="${escapeHtml(tool.id)}" tabindex="0" aria-label="Chọn tool ${escapeHtml(tool.title)}" class="tool ${chosen?.id===tool.id?'active':''}"><div class="tool-topline"><span>(${String(tools.indexOf(tool)+1).padStart(2,'0')}) / REPORT TOOL</span><button type="button" data-favorite="${escapeHtml(tool.id)}" aria-label="Yêu thích ${escapeHtml(tool.title)}" aria-pressed="${favorites.includes(tool.id)}">${favorites.includes(tool.id)?'★':'☆'}</button></div><div class="tool-object">${toolVisual(tool.id)}</div><div class="run-head"><button type="button" class="tool-select" data-tool="${escapeHtml(tool.id)}">${escapeHtml(tool.title)}</button><span aria-hidden="true">↗</span></div><p>${escapeHtml(tool.description)}</p><span class="tag">${tool.ready?'Sẵn sàng':'Cần cấu hình'}${recent.includes(tool.id)?' · Gần đây':''}</span></article>`).join('')||'<p class="tools-empty">Không tìm thấy tool phù hợp.</p>';
 document.querySelectorAll('[data-tool-card]').forEach(card=>{card.onclick=event=>{if(!event.target.closest('[data-favorite]'))selectTool(card.dataset.toolCard);};card.onkeydown=event=>{if(event.target===card&&(event.key==='Enter'||event.key===' ')){event.preventDefault();selectTool(card.dataset.toolCard);}};});
 document.querySelectorAll('[data-favorite]').forEach(node=>node.onclick=event=>{event.stopPropagation();const id=node.dataset.favorite;store('sfc-tool-favorites',favorites.includes(id)?favorites.filter(x=>x!==id):[...favorites,id]);renderTools();});
}
function renderAutomaticInputs(period){
 if(quarterlySfc()){
  const cutoff=escapeHtml(el('cutoff').value);
  el('groups').innerHTML=`<div class="group"><h3>HFPA / QAStation và MES410</h3><p>Kiểm tra dữ liệu đã lưu trong sheet VH FTT. Phần còn thiếu lấy QAStation và MES410 cùng kỳ, đối chiếu theo SOP HFPA.</p><p class="note">VH · VH2 · JV · JV2</p></div>
  <div class="group"><h3>FTT Internal</h3><p>Dùng dữ liệu Assembly đã lưu; lấy phần còn thiếu từ FTT Internal và gộp số đôi lặp trên các dòng lỗi.</p><p class="note">VH · VH2 · JV · JV2</p></div>
  <p class="note auto-note">${escapeHtml(period)} · Từ ${escapeHtml(el('cutoff').min)} đến ${cutoff}. Bổ sung phần thiếu vào sheet VH FTT theo ngày chốt. Các dòng B/, C/ phục vụ biểu đồ được đặt cuối sheet. Số quý trước dùng dữ liệu đã lưu trong mẫu báo cáo. Các file SFC đúng khoảng ngày được dùng lại khi có sẵn.</p>`;
  return;
 }
 const cards=chosen.id==='reinspection'?[
  ['Re-Inspection · 4 nhà máy','Báo cáo kiểm tra lại, đủ kỳ tháng; không giới hạn chuyền và loại.','VH · VH2 · JV · JV2'],
  ['FTT Internal · 4 nhà máy','FTT nội bộ, đủ kỳ tháng; không chọn trạm.','VH · VH2 · JV · JV2'],
  ['HFPA / MES410 · 4 nhà máy','MES410 nội bộ; tắt NIKE và không giới hạn chuyền.','VH · VH2 · JV · JV2']
 ]:chosen.id==='hfpa'?[
  ['QAStation / FTT · 4 nhà máy','QAStation, chỉ chọn trạm HFPA và xác nhận OK.','VH · VH2 · JV · JV2'],
  ['HFPA / MES410 · 4 nhà máy','MES410 nội bộ; tắt NIKE và không giới hạn chuyền.','VH · VH2 · JV · JV2']
 ]:[
  ...(el('mode').value!=='ftt'?[['COPQ Data · 7 nhà máy','COPQ chi tiết, đủ kỳ tháng; không giới hạn Category và Type.','VH · VH2 · VH3 · VH4 · JV · JV2 · JVB/JV3']]:[]),
  ...(el('mode').value!=='combine'?[['FTT Internal · 4 nhà máy','FTT nội bộ, đủ kỳ tháng; không chọn trạm.','VH · VH2 · JV · JV2']]:[])
 ];
 el('groups').innerHTML=cards.map(([title,description,plants])=>`<div class="group"><h3>${escapeHtml(title)}</h3><p>${escapeHtml(description)}</p><p class="note">${escapeHtml(plants)} · Cùng kỳ tháng ${escapeHtml(period)}</p></div>`).join('')+'<p class="note auto-note">Tự kiểm tra file đúng mẫu và đủ kỳ tháng. Chỉ tải phần còn thiếu; không cần chạy BAT hoặc chép file vào thư mục Input. Đăng nhập SFC trên Chrome kết nối Dashboard nếu được yêu cầu.</p>';
}
function renderInputs(){if(!chosen)return;updateCutoff();el('source-note').textContent=quarterlySfc()
 ?'Chọn ngày chốt để xuất Excel và PowerPoint theo hai mẫu gốc. Khi cần tải, dùng Chrome đã kết nối SFC.'
 :chosen.upload_only?'Chọn workbook đã lưu để cập nhật và xuất cả Excel lẫn PowerPoint theo hai mẫu gốc.':originalSourceNote;
 const source=el('source').value,period=periodField().value;if(source==='auto'){renderAutomaticInputs(period);return;}const groups=chosen.groups.filter(group=>(chosen.required_groups?.[el('mode').value]||chosen.groups.map(group=>group.id)).includes(group.id));el('groups').innerHTML=groups.map(group=>{const files=chosen.inputs.filter(item=>item.source===source&&item.group===group.id&&(!item.period||item.period===period));return `<div class="group"><h3>${escapeHtml(group.title)}</h3>${source==='upload'?`<label class="dropzone" data-drop="${escapeHtml(group.id)}">Thả file vào đây hoặc chọn file<input type="file" ${chosen.upload_only?'':'multiple'} accept="${(group.formats||['.xls','.xlsx']).join(',')}" data-upload="${escapeHtml(group.id)}"></label><p class="note" data-upload-list="${escapeHtml(group.id)}">Tối đa ${size(chosen.max_file_bytes)} mỗi file. ${chosen.upload_only?'Chọn một workbook tổng hợp.':'Tên file có mã nhà máy, ví dụ -VH.xlsx.'}</p>`:`<label><input type="checkbox" ${source==='local'?'checked':''} data-all="${escapeHtml(group.id)}"> Chọn tất cả (${files.length})</label><div class="files">${files.map(item=>`<label class="file"><input type="checkbox" ${source==='local'?'checked':''} data-input="${escapeHtml(item.id)}" data-group="${escapeHtml(group.id)}"><span>${escapeHtml(item.name)}<br><span class="note">${size(item.size)} · ${escapeHtml(item.plant||'Chưa rõ nhà máy')}</span></span></label>`).join('')||'<p class="empty">Chưa có file đúng tháng trong nhóm này.</p>'}</div>`}</div>`;}).join('');document.querySelectorAll('[data-all]').forEach(box=>box.onchange=()=>{document.querySelectorAll('[data-input]').forEach(item=>{if(item.dataset.group===box.dataset.all)item.checked=box.checked;});invalidate();});document.querySelectorAll('[data-input]').forEach(box=>box.onchange=invalidate);document.querySelectorAll('[data-upload]').forEach(input=>{input.onchange=()=>{invalidate();const label=document.querySelector(`[data-upload-list="${input.dataset.upload}"]`);label.textContent=Array.from(input.files).map(f=>`${f.name} · ${size(f.size)}`).join(' | ');};const zone=input.closest('[data-drop]');zone.ondragover=e=>{e.preventDefault();zone.classList.add('dragging');};zone.ondragleave=()=>zone.classList.remove('dragging');zone.ondrop=e=>{e.preventDefault();zone.classList.remove('dragging');input.files=e.dataTransfer.files;input.onchange();};});}
function formData(){const data=new FormData();data.append('period',periodField().value);data.append('mode',el('mode').value);if(quarterlySfc())data.append('cutoff',el('cutoff').value);data.append('selections',JSON.stringify(Array.from(document.querySelectorAll('[data-input]:checked')).map(box=>({id:box.dataset.input,group:box.dataset.group}))));data.append('acquire_missing',el('source').value==='auto'?'true':'false');document.querySelectorAll('[data-upload]').forEach(input=>{for(const file of input.files)data.append(input.dataset.upload,file);});return data;}
function rowDeleteHTML(report,issue,index){
 if(report.issues.slice(0,index).some(previous=>previous.edit_token&&previous.sheet===issue.sheet&&previous.row===issue.row))return '';
 return ` <input type="checkbox" data-pick-row data-sheet="${escapeHtml(issue.sheet)}" aria-label="Chọn dòng ${issue.row} để xóa"> <button type="button" data-delete-row="${escapeHtml(issue.edit_token)}" data-file="${escapeHtml(report.name)}" data-sheet="${escapeHtml(issue.sheet)}" data-row="${issue.row}" aria-label="Xóa dòng ${issue.row}">Xóa dòng</button>`;
}
function batchDeleteHTML(report){
 const sheets=new Set(report.issues.filter(issue=>issue.edit_token).map(issue=>issue.sheet));
 return [...sheets].map(sheet=>`<button type="button" data-delete-selected data-sheet="${escapeHtml(sheet)}">Xóa các dòng đã chọn · ${escapeHtml(sheet)}</button>`).join(' ');
}
function validationHTML(reports, editable=false){return reports.map(report=>`<details ${report.valid?'':'open'}><summary>${report.valid?'✓':'!'} ${escapeHtml(report.name)} · ${report.rows} dòng · ${report.error_count} lỗi</summary>${report.sheets.map(s=>`<p class="note">Sheet ${escapeHtml(s.name)} · tiêu đề dòng ${s.header_row} · ${s.rows} bản ghi</p>`).join('')}${(report.excluded_summary_rows||[]).map(s=>`<p class="note">Sheet ${escapeHtml(s.sheet)}, dòng ${s.row}: ${escapeHtml(s.reason)}</p>`).join('')}${report.business_exclusions?.rework_rows?`<p class="note">COPQ_Clean loại ${report.business_exclusions.rework_rows} dòng Rework theo quy tắc cũ; bản sao đầu vào vẫn giữ nguyên.</p>`:''}${report.blank_rows?`<p class="note">${report.blank_rows} dòng trống không chứa bản ghi.</p>`:''}<ul>${report.issues.map((issue,index)=>`<li>${escapeHtml(issue.sheet||'File')}${issue.row?' · dòng '+issue.row:''}${issue.column?' · cột '+escapeHtml(issue.column):''}: ${escapeHtml(issue.message)}${editable&&issue.edit_token?` <button type="button" data-edit-cell="${escapeHtml(issue.edit_token)}" aria-label="Sửa ô ${escapeHtml(issue.column)}${issue.row}">Sửa</button>${rowDeleteHTML(report,issue,index)}`:''}</li>`).join('')}</ul>${editable?batchDeleteHTML(report):''}${report.error_count>report.issues.length?'<p>Hiển thị 100 lỗi đầu tiên; toàn bộ dòng đã được kiểm tra.</p>':''}</details>`).join('');}
function renderActions(){const active=runItems.some(run=>['queued','running'].includes(run.status));el('start').disabled=submitting||active||!chosen?.ready||(!validated&&!legacyBackend());el('validate').hidden=legacyBackend();el('validate').disabled=submitting||active||!chosen?.ready;el('start').textContent=submitting?'Đang kiểm tra…':active?'Đang xử lý…':legacyBackend()?'Chạy báo cáo':'3. Chạy tool';}
function renderRuns(){
 const diagnosticsOpen=new Set(Array.from(el('runs').querySelectorAll('[data-diagnostics][open]'),
  node=>node.closest('.run').dataset.run));
 const opened=new Set(Array.from(document.querySelectorAll('.run')).filter(node=>node.querySelector(':scope > details')?.open).map(node=>node.dataset.run));const acknowledged=new Set(Array.from(document.querySelectorAll('[data-ack]:checked')).map(node=>node.dataset.ack));renderActions();el('runs').replaceChildren();if(!runItems.length){el('runs').textContent='Chưa có lượt chạy. Chọn tool và dữ liệu để bắt đầu.';return;}for(const run of runItems){const section=document.createElement('article');section.className='run';section.dataset.run=run.id;const labels={queued:'Đang khởi động',running:'Đang xử lý',done:'Hoàn tất',error:'Chưa hoàn tất',cancelled:'Đã hủy',pending:'Chờ xử lý'};const review=run.export_review;const needsAck=run.requires_ack;section.innerHTML=`<div class="run-head"><strong>${escapeHtml(run.title)} · ${escapeHtml(run.period)}</strong><span class="status ${escapeHtml(run.status)}">${escapeHtml(labels[run.status]||run.status)}</span></div><p>${escapeHtml(run.error||run.stage)}</p>${run.progress?`<p class="note">${run.progress.completed}/${run.progress.total} ${escapeHtml(run.progress.unit)}</p>`:''}<ol class="steps">${(run.steps||[]).map(step=>`<li class="${escapeHtml(step.status)}">${escapeHtml(step.title)} — ${escapeHtml(labels[step.status]||step.status)}${step.finished_at?' · đã kết thúc':''}</li>`).join('')}</ol>${run.can_cancel?'<button data-cancel>Hủy tại điểm an toàn</button>':''}${run.cancel_requested&&['queued','running'].includes(run.status)?'<p>Đã yêu cầu hủy. Đang chờ bước hiện tại kết thúc để giữ file và trả trạng thái SFC.</p>':''}<details ${opened.has(run.id)?'open':''}><summary>4. Xem kết quả / Export center · ${escapeHtml(run.started_at?.slice(0,19).replace('T',' '))}</summary>${review?`<ul>${review.checks.map((c,index)=>`<li class="${c.status==='warning'?'error':''}">${c.status==='passed'?'✓':'!'} ${escapeHtml(c.message)}${c.expected!==undefined?` · trước: ${escapeHtml(c.expected)} → sau: ${escapeHtml(c.actual)}`:''}${c.mismatches!==undefined?' · '+c.mismatches+' bản ghi cần xem lại':''}${reportBriefHTML(run,c,index)}</li>`).join('')}</ul><p class="note">${escapeHtml(review.scope)}</p>`:''}${needsAck?`<label class="warning"><input type="checkbox" data-ack="${run.id}" ${acknowledged.has(run.id)?'checked':''}> Tôi đã xem cảnh báo; chấp nhận tải kết quả cần kiểm tra.</label>`:''}<div class="links">${run.bundle?`<a data-download href="${escapeHtml(run.bundle.url)}" download>↓ Tải gói ZIP giữ liên kết</a>`:''}</div><div class="export-table"><table><thead><tr><th>File đầu ra</th><th>Kích thước / Nội dung</th><th>SHA-256</th></tr></thead><tbody>${(run.files||[]).map(file=>`<tr><td><a data-download href="${escapeHtml(file.url)}" download>${escapeHtml(file.name)}</a></td><td>${size(file.size)}${(file.sheets||[]).map(s=>`<br>${escapeHtml(s.name)}: ${s.rows} dòng`).join('')}</td><td class="hash">${escapeHtml(file.sha256||'Lượt cũ chưa có checksum')}</td></tr>`).join('')}</tbody></table></div><details><summary>Đầu vào đã kiểm tra</summary>${validationHTML(run.validation||[])}</details><details data-diagnostics ${diagnosticsOpen.has(run.id)?'open':''}><summary>Chẩn đoán kỹ thuật (không chứa dữ liệu ô)</summary><pre>${escapeHtml(JSON.stringify(run.diagnostics||{},null,2))}\n${escapeHtml(run.log||'')}</pre></details></details>`;const cancel=section.querySelector('[data-cancel]');if(cancel)cancel.onclick=async()=>{cancel.disabled=true;try{await jsonFetch('/api/report-tools/runs/'+run.id+'/cancel',{method:'POST'});await poll();}catch(error){el('notice').textContent=error.message;cancel.disabled=false;}};section.querySelectorAll('[data-download]').forEach(link=>link.onclick=e=>{if(needsAck){if(!section.querySelector('[data-ack]').checked){e.preventDefault();el('notice').textContent='Xem cảnh báo và đánh dấu xác nhận trước khi tải.';}else{const url=new URL(link.href,location.href);url.searchParams.set('ack','1');link.href=url.toString();}}});el('runs').appendChild(section);bindBriefControls(section);}}
async function refresh(){try{const data=await jsonFetch('/api/report-tools?catalog=0');tools=data.tools;runItems=data.runs;const requested=new URLSearchParams(location.search).get('tool');const preferred=chosen?.id||requested||saved('sfc-report-tool',null);const current=tools.some(tool=>tool.id===preferred)?preferred:tools[0]?.id;if(!chosen)selectTool(current);else{chosen=tools.find(tool=>tool.id===current);renderTools();}renderRuns();schedule();}catch(error){el('notice').textContent=error.message;}}
function schedule(){clearTimeout(timer);if(runItems.some(run=>['queued','running'].includes(run.status)))timer=setTimeout(poll,2000);}
async function poll(){try{for(const item of runItems.filter(run=>['queued','running'].includes(run.status))){const fresh=await jsonFetch('/api/report-tools/runs/'+item.id);runItems=runItems.map(run=>run.id===fresh.id?fresh:run);}renderRuns();schedule();}catch(error){el('notice').textContent=error.message+' Tác vụ vẫn được giữ lại; bấm Làm mới để xem tiếp.';}}
function showValidation(reports){
 el('validation').innerHTML=validationHTML(reports,true);
 const coverage=reports.find(report=>report.valid&&report.coverage)?.coverage;
 if(coverage){
  const note=document.createElement('p');
  note.className='note';
  note.textContent=`Kỳ có dữ liệu: ${coverage.months.join(', ')} (${coverage.months.length}/${coverage.expected.length} tháng).`;
  el('validation').prepend(note);
 }
 if(el('source').value==='upload'&&reports.some(report=>!report.valid)){
  const note=document.createElement('p');
  note.className='note';
  note.textContent=chosen.upload_only
   ?'Sửa trong Excel rồi chọn lại file.'
   :'Với file tải lên từ máy, hãy sửa trong Excel rồi chọn lại file. Để sửa ngay tại đây, chọn file nguồn từ máy chủ Dashboard.';
  el('validation').prepend(note);
 }
}

el('validation').addEventListener('click',async event=>{
 const button=event.target.closest('[data-edit-cell],[data-delete-row],[data-delete-selected]');
 if(!button||submitting||!chosen)return;
 const editVersion=version;
 const url='/api/report-tools/'+chosen.id+'/cell';
 const batch=button.hasAttribute('data-delete-selected');
 const deletion=batch||button.hasAttribute('data-delete-row');
 const rows=batch?Array.from(button.closest('details').querySelectorAll('[data-pick-row]:checked'))
  .filter(input=>input.dataset.sheet===button.dataset.sheet).map(input=>input.nextElementSibling):[button];
 if(!rows.length){
  el('notice').textContent='Đánh dấu các dòng cần xóa trong sheet này.';
  return;
 }
 const token=rows[0].dataset.editCell||rows[0].dataset.deleteRow;
 const requestCell=data=>jsonFetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({token,...data})});
 if(deletion){
  const {file,sheet}=rows[0].dataset;
  const numbers=rows.map(item=>item.dataset.row).join(', ');
  if(!window.confirm(`Xóa ${rows.length} dòng (${numbers}) trong sheet “${sheet}” của file “${file}”? Dashboard sẽ sao lưu file và cập nhật dòng tổng nếu có.`))return;
  submitting=true;
  button.disabled=true;
  el('notice').textContent='Đang xóa dòng…';
  renderActions();
  let deleted=false;
  try{
   await requestCell({delete_row:true,row_tokens:rows.slice(1).map(item=>item.dataset.deleteRow)});
   deleted=true;
  }catch(error){
   if(version===editVersion)el('notice').textContent=error.message;
  }finally{
   submitting=false;
   button.disabled=false;
   renderActions();
  }
  if(deleted&&version===editVersion){
   invalidate();
   await submit(true,`Đã xóa ${rows.length} dòng và sao lưu file. `);
  }else if(deleted){
   el('notice').textContent='Đã xóa dòng và sao lưu file của lựa chọn trước.';
  }
  return;
 }
 button.disabled=true;
 try{
  const {current}=await requestCell({});
  if(version!==editVersion||!button.isConnected)return;
  const form=document.createElement('form');
  form.className='cell-editor';
  form.innerHTML='<p class="note" data-current></p><label>Giá trị mới <input type="text" required maxlength="32767" autocomplete="off"></label> <button type="submit">Lưu</button> <button type="button" data-cancel-edit>Hủy</button><p class="note" role="status" data-edit-status></p>';
  form.querySelector('[data-current]').textContent='Giá trị hiện tại: '+(current.trim()?current:'(trống)');
  const input=form.querySelector('input');
  const cancel=form.querySelector('[data-cancel-edit]');
  const save=form.querySelector('[type="submit"]');
  const status=form.querySelector('[data-edit-status]');
  cancel.onclick=()=>{form.remove();button.disabled=false;button.focus();};
  form.onsubmit=async e=>{
   e.preventDefault();
   if(submitting||version!==editVersion||!input.reportValidity())return;
   submitting=true;
   input.disabled=save.disabled=cancel.disabled=true;
   status.textContent='Đang lưu…';
   renderActions();
   let saved=false;
   try{
    await requestCell({value:input.value});
    saved=true;
   }catch(error){
    status.textContent=error.message;
   }finally{
    submitting=false;
    input.disabled=save.disabled=cancel.disabled=false;
    renderActions();
   }
   if(saved&&version===editVersion){
    invalidate();
    await submit(true,'Đã lưu ô và sao lưu file. ');
   }else if(saved){
    el('notice').textContent='Đã lưu ô và sao lưu file của lựa chọn trước.';
   }
  };
  button.after(form);
  input.focus();
 }catch(error){
  if(version===editVersion){el('notice').textContent=error.message;button.disabled=false;}
 }
});

async function submit(validateOnly,prefix=''){if(submitting||runItems.some(run=>['queued','running'].includes(run.status))||!chosen||(!validateOnly&&!validated&&!legacyBackend()))return;const submittedVersion=version;submitting=true;el('notice').textContent='';renderActions();try{const result=await jsonFetch('/api/report-tools/'+chosen.id+(validateOnly?'/validate':'/runs'),{method:'POST',body:formData()});if(validateOnly){if(submittedVersion===version){validated=result.valid;showValidation(result.validation);el('notice').textContent=prefix+(quarterlySfc()?`Kỳ hợp lệ. Tool sẽ lấy và kiểm tra ${result.will_download.length} file SFC, dùng lại file đúng kỳ khi có sẵn.`:result.will_download?.length?`Đầu vào hiện có hợp lệ. Sẽ tải và kiểm tra thêm ${result.will_download.length} file từ SFC sau khi bấm Chạy.`:'Đầu vào hợp lệ. Sẵn sàng chạy tool.');}}else{runItems.unshift(result);store('sfc-tool-last-successful-settings-'+chosen.id,{mode:el('mode').value,source:el('source').value,period:periodField().value,...(quarterlySfc()?{cutoff:el('cutoff').value}:{})});schedule();}}catch(error){if(submittedVersion===version){validated=false;if(error.data?.validation)showValidation(error.data.validation);el('notice').textContent=prefix+error.message;}}finally{submitting=false;renderRuns();}}
el('refresh').onclick=async()=>{invalidate();await refresh();if(['local','sfc'].includes(el('source').value)){const period=periodField().value;await el('source').onchange();periodField().value=period;}renderInputs();};el('search').oninput=renderTools;el('favorites-only').onchange=renderTools;
for(const id of ['period','mode'])el(id).onchange=()=>{invalidate();renderInputs();};
el('quarter-period').onchange=()=>{invalidate();if(quarterlySfc())renderInputs();};
el('cutoff').onchange=()=>{invalidate();if(quarterlySfc())renderAutomaticInputs(periodField().value);};
el('source').onchange=async()=>{invalidate();const tool=chosen,source=el('source').value;if(['local','sfc'].includes(source)){el('groups').textContent='Đang đọc danh sách file từ ổ đĩa…';try{const data=legacyBackend()?{inputs:tool.inputs.filter(item=>item.source===source)}:await jsonFetch('/api/report-tools/'+tool.id+'/inputs?source='+source);tool.inputs=[...tool.inputs.filter(item=>item.source!==source),...data.inputs];if(chosen!==tool||el('source').value!==source)return;const periods=data.inputs.map(item=>item.period).filter(p=>p&&p!=='mixed').sort();if(source==='local')periodField().value=tool.default_period;else if(periods.length)periodField().value=periods[periods.length-1];}catch(error){el('notice').textContent=error.message;}}renderInputs();};
el('save-preset').onclick=()=>{if(chosen){store('sfc-tool-preset-'+chosen.id,{mode:el('mode').value,source:el('source').value,period:periodField().value,...(quarterlySfc()?{cutoff:el('cutoff').value}:{})});el('notice').textContent='Đã lưu cấu hình trên trình duyệt này; không lưu file hoặc dữ liệu đầu vào.';}};
el('load-preset').onclick=async()=>{const preset=saved('sfc-tool-preset-'+chosen?.id,null);if(!preset){el('notice').textContent='Chưa có preset cho tool này.';return;}el('mode').value=chosen.modes.some(mode=>mode.id===preset.mode)?preset.mode:chosen.modes[0].id;el('source').value=preset.source;await el('source').onchange();periodField().value=preset.period;if(quarterlySfc())el('cutoff').value=preset.cutoff||'';invalidate();renderInputs();};
el('validate').onclick=()=>submit(true);el('start').onclick=()=>submit(false);refresh();
