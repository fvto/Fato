'use strict';
(() => {
    const $ = id => document.getElementById(id);
    const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    const read = (key, fallback) => {try {return JSON.parse(localStorage.getItem(key)) ?? fallback;} catch {return fallback;}};
    const save = (key, value) => {try {localStorage.setItem(key, JSON.stringify(value));} catch {notice('Bộ nhớ trình duyệt đã đầy; chưa lưu được thiết lập.', true);}};
    const operations = {average:'Trung bình',sum:'Tổng',min:'Nhỏ nhất',max:'Lớn nhất',count:'Đếm số',weighted_average:'Trung bình có trọng số',ratio_of_sums:'Tỷ lệ tổng'};
    const statuses = {match:'Khớp',rounding:'Chỉ khác do làm tròn',within_tolerance:'Trong sai số cho phép',difference:'Chênh lệch',missing_value:'Thiếu giá trị',missing_input:'Thiếu dữ liệu đầu vào',formula_error:'Lỗi Excel',formula_uncached:'Công thức chưa lưu kết quả',invalid_input:'Đầu vào không hợp lệ',invalid_value:'Giá trị không hợp lệ',ambiguous_key:'Trùng khóa',invalid_key:'Khóa không hợp lệ',unmatched:'Dòng chưa ghép',no_data:'Không có dữ liệu'};
    const issues = new Set(['difference','missing_value','missing_input','formula_error','formula_uncached','invalid_input','invalid_value','ambiguous_key','invalid_key','unmatched']);
    const reasonTranslations = new Map([
        ['Values match.', 'Các giá trị khớp nhau.'],
        ['Values differ. Tolerance 0.', 'Các giá trị chênh lệch. Sai số cho phép bằng 0.'],
        ['The total weight / denominator is zero.', 'Tổng trọng số / mẫu số bằng 0.'],
        ['No valid numeric input. Missing values are not treated as zero.', 'Không có đầu vào số hợp lệ. Giá trị thiếu không được coi là 0.'],
        ['Expected-value cell contains an Excel error.', 'Ô kết quả báo cáo chứa lỗi Excel.'],
        ['Formula has no saved result. Recalculate and save it in Excel / WPS before checking.', 'Công thức chưa có kết quả được lưu. Tính lại và lưu file trong Excel / WPS trước khi kiểm tra.'],
        ['Numeric inputs exist, but the expected-value cell is blank or N/A.', 'Có đầu vào số, nhưng ô kết quả báo cáo trống hoặc N/A.'],
        ['Expected-value cell is not a valid number in the chosen number interpretation.', 'Ô kết quả báo cáo không phải số hợp lệ theo cách đọc số đã chọn.'],
        ['A numeric result is present, but there are no usable inputs / denominator.', 'Có kết quả số nhưng không có đầu vào / mẫu số hợp lệ để kiểm chứng.'],
        ['Blank / uncached matching key; rows were not paired.', 'Khóa ghép trống hoặc chưa lưu kết quả công thức; các dòng chưa được ghép.'],
        ['Duplicate key. Add another key column; no first-match or many-to-many pairing was used.', 'Khóa bị trùng. Hãy thêm cột khóa; hệ thống chưa tự chọn dòng đầu tiên hoặc ghép nhiều dòng với nhau.'],
        ['Row exists only in reference.', 'Dòng chỉ có ở file nguồn.'],
        ['Row exists only in comparison.', 'Dòng chỉ có ở file đối chiếu.'],
        ['A formula has no saved result; recalculate and save the workbook first.', 'Công thức chưa lưu kết quả; hãy tính lại và lưu file trước.'],
        ['One of the cells contains an Excel error.', 'Một trong các ô chứa lỗi Excel.'],
        ['Both cells are blank / N/A.', 'Cả hai ô đều trống / N/A.'],
        ['Only one side has a value.', 'Chỉ một bên có giá trị.'],
        ['Text values match.', 'Các giá trị văn bản khớp nhau.'],
        ['Text values differ.', 'Các giá trị văn bản chênh lệch.'],
        ['Numeric value cannot be interpreted consistently. Check the number interpretation.', 'Không đọc được giá trị số nhất quán. Hãy kiểm tra cách đọc giá trị số.'],
    ]);
    function resultReason(value) {
        let message=String(value||''), prefix='';
        const aggregate=message.match(/^(Average|Sum|Minimum|Maximum|Numeric count|Weighted average|Ratio of sums): (\d+) numeric inputs, (\d+) skipped\. /);
        if(aggregate){
            const names=['Trung bình','Tổng','Nhỏ nhất','Lớn nhất','Đếm số','Trung bình có trọng số','Tỷ lệ tổng'];
            const methods=['Average','Sum','Minimum','Maximum','Numeric count','Weighted average','Ratio of sums'];
            prefix=`${names[methods.indexOf(aggregate[1])]}: ${aggregate[2]} đầu vào số, bỏ qua ${aggregate[3]} ô. `;
            message=message.slice(aggregate[0].length);
        }
        message=message.replace(/^Values (match|differ) after rounding to (\d+) decimals; tolerance 0\./,(_,state,digits)=>`Các giá trị ${state==='match'?'khớp nhau':'chênh lệch'} sau khi làm tròn đến ${digits} chữ số thập phân; sai số cho phép bằng 0.`);
        for(const [source,translation] of reasonTranslations){
            if(message.startsWith(source)){message=translation+message.slice(source.length);break;}
        }
        // Only translate engine-generated cell diagnostics; preserve cell contents.
        message=message.replace(/(^|; )([A-Z]+\d+): missing or invalid weight \/ denominator/g,'$1$2: thiếu hoặc không hợp lệ ở trọng số / mẫu số')
            .replace(/(^|; )([A-Z]+\d+): missing value/g,'$1$2: thiếu giá trị')
            .replace(/(^|; )([A-Z]+\d+): (invalid|error|uncached) value /g,(_,separator,address,kind)=>`${separator}${address}: ${kind==='error'?'lỗi Excel':kind==='uncached'?'công thức chưa lưu kết quả':'giá trị không hợp lệ'} `);
        return prefix+message.replace('Trung bình MQAA% theo dòng export; giữ nguyên dòng lặp.','Trung bình MQAA% theo dòng dữ liệu đã xuất; giữ nguyên dòng lặp.');
    }
    function builtInRuleName(name) {
        return String(name||'')
            .replace(/^QC nội bộ · Average từng dòng/, 'QC nội bộ · Trung bình từng dòng')
            .replace(/^QC nội bộ · Grand Total từng cột/, 'QC nội bộ · Tổng hợp từng cột')
            .replace(/^Nguồn → QC · Line × Audit form/, 'Nguồn → QC · Chuyền × Biểu mẫu kiểm toán')
            .replace(/^Check ([A-Z]+) · (average|avg|mean|total|sum|minimum|maximum|count)$/,(_,column,method)=>`Kiểm tra cột ${column} · ${{average:'trung bình',avg:'trung bình',mean:'trung bình',total:'tổng',sum:'tổng',minimum:'nhỏ nhất',maximum:'lớn nhất',count:'đếm số'}[method]}`)
            .replace(/^Footer totals · review Average \/ Sum method$/, 'Tổng cuối bảng · kiểm tra cách tính trung bình / tổng');
    }
    let books = [], rules = [], result = null, busy = false, deletingExport = false;
    let exportIds = [];
    let exportingHtml = false;
    let workflowPlan = null, workflowBusy = false;
    const linkedChat = new URLSearchParams(location.search).get('chat');
    const storedChat = read('compare-chat-id',null);
    let chat = {messages:[],jobs:{}}, chatId = /^[a-f0-9]{32}$/.test(linkedChat||'')?linkedChat:storedChat, chatReady = false, chatBusy = false, chatTimer = null;
    let chatMode=read('compare-analysis-mode','offline');
    if(!['offline','chatgpt','codex'].includes(chatMode))chatMode='offline';
    let pendingChat = read('compare-chat-pending',null);
    async function api(url, options) {
        const response = await fetch(url, options);
        const data = await response.json().catch(() => ({error:'Không đọc được phản hồi từ máy chủ.'}));
        if (!response.ok) throw new Error(data.error || 'Yêu cầu không thành công.');
        return data;
    }
    const post = data => ({method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
    function updateCatalogControls() {
        const disabled=busy||deletingExport;
        $('compare-catalog').disabled=disabled;
        $('compare-import').disabled=disabled||!$('compare-catalog').value;
        $('compare-delete-all').disabled=disabled||!exportIds.length;
    }
    async function refreshCatalog() {
        const data=await api('/api/compare/catalog');
        exportIds=data.delete_ids;
        const selected=$('compare-catalog').value;
        $('compare-catalog').innerHTML='<option value="">Chọn file đã xuất…</option>'+options(data.files.map(file=>[file.id,`${file.name} · ${Math.round(file.size/1024)} KB`]),selected);
        $('compare-catalog-count').textContent=data.total>data.files.length?`${data.total} file đã xuất · hiển thị ${data.files.length} file mới nhất. Xóa tất cả áp dụng cho ${data.total} file.`:`${data.total} file đã xuất đang lưu.`;
        updateCatalogControls();
    }
    function notice(message, error=false) {$('compare-notice').textContent=message; $('compare-notice').classList.toggle('error',error);}
    function persist() {save('compare-workbooks',books.map(book=>book.id)); save('compare-rules',rules);}
    function options(items, value) {return items.map(([key,label])=>`<option value="${esc(key)}" ${String(value)===String(key)?'selected':''}>${esc(label)}</option>`).join('');}
    function field(label, key, value, type='text', extra='') {return `<label ${extra}>${esc(label)}<input type="${type}" data-field="${esc(key)}" value="${esc(value)}" ${type==='number'?'step="any" min="0"':''}></label>`;}
    function select(label,key,items,value,extra='') {return `<label ${extra}>${esc(label)}<select data-field="${esc(key)}">${options(items,value)}</select></label>`;}
    function bookFor(spec) {return books.find(book=>book.id===spec?.workbook);}
    function sheetFor(spec) {return bookFor(spec)?.sheets.find(sheet=>sheet.name===spec?.sheet);}
    function fileFields(rule,side,label) {
        const spec=rule[side]||{};
        const sheets=(bookFor(spec)?.sheets||[]).map(sheet=>[sheet.name,`${sheet.name} · ${sheet.rows} dòng`]);
        const exists=sheets.some(([name])=>name===spec.sheet);
        return select('File '+label,side+'.workbook',books.map(book=>[book.id,book.name]),spec.workbook)+
            select('Trang tính '+label,side+'.sheet',exists?sheets:[['','Chọn trang tính…'],...sheets],exists?spec.sheet:'');
    }
    function commonFields(rule) {
        return select('Cách đọc giá trị số','number_mode',[
            ['number','Số theo giá trị đã lưu'],['percent','Phần trăm: số thập phân hoặc văn bản có %'],['percentage_points','Phần trăm: số từ 0–100 chưa định dạng']],rule.number_mode||'number')+
            select('Làm tròn trước khi so sánh','comparison_rounding_digits',[['','Không làm tròn'],['1','1 chữ số'],['2','2 chữ số'],['3','3 chữ số']],rule.comparison_rounding_digits??'')+
            '<p class="note">Sai số 0 sau bước làm tròn đã chọn.</p>';
    }
    function renderRules() {
        $('compare-rules').innerHTML=rules.map((rule,index)=>{
            let content=field('Tên phần / quy tắc','name',builtInRuleName(rule.name))+fileFields(rule,'left',rule.kind==='comparison'?'A':'nguồn');
            if(rule.kind==='calculation') {
                const proposed=sheetFor(rule.left)?.suggestions||[];
                content+=select('Phép tính','operation',Object.entries(operations),rule.operation)+
                    select('Tính theo','direction',[['rows','Từng dòng'],['columns','Từng cột']],rule.direction)+
                    field('Vùng dữ liệu đầu vào','source_range',rule.source_range)+field('Vùng kết quả báo cáo','target_range',rule.target_range)+
                    field('Vùng nhãn dòng / cột (tùy chọn)','label_range',rule.label_range||'')+
                    (['weighted_average','ratio_of_sums'].includes(rule.operation)?field(rule.operation==='weighted_average'?'Vùng trọng số (cùng kích thước với đầu vào)':'Vùng mẫu số (cùng kích thước với đầu vào)','weights_range',rule.weights_range||''):'')+
                    commonFields(rule)+select('Xử lý ô trống và N/A','missing_policy',[['skip','Bỏ qua đầu vào thiếu'],['error','Báo đầu vào thiếu']],rule.missing_policy||'skip')+
                    `<label class="check-option"><input type="checkbox" data-field="ignore_zero" ${rule.ignore_zero?'checked':''}>Loại số 0 (mặc định: vẫn tính)</label>`+
                    `<div class="full rule-suggestion">${select('Gợi ý quy tắc — kiểm tra trước khi dùng','suggestion',proposed.map((p,i)=>[i,builtInRuleName(p.name)]),Math.max(0,proposed.findIndex(p=>p.source_range===rule.source_range&&p.target_range===rule.target_range&&p.direction===rule.direction)))}<button type="button" data-use-suggestion="${index}" ${!proposed.length?'disabled':''}>Dùng gợi ý</button></div>`;
            } else {
                content+=fileFields(rule,'right','B')+field('Vùng dữ liệu A (bỏ dòng tiêu đề, ví dụ A2:D100)','left_range',rule.left_range)+
                    field('Vùng dữ liệu B (bỏ dòng tiêu đề)','right_range',rule.right_range)+
                    select('Ghép dòng theo','matching',[['keys','Cột khóa · ghép theo mã'],['position','Thứ tự dòng · cùng bố cục']],rule.matching||'keys')+
                    (rule.matching!=='position'?field('Cột khóa A (ví dụ A hoặc A,B)','left_keys',rule.left_keys||'')+field('Cột khóa B','right_keys',rule.right_keys||''):'')+
                    field('Cột giá trị A (ví dụ C,D)','left_columns',rule.left_columns||'')+field('Cột giá trị B (cùng thứ tự đối chiếu)','right_columns',rule.right_columns||'')+
                    select('Kiểu giá trị','value_type',[['auto','Số nếu có thể, còn lại là văn bản'],['numeric','Số'],['text','Văn bản']],rule.value_type||'auto')+
                    commonFields(rule)+`<label class="check-option"><input type="checkbox" data-field="ignore_case" ${rule.ignore_case!==false?'checked':''}>Không phân biệt chữ hoa, chữ thường trong khóa và văn bản</label>`;
            }
            return `<article class="check-rule" data-rule="${index}"><div class="rule-top"><strong>${String(index+1).padStart(2,'0')} / ${rule.kind==='calculation'?'KIỂM TRA PHÉP TÍNH':'ĐỐI CHIẾU DỮ LIỆU'}</strong><button type="button" data-remove-rule="${index}" aria-label="Xóa quy tắc ${index+1}">Xóa ×</button></div><div class="rule-grid">${content}</div><p class="note method-note">${rule.kind==='calculation'?'Số 0 được tính trừ khi bạn chọn loại bỏ. Chế độ phần trăm nghiêm ngặt sẽ báo số chưa định dạng lớn hơn 1 (ví dụ 30), không tự hiểu thành 30%. Trung bình có trọng số = Σ(giá trị × trọng số) / Σ(trọng số); tỷ lệ tổng = Σ(tử số) / Σ(mẫu số).':'Có thể dùng nhiều cột làm khóa ghép. Khóa trùng được báo kèm số dòng gốc để kiểm tra; hệ thống không tự chọn dòng khớp đầu tiên.'}</p></article>`;
        }).join('') || '<p class="note">Nhập file Excel, sau đó thêm quy tắc kiểm tra phép tính hoặc đối chiếu dữ liệu.</p>';
        $('run-compare').disabled=busy||!rules.length||!books.length;
        $('add-calculation').disabled=busy||!books.length;
        $('add-comparison').disabled=busy||books.length<2;
        $('offline-add-comparison').disabled=busy||books.length<2;
        updateWorkflowControls();
        $('compare-rules').querySelectorAll('[data-field]').forEach(input=>{
          const update=()=>{
            const index=Number(input.closest('[data-rule]').dataset.rule), rule=rules[index], key=input.dataset.field;
            if(key==='suggestion')return;
            const value=input.type==='checkbox'?input.checked:input.value;
            if(key.includes('.')) {
                const [side,property]=key.split('.'); rule[side]={...(rule[side]||{}),[property]:value};
                if(property==='workbook')rule[side].sheet=bookFor(rule[side])?.sheets[0]?.name||'';
                const sheet=sheetFor(rule[side]);
                if(rule.kind==='comparison'&&sheet)rule[side==='left'?'left_range':'right_range']=sheet.data_range;
            } else rule[key]=value;
            persist();
            if(['left.workbook','left.sheet','right.workbook','right.sheet','operation','matching','number_mode'].includes(key))renderRules();
            if(result)notice('Thiết lập đã thay đổi. Chạy lại đối chiếu để cập nhật kết quả.');
          };
          input.addEventListener('change',update);
          if(input.tagName==='INPUT'&&input.type!=='checkbox')input.addEventListener('input',update);
        });
        $('compare-rules').querySelectorAll('[data-remove-rule]').forEach(button=>button.onclick=()=>{rules.splice(Number(button.dataset.removeRule),1);persist();renderRules();});
        $('compare-rules').querySelectorAll('[data-use-suggestion]').forEach(button=>button.onclick=()=>{
            const index=Number(button.dataset.useSuggestion), article=button.closest('[data-rule]');
            const suggestion=sheetFor(rules[index].left)?.suggestions[Number(article.querySelector('[data-field="suggestion"]').value)];
            if(suggestion){rules[index]=structuredClone(suggestion);persist();renderRules();notice('Đã áp dụng vùng gợi ý. Kiểm tra phép tính, vùng đầu vào và vùng kết quả trước khi chạy.');}
        });
    }
    function preferredSheet(book) {return book?.sheets.find(sheet=>sheet.suggestions.length)||book?.sheets[0];}
    function addRule(kind) {
        if(kind==='comparison'&&books.length<2){notice('Hãy nhập hai file trước khi thiết lập so sánh.',true);return;}
        if(rules.length>=20){notice('Mỗi lượt hỗ trợ tối đa 20 quy tắc.',true);return;}
        const leftBook=books[0], leftSheet=preferredSheet(leftBook);
        if(!leftBook||!leftSheet)return;
        const rightBook=books[1]||leftBook, rightSheet=rightBook.sheets.find(sheet=>sheet.name===leftSheet.name)||preferredSheet(rightBook);
        const base={kind,name:kind==='calculation'?'Kiểm tra phép tính':'Đối chiếu dữ liệu',left:{workbook:leftBook.id,sheet:leftSheet.name},number_mode:'number',tolerance:'0',rounding_digits:''};
        if(kind==='calculation')rules.push(leftSheet.suggestions[0]?structuredClone(leftSheet.suggestions[0]):{...base,operation:'average',direction:'rows',source_range:'',target_range:'',label_range:'',missing_policy:'skip',ignore_zero:false});
        else rules.push({...base,name:'So sánh 2 file',right:{workbook:rightBook.id,sheet:rightSheet.name},left_range:leftSheet.data_range,right_range:rightSheet.data_range,matching:'keys',left_keys:'',right_keys:'',left_columns:'',right_columns:'',value_type:'auto',ignore_case:true});
        if(rules.length)rules[rules.length-1].comparison_rounding_digits=$('workflow-rounding').value;
        persist();renderRules();
    }
    function renderFiles() {
        $('workbook-count').textContent=`${books.length} file`;
        $('compare-files').innerHTML=books.map(book=>`<div class="file-tag"><span>${esc(book.name)} · ${book.sheets.length} trang tính</span><button type="button" data-remove-book="${book.id}" aria-label="Bỏ ${esc(book.name)}">×</button></div>`).join('');
        $('compare-files').querySelectorAll('[data-remove-book]').forEach(button=>button.onclick=()=>{
            books=books.filter(book=>book.id!==button.dataset.removeBook);
            rules=rules.filter(rule=>rule.left?.workbook!==button.dataset.removeBook&&rule.right?.workbook!==button.dataset.removeBook);
            persist();renderFiles();renderRules();
        });
        $('preview-book').innerHTML=options(books.map(book=>[book.id,book.name]),$('preview-book').value||books[0]?.id);
        previewSheets();renderRules();renderHandoffFiles();updateChatControls();renderWorkflowFiles();
    }
    function invalidateWorkflow() {
        workflowPlan=null;$('workflow-plan').hidden=true;$('workflow-ack').checked=false;updateWorkflowControls();
    }
    function renderWorkflowFiles() {
        const source=$('workflow-source').value, report=$('workflow-report').value;
        const entries=books.map(book=>[book.id,book.name]);
        $('workflow-source').innerHTML=options(entries,books.some(book=>book.id===source)?source:books[0]?.id);
        $('workflow-report').innerHTML=options(entries,books.some(book=>book.id===report)?report:(books[1]||books[0])?.id);
        if(workflowPlan&&(workflowPlan.inputs[0].id!==$('workflow-source').value||workflowPlan.inputs[1].id!==$('workflow-report').value))invalidateWorkflow();
        updateWorkflowControls();
    }
    function updateWorkflowControls() {
        const disabled=busy||workflowBusy;
        for(const id of ['workflow-source','workflow-report','workflow-profile','workflow-factory','workflow-rounding'])$(id).disabled=disabled;
        $('workflow-preview').disabled=disabled||books.length<2||$('workflow-source').value===$('workflow-report').value;
        $('workflow-run').disabled=disabled||!workflowPlan||!$('workflow-ack').checked;
    }
    function workflowScope(scope) {
        return scope.factories?`Nhà máy ${scope.factories.join(', ')} · ${scope.period.join(' đến ')} · ${scope.scoped_rows}/${scope.source_rows} dòng nguồn trong phạm vi · ${scope.matrices.join(', ')}`:`${scope.paired_tables} cặp bảng · ${scope.skipped_sheets.length} trang tính chưa ghép.`;
    }
    function renderWorkflowPlan() {
        const plan=workflowPlan;
        $('workflow-plan-content').innerHTML=`<h3>${esc(plan.profile_label)}</h3><p><strong>Nguồn:</strong> ${esc(plan.inputs[0].name)}<br><strong>Báo cáo:</strong> ${esc(plan.inputs[1].name)}</p><p><strong>Phạm vi:</strong> ${esc(workflowScope(plan.scope))}</p><p><strong>${plan.rules.length} quy tắc được lập:</strong></p><ul>${plan.descriptions.map(item=>`<li>${esc(item)}</li>`).join('')}</ul><p><strong>Phương pháp &amp; điểm cần xác nhận:</strong></p><ul>${plan.warnings.map(item=>`<li>${esc(item)}</li>`).join('')}</ul><p><strong>Chưa bao gồm:</strong></p><ul>${plan.exclusions.map(item=>`<li>${esc(item)}</li>`).join('')}</ul>`;
        $('workflow-plan').hidden=false;$('workflow-ack').checked=false;updateWorkflowControls();
    }
    function previewSheets() {
        const book=books.find(book=>book.id===$('preview-book').value);
        $('preview-sheet').innerHTML=options((book?.sheets||[]).map(sheet=>[sheet.name,sheet.name]),preferredSheet(book)?.name);
        renderPreview();
    }
    function renderPreview() {
        const book=books.find(book=>book.id===$('preview-book').value), sheet=book?.sheets.find(sheet=>sheet.name===$('preview-sheet').value);
        $('preview-size').textContent=sheet?`${sheet.rows} dòng × ${sheet.columns} cột · dòng tiêu đề dự kiến: ${sheet.header_row}`:'';
        $('compare-preview').innerHTML=sheet?`<table><tbody>${sheet.preview.map(row=>`<tr>${row.map(item=>`<td><small>${esc(item.cell)}</small>${esc(item.value??(item.formula?'[Công thức chưa lưu kết quả]':''))}</td>`).join('')}</tr>`).join('')}</tbody></table>`:'';
    }
    async function importBooks(options) {
        if(busy||deletingExport)return;
        busy=true;$('run-compare').disabled=true;updateCatalogControls();$('compare-upload').disabled=true;notice('Đang đọc cấu trúc file và giá trị đã lưu…');
        try {
            const data=await api('/api/compare/workbooks',options);
            books.push(...data.workbooks);persist();renderFiles();
            notice(chatMode==='offline'?`Đã nhập ${data.workbooks.length} file. Chọn “Thiết lập so sánh 2 file” để chọn vùng dữ liệu và cột đối chiếu.`:`Đã nhập ${data.workbooks.length} file. Nhập yêu cầu kiểm tra vào khung trò chuyện bên dưới.`);
        } catch(error){notice(error.message,true);}
        finally{busy=false;updateCatalogControls();$('compare-upload').disabled=false;renderRules();updateChatControls();}
    }
    function showValue(value, mode) {
        if(value===null||value===undefined||value==='')return '—';
        if(typeof value==='number')return mode!=='number'?Number((value*100).toFixed(3)).toString()+'%':value.toLocaleString(undefined,{maximumFractionDigits:3});
        if(typeof value==='string'&&/^[+-]?\d+(?:[.,]\d+)?\s*%$/.test(value.trim()))return Number(Number(value.trim().replace('%','').replace(',','.')).toFixed(3)).toString()+'%';
        return String(value);
    }
    function resultSheetKey(spec) {
        return JSON.stringify([spec?.workbook||'',spec?.sheet||'']);
    }
    function resultLocation(file, sheet, cell, compared=false) {
        return `<div class="result-location${compared?' compared':''}"><span class="result-file">${compared?'↔ ':''}${esc(file)}</span><code>${esc(sheet)}${cell?'!'+esc(cell):''}</code></div>`;
    }
    function renderResult() {
        if(!result)return;
        $('compare-results-panel').hidden=false;
        $('compare-html').disabled=exportingHtml||!['exact','rounded'].includes(result.comparison_policy);
        const counts=result.counts, exact=counts.match||0;
        $('compare-summary').innerHTML=[[result.total,'Tổng mục kiểm tra'],[result.issue_count,'Cần xem lại'],[exact,'Khớp hoàn toàn'],[0,'Sai số cho phép']].map(([number,label])=>`<div class="summary-number"><strong>${number.toLocaleString('vi-VN')}</strong><span>${label}</span></div>`).join('');
        $('compare-no-data').textContent=counts.no_data?`${counts.no_data} mục không có dữ liệu số. `:'';
        $('compare-csv').href=result.csv_url;$('compare-json').href=result.json_url;
        $('compare-html').dataset.url=result.html_url||`/api/compare/results/${encodeURIComponent(result.id)}/download.html`;
        $('compare-pagination').textContent=result.filtered_total?`${result.offset+1}–${Math.min(result.offset+200,result.filtered_total)} / ${result.filtered_total}`:'Không có kết quả theo bộ lọc';
        $('compare-previous').disabled=result.offset===0;$('compare-next').disabled=result.offset+200>=result.filtered_total;
        let previousSheet=null;
        const sheetKeys=[...new Set(result.rules.map(rule=>resultSheetKey(rule.right||rule.left)))];
        $('compare-result-rows').innerHTML=result.results.map((row,index)=>{
            const type=issues.has(row.status)?'issue':['match','rounding','within_tolerance'].includes(row.status)?'ok':'';
            const rule=result.rules[row.rule_index-1];
            const primary=rule.right||rule.left, sheetKey=resultSheetKey(primary);
            const groupStart=sheetKey!==previousSheet;
            previousSheet=sheetKey;
            let groupHeading='';
            if(groupStart){
                let end=index+1;
                while(end<result.results.length){
                    const nextRule=result.rules[result.results[end].rule_index-1];
                    if(resultSheetKey(nextRule.right||nextRule.left)!==sheetKey)break;
                    end++;
                }
                groupHeading=`<tr class="sheet-group"><th colspan="6" scope="rowgroup"><div class="sheet-group-heading"><span class="sheet-group-index">TRANG TÍNH ${String(sheetKeys.indexOf(sheetKey)+1).padStart(2,'0')}</span><h3>${esc(primary.sheet)}</h3><span class="sheet-group-count">${end-index} mục trên trang</span></div></th></tr>`;
            }
            const location=resultLocation(row.file,row.sheet,row.cell);
            const other=row.other_file?resultLocation(row.other_file,row.other_sheet,row.other_cell,true):'';
            const suffix=' · '+primary.sheet;
            const ruleLabel=builtInRuleName(row.rule.endsWith(suffix)?row.rule.slice(0,-suffix.length):row.rule);
            const evidence=(row.inputs?`<p class="note">${esc(operations[row.operation]||row.operation)} · ${esc(rule.source_range)} → ${esc(rule.target_range)}<br>Số hợp lệ: ${row.valid_count} · Bỏ qua: ${row.ignored_count}</p><ul class="input-evidence">${row.inputs.map(input=>`<li><code>${esc(input.cell)}</code> <span>${esc(input.value??'Trống')}${input.role?' ('+esc(input.role==='weight / denominator'?'trọng số / mẫu số':input.role)+')':''}</span></li>`).join('')}</ul>`:'')+
                (row.formula?`<p>Công thức: <code>${esc(row.formula)}</code></p>`:'')+
                (row.other_formula?`<p>Công thức đối chiếu: <code>${esc(row.other_formula)}</code></p>`:'')+
                (row.left_rows?`<p>Dòng nguồn: ${esc(row.left_rows.join(', ')||'Không có')}<br>Dòng đối chiếu: ${esc((row.right_rows||[]).join(', ')||'Không có')}</p>`:'');
            const delta=row.delta===null||row.delta===undefined?'—':showValue(row.number_mode!=='number'?row.delta*100:row.delta,'number')+(row.number_mode!=='number'?'%':'');
            return groupHeading+`<tr class="result-row"><td><span class="status-tag ${type}">${esc(statuses[row.status]||row.status)}</span><p>${esc(row.rule_index+'. '+ruleLabel)}</p></td><td><strong>${esc(row.label||'')}</strong>${location}${other}</td><td>${esc(showValue(row.expected,row.number_mode))}</td><td>${esc(showValue(row.actual,row.number_mode))}</td><td>${esc(delta)}</td><td>${esc(resultReason(row.reason))}<details><summary>Bằng chứng và quy tắc</summary>${evidence}<p class="note">Sai số cho phép: ${esc(rule.tolerance??'0')}${row.number_mode!=='number'?' điểm phần trăm':' đơn vị tuyệt đối'}. Số chữ số hiển thị: ${esc(rule.rounding_digits||'theo định dạng ô')}.</p></details></td></tr>`;
        }).join('')||'<tr><td colspan="6">Không có kết quả trong nhóm này. Chọn “Tất cả kết quả” để xem toàn bộ.</td></tr>';
    }
    async function fetchResult(offset=0) {
        if(!result)return;
        try{result=await api(`/api/compare/results/${result.id}?status=${encodeURIComponent($('compare-status').value)}&offset=${offset}`);renderResult();}catch(error){notice(error.message,true);}
    }
    function renderPresets() {
        const presets=read('compare-presets',{});
        $('compare-presets').innerHTML='<option value="">Chọn mẫu quy tắc…</option>'+options(Object.keys(presets).map(name=>[name,name]),$('compare-presets').value);
    }
    function updateChatControls() {
        const web=chatMode==='chatgpt';
        $('chat-send').disabled=(!web&&!chatReady)||chatBusy||busy||!books.length||!$('chat-request').value.trim();
        $('chat-send').textContent=web?'Sao chép yêu cầu':chatBusy?'Đang phân tích…':'Gửi yêu cầu →';
        $('chat-new').disabled=chatBusy;
        $('chat-new').hidden=web||chatMode==='offline';
        $('chat-request').disabled=chatBusy;
        $('chat-mode').disabled=chatBusy;
        $('chat-import-submit').disabled=chatBusy||busy||!books.length||!$('chat-import-text').value.trim();
    }
    function previousContext() {
        const previous=chat.messages.findLast(item=>item.role==='assistant');
        const origin=chat.messages.find(item=>item.role==='user'&&item.job_id===previous?.job_id);
        return origin?.files?.length===books.length&&books.every(book=>origin.files.includes(book.id))?previous:null;
    }
    function renderHandoffFiles() {
        $('chat-handoff-files').innerHTML=books.length?'<p class="note">Tải bản gốc nếu bạn chưa có file để đính kèm:</p>'+books.map(book=>
            `<a href="/api/compare/workbooks/${encodeURIComponent(book.id)}/download" download>${esc(book.name)} ↓</a>`).join(''):'<p class="note">Chọn file ở bước 01 để chuẩn bị yêu cầu.</p>';
        $('chat-include-context-label').hidden=!previousContext();
        if(!previousContext())$('chat-include-context').checked=false;
    }
    function clearHandoff() {
        $('chat-handoff-status').textContent='';$('chat-copy-fallback').hidden=true;$('chat-copy-text').value='';
    }
    async function ensureChatConnection() {
        try{const state=await api('/api/compare/chat/status');chatReady=state.ready;$('chat-connection').textContent=state.message;}catch(error){$('chat-connection').textContent=error.message;}
        updateChatControls();
    }
    function applyChatMode() {
        const web=chatMode==='chatgpt', offline=chatMode==='offline';
        $('chat-mode').value=chatMode;
        $('offline-guide').hidden=!offline;$('chat-online-controls').hidden=offline;
        $('advanced-checks').open=offline;
        if(offline&&chatTimer){clearTimeout(chatTimer);chatTimer=null;}
        $('chat-handoff').hidden=!web;$('chat-import-panel').hidden=!web;
        $('chat-open-web').hidden=!web;$('chat-connection').hidden=web;
        $('chat-local-history').open=!web;
        $('chat-shortcut').textContent=web?'Ctrl + Enter để sao chép yêu cầu':'Enter để xuống dòng · Ctrl + Enter để gửi';
        updateChatControls();
    }
    function evidenceLinks(job) {
        return (job?.evidence||[]).map(item=>`<div class="chat-evidence"><a href="${esc(item.url)}" download>${esc(item.label)} · CSV ↗</a>${item.json_url?`<a href="${esc(item.json_url)}" download>Yêu cầu và nguồn dữ liệu đầy đủ ↗</a>`:''}${item.result_id?`<button type="button" data-chat-result="${esc(item.result_id)}">Xem các ô →</button>`:''}</div>`).join('');
    }
    function renderChat() {
        $('chat-local-history').hidden=!chat.messages.length;
        $('chat-messages').innerHTML=chat.messages.map(item=>{
            const job=chat.jobs[item.job_id];
            const fileNames=item.role==='user'?(job?.attachment_names||[]):[];
            const comparisons=item.role==='assistant'?window.SpreadsheetChatFormat.renderComparisons(item.comparisons,item.answer_language):'';
            const imported=item.source==='chatgpt_web';
            return `<article class="chat-message ${esc(item.role)}"><p class="eyebrow">${item.role==='user'?'BẠN':item.role==='error'?'KẾT NỐI / LỖI':imported?'CHATGPT WEB · KẾT QUẢ ĐÃ NHẬP':'TRỢ LÝ ĐỐI CHIẾU'}</p>${imported?'<p class="note">Kết quả từ ChatGPT do bạn nhập. Dashboard hiển thị lại và chưa tự xác minh số liệu.</p>':''}<div class="chat-body">${window.SpreadsheetChatFormat.render(item.content)}${comparisons}</div>${fileNames.length?`<p class="note">Đính kèm: ${esc(fileNames.join(' · '))}</p>`:''}${item.role==='assistant'?evidenceLinks(job)+(item.checked_scope?.length?`<details class="chat-scope"><summary>${imported?'Phạm vi ChatGPT báo đã kiểm tra':'Phạm vi kiểm tra và giả định'}</summary><ul>${item.checked_scope.map(scope=>`<li>${esc(scope)}</li>`).join('')}</ul></details>`:''):''}${item.role==='error'?evidenceLinks(job):''}</article>`;
        }).join('') || '<p class="chat-empty">Nhập câu hỏi để bắt đầu. Bạn có thể hỏi tiếp về các file đã chọn.</p>';
        $('chat-messages').querySelectorAll('[data-chat-result]').forEach(button=>button.onclick=async()=>{
            try{result=await api('/api/compare/results/'+button.dataset.chatResult);save('compare-last-result',result.id);$('compare-status').value='issues';renderResult();$('comparison-results-heading').scrollIntoView({behavior:'smooth'});}catch(error){notice(error.message,true);}
        });
        const running=Object.values(chat.jobs).find(job=>['queued','running'].includes(job.status));
        chatBusy=Boolean(running);
        $('chat-progress').hidden=!running;
        if(running)$('chat-progress').textContent=running.stage+(running.completed_tools?` · ${running.completed_tools} kết quả tính toán đã sẵn sàng`:'');
        renderHandoffFiles();updateChatControls();
    }
    async function restoreChat() {
        if(chatTimer)clearTimeout(chatTimer);
        if(!chatId){renderChat();return;}
        try{
            chat=await api('/api/compare/chat/conversations/'+encodeURIComponent(chatId));
            renderChat();
            if(chatBusy)chatTimer=setTimeout(restoreChat,2500);
        }catch(error){notice(error.message,true);chatTimer=setTimeout(restoreChat,5000);}
    }
    async function initializeChat() {
        applyChatMode();
        if(chatMode==='offline')return;
        if(chatMode==='codex')await ensureChatConnection();
        if(pendingChat){$('chat-request').value=pendingChat.message;notice('Lần gửi trước chưa được xác nhận. Gửi lại để khôi phục yêu cầu mà không tạo bản trùng.');}
        updateChatControls();await restoreChat();
        if(chatId)save('compare-chat-id',chatId);
        if((linkedChat===chatId&&linkedChat!==storedChat)||!books.length){
            const lastUser=chat.messages.findLast(item=>item.role==='user');
            if(lastUser?.files){
                const restored=[];
                for(const id of lastUser.files){try{restored.push(await api('/api/compare/workbooks/'+encodeURIComponent(id)));}catch{/* Unavailable attachment */}}
                books=restored;persist();renderFiles();
            }
        }
    }
    $('chat-form').onsubmit=async event=>{
        event.preventDefault();
        const message=$('chat-request').value.trim();
        if(chatMode==='offline'||!message||chatBusy||busy||!books.length)return;
        if(chatMode==='chatgpt') {
            const prompt=window.SpreadsheetChatGPT.buildPrompt(message,books,$('chat-include-context').checked?previousContext():null);
            try{
                await navigator.clipboard.writeText(prompt);
                $('chat-handoff-status').textContent='Đã sao chép. Mở ChatGPT, đính kèm file rồi dán yêu cầu và gửi. Khi có kết quả, nhập file dashboard-result.json bên dưới để xem trên Dashboard.';
            }catch{
                $('chat-copy-text').value=prompt;$('chat-copy-fallback').hidden=false;$('chat-copy-fallback').open=true;
                $('chat-copy-text').focus();$('chat-copy-text').select();
                $('chat-handoff-status').textContent='Trình duyệt chưa cho sao chép tự động. Bấm Ctrl + C ở nội dung đã chọn, rồi dán vào ChatGPT.';
            }
            return;
        }
        if(!chatReady)return;
        chatBusy=true;updateChatControls();
        const ids=books.map(book=>book.id);
        const same=pendingChat&&pendingChat.message===message&&JSON.stringify(pendingChat.workbooks)===JSON.stringify(ids)&&pendingChat.chat_id===chatId;
        if(!same)pendingChat={chat_id:chatId,message,workbooks:ids,request_id:crypto.randomUUID().replaceAll('-','')};
        save('compare-chat-pending',pendingChat);
        try{
            const job=await api('/api/compare/chat/messages',post(pendingChat));
            chatId=job.chat_id;save('compare-chat-id',chatId);pendingChat=null;save('compare-chat-pending',null);
            history.replaceState(null,'','/compare?chat='+chatId);
            $('compare-results-panel').hidden=true;
            $('chat-request').value='';await restoreChat();
        }catch(error){chatBusy=false;updateChatControls();notice(error.message,true);}
    };
    $('chat-request').oninput=()=>{clearHandoff();updateChatControls();};
    $('chat-include-context').onchange=clearHandoff;
    $('chat-mode').onchange=async()=>{chatMode=$('chat-mode').value;save('compare-analysis-mode',chatMode);applyChatMode();if(chatMode==='codex'&&!chatReady)await ensureChatConnection();if(chatMode!=='offline')await restoreChat();};
    $('chat-import-text').oninput=()=>updateChatControls();
    $('chat-import-file').onchange=async()=>{
        const file=$('chat-import-file').files[0];if(!file)return;
        if(file.size>300000){notice('File kết quả phải nhỏ hơn 300 KB.',true);return;}
        $('chat-import-text').value=await file.text();updateChatControls();
    };
    $('chat-import-submit').onclick=async()=>{
        if(chatBusy||busy||!books.length)return;
        let text=$('chat-import-text').value.trim().replace(/^```(?:json)?\s*/i,'').replace(/\s*```$/,'');
        let answer;try{answer=JSON.parse(text);}catch{notice('Hãy chọn file JSON ChatGPT xuất hoặc dán đúng nội dung JSON kết quả.',true);return;}
        busy=true;updateChatControls();
        try{
            const job=await api('/api/compare/chat/import',post({chat_id:chatId,workbooks:books.map(book=>book.id),answer}));
            chatId=job.chat_id;save('compare-chat-id',chatId);history.replaceState(null,'','/compare?chat='+chatId);
            await restoreChat();$('chat-local-history').open=true;$('chat-import-text').value='';$('chat-import-file').value='';
            notice('Đã đưa kết quả ChatGPT vào Dashboard. Không chạy phân tích Codex.');
            $('chat-local-history').scrollIntoView({behavior:'smooth',block:'start'});
        }catch(error){notice(error.message,true);}finally{busy=false;updateChatControls();}
    };
    $('chat-request').onkeydown=event=>{if((event.ctrlKey||event.metaKey)&&event.key==='Enter'){event.preventDefault();$('chat-form').requestSubmit();}};
    $('chat-new').onclick=()=>{if(chatBusy)return;chatId=null;pendingChat=null;save('compare-chat-id',null);save('compare-chat-pending',null);history.replaceState(null,'','/compare');chat={messages:[],jobs:{}};$('compare-results-panel').hidden=true;renderChat();$('chat-request').focus();};
    document.querySelectorAll('[data-chat-example]').forEach(button=>button.onclick=()=>{$('chat-request').value=button.dataset.chatExample;clearHandoff();updateChatControls();$('chat-request').focus();});
    $('preview-book').onchange=previewSheets;$('preview-sheet').onchange=renderPreview;
    $('compare-upload').onchange=()=>{const files=Array.from($('compare-upload').files);if(!files.length)return;const data=new FormData();files.forEach(file=>data.append('files',file));importBooks({method:'POST',body:data});};
    $('compare-import').onclick=()=>{if(!$('compare-catalog').value){notice('Chọn file đã xuất.',true);return;}importBooks(post({catalog_id:$('compare-catalog').value}));};
    $('compare-catalog').onchange=updateCatalogControls;
    $('compare-delete-all').onclick=async()=>{
        if(busy||deletingExport)return;
        deletingExport=true;updateCatalogControls();
        try{
            await refreshCatalog();
            if(!exportIds.length){notice('Không còn file đã xuất để xóa.');return;}
            const ids=[...exportIds];
            if(!window.confirm(`Xóa TẤT CẢ ${ids.length} file đã xuất đang lưu trên Dashboard?\n\nBao gồm các file ngoài 80 mục đang hiển thị. File sẽ bị xóa khỏi ổ đĩa.\n\nBản sao đã nhập để phân tích vẫn được giữ lại.`))return;
            $('compare-delete-all').textContent='Đang xóa…';
            const data=await api('/api/compare/catalog/delete-all',post({catalog_ids:ids}));
            const message=`Đã xóa ${data.deleted_count} file đã xuất.`+(data.missing_count?` ${data.missing_count} file đã được xóa trước đó.`:'')+(data.failed.length?` Không xóa được ${data.failed.length} file: ${data.failed.map(file=>file.name).join(', ')}. Hãy đóng file trong Excel / WPS rồi thử lại.`:'');
            notice(message,Boolean(data.failed.length));
            try{await refreshCatalog();}catch{notice(message+' Chưa tải lại được danh sách; hãy tải lại trang.',true);}
        }catch(error){notice(error.message,true);}
        finally{deletingExport=false;$('compare-delete-all').textContent='Xóa tất cả file';updateCatalogControls();}
    };
    $('offline-add-comparison').onclick=()=>{addRule('comparison');$('advanced-checks').open=true;$('advanced-checks').scrollIntoView({behavior:'smooth',block:'start'});};
    for(const id of ['workflow-source','workflow-report','workflow-profile','workflow-factory','workflow-rounding'])$(id).onchange=invalidateWorkflow;
    $('workflow-rounding').onchange=()=>{save('compare-workflow-rounding',$('workflow-rounding').value);invalidateWorkflow();};
    $('workflow-factory').oninput=invalidateWorkflow;
    $('workflow-ack').onchange=updateWorkflowControls;
    $('workflow-preview').onclick=async()=>{
        if(workflowBusy||busy)return;
        invalidateWorkflow();workflowBusy=true;updateWorkflowControls();$('workflow-preview').textContent='Đang nhận diện…';
        try{
            workflowPlan=await api('/api/compare/workflows/preview',post({source_id:$('workflow-source').value,report_id:$('workflow-report').value,profile:$('workflow-profile').value,factory_scope:$('workflow-factory').value.trim()||null,comparison_rounding_digits:$('workflow-rounding').value}));
            renderWorkflowPlan();notice('Đã lập kế hoạch. Chưa chạy đối chiếu; hãy kiểm tra phạm vi và phương pháp bên dưới.');
        }catch(error){notice(error.message,true);}
        finally{workflowBusy=false;$('workflow-preview').textContent='Nhận diện & lập kế hoạch';updateWorkflowControls();}
    };
    $('workflow-run').onclick=async()=>{
        if(workflowBusy||busy||!workflowPlan||!$('workflow-ack').checked)return;
        workflowBusy=true;busy=true;updateWorkflowControls();updateCatalogControls();renderRules();$('workflow-run').textContent='Đang chạy quy trình…';
        try{
            result=await api(`/api/compare/workflows/${encodeURIComponent(workflowPlan.id)}/run`,post({acknowledged:true}));
            save('compare-last-result',result.id);$('compare-status').value='issues';renderResult();
            notice(`Hoàn tất quy trình ${workflowPlan.profile_label}: ${result.total} mục kiểm tra, ${result.issue_count} mục cần xem lại. Có thể xuất HTML kèm ảnh.`);
            $('comparison-results-heading').scrollIntoView({behavior:'smooth',block:'start'});
        }catch(error){notice(error.message,true);}
        finally{busy=false;workflowBusy=false;$('workflow-run').textContent='Xác nhận & chạy quy trình';updateWorkflowControls();updateCatalogControls();renderRules();}
    };
    $('add-calculation').onclick=()=>addRule('calculation');$('add-comparison').onclick=()=>addRule('comparison');
    $('run-compare').onclick=async()=>{
        if(busy||!rules.length)return;
        // Commit a field that is still focused before reading the rule list.
        document.activeElement?.blur();busy=true;$('run-compare').disabled=true;updateCatalogControls();$('run-compare').textContent='Đang so sánh…';notice('Đang đối chiếu dữ liệu trên máy…');
        try{result=await api('/api/compare/run',post({rules}));save('compare-last-result',result.id);$('compare-status').value='issues';renderResult();notice(`Đã kiểm tra ${result.total} giá trị. ${result.issue_count} mục cần xem lại.`);$('comparison-results-heading').scrollIntoView({behavior:'smooth',block:'start'});}
        catch(error){notice(error.message,true);}finally{busy=false;$('run-compare').textContent='Chạy so sánh ngoại tuyến →';renderRules();updateCatalogControls();}
    };
    $('compare-status').onchange=()=>fetchResult(0);$('compare-previous').onclick=()=>fetchResult(Math.max(0,result.offset-200));$('compare-next').onclick=()=>fetchResult(result.offset+200);
    $('compare-html').onclick=async()=>{
        if(exportingHtml||!result)return;
        if(!['exact','rounded'].includes(result.comparison_policy)){notice('Kết quả cũ chưa theo quy tắc hiện hành. Hãy chạy lại.',true);return;}
        exportingHtml=true;const button=$('compare-html'), id=result.id, format=$('compare-export-format').value, endpoint=`/api/compare/results/${encodeURIComponent(id)}/download.${format}`;
        button.disabled=true;$('compare-export-format').disabled=true;button.textContent='Đang xuất…';notice(`Đang tạo báo cáo ${format.toUpperCase()} và ảnh đánh dấu trên máy…`);
        try{
            const response=await fetch(endpoint);
            if(!response.ok){const error=await response.json().catch(()=>({}));throw new Error(error.error||'Không xuất được báo cáo HTML.');}
            const url=URL.createObjectURL(await response.blob()), link=document.createElement('a');
            link.href=url;link.download=`Comparison_Report_${id.slice(0,8)}.${format}`;document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
            notice(`Đã xuất báo cáo ${format.toUpperCase()} kèm ảnh đánh dấu.`);
        }catch(error){notice(error.message,true);}
        finally{exportingHtml=false;button.disabled=false;$('compare-export-format').disabled=false;button.textContent='Xuất báo cáo ↓';}
    };
    $('save-compare-preset').onclick=()=>{
        document.activeElement?.blur();const name=$('preset-name').value.trim();if(!name||!rules.length){notice('Nhập tên mẫu và thêm quy tắc trước.',true);return;}
        const presets=read('compare-presets',{});
        presets[name]=rules.map(rule=>{const copy=structuredClone(rule);for(const side of ['left','right'])if(copy[side]){copy[side].slot=Math.max(0,books.findIndex(book=>book.id===copy[side].workbook));delete copy[side].workbook;}return copy;});
        save('compare-presets',presets);renderPresets();notice('Đã lưu mẫu quy tắc và vùng dữ liệu. Nội dung file không được lưu trong mẫu.');
    };
    $('load-compare-preset').onclick=()=>{
        const preset=read('compare-presets',{})[$('compare-presets').value];if(!preset||!books.length){notice('Nhập file và chọn mẫu quy tắc trước.',true);return;}
        rules=structuredClone(preset).map(rule=>{rule.tolerance='0';rule.rounding_digits='';for(const side of ['left','right'])if(rule[side]){const book=books[rule[side].slot]||books[0];rule[side].workbook=book.id;delete rule[side].slot;}return rule;});persist();renderRules();notice('Đã áp dụng mẫu. Sai số đã đặt về 0.');
    };
    $('delete-compare-preset').onclick=()=>{const name=$('compare-presets').value;if(!name)return;const presets=read('compare-presets',{});delete presets[name];save('compare-presets',presets);renderPresets();notice('Đã xóa mẫu.');};
    async function initialize() {
        if(read('compare-rounding-policy-version',0)<1){save('compare-workflow-rounding','1');save('compare-rounding-policy-version',1);}
        $('workflow-rounding').value=read('compare-workflow-rounding','1');
        renderPresets();renderFiles();
        try {
            await refreshCatalog();
            for(const id of read('compare-workbooks',[]).slice(0,10)){try{books.push(await api('/api/compare/workbooks/'+encodeURIComponent(id)));}catch{/* An expired / removed workbook is left out. */}}
            rules=read('compare-rules',[]).map(rule=>({...rule,tolerance:'0',rounding_digits:''})).filter(rule=>books.some(book=>book.id===rule.left?.workbook)&&(!rule.right||books.some(book=>book.id===rule.right.workbook)));
            renderFiles();
            const previous=read('compare-last-result',null);
            if(previous&&(chatMode==='offline'||!chatId)){try{result=await api(`/api/compare/results/${encodeURIComponent(previous)}`);renderResult();notice(['exact','rounded'].includes(result.comparison_policy)?'Đã khôi phục kết quả trước đó. Chưa chạy lại đối chiếu.':'Kết quả cũ chưa theo quy tắc hiện hành. Hãy chạy lại.');}catch{notice('Kết quả trước không còn khả dụng. Hãy chạy lại đối chiếu.');}}
        }catch(error){notice(error.message,true);}
    }
    initialize().then(initializeChat);
})();
