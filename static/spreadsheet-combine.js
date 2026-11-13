(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  let job = null, current = null, busy = false, searchIndex = null, searchPage = 1, searchMatched = 0, lastFilters = [], filterDirty = false;
  const notice = (message, error = false) => { $('combine-notice').textContent = message; $('combine-notice').classList.toggle('error', error); $('combine-file-errors').replaceChildren(); $('combine-file-errors').hidden = true; };
  const node = (tag, text) => { const element = document.createElement(tag); if (text != null) element.textContent = text; return element; };
  function showError(error) {
    notice(error.message, true);
    for (const item of error.files || []) $('combine-file-errors').append(node('li', `${item.file}: ${item.error}`));
    $('combine-file-errors').hidden = !(error.files || []).length;
  }
  async function api(url, options = {}) {
    const response = await fetch(url, options);
    let data;
    try { data = await response.json(); }
    catch (_) { throw new Error('Máy chủ chưa trả về kết quả hợp lệ. Hãy thử lại hoặc chia nhỏ lượt gộp.'); }
    if (!response.ok) { const error = new Error(data.error || 'Chưa xử lý được file. Hãy thử lại.'); error.files = data.errors || []; throw error; }
    if (data.task_url) {
      while (true) {
        await new Promise(resolve => setTimeout(resolve, 1000));
        const task = await api(data.task_url);
        if (task.status === 'done') return task.result;
        if (task.status === 'error') { const error = new Error(task.error || 'Chưa xử lý được file.'); error.files = task.errors || []; throw error; }
        notice(task.message || 'Đang xử lý dữ liệu lớn…');
      }
    }
    return data;
  }
  function lock(value) {
    busy = value;
    $('combine-upload').disabled = value;
    $('combine-output-name').disabled = value;
    $('combine-schema-mode').disabled = value;
    $('combine-footer-mode').disabled = value;
    $('combine-csv-encoding').disabled = value;
    $('combine-csv-delimiter').disabled = value;
    $('combine-preview').disabled = value || !job;
    $('combine-reset').disabled = value || !job;
    $('combine-run').disabled = value || !current;
    $('combine-search-create').disabled = value || !current;
    $('combine-search-ready').querySelectorAll('input,select,button').forEach(control => { control.disabled = value; });
    $('combine-search-previous').disabled = value || filterDirty || searchPage <= 1;
    $('combine-search-next').disabled = value || filterDirty || searchPage * 100 >= searchMatched;
    $('combine-files').querySelectorAll('input').forEach(input => { input.disabled = value; });
  }
  function invalidate() {
    current = null;
    searchIndex = null; searchPage = 1; searchMatched = 0; lastFilters = []; filterDirty = false;
    $('combine-search-ready').hidden = true; $('combine-filters').replaceChildren();
    $('combine-result').hidden = true;
    $('combine-download').hidden = true;
    $('combine-download').removeAttribute('href');
    notice('Đã thay đổi lựa chọn. Bấm Xem trước để cập nhật bản gộp.');
  }
  function selections() {
    return [...document.querySelectorAll('.combine-sheet')].filter(row => row.querySelector('[type=checkbox]').checked).map(row => ({file_id: Number(row.dataset.file), sheet: row.dataset.sheet, header_row: Number(row.querySelector('[data-field=header]').value), end_row: Number(row.querySelector('[data-field=end]').value)}));
  }
  function options() { return {schema_mode:$('combine-schema-mode').value, footer_mode:$('combine-footer-mode').value}; }
  function showFiles() {
    $('combine-files').replaceChildren();
    for (const file of job.files) {
      const section = node('div'); section.className = 'combine-file';
      const title = node('p', file.name); title.className = 'combine-file-title'; section.append(title);
      if (file.csv) {
        const delimiter = file.csv.delimiter === '\t' ? 'Tab' : file.csv.delimiter;
        const detail = node('p', `CSV · ${file.csv.encoding} · Dấu phân cách: ${delimiter}`); detail.className = 'combine-help'; section.append(detail);
      }
      for (const warning of file.warnings || []) { const detail = node('p', warning); detail.className = 'combine-help'; section.append(detail); }
      file.sheets.forEach((sheet, index) => {
        const row = node('div'); row.className = 'combine-sheet'; row.dataset.file = file.id; row.dataset.sheet = sheet.name;
        const label = node('label'); const check = node('input'); check.type = 'checkbox'; check.checked = sheet.default_selected ?? index === 0;
        label.append(check, document.createTextNode(sheet.name));
        const headerLabel = node('label', 'Dòng tiêu đề'); const header = node('input'); header.type = 'number'; header.dataset.field = 'header'; header.min = 1; header.max = sheet.rows; header.value = sheet.header_row; header.setAttribute('aria-label', `Dòng tiêu đề ${file.name} / ${sheet.name}`); headerLabel.append(header);
        const endLabel = node('label', 'Dòng cuối'); const end = node('input'); end.type = 'number'; end.dataset.field = 'end'; end.min = 1; end.max = sheet.rows; end.value = sheet.rows; end.setAttribute('aria-label', `Dòng cuối ${file.name} / ${sheet.name}`); endLabel.append(end);
        row.append(label, headerLabel, endLabel, node('span', `${sheet.rows} dòng nguồn · ${sheet.columns} cột · Tiêu đề gợi ý: dòng ${sheet.header_row}`));
        row.addEventListener('input', invalidate); section.append(row);
      });
      $('combine-files').append(section);
    }
    $('combine-file-count').textContent = `${job.files.length} file`;
  }
  async function upload(files) {
    if (busy || !files.length) return;
    const maxFiles = Number($('combine-upload').dataset.maxFiles);
    if (files.length > maxFiles) {
      notice(`Tối đa ${maxFiles} file mỗi lượt gộp. Bạn đã chọn ${files.length} file.`, true);
      $('combine-upload').value = '';
      return;
    }
    const errors = [];
    const maxBytes = Number($('combine-upload').dataset.maxBytes);
    const maxTotalBytes = Number($('combine-upload').dataset.maxTotalBytes);
    for (const file of files) {
      if (!/\.(xlsx|xls|xlsm|csv)$/i.test(file.name)) errors.push({file:file.name, error:'Chỉ hỗ trợ .xls, .xlsx, .xlsm và .csv.'});
      if (file.size > maxBytes) errors.push({file:file.name, error:`Vượt giới hạn ${maxBytes / (1024 * 1024)} MB mỗi file.`});
    }
    if (errors.length || files.reduce((sum, file) => sum + file.size, 0) > maxTotalBytes) {
      showError({message:errors.length ? 'Kiểm tra các file được liệt kê rồi chọn lại.' : 'Tổng file vượt giới hạn 5 GB. Hãy chia nhỏ lượt gộp.', files:errors});
      $('combine-upload').value = '';
      return;
    }
    job = null; invalidate(); $('combine-files').replaceChildren(); $('combine-file-count').textContent = '0 file'; lock(true);
    notice('Đang đọc cấu trúc các file Excel / CSV…');
    try {
      const form = new FormData(); for (const file of files) form.append('files', file);
      const delimiter = $('combine-csv-delimiter').value;
      form.append('csv_options', JSON.stringify({encoding:$('combine-csv-encoding').value, delimiter:delimiter === 'tab' ? '\t' : delimiter}));
      job = await api('/api/combine/uploads', {method:'POST', body: form}); showFiles();
      history.replaceState(null,'',`/combine?job=${encodeURIComponent(job.id)}`);
      notice(`Đã đọc ${job.files.length} file. ${job.estimated_rows > job.xlsx_row_threshold ? 'Ước tính trên 1 triệu dòng: sẽ xuất gói Power Query. ' : ''}Chọn trang tính và dòng tiêu đề, sau đó xem trước.`);
    } catch (error) { showError(error); }
    finally { lock(false); $('combine-upload').value = ''; }
  }
  function display(value, format = '') {
    if (value == null) return '';
    if (typeof value === 'number' && /yyyy-mm-dd/i.test(format)) {
      const iso = new Date(Date.UTC(1899,11,30) + value * 86400000).toISOString();
      return /hh:mm/i.test(format) ? iso.slice(0,19).replace('T',' ') : iso.slice(0,10);
    }
    if (typeof value === 'number' && /^0(?:\.0+)?%$/.test(format)) return `${(value * 100).toFixed((format.split('.')[1] || '').replace('%','').length)}%`;
    return String(value);
  }
  function drawTable(container, headers, rows, formats = []) {
    const table = node('table'), head = node('thead'), tr = node('tr');
    headers.forEach(text => tr.append(node('th', text))); head.append(tr); table.append(head);
    const body = node('tbody'); rows.forEach((values, r) => { const row = node('tr'); values.forEach((value, c) => {const text = display(value, formats[r]?.[c]); const cell = node('td', text); cell.title = text; row.append(cell);}); body.append(row); });
    table.append(body); container.replaceChildren(table);
  }
  function showPreview(result) {
    const powerQuery=result.export_mode==='power_query';
    $('combine-run').textContent=powerQuery ? 'Tạo gói Power Query →' : 'Gộp & tạo file Excel →';
    $('combine-download').textContent=powerQuery ? 'Tải gói Power Query ↓' : 'Tải file Excel đã gộp ↓';
    $('combine-name-help').textContent=powerQuery ? 'Trên 1 triệu dòng: xuất ZIP gồm dữ liệu, truy vấn .pq và hướng dẫn dùng Microsoft Excel có Power Query; không xuất bảng XLSX khổng lồ.' : 'Bạn có thể sửa tên trước khi xuất; định dạng là .xlsx.';
    $('combine-output-name').value = result.output_name;
    $('combine-naming-basis').textContent = result.naming_basis || '';
    $('combine-stats').textContent = `${result.file_count} file · ${result.row_count} dòng · ${result.column_count} cột`;
    $('combine-method').textContent = `Ghép theo tên cột; ${result.footer_rows} dòng tổng đã đối chiếu và ${result.repeated_headers} tiêu đề lặp đã được bỏ. Ba cột truy vết cuối bảng: ${(result.source_columns || ['Source File','Source Sheet','Source Row']).join(', ')}.`;
    $('combine-warnings').textContent = (result.warnings || []).join(' ');
    $('combine-warnings').hidden = !(result.warnings || []).length;
    drawTable($('combine-preview-table'), result.columns, result.preview, result.preview_formats);
    drawTable($('combine-sources'), ['File', 'Trang tính', 'Tiêu đề', 'Dòng cuối', 'Dòng dữ liệu', 'Cột nguồn', 'Cột thiếu so với bảng gộp', 'Dòng tổng bỏ qua'], result.sources.map(s => [s.file, s.sheet, s.header_row, s.end_row, s.rows, s.columns, (s.missing_columns || []).join(', '), s.footer_rows]));
    $('combine-result').hidden = false;
  }
  function markSearchDirty() {
    filterDirty=true;
    $('combine-search-stats').textContent='Điều kiện đã thay đổi. Bấm Tìm và tính tổng để cập nhật.';
    $('combine-size-totals').hidden=true; $('combine-search-table').hidden=true;
    $('combine-top-issues').hidden=true;
    $('combine-search-page').textContent='';
    $('combine-search-previous').disabled=true; $('combine-search-next').disabled=true;
  }
  function addFilter() {
    if ($('combine-filters').children.length >= 30) { notice('Tối đa 30 điều kiện lọc.',true); return; }
    const line = node('div'); line.className = 'combine-filter';
    const column = node('select'); column.setAttribute('aria-label','Cột cần tìm');
    searchIndex.columns.forEach((name,index) => { const option=node('option',name); option.value=String(index); column.append(option); });
    const preferred=searchIndex.columns.findIndex(name => name.trim().toLowerCase()==='mold code');
    if (!$('combine-filters').children.length && preferred>=0) column.value=String(preferred);
    const mode=node('select'); mode.setAttribute('aria-label','Cách tìm');
    for (const [value,label] of [['exact','Khớp chính xác'],['contains','Có chứa']]) { const option=node('option',label); option.value=value; mode.append(option); }
    const value=node('input'); value.type='text'; value.maxLength=200; value.placeholder='Nhập giá trị cần tìm'; value.setAttribute('aria-label','Giá trị cần tìm');
    const remove=node('button','Xóa'); remove.type='button'; remove.addEventListener('click',() => { line.remove(); markSearchDirty(); });
    [column,mode,value].forEach(control => control.addEventListener('input',markSearchDirty));
    line.append(column,mode,value,remove); $('combine-filters').append(line);
    markSearchDirty();
  }
  function searchFilters() {
    return [...$('combine-filters').children].map(line => ({column:Number(line.children[0].value),mode:line.children[1].value,value:line.children[2].value.trim()})).filter(filter => filter.value);
  }
  function showSearch(result, initial=false) {
    filterDirty=false; $('combine-size-totals').hidden=false; $('combine-search-table').hidden=false;
    $('combine-top-issues').hidden=false;
    searchPage=result.page || 1; searchMatched=initial ? result.row_count : result.matched;
    $('combine-search-stats').textContent = `${searchMatched.toLocaleString('vi-VN')} dòng phù hợp`;
    const messages=[];
    if (result.invalid_quantities) messages.push(`${result.invalid_quantities.toLocaleString('vi-VN')} dòng có Number of Pairs không phải số; các dòng này chưa được cộng vào tổng. Kiểm tra nguồn.`);
    if (result.omitted_sizes) messages.push('Bảng chỉ hiển thị 1.000 nhóm size đầu. Hãy thêm điều kiện để thu hẹp kết quả.');
    if (result.top_issues_truncated) messages.push('Top Issue chỉ hiển thị 1.000 nhóm đầu. Hãy lọc mã giày hoặc Size Range.');
    $('combine-search-warning').textContent=messages.join(' '); $('combine-search-warning').hidden=!messages.length;
    const withRange=Boolean(result.range_column);
    drawTable($('combine-size-totals'),withRange ? ['Size','Size Range','Số dòng','Grand Total (đôi)'] : ['Size','Số dòng','Grand Total (đôi)'],result.groups.map(group => withRange ? [group.size,group.size_range,group.rows,group.total] : [group.size,group.rows,group.total]));
    drawTable($('combine-top-issues'),['Size','Size Range','Hạng','Issue','Số đôi DRR'],(result.top_issues || []).map(item => [item.size,item.size_range,item.rank,item.issue,item.total]));
    $('combine-top-issues-help').textContent=initial ? 'Nhập bộ lọc và bấm Tìm để tính top 3 Issue DRR.' : result.top_issues_available ? 'Chỉ Unit Type = DRR, xếp theo tổng Number of Pairs giảm dần trong phạm vi bộ lọc. Khi bằng nhau, sắp theo tên Issue.' : 'Cần có cột Unit Type và Issue để tính top 3.';
    drawTable($('combine-search-table'),searchIndex.columns,result.rows || []);
    $('combine-search-page').textContent=initial ? 'Nhập điều kiện rồi bấm Tìm' : `Trang ${searchPage} / ${Math.max(1,Math.ceil(searchMatched / 100))}`;
    if (initial) searchMatched=0;
  }
  $('combine-search-create').addEventListener('click',async () => {
    if (!job || !current || busy) return;
    lock(true); notice('Đang chuẩn bị dữ liệu tra cứu…');
    try {
      searchIndex=await api(`/api/combine/${job.id}/search-index`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({selections:selections(),options:options()})});
      $('combine-filters').replaceChildren(); addFilter(); $('combine-search-ready').hidden=false;
      $('combine-search-scope').textContent=`Đã chuẩn bị ${searchIndex.row_count.toLocaleString('vi-VN')} dòng từ các bảng đã chọn. Các điều kiện được áp dụng đồng thời.`;
      showSearch(searchIndex,true); notice('Đã chuẩn bị tra cứu. Nhập mã giày hoặc thêm điều kiện theo cột.');
    } catch (error) { showError(error); }
    finally { lock(false); }
  });
  async function runSearch(page, filters) {
    if (!searchIndex || busy) return;
    lock(true); notice('Đang tìm và tính tổng số đôi…');
    try {
      const result=await api(`/api/combine/${job.id}/search`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({index_id:searchIndex.index_id,filters,page})});
      lastFilters=filters; showSearch(result); notice('Đã cập nhật kết quả tra cứu.');
    } catch (error) { showError(error); }
    finally { lock(false); }
  }
  $('combine-filter-add').addEventListener('click',addFilter);
  $('combine-search-run').addEventListener('click',() => runSearch(1,searchFilters()));
  $('combine-search-clear').addEventListener('click',() => { $('combine-filters').replaceChildren();addFilter();lastFilters=[];showSearch(searchIndex,true);lock(false); });
  $('combine-search-previous').addEventListener('click',() => runSearch(searchPage-1,lastFilters));
  $('combine-search-next').addEventListener('click',() => runSearch(searchPage+1,lastFilters));
  $('combine-upload').addEventListener('change', event => upload([...event.target.files]));
  $('combine-reset').addEventListener('click', () => { job = null; invalidate(); $('combine-files').replaceChildren(); $('combine-file-count').textContent = '0 file'; history.replaceState(null,'','/combine'); notice('Chọn các file cho lượt gộp mới.'); lock(false); });
  const drop = $('combine-drop');
  ['dragenter','dragover'].forEach(name => drop.addEventListener(name, event => { event.preventDefault(); if (!busy) drop.classList.add('dragging'); }));
  ['dragleave','drop'].forEach(name => drop.addEventListener(name, event => { event.preventDefault(); drop.classList.remove('dragging'); }));
  drop.addEventListener('drop', event => upload([...event.dataTransfer.files]));
  $('combine-preview').addEventListener('click', async () => {
    if (!job || busy) return;
    invalidate(); lock(true); notice('Đang ghép tên cột và kiểm tra dữ liệu…');
    try {
      current = await api(`/api/combine/${job.id}/preview`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({selections:selections(), options:options()})});
      showPreview(current); notice(current.export_mode==='power_query' ? 'Trên 1 triệu dòng: bạn có thể tra cứu trên Dashboard hoặc tạo gói Power Query.' : 'Bản xem trước đã sẵn sàng. Bạn có thể tra cứu hoặc tạo file Excel.');
    } catch (error) { notice(error.message, true); }
    finally { lock(false); }
  });
  $('combine-run').addEventListener('click', async () => {
    if (!job || !current || busy) return;
    lock(true); $('combine-download').hidden = true; notice(current.export_mode==='power_query' ? 'Đang tạo và kiểm tra gói Power Query…' : 'Đang tạo và kiểm tra file Excel…');
    try {
      const result = await api(`/api/combine/${job.id}/run`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({selections:selections(), output_name:$('combine-output-name').value, options:options()})});
      $('combine-output-name').value = result.kind==='power_query' ? result.output_name.replace(/\.zip$/,'.xlsx') : result.output_name;
      $('combine-download').href = result.download_url; $('combine-download').download = result.output_name; $('combine-download').hidden = false;
      notice(`Đã gộp ${result.file_count} file thành ${result.row_count} dòng dữ liệu. Tên file: ${result.output_name}`);
    } catch (error) { notice(error.message, true); }
    finally { lock(false); }
  });
  ['combine-schema-mode','combine-footer-mode'].forEach(id => $(id).addEventListener('change', invalidate));
  $('combine-output-name').addEventListener('input', () => {
    $('combine-download').hidden = true;
    $('combine-download').removeAttribute('href');
    notice('Đã đổi tên file. Bấm Gộp & tạo file Excel để xuất với tên mới.');
  });
  api('/api/combine/example').then(({example}) => {
    if (!example) return;
    $('combine-example-detail').textContent = `${example.file_count} file · ${example.row_count} dòng dữ liệu · ${example.column_count} cột · Đã đối chiếu và bỏ ${example.footer_rows} dòng tổng cuối file.`;
    $('combine-example-download').href = example.download_url;
    $('combine-example-download').download = example.output_name;
    $('combine-example-download').title = example.output_name;
    $('combine-example').hidden = false;
  }).catch(() => {});
  const previousJob=new URLSearchParams(location.search).get('job');
  if (previousJob) {
    lock(true);notice('Đang khôi phục các file đã tải…');
    api(`/api/combine/${encodeURIComponent(previousJob)}/inputs`).then(value => {
      job=value;showFiles();notice('Đã khôi phục file nguồn. Bấm Xem trước rồi Chuẩn bị tra cứu.');
    }).catch(showError).finally(() => lock(false));
  }
})();
