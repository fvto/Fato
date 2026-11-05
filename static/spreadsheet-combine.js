(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  let job = null, current = null, busy = false;
  const notice = (message, error = false) => { $('combine-notice').textContent = message; $('combine-notice').classList.toggle('error', error); };
  const node = (tag, text) => { const element = document.createElement(tag); if (text != null) element.textContent = text; return element; };
  async function api(url, options = {}) {
    const response = await fetch(url, options);
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Chưa xử lý được file. Hãy thử lại.');
    return data;
  }
  function lock(value) {
    busy = value;
    $('combine-upload').disabled = value;
    $('combine-output-name').disabled = value;
    $('combine-schema-mode').disabled = value;
    $('combine-footer-mode').disabled = value;
    $('combine-preview').disabled = value || !job;
    $('combine-reset').disabled = value || !job;
    $('combine-run').disabled = value || !current;
    $('combine-files').querySelectorAll('input').forEach(input => { input.disabled = value; });
  }
  function invalidate() {
    current = null;
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
      file.sheets.forEach((sheet, index) => {
        const row = node('div'); row.className = 'combine-sheet'; row.dataset.file = file.id; row.dataset.sheet = sheet.name;
        const label = node('label'); const check = node('input'); check.type = 'checkbox'; check.checked = index === 0;
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
    job = null; invalidate(); $('combine-files').replaceChildren(); $('combine-file-count').textContent = '0 file'; lock(true);
    notice('Đang đọc cấu trúc các file Excel…');
    try {
      const form = new FormData(); for (const file of files) form.append('files', file);
      job = await api('/api/combine/uploads', {method:'POST', body: form}); showFiles();
      notice(`Đã đọc ${job.files.length} file. Chọn trang tính và dòng tiêu đề, sau đó xem trước.`);
    } catch (error) { notice(error.message, true); }
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
  $('combine-upload').addEventListener('change', event => upload([...event.target.files]));
  $('combine-reset').addEventListener('click', () => { job = null; invalidate(); $('combine-files').replaceChildren(); $('combine-file-count').textContent = '0 file'; notice('Chọn các file cho lượt gộp mới.'); lock(false); });
  const drop = $('combine-drop');
  ['dragenter','dragover'].forEach(name => drop.addEventListener(name, event => { event.preventDefault(); if (!busy) drop.classList.add('dragging'); }));
  ['dragleave','drop'].forEach(name => drop.addEventListener(name, event => { event.preventDefault(); drop.classList.remove('dragging'); }));
  drop.addEventListener('drop', event => upload([...event.dataTransfer.files]));
  $('combine-preview').addEventListener('click', async () => {
    if (!job || busy) return;
    invalidate(); lock(true); notice('Đang ghép tên cột và kiểm tra dữ liệu…');
    try {
      current = await api(`/api/combine/${job.id}/preview`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({selections:selections(), options:options()})});
      showPreview(current); notice('Bản xem trước đã sẵn sàng. Bạn có thể tạo file Excel.');
    } catch (error) { notice(error.message, true); }
    finally { lock(false); }
  });
  $('combine-run').addEventListener('click', async () => {
    if (!job || !current || busy) return;
    lock(true); $('combine-download').hidden = true; notice('Đang tạo và kiểm tra file Excel…');
    try {
      const result = await api(`/api/combine/${job.id}/run`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({selections:selections(), output_name:$('combine-output-name').value, options:options()})});
      $('combine-output-name').value = result.output_name;
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
})();
