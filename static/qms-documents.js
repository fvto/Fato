'use strict';
window.QmsDocuments = (() => {
    const byId = id => document.getElementById(id);
    const pageSize = 20;
    let catalog = null, page = 0, requestNumber = 0, refreshing = false;
    let bulkJob = null, bulkStarting = false, bulkTimer = null, bulkScope = '', bulkRequest = 0;
    const sourceName = () => byId('qa-pdf-source').value === 'manual' ? 'QA Manual' : 'QA Document';
    const scopeKey = () => `qms-pdf-bulk:${byId('qa-pdf-plant').value}:${byId('qa-pdf-source').value}`;
    const textNode = (tag, text, className) => {
        const node = document.createElement(tag);
        node.textContent = text;
        if (className) node.className = className;
        return node;
    };
    const collectionOf = item => item.collection || (item.source === 'document' || item.code ? 'records' : item.kind === 'BDTS' ? 'bdts' : 'hfpa');
    function scopeItems() {
        const source = byId('qa-pdf-source').value, collection = byId('qa-pdf-collection').value;
        return (catalog?.documents || []).filter(item => item.source === source &&
            (!collection || collectionOf(item) === collection));
    }
    function filtered() {
        const terms = byId('qa-pdf-search').value.toLowerCase().trim().split(/\s+/).filter(Boolean);
        const language = byId('qa-pdf-language').value;
        return scopeItems().filter(item => (!language || item.language === language) && terms.every(term =>
                `${item.name} ${item.code} ${item.kind} ${item.project} ${item.shoe} ${item.process} ${item.date}`.toLowerCase().includes(term)));
    }
    function reconciliation() {
        const source = byId('qa-pdf-source').value;
        const summaries = [];
        ['manual', 'document'].filter(key => !source || source === key).forEach(key => {
            const items = (catalog?.documents || []).filter(item => item.source === key);
            const rows = catalog?.sources?.[key]?.rows;
            const count = items.filter(item => collectionOf(item) === 'records').length;
            summaries.push(`${key === 'manual' ? 'QA Manual' : 'QA Document'}: ${rows == null ? '' : rows + ' dòng SFC · '}${count} liên kết PDF trong bảng`);
            if (key === 'manual') summaries.push(`Tài liệu bổ sung riêng: HFPA Guide ${items.filter(item => collectionOf(item) === 'hfpa').length} · BDTS ${items.filter(item => collectionOf(item) === 'bdts').length}`);
        });
        byId('qa-pdf-reconciliation').textContent = catalog?.refreshed_at ? summaries.join('. ') +
            (source === 'document' ? '. Một dòng có thể có PDF ở nhiều ngôn ngữ.' : '.') : '';
    }
    function render() {
        const items = filtered(), pages = Math.max(1, Math.ceil(items.length / pageSize));
        page = Math.min(page, pages - 1);
        const list = byId('qa-pdf-list');
        list.replaceChildren();
        items.slice(page * pageSize, (page + 1) * pageSize).forEach(item => {
            const row = textNode('div', '', 'qa-pdf-row');
            const info = textNode('div', '', 'qa-pdf-info');
            info.appendChild(textNode('strong', `${item.code ? item.code + ' · ' : ''}${item.name}`));
            info.appendChild(textNode('small', [item.source === 'manual' ? 'QA Manual' : 'QA Document',
                item.kind, item.project, item.language || 'Original', item.date].filter(Boolean).join(' · ')));
            const download = textNode('button', '↓ PDF', 'btn-action qa-pdf-download');
            download.type = 'button';
            download.title = `Download ${item.name}${item.language ? ' (' + item.language + ')' : ''}`;
            download.addEventListener('click', () => downloadPdf(item, download));
            row.append(info, download);
            list.appendChild(row);
        });
        if (!items.length) list.appendChild(textNode('p', catalog?.refreshed_at
            ? 'No PDFs match these filters.' : 'Chưa có danh mục cho nhà máy này. Bấm “Cập nhật từ SFC” để lấy tài liệu.'));
        if (!refreshing) byId('qa-pdf-status').textContent = catalog?.refreshed_at
            ? `${items.length} / ${scopeItems().length} liên kết PDF trong nhóm đã chọn · ${catalog.plant} · Updated: ${new Date(catalog.refreshed_at).toLocaleString('vi-VN', {timeZone: 'Asia/Ho_Chi_Minh'})}`
            : 'Cần đăng nhập SFC để cập nhật danh mục và tải PDF.';
        byId('qa-pdf-page').textContent = `${page + 1} / ${pages}`;
        byId('qa-pdf-prev').disabled = page === 0;
        byId('qa-pdf-next').disabled = page + 1 >= pages;
        reconciliation();
        byId('qa-pdf-download-all').textContent = `↓ Tải tất cả ${sourceName()} (${items.length} PDF)`;
        byId('qa-pdf-download-all').disabled = !items.length || refreshing || bulkStarting || ['queued', 'running'].includes(bulkJob?.status);
        byId('qa-pdf-bulk-resume').disabled = refreshing || bulkStarting;
    }
    async function load() {
        const number = ++requestNumber;
        byId('qa-pdf-status').textContent = 'Đang đọc danh mục PDF…';
        try {
            const response = await fetch(`/api/qms/documents?plant=${encodeURIComponent(byId('qa-pdf-plant').value)}`, {cache: 'no-store'});
            const data = await response.json();
            if (!response.ok) throw new Error(data.error || 'Chưa đọc được danh mục PDF.');
            if (number !== requestNumber) return;
            catalog = data;
            page = 0;
            render();
            restoreBulk();
        } catch (error) {
            if (number !== requestNumber) return;
            catalog = null;
            render();
            byId('qa-pdf-status').textContent = error.message;
        }
    }
    function restoreBulk() {
        const scope = scopeKey();
        if (scope === bulkScope) return;
        bulkScope = scope;
        bulkRequest++;
        clearTimeout(bulkTimer);
        bulkJob = null;
        byId('qa-pdf-bulk-link').hidden = true;
        byId('qa-pdf-bulk-resume').hidden = true;
        byId('qa-pdf-bulk-status').textContent = '';
        const saved = localStorage.getItem(scope);
        if (saved) watchBulk(saved, scope);
        render();
    }
    function showBulk(job) {
        bulkJob = job;
        const link = byId('qa-pdf-bulk-link');
        link.hidden = !['done', 'partial'].includes(job.status);
        byId('qa-pdf-bulk-resume').hidden = !job.resumable;
        if (job.status === 'done') {
            link.href = job.url;
            link.textContent = `↓ Tải ZIP ${sourceName()} (${job.completed} PDF)`;
            byId('qa-pdf-bulk-status').textContent = `Đã lấy đủ ${job.completed}/${job.total} PDF · ${sourceName()} · ${job.plant}. ZIP đã sẵn sàng.`;
        } else if (job.status === 'partial') {
            link.href = job.url;
            link.textContent = `↓ Tải ZIP ${sourceName()} CHƯA ĐỦ (${job.completed}/${job.total} PDF)`;
            const failures = (job.failed || []).map(item => `${item.code || item.name}: ${item.message}`).join(' ');
            byId('qa-pdf-bulk-status').textContent = `Đã lưu ${job.completed}/${job.total} PDF · ${job.plant}, còn thiếu ${job.total - job.completed}. ${failures || job.error || ''} ZIP có các file tải được và danh sách PDF_CHUA_TAI_DUOC.csv. Bấm Tiếp tục chỉ lấy phần còn thiếu.`;
        } else if (job.status === 'error') {
            byId('qa-pdf-bulk-status').textContent = `Đã lấy ${job.completed}/${job.total} PDF · ${job.plant}. ${job.error || 'Lượt tải bị lỗi.'}${job.current ? ' File đang xử lý: ' + job.current + '.' : ''} Chưa tạo ZIP hoàn chỉnh.${job.cached != null ? ' PDF thành công được lưu riêng để tiếp tục.' : ''}`;
        } else {
            byId('qa-pdf-bulk-status').textContent = `Đang tải ${job.completed}/${job.total} PDF · ${job.plant}${job.current ? ' · ' + job.current : ''}. Có thể tải lại trang để tiếp tục theo dõi.`;
        }
        render();
    }
    async function resumeBulk() {
        if (!bulkJob?.resumable || bulkStarting) return;
        const id = bulkJob.id, scope = scopeKey();
        bulkStarting = true;
        render();
        try {
            const response = await fetch(`/api/qms/bulk/${encodeURIComponent(id)}/resume`, {method: 'POST'});
            const data = await response.json();
            if (!response.ok) throw new Error(data.error || 'Chưa tiếp tục được lượt tải PDF.');
            if (scope !== scopeKey()) return;
            showBulk(data);
            watchBulk(id, scope);
        } catch (error) {
            if (scope === scopeKey()) byId('qa-pdf-bulk-status').textContent = error.message;
        } finally {
            bulkStarting = false;
            render();
        }
    }
    async function watchBulk(id, scope = scopeKey()) {
        if (scope !== scopeKey()) return;
        const number = ++bulkRequest;
        clearTimeout(bulkTimer);
        try {
            const response = await fetch(`/api/qms/bulk/${encodeURIComponent(id)}`, {cache: 'no-store'});
            const data = await response.json();
            if (number !== bulkRequest || scope !== scopeKey()) return;
            if (!response.ok) {
                if (response.status === 404) localStorage.removeItem(scope);
                throw new Error(data.error || 'Chưa đọc được tiến độ tải PDF.');
            }
            if (data.source !== byId('qa-pdf-source').value || data.plant !== byId('qa-pdf-plant').value) {
                localStorage.removeItem(scope);
                return;
            }
            showBulk(data);
            if (['queued', 'running'].includes(data.status)) bulkTimer = setTimeout(() => watchBulk(id, scope), 2500);
        } catch (error) {
            if (number !== bulkRequest || scope !== scopeKey()) return;
            byId('qa-pdf-bulk-status').textContent = error.message;
            // A disconnected poll never starts or repeats a download job.
            if (bulkJob && ['queued', 'running'].includes(bulkJob.status)) bulkTimer = setTimeout(() => watchBulk(id, scope), 5000);
        }
    }
    async function downloadAll() {
        const items = filtered();
        if (!items.length || bulkStarting || ['queued', 'running'].includes(bulkJob?.status)) return;
        const source = byId('qa-pdf-source').value, scope = scopeKey();
        bulkStarting = true;
        render();
        byId('qa-pdf-bulk-status').textContent = `Đang chuẩn bị ${items.length} PDF trên tất cả các trang…`;
        try {
            const response = await fetch('/api/qms/bulk', {method: 'POST', headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({plant: catalog.plant, source, ids: items.map(item => item.id)})});
            const data = await response.json();
            if (!response.ok) throw new Error(data.error || 'Chưa khởi động được lượt tải tất cả PDF.');
            localStorage.setItem(scope, data.id);
            if (scope !== scopeKey()) return;
            showBulk(data);
            watchBulk(data.id, scope);
        } catch (error) {
            if (scope === scopeKey()) byId('qa-pdf-bulk-status').textContent = error.message;
        } finally {
            bulkStarting = false;
            render();
        }
    }
    async function refresh() {
        if (refreshing) return;
        refreshing = true;
        byId('qa-pdf-refresh').disabled = true;
        byId('qa-pdf-plant').disabled = true;
        byId('qa-pdf-status').textContent = 'Đang cào lại QMS và toàn bộ trang tài liệu. Vui lòng chờ…';
        try {
            const response = await fetch('/api/qms/refresh', {method: 'POST', headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({plant: byId('qa-pdf-plant').value})});
            const data = await response.json();
            if (!response.ok) throw new Error(data.error || 'Chưa cập nhật được QMS.');
            refreshing = false;
            await Promise.all([load(), window.loadMemory()]);
        } catch (error) {
            byId('qa-pdf-status').textContent = error.message;
        } finally {
            refreshing = false;
            byId('qa-pdf-refresh').disabled = false;
            byId('qa-pdf-plant').disabled = false;
        }
    }
    async function downloadPdf(item, button) {
        button.disabled = true;
        button.textContent = '…';
        try {
            const response = await fetch(item.url, {cache: 'no-store'});
            if (!response.ok) {
                const data = await response.json();
                throw new Error(data.error || 'Chưa tải được PDF.');
            }
            const blob = await response.blob();
            if (blob.type !== 'application/pdf') throw new Error('Máy chủ chưa trả về file PDF.');
            const disposition = response.headers.get('Content-Disposition') || '';
            const encoded = disposition.match(/filename\*=UTF-8''([^;]+)/i);
            const plain = disposition.match(/filename="([^"]+)"|filename=([^;]+)/i);
            const filename = encoded ? decodeURIComponent(encoded[1]) : plain ? (plain[1] || plain[2]).trim() : `${item.code || item.name}.pdf`;
            const url = URL.createObjectURL(blob), link = document.createElement('a');
            link.href = url;
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            link.remove();
            setTimeout(() => URL.revokeObjectURL(url), 60000);
            byId('qa-pdf-status').textContent = `Đã tải: ${filename}`;
        } catch (error) {
            byId('qa-pdf-status').textContent = error.message;
        } finally {
            button.disabled = false;
            button.textContent = '↓ PDF';
        }
    }
    function open(source) {
        byId('qa-pdf-source').value = source;
        byId('qa-pdf-search').value = '';
        byId('qa-pdf-language').value = '';
        byId('qa-pdf-collection').value = 'records';
        page = 0;
        const panel = byId('qa-pdf-library');
        panel.open = true;
        if (catalog) render();
        restoreBulk();
        panel.scrollIntoView({behavior: 'smooth', block: 'nearest'});
        byId('qa-pdf-search').focus({preventScroll: true});
    }
    document.addEventListener('DOMContentLoaded', () => {
        byId('qa-pdf-library').addEventListener('toggle', () => {
            if (byId('qa-pdf-library').open && !catalog) load();
        });
        ['qa-pdf-collection', 'qa-pdf-language', 'qa-pdf-search'].forEach(id =>
            byId(id).addEventListener(id.endsWith('search') ? 'input' : 'change', () => {page = 0; render();}));
        byId('qa-pdf-plant').addEventListener('change', () => {catalog = null; restoreBulk(); render(); load();});
        byId('qa-pdf-source').addEventListener('change', () => {byId('qa-pdf-collection').value = 'records'; page = 0; restoreBulk(); render();});
        byId('qa-pdf-refresh').addEventListener('click', refresh);
        byId('qa-pdf-download-all').addEventListener('click', downloadAll);
        byId('qa-pdf-bulk-resume').addEventListener('click', resumeBulk);
        byId('qa-pdf-prev').addEventListener('click', () => {page--; render();});
        byId('qa-pdf-next').addEventListener('click', () => {page++; render();});
    });
    return {open};
})();
