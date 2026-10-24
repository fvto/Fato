
        let currentActivePlant = 'VH';

        // Theme switch
        const themeBtn = document.getElementById('theme-toggle');
        themeBtn.addEventListener('click', () => {
            const current = document.body.getAttribute('data-theme');
            const target = current === 'dark' ? 'light' : 'dark';
            document.body.setAttribute('data-theme', target);
            themeBtn.innerHTML = target === 'dark' ? '<i class="fa-solid fa-moon"></i>' : '<i class="fa-solid fa-sun"></i>';
            showToast(`Đã chuyển sang chế độ: ${target === 'dark' ? 'Dark Mode' : 'Light Mode'}`);
        });

        function showToast(msg, isSuccess = true) {
            const toast = document.getElementById('toast');
            const txt = document.getElementById('toast-text');
            txt.innerText = msg;
            toast.style.borderLeftColor = isSuccess ? 'var(--accent-green)' : 'var(--accent-red)';
            toast.classList.add('show');
            setTimeout(() => toast.classList.remove('show'), 3500);
        }

        function addLog(msg, type = 'info') {
            const box = document.getElementById('activity-log');
            const entry = document.createElement('div');
            entry.className = 'log-entry';
            const time = new Date().toLocaleTimeString();
            entry.innerHTML = `<span class="log-time">[${time}]</span> <span class="log-text ${type}">${msg}</span>`;
            box.prepend(entry);
        }

        function applyPreset(text) {
            document.getElementById('ai-prompt-input').value = text;
            executeCommand();
        }

        // Switch between the 7 Factories
        async function switchPlant(plantName) {
            if (!plantName) return;
            addLog(`Đang chuyển sang Nhà máy: "${plantName}"...`, 'info');

            // Update UI buttons immediately
            document.querySelectorAll('.plant-btn').forEach(btn => btn.classList.remove('active'));
            const activeBtn = document.getElementById(`plant-btn-${plantName}`);
            if (activeBtn) activeBtn.classList.add('active');

            try {
                const res = await fetch('/api/switch_factory', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ factory: plantName })
                });
                const data = await res.json();
                if (data.success) {
                    currentActivePlant = data.current_plant || plantName;
                    document.getElementById('active-plant-text').innerText = `Nhà máy: ${currentActivePlant}`;
                    showToast(data.message || `Đã chuyển sang Nhà máy: ${currentActivePlant}`);
                    addLog(data.message || `Đã chuyển sang Nhà máy: ${currentActivePlant}`, 'success');
                    if (data.screenshot) {
                        document.getElementById('mirror-img').src = data.screenshot;
                    }
                    setTimeout(refreshStatus, 1500);
                } else {
                    showToast(data.error || 'Lỗi khi chuyển Nhà máy', false);
                    addLog(`Lỗi: ${data.error || 'Không đổi được nhà máy'}`, 'warn');
                }
            } catch(e) {
                showToast('Lỗi kết nối máy chủ', false);
            }
        }

        let nikeBusy = false;
        let confirmedNike = false;
        let nikeRevision = 0;

        function updateNikeMode(checked) {
            document.getElementById('nike-toggle').checked = checked;
            document.getElementById('nike-mode-text').innerText = checked ? 'External (NIKE)' : 'Nội bộ';
        }

        async function toggleNike() {
            const checkbox = document.getElementById('nike-toggle');
            const requested = checkbox.checked;
            nikeBusy = true;
            nikeRevision++;
            checkbox.disabled = true;
            document.getElementById('btn-run-search').disabled = true;
            document.getElementById('btn-run-excel').disabled = true;
            document.getElementById('nike-mode-text').innerText = 'Đang đồng bộ...';
            try {
                const response = await fetch('/api/nike', {
                    method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ nike: requested })
                });
                const data = await response.json();
                if (!response.ok || !data.success) throw new Error(data.error || 'Không đổi được chế độ NIKE');
                confirmedNike = data.filter_result.nike;
                updateNikeMode(confirmedNike);
                if (data.screenshot) document.getElementById('mirror-img').src = data.screenshot;
                const message = `Đã chuyển chế độ: ${confirmedNike ? 'External (NIKE)' : 'Nội bộ'}`;
                showToast(message);
                addLog(message, 'success');
            } catch (error) {
                updateNikeMode(confirmedNike);
                showToast(error.message, false);
                addLog(error.message, 'warn');
            } finally {
                nikeBusy = false;
                checkbox.disabled = false;
                document.getElementById('btn-run-search').disabled = false;
                document.getElementById('btn-run-excel').disabled = false;
                refreshStatus();
            }
        }

        // Fetch Live Status
        async function refreshStatus() {
            const requestedRevision = nikeRevision;
            try {
                const res = await fetch('/api/status');
                const data = await res.json();
                if (!nikeBusy && requestedRevision === nikeRevision && data.form && typeof data.form.nike === 'boolean') {
                    confirmedNike = data.form.nike;
                    updateNikeMode(confirmedNike);
                }
                if (data.connected) {
                    document.getElementById('mirror-url').innerText = data.url;
                    document.getElementById('mirror-title').innerText = data.title || 'SFC';
                    if (data.screenshot) {
                        document.getElementById('mirror-img').src = data.screenshot;
                    }
                    if (data.current_plant) {
                        currentActivePlant = data.current_plant;
                        document.getElementById('active-plant-text').innerText = `Nhà máy: ${currentActivePlant}`;
                        document.querySelectorAll('.plant-btn').forEach(btn => btn.classList.remove('active'));
                        const btn = document.getElementById(`plant-btn-${currentActivePlant}`);
                        if (btn) btn.classList.add('active');
                    }
                    if (data.form && data.form.selectedDept) {
                        const sel = document.getElementById('factory-select');
                        for (let opt of sel.options) {
                            if (opt.text === data.form.selectedDept) {
                                opt.selected = true;
                                break;
                            }
                        }
                    }
                    if (data.mqaa_status && data.mqaa_status.progress) {
                        document.getElementById('mqaa-status-text').innerText = data.mqaa_status.progress;
                    }
                } else {
                    document.getElementById('mirror-url').innerText = "Chưa kết nối được Chrome cổng 9222";
                }
            } catch(e) {
                console.error(e);
            }
        }

        // Load Lines into Dropdown
        async function loadFactories() {
            try {
                const res = await fetch('/api/factories');
                const data = await res.json();
                const sel = document.getElementById('factory-select');
                sel.innerHTML = '<option value="">-- Chọn Chuyền / Xưởng sản xuất --</option>';
                if (data.factories) {
                    data.factories.forEach(f => {
                        const opt = document.createElement('option');
                        opt.value = f.text;
                        opt.innerText = f.text;
                        sel.appendChild(opt);
                    });
                }
            } catch(e) {
                console.error(e);
            }
        }

        // Load Memory & QMS modules
        async function loadMemory() {
            try {
                const res = await fetch('/api/memory');
                const data = await res.json();
                const qmsList = data.qms_modules || [];
                const qmsSel = document.getElementById('qms-jump-select');
                qmsSel.innerHTML = '<option value="">-- Chọn tính năng QMS muốn mở --</option>';
                
                qmsList.forEach(m => {
                    const opt = document.createElement('option');
                    opt.value = m.href;
                    opt.innerText = m.name;
                    qmsSel.appendChild(opt);
                });

                document.getElementById('memory-count').innerText = `7 Nhà máy + ${qmsList.length} QMS + ${data.factories_and_lines ? data.factories_and_lines.length : 138} Chuyền`;
            } catch(e) {
                console.error(e);
            }
        }

        async function jumpToQmsModule() {
            const sel = document.getElementById('qms-jump-select');
            const targetUrl = sel.value;
            const targetName = sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text : '';
            if (!targetUrl) {
                showToast("Vui lòng chọn tính năng QMS trong danh sách!", false);
                return;
            }

            addLog(`Đang chuyển đến: "${targetName}"...`, 'info');
            try {
                const res = await fetch('/api/navigate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ destination: targetUrl })
                });
                const data = await res.json();
                if (data.success) {
                    showToast(`Đã mở: ${targetName}`);
                    addLog(`Đã chuyển đến: ${targetName}`, 'success');
                    if (data.screenshot) {
                        document.getElementById('mirror-img').src = data.screenshot;
                    }
                    setTimeout(refreshStatus, 1500);
                } else {
                    showToast(data.error || 'Lỗi điều hướng', false);
                }
            } catch(e) {
                showToast('Lỗi kết nối máy chủ', false);
            }
        }

        // Load MQAA Summary
        async function loadMqaaSummary() {
            try {
                const res = await fetch('/api/mqaa/summary');
                const data = await res.json();
                if (data.has_data && data.rows) {
                    document.getElementById('scorecard-stat').innerText = `${data.perfect_count}/${data.total_models} Đạt chuẩn 100%`;
                    const tbody = document.getElementById('mqaa-table-body');
                    tbody.innerHTML = '';
                    data.rows.forEach(r => {
                        const tr = document.createElement('tr');
                        const isOk = (r.status || '').includes('100%');
                        tr.innerHTML = `
                            <td>${r.stt}</td>
                            <td><span style="font-weight:600; color: ${r.cat === 'Bottom' ? '#38bdf8' : '#c084fc'};">${r.cat}</span></td>
                            <td><strong>${r.model}</strong></td>
                            <td style="text-align:center;">${r.match_q}</td>
                            <td style="text-align:center;">${r.match_rmk}</td>
                            <td><span class="status-badge-sm ${isOk ? 'status-ok' : 'status-warn'}">${r.status}</span></td>
                        `;
                        tbody.appendChild(tr);
                    });
                }
            } catch(e) {
                console.error(e);
            }
        }

        let commandBusy = false;
        let commandTerminal = false;

        async function consumeCommandStream(res) {
            if (!res.ok) {
                const error = await res.json().catch(() => ({}));
                throw new Error(error.error || `HTTP ${res.status}`);
            }
            const reader = res.body.getReader();
            const decoder = new TextDecoder('utf-8');
            let buffer = '';
            const consume = chunk => {
                const line = chunk.trim();
                if (line.startsWith('data: ')) handleStreamStep(JSON.parse(line.slice(6)));
            };
            while (true) {
                const {done, value} = await reader.read();
                if (done) break;
                buffer += decoder.decode(value, {stream:true});
                const chunks = buffer.split('\n\n');
                buffer = chunks.pop();
                chunks.forEach(consume);
            }
            buffer += decoder.decode();
            if (buffer.trim()) consume(buffer);
            if (!commandTerminal) throw new Error('Luồng tiến độ bị ngắt. Tác vụ vẫn chạy nền; tải lại dashboard để kết nối lại.');
        }

        // Execute or reconnect to an existing job; reconnect never reruns exports.
        async function executeCommand(resumeJobId = null) {
            if (commandBusy || nikeBusy) return;
            const resume = typeof resumeJobId === 'string' && resumeJobId.length > 0;
            const input = document.getElementById('ai-prompt-input');
            const prompt = input.value.trim();
            if (!prompt) {
                showToast('Vui lòng nhập lệnh trước khi bấm thực thi', false);
                return;
            }

            commandBusy = true;
            commandTerminal = false;
            localStorage.setItem('sfc-last-prompt', prompt);
            if (!resume) localStorage.removeItem('sfc-last-job');
            document.getElementById('command-files').replaceChildren();

            const btn = document.getElementById('btn-prompt-exec');
            const originalHtml = btn.innerHTML;
            btn.innerHTML = '<span class="spinner"></span> Đang chạy...';
            btn.disabled = true;

            // Show and reset Exec Console
            const consoleEl = document.getElementById('ai-exec-console');
            consoleEl.style.display = 'block';
            document.getElementById('exec-prompt-echo').innerText = `"${prompt}"`;
            const badge = document.getElementById('exec-status-badge');
            badge.className = 'live-pulse-badge';
            badge.innerHTML = '<span class="pulse-dot-red"></span> ĐANG XỬ LÝ TRỰC TIẾP';
            document.getElementById('intent-chips-bar').innerHTML = '<span class="intent-pill"><i class="fa-solid fa-spinner fa-spin"></i> Đang phân tích câu lệnh...</span>';

            // Reset all 5 steps to waiting
            for (let s = 1; s <= 5; s++) {
                const card = document.getElementById(`step-card-${s}`);
                const msg = document.getElementById(`step-msg-${s}`);
                const statusBadge = document.getElementById(`step-status-${s}`);
                if (card) {
                    card.className = s === 1 ? 'step-card running' : 'step-card waiting';
                }
                if (msg) msg.innerText = s === 1 ? '🧠 Đang bóc tách ý định câu lệnh...' : 'Đang chờ...';
                if (statusBadge) statusBadge.innerHTML = s === 1 ? '<i class="fa-solid fa-spinner fa-spin" style="color:#60a5fa;"></i>' : '<i class="fa-solid fa-clock" style="color:#64748b;"></i>';
            }

            addLog(`Đang gửi lệnh: "${prompt}"...`, 'info');

            try {
                const res = resume ? await fetch(`/api/command/jobs/${encodeURIComponent(resumeJobId)}/stream`) : await fetch('/api/command/stream', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ prompt })
                });

                await consumeCommandStream(res);

            } catch (e) {
                showToast(`Lỗi kết nối máy chủ: ${e.message}`, false);
                addLog(`Lỗi xử lý: ${e.message}`, 'warn');
                if (badge) {
                    badge.className = 'live-pulse-badge';
                    badge.innerHTML = '<i class="fa-solid fa-circle-exclamation"></i> GẶP SỰ CỐ';
                }
            } finally {
                commandBusy = false;
                btn.innerHTML = originalHtml;
                btn.disabled = false;
            }
        }

        function handleStreamStep(data) {
            if (data.job_id) localStorage.setItem('sfc-last-job', data.job_id);
            const stepNum = data.step;
            const card = document.getElementById(`step-card-${stepNum}`);
            const msg = document.getElementById(`step-msg-${stepNum}`);
            const statusBadge = document.getElementById(`step-status-${stepNum}`);

            if (card) {
                card.className = `step-card ${data.status}`;
            }
            if (msg && data.msg) {
                msg.innerText = data.msg;
            }

            // Status icon update
            if (statusBadge) {
                if (data.status === 'running') {
                    statusBadge.innerHTML = '<i class="fa-solid fa-spinner fa-spin" style="color:#60a5fa;"></i>';
                } else if (data.status === 'success') {
                    statusBadge.innerHTML = '<i class="fa-solid fa-circle-check" style="color:#34d399;"></i>';
                } else if (data.status === 'skipped') {
                    statusBadge.innerHTML = '<i class="fa-solid fa-forward-step" style="color:#94a3b8;"></i>';
                } else if (data.status === 'error') {
                    statusBadge.innerHTML = '<i class="fa-solid fa-circle-exclamation" style="color:#f87171;"></i>';
                } else if (data.status === 'warn') {
                    statusBadge.innerHTML = '<i class="fa-solid fa-triangle-exclamation" style="color:#fbbf24;"></i>';
                } else if (data.status === 'done') {
                    statusBadge.innerHTML = '<i class="fa-solid fa-trophy" style="color:#34d399;"></i>';
                }
            }

            // Step 1: Render Intent Pills
            if (stepNum === 1 && data.intent) {
                const intent = data.intent;
                const bar = document.getElementById('intent-chips-bar');
                bar.innerHTML = '';

                if (intent.target_plant) {
                    bar.innerHTML += `<span class="intent-pill plant"><i class="fa-solid fa-house-chimney"></i> Nhà máy: <strong>${(intent.target_plants || [intent.target_plant]).join(', ')}</strong></span>`;
                }
                if (intent.target_page) {
                    bar.innerHTML += `<span class="intent-pill page"><i class="fa-solid fa-compass"></i> Trang: <strong>${intent.target_page.toUpperCase()}</strong></span>`;
                }
                if (intent.target_line) {
                    bar.innerHTML += `<span class="intent-pill line"><i class="fa-solid fa-network-wired"></i> Chuyền: <strong>${intent.target_line}</strong></span>`;
                }
                if (intent.date_from) {
                    const dateStr = intent.date_to && intent.date_to !== intent.date_from ? `${intent.date_from} ➔ ${intent.date_to}` : intent.date_from;
                    bar.innerHTML += `<span class="intent-pill date"><i class="fa-solid fa-calendar-day"></i> Ngày: <strong>${dateStr}</strong></span>`;
                }
                if (intent.action) {
                    const actLabel = intent.action === 'excel' ? 'Xuất Excel' : 'Tìm kiếm dữ liệu';
                    bar.innerHTML += `<span class="intent-pill"><i class="fa-solid fa-bolt"></i> Thao tác: <strong>${actLabel}</strong></span>`;
                }
                if (typeof intent.nike === 'boolean') {
                    bar.innerHTML += `<span class="intent-pill">${intent.nike ? 'NIKE (External)' : 'Nội bộ'}</span>`;
                }
                if (intent.monthly) bar.innerHTML += '<span class="intent-pill">Chia từng tháng</span>';
                if (intent.combine) bar.innerHTML += '<span class="intent-pill">Gộp riêng từng nhà máy</span>';
                if (!bar.innerHTML) {
                    bar.innerHTML = '<span class="intent-pill"><i class="fa-solid fa-check"></i> Đã nhận diện yêu cầu</span>';
                }
            }

            // Step 2: Plant Switch Sync
            if (data.plant) {
                document.querySelectorAll('.plant-btn').forEach(btn => {
                    btn.classList.toggle('active', btn.id === `plant-btn-${data.plant}`);
                });
            }

            // Realtime Screenshot Mirror Update
            if (data.screenshot) {
                const mirrorImg = document.getElementById('mirror-img');
                if (mirrorImg) {
                    mirrorImg.src = data.screenshot;
                }
            }

            // Append to Activity Log
            if (data.msg) {
                const logType = (data.status === 'success' || data.status === 'done') ? 'success' : (data.status === 'warn' ? 'warn' : 'info');
                addLog(data.msg, logType);
            }

            if (data.files) {
                const files = document.getElementById('command-files');
                for (const file of data.files) {
                    if (!file.url || !file.url.startsWith('/api/exports/')) continue;
                    if ([...files.querySelectorAll('a')].some(a => a.getAttribute('href') === file.url)) continue;
                    const link = document.createElement('a');
                    link.href = file.url;
                    link.className = 'intent-pill';
                    link.textContent = `↓ ${file.name}${typeof file.rows === 'number' ? ` (${file.rows.toLocaleString()} dòng)` : ''}`;
                    link.setAttribute('download', file.name);
                    files.appendChild(link);
                }
            }
            if (data.status === 'error') {
                commandTerminal = true;
                const badge = document.getElementById('exec-status-badge');
                badge.className = 'live-pulse-badge';
                badge.textContent = 'CHƯA HOÀN TẤT';
                showToast(data.msg || 'Tác vụ thất bại', false);
            }

            // Step 5: Finished
            if (stepNum === 5 && data.status === 'done') {
                commandTerminal = true;
                const badge = document.getElementById('exec-status-badge');
                if (badge) {
                    badge.className = 'live-pulse-badge done';
                    badge.innerHTML = '<i class="fa-solid fa-circle-check"></i> HOÀN THÀNH TOÀN BỘ';
                }
                showToast('🎉 Đã thực hiện xong lệnh của bạn!');
                setTimeout(refreshStatus, 1500);
            }
        }

        function closeExecConsole() {
            const consoleEl = document.getElementById('ai-exec-console');
            if (consoleEl) consoleEl.style.display = 'none';
        }

        document.getElementById('btn-prompt-exec').addEventListener('click', executeCommand);
        document.getElementById('ai-prompt-input').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') executeCommand();
        });

        // Navigate Quick
        async function navTo(dest) {
            addLog(`Đang chuyển hướng sang: ${dest.toUpperCase()}...`, 'info');
            try {
                const res = await fetch('/api/navigate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ destination: dest })
                });
                const data = await res.json();
                if (data.success) {
                    showToast(`Đã chuyển đến: ${data.title}`);
                    addLog(`Đã chuyển đến trang: ${data.title}`, 'success');
                    if (data.screenshot) {
                        document.getElementById('mirror-img').src = data.screenshot;
                    }
                    setTimeout(refreshStatus, 1000);
                } else {
                    showToast(data.error || 'Lỗi điều hướng', false);
                }
            } catch(e) {
                showToast('Lỗi điều hướng', false);
            }
        }

        // Trigger Filter Action (Search / Excel)
        async function triggerAction(action) {
            if (nikeBusy) return;
            const nike = document.getElementById('nike-toggle').checked;
            const dateFrom = document.getElementById('date-from').value;
            const dateTo = document.getElementById('date-to').value;
            const factory = document.getElementById('factory-select').value;

            const btn = action === 'search' ? document.getElementById('btn-run-search') : document.getElementById('btn-run-excel');
            const orig = btn.innerHTML;
            btn.innerHTML = '<span class="spinner"></span> Đang gửi lệnh...';
            btn.disabled = true;

            addLog(`Đang thực hiện ${action === 'search' ? 'Tìm kiếm' : 'Xuất Excel'} Chuyền [${factory || 'Tất cả'} | ${dateFrom} -> ${dateTo}]...`, 'info');

            try {
                const res = await fetch('/api/action', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ dateFrom, dateTo, factory, action, nike })
                });
                const data = await res.json();
                if (data.success) {
                    showToast(data.action_result || 'Thực hiện thành công!');
                    addLog(data.action_result || 'Đã hoàn tất thao tác', 'success');
                    if (data.screenshot) {
                        document.getElementById('mirror-img').src = data.screenshot;
                    }
                    setTimeout(refreshStatus, 2000);
                } else {
                    showToast(data.error || 'Lỗi thực thi', false);
                }
            } catch(e) {
                showToast('Lỗi kết nối máy chủ', false);
            } finally {
                btn.innerHTML = orig;
                btn.disabled = false;
            }
        }

        // MQAA Actions
        async function startMqaaScan() {
            addLog("Đang kích hoạt quy trình quét 21 Form MQAA...", 'info');
            try {
                const res = await fetch('/api/mqaa/scan', { method: 'POST' });
                const data = await res.json();
                showToast(data.message);
                addLog(data.message, 'success');
            } catch(e) {
                showToast("Lỗi khi bắt đầu quét MQAA", false);
            }
        }

        async function openMasterExcel() {
            try {
                const res = await fetch('/api/mqaa/open_excel', { method: 'POST' });
                const data = await res.json();
                if (data.success) {
                    showToast(`Đã mở file: ${data.file}`);
                    addLog(`Đã mở file: ${data.file}`, 'success');
                } else {
                    showToast(data.error, false);
                }
            } catch(e) {
                showToast("Lỗi mở file", false);
            }
        }

        // Initial Boot
        window.addEventListener('DOMContentLoaded', () => {
            const savedPrompt = localStorage.getItem('sfc-last-prompt');
            const savedJob = localStorage.getItem('sfc-last-job');
            if (savedPrompt) document.getElementById('ai-prompt-input').value = savedPrompt;
            if (savedJob && savedPrompt) executeCommand(savedJob);
            refreshStatus();
            loadFactories();
            loadMemory();
            loadMqaaSummary();
            setInterval(refreshStatus, 4000);
        });
    