/* Interactive Preview, 3-Tier Parallel Matrix & Drill-Down for Quarterly NIKE Defect Review */
(function () {
  'use strict';

  const cache = new Map();
  const uiState = new Map();

  // Keep the displayed plan aligned with the actual worker events, including old runs.
  function progressSteps(run) {
    const plan = [
      ['validation', 'Kiểm tra nguồn NIKE độc lập'],
      ['hfpa_parse', 'Chuẩn hóa dữ liệu HFPA NIKE'],
      ['ftt_parse', 'Chuẩn hóa dữ liệu FTT NIKE'],
      ['area_mapping', 'Phân khu VH / JV / VB từ Factory Sandals'],
      ['analysis', 'Phân tích 3 tầng lỗi và Top model'],
      ['excel', 'Tạo Excel và biểu đồ nguồn'],
      ...(run.mode === 'excel' ? [] : [['presentation', 'Xuất PowerPoint 7 slide và kiểm tra mở lại']]),
      ['export', 'Đối chiếu và đóng gói kết quả'],
    ];
    const recorded = new Map((run.steps || []).map(step => [step.id, step]));
    if (recorded.has('acquisition')) plan.unshift(['acquisition', 'Thu thập dữ liệu SFC']);
    const steps = plan.map(([id, title]) => ({status: 'pending', ...recorded.get(id), id, title}));
    // Preserve unexpected executed events; discard only the obsolete unused placeholder.
    for (const step of run.steps || []) {
      if (step.id === 'processing' && step.status === 'pending') continue;
      if (run.mode === 'excel' && step.id === 'presentation' && step.status === 'pending') continue;
      if (!plan.some(([id]) => id === step.id)) steps.push({...step});
    }
    return steps;
  }

  function esc(v) {
    return String(v ?? '').replace(/[&<>"']/g, (c) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
    }[c]));
  }

  function fmtNum(n) {
    if (n === null || n === undefined || Number.isNaN(Number(n))) return '-';
    return Math.round(Number(n)).toLocaleString('en-US');
  }

  function fmtPct(p) {
    if (p === null || p === undefined || Number.isNaN(Number(p))) return 'N/A';
    return Number(p).toFixed(1) + '%';
  }

  function computeQuarterDates(period) {
    if (!/^FY\d{2}-Q[1-4]$/.test(period || '')) return ['2026-06-01', '2026-08-31'];
    const fyEnd = 2000 + Number(period.slice(2, 4));
    const q = Number(period.slice(-1));
    if (q === 1) return [`${fyEnd - 1}-06-01`, `${fyEnd - 1}-08-31`];
    if (q === 2) return [`${fyEnd - 1}-09-01`, `${fyEnd - 1}-11-30`];
    if (q === 3) {
      const febLast = new Date(Date.UTC(fyEnd, 2, 0)).getUTCDate();
      return [`${fyEnd - 1}-12-01`, `${fyEnd}-02-${String(febLast).padStart(2, '0')}`];
    }
    return [`${fyEnd}-03-01`, `${fyEnd}-05-31`];
  }

  function renderScopeCards(metricData) {
    const isHfpa = metricData.metric === 'HFPA';
    const order = ['CLG', 'VH', 'JV', 'VB'];
    const titles = {
      CLG: 'CLG (VH + JV + VB)',
      VH: 'VH (VH+VH2 trừ VB)',
      JV: 'JV (JV+JV2)',
      VB: 'VB (Factory Sandals)',
    };
    return `<div class="qdr-cards">${order.map((s) => {
      const st = metricData.scope_totals[s] || {};
      const missing = st.scope_status === 'missing_data';
      const vol = isHfpa ? `Sample: ${fmtNum(st.inspected_or_produced_qty)} · DefPair: ${fmtNum(st.defective_pairs_qty)}` : `Produced(A+B+C): ${fmtNum(st.inspected_or_produced_qty)} · Rework: ${fmtNum(st.rework_pairs_qty)}`;
      return `<div class="qdr-card ${missing ? 'is-warn' : ''}">
        <div class="qdr-card-head"><strong>${esc(titles[s])}</strong><span class="qdr-pill">${missing ? 'Thiếu dữ liệu' : st.scope_status === 'no_defect' ? 'Không có lỗi' : 'Đủ nguồn'}</span></div>
        <div class="qdr-big">${missing ? '-' : fmtNum(st.total_defect_issues)} <small>issues</small></div>
        <div class="qdr-sub">${missing ? 'Chưa có dữ liệu nguồn khu vực này' : esc(vol)}</div>
      </div>`;
    }).join('')}</div>`;
  }

  function renderTopGroups(metricData) {
    const groups = metricData.top_groups || [];
    if (!groups.length) return '<p class="note">Chưa phát sinh nhóm lỗi dương trong kỳ đã chọn.</p>';
    return `<div class="export-table"><table><thead><tr>
      <th>Hạng</th><th>Nhóm chuẩn HFPA (Top 3 CLG)</th><th>VH</th><th>JV</th><th>VB</th><th>CLG</th><th>Đóng góp VH/JV/VB</th>
    </tr></thead><tbody>${groups.map((g) => `<tr>
      <td>#${g.rank}</td>
      <td><strong>${esc(g.group_label)}</strong></td>
      <td>${fmtNum(g.vh_qty)}</td>
      <td>${fmtNum(g.jv_qty)}</td>
      <td>${fmtNum(g.vb_qty)}</td>
      <td><strong>${fmtNum(g.clg_qty)}</strong></td>
      <td>${fmtPct(g.vh_contrib_pct)} / ${fmtPct(g.jv_contrib_pct)} / ${fmtPct(g.vb_contrib_pct)}</td>
    </tr>`).join('')}</tbody></table></div>`;
  }

  function formatDetailCell(cell) {
    if (!cell || cell.cell_status === 'missing_scope_data') return `<span class="note">Thiếu dữ liệu nguồn</span>`;
    if (cell.cell_status === 'no_defect') return `<span class="note">Không có lỗi (0)</span>`;
    const items = cell.top_details || [];
    if (!items.length) return `<span class="qdr-warn-text">${esc(cell.status_note || 'Thiếu dữ liệu chi tiết')}</span>`;
    const lines = items.map((d) => `${d.rank}. ${esc(d.detail_label)}: <strong>${fmtNum(d.qty)}</strong> (${fmtPct(d.share_of_group_in_scope_pct)})`);
    if (cell.gap_qty > 0.5) {
      lines.push(`<span class="note">Chưa có chi tiết: ${fmtNum(cell.gap_qty)}</span>`);
    }
    return lines.join('<br>');
  }

  function formatModelCell(mcell) {
    if (!mcell || mcell.cell_status === 'missing_scope_data') return `<span class="note">Thiếu dữ liệu</span>`;
    if (mcell.cell_status === 'no_defect') return `<span class="note">Không có lỗi (0)</span>`;
    const items = mcell.top_models || [];
    if (!items.length) return `<span class="note">(Không có model lỗi)</span>`;
    return items.map((m) => `${m.rank}. ${esc(m.model_label)}: <strong>${fmtNum(m.qty)}</strong> (${fmtPct(m.share_of_group_in_scope_pct)})`).join('<br>');
  }

  function renderMatrix(metricData) {
    const groups = metricData.top_groups || [];
    if (!groups.length) return '';
    const scopes = ['CLG', 'VH', 'JV', 'VB'];
    const detRows = groups.map((g) => `<tr>
      <td><strong>#${g.rank} ${esc(g.group_label)}</strong><br><span class="note">CLG: ${fmtNum(g.clg_qty)}</span></td>
      ${scopes.map((s) => `<td>${formatDetailCell(metricData.details_by_group?.[g.group_id]?.[s])}</td>`).join('')}
    </tr>`).join('');
    const modRows = groups.map((g) => `<tr>
      <td><strong>#${g.rank} ${esc(g.group_label)}</strong></td>
      ${scopes.map((s) => `<td>${formatModelCell(metricData.models_by_group?.[g.group_id]?.[s])}</td>`).join('')}
    </tr>`).join('');
    return `
      <h4>Tầng 2 — Top ≤3 Lỗi Chi Tiết Trong Từng Nhóm × Khu Vực (% trên tổng nhóm tại khu đó)</h4>
      <div class="export-table"><table><thead><tr><th>Nhóm cha</th><th>CLG</th><th>VH</th><th>JV</th><th>VB</th></tr></thead><tbody>${detRows}</tbody></table></div>
      <h4>Chiều Song Song — Top 3 Model Trong Từng Nhóm × Khu Vực (không thay thế Top lỗi chi tiết)</h4>
      <div class="export-table"><table><thead><tr><th>Nhóm cha</th><th>CLG</th><th>VH</th><th>JV</th><th>VB</th></tr></thead><tbody>${modRows}</tbody></table></div>
    `;
  }

  function renderDrilldown(metricData, stateObj) {
    const groups = metricData.top_groups || [];
    if (!groups.length) return '';
    const selGroup = stateObj.groupId && groups.some((g) => g.group_id === stateObj.groupId) ? stateObj.groupId : groups[0].group_id;
    const selScope = stateObj.scope || 'CLG';
    const dd = metricData.drilldown_by_group?.[selGroup];
    if (!dd) return '';

    const detailRows = (dd.details || []).map((d) => {
      const qtyScope = selScope === 'CLG' ? d.clg_qty : selScope === 'VH' ? d.vh_qty : selScope === 'JV' ? d.jv_qty : d.vb_qty;
      const gScope = selScope === 'CLG' ? dd.group_total_clg : (metricData.details_by_group?.[selGroup]?.[selScope]?.group_total || 0);
      const sharePct = gScope > 0 ? (qtyScope * 100 / gScope) : null;
      const areaContrib = d.clg_qty > 0 && selScope !== 'CLG' ? (qtyScope * 100 / d.clg_qty) : (selScope === 'CLG' && d.clg_qty > 0 ? 100.0 : null);
      const topMods = (d.top_models_for_detail || []).slice(0, 3).map(([nm, q]) => `${nm} (${fmtNum(q)})`).join(', ');
      const metaBits = [
        d.lines_count ? `${fmtNum(d.lines_count)} chuyền` : '',
        topMods ? `Model: ${topMods}` : '',
      ].filter(Boolean).join(' · ');
      const provList = (d.provenance_samples || d.provenance_rows || []).slice(0, 2);
      return `<tr>
        <td>#${d.rank_clg}</td>
        <td>${esc(d.detail_label)}${metaBits ? `<br><span class="note">${esc(metaBits)}</span>` : ''}</td>
        <td>${esc(d.nike_subtype_label || '-')}</td>
        <td><strong>${fmtNum(qtyScope)}</strong></td>
        <td>${fmtPct(sharePct)}</td>
        <td>${selScope === 'CLG' ? `VH ${fmtPct(d.vh_contrib_to_detail_pct)} · JV ${fmtPct(d.jv_contrib_to_detail_pct)} · VB ${fmtPct(d.vb_contrib_to_detail_pct)}` : fmtPct(areaContrib)}</td>
        <td class="note">${esc(provList.join('; ') || '-')}</td>
      </tr>`;
    }).join('');

    const rawModels = selScope === 'CLG'
      ? ((dd.models_clg && dd.models_clg.length) ? dd.models_clg : (metricData.models_by_group?.[selGroup]?.CLG?.top_models || []))
      : (metricData.models_by_group?.[selGroup]?.[selScope]?.top_models || []);
    const modelRows = rawModels.map((m, i) => {
      const shareVal = m.share_of_clg_group_pct ?? m.share_of_group_in_scope_pct ?? m.share_of_group_pct;
      const lineOrVol = (m.lines && m.lines.length)
        ? m.lines.slice(0, 4).join(', ')
        : (m.ref_inspected_or_produced_qty ? `SL kiểm/SX: ${fmtNum(m.ref_inspected_or_produced_qty)}` : '-');
      return `<tr>
        <td>#${m.rank || (i + 1)}</td>
        <td><strong>${esc(m.model_label)}</strong></td>
        <td>${fmtNum(m.qty)}</td>
        <td>${fmtPct(shareVal)}</td>
        <td class="note">${esc(lineOrVol)}</td>
      </tr>`;
    }).join('');

    return `<div class="qdr-drill">
      <div class="qdr-drill-bar">
        <strong>Drill-Down Từng Ô (Tách Biệt 2 Mẫu Số)</strong>
        <label>Nhóm lỗi <select data-qdr-group>${groups.map((g) => `<option value="${esc(g.group_id)}" ${g.group_id === selGroup ? 'selected' : ''}>#${g.rank} ${esc(g.group_label)}</option>`).join('')}</select></label>
        <label>Phạm vi <select data-qdr-scope>${['CLG', 'VH', 'JV', 'VB'].map((s) => `<option value="${s}" ${s === selScope ? 'selected' : ''}>${s}</option>`).join('')}</select></label>
      </div>
      <div class="export-table"><table><thead><tr>
        <th>Hạng</th><th>Lỗi chi tiết</th><th>Loại FTT</th><th>Số lỗi (${esc(selScope)})</th><th>Tỷ trọng trong nhóm tại ${esc(selScope)} (%)</th><th>Đóng góp khu vào cùng lỗi CLG (%)</th><th>Nguồn đối chiếu</th>
      </tr></thead><tbody>${detailRows || `<tr><td colspan="7" class="note">Không có dòng lỗi chi tiết ở mức hạt ${esc(metricData.detail_support_level)}.</td></tr>`}</tbody></table></div>
      <div class="export-table"><table><thead><tr>
        <th>Hạng</th><th>Model (${esc(selScope)})</th><th>Số lỗi trong nhóm</th><th>Tỷ trọng trong nhóm (%)</th><th>Chuyền ghi nhận</th>
      </tr></thead><tbody>${modelRows || '<tr><td colspan="5" class="note">Không có model phát sinh lỗi trong phạm vi này.</td></tr>'}</tbody></table></div>
    </div>`;
  }

  function renderReviewAndOverrides(summary, metricData, runId) {
    const items = metricData.review?.items || [];
    const unmapped = items.filter((it) => it.review_type === 'unmapped_defect');
    const groupOpts = (summary.hfpa?.all_groups || []).map((g) => `<option value="${esc(g.group_id)}">${esc(g.group_label)}</option>`).join('');
    const rows = items.slice(0, 25).map((it) => `<tr>
      <td><span class="qdr-pill">${esc(it.review_type)}</span></td>
      <td>${esc(it.factory)} ${esc(it.plant)} ${esc(it.line_std)}</td>
      <td><strong>${esc(it.raw_label)}</strong></td>
      <td>${fmtNum(it.affected_quantity)}</td>
      <td>${esc(it.reason)}</td>
      <td>${it.review_type === 'unmapped_defect' ? `<div class="qdr-map-row">
        <select data-map-group data-raw-key="${esc(it.raw_key)}" data-raw-label="${esc(it.raw_label)}" data-metric="${esc(it.metric)}">${groupOpts}</select>
        <button type="button" data-save-map>Lưu quy tắc</button>
      </div>` : esc(it.suggested_or_confirmed_mapping)}</td>
    </tr>`).join('');
    return `<details class="qdr-review-box" ${unmapped.length ? 'open' : ''}>
      <summary>Kiểm tra mapping & cảnh báo ${esc(metricData.metric)} (${items.length} mục)</summary>
      ${rows ? `<div class="export-table"><table><thead><tr><th>Loại</th><th>Khu/Chuyền</th><th>Nhãn gốc</th><th>Số lỗi ảnh hưởng</th><th>Lý do</th><th>Hướng xử lý / Chốt mapping</th></tr></thead><tbody>${rows}</tbody></table></div>` : '<p class="note">Không có nhãn chưa ánh xạ hoặc xung đột.</p>'}
      <p class="note" data-qdr-map-status="${esc(runId)}"></p>
    </details>`;
  }

  function renderSlides(summary) {
    const previews = summary.preview_urls || [];
    if (!previews.length) return '<p class="note">Chưa có ảnh xem trước slide (chế độ chỉ xuất Excel hoặc chưa xuất PPTX).</p>';
    return `<div class="qdr-slides">${previews.map((p) => `<figure class="qdr-slide-card">
      <figcaption>${esc(p.name)}</figcaption>
      <img src="${esc(p.url)}" alt="${esc(p.name)}" loading="lazy">
    </figure>`).join('')}</div>`;
  }

  function draw(container, run, summary) {
    const st = uiState.get(run.id) || { tab: 'HFPA', groupId: '', scope: 'CLG' };
    uiState.set(run.id, st);
    const activeData = st.tab === 'FTT' ? summary.ftt : summary.hfpa;
    const catCount = summary.catalog?.count || 0;

    container.innerHTML = `<div class="qdr-panel">
      <div class="qdr-banner">
        <div><strong>Kỳ: ${esc(summary.period)} (${esc(summary.date_from)} → ${esc(summary.date_to)})</strong> · Chỉ nguồn <strong>NIKE</strong> (HFPA & FTT độc lập) · Danh mục Factory Sandals (VB): <strong>${catCount} line</strong></div>
        <div class="qdr-dr-note"><strong>Trạng thái DR:</strong> ${esc(summary.dr_config?.status?.toUpperCase())} — ${esc(summary.dr_config?.note)}</div>
      </div>
      <div class="qdr-tabs" role="tablist">
        <button type="button" class="${st.tab === 'HFPA' ? 'is-active' : ''}" data-qdr-tab="HFPA">HFPA NIKE (${fmtNum(summary.hfpa?.scope_totals?.CLG?.total_defect_issues)} lỗi)</button>
        <button type="button" class="${st.tab === 'FTT' ? 'is-active' : ''}" data-qdr-tab="FTT">FTT NIKE (${fmtNum(summary.ftt?.scope_totals?.CLG?.total_defect_issues)} lỗi)</button>
        <button type="button" class="${st.tab === 'SLIDES' ? 'is-active' : ''}" data-qdr-tab="SLIDES">Xem trước 7 Slide PPTX (${(summary.preview_urls || []).length})</button>
      </div>
      ${st.tab === 'SLIDES' ? renderSlides(summary) : `
        <div class="qdr-metric-body">
          <p class="note">Mức chi tiết nguồn <strong>${esc(activeData.metric)}</strong>: <code>${esc(activeData.detail_support_level)}</code> · Trạng thái: <strong>${esc(activeData.review?.status_label)}</strong></p>
          ${renderScopeCards(activeData)}
          ${renderTopGroups(activeData)}
          ${renderMatrix(activeData)}
          ${renderDrilldown(activeData, st)}
          ${renderReviewAndOverrides(summary, activeData, run.id)}
        </div>
      `}
    </div>`;

    container.querySelectorAll('[data-qdr-tab]').forEach((btn) => {
      btn.onclick = () => {
        st.tab = btn.dataset.qdrTab;
        st.groupId = '';
        draw(container, run, summary);
      };
    });
    const gSel = container.querySelector('[data-qdr-group]');
    if (gSel) gSel.onchange = () => { st.groupId = gSel.value; draw(container, run, summary); };
    const sSel = container.querySelector('[data-qdr-scope]');
    if (sSel) sSel.onchange = () => { st.scope = sSel.value; draw(container, run, summary); };

    container.querySelectorAll('[data-save-map]').forEach((btn) => {
      btn.onclick = async () => {
        const row = btn.closest('.qdr-map-row');
        const sel = row.querySelector('[data-map-group]');
        btn.disabled = true;
        try {
          const resp = await fetch('/api/report-tools/quarterly_defect_review/overrides', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              rules: [{
                source_metric: sel.dataset.metric,
                raw_key: sel.dataset.rawKey,
                raw_label: sel.dataset.rawLabel,
                standard_group_id: sel.value,
                confirmed_by: 'user_ui',
              }],
            }),
          });
          const data = await resp.json();
          const msgEl = container.querySelector(`[data-qdr-map-status="${run.id}"]`);
          if (msgEl) msgEl.textContent = resp.ok ? `Đã lưu quy tắc mapping (${data.mapping_version}). Bấm "Chạy báo cáo" để cập nhật kết quả mới.` : (data.error || 'Chưa lưu được mapping.');
        } finally {
          btn.disabled = false;
        }
      };
    });
  }

  async function mount(runSection, run) {
    if (!runSection || !run || run.tool !== 'quarterly_defect_review' || run.status !== 'done') return;
    let host = runSection.querySelector('.qdr-host');
    if (!host) {
      host = document.createElement('div');
      host.className = 'qdr-host';
      const details = runSection.querySelector('details');
      runSection.insertBefore(host, details || null);
    }
    const cacheKey = `${run.id}:${run.finished_at || ''}`;
    if (cache.has(cacheKey)) {
      draw(host, run, cache.get(cacheKey));
      return;
    }
    host.innerHTML = '<p class="note">Đang tải bảng phân tích HFPA NIKE & FTT NIKE...</p>';
    try {
      const resp = await fetch(`/api/report-tools/runs/${encodeURIComponent(run.id)}/defect-review`);
      const data = await resp.json();
      if (data.ready && data.review_summary) {
        cache.set(cacheKey, data.review_summary);
        draw(host, run, data.review_summary);
      }
    } catch (_err) {
      host.innerHTML = '<p class="error">Chưa tải được bản xem trước phân tích.</p>';
    }
  }

  window.QDRPanel = { mount, computeQuarterDates, progressSteps };
})();

