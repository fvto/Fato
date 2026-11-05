'use strict';
// English labels are keyed by destination so refreshing the crawled memory cannot
// bring translated labels back or change the original navigation URLs.
window.QmsNavigation = (() => {
    const groups = ['Control Plan', 'Basic Information', 'QA Document', 'Collect Data',
        'Quality Database', 'Performance Management', 'Other Modules'];
    const catalog = [
        ['id=Q101', 'Quality Risk', 0], ['id=Q102', 'PFMEA', 0], ['id=Q103', 'Order Information', 0],
        ['QAReason.aspx', 'Inspection Basic Data', 1], ['id=Q202', 'BDTS Rule', 1], ['id=Q203', 'Notification', 1],
        ['id=Q301', 'Inspection Strategy', 2], ['id=Q302', 'CTQ & CTP', 2],
        ['QAReasonBasic.aspx', 'QA Manual', 2], ['id=Q304', 'Play Book/Guideline', 2],
        ['AQLTracking.aspx', 'AQL', 3], ['QARecordCOPQ.aspx', 'Cost Of Poor Quality (CoPQ)', 3],
        ['BondTest.aspx', 'Bonding Test', 3], ['GoreTexTest.aspx', 'Gore-Tex', 3],
        ['LearningCase.aspx', 'CFM', 3], ['/Update/Reason.aspx', 'Abnormal Tracking', 3],
        ['QABCDSTRRecord.aspx', 'Destruction Management', 3], ['MesQAPPC.aspx', 'Product Process Calibration', 3],
        ['MQAACheck.aspx', 'MQAA Audit', 3],
        ['id=Q501', 'AQL Inspection', 4], ['QAAnalysis.aspx', 'Quality Analysis', 4],
        ['QCHFPA.aspx', 'HFPA', 4], ['PackingCheck.aspx', 'Packing', 4],
        ['BondTestAnalysis.aspx', 'Bond Test Analysis', 4], ['DefectiveReturn.aspx', 'Upper Quality Data', 4],
        ['QCFGWH.aspx', 'FGWH', 4], ['QCRECHK.aspx', 'Re-Inspection', 4], ['id=Q508', 'IMQC', 4],
        ['QCPPC.aspx', 'Product Process Calibration Report', 4], ['MQAAReport.aspx', 'MQAA Report', 4],
        ['QAReportTrackingCOPQ.aspx', 'COPQ Report', 5], ['QAReportNIKE.aspx', 'BDTS Report', 5],
        ['QADashboard.aspx', 'Dashboard', 5], ['/QA/QA.aspx', 'FTT', 5],
        ['QMSMain.aspx', 'QMS Home', 6], ['EDisplayQA.aspx', 'QA E-Display', 6],
        ['EDisplaySF.aspx', 'Stockfitting E-Display', 6], ['MesQACheck.aspx', 'QC CTB', 6],
        ['MesQACheckPackingOQC.aspx', 'QC HFPA (Packing OQC)', 6], ['MesQACheckPackingFGWH.aspx', 'QC FGWH', 6],
        ['MesQACheckPackingRECHK.aspx', 'HFPA Re-Inspection', 6], ['ArticleInspection.aspx', 'OMD (Article Inspection)', 6],
        ['MesQAIPCheck.aspx', 'QC IP', 6], ['MesQADPCheck.aspx', 'QC DP', 6], ['MesQAOSCheck.aspx', 'QC RB/OS', 6],
        ['MesQAPUCheck.aspx', 'QC PU', 6], ['MesQACMPCheck.aspx', 'QC CMP', 6],
        ['MesQATransferFrom.aspx', 'Upper QC Transfer', 6], ['MesQATransferBottom.aspx', 'Bottom QC Transfer', 6],
        ['BCGrade.aspx', 'B/C Grade QA', 6], ['PackingUCC.aspx', 'UCC Packing', 6]
    ];
    function prepare(modules) {
        return modules.map(module => {
            const url = new URL(module.href, window.location.href);
            const match = catalog.find(([key]) => key.startsWith('id=')
                ? url.searchParams.get('id') === key.slice(3)
                : url.pathname.toLowerCase().endsWith(key.toLowerCase()));
            const code = url.searchParams.get('id') || url.searchParams.get('tree_id') || '';
            return {...module, name: match ? match[1] : module.name,
                category: match ? groups[match[2]] : groups[6], code};
        });
    }
    function filter(modules, query, category = '') {
        const terms = query.toLowerCase().trim().split(/\s+/).filter(Boolean);
        return modules.filter(module => (!category || module.category === category) &&
            terms.every(term => `${module.name} ${module.category} ${module.code} ${module.href}`.toLowerCase().includes(term)));
    }
    function englishTile(tile) {
        const labels = {
            'Kiểm tra lại': 'Re-Inspection', 'Báo cáo MQAA': 'MQAA Report',
            'Mã Q504': 'Code Q504', '3 Dạng Xuất': '3 export options',
            '3 Tùy chọn Xuất: 1.Report COPQ, 2.COPQ Data, 3.Monthly': 'Report COPQ / COPQ Data / Monthly',
            'Kiểm toán 21 Mẫu': '21-form audit', 'Cây Quản trị QA': 'QA module tree',
            'Trang chủ FlowChart': 'FlowChart home'
        };
        return {...tile, title: labels[tile.title] || tile.title, sub: labels[tile.sub] || tile.sub};
    }
    function render() {
        const search = document.getElementById('qms-search');
        const category = document.getElementById('qms-category');
        const container = document.getElementById('qms-module-groups');
        if (!container) return;
        const modules = window.cachedQmsModules || [];
        const matches = filter(modules, search.value, category.value);
        container.replaceChildren();
        // Match the reference's three-column order, with related sections stacked.
        [[0, 1, 2], [3, 6], [4, 5]].forEach(indices => {
            const column = document.createElement('div');
            column.className = 'qms-menu-column';
            indices.forEach(index => {
                const items = matches.filter(module => module.category === groups[index]);
                if (!items.length) return;
                const section = document.createElement('section');
                section.className = 'qms-menu-group';
                const heading = document.createElement('h3');
                heading.textContent = groups[index];
                section.appendChild(heading);
                const list = document.createElement('ul');
                items.forEach(module => {
                    const item = document.createElement('li');
                    const link = document.createElement('a');
                    link.href = module.href;
                    link.textContent = module.name;
                    link.title = `Open ${module.name}${module.code ? ` (${module.code})` : ''}`;
                    link.addEventListener('click', event => {
                        if (event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
                        event.preventDefault();
                        window.jumpToQmsModule(module.href, module.name);
                    });
                    item.appendChild(link);
                    list.appendChild(item);
                });
                section.appendChild(list);
                column.appendChild(section);
            });
            if (column.children.length) container.appendChild(column);
        });
        document.getElementById('qms-search-count').textContent = `${matches.length} of ${modules.length} modules`;
        document.getElementById('qms-search-clear').hidden = !search.value;
        document.getElementById('qms-empty').hidden = matches.length > 0;
    }
    function initialize() {
        const select = document.getElementById('qms-category');
        select.replaceChildren(new Option('All categories', ''));
        groups.forEach(group => select.add(new Option(group, group)));
        document.getElementById('qms-search').disabled = false;
        select.disabled = false;
        render();
    }
    function populateSelect(select, modules) {
        select.replaceChildren(new Option('-- Choose a QMS module --', ''));
        groups.forEach(group => {
            const items = modules.filter(module => module.category === group);
            if (!items.length) return;
            const optgroup = document.createElement('optgroup');
            optgroup.label = group;
            items.forEach(module => {
                const option = new Option(`${module.name}${module.code ? ` (${module.code})` : ''}`, module.href);
                option.setAttribute('data-name', module.name);
                optgroup.appendChild(option);
            });
            select.appendChild(optgroup);
        });
    }
    return {prepare, filter, englishTile, render, initialize, populateSelect};
})();
