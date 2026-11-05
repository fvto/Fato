// QMS catalog coverage, search behavior and original destination preservation.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const root = path.join(__dirname, '..');
const memory = JSON.parse(fs.readFileSync(path.join(root, 'sfc_memory_knowledge_base.json'), 'utf8'));
const context = vm.createContext({window: {location: {href: 'http://localhost:5050/'}}, URL});
vm.runInContext(fs.readFileSync(path.join(root, 'static/qms-navigation.js'), 'utf8'), context);
const nav = context.window.QmsNavigation;
const original = JSON.stringify(memory.qms_modules);
const modules = nav.prepare(memory.qms_modules);
assert.equal(modules.length, 51);
assert.equal(JSON.stringify(memory.qms_modules), original, 'Display translation must not mutate crawled memory');
assert.deepEqual(Array.from(modules, m => m.href), memory.qms_modules.map(m => m.href));
assert.ok(modules.every(m => /^[\x20-\x7e]+$/.test(m.name)), 'All 51 displayed names must be English');
const counts = {};
for (const module of modules) counts[module.category] = (counts[module.category] || 0) + 1;
assert.deepEqual(counts, {'Control Plan': 3, 'Basic Information': 3, 'QA Document': 4,
    'Collect Data': 9, 'Quality Database': 11, 'Performance Management': 4, 'Other Modules': 17});
assert.equal(nav.filter(modules, '   ').length, 51);
assert.equal(nav.filter(modules, 'q504')[0].name, 'HFPA');
assert.equal(nav.filter(modules, 'Q101')[0].category, 'Control Plan');
assert.equal(nav.filter(modules, 'q05')[0].name, 'QA Manual');
assert.equal(nav.filter(modules, 'cOpQ').length, 2, 'Search must include both data entry and reporting');
assert.equal(nav.filter(modules, 'performance management').length, 4);
assert.equal(nav.filter(modules, 'calibration report')[0].name, 'Product Process Calibration Report');
assert.equal(nav.filter(modules, 'COPQ', 'Collect Data')[0].name, 'Cost Of Poor Quality (CoPQ)');
assert.equal(nav.filter(modules, 'COPQ', 'QA Document').length, 0);
assert.equal(nav.filter(modules, '<script>').length, 0);
assert.equal(nav.filter(modules, 'zz-no-such-module').length, 0);
assert.equal(nav.filter(modules, 'Re-Inspection')[0].href, 'https://sfc.chingluh.com/vhmes/QA/QCRECHK.aspx');
function element() {
    return {children: [], value: '', appendChild(child) {this.children.push(child);},
        replaceChildren() {this.children = [];}, addEventListener(type, callback) {this[type] = callback;}};
}
const elements = new Map(['qms-search', 'qms-category', 'qms-module-groups', 'qms-search-count',
    'qms-search-clear', 'qms-empty'].map(id => [id, element()]));
context.document = {getElementById: id => elements.get(id), createElement: element};
context.window.cachedQmsModules = modules;
elements.get('qms-search').value = 'q504';
nav.render();
assert.equal(elements.get('qms-search-count').textContent, '1 of 51 modules');
const link = elements.get('qms-module-groups').children[0].children[0].children[1].children[0].children[0];
let opened;
let prevented = false;
context.window.jumpToQmsModule = (href, name) => {opened = {href, name};};
link.click({preventDefault() {prevented = true;}});
assert.ok(prevented, 'Ordinary clicks must navigate through the dashboard connection');
assert.deepEqual(opened, {href: 'https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504', name: 'HFPA'});
elements.get('qms-search').value = 'zz-no-such-module';
nav.render();
assert.equal(elements.get('qms-empty').hidden, false);
assert.equal(elements.get('qms-module-groups').children.length, 0);
console.log('QMS: 51 English labels, categories, URL preservation, combined search, empty state and filtered-result navigation passed');
