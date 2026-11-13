const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const root = path.join(__dirname, '..');
const context = vm.createContext({window: {}});
vm.runInContext(fs.readFileSync(path.join(root, 'static/quarterly_defect_review.js'), 'utf8'), context);
const progress = context.window.QDRPanel.progressSteps;
const registry = JSON.parse(fs.readFileSync(path.join(root, 'report_registry.json'), 'utf8'));
const expected = ['validation', 'hfpa_parse', 'ftt_parse', 'area_mapping', 'analysis', 'excel', 'presentation', 'export'];
assert.deepEqual(registry.quarterly_defect_review.steps.all.map(s => s.id), expected);
assert.deepEqual(registry.quarterly_defect_review.steps.excel.map(s => s.id), expected.filter(id => id !== 'presentation'));

// An old run appends parse events after its unused coarse processing placeholder.
const run = {mode: 'all', steps: [
  {id: 'acquisition', status: 'done', finished_at: '13:59'},
  {id: 'validation', status: 'done'},
  {id: 'processing', status: 'pending'},
  {id: 'presentation', status: 'pending'},
  {id: 'export', status: 'pending'},
  {id: 'hfpa_parse', status: 'done', finished_at: '14:03'},
  {id: 'ftt_parse', status: 'running', started_at: '14:03'},
]};
const before = JSON.stringify(run);
const shown = progress(run);
assert.deepEqual(Array.from(shown, s => s.id), ['acquisition', ...expected]);
assert.equal(shown.find(s => s.id === 'ftt_parse').status, 'running');
assert.equal(shown.find(s => s.id === 'hfpa_parse').finished_at, '14:03');
assert.equal(shown.find(s => s.id === 'analysis').status, 'pending');
assert.equal(JSON.stringify(run), before, 'Rendering must preserve recorded run state');

const failed = {...run, steps: run.steps.map(s => s.id === 'presentation' ? {...s, status: 'error', finished_at: '14:05'} : s)};
assert.equal(progress(failed).find(s => s.id === 'presentation').status, 'error');
assert.deepEqual(Array.from(progress({mode: 'excel', steps: run.steps}), s => s.id), ['acquisition', ...expected.filter(id => id !== 'presentation')]);
assert.deepEqual(Array.from(progress({mode: 'all', steps: []}), s => s.id), expected);
const other = {...run, steps: [...run.steps, {id: 'unexpected', status: 'error'}]};
assert.equal(progress(other).at(-1).id, 'unexpected', 'Unexpected recorded errors must stay visible');

// Every detailed stage emitted by this report has a planned row before execution.
const report = fs.readFileSync(path.join(root, 'report_tools/quarterly_defect_review/report.py'), 'utf8');
for (const match of report.matchAll(/notify\([^\n]+, '([^']+)'\)/g)) {
  assert.ok(expected.includes(match[1]), `Unplanned stage ${match[1]}`);
}
const renderer = fs.readFileSync(path.join(root, 'static/report_tools.js'), 'utf8');
assert.ok(renderer.includes("run.tool==='quarterly_defect_review'&&window.QDRPanel?.progressSteps?window.QDRPanel.progressSteps(run):(run.steps||[])"));
console.log('Quarterly NIKE progress: order, old runs, failure, Excel mode and stage coverage passed.');
