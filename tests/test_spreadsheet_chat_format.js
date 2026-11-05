'use strict';
const assert=require('node:assert/strict');
const {render,renderComparisons}=require('../static/spreadsheet-chat-format.js');
const malicious=render('<img src=x onerror="alert(1)"> **<script>alert(2)</script>**');
assert.ok(!malicious.includes('<img'));
assert.ok(!malicious.includes('<script>'));
assert.ok(malicious.includes('&lt;img'));
const table=render('| Cell | Value |\n|---|---|\n| **D4** | <b>40</b> |');
assert.ok(table.includes('<table>'));
assert.ok(table.includes('<strong>D4</strong>'));
assert.ok(table.includes('&lt;b&gt;40&lt;/b&gt;'));
assert.ok(render('- A\n- B').includes('<ul><li>A</li><li>B</li></ul>'));
console.log('Chat rendering: safe escaping, tables and bullets passed.');
const comparison={section:'Phần A · Kho nguyên liệu',item:'Điểm đánh giá',status:'difference',
    left:{label:'File A',file:'Audit A.xlsx',section:'Kho nguyên liệu',value:'90%',location:'Audit!D12'},
    right:{label:'File B',file:'Audit B.xlsx',section:'Kho nguyên liệu',value:'95%',location:'Results!H24'},
    explanation:'Hai file khác điểm ở cùng tiêu chí.'};
const cards=renderComparisons([comparison,{...comparison,item:'Ghi chú',status:'review'},
    {...comparison,section:'Phần B',status:'match'}]);
assert.equal((cards.match(/<h3>Phần A/g)||[]).length,1,'Same section groups together');
assert.equal((cards.match(/class="chat-comparison-group"/g)||[]).length,2);
assert.ok(cards.indexOf('Audit A.xlsx')<cards.indexOf('Audit B.xlsx'));
assert.ok(cards.includes('90%')&&cards.includes('95%')&&cards.includes('Cần kiểm tra'));
assert.ok(cards.includes('<summary>Xem ô và nguồn đối chiếu</summary>'));
assert.ok(cards.indexOf('<details')<cards.indexOf('Audit!D12'),'Coordinates stay inside collapsed evidence');
assert.ok(!cards.includes('<details open'));
assert.equal(renderComparisons(undefined),'','Legacy replies have no cards');
assert.equal(renderComparisons([]),'');
const hostile={section:'<img src=x>',item:'<script>x</script>',status:'" onclick="x',
    left:{label:'<svg>',file:'<iframe>',section:'<img>',value:'<script>',location:'<input>'},
    right:null,explanation:'<button onclick="x">'};
const safeCards=renderComparisons([hostile]);
for(const tag of ['img','script','svg','iframe','input','button'])assert.ok(!safeCards.includes('<'+tag));
assert.ok(safeCards.includes('&lt;script&gt;'));
assert.ok(safeCards.includes('class="chat-comparison review"'));
assert.ok(renderComparisons([comparison],'en').includes('View cells & sources'));
const internal=renderComparisons([{...comparison,left:{...comparison.left,label:'Tổng hợp'},
    right:{...comparison.right,label:'Chi tiết / tính lại',file:comparison.left.file}}]);
assert.ok(internal.includes('Chi tiết / tính lại'));
assert.ok(!internal.includes('File B'));
console.log('Comparison cards: grouping, paired values, internal checks, evidence and escaping passed.');
