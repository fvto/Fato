/* Small, deliberately limited renderer. All model text is escaped before markup. */
(function(root) {
    'use strict';
    const escape=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    const inline=value=>escape(value).replace(/\*\*([^*\n]+)\*\*/g,'<strong>$1</strong>').replace(/`([^`\n]+)`/g,'<code>$1</code>');
    const cells=line=>line.trim().replace(/^\|/,'').replace(/\|$/,'').split('|').map(value=>value.trim());
    function render(text) {
        return String(text??'').split(/\n\s*\n/).map(block=>{
            const lines=block.split('\n');
            if(lines.length>=2&&lines[0].includes('|')&&/^\s*\|?[\s:|-]+\|?\s*$/.test(lines[1])&&lines[1].includes('---')) {
                return '<div class="chat-table-scroll"><table><thead><tr>'+cells(lines[0]).map(cell=>'<th>'+inline(cell)+'</th>').join('')+'</tr></thead><tbody>'+lines.slice(2).map(line=>'<tr>'+cells(line).map(cell=>'<td>'+inline(cell)+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>';
            }
            if(lines.every(line=>/^\s*[-*] /.test(line)))return '<ul>'+lines.map(line=>'<li>'+inline(line.replace(/^\s*[-*] /,''))+'</li>').join('')+'</ul>';
            return '<p>'+lines.map(inline).join('<br>')+'</p>';
        }).join('');
    }
    function renderComparisons(items, language='vi') {
        if(!Array.isArray(items)||!items.length)return '';
        const words=language==='en'
            ?{difference:'Difference',match:'Matches',review:'Needs review',evidence:'View cells & sources',left:'File A',right:'File B'}
            :{difference:'Có chênh lệch',match:'Khớp',review:'Cần kiểm tra',evidence:'Xem ô và nguồn đối chiếu',left:'File A',right:'File B'};
        const groups=new Map();
        for(const item of items) {
            if(!item||typeof item!=='object')continue;
            const section=String(item.section|| (language==='en'?'Comparison':'Đối chiếu'));
            if(!groups.has(section))groups.set(section,[]);
            groups.get(section).push(item);
        }
        const side=(value,label)=>{
            const data=value||{};
            return '<div class="chat-comparison-side"><p class="chat-side-label">'+escape(data.label||label)+'</p>'+
                '<p class="chat-side-file">'+escape(data.file)+'</p><p class="chat-side-section">'+escape(data.section)+'</p>'+
                '<p class="chat-side-value">'+escape(data.value)+'</p></div>';
        };
        return '<div class="chat-comparisons">'+Array.from(groups,([section,rows])=>
            '<section class="chat-comparison-group"><h3>'+escape(section)+'</h3>'+rows.map(item=>{
                const status=['difference','match','review'].includes(item.status)?item.status:'review';
                const sources=[item.left,item.right].filter(value=>value?.location).map(value=>
                    '<li><strong>'+escape(value.label)+' · '+escape(value.file)+'</strong><code>'+escape(value.location)+'</code></li>').join('');
                return '<article class="chat-comparison '+status+'"><div class="chat-comparison-heading"><h4>'+escape(item.item)+'</h4>'+
                    '<span class="chat-comparison-status">'+words[status]+'</span></div><div class="chat-comparison-pair">'+
                    side(item.left,words.left)+side(item.right,words.right)+'</div><p class="chat-comparison-explanation">'+escape(item.explanation)+'</p>'+
                    (sources?'<details class="chat-comparison-evidence"><summary>'+words.evidence+'</summary><ul>'+sources+'</ul></details>':'')+'</article>';
            }).join('')+'</section>').join('')+'</div>';
    }
    if(typeof module==='object'&&module.exports)module.exports={render,renderComparisons};
    else root.SpreadsheetChatFormat={render,renderComparisons};
})(typeof window==='object'?window:this);
