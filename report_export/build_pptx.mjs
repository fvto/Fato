import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {Presentation,PresentationFile} from '@oai/artifact-tool';

const [inputPath,outputPath,runtimePython,skillDir]=process.argv.slice(2);
if(!inputPath||!outputPath||!runtimePython||!skillDir)throw new Error('Missing PowerPoint export arguments.');
const model=JSON.parse(await fs.readFile(inputPath,'utf8'));
const workspaceDir=path.dirname(inputPath);
const previewDir=path.join(workspaceDir,'previews');
await fs.mkdir(previewDir,{recursive:true});
await fs.mkdir(path.dirname(outputPath),{recursive:true});
const presentation=Presentation.create({slideSize:{width:1600,height:900}});
const font='Arial',ink='#172B42',muted='#526274';
presentation.theme.defaultFont=font;
const tableOwners=[];
function text(slide,value,left,top,width,height,size=22,bold=false,color=ink){
  const shape=slide.shapes.add({geometry:'textbox',position:{left,top,width,height},fill:'none',line:{fill:'none',width:0}});
  shape.text=String(value);shape.text.style={typeface:font,fontSize:size,bold,color};
  return shape;
}
function slide(title,subtitle=''){
  const s=presentation.slides.add();s.background.fill='#FFFFFF';
  text(s,title,64,40,1472,62,38,true);
  if(subtitle)text(s,subtitle,64,106,1472,60,20,false,muted);
  text(s,`Sai số 0 · ${presentation.slides.items.length}`,64,852,1472,28,17,false,muted);
  s.speakerNotes.text=JSON.stringify({scope:model.scope,policy:model.policy,sources:model.sources},null,2);
  return s;
}
function table(s,values,widths,{top=184,rowHeight=52,fontSize=22}={}){
  const t=s.tables.add({rows:values.length,columns:widths.length,left:64,top,width:1472,height:rowHeight*values.length,columnWidths:widths,values});
  t.borders.assign({fill:'#CCD6DF',width:1,style:'solid'});
  for(let r=0;r<values.length;r++){
    t.rows[r].height=rowHeight;
    for(let c=0;c<widths.length;c++){
      const cell=t.getCell(r,c);
      cell.fill=r===0?'#172B42':r%2?'#FFFFFF':'#F1F5F8';
      cell.text.style={typeface:font,fontSize,bold:r===0,color:r===0?'#FFFFFF':ink};
    }
  }
  tableOwners.push(presentation.slides.items.length);
  return t;
}
const groupChunks=[];
for(let i=0;i<model.groups.length;i+=10)groupChunks.push(model.groups.slice(i,i+10));
if(!groupChunks.length)groupChunks.push([]);
for(const [index,groups]of groupChunks.entries()){
  const s=slide(index?'Tổng hợp chênh lệch (tiếp)':'Báo cáo chênh lệch',`${model.count} mục khác nhau · ${model.groups.length} nhóm. ${model.scope}\n${model.comparison_policy}`);
  table(s,[['Phần kiểm tra','Vấn đề','Số mục'],...groups.map(g=>[g.name,g.status,String(g.rows.length)])],[1010,340,122],{rowHeight:55,fontSize:22});
  text(s,'Chi tiết giá trị ở các trang sau. Ảnh highlight được gom vào phụ lục.',64,812,1472,30,19,false,muted);
}
function short(value,max=30){
  const s=String(value);
  if(s.length<=max)return s;
  const match=s.match(/^(-?\d+\.\d+)/);
  if(match){const suffix=s.endsWith('%')?'%':'';return match[1].slice(0,max-4)+'…'+suffix;}
  return s.slice(0,max-1)+'…';
}
for(let i=0;i<model.findings.length;i+=15){
  const rows=model.findings.slice(i,i+15);
  const s=slide(`Chi tiết ${i+1}–${i+rows.length}`, `Hiển thị tối đa 3 số lẻ, bỏ số 0 thừa. ${model.comparison_policy}. Giá trị gốc nằm trong ghi chú.`);
  table(s,[['Mục / vị trí','Chỉ tiêu','Báo cáo','Tính lại / nguồn','Chênh lệch'],...rows.map(r=>[`${r.index}. ${r.location}`,short(r.label,36),short(r.reported,20),short(r.expected,22),short(r.delta,16)])],[300,430,235,285,222],{top:176,rowHeight:40,fontSize:20});
  s.speakerNotes.text=JSON.stringify({sources:model.sources,scope:model.scope,findings:rows},null,2);
}
for(let i=0;i<model.evidence.length;i+=2){
  const pair=model.evidence.slice(i,i+2);
  const s=slide(`Ảnh highlight ${i+1}–${i+pair.length}`, 'Vàng: ô cần kiểm tra. Xanh: đầu vào. Ảnh dựng từ dữ liệu Excel.');
  for(const [j,entry]of pair.entries()){
    const left=pair.length===1?64:64+j*752;
    const width=pair.length===1?1472:720;
    text(s,entry.caption,left,168,width,46,21,true);
    s.images.add({blob:new Uint8Array(Buffer.from(entry.png,'base64')),contentType:'image/png',alt:entry.caption,fit:'contain',position:{left,top:220,width,height:602}});
  }
  s.speakerNotes.text=JSON.stringify({sources:model.sources,evidence:pair.map(x=>({id:x.id,caption:x.caption}))},null,2);
}
for(const [index,s]of presentation.slides.items.entries()){
  const png=await presentation.export({slide:s,format:'png',scale:1});
  await fs.writeFile(path.join(previewDir,`slide-${String(index+1).padStart(2,'0')}.png`),new Uint8Array(await png.arrayBuffer()));
}
const candidatePath=path.join(workspaceDir,'candidate.pptx');
await(await PresentationFile.exportPptx(presentation)).save(candidatePath);
const {finalizePresentation}=await import(pathToFileURL(path.join(skillDir,'container_tools/artifact_tool_utils.mjs')).href);
await finalizePresentation({workspaceDir,candidatePath,finalPath:outputPath,pythonExecutable:runtimePython,
  integrityValidatorPath:path.join(skillDir,'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath:path.join(skillDir,'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs:['--expected-slide-size-emu','15240000,8572500','--validate-heading-fit',...tableOwners.flatMap(n=>['--require-native-table-slide',String(n)])],
  requiredNativeTableOwnerSlides:tableOwners,fontPolicy:{basis:'design',families:[font]},verifyArtifactToolImport:true,
  receiptPath:path.join(workspaceDir,'validation.json')});
await fs.writeFile(path.join(workspaceDir,'manifest.json'),JSON.stringify({slides:presentation.slides.items.length,tableOwners,previewDir,outputPath}));
// All artifacts/validation receipts are complete; exit before native renderer teardown.
process.stdout.write(JSON.stringify({slides:presentation.slides.items.length,outputPath}),()=>process.exit(0));
