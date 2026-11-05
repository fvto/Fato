const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const elements = new Map();
let clicks = 0;
function element() {
    return {value:'', style:{}, classList:{add(){}, remove(){}, toggle(){}},
        addEventListener(){}, setAttribute(){}, getAttribute(){return 'dark';},
        appendChild(){}, removeChild(){}, click(){clicks++;}};
}
const context = vm.createContext({console, setTimeout(){}, setInterval(){},
    Date: class extends Date {constructor(...args) {super(...(args.length ? args : [2026,9,24]));}},
    localStorage:{getItem(){return null;}}, window:{addEventListener(){}},
    document:{body:element(),createElement:element,getElementById(id){
        if (!elements.has(id)) elements.set(id,element());
        return elements.get(id);
    }}});
const html = fs.readFileSync(path.join(__dirname,'../templates/index.html'),'utf8');
const script = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)].map(x=>x[1]).join('\n');
vm.runInContext(script, context);
const saved = [];
context.addSavedFilePill = (name,url)=>saved.push({name,url});
context.addLog = ()=>{};
context.result = {success:true,downloaded_files:['VH','VH2','VH3','VH4','JV','JV2','JVB'].map(plant=>
    ({name:`COPQ-${plant}.xls`,url:`/api/exports/batch/COPQ-${plant}.xls`}))};
assert.equal(vm.runInContext('showDynamicExportFiles(result)',context),7);
assert.equal(saved.length,7,'Every selected factory must have a download link');
assert.equal(clicks,0,'Batch links should not rely on browser permission for several automatic downloads');
context.result = {success:true,downloaded_file:saved[0]};
assert.equal(vm.runInContext('showDynamicExportFiles(result)',context),1);
assert.equal(clicks,1,'A single completed Excel downloads directly');
vm.runInContext('showDynamicExportFiles(result)',context);
assert.equal(clicks,1,'Replaying the same result must never download the same file again');
assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calculateDateRange("last_month"))',context)),
    {from:'2026-09-01',to:'2026-09-30'});
vm.runInContext(`currentInspectionData={date_inputs:[{id:'dtDateFrom'},{id:'dtDateTo'}]}`,context);
context.document.getElementById('adaptive-date-dtDateFrom').value='2026-09-01';
context.document.getElementById('adaptive-date-dtDateTo').value='2026-09-30';
assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(dynamicExportParams({index:0}))',context)),
    {date_from:'2026-09-01',date_to:'2026-09-30',index:0});
context.inspection = {page:{title:'Report COPQ',url:'https://sfc.chingluh.com/vhmes/QA/QAReportCOPQ.aspx'},
    date_inputs:[], buttons:[], dropdowns:[
        {id:'cboCategory',type:'combobox',label:'Category',supports_all:false,supports_clear:true,selected_value:'EXT',
            options:[{value:'',text:'-- Không chọn --',is_clear:true},{value:'ALL',text:'ALL',is_all:true},{value:'EXT',text:'External'},{value:'INT',text:'Internal'}]},
        {id:'cboType',type:'combobox',label:'Type',supports_all:false,supports_clear:true,selected_value:null,requires_selection:true,
            options:[{value:'',text:'-- Không chọn --',is_clear:true},{value:'ALL',text:'ALL',is_all:true},{value:'BC',text:'B/C'},{value:'REI',text:'Reinspection'},{value:'REP',text:'Rework'}]}]};
vm.runInContext('applyDynamicInspection({inspection})',context);
const rendered = elements.get('adaptive-controls-body').innerHTML;
assert.doesNotMatch(rendered,/ALL -/,'COPQ must use the native Category and Type options');
assert.match(rendered,/value="EXT" selected/,'Scanned server selection must be preserved');
assert.match(rendered,/-- Chọn Type --/,'Changing Category must require a valid Type when the old value disappears');
assert.equal(elements.get('mirror-url').innerText,context.inspection.page.url);
assert.match(elements.get('detected-counts-text').innerText,/2 Dropdown/);

context.inspection.dropdowns[0].selected_value='';
context.inspection.dropdowns[1].selected_value='';
context.inspection.dropdowns[1].requires_selection=false;
vm.runInContext('applyDynamicInspection({inspection})',context);
assert.match(elements.get('adaptive-controls-body').innerHTML,/value="" selected>-- Không chọn --/);
assert.match(elements.get('adaptive-controls-body').innerHTML,/Xóa tất cả lựa chọn/);
assert.match(elements.get('adaptive-controls-body').innerHTML,/value="ALL"/);
vm.runInContext("selectedDropdownOptions.cboType=inspection.dropdowns[1].options.find(o=>o.value==='ALL');applyDynamicInspection({inspection})",context);
assert.match(elements.get('adaptive-controls-body').innerHTML,/value="ALL" selected>ALL/);
context.result = {success:true,requires_filter_selection:true,inspection:context.inspection,downloaded_files:[]};
context.fetch = async()=>({json:async()=>context.result});
let notice='';
context.showToast = msg=>{notice=msg;};
vm.runInContext(`currentInspectionData={dropdowns:[{id:'btnExportExcel',type:'menu_dropdown',options:[
    {index:1,value:'EXCEL2',text:'COPQ Data',id:'btnExportExcel_DXI0i1_'}]}]};
    selectedDropdownOptions={};`,context);
vm.runInContext('triggerDynamicMenuExport("btnExportExcel")',context).then(async()=>{
    assert.match(notice,/Chọn Category, Type/,'Opening COPQ Data must prompt for filters, not claim an export');
    assert.equal(clicks,1,'Navigation must never automatically export the default Type');
    let clearAction;
    context.fetch = async(url,options)=>{clearAction=JSON.parse(options.body);return {json:async()=>context.result};};
    await vm.runInContext('clearCopqFilterSelections()',context);
    assert.equal(clearAction.action_type,'clear_copq_filters');
    assert.match(notice,/Đã xóa Category và Type/);
    let finish, requests=0;
    context.fetch = ()=>{requests++;return new Promise(resolve=>{finish=resolve;});};
    vm.runInContext('currentInspectionData=inspection',context);
    const first = vm.runInContext('triggerDynamicButtonClick("rpnlMain_ContentPlaceHolder1_btnExcel", "Excel")',context);
    await vm.runInContext('triggerDynamicButtonClick("rpnlMain_ContentPlaceHolder1_btnExcel", "Excel")',context);
    await vm.runInContext('executeCommand()',context);
    assert.equal(requests,1,'A repeated click or command cannot queue another export');
    finish({json:async()=>context.result});
    await first;
    assert.equal(vm.runInContext('dynamicActionBusy',context),false);
    context.inspection.buttons = [{id:'btnExcel',label:'Export Excel',btn_type:'export',icon:'file-excel'}];
    vm.runInContext('applyDynamicInspection({inspection})',context);
    assert.match(elements.get('adaptive-controls-body').innerHTML,/btn-dyn-combine-btnExcel/);
    assert.match(elements.get('adaptive-controls-body').innerHTML,/Xuất Excel \+ Gộp/);
    let combineRequest;
    context.fetch=async(url,options)=>{
        combineRequest=JSON.parse(options.body);
        return {json:async()=>({success:true,
            downloaded_files:[{name:'Combined.xlsx',url:'/api/exports/batch/Combined.xlsx'}],
            source_files:saved.slice(0,7)})};
    };
    const clicksBefore=clicks;
    await vm.runInContext('triggerDynamicButtonClick("btnExcel", "Xuất Excel + Gộp", true)',context);
    assert.equal(combineRequest.params.combine,true);
    assert.equal(combineRequest.action_type,'click_button');
    assert.equal(combineRequest.target_id,'btnExcel');
    assert.equal(clicks,clicksBefore+1,'Only the combined workbook should download automatically');
    assert.equal(elements.get('btn-dyn-btnExcel').disabled,false,'Original export must be enabled after merge');
    assert.ok(saved.some(file=>file.name==='Combined.xlsx'));
    let internalRequest;
    context.fetch=async(url,options)=>{internalRequest=JSON.parse(options.body);return {json:async()=>({success:true,downloaded_files:[]})};};
    vm.runInContext('selectedFttStations.clear()',context);
    await vm.runInContext('executeFttStationExport("internal")',context);
    assert.equal(internalRequest.params.station_selection,'none');
    assert.equal(internalRequest.params.stations.length,0);

    context.mqaaCheck = {page:{pathname:'/vhmes/QA/MQAACheck.aspx'},
        date_inputs:[{id:'dtDate',label:'Date',value:'2026-09-30'}],buttons:[],dropdowns:[]};
    vm.runInContext('applyDynamicInspection({inspection:mqaaCheck})',context);
    const mqaaControls = elements.get('adaptive-controls-body').innerHTML;
    assert.match(mqaaControls,/Date From/);
    assert.match(mqaaControls,/Date To/);
    assert.doesNotMatch(mqaaControls,/adaptive-date-preset-select/);
    context.document.getElementById('adaptive-date-dtDate').value='2026-09-01';
    context.document.getElementById('adaptive-date-mqaa-date-to').value='2026-09-30';
    internalRequest = null;
    await vm.runInContext('triggerDynamicButtonClick("btnSearch", "Search")',context);
    assert.equal(internalRequest,null,'MQAA Check must never silently search one day for a multi-day selection');
    assert.match(notice,/MQAA Report/);

    context.navTo = async destination => {
        assert.equal(destination,'mqaa_report');
        vm.runInContext(`currentInspectionData={page:{pathname:'/vhmes/QA/MQAAReport.aspx'},
            date_inputs:[{id:'dtDateFrom'},{id:'dtDateTo'}]}`,context);
    };
    await vm.runInContext('openMqaaDateRange()',context);
    assert.equal(internalRequest.action_type,'set_dates');
    assert.deepEqual(internalRequest.params,{date_from:'2026-09-01',date_to:'2026-09-30'});
    console.log('COPQ UI: Internal without stations, export deduplication, concurrent clicks, empty selections, clear all, seven links, dates, native filters, dependent Type and navigation scan passed');
}).catch(error=>{console.error(error);process.exitCode=1;});
