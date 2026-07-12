const DATA_PATH='../data/';
const IDX={id:0,risk:1,sub:2,sev:3,reportUnit:4,place:6,summary:12,keyword:13,detail:14,initial:15,suggest:16,mainUnit:17,jointUnit:18,date:27};
const REQUIRED_HEADER='รหัส: เรื่องอุบัติการณ์';
const DEFAULT_SETTINGS={hospital:'โรงพยาบาลดอนตูม',system:'Don Tum Risk Management System',defaultYear:'2569',previewLimit:50,blockErrors:false};

let dataset={fiscalYear:2569,sourceFile:'',sheetName:'',headers:[],rows:[],loadedAt:'',quality:null};
let rcaState={rows:[],manifest:[],selectedIncident:null};
let settings=loadSettings();

const $=s=>document.querySelector(s);
const $$=s=>[...document.querySelectorAll(s)];
const norm=v=>String(v??'').replace(/\r\n/g,'\n').replace(/\r/g,'\n').trim();
const esc=v=>String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
function toast(t){const e=$('#toast');e.textContent=t;e.style.display='block';clearTimeout(window.__toast);window.__toast=setTimeout(()=>e.style.display='none',2600)}
function download(content,name,type){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([content],{type}));a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1500)}
function bytes(n){if(!Number.isFinite(n))return'–';if(n<1024)return`${n} B`;if(n<1048576)return`${(n/1024).toFixed(1)} KB`;return`${(n/1048576).toFixed(1)} MB`}
function nowTh(){return new Date().toLocaleString('th-TH')}
function fmtDate(v){if(!v)return'–';const d=parseDate(v);return d?d.toLocaleDateString('th-TH'):String(v)}
function parseDate(v){
  if(v instanceof Date&&!Number.isNaN(v.getTime()))return v;
  if(typeof v==='number'&&v>0){
    const p=XLSX?.SSF?.parse_date_code?.(v);
    if(p)return new Date(p.y,p.m-1,p.d);
  }
  const s=norm(v);
  if(!s)return null;
  if(/^\d{4}-\d{2}-\d{2}/.test(s)){const d=new Date(s.slice(0,10)+'T00:00:00');return Number.isNaN(d.getTime())?null:d}
  const m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if(m){let y=+m[3];if(y>2400)y-=543;const d=new Date(y,+m[2]-1,+m[1]);return Number.isNaN(d.getTime())?null:d}
  const d=new Date(s);return Number.isNaN(d.getTime())?null:d
}
function isoDate(v){const d=parseDate(v);if(!d)return norm(v)||null;return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`}
function riskType(r){const code=norm(r[IDX.risk]).toUpperCase();if(code.startsWith('C'))return'Clinical';if(code.startsWith('G')||code.startsWith('N'))return'Non-clinical';return'Other'}
function riskCode(r){return norm(r[IDX.risk]).split(':')[0].trim()}
function isHigh(v){return ['E','F','G','H','I','3','4','5'].includes(norm(v).toUpperCase())}
function expectedFiscalYear(dateValue){
  const d=parseDate(dateValue);if(!d)return null;
  let buddhist=d.getFullYear()+543;
  return d.getMonth()+1>=10?buddhist+1:buddhist
}
function loadSettings(){
  try{return {...DEFAULT_SETTINGS,...JSON.parse(localStorage.getItem('drms_admin_settings')||'{}')}}catch{return {...DEFAULT_SETTINGS}}
}
function applySettings(){
  $('#settingHospital').value=settings.hospital;
  $('#settingSystem').value=settings.system;
  $('#settingDefaultYear').value=settings.defaultYear;
  $('#settingPreviewLimit').value=String(settings.previewLimit);
  $('#settingBlockErrors').checked=!!settings.blockErrors;
  $('#fiscalYear').value=settings.defaultYear;
}
function saveSettings(){
  settings={
    hospital:norm($('#settingHospital').value)||DEFAULT_SETTINGS.hospital,
    system:norm($('#settingSystem').value)||DEFAULT_SETTINGS.system,
    defaultYear:$('#settingDefaultYear').value,
    previewLimit:+$('#settingPreviewLimit').value||50,
    blockErrors:$('#settingBlockErrors').checked
  };
  localStorage.setItem('drms_admin_settings',JSON.stringify(settings));
  toast('บันทึกการตั้งค่าแล้ว')
}
function resetSettings(){settings={...DEFAULT_SETTINGS};localStorage.setItem('drms_admin_settings',JSON.stringify(settings));applySettings();toast('คืนค่าเริ่มต้นแล้ว')}

function showTab(id){
  $$('.dc-tab').forEach(x=>x.classList.toggle('active',x.id===id));
  $$('.dc-nav').forEach(x=>x.classList.toggle('active',x.dataset.tab===id))
}
function on(selector,event,handler){
  const el=$(selector);
  if(el)el.addEventListener(event,handler);
}
function bind(){
  $$('.dc-nav').forEach(x=>x.onclick=()=>showTab(x.dataset.tab));
  const openExcel=()=>$('#excelFile')?.click();
  on('#chooseExcel','click',openExcel);on('#browseExcel','click',openExcel);
  on('#excelFile','change',e=>{if(e.target.files[0])readExcel(e.target.files[0])});
  const dz=$('#dropZone');
  if(dz){
    dz.ondragover=e=>{e.preventDefault();e.currentTarget.classList.add('dragover')};
    dz.ondragleave=e=>e.currentTarget.classList.remove('dragover');
    dz.ondrop=e=>{e.preventDefault();e.currentTarget.classList.remove('dragover');const f=e.dataTransfer.files[0];if(f)readExcel(f)}
  }
  on('#loadCurrent','click',loadCurrent);on('#restoreCurrent','click',loadCurrent);
  on('#exportJson','click',exportJson);on('#exportErrors','click',exportIssues);on('#exportBundle','click',exportDashboardBundle);
  on('#previewIssueFilter','change',()=>renderPreview(dataset.quality||qualityCheck([],dataset.fiscalYear)));
  ['#excludeDuplicates','#excludeOutOfYear','#excludeMissingRisk','#excludeMissingUnit'].forEach(s=>on(s,'change',updateExportPolicy));
  on('#clearLog','click',()=>{localStorage.removeItem('drms_admin_log');renderLog();toast('ล้างประวัติแล้ว')});
  on('#downloadBackup','click',downloadBackup);
  on('#chooseRestore','click',()=>$('#restoreFile')?.click());
  on('#restoreFile','change',e=>{if(e.target.files[0])restoreJson(e.target.files[0])});
  on('#fiscalYear','change',()=>{dataset.fiscalYear=+$('#fiscalYear').value;const y=$('#sumYear');if(y)y.textContent=dataset.fiscalYear;if(dataset.rows.length)validateAndRender()});
  on('#saveSettings','click',saveSettings);on('#resetSettings','click',resetSettings);
  on('#loadRcaIncidents','click',loadRcaIncidents);
  on('#refreshRcaManifest','click',loadRcaManifest);
  on('#rcaSearch','input',renderRcaRows);
  on('#rcaAvailability','change',renderRcaRows);
  on('#clearRcaFilter','click',()=>{$('#rcaSearch').value='';$('#rcaAvailability').value='';renderRcaRows()});
  on('#rcaFileInput','change',e=>{if(e.target.files[0])uploadSelectedRca(e.target.files[0]);e.target.value=''});
}

async function readExcel(file){
  if(typeof XLSX==='undefined'){toast('โหลดตัวอ่าน Excel ไม่สำเร็จ กรุณาเชื่อมต่ออินเทอร์เน็ตแล้วรีเฟรช');return}
  try{
    const buffer=await file.arrayBuffer();
    const wb=XLSX.read(buffer,{type:'array',cellDates:true,raw:false});
    const preferred=wb.SheetNames[0],ws=wb.Sheets[preferred];
    const matrix=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:false});
    const headerIndex=findHeaderRow(matrix);
    if(headerIndex<0)throw new Error(`ไม่พบหัวคอลัมน์ "${REQUIRED_HEADER}"`);
    const headers=matrix[headerIndex].map((v,i)=>norm(v)||`คอลัมน์_${i+1}`);
    const rows=matrix.slice(headerIndex+1).map(row=>normalizeRow(row,headers)).filter(row=>row.some(v=>v!==null&&v!==''));
    dataset={fiscalYear:+$('#fiscalYear').value,sourceFile:file.name,sheetName:preferred,headers,rows,loadedAt:new Date().toISOString(),quality:null};
    $('#fileName').textContent=file.name;$('#fileSize').textContent=bytes(file.size);$('#sheetName').textContent=preferred;$('#headerRow').textContent=headerIndex+1;
    validateAndRender();addLog('Import Excel',rows.length,file.name);toast(`อ่านข้อมูล ${rows.length.toLocaleString()} รายการแล้ว`)
  }catch(e){console.error(e);toast(`อ่าน Excel ไม่สำเร็จ: ${e.message}`)}
}
function findHeaderRow(matrix){
  return matrix.slice(0,20).findIndex(row=>row.some(v=>norm(v).includes(REQUIRED_HEADER)))
}
function normalizeRow(row,headers){
  const out=Array(headers.length).fill(null);
  for(let i=0;i<headers.length;i++){
    let v=row[i];
    if(v instanceof Date)v=isoDate(v);
    else if(headers[i].startsWith('วันที่'))v=isoDate(v);
    else if(typeof v==='string'){v=norm(v)||null}
    out[i]=v??null
  }
  while(out.length<28)out.push(null);
  return out
}

async function loadCurrent(){
  const year=+$('#fiscalYear').value;
  try{
    const res=await fetch(`${DATA_PATH}incidents_${year}.json?v=${Date.now()}`);
    if(!res.ok)throw new Error(`HTTP ${res.status}`);
    const obj=await res.json();
    dataset={
      fiscalYear:year,
      sourceFile:obj.sourceFile||`incidents_${year}.json`,
      sheetName:'JSON ปัจจุบัน',
      headers:obj.headers||[],
      rows:Array.isArray(obj.rows)?obj.rows:[],
      loadedAt:new Date().toISOString(),
      quality:null
    };
    $('#fileName').textContent=dataset.sourceFile;$('#fileSize').textContent='–';$('#sheetName').textContent='JSON ปัจจุบัน';$('#headerRow').textContent='–';
    validateAndRender();addLog('Load current',dataset.rows.length,dataset.sourceFile);toast(`โหลดข้อมูลปัจจุบัน ${dataset.rows.length.toLocaleString()} รายการ`)
  }catch(e){console.error(e);toast('โหลดข้อมูลปัจจุบันไม่สำเร็จ')}
}

function validateAndRender(){
  dataset.fiscalYear=+$('#fiscalYear').value;
  dataset.quality=qualityCheck(dataset.rows,dataset.fiscalYear);
  renderAll()
}
function qualityCheck(rows,year){
  const issues=[];
  const duplicateKeys=new Map();
  rows.forEach((r,i)=>{
    const reportId=norm(r[IDX.id]);
    const composite=[isoDate(r[IDX.date]),riskCode(r),norm(r[IDX.reportUnit]),norm(r[IDX.detail]).slice(0,100)].join('|');
    const key=reportId?`ID:${reportId}`:`C:${composite}`;
    if(key!=='C:|||'){if(duplicateKeys.has(key)){issues.push(issue(i,'duplicate','ข้อมูลซ้ำ',`ซ้ำกับรายการที่ ${duplicateKeys.get(key)+1}`))}else duplicateKeys.set(key,i)}
    if(!norm(r[IDX.risk]))issues.push(issue(i,'risk','ไม่มีรหัส/เรื่องความเสี่ยง','ควรตรวจสอบก่อนเผยแพร่'));
    if(!norm(r[IDX.mainUnit]))issues.push(issue(i,'mainUnit','ไม่มีหน่วยงานหลักที่แก้ไข','ไม่สามารถจัดกลุ่ม Risk Profile/Register ได้'));
    if(!norm(r[IDX.sev]))issues.push(issue(i,'severity','ไม่มีระดับความรุนแรง','ควรระบุระดับ'));
    if(!norm(r[IDX.date]))issues.push(issue(i,'date','ไม่มีวันที่เกิดอุบัติการณ์','ควรตรวจสอบวันที่'));
    else if(!parseDate(r[IDX.date]))issues.push(issue(i,'invalidDate','รูปแบบวันที่ไม่ถูกต้อง',norm(r[IDX.date])));
    else{
      const expected=expectedFiscalYear(r[IDX.date]);
      if(expected!==year)issues.push(issue(i,'fiscal','วันที่อยู่นอกช่วงปีงบประมาณ',`วันที่นี้อยู่ในปีงบประมาณ ${expected} แต่ชุดข้อมูลกำหนดเป็น ${year}`))
    }
  });
  const byType={};
  issues.forEach(x=>(byType[x.type]??=[]).push(x));
  const dates=rows.map(r=>parseDate(r[IDX.date])).filter(Boolean).sort((a,b)=>a-b);
  return {
    issues,byType,
    total:rows.length,
    clinical:rows.filter(r=>riskType(r)==='Clinical').length,
    nonClinical:rows.filter(r=>riskType(r)==='Non-clinical').length,
    high:rows.filter(r=>isHigh(r[IDX.sev])).length,
    minDate:dates[0]||null,maxDate:dates.at(-1)||null,
    units:new Set(rows.map(r=>norm(r[IDX.mainUnit])).filter(Boolean)).size,
    codes:new Set(rows.map(r=>riskCode(r)).filter(Boolean)).size
  }
}
function issue(index,type,label,detail){return{index,type,label,detail}}

function renderAll(){
  const q=dataset.quality||qualityCheck([],dataset.fiscalYear);
  $('#sideTotal').textContent=q.total.toLocaleString();
  $('#dataState').textContent=q.total?`พร้อมส่งออก ${q.total.toLocaleString()} รายการ`:'ยังไม่ได้โหลดข้อมูล';
  $('#dataState').classList.toggle('dirty',q.total>0);
  $('#qTotal').textContent=q.total.toLocaleString();$('#qClinical').textContent=q.clinical.toLocaleString();$('#qNon').textContent=q.nonClinical.toLocaleString();$('#qHigh').textContent=q.high.toLocaleString();$('#qWarnings').textContent=q.issues.length.toLocaleString();
  $('#sumYear').textContent=dataset.fiscalYear;$('#sumMinDate').textContent=fmtDate(q.minDate);$('#sumMaxDate').textContent=fmtDate(q.maxDate);$('#sumUnits').textContent=q.units.toLocaleString();$('#sumCodes').textContent=q.codes.toLocaleString();$('#sumDuplicates').textContent=(q.byType.duplicate?.length||0).toLocaleString();
  renderQuality(q);renderPreview(q);renderBackupStatus();renderLog();updateExportPolicy()
}
function renderQuality(q){
  const defs=[
    ['duplicate','ข้อมูลซ้ำ'],
    ['risk','ไม่มีรหัส/เรื่องความเสี่ยง'],
    ['mainUnit','ไม่มีหน่วยงานหลักที่แก้ไข'],
    ['severity','ไม่มีระดับความรุนแรง'],
    ['date','ไม่มีวันที่เกิดอุบัติการณ์'],
    ['invalidDate','รูปแบบวันที่ไม่ถูกต้อง'],
    ['fiscal','วันที่อยู่นอกช่วงปีงบประมาณ']
  ];
  $('#qualityList').innerHTML=defs.map(([type,label])=>{
    const n=q.byType[type]?.length||0;
    return `<div class="quality-row ${n?'warn':'ok'}"><span class="quality-icon">${n?'!':'✓'}</span><div><b>${esc(label)}</b><small>${n?`พบ ${n.toLocaleString()} รายการ`:'ไม่พบปัญหา'}</small></div><strong>${n.toLocaleString()}</strong></div>`
  }).join('');
  const critical=(q.byType.risk?.length||0)+(q.byType.mainUnit?.length||0);
  const badge=$('#qualityBadge');
  badge.className=`quality-badge ${critical?'bad':q.issues.length?'warn':'good'}`;
  badge.textContent=critical?'ควรแก้ไข':q.issues.length?'มีข้อสังเกต':'ผ่านการตรวจ'
}
function rowIssues(q,index){return q.issues.filter(x=>x.index===index)}
function renderPreview(q){
  const limit=settings.previewLimit||50;
  const filter=$('#previewIssueFilter')?.value||'';
  const issueIndexes=new Set(q.issues.map(x=>x.index));
  let indexes=dataset.rows.map((_,i)=>i);
  if(filter==='issues')indexes=indexes.filter(i=>issueIndexes.has(i));
  else if(filter)indexes=indexes.filter(i=>q.issues.some(x=>x.index===i&&x.type===filter));
  else indexes=[...issueIndexes,...indexes.filter(i=>!issueIndexes.has(i))];
  const order=indexes.slice(0,limit);
  $('#previewText').textContent=dataset.rows.length?`แสดง ${order.length} จาก ${indexes.length.toLocaleString()} รายการตามตัวกรอง • ข้อมูลทั้งหมด ${dataset.rows.length.toLocaleString()}`:'ยังไม่มีข้อมูล';
  $('#previewRows').innerHTML=order.map(i=>{
    const r=dataset.rows[i],problems=rowIssues(q,i);
    return `<tr class="${problems.length?'issue-row':''}">
      <td>${i+1}</td><td>${esc(fmtDate(r[IDX.date]))}</td><td>${esc(r[IDX.id])}</td>
      <td><span class="type-pill ${riskType(r)==='Clinical'?'clinical':'non'}">${esc(riskType(r))}</span></td>
      <td>${esc(r[IDX.risk])}</td><td><b>${esc(r[IDX.sev])}</b></td><td>${esc(r[IDX.reportUnit])}</td><td>${esc(r[IDX.mainUnit])}</td>
      <td>${problems.length?problems.map(x=>`<span class="issue-pill">${esc(x.label)}</span>`).join(' '):'<span class="ok-pill">ผ่าน</span>'}</td>
    </tr>`
  }).join('')||'<tr><td colspan="9" class="quality-empty">ไม่พบข้อมูลตามตัวกรอง</td></tr>'
}
function renderBackupStatus(){
  $('#backupYear').textContent=dataset.rows.length?dataset.fiscalYear:'–';
  $('#backupRows').textContent=dataset.rows.length.toLocaleString();
  $('#backupSource').textContent=dataset.sourceFile||'–';
  $('#backupUpdated').textContent=dataset.loadedAt?new Date(dataset.loadedAt).toLocaleString('th-TH'):'–'
}


function exportPolicy(){
  return {
    duplicate:$('#excludeDuplicates')?.checked||false,
    fiscal:$('#excludeOutOfYear')?.checked||false,
    risk:$('#excludeMissingRisk')?.checked||false,
    mainUnit:$('#excludeMissingUnit')?.checked||false
  }
}
function excludedIndexes(){
  const p=exportPolicy(),q=dataset.quality||{issues:[]},set=new Set();
  q.issues.forEach(x=>{if(p[x.type])set.add(x.index)});
  return set
}
function exportRows(){
  const excluded=excludedIndexes();
  return dataset.rows.filter((_,i)=>!excluded.has(i))
}
function updateExportPolicy(){
  const n=exportRows().length,b=$('#exportCountBadge');
  if(b){b.textContent=`พร้อมส่งออก ${n.toLocaleString()} รายการ`;b.className='quality-badge '+(n?'good':'neutral')}
}
function exportObject(){
  const rows=exportRows();
  return {
    fiscalYear:dataset.fiscalYear,
    sourceFile:dataset.sourceFile||'Imported from DRMS Data Center',
    updatedAt:new Date().toISOString(),
    exportPolicy:exportPolicy(),
    originalRows:dataset.rows.length,
    excludedRows:dataset.rows.length-rows.length,
    headers:dataset.headers,
    rows
  }
}
function dashboardMeta(){
  const rows=exportRows();
  return {
    generatedAt:new Date().toISOString(),
    activeFiscalYear:dataset.fiscalYear,
    years:{
      [dataset.fiscalYear]:{
        total:rows.length,
        clinical:rows.filter(r=>riskType(r)==='Clinical').length,
        nonClinical:rows.filter(r=>riskType(r)==='Non-clinical').length,
        highRisk:rows.filter(r=>isHigh(r[IDX.sev])).length,
        sourceFile:dataset.sourceFile||''
      }
    }
  }
}
function exportDashboardBundle(){
  if(!dataset.rows.length){toast('ยังไม่มีข้อมูลสำหรับส่งออก');return}
  const obj=exportObject(),year=dataset.fiscalYear;
  download(JSON.stringify(obj),`incidents_${year}.json`,'application/json');
  setTimeout(()=>download(JSON.stringify(dashboardMeta(),null,2),`meta_${year}.json`,'application/json'),450);
  addLog('Export dashboard bundle',obj.rows.length,`incidents_${year}.json + meta_${year}.json`);
  toast(`ส่งออกชุด Dashboard ${obj.rows.length.toLocaleString()} รายการแล้ว`)
}

function exportJson(){
  if(!dataset.rows.length){toast('ยังไม่มีข้อมูลสำหรับส่งออก');return}
  const critical=(dataset.quality?.byType.risk?.length||0)+(dataset.quality?.byType.mainUnit?.length||0);
  if(settings.blockErrors&&critical){toast(`ยังส่งออกไม่ได้: พบข้อมูลสำคัญไม่ครบ ${critical} รายการ`);return}
  const obj=exportObject();
  download(JSON.stringify(obj),`incidents_${dataset.fiscalYear}.json`,'application/json');
  addLog('Export JSON',obj.rows.length,`incidents_${dataset.fiscalYear}.json`);
  toast(`ส่งออก JSON ${obj.rows.length.toLocaleString()} รายการแล้ว`)
}
function exportIssues(){
  const issues=dataset.quality?.issues||[];
  if(!issues.length){toast('ไม่พบรายการที่ต้องตรวจสอบ');return}
  const header=['ลำดับข้อมูล','ประเภทปัญหา','รายละเอียด','รหัสรายงาน','วันที่','รหัส/เรื่อง','ความรุนแรง','หน่วยงานรายงาน','หน่วยงานหลัก'];
  const body=issues.map(x=>{const r=dataset.rows[x.index]||[];return[x.index+1,x.label,x.detail,r[IDX.id],fmtDate(r[IDX.date]),r[IDX.risk],r[IDX.sev],r[IDX.reportUnit],r[IDX.mainUnit]]});
  const csv='\ufeff'+[header,...body].map(row=>row.map(v=>`"${String(v??'').replaceAll('"','""')}"`).join(',')).join('\n');
  download(csv,`data_quality_${dataset.fiscalYear}.csv`,'text/csv');
  addLog('Export issues',issues.length,`data_quality_${dataset.fiscalYear}.csv`)
}
function downloadBackup(){
  if(!dataset.rows.length){toast('ยังไม่มีข้อมูลสำหรับสำรอง');return}
  const obj={backupVersion:'DRMS-V6.1',createdAt:new Date().toISOString(),dataset,settings};
  download(JSON.stringify(obj,null,0),`drms_backup_${dataset.fiscalYear}_${new Date().toISOString().slice(0,10)}.json`,'application/json');
  addLog('Backup',dataset.rows.length,'DRMS backup');toast('ดาวน์โหลด Backup แล้ว')
}
async function restoreJson(file){
  try{
    const obj=JSON.parse(await file.text());
    const source=obj.dataset||obj;
    if(!Array.isArray(source.rows))throw new Error('ไม่พบ rows ในไฟล์');
    dataset={
      fiscalYear:+(source.fiscalYear||$('#fiscalYear').value),
      sourceFile:source.sourceFile||file.name,
      sheetName:source.sheetName||'Restored JSON',
      headers:source.headers||[],
      rows:source.rows,
      loadedAt:new Date().toISOString(),
      quality:null
    };
    $('#fiscalYear').value=String(dataset.fiscalYear);
    $('#fileName').textContent=file.name;$('#fileSize').textContent=bytes(file.size);$('#sheetName').textContent=dataset.sheetName;$('#headerRow').textContent='–';
    validateAndRender();addLog('Restore',dataset.rows.length,file.name);showTab('dataCenter');toast(`กู้คืน ${dataset.rows.length.toLocaleString()} รายการแล้ว`)
  }catch(e){console.error(e);toast(`กู้คืนไม่สำเร็จ: ${e.message}`)}
}



function isRcaSeverity(v){return ['E','F','G','H','I','3','4','5'].includes(norm(v).toUpperCase())}
function incidentId(r,index){
  const id=norm(r[IDX.id]);
  return id?id.replace(/[^\w\-]/g,'_'):`INC_${dataset.fiscalYear}_${String(index+1).padStart(5,'0')}`
}
async function loadRcaIncidents(){
  const year=+$('#fiscalYear').value;
  try{
    const res=await fetch(`${DATA_PATH}incidents_${year}.json?v=${Date.now()}`);
    if(!res.ok)throw new Error(`HTTP ${res.status}`);
    const obj=await res.json(),rows=Array.isArray(obj.rows)?obj.rows:[];
    rcaState.rows=rows.map((r,i)=>({raw:r,index:i,id:incidentId(r,i)})).filter(x=>isRcaSeverity(x.raw[IDX.sev]));
    await loadRcaManifest(false);
    renderRcaRows();
    toast(`โหลดรายการที่ต้องมี RCA ${rcaState.rows.length.toLocaleString()} รายการ`)
  }catch(e){console.error(e);toast('โหลดรายการ RCA ไม่สำเร็จ')}
}
async function loadRcaManifest(showToast=true){
  try{
    const res=await fetch(`${DATA_PATH}rca.json?v=${Date.now()}`);
    if(res.ok){
      const obj=await res.json();
      rcaState.manifest=Array.isArray(obj)?obj:(obj.items||[])
    }else rcaState.manifest=[];
    renderRcaRows();
    if(showToast)toast('รีเฟรชรายการ RCA แล้ว')
  }catch(e){console.warn(e);rcaState.manifest=[];renderRcaRows()}
}
function rcaEntry(id,year){
  return rcaState.manifest.find(x=>String(x.incident)===String(id)&&Number(x.year)===Number(year))
}
function filteredRcaRows(){
  const q=norm($('#rcaSearch')?.value).toLowerCase(),availability=$('#rcaAvailability')?.value||'',year=+$('#fiscalYear').value;
  return rcaState.rows.filter(x=>{
    const r=x.raw,has=!!rcaEntry(x.id,year);
    const text=[x.id,r[IDX.risk],r[IDX.mainUnit],r[IDX.sev],r[IDX.date]].join(' ').toLowerCase();
    return (!q||text.includes(q))&&(!availability||(availability==='has'?has:!has))
  })
}
function renderRcaRows(){
  if(!$('#rcaRows'))return;
  const year=+$('#fiscalYear').value,rows=filteredRcaRows(),hasCount=rcaState.rows.filter(x=>rcaEntry(x.id,year)).length;
  $('#rcaRequired').textContent=rcaState.rows.length.toLocaleString();
  $('#rcaHas').textContent=hasCount.toLocaleString();
  $('#rcaMissing').textContent=(rcaState.rows.length-hasCount).toLocaleString();
  $('#rcaRows').innerHTML=rows.map(x=>{
    const r=x.raw,e=rcaEntry(x.id,year);
    return `<tr>
      <td>${esc(fmtDate(r[IDX.date]))}</td><td><b>${esc(x.id)}</b></td>
      <td>${esc(r[IDX.risk])}</td><td>${esc(r[IDX.mainUnit])}</td>
      <td><span class="severity-pill">${esc(r[IDX.sev])}</span></td>
      <td>${e?'<span class="rca-has">✓ มี RCA</span>':'<span class="rca-missing">✕ ไม่มี RCA</span>'}</td>
      <td>${e?`<a class="download-link" href="${esc(e.file)}" target="_blank" rel="noopener">ดาวน์โหลด</a>`:'–'}</td>
      <td><button class="prepare-rca" data-incident="${esc(x.id)}">${e?'เตรียมไฟล์แทนที่':'เตรียมไฟล์'}</button></td>
    </tr>`
  }).join('')||'<tr><td colspan="8" class="quality-empty">ไม่พบรายการ</td></tr>';
  $$('.prepare-rca').forEach(b=>b.onclick=()=>{
    rcaState.selectedIncident=b.dataset.incident;
    const path=$('#rcaUploadPath');
    if(path)path.textContent=`drms-v6/rca/${year}/${b.dataset.incident}.pdf`;
    $('#rcaFileInput').click()
  })
}
async function uploadSelectedRca(file){
  const incident=rcaState.selectedIncident,year=+$('#fiscalYear').value;
  if(!incident){toast('ไม่พบ Incident ที่เลือก');return}
  if(!/\.(pdf|doc|docx)$/i.test(file.name)){toast('รองรับเฉพาะ PDF, DOC และ DOCX');return}
  if(file.size>20*1024*1024){toast('ไฟล์ต้องมีขนาดไม่เกิน 20 MB');return}

  const ext=file.name.split('.').pop().toLowerCase();
  const renamed=`${incident}.${ext}`;
  const publicPath=`rca/${year}/${renamed}`;

  // Download renamed RCA file
  download(await file.arrayBuffer(),renamed,file.type||'application/octet-stream');

  // Create updated manifest without touching GitHub directly
  let manifest=Array.isArray(rcaState.manifest)?[...rcaState.manifest]:[];
  manifest=manifest.filter(x=>!(String(x.incident)===String(incident)&&Number(x.year)===year));
  manifest.push({
    incident,
    year,
    file:publicPath,
    filename:renamed,
    originalFilename:file.name,
    preparedAt:new Date().toISOString()
  });
  manifest.sort((a,b)=>Number(a.year)-Number(b.year)||String(a.incident).localeCompare(String(b.incident)));

  setTimeout(()=>download('\ufeff'+JSON.stringify(manifest,null,2),'rca.json','application/json'),500);

  rcaState.manifest=manifest;
  renderRcaRows();
  addLog('Prepare RCA files',1,`${renamed} + rca.json`);
  toast(`เตรียม ${renamed} และ rca.json แล้ว`)
}

function addLog(action,count,file){
  const logs=getLogs();logs.unshift({at:nowTh(),action,count,file,year:dataset.fiscalYear});
  localStorage.setItem('drms_admin_log',JSON.stringify(logs.slice(0,50)));renderLog()
}
function getLogs(){try{return JSON.parse(localStorage.getItem('drms_admin_log')||'[]')}catch{return[]}}
function renderLog(){
  const logs=getLogs();
  $('#importLog').innerHTML=logs.length?logs.map(x=>`<div class="log-row"><div><b>${esc(x.action)}</b><small>${esc(x.file||'')}</small></div><span>ปี ${esc(x.year)} • ${Number(x.count||0).toLocaleString()} รายการ</span><time>${esc(x.at)}</time></div>`).join(''):'<div class="quality-empty">ยังไม่มีประวัติ</div>'
}

applySettings();bind();renderAll();