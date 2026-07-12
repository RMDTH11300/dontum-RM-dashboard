const DATA_PATH='../data/';
const IDX={id:0,risk:1,sub:2,sev:3,reportUnit:4,place:6,summary:12,keyword:13,detail:14,initial:15,suggest:16,mainUnit:17,jointUnit:18,date:27};
const REQUIRED_HEADER='รหัส: เรื่องอุบัติการณ์';
const DEFAULT_SETTINGS={hospital:'โรงพยาบาลดอนตูม',system:'Don Tum Risk Management System',defaultYear:'2569',previewLimit:50,blockErrors:false};

let dataset={fiscalYear:2569,sourceFile:'',sheetName:'',headers:[],rows:[],loadedAt:'',quality:null};
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
function bind(){
  $$('.dc-nav').forEach(x=>x.onclick=()=>showTab(x.dataset.tab));
  $('#chooseExcel').onclick=$('#browseExcel').onclick=()=>$('#excelFile').click();
  $('#excelFile').onchange=e=>{if(e.target.files[0])readExcel(e.target.files[0])};
  $('#dropZone').ondragover=e=>{e.preventDefault();e.currentTarget.classList.add('dragover')};
  $('#dropZone').ondragleave=e=>e.currentTarget.classList.remove('dragover');
  $('#dropZone').ondrop=e=>{e.preventDefault();e.currentTarget.classList.remove('dragover');const f=e.dataTransfer.files[0];if(f)readExcel(f)};
  $('#loadCurrent').onclick=$('#restoreCurrent').onclick=loadCurrent;
  $('#exportJson').onclick=exportJson;
  $('#exportErrors').onclick=exportIssues;
  $('#clearLog').onclick=()=>{localStorage.removeItem('drms_admin_log');renderLog();toast('ล้างประวัติแล้ว')};
  $('#downloadBackup').onclick=downloadBackup;
  $('#chooseRestore').onclick=()=>$('#restoreFile').click();
  $('#restoreFile').onchange=e=>{if(e.target.files[0])restoreJson(e.target.files[0])};
  $('#fiscalYear').onchange=()=>{dataset.fiscalYear=+$('#fiscalYear').value;$('#sumYear').textContent=dataset.fiscalYear;if(dataset.rows.length)validateAndRender()};
  $('#saveSettings').onclick=saveSettings;
  $('#resetSettings').onclick=resetSettings
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
  renderQuality(q);renderPreview(q);renderBackupStatus();renderLog()
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
  const issueIndexes=new Set(q.issues.map(x=>x.index));
  const order=[...issueIndexes,...dataset.rows.map((_,i)=>i).filter(i=>!issueIndexes.has(i))].slice(0,limit);
  $('#previewText').textContent=dataset.rows.length?`แสดง ${order.length} จาก ${dataset.rows.length.toLocaleString()} รายการ โดยนำรายการที่ต้องตรวจสอบขึ้นก่อน`:'ยังไม่มีข้อมูล';
  $('#previewRows').innerHTML=order.map((i,pos)=>{
    const r=dataset.rows[i],problems=rowIssues(q,i);
    return `<tr class="${problems.length?'issue-row':''}">
      <td>${i+1}</td><td>${esc(fmtDate(r[IDX.date]))}</td><td>${esc(r[IDX.id])}</td>
      <td><span class="type-pill ${riskType(r)==='Clinical'?'clinical':'non'}">${esc(riskType(r))}</span></td>
      <td>${esc(r[IDX.risk])}</td><td><b>${esc(r[IDX.sev])}</b></td><td>${esc(r[IDX.reportUnit])}</td><td>${esc(r[IDX.mainUnit])}</td>
      <td>${problems.length?problems.map(x=>`<span class="issue-pill">${esc(x.label)}</span>`).join(' '):'<span class="ok-pill">ผ่าน</span>'}</td>
    </tr>`
  }).join('')||'<tr><td colspan="9" class="quality-empty">ยังไม่มีข้อมูลสำหรับแสดง</td></tr>'
}
function renderBackupStatus(){
  $('#backupYear').textContent=dataset.rows.length?dataset.fiscalYear:'–';
  $('#backupRows').textContent=dataset.rows.length.toLocaleString();
  $('#backupSource').textContent=dataset.sourceFile||'–';
  $('#backupUpdated').textContent=dataset.loadedAt?new Date(dataset.loadedAt).toLocaleString('th-TH'):'–'
}

function exportJson(){
  if(!dataset.rows.length){toast('ยังไม่มีข้อมูลสำหรับส่งออก');return}
  const critical=(dataset.quality?.byType.risk?.length||0)+(dataset.quality?.byType.mainUnit?.length||0);
  if(settings.blockErrors&&critical){toast(`ยังส่งออกไม่ได้: พบข้อมูลสำคัญไม่ครบ ${critical} รายการ`);return}
  const obj={fiscalYear:dataset.fiscalYear,sourceFile:dataset.sourceFile||'Imported from DRMS Data Center',updatedAt:new Date().toISOString(),headers:dataset.headers,rows:dataset.rows};
  download(JSON.stringify(obj,null,0),`incidents_${dataset.fiscalYear}.json`,'application/json');
  addLog('Export JSON',dataset.rows.length,`incidents_${dataset.fiscalYear}.json`);
  toast('ส่งออก JSON แล้ว')
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