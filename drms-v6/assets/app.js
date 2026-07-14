const PARENT='../data/';
const IDX={id:0,risk:1,sub:2,sev:3,reportUnit:4,place:6,summary:12,keyword:13,detail:14,initial:15,suggest:16,mainUnit:17,status:26,date:27};
const state={year:'all',years:[],rows:[],filtered:[],units:[],selected:new Set(),profiles:{},registers:{},orgTree:{groups:[]},essentialStandards:[],selectedEssential:null,rcaManifest:[],allYearRows:{},page:1,pageSize:25,view:'dashboard',profileMode:'summary',registerMode:'summary'};
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const esc=v=>String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
function toast(t){const e=$('#toast');e.textContent=t;e.style.display='block';clearTimeout(window.__toast);window.__toast=setTimeout(()=>e.style.display='none',2400)}
async function json(url){const r=await fetch(url+'?v='+Date.now());if(!r.ok)throw Error(`${r.status} ${url}`);return r.json()}
async function init(){
  try{
    const meta=await json(PARENT+'meta.json');
    state.profiles=await json(PARENT+'profiles.json');
    state.registers=await json(PARENT+'registers.json');

    try{state.orgTree=await json('config/departments.json')}
    catch(e){console.warn('ไม่พบ departments.json ใช้รายชื่อหน่วยงานแบบเดิม',e);state.orgTree={groups:[]}}

    try{
      const e=await json('config/essential-standards.json');
      state.essentialStandards=e.standards||[]
    }catch(e){
      console.warn('ไม่พบ essential-standards.json',e);
      state.essentialStandards=[]
    }

    try{
      const r=await json(PARENT+'rca.json');
      state.rcaManifest=Array.isArray(r)?r:(r.items||[])
    }catch(e){state.rcaManifest=[]}

    state.years=(meta.years||[]).map(Number).filter(Boolean);
    $('#year').innerHTML='<option value="all">ทุกปีงบประมาณ</option>'+
      state.years.map(y=>`<option value="${y}">${y}</option>`).join('');
    state.year='all';
    $('#year').value='all';

    buildMonths();
    bind();
    await loadAllYears();
  }catch(e){
    console.error(e);
    toast('เริ่มระบบไม่สำเร็จ ตรวจสอบไฟล์ data')
  }
}
function buildMonths(){const vals=[10,11,12,1,2,3,4,5,6,7,8,9],names=['ตุลาคม','พฤศจิกายน','ธันวาคม','มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน'];$('#month').innerHTML='<option value="">ทุกเดือนงบประมาณ</option>'+vals.map((m,i)=>`<option value="${m}">${names[i]}</option>`).join('')}
function bind(){
  $$('.nav[data-view]').forEach(b=>b.onclick=()=>show(b.dataset.view));
  $('#year').onchange=async e=>{
    const value=e.target.value;
    state.year=value==='all'?'all':Number(value);
    state.selected.clear();
    state.page=1;
    if(state.year==='all'){
      show('dashboard');
      await loadAllYears()
    }else{
      await loadYear()
    }
  };
  $('#allUnits').onclick=()=>{state.selected=new Set(state.units);renderUnits();apply()};
  $('#expandUnits').onclick=()=>{
    const details=$$('#units details');
    const shouldOpen=details.some(x=>!x.open);
    details.forEach(x=>x.open=shouldOpen);
    $('#expandUnits').textContent=shouldOpen?'ย่อทั้งหมด':'ขยายทั้งหมด'
  };
  $('#clearUnits').onclick=()=>{state.selected.clear();renderUnits();apply()};
  $('#unitSearch').oninput=renderUnits;$('#refresh').onclick=()=>state.year==='all'?loadAllYears():loadYear();
  ['#q','#sev','#type','#month'].forEach(s=>$(s).addEventListener(s==='#q'?'input':'change',()=>{state.page=1;apply()}));
  $('#reset').onclick=()=>{$('#q').value='';$('#sev').value='';$('#type').value='';$('#month').value='';state.selected.clear();state.page=1;renderUnits();apply()};
  $('#csv').onclick=exportCsv;
  $('#printProfile').onclick=()=>window.print();$('#printRegister').onclick=()=>window.print();$('#printAnalytics').onclick=()=>window.print();$('#printHa').onclick=()=>window.print();
  $('#exportHaCsv').onclick=exportHaCsv;$('#refreshHaReport').onclick=()=>{renderHaReport(true);toast('สร้างสรุปรายงานใหม่แล้ว')};$('#saveHaDraft').onclick=saveHaReportDraft;$('#exportHaWord').onclick=()=>exportHaReport('word');$('#exportHaExcel').onclick=()=>exportHaReport('excel');$('#exportProfileCsv').onclick=exportProfileCsv;$('#exportRegisterCsv').onclick=exportRegisterCsv;
  $('#profileSummaryMode').onclick=()=>setProfileMode('summary');$('#profileHaMode').onclick=()=>setProfileMode('ha');
  $('#registerSummaryMode').onclick=()=>setRegisterMode('summary');$('#registerHaMode').onclick=()=>setRegisterMode('ha');
  $('#saveProfileDraft').onclick=()=>saveHaDrafts('profile');$('#saveRegisterDraft').onclick=()=>saveHaDrafts('register');
  $('#exportProfileWord').onclick=()=>exportTableDocument('profile','word');
  $('#exportProfileExcel').onclick=()=>exportTableDocument('profile','excel');
  $('#exportRegisterWord').onclick=()=>exportTableDocument('register','word');
  $('#exportRegisterExcel').onclick=()=>exportTableDocument('register','excel');
  $('#exportEssentialCsv').onclick=exportEssentialCsv;
  $('#exportEssentialExcel').onclick=exportEssentialExcel;
  $('#printEssential').onclick=()=>window.print();

  const registerTab=$('#essentialRegisterTab');
  const matrixTab=$('#essentialMatrixTab');
  if(registerTab)registerTab.onclick=()=>showEssentialMode('register');
  if(matrixTab)matrixTab.onclick=()=>showEssentialMode('matrix');
}
function show(v){state.view=v;$$('.view').forEach(x=>x.classList.toggle('active',x.id===v));$$('.nav[data-view]').forEach(x=>x.classList.toggle('active',x.dataset.view===v));if(v==='profile'){renderProfile();renderProfileHa()}if(v==='register'){renderRegister();renderRegisterHa()}if(v==='analytics')renderAnalytics();if(v==='essential')renderEssentialStandards();if(v==='rca')renderPublicRca();if(v==='haReport')renderHaReport();window.scrollTo({top:0,behavior:'smooth'})}

async function loadAllYears(){
  const years=(state.years||[]).map(Number).filter(Boolean);
  const all=[];
  state.allYearRows=state.allYearRows||{};

  for(const y of years){
    try{
      let rows=state.allYearRows[y];
      if(!rows){
        const obj=await json(PARENT+`incidents_${y}.json`);
        rows=Array.isArray(obj.rows)?obj.rows:[];
        state.allYearRows[y]=rows;
      }
      rows.forEach(r=>{
        const copy=[...r];
        copy.__fiscalYear=y;
        all.push(copy)
      })
    }catch(e){
      console.warn(`โหลดปี ${y} ไม่สำเร็จ`,e)
    }
  }

  state.rows=all;
  state.units=[...new Set(all.map(r=>normUnit(r[IDX.mainUnit])).filter(Boolean))]
    .sort((a,b)=>a.localeCompare(b,'th'));

  const sevs=[...new Set(all.map(r=>String(r[IDX.sev]||'').trim()).filter(Boolean))]
    .sort((a,b)=>a.localeCompare(b,'th'));
  $('#sev').innerHTML='<option value="">ทุกระดับ</option>'+
    sevs.map(x=>`<option>${esc(x)}</option>`).join('');

  renderUnits();
  apply();
  toast(`โหลดทุกปีงบประมาณ ${years.length} ปี รวม ${all.length.toLocaleString()} รายการ`)
}

async function loadYear(){
  if(state.year==='all')return loadAllYears();

  try{
    const d=await json(PARENT+`incidents_${state.year}.json`);
    state.rows=d.rows||[];
    state.allYearRows[state.year]=state.rows;
    state.units=[...new Set(state.rows.map(r=>normUnit(r[IDX.mainUnit])).filter(Boolean))]
      .sort((a,b)=>a.localeCompare(b,'th'));

    const sevs=[...new Set(state.rows.map(r=>String(r[IDX.sev]||'').trim()).filter(Boolean))]
      .sort((a,b)=>a.localeCompare(b,'th'));
    $('#sev').innerHTML='<option value="">ทุกระดับ</option>'+
      sevs.map(x=>`<option>${esc(x)}</option>`).join('');

    renderUnits();
    apply();
    toast(`โหลดปี ${state.year} จำนวน ${state.rows.length.toLocaleString()} รายการ`)
  }catch(e){
    console.error(e);
    toast('โหลดข้อมูลไม่สำเร็จ ตรวจสอบโครงสร้างไฟล์')
  }
}
function normUnit(v){return String(v||'ไม่ระบุ').trim()||'ไม่ระบุ'}
function allDescendantUnits(node){
  if(Array.isArray(node.units))return node.units.filter(u=>state.units.includes(u));
  if(Array.isArray(node.departments))return node.departments.flatMap(allDescendantUnits);
  return[]
}
function checkboxState(units){
  const count=units.filter(u=>state.selected.has(u)).length;
  return{checked:units.length>0&&count===units.length,partial:count>0&&count<units.length}
}

function treeMatches(group,q){
  if(!q)return true;
  const text=[group.name,...group.departments.flatMap(d=>[d.name,...d.units])].join(' ').toLowerCase();
  return text.includes(q)
}
function renderUnits(){
  const q=$('#unitSearch').value.trim().toLowerCase();
  $('#selectedCount').textContent=state.selected.size?`${state.selected.size} งานย่อยที่เลือก`:'ทั้งหมด';
  const groups=(state.orgTree.groups||[]).filter(g=>treeMatches(g,q));
  let out=groups.map((g,gi)=>{
    const gUnits=allDescendantUnits(g),gs=checkboxState(gUnits);
    const departments=g.departments.filter(d=>!q||[d.name,...d.units].join(' ').toLowerCase().includes(q)||g.name.toLowerCase().includes(q));
    const deptHtml=departments.map((d,di)=>{
      const dUnits=allDescendantUnits(d),ds=checkboxState(dUnits);
      const leaves=d.units.filter(u=>state.units.includes(u)&&(!q||[g.name,d.name,u].join(' ').toLowerCase().includes(q)));
      if(!leaves.length)return'';
      return `<details class="org-dept" open>
        <summary><label class="org-parent"><input type="checkbox" data-kind="dept" data-units="${esc(dUnits.join('||'))}" ${ds.checked?'checked':''}><span>${esc(d.name)}</span></label></summary>
        <div class="org-leaves">${leaves.map(u=>`<label class="org-leaf"><input type="checkbox" data-kind="leaf" value="${esc(u)}" ${state.selected.has(u)?'checked':''}><span>${esc(u)}</span></label>`).join('')}</div>
      </details>`
    }).join('');
    if(!deptHtml)return'';
    return `<details class="org-group" open>
      <summary><label class="org-parent group"><input type="checkbox" data-kind="group" data-units="${esc(gUnits.join('||'))}" ${gs.checked?'checked':''}><span>${esc(g.name)}</span></label></summary>
      <div class="org-departments">${deptHtml}</div>
    </details>`
  }).join('');
  $('#units').innerHTML=out||'<p class="empty">ไม่พบหน่วยงานใน Master</p>';

  $$('#units input[data-kind="leaf"]').forEach(x=>x.onchange=()=>{
    x.checked?state.selected.add(x.value):state.selected.delete(x.value);
    state.page=1;renderUnits();apply()
  });
  $$('#units input[data-kind="group"],#units input[data-kind="dept"]').forEach(x=>{
    const units=(x.dataset.units||'').split('||').filter(Boolean);
    const st=checkboxState(units);x.indeterminate=st.partial;
    x.onchange=()=>{
      units.forEach(u=>x.checked?state.selected.add(u):state.selected.delete(u));
      state.page=1;renderUnits();apply()
    }
  })
}
function riskType(r){const text=String(r[IDX.risk]||'').toUpperCase();if(text.startsWith('C'))return'Clinical';if(text.startsWith('G')||text.startsWith('N'))return'Non-clinical';return'อื่นๆ'}
function isHigh(s){return ['E','F','G','H','I','3','4','5'].includes(String(s||'').trim().toUpperCase())}
function monthOf(v){if(!v)return'';const d=new Date(v);return Number.isNaN(d.getTime())?'':d.getMonth()+1}
function rowMatchesSelectedUnit(v){if(!state.selected.size)return true;const text=normUnit(v);if(state.selected.has(text))return true;return [...state.selected].some(u=>text.split(/[\n,;]+/).map(x=>x.trim()).includes(u))}
function apply(){const q=($('#q')?.value||'').trim().toLowerCase(),sev=$('#sev')?.value||'',type=$('#type')?.value||'',month=+($('#month')?.value||0);state.filtered=state.rows.filter(r=>rowMatchesSelectedUnit(r[IDX.mainUnit])&&(!sev||String(r[IDX.sev]||'')===sev)&&(!type||riskType(r)===type)&&(!month||monthOf(r[IDX.date])===month)&&(!q||r.some(v=>String(v??'').toLowerCase().includes(q))));renderDashboard();renderIncidents();if(state.view==='profile'){renderProfile();renderProfileHa()}if(state.view==='register'){renderRegister();renderRegisterHa()}if(state.view==='analytics')renderAnalytics();if(state.view==='essential')renderEssentialStandards();if(state.view==='rca')renderPublicRca();if(state.view==='haReport')renderHaReport()}
function renderDashboard(){const a=state.filtered;$('#kTotal').textContent=a.length.toLocaleString();$('#kClinical').textContent=a.filter(r=>riskType(r)==='Clinical').length.toLocaleString();$('#kNon').textContent=a.filter(r=>riskType(r)==='Non-clinical').length.toLocaleString();$('#kHigh').textContent=a.filter(r=>isHigh(r[IDX.sev])).length.toLocaleString();const months=[10,11,12,1,2,3,4,5,6,7,8,9],names=['ต.ค.','พ.ย.','ธ.ค.','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.'];barChart('#monthly',months.map((m,i)=>({label:names[i],count:a.filter(r=>monthOf(r[IDX.date])===m).length})));renderSeveritySplit(a);rank('#topRisks',a.map(r=>String(r[IDX.risk]||'ไม่ระบุ').split(':')[0].trim()));rank('#topUnits',a.map(r=>normUnit(r[IDX.mainUnit]))) }

function renderSeveritySplit(rows){
  const clinicalLevels=['A','B','C','D','E','F','G','H','I'];
  const nonClinicalLevels=['1','2','3','4','5'];

  const clinicalRows=rows.filter(r=>riskType(r)==='Clinical');
  const nonClinicalRows=rows.filter(r=>riskType(r)==='Non-clinical');

  const clinicalCounts=Object.fromEntries(clinicalLevels.map(x=>[x,0]));
  const nonClinicalCounts=Object.fromEntries(nonClinicalLevels.map(x=>[x,0]));

  clinicalRows.forEach(r=>{
    const level=String(r[IDX.sev]||'').trim().toUpperCase();
    if(level in clinicalCounts)clinicalCounts[level]++
  });
  nonClinicalRows.forEach(r=>{
    const level=String(r[IDX.sev]||'').trim();
    if(level in nonClinicalCounts)nonClinicalCounts[level]++
  });

  severityBarChart('#severityClinical',clinicalLevels.map(level=>({
    label:level,
    count:clinicalCounts[level],
    color:clinicalSeverityColor(level)
  })));

  severityBarChart('#severityNonClinical',nonClinicalLevels.map(level=>({
    label:level,
    count:nonClinicalCounts[level],
    color:nonClinicalSeverityColor(level)
  })));
}
function clinicalSeverityColor(level){
  return ({
    A:'#22c55e',
    B:'#4ade80',
    C:'#84cc16',
    D:'#eab308',
    E:'#f59e0b',
    F:'#f97316',
    G:'#ef4444',
    H:'#dc2626',
    I:'#991b1b'
  })[level]||'#94a3b8'
}
function nonClinicalSeverityColor(level){
  return ({
    '1':'#22c55e',
    '2':'#84cc16',
    '3':'#f59e0b',
    '4':'#f97316',
    '5':'#dc2626'
  })[level]||'#94a3b8'
}
function severityBarChart(sel,data){
  const max=Math.max(1,...data.map(x=>x.count));
  $(sel).innerHTML=data.map(x=>`
    <div class="bar-row severity-row">
      <span class="severity-level" style="border-color:${x.color};color:${x.color}">${esc(x.label)}</span>
      <div class="bar-track">
        <div class="bar-fill severity-fill" style="width:${x.count/max*100}%;background:${x.color}"></div>
      </div>
      <b>${x.count.toLocaleString()}</b>
    </div>
  `).join('')
}

function barChart(sel,data){const max=Math.max(1,...data.map(x=>x.count));$(sel).innerHTML=data.map(x=>`<div class="bar-row"><span>${esc(x.label)}</span><div class="bar-track"><div class="bar-fill" style="width:${x.count/max*100}%"></div></div><b>${x.count.toLocaleString()}</b></div>`).join('')}
function countChart(sel,arr){const m={};arr.forEach(x=>m[x]=(m[x]||0)+1);barChart(sel,Object.entries(m).sort((a,b)=>b[1]-a[1]).map(([label,count])=>({label,count})))}
function rank(sel,arr){const m={};arr.forEach(x=>m[x]=(m[x]||0)+1);$(sel).innerHTML=Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([x,n],i)=>`<div class="rank"><b>${i+1}</b><span>${esc(x)}</span><b>${n.toLocaleString()}</b></div>`).join('')||'<p class="empty">ไม่พบข้อมูล</p>'}
function renderIncidents(){const total=state.filtered.length,max=Math.max(1,Math.ceil(total/state.pageSize));state.page=Math.min(state.page,max);const start=(state.page-1)*state.pageSize;$('#resultCount').textContent=`${total.toLocaleString()} รายการ`;$('#incidentRows').innerHTML=state.filtered.slice(start,start+state.pageSize).map((r,i)=>`<tr><td>${start+i+1}</td><td>${esc(fmtDate(r[IDX.date]))}</td><td>${esc(r[IDX.risk])}</td><td>${esc(r[IDX.sub])}</td><td><span class="badge ${isHigh(r[IDX.sev])?'high':''}">${esc(r[IDX.sev])}</span></td><td>${esc(r[IDX.reportUnit])}</td><td>${esc(normUnit(r[IDX.mainUnit]))}</td><td class="detail-cell">${esc(r[IDX.detail])}</td></tr>`).join('')||'<tr><td colspan="8" class="empty">ไม่พบข้อมูลตามตัวกรอง</td></tr>';const pages=[1,state.page-1,state.page,state.page+1,max].filter((x,i,a)=>x>=1&&x<=max&&a.indexOf(x)===i).sort((a,b)=>a-b);$('#pager').innerHTML=pages.map(p=>`<button class="${p===state.page?'active':''}" data-p="${p}">${p}</button>`).join('');$$('#pager button').forEach(b=>b.onclick=()=>{state.page=+b.dataset.p;renderIncidents()})}
function fmtDate(s){if(!s)return'';const d=new Date(s);return Number.isNaN(d.getTime())?s:d.toLocaleDateString('th-TH')}
function contextText(){return state.selected.size?[...state.selected].join(' • '):'ภาพรวมโรงพยาบาล'}
function codeOf(v){return String(v||'').split(':')[0].trim().split(/\s+/).pop()||''}
function relatedRiskCodes(){return new Set(state.filtered.map(r=>codeOf(r[IDX.risk])).filter(Boolean))}
function matchGeneratedText(obj){if(!state.selected.size)return true;const s=JSON.stringify(obj).toLowerCase();return [...state.selected].some(u=>s.includes(u.toLowerCase()))}
function frequencyScore(n){if(n<=2)return 1;if(n<=20)return 2;if(n<=50)return 3;if(n<=100)return 4;return 5}
function severityImpact(s){const x=String(s||'').trim().toUpperCase();if(['I','H','G','5'].includes(x))return 5;if(['F','E','4'].includes(x))return 4;if(['D','3'].includes(x))return 3;if(['C','2'].includes(x))return 2;return 1}
function riskLevel(score){if(score>=16)return'สูงมาก';if(score>=10)return'สูง';if(score>=5)return'ปานกลาง';return'ต่ำ'}
function riskParts(v){const raw=String(v||'ไม่ระบุ').trim();const m=raw.match(/^([^:]+):\s*(.*)$/s);return m?{code:m[1].trim(),title:m[2].trim()}:{code:codeOf(raw),title:raw}}
function baselineProfileMap(){const m=new Map();Object.values(state.profiles||{}).flat().forEach(x=>{const c=codeOf(x.risk);if(c&&!m.has(c))m.set(c,x)});return m}
function genericControl(type){return type==='Clinical'?'1. ทบทวนแนวทาง/มาตรฐานการดูแลในจุดเสี่ยง\n2. ใช้ checklist หรือ double check ในขั้นตอนสำคัญ\n3. ทบทวนเหตุการณ์และติดตามผลการปรับปรุงในหน่วยงาน':'1. วิเคราะห์สาเหตุและแก้ไขจุดเสี่ยงเชิงระบบ\n2. กำหนดผู้รับผิดชอบและระยะเวลาติดตามชัดเจน\n3. คืนข้อมูลและติดตามผลในคณะกรรมการ/หน่วยงานเจ้าภาพ'}
function generatedProfileRows(){
 const base=baselineProfileMap(), groups=new Map();
 state.filtered.forEach(r=>{const p=riskParts(r[IDX.risk]),type=riskType(r);if(!['Clinical','Non-clinical'].includes(type))return;const key=type+'|'+p.code;let g=groups.get(key);if(!g){g={type,code:p.code,title:p.title,count:0,maxImpact:1};groups.set(key,g)}g.count++;g.maxImpact=Math.max(g.maxImpact,severityImpact(r[IDX.sev]))});
 const rows=[...groups.values()].map(g=>{const b=base.get(g.code),likelihood=frequencyScore(g.count),impact=Number(b?.impact)||g.maxImpact,score=likelihood*impact;return {...g,likelihood,impact,score,level:riskLevel(score),control:b?.control||genericControl(g.type)}});
 const sort=(a,b)=>b.score-a.score||b.count-a.count||a.code.localeCompare(b.code,'th');
 return {clinical:rows.filter(x=>x.type==='Clinical').sort(sort).slice(0,10),non:rows.filter(x=>x.type==='Non-clinical').sort(sort).slice(0,10)}
}
function profileRowHtml(x,i){return `<tr><td class="center">${i+1}</td><td><b>${esc(x.code)}:</b> ${esc(x.title)}<br><small>(พบ ${x.count.toLocaleString()} ครั้ง)</small></td><td class="center score-cell">${x.likelihood}</td><td class="center score-cell">${x.impact}</td><td class="center score-cell"><b>${x.score}</b></td><td class="center"><span class="risk-level level-${x.level}">${esc(x.level)}</span></td><td class="control-cell">${esc(x.control).replaceAll('\n','<br>')}</td></tr>`}
function renderProfile(){const p=generatedProfileRows(),total=p.clinical.length+p.non.length;$('#profileUnitTitle').textContent=`หน่วยงาน: ${contextText()}`;$('#profilePeriod').textContent=`ประจำปีงบประมาณ ${state.year}`;$('#profileContext').textContent=`ใช้ข้อมูล Incident ที่ผ่านตัวกรองปัจจุบัน ${state.filtered.length.toLocaleString()} รายการ • สร้าง Clinical ${p.clinical.length} และ Non-clinical ${p.non.length} รายการ`;$('#profileClinicalRows').innerHTML=p.clinical.map(profileRowHtml).join('')||'<tr><td colspan="7" class="empty">ไม่พบความเสี่ยงด้านคลินิก</td></tr>';$('#profileNonRows').innerHTML=p.non.map(profileRowHtml).join('')||'<tr><td colspan="7" class="empty">ไม่พบความเสี่ยงด้านทั่วไป</td></tr>'}
function getProfileRows(){const p=generatedProfileRows();return [...p.clinical,...p.non].map((x,i)=>({rank:i+1,type:x.type,risk:`${x.code}: ${x.title} (พบ ${x.count} ครั้ง)`,likelihood:x.likelihood,impact:x.impact,score:x.score,level:x.level,control:x.control}))}

function contextKey(){return `${state.year}|${[...state.selected].sort().join('||')||'ALL'}`}
function draftStoreKey(kind){return `drms_${kind}_draft_${contextKey()}`}
function historyStoreKey(kind){return `drms_${kind}_history_${contextKey()}`}
function readStore(key,fallback={}){try{return JSON.parse(localStorage.getItem(key)||JSON.stringify(fallback))}catch(e){return fallback}}
function setProfileMode(mode){state.profileMode=mode;$('#profileSummaryPane').classList.toggle('hidden',mode!=='summary');$('#profileHaPane').classList.toggle('hidden',mode!=='ha');$('#profileSummaryMode').classList.toggle('active',mode==='summary');$('#profileHaMode').classList.toggle('active',mode==='ha');if(mode==='ha')renderProfileHa()}
function setRegisterMode(mode){state.registerMode=mode;$('#registerSummaryPane').classList.toggle('hidden',mode!=='summary');$('#registerHaPane').classList.toggle('hidden',mode!=='ha');$('#registerSummaryMode').classList.toggle('active',mode==='summary');$('#registerHaMode').classList.toggle('active',mode==='ha');if(mode==='ha')renderRegisterHa()}
function editableCell(kind,code,field,value){return `<td class="editable-cell" contenteditable="true" data-kind="${kind}" data-code="${esc(code)}" data-field="${field}" spellcheck="false">${esc(value||'')}</td>`}
function profileDefaults(x){return {objective:`ลดการเกิด ${x.code} และลดผลกระทบต่อผู้รับบริการ/บุคลากร`,indicator:`จำนวนอุบัติการณ์ ${x.code} และร้อยละการดำเนินการตามมาตรการ`,target:'ไม่เกิดเหตุระดับรุนแรง และแนวโน้มลดลงจากรอบก่อน',additional:'ทบทวนสาเหตุ กำหนด strong action และสื่อสารมาตรการในหน่วยงาน',owner:state.selected.size?[...state.selected].join(', '):'หน่วยงานหลักที่แก้ไข',deadline:`ภายในปีงบประมาณ ${state.year}`,followup:'รอติดตามผล'}}
function renderProfileHa(){if(!$('#profileHaRows'))return;const p=generatedProfileRows(),rows=[...p.clinical,...p.non],drafts=readStore(draftStoreKey('profile'),{});$('#profileHaContext').textContent=`ปีงบประมาณ ${state.year} • ${contextText()} • ${rows.length} ประเด็นความเสี่ยง`;$('#profileHaRows').innerHTML=rows.map((x,i)=>{const d={...profileDefaults(x),...(drafts[x.code]||{})};return `<tr><td class="center">${i+1}</td><td>${esc(x.type)}</td><td><b>${esc(x.code)}:</b> ${esc(x.title)}<br><small>พบ ${x.count.toLocaleString()} ครั้ง</small></td><td class="center">${x.likelihood}</td><td class="center">${x.impact}</td><td class="center"><b>${x.score}</b><br><span class="risk-level level-${x.level}">${esc(x.level)}</span></td>${editableCell('profile',x.code,'objective',d.objective)}${editableCell('profile',x.code,'indicator',d.indicator)}${editableCell('profile',x.code,'target',d.target)}<td>${esc(x.control).replaceAll('\n','<br>')}</td>${editableCell('profile',x.code,'additional',d.additional)}${editableCell('profile',x.code,'owner',d.owner)}${editableCell('profile',x.code,'deadline',d.deadline)}${editableCell('profile',x.code,'followup',d.followup)}</tr>`}).join('')||'<tr><td colspan="14" class="empty">ไม่พบข้อมูล</td></tr>';renderHistory('profile')}
function registerDefaults(x){return {prevention:x.prevention,indicator:x.monitor,target:'ผลการติดตามเป็นไปตามเป้าหมายและไม่เกิดเหตุรุนแรงซ้ำ',owner:x.owner,deadline:`Q1-Q4 ปีงบประมาณ ${state.year}`,followup:'รอติดตามผล',evidence:'รายงานประชุม/ภาพถ่าย/แนวทาง/ผล audit'}}
function renderRegisterHa(){if(!$('#registerHaRows'))return;const rows=generatedRegisterRows(),drafts=readStore(draftStoreKey('register'),{});$('#registerHaContext').textContent=`ปีงบประมาณ ${state.year} • ${contextText()} • ${rows.length} รายการ`;$('#registerHaRows').innerHTML=rows.map((x,i)=>{const code=codeOf(x.title)||`RR-${i+1}`,d={...registerDefaults(x),...(drafts[code]||{})};return `<tr><td>${esc(x.riskId).replaceAll('\n','<br>')}</td><td><b>${esc(x.title)}</b></td><td>${esc(x.description).replaceAll('\n','<br>')}</td><td class="center">${x.likelihood}</td><td class="center">${x.impact}</td><td class="center"><span class="risk-level level-${riskLevel(Number(x.likelihood)*Number(x.impact))}">${esc(x.level).replaceAll('\n','<br>')}</span></td>${editableCell('register',code,'prevention',d.prevention)}${editableCell('register',code,'indicator',d.indicator)}${editableCell('register',code,'target',d.target)}${editableCell('register',code,'owner',d.owner)}${editableCell('register',code,'deadline',d.deadline)}${editableCell('register',code,'followup',d.followup)}${editableCell('register',code,'evidence',d.evidence)}</tr>`}).join('')||'<tr><td colspan="13" class="empty">ไม่พบข้อมูล</td></tr>';renderHistory('register')}
function saveHaDrafts(kind){const data=readStore(draftStoreKey(kind),{});$$(`[data-kind="${kind}"]`).forEach(cell=>{const code=cell.dataset.code,field=cell.dataset.field;data[code]=data[code]||{};data[code][field]=cell.innerText.trim()});const now=new Date().toLocaleString('th-TH');data.__meta={updatedAt:now,year:state.year,units:contextText()};localStorage.setItem(draftStoreKey(kind),JSON.stringify(data));const hist=readStore(historyStoreKey(kind),[]);hist.unshift({at:now,year:state.year,units:contextText(),count:Object.keys(data).filter(k=>k!=='__meta').length});localStorage.setItem(historyStoreKey(kind),JSON.stringify(hist.slice(0,20)));renderHistory(kind);toast(`บันทึกร่าง ${kind==='profile'?'Risk Profile':'Risk Register'} แล้ว`)}
function renderHistory(kind){const data=readStore(draftStoreKey(kind),{}),hist=readStore(historyStoreKey(kind),[]),status=$(`#${kind}DraftStatus`),box=$(`#${kind}History`);if(status)status.textContent=data.__meta?.updatedAt?`บันทึกล่าสุด ${data.__meta.updatedAt}`:'ยังไม่บันทึกร่าง';if(box)box.innerHTML=hist.length?hist.map(x=>`<div class="history-item"><b>${esc(x.at)}</b><span>${esc(x.units)} • ${x.count} รายการ</span></div>`).join(''):'<p class="empty">ยังไม่มีประวัติการบันทึก</p>'}
function documentHtml(title,tableHtml){return `<!doctype html><html><head><meta charset="utf-8"><title>${esc(title)}</title><style>body{font-family:Tahoma,Arial,sans-serif;font-size:11pt}h1,h2{text-align:center}table{border-collapse:collapse;width:100%}th,td{border:1px solid #333;padding:5px;vertical-align:top}th{background:#dbe7f5}.risk-level{font-weight:bold}</style></head><body><h2>${esc(title)}</h2><p>โรงพยาบาลดอนตูม • ปีงบประมาณ ${state.year} • ${esc(contextText())}</p>${tableHtml}</body></html>`}
function exportTableDocument(kind,format){const isProfile=kind==='profile',mode=isProfile?state.profileMode:state.registerMode,tableId=isProfile?(mode==='ha'?'profileHaTable':null):(mode==='ha'?'registerHaTable':null);let table;if(tableId)table=document.getElementById(tableId);else{const pane=document.getElementById(isProfile?'profileSummaryPane':'registerSummaryPane');table=pane?.querySelector('table')}if(!table){toast('ไม่พบตารางสำหรับส่งออก');return}const title=isProfile?'Risk Profile':'Risk Register',html='\ufeff'+documentHtml(title,table.outerHTML);if(format==='word')download(html,`${title.replace(' ','_')}_${state.year}.doc`,'application/msword');else download(html,`${title.replace(' ','_')}_${state.year}.xls`,'application/vnd.ms-excel')}

function exportProfileCsv(){const p=generatedProfileRows(),rows=[...p.clinical,...p.non];const lines=[['บัญชีรายการความเสี่ยง (Risk Profile)'],['ปีงบประมาณ',state.year],['หน่วยงาน',contextText()],[],['ประเภท','อันดับ','รหัส','รายการความเสี่ยง','จำนวนครั้ง','โอกาส','ผลกระทบ','ผลคูณ','ระดับความเสี่ยง','มาตรการควบคุม/ป้องกัน/แก้ไข'],...rows.map((x,i)=>[x.type,i+1,x.code,x.title,x.count,x.likelihood,x.impact,x.score,x.level,x.control])];const csv='\ufeff'+lines.map(row=>row.map(v=>'"'+String(v??'').replaceAll('"','""')+'"').join(',')).join('\n');download(csv,`Risk_Profile_${state.year}.csv`,'text/csv')}
function baselineRegisterMap(){const m=new Map();Object.values(state.registers||{}).flat().forEach(x=>{const c=codeOf(x.title)||codeOf(x.riskId);if(c&&!m.has(c))m.set(c,x)});return m}
function generatedRegisterRows(){const p=generatedProfileRows(), base=baselineRegisterMap(), rows=[...p.clinical,...p.non];return rows.map((x,i)=>{const b=base.get(x.code)||{};const owner=state.selected.size?[...state.selected].join(', '):'หน่วยงานหลักที่แก้ไขตามข้อมูล Incident';return {riskId:`RR${String(state.year).slice(-2)}-${String(i+1).padStart(2,'0')}\n${x.code}`,title:`${x.code}: ${x.title}`,description:`พบ ${x.count.toLocaleString()} ครั้ง ในปีงบประมาณ ${state.year}\nประเภท: ${x.type}\nหน่วยงานเจ้าภาพ: ${owner}`,quarter:b.quarter||'Q1-Q4',likelihood:x.likelihood,impact:x.impact,level:`${x.score}\n(${x.level})`,prevention:b.prevention||x.control,monitor:b.monitor||`ติดตามจำนวนอุบัติการณ์ของ ${x.code} รายเดือน/รายไตรมาส วิเคราะห์แนวโน้ม และรายงานในที่ประชุมหน่วยงาน/RM`,mitigation:b.mitigation||'ควบคุมความเสียหายทันที แจ้งหัวหน้างานและทีม RM ทบทวนสาเหตุ กำหนด strong action และติดตามผล',owner};})}
function renderRegister(){const rows=generatedRegisterRows();$('#registerContext').textContent=`ปีงบประมาณ ${state.year} • ${contextText()} • สร้างจาก Incident ${state.filtered.length.toLocaleString()} รายการ • Risk Register ${rows.length.toLocaleString()} รายการ`;$('#registerRows').innerHTML=rows.map((x,i)=>`<tr><td class="center">${i+1}<br><small>${esc(x.riskId).replaceAll('\n','<br>')}</small></td><td><b>${esc(x.title)}</b></td><td>${esc(x.description).replaceAll('\n','<br>')}</td><td class="center">${esc(x.quarter)}</td><td class="center score-cell">${esc(x.likelihood)}</td><td class="center score-cell">${esc(x.impact)}</td><td class="center"><span class="risk-level level-${riskLevel(Number(x.likelihood)*Number(x.impact))}">${esc(x.level).replaceAll('\n','<br>')}</span></td><td class="control-cell">${esc(x.prevention).replaceAll('\n','<br>')}</td><td class="control-cell">${esc(x.monitor).replaceAll('\n','<br>')}</td><td class="control-cell">${esc(x.mitigation).replaceAll('\n','<br>')}</td><td>${esc(x.owner)}</td></tr>`).join('')||'<tr><td colspan="11" class="empty">ไม่พบข้อมูลสำหรับสร้าง Risk Register</td></tr>'}
function getRegisterRows(){return generatedRegisterRows()}
function exportRegisterCsv(){const rows=generatedRegisterRows();const lines=[['ทะเบียนความเสี่ยง (Risk Register)'],['ปีงบประมาณ',state.year],['หน่วยงาน',contextText()],[],['ลำดับ','รหัส','ชื่อความเสี่ยง','รายละเอียดและหลักฐาน','ช่วงเวลา','L','I','ระดับ','การป้องกัน/ควบคุม','ตัวชี้วัดและการติดตาม','การลดผลกระทบ','ผู้รับผิดชอบ'],...rows.map((x,i)=>[i+1,x.riskId,x.title,x.description,x.quarter,x.likelihood,x.impact,x.level,x.prevention,x.monitor,x.mitigation,x.owner])];const csv='\ufeff'+lines.map(row=>row.map(v=>'"'+String(v??'').replaceAll('"','""')+'"').join(',')).join('\n');download(csv,`Risk_Register_${state.year}.csv`,'text/csv')}
function exportCsv(){const h=['ปีงบประมาณ','วันที่เกิด','รหัส/เรื่อง','เรื่องย่อย','ความรุนแรง','หน่วยงานรายงาน','หน่วยงานหลักที่แก้ไข','รายละเอียด'];const body=state.filtered.map(r=>[state.year,r[IDX.date],r[IDX.risk],r[IDX.sub],r[IDX.sev],r[IDX.reportUnit],normUnit(r[IDX.mainUnit]),r[IDX.detail]]);const csv='\ufeff'+[h,...body].map(row=>row.map(v=>'"'+String(v??'').replaceAll('"','""')+'"').join(',')).join('\n');download(csv,`drms_incidents_${state.year}.csv`,'text/csv')}
function download(data,name,type){const a=document.createElement('a');const url=URL.createObjectURL(new Blob([data],{type}));a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),500)}

function isNearMiss(s){return ['A','B','1','2'].includes(String(s||'').trim().toUpperCase())}
function countMap(arr){const m={};arr.forEach(x=>m[x]=(m[x]||0)+1);return m}

function currentFilterValues(){return {sev:$('#sev')?.value||'',type:$('#type')?.value||'',month:+($('#month')?.value||0)}}
function filterComparableRows(rows){const f=currentFilterValues();return (rows||[]).filter(r=>(!state.selected.size||state.selected.has(normUnit(r[IDX.mainUnit])))&&(!f.sev||String(r[IDX.sev]||'')===f.sev)&&(!f.type||riskType(r)===f.type)&&(!f.month||monthOf(r[IDX.date])===f.month))}
async function ensureYearRows(year){if(state.allYearRows[year])return state.allYearRows[year];try{const d=await json(PARENT+`incidents_${year}.json`);state.allYearRows[year]=d.rows||[];return state.allYearRows[year]}catch(e){console.warn('load trend year failed',year,e);return[]}}
function renderRiskMatrix(){
  const rows=getProfileRows();const cells={};rows.forEach(x=>{const l=Number(x.likelihood)||1,i=Number(x.impact)||1;cells[`${i}-${l}`]=(cells[`${i}-${l}`]||0)+1});
  let out='<div class="rm-axis-title">ผลกระทบ (Impact)</div><div class="rm-grid"><div class="rm-corner"></div>';
  for(let l=1;l<=5;l++)out+=`<div class="rm-head">L${l}</div>`;
  for(let i=5;i>=1;i--){out+=`<div class="rm-head rm-side">I${i}</div>`;for(let l=1;l<=5;l++){const score=i*l,n=cells[`${i}-${l}`]||0;const cls=score>=16?'rm-vhigh':score>=10?'rm-high':score>=5?'rm-med':'rm-low';out+=`<div class="rm-cell ${cls}" title="Impact ${i} × Likelihood ${l} = ${score}"><b>${n}</b><small>${score}</small></div>`}}out+='</div><div class="rm-x-title">โอกาสเกิด (Likelihood)</div><div class="rm-legend"><span class="rm-low">ต่ำ</span><span class="rm-med">ปานกลาง</span><span class="rm-high">สูง</span><span class="rm-vhigh">สูงมาก</span></div>';
  $('#riskMatrix').innerHTML=out;
}
async function renderYearTrend(){
  const years=state.years.length?state.years:[2566,2567,2568,2569];
  const values=[];for(const y of years){const rows=await ensureYearRows(y);values.push({label:String(y),count:filterComparableRows(rows).length})}
  if($('#yearTrend'))barChart('#yearTrend',values);
}
function renderReadiness(){
  const a=state.filtered,profile=getProfileRows(),register=getRegisterRows(),monthCount=new Set(a.map(r=>monthOf(r[IDX.date])).filter(Boolean)).size;
  const checks=[
    {ok:a.length>0,title:'มีข้อมูล Incident',detail:`${a.length.toLocaleString()} รายการ`},
    {ok:state.selected.size>0,title:'ระบุหน่วยงานเจ้าภาพ',detail:state.selected.size?`${state.selected.size} หน่วยงาน`:'ยังเป็นภาพรวมโรงพยาบาล'},
    {ok:profile.length>0,title:'สร้าง Risk Profile ได้',detail:`${profile.length} ประเด็น`},
    {ok:register.length>0,title:'สร้าง Risk Register ได้',detail:`${register.length} รายการ`},
    {ok:monthCount>=3,title:'มีข้อมูลแนวโน้มตามเวลา',detail:`ครอบคลุม ${monthCount} เดือนงบประมาณ`},
    {ok:a.filter(r=>isHigh(r[IDX.sev])).length===0||register.length>0,title:'High Risk มีแผนติดตาม',detail:`High Risk ${a.filter(r=>isHigh(r[IDX.sev])).length.toLocaleString()} รายการ`}
  ];
  $('#haReadiness').innerHTML=checks.map(x=>`<div class="readiness-item ${x.ok?'ready':'pending'}"><span>${x.ok?'✓':'!'}</span><div><b>${esc(x.title)}</b><small>${esc(x.detail)}</small></div></div>`).join('');
}
function buildHaNarrative(){
  const a=state.filtered,total=a.length,clinical=a.filter(r=>riskType(r)==='Clinical').length,non=a.filter(r=>riskType(r)==='Non-clinical').length,high=a.filter(r=>isHigh(r[IDX.sev])).length;
  const top=Object.entries(countMap(a.map(r=>String(r[IDX.risk]||'ไม่ระบุ').split(':')[0].trim()))).sort((x,y)=>y[1]-x[1]).slice(0,3);
  const months=[10,11,12,1,2,3,4,5,6,7,8,9],names={10:'ตุลาคม',11:'พฤศจิกายน',12:'ธันวาคม',1:'มกราคม',2:'กุมภาพันธ์',3:'มีนาคม',4:'เมษายน',5:'พฤษภาคม',6:'มิถุนายน',7:'กรกฎาคม',8:'สิงหาคม',9:'กันยายน'};
  const peak=months.map(m=>({m,n:a.filter(r=>monthOf(r[IDX.date])===m).length})).sort((x,y)=>y.n-x.n)[0]||{m:'',n:0};
  const topText=top.length?top.map(([x,n])=>`${x} (${n.toLocaleString()} ครั้ง)`).join(', '):'ยังไม่มีข้อมูล';
  return `<p>ปีงบประมาณ <b>${state.year}</b> ${esc(contextText())} มีอุบัติการณ์ทั้งหมด <b>${total.toLocaleString()}</b> รายการ แบ่งเป็น Clinical <b>${clinical.toLocaleString()}</b> รายการ และ Non-clinical <b>${non.toLocaleString()}</b> รายการ โดยพบ High Risk <b>${high.toLocaleString()}</b> รายการ</p><p>ประเด็นความเสี่ยงที่พบสูงสุด ได้แก่ <b>${esc(topText)}</b>${peak.n?` และเดือนที่พบเหตุการณ์สูงสุดคือ <b>${names[peak.m]}</b> จำนวน <b>${peak.n.toLocaleString()}</b> รายการ`:''}</p><p>หน่วยงานควรทบทวน Risk Profile และ Risk Register ที่ระบบสร้างเป็นร่าง ตรวจสอบมาตรการควบคุม ตัวชี้วัด ผู้รับผิดชอบ และหลักฐานการติดตามก่อนใช้ในการประเมิน HA</p>`;
}
function renderAnalytics(){
  const a=state.filtered;
  $('#aTotal').textContent=a.length.toLocaleString();
  $('#aHigh').textContent=a.filter(r=>isHigh(r[IDX.sev])).length.toLocaleString();
  $('#aNear').textContent=a.filter(r=>isNearMiss(r[IDX.sev])).length.toLocaleString();
  $('#aUnits').textContent=new Set(a.map(r=>normUnit(r[IDX.mainUnit]))).size.toLocaleString();
  const risks=Object.entries(countMap(a.map(r=>String(r[IDX.risk]||'ไม่ระบุ').split(':')[0].trim()))).sort((x,y)=>y[1]-x[1]).slice(0,12);
  const total=Math.max(1,risks.reduce((n,x)=>n+x[1],0));let cum=0;
  $('#paretoRisks').innerHTML=risks.map(([name,n],i)=>{cum+=n;return `<div class="pareto-row"><b>${i+1}</b><span>${esc(name)}</span><div class="pareto-bar"><i style="width:${n/risks[0][1]*100}%"></i></div><strong>${n.toLocaleString()}</strong><small>${(cum/total*100).toFixed(1)}%</small></div>`}).join('')||'<p class="empty">ไม่พบข้อมูล</p>';
  const clinical=a.filter(r=>riskType(r)==='Clinical').length, non=a.filter(r=>riskType(r)==='Non-clinical').length, all=Math.max(1,a.length);
  $('#typeSummary').innerHTML=`<div class="donut" style="--p:${clinical/all*100}"><span>${(clinical/all*100).toFixed(1)}%</span></div><div class="legend"><p><i class="dot c"></i>Clinical <b>${clinical.toLocaleString()}</b></p><p><i class="dot n"></i>Non-clinical <b>${non.toLocaleString()}</b></p></div>`;
  const sevs=[...new Set(a.map(r=>String(r[IDX.sev]||'ไม่ระบุ')))].sort((x,y)=>x.localeCompare(y,'th'));
  $('#severityByType').innerHTML=`<table><thead><tr><th>ระดับ</th><th>Clinical</th><th>Non-clinical</th><th>รวม</th></tr></thead><tbody>${sevs.map(s=>{const c=a.filter(r=>String(r[IDX.sev]||'ไม่ระบุ')===s&&riskType(r)==='Clinical').length;const n=a.filter(r=>String(r[IDX.sev]||'ไม่ระบุ')===s&&riskType(r)==='Non-clinical').length;return `<tr><td><span class="badge ${isHigh(s)?'high':''}">${esc(s)}</span></td><td>${c.toLocaleString()}</td><td>${n.toLocaleString()}</td><td><b>${(c+n).toLocaleString()}</b></td></tr>`}).join('')}</tbody></table>`;
  const months=[10,11,12,1,2,3,4,5,6,7,8,9],names=['ต.ค.','พ.ย.','ธ.ค.','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.'];
  barChart('#peakMonths',months.map((m,i)=>({label:names[i],count:a.filter(r=>monthOf(r[IDX.date])===m).length})).sort((x,y)=>y.count-x.count));
  renderRiskMatrix();renderReadiness();renderYearTrend();
}
function haReportKey(){return `drms_ha_report_${contextKey()}`}
function haReportHistoryKey(){return `drms_ha_report_history_${contextKey()}`}
function defaultActionRows(){
  const rr=getRegisterRows().slice(0,5);
  return rr.map((x,i)=>({
    issue:x.title||`ประเด็นพัฒนา ${i+1}`,
    action:x.prevention||'ทบทวนกระบวนการและกำหนดมาตรการป้องกันเชิงระบบ',
    indicator:x.monitor||'ติดตามจำนวนอุบัติการณ์และผลการดำเนินงานรายเดือน',
    target:'แนวโน้มลดลงและไม่เกิดเหตุรุนแรงซ้ำ',
    owner:x.owner||contextText(),
    deadline:`ภายในปีงบประมาณ ${state.year}`,
    followup:'รอติดตามผล'
  }))
}
function editableActionCell(row,field,value){return `<td class="editable-cell" contenteditable="true" data-ha-row="${row}" data-ha-field="${field}" spellcheck="false">${esc(value||'')}</td>`}
function renderHaActionRows(saved={}){
  const rows=(saved.actions&&saved.actions.length?saved.actions:defaultActionRows());
  $('#haActionRows').innerHTML=rows.map((x,i)=>`<tr>${editableActionCell(i,'issue',x.issue)}${editableActionCell(i,'action',x.action)}${editableActionCell(i,'indicator',x.indicator)}${editableActionCell(i,'target',x.target)}${editableActionCell(i,'owner',x.owner)}${editableActionCell(i,'deadline',x.deadline)}${editableActionCell(i,'followup',x.followup)}</tr>`).join('')||'<tr><td colspan="7" class="empty">ไม่พบข้อมูลสำหรับสร้างแผนพัฒนา</td></tr>'
}
function renderHaEvidence(saved={}){
  const profileDraft=readStore(draftStoreKey('profile'),{}),registerDraft=readStore(draftStoreKey('register'),{});
  const items=[
    ['incident','รายการ Incident และหลักฐานการทบทวน',state.filtered.length>0],
    ['profile','Risk Profile ที่หน่วยงานทบทวนแล้ว',Object.keys(profileDraft).some(k=>k!=='__meta')],
    ['register','Risk Register ที่กำหนดมาตรการและผู้รับผิดชอบแล้ว',Object.keys(registerDraft).some(k=>k!=='__meta')],
    ['minutes','รายงานประชุม/การทบทวนของหน่วยงาน',false],
    ['rca','RCA / mini RCA สำหรับเหตุการณ์สำคัญ',state.filtered.filter(r=>isHigh(r[IDX.sev])).length===0],
    ['audit','ผล audit / tracer / การติดตามตัวชี้วัด',false],
    ['communication','หลักฐานการสื่อสารมาตรการแก่ผู้เกี่ยวข้อง',false],
    ['outcome','หลักฐานผลลัพธ์หลังดำเนินการ',false]
  ];
  const checked=new Set(saved.evidence||[]);
  $('#haEvidenceChecklist').innerHTML=items.map(([id,label,auto])=>`<label class="evidence-item"><input type="checkbox" data-evidence="${id}" ${(checked.has(id)||auto)?'checked':''}><span><b>${esc(label)}</b><small>${auto?'ระบบพบข้อมูลเบื้องต้นแล้ว':'ให้หน่วยงานตรวจสอบและแนบหลักฐาน'}</small></span></label>`).join('')
}
function readHaDraft(){return readStore(haReportKey(),{})}
function renderHaReport(force=false){
  const a=state.filtered, pr=getProfileRows(), rr=getRegisterRows(),saved=force?{}:readHaDraft();
  $('#haContext').textContent=`ปีงบประมาณ ${state.year} • ${contextText()} • จัดทำจากข้อมูลที่ผ่านตัวกรองปัจจุบัน`;
  $('#haNarrative').innerHTML=saved.narrative||buildHaNarrative();
  $('#haStrengths').innerHTML=saved.strengths||'ระบุจุดเด่นของระบบบริหารความเสี่ยง ผลลัพธ์ที่ดีขึ้น และแนวปฏิบัติที่หน่วยงานภาคภูมิใจ';
  $('#haOpportunities').innerHTML=saved.opportunities||'ระบุประเด็นที่ต้องปรับปรุง ความเสี่ยงซ้ำ แนวโน้มที่ต้องเฝ้าระวัง และช่องว่างของมาตรการควบคุม';
  $('#haTotal').textContent=a.length.toLocaleString();$('#haClinical').textContent=a.filter(r=>riskType(r)==='Clinical').length.toLocaleString();$('#haNon').textContent=a.filter(r=>riskType(r)==='Non-clinical').length.toLocaleString();$('#haHigh').textContent=a.filter(r=>isHigh(r[IDX.sev])).length.toLocaleString();
  rank('#haTopRisks',a.map(r=>String(r[IDX.risk]||'ไม่ระบุ').split(':')[0].trim()));
  const months=[10,11,12,1,2,3,4,5,6,7,8,9],names=['ต.ค.','พ.ย.','ธ.ค.','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.'];barChart('#haMonthly',months.map((m,i)=>({label:names[i],count:a.filter(r=>monthOf(r[IDX.date])===m).length})));
  $('#haProfileRows').innerHTML=pr.map(x=>`<tr><td>${esc(x.rank)}</td><td>${esc(x.type)}</td><td>${esc(x.risk)}</td><td>${esc(x.likelihood)}</td><td>${esc(x.impact)}</td><td>${esc(x.score)}</td><td><span class="level">${esc(x.level)}</span></td><td>${esc(x.control)}</td></tr>`).join('')||'<tr><td colspan="8" class="empty">ไม่พบข้อมูล</td></tr>';
  $('#haRegisterRows').innerHTML=rr.map(x=>`<tr><td>${esc(x.riskId)}</td><td>${esc(x.title)}</td><td>${esc(x.description)}</td><td>${esc(x.quarter)}</td><td>${esc(x.likelihood)}</td><td>${esc(x.impact)}</td><td><span class="level">${esc(x.level)}</span></td><td>${esc(x.prevention)}</td><td>${esc(x.monitor)}</td><td>${esc(x.mitigation)}</td></tr>`).join('')||'<tr><td colspan="10" class="empty">ไม่พบข้อมูล</td></tr>';
  renderHaActionRows(saved);renderHaEvidence(saved);renderHaReportHistory()
}
function collectHaActions(){
  const out={};$$('[data-ha-row]').forEach(cell=>{const i=+cell.dataset.haRow,field=cell.dataset.haField;out[i]=out[i]||{};out[i][field]=cell.innerText.trim()});return Object.keys(out).sort((a,b)=>a-b).map(k=>out[k])
}
function saveHaReportDraft(){
  const evidence=$$('[data-evidence]:checked').map(x=>x.dataset.evidence);
  const now=new Date().toLocaleString('th-TH');
  const data={narrative:$('#haNarrative').innerHTML,strengths:$('#haStrengths').innerHTML,opportunities:$('#haOpportunities').innerHTML,actions:collectHaActions(),evidence,__meta:{updatedAt:now,year:state.year,units:contextText()}};
  localStorage.setItem(haReportKey(),JSON.stringify(data));
  const hist=readStore(haReportHistoryKey(),[]);hist.unshift({at:now,year:state.year,units:contextText(),actions:data.actions.length,evidence:evidence.length});localStorage.setItem(haReportHistoryKey(),JSON.stringify(hist.slice(0,20)));renderHaReportHistory();toast('บันทึกร่างรายงาน HA แล้ว')
}
function renderHaReportHistory(){
  const saved=readHaDraft(),hist=readStore(haReportHistoryKey(),[]);
  $('#haDraftStatus').textContent=saved.__meta?.updatedAt?`บันทึกล่าสุด ${saved.__meta.updatedAt}`:'ยังไม่บันทึกร่างรายงาน';
  $('#haReportHistory').innerHTML=hist.length?hist.map(x=>`<div class="history-item"><b>${esc(x.at)}</b><span>${esc(x.units)} • แผน ${x.actions} รายการ • หลักฐาน ${x.evidence} รายการ</span></div>`).join(''):'<p class="empty">ยังไม่มีประวัติการบันทึก</p>'
}
function haReportDocument(){
  const evidence=$$('[data-evidence]').map(x=>`<li>${x.checked?'☑':'☐'} ${esc(x.closest('label').querySelector('b').innerText)}</li>`).join('');
  return `<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Tahoma,Arial,sans-serif;font-size:11pt}h1,h2,h3{text-align:center}table{border-collapse:collapse;width:100%;margin:10px 0}th,td{border:1px solid #333;padding:5px;vertical-align:top}th{background:#dbe7f5}.section{margin:20px 0}.kpis{display:flex;gap:10px}.kpis div{border:1px solid #333;padding:8px}.bars,.actions{display:none}</style></head><body><h2>รายงานบริหารความเสี่ยงเพื่อการพัฒนาคุณภาพและการประเมิน HA</h2><p><b>โรงพยาบาลดอนตูม</b><br>ปีงบประมาณ ${state.year}<br>${esc(contextText())}</p><div class="section"><h3>1. บริบทและภาพรวม</h3>${$('#haNarrative').innerHTML}</div><div class="section"><h3>2. จุดเด่นและผลลัพธ์สำคัญ</h3>${$('#haStrengths').innerHTML}</div><div class="section"><h3>3. ประเด็นพัฒนา</h3>${$('#haOpportunities').innerHTML}</div><div class="section"><h3>4. แผนพัฒนาและการติดตาม</h3>${$('#haActionTable').outerHTML}</div><div class="section"><h3>5. Checklist หลักฐาน</h3><ul>${evidence}</ul></div><div class="section"><h3>6. Risk Profile</h3>${$('#haProfileTable').outerHTML}</div><div class="section"><h3>7. Risk Register</h3>${$('#haRegisterTable').outerHTML}</div></body></html>`
}
function exportHaReport(format){
  const html='\ufeff'+haReportDocument();
  if(format==='word')download(html,`HA_Report_${state.year}.doc`,'application/msword');
  else download(html,`HA_Report_${state.year}.xls`,'application/vnd.ms-excel')
}

function publicRcaSeverity(v){return ['E','F','G','H','I','3','4','5'].includes(String(v??'').trim().toUpperCase())}
function publicIncidentId(r,index){
  const id=String(r[IDX.id]??'').trim();
  return id?id.replace(/[^\w\-]/g,'_'):`INC_${state.year}_${String(index+1).padStart(5,'0')}`
}
function publicRcaEntry(id){
  return (state.rcaManifest||[]).find(x=>String(x.incident)===String(id)&&Number(x.year)===Number(state.year))
}
function renderPublicRca(){
  if(!$('#pubRcaRows'))return;
  const rows=state.filtered.map((r,i)=>({r,id:publicIncidentId(r,i)})).filter(x=>publicRcaSeverity(x.r[IDX.sev]));
  const has=rows.filter(x=>publicRcaEntry(x.id)).length;
  $('#pubRcaRequired').textContent=rows.length.toLocaleString();
  $('#pubRcaHas').textContent=has.toLocaleString();
  $('#pubRcaMissing').textContent=(rows.length-has).toLocaleString();
  $('#pubRcaRows').innerHTML=rows.map(x=>{
    const e=publicRcaEntry(x.id);
    return `<tr><td>${esc(x.r[IDX.date]||'')}</td><td><b>${esc(x.id)}</b></td>
      <td>${esc(x.r[IDX.risk]||'')}</td><td>${esc(x.r[IDX.mainUnit]||'')}</td>
      <td><span class="severity-pill">${esc(x.r[IDX.sev]||'')}</span></td>
      <td>${e?'<span class="rca-has">✓ มี RCA</span>':'<span class="rca-missing">✕ ไม่มี RCA</span>'}</td>
      <td>${e?`<a class="download-link" href="${esc(e.file)}" target="_blank" rel="noopener">ดาวน์โหลด</a>`:'–'}</td></tr>`
  }).join('')||'<tr><td colspan="7" class="empty">ไม่พบข้อมูลระดับ E–I หรือ 3–5</td></tr>'
}


function getIncidentCode(row){
  const text=[row[IDX.risk],row[IDX.sub],row[IDX.keyword]].map(x=>String(x||'').toUpperCase()).join(' ');
  const match=text.match(/\b[A-Z]{3}\d{3}\b/);
  return match?match[0]:''
}
function essentialCodeList(standard){return (standard?.codes||[]).map(x=>typeof x==='string'?x:x.code)}
function essentialRowsByStandard(standard){
  const allowed=new Set(essentialCodeList(standard));
  return state.filtered.filter(row=>allowed.has(getIncidentCode(row)))
}
function severityRank(value){
  const s=String(value||'').trim().toUpperCase();
  const clinical=['A','B','C','D','E','F','G','H','I'];
  const general=['1','2','3','4','5'];
  const ci=clinical.indexOf(s);if(ci>=0)return ci;
  const gi=general.indexOf(s);if(gi>=0)return gi+10;
  return -1
}
function maximumSeverity(rows){
  let best='',score=-1;
  rows.forEach(r=>{const s=String(r[IDX.sev]||'').trim().toUpperCase(),n=severityRank(s);if(n>score){score=n;best=s}});
  return best
}
function fiscalYearForRow(row){return Number(row.__fiscalYear||state.year)||0}
function hasRcaForRow(row,index){
  const id=publicIncidentId(row,index),year=fiscalYearForRow(row);
  return (state.rcaManifest||[]).some(x=>String(x.incident)===String(id)&&Number(x.year)===year)
}
function renderEssentialStandards(){
  const cards=$('#essentialCards');if(!cards)return;
  const standards=state.essentialStandards||[];
  cards.innerHTML=standards.map((standard,index)=>{
    const rows=essentialRowsByStandard(standard),high=rows.filter(r=>isHigh(r[IDX.sev])).length;
    const active=state.selectedEssential?.id===standard.id?' active':'';
    return `<button class="essential-card${active}" data-essential="${esc(standard.id)}">
      <span class="essential-number">${standard.number||index+1}</span>
      <div class="essential-card-text"><b>${esc(standard.shortTitle)}</b><small>${esc(standard.title)}</small><span>${rows.length.toLocaleString()} รายการ</span></div>
      <strong class="${high?'essential-high':'essential-safe'}">${high?`${high.toLocaleString()} High Risk`:'ไม่พบ High Risk'}</strong>
    </button>`
  }).join('')||'<p class="empty">ไม่พบข้อมูลมาตรฐาน</p>';
  $$('.essential-card').forEach(button=>button.onclick=()=>{
    state.selectedEssential=standards.find(x=>x.id===button.dataset.essential)||null;
    renderEssentialStandards()
  });
  if(!state.selectedEssential&&standards.length)state.selectedEssential=standards[0];
  if(state.selectedEssential)renderEssentialDetail(state.selectedEssential)
}
function renderEssentialDetail(standard){
  const rows=essentialRowsByStandard(standard),codes=standard.codes||[];
  $('#essentialBadge').textContent=standard.id;
  $('#essentialTitle').textContent=standard.title;
  $('#essentialCodes').textContent=`รหัส: ${essentialCodeList(standard).join(', ')}`;
  $('#essentialTotal').textContent=rows.length.toLocaleString();
  $('#essentialHigh').textContent=rows.filter(r=>isHigh(r[IDX.sev])).length.toLocaleString();
  $('#essentialRca').textContent=rows.filter((r,i)=>hasRcaForRow(r,i)).length.toLocaleString();
  $('#essentialCodeCount').textContent=new Set(rows.map(getIncidentCode).filter(Boolean)).size.toLocaleString();
  $('#essentialRows').innerHTML=codes.map(item=>{
    const code=typeof item==='string'?item:item.code,description=typeof item==='string'?item:item.description;
    const codeRows=rows.filter(r=>getIncidentCode(r)===code);
    const units=[...new Set(codeRows.flatMap(r=>String(r[IDX.mainUnit]||'').split(/[\n,;]+/).map(normUnit)).filter(Boolean))];
    const high=codeRows.filter(r=>isHigh(r[IDX.sev])).length;
    const rca=codeRows.filter((r,i)=>hasRcaForRow(r,i)).length;
    return `<tr><td><b>${esc(code)}</b></td><td>${esc(description||'-')}</td><td>${codeRows.length.toLocaleString()}</td><td><span class="severity-pill">${esc(maximumSeverity(codeRows)||'-')}</span></td><td>${high.toLocaleString()}</td><td>${esc(units.slice(0,6).join(', ')||'-')}</td><td>${rca.toLocaleString()}</td></tr>`
  }).join('')

  renderEssentialMatrix(standard);
}
function essentialExportRows(){
  const standard=state.selectedEssential;if(!standard)return[];
  return essentialRowsByStandard(standard).map((row,index)=>[
    standard.id,standard.title,getIncidentCode(row),row[IDX.date]||'',row[IDX.risk]||'',row[IDX.sub]||'',row[IDX.sev]||'',row[IDX.mainUnit]||'',hasRcaForRow(row,index)?'มี':'ไม่มี'
  ])
}

function showEssentialMode(mode){
  const registerView=$('#essentialRegisterView');
  const matrixView=$('#essentialMatrixView');
  const registerTab=$('#essentialRegisterTab');
  const matrixTab=$('#essentialMatrixTab');
  if(!registerView||!matrixView)return;

  const isMatrix=mode==='matrix';
  registerView.classList.toggle('hidden',isMatrix);
  matrixView.classList.toggle('hidden',!isMatrix);
  registerTab?.classList.toggle('active',!isMatrix);
  matrixTab?.classList.toggle('active',isMatrix);

  if(isMatrix&&state.selectedEssential){
    renderEssentialMatrix(state.selectedEssential)
  }
}

function likelihoodFromCount(count){
  if(count<=1)return count===0?0:1;
  if(count<=3)return 2;
  if(count<=6)return 3;
  if(count<=12)return 4;
  return 5
}

function severityScore(level){
  const value=String(level||'').trim().toUpperCase();
  const clinical={A:1,B:1,C:2,D:2,E:3,F:3,G:4,H:4,I:5};
  const general={'1':1,'2':2,'3':3,'4':4,'5':5};
  return clinical[value]||general[value]||1
}

function severityLabel(level){
  const value=String(level||'').trim().toUpperCase();
  if(['A','B'].includes(value))return 'A-B';
  if(['C','D'].includes(value))return 'C-D';
  if(['E','F'].includes(value))return 'E-F';
  if(['G','H'].includes(value))return 'G-H';
  if(value==='I')return 'I';
  return value||'-'
}

function matrixRiskClass(score){
  if(score<=4)return 'risk-low';
  if(score<=9)return 'risk-moderate';
  if(score<=16)return 'risk-high';
  return 'risk-extreme'
}

function renderEssentialMatrix(standard){
  const tbody=$('#essentialMatrixRows');
  if(!tbody||!standard)return;

  const rows=essentialRowsByStandard(standard);
  const grouped={};

  rows.forEach(row=>{
    const code=getIncidentCode(row);
    if(!code)return;
    if(!grouped[code])grouped[code]=[];
    grouped[code].push(row)
  });

  const columnTotals=[0,0,0,0,0];
  let grandTotal=0;
  let maxScore=0;
  let extremeCount=0;

  tbody.innerHTML=standard.codes.map(code=>{
    const codeRows=grouped[code]||[];
    const count=codeRows.length;
    const likelihood=likelihoodFromCount(count);
    grandTotal+=count;

    const riskText=codeRows.length
      ? String(codeRows[0][IDX.risk]||code)
      : code;

    let maxSeverityLevel='';
    let maxSeverityScore=0;
    const severityCounts={};

    codeRows.forEach(row=>{
      const level=String(row[IDX.sev]||'').trim().toUpperCase();
      if(!level)return;
      severityCounts[level]=(severityCounts[level]||0)+1;
      const score=severityScore(level);
      if(score>maxSeverityScore){
        maxSeverityScore=score;
        maxSeverityLevel=level
      }
    });

    const riskScore=likelihood*maxSeverityScore;
    maxScore=Math.max(maxScore,riskScore);
    if(riskScore>=17)extremeCount++;

    const cells=[1,2,3,4,5].map(col=>{
      if(col!==likelihood||count===0)return '<td class="matrix-empty">-</td>';

      columnTotals[col-1]+=count;
      const groupedLabels={};
      Object.entries(severityCounts).forEach(([level,n])=>{
        const label=severityLabel(level);
        groupedLabels[label]=(groupedLabels[label]||0)+n
      });

      const detail=Object.entries(groupedLabels)
        .sort((a,b)=>a[0].localeCompare(b[0],'th'))
        .map(([label,n])=>`${esc(label)}: ${n.toLocaleString()}`)
        .join('<br>');

      return `<td class="${matrixRiskClass(riskScore)}">
        <b>${detail||count.toLocaleString()}</b>
        <small>คะแนน ${riskScore}</small>
      </td>`
    }).join('');

    return `<tr>
      <th class="matrix-risk-title">
        <b>${esc(code)}</b>
        <span>${esc(riskText.replace(code,'').replace(/^[:\s-]+/,''))}</span>
        <small>ระดับสูงสุด ${esc(maxSeverityLevel||'-')} • โอกาสเกิด ${likelihood||'-'}</small>
      </th>
      ${cells}
      <td class="matrix-total">${count.toLocaleString()}</td>
    </tr>`
  }).join('');

  $('#matrixRiskCodes').textContent=standard.codes.length.toLocaleString();
  $('#matrixIncidentTotal').textContent=grandTotal.toLocaleString();
  $('#matrixMaxScore').textContent=maxScore.toLocaleString();
  $('#matrixExtremeCount').textContent=extremeCount.toLocaleString();

  columnTotals.forEach((n,i)=>{
    const el=$(`#matrixCol${i+1}`);
    if(el)el.textContent=n.toLocaleString()
  });
  $('#matrixGrandTotal').textContent=grandTotal.toLocaleString()
}

function exportEssentialCsv(){
  const standard=state.selectedEssential;if(!standard)return toast('กรุณาเลือกมาตรฐาน');
  const lines=[['มาตรฐาน','ชื่อมาตรฐาน','รหัส Incident','วันที่','รายการความเสี่ยง','เรื่องย่อย','ระดับ','หน่วยงานหลักที่แก้ไข','RCA'],...essentialExportRows()];
  const csv='\ufeff'+lines.map(row=>row.map(v=>'"'+String(v??'').replaceAll('"','""')+'"').join(',')).join('\n');
  download(csv,`Essential_${standard.id}_${state.year}.csv`,'text/csv')
}
function exportEssentialExcel(){
  const standard=state.selectedEssential;if(!standard)return toast('กรุณาเลือกมาตรฐาน');
  const rows=essentialExportRows();
  const body=rows.map(r=>`<tr>${r.map(v=>`<td>${esc(v)}</td>`).join('')}</tr>`).join('');
  const doc=`\ufeff<html><head><meta charset="utf-8"><style>table{border-collapse:collapse}th,td{border:1px solid #333;padding:5px}th{background:#dbe7f5}</style></head><body><h2>${esc(standard.title)}</h2><p>ปีงบประมาณ ${esc(state.year==='all'?'ทุกปี':state.year)}</p><table><thead><tr><th>มาตรฐาน</th><th>ชื่อมาตรฐาน</th><th>รหัส Incident</th><th>วันที่</th><th>รายการความเสี่ยง</th><th>เรื่องย่อย</th><th>ระดับ</th><th>หน่วยงานหลักที่แก้ไข</th><th>RCA</th></tr></thead><tbody>${body}</tbody></table></body></html>`;
  download(doc,`Essential_${standard.id}_${state.year}.xls`,'application/vnd.ms-excel')
}

function exportHaCsv(){
 const top=Object.entries(countMap(state.filtered.map(r=>String(r[IDX.risk]||'ไม่ระบุ').split(':')[0].trim()))).sort((a,b)=>b[1]-a[1]).slice(0,10);
 const lines=[['รายงาน HA หน่วยงาน'],['ปีงบประมาณ',state.year],['หน่วยงานหลักที่แก้ไข',contextText()],['อุบัติการณ์ทั้งหมด',state.filtered.length],['Clinical',state.filtered.filter(r=>riskType(r)==='Clinical').length],['Non-clinical',state.filtered.filter(r=>riskType(r)==='Non-clinical').length],['High Risk',state.filtered.filter(r=>isHigh(r[IDX.sev])).length],[],['อันดับ','ความเสี่ยงสำคัญ','จำนวน'],...top.map(([x,n],i)=>[i+1,x,n])];
 const csv='\ufeff'+lines.map(row=>row.map(v=>'"'+String(v??'').replaceAll('"','""')+'"').join(',')).join('\n');download(csv,`HA_Report_${state.year}.csv`,'text/csv')
}

init();





function currentYearLabel(){
  return state.year==='all' ? 'ทุกปีงบประมาณ' : `ปีงบประมาณ ${state.year}`;
}
