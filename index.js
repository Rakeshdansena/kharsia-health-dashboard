// Kharsia Health Dashboard main controller

const SPREADSHEET_ID = "1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0";

const REPORTS=[
 {name:"Ayushman Card",icon:"💳",gid:"925649620",description:"आयुष्मान कार्ड एवं वय वंदन कार्ड की प्रगति"},
 {name:"RCH 2.0",icon:"👩‍🍼",gid:"1044088930",description:"RCH 2.0 PW Registration Detail"},
 {name:"NCD",icon:"❤️",gid:"1254412412",description:"NCD Status 2026-27"},
 {name:"JAS Meeting",icon:"🤝",gid:"1018164338",description:"JAS Meeting Reporting"},
 {name:"Health & Wellness Center",icon:"🏥",gid:"0",description:"Health Wellness Center - Block Kharsia"},
 {name:"Ayushman Shivir",icon:"🏕️",gid:"1262815420",description:"Ayushman Shivir Reporting FY 2026-27"},
 {name:"Wellness Activity",icon:"🩺",gid:"447031017",description:"Ayushman Arogya Mandir Wellness Activity"},
 {name:"RBSK",icon:"👶",gid:"1502752823",description:"राष्ट्रीय बाल स्वास्थ्य कार्यक्रम"},
 {name:"NRC Kharsia",icon:"🏥",gid:"1010102020",description:"NRC Kharsia Reporting"},
 {name:"Blindness Control",icon:"👁️",gid:"1002009767",description:"National Blindness Control Programme"},
 {name:"NQAS Certification",icon:"🏅",gid:"728123647",description:"NQAS Certification Detail Kharsia"},
 {name:"Dialysis",icon:"💧",gid:"781496964",description:"प्रधानमंत्री राष्ट्रीय डायलिसिस कार्यक्रम"},
 {name:"NLEP",icon:"🦠",gid:"1536656599",description:"National Leprosy Eradication Programme"}
];

let currentReportIndex=null;
let currentData=null;
let currentRCHData=null;
const RCH_SECTORS=["Barra","Binjkot","Gorpar","Jobi","Sarwani","Sondka","Turekela","Urban Kharsia"];

function showDashboard(){
 document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
 const page=document.getElementById('dashboardPage'); if(page) page.classList.add('active');
 document.querySelectorAll('.menu-btn').forEach(b=>b.classList.remove('active'));
 const b=document.getElementById('dashboardBtn'); if(b) b.classList.add('active');
 currentReportIndex=null;
 loadBlockSummary();
}

function createMenu(){
 const box=document.getElementById('menuButtons');
 if(!box) return;
 box.innerHTML='';
 REPORTS.forEach((r,i)=>{
  const b=document.createElement('button');
  b.className='menu-btn';
  b.innerHTML=`${r.icon} ${r.name}`;
  b.onclick=()=>openReport(i);
  box.appendChild(b);
 });
 const cards=document.getElementById('dashboardCards');
 if(cards){
  cards.innerHTML=REPORTS.map((r,i)=>`<div class="card" onclick="openReport(${i})"><div class="card-icon">${r.icon}</div><h3>${r.name}</h3><p>${r.description}</p></div>`).join('');
 }
 loadBlockSummary();
}

function gvizUrl(gid){
 return `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&gid=${encodeURIComponent(gid)}`;
}

async function fetchSheetRows(report){
 if(!report || !report.gid || report.gid==='0') throw new Error('Google Sheet GID उपलब्ध नहीं है');
 const res=await fetch(gvizUrl(report.gid),{cache:'no-store'});
 if(!res.ok) throw new Error(`HTTP ${res.status}`);
 const text=await res.text();
 const start=text.indexOf('{'), end=text.lastIndexOf('}');
 if(start<0||end<start) throw new Error('Google Sheet response invalid');
 const json=JSON.parse(text.slice(start,end+1));
 const table=json.table||{};
 const cols=(table.cols||[]).map((c,i)=>c.label||`Column ${i+1}`);
 const rows=(table.rows||[]).map(r=>(r.c||[]).map(c=>c ? (c.v ?? c.f ?? '') : ''));
 return {cols,rows};
}

function num(v){
 if(typeof v==='number' && Number.isFinite(v)) return v;
 const s=String(v??'').replace(/,/g,'').replace(/%/g,'').trim();
 if(!s || !/^[-+]?\d*\.?\d+$/.test(s)) return null;
 const n=Number(s); return Number.isFinite(n)?n:null;
}
function fmt(n){
 return Number.isInteger(n)?n.toLocaleString('en-IN'):n.toLocaleString('en-IN',{maximumFractionDigits:2});
}
function cleanLabel(s){return String(s||'').replace(/\s+/g,' ').trim();}

function buildModuleSummary(report,data){
 const rows=data.rows.filter(r=>r.some(v=>String(v??'').trim()!==''));
 const cols=data.cols;
 const metricCandidates=[];
 for(let c=0;c<cols.length;c++){
  const label=cleanLabel(cols[c]);
  const low=label.toLowerCase();
  if(!label || /%|percent|percentage|rate|क्रम|sn\b|serial/.test(low)) continue;
  let total=0,count=0;
  for(const row of rows){
   const n=num(row[c]);
   if(n!==null){total+=n;count++;}
  }
  if(count>=1 && !/^(year|fy|month|date|दिन|वर्ष|facility code|code)$/i.test(label)) metricCandidates.push({label,total,count});
 }
 // Prefer meaningful health-programme columns, then largest numeric totals.
 const priority=/(total|कुल|target|लक्ष्य|card|registration|screening|treatment|follow|control|patient|beneficiar|activity|meeting|camp|visit|abha|delivery|pregnan|immun|referr|case|test|service)/i;
 metricCandidates.sort((a,b)=>{
  const ap=priority.test(a.label)?1:0, bp=priority.test(b.label)?1:0;
  if(ap!==bp) return bp-ap;
  return b.total-a.total;
 });
 const metrics=metricCandidates.slice(0,3);
 const uniqueFacilities=new Set();
 rows.forEach(r=>{
  const candidates=r.slice(0,Math.min(3,r.length)).map(v=>cleanLabel(v)).filter(v=>v && !/^total$|^योग$|^कुल$/i.test(v));
  const name=candidates.find(v=>/[A-Za-zअ-ह]/.test(v));
  if(name) uniqueFacilities.add(name);
 });
 return {rows:rows.length,facilities:uniqueFacilities.size,metrics};
}

async function loadBlockSummary(){
 const grid=document.getElementById('blockSummaryGrid');
 const status=document.getElementById('blockSummaryStatus');
 if(!grid||!status) return;
 status.className='block-summary-status';
 status.textContent='सभी programme का live summary load हो रहा है...';
 grid.innerHTML=REPORTS.map(r=>`<div class="summary-module-card"><div class="summary-module-title"><span class="summary-module-icon">${r.icon}</span>${r.name}</div><div class="summary-stats"><div class="summary-stat"><div class="summary-stat-label">Status</div><div class="summary-stat-value">Loading…</div></div><div class="summary-stat"><div class="summary-stat-label">Google Sheet</div><div class="summary-stat-value">Live</div></div></div></div>`).join('');
 let ok=0;
 const results=await Promise.all(REPORTS.map(async r=>{
  try{return {r,data:await fetchSheetRows(r)}}catch(e){return {r,error:e.message||'Load failed'}}
 }));
 grid.innerHTML='';
 results.forEach(({r,data,error})=>{
  const card=document.createElement('div');
  card.className='summary-module-card'+(error?' summary-error':'');
  card.onclick=()=>openReport(REPORTS.indexOf(r));
  if(error){
   card.innerHTML=`<div class="summary-module-title"><span class="summary-module-icon">${r.icon}</span>${r.name}</div><div style="font-size:12px;line-height:1.5">Data load नहीं हुआ: ${error}<br><span style="color:#64748b">Click करके module खोलें</span></div>`;
   grid.appendChild(card); return;
  }
  ok++;
  const s=buildModuleSummary(r,data);
  const metrics=s.metrics.map(m=>`<div class="summary-metric"><span class="summary-metric-name" title="${m.label}">${m.label}</span><span class="summary-metric-value">${fmt(m.total)}</span></div>`).join('');
  card.innerHTML=`<div class="summary-module-title"><span class="summary-module-icon">${r.icon}</span>${r.name}</div><div class="summary-stats"><div class="summary-stat"><div class="summary-stat-label">Data Rows</div><div class="summary-stat-value">${fmt(s.rows)}</div></div><div class="summary-stat"><div class="summary-stat-label">Facilities / Entries</div><div class="summary-stat-value">${fmt(s.facilities)}</div></div></div>${metrics||'<div style="font-size:12px;color:#64748b">Numeric summary उपलब्ध नहीं</div>'}`;
  grid.appendChild(card);
 });
 status.textContent=`Block Kharsia Summary: ${ok}/${REPORTS.length} programmes का live data load हुआ। Last refresh: ${new Date().toLocaleTimeString('en-IN')}`;
}

async function openReport(index){
 currentReportIndex=index;
 const report=REPORTS[index];
 document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
 document.getElementById('reportPage')?.classList.add('active');
 document.querySelectorAll('.menu-btn').forEach(b=>b.classList.remove('active'));
 const buttons=[...document.querySelectorAll('.menu-btn')];
 const target=buttons.find(b=>b.textContent.includes(report.name)); if(target) target.classList.add('active');
 document.getElementById('reportTitle').textContent=report.name;
 document.getElementById('reportDescription').textContent=report.description;
 const status=document.getElementById('status');
 if(status){status.className='status';status.textContent='Google Sheet से data load हो रहा है...';}
 try{
  const data=await fetchSheetRows(report);
  currentData=data;
  if(typeof clearTable==='function') clearTable();
  if(typeof renderNormalTable==='function') renderNormalTable({
   getNumberOfColumns:()=>data.cols.length,
   getNumberOfRows:()=>data.rows.length,
   getColumnLabel:c=>data.cols[c],
   getFormattedValue:(r,c)=>String(data.rows[r]?.[c]??'')
  });
  if(status){status.className='status success';status.textContent='✓ Google Sheet से data successfully loaded';}
 }catch(e){
  if(status){status.className='status error';status.textContent='❌ Google Sheet से data load नहीं हुआ: '+(e.message||e);}
 }
}

function refreshReport(){ if(currentReportIndex!==null) openReport(currentReportIndex); }

function googleSheetLoad(){ return loadBlockSummary(); }

// Start dashboard
window.addEventListener('DOMContentLoaded',()=>createMenu());
