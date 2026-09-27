/* JAS LIVE FIX — Sector Wise first, Facility Wise second */
(function(){
'use strict';

const SID='1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0';
const GID='1018164338';
const TITLE_FALLBACK='Jan Arogya Samiti Meeting FY 2026-27 Till August 2026';

const clean=v=>String(v??'').trim();
const num=v=>{
  const n=Number(clean(v).replace(/,/g,'').replace('%',''));
  return Number.isFinite(n)?n:0;
};
const esc=v=>String(v??'').replace(/[&<>"']/g,m=>({
  '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
}[m]));

function css(){
  if(document.getElementById('jas-sector-css')) return;
  const s=document.createElement('style');
  s.id='jas-sector-css';
  s.textContent=`
#jasMeetingModule{
 display:none;
 width:100%;
 background:#fff;
 padding:12px;
 border-radius:12px;
 box-sizing:border-box;
}
.jas-summary{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin:8px 0 18px}.jas-card{background:linear-gradient(135deg,#eff6ff,#ecfeff);border:1px solid #bae6fd;border-radius:10px;padding:10px;text-align:center;box-shadow:0 2px 7px rgba(15,23,42,.08)}.jas-card b{display:block;font-size:10px;color:#475569;letter-spacing:.6px}.jas-card strong{display:block;font-size:23px;color:#075985;margin-top:4px}.jas-title{
 text-align:center;
 font-weight:800;
 color:#075985;
 font-size:17px;
 margin:4px 0 12px;
}
.jas-section-title{
 text-align:left;
 font-size:15px;
 font-weight:800;
 color:#075985;
 margin:18px 0 8px;
}
.jas-page{
 background:#fff;
 margin-bottom:18px;
 overflow:auto;
}
.jas-table{
 width:100%;
 border-collapse:collapse;
 table-layout:fixed;
 font-size:11px;
}
.jas-table th,.jas-table td{
 border:1px solid #334155;
 padding:6px 4px;
 text-align:center;
 vertical-align:middle;
 word-break:break-word;
}
.jas-table th{
 background:#075985;
 color:#fff;
 font-weight:800;
}
.jas-good{background:#16a34a!important;color:#000!important;font-weight:900!important}
.jas-mid{background:#eab308!important;color:#111827!important;font-weight:900!important}
.jas-low{background:#dc2626!important;color:#000!important;font-weight:900!important}

.jas-table td:nth-child(2),
.jas-table td:nth-child(3),
.jas-table td:nth-child(4){
 text-align:left;
}
.jas-total td{
 font-weight:800;
 background:#dbeafe;
}
.jas-good,.jas-mid,.jas-low{font-weight:900!important;color:#000!important;}
.jas-tools{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:0 0 12px}
.jas-tools input{flex:1;min-width:220px;padding:9px 11px;border:1px solid #94a3b8;border-radius:7px;font-size:13px}
.jas-tools button{padding:9px 12px;border:0;border-radius:7px;background:#075985;color:#fff;font-weight:700;cursor:pointer}
.jas-tools button:hover{opacity:.9}
.jas-empty{
 padding:18px;
 text-align:center;
 color:#b91c1c;
 font-weight:700;
}
@media(max-width:700px){
 .jas-table{font-size:9px}
 .jas-table th,.jas-table td{padding:5px 3px}
}
`;
  document.head.appendChild(s);
}

function getRows(dt){
  const out=[];
  for(let r=0;r<dt.getNumberOfRows();r++){
    const row=[];
    for(let c=0;c<Math.min(13,dt.getNumberOfColumns());c++) row.push(clean(dt.getFormattedValue(r,c))); while(row.length<13) row.push('');
    out.push(row);
  }
  return out;
}

function isTotal(v){
  return /^(total|योग|कुल)$/i.test(clean(v));
}

function looksLikeNumber(v){
  const t=clean(v).replace(/,/g,'').replace('%','');
  return t!=='' && Number.isFinite(Number(t));
}

function isHeaderRow(v){
  const x=v.slice(0,8).map(norm).join(' ');
  return /target/.test(x) && /achiev/.test(x);
}

function findTitle(a){
  for(const r of a){
    for(const v of r){
      const s=clean(v);
      if(/jan\s*arogya\s*samiti\s*meeting/i.test(s)) return s;
    }
  }
  return TITLE_FALLBACK;
}

function norm(v){return clean(v).toLowerCase().replace(/[^a-z0-9]+/g,' ');}

function isAchievement(v){return /achiev/.test(norm(v));}

function isSectorHeader(v){
  const x=v.slice(0,7).map(norm).join(' ');
  return /\bsn\b/.test(x) &&
         /sector/.test(x) &&
         /target/.test(x) &&
         isAchievement(x) &&
         !/facility/.test(x) ||
         (/\bsn\b/.test(x) && /aam/.test(x) && /target/.test(x) && isAchievement(x) && !/name of facility/.test(x));
}

function isFacilityHeader(v){
  const x=v.slice(0,7).map(norm).join(' ');
  return /\bsn\b/.test(x) &&
         /nin/.test(x) &&
         /sector/.test(x) &&
         /facility/.test(x) &&
         /target/.test(x) &&
         isAchievement(x);
}

function findSections(a){
  let sectorHeader=-1, facilityHeader=-1;
  for(let i=0;i<a.length;i++){
    const x=a[i].map(norm).join(' ');
    if(!isHeaderRow(a[i])) continue;

    const hasFacility=/name of facility|facility name|facility/.test(x);
    const hasSector=/name of sector|sector/.test(x);
    const hasNIN=/\bnin\b/.test(x);

    if(hasFacility && hasNIN){
      facilityHeader=i;
    }else if(hasSector && !hasFacility && sectorHeader<0){
      sectorHeader=i;
    }
  }
  return {sectorHeader,facilityHeader};
}

function collectSection(a,start,end,type){
  const rows=[];
  const from=start>=0?start+1:0;
  const to=end>=0?end:a.length;

  for(let r=from;r<to;r++){
    const v=a[r].slice(0,7);
    while(v.length<7) v.push('');

    const sn=clean(v[0]);
    if(!/^\d+(?:\.0+)?$/.test(sn)) continue;

    const nin=clean(v[1]);
    const sector=clean(v[2]);
    const facility=clean(v[3]);
    const target=clean(v[4]);
    const achievement=clean(v[5]);

    if(!target && !achievement) continue;

    if(type==='sector' && sector && looksLikeNumber(v[3]) &&
       looksLikeNumber(target) && looksLikeNumber(achievement)){
      rows.push(v);
    }

    if(type==='facility' && nin && sector && facility &&
       looksLikeNumber(target) && looksLikeNumber(achievement)){
      rows.push(v);
    }
  }
  return rows;
}

function pctClass(p){ return p>=100?'jas-good':p>90?'jas-mid':'jas-low'; }

function renderSector(rows){
  let h='<div class="jas-section-title">Sector Wise</div>';
  h+='<div class="jas-page"><table class="jas-table"><thead><tr>'+
     '<th>SN</th><th>Sector</th><th>NIN</th><th>No of AAM Facility</th>'+
     '<th>Target till Aug 26</th><th>Achievment</th><th>%</th></tr></thead><tbody>';

  let fac=0,target=0,ach=0;
  rows.forEach(v=>{
    const F=num(v[3]),T=num(v[4]),A=num(v[5]),P=num(v[6]);
    fac+=F; target+=T; ach+=A;
    h+='<tr>'+
      '<td>'+esc(v[0])+'</td>'+
      '<td>'+esc(v[1])+'</td>'+
      '<td>'+esc(v[2])+'</td>'+
      '<td>'+F+'</td>'+
      '<td>'+T+'</td>'+
      '<td>'+A+'</td>'+
      '<td class="'+pctClass(P)+'">'+P+'%</td>'+
      '</tr>';
  });

  h+='<tr class="jas-total">'+
     '<td colspan="3">Total</td><td>'+fac+'</td><td>'+target+'</td><td>'+ach+'</td><td class="'+pctClass(target?ach/target*100:0)+'">'+(target?(ach/target*100).toFixed(1):0)+'%</td></tr>'+
     '</tbody></table></div>';
  return h;
}

function renderFacility(rows){
  let h='<div class="jas-section-title">Facility Wise</div>';
  h+='<div class="jas-page"><table class="jas-table"><thead><tr>'+
     '<th>SN</th><th>NIN</th><th>Name of Sector</th><th>Name of Facility</th>'+
     '<th>Target till Aug 26</th><th>Achievment</th><th>%</th></tr></thead><tbody>';

  let target=0,ach=0;
  rows.forEach(v=>{
    const T=num(v[4]),A=num(v[5]),P=num(v[6]);
    target+=T; ach+=A;
    h+='<tr>'+
      '<td>'+esc(v[0])+'</td>'+
      '<td>'+esc(v[1])+'</td>'+
      '<td>'+esc(v[2])+'</td>'+
      '<td>'+esc(v[3])+'</td>'+
      '<td>'+T+'</td>'+
      '<td>'+A+'</td>'+
      '<td class="'+pctClass(P)+'">'+P+'%</td>'+
      '</tr>';
  });

  h+='<tr class="jas-total">'+
     '<td colspan="4">Total</td><td>'+target+'</td><td>'+ach+'</td><td class="'+pctClass(target?ach/target*100:0)+'">'+(target?(ach/target*100).toFixed(1):0)+'%</td></tr>'+
     '</tbody></table></div>';
  return h;
}

function addJASTools(){
  return '<div class="jas-tools">'+
    '<input id="jasSearchInput" type="search" placeholder="Search Sector / Facility / NIN...">'+
    '<button type="button" onclick="window.jasExportExcel&&window.jasExportExcel()">Excel</button>'+
    '<button type="button" onclick="window.jasExportPDF&&window.jasExportPDF()">PDF</button>'+
  '</div>';
}

function filterJASTables(value){
  const q=clean(value).toLowerCase();
  document.querySelectorAll('#jasMeetingModule .jas-table tbody tr').forEach(tr=>{
    const isTotal=tr.classList.contains('jas-total');
    if(isTotal) return;
    tr.style.display=!q || tr.innerText.toLowerCase().includes(q)?'':'none';
  });
}

function setupJASTools(){
  const input=document.getElementById('jasSearchInput');
  if(input) input.oninput=()=>filterJASTables(input.value);
}

function exportJASTable(format){
  const mod=document.getElementById('jasMeetingModule');
  if(!mod) return;
  const title=findTitle(getRows(window.__jasLastData||{getNumberOfRows:()=>0,getNumberOfColumns:()=>0,getFormattedValue:()=>''}));
  const tables=Array.from(mod.querySelectorAll('.jas-table')).map(t=>t.outerHTML).join('<br>');
  if(format==='excel'){
    const html='<!doctype html><html><head><meta charset="utf-8"></head><body><h2>'+esc(title)+'</h2>'+tables+'</body></html>';
    const blob=new Blob([html],{type:'application/vnd.ms-excel'});
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
    a.download='JAS_Meeting_'+title.replace(/[^a-z0-9]+/gi,'_')+'.xls'; a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href),1000);
  }else{
    const w=window.open('','_blank');
    if(!w) return;
    w.document.write('<html><head><title>'+esc(title)+'</title><style>body{font-family:Arial;padding:20px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #333;padding:6px;text-align:center}th{background:#075985;color:#fff}h2{text-align:center}</style></head><body><h2>'+esc(title)+'</h2>'+tables+'</body></html>');
    w.document.close(); w.focus(); setTimeout(()=>w.print(),400);
  }
}
window.jasExportExcel=()=>exportJASTable('excel');
window.jasExportPDF=()=>exportJASTable('pdf');

function renderSummary(sectorRows,facilityRows){
  const sectors=sectorRows.filter(r=>clean(r[1]));
  const facilities=facilityRows.filter(r=>clean(r[3]));
  const target=sectors.length
    ? sectors.reduce((s,r)=>s+num(r[4]),0)
    : facilities.reduce((s,r)=>s+num(r[4]),0);
  const achievement=sectors.length
    ? sectors.reduce((s,r)=>s+num(r[5]),0)
    : facilities.reduce((s,r)=>s+num(r[5]),0);
  const overall=target?((achievement/target)*100):0;
  return '<div class="jas-summary">'+
    '<div class="jas-card"><b>SECTORS</b><strong>'+sectors.length+'</strong></div>'+
    '<div class="jas-card"><b>FACILITIES</b><strong>'+facilities.length+'</strong></div>'+
    '<div class="jas-card"><b>TARGET</b><strong>'+target+'</strong></div>'+
    '<div class="jas-card"><b>ACHIEVEMENT</b><strong>'+achievement+'</strong></div>'+
    '<div class="jas-card"><b>OVERALL %</b><strong>'+overall.toFixed(1)+'%</strong></div>'+
  '</div>';
}

function buildSectorRowsFromFacilities(facilityRows){
  const map=new Map();
  facilityRows.forEach(v=>{
    const sector=clean(v[2]);
    if(!sector) return;
    if(!map.has(sector)) map.set(sector,{sector,facilities:0,target:0,achievement:0});
    const x=map.get(sector);
    x.facilities++;
    x.target+=num(v[4]);
    x.achievement+=num(v[5]);
  });
  let sn=1;
  return Array.from(map.values()).map(x=>[
    String(sn++), x.sector, '-', String(x.facilities),
    String(x.target), String(x.achievement),
    x.target ? String((x.achievement/x.target*100).toFixed(1)) : '0'
  ]);
}

function scanJASRows(a){
  const facilityRows=[];
  const sectorRows=[];
  for(let r=0;r<a.length;r++){
    const v=a[r].slice(0,7);
    while(v.length<7) v.push('');
    const sn=clean(v[0]);
    if(!/^\d+(?:\.0+)?$/.test(sn)) continue;

    const c1=clean(v[1]), c2=clean(v[2]), c3=clean(v[3]);
    const c4=clean(v[4]), c5=clean(v[5]), c6=clean(v[6]);

    // Facility row: SN, NIN, Sector, Facility, Target, Achievement, %
    if(c1 && c2 && c3 && looksLikeNumber(c4) && looksLikeNumber(c5)){
      facilityRows.push(v);
      continue;
    }

    // Sector row: SN, Sector, NIN, No. of AAM Facility, Target, Achievement, %
    if(c1 && looksLikeNumber(c3) && looksLikeNumber(c4) && looksLikeNumber(c5)){
      sectorRows.push(v);
    }
  }
  return {sectorRows,facilityRows};
}

function render(dt){
  css();
  const a=getRows(dt);
  const sec=findSections(a);

  // First use the known section headers.
  let facilityRows=sec.facilityHeader>=0
    ? collectSection(a,sec.facilityHeader,-1,'facility')
    : [];

  let sectorRows=sec.sectorHeader>=0
    ? collectSection(a,sec.sectorHeader,sec.facilityHeader,'sector')
    : [];

  // Fallback: the JAS sheet layout can change (extra title rows,
  // merged headers, or different header wording). Scan the actual
  // 7 data columns instead of depending on header text.
  if(!facilityRows.length || !sectorRows.length){
    const scanned=scanJASRows(a);
    if(!facilityRows.length) facilityRows=scanned.facilityRows;
    if(!sectorRows.length) sectorRows=scanned.sectorRows;
  }

  // If the sheet contains only facility-wise data, derive Sector Wise.
  if(!sectorRows.length && facilityRows.length){
    sectorRows=buildSectorRowsFromFacilities(facilityRows);
  }

  let q=document.getElementById('jasMeetingModule');
  const rp=document.getElementById('reportPage');
  if(!q && rp){
    q=document.createElement('div');
    q.id='jasMeetingModule';
    rp.appendChild(q);
  }
  if(!q) return;

  window.__jasLastData=dt;
  const title=findTitle(a);
  let h='<div class="jas-title">'+esc(title)+'</div>';
  h+=renderSummary(sectorRows,facilityRows);
  h+=addJASTools();
  if(!sectorRows.length && !facilityRows.length){
    h+='<div class="jas-empty">JAS Meeting data नहीं मिला। Google Sheet की पहली 7 columns में सही header check करें।</div>';
  }else{
    // Required layout: Sector Wise first, Facility Wise second.
    if(sectorRows.length) h+=renderSector(sectorRows);
    if(facilityRows.length) h+=renderFacility(facilityRows);
  }
  q.innerHTML=h;
  q.style.display='block';
  setupJASTools();

  if(rp){
    Array.from(rp.children).forEach(e=>{
      if(e!==q) e.style.display='none';
    });
  }
}

function load(){
  if(!window.google?.visualization?.Query){
    setTimeout(load,300);
    return;
  }
  const q=new google.visualization.Query(
    'https://docs.google.com/spreadsheets/d/'+SID+'/gviz/tq?gid='+GID+'&headers=0'
  );
  /* IMPORTANT: JAS uses 7 columns A:G. */
  q.setQuery('select A,B,C,D,E,F,G where A is not null');
  q.send(r=>{
    if(r.isError()){
      console.error('JAS Sheet error:',r.getMessage());
      const qx=document.getElementById('jasMeetingModule');
      if(qx) qx.innerHTML='<div class="jas-empty">Google Sheet data load error: '+esc(r.getMessage())+'</div>';
      return;
    }
    render(r.getDataTable());
  });
}

window.renderJASFromData=render;

function openJAS(){
  const status=document.getElementById('status');
  if(status) status.style.display='none';

  document.getElementById('dashboardPage')?.classList.remove('active');
  document.getElementById('reportPage')?.classList.add('active');

  const title=document.getElementById('reportTitle');
  if(title) title.innerText='🤝 JAS Meeting';

  const pdfTitle=document.getElementById('pdfTitle');
  if(pdfTitle) pdfTitle.innerText='Kharsia Health Dashboard — JAS Meeting';

  const desc=document.getElementById('reportDescription');
  if(desc) desc.innerText='Jan Arogya Samiti Meeting Reporting';

  const search=document.getElementById('searchBox');
  if(search) search.value='';

  document.querySelectorAll('.menu-btn').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('#menuButtons .menu-btn').forEach(b=>{
    if(/jas meeting/i.test(b.textContent||'')) b.classList.add('active');
  });

  const tableBox=document.querySelector('.table-box');
  if(tableBox) tableBox.style.display='none';

  const rch=document.getElementById('rchContainer');
  if(rch) rch.style.display='none';

  load();
}

function hook(){
  if(window.__jasSectorHook) return;
  const old=window.openReport;
  if(typeof old!=='function') return;
  window.__jasSectorHook=true;
  window.openReport=function(i){
    const b=document.querySelectorAll('#menuButtons .menu-btn')[i];
    if(b && /jas meeting/i.test(b.textContent||'')){
      return old.apply(this,arguments);
    }
    return old.apply(this,arguments);
  };
  document.addEventListener('click',e=>{
    const b=e.target.closest('#menuButtons .menu-btn');
    if(b && /jas meeting/i.test(b.textContent||'')){
      e.preventDefault();
      e.stopImmediatePropagation();
      const buttons=Array.from(document.querySelectorAll('#menuButtons .menu-btn'));
      const idx=buttons.indexOf(b);
      if(idx>=0 && typeof window.openReport==='function'){
        window.openReport(idx);
      }
    }
  },true);
}

css();
let tries=0;
const timer=setInterval(()=>{
  hook();
  if(++tries>80) clearInterval(timer);
},500);

})();
