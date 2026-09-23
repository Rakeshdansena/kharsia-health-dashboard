/* JAS LIVE FIX — Sector Wise first, Facility Wise second */
(function(){
'use strict';

const SID='1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0';
const GID='1018164338';
const TITLE_FALLBACK='Jan Arogya Samiti Meeting FY 2026-27 Till July 2026';

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
 display:block!important;
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
.jas-table td:nth-child(2),
.jas-table td:nth-child(3),
.jas-table td:nth-child(4){
 text-align:left;
}
.jas-total td{
 font-weight:800;
 background:#dbeafe;
}
.jas-good{
 color:#15803d!important;
 font-weight:900;
}
.jas-mid{color:#b45309!important;font-weight:900}
.jas-low{
 color:#dc2626!important;
 font-weight:900;
}
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
    for(let c=0;c<13;c++) row.push(clean(dt.getFormattedValue(r,c)));
    out.push(row);
  }
  return out;
}

function isTotal(v){
  return /^(total|योग|कुल)$/i.test(clean(v));
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
    if(/\bsn\b/.test(x) && /target/.test(x) && /achiev/.test(x)){
      if(/name of facility|facility/.test(x)) facilityHeader=i;
      else if(sectorHeader<0) sectorHeader=i;
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
    const sn=clean(v[0]), d=clean(v[3]);
    if(!/^\d+(?:\.0+)?$/.test(sn)) continue;
    if(type==='sector' && v[1] && v[2] && d && Number.isFinite(num(d))) rows.push(v);
    if(type==='facility' && v[1] && v[2] && d && !Number.isFinite(num(d))) rows.push(v);
  }
  return rows;
}

function pctClass(p){ return p>=100?'jas-good':p>=80?'jas-mid':'jas-low'; }

function renderSector(rows){
  let h='<div class="jas-section-title">Sector Wise</div>';
  h+='<div class="jas-page"><table class="jas-table"><thead><tr>'+
     '<th>SN</th><th>Sector</th><th>NIN</th><th>No of AAM Facility</th>'+
     '<th>Target till July 26</th><th>Achievment</th><th>%</th></tr></thead><tbody>';

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
     '<th>Target till July 26</th><th>Achievment</th><th>%</th></tr></thead><tbody>';

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

function renderSummary(sectorRows,facilityRows){
  const sT=sectorRows.reduce((a,v)=>a+num(v[4]),0), sA=sectorRows.reduce((a,v)=>a+num(v[5]),0);
  const pct=sT?((sA/sT)*100):0;
  return '<div class="jas-summary">'+
    '<div class="jas-card"><b>SECTORS</b><strong>'+sectorRows.length+'</strong></div>'+
    '<div class="jas-card"><b>FACILITIES</b><strong>'+facilityRows.length+'</strong></div>'+
    '<div class="jas-card"><b>TARGET</b><strong>'+sT+'</strong></div>'+
    '<div class="jas-card"><b>ACHIEVEMENT</b><strong>'+sA+'</strong></div>'+
    '<div class="jas-card"><b>OVERALL %</b><strong class="'+pctClass(pct)+'">'+pct.toFixed(1)+'%</strong></div>'+
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

function render(dt){
  css();
  const a=getRows(dt);
  const sec=findSections(a);
  const facilityRows=sec.facilityHeader>=0
    ? collectSection(a,sec.facilityHeader,-1,'facility')
    : [];

  // Current Google Sheet is facility-wise only. Build Sector Wise
  // automatically from the facility rows so both views are always shown.
  let sectorRows=sec.sectorHeader>=0
    ? collectSection(a,sec.sectorHeader,sec.facilityHeader,'sector')
    : [];
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

  const title=findTitle(a);
  let h='<div class="jas-title">'+esc(title)+'</div>';
  h+=renderSummary(sectorRows,facilityRows);
  if(!sectorRows.length && !facilityRows.length){
    h+='<div class="jas-empty">JAS Meeting data नहीं मिला। Google Sheet की पहली 7 columns में सही header check करें।</div>';
  }else{
    if(sectorRows.length) h+=renderSector(sectorRows);
    if(facilityRows.length) h+=renderFacility(facilityRows);
  }
  q.innerHTML=h;

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
  q.setQuery('select A,B,C,D,E,F,G,H,I,J,K,L,M where A is not null');
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

function openJAS(){
  document.getElementById('dashboardPage')?.classList.remove('active');
  document.getElementById('reportPage')?.classList.add('active');
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
      openJAS();
      return;
    }
    return old.apply(this,arguments);
  };
  document.addEventListener('click',e=>{
    const b=e.target.closest('#menuButtons .menu-btn');
    if(b && /jas meeting/i.test(b.textContent||'')){
      e.preventDefault();
      e.stopImmediatePropagation();
      openJAS();
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
