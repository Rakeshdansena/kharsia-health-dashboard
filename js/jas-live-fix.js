/* JAS LIVE FIX — Sector Wise first, Facility Wise second */
(function(){
'use strict';

const SID='1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0';
const GID='1018164338';
const TITLE='Jan Arogya Samiti Meeting FY 2026-27 Till July 2026';

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
.jas-title{
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
    for(let c=0;c<7;c++) row.push(clean(dt.getFormattedValue(r,c)));
    out.push(row);
  }
  return out;
}

function isTotal(v){
  return /^(total|योग|कुल)$/i.test(clean(v));
}

function isSectorHeader(v){
  const x=v.slice(0,7).join(' ').toLowerCase();
  return x.includes('sn') &&
         x.includes('sector') &&
         x.includes('no of aam facility') &&
         x.includes('target') &&
         x.includes('achievment');
}

function isFacilityHeader(v){
  const x=v.slice(0,7).join(' ').toLowerCase();
  return x.includes('sn') &&
         x.includes('nin') &&
         x.includes('name of sector') &&
         x.includes('name of facility') &&
         x.includes('target') &&
         x.includes('achievment');
}

function findSections(a){
  let sectorHeader=-1;
  let facilityHeader=-1;
  for(let i=0;i<a.length;i++){
    if(sectorHeader<0 && isSectorHeader(a[i])) sectorHeader=i;
    else if(facilityHeader<0 && isFacilityHeader(a[i])) facilityHeader=i;
  }
  return {sectorHeader,facilityHeader};
}

function collectSection(a,start,end,type){
  const rows=[];
  for(let r=start+1;r<(end<0?a.length:end);r++){
    const v=a[r].slice(0,7);
    if(isTotal(v[0])) break;
    if(/^\d+$/.test(v[0])){
      if(type==='sector' && v[1]) rows.push(v);
      if(type==='facility' && v[1] && v[2] && v[3]) rows.push(v);
    }
  }
  return rows;
}

function pctClass(p){ return p>=100?'jas-good':'jas-low'; }

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
     '<td colspan="3">Total</td><td>'+fac+'</td><td>'+target+'</td><td>'+ach+'</td><td></td></tr>'+
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
     '<td colspan="4">Total</td><td>'+target+'</td><td>'+ach+'</td><td></td></tr>'+
     '</tbody></table></div>';
  return h;
}

function render(dt){
  css();
  const a=getRows(dt);
  const sec=findSections(a);
  const sectorRows=sec.sectorHeader>=0
    ? collectSection(a,sec.sectorHeader,sec.facilityHeader,'sector')
    : [];
  const facilityRows=sec.facilityHeader>=0
    ? collectSection(a,sec.facilityHeader,-1,'facility')
    : [];

  let q=document.getElementById('jasMeetingModule');
  const rp=document.getElementById('reportPage');
  if(!q && rp){
    q=document.createElement('div');
    q.id='jasMeetingModule';
    rp.appendChild(q);
  }
  if(!q) return;

  let h='<div class="jas-title">'+TITLE+'</div>';
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
  q.setQuery('select A,B,C,D,E,F,G');
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
