/* JAS Meeting - locked final layout / robust renderer */
(function(){
'use strict';

const GID='1018164338';
// index_updated.html defines SPREADSHEET_ID with const, so it is not a window property.
const SID=(typeof SPREADSHEET_ID!=='undefined' ? SPREADSHEET_ID : window.SPREADSHEET_ID);
const TITLE='Jan Arogya Samiti Meeting FY 2026-27 Till July 2026';

const N=v=>Number(String(v??'').replace(/,/g,'').replace('%','').trim())||0;
const E=v=>String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));

function addStyle(){
 if(document.getElementById('jas-final-css')) return;
 const s=document.createElement('style');
 s.id='jas-final-css';
 s.textContent=`
 #jasMeetingModule{display:block!important;width:100%;}
 #reportPage.jas-active .report-header,
 #reportPage.jas-active .status,
 #reportPage.jas-active .table-box,
 #reportPage.jas-active .report-actions{display:none!important;}
 .jas-page{background:#fff;border-radius:12px;padding:10px;margin:0 0 14px;box-shadow:0 3px 12px rgba(0,0,0,.07);overflow:auto;}
 .jas-table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:11px;}
 .jas-table th,.jas-table td{border:1px solid #cbd5e1;padding:6px 4px;text-align:center;vertical-align:middle;word-break:break-word;}
 .jas-table .jas-title{background:#e0f2fe;color:#075985;font-size:15px;font-weight:900;padding:8px;}
 .jas-table th{background:#075985;color:#fff;font-weight:800;}
 .jas-table .name{text-align:left;font-weight:700;}
 .jas-total{background:#dbeafe!important;font-weight:900;}
 .jas-good{background:#dcfce7!important;color:#166534!important;font-weight:900;}
 .jas-low{background:#fecaca!important;color:#991b1b!important;font-weight:900;}
 .jas-zero{background:#fef3c7!important;color:#92400e!important;font-weight:900;}
 .jas-fac th:nth-child(1){width:7%}.jas-fac th:nth-child(2){width:17%}.jas-fac th:nth-child(3){width:27%}.jas-fac th:nth-child(4){width:17%}.jas-fac th:nth-child(5){width:17%}.jas-fac th:nth-child(6){width:15%}
 .jas-sec th:nth-child(1){width:7%}.jas-sec th:nth-child(2){width:23%}.jas-sec th:nth-child(3){width:20%}.jas-sec th:nth-child(4){width:18%}.jas-sec th:nth-child(5){width:17%}.jas-sec th:nth-child(6){width:15%}
 @media(max-width:700px){.jas-table{font-size:9px}.jas-table th,.jas-table td{padding:4px 2px}.jas-page{padding:6px}}
 @media print{.jas-page{box-shadow:none;border-radius:0;margin:0;padding:0}.jas-table{font-size:7px}.jas-table th,.jas-table td{padding:3px 2px}}
 `;
 document.head.appendChild(s);
}

function pct(v){
 const x=N(v);
 const c=x===0?'jas-zero':x<100?'jas-low':'jas-good';
 return '<td class="'+c+'">'+x+'%</td>';
}

function parse(dt){
 const a=[];
 for(let r=0;r<dt.getNumberOfRows();r++){
  const v=[];
  for(let c=0;c<9;c++) v.push(String(dt.getFormattedValue(r,c)||'').trim());
  // Ignore Google Sheet title/header rows; accept only actual numbered facility rows.
  if(/^\d+$/.test(v[0]) && v[2]){
   a.push([v[0],v[1],v[2],N(v[3]),N(v[4]),N(v[5]),v[6],N(v[7]),N(v[8])]);
  }
 }
 return a;
}

function facility(rows){
 let h='<table class="jas-table jas-fac"><thead>'+
 '<tr><th colspan="6" class="jas-title">'+TITLE+'</th></tr>'+
 '<tr><th>SN</th><th>NIN</th><th>NAME</th><th>Target<br>till July 26</th><th>Achievment</th><th>%</th></tr>'+
 '</thead><tbody>';
 rows.forEach(r=>h+='<tr><td>'+E(r[0])+'</td><td>'+E(r[1])+'</td><td class="name">'+E(r[2])+'</td><td>'+r[3]+'</td><td>'+r[4]+'</td>'+pct(r[5])+'</tr>');
 const T=rows.reduce((a,r)=>a+r[3],0), A=rows.reduce((a,r)=>a+r[4],0);
 return h+'<tr class="jas-total"><td colspan="2">Total</td><td>'+rows.length+' Facilities</td><td>'+T+'</td><td>'+A+'</td>'+pct(T?Math.round(A/T*100):0)+'</tr></tbody></table>';
}

function sector(rows){
 const m={};
 rows.forEach(r=>{
  const k=r[6]||'Other';
  if(!m[k]) m[k]={f:0,t:0,a:0};
  m[k].f++;m[k].t+=r[3];m[k].a+=r[4];
 });
 let h='<table class="jas-table jas-sec"><thead>'+
 '<tr><th colspan="6" class="jas-title">'+TITLE+'</th></tr>'+
 '<tr><th>SN</th><th>Sector</th><th>No of AAM Facility</th><th>Target<br>till July 26</th><th>Achievment</th><th>%</th></tr>'+
 '</thead><tbody>',i=1,T=0,A=0;
 Object.entries(m).forEach(([k,x])=>{
  T+=x.t;A+=x.a;
  h+='<tr><td>'+i+++'</td><td class="name">'+E(k)+'</td><td>'+x.f+'</td><td>'+x.t+'</td><td>'+x.a+'</td>'+pct(x.t?Math.round(x.a/x.t*100):0)+'</tr>';
 });
 return h+'<tr class="jas-total"><td colspan="2">Total</td><td>'+rows.length+'</td><td>'+T+'</td><td>'+A+'</td>'+pct(T?Math.round(A/T*100):0)+'</tr></tbody></table>';
}

function render(dt){
 addStyle();
 const rows=parse(dt);
 const rp=document.getElementById('reportPage');
 if(!rp) return;
 rp.classList.add('jas-active');
 rp.querySelectorAll('.report-header,.status,.table-box,.report-actions').forEach(e=>e.style.display='none');
 let q=document.getElementById('jasMeetingModule');
 if(!q){q=document.createElement('div');q.id='jasMeetingModule';rp.appendChild(q);}
 q.innerHTML=rows.length ? '<div class="jas-page">'+facility(rows)+'</div><div class="jas-page">'+sector(rows)+'</div>' : '<div class="jas-page">JAS Meeting data नहीं मिला।</div>';
}

function cleanNormal(){
 const rp=document.getElementById('reportPage');
 if(!rp) return;
 rp.classList.remove('jas-active');
 const q=document.getElementById('jasMeetingModule');
 if(q) q.remove();
}

function loadDirect(){
 addStyle();
 if(!window.google?.visualization?.Query || !SID) return;
 const q=new google.visualization.Query('https://docs.google.com/spreadsheets/d/'+SID+'/gviz/tq?gid='+GID+'&headers=0');
 q.setQuery('select A,B,C,D,E,F,G,H,I');
 q.send(r=>{
  if(!r.isError()) render(r.getDataTable());
  else {
   const rp=document.getElementById('reportPage');
   if(rp){
    rp.classList.add('jas-active');
    let qbox=document.getElementById('jasMeetingModule');
    if(!qbox){qbox=document.createElement('div');qbox.id='jasMeetingModule';rp.appendChild(qbox);}
    qbox.innerHTML='<div class="jas-page">JAS Meeting data load नहीं हो पाया: '+E(r.getMessage())+'</div>';
   }
  }
 });
}

function install(){
 addStyle();
 if(!window.openReport || !window.REPORTS) return;

 // Primary protection: intercept report opening so generic renderer never owns JAS.
 if(!window.openReport.__jasFinal){
  const oldOpen=window.openReport;
  const wrappedOpen=function(i){
   if(window.REPORTS[i] && window.REPORTS[i].name==='JAS Meeting'){
    document.getElementById('dashboardPage')?.classList.remove('active');
    document.getElementById('reportPage')?.classList.add('active');
    const t=document.getElementById('reportTitle');
    const d=document.getElementById('reportDescription');
    if(t)t.innerText='🤝 JAS Meeting';
    if(d)d.innerText='Jan Arogya Samiti Meeting Reporting';
    document.getElementById('searchBox')?.setAttribute('value','');
    document.querySelectorAll('#menuButtons .menu-btn').forEach((b,n)=>b.classList.toggle('active',n===i));
    document.getElementById('dashboardBtn')?.classList.remove('active');
    document.querySelector('.table-box')?.style.setProperty('display','none','important');
    document.querySelector('.report-actions')?.style.setProperty('display','none','important');
    document.getElementById('status')?.style.setProperty('display','none','important');
    loadDirect();
    return;
   }
   cleanNormal();
   return oldOpen.apply(this,arguments);
  };
  wrappedOpen.__jasFinal=true;
  window.openReport=wrappedOpen;
 }

 // Secondary protection: if the base app calls its loader before/after the wrapper,
 // render the JAS sheet ourselves and suppress the generic Column 1/2/3 table.
 if(window.loadGoogleSheet && !window.loadGoogleSheet.__jasFinal){
  const oldLoad=window.loadGoogleSheet;
  const wrappedLoad=function(gid){
   const idx=window.currentReportIndex;
   const report=window.REPORTS?.[idx];
   if(report && report.name==='JAS Meeting') return loadDirect();
   cleanNormal();
   return oldLoad.apply(this,arguments);
  };
  wrappedLoad.__jasFinal=true;
  window.loadGoogleSheet=wrappedLoad;
 }
}

window.renderJASMeeting=loadDirect;
[0,100,300,800,1500,2500,4000].forEach(t=>setTimeout(install,t));
})();
