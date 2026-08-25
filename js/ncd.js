/* NCD MODULE — Google Sheet renderer, exact merged HTN/DM headers */
(function(){
'use strict';
const SHEET_ID='1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0';
const NCD_GID='1254412412';
const LAST_UPDATED='17-Aug-2026 23:06:53';
const clean=v=>String(v??'').replace(/\s+/g,' ').trim();
function rows(data){const out=[];for(let r=0;r<data.getNumberOfRows();r++){const a=[];for(let c=0;c<data.getNumberOfColumns();c++){let v='';try{v=data.getFormattedValue(r,c)||data.getValue(r,c)||''}catch(e){}a.push(clean(v))}out.push(a)}return out}
function title(r){return /NCD\s*status\s*2026-27/i.test(r.join(' '))}
function total(r){return /^(total|योग|कुल)$/i.test(clean(r[0]))}
function header(r){const s=r.join(' ').toLowerCase();return /subcenter|sector|total\s+population|screening\s+target|enrollment\s*30|abha\s*link/.test(s)}
const HTN_HEADERS=['Screening','%','Estimated Hypertensive Patient (NFHS - 29%)','Under Treatment','%','Followup','%','Under Control','%'];
const DM_HEADERS=['Screening','%','Estimated Diabetes Patients (NFHS - 10.8%)','Under Treatment','%','Followup','%','Under Control','%'];
function normalizeRow(r){const out=r.slice();while(out.length<23)out.push('');return out.slice(0,23)}
function table(dataRows){
 const t=document.createElement('table');t.className='ncd-final-table';
 const thd=document.createElement('thead');
 const r1=document.createElement('tr');
 ['Subcenter','Total Population','Screening Target (47.8% of Total Population)','Enrollment 30+ (Progressive)','%','ABHA Link (Progressive)','ABHA Link %'].forEach(v=>{const th=document.createElement('th');th.textContent=v;th.rowSpan=2;r1.appendChild(th)});
 let th=document.createElement('th');th.textContent='HTN';th.colSpan=9;th.rowSpan=1;r1.appendChild(th);
 th=document.createElement('th');th.textContent='DM';th.colSpan=9;th.rowSpan=1;r1.appendChild(th);
 thd.appendChild(r1);
 const r2=document.createElement('tr');
 HTN_HEADERS.forEach(v=>{const x=document.createElement('th');x.textContent=v;r2.appendChild(x)});
 DM_HEADERS.forEach(v=>{const x=document.createElement('th');x.textContent=v;r2.appendChild(x)});
 thd.appendChild(r2);t.appendChild(thd);
 const tb=document.createElement('tbody');
 dataRows.forEach(raw=>{if(!raw.some(Boolean))return;const r=normalizeRow(raw);const tr=document.createElement('tr');if(total(r))tr.className='ncd-total-row';for(let c=0;c<23;c++){const td=document.createElement('td');td.textContent=r[c]||'';tr.appendChild(td)}tb.appendChild(tr)});
 t.appendChild(tb);return t;
}
function css(){if(document.getElementById('ncd-final-css'))return;const s=document.createElement('style');s.id='ncd-final-css';s.textContent=`#ncdModuleContainer{display:block!important;width:100%;box-sizing:border-box}.ncd-final-heading{width:100%;text-align:center;background:#e0f2fe;color:#075985;border:1px solid #93c5fd;border-radius:7px;padding:11px 8px;margin:0 0 14px;font-size:17px;font-weight:900;box-sizing:border-box}.ncd-final-box{background:#fff;border-radius:10px;padding:10px;margin-bottom:18px;box-shadow:0 3px 12px rgba(0,0,0,.07);overflow:auto}.ncd-final-title{text-align:center;background:#e0f2fe;border-left:5px solid #075985;color:#075985;font-size:18px;font-weight:900;padding:9px;margin-bottom:9px}.ncd-final-table{width:100%;min-width:2500px;border-collapse:collapse;font-size:12px;table-layout:auto}.ncd-final-table th{background:#075985;color:#fff;border:1px solid #cbd5e1;padding:8px 6px;text-align:center;vertical-align:middle;font-weight:800;line-height:1.25}.ncd-final-table thead tr:first-child th{font-size:13px;background:#0c4a6e}.ncd-final-table thead tr:nth-child(2) th{font-size:11px;background:#075985}.ncd-final-table td{border:1px solid #cbd5e1;padding:7px 6px;text-align:center;vertical-align:middle;white-space:normal}.ncd-final-table td:first-child{text-align:left;font-weight:700;min-width:150px}.ncd-final-table tbody tr:nth-child(even){background:#f8fafc}.ncd-final-table .ncd-total-row{background:#dbeafe!important;font-weight:900}@media(max-width:800px){.ncd-final-table{min-width:2500px}.ncd-final-heading{font-size:15px}}`;document.head.appendChild(s)}
function draw(data){css();const old=document.getElementById('ncdModuleContainer');if(old)old.remove();const tb=document.querySelector('.table-box');if(tb)tb.style.display='none';const rc=document.getElementById('rchContainer');if(rc)rc.style.display='none';const out=document.createElement('div');out.id='ncdModuleContainer';(document.getElementById('pdfArea')||document.body).appendChild(out);const rs=rows(data);let ti=[];rs.forEach((r,i)=>{if(title(r))ti.push(i)});const t1=ti[0]??-1,t2=ti[1]??-1;const findHead=(a,b)=>{for(let i=Math.max(0,a);i<(b<0?rs.length:b);i++)if(header(rs[i]))return i;return-1};const fh=findHead(t1+1,t2),sh=findHead(t2+1,rs.length);const end=(a,b)=>{for(let i=a;i<(b<0?rs.length:b);i++)if(total(rs[i]))return i;return b<0?rs.length:b};const fe=end(fh+1,t2),se=end(sh+1,rs.length);const frows=fh>=0?rs.slice(fh+1,fe):[],srows=sh>=0?rs.slice(sh+1,se):[];const h=document.createElement('div');h.className='ncd-final-heading';h.textContent='NCD Status 2026-27 As On Last Updated On : '+LAST_UPDATED;out.appendChild(h);[['Sector Wise',srows],['Facility Wise',frows]].forEach(x=>{const b=document.createElement('div');b.className='ncd-final-box';const q=document.createElement('div');q.className='ncd-final-title';q.textContent=x[0];b.appendChild(q);b.appendChild(table(x[1]));out.appendChild(b)})}
function renderNCDTable(){if(window.google&&google.visualization){const q=new google.visualization.Query('https://docs.google.com/spreadsheets/d/'+SHEET_ID+'/gviz/tq?gid='+NCD_GID+'&headers=0');q.setQuery('select *');q.send(r=>{if(r.isError()){console.error('NCD Sheet:',r.getMessage());return}draw(r.getDataTable())})}}
window.renderNCDTable=renderNCDTable;
})();
