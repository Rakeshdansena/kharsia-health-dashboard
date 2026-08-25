/* NCD MODULE — exact Google Sheet layout: Sector Wise first, Facility Wise second */
(function(){
'use strict';
const LAST_UPDATED='17-Aug-2026 23:06:53';
const clean=v=>String(v??'').replace(/\s+/g,' ').trim();
function row(data,r){let a=[];for(let c=0;c<data.getNumberOfColumns();c++){let v='';try{v=data.getFormattedValue(r,c)||data.getValue(r,c)||''}catch(e){}a.push(clean(v))}return a}
function text(r){return r.join(' ').toLowerCase()}
function isHeader(r){let s=text(r);return /subcenter|total\s+population|screening\s+target|enrollment\s*30|abha\s*link/i.test(s)}
function isSectorHeader(r){return /^sector$/i.test(clean(r[0])) || /\bsector\b/i.test(text(r)) && /total\s+population/i.test(text(r))}
function isTotal(r){return /^total$/i.test(clean(r[0]))}
function makeTable(head,rows){let t=document.createElement('table');t.className='ncd-table';let thead=document.createElement('thead'),tr=document.createElement('tr');head.forEach(v=>{let th=document.createElement('th');th.textContent=v;tr.appendChild(th)});thead.appendChild(tr);t.appendChild(thead);let tbody=document.createElement('tbody');rows.forEach(r=>{if(!r.some(Boolean)||isHeader(r)||isSectorHeader(r))return;let tr=document.createElement('tr');if(isTotal(r))tr.className='ncd-total-row';for(let i=0;i<head.length;i++){let td=document.createElement('td');td.textContent=r[i]||'';tr.appendChild(td)}tbody.appendChild(tr)});t.appendChild(tbody);return t}
function css(){if(document.getElementById('ncd-exact-style'))return;let s=document.createElement('style');s.id='ncd-exact-style';s.textContent=`#ncdModuleContainer{display:none}.ncd-status-heading{width:100%;box-sizing:border-box;text-align:center;font-size:16px;font-weight:900;color:#075985;background:#e0f2fe;border:1px solid #93c5fd;padding:10px 8px;margin:0 0 14px;border-radius:6px}.ncd-section{background:#fff;padding:10px;margin:0 0 18px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.08);overflow:auto}.ncd-section-title{font-size:17px;font-weight:900;color:#075985;background:#e0f2fe;border-left:5px solid #075985;padding:9px 10px;margin:0 0 8px}.ncd-table{width:100%;border-collapse:collapse;font-size:12px;min-width:1500px}.ncd-table th{background:#075985;color:#fff;border:1px solid #cbd5e1;padding:8px 6px;text-align:center;vertical-align:middle;white-space:normal}.ncd-table td{border:1px solid #cbd5e1;padding:7px 6px;text-align:center;vertical-align:middle}.ncd-table td:first-child{text-align:left;font-weight:700}.ncd-total-row{background:#dbeafe!important;font-weight:900}.ncd-section-title + .ncd-table{margin-bottom:0}`;document.head.appendChild(s)}
function renderNCDTable(data){
 css();
 let rch=document.getElementById('rchContainer');if(rch)rch.style.display='none';
 let box=document.querySelector('.table-box');if(box)box.style.display='none';
 let out=document.getElementById('ncdModuleContainer');if(!out){out=document.createElement('div');out.id='ncdModuleContainer';(document.getElementById('pdfArea')||document.body).appendChild(out)}
 out.innerHTML='';out.style.display='block';
 let rows=[];for(let r=0;r<data.getNumberOfRows();r++)rows.push(row(data,r));
 let headerIndex=rows.findIndex(isHeader);if(headerIndex<0){headerIndex=rows.findIndex(r=>/subcenter/i.test(clean(r[0])))}
 if(headerIndex<0){let msg=document.createElement('div');msg.className='ncd-status-heading';msg.textContent='NCD data header नहीं मिला';out.appendChild(msg);return}
 let head=rows[headerIndex];
 let sectorStart=-1,sectorHeader=-1,facilityEnd=-1;
 for(let i=headerIndex+1;i<rows.length;i++){if(isTotal(rows[i])){if(facilityEnd<0)facilityEnd=i;else break}if(isSectorHeader(rows[i])){sectorHeader=i;sectorStart=i+1;break}}
 let facilityEndIndex=facilityEnd>=0?facilityEnd:sectorHeader>=0?sectorHeader-1:rows.length;
 let facility=rows.slice(headerIndex+1,facilityEndIndex).filter(r=>r.some(Boolean));
 let sector=[];
 if(sectorHeader>=0){for(let i=sectorStart;i<rows.length;i++){if(isTotal(rows[i]))break;if(rows[i].some(Boolean))sector.push(rows[i])}}
 if(!sector.length){
   // Fallback: locate known sector names/second copy of the NCD header.
   let second=rows.findIndex((r,i)=>i>headerIndex+1 && isSectorHeader(r));
   if(second>=0)sector=rows.slice(second+1).filter(r=>r.some(Boolean)&&!isTotal(r));
 }
 let status=document.createElement('div');status.className='ncd-status-heading';status.textContent='NCD Status 2026-27 As On Last Updated On : '+LAST_UPDATED;out.appendChild(status);
 let sec=document.createElement('div');sec.className='ncd-section';let st=document.createElement('div');st.className='ncd-section-title';st.textContent='Sector Wise';sec.appendChild(st);sec.appendChild(makeTable(head,sector));out.appendChild(sec);
 let fac=document.createElement('div');fac.className='ncd-section';let ft=document.createElement('div');ft.className='ncd-section-title';ft.textContent='Facility Wise';fac.appendChild(ft);fac.appendChild(makeTable(head,facility));out.appendChild(fac);
}
window.renderNCDTable=renderNCDTable;
})();
