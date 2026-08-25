/* NCD MODULE — exact Google Sheet structure */
(function(){
'use strict';
const LAST_UPDATED='17-Aug-2026 23:06:53';
const clean=v=>String(v??'').replace(/\s+/g,' ').trim();
function getRows(data){const rows=[];for(let r=0;r<data.getNumberOfRows();r++){const a=[];for(let c=0;c<data.getNumberOfColumns();c++){let v='';try{v=data.getFormattedValue(r,c)||data.getValue(r,c)||''}catch(e){}a.push(clean(v))}rows.push(a)}return rows}
function isTitle(r){return /NCD\s*status\s*2026-27/i.test(r.join(' '))}
function isTotal(r){return /^(total|योग|कुल)$/i.test(clean(r[0]))}
function isHeader(r){const s=r.join(' ').toLowerCase();return /subcenter|sector|total\s+population|screening\s+target|enrollment\s*30|abha\s*link/.test(s)}
function makeTable(header,rows){const table=document.createElement('table');table.className='ncd-exact-table';const thead=document.createElement('thead');const trh=document.createElement('tr');header.forEach(v=>{const th=document.createElement('th');th.textContent=v||'';trh.appendChild(th)});thead.appendChild(trh);table.appendChild(thead);const tbody=document.createElement('tbody');rows.forEach(r=>{if(!r.some(Boolean))return;const tr=document.createElement('tr');if(isTotal(r))tr.className='ncd-total-row';for(let c=0;c<header.length;c++){const td=document.createElement('td');td.textContent=r[c]||'';tr.appendChild(td)}tbody.appendChild(tr)});table.appendChild(tbody);return table}
function addSection(out,title,header,rows){const box=document.createElement('div');box.className='ncd-exact-section';const h=document.createElement('div');h.className='ncd-exact-section-title';h.textContent=title;box.appendChild(h);box.appendChild(makeTable(header,rows));out.appendChild(box)}
function css(){if(document.getElementById('ncd-exact-style'))return;const s=document.createElement('style');s.id='ncd-exact-style';s.textContent=`#ncdModuleContainer{display:block!important;width:100%;box-sizing:border-box}.ncd-exact-heading{width:100%;box-sizing:border-box;text-align:center;font-weight:900;font-size:17px;color:#075985;background:#e0f2fe;border:1px solid #93c5fd;border-radius:7px;padding:11px 8px;margin:0 0 14px}.ncd-exact-section{background:#fff;border-radius:10px;box-shadow:0 3px 12px rgba(0,0,0,.07);padding:10px;margin:0 0 18px;overflow:auto}.ncd-exact-section-title{text-align:center;font-size:18px;font-weight:900;color:#075985;background:#e0f2fe;border-left:5px solid #075985;padding:9px;margin:0 0 9px}.ncd-exact-table{border-collapse:collapse;width:100%;min-width:1500px;font-size:12px;table-layout:auto}.ncd-exact-table th{background:#075985;color:#fff;border:1px solid #cbd5e1;padding:8px 6px;text-align:center;vertical-align:middle;white-space:normal;font-weight:800}.ncd-exact-table td{border:1px solid #cbd5e1;padding:7px 6px;text-align:center;vertical-align:middle;white-space:normal}.ncd-exact-table td:first-child{text-align:left;font-weight:700;min-width:180px}.ncd-exact-table tbody tr:nth-child(even){background:#f8fafc}.ncd-exact-table .ncd-total-row{background:#dbeafe!important;font-weight:900}.ncd-exact-table .ncd-total-row td{font-weight:900}@media(max-width:800px){.ncd-exact-table{min-width:1500px}.ncd-exact-heading{font-size:15px}}`;document.head.appendChild(s)}
function renderNCDTable(data){
 css();
 const old=document.getElementById('ncdModuleContainer');if(old)old.remove();
 const tableBox=document.querySelector('.table-box');if(tableBox)tableBox.style.display='none';
 const rch=document.getElementById('rchContainer');if(rch)rch.style.display='none';
 const out=document.createElement('div');out.id='ncdModuleContainer';const pdf=document.getElementById('pdfArea');(pdf||document.body).appendChild(out);
 const rows=getRows(data);
 let titles=[];rows.forEach((r,i)=>{if(isTitle(r))titles.push(i)});
 const firstTitle=titles[0]??-1;const secondTitle=titles[1]??-1;
 function nextHeader(start,end){for(let i=Math.max(0,start);i<(end<0?rows.length:end);i++){if(isHeader(rows[i]))return i}return -1}
 const facilityHeader=nextHeader(firstTitle+1,secondTitle);
 const sectorHeader=nextHeader(secondTitle+1,rows.length);
 function endAtTotal(start,end){for(let i=start;i<(end<0?rows.length:end);i++){if(isTotal(rows[i]))return i}return end<0?rows.length:end}
 const facilityEnd=endAtTotal(facilityHeader>=0?facilityHeader+1:0,secondTitle);
 const sectorEnd=endAtTotal(sectorHeader>=0?sectorHeader+1:0,rows.length);
 const facilityHead=facilityHeader>=0?rows[facilityHeader]:[];
 const sectorHead=sectorHeader>=0?rows[sectorHeader]:facilityHead;
 const facilityRows=facilityHeader>=0?rows.slice(facilityHeader+1,facilityEnd):[];
 const sectorRows=sectorHeader>=0?rows.slice(sectorHeader+1,sectorEnd):[];
 const heading=document.createElement('div');heading.className='ncd-exact-heading';heading.textContent='NCD Status 2026-27 As On Last Updated On : '+LAST_UPDATED;out.appendChild(heading);
 addSection(out,'Sector Wise',sectorHead,sectorRows);
 addSection(out,'Facility Wise',facilityHead,facilityRows);
}
window.renderNCDTable=renderNCDTable;
})();
