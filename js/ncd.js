/* NCD MODULE — Kharsia Health Dashboard — FY 2026-27 */
(function(){
'use strict';
const PARTS=[{key:'enrollment',title:'Part 1 — Enrollment & ABHA',terms:/enrollment|abha|30\+|population|screening/i},{key:'htn',title:'Part 2 — HTN',terms:/\bhtn\b|hypertension|hypertensive/i},{key:'dm',title:'Part 3 — DM',terms:/\bdm\b|diabetes|diabetic/i}];
const clean=v=>String(v??'').replace(/\s+/g,' ').trim();
function vals(data,r){let a=[];for(let c=0;c<data.getNumberOfColumns();c++){let v='';try{v=data.getFormattedValue(r,c)||data.getValue(r,c)||''}catch(e){}a.push(clean(v))}return a}
function txt(r){return r.map(clean).join(' ').toLowerCase()}
function pct(v){let n=parseFloat(clean(v).replace(/,/g,'').replace(/%/g,''));return n>=90?'good':n>=70?'mid':'low'}
function isTotal(r){return /^(total|योग|कुल|grand total|block total)/i.test(clean(r[0]))||/\b(total|योग|कुल)\b/i.test(txt(r))}
function findPart(r,current){let s=txt(r);if(/part\s*1|part-1|enrollment.*abha/i.test(s))return'enrollment';if(/part\s*2|part-2|\bhtn\b|hypertension/i.test(s))return'htn';if(/part\s*3|part-3|\bdm\b|diabetes/i.test(s))return'dm';return current}
function findHeader(rows){
  // Google Sheet often has title/merged rows before the actual column headings.
  // Find the first row containing several known NCD column labels.
  for(let i=0;i<Math.min(rows.length,20);i++){
    let s=txt(rows[i]), n=rows[i].filter(Boolean).length;
    if(n>=2 && /sector|facility|30\+|population|screening|enrollment|abha|htn|hypertension|estimated|under treatment|treatment|follow|control|diabetes|dm/i.test(s)) return i;
  }
  return -1;
}
function header(rows,hi){
  if(hi<0)return [];
  // Preserve every real header cell; never generate Column N.
  return rows[hi].map(clean);
}
function table(headerRow,dataRows){
 const t=document.createElement('table');t.className='ncd-part-table';
 const thd=document.createElement('thead'),trh=document.createElement('tr');
 headerRow.forEach(v=>{let th=document.createElement('th');th.textContent=v;trh.appendChild(th)});thd.appendChild(trh);t.appendChild(thd);
 const bd=document.createElement('tbody');
 dataRows.forEach(r=>{if(!r.some(Boolean))return;let tr=document.createElement('tr');if(isTotal(r))tr.className='ncd-total-row';r.forEach(v=>{let td=document.createElement('td');td.textContent=v;if(/%$/.test(v))td.classList.add('ncd-percent-'+pct(v));tr.appendChild(td)});bd.appendChild(tr)});t.appendChild(bd);return t;
}
function section(rows,mode){let out=[],state='';for(const r of rows){let s=txt(r);if(/\bsector\b|सेक्टर/i.test(s)){state='sector';continue}if(/\bfacility\b|facility wise|sub.?centre|phc|shc|hwc|uphc|आरोग्य|उप.?केंद्र/i.test(s)){state='facility';continue}if(state===mode)out.push(r)}return out}
function css(){if(document.getElementById('ncd-style'))return;let s=document.createElement('style');s.id='ncd-style';s.textContent=`#ncdModuleContainer{display:none}.ncd-part{background:#fff;padding:12px;margin:0 0 18px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.08);overflow:auto}.ncd-part-title{text-align:center;font-size:18px;font-weight:900;color:#075985;margin:5px}.ncd-section-title{font-size:15px;font-weight:900;color:#075985;background:#e0f2fe;border-left:5px solid #075985;padding:8px 10px;margin:12px 0 7px}.ncd-part-table{width:100%;border-collapse:collapse;font-size:12px}.ncd-part-table th{background:#075985;color:#fff;border:1px solid #cbd5e1;padding:8px 6px;text-align:center}.ncd-part-table td{border:1px solid #cbd5e1;padding:7px 6px;text-align:center}.ncd-part-table td:first-child{text-align:left;font-weight:700}.ncd-total-row{background:#dbeafe!important;font-weight:900}.ncd-percent-good{background:#dcfce7;color:#166534;font-weight:900}.ncd-percent-mid{background:#fef3c7;color:#92400e;font-weight:900}.ncd-percent-low{background:#fecaca;color:#991b1b;font-weight:900}@media(max-width:800px){.ncd-part-table{min-width:900px}}`;document.head.appendChild(s)}
function renderNCDTable(data){
 css();let rch=document.getElementById('rchContainer');if(rch)rch.style.display='none';let box=document.querySelector('.table-box');if(box)box.style.display='none';let out=document.getElementById('ncdModuleContainer');if(!out){out=document.createElement('div');out.id='ncdModuleContainer';let area=document.getElementById('pdfArea');(area||document.body).appendChild(out)}out.innerHTML='';out.style.display='block';
 let groups=PARTS.map(p=>({key:p.key,title:p.title,rows:[]})),current=null;for(let r=0;r<data.getNumberOfRows();r++){let x=vals(data,r),p=findPart(x,current);if(p)current=p;if(current)groups.find(g=>g.key===current).rows.push(x)}
 groups.filter(g=>g.rows.length).forEach(g=>{let p=document.createElement('div');p.className='ncd-part';let h=document.createElement('div');h.className='ncd-part-title';h.textContent=g.title;p.appendChild(h);let hi=findHeader(g.rows),hdr=header(g.rows,hi),body=hi>=0?g.rows.slice(hi+1):g.rows;let sec=section(body,'sector'),fac=section(body,'facility');let a=document.createElement('div');a.className='ncd-section-title';a.textContent='Sector Wise';p.appendChild(a);p.appendChild(table(hdr,sec));let b=document.createElement('div');b.className='ncd-section-title';b.textContent='Facility Wise';p.appendChild(b);p.appendChild(table(hdr,fac));out.appendChild(p)});
}
window.renderNCDTable=renderNCDTable;
})();
