/* Ayushman Shivir FINAL - 9 columns + calculated Total rows */
(function(){'use strict';
const TITLE='Ayushman Shivir Reporting FY 2026-27 till 18 August 2026';
// Google Sheet se deleted/empty last column ko completely remove kiya gaya hai.
const COLS=9;
const FH=['Sn','NIN ID','Sector','Facility Name','Target Upto 18 August 2026','Shivir Reporting','%','Total Footfall','Avg Footfall/Shivir'];
const SH=['Sn','NIN ID','Sector','No. of Facility','Target Upto 18 August 2026','Shivir Reporting','%','Total Footfall','Avg Footfall/Shivir'];
function v(x){return String(x==null?'':x).replace(/\s+/g,' ').trim()}
function num(x){const n=Number(String(x??'').replace(/,/g,''));return Number.isFinite(n)?n:0}
function getRows(d){const a=[];for(let r=0;r<d.getNumberOfRows();r++){const x=[];for(let c=0;c<Math.min(d.getNumberOfColumns(),COLS);c++)x.push(v(d.getFormattedValue(r,c)));a.push(x)}return a}
function facility(r){return /^\d+$/.test(v(r[0]))&&/^\d{8,}$/.test(v(r[1]))&&v(r[2])&&v(r[3])}
function sector(r){return /^\d+$/.test(v(r[0]))&&v(r[2])&&/^\d+$/.test(v(r[3]))}
function isTotal(r){return v(r[0]).toLowerCase()==='total'||v(r[3]).toLowerCase()==='total'||v(r[2]).toLowerCase()==='total'}
function makeTotal(rows,isSector){
  const valid=rows.filter(r=>!isTotal(r));
  const target=valid.reduce((a,r)=>a+num(r[4]),0);
  const rep=valid.reduce((a,r)=>a+num(r[5]),0);
  const foot=valid.reduce((a,r)=>a+num(r[7]),0);
  const pct=target?Math.round(rep/target*100)+'%':'0%';
  const avg=rep?Math.round(foot/rep):0;
  if(isSector){
    const facilities=valid.reduce((a,r)=>a+num(r[3]),0);
    return ['Total','','Total',String(facilities),String(target),String(rep),pct,String(foot),String(avg)];
  }
  return ['Total','','','',''+target,''+rep,pct,''+foot,''+avg];
}
function box(title,heads,data,isSector){
  const b=document.createElement('section');b.className='shivir-box';
  const h=document.createElement('h3');h.className='shivir-title';h.textContent=title;b.appendChild(h);
  const w=document.createElement('div');w.className='shivir-wrap';
  const t=document.createElement('table');t.className='shivir-table';
  const hr=document.createElement('tr');heads.slice(0,COLS).forEach(x=>{const e=document.createElement('th');e.textContent=x;hr.appendChild(e)});
  const thead=document.createElement('thead');thead.appendChild(hr);t.appendChild(thead);
  const tb=document.createElement('tbody');
  data.forEach(r=>{const tr=document.createElement('tr');for(let i=0;i<COLS;i++){const td=document.createElement('td');td.textContent=v(r[i]);if(i===6)td.className='shivir-pct';tr.appendChild(td)}if(isTotal(r))tr.className='shivir-total';tb.appendChild(tr)});
  const totalRow=makeTotal(data,isSector);const tr=document.createElement('tr');tr.className='shivir-total';for(let i=0;i<COLS;i++){const td=document.createElement('td');td.textContent=v(totalRow[i]);if(i===6)td.className='shivir-pct';tr.appendChild(td)}tb.appendChild(tr);
  t.appendChild(tb);w.appendChild(t);b.appendChild(w);return b;
}
function summary(f){let target=0,rep=0,foot=0;f.filter(r=>!isTotal(r)).forEach(r=>{target+=num(r[4]);rep+=num(r[5]);foot+=num(r[7])});const b=document.createElement('div');b.className='shivir-summary';[['Total Facility',f.filter(r=>!isTotal(r)).length],['Target Shivir',target],['Shivir Reporting',rep],['Overall Achievement',target?Math.round(rep/target*100)+'%':'0%'],['Total Footfall',foot],['Avg Footfall/Shivir',rep?Math.round(foot/rep):0]].forEach(a=>{const c=document.createElement('div');c.className='shivir-card';c.innerHTML='<span>'+a[0]+'</span><strong>'+a[1]+'</strong>';b.appendChild(c)});return b}
function render(data){const pdf=document.getElementById('pdfArea');if(!pdf||!data)return false;const all=getRows(data),f=all.filter(facility),s=all.filter(sector);if(!f.length)return false;pdf.querySelectorAll('.table-box,.shivir-final').forEach(x=>x.style.display='none');let root=pdf.querySelector('.shivir-final');if(!root){root=document.createElement('div');root.className='shivir-final';pdf.appendChild(root)}root.innerHTML='';const title=document.createElement('h2');title.className='shivir-main-heading';title.textContent=TITLE;root.appendChild(title);root.appendChild(summary(f));if(s.length)root.appendChild(box('Sector Wise',SH,s,true));root.appendChild(box('Facility Wise',FH,f,false));root.style.display='block';return true}
const old=window.renderNormalTable;window.renderNormalTable=function(data){if(render(data)){if(typeof window.setStatus==='function')window.setStatus('✓ Ayushman Shivir data loaded','success');return}if(typeof old==='function')return old.apply(this,arguments)};
const css=document.createElement('style');css.textContent='.shivir-final{display:block!important}.shivir-main-heading{text-align:center;margin:12px 0 18px;font-size:22px;font-weight:800;color:#075985}.shivir-summary{display:grid;grid-template-columns:repeat(6,minmax(120px,1fr));gap:10px;margin:0 0 18px}.shivir-card{background:#fff;border:1px solid #dbe3ec;border-radius:12px;padding:10px;text-align:center}.shivir-card span{display:block;font-size:11px;color:#64748b}.shivir-card strong{display:block;font-size:21px;color:#075985;margin-top:4px}.shivir-box{background:#fff;border-radius:12px;padding:10px;margin-bottom:18px}.shivir-title{text-align:center;color:#075985;font-size:18px;font-weight:800;margin:5px}.shivir-wrap{width:100%;overflow-x:auto}.shivir-table{width:100%;min-width:900px;border-collapse:collapse;font-size:13px}.shivir-table th{background:#075985;color:#fff;border:1px solid #cbd5e1;padding:8px;text-align:center}.shivir-table td{border:1px solid #cbd5e1;padding:7px;text-align:center;white-space:nowrap}.shivir-pct{font-weight:800}.shivir-total{font-weight:800;background:#dbeafe!important}@media(max-width:800px){.shivir-summary{grid-template-columns:repeat(2,1fr)}.shivir-main-heading{font-size:18px}}';document.head.appendChild(css);
})();