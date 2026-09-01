/* Ayushman Shivir — safe renderer
   Important: this file NEVER replaces openReport/loadGoogleSheet.
   It only takes over the generic table renderer when Shivir is active. */
(function(){
  'use strict';

  const GID='1262815420';
  const TITLE='Ayushman Shivir Reporting FY 2026-27 till 18 August 2026';
  const H1=['Sn','NIN ID','Sector','Facility Name','Target Upto 18 August 2026','Shivir Reporting','%','Total Footfall','Avg Footfall/Shivir','Total Footfall','Avg Footfall'];
  const H2=['Sn','NIN ID','Sector','No. of Facility','Target Upto 18 August 2026','Shivir Reporting','%','Total Footfall','Avg Footfall/Shivir','Total Footfall','Avg Footfall'];

  function active(){
    const t=document.getElementById('reportTitle');
    return !!t && /Ayushman Shivir/i.test(t.textContent||'');
  }
  function text(v){return String(v==null?'':v).replace(/\s+/g,' ').trim();}
  function row(data,r){
    const a=[];
    for(let c=0;c<data.getNumberOfColumns();c++) a.push(text(data.getFormattedValue(r,c)));
    return a;
  }
  function isTitle(r){return r.join(' ').toLowerCase().includes('ayushman shivir reporting');}
  function sn(r){
    for(let i=0;i<5&&i<r.length;i++) if(/^\d+$/.test(text(r[i]))) return text(r[i]);
    return '';
  }
  function dataRow(r){return !!sn(r)&&!!text(r[2])&&!!text(r[3]);}
  function totalRow(r){return r.some(v=>text(v).toLowerCase()==='total');}
  function normalize(r){
    const n=sn(r); let s=0;
    for(let i=0;i<5&&i<r.length;i++){if(text(r[i])===n){s=i;break;}}
    const out=r.slice(s,s+11);
    while(out.length<11) out.push('');
    out[0]=n;
    return out;
  }
  function td(tag,v){const e=document.createElement(tag);e.textContent=v;return e;}

  function makeTable(rows,sectorWise){
    const box=document.createElement('div'); box.className='shivir-report-box';
    const title=document.createElement('div'); title.className='shivir-section-title'; title.textContent=TITLE; box.appendChild(title);
    const wrap=document.createElement('div'); wrap.className='shivir-table-wrap';
    const table=document.createElement('table'); table.className='shivir-table';
    const headers=sectorWise?H2:H1;
    const thead=document.createElement('thead'), hr=document.createElement('tr');
    headers.forEach(h=>hr.appendChild(td('th',h))); thead.appendChild(hr); table.appendChild(thead);
    const tbody=document.createElement('tbody');
    rows.forEach(r=>{
      const tr=document.createElement('tr');
      const vals=totalRow(r)?r:normalize(r);
      for(let i=0;i<11;i++){
        const cell=td('td',text(vals[i]));
        if(i===6 && text(vals[i]).includes('%')) cell.className='shivir-percent';
        tr.appendChild(cell);
      }
      if(totalRow(r)) tr.className='shivir-total';
      tbody.appendChild(tr);
    });
    table.appendChild(tbody); wrap.appendChild(table); box.appendChild(wrap);
    return box;
  }

  function render(data){
    if(!active() || !data) return false;
    const all=[];
    for(let r=0;r<data.getNumberOfRows();r++) all.push(row(data,r));
    let section=0, facility=[], sector=[];
    all.forEach(r=>{
      if(isTitle(r)){section++;return;}
      if(section===1 && (dataRow(r)||totalRow(r))) facility.push(r);
      else if(section>=2 && (dataRow(r)||totalRow(r))) sector.push(r);
    });
    if(!facility.length) return false;

    const pdf=document.getElementById('pdfArea');
    if(!pdf) return false;
    const generic=pdf.querySelector('.table-box');
    if(generic) generic.style.display='none';
    const old=pdf.querySelector('.shivir-container');
    if(old) old.remove();
    const container=document.createElement('div'); container.className='shivir-container';
    container.appendChild(makeTable(facility,false));
    if(sector.length) container.appendChild(makeTable(sector,true));
    pdf.appendChild(container);
    return true;
  }

  /* Take over ONLY the generic renderer while Shivir is selected. */
  const original=window.renderNormalTable;
  window.renderNormalTable=function(data){
    if(active() && render(data)){
      if(typeof window.setStatus==='function') window.setStatus('✓ Ayushman Shivir data successfully loaded','success');
      return;
    }
    return original.apply(this,arguments);
  };

  /* When switching away, restore the normal table visibility. */
  setInterval(function(){
    if(active()){
      const b=document.querySelector('#pdfArea .table-box');
      if(b) b.style.display='none';
    }
  },700);

  const css=document.createElement('style');
  css.id='ayushman-shivir-safe-css';
  css.textContent=''+
  '.shivir-container{display:block!important;margin-top:0}'+
  '.shivir-report-box{background:#fff;border-radius:14px;box-shadow:0 3px 12px rgba(0,0,0,.07);padding:12px;margin-bottom:18px}'+
  '.shivir-section-title{text-align:center;color:#075985;font-size:18px;font-weight:800;padding:10px 6px 14px;line-height:1.35}'+
  '.shivir-table-wrap{width:100%;overflow-x:auto}'+
  '.shivir-table{width:100%;min-width:1100px;border-collapse:collapse;font-size:13px}'+
  '.shivir-table th{background:#075985;color:#fff;border:1px solid #cbd5e1;padding:9px 7px;text-align:center;vertical-align:middle;white-space:normal}'+
  '.shivir-table td{border:1px solid #cbd5e1;padding:7px 6px;text-align:center;vertical-align:middle;white-space:nowrap}'+
  '.shivir-table tbody tr:nth-child(even){background:#f8fafc}'+
  '.shivir-percent{font-weight:800;background:#dcfce7;color:#166534}'+
  '.shivir-total{font-weight:800;background:#dbeafe!important}'+
  '@media(max-width:800px){.shivir-table{min-width:1100px}.shivir-section-title{font-size:16px}}';
  document.head.appendChild(css);
})();