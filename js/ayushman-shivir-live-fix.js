/* Ayushman Shivir — dedicated renderer
   Summary -> Sector Wise -> Facility Wise.
   It does not replace the dashboard's Google Sheet loader. */
(function(){
  'use strict';
  const TITLE='Ayushman Shivir Reporting FY 2026-27 till 18 August 2026';
  const H1=['Sn','NIN ID','Sector','Facility Name','Target Upto 18 August 2026','Shivir Reporting','%','Total Footfall','Avg Footfall/Shivir','Total Footfall','Avg Footfall'];
  const H2=['Sn','NIN ID','Sector','No. of Facility','Target Upto 18 August 2026','Shivir Reporting','%','Total Footfall','Avg Footfall/Shivir','Total Footfall','Avg Footfall'];

  function active(){const t=document.getElementById('reportTitle');return !!t&&/Ayushman Shivir/i.test(t.textContent||'');}
  function text(v){return String(v==null?'':v).replace(/\s+/g,' ').trim();}
  function getRows(data){const a=[];for(let r=0;r<data.getNumberOfRows();r++){const x=[];for(let c=0;c<data.getNumberOfColumns();c++)x.push(text(data.getFormattedValue(r,c)));a.push(x);}return a;}
  function isTitle(r){return r.join(' ').toLowerCase().includes('ayushman shivir reporting');}
  function sn(r){for(let i=0;i<Math.min(5,r.length);i++)if(/^\d+$/.test(text(r[i])))return text(r[i]);return '';}
  function isData(r){return !!sn(r)&&!!text(r[2])&&!!text(r[3]);}
  function isTotal(r){return r.some(v=>/^total$/i.test(text(v)));}
  function normalize(r){
    const n=sn(r);let start=0;
    for(let i=0;i<Math.min(5,r.length);i++){if(text(r[i])===n){start=i;break;}}
    const out=r.slice(start,start+11);while(out.length<11)out.push('');out[0]=n;return out;
  }
  function cell(tag,v){const e=document.createElement(tag);e.textContent=text(v);return e;}

  function makeTable(rows,sectorWise){
    const box=document.createElement('div');box.className='shivir-report-box';
    const title=document.createElement('div');title.className='shivir-section-title';title.textContent=sectorWise?'Sector Wise': 'Facility Wise';box.appendChild(title);
    const wrap=document.createElement('div');wrap.className='shivir-table-wrap';
    const table=document.createElement('table');table.className='shivir-table';
    const headers=sectorWise?H2:H1;const thead=document.createElement('thead');const hr=document.createElement('tr');
    headers.forEach(h=>hr.appendChild(cell('th',h)));thead.appendChild(hr);table.appendChild(thead);
    const tbody=document.createElement('tbody');
    rows.forEach(r=>{const tr=document.createElement('tr');const vals=isTotal(r)?r:normalize(r);for(let i=0;i<11;i++){const td=cell('td',vals[i]);if(i===6&&/%/.test(text(vals[i])))td.className='shivir-percent';tr.appendChild(td);}if(isTotal(r))tr.className='shivir-total';tbody.appendChild(tr);});
    table.appendChild(tbody);wrap.appendChild(table);box.appendChild(wrap);return box;
  }

  function makeSummary(facility,sector){
    let total=facility.find(isTotal)||sector.find(isTotal)||null;
    let target='',report='',pct='',foot='',avg='';
    if(total){
      const v=normalize(total); target=v[4];report=v[5];pct=v[6];foot=v[7];avg=v[8];
      if(!target&&sector.length){target=String(sector.filter(x=>!isTotal(x)).reduce((s,r)=>s+(parseFloat(normalize(r)[4])||0),0));}
    }
    const facilities=facility.filter(isData).length;
    if(!target)target=String(sector.filter(isData).reduce((s,r)=>s+(parseFloat(normalize(r)[4])||0),0));
    if(!report)report=String(sector.filter(isData).reduce((s,r)=>s+(parseFloat(normalize(r)[5])||0),0));
    if(!pct&&target)pct=Math.round((parseFloat(report||0)/parseFloat(target))*100)+'%';
    if(!foot)foot=String(sector.filter(isData).reduce((s,r)=>s+(parseFloat(normalize(r)[7])||0),0));
    if(!avg&&report)avg=(parseFloat(foot||0)/parseFloat(report||1)).toFixed(1);

    const box=document.createElement('div');box.className='shivir-summary-box';
    const title=document.createElement('div');title.className='shivir-summary-title';title.textContent='Ayushman Shivir Summary';box.appendChild(title);
    const grid=document.createElement('div');grid.className='shivir-summary-grid';
    [['Total Facility',facilities],['Target Shivir',target],['Shivir Reporting',report],['Achievement',pct],['Total Footfall',foot],['Avg Footfall/Shivir',avg]].forEach(([k,v])=>{const c=document.createElement('div');c.className='shivir-summary-card';c.innerHTML='<div class="shivir-summary-label"></div><div class="shivir-summary-value"></div>';c.querySelector('.shivir-summary-label').textContent=k;c.querySelector('.shivir-summary-value').textContent=text(v)||'0';grid.appendChild(c);});
    box.appendChild(grid);return box;
  }

  function render(data){
    if(!active()||!data)return false;
    const all=getRows(data);let section=0,facility=[],sector=[];
    all.forEach(r=>{if(isTitle(r)){section++;return;}if(section===1&&(isData(r)||isTotal(r)))facility.push(r);if(section>=2&&(isData(r)||isTotal(r)))sector.push(r);});
    if(!facility.length&&!sector.length)return false;
    const pdf=document.getElementById('pdfArea');if(!pdf)return false;
    const generic=pdf.querySelector('.table-box');if(generic)generic.style.display='none';
    const old=pdf.querySelector('.shivir-container');if(old)old.remove();
    const container=document.createElement('div');container.className='shivir-container';
    container.appendChild(makeSummary(facility,sector));
    if(sector.length)container.appendChild(makeTable(sector,true));
    if(facility.length)container.appendChild(makeTable(facility,false));
    pdf.appendChild(container);return true;
  }

  function hook(){
    if(typeof window.renderNormalTable!=='function'){setTimeout(hook,300);return;}
    if(window.__shivirHooked)return;window.__shivirHooked=true;
    const original=window.renderNormalTable;
    window.renderNormalTable=function(data){
      if(active()){if(render(data)){if(typeof window.setStatus==='function')window.setStatus('✓ Ayushman Shivir data successfully loaded','success');return;}}
      return original.apply(this,arguments);
    };
  }
  hook();

  const css=document.createElement('style');css.id='ayushman-shivir-safe-css';css.textContent=
  '.shivir-container{display:block!important;margin-top:0}.shivir-report-box,.shivir-summary-box{background:#fff;border-radius:14px;box-shadow:0 3px 12px rgba(0,0,0,.07);padding:12px;margin-bottom:18px}.shivir-summary-title,.shivir-section-title{text-align:center;color:#075985;font-size:18px;font-weight:800;padding:8px 6px 14px}.shivir-summary-grid{display:grid;grid-template-columns:repeat(6,minmax(130px,1fr));gap:10px}.shivir-summary-card{background:#f8fafc;border:1px solid #dbe3ec;border-radius:10px;padding:12px 8px;text-align:center}.shivir-summary-label{color:#64748b;font-size:12px;font-weight:600}.shivir-summary-value{color:#075985;font-size:22px;font-weight:800;margin-top:5px}.shivir-table-wrap{width:100%;overflow-x:auto}.shivir-table{width:100%;min-width:1100px;border-collapse:collapse;font-size:13px}.shivir-table th{background:#075985;color:#fff;border:1px solid #cbd5e1;padding:9px 7px;text-align:center;vertical-align:middle;white-space:normal}.shivir-table td{border:1px solid #cbd5e1;padding:7px 6px;text-align:center;vertical-align:middle;white-space:nowrap}.shivir-table tbody tr:nth-child(even){background:#f8fafc}.shivir-percent{font-weight:800;background:#dcfce7;color:#166534}.shivir-total{font-weight:800;background:#dbeafe!important}@media(max-width:800px){.shivir-summary-grid{grid-template-columns:repeat(2,minmax(140px,1fr))}.shivir-table{min-width:1100px}.shivir-summary-title,.shivir-section-title{font-size:16px}}';document.head.appendChild(css);
})();