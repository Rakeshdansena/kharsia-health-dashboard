/* Ayushman Shivir — fixed renderer
   Reads the existing Shivir sheet directly and never uses the generic renderer. */
(function(){
  'use strict';
  const GID='1262815420';
  const DEFAULT_TITLE='Ayushman Shivir Reporting FY 2026-27 till 18 August 2026';
  function t(v){return String(v==null?'':v).replace(/\s+/g,' ').trim();}
  function vals(data,r){const a=[];for(let c=0;c<data.getNumberOfColumns();c++){try{a.push(t(data.getFormattedValue(r,c)));}catch(e){a.push('');}}return a;}
  function isTitle(r){return r.join(' ').toLowerCase().includes('ayushman shivir reporting');}
  function sn(r){for(let i=0;i<Math.min(4,r.length);i++){const v=t(r[i]);if(/^\d+$/.test(v))return v;}return '';}
  function isData(r){return !!sn(r) && !!t(r[2]) && !!t(r[3]);}
  function isTotal(r){return t(r[0]).toLowerCase()==='total' || r.some(v=>t(v).toLowerCase()==='total');}
  function normalized(r){
    const n=sn(r);
    let start=0;
    if(n && t(r[0])!==n){for(let i=0;i<4;i++)if(t(r[i])===n){start=i;break;}}
    const out=r.slice(start,start+11);
    while(out.length<11)out.push('');
    out[0]=n;
    return out;
  }
  function cell(tag,v){const e=document.createElement(tag);e.textContent=v;return e;}
  function makeSection(rows,title,sectorWise){
    const box=document.createElement('div');box.className='shivir-report-box';
    const h=document.createElement('div');h.className='shivir-section-title';h.textContent=title||DEFAULT_TITLE;box.appendChild(h);
    const wrap=document.createElement('div');wrap.className='shivir-table-wrap';
    const table=document.createElement('table');table.className='shivir-table';
    const headers=sectorWise?['Sn','NIN ID','Sector','No. of Facility','Target Upto 18 August 2026','Shivir Reporting','%','Total Footfall','Avg Footfall/Shivir','Total Footfall','Avg Footfall']:['Sn','NIN ID','Sector','Facility Name','Target Upto 18 August 2026','Shivir Reporting','%','Total Footfall','Avg Footfall/Shivir','Total Footfall','Avg Footfall'];
    const thead=document.createElement('thead'),hr=document.createElement('tr');headers.forEach(x=>hr.appendChild(cell('th',x)));thead.appendChild(hr);table.appendChild(thead);
    const tbody=document.createElement('tbody');
    rows.forEach(r=>{const tr=document.createElement('tr');const x=isTotal(r)?r:normalized(r);x.slice(0,11).forEach((v,i)=>{const td=cell('td',v);if(i===6&&/%/.test(v))td.className='shivir-percent';tr.appendChild(td);});if(isTotal(r))tr.className='shivir-total';tbody.appendChild(tr);});
    table.appendChild(tbody);wrap.appendChild(table);box.appendChild(wrap);return box;
  }
  function render(data){
    const pdf=document.getElementById('pdfArea');if(!pdf)return false;
    const rows=[];for(let r=0;r<data.getNumberOfRows();r++)rows.push(vals(data,r));
    const first=[],second=[];let section=0;
    rows.forEach(r=>{if(isTitle(r)){section++;return;}if(section===1&&(isData(r)||isTotal(r)))first.push(r);else if(section>=2&&(isData(r)||isTotal(r)))second.push(r);});
    if(!first.length)return false;
    const old=pdf.querySelector('.shivir-container');if(old)old.remove();
    const container=document.createElement('div');container.className='shivir-container';
    container.appendChild(makeSection(first,DEFAULT_TITLE,false));
    if(second.length)container.appendChild(makeSection(second,DEFAULT_TITLE,true));
    const tb=document.querySelector('.table-box');if(tb)tb.style.display='none';
    pdf.appendChild(container);return true;
  }
  function load(){
    const title=document.getElementById('reportTitle');if(!title||!/Ayushman Shivir/i.test(title.textContent||''))return;
    const pdf=document.getElementById('pdfArea');if(!pdf||pdf.dataset.shivirLoading==='1'||pdf.querySelector('.shivir-container'))return;
    if(!(window.google&&google.visualization))return;
    pdf.dataset.shivirLoading='1';
    const url='https://docs.google.com/spreadsheets/d/'+window.SPREADSHEET_ID+'/gviz/tq?gid='+GID+'&headers=0';
    const q=new google.visualization.Query(url);q.setQuery('select *');
    q.send(function(res){pdf.dataset.shivirLoading='0';if(res.isError()){if(window.setStatus)setStatus('Google Sheet से Ayushman Shivir data load नहीं हो पाया: '+res.getMessage(),'error');return;}if(!render(res.getDataTable())&&window.setStatus)setStatus('Ayushman Shivir data rows नहीं मिलीं — Sheet structure check करें','error');else if(window.setStatus)setStatus('✓ Ayushman Shivir data successfully loaded','success');});
  }
  const css=document.createElement('style');css.id='ayushman-shivir-fix-css';css.textContent=`
    .shivir-container{display:block!important;margin-top:0}.shivir-report-box{background:#fff;border-radius:14px;box-shadow:0 3px 12px rgba(0,0,0,.07);padding:12px;margin-bottom:18px}.shivir-section-title{text-align:center;color:#075985;font-size:18px;font-weight:800;padding:10px 6px 14px;line-height:1.35}.shivir-table-wrap{width:100%;overflow-x:auto}.shivir-table{width:100%;min-width:1100px;border-collapse:collapse;font-size:13px}.shivir-table th{background:#075985;color:#fff;border:1px solid #cbd5e1;padding:9px 7px;text-align:center;vertical-align:middle;white-space:normal}.shivir-table td{border:1px solid #cbd5e1;padding:7px 6px;text-align:center;vertical-align:middle;white-space:nowrap}.shivir-table tbody tr:nth-child(even){background:#f8fafc}.shivir-percent{font-weight:800;background:#dcfce7;color:#166534}.shivir-total{font-weight:800;background:#dbeafe!important}@media(max-width:800px){.shivir-table{min-width:1100px}.shivir-section-title{font-size:16px}}@media print{.shivir-report-box{box-shadow:none;border-radius:0;margin:0 0 8mm;padding:0}.shivir-section-title{font-size:13px}.shivir-table{min-width:0;width:100%;font-size:7px}.shivir-table th,.shivir-table td{padding:2px;white-space:normal}.shivir-report-box{page-break-after:always}.shivir-report-box:last-child{page-break-after:auto}}
  `;document.head.appendChild(css);
  setInterval(load,500);
})();