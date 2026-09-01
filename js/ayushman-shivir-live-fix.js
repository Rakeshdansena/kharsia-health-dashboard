/* Ayushman Shivir — dedicated renderer
   Keeps the Shivir report independent from JAS/HWC/Wellness/RCH/Ayushman Card. */
(function(){
  'use strict';

  const SHIVIR_NAME = 'Ayushman Shivir';
  const SHIVIR_GID = '1262815420';

  function txt(v){ return String(v == null ? '' : v).replace(/\s+/g,' ').trim(); }
  function rowValues(data,r){
    const a=[];
    for(let c=0;c<data.getNumberOfColumns();c++){
      try { a.push(txt(data.getFormattedValue(r,c))); }
      catch(e){ a.push(''); }
    }
    return a;
  }
  function isHeader(row){
    const s=row.join(' ').toLowerCase();
    return s.includes('nin') && s.includes('sector') && (s.includes('shivir') || s.includes('target'));
  }
  function isData(row){ return /^\d+$/.test(txt(row[0])); }
  function isTitle(row){
    const s=row.join(' ').toLowerCase();
    return s.includes('ayushman shivir reporting');
  }

  function makeCell(tag,value,cls){
    const el=document.createElement(tag);
    el.textContent=value;
    if(cls) el.className=cls;
    return el;
  }

  function renderSection(tableRows,titleText,headers){
    const box=document.createElement('div');
    box.className='shivir-report-box';

    const title=document.createElement('div');
    title.className='shivir-section-title';
    title.textContent=titleText || 'Ayushman Shivir Reporting FY 2026-27';
    box.appendChild(title);

    const wrap=document.createElement('div');
    wrap.className='shivir-table-wrap';
    const table=document.createElement('table');
    table.className='shivir-table';
    const thead=document.createElement('thead');
    const hr=document.createElement('tr');
    headers.forEach(h=>hr.appendChild(makeCell('th',h)));
    thead.appendChild(hr); table.appendChild(thead);

    const tbody=document.createElement('tbody');
    tableRows.forEach(row=>{
      const tr=document.createElement('tr');
      row.forEach((v,i)=>{
        const td=makeCell('td',v);
        if(i===6 && /%/.test(v)) td.classList.add('shivir-percent');
        tr.appendChild(td);
      });
      if(txt(row[0]).toLowerCase()==='total') tr.classList.add('shivir-total');
      tbody.appendChild(tr);
    });
    table.appendChild(tbody); wrap.appendChild(table); box.appendChild(wrap);
    return box;
  }

  function renderShivir(data){
    const table=document.getElementById('reportTable');
    const tableBox=document.querySelector('.table-box');
    if(!table || !tableBox) return false;

    const rows=[];
    for(let r=0;r<data.getNumberOfRows();r++) rows.push(rowValues(data,r));

    let firstHeader=-1, secondHeader=-1;
    rows.forEach((row,i)=>{
      if(isHeader(row)) {
        if(firstHeader<0) firstHeader=i; else if(secondHeader<0) secondHeader=i;
      }
    });
    if(firstHeader<0) return false;

    const defaultHeaders=[
      'Sn','NIN ID','Sector','Facility Name','Target Upto 18 August 2026',
      'Shivir Reporting','%','Total Footfall','Avg Footfall/Shivir','Total Footfall','Avg Footfall'
    ];
    const sectorHeaders=['Sn','NIN ID','Sector','No. of Facility','Target Upto 18 August 2026','Shivir Reporting','%','Total Footfall','Avg Footfall/Shivir','Total Footfall','Avg Footfall'];

    const firstData=[];
    const secondData=[];
    let firstTitle='Ayushman Shivir Reporting FY 2026-27 till 18 August 2026';
    let secondTitle=firstTitle;

    for(let i=0;i<rows.length;i++){
      if(isTitle(rows[i])){
        if(i<=firstHeader) firstTitle=txt(rows[i].join(' ')) || firstTitle;
        else secondTitle=txt(rows[i].join(' ')) || secondTitle;
      }
    }

    const firstEnd=secondHeader>firstHeader?secondHeader:rows.length;
    for(let i=firstHeader+1;i<firstEnd;i++) if(isData(rows[i]) || txt(rows[i][0]).toLowerCase()==='total') firstData.push(rows[i].slice(0,11));
    if(secondHeader>=0){
      for(let i=secondHeader+1;i<rows.length;i++) if(isData(rows[i]) || txt(rows[i][0]).toLowerCase()==='total') secondData.push(rows[i].slice(0,11));
    }

    const pdfArea=document.getElementById('pdfArea');
    const old=pdfArea.querySelector('.shivir-container');
    if(old) old.remove();
    const container=document.createElement('div');
    container.className='shivir-container';
    container.appendChild(renderSection(firstData,firstTitle,defaultHeaders));
    if(secondHeader>=0 && secondData.length) container.appendChild(renderSection(secondData,secondTitle,sectorHeaders));

    tableBox.style.display='none';
    pdfArea.appendChild(container);
    return true;
  }

  function hook(){
    if(!window.REPORTS || !window.loadGoogleSheet) return;
    if(window.__shivirHooked) return;
    window.__shivirHooked=true;

    const original=window.loadGoogleSheet;
    window.loadGoogleSheet=function(gid){
      if(String(gid)!==SHIVIR_GID) return original.apply(this,arguments);
      const url='https://docs.google.com/spreadsheets/d/'+window.SPREADSHEET_ID+'/gviz/tq?gid='+encodeURIComponent(gid)+'&headers=0';
      const query=new google.visualization.Query(url);
      query.setQuery('select *');
      query.send(function(response){
        if(response.isError()){
          if(window.setStatus) setStatus('Google Sheet से Ayushman Shivir data load नहीं हो पाया: '+response.getMessage(),'error');
          return;
        }
        window.currentData=response.getDataTable();
        const ok=renderShivir(window.currentData);
        if(!ok && window.renderNormalTable) window.renderNormalTable(window.currentData);
        if(window.setStatus) setStatus('✓ Ayushman Shivir data successfully loaded','success');
      });
    };
  }

  const css=document.createElement('style');
  css.id='ayushman-shivir-fix-css';
  css.textContent=`
    .shivir-container{display:block;margin-top:0}
    .shivir-report-box{background:#fff;border-radius:14px;box-shadow:0 3px 12px rgba(0,0,0,.07);padding:12px;margin-bottom:18px}
    .shivir-section-title{text-align:center;color:#075985;font-size:18px;font-weight:800;padding:10px 6px 14px;line-height:1.35}
    .shivir-table-wrap{overflow:auto}
    .shivir-table{width:100%;min-width:1100px;border-collapse:collapse;font-size:13px}
    .shivir-table th{background:#075985;color:#fff;border:1px solid #cbd5e1;padding:9px 7px;text-align:center;vertical-align:middle;white-space:normal}
    .shivir-table td{border:1px solid #cbd5e1;padding:7px 6px;text-align:center;vertical-align:middle;white-space:nowrap}
    .shivir-table tbody tr:nth-child(even){background:#f8fafc}
    .shivir-table tbody tr:hover{background:#e0f2fe}
    .shivir-percent{font-weight:800;background:#dcfce7;color:#166534}
    .shivir-total{font-weight:800;background:#dbeafe!important}
    @media(max-width:800px){.shivir-table{min-width:1100px}.shivir-section-title{font-size:16px}}
    @media print{.shivir-report-box{box-shadow:none;border-radius:0;margin:0 0 8mm;padding:0}.shivir-section-title{font-size:13px;padding:3mm}.shivir-table{min-width:0;width:100%;font-size:7px}.shivir-table th{padding:3px}.shivir-table td{padding:2px;white-space:normal}.shivir-report-box{page-break-after:always}.shivir-report-box:last-child{page-break-after:auto}}
    .pdf-mode .shivir-report-box{box-shadow:none;border-radius:0;margin:0;padding:0 0 10px}.pdf-mode .shivir-section-title{font-size:14px}.pdf-mode .shivir-table{min-width:0;width:100%;font-size:7px}.pdf-mode .shivir-table th{padding:3px}.pdf-mode .shivir-table td{padding:2px;white-space:normal}
  `;
  document.head.appendChild(css);

  const timer=setInterval(hook,300);
  setTimeout(()=>clearInterval(timer),15000);
})();
