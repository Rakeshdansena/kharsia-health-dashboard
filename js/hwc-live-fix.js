(function(){
  'use strict';
  function isHWC(){
    const t=document.getElementById('reportTitle');
    return t && /Health\s*&\s*Wellness\s*Center/i.test(t.innerText||'');
  }
  function install(){
    if(typeof window.renderNormalTable!=='function') return false;
    if(window.__hwcFixInstalled) return true;
    window.__hwcFixInstalled=true;
    const original=window.renderNormalTable;
    window.renderNormalTable=function(data){
      if(isHWC()){ renderHWC(data); return; }
      return original(data);
    };
    return true;
  }
  function text(data,r,c){try{return data.getFormattedValue(r,c)||'';}catch(e){return '';}}
  function renderHWC(data){
    const table=document.getElementById('reportTable'); if(!table)return;
    table.className='hwc-table'; table.innerHTML='';
    const cg=document.createElement('colgroup');
    ['8%','35%','20%','20%','17%'].forEach(w=>{const c=document.createElement('col');c.style.width=w;cg.appendChild(c);});
    table.appendChild(cg);
    const thead=document.createElement('thead');
    const titleRow=document.createElement('tr'); const title=document.createElement('th');
    title.colSpan=5; title.className='hwc-title-cell';
    title.innerText='Health Wellness Center - Block Kharsia Dist - Raigarh(C.G)';
    titleRow.appendChild(title); thead.appendChild(titleRow);
    const hr=document.createElement('tr');
    ['Sn','No of Health Wellness Center','Facility type','Function','%'].forEach(x=>{const th=document.createElement('th');th.innerText=x;hr.appendChild(th);});
    thead.appendChild(hr); table.appendChild(thead);
    const tbody=document.createElement('tbody');
    for(let r=0;r<data.getNumberOfRows();r++){
      const row=[]; for(let c=0;c<Math.min(5,data.getNumberOfColumns());c++) row.push(text(data,r,c).trim());
      const joined=row.join(' ').trim(); if(!joined)continue;
      if(/Health Wellness Center/i.test(joined))continue;
      if(/^Column\s*\d+$/i.test(row[0]))continue;
      if(/^Sn$/i.test(row[0]) || /^Facility type$/i.test(row[2]))continue;
      const tr=document.createElement('tr');
      for(let c=0;c<5;c++){const td=document.createElement('td');td.innerText=row[c]||'';if(c===0)td.className='hwc-sn';if(c===4)td.className='hwc-percent';tr.appendChild(td);}
      if(/^Total$|^योग$|^कुल$/i.test(row[0]))tr.className='hwc-total-row';
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
  }
  function addStyle(){
    if(document.getElementById('hwc-live-fix-style'))return;
    const s=document.createElement('style');s.id='hwc-live-fix-style';s.textContent=`
      #reportTable.hwc-table{width:100%;min-width:0;border-collapse:collapse;table-layout:fixed;font-size:13px}
      #reportTable.hwc-table th,#reportTable.hwc-table td{border:1px solid #cbd5e1;text-align:center;vertical-align:middle;padding:8px 6px;white-space:normal}
      #reportTable.hwc-table .hwc-title-cell{background:#075985;color:#fff;font-size:18px;font-weight:800;padding:12px}
      #reportTable.hwc-table thead tr:nth-child(2) th{background:#0f766e;color:#fff;font-weight:800}
      #reportTable.hwc-table tbody tr:nth-child(even){background:#f8fafc}
      #reportTable.hwc-table tbody tr:hover{background:#e0f2fe}
      #reportTable.hwc-table .hwc-total-row{background:#dbeafe!important;font-weight:800}
      #reportTable.hwc-table .hwc-percent{font-weight:800}
      #reportTable.hwc-table .hwc-sn{font-weight:700}
    `;document.head.appendChild(s);
  }
  function boot(){addStyle();if(!install())setTimeout(boot,100);} boot();
})();