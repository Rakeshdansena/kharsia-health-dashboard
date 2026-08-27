(function(){
  'use strict';

  function install(){
    if(typeof window.renderNormalTable !== 'function') return false;
    if(window.__hwcFixInstalled) return true;
    window.__hwcFixInstalled = true;

    const originalRenderNormalTable = window.renderNormalTable;

    window.renderNormalTable = function(data){
      const report = (window.REPORTS && window.currentReportIndex !== null && window.currentReportIndex !== undefined)
        ? window.REPORTS[window.currentReportIndex] : null;

      if(!report || report.name !== 'Health & Wellness Center'){
        return originalRenderNormalTable(data);
      }

      renderHWC(data);
    };

    return true;
  }

  function text(data,r,c){
    try { return data.getFormattedValue(r,c) || ''; } catch(e){ return ''; }
  }

  function renderHWC(data){
    const table = document.getElementById('reportTable');
    if(!table) return;

    if(typeof window.clearTable === 'function') window.clearTable();
    table.className = 'hwc-table';

    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    thead.innerHTML = '';
    tbody.innerHTML = '';

    const cols = Math.min(5, data.getNumberOfColumns());

    // Merged title row: Health Wellness Center - Block Kharsia Dist - Raigarh(C.G)
    const titleRow = document.createElement('tr');
    const titleCell = document.createElement('th');
    titleCell.colSpan = 5;
    titleCell.className = 'hwc-title-cell';
    titleCell.innerText = 'Health Wellness Center - Block Kharsia Dist - Raigarh(C.G)';
    titleRow.appendChild(titleCell);
    thead.appendChild(titleRow);

    // Exact five headings requested by the user.
    const headerRow = document.createElement('tr');
    ['Sn','No of Health Wellness Center','Facility type','Function','%'].forEach(function(label){
      const th = document.createElement('th');
      th.innerText = label;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    for(let r=0; r<data.getNumberOfRows(); r++){
      const row = [];
      for(let c=0; c<cols; c++) row.push(text(data,r,c).trim());

      const joined = row.join(' ').trim();
      if(!joined) continue;
      if(/Health Wellness Center/i.test(joined)) continue;
      if(row.some(v => /^Sn$/i.test(v))) continue;
      if(row.every(v => /^Column\s*\d+$/i.test(v) || v === '')) continue;

      const tr = document.createElement('tr');
      row.forEach(function(v,c){
        const td = document.createElement('td');
        td.innerText = v;
        if(c === 0) td.className = 'hwc-sn';
        if(c === 4) td.className = 'hwc-percent';
        tr.appendChild(td);
      });

      if(/^Total$|^योग$|^कुल$/i.test(row[0])) tr.classList.add('hwc-total-row');
      tbody.appendChild(tr);
    }

    // Ensure the five visible columns are exactly five even if the sheet contains more.
    table.querySelectorAll('tr').forEach(function(tr){
      while(tr.children.length > 5) tr.removeChild(tr.lastElementChild);
    });
  }

  function addStyle(){
    if(document.getElementById('hwc-live-fix-style')) return;
    const style = document.createElement('style');
    style.id = 'hwc-live-fix-style';
    style.textContent = `
      #reportTable.hwc-table{width:100%;min-width:0;border-collapse:collapse;table-layout:fixed;font-size:13px}
      #reportTable.hwc-table th,#reportTable.hwc-table td{border:1px solid #cbd5e1;text-align:center;vertical-align:middle;padding:8px 6px;white-space:normal}
      #reportTable.hwc-table .hwc-title-cell{background:#075985;color:#fff;font-size:18px;font-weight:800;padding:12px}
      #reportTable.hwc-table thead tr:nth-child(2) th{background:#0f766e;color:#fff;font-weight:800}
      #reportTable.hwc-table tbody tr:nth-child(even){background:#f8fafc}
      #reportTable.hwc-table tbody tr:hover{background:#e0f2fe}
      #reportTable.hwc-table .hwc-total-row{background:#dbeafe!important;font-weight:800}
      #reportTable.hwc-table .hwc-percent{font-weight:800}
      #reportTable.hwc-table .hwc-sn{font-weight:700}
      #reportTable.hwc-table th:nth-child(1),#reportTable.hwc-table td:nth-child(1){width:8%}
      #reportTable.hwc-table th:nth-child(2),#reportTable.hwc-table td:nth-child(2){width:35%}
      #reportTable.hwc-table th:nth-child(3),#reportTable.hwc-table td:nth-child(3){width:20%}
      #reportTable.hwc-table th:nth-child(4),#reportTable.hwc-table td:nth-child(4){width:20%}
      #reportTable.hwc-table th:nth-child(5),#reportTable.hwc-table td:nth-child(5){width:17%}
    `;
    document.head.appendChild(style);
  }

  function boot(){
    addStyle();
    if(!install()) setTimeout(boot,100);
  }

  boot();
})();