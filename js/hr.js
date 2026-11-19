/* Human Resource module for Kharsia Health Dashboard
   Source: Google Spreadsheet tab "Human Resource" (gid 0)
   Reads live data only; does not modify the source sheet or existing modules. */
(function () {
  'use strict';
  const HR_SHEET_ID = '16pTrmZQgc5S6-wTaa6_cKydJDHBr7d1juIdWozBqN5I';
  const HR_GID = '0';
  const HR_REFRESH_MS = 5 * 60 * 1000;
  const HR_HEADERS = [
    'Sn', 'Officer/Employee Name', 'Post', 'Posting Facility Type',
    'Posting Place', 'Mobile No', 'Job Type', 'Remarks'
  ];
  let hrRows = [];
  let hrLastUpdated = null;
  let hrLoading = false;

  const esc = value => String(value == null ? '' : value).replace(/[&<>"']/g, ch => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[ch]));
  const clean = value => String(value == null ? '' : value).trim();
  const key = value => clean(value).toLocaleLowerCase('en-IN');

  function gvizUrl() {
    return 'https://docs.google.com/spreadsheets/d/' + HR_SHEET_ID +
      '/gviz/tq?tqx=out:json&gid=' + encodeURIComponent(HR_GID) +
      '&_=' + Date.now();
  }

  async function fetchHRData() {
    const response = await fetch(gvizUrl(), { cache: 'no-store' });
    if (!response.ok) throw new Error('Google Sheet HTTP ' + response.status);
    const raw = await response.text();
    const start = raw.indexOf('{');
    const end = raw.lastIndexOf('}');
    if (start < 0 || end <= start) throw new Error('Google Sheet से मान्य डेटा नहीं मिला।');
    const parsed = JSON.parse(raw.slice(start, end + 1));
    const table = parsed.table || {};
    const labels = (table.cols || []).map(c => clean(c.label));
    const rows = (table.rows || []).map(row =>
      (row.c || []).map(cell => cell ? clean(cell.f != null ? cell.f : cell.v) : '')
    ).filter(row => row.some(clean));

    // Prefer header-name mapping, but retain the agreed eight-column order as fallback.
    const normalized = labels.map(key);
    const aliases = [
      ['sn','s no','serial no','क्रमांक'],
      ['officer/employee name','employee name','name','अधिकारी/कर्मचारी का नाम'],
      ['post','designation','पद','पदनाम'],
      ['posting facility type','facility type','संस्था प्रकार'],
      ['posting place','facility','facility name','posting facility','पदस्थापना स्थान'],
      ['mobile no','mobile number','mobile','मोबाइल नंबर'],
      ['job type','employment type','नियुक्ति प्रकार'],
      ['remarks','remark','टिप्पणी']
    ];
    const indexes = aliases.map((names, i) => {
      const found = normalized.findIndex(label => names.includes(label));
      return found >= 0 ? found : i;
    });
    if (!rows.length) return [];
    const first = rows[0].map(key);
    const hasHeader = first.some(v => aliases.some(group => group.includes(v)));
    const dataRows = hasHeader ? rows.slice(1) : rows;
    return dataRows.map((row, index) => {
      const item = {};
      HR_HEADERS.forEach((header, i) => item[header] = clean(row[indexes[i]]));
      if (!item['Sn']) item['Sn'] = String(index + 1);
      return item;
    }).filter(item => item['Officer/Employee Name'] || item.Post || item['Posting Place']);
  }

  function countBy(field) {
    const map = new Map();
    hrRows.forEach(row => {
      const label = clean(row[field]) || 'उल्लेख नहीं';
      map.set(label, (map.get(label) || 0) + 1);
    });
    return [...map.entries()].sort((a,b) => b[1]-a[1] || a[0].localeCompare(b[0], 'hi'));
  }

  function summaryTable(title, pairs) {
    const total = pairs.reduce((sum, item) => sum + item[1], 0);
    return '<div class="hr-summary-panel"><h3>' + esc(title) + '</h3>' +
      '<div class="hr-summary-scroll"><table class="hr-table hr-small-table"><thead><tr><th>' +
      esc(title.replace('Summary','').trim()) + '</th><th>कर्मचारी संख्या</th></tr></thead><tbody>' +
      pairs.map(([name, count]) => '<tr><td>' + esc(name) + '</td><td class="hr-number">' + count + '</td></tr>').join('') +
      '<tr class="hr-total"><td>Total</td><td class="hr-number">' + total + '</td></tr>' +
      '</tbody></table></div></div>';
  }

  function uniqueOptions(field) {
    return [...new Set(hrRows.map(r => clean(r[field])).filter(Boolean))]
      .sort((a,b) => a.localeCompare(b, 'hi'));
  }

  function render() {
    const root = document.getElementById('hrModule');
    if (!root) return;
    const regular = hrRows.filter(r => key(r['Job Type']).includes('regular')).length;
    const contractual = hrRows.filter(r => /contract|संविदा|कॉन्ट्रैक्ट/i.test(r['Job Type'])).length;
    const facilities = countBy('Posting Place');
    const designations = countBy('Post');
    const facilityType = uniqueOptions('Posting Facility Type');
    const posts = uniqueOptions('Post');
    const places = uniqueOptions('Posting Place');
    const jobs = uniqueOptions('Job Type');
    root.innerHTML = `
      <div class="hr-head">
        <div><div class="hr-eyebrow">BLOCK KHARSIA · HUMAN RESOURCE</div><h2>👥 Human Resource Management</h2>
        <p>कर्मचारी विवरण एवं Facility / Designation Wise Summary</p></div>
        <div class="hr-head-actions"><button class="hr-btn" id="hrRefresh">↻ Refresh</button>
        <button class="hr-btn hr-excel" id="hrExcel">⬇ Excel</button><button class="hr-btn hr-print" id="hrPrint">⎙ PDF / Print</button></div>
      </div>
      <div class="hr-updated" id="hrUpdated">डेटा लोड हो रहा है...</div>
      <div class="hr-kpis">
        <div class="hr-kpi"><span>Total Employees</span><strong>\${hrRows.length}</strong></div>
        <div class="hr-kpi"><span>Regular Employees</span><strong>\${regular}</strong></div>
        <div class="hr-kpi"><span>Contractual Employees</span><strong>\${contractual}</strong></div>
        <div class="hr-kpi"><span>Total Facilities</span><strong>\${facilities.length}</strong></div>
        <div class="hr-kpi"><span>Total Designations</span><strong>\${designations.length}</strong></div>
      </div>
      <div class="hr-summary-grid">
        \${summaryTable('Facility Wise Summary', facilities)}
        \${summaryTable('Designation Wise Summary', designations)}
      </div>
      <div class="hr-list-panel">
        <div class="hr-list-head"><div><h3>Employee Details</h3><p>Google Sheet: Human Resource</p></div>
          <input id="hrSearch" type="search" placeholder="नाम, पद, मोबाइल खोजें..." aria-label="कर्मचारी खोजें"></div>
        <div class="hr-filters">
          <select id="hrFacilityFilter"><option value="">सभी Posting Place</option>\${places.map(v=>'<option value="'+esc(v)+'">'+esc(v)+'</option>').join('')}</select>
          <select id="hrPostFilter"><option value="">सभी Designation</option>\${posts.map(v=>'<option value="'+esc(v)+'">'+esc(v)+'</option>').join('')}</select>
          <select id="hrTypeFilter"><option value="">सभी Facility Type</option>\${facilityType.map(v=>'<option value="'+esc(v)+'">'+esc(v)+'</option>').join('')}</select>
          <select id="hrJobFilter"><option value="">सभी Job Type</option>\${jobs.map(v=>'<option value="'+esc(v)+'">'+esc(v)+'</option>').join('')}</select>
        </div>
        <div class="hr-table-wrap"><table class="hr-table" id="hrEmployeeTable"><thead><tr>\${HR_HEADERS.map(h=>'<th>'+esc(h)+'</th>').join('')}</tr></thead><tbody></tbody></table></div>
        <div class="hr-count" id="hrCount"></div>
      </div>`;
    document.getElementById('hrUpdated').textContent = hrLastUpdated
      ? 'Last Updated: ' + hrLastUpdated.toLocaleString('en-IN') + ' · Auto refresh every 5 minutes'
      : 'Google Sheet से डेटा प्राप्त किया गया।';
    ['hrSearch','hrFacilityFilter','hrPostFilter','hrTypeFilter','hrJobFilter'].forEach(id => {
      document.getElementById(id).addEventListener(id === 'hrSearch' ? 'input' : 'change', filterTable);
    });
    document.getElementById('hrRefresh').addEventListener('click', load);
    document.getElementById('hrExcel').addEventListener('click', exportExcel);
    document.getElementById('hrPrint').addEventListener('click', () => window.print());
    filterTable();
  }

  function filterTable() {
    const body = document.querySelector('#hrEmployeeTable tbody');
    if (!body) return;
    const search = key(document.getElementById('hrSearch')?.value);
    const facility = clean(document.getElementById('hrFacilityFilter')?.value);
    const post = clean(document.getElementById('hrPostFilter')?.value);
    const type = clean(document.getElementById('hrTypeFilter')?.value);
    const job = clean(document.getElementById('hrJobFilter')?.value);
    const filtered = hrRows.filter(row => {
      const all = HR_HEADERS.map(h => row[h]).join(' ').toLocaleLowerCase('en-IN');
      return (!search || all.includes(search)) &&
        (!facility || row['Posting Place'] === facility) &&
        (!post || row.Post === post) &&
        (!type || row['Posting Facility Type'] === type) &&
        (!job || row['Job Type'] === job);
    });
    body.innerHTML = filtered.map(row => '<tr>' + HR_HEADERS.map(h => '<td>' + esc(row[h]) + '</td>').join('') + '</tr>').join('');
    const count = document.getElementById('hrCount');
    if (count) count.textContent = 'दिखाए गए कर्मचारी: ' + filtered.length + ' / कुल ' + hrRows.length;
  }

  function exportExcel() {
    if (!window.XLSX) { alert('Excel सुविधा अभी उपलब्ध नहीं है।'); return; }
    const rows = [HR_HEADERS, ...hrRows.map(row => HR_HEADERS.map(h => row[h]))];
    const wb = XLSX.utils.aoa_to_sheet(rows);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, wb, 'Human Resource');
    XLSX.writeFile(book, 'Kharsia_Human_Resource.xlsx');
  }

  async function load() {
    if (hrLoading) return;
    hrLoading = true;
    const status = document.getElementById('hrUpdated');
    if (status) status.textContent = '⏳ Google Sheet से डेटा लोड हो रहा है...';
    try {
      hrRows = await fetchHRData();
      hrLastUpdated = new Date();
      render();
      if (!hrRows.length) {
        const msg = document.getElementById('hrUpdated');
        if (msg) msg.textContent = 'Sheet में कर्मचारी डेटा नहीं मिला। Header और Sheet sharing जाँचें।';
      }
    } catch (error) {
      console.error('HR data load error:', error);
      if (status) status.textContent = 'डेटा लोड नहीं हुआ: ' + error.message +
        ' · Sheet को Anyone with the link – Viewer या Publish to web करें।';
    } finally { hrLoading = false; }
  }

  function openHR() {
    document.querySelectorAll('.page').forEach(page => {
      page.classList.remove('active');
      page.style.display = 'none';
    });
    const root = document.getElementById('hrModule');
    root.classList.add('active');
    root.style.display = 'block';
    document.querySelectorAll('#menuButtons .menu-btn').forEach(btn => btn.classList.remove('active'));
    const btn = document.getElementById('hrMenuButton');
    if (btn) btn.classList.add('active');
    const dash = document.getElementById('dashboardPage');
    const report = document.getElementById('reportPage');
    if (dash) dash.classList.remove('active');
    if (report) report.classList.remove('active');
    window.scrollTo(0,0);
    if (!hrLastUpdated) load();
  }

  function install() {
    const menu = document.getElementById('menuButtons');
    const cards = document.getElementById('dashboardCards');
    const content = document.querySelector('main.content');
    if (!menu || !cards || !content) return false;
    if (!document.getElementById('hrModule')) {
      const section = document.createElement('section');
      section.id = 'hrModule';
      section.className = 'page';
      section.style.display = 'none';
      content.appendChild(section);
    }
    if (!document.getElementById('hrMenuButton')) {
      const button = document.createElement('button');
      button.type = 'button';
      button.id = 'hrMenuButton';
      button.className = 'menu-btn';
      button.textContent = '👥 Human Resource';
      button.addEventListener('click', openHR);
      menu.appendChild(button);
    }
    if (!document.getElementById('hrDashboardCard')) {
      const card = document.createElement('div');
      card.id = 'hrDashboardCard';
      card.className = 'card';
      card.innerHTML = '<div class="card-icon">👥</div><h3>Human Resource</h3><p>Employee List · Facility Wise · Designation Wise Summary</p>';
      card.addEventListener('click', openHR);
      cards.appendChild(card);
    }
    return true;
  }

  function addStyles() {
    if (document.getElementById('hr-module-styles')) return;
    const style = document.createElement('style');
    style.id = 'hr-module-styles';
    style.textContent = `
      #hrModule{padding:2px 0 28px;color:#172b3c}
      .hr-head{display:flex;justify-content:space-between;align-items:center;gap:16px;background:linear-gradient(120deg,#fff,#edf8f8);border:1px solid #d7e7ee;border-radius:16px;padding:20px;margin-bottom:12px}
      .hr-eyebrow{font-size:10px;font-weight:800;letter-spacing:1.2px;color:#0f766e}.hr-head h2{font-size:24px;margin:5px 0;color:#12334b}.hr-head p,.hr-list-head p{font-size:12px;color:#64748b}
      .hr-head-actions{display:flex;gap:7px;flex-wrap:wrap}.hr-btn{border:0;border-radius:8px;background:#0f766e;color:#fff;padding:9px 12px;font-weight:700;cursor:pointer}.hr-excel{background:#15803d}.hr-print{background:#334155}
      .hr-updated{font-size:12px;color:#64748b;padding:8px 2px}.hr-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(145px,1fr));gap:10px;margin:10px 0 16px}
      .hr-kpi{background:#fff;border:1px solid #dce7ef;border-radius:12px;padding:14px;box-shadow:0 3px 10px rgba(15,40,60,.04)}.hr-kpi span{display:block;font-size:11px;color:#64748b}.hr-kpi strong{display:block;font-size:25px;color:#0f766e;margin-top:5px}
      .hr-summary-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin-bottom:16px}.hr-summary-panel,.hr-list-panel{background:#fff;border:1px solid #dce7ef;border-radius:14px;padding:15px;box-shadow:0 3px 12px rgba(15,40,60,.04);min-width:0}.hr-summary-panel h3,.hr-list-head h3{font-size:16px;color:#173b56;margin-bottom:10px}
      .hr-summary-scroll,.hr-table-wrap{overflow:auto;max-width:100%}.hr-table{border-collapse:collapse;width:100%;min-width:850px;background:#fff}.hr-table th{background:#0f536f;color:#fff;font-size:12px;padding:10px 8px;white-space:nowrap}.hr-table td{border:1px solid #e2e8f0;padding:8px;font-size:12px}.hr-table tbody tr:nth-child(even){background:#f7fafc}.hr-small-table{min-width:0}.hr-small-table th,.hr-small-table td{padding:8px}.hr-number{text-align:center;font-weight:800}.hr-total{background:#eaf5f4!important;font-weight:800}
      .hr-list-head{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:12px}.hr-list-head input,.hr-filters select{border:1px solid #cbd5e1;border-radius:8px;padding:9px 10px;background:#fff;min-width:150px}.hr-filters{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px}.hr-filters select{flex:1}.hr-count{font-size:12px;color:#64748b;padding-top:10px}
      @media(max-width:760px){.hr-head,.hr-list-head{align-items:flex-start;flex-direction:column}.hr-summary-grid{grid-template-columns:1fr}.hr-head h2{font-size:20px}.hr-head-actions{width:100%}.hr-btn{flex:1}.hr-list-head input{width:100%}}
      @media print{body *{visibility:hidden!important}#hrModule,#hrModule *{visibility:visible!important}#hrModule{position:absolute;left:0;top:0;width:100%;padding:0}.hr-head-actions,.hr-filters,.hr-list-head input{display:none!important}.hr-summary-grid{grid-template-columns:1fr 1fr}.hr-table{min-width:0;font-size:9px}.hr-table th,.hr-table td{padding:4px}}
    `;
    document.head.appendChild(style);
  }

  function boot() {
    addStyles();
    install();
    // Keep the HR page isolated when navigating back to the dashboard or another report.
    if (typeof window.showDashboard === 'function' && !window.showDashboard.__hrWrapped) {
      const originalShowDashboard = window.showDashboard;
      const wrappedShowDashboard = function () {
        const hr = document.getElementById('hrModule');
        if (hr) { hr.classList.remove('active'); hr.style.display = 'none'; }
        return originalShowDashboard.apply(this, arguments);
      };
      wrappedShowDashboard.__hrWrapped = true;
      window.showDashboard = wrappedShowDashboard;
    }
    if (typeof window.openReport === 'function' && !window.openReport.__hrWrapped) {
      const originalOpenReport = window.openReport;
      const wrappedOpenReport = function () {
        const hr = document.getElementById('hrModule');
        if (hr) { hr.classList.remove('active'); hr.style.display = 'none'; }
        return originalOpenReport.apply(this, arguments);
      };
      wrappedOpenReport.__hrWrapped = true;
      window.openReport = wrappedOpenReport;
    }
    if (!window.__hrAutoRefreshStarted) {
      window.__hrAutoRefreshStarted = true;
      setInterval(() => {
        if (document.getElementById('hrModule')?.classList.contains('active')) load();
      }, HR_REFRESH_MS);
    }
    // Dashboard's existing Google Charts callback builds its menu asynchronously.
    const original = window.createMenu;
    if (typeof original === 'function' && !original.__hrWrapped) {
      const wrapped = function () { original.apply(this, arguments); install(); };
      wrapped.__hrWrapped = true;
      window.createMenu = wrapped;
    }
    let attempts = 0;
    const timer = setInterval(() => {
      install();
      if (++attempts > 30) clearInterval(timer);
    }, 300);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
  window.openHumanResource = openHR;
})();
