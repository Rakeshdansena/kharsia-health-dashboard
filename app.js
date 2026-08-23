let currentReportTitle = "";

function showDashboardView() {
  updateActiveMenu('dashboard-menu');

  const content = document.getElementById('main-content');
  content.innerHTML = `
    <h2 class="dashboard-title">Kharsia Health Programme Dashboard</h2>
    <p class="dashboard-subtitle">विभिन्न स्वास्थ्य कार्यक्रमों की प्रगति रिपोर्ट देखने के लिए नीचे दिए कार्ड या साइडबार पर क्लिक करें।</p>
    
    <div class="cards-grid">
      <div class="card" onclick="loadReportView('ayushman')">
        <div class="card-icon">💳</div>
        <h3>Ayushman Card</h3>
        <p>आयुष्मान कार्ड एवं वय वंदन कार्ड की प्रगति</p>
      </div>
      <div class="card" onclick="loadReportView('rch')">
        <div class="card-icon">🤱</div>
        <h3>RCH 2.0</h3>
        <p>RCH 2.0 PW Registration Detail</p>
      </div>
      <div class="card" onclick="loadReportView('ncd')">
        <div class="card-icon">❤️</div>
        <h3>NCD Status</h3>
        <p>NCD Status 2026-27 Reporting</p>
      </div>
      <div class="card" onclick="loadReportView('jas')">
        <div class="card-icon">🤝</div>
        <h3>JAS Meeting</h3>
        <p>JAS Meeting Progress</p>
      </div>
      <div class="card" onclick="loadReportView('hwc')">
        <div class="card-icon">🏥</div>
        <h3>Health & Wellness Center</h3>
        <p>Block Kharsia HWC Progress</p>
      </div>
      <div class="card" onclick="loadReportView('shivir')">
        <div class="card-icon">🏕️</div>
        <h3>Ayushman Shivir</h3>
        <p>Shivir Reporting FY 2026-27</p>
      </div>
      <div class="card" onclick="loadReportView('activity')">
        <div class="card-icon">🩺</div>
        <h3>Wellness Activity</h3>
        <p>Arogya Mandir Wellness Activity</p>
      </div>
      <div class="card" onclick="loadReportView('rbsk')">
        <div class="card-icon">👶</div>
        <h3>RBSK</h3>
        <p>राष्ट्रीय बाल स्वास्थ्य कार्यक्रम</p>
      </div>
      <div class="card" onclick="loadReportView('nrc')">
        <div class="card-icon">🏢</div>
        <h3>NRC Kharsia</h3>
        <p>पोषण पुनर्वास केंद्र रिपोर्ट</p>
      </div>
      <div class="card" onclick="loadReportView('blindness')">
        <div class="card-icon">👁️</div>
        <h3>Blindness Control</h3>
        <p>राष्ट्रीय दृष्टिहीनता नियंत्रण कार्यक्रम</p>
      </div>
      <div class="card" onclick="loadReportView('nqas')">
        <div class="card-icon">🏅</div>
        <h3>NQAS Certification</h3>
        <p>गुणवत्ता प्रमाणन स्थिति</p>
      </div>
      <div class="card" onclick="loadReportView('dialysis')">
        <div class="card-icon">💧</div>
        <h3>Dialysis</h3>
        <p>प्रधानमंत्री डायलिसिस कार्यक्रम</p>
      </div>
      <div class="card" onclick="loadReportView('nlep')">
        <div class="card-icon">🩺</div>
        <h3>NLEP</h3>
        <p>राष्ट्रीय कुष्ठ उन्मूलन कार्यक्रम</p>
      </div>
    </div>
  `;
}

function loadReportView(reportKey) {
  const config = REPORT_CONFIG[reportKey];
  if (!config) return;

  currentReportTitle = config.title;
  updateActiveMenu(`menu-${reportKey}`);

  const content = document.getElementById('main-content');
  content.innerHTML = `
    <div class="table-card">
      <div class="table-toolbar">
        <div class="table-title">${config.title}</div>
        <div class="table-actions">
          <input type="text" id="tableSearch" class="search-input" onkeyup="filterTable()" placeholder="🔍 डेटा खोजें...">
          <button class="btn btn-excel" onclick="exportToExcel()">📊 Excel डाउनलोड</button>
          <button class="btn btn-pdf" onclick="exportToPDF()">📄 PDF डाउनलोड</button>
        </div>
      </div>
      <div class="table-responsive">
        <div id="data-container" style="padding: 40px; text-align: center; color: #64748B;">
          ⏳ गूगल शीट से लाइव डेटा लोड हो रहा है...
        </div>
      </div>
    </div>
  `;

  const proxyUrl = `https://corsproxy.io/?` + encodeURIComponent(config.csvUrl);

  fetch(proxyUrl)
    .then(res => res.text())
    .then(csvText => renderTable(csvText))
    .catch(() => {
      fetch(`https://api.allorigins.win/raw?url=` + encodeURIComponent(config.csvUrl))
        .then(res => res.text())
        .then(csvText => renderTable(csvText))
        .catch(() => {
          document.getElementById('data-container').innerHTML = `
            <div style="color: #DC2626; padding: 20px;">
              ❌ डेटा लोड करने में समस्या आई। गूगल शीट Publish to Web सेट है या नहीं यह जांचें।
            </div>`;
        });
    });
}

function renderTable(csvText) {
  const lines = csvText.trim().split('\n');
  if (lines.length === 0 || lines[0].trim() === '') {
    document.getElementById('data-container').innerHTML = '<div style="padding: 20px;">कोई डेटा उपलब्ध नहीं है।</div>';
    return;
  }

  let html = '<table id="reportTable">';
  lines.forEach((line, rIdx) => {
    const cells = line.match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g) || line.split(',');
    html += '<tr>';
    cells.forEach(cell => {
      let clean = cell ? cell.replace(/^"|"$/g, '').trim() : '';
      if (rIdx === 0) {
        html += `<th>${clean}</th>`;
      } else {
        html += `<td>${clean}</td>`;
      }
    });
    html += '</tr>';
  });
  html += '</table>';
  document.getElementById('data-container').innerHTML = html;
}

// 🔍 लाइव सर्च फ़ंक्शन
function filterTable() {
  const input = document.getElementById("tableSearch");
  const filter = input.value.toUpperCase();
  const table = document.getElementById("reportTable");
  if (!table) return;

  const tr = table.getElementsByTagName("tr");
  for (let i = 1; i < tr.length; i++) {
    let visible = false;
    const td = tr[i].getElementsByTagName("td");
    for (let j = 0; j < td.length; j++) {
      if (td[j]) {
        if (td[j].textContent.toUpperCase().indexOf(filter) > -1) {
          visible = true;
          break;
        }
      }
    }
    tr[i].style.display = visible ? "" : "none";
  }
}

// 📊 Excel डाउनलोड फ़ंक्शन
function exportToExcel() {
  const table = document.getElementById("reportTable");
  if (!table) return;
  const wb = XLSX.utils.table_to_book(table, { sheet: "Report" });
  XLSX.writeFile(wb, `${currentReportTitle}.xlsx`);
}

// 📄 PDF डाउनलोड फ़ंक्शन
function exportToPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF('l', 'pt', 'a4');
  
  doc.setFontSize(14);
  doc.text(currentReportTitle, 40, 30);

  doc.autoTable({
    html: '#reportTable',
    startY: 45,
    styles: { fontSize: 8, cellPadding: 4 },
    headStyles: { fillColor: [30, 41, 59] }
  });

  doc.save(`${currentReportTitle}.pdf`);
}

function updateActiveMenu(activeId) {
  document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
  const el = document.getElementById(activeId);
  if (el) el.classList.add('active');
}

window.onload = showDashboardView;
