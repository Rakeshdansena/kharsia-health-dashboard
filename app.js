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

  updateActiveMenu(`menu-${reportKey}`);

  const content = document.getElementById('main-content');
  content.innerHTML = `
    <div class="table-card">
      <div class="table-header">${config.title}</div>
      <div class="table-responsive">
        <div id="data-container" style="padding: 40px; text-align: center;">
          डेटा लोड हो रहा है...
        </div>
      </div>
    </div>
  `;

  const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(config.csvUrl)}`;

  fetch(proxyUrl)
    .then(res => {
      if (!res.ok) throw new Error('Network response error');
      return res.text();
    })
    .then(csvText => {
      const rows = csvText.trim().split('\n').map(row => row.match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g) || row.split(','));
      
      if(rows.length === 0 || (rows.length === 1 && rows[0][0] === "")) {
        document.getElementById('data-container').innerHTML = '<div style="padding: 20px;">कोई डेटा उपलब्ध नहीं है।</div>';
        return;
      }

      let html = '<table>';
      rows.forEach((row, rIdx) => {
        html += '<tr>';
        row.forEach(cell => {
          let clean = cell ? cell.replace(/^"|"$/g, '').trim() : '';
          html += rIdx === 0 ? `<th>${clean}</th>` : `<td>${clean}</td>`;
        });
        html += '</tr>';
      });
      html += '</table>';
      document.getElementById('data-container').innerHTML = html;
    })
    .catch(() => {
      document.getElementById('data-container').innerHTML = `
        <div style="color: #DC2626; padding: 20px;">
          ❌ डेटा लोड करने में समस्या आई। कृपया Google Sheet की Publish setting या CSV URL जांचें।
        </div>`;
    });
}

function updateActiveMenu(activeId) {
  document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
  const el = document.getElementById(activeId);
  if (el) el.classList.add('active');
}

window.onload = showDashboardView;
