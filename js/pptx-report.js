/* Kharsia Health Dashboard — Progressive PPTX Report
   Stable standalone generator for the Dashboard PPTX button.\n   Cover slide follows the approved sample-style healthcare layout.
*/
(function () {
  'use strict';

  var SHEET_ID = '1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0';
  var RCH_GID = '1044088930';

  function wait(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }

  function clean(v) {
    return String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
  }

  function number(v) {
    var n = Number(String(v == null ? '' : v).replace(/,/g, '').replace(/%/g, ''));
    return isFinite(n) ? n : 0;
  }

  function setButton(label, busy) {
    var b = document.getElementById('dashboardPptxBtn');
    if (!b) return;
    b.disabled = !!busy;
    b.textContent = label;
    b.style.opacity = busy ? '0.7' : '1';
  }

  function loadPptx() {
    return new Promise(function (resolve, reject) {
      if (window.PptxGenJS) return resolve();
      var s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
      s.onload = function () {
        if (window.PptxGenJS) resolve();
        else reject(new Error('PowerPoint library load हुई लेकिन PptxGenJS नहीं मिला।'));
      };
      s.onerror = function () {
        reject(new Error('PowerPoint library load नहीं हुई। Internet connection check करें।'));
      };
      document.head.appendChild(s);
    });
  }

  function waitForGoogle() {
    return new Promise(function (resolve, reject) {
      var tries = 0;
      function check() {
        if (window.google && google.visualization && google.visualization.Query) {
          resolve();
          return;
        }
        tries++;
        if (tries > 60) {
          reject(new Error('Google Sheets service उपलब्ध नहीं है।'));
          return;
        }
        setTimeout(check, 250);
      }
      check();
    });
  }

  function getRchData() {
    return new Promise(function (resolve, reject) {
      var q = new google.visualization.Query(
        'https://docs.google.com/spreadsheets/d/' +
        encodeURIComponent(SHEET_ID) +
        '/gviz/tq?gid=' + RCH_GID + '&headers=0'
      );
      q.setQuery('select *');
      q.send(function (response) {
        try {
          if (response.isError()) {
            throw new Error('Janani Portal data load नहीं हुआ: ' + response.getMessage());
          }
          if (typeof window.parseRCHData !== 'function') {
            throw new Error('Janani Portal parser उपलब्ध नहीं है।');
          }
          var parsed = window.parseRCHData(response.getDataTable());
          if (!parsed || !parsed.facilityRows || !parsed.facilityRows.length) {
            throw new Error('Janani Portal में कोई facility data नहीं मिला।');
          }
          window.currentRCHData = parsed;
          resolve(parsed);
        } catch (e) {
          reject(e);
        }
      });
    });
  }

  function addHeader(slide, title, subtitle) {
    slide.background = { color: 'F8FAFC' };
    slide.addShape('rect', {
      x: 0, y: 0, w: 13.333, h: 0.72,
      fill: { color: '0B4F6C' }, line: { color: '0B4F6C' }
    });
    slide.addShape('rect', {
      x: 0, y: 0.72, w: 13.333, h: 0.08,
      fill: { color: '14B8A6' }, line: { color: '14B8A6' }
    });
    slide.addText(title, {
      x: 0.48, y: 0.16, w: 8.8, h: 0.34,
      fontSize: 25, bold: true, color: 'FFFFFF', margin: 0
    });
    slide.addText(subtitle || '', {
      x: 9.0, y: 0.20, w: 3.75, h: 0.25,
      fontSize: 11, bold: true, color: 'DFF7F3', align: 'right', margin: 0
    });
    slide.addText('KHARSIA HEALTH DEPARTMENT  •  BLOCK KHARSIA  •  DISTRICT RAIGARH', {
      x: 0.48, y: 7.12, w: 9.5, h: 0.18,
      fontSize: 7, color: '64748B', margin: 0
    });
    slide.addText('FY 2026–27', {
      x: 11.0, y: 7.12, w: 1.8, h: 0.18,
      fontSize: 7, bold: true, color: '0B4F6C', align: 'right', margin: 0
    });
  }

  function addSectionTitle(slide, text, x, y, w) {
    slide.addText(text, {
      x:x, y:y, w:w, h:0.32, fontSize:19, bold:true,
      color:'0B4F6C', margin:0
    });
    slide.addShape('rect', {
      x:x, y:y+0.38, w:0.65, h:0.045,
      fill:{color:'14B8A6'}, line:{color:'14B8A6'}
    });
  }

  function addCard(slide, x, y, w, h, label, value) {
    slide.addShape('roundRect', {
      x:x, y:y, w:w, h:h,
      rectRadius:0.08,
      fill:{color:'FFFFFF'},
      line:{color:'D7E2EA', pt:1},
      shadow:{type:'outer', color:'94A3B8', blur:1, angle:45, distance:1, opacity:0.16}
    });
    slide.addShape('rect', {
      x:x, y:y, w:0.08, h:h,
      fill:{color:'14B8A6'}, line:{color:'14B8A6'}
    });
    slide.addText(label, {
      x:x+0.22, y:y+0.15, w:w-0.35, h:0.25,
      fontSize:11, bold:true, color:'475569', margin:0, fit:'shrink'
    });
    slide.addText(String(value), {
      x:x+0.22, y:y+0.48, w:w-0.35, h:0.42,
      fontSize:28, bold:true, color:'075985', margin:0, fit:'shrink'
    });
  }

  function addPill(slide, text, x, y, w, color) {
    slide.addShape('roundRect', {
      x:x, y:y, w:w, h:0.34,
      fill:{color:color}, line:{color:color}
    });
    slide.addText(text, {
      x:x, y:y+0.08, w:w, h:0.15,
      fontSize:7.5, bold:true, color:'FFFFFF', align:'center', margin:0
    });
  }

  function groupSectors(rows) {
    var map = {};
    rows.forEach(function (r) {
      var s = clean(r.sector) || 'Other';
      if (!map[s]) map[s] = { sector: s, hmis: 0, rch: 0, temp: 0, backlog: 0, highRisk: 0, facilities: 0 };
      map[s].hmis += number(r.hmis);
      map[s].rch += number(r.rch);
      map[s].temp += number(r.temp);
      map[s].backlog += number(r.backlog);
      map[s].highRisk += number(r.highRisk);
      map[s].facilities++;
    });
    return Object.keys(map).map(function (k) {
      var x = map[k];
      x.percent = x.hmis > 0 ? Math.round(x.rch / x.hmis * 100) : 0;
      return x;
    });
  }

  function analysis(sector) {
    var a = [];
    if (sector.percent >= 90) a.push('Registration achievement is 90% or above.');
    else if (sector.percent >= 70) a.push('Registration achievement is in the 70–89% range.');
    else a.push('Registration achievement is below 70%; facility-level gaps should be reviewed.');
    if (sector.backlog < 0) a.push('Backlog is negative.');
    else if (sector.backlog === 0) a.push('Backlog is zero.');
    else a.push('Positive backlog is recorded and should be followed up.');
    if (sector.highRisk > 0) a.push('High-risk cases are reported in this sector.');
    a.push('Facilities covered: ' + sector.facilities + '.');
    return a;
  }

  function addTable(slide, rows, x, y, w, h) {
    if (!rows.length) return;
    var cols = 0;
    rows.forEach(function (r) { cols = Math.max(cols, r.length); });
    var width = w / Math.max(cols, 1);
    var data = rows.map(function (r, ri) {
      return r.map(function (v) {
        return {
          text: clean(v),
          options: {
            bold: ri === 0,
            fontSize: Math.max(6, Math.min(10, 10 - cols * 0.35)),
            color: ri === 0 ? 'FFFFFF' : '172033',
            fill: { color: ri === 0 ? '075985' : 'FFFFFF' },
            align: 'center', valign: 'mid', margin: 2
          }
        };
      });
    });
    slide.addTable(data, {
      x: x, y: y, w: w, h: h,
      border: { type: 'solid', color: 'CBD5E1', pt: 1 },
      autoFit: false, colW: Array(cols).fill(width), rowH: 0.3, margin: 2
    });
  }


  function showProgress(title, percent, detail) {
    var host = document.getElementById('dashboardPptxBtn');
    if (!host) return;
    var box = document.getElementById('pptxProgressBox');
    if (!box) {
      box = document.createElement('div');
      box.id = 'pptxProgressBox';
      box.style.cssText = 'margin:12px 0;padding:14px 16px;background:#eff6ff;border:2px solid #3b82f6;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,.08);font-weight:700;';
      host.parentElement.appendChild(box);
    }
    box.innerHTML =
      '<div style="display:flex;justify-content:space-between;gap:12px;margin-bottom:7px">' +
      '<span>' + title + '</span><span>' + percent + '%</span></div>' +
      '<div style="height:12px;background:#dbeafe;border-radius:8px;overflow:hidden">' +
      '<div style="height:100%;width:' + percent + '%;background:#2563eb;transition:width .25s"></div></div>' +
      '<div style="margin-top:7px;font-size:12px;font-weight:500;color:#475569">' + detail + '</div>';
  }

  var COVER_PHOTO_URL = 'https://images.unsplash.com/photo-1758691462878-6edc3d3da1be?auto=format&fit=crop&fm=jpg&ixid=M3wxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8fA%3D%3D&ixlib=rb-4.1.0&q=80&w=1800';

  function fetchImageData(url) {
    return fetch(url, { mode: 'cors' }).then(function (r) {
      if (!r.ok) throw new Error('Cover photo load failed');
      return r.blob();
    }).then(function (blob) {
      return new Promise(function (resolve, reject) {
        var reader = new FileReader();
        reader.onloadend = function () { resolve(reader.result); };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
    });
  }

  function queryGenericSheet(gid) {
    return new Promise(function (resolve, reject) {
      if (!gid || gid === '0') { resolve(null); return; }
      var q = new google.visualization.Query(
        'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(SHEET_ID) +
        '/gviz/tq?gid=' + encodeURIComponent(gid) + '&headers=0'
      );
      q.setQuery('select *');
      q.send(function (response) {
        if (response.isError()) { resolve(null); return; }
        try {
          var dt = response.getDataTable();
          var headers = [], rows = [];
          for (var c = 0; c < dt.getNumberOfColumns(); c++) headers.push(clean(dt.getFormattedValue(0, c)) || ('Column ' + (c + 1)));
          for (var r = 1; r < dt.getNumberOfRows(); r++) {
            var row = [];
            for (var cc = 0; cc < dt.getNumberOfColumns(); cc++) row.push(clean(dt.getFormattedValue(r, cc)));
            if (row.some(function(v){ return v !== ''; })) rows.push(row);
          }
          resolve({ headers: headers, rows: rows });
        } catch (e) { resolve(null); }
      });
    });
  }

  function addGenericModuleSlide(pptx, module, data, index) {
    if (!data || !data.rows || !data.rows.length) return;
    var slide = pptx.addSlide();
    addHeader(slide, module.icon + ' ' + module.name, 'MODULE PROGRESS REPORT');
    addSectionTitle(slide, 'MODULE SUMMARY', 0.55, 1.05, 3.2);

    var numeric = [];
    data.headers.forEach(function(h, ci) {
      var total = 0, count = 0;
      data.rows.forEach(function(row) {
        var raw = row[ci];
        if (raw !== '' && isNumberLike(String(raw).replace(/,/g,''))) {
          total += number(raw); count++;
        }
      });
      if (count) numeric.push({ header:h, total:total, count:count, ci:ci });
    });
    numeric = numeric.slice(0, 4);
    numeric.forEach(function(m, i) {
      addCard(slide, 0.55 + i * 3.1, 1.48, 2.8, 1.15, m.header, Math.round(m.total * 100) / 100);
    });

    var chartMetric = numeric[0];
    if (chartMetric) {
      addSectionTitle(slide, 'PERFORMANCE CHART', 0.55, 2.95, 5.0);
      var chartData = data.rows.slice(0, 12).map(function(row, i) {
        return { name: clean(row[0]) || ('Item ' + (i + 1)), value: number(row[chartMetric.ci]) };
      });
      try {
        slide.addChart(pptx.ChartType.bar, [{
          name: chartMetric.header,
          labels: chartData.map(function(x){return x.name;}),
          values: chartData.map(function(x){return x.value;})
        }], {
          x:0.45,y:3.35,w:7.35,h:3.15,showLegend:false,showTitle:false,showValue:true,
          catAxisLabelFontSize:12,valAxisLabelFontSize:11,
          chartColors:['0F766E','2563EB','F59E0B','DC2626','7C3AED','059669','EA580C','0891B2','DB2777','65A30D','0284C7','9333EA'],
          valGridLine:{color:'D6E3EC',pt:1},valAxisMinVal:0,dataLabelPosition:'outEnd'
        });
      } catch(e) {}
    }

    addSectionTitle(slide, 'KEY DATA / OBSERVATIONS', 8.05, 2.95, 4.5);
    var obs = [];
    obs.push('• Reporting rows available: ' + data.rows.length + '.');
    if (numeric.length) numeric.forEach(function(m){ obs.push('• ' + m.header + ': ' + Math.round(m.total * 100) / 100 + '.'); });
    if (data.headers.length) obs.push('• Data fields: ' + data.headers.length + '.');
    slide.addText(obs.join('\n'), {
      x:8.1,y:3.45,w:4.55,h:2.5,fontSize:15,bold:true,color:'172033',margin:0.03,fit:'shrink'
    });

    var table = [data.headers.slice(0, 6)];
    data.rows.slice(0, 7).forEach(function(row){ table.push(row.slice(0,6)); });
    addTable(slide, table, 8.05, 5.35, 4.65, 1.2);
    slide.addText('Source: live Google Sheet | Module: ' + module.name, {
      x:8.05,y:6.7,w:4.5,h:0.2,fontSize:7.5,color:'64748B',margin:0
    });
  }

  async function generate() {
    setButton('⏳ PPTX तैयार हो रहा है...', true);
    try {
      await waitForGoogle();
      await loadPptx();
      var parsed = await getRchData();
      var rows = parsed.facilityRows || [];
      var sectors = groupSectors(rows);
      showProgress('Data तैयार है', 45, rows.length + ' facility records और ' + sectors.length + ' sectors मिले।');
      var total = { hmis: 0, rch: 0, temp: 0, backlog: 0, highRisk: 0, facilities: rows.length };
      rows.forEach(function (r) {
        total.hmis += number(r.hmis);
        total.rch += number(r.rch);
        total.temp += number(r.temp);
        total.backlog += number(r.backlog);
        total.highRisk += number(r.highRisk);
      });
      total.percent = total.hmis > 0 ? Math.round(total.rch / total.hmis * 100) : 0;

      var pptx = new window.PptxGenJS();
      pptx.layout = 'LAYOUT_WIDE';
      pptx.author = 'Kharsia Health Dashboard';
      pptx.company = 'Kharsia Health Dashboard';
      pptx.subject = 'Progressive Health Report';
      pptx.title = 'Kharsia Health Progressive Report';
      pptx.lang = 'en-IN';

      // 1. COVER — final approved whole-programme cover.
      // Render Slide 1 as one SVG image so the downloaded PPTX matches the approved preview.
      // Only Slide 1 uses this design; all later slides keep their existing layouts.
      var coverPhoto = null;
      try { coverPhoto = await fetchImageData(COVER_PHOTO_URL); } catch (photoErr) { console.warn(photoErr); }

      var coverDate = new Date().toLocaleDateString('en-IN', {
        day:'2-digit', month:'long', year:'numeric'
      });

      function escSvg(v) {
        return String(v == null ? '' : v)
          .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
          .replace(/"/g,'&quot;');
      }

      function svgDataUri(svg) {
        return 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svg)));
      }

      var photoHref = coverPhoto || '';
      var coverSvg =
        '<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 1600 900">' +
        '<defs>' +
          '<linearGradient id="top" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#ffffff"/><stop offset="1" stop-color="#eaf6ff"/></linearGradient>' +
          '<linearGradient id="blue" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#0b4f8a"/><stop offset="1" stop-color="#075985"/></linearGradient>' +
          '<clipPath id="photoClip"><path d="M1080 128 C1320 70 1490 135 1600 250 L1600 670 C1480 730 1260 750 1080 680 Z"/></clipPath>' +
          '<filter id="shadow"><feDropShadow dx="0" dy="8" stdDeviation="10" flood-opacity=".16"/></filter>' +
        '</defs>' +
        '<rect width="1600" height="900" fill="#f7fbff"/>' +
        '<rect width="1600" height="142" fill="url(#top)"/>' +
        '<rect y="136" width="1600" height="8" fill="#59aee8"/>' +
        '<text x="68" y="62" font-family="Arial, Noto Sans, sans-serif" font-size="27" font-weight="800" fill="#073b75">छत्तीसगढ़ शासन</text>' +
        '<text x="68" y="101" font-family="Arial, Noto Sans, sans-serif" font-size="21" font-weight="700" fill="#1f4e79">स्वास्थ्य एवं परिवार कल्याण विभाग</text>' +
        '<circle cx="31" cy="77" r="25" fill="#ffffff" stroke="#075985" stroke-width="4"/>' +
        '<text x="31" y="86" text-anchor="middle" font-family="Arial" font-size="25" font-weight="800" fill="#075985">+</text>' +
        '<text x="1470" y="61" text-anchor="middle" font-family="Arial" font-size="18" font-weight="800" fill="#0b4f6c">NATIONAL HEALTH</text>' +
        '<text x="1470" y="87" text-anchor="middle" font-family="Arial" font-size="18" font-weight="800" fill="#0b4f6c">MISSION</text>' +
        '<circle cx="1470" cy="111" r="18" fill="#ef4444"/><text x="1470" y="118" text-anchor="middle" font-family="Arial" font-size="18" fill="#fff">+</text>' +
        '<path d="M0 185 L1080 185 C1130 185 1160 205 1190 245 L1040 690 L0 690 Z" fill="url(#blue)" opacity=".98"/>' +
        '<path d="M0 690 C260 620 560 650 850 665 C990 672 1060 635 1110 590 L1045 900 L0 900 Z" fill="#e7f4ff"/>' +
        '<path d="M100 185 L1080 185 C1110 185 1130 198 1155 220" fill="none" stroke="#4aa8e8" stroke-width="18" opacity=".8"/>' +
        '<g filter="url(#shadow)">' +
          '<rect x="110" y="207" width="920" height="102" rx="18" fill="#073b75"/>' +
          '<text x="570" y="278" text-anchor="middle" font-family="Arial, Noto Sans, sans-serif" font-size="56" font-weight="900" fill="#ffffff">KHARSIA HEALTH DEPARTMENT</text>' +
          '<rect x="110" y="328" width="920" height="86" rx="18" fill="#ffffff" opacity=".97"/>' +
          '<text x="570" y="386" text-anchor="middle" font-family="Arial, Noto Sans, sans-serif" font-size="48" font-weight="900" fill="#18783c">PROGRESSIVE REPORT</text>' +
          '<rect x="330" y="435" width="480" height="70" rx="18" fill="#0d63c9"/>' +
          '<text x="570" y="483" text-anchor="middle" font-family="Arial" font-size="38" font-weight="900" fill="#ffffff">FY 2026 – 27</text>' +
        '</g>' +
        '<rect x="95" y="530" width="900" height="70" rx="20" fill="#ffffff" opacity=".94"/>' +
        '<text x="545" y="575" text-anchor="middle" font-family="Arial, Noto Sans, sans-serif" font-size="29" font-weight="800" fill="#0f2d52">Block Kharsia  |  District Raigarh  |  Chhattisgarh</text>' +
        '<rect x="265" y="610" width="560" height="56" rx="17" fill="#fff0f3"/>' +
        '<text x="545" y="648" text-anchor="middle" font-family="Arial, Noto Sans, sans-serif" font-size="24" font-weight="800" fill="#be185d">As On Date : ' + escSvg(coverDate) + '</text>' +
        (photoHref ?
          '<image href="' + photoHref + '" x="1050" y="130" width="550" height="560" preserveAspectRatio="xMidYMid slice" clip-path="url(#photoClip)"/>' :
          '<path d="M1080 128 C1320 70 1490 135 1600 250 L1600 670 C1480 730 1260 750 1080 680 Z" fill="#dff2ff"/>') +
        '<path d="M1080 128 C1320 70 1490 135 1600 250 L1600 670 C1480 730 1260 750 1080 680 Z" fill="none" stroke="#ffffff" stroke-width="12"/>' +
        '<g font-family="Arial, Noto Sans, sans-serif" font-weight="800" font-size="18" text-anchor="middle">' +
          '<g><rect x="45" y="706" width="116" height="76" rx="16" fill="#eadcff"/><text x="103" y="750" fill="#4338ca">RCH 2.0</text></g>' +
          '<g><rect x="173" y="706" width="116" height="76" rx="16" fill="#ffd7e8"/><text x="231" y="750" fill="#be185d">Ayushman</text></g>' +
          '<g><rect x="301" y="706" width="116" height="76" rx="16" fill="#d8f0ff"/><text x="359" y="750" fill="#075985">NCD</text></g>' +
          '<g><rect x="429" y="706" width="116" height="76" rx="16" fill="#dcfce7"/><text x="487" y="750" fill="#166534">AAM</text></g>' +
          '<g><rect x="557" y="706" width="116" height="76" rx="16" fill="#fff1b8"/><text x="615" y="750" fill="#92400e">JAS Meeting</text></g>' +
          '<g><rect x="685" y="706" width="116" height="76" rx="16" fill="#ffe0cf"/><text x="743" y="750" fill="#c2410c">Ayushman Shivir</text></g>' +
          '<g><rect x="813" y="706" width="116" height="76" rx="16" fill="#f6d7ff"/><text x="871" y="750" fill="#7e22ce">Wellness Activity</text></g>' +
          '<g><rect x="941" y="706" width="116" height="76" rx="16" fill="#e2dcff"/><text x="999" y="750" fill="#4338ca">RBSK</text></g>' +
          '<g><rect x="1069" y="706" width="116" height="76" rx="16" fill="#d9efff"/><text x="1127" y="750" fill="#075985">Dialysis</text></g>' +
          '<g><rect x="1197" y="706" width="116" height="76" rx="16" fill="#efffbf"/><text x="1255" y="750" fill="#3f6212">Blindness</text></g>' +
          '<g><rect x="1325" y="706" width="116" height="76" rx="16" fill="#fff1b8"/><text x="1383" y="750" fill="#854d0e">NQAS</text></g>' +
          '<g><rect x="1453" y="706" width="116" height="76" rx="16" fill="#d9f7fb"/><text x="1511" y="750" fill="#155e75">NLEP</text></g>' +
        '</g>' +
        '<rect y="820" width="1600" height="80" fill="#073b75"/>' +
        '<text x="70" y="870" font-family="Arial, Noto Sans, sans-serif" font-size="23" font-weight="800" fill="#ffffff">Health Department  |  District Raigarh  |  Chhattisgarh</text>' +
        '<text x="1355" y="870" text-anchor="end" font-family="Arial, Noto Sans, sans-serif" font-size="23" font-weight="900" fill="#ffd900">Monthly Performance Review</text><rect x="1435" y="835" width="120" height="43" rx="18" fill="#ffffff"/><text x="1495" y="864" text-anchor="middle" font-family="Arial, Noto Sans, sans-serif" font-size="20" font-weight="900" fill="#073b75">Page 01</text>' +
        '</svg>';

      var slide = pptx.addSlide();
      slide.background = { color:'F8FBFF' };
      slide.addImage({ data: svgDataUri(coverSvg), x:0, y:0, w:13.333, h:7.5 });

      // 2. INDEX — created after all content slides so page ranges are always exact.
      // The index slide is moved to position 2 after all modules are generated.
      var moduleRanges = [];
      var otherModules = [
        {name:'Ayushman Card',icon:'💳',gid:'925649620'},
        {name:'NCD',icon:'❤️',gid:'1254412412'},
        {name:'JAS Meeting',icon:'🤝',gid:'1018164338'},
        {name:'Health & Wellness Center',icon:'🏥',gid:'0'},
        {name:'Ayushman Shivir',icon:'🏕️',gid:'1262815420'},
        {name:'Wellness Activity',icon:'🩺',gid:'447031017'},
        {name:'RBSK',icon:'👶',gid:'1502752823'},
        {name:'NRC Kharsia',icon:'🏥',gid:'1010102020'},
        {name:'Blindness Control',icon:'👁️',gid:'1002009767'},
        {name:'NQAS Certification',icon:'🏅',gid:'728123647'},
        {name:'Dialysis',icon:'💧',gid:'781496964'},
        {name:'NLEP',icon:'🦠',gid:'1536656599'}
      ];

      showProgress('Slides बन रही हैं', 55, 'Block Summary Dashboard तैयार हो रहा है...');
      // 3. SECTOR WISE DATA — first RCH 2.0 content slide (Page 03).
      var rchStartPage = pptx.slides.length + 2;
      showProgress('Slides बन रही हैं', 60, 'RCH 2.0 Sector Wise Data और color coding तैयार हो रही है...');

      slide = pptx.addSlide();
      addHeader(slide, 'Janani Portal (RCH 2.0)', 'SECTOR WISE DATA | FY 2026–27');
      var chartRows = sectors.slice().sort(function (a, b) { return b.percent - a.percent; });

      // Colored sector table
      var tx = 0.45, ty = 1.18, tw = 12.45;
      var rowH = 0.59;
      var widths = [3.05, 1.78, 2.00, 1.68, 1.92, 2.02];
      var headers = ['Sector', 'HMIS PW', 'RCH 2.0 PW', 'Achievement', 'Backlog', 'High Risk'];
      var cx = tx;
      headers.forEach(function(h, i) {
        slide.addShape('rect', {
          x:cx, y:ty, w:widths[i], h:rowH,
          fill:{color:'075985'}, line:{color:'FFFFFF', pt:1}
        });
        slide.addText(h, {
          x:cx+0.04, y:ty+0.16, w:widths[i]-0.08, h:0.12,
          fontSize:14, bold:true, color:'FFFFFF', align:'center', margin:0, fit:'shrink'
        });
        cx += widths[i];
      });

      chartRows.forEach(function(s, ri) {
        var y = ty + rowH + ri * rowH;
        var pct = s.percent;
        var pctColor = pct >= 90 ? '16A34A' : (pct >= 70 ? 'F59E0B' : 'DC2626');
        var pctFill = pct >= 90 ? 'DCFCE7' : (pct >= 70 ? 'FEF3C7' : 'FEE2E2');
        var backlogColor = s.backlog > 0 ? '16A34A' : (s.backlog === 0 ? 'CA8A04' : 'DC2626');
        var backlogFill = s.backlog > 0 ? 'DCFCE7' : (s.backlog === 0 ? 'FEF9C3' : 'FEE2E2');
        var vals = [s.sector, s.hmis, s.rch, pct + '%', s.backlog, s.highRisk];
        cx = tx;
        vals.forEach(function(v, i) {
          var fill = ri % 2 === 0 ? 'FFFFFF' : 'F8FBFF';
          slide.addShape('rect', {
            x:cx, y:y, w:widths[i], h:rowH,
            fill:{color:fill}, line:{color:'D7E2EA', pt:0.8}
          });
          if (i === 3) {
            slide.addShape('roundRect', {
              x:cx+0.23, y:y+0.075, w:widths[i]-0.46, h:0.33,
              fill:{color:pctFill}, line:{color:pctColor, pt:1}
            });
            slide.addText(String(v), {
              x:cx+0.23, y:y+0.17, w:widths[i]-0.46, h:0.10,
              fontSize:14, bold:true, color:pctColor, align:'center', margin:0
            });
          } else if (i === 4) {
            slide.addShape('roundRect', {
              x:cx+0.25, y:y+0.075, w:widths[i]-0.5, h:0.33,
              fill:{color:backlogFill}, line:{color:backlogColor, pt:1}
            });
            slide.addText(String(v), {
              x:cx+0.25, y:y+0.17, w:widths[i]-0.5, h:0.10,
              fontSize:14, bold:true, color:backlogColor, align:'center', margin:0
            });
          } else {
            slide.addText(String(v), {
              x:cx+0.05, y:y+0.17, w:widths[i]-0.10, h:0.10,
              fontSize:i===0?14:13, bold:i===0, color:'172033',
              align:i===0?'left':'center', margin:0, fit:'shrink'
            });
          }
          cx += widths[i];
        });
      });

      // Block Total row at the bottom of the sector table
      var totalY = ty + rowH + chartRows.length * rowH;
      var totalVals = ['BLOCK TOTAL', total.hmis, total.rch, total.percent + '%', total.backlog, total.highRisk];
      cx = tx;
      totalVals.forEach(function(v, i) {
        slide.addShape('rect', {
          x:cx, y:totalY, w:widths[i], h:rowH,
          fill:{color:'073B75'}, line:{color:'FFFFFF', pt:1}
        });
        if (i === 3) {
          slide.addText(String(v), {
            x:cx+0.05, y:totalY+0.17, w:widths[i]-0.10, h:0.12,
            fontSize:14, bold:true, color:'FFFFFF', align:'center', margin:0
          });
        } else if (i === 4) {
          var totalBacklogColor = total.backlog > 0 ? '86EFAC' : (total.backlog === 0 ? 'FDE68A' : 'FCA5A5');
          slide.addText(String(v), {
            x:cx+0.05, y:totalY+0.17, w:widths[i]-0.10, h:0.12,
            fontSize:12, bold:true, color:totalBacklogColor, align:'center', margin:0
          });
        } else {
          slide.addText(String(v), {
            x:cx+0.05, y:totalY+0.17, w:widths[i]-0.10, h:0.12,
            fontSize:12, bold:true, color:'FFFFFF',
            align:i===0?'left':'center', margin:0, fit:'shrink'
          });
        }
        cx += widths[i];
      });

      // 4. BLOCK SUMMARY REMOVED — RCH Sector Wise is now the first RCH detail slide.
      // 5. FACILITY WISE / GRAPH
      showProgress('Slides बन रही हैं', 70, 'Sector-wise graph और facility analysis तैयार हो रहे हैं...');
      // 5. FACILITY WISE — grouped by SECTOR, with a TOTAL + Achievement % row after every sector.
      // Grouping is based on r.sector (not facility name), so every facility remains under its correct sector.
      var facilityGroups = [
        { title:'Barra • Jobi • Gorpar', sectors:['barra','jobi','gorpar'] },
        { title:'Sarwani • Turekela', sectors:['sarwani','turekela'] },
        { title:'Sondka • Binjkot', sectors:['sondka','binjkot'] }
      ];

      function sectorMatchesGroup(sectorName, group) {
        var n = clean(sectorName).toLowerCase();
        return group.sectors.some(function(name){ return n === name || n.indexOf(name) >= 0; });
      }

      var assignedFacilities = {};
      facilityGroups.forEach(function(group, groupIndex) {
        showProgress('Facility slides बन रही हैं', 75 + groupIndex * 5, group.title + ' का Sector-wise Facility Data तैयार हो रहा है...');

        var groupRows = rows.filter(function(r) {
          if (!clean(r.facility)) return false;
          if (!sectorMatchesGroup(r.sector, group)) return false;
          assignedFacilities[clean(r.facility).toLowerCase()] = true;
          return true;
        });

        slide = pptx.addSlide();
        addHeader(slide, 'Janani Portal (RCH 2.0)', 'FACILITY WISE DATA | ' + group.title);

        var tx = 0.38, ty = 1.05;
        var widths = [3.45, 1.75, 1.90, 1.70, 1.75, 1.94];
        var headers = ['Facility', 'HMIS PW', 'RCH 2.0 PW', 'Achievement', 'Backlog', 'High Risk'];
        var sectorOrder = group.sectors;
        var grouped = {};
        groupRows.forEach(function(r) {
          var key = clean(r.sector) || 'Other';
          if (!grouped[key]) grouped[key] = [];
          grouped[key].push(r);
        });

        var orderedSectors = [];
        sectorOrder.forEach(function(s) {
          Object.keys(grouped).forEach(function(k) {
            if (k.toLowerCase() === s) orderedSectors.push(k);
          });
        });
        Object.keys(grouped).forEach(function(k) {
          if (orderedSectors.indexOf(k) < 0) orderedSectors.push(k);
        });

        var lineCount = 1;
        orderedSectors.forEach(function(k) { lineCount += grouped[k].length + 1; });
        var rowH = Math.max(0.25, Math.min(0.42, 5.95 / Math.max(lineCount,1)));
        var sectionH = rowH;

        var cx = tx;
        headers.forEach(function(h,i) {
          slide.addShape('rect',{x:cx,y:ty,w:widths[i],h:rowH,fill:{color:'075985'},line:{color:'FFFFFF',pt:1}});
          slide.addText(h,{x:cx+0.03,y:ty+rowH*0.28,w:widths[i]-0.06,h:rowH*0.38,fontSize:Math.max(9,Math.min(12,rowH*30)),bold:true,color:'FFFFFF',align:'center',margin:0,fit:'shrink'});
          cx += widths[i];
        });

        var currentY = ty + rowH;
        orderedSectors.forEach(function(sectorName) {
          var sectorRows = grouped[sectorName] || [];
          // Sector band
          slide.addShape('rect',{x:tx,y:currentY,w:12.49,h:sectionH,fill:{color:'E0F2FE'},line:{color:'B6D7EA',pt:1}});
          slide.addText('SECTOR: ' + sectorName.toUpperCase(),{
            x:tx+0.08,y:currentY+rowH*0.26,w:12.25,h:rowH*0.38,
            fontSize:Math.max(8,Math.min(11,rowH*28)),bold:true,color:'075985',margin:0,fit:'shrink'
          });
          currentY += sectionH;

          var sectorTotal = {hmis:0,rch:0,backlog:0,highRisk:0};
          sectorRows.forEach(function(r,ri) {
            var hmis = number(r.hmis);
            var rch = number(r.rch);
            var backlog = number(r.backlog);
            var highRisk = number(r.highRisk);
            sectorTotal.hmis += hmis;
            sectorTotal.rch += rch;
            sectorTotal.backlog += backlog;
            sectorTotal.highRisk += highRisk;

            var pct = hmis > 0 ? Math.round(rch / hmis * 100) : 0;
            var pctColor = pct >= 90 ? '16A34A' : (pct >= 70 ? 'F59E0B' : 'DC2626');
            var pctFill = pct >= 90 ? 'DCFCE7' : (pct >= 70 ? 'FEF3C7' : 'FEE2E2');
            var backlogColor = backlog > 0 ? '16A34A' : (backlog === 0 ? 'CA8A04' : 'DC2626');
            var backlogFill = backlog > 0 ? 'DCFCE7' : (backlog === 0 ? 'FEF9C3' : 'FEE2E2');
            var vals = [clean(r.facility), hmis, rch, pct + '%', backlog, highRisk];
            cx = tx;

            vals.forEach(function(v,i) {
              var fill = ri % 2 === 0 ? 'FFFFFF' : 'F8FBFF';
              slide.addShape('rect',{x:cx,y:currentY,w:widths[i],h:rowH,fill:{color:fill},line:{color:'D7E2EA',pt:0.7}});
              if(i===3) {
                slide.addShape('roundRect',{x:cx+0.18,y:currentY+rowH*0.13,w:widths[i]-0.36,h:rowH*0.70,fill:{color:pctFill},line:{color:pctColor,pt:0.8}});
                slide.addText(String(v),{x:cx+0.18,y:currentY+rowH*0.31,w:widths[i]-0.36,h:rowH*0.30,fontSize:Math.max(8,Math.min(12,rowH*28)),bold:true,color:pctColor,align:'center',margin:0,fit:'shrink'});
              } else if(i===4) {
                slide.addShape('roundRect',{x:cx+0.18,y:currentY+rowH*0.13,w:widths[i]-0.36,h:rowH*0.70,fill:{color:backlogFill},line:{color:backlogColor,pt:0.8}});
                slide.addText(String(v),{x:cx+0.18,y:currentY+rowH*0.31,w:widths[i]-0.36,h:rowH*0.30,fontSize:Math.max(8,Math.min(12,rowH*28)),bold:true,color:backlogColor,align:'center',margin:0,fit:'shrink'});
              } else {
                slide.addText(String(v),{x:cx+0.04,y:currentY+rowH*0.31,w:widths[i]-0.08,h:rowH*0.30,fontSize:Math.max(8,Math.min(12,rowH*28)),bold:i===0,color:'172033',align:i===0?'left':'center',margin:0,fit:'shrink'});
              }
              cx += widths[i];
            });
            currentY += rowH;
          });

          // Mandatory TOTAL row for every sector, including its sector Achievement %.
          var sectorPct = sectorTotal.hmis > 0 ? Math.round(sectorTotal.rch / sectorTotal.hmis * 100) : 0;
          var totalVals = [sectorName.toUpperCase() + ' TOTAL', sectorTotal.hmis, sectorTotal.rch, sectorPct + '%', sectorTotal.backlog, sectorTotal.highRisk];
          cx = tx;
          totalVals.forEach(function(v,i) {
            var fillColor = i === 3 ? (sectorPct >= 90 ? '16A34A' : (sectorPct >= 70 ? 'F59E0B' : 'DC2626')) : '073B75';
            slide.addShape('rect',{x:cx,y:currentY,w:widths[i],h:rowH,fill:{color:'073B75'},line:{color:'FFFFFF',pt:1}});
            slide.addText(String(v),{
              x:cx+0.04,y:currentY+rowH*0.29,w:widths[i]-0.08,h:rowH*0.34,
              fontSize:Math.max(8,Math.min(12,rowH*28)),bold:true,
              color:i===3 ? 'FFFFFF' : 'FFFFFF',align:i===0?'left':'center',margin:0,fit:'shrink'
            });
            if (i === 3) {
              slide.addShape('roundRect',{x:cx+0.18,y:currentY+rowH*0.13,w:widths[i]-0.36,h:rowH*0.70,fill:{color:fillColor},line:{color:'FFFFFF',pt:0.8}});
              slide.addText(String(v),{x:cx+0.18,y:currentY+rowH*0.31,w:widths[i]-0.36,h:rowH*0.30,fontSize:Math.max(8,Math.min(12,rowH*28)),bold:true,color:'FFFFFF',align:'center',margin:0,fit:'shrink'});
            }
            cx += widths[i];
          });
          currentY += rowH;
        });

        slide.addText('Each sector includes facility-wise details followed by Sector TOTAL and Achievement %.', {
          x:0.42,y:6.72,w:9.8,h:0.18,fontSize:7.5,color:'64748B',margin:0
        });
      });

      // Any facility not covered by the three requested groups is shown separately,
      // with the same sector-total format. This includes Urban Kharsia.
      var remainingRows = rows.filter(function(r) {
        var key = clean(r.facility).toLowerCase();
        return clean(r.facility) && !assignedFacilities[key];
      });
      if (remainingRows.length) {
        showProgress('Facility slides बन रही हैं', 90, 'बाकी Facility Wise Data और Sector Total तैयार हो रहा है...');
        slide = pptx.addSlide();
        addHeader(slide, 'Janani Portal (RCH 2.0)', 'FACILITY WISE DATA | OTHER SECTORS');

        var tx2=0.38, ty2=1.05;
        var w2=[3.45,1.75,1.90,1.70,1.75,1.94];
        var h2=['Facility','HMIS PW','RCH 2.0 PW','Achievement','Backlog','High Risk'];
        var groupedOther={};
        remainingRows.forEach(function(r){
          var key=clean(r.sector)||'Other';
          if(!groupedOther[key]) groupedOther[key]=[];
          groupedOther[key].push(r);
        });
        var otherKeys=Object.keys(groupedOther);
        var totalLines=1;
        otherKeys.forEach(function(k){ totalLines += groupedOther[k].length + 2; });
        var rh2=Math.max(0.28,Math.min(0.42,5.95/Math.max(totalLines,1)));
        var cx2=tx2;
        h2.forEach(function(h,i){
          slide.addShape('rect',{x:cx2,y:ty2,w:w2[i],h:rh2,fill:{color:'075985'},line:{color:'FFFFFF',pt:1}});
          slide.addText(h,{x:cx2+0.03,y:ty2+rh2*0.28,w:w2[i]-0.06,h:rh2*0.38,fontSize:Math.max(9,Math.min(12,rh2*30)),bold:true,color:'FFFFFF',align:'center',margin:0,fit:'shrink'});
          cx2+=w2[i];
        });
        var cy2=ty2+rh2;
        otherKeys.forEach(function(sectorName){
          var sr=groupedOther[sectorName];
          slide.addShape('rect',{x:tx2,y:cy2,w:12.49,h:rh2,fill:{color:'E0F2FE'},line:{color:'B6D7EA',pt:1}});
          slide.addText('SECTOR: '+sectorName.toUpperCase(),{x:tx2+0.08,y:cy2+rh2*0.26,w:12.25,h:rh2*0.38,fontSize:Math.max(8,Math.min(11,rh2*28)),bold:true,color:'075985',margin:0,fit:'shrink'});
          cy2+=rh2;
          var st={hmis:0,rch:0,backlog:0,highRisk:0};
          sr.forEach(function(r,ri){
            var hmis=number(r.hmis),rch=number(r.rch),backlog=number(r.backlog),hr=number(r.highRisk);
            st.hmis+=hmis;st.rch+=rch;st.backlog+=backlog;st.highRisk+=hr;
            var pct=hmis>0?Math.round(rch/hmis*100):0;
            var pc=pct>=90?'16A34A':(pct>=70?'F59E0B':'DC2626');
            var pf=pct>=90?'DCFCE7':(pct>=70?'FEF3C7':'FEE2E2');
            var bc=backlog>0?'16A34A':(backlog===0?'CA8A04':'DC2626');
            var bf=backlog>0?'DCFCE7':(backlog===0?'FEF9C3':'FEE2E2');
            var vals=[clean(r.facility),hmis,rch,pct+'%',backlog,hr];
            cx2=tx2;
            vals.forEach(function(v,i){
              slide.addShape('rect',{x:cx2,y:cy2,w:w2[i],h:rh2,fill:{color:ri%2?'F8FBFF':'FFFFFF'},line:{color:'D7E2EA',pt:0.7}});
              if(i===3){
                slide.addShape('roundRect',{x:cx2+0.18,y:cy2+rh2*0.13,w:w2[i]-0.36,h:rh2*0.70,fill:{color:pf},line:{color:pc,pt:0.8}});
                slide.addText(String(v),{x:cx2+0.18,y:cy2+rh2*0.31,w:w2[i]-0.36,h:rh2*0.30,fontSize:Math.max(8,Math.min(12,rh2*28)),bold:true,color:pc,align:'center',margin:0,fit:'shrink'});
              } else if(i===4){
                slide.addShape('roundRect',{x:cx2+0.18,y:cy2+rh2*0.13,w:w2[i]-0.36,h:rh2*0.70,fill:{color:bf},line:{color:bc,pt:0.8}});
                slide.addText(String(v),{x:cx2+0.18,y:cy2+rh2*0.31,w:w2[i]-0.36,h:rh2*0.30,fontSize:Math.max(8,Math.min(12,rh2*28)),bold:true,color:bc,align:'center',margin:0,fit:'shrink'});
              } else {
                slide.addText(String(v),{x:cx2+0.04,y:cy2+rh2*0.31,w:w2[i]-0.08,h:rh2*0.30,fontSize:Math.max(8,Math.min(12,rh2*28)),bold:i===0,color:'172033',align:i===0?'left':'center',margin:0,fit:'shrink'});
              }
              cx2+=w2[i];
            });
            cy2+=rh2;
          });
          var sp=st.hmis>0?Math.round(st.rch/st.hmis*100):0;
          var tv=[sectorName.toUpperCase()+' TOTAL',st.hmis,st.rch,sp+'%',st.backlog,st.highRisk];
          cx2=tx2;
          tv.forEach(function(v,i){
            slide.addShape('rect',{x:cx2,y:cy2,w:w2[i],h:rh2,fill:{color:'073B75'},line:{color:'FFFFFF',pt:1}});
            if(i===3){
              var tc=sp>=90?'16A34A':(sp>=70?'F59E0B':'DC2626');
              slide.addShape('roundRect',{x:cx2+0.18,y:cy2+rh2*0.13,w:w2[i]-0.36,h:rh2*0.70,fill:{color:tc},line:{color:'FFFFFF',pt:0.8}});
              slide.addText(String(v),{x:cx2+0.18,y:cy2+rh2*0.31,w:w2[i]-0.36,h:rh2*0.30,fontSize:Math.max(8,Math.min(12,rh2*28)),bold:true,color:'FFFFFF',align:'center',margin:0,fit:'shrink'});
            } else {
              slide.addText(String(v),{x:cx2+0.04,y:cy2+rh2*0.31,w:w2[i]-0.08,h:rh2*0.30,fontSize:Math.max(8,Math.min(12,rh2*28)),bold:true,color:'FFFFFF',align:i===0?'left':'center',margin:0,fit:'shrink'});
            }
            cx2+=w2[i];
          });
          cy2+=rh2;
        });
      }

      moduleRanges.push({
        name: 'RCH 2.0',
        start: rchStartPage,
        end: pptx.slides.length + 1
      });

      // 6 onward. OTHER MODULE REPORTS — live Google Sheet + colorful chart
      showProgress('Other module reports', 91, 'बाकी सभी modules के reports और color charts तैयार हो रहे हैं...');
      for (var mi = 0; mi < otherModules.length; mi++) {
        var mod = otherModules[mi];
        var progress = 91 + Math.round((mi / Math.max(otherModules.length,1)) * 4);
        showProgress('Module report ' + (mi + 1) + '/' + otherModules.length, progress, mod.name + ' का live data और color chart तैयार हो रहा है...');
        var moduleStart = pptx.slides.length + 2;
        var genericData = await queryGenericSheet(mod.gid);
        addGenericModuleSlide(pptx, mod, genericData, mi);
        var moduleEnd = pptx.slides.length + 1;
        if (moduleEnd >= moduleStart) {
          moduleRanges.push({ name: mod.name, start: moduleStart, end: moduleEnd });
        }
      }

      showProgress('Final analysis', 92, 'Overall Data Analysis slide तैयार हो रही है...');
      // OVERALL ANALYSIS / KEY OBSERVATIONS slide removed as requested.
      // Build the final Index as a full-bleed visual slide to match the approved sample style.
      var indexSlide = pptx.addSlide();
      indexSlide.background = { color:'F7FBFF' };

      function pageRangeText(item) {
        return item.start === item.end
          ? 'Page ' + String(item.start).padStart(2,'0')
          : 'Page ' + String(item.start).padStart(2,'0') + '–' + String(item.end).padStart(2,'0');
      }

      var desiredModules = [
        {name:'RCH 2.0', icon:'R', color:'#7C3AED', fill:'#EDE9FE'},
        {name:'Ayushman Card', icon:'A', color:'#DB2777', fill:'#FCE7F3'},
        {name:'NCD', icon:'♥', color:'#0284C7', fill:'#E0F2FE'},
        {name:'Ayushman Arogya Mandir (AAM)', icon:'+', color:'#16A34A', fill:'#DCFCE7'},
        {name:'JAS Meeting', icon:'J', color:'#D97706', fill:'#FEF3C7'},
        {name:'Health & Wellness Centre', icon:'H', color:'#EA580C', fill:'#FFEDD5'},
        {name:'Ayushman Shivir', icon:'S', color:'#DB2777', fill:'#FCE7F3'},
        {name:'Wellness Activity', icon:'W', color:'#9333EA', fill:'#F3E8FF'},
        {name:'RBSK', icon:'B', color:'#2563EB', fill:'#DBEAFE'},
        {name:'NRC Kharsia', icon:'N', color:'#16A34A', fill:'#DCFCE7'},
        {name:'Blindness Control', icon:'E', color:'#CA8A04', fill:'#FEF9C3'},
        {name:'NQAS Certification', icon:'Q', color:'#EA580C', fill:'#FFEDD5'},
        {name:'Dialysis', icon:'D', color:'#DB2777', fill:'#FCE7F3'},
      ];
      desiredModules.push({name:'NLEP', icon:'L', color:'#0284C7', fill:'#E0F2FE'});

      var indexItems = desiredModules.map(function(dm) {
        var found = moduleRanges.find(function(r){ return r.name === dm.name; });
        return {name:dm.name, icon:dm.icon, color:dm.color, fill:dm.fill,
                start:found ? found.start : null, end:found ? found.end : null};
      });

      function esc(v) {
        return String(v == null ? '' : v)
          .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
          .replace(/"/g,'&quot;');
      }

      var svg = '<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 1600 900">' +
        '<defs>' +
          '<linearGradient id="bg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#F8FCFF"/><stop offset="1" stop-color="#EAF5FF"/></linearGradient>' +
          '<linearGradient id="navy" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#073B75"/><stop offset="1" stop-color="#075985"/></linearGradient>' +
          '<filter id="shadow"><feDropShadow dx="0" dy="5" stdDeviation="7" flood-opacity=".14"/></filter>' +
          '<clipPath id="idxPhoto"><path d="M1190 165 C1370 115 1515 145 1600 235 L1600 690 C1480 750 1320 720 1190 650 Z"/></clipPath>' +
        '</defs>' +
        '<rect width="1600" height="900" fill="url(#bg)"/>' +
        '<rect width="1600" height="138" fill="#FFFFFF"/>' +
        '<rect y="132" width="1600" height="7" fill="#59AEE8"/>' +
        '<text x="70" y="57" font-family="Arial, Noto Sans, sans-serif" font-size="28" font-weight="800" fill="#073B75">छत्तीसगढ़ शासन</text>' +
        '<text x="70" y="99" font-family="Arial, Noto Sans, sans-serif" font-size="22" font-weight="700" fill="#1F4E79">स्वास्थ्य एवं परिवार कल्याण विभाग</text>' +
        '<circle cx="35" cy="77" r="27" fill="#FFFFFF" stroke="#075985" stroke-width="4"/><text x="35" y="87" text-anchor="middle" font-family="Arial" font-size="28" font-weight="900" fill="#075985">+</text>' +
        '<text x="1470" y="53" text-anchor="middle" font-family="Arial" font-size="17" font-weight="800" fill="#0B4F6C">NATIONAL HEALTH MISSION</text>' +
        '<circle cx="1470" cy="91" r="22" fill="#EF4444"/><text x="1470" y="99" text-anchor="middle" font-family="Arial" font-size="22" font-weight="900" fill="#FFFFFF">+</text>' +
        '<rect x="65" y="166" width="1020" height="86" rx="24" fill="url(#navy)" filter="url(#shadow)"/>' +
        '<text x="575" y="228" text-anchor="middle" font-family="Arial" font-size="54" font-weight="900" fill="#FFFFFF">INDEX</text>' +
        '<text x="1110" y="222" text-anchor="middle" font-family="Arial" font-size="17" font-weight="800" fill="#075985">PROGRESSIVE REPORT • FY 2026–27</text>';

      indexItems.forEach(function(item,i){
        var col = i < 7 ? 0 : 1;
        var row = i < 7 ? i : i - 7;
        var x = col === 0 ? 55 : 590;
        var y = 278 + row * 65;
        var page = item.start ? pageRangeText(item) : 'Data Pending';
        svg += '<g filter="url(#shadow)">' +
          '<rect x="'+x+'" y="'+y+'" width="510" height="53" rx="14" fill="'+item.fill+'" stroke="'+item.color+'" stroke-opacity=".28" stroke-width="2"/>' +
          '<rect x="'+(x+9)+'" y="'+(y+8)+'" width="42" height="37" rx="10" fill="'+item.color+'"/>' +
          '<text x="'+(x+30)+'" y="'+(y+32)+'" text-anchor="middle" font-family="Arial" font-size="16" font-weight="900" fill="#FFFFFF">'+String(i+1).padStart(2,'0')+'</text>' +
          '<circle cx="'+(x+76)+'" cy="'+(y+26)+'" r="18" fill="#FFFFFF" stroke="'+item.color+'" stroke-width="2"/>' +
          '<text x="'+(x+76)+'" y="'+(y+33)+'" text-anchor="middle" font-family="Arial" font-size="18" font-weight="900" fill="'+item.color+'">'+esc(item.icon)+'</text>' +
          '<text x="'+(x+108)+'" y="'+(y+33)+'" font-family="Arial, Noto Sans, sans-serif" font-size="'+(item.name.length>25?18:20)+'" font-weight="800" fill="#0F2D52">'+esc(item.name)+'</text>' +
          '<rect x="'+(x+395)+'" y="'+(y+9)+'" width="105" height="35" rx="11" fill="#FFFFFF" stroke="'+item.color+'" stroke-opacity=".30"/>' +
          '<text x="'+(x+447)+'" y="'+(y+31)+'" text-anchor="middle" font-family="Arial" font-size="12" font-weight="800" fill="'+(item.start?'#075985':'#C2410C')+'">'+esc(page)+'</text>' +
        '</g>';
      });

      svg += '__PHOTO__' +
        '<path d="M1190 165 C1370 115 1515 145 1600 235 L1600 690 C1480 750 1320 720 1190 650 Z" fill="none" stroke="#FFFFFF" stroke-width="12"/>' +
        '<rect x="1170" y="570" width="360" height="72" rx="18" fill="#073B75" fill-opacity=".93"/>' +
        '<text x="1350" y="600" text-anchor="middle" font-family="Arial" font-size="18" font-weight="900" fill="#FFFFFF">BLOCK KHARSIA</text>' +
        '<text x="1350" y="626" text-anchor="middle" font-family="Arial" font-size="16" font-weight="700" fill="#D9F2FF">DISTRICT RAIGARH • CHHATTISGARH</text>' +
        '<path d="M1080 690 C1250 625 1450 680 1600 610 L1600 820 L1080 820 Z" fill="#DCEFFF" opacity=".9"/>' +
        '<rect y="820" width="1600" height="80" fill="#073B75"/>' +
        '<text x="65" y="869" font-family="Arial" font-size="21" font-weight="800" fill="#FFFFFF">Health Department | District Raigarh | Chhattisgarh</text>' +
        '<text x="1255" y="869" text-anchor="end" font-family="Arial" font-size="19" font-weight="900" fill="#FFD900">Monthly Performance Review | FY 2026–27</text>' +
        '<rect x="1390" y="838" width="145" height="42" rx="16" fill="#FFFFFF"/><text x="1462" y="865" text-anchor="middle" font-family="Arial" font-size="19" font-weight="900" fill="#073B75">Page 02</text>' +
        '</svg>';

      var photoSvg = photoHref
        ? '<image href="'+photoHref+'" x="1115" y="150" width="485" height="540" preserveAspectRatio="xMidYMid slice" clip-path="url(#idxPhoto)"/>'
        : '<path d="M1190 165 C1370 115 1515 145 1600 235 L1600 690 C1480 750 1320 720 1190 650 Z" fill="#DDF2FF"/>';
      svg = svg.replace('__PHOTO__', photoSvg);

      indexSlide.addImage({
        data:'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svg))),
        x:0, y:0, w:13.333, h:7.5
      });

      // Move the completed index to slide 2 (after the cover).
      pptx.slides.splice(1, 0, pptx.slides.pop());

      // Add visible page numbers to every slide. Cover already has its own footer area,
      // while all report slides use the native PptxGenJS slide-number field.
      pptx.slides.forEach(function(s, i) {
        if (i === 0) return;
        s.slideNumber = {
          x:12.15, y:7.03, w:0.75, h:0.20,
          fontFace:'Arial', fontSize:8, bold:true,
          color:'0B4F6C', align:'right'
        };
      });

      var date = new Date().toISOString().slice(0, 10);
      var fileName = 'Kharsia Health Progressive Report_' + date + '.pptx';
      showProgress('PPTX file बन रही है', 96, 'PowerPoint file को final .pptx format में बनाया जा रहा है...');
      var blob = await pptx.write({ outputType: 'blob' });
      var url = URL.createObjectURL(blob);
      var old = document.getElementById('pptxDownloadFallback');
      if (old) old.remove();

      var host = document.getElementById('dashboardPptxBtn');
      host = host ? host.parentElement : document.body;
      var box = document.createElement('div');
      box.id = 'pptxDownloadFallback';
      box.style.cssText = 'margin:10px 0;padding:12px 16px;background:#ecfdf5;border:2px solid #10b981;border-radius:10px;font-weight:700;';
      var link = document.createElement('a');
      link.href = url;
      link.download = fileName;
      link.textContent = '⬇️ Download ' + fileName;
      link.target = '_blank';
      link.rel = 'noopener';
      link.onclick = function () { showProgress('Download शुरू हो रहा है', 100, 'अगर browser पूछे तो download को Allow करें।'); };
      try { window.open(url, '_blank'); } catch (openErr) { console.warn('Auto-open blocked; use download link.', openErr); }
      link.style.cssText = 'color:#065f46;text-decoration:none;font-size:15px;';
      box.appendChild(link);
      host.appendChild(box);

      showProgress('PPTX तैयार है', 100, 'Download link नीचे दिखाई दे रहा है।');
      setButton('📊 Generate PPTX', false);
      alert('PPTX तैयार है। नीचे दिख रहे Download link पर क्लिक करें।');
    } catch (e) {
      var pb = document.getElementById('pptxProgressBox');
      if (pb) { pb.style.background='#fef2f2'; pb.style.borderColor='#ef4444'; }
      console.error(e);
      setButton('❌ PPTX में error — फिर प्रयास करें', false);
      alert('PPTX download नहीं हुआ: ' + (e && e.message ? e.message : e));
    }
  }

  window.generate = generate;
  window.generatePPTX = generate;
})();