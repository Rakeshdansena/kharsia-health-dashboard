/* Kharsia Health Dashboard — Progressive PPTX Report
   Stable standalone generator for the Dashboard PPTX button.
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
    slide.background = { color: 'F7FAFC' };
    slide.addShape('rect', {
      x: 0, y: 0, w: 13.333, h: 0.55,
      fill: { color: '075985' }, line: { color: '075985' }
    });
    slide.addText(title, {
      x: 0.45, y: 0.72, w: 12.2, h: 0.42,
      fontSize: 23, bold: true, color: '0F172A', margin: 0
    });
    slide.addText(subtitle || '', {
      x: 0.45, y: 1.15, w: 12.2, h: 0.28,
      fontSize: 10, color: '475569', margin: 0
    });
    slide.addText('Kharsia Health Dashboard | Block Kharsia | District Raigarh | Chhattisgarh', {
      x: 0.45, y: 7.15, w: 12, h: 0.18,
      fontSize: 7.5, color: '64748B', align: 'right', margin: 0
    });
  }

  function addCard(slide, x, y, w, h, label, value) {
    slide.addShape('roundRect', {
      x: x, y: y, w: w, h: h,
      fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', pt: 1 }
    });
    slide.addText(label, {
      x: x + 0.08, y: y + 0.12, w: w - 0.16, h: 0.24,
      fontSize: 8.5, bold: true, color: '64748B',
      align: 'center', margin: 0, fit: 'shrink'
    });
    slide.addText(String(value), {
      x: x + 0.08, y: y + 0.46, w: w - 0.16, h: 0.4,
      fontSize: 20, bold: true, color: '075985',
      align: 'center', margin: 0, fit: 'shrink'
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

  async function generate() {
    setButton('⏳ PPTX तैयार हो रहा है...', true);
    try {
      await waitForGoogle();
      await loadPptx();
      var parsed = await getRchData();
      var rows = parsed.facilityRows || [];
      var sectors = groupSectors(rows);
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

      // 1. COVER
      var slide = pptx.addSlide();
      slide.background = { color: '075985' };
      slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: 13.333, h: 7.5,
        fill: { color: '075985' }, line: { color: '075985' }
      });
      slide.addText('KHARSIA HEALTH DASHBOARD', {
        x: 0.7, y: 1.0, w: 11.9, h: 0.35,
        fontSize: 14, bold: true, color: 'BAE6FD',
        align: 'center', charSpacing: 2, margin: 0
      });
      slide.addText('PROGRESSIVE REPORT', {
        x: 0.7, y: 1.75, w: 11.9, h: 0.75,
        fontSize: 38, bold: true, color: 'FFFFFF', align: 'center', margin: 0
      });
      slide.addText('Janani Portal (RCH 2.0)', {
        x: 0.7, y: 2.65, w: 11.9, h: 0.42,
        fontSize: 21, color: 'E0F2FE', align: 'center', margin: 0
      });
      slide.addText('FY 2026–27', {
        x: 0.7, y: 3.25, w: 11.9, h: 0.35,
        fontSize: 16, bold: true, color: 'FFFFFF', align: 'center', margin: 0
      });
      slide.addText('As On Date: ' + clean(parsed.asOnDate || 'Current Date'), {
        x: 2.5, y: 4.2, w: 8.3, h: 0.5,
        fontSize: 18, bold: true, color: 'FFFFFF',
        align: 'center', margin: 0
      });
      slide.addText('Block Kharsia | District Raigarh | Chhattisgarh', {
        x: 0.7, y: 6.55, w: 11.9, h: 0.28,
        fontSize: 12, color: 'BAE6FD', align: 'center', margin: 0
      });

      // 2. INDEX
      slide = pptx.addSlide();
      addHeader(slide, 'INDEX', 'Progressive report structure');
      [
        ['01', 'Cover Page'],
        ['02', 'Index'],
        ['03', 'Block Summary Dashboard'],
        ['04', 'Sector Wise Data + Graph'],
        ['05 onward', 'Sector-wise Facility Data + Analysis'],
        ['Last', 'Overall Data Analysis']
      ].forEach(function (item, i) {
        var y = 1.55 + i * 0.78;
        slide.addShape(pptx.ShapeType.roundRect, {
          x: 1, y: y, w: 1.0, h: 0.45,
          fill: { color: '0F766E' }, line: { color: '0F766E' }
        });
        slide.addText(item[0], {
          x: 1, y: y + 0.1, w: 1, h: 0.2,
          fontSize: 9, bold: true, color: 'FFFFFF',
          align: 'center', margin: 0
        });
        slide.addText(item[1], {
          x: 2.2, y: y + 0.05, w: 9.7, h: 0.28,
          fontSize: 15, bold: i === 2,
          color: '0F172A', margin: 0
        });
      });

      // 3. BLOCK SUMMARY
      slide = pptx.addSlide();
      addHeader(slide, 'Janani Portal (RCH 2.0)', 'BLOCK SUMMARY DASHBOARD | FY 2026–27');
      addCard(slide, 0.65, 1.6, 2.7, 1.25, 'Facilities', total.facilities);
      addCard(slide, 3.55, 1.6, 2.7, 1.25, 'HMIS PW Registration', total.hmis);
      addCard(slide, 6.45, 1.6, 2.7, 1.25, 'RCH 2.0 PW Registration', total.rch);
      addCard(slide, 9.35, 1.6, 2.7, 1.25, 'Achievement %', total.percent + '%');
      addCard(slide, 2.1, 3.35, 2.7, 1.25, 'Temporary Registration', total.temp);
      addCard(slide, 5.25, 3.35, 2.7, 1.25, 'Backlog', total.backlog);
      addCard(slide, 8.4, 3.35, 2.7, 1.25, 'High Risk', total.highRisk);
      slide.addText('KEY POINTS', {
        x: 0.65, y: 5.15, w: 3.2, h: 0.35,
        fontSize: 18, bold: true, color: '0F766E', margin: 0
      });
      slide.addText([
        '• Registration achievement: ' + total.percent + '%.',
        '• Total RCH 2.0 registrations: ' + total.rch + '.',
        '• Total backlog: ' + total.backlog + '.',
        '• High-risk cases: ' + total.highRisk + '.'
      ].join('\n'), {
        x: 0.75, y: 5.65, w: 11.2, h: 1.05,
        fontSize: 14, bold: true, color: '172033',
        breakLine: false, margin: 0.03
      });

      // 4. SECTOR WISE + GRAPH
      slide = pptx.addSlide();
      addHeader(slide, 'Janani Portal (RCH 2.0)', 'SECTOR WISE DATA + GRAPH');
      var chartRows = sectors.slice().sort(function (a, b) { return b.percent - a.percent; });
      try {
        slide.addChart(pptx.ChartType.bar, [{
          name: 'Achievement %',
          labels: chartRows.map(function (s) { return s.sector; }),
          values: chartRows.map(function (s) { return s.percent; })
        }], {
          x: 0.55, y: 1.5, w: 7.1, h: 4.9,
          showLegend: false, showTitle: false, showValue: true,
          catAxisLabelFontSize: 11, valAxisLabelFontSize: 9,
          chartColors: ['0F766E'], valGridLine: { color: 'D6E3EC', pt: 1 },
          valAxisMinVal: 0, dataLabelPosition: 'outEnd'
        });
      } catch (e) {
        addTable(slide, [['Sector', 'Achievement %']].concat(chartRows.map(function (s) {
          return [s.sector, s.percent + '%'];
        })), 0.65, 1.55, 6.7, 4.7);
      }
      var sectorTable = [['Sector', 'HMIS PW', 'RCH 2.0 PW', '%', 'Backlog', 'High Risk']];
      chartRows.forEach(function (s) {
        sectorTable.push([s.sector, s.hmis, s.rch, s.percent + '%', s.backlog, s.highRisk]);
      });
      addTable(slide, sectorTable, 7.9, 1.5, 4.85, 4.95);

      // 5 onward. FACILITY WISE BY SECTOR + ANALYSIS
      sectors.forEach(function (s) {
        slide = pptx.addSlide();
        addHeader(slide, 'Sector: ' + s.sector, 'FACILITY WISE DATA + ANALYSIS');
        var facilityRows = rows.filter(function (r) { return clean(r.sector) === s.sector; });
        var tableRows = [['SN', 'Facility', 'HMIS PW', 'RCH 2.0 PW', '%', 'Temp.', 'Backlog', 'High Risk']];
        facilityRows.forEach(function (r, i) {
          var hmis = number(r.hmis);
          var rch = number(r.rch);
          var pct = hmis > 0 ? Math.round(rch / hmis * 100) : 0;
          tableRows.push([
            i + 1, clean(r.facility), hmis, rch, pct + '%',
            number(r.temp), number(r.backlog), number(r.highRisk)
          ]);
        });
        addTable(slide, tableRows.slice(0, 20), 0.45, 1.45, 8.0, 5.35);
        slide.addText('ANALYSIS', {
          x: 8.75, y: 1.5, w: 3.5, h: 0.35,
          fontSize: 19, bold: true, color: '0F766E', margin: 0
        });
        slide.addText(analysis(s).map(function (x) { return '• ' + x; }).join('\n'), {
          x: 8.75, y: 2.0, w: 3.9, h: 2.9,
          fontSize: 13, bold: true, color: '172033',
          breakLine: false, margin: 0.03, fit: 'shrink'
        });
        addCard(slide, 8.8, 5.1, 1.75, 1.05, 'Achievement', s.percent + '%');
        addCard(slide, 10.75, 5.1, 1.75, 1.05, 'Facilities', s.facilities);
      });

      // LAST. OVERALL ANALYSIS
      slide = pptx.addSlide();
      addHeader(slide, 'Janani Portal (RCH 2.0)', 'OVERALL DATA ANALYSIS');
      var highest = sectors.slice().sort(function (a, b) { return b.percent - a.percent; })[0];
      var lowest = sectors.slice().sort(function (a, b) { return a.percent - b.percent; })[0];
      var maxBacklog = sectors.slice().sort(function (a, b) { return b.backlog - a.backlog; })[0];
      var minBacklog = sectors.slice().sort(function (a, b) { return a.backlog - b.backlog; })[0];
      slide.addText('OBSERVATIONS', {
        x: 0.65, y: 1.5, w: 4, h: 0.35,
        fontSize: 20, bold: true, color: '0F766E', margin: 0
      });
      slide.addText([
        '• Block achievement: ' + total.percent + '%.',
        '• Total facilities: ' + total.facilities + '.',
        '• Sector-wise achievement values range from ' + (lowest ? lowest.percent : 0) + '% to ' + (highest ? highest.percent : 0) + '%.',
        '• Highest reported sector achievement: ' + (highest ? highest.sector : 'N/A') + ' (' + (highest ? highest.percent : 0) + '%).',
        '• Lowest reported sector achievement: ' + (lowest ? lowest.sector : 'N/A') + ' (' + (lowest ? lowest.percent : 0) + '%).',
        '• Largest positive sector backlog: ' + (maxBacklog ? maxBacklog.sector : 'N/A') + ' (' + (maxBacklog ? maxBacklog.backlog : 0) + ').',
        '• Lowest sector backlog: ' + (minBacklog ? minBacklog.sector : 'N/A') + ' (' + (minBacklog ? minBacklog.backlog : 0) + ').'
      ].join('\n'), {
        x: 0.75, y: 2.0, w: 11.5, h: 3.2,
        fontSize: 15, color: '172033',
        breakLine: false, margin: 0.03, fit: 'shrink'
      });
      slide.addText('Source: live Google Sheet data | Generated from the Dashboard PPTX button', {
        x: 0.65, y: 6.75, w: 8, h: 0.2,
        fontSize: 8, color: '94A3B8', margin: 0
      });

      var date = new Date().toISOString().slice(0, 10);
      var fileName = 'Kharsia Health Progressive Report_' + date + '.pptx';
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
      link.style.cssText = 'color:#065f46;text-decoration:none;font-size:15px;';
      box.appendChild(link);
      host.appendChild(box);

      setButton('📊 Generate PPTX', false);
      alert('PPTX तैयार है। नीचे दिख रहे Download link पर क्लिक करें।');
    } catch (e) {
      console.error(e);
      setButton('❌ PPTX में error — फिर प्रयास करें', false);
      alert('PPTX download नहीं हुआ: ' + (e && e.message ? e.message : e));
    }
  }

  window.generate = generate;
  window.generatePPTX = generate;
})();