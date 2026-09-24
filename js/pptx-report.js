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
    slide.addText('KHARSIA HEALTH DASHBOARD  •  BLOCK KHARSIA  •  DISTRICT RAIGARH', {
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

      // 1. COVER — meeting presentation style with photo
      var coverPhoto = null;
      try { coverPhoto = await fetchImageData(COVER_PHOTO_URL); } catch (photoErr) { console.warn(photoErr); }
      var slide = pptx.addSlide();
      slide.background = { color: 'F8FAFC' };
      slide.addShape('rect', { x:0,y:0,w:13.333,h:7.5,fill:{color:'F8FAFC'},line:{color:'F8FAFC'} });
      slide.addShape('rect', { x:0,y:0,w:13.333,h:0.85,fill:{color:'075985'},line:{color:'075985'} });
      slide.addShape('rect', { x:0,y:7.05,w:13.333,h:0.45,fill:{color:'075985'},line:{color:'075985'} });
      if (coverPhoto) {
        slide.addImage({ data: coverPhoto, x:7.55, y:1.15, w:5.15, h:4.65, transparency:4 });
        slide.addShape('rect',{x:7.55,y:1.15,w:5.15,h:4.65,fill:{color:'075985',transparency:78},line:{color:'FFFFFF',transparency:100}});
      } else {
        slide.addShape('roundRect',{x:7.55,y:1.15,w:5.15,h:4.65,fill:{color:'DFF7F3'},line:{color:'14B8A6',pt:1}});
      }
      slide.addText('KHARSIA HEALTH DASHBOARD', {x:0.65,y:1.0,w:6.35,h:0.38,fontSize:17,bold:true,color:'075985',margin:0});
      slide.addText('PROGRESSIVE\nREPORT', {x:0.65,y:1.75,w:6.4,h:1.25,fontSize:35,bold:true,color:'0F172A',margin:0,breakLine:false});
      slide.addText('Janani Portal (RCH 2.0)', {x:0.68,y:3.35,w:6.2,h:0.45,fontSize:22,bold:true,color:'0F766E',margin:0});
      slide.addText('FY 2026–27', {x:0.68,y:4.0,w:3.5,h:0.38,fontSize:18,bold:true,color:'334155',margin:0});
      slide.addText('As On Date: ' + clean(parsed.asOnDate || 'Current Date'), {x:0.68,y:4.65,w:5.9,h:0.45,fontSize:18,bold:true,color:'075985',margin:0});
      slide.addText('Block Kharsia  |  District Raigarh  |  Chhattisgarh', {x:0.68,y:5.55,w:6.3,h:0.3,fontSize:11,bold:true,color:'475569',margin:0});
      slide.addText('Healthcare photo: Unsplash / Vitaly Gariev', {x:7.65,y:5.93,w:4.9,h:0.2,fontSize:7.5,color:'64748B',align:'right',margin:0});

      // 2. INDEX — all dashboard modules
      slide = pptx.addSlide();
      addHeader(slide, 'INDEX', 'Progressive Report Modules');
      var indexItems = [
        ['01','Cover Page'],
        ['02','Index'],
        ['03','Janani Portal (RCH 2.0) — Block Summary'],
        ['04','Janani Portal — Sector Wise Data + Color Graph'],
        ['05','Janani Portal — Sector-wise Facility Data + Analysis'],
        ['06','Janani Portal — Overall Analysis']
      ];
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
      otherModules.forEach(function(m,i){ indexItems.push([String(i+7).padStart(2,'0'),m.icon+' '+m.name+' — Color Chart & Report']); });
      indexItems.forEach(function(item,i){
        var col=i<9?0:1, row=i<9?i:i-9, x=0.65+col*6.25, y=1.12+row*0.57;
        slide.addShape('roundRect',{x:x,y:y,w:0.72,h:0.36,fill:{color:i<6?'0F766E':'2563EB'},line:{color:i<6?'0F766E':'2563EB'}});
        slide.addText(item[0],{x:x,y:y+0.09,w:0.72,h:0.15,fontSize:8.5,bold:true,color:'FFFFFF',align:'center',margin:0});
        slide.addText(item[1],{x:x+0.9,y:y+0.02,w:5.0,h:0.3,fontSize:12.5,bold:i<6,color:'172033',margin:0,fit:'shrink'});
      });

      showProgress('Slides बन रही हैं', 55, 'Block Summary Dashboard तैयार हो रहा है...');
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
        fontSize: 17, bold: true, color: '172033',
        breakLine: false, margin: 0.03
      });

      showProgress('Slides बन रही हैं', 65, 'Sector Wise graph और table तैयार हो रहे हैं...');
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
          catAxisLabelFontSize: 15, valAxisLabelFontSize: 12,
          chartColors: ['0F766E','2563EB','F59E0B','DC2626','7C3AED','059669','EA580C','0891B2'], valGridLine: { color: 'D6E3EC', pt: 1 },
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

      showProgress('Slides बन रही हैं', 75, 'Sector-wise Facility Data और Analysis तैयार हो रहा है...');
      // 5 onward. FACILITY WISE BY SECTOR + ANALYSIS
      sectors.forEach(function (s, sectorIndex) {
        showProgress('Facility slides बन रही हैं', Math.min(90, 75 + Math.round((sectorIndex / Math.max(sectors.length,1)) * 15)), 'Sector ' + (sectorIndex + 1) + '/' + sectors.length + ': ' + s.sector);
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
          fontSize: 16, bold: true, color: '172033',
          breakLine: false, margin: 0.03, fit: 'shrink'
        });
        addCard(slide, 8.8, 5.1, 1.75, 1.05, 'Achievement', s.percent + '%');
        addCard(slide, 10.75, 5.1, 1.75, 1.05, 'Facilities', s.facilities);
      });

      // 6 onward. OTHER MODULE REPORTS — live Google Sheet + colorful chart
      showProgress('Other module reports', 91, 'बाकी सभी modules के reports और color charts तैयार हो रहे हैं...');
      for (var mi = 0; mi < otherModules.length; mi++) {
        var mod = otherModules[mi];
        var progress = 91 + Math.round((mi / Math.max(otherModules.length,1)) * 4);
        showProgress('Module report ' + (mi + 1) + '/' + otherModules.length, progress, mod.name + ' का live data और color chart तैयार हो रहा है...');
        var genericData = await queryGenericSheet(mod.gid);
        addGenericModuleSlide(pptx, mod, genericData, mi);
      }

      showProgress('Final analysis', 92, 'Overall Data Analysis slide तैयार हो रही है...');
      // LAST. OVERALL ANALYSIS
      slide = pptx.addSlide();
      addHeader(slide, 'Janani Portal (RCH 2.0)', 'OVERALL DATA ANALYSIS');
      var highest = sectors.slice().sort(function (a, b) { return b.percent - a.percent; })[0];
      var lowest = sectors.slice().sort(function (a, b) { return a.percent - b.percent; })[0];
      var maxBacklog = sectors.slice().sort(function (a, b) { return b.backlog - a.backlog; })[0];
      var minBacklog = sectors.slice().sort(function (a, b) { return a.backlog - b.backlog; })[0];
      addSectionTitle(slide, 'KEY OBSERVATIONS', 0.65, 1.35, 4.5);
      slide.addText('OBSERVATIONS', {
        x: 0.65, y: 1.72, w: 4, h: 0.25,
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
        x: 0.75, y: 2.15, w: 11.5, h: 3.15,
        fontSize: 17, color: '172033',
        breakLine: false, margin: 0.03, fit: 'shrink'
      });
      slide.addText('Source: live Google Sheet data | Generated from the Dashboard PPTX button', {
        x: 0.65, y: 6.75, w: 8, h: 0.2,
        fontSize: 8, color: '94A3B8', margin: 0
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