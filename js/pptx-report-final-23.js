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
      function ready() {
        return !!(window.google && google.visualization && google.visualization.Query);
      }

      // The dashboard page does not need to preload Google Charts.
      // Load it here when the user clicks Generate PPTX, then wait until
      // the Visualization Query API is actually ready.
      if (ready()) {
        resolve();
        return;
      }

      var existing = document.querySelector('script[data-kharsia-google-charts="1"]');
      if (!existing) {
        existing = document.createElement('script');
        existing.src = 'https://www.gstatic.com/charts/loader.js';
        existing.async = true;
        existing.setAttribute('data-kharsia-google-charts', '1');
        document.head.appendChild(existing);
      }

      var tries = 0;
      var chartLoadStarted = false;

      function check() {
        if (ready()) {
          resolve();
          return;
        }

        if (window.google && google.charts && !chartLoadStarted) {
          chartLoadStarted = true;
          try {
            google.charts.load('current', { packages: ['corechart', 'table'] });
            google.charts.setOnLoadCallback(function () {
              if (ready()) resolve();
            });
          } catch (e) {
            // Keep polling; the loader can become ready asynchronously.
          }
        }

        tries++;
        if (tries > 120) {
          reject(new Error('Google Sheets service उपलब्ध नहीं है। Internet connection या Google Charts loading check करें।'));
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
      x:8.05,y:6.7,w:4.5,h:0.2,fontSize:8,color:'64748B',margin:0
    });
  }


  function queryNCDSheet() {
    return new Promise(function(resolve, reject) {
      var gid = '1254412412';
      var q = new google.visualization.Query(
        'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(SHEET_ID) +
        '/gviz/tq?gid=' + gid + '&headers=0'
      );
      q.setQuery('select *');
      q.send(function(response) {
        try {
          if (response.isError()) throw new Error('NCD data load नहीं हुआ: ' + response.getMessage());
          var dt = response.getDataTable(), rs = [];
          for (var r=0;r<dt.getNumberOfRows();r++) {
            var row=[];
            for (var col=0;col<dt.getNumberOfColumns();col++) {
              row.push(clean(dt.getFormattedValue(r,col) || dt.getValue(r,col) || ''));
            }
            rs.push(row);
          }
          resolve(rs);
        } catch(e) { reject(e); }
      });
    });
  }

  function ncdIsTitle(r) {
    return /NCD\s*status\s*2026-27/i.test((r||[]).join(' '));
  }
  function ncdIsHeader(r) {
    return /subcenter|sector|facility|total\s+population|screening\s+target|enrollment\s*30|abha\s*link|estimated|under\s+treatment|follow.?up|under\s+control/i.test((r||[]).join(' ').toLowerCase());
  }
  function ncdIsTotal(r) {
    return /^(total|योग|कुल)$/i.test(clean((r||[])[0]));
  }
  function ncdUpdated(rs) {
    for (var i=0;i<rs.length;i++) if(ncdIsTitle(rs[i])) {
      var t=clean(rs[i].join(' '));
      return t || 'NCD Status 2026-27';
    }
    return 'NCD Status 2026-27';
  }
  function ncdRowsAfterHeader(rs, start, end) {
    var data=[], total=null;
    for(var i=start;i<end;i++) {
      if(!rs[i] || !rs[i].some(function(v){return clean(v)!=='';})) continue;
      if(ncdIsHeader(rs[i])) continue;
      if(ncdIsTotal(rs[i])) { total=rs[i]; break; }
      if(rs[i].length>=2) data.push(rs[i]);
    }
    return {data:data,total:total};
  }
  function ncdLabels(rs, hi, fallback) {
    var labels=[];
    for(var c=0;c<25;c++) labels[c]=clean((rs[hi]||[])[c]) || fallback[c] || ('Column '+(c+1));
    return labels;
  }

  function addNCDTableSlide(pptx, title, subtitle, labels, dataRows, totalRow, part, groupTotals) {
    var slide=pptx.addSlide();
    addHeader(slide, '❤️ NCD — '+title, subtitle);
    var idx = part==='enroll' ? [0,1,2,3,4,5,6] :
              part==='htn' ? [0,7,8,9,10,11,12,13,14,15] :
              [0,16,17,18,19,20,21,22,23,24];
    var fallback = part==='enroll'
      ? ['Sector / Facility','30+ Population','Screening Target','Enrollment 30+','Enrollment %','ABHA Link','ABHA Link %']
      : part==='htn'
      ? ['Sector / Facility','HTN Screening','HTN Screening %','Estimated Hypertensive Patient','Under Treatment','Treatment %','Follow-up','Follow-up %','Under Control','Control %']
      : ['Sector / Facility','DM Screening','DM Screening %','Estimated Diabetes Patients','Under Treatment','Treatment %','Follow-up','Follow-up %','Under Control','Control %'];
    var head = idx.map(function(i,k){ return clean((labels||[])[i]) || fallback[k]; });
    var rows=[head];
    (dataRows||[]).forEach(function(r){ rows.push(idx.map(function(i){return clean(r[i]);})); });
    if(groupTotals && groupTotals.length){
      // groupTotals are already placed in dataRows in the desired order.
    } else if(totalRow) {
      rows.push(idx.map(function(i){return clean(totalRow[i]);}));
    }
    if(rows.length<=1) return slide;

    var x=0.34, w=12.66, y=1.10;
    var dataCount=Math.max(1,rows.length-1);
    var headerH=0.70;
    // Keep every Facility Wise row, including each Sector Total, inside the slide.
    // This is especially important for the combined Barra + Jobi + Gorpar slide.
    var availableH=5.55-headerH;
    var dataH=Math.min(0.46,Math.max(0.31,availableH/dataCount));
    if (dataH * dataCount > availableH) dataH=availableH/dataCount;

    var colW=part==='enroll'
      ? [2.65,1.55,1.55,1.55,1.20,1.45,1.35]
      : [2.20,1.12,1.05,1.55,1.25,1.00,1.05,1.00,1.05,1.02];
    var sum=colW.reduce(function(a,b){return a+b;},0), scale=w/sum;
    colW=colW.map(function(v){return v*scale;});

    var yy=y;
    rows.forEach(function(row,ri){
      var rowH=ri===0?headerH:dataH, xx=x;
      var isGroupTotal = ri>0 && /\bTotal\b/i.test(String(row[0]||''));
      row.forEach(function(v,ci){
        var fill=ri===0?'075985':(isGroupTotal?'0F766E':(ri%2?'FFFFFF':'F8FBFF'));
        slide.addShape('rect',{x:xx,y:yy,w:colW[ci],h:rowH,fill:{color:fill},line:{color:'CBD5E1',pt:0.7}});
        var isPct=/%/.test(head[ci])||/percent/i.test(head[ci]);
        var txtColor=ri===0?'FFFFFF':'172033';
        slide.addText(String(v||''),{
          x:xx+0.035,y:yy+0.04,w:colW[ci]-0.07,h:rowH-0.08,
          fontSize:ri===0?11.5:Math.max(11,Math.min(14,dataH*30)),
          bold:ri===0||isGroupTotal||ci===0,
          color:isGroupTotal?'FFFFFF':txtColor,align:ci===0?'left':'center',valign:'mid',margin:0.01,fit:'shrink',breakLine:false
        });
        if(ri>0&&!isGroupTotal&&isPct&&v!==''){
          var n=number(String(v).replace('%',''));
          var pc=n>=80?'16A34A':(n>=50?'CA8A04':'DC2626');
          var pf=n>=80?'DCFCE7':(n>=50?'FEF3C7':'FEE2E2');
          slide.addShape('roundRect',{x:xx+0.12,y:yy+rowH*0.16,w:colW[ci]-0.24,h:rowH*0.68,fill:{color:pf},line:{color:pc,pt:0.7}});
          slide.addText(String(v),{x:xx+0.12,y:yy+rowH*0.30,w:colW[ci]-0.24,h:rowH*0.30,fontSize:Math.max(10,Math.min(13,dataH*30)),bold:true,color:pc,align:'center',margin:0,fit:'shrink'});
        }
        xx+=colW[ci];
      });
      yy+=rowH;
    });
    slide.addText('Source: live Google Sheet • '+ncdUpdated(window.__ncdRawRows||[]),{x:0.38,y:6.82,w:10.5,h:0.18,fontSize:7.5,color:'64748B',margin:0});
    return slide;
  }

  function ncdSectorTotalRow(rows, sectorName, sectorTotals) {
    var target=ncdNormalizeFacility(sectorName);
    var found=(sectorTotals||[]).find(function(r){return ncdNormalizeFacility(r&&r[0])===target;});
    if(found) {
      // Always mark the row explicitly as a Sector Total so the PPT renderer
      // can apply the dark-teal total-row styling.
      var foundOut=found.slice ? found.slice() : Array.prototype.slice.call(found);
      foundOut[0]=sectorName+' Total';
      return foundOut;
    }
    var out=new Array(25).fill('');
    out[0]=sectorName+' Total';
    var numeric=[1,2,3,5,7,9,10,12,14,16,18,19,21,23];
    numeric.forEach(function(c){out[c]=(rows||[]).reduce(function(sum,r){return sum+number(r[c]);},0);});
    function pct(num,den){return den?Math.round(num/den*100)+'%':'-';}
    out[4]=pct(out[3],out[2]);
    out[6]=pct(out[5],out[3]);
    out[8]=pct(out[7],out[9]);
    out[11]=pct(out[10],out[9]);
    out[13]=pct(out[12],out[10]);
    out[15]=pct(out[14],out[9]);
    out[17]=pct(out[16],out[18]);
    out[20]=pct(out[19],out[18]);
    out[22]=pct(out[21],out[19]);
    out[24]=pct(out[23],out[18]);
    return out;
  }

  function addNCDPerformanceSlide(pptx, facilityRows, mode) {
    var isTop = mode === 'top';
    var slide=pptx.addSlide();
    slide.background={color:'F7FBFF'};

    slide.addShape('roundRect',{x:0.35,y:0.22,w:12.63,h:0.78,fill:{color:isTop?'166534':'073B75'},line:{color:isTop?'166534':'073B75'}});
    slide.addText(isTop?'🏆  NCD — TOP 10 FACILITIES':'❤️  NCD — LOWEST 10 FACILITIES',{x:0.65,y:0.36,w:7.1,h:0.32,fontSize:25,bold:true,color:'FFFFFF',margin:0,fit:'shrink'});
    slide.addText('ABHA LINK • HTN SCREENING • DM SCREENING',{x:7.55,y:0.42,w:5.0,h:0.22,fontSize:11,bold:true,color:'D9F2FF',align:'right',margin:0,fit:'shrink'});
    slide.addText('FY 2026–27  |  Facility-wise percentage analysis',{x:0.65,y:0.77,w:7.0,h:0.14,fontSize:8.5,bold:true,color:'BFE8FF',margin:0});

    var metrics=[
      {title:'ABHA LINK %',sub:isTop?'Top 10':'Lowest 10',col:6,accent:isTop?'16A34A':'7C3AED',soft:isTop?'DCFCE7':'F3E8FF'},
      {title:'HTN SCREENING %',sub:isTop?'Top 10':'Lowest 10',col:8,accent:isTop?'0F766E':'EA580C',soft:isTop?'CCFBF1':'FFEDD5'},
      {title:'DM SCREENING %',sub:isTop?'Top 10':'Lowest 10',col:17,accent:isTop?'2563EB':'DC2626',soft:isTop?'DBEAFE':'FEE2E2'}
    ];
    var groups=ncdGroupFacilities(facilityRows||[]);
    var all=[];
    Object.keys(groups).forEach(function(k){(groups[k]||[]).forEach(function(r){all.push(r);});});
    var seen={};
    all=all.filter(function(r){
      var n=ncdNormalizeFacility(r&&r[0]);
      if(!n||NCD_EXCLUDED_FACILITIES[n]||n.indexOf('CHAPLE')>=0||n.indexOf('CHC CHAPLE')>=0||seen[n]) return false;
      seen[n]=true; return true;
    });

    var positions=[0.35,4.49,8.63], boxW=3.86, boxY=1.20, boxH=5.62;
    metrics.forEach(function(metric,mi){
      var rows=all.filter(function(r){
        return r && r[metric.col]!=='' && !isNaN(number(String(r[metric.col]).replace('%','')));
      }).sort(function(a,b){
        var av=number(String(a[metric.col]).replace('%','')), bv=number(String(b[metric.col]).replace('%',''));
        return isTop ? bv-av : av-bv;
      }).slice(0,10);

      var x=positions[mi];
      slide.addShape('roundRect',{x:x,y:boxY,w:boxW,h:boxH,rectRadius:0.08,fill:{color:'FFFFFF'},line:{color:metric.accent,pt:1.4},shadow:{type:'outer',color:'94A3B8',blur:1,angle:45,distance:1,opacity:0.13}});
      slide.addShape('roundRect',{x:x+0.04,y:boxY+0.04,w:boxW-0.08,h:0.62,rectRadius:0.06,fill:{color:metric.accent},line:{color:metric.accent}});
      slide.addText(metric.title,{x:x+0.18,y:boxY+0.15,w:2.5,h:0.24,fontSize:16,bold:true,color:'FFFFFF',margin:0,fit:'shrink'});
      slide.addText(metric.sub,{x:x+2.55,y:boxY+0.18,w:1.05,h:0.18,fontSize:9,bold:true,color:'FFFFFF',align:'right',margin:0});

      var rh=0.46, y=boxY+0.78;
      rows.forEach(function(r,ri){
        var pct=number(String(r[metric.col]).replace('%',''));
        var pctFill=isTop?(pct>=90?'DCFCE7':(pct>=70?'FEF3C7':'FEE2E2')):(pct<30?'FEE2E2':(pct<50?'FEF3C7':'FFF7ED'));
        var pctColor=isTop?(pct>=90?'15803D':(pct>=70?'A16207':'B91C1C')):(pct<30?'B91C1C':(pct<50?'A16207':'C2410C'));

        slide.addShape('rect',{x:x+0.10,y:y,w:boxW-0.20,h:rh,fill:{color:ri%2?'FFFFFF':'F8FBFF'},line:{color:'E2E8F0',pt:0.45}});
        slide.addShape('ellipse',{x:x+0.16,y:y+0.065,w:0.32,h:0.32,fill:{color:ri<3?metric.accent:'E2E8F0'},line:{color:ri<3?metric.accent:'CBD5E1',pt:0.5}});
        slide.addText(String(ri+1),{x:x+0.16,y:y+0.115,w:0.32,h:0.12,fontSize:8.5,bold:true,color:ri<3?'FFFFFF':'475569',align:'center',margin:0});
        slide.addText(clean(r[0]),{x:x+0.58,y:y+0.07,w:2.18,h:0.30,fontSize:14,bold:true,color:'172033',margin:0,fit:'shrink'});
        slide.addShape('roundRect',{x:x+2.82,y:y+0.055,w:0.86,h:0.34,fill:{color:pctFill},line:{color:pctColor,pt:0.8}});
        slide.addText(String(r[metric.col]),{x:x+2.86,y:y+0.105,w:0.78,h:0.16,fontSize:13,bold:true,color:pctColor,align:'center',margin:0,fit:'shrink'});
        y+=rh;
      });
      if(!rows.length) slide.addText('No facility data available',{x:x+0.35,y:3.55,w:3.15,h:0.35,fontSize:14,bold:true,color:'64748B',align:'center',margin:0});
    });

    slide.addShape('roundRect',{x:0.48,y:6.90,w:12.35,h:0.30,fill:{color:isTop?'ECFDF5':'EAF5FF'},line:{color:isTop?'BBF7D0':'CFE8F8',pt:0.6}});
    slide.addText(isTop
      ? 'Higher percentage = stronger performance  •  CHC Chaple Kharsia excluded  •  Source: live Google Sheet'
      : 'Lower percentage = higher priority for review  •  CHC Chaple Kharsia excluded  •  Source: live Google Sheet',
      {x:0.62,y:6.98,w:12.05,h:0.12,fontSize:8.5,bold:true,color:isTop?'166534':'075985',align:'center',margin:0,fit:'shrink'});
    return slide;
  }

  function addNCDLowest10Slide(pptx, facilityRows) {
    return addNCDPerformanceSlide(pptx, facilityRows, 'low');
  }

  function addNCDTop10Slide(pptx, facilityRows) {
    return addNCDPerformanceSlide(pptx, facilityRows, 'top');
  }

  function addNCDSummarySlide(pptx, labels, facilityTotal, sectorTotal, updated) {
    var slide=pptx.addSlide();
    addHeader(slide,'❤️ NCD — Summary Dashboard','NCD STATUS 2026–27');
    addSectionTitle(slide,'LAST UPDATED',0.55,1.02,3.0);
    slide.addText(updated,{x:0.55,y:1.42,w:12.1,h:0.35,fontSize:14,bold:true,color:'075985',margin:0,fit:'shrink'});
    var row=sectorTotal||facilityTotal||[];
    var cards=[
      ['30+ Population',row[1]],['Screening Target',row[2]],['Enrollment 30+',row[3]],['Enrollment %',row[4]],
      ['ABHA Link',row[5]],['ABHA Link %',row[6]],['HTN Screening',row[7]],['HTN Treatment',row[10]],
      ['HTN Control',row[14]],['DM Screening',row[16]],['DM Treatment',row[19]],['DM Control',row[23]]
    ];
    cards.forEach(function(c,i){
      var col=i%4, rr=Math.floor(i/4);
      addCard(slide,0.55+col*3.15,1.98+rr*1.05,2.82,0.82,c[0],c[1]||'-');
    });
    addSectionTitle(slide,'NCD PRESENTATION',0.55,5.30,4.0);
    slide.addText('Part 1 — Enrollment & ABHA\nPart 2 — Hypertension (HTN)\nPart 3 — Diabetes Mellitus (DM)\nSector Wise + Facility Wise data with percentage analysis',{
      x:0.62,y:5.75,w:6.2,h:1.0,fontSize:16,bold:true,color:'172033',breakLine:false,fit:'shrink',margin:0
    });
    slide.addText('Source: live Google Sheet',{x:8.0,y:6.62,w:4.0,h:0.2,fontSize:9,color:'64748B',align:'right',margin:0});
    return slide;
  }

  // Exact NCD Sector -> Facility mapping supplied by the user.
  // Exact NCD Sector -> Facility mapping supplied by the user.
  var NCD_SECTOR_FACILITY_ORDER = {
    'Barra':['FARKANARA','BARRA','DEHJARI'],
    'Binjkot':['BHUPDEOPUR','LODHAJHAR','PAMGARH','BADE DUMARPALI','BARBHAUNA','BINJKOT','GURDA','MURA','NAHARPALI'],
    'Gorpar':['KHADGAON','GANDAPALI','GORPAR'],
    'Jobi':['NAGOI','KHAMAR','NANDGAON'],
    'Sarwani':['AURDA','BARGARH','PARASKOL','BOTALDA','CHHOTE DEOGAON','MADANPUR','SARWANI','TIUR'],
    'Sondka':['BANIPATHAR','GIDHA','JAIMURA','BASNAJHAR','KUNKUNI','SONDAKA','SONBARSA','TEMTEMA'],
    'Turekela':['BAKELI','MAHUAPALI','DUMARBHATA','HALAHULI','MAKARI','TELIKOT','TUREKELA']
  };

  // Chaple is explicitly excluded from NCD Facility Wise reporting.
  var NCD_EXCLUDED_FACILITIES = {'CHAPLE':true};

  function ncdNormalizeFacility(v) {
    return clean(v).toUpperCase()
      .replace(/[^A-Z0-9]+/g,' ')
      .replace(/\b(SHC|SUBCENTER|SUB CENTRE|HEALTH FACILITY|FACILITY|SC)\b/g,' ')
      .replace(/\s+/g,' ')
      .trim();
  }

  function ncdGroupFacilities(rows) {
    var groups = {};
    var used = {};
    Object.keys(NCD_SECTOR_FACILITY_ORDER).forEach(function(sector){ groups[sector]=[]; });

    Object.keys(NCD_SECTOR_FACILITY_ORDER).forEach(function(sector){
      NCD_SECTOR_FACILITY_ORDER[sector].forEach(function(target){
        var targetN = ncdNormalizeFacility(target);
        for(var i=0;i<rows.length;i++){
          if(used[i]) continue;
          var rowN = ncdNormalizeFacility(rows[i] && rows[i][0]);
          if(NCD_EXCLUDED_FACILITIES[rowN] || rowN.indexOf('CHAPLE') === 0 || rowN.indexOf('CHAPLE') >= 0) { used[i]=true; continue; }

          // Exact match after removing SHC/punctuation, plus the two spelling variants
          // present in the user's mapping.
          var match = rowN === targetN;
          if(!match && ((rowN==='KHAMHAR' && targetN==='KHAMAR') || (rowN==='KHAMAR' && targetN==='KHAMHAR'))) match=true;
          if(!match && ((rowN==='DUMARBHANTHA' && targetN==='DUMARBHATA') || (rowN==='DUMARBHATA' && targetN==='DUMARBHANTHA'))) match=true;
          if(!match && ((rowN==='MAKRI' && targetN==='MAKARI') || (rowN==='MAKARI' && targetN==='MAKRI'))) match=true;
          if(!match && ((rowN==='SONDKA' && targetN==='SONDAKA') || (rowN==='SONDAKA' && targetN==='SONDKA'))) match=true;

          if(match){
            groups[sector].push(rows[i]);
            used[i]=true;
            break;
          }
        }
      });
    });

    // Unmatched facilities are intentionally omitted from NCD Facility Wise PPT.
    // Only the seven approved sectors are shown; Chaple is excluded.
    return groups;
  }

  async function addNCDPresentation(pptx) {
    var rs=await queryNCDSheet();
    window.__ncdRawRows=rs;

    var titleIdx=-1,nextTitle=rs.length;
    for(var i=0;i<rs.length;i++) if(ncdIsTitle(rs[i])) {titleIdx=i;break;}
    if(titleIdx<0) titleIdx=0;
    for(var j=titleIdx+1;j<rs.length;j++) if(ncdIsTitle(rs[j])) {nextTitle=j;break;}

    var h1=-1,h2=-1;
    for(var k=titleIdx+1;k<nextTitle;k++) if(ncdIsHeader(rs[k])) {h1=k;break;}
    for(var m=nextTitle+1;m<rs.length;m++) if(ncdIsHeader(rs[m])) {h2=m;break;}
    if(h1<0) throw new Error('NCD Facility Wise header नहीं मिला।');
    if(h2<0) h2=rs.length;

    var facility=ncdRowsAfterHeader(rs,h1+1,nextTitle);
    var sector=ncdRowsAfterHeader(rs,h2+1,rs.length);
    var fallback=['Sector / Facility','30+ Population','Screening Target','Enrollment 30+','Enrollment %','ABHA Link','ABHA Link %','HTN Screening','HTN Screening %','Estimated Hypertensive Patient','Under Treatment','Treatment %','Follow-up','Follow-up %','Under Control','Control %','DM Screening','DM Screening %','Estimated Diabetes Patients','Under Treatment','Treatment %','Follow-up','Follow-up %','Under Control','Control %'];
    var fl=ncdLabels(rs,h1,fallback), sl=ncdLabels(rs,Math.min(h2,rs.length-1),fallback);
    var grouped=ncdGroupFacilities(facility.data);
    var sectorTotals=sector.data||[];

    function rowsWithSectorTotal(sectorName){
      var data=(grouped[sectorName]||[]).slice();
      if(data.length) data.push(ncdSectorTotalRow(data,sectorName,sectorTotals));
      return data;
    }
    function combinedWithTotals(){
      var out=[];
      ['Barra','Jobi','Gorpar'].forEach(function(sec){
        out=out.concat(rowsWithSectorTotal(sec));
      });
      return out;
    }

    ['enroll','htn','dm'].forEach(function(part){
      var name=part==='enroll'?'Enrollment & ABHA':part==='htn'?'Hypertension (HTN)':'Diabetes Mellitus (DM)';
      addNCDTableSlide(pptx,name+' — Sector Wise','SECTOR WISE DATA | FY 2026–27',sl,sector.data,sector.total,part);

      // Barra + Jobi + Gorpar in one slide, each with its own Sector Total.
      var combined=combinedWithTotals();
      if(combined.length) addNCDTableSlide(pptx,name+' — Facility Wise | Barra • Jobi • Gorpar','FACILITY WISE DATA | FY 2026–27',fl,combined,null,part);

      ['Binjkot','Sarwani','Sondka','Turekela'].forEach(function(sec){
        var data=rowsWithSectorTotal(sec);
        if(data.length) addNCDTableSlide(pptx,name+' — Facility Wise | '+sec,'FACILITY WISE DATA | FY 2026–27',fl,data,null,part);
      });
    });

    // Final slide: lowest 10 facilities for ABHA Link %, HTN Screening %, and DM Screening %.
    addNCDLowest10Slide(pptx,facility.data);
    addNCDTop10Slide(pptx,facility.data);
    return true;
  }

  
  // ---------------- JAS MEETING PPTX ----------------
  function queryJASSheet() {
    return new Promise(function(resolve, reject) {
      var gid = '1018164338';
      var q = new google.visualization.Query(
        'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(SHEET_ID) +
        '/gviz/tq?gid=' + gid + '&headers=0'
      );
      q.setQuery('select *');
      q.send(function(response) {
        try {
          if (response.isError()) throw new Error('JAS Meeting data load नहीं हुआ: ' + response.getMessage());
          var dt=response.getDataTable(), rs=[];
          for(var r=0;r<dt.getNumberOfRows();r++){
            var row=[];
            for(var c=0;c<dt.getNumberOfColumns();c++) row.push(clean(dt.getFormattedValue(r,c) || dt.getValue(r,c) || ''));
            rs.push(row);
          }
          resolve(rs);
        } catch(e){ reject(e); }
      });
    });
  }

  function jasTitle(rs) {
    for(var i=0;i<rs.length;i++){
      for(var c=0;c<(rs[i]||[]).length;c++){
        var s=clean(rs[i][c]);
        if(/jan\s*arogya\s*samiti\s*meeting/i.test(s)) return s;
      }
    }
    return 'Jan Arogya Samiti Meeting FY 2026-27';
  }

  function jasIsFacilityHeader(r){
    var x=(r||[]).slice(0,7).join(' ').toLowerCase();
    return /sn/.test(x) && /nin/.test(x) && /sector/.test(x) && /facility/.test(x) && /target/.test(x) && /achiev/.test(x);
  }

  function jasIsSectorHeader(r){
    var x=(r||[]).slice(0,7).join(' ').toLowerCase();
    return /sn/.test(x) && /sector/.test(x) && /aam/.test(x) && /target/.test(x) && /achiev/.test(x);
  }

  function jasIsTotal(r){
    return /^(total|योग|कुल)$/i.test(clean((r||[])[0]));
  }

  function jasLooksNumber(v){
    var t=clean(v).replace(/,/g,'').replace('%','');
    return t!=='' && Number.isFinite(Number(t));
  }

  function jasNorm(v){ return clean(v).toLowerCase().replace(/[^a-z0-9]+/g,' '); }

  function jasHeaderMap(row, facilityMode){
    var m={sn:-1,nin:-1,sector:-1,facility:-1,aam:-1,target:-1,achievement:-1,pct:-1};
    (row||[]).forEach(function(v,i){
      var x=jasNorm(v);
      if(m.sn<0 && /^(sn|s no|s no\.)$/.test(x)) m.sn=i;
      if(m.nin<0 && /\\bnin\\b/.test(x)) m.nin=i;
      // "AAM Facility" appears in the Sector header too. Do not treat
      // that as the Facility-name column.
      if(m.facility<0 && /^(name of facility|facility name|facility)$/.test(x)) m.facility=i;
      if(m.sector<0 && /sector/.test(x) && !/facility/.test(x)) m.sector=i;
      if(m.aam<0 && /aam/.test(x)) m.aam=i;
      if(m.target<0 && /target/.test(x)) m.target=i;
      if(m.achievement<0 && /achiev/.test(x)) m.achievement=i;
      if(m.pct<0 && /^(%|percent|percentage|achievement %|achievement percentage)$/.test(x)) m.pct=i;
    });
    if(m.pct<0){
      for(var j=0;j<(row||[]).length;j++){ if(j!==m.target&&j!==m.achievement&&/%/.test(clean(row[j]))) {m.pct=j;break;} }
    }
    var valid=facilityMode
      ? m.sn>=0&&m.nin>=0&&m.sector>=0&&m.facility>=0&&m.target>=0&&m.achievement>=0
      : m.sn>=0&&m.sector>=0&&m.nin>=0&&m.aam>=0&&m.target>=0&&m.achievement>=0;
    return valid?m:null;
  }

  function getJASRows(rs){
    var facility=[], sector=[], facilityMap=null, sectorMap=null;
    // Prefer explicit header mapping so inserted/title/blank columns do not break PPTX.
    for(var i=0;i<rs.length;i++){
      var fm=jasHeaderMap(rs[i],true), sm=jasHeaderMap(rs[i],false);
      if(fm && !facilityMap) facilityMap=fm;
      if(sm && !sectorMap) sectorMap=sm;
    }

    function parseMapped(map, mode, out){
      if(!map) return;
      for(var r=0;r<rs.length;r++){
        var row=rs[r]||[];
        var sn=clean(row[map.sn]);
        if(!/^\\d+(?:\\.0+)?$/.test(sn)) continue;
        var target=clean(row[map.target]), ach=clean(row[map.achievement]);
        if(!jasLooksNumber(target)||!jasLooksNumber(ach)) continue;
        var pct=map.pct>=0?clean(row[map.pct]):'';
        if(!jasLooksNumber(pct)) pct=number(target)>0?(number(ach)/number(target)*100):0;
        if(mode==='facility'){
          var v=[sn,clean(row[map.nin]),clean(row[map.sector]),clean(row[map.facility]),target,ach,String(Math.round(number(pct)*10)/10)];
          if(v[1]&&v[2]&&v[3]) out.push(v);
        }else{
          var sv=[sn,clean(row[map.sector]),clean(row[map.nin]),clean(row[map.aam]),target,ach,String(Math.round(number(pct)*10)/10)];
          if(sv[1]&&jasLooksNumber(sv[3])) out.push(sv);
        }
      }
    }
    parseMapped(facilityMap,'facility',facility);
    parseMapped(sectorMap,'sector',sector);

    // Use the same robust 7-column scan as the working JAS web module.
    // Header mapping can miss rows when Google Sheets has merged/variant headers.
    var scannedFacility=[], scannedSector=[];
    for(var k=0;k<rs.length;k++){
      var v=(rs[k]||[]).slice(0,7); while(v.length<7)v.push('');
      var sn2=clean(v[0]); if(!/^\\d+(?:\\.0+)?$/.test(sn2)) continue;
      if(clean(v[1])&&clean(v[2])&&clean(v[3])&&jasLooksNumber(v[4])&&jasLooksNumber(v[5])){
        scannedFacility.push([sn2,clean(v[1]),clean(v[2]),clean(v[3]),clean(v[4]),clean(v[5]),jasLooksNumber(v[6])?clean(v[6]):String(number(v[4])?number(v[5])/number(v[4])*100:0)]);
      } else if(clean(v[1])&&jasLooksNumber(v[3])&&jasLooksNumber(v[4])&&jasLooksNumber(v[5])){
        scannedSector.push([sn2,clean(v[1]),clean(v[2]),clean(v[3]),clean(v[4]),clean(v[5]),jasLooksNumber(v[6])?clean(v[6]):String(number(v[4])?number(v[5])/number(v[4])*100:0)]);
      }
    }
    if(scannedFacility.length>facility.length) facility=scannedFacility;
    if(scannedSector.length>sector.length) sector=scannedSector;
    return {facility:facility,sector:sector};
  }

  function jasPctColor(p) {
    return p >= 100 ? '16A34A' : (p > 90 ? 'EAB308' : 'DC2626');
  }
  function jasPctSoft(p) {
    return p >= 100 ? 'DCFCE7' : (p > 90 ? 'FEF3C7' : 'FEE2E2');
  }

  function jasGroupFacilities(rows){
    var map={};
    (rows||[]).forEach(function(r){
      var s=clean(r[2])||'Other';
      if(!map[s]) map[s]=[];
      map[s].push(r);
    });
    return map;
  }

  function jasAddPctCell(slide,x,y,w,h,value){
    var p=number(value);
    var c=jasPctColor(p), soft=jasPctSoft(p);
    slide.addShape('roundRect',{x:x+0.10,y:y+h*0.13,w:w-0.20,h:h*0.74,fill:{color:soft},line:{color:c,pt:0.8}});
    slide.addText(String(value)+'%',{x:x+0.10,y:y+h*0.31,w:w-0.20,h:h*0.30,fontSize:Math.max(9,Math.min(15,h*34)),bold:true,color:'000000',align:'center',valign:'mid',margin:0,fit:'shrink'});
  }

  function jasAddSummarySlide(pptx, title, sectorRows, facilityRows){
    var slide=pptx.addSlide();
    addHeader(slide,'🤝 JAS Meeting','JAN AROGYA SAMITI | FY 2026–27');
    addSectionTitle(slide,'SUMMARY DASHBOARD',0.55,1.02,3.5);

    var sectors=sectorRows.length;
    var facilities=facilityRows.length;
    var target=sectorRows.reduce(function(s,r){return s+number(r[4]);},0);
    var ach=sectorRows.reduce(function(s,r){return s+number(r[5]);},0);
    var overall=target?ach/target*100:0;

    addCard(slide,0.55,1.45,2.35,1.12,'SECTORS',sectors);
    addCard(slide,3.05,1.45,2.35,1.12,'FACILITIES',facilities);
    addCard(slide,5.55,1.45,2.35,1.12,'TARGET',target);
    addCard(slide,8.05,1.45,2.35,1.12,'ACHIEVEMENT',ach);
    addCard(slide,10.55,1.45,2.25,1.12,'OVERALL %',overall.toFixed(1)+'%');

    addSectionTitle(slide,'SECTOR PERFORMANCE',0.55,2.95,4.0);
    var chartData=(sectorRows||[]).map(function(r){return {name:clean(r[1]),value:number(r[6])};});
    try{
      slide.addChart(pptx.ChartType.bar,[{
        name:'Achievement %',
        labels:chartData.map(function(x){return x.name;}),
        values:chartData.map(function(x){return x.value;})
      }],{
        x:0.48,y:3.35,w:7.15,h:3.05,showLegend:false,showTitle:false,showValue:true,
        catAxisLabelFontSize:12,valAxisLabelFontSize:10,chartColors:['0F766E'],
        valGridLine:{color:'D6E3EC',pt:1},valAxisMinVal:0,valAxisMaxVal:120,
        dataLabelPosition:'outEnd'
      });
    }catch(e){}

    addSectionTitle(slide,'KEY OBSERVATIONS',7.95,2.95,4.4);
    var obs=[];
    (sectorRows||[]).forEach(function(r){
      var p=number(r[6]);
      obs.push('• '+clean(r[1])+': '+p+'% achievement.');
    });
    obs.push('• Overall achievement: '+overall.toFixed(1)+'%.');
    obs.push('• 100% achievement: '+sectorRows.filter(function(r){return number(r[6])>=100;}).length+' sector(s).');
    obs.push('• Above 90% but below 100%: '+sectorRows.filter(function(r){var p=number(r[6]);return p>90&&p<100;}).length+' sector(s).');
    obs.push('• 90% or below: '+sectorRows.filter(function(r){return number(r[6])<=90;}).length+' sector(s).');
    slide.addText(obs.join('\n'),{
      x:8.0,y:3.38,w:4.65,h:2.85,fontSize:13,bold:true,color:'172033',margin:0.03,breakLine:false,fit:'shrink'
    });
    slide.addText('Source: live Google Sheet • '+title,{
      x:0.55,y:6.78,w:12.1,h:0.18,fontSize:8,color:'64748B',margin:0,fit:'shrink'
    });
    return slide;
  }

  function jasAddSectorSlide(pptx,title,sectorRows){
    var slide=pptx.addSlide();
    addHeader(slide,'🤝 JAS Meeting — Sector Wise','SECTOR WISE DATA | FY 2026–27');
    addSectionTitle(slide,'SECTOR WISE DATA',0.28,0.94,4.2);

    var x=0.22, y=1.34;
    var w=[0.62,2.62,1.72,2.20,2.20,2.34];
    var headers=['SN','Sector','No of AAM Facility','Target till Aug 26','Achievement','%'];
    var headerH=0.62;
    var dataH=Math.min(0.57,5.35/Math.max(1,sectorRows.length));
    var totalH=0.58;
    var pctVals=sectorRows.map(function(r){return number(r[6]);});
    var pmin=pctVals.length?Math.min.apply(null,pctVals):0;
    var pmax=pctVals.length?Math.max.apply(null,pctVals):0;

    function scaleColor(value){
      var p=number(value);
      if(pmax<=pmin) return {bg:'FEF3C7',line:'CA8A04'};
      var t=Math.max(0,Math.min(1,(p-pmin)/(pmax-pmin)));
      if(t<=0.5){
        var q=t*2;
        var rr=Math.round(220+(255-220)*q);
        var gg=Math.round(38+(193-38)*q);
        var bb=Math.round(38+(7-38)*q);
        return {bg:('0'+rr.toString(16)).slice(-2)+('0'+gg.toString(16)).slice(-2)+('0'+bb.toString(16)).slice(-2),line:'B91C1C'};
      }
      var q2=(t-0.5)*2;
      var rr2=Math.round(255+(22-255)*q2);
      var gg2=Math.round(193+(163-193)*q2);
      var bb2=Math.round(7+(74-7)*q2);
      return {bg:('0'+rr2.toString(16)).slice(-2)+('0'+gg2.toString(16)).slice(-2)+('0'+bb2.toString(16)).slice(-2),line:'15803D'};
    }

    var cx=x;
    headers.forEach(function(h,i){
      slide.addShape('rect',{x:cx,y:y,w:w[i],h:headerH,fill:{color:'075985'},line:{color:'FFFFFF',pt:1}});
      slide.addText(h,{x:cx+0.03,y:y+0.09,w:w[i]-0.06,h:headerH-0.14,fontSize:13.5,bold:true,color:'FFFFFF',align:'center',valign:'mid',margin:0,fit:'shrink'});
      cx+=w[i];
    });

    sectorRows.forEach(function(r,ri){
      var cy=y+headerH+ri*dataH, vals=[r[0],r[1],number(r[3]),number(r[4]),number(r[5])];
      cx=x;
      vals.forEach(function(v,ci){
        slide.addShape('rect',{x:cx,y:cy,w:w[ci],h:dataH,fill:{color:ri%2?'F8FBFF':'FFFFFF'},line:{color:'CBD5E1',pt:0.8}});
        slide.addText(String(v),{x:cx+0.04,y:cy+0.09,w:w[ci]-0.08,h:dataH-0.15,fontSize:13.5,bold:ci===1,color:'172033',align:ci===1?'left':'center',valign:'mid',margin:0,fit:'shrink'});
        cx+=w[ci];
      });
      var sc=scaleColor(r[6]);
      slide.addShape('roundRect',{x:cx+0.10,y:cy+dataH*0.13,w:w[5]-0.20,h:dataH*0.74,fill:{color:sc.bg},line:{color:sc.line,pt:0.9}});
      slide.addText(number(r[6]).toFixed(1)+'%',{x:cx+0.10,y:cy+dataH*0.31,w:w[5]-0.20,h:dataH*0.30,fontSize:13.5,bold:true,color:'111827',align:'center',valign:'mid',margin:0,fit:'shrink'});
    });

    var target=sectorRows.reduce(function(s,r){return s+number(r[4]);},0);
    var ach=sectorRows.reduce(function(s,r){return s+number(r[5]);},0);
    var facilities=sectorRows.reduce(function(s,r){return s+number(r[3]);},0);
    var totalPct=target?ach/target*100:0;
    var ty=y+headerH+(sectorRows.length*dataH), tc=x;
    var tv=['TOTAL','',facilities,target,ach,totalPct];
    tv.forEach(function(v,ci){
      var ww=w[ci];
      slide.addShape('rect',{x:tc,y:ty,w:ww,h:totalH,fill:{color:'0F766E'},line:{color:'FFFFFF',pt:1}});
      if(ci===5){
        slide.addShape('roundRect',{x:tc+0.10,y:ty+0.08,w:ww-0.20,h:totalH-0.16,fill:{color:'16A34A'},line:{color:'FFFFFF',pt:0.8}});
        slide.addText(Number(v).toFixed(1)+'%',{x:tc+0.10,y:ty+0.17,w:ww-0.20,h:totalH-0.24,fontSize:13.5,bold:true,color:'FFFFFF',align:'center',valign:'mid',margin:0,fit:'shrink'});
      }else{
        slide.addText(String(v),{x:tc+0.04,y:ty+0.11,w:ww-0.08,h:totalH-0.18,fontSize:13.5,bold:true,color:'FFFFFF',align:ci===0?'left':'center',valign:'mid',margin:0,fit:'shrink'});
      }
      tc+=ww;
    });

    slide.addText('Percentage color scale: lower performance → red, middle → yellow, higher performance → green.',{
      x:0.35,y:6.86,w:12.55,h:0.18,fontSize:8.5,bold:true,color:'475569',margin:0,fit:'shrink',align:'center'
    });
    slide.addText('NIN column removed • Source: live Google Sheet • '+title,{
      x:0.45,y:7.10,w:12.0,h:0.14,fontSize:7.2,color:'64748B',margin:0,fit:'shrink',align:'center'
    });
    return slide;
  }
  function jasAddAllFacilitySlide(pptx,title,rows){
    var slide=pptx.addSlide();
    addHeader(slide,'🤝 JAS Meeting — Analysis Report','BLOCK SUMMARY + FACILITY MEETING LIST | FY 2026–27');
    addSectionTitle(slide,'JAS MEETING — BLOCK SUMMARY',0.34,0.92,5.5);

    var data=(rows||[]).filter(function(r){
      return clean(r[3]) && (jasLooksNumber(r[4]) || jasLooksNumber(r[5]));
    }).map(function(r){
      var target=jasLooksNumber(r[4])?number(r[4]):0;
      var held=jasLooksNumber(r[5])?number(r[5]):0;
      return {facility:clean(r[3]),target:target,held:held,gap:Math.max(target-held,0)};
    });

    var groups={};
    data.forEach(function(r){
      var k=String(Math.round(r.held));
      if(!groups[k]) groups[k]=[];
      groups[k].push(r);
    });

    var facilityCount=data.length;
    var totalTarget=data.reduce(function(s,r){return s+r.target;},0);
    var totalHeld=data.reduce(function(s,r){return s+r.held;},0);
    var totalGap=Math.max(totalTarget-totalHeld,0);
    var overall=totalTarget?totalHeld/totalTarget*100:0;

    addCard(slide,0.34,1.22,2.32,0.88,'TOTAL FACILITIES',facilityCount);
    addCard(slide,2.80,1.22,2.32,0.88,'TARGET MEETINGS',totalTarget);
    addCard(slide,5.26,1.22,2.32,0.88,'MEETINGS HELD',totalHeld);
    addCard(slide,7.72,1.22,2.32,0.88,'GAP',totalGap);
    addCard(slide,10.18,1.22,2.80,0.88,'OVERALL %',overall.toFixed(1)+'%');

    addSectionTitle(slide,'FACILITY LIST BY MEETINGS HELD',0.34,2.30,6.0);

    function titleColor(k){ return k===0?'DC2626':(k<=4?'CA8A04':'16A34A'); }
    function softColor(k){ return k===0?'FEE2E2':(k<=4?'FEF3C7':'DCFCE7'); }

    function drawSmallGroup(k,x,y,w,h){
      var list=(groups[String(k)]||[]).slice().sort(function(a,b){return a.facility.localeCompare(b.facility);});
      slide.addShape('roundRect',{x:x,y:y,w:w,h:h,fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:0.8}});
      slide.addShape('rect',{x:x,y:y,w:1.14,h:h,fill:{color:softColor(k)},line:{color:softColor(k)}});
      slide.addText(String(k)+' meetings',{x:x+0.05,y:y+0.10,w:1.04,h:0.22,fontSize:12.5,bold:true,color:titleColor(k),align:'center',margin:0,fit:'shrink'});
      slide.addText(String(list.length)+' facilities',{x:x+0.05,y:y+0.39,w:1.04,h:0.18,fontSize:9.5,bold:true,color:'475569',align:'center',margin:0,fit:'shrink'});

      var usableX=x+1.25, usableW=w-1.36;
      var cols=list.length>4?2:1;
      var perCol=Math.ceil(Math.max(list.length,1)/cols);
      var gapX=0.10;
      var colW=(usableW-gapX*(cols-1))/cols;
      for(var c=0;c<cols;c++){
        var part=list.slice(c*perCol,(c+1)*perCol);
        if(!part.length) continue;
        slide.addText(part.map(function(r){return r.facility;}).join('\n'),{
          x:usableX+c*(colW+gapX),y:y+0.08,w:colW,h:h-0.16,
          fontSize:list.length>=8?8.6:9.4,
          bold:true,color:'172033',align:'left',valign:'mid',
          margin:0.01,fit:'shrink',breakLine:false
        });
      }
    }

    function drawFiveGroup(x,y,w,h){
      var list=(groups['5']||[]).slice().sort(function(a,b){return a.facility.localeCompare(b.facility);});
      slide.addShape('roundRect',{x:x,y:y,w:w,h:h,fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:0.9}});
      slide.addShape('rect',{x:x,y:y,w:1.42,h:h,fill:{color:'DCFCE7'},line:{color:'DCFCE7'}});
      slide.addText('5 meetings',{x:x+0.10,y:y+0.20,w:1.22,h:0.26,fontSize:14.5,bold:true,color:'16A34A',align:'center',margin:0,fit:'shrink'});
      slide.addText(String(list.length)+' facilities',{x:x+0.10,y:y+0.58,w:1.22,h:0.20,fontSize:10.5,bold:true,color:'166534',align:'center',margin:0,fit:'shrink'});
      var mid=Math.ceil(list.length/2);
      var left=list.slice(0,mid), right=list.slice(mid);
      slide.addText(left.map(function(r){return r.facility;}).join('\n'),{
        x:x+1.60,y:y+0.10,w:(w-1.86)/2,h:h-0.18,fontSize:10.5,bold:true,color:'172033',align:'left',valign:'mid',margin:0.01,fit:'shrink',breakLine:false
      });
      slide.addText(right.map(function(r){return r.facility;}).join('\n'),{
        x:x+1.70+(w-1.86)/2,y:y+0.10,w:(w-2.00)/2,h:h-0.18,fontSize:10.5,bold:true,color:'172033',align:'left',valign:'mid',margin:0.01,fit:'shrink',breakLine:false
      });
    }

    var smallY=2.68, boxH=0.80, gap=0.12;
    drawSmallGroup(0,0.34,smallY,6.08,boxH);
    drawSmallGroup(1,0.34,smallY+boxH+gap,6.08,boxH);
    drawSmallGroup(2,0.34,smallY+2*(boxH+gap),6.08,boxH);
    drawSmallGroup(3,6.72,smallY,6.08,boxH);
    drawSmallGroup(4,6.72,smallY+boxH+gap,6.08,boxH);
    drawFiveGroup(6.72,smallY+2*(boxH+gap),6.08,1.84);

    slide.addShape('roundRect',{x:0.34,y:smallY+2*(boxH+gap),w:6.08,h:1.84,fill:{color:'F8FBFF'},line:{color:'CBD5E1',pt:0.8}});
    slide.addText('MEETING GAP SUMMARY',{x:0.56,y:smallY+2*(boxH+gap)+0.18,w:5.65,h:0.22,fontSize:13,bold:true,color:'075985',margin:0,fit:'shrink'});
    slide.addText(
      '0 meetings: '+(groups['0']||[]).length+' facilities\n'+
      '1 meeting: '+(groups['1']||[]).length+' facilities\n'+
      '2 meetings: '+(groups['2']||[]).length+' facilities\n'+
      '3 meetings: '+(groups['3']||[]).length+' facilities\n'+
      '4 meetings: '+(groups['4']||[]).length+' facilities\n'+
      '5 meetings: '+(groups['5']||[]).length+' facilities',
      {x:0.56,y:smallY+2*(boxH+gap)+0.52,w:5.55,h:1.15,fontSize:11.5,bold:true,color:'172033',margin:0.02,fit:'shrink',breakLine:false}
    );

    slide.addText('NIN number excluded • Facility names only • Meeting Held = Achievement column',{
      x:0.42,y:7.05,w:12.0,h:0.15,fontSize:7.2,color:'64748B',margin:0,fit:'shrink',align:'center'
    });
    slide.addText('Source: live Google Sheet • '+title,{
      x:0.42,y:7.22,w:12.0,h:0.13,fontSize:6.8,color:'64748B',margin:0,fit:'shrink',align:'center'
    });
    return slide;
  }
  async function addJASPresentation(pptx){
    var live=null;

    // First use the exact data already parsed by the dashboard's JAS module.
    if(typeof window.getJASPresentationData==='function'){
      live=await window.getJASPresentationData();
    }
    if((!live || (!(live.sector&&live.sector.length) && !(live.facility&&live.facility.length))) &&
       typeof window.__loadJASForPPTX==='function'){
      live=await window.__loadJASForPPTX();
    }
    if(!live && window.__jasPptxData){
      live=window.__jasPptxData;
    }

    var title=live && live.title ? live.title : '';
    var parsed={
      sector:(live && live.sector)||[],
      facility:(live && live.facility)||[]
    };

    // Last-resort fallback for environments where the live JAS bridge is unavailable.
    if(!parsed.sector.length && !parsed.facility.length &&
       typeof window.__openJASForPPTX==='function'){
      live=await window.__openJASForPPTX();
      title=live && live.title ? live.title : title;
      parsed={
        sector:(live && live.sector)||[],
        facility:(live && live.facility)||[]
      };
    }

    // Final fallback: read the exact rendered JAS tables from the dashboard DOM.
    if(!parsed.sector.length && !parsed.facility.length){
      var jasModule=document.getElementById('jasMeetingModule');
      if(jasModule){
        var tables=jasModule.querySelectorAll('.jas-table');
        for(var ti=0;ti<tables.length;ti++){
          var trs=tables[ti].querySelectorAll('tbody tr');
          for(var tr=0;tr<trs.length;tr++){
            var cells=trs[tr].querySelectorAll('td');
            if(cells.length<7) continue;
            var vals=[];
            for(var ci=0;ci<cells.length;ci++) vals.push(clean(cells[ci].innerText));
            if(!/^\\d+(?:\\.0+)?$/.test(vals[0])) continue;
            if(tables[ti].innerText.indexOf('Name of Facility')>=0){
              parsed.facility.push(vals.slice(0,7));
            }else{
              parsed.sector.push(vals.slice(0,7));
            }
          }
        }
      }
    }

    if(!parsed.sector.length && !parsed.facility.length){
      var rs=await queryJASSheet();
      title=jasTitle(rs);
      parsed=getJASRows(rs);
    }

    if(!parsed.sector.length && !parsed.facility.length) throw new Error('JAS Meeting में कोई data नहीं मिला।');

    var start=pptx.slides.length+1;

    // JAS PPTX requested layout:
    // 1) Analysis Report / Block Summary + Facility Meeting List — first JAS slide
    // 2) Sector Wise Data
    jasAddAllFacilitySlide(pptx,title,parsed.facility);
    jasAddSectorSlide(pptx,title,parsed.sector);

    return {start:start,end:pptx.slides.length};
  }


  // ---------------- WELLNESS ACTIVITY PPTX ----------------
  function queryWellnessActivitySheet() {
    return new Promise(function(resolve, reject) {
      var gid = '447031017';
      var q = new google.visualization.Query(
        'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(SHEET_ID) +
        '/gviz/tq?gid=' + gid + '&headers=0'
      );
      q.setQuery('select A,B,C,D,E,F,G');
      q.send(function(response) {
        try {
          if (response.isError()) throw new Error('Wellness Activity data load नहीं हुआ: ' + response.getMessage());
          var dt=response.getDataTable(), rs=[];
          for(var r=0;r<dt.getNumberOfRows();r++){
            var row=[];
            for(var c=0;c<7;c++) row.push(clean(dt.getFormattedValue(r,c) || dt.getValue(r,c) || ''));
            rs.push(row);
          }
          resolve(rs);
        } catch(e) { reject(e); }
      });
    });
  }

  function wellnessTitle(rs){
    for(var i=0;i<rs.length;i++){
      var t=clean((rs[i]||[]).join(' '));
      if(/wellness|activity|ayushmans*arogya|fys*2026-27/i.test(t) && t.length>8) return t;
    }
    return 'Ayushman Arogya Mandir Wellness Activity FY 2026-27';
  }

  function wellnessNum(v){
    return number(String(v==null?'':v).replace(/%/g,''));
  }

  function wellnessPct(r){
    var p=wellnessNum(r[6]);
    if(!p && wellnessNum(r[4])>0) p=wellnessNum(r[5])/wellnessNum(r[4])*100;
    return p;
  }

  function wellnessTotal(r){
    return /^(total|योग|कुल)$/i.test(clean(r[0])) ||
           /^total$/i.test(clean(r[2]));
  }

  function parseWellnessActivityRows(rs){
    var sector=[], facility=[];

    function isNum(v){
      var t=clean(v).replace(/,/g,'').replace(/%/g,'').trim();
      return t!=='' && isFinite(Number(t));
    }

    for(var i=0;i<rs.length;i++){
      var r=rs[i]||[];
      while(r.length<7) r.push('');
      var sn=clean(r[0]), sec=clean(r[2]), fac=clean(r[3]);

      if(!isNum(sn) || !sec || wellnessTotal(r)) continue;
      if(!isNum(r[4]) && !isNum(r[5])) continue;

      // Numeric column D = No. of Facility -> Sector row.
      if(isNum(r[3])) sector.push(r.slice(0,7));
      else if(fac) facility.push(r.slice(0,7));
    }

    function uniq(rows,keyIndex){
      var seen={},out=[];
      rows.forEach(function(r){
        var k=clean(r[keyIndex]||'').toLowerCase();
        if(!k || seen[k]) return;
        seen[k]=true; out.push(r);
      });
      return out;
    }

    sector=uniq(sector,2);
    facility=uniq(facility,3);

    if(!sector.length && facility.length){
      var map={};
      facility.forEach(function(r){
        var s=clean(r[2])||'Other';
        if(!map[s]) map[s]={sector:s,facilities:0,target:0,total:0};
        map[s].facilities++;
        map[s].target+=wellnessNum(r[4]);
        map[s].total+=wellnessNum(r[5]);
      });
      Object.keys(map).forEach(function(k){
        var x=map[k];
        sector.push([
          '',x.sector,String(x.facilities),String(x.facilities),
          String(x.target),String(x.total),
          String(x.target?x.total/x.target*100:0)
        ]);
      });
    }

    return {sector:sector,facility:facility};
  }

  function wellnessScale(v,min,max){
    var n=wellnessNum(v);
    if(max<=min) return {fill:'FEF3C7',line:'CA8A04',text:'172033'};
    var t=Math.max(0,Math.min(1,(n-min)/(max-min)));
    if(t<=0.5){
      var q=t*2;
      var r=Math.round(220+(255-220)*q).toString(16).padStart(2,'0');
      var g=Math.round(38+(193-38)*q).toString(16).padStart(2,'0');
      var b=Math.round(38+(7-38)*q).toString(16).padStart(2,'0');
      return {fill:r+g+b,line:'B91C1C',text:'111827'};
    }
    var q2=(t-0.5)*2;
    var r2=Math.round(255+(22-255)*q2).toString(16).padStart(2,'0');
    var g2=Math.round(193+(163-193)*q2).toString(16).padStart(2,'0');
    var b2=Math.round(7+(74-7)*q2).toString(16).padStart(2,'0');
    return {fill:r2+g2+b2,line:'15803D',text:'111827'};
  }

  function wellnessMetricCell(slide,x,y,w,h,value,min,max){
    var sc=wellnessScale(value,min,max);
    slide.addShape('roundRect',{
      x:x+0.04,y:y+0.03,w:w-0.08,h:h-0.06,
      fill:{color:sc.fill},line:{color:sc.line,pt:0.8}
    });
    slide.addText(String(value),{
      x:x+0.04,y:y+h*0.24,w:w-0.08,h:h*0.44,
      fontSize:Math.max(8.5,Math.min(12,h*44)),
      bold:true,color:sc.text,align:'center',valign:'mid',margin:0,fit:'shrink'
    });
  }

  function wellnessAddSectorSlide(pptx,title,sectorRows){
    var slide=pptx.addSlide();
    addHeader(slide,'🩺 Wellness Activity — Sector Wise','SECTOR WISE DATA | FY 2026–27');
    addSectionTitle(slide,'SECTOR WISE DATA',0.28,0.94,5.8);

    var rows=sectorRows.slice();
    var x=0.22,y=1.34,w=[2.70,1.75,2.05,2.05,2.45];
    var heads=['Sector','No. of Facility','Target / Month','Total Activity','%'];
    var hh=0.62, tableH=5.56, bodyRows=rows.length+1, rh=Math.min(0.80,Math.max(0.45,(tableH-hh)/Math.max(1,bodyRows)));
    var pv=rows.map(wellnessPct), pmin=pv.length?Math.min.apply(null,pv):0,pmax=pv.length?Math.max.apply(null,pv):0;

    var cx=x;
    heads.forEach(function(h,i){
      slide.addShape('rect',{x:cx,y:y,w:w[i],h:hh,fill:{color:'075985'},line:{color:'FFFFFF',pt:1}});
      slide.addText(h,{x:cx+0.03,y:y+0.10,w:w[i]-0.06,h:0.36,fontSize:14,bold:true,color:'FFFFFF',align:'center',valign:'mid',margin:0,fit:'shrink'});
      cx+=w[i];
    });

    rows.forEach(function(r,ri){
      var cy=y+hh+ri*rh,cx2=x;
      var vals=[clean(r[2]),number(r[3]),number(r[4]),number(r[5])];
      vals.forEach(function(v,ci){
        slide.addShape('rect',{x:cx2,y:cy,w:w[ci],h:rh,fill:{color:ri%2?'F8FBFF':'FFFFFF'},line:{color:'CBD5E1',pt:0.8}});
        slide.addText(String(v),{
          x:cx2+0.04,y:cy+rh*0.24,w:w[ci]-0.08,h:rh*0.42,
          fontSize:13.5,bold:ci===0,color:'172033',
          align:ci===0?'left':'center',valign:'mid',margin:0,fit:'shrink'
        });
        cx2+=w[ci];
      });
      slide.addShape('rect',{x:cx2,y:cy,w:w[4],h:rh,fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:0.8}});
      wellnessMetricCell(slide,cx2,cy,w[4],rh,wellnessPct(r).toFixed(1)+'%',pmin,pmax);
    });

    var totalTarget=rows.reduce(function(s,r){return s+wellnessNum(r[4]);},0);
    var total=rows.reduce(function(s,r){return s+wellnessNum(r[5]);},0);
    var facilities=rows.reduce(function(s,r){return s+wellnessNum(r[3]);},0);
    var totalPct=totalTarget?total/totalTarget*100:0;
    var ty=y+hh+rows.length*rh,c=x;
    var tv=['BLOCK TOTAL',facilities,totalTarget,total,totalPct];

    tv.forEach(function(v,ci){
      slide.addShape('rect',{x:c,y:ty,w:w[ci],h:rh,fill:{color:'073B75'},line:{color:'FFFFFF',pt:1}});
      if(ci===4){
        slide.addShape('roundRect',{x:c+0.10,y:ty+rh*0.13,w:w[ci]-0.20,h:rh*0.70,fill:{color:'16A34A'},line:{color:'FFFFFF',pt:0.8}});
        slide.addText(Number(v).toFixed(1)+'%',{x:c+0.10,y:ty+rh*0.30,w:w[ci]-0.20,h:rh*0.30,fontSize:13.5,bold:true,color:'FFFFFF',align:'center',margin:0,fit:'shrink'});
      }else{
        slide.addText(String(v),{x:c+0.04,y:ty+rh*0.25,w:w[ci]-0.08,h:rh*0.42,fontSize:13.5,bold:true,color:'FFFFFF',align:ci===0?'left':'center',margin:0,fit:'shrink'});
      }
      c+=w[ci];
    });

    slide.addText('Target included in Sector Wise • NIN removed • Percentage color scale: low → red, middle → yellow, high → green • Source: live Google Sheet • '+title,{
      x:0.35,y:7.10,w:12.6,h:0.14,fontSize:7.2,color:'64748B',margin:0,align:'center',fit:'shrink'
    });
    return slide;
  }

function wellnessAddFacilitySlide(pptx,title,facilityRows){
    var slide=pptx.addSlide();
    addHeader(slide,'🩺 Wellness Activity — Facility Wise','FACILITY PERFORMANCE | HIGHEST → LOWEST');
    addSectionTitle(slide,'FACILITY WISE — HIGHEST TO LOWEST PERFORMANCE',0.25,0.94,7.5);

    var rows=facilityRows.slice().sort(function(a,b){
      return wellnessPct(b)-wellnessPct(a) || clean(a[3]).localeCompare(clean(b[3]));
    });

    var allPct=rows.map(wellnessPct);
    var pmin=allPct.length?Math.min.apply(null,allPct):0;
    var pmax=allPct.length?Math.max.apply(null,allPct):0;

    var cols=2,gap=0.16,panelW=(12.86-gap)/2,panelX=[0.20,0.20+panelW+gap];
    var y=1.34,panelH=5.72,headerH=0.45,per=Math.ceil(rows.length/2);

    function drawPanel(list,px,startRank){
      slide.addShape('roundRect',{x:px,y:y,w:panelW,h:panelH,fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:0.9}});
      var widths=[0.46,3.18,1.18,0.88];
      var heads=['#','Facility','Total','%'];
      var scale=panelW/widths.reduce(function(a,b){return a+b;},0);
      widths=widths.map(function(v){return v*scale;});
      var rh=Math.min(0.29,(panelH-headerH-0.05)/Math.max(1,per)),cx=px;

      heads.forEach(function(h,i){
        slide.addShape('rect',{x:cx,y:y,w:widths[i],h:headerH,fill:{color:'075985'},line:{color:'FFFFFF',pt:0.8}});
        slide.addText(h,{x:cx+0.015,y:y+0.07,w:widths[i]-0.03,h:0.26,fontSize:10.5,bold:true,color:'FFFFFF',align:'center',valign:'mid',margin:0,fit:'shrink'});
        cx+=widths[i];
      });

      list.forEach(function(r,ri){
        var cy=y+headerH+ri*rh,cx2=px,rank=startRank+ri;
        var vals=[rank+'.',clean(r[3]),number(r[5])];
        vals.forEach(function(v,ci){
          slide.addShape('rect',{x:cx2,y:cy,w:widths[ci],h:rh,fill:{color:ri%2?'F8FBFF':'FFFFFF'},line:{color:'CBD5E1',pt:0.55}});
          slide.addText(String(v),{x:cx2+0.018,y:cy+0.025,w:widths[ci]-0.036,h:rh-0.05,fontSize:ci===1?10.2:9.8,bold:ci===1,color:'172033',align:ci===1?'left':'center',valign:'mid',margin:0,fit:'shrink'});
          cx2+=widths[ci];
        });
        slide.addShape('rect',{x:cx2,y:cy,w:widths[3],h:rh,fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:0.55}});
        wellnessMetricCell(slide,cx2,cy,widths[3],rh,wellnessPct(r).toFixed(1)+'%',pmin,pmax);
      });
    }

    drawPanel(rows.slice(0,per),panelX[0],1);
    drawPanel(rows.slice(per),panelX[1],per+1);

    slide.addText('Facility names only • NIN, Sector, Target and Block Summary removed • Ranked by Achievement % highest → lowest • One common % color scale for all facilities',{
      x:0.35,y:7.11,w:12.6,h:0.14,fontSize:7.3,color:'64748B',margin:0,align:'center',fit:'shrink'
    });
    slide.addText('Source: live Google Sheet • '+title,{x:0.45,y:7.24,w:12.0,h:0.11,fontSize:6.6,color:'64748B',margin:0,align:'center',fit:'shrink'});
    return slide;
  }

  async function addWellnessActivityPresentation(pptx){
    var rs=await queryWellnessActivitySheet();
    var title=wellnessTitle(rs);
    var parsed=parseWellnessActivityRows(rs);
    if(!parsed.facility.length && !parsed.sector.length) throw new Error('Wellness Activity में कोई data नहीं मिला।');
    var start=pptx.slides.length+1;
    if(parsed.sector.length) wellnessAddSectorSlide(pptx,title,parsed.sector);
    wellnessAddFacilitySlide(pptx,title,parsed.facility);
    return {start:start,end:pptx.slides.length};
  }



  // ---------------- TELEMEDICINE REPORT PPTX ----------------
  var TELEMEDICINE_SHEET_ID = '1pnVA9hfSR6bfKUkyctcDWFYANQmHyWNFL6NfKue3dA4';
  var TELEMEDICINE_SECTOR_GID = '318264987';
  var TELEMEDICINE_FACILITY_GID = '1810450627';

  function queryTelemedicineSheet(gid) {
    return new Promise(function(resolve, reject) {
      var q = new google.visualization.Query(
        'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(TELEMEDICINE_SHEET_ID) +
        '/gviz/tq?gid=' + encodeURIComponent(gid) + '&headers=0'
      );
      q.setQuery('select A,B,C,D,E,F,G,H');
      q.send(function(response) {
        try {
          if (response.isError()) throw new Error(response.getMessage());
          var dt=response.getDataTable(), rs=[];
          for(var r=0;r<dt.getNumberOfRows();r++){
            var row=[];
            for(var cc=0;cc<dt.getNumberOfColumns();cc++){
              var fv=dt.getFormattedValue(r,cc);
              var rv=(fv!=='' && fv!=null) ? fv : dt.getValue(r,cc);
              row.push(clean(rv));
            }
            if(row.some(function(v){return v!=='';})) rs.push(row);
          }
          resolve(rs);
        } catch(e) { reject(e); }
      });
    });
  }

  function telePptxNum(v) {
    var n=Number(String(v==null?'':v).replace(/,/g,'').replace(/%/g,'').trim());
    return isFinite(n)?n:0;
  }

  function telePptxPct(target, completed, sourcePct) {
    var t=telePptxNum(target), c=telePptxNum(completed);
    if(t>0) return c/t*100;
    return telePptxNum(sourcePct);
  }

  function telePptxTitle(rs) {
    for(var i=0;i<rs.length;i++){
      var t=clean((rs[i]||[]).join(' '));
      if(/e[-\s]?sanjeevani/i.test(t) || /telemedicine/i.test(t)) return t;
    }
    return 'E-Sanjeevani Telemedicine Status FY 2026-27';
  }

  function telePptxDate(rs) {
    for(var i=0;i<rs.length;i++){
      var t=clean((rs[i]||[]).join(' '));
      if(/as\s*on\s*date/i.test(t)) return t;
    }
    return '';
  }

  function telePptxIsTotal(r) {
    var s=(r||[]).map(function(v){return clean(v).toLowerCase();});
    return s.some(function(v){return v==='total'||v==='योग'||v==='कुल';});
  }

  function telePptxIsPHCGorpar(r, mode) {
    if(mode!=='sector') return false;
    var e=clean((r||[])[1]).toLowerCase();
    return e==='phc gorpar' || e==='phc-gorpar' || /\bphc\s*gorpar\b/.test(e);
  }

  function telePptxHeaderIndex(rs, mode) {
    for(var i=0;i<rs.length;i++){
      var t=clean((rs[i]||[]).join(' ')).toLowerCase();
      var first=clean((rs[i]||[])[0]);
      var hasFirst=/^(sn|s\.?\s*no\.?|क्रम|क्रमांक)$/i.test(first);
      var entity=mode==='sector'?/sector/.test(t):/facility/.test(t);
      var metrics=/annual\s*target|per\s*month|completed|in\s*process/.test(t);
      if((hasFirst||/\bsn\b/.test(t)) && entity && metrics) return i;
    }
    return -1;
  }

  function telePptxRows(rs, mode) {
    var h=telePptxHeaderIndex(rs, mode);
    var data=[];
    var start=h>=0?h+1:0;

    for(var i=start;i<rs.length;i++){
      var r=rs[i]||[];
      if(!r.some(function(v){return clean(v)!=='';})) continue;
      if(telePptxIsTotal(r)) continue;
      if(telePptxIsPHCGorpar(r,mode)) continue;
      var first=clean(r[0]);
      if(/^\d+$/.test(first) && clean(r[1])) data.push(r);
    }

    // Fallback when the source header is not detected.
    if(!data.length){
      for(var j=0;j<rs.length;j++){
        var rr=rs[j]||[];
        if(telePptxIsTotal(rr) || telePptxIsPHCGorpar(rr,mode)) continue;
        if(/^\d+$/.test(clean(rr[0])) && clean(rr[1])) data.push(rr);
      }
    }

    var title=telePptxTitle(rs);
    var date=telePptxDate(rs);
    var total;

    if(mode==='sector'){
      var aam=0,target=0,totalCount=0,completed=0,inProcess=0;
      data.forEach(function(r){
        aam+=telePptxNum(r[2]);
        target+=telePptxNum(r[3]);
        totalCount+=telePptxNum(r[4]);
        completed+=telePptxNum(r[5]);
        inProcess+=telePptxNum(r[6]);
      });
      total={aam:aam,target:target,total:totalCount,completed:completed,inProcess:inProcess,pct:target?completed/target*100:0};
    }else{
      var ft=0, ftotal=0, fcompleted=0, fprocess=0;
      data.forEach(function(r){
        ft+=telePptxNum(r[2]);
        ftotal+=telePptxNum(r[3]);
        fcompleted+=telePptxNum(r[4]);
        fprocess+=telePptxNum(r[5]);
      });
      total={target:ft,total:ftotal,completed:fcompleted,inProcess:fprocess,pct:ft?fcompleted/ft*100:0};
    }

    return {title:title,date:date,rows:data,total:total};
  }

  function telePptxColor(pct) {
    return pct>=100
      ? {fill:'DCFCE7',line:'16A34A',text:'166534'}
      : (pct>=90
        ? {fill:'FEF3C7',line:'D97706',text:'92400E'}
        : {fill:'FEE2E2',line:'DC2626',text:'991B1B'});
  }

  function teleAddMetricCell(slide,x,y,w,h,value,pct) {
    var sc=telePptxColor(pct);
    slide.addShape('roundRect',{
      x:x+0.08,y:y+h*0.14,w:w-0.16,h:h*0.72,
      fill:{color:sc.fill},line:{color:sc.line,pt:0.8}
    });
    slide.addText(String(value),{
      x:x+0.08,y:y+h*0.31,w:w-0.16,h:h*0.30,
      fontSize:Math.max(10,Math.min(16,h*40)),bold:true,
      color:sc.text,align:'center',valign:'mid',margin:0,fit:'shrink'
    });
  }

  function teleAddSectorSlide(pptx,title,date,sectorRows,total) {
    var slide=pptx.addSlide();
    addHeader(slide,'🩻 Telemedicine Report','SECTOR WISE DATA | FY 2026–27');
    addSectionTitle(slide,'SECTOR WISE DATA',0.28,0.94,5.2);

    if(date) slide.addText(date,{
      x:7.2,y:0.97,w:5.75,h:0.24,fontSize:10,bold:true,color:'475569',align:'right',margin:0
    });

    var rows=sectorRows.slice();
    var x=0.22,y=1.34,fullW=12.88;
    var widths=[1.95,1.45,2.25,1.55,1.55,1.65,2.48];
    var heads=['Sector','No. of AAM Facility','Annual Target / Facility','Total','Completed','In Process','%'];
    var hh=0.55;
    var tableRows=rows.length+1;
    var rh=Math.min(0.62,(5.56-hh)/Math.max(1,tableRows));
    var cx=x;

    heads.forEach(function(h,i){
      slide.addShape('rect',{x:cx,y:y,w:widths[i],h:hh,fill:{color:'075985'},line:{color:'FFFFFF',pt:1}});
      slide.addText(h,{x:cx+0.03,y:y+0.11,w:widths[i]-0.06,h:0.30,fontSize:11.5,bold:true,color:'FFFFFF',align:'center',valign:'mid',margin:0,fit:'shrink'});
      cx+=widths[i];
    });

    rows.forEach(function(r,ri){
      var cy=y+hh+ri*rh;
      var pct=telePptxPct(r[3],r[5],r[7]);
      var vals=[clean(r[1]),telePptxNum(r[2]),telePptxNum(r[3]),telePptxNum(r[4]),telePptxNum(r[5]),telePptxNum(r[6])];
      cx=x;
      vals.forEach(function(v,ci){
        slide.addShape('rect',{x:cx,y:cy,w:widths[ci],h:rh,fill:{color:ri%2?'F8FBFF':'FFFFFF'},line:{color:'CBD5E1',pt:0.75}});
        slide.addText(String(v),{
          x:cx+0.03,y:cy+rh*0.25,w:widths[ci]-0.06,h:rh*0.42,
          fontSize:ci===0?12.2:11.5,bold:ci===0,color:'172033',
          align:ci===0?'left':'center',valign:'mid',margin:0,fit:'shrink'
        });
        cx+=widths[ci];
      });
      slide.addShape('rect',{x:cx,y:cy,w:widths[6],h:rh,fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:0.75}});
      teleAddMetricCell(slide,cx,cy,widths[6],rh,pct.toFixed(0)+'%',pct);
    });

    var ty=y+hh+rows.length*rh;
    cx=x;
    var totalVals=['TOTAL',total.aam,total.target,total.total,total.completed,total.inProcess];
    totalVals.forEach(function(v,ci){
      slide.addShape('rect',{x:cx,y:ty,w:widths[ci],h:rh,fill:{color:'073B75'},line:{color:'FFFFFF',pt:1}});
      slide.addText(String(v),{
        x:cx+0.03,y:ty+rh*0.25,w:widths[ci]-0.06,h:rh*0.42,
        fontSize:12.2,bold:true,color:'FFFFFF',align:ci===0?'left':'center',valign:'mid',margin:0,fit:'shrink'
      });
      cx+=widths[ci];
    });
    slide.addShape('rect',{x:cx,y:ty,w:widths[6],h:rh,fill:{color:'073B75'},line:{color:'FFFFFF',pt:1}});
    teleAddMetricCell(slide,cx,ty,widths[6],rh,total.pct.toFixed(0)+'%',total.pct);

    slide.addText('PHC Gorpar removed as requested • Total calculated from displayed sectors • Source: separate Telemedicine Google Sheet'+(date?' • '+date:''),{
      x:0.30,y:6.92,w:12.72,h:0.14,fontSize:7.1,color:'64748B',align:'center',margin:0,fit:'shrink'
    });
    return slide;
  }

  function teleAddFacilitySlide(pptx,title,date,facilityRows,total) {
    var slide=pptx.addSlide();
    addHeader(slide,'🩻 Telemedicine Report','FACILITY WISE DATA | FY 2026–27');
    addSectionTitle(slide,'FACILITY WISE DATA',0.25,0.94,5.8);
    if(date) slide.addText(date,{
      x:7.2,y:0.97,w:5.8,h:0.24,fontSize:10,bold:true,color:'475569',align:'right',margin:0
    });

    var rows=facilityRows.slice();
    var panelGap=0.16,panelW=(12.86-panelGap)/2,panelX=[0.18,0.18+panelW+panelGap];
    var y=1.34,panelH=5.45,headerH=0.42;
    var per=Math.ceil(rows.length/2);
    var widths=[0.44,2.70,1.12,1.15,1.15,1.15,0.95,1.25];
    var heads=['#','Facility','Target','Total','Completed','In Process','%','Remark'];

    function panel(list,px,startRank){
      slide.addShape('roundRect',{x:px,y:y,w:panelW,h:panelH,fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:0.9}});
      var scale=panelW/widths.reduce(function(a,b){return a+b;},0);
      var w=widths.map(function(v){return v*scale;});
      var rh=Math.min(0.26,(panelH-headerH-0.05)/Math.max(1,per));
      var cx=px;

      heads.forEach(function(h,i){
        slide.addShape('rect',{x:cx,y:y,w:w[i],h:headerH,fill:{color:'075985'},line:{color:'FFFFFF',pt:0.7}});
        slide.addText(h,{x:cx+0.01,y:y+0.06,w:w[i]-0.02,h:0.24,fontSize:8.8,bold:true,color:'FFFFFF',align:'center',valign:'mid',margin:0,fit:'shrink'});
        cx+=w[i];
      });

      list.forEach(function(r,ri){
        var cy=y+headerH+ri*rh, rank=startRank+ri;
        var pct=telePptxPct(r[2],r[4],r[6]);
        var missing=!clean(r[3])&&!clean(r[4])&&!clean(r[5]);
        var vals=[rank+'.',clean(r[1]),telePptxNum(r[2]),telePptxNum(r[3]),telePptxNum(r[4]),telePptxNum(r[5])];
        cx=px;
        vals.forEach(function(v,ci){
          slide.addShape('rect',{x:cx,y:cy,w:w[ci],h:rh,fill:{color:ri%2?'F8FBFF':'FFFFFF'},line:{color:'CBD5E1',pt:0.55}});
          slide.addText(String(v),{
            x:cx+0.015,y:cy+0.02,w:w[ci]-0.03,h:rh-0.04,
            fontSize:ci===1?8.7:8.2,bold:ci===1,color:'172033',
            align:ci===1?'left':'center',valign:'mid',margin:0,fit:'shrink'
          });
          cx+=w[ci];
        });
        slide.addShape('rect',{x:cx,y:cy,w:w[6],h:rh,fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:0.55}});
        teleAddMetricCell(slide,cx,cy,w[6],rh,pct.toFixed(0)+'%',pct);
        cx+=w[6];
        slide.addShape('rect',{x:cx,y:cy,w:w[7],h:rh,fill:{color:missing?'FFF7ED':'FFFFFF'},line:{color:'CBD5E1',pt:0.55}});
        slide.addText(missing?'Not Mapped Portal':'',{
          x:cx+0.02,y:cy+0.02,w:w[7]-0.04,h:rh-0.04,
          fontSize:7.1,bold:true,color:'B45309',align:'center',valign:'mid',margin:0,fit:'shrink'
        });
      });
    }

    var left=rows.slice(0,per), right=rows.slice(per);
    panel(left,panelX[0],1);
    panel(right,panelX[1],per+1);

    // Total strip at bottom
    var ty=6.86;
    slide.addShape('roundRect',{x:0.18,y:ty,w:12.98,h:0.28,fill:{color:'073B75'},line:{color:'073B75'}});
    slide.addText(
      'TOTAL  •  Facilities: '+rows.length+
      '  •  Target: '+total.target+
      '  •  Total: '+total.total+
      '  •  Completed: '+total.completed+
      '  •  In Process: '+total.inProcess+
      '  •  '+total.pct.toFixed(0)+'%',
      {x:0.30,y:6.925,w:12.72,h:0.12,fontSize:8.2,bold:true,color:'FFFFFF',align:'center',margin:0,fit:'shrink'}
    );
    slide.addText('Missing facility data = Not Mapped Portal • Source: separate Telemedicine Google Sheet'+(date?' • '+date:''),{
      x:0.30,y:7.27,w:12.72,h:0.10,fontSize:6.6,color:'64748B',align:'center',margin:0,fit:'shrink'
    });
    return slide;
  }

  async function addTelemedicinePresentation(pptx) {
    var sectorRs=await queryTelemedicineSheet(TELEMEDICINE_SECTOR_GID);
    var facilityRs=await queryTelemedicineSheet(TELEMEDICINE_FACILITY_GID);
    var sector=telePptxRows(sectorRs,'sector');
    var facility=telePptxRows(facilityRs,'facility');

    if(!sector.rows.length && !facility.rows.length){
      throw new Error('Telemedicine Report में कोई data नहीं मिला।');
    }

    var start=pptx.slides.length+1;
    if(sector.rows.length) teleAddSectorSlide(pptx,sector.title,sector.date,sector.rows,sector.total);
    if(facility.rows.length) teleAddFacilitySlide(pptx,facility.title,facility.date,facility.rows,facility.total);
    return {start:start,end:pptx.slides.length};
  }

  // ---------------- AYUSHMAN SHIVIR PPTX ----------------
  function queryAyushmanShivirSheet() {
    return new Promise(function(resolve, reject) {
      var gid = '1262815420';
      var q = new google.visualization.Query(
        'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(SHEET_ID) +
        '/gviz/tq?gid=' + gid + '&headers=0'
      );
      q.setQuery('select *');
      q.send(function(response) {
        try {
          if (response.isError()) throw new Error('Ayushman Shivir data load नहीं हुआ: ' + response.getMessage());
          var dt = response.getDataTable(), rs = [];
          for (var r=0; r<dt.getNumberOfRows(); r++) {
            var row = [];
            for (var col=0; col<Math.min(9,dt.getNumberOfColumns()); col++) {
              row.push(clean(dt.getFormattedValue(r,col) || dt.getValue(r,col) || ''));
            }
            while(row.length<9) row.push('');
            rs.push(row);
          }
          resolve(rs);
        } catch(e) { reject(e); }
      });
    });
  }

  function shivirPptxTitle(rs) {
    for(var i=0;i<rs.length;i++){
      var t=clean((rs[i]||[]).join(' '));
      if(/ayushman\\s*shivir/i.test(t) || /shivir\\s*reporting/i.test(t)) return t;
    }
    return 'Ayushman Shivir Reporting FY 2026-27';
  }

  function shivirPptxNum(v) {
    return number(String(v==null?'':v).replace(/%/g,''));
  }

  function shivirPptxPct(r) {
    var p=shivirPptxNum(r[6]);
    if(!p && shivirPptxNum(r[4])>0) p=shivirPptxNum(r[5])/shivirPptxNum(r[4])*100;
    return p;
  }

  function shivirIsTotal(r) {
    return /^(total|योग|कुल)$/i.test(clean(r[0])) || /^total$/i.test(clean(r[2]));
  }

  function parseAyushmanShivirRows(rs) {
    var sector=[], facility=[];

    function isNum(v){
      var t=clean(v).replace(/,/g,'').replace(/%/g,'').trim();
      return t !== '' && isFinite(Number(t));
    }

    for(var i=0;i<rs.length;i++){
      var r=rs[i]||[];
      while(r.length<9) r.push('');

      var sn=clean(r[0]);
      var sec=clean(r[2]);
      var fac=clean(r[3]);

      if(!isNum(sn)) continue;
      if(!sec) continue;
      if(shivirIsTotal(r)) continue;
      if(!isNum(r[4]) && !isNum(r[5])) continue;

      // IMPORTANT: Sector rows contain a numeric No. of Facility in column 4.
      // Those rows must never appear as a facility named "7", "4", "9", etc.
      if(isNum(r[3])){
        sector.push(r.slice(0,9));
      }else if(fac){
        facility.push(r.slice(0,9));
      }
    }

    function uniq(rows,keyIndex){
      var seen={}, out=[];
      rows.forEach(function(r){
        var k=clean(r[keyIndex]||'').toLowerCase();
        if(!k || seen[k]) return;
        seen[k]=true;
        out.push(r);
      });
      return out;
    }

    sector=uniq(sector,2);
    facility=uniq(facility,3);

    if(!sector.length && facility.length){
      var map={};
      facility.forEach(function(r){
        var s=clean(r[2])||'Other';
        if(!map[s]) map[s]={sector:s,facilities:0,target:0,report:0,foot:0};
        map[s].facilities++;
        map[s].target+=shivirPptxNum(r[4]);
        map[s].report+=shivirPptxNum(r[5]);
        map[s].foot+=shivirPptxNum(r[7]);
      });
      Object.keys(map).forEach(function(k){
        var x=map[k];
        sector.push([
          '',x.sector,String(x.facilities),
          String(x.facilities),String(x.target),String(x.report),
          String(x.target?x.report/x.target*100:0),
          String(x.foot),String(x.report?x.foot/x.report:0)
        ]);
      });
    }

    return {sector:sector,facility:facility};
  }
  function shivirPptxScale(v,min,max){
    if(max<=min) return {fill:'FEF3C7',line:'CA8A04',text:'172033'};
    var t=Math.max(0,Math.min(1,(v-min)/(max-min)));
    if(t>=0.67) return {fill:'DCFCE7',line:'16A34A',text:'166534'};
    if(t>=0.34) return {fill:'FEF3C7',line:'CA8A04',text:'92400E'};
    return {fill:'FEE2E2',line:'DC2626',text:'991B1B'};
  }

  function shivirAddMetricCell(slide,x,y,w,h,value,min,max){
    var sc=shivirPptxScale(shivirPptxNum(value),min,max);
    slide.addShape('roundRect',{
      x:x+0.035,y:y+0.025,w:w-0.07,h:h-0.05,
      fill:{color:sc.fill},line:{color:sc.line,pt:0.8}
    });
    slide.addText(String(value),{
      x:x+0.04,y:y+Math.max(0.025,h*0.22),w:w-0.08,h:Math.max(0.10,h*0.48),
      fontSize:Math.max(8.5,Math.min(11.5,h*46)),
      bold:true,color:sc.text,align:'center',valign:'mid',margin:0,fit:'shrink'
    });
  }
  function shivirAddSummarySlide(pptx,title,sectorRows,facilityRows){
    var slide=pptx.addSlide();
    addHeader(slide,'🏕️ Ayushman Shivir','LIVE GOOGLE SHEET | FY 2026–27');
    addSectionTitle(slide,'AYUSHMAN SHIVIR — SUMMARY DASHBOARD',0.45,1.02,6.4);

    var target=facilityRows.reduce(function(s,r){return s+shivirPptxNum(r[4]);},0);
    var report=facilityRows.reduce(function(s,r){return s+shivirPptxNum(r[5]);},0);
    var foot=facilityRows.reduce(function(s,r){return s+shivirPptxNum(r[7]);},0);
    var pct=target?report/target*100:0;
    var avg=report?foot/report:0;

    addCard(slide,0.45,1.45,2.30,1.10,'FACILITIES',facilityRows.length);
    addCard(slide,2.90,1.45,2.30,1.10,'TARGET SHIVIR',target);
    addCard(slide,5.35,1.45,2.30,1.10,'SHIVIR REPORTING',report);
    addCard(slide,7.80,1.45,2.30,1.10,'ACHIEVEMENT %',pct.toFixed(1)+'%');
    addCard(slide,10.25,1.45,2.50,1.10,'TOTAL FOOTFALL',foot);

    addSectionTitle(slide,'SECTOR PERFORMANCE',0.45,2.95,4.0);
    var sorted=sectorRows.slice().sort(function(a,b){return shivirPptxPct(b)-shivirPptxPct(a);});
    var chart=sorted.map(function(r){return {name:clean(r[1]||r[2]),value:shivirPptxPct(r)};});
    if(chart.length){
      try{
        slide.addChart(pptx.ChartType.bar,[{
          name:'Achievement %',
          labels:chart.map(function(x){return x.name;}),
          values:chart.map(function(x){return x.value;})
        }],{
          x:0.42,y:3.35,w:7.20,h:3.15,
          showLegend:false,showTitle:false,showValue:true,
          catAxisLabelFontSize:12,valAxisLabelFontSize:10,
          chartColors:['0F766E'],
          valGridLine:{color:'D6E3EC',pt:1},
          valAxisMinVal:0,valAxisMaxVal:120,
          dataLabelPosition:'outEnd'
        });
      }catch(e){}
    }

    addSectionTitle(slide,'KEY ANALYSIS',8.00,2.95,4.2);
    var best=sorted.length?sorted[0]:null, low=sorted.length?sorted[sorted.length-1]:null;
    var obs=[
      '• Block achievement: '+pct.toFixed(1)+'%.',
      '• Total target: '+target+' shivir; reporting: '+report+'.',
      '• Total footfall: '+foot+'; average per reported shivir: '+avg.toFixed(1)+'.',
      best ? '• Highest sector achievement: '+clean(best[1]||best[2])+' — '+shivirPptxPct(best).toFixed(1)+'%.' : '',
      low ? '• Lowest sector achievement: '+clean(low[1]||low[2])+' — '+shivirPptxPct(low).toFixed(1)+'%.' : ''
    ].filter(function(x){return x;});
    slide.addText(obs.join('\n'),{
      x:8.05,y:3.40,w:4.55,h:2.70,
      fontSize:14,bold:true,color:'172033',
      margin:0.03,fit:'shrink'
    });
    slide.addText('Source: live Google Sheet • '+title,{
      x:0.50,y:6.82,w:12.0,h:0.16,fontSize:7.5,color:'64748B',margin:0
    });
    return slide;
  }

  function shivirAddSectorSlide(pptx,title,sectorRows){
    var slide=pptx.addSlide();
    addHeader(slide,'🏕️ Ayushman Shivir — Sector Wise','SECTOR WISE DATA | FY 2026–27');
    addSectionTitle(slide,'SECTOR WISE DATA',0.28,0.94,5.5);

    var rows=sectorRows.slice();
    var x=0.22,y=1.34;
    var w=[2.55,1.75,1.80,1.20,2.45,2.95];
    var heads=['Sector','No. of Facility','Shivir Reporting','%','Total Footfall','Avg Footfall/Shivir'];
    var hh=0.62;
    var availableH=5.35;
    var rh=Math.min(0.62,availableH/Math.max(1,rows.length+1));
    var pctVals=rows.map(shivirPptxPct), pmin=pctVals.length?Math.min.apply(null,pctVals):0,pmax=pctVals.length?Math.max.apply(null,pctVals):0;
    var avgVals=rows.map(function(r){return shivirPptxNum(r[8]);}), amin=avgVals.length?Math.min.apply(null,avgVals):0,amax=avgVals.length?Math.max.apply(null,avgVals):0;

    var cx=x;
    heads.forEach(function(h,i){
      slide.addShape('rect',{x:cx,y:y,w:w[i],h:hh,fill:{color:'075985'},line:{color:'FFFFFF',pt:1}});
      slide.addText(h,{x:cx+0.03,y:y+0.10,w:w[i]-0.06,h:0.36,fontSize:14,bold:true,color:'FFFFFF',align:'center',valign:'mid',margin:0,fit:'shrink'});
      cx+=w[i];
    });

    rows.forEach(function(r,ri){
      var cy=y+hh+ri*rh,cx2=x;
      var vals=[clean(r[1]||r[2]),number(r[3]),number(r[5]),shivirPptxPct(r).toFixed(1)+'%',number(r[7]),number(r[8]).toFixed(1)];
      vals.forEach(function(v,ci){
        slide.addShape('rect',{x:cx2,y:cy,w:w[ci],h:rh,fill:{color:ri%2?'F8FBFF':'FFFFFF'},line:{color:'CBD5E1',pt:0.8}});
        if(ci===3){
          shivirAddMetricCell(slide,cx2,cy,w[ci],rh,shivirPptxPct(r).toFixed(1)+'%',pmin,pmax);
        } else if(ci===5){
          shivirAddMetricCell(slide,cx2,cy,w[ci],rh,number(r[8]).toFixed(1),amin,amax);
        } else {
          slide.addText(String(v),{x:cx2+0.03,y:cy+rh*0.25,w:w[ci]-0.06,h:rh*0.43,fontSize:13.5,bold:ci===0,color:'172033',align:ci===0?'left':'center',valign:'mid',margin:0,fit:'shrink'});
        }
        cx2+=w[ci];
      });
    });

    var rep=rows.reduce(function(s,r){return s+shivirPptxNum(r[5]);},0);
    var foot=rows.reduce(function(s,r){return s+shivirPptxNum(r[7]);},0);
    var target=rows.reduce(function(s,r){return s+shivirPptxNum(r[4]);},0);
    var avg=rep?foot/rep:0;
    var ty=y+hh+rows.length*rh,c=x;
    var tv=['BLOCK TOTAL','',rep,target?rep/target*100:0,foot,avg];
    tv.forEach(function(v,ci){
      slide.addShape('rect',{x:c,y:ty,w:w[ci],h:rh,fill:{color:'073B75'},line:{color:'FFFFFF',pt:1}});
      if(ci===3){
        slide.addShape('roundRect',{x:c+0.10,y:ty+rh*0.13,w:w[ci]-0.20,h:rh*0.70,fill:{color:'16A34A'},line:{color:'FFFFFF',pt:0.8}});
        slide.addText(Number(v).toFixed(1)+'%',{x:c+0.10,y:ty+rh*0.29,w:w[ci]-0.20,h:rh*0.32,fontSize:13.5,bold:true,color:'FFFFFF',align:'center',margin:0,fit:'shrink'});
      }else{
        slide.addText(String(ci===5?Number(v).toFixed(1):v),{x:c+0.03,y:ty+rh*0.25,w:w[ci]-0.06,h:rh*0.42,fontSize:13.5,bold:true,color:'FFFFFF',align:ci===0?'left':'center',margin:0,fit:'shrink'});
      }
      c+=w[ci];
    });
    slide.addText('Source: live Google Sheet • '+title,{x:0.45,y:7.10,w:12.0,h:0.14,fontSize:7.2,color:'64748B',margin:0,align:'center'});
    return slide;
  }
  function shivirPackFacilityRows(rows,maxRows){
    var sorted=rows.slice().sort(function(a,b){
      var s=clean(a[2]).localeCompare(clean(b[2]));
      return s || clean(a[3]).localeCompare(clean(b[3]));
    });
    var packs=[], cur=[], curSectors=[];
    sorted.forEach(function(r){
      if(cur.length>=maxRows){
        packs.push({rows:cur,sectors:curSectors});
        cur=[];curSectors=[];
      }
      cur.push(r);
      var sec=clean(r[2]);
      if(curSectors.indexOf(sec)<0) curSectors.push(sec);
    });
    if(cur.length) packs.push({rows:cur,sectors:curSectors});
    return packs;
  }

  function shivirAddFacilitySlides(pptx,title,facilityRows){
    var slide=pptx.addSlide();
    addHeader(slide,'🏕️ Ayushman Shivir — Facility Wise','42 FACILITIES | HIGHEST → LOWEST PERFORMANCE');
    addSectionTitle(slide,'FACILITY WISE — HIGHEST TO LOWEST PERFORMANCE',0.25,0.94,7.4);

    var rows=facilityRows.slice().sort(function(a,b){
      return shivirPptxPct(b)-shivirPptxPct(a) || clean(a[3]).localeCompare(clean(b[3]));
    });

    var cols=2, gap=0.16, panelW=(12.86-gap)/2, panelX=[0.20,0.20+panelW+gap];
    var y=1.34, panelH=5.72, headerH=0.45;
    var per=Math.ceil(rows.length/cols);

    // One common percentage scale across ALL facilities (rank 1–42),
    // so both panels use identical color meaning.
    var allPct=rows.map(shivirPptxPct);
    var globalPmin=allPct.length?Math.min.apply(null,allPct):0;
    var globalPmax=allPct.length?Math.max.apply(null,allPct):0;

    function drawPanel(list,px,startRank){
      slide.addShape('roundRect',{
        x:px,y:y,w:panelW,h:panelH,
        fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:0.9}
      });

      // Wider Facility + KPI columns so the slide remains readable with 21 rows per panel.
      var widths=[0.46,2.92,1.02,0.78,1.36,1.22];
      var heads=['#','Facility','Shivir','%','Footfall','Avg/Shivir'];
      var scale=panelW/widths.reduce(function(a,b){return a+b;},0);
      widths=widths.map(function(v){return v*scale;});

      var rh=Math.min(0.29,(panelH-headerH-0.05)/Math.max(1,per));
      var cx=px;

      heads.forEach(function(h,i){
        slide.addShape('rect',{
          x:cx,y:y,w:widths[i],h:headerH,
          fill:{color:'075985'},line:{color:'FFFFFF',pt:0.8}
        });
        slide.addText(h,{
          x:cx+0.015,y:y+0.07,w:widths[i]-0.03,h:0.26,
          fontSize:10.5,bold:true,color:'FFFFFF',
          align:'center',valign:'mid',margin:0,fit:'shrink'
        });
        cx+=widths[i];
      });

      var av=list.map(function(r){return shivirPptxNum(r[8]);});
      var amin=av.length?Math.min.apply(null,av):0;
      var amax=av.length?Math.max.apply(null,av):0;

      list.forEach(function(r,ri){
        var cy=y+headerH+ri*rh, cx2=px, rank=startRank+ri;
        var vals=[rank+'.',clean(r[3]),number(r[5]),shivirPptxPct(r),number(r[7]),number(r[8])];

        vals.forEach(function(v,ci){
          slide.addShape('rect',{
            x:cx2,y:cy,w:widths[ci],h:rh,
            fill:{color:ri%2?'F8FBFF':'FFFFFF'},
            line:{color:'CBD5E1',pt:0.55}
          });

          if(ci===3){
            shivirAddMetricCell(slide,cx2,cy,widths[ci],rh,
              shivirPptxPct(r).toFixed(1)+'%',globalPmin,globalPmax);
          }else if(ci===5){
            shivirAddMetricCell(slide,cx2,cy,widths[ci],rh,
              number(r[8]).toFixed(1),amin,amax);
          }else{
            slide.addText(String(v),{
              x:cx2+0.018,y:cy+0.025,w:widths[ci]-0.036,h:rh-0.05,
              fontSize:ci===1?10.2:9.8,
              bold:ci===1,color:'172033',
              align:ci===1?'left':'center',
              valign:'mid',margin:0,fit:'shrink'
            });
          }
          cx2+=widths[ci];
        });
      });
    }

    // For the expected 42 facilities this is exactly 21 + 21.
    drawPanel(rows.slice(0,per),panelX[0],1);
    drawPanel(rows.slice(per,per*2),panelX[1],per+1);

    slide.addText(
      'Facility names only • Sector and Target removed • Ranked 1–42 by Achievement % (Highest → Lowest) • One common % color scale for all 42 facilities',
      {
        x:0.35,y:7.11,w:12.6,h:0.14,
        fontSize:7.3,color:'64748B',margin:0,align:'center',fit:'shrink'
      }
    );
    slide.addText('Source: live Google Sheet • '+title,{
      x:0.45,y:7.24,w:12.0,h:0.11,
      fontSize:6.6,color:'64748B',margin:0,align:'center',fit:'shrink'
    });
    return 1;
  }  async function addAyushmanShivirPresentation(pptx){
    var rs=await queryAyushmanShivirSheet();
    var title=shivirPptxTitle(rs);
    var parsed=parseAyushmanShivirRows(rs);
    if(!parsed.facility.length && !parsed.sector.length) throw new Error('Ayushman Shivir में कोई data नहीं मिला।');

    var start=pptx.slides.length+1;
    // Requested layout: Sector Wise + one Facility Wise slide only.
    if(parsed.sector.length) shivirAddSectorSlide(pptx,title,parsed.sector);
    shivirAddFacilitySlides(pptx,title,parsed.facility);
    return {start:start,end:pptx.slides.length};
  }


  // ---------------- RBSK PPTX ----------------
  function addRBSKPresentation(pptx) {
    var main = [
      ['TEAM','TOTAL SCHOOL','SCHOOL VISIT','%','SCHOOL SCREENING TARGET','SCHOOL SCREENED','%','TOTAL AWC','AWC VISIT','%','AWC SCREENING TARGET','AWC SCREENED','%','SICK CHILDREN','TREATED','REFERRED','REFERRAL TREATED'],
      ['TEAM A',161,79,'49%',11381,4566,'40%',212,212,'100%',5322,5267,'99%',5322,5267,1000,895],
      ['TEAM B',160,74,'46%',10036,3840,'38%',210,210,'100%',5526,5405,'98%',5526,5405,1419,1322],
      ['TOTAL',321,153,'48%',21417,8406,'39%',422,422,'100%',10848,10672,'98%',10848,10672,2419,2217]
    ];
    function addRHeader(slide,t,sub){
      addHeader(slide,t,sub);
      slide.background={color:'F8FBFF'};
    }
    function addRTable(slide,rows,x,y,w,h,widths,fs,pctCols){
      var cols=rows[0].length;
      var tdata=rows.map(function(r,ri){
        return r.map(function(v,ci){
          var fill=ri===0?'075985':(r[0]==='TOTAL'?'E0F2FE':(ri%2?'FFFFFF':'F8FBFF'));
          var color=ri===0?'FFFFFF':'172033';
          if(ri>0 && pctCols.indexOf(ci)>=0){
            var n=number(String(v).replace('%',''));
            fill=n>=70?'DCFCE7':(n>=40?'FEF3C7':'FEE2E2');
          }
          return {text:String(v),options:{bold:ri===0||r[0]==='TOTAL'||ci===0,fontSize:fs,color:color,fill:{color:fill},align:'center',valign:'mid',margin:2}};
        });
      });
      slide.addTable(tdata,{x:x,y:y,w:w,h:h,border:{type:'solid',color:'B8C9D8',pt:1.1},autoFit:false,colW:widths,rowH:0.50,margin:2.5,fill:'FFFFFF'});
    }

    // Slide 1: School performance
    var s=pptx.addSlide();
    addRHeader(s,'👶 RBSK — Part 1 | School Performance','STATUS AS ON 31-08-2026');
    addSectionTitle(s,'SCHOOL VISIT & SCHOOL SCREENING',0.45,1.02,5.2);
    addRTable(s,[
      ['Team','Total School','School Visit','Visit %','Screening Target','Screened','Screening %'],
      main[1].slice(0,7),main[2].slice(0,7),main[3].slice(0,7)
    ],0.45,1.40,12.45,2.55,[1.65,1.55,1.55,1.20,2.05,1.55,1.45],15,[3,6]);
    addCard(s,0.65,4.35,2.7,1.05,'TOTAL SCHOOL',321);
    addCard(s,3.55,4.35,2.7,1.05,'SCHOOL VISIT',153);
    addCard(s,6.45,4.35,2.7,1.05,'SCREENED',8406);
    addCard(s,9.35,4.35,2.7,1.05,'SCREENING %','39%');
    s.addText('Color scale: red = lower • yellow = middle • green = higher',{x:0.7,y:5.75,w:11.9,h:0.22,fontSize:9,bold:true,color:'64748B',align:'center',margin:0});

    // Slide 2: Anganwadi
    s=pptx.addSlide();
    addRHeader(s,'👶 RBSK — Part 1 | Anganwadi Performance','ANGANWADI FIRST VISIT • 31-08-2026');
    addSectionTitle(s,'ANGANWADI VISIT & SCREENING',0.45,1.02,5.2);
    addRTable(s,[
      ['Team','Total AWC','AWC Visit','Visit %','Screening Target','Screened','Screening %'],
      [main[1][0],main[1][7],main[1][8],main[1][9],main[1][10],main[1][11],main[1][12]],
      [main[2][0],main[2][7],main[2][8],main[2][9],main[2][10],main[2][11],main[2][12]],
      [main[3][0],main[3][7],main[3][8],main[3][9],main[3][10],main[3][11],main[3][12]]
    ],0.45,1.40,12.45,2.55,[1.65,1.55,1.55,1.20,2.05,1.55,1.45],15,[3,6]);
    addCard(s,0.65,4.35,2.7,1.05,'TOTAL AWC',422);
    addCard(s,3.55,4.35,2.7,1.05,'AWC VISIT',422);
    addCard(s,6.45,4.35,2.7,1.05,'SCREENED',10672);
    addCard(s,9.35,4.35,2.7,1.05,'SCREENING %','98%');

    // Slide 3: Referral & Treatment
    s=pptx.addSlide();
    addRHeader(s,'👶 RBSK — Part 1 | Referral & Treatment','STATUS AS ON 31-08-2026');
    addSectionTitle(s,'REFERRAL & TREATMENT',0.45,1.02,4.2);
    addRTable(s,[
      ['Team','Sick Children','Treated','Referred','Referral Treated'],
      [main[1][0],main[1][13],main[1][14],main[1][15],main[1][16]],
      [main[2][0],main[2][13],main[2][14],main[2][15],main[2][16]],
      [main[3][0],main[3][13],main[3][14],main[3][15],main[3][16]]
    ],0.55,1.45,12.2,2.65,[2.15,2.45,2.45,2.45,2.70],15.5,[]);
    addCard(s,0.8,4.45,2.7,1.05,'SICK CHILDREN',10848);
    addCard(s,3.7,4.45,2.7,1.05,'TREATED',10672);
    addCard(s,6.6,4.45,2.7,1.05,'REFERRED',2419);
    addCard(s,9.5,4.45,2.7,1.05,'REFERRAL TREATED',2217);

    var catData = {
      'A':[[2,1,1,'50%',1],[1,0,1,'100%',0],[3,1,2,'67%',1]],
      'B':[[16,3,13,'81%',16],[50,7,43,'86%',50],[66,10,56,'85%',66]],
      'C':[[7,6,1,'14%',7],[5,2,3,'60%',5],[12,8,4,'33%',12]],
      'D':[[4,0,0,'0%',4],[1,1,20,'20%',0],[5,1,8,'8%',0]],
      'E':[[77,1,76,'99%',1],[106,17,89,'84%',17],[183,18,165,'90%',18]],
      'All Category':[[48,2,46,'96%',48],[49,7,42,'86%',49],[97,9,88,'91%',97]]
    };
    function catSlide(title, key, idx){
      var ss=pptx.addSlide();
      addRHeader(ss,'💊 RBSK — Part 2 | '+title,'01/04/2026 TO 31/03/2027');
      addSectionTitle(ss,'CATEGORY: '+key.toUpperCase(),0.45,1.02,4.8);
      var v=catData[key];
      addRTable(ss,[
        ['Team','Identified','Under Treatment','Treatment Over','Treatment %','Pending'],
        ['Team A',v[0][0],v[0][1],v[0][2],v[0][3],v[0][4]],
        ['Team B',v[1][0],v[1][1],v[1][2],v[1][3],v[1][4]],
        ['TOTAL',v[2][0],v[2][1],v[2][2],v[2][3],v[2][4]]
      ],0.55,1.45,12.2,2.80,[2.05,1.85,2.10,2.10,1.85,2.25],14.5,[4]);
      addCard(ss,0.65,4.55,3.75,1.15,'TOTAL IDENTIFIED',v[2][0]);
      addCard(ss,4.80,4.55,3.75,1.15,'TREATMENT OVER',v[2][2]);
      addCard(ss,8.95,4.55,3.75,1.15,'TREATMENT %',v[2][3]);
    }
    ['A','B','C','D','E','All Category'].forEach(function(k){catSlide('Part 2 | Category Wise Periodical Treatment',k);});
    return true;
  }

  function queryRawSheet(gid) {
    return new Promise(function (resolve, reject) {
      var q = new google.visualization.Query(
        'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(SHEET_ID) +
        '/gviz/tq?gid=' + encodeURIComponent(gid) + '&headers=0'
      );
      q.setQuery('select *');
      q.send(function (response) {
        if (response.isError()) { reject(new Error(response.getMessage())); return; }
        var dt=response.getDataTable(), rows=[];
        for(var r=0;r<dt.getNumberOfRows();r++){
          var row=[];
          for(var c=0;c<dt.getNumberOfColumns();c++) row.push(clean(dt.getFormattedValue(r,c)));
          rows.push(row);
        }
        resolve(rows);
      });
    });
  }

  function pptxTable(slide, headers, rows, x, y, w, h, fontSize) {
    var data=[headers].concat(rows);
    var cols=headers.length, colW=[];
    for(var i=0;i<cols;i++) colW.push(w/cols);
    slide.addTable(data.map(function(r,ri){
      return r.map(function(v){return {text:clean(v),options:{
        bold:ri===0,fontSize:fontSize||9,color:ri===0?'FFFFFF':'172033',
        fill:{color:ri===0?'075985':'FFFFFF'},align:'center',valign:'mid',margin:2
      }};});
    }),{x:x,y:y,w:w,h:h,border:{type:'solid',color:'CBD5E1',pt:1},autoFit:false,colW:colW,rowH:0.38,margin:2});
  }

  async function addNRCPresentation(pptx) {
    var rows=await queryRawSheet('1010102020'), h=-1, data=[];
    for(var i=0;i<rows.length;i++) if(rows[i].some(function(v){return /^NRC Name$/i.test(clean(v));})){h=i;break;}
    if(h>=0) for(var r=h+1;r<rows.length;r++) if(clean(rows[r][0])) data.push(rows[r].slice(0,8));
    var first=data[0]||[], slide=pptx.addSlide();
    addHeader(slide,'🏥 NRC Kharsia','NRC REPORT | FY 2026–27');
    addSectionTitle(slide,'NRC KHARSIA — PERFORMANCE SUMMARY',0.45,1.00,6.5);
    addCard(slide,0.45,1.48,2.8,1.12,'TOTAL CHILDREN',first[1]||'');
    addCard(slide,3.45,1.48,2.8,1.12,'CHILDREN DISCHARGED',first[2]||'');
    addCard(slide,6.45,1.48,2.8,1.12,'BED OCCUPANCY RATE',first[6]||'');
    addCard(slide,9.45,1.48,2.8,1.12,'CURE RATE',first[7]||'');
    addSectionTitle(slide,'FACILITY WISE NRC REPORT',0.45,2.95,6.0);
    pptxTable(slide,['NRC Name','Total Children','Discharged','<7 Days','7–15 Days','>15 Days','Bed Occupancy %','Cure Rate'],data,0.45,3.38,12.4,2.35,8.5);
  }

  async function addBlindnessPresentation(pptx) {
    var rows=await queryRawSheet('1002009767');
    var re=/^(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*[-/]\s*\d{2,4}$/i;
    var months=rows.filter(function(r){return re.test(clean(r[1]));});
    var total=rows.find(function(r){return r.some(function(v){return /^total$/i.test(clean(v));});});
    var t=total||months[months.length-1]||[], slide=pptx.addSlide();
    addHeader(slide,'👁️ Blindness Control','MONTH WISE REPORT | FY 2026–27');
    addSectionTitle(slide,'BLINDNESS CONTROL — TOTAL PERFORMANCE',0.45,1.00,6.5);
    addCard(slide,0.45,1.48,2.8,1.12,'CATARACT TARGET',t[2]||'');
    addCard(slide,3.45,1.48,2.8,1.12,'CATARACT ACHIEVEMENT',t[3]||'');
    addCard(slide,6.45,1.48,2.8,1.12,'CATARACT TOTAL %',t[4]||'');
    addCard(slide,9.45,1.48,2.8,1.12,'STUDENT TARGET',t[8]||'');
    addSectionTitle(slide,'MONTH WISE PERFORMANCE',0.45,2.95,5.5);
    var heads=['Sn','Month','Cataract Target','Achievement','%','School Target','Visit School','%','Student Target','Achievement','%','Refractive Error','Spectle Provide'];
    var visible=months.slice(); if(total) visible.push(total);
    pptxTable(slide,heads,visible.map(function(r){var a=r.slice(0,13);while(a.length<13)a.push('');return a;}),0.18,3.35,12.95,2.9,6.5);
  }

  async function addNQASPresentation(pptx) {
    var rows=await queryRawSheet('728123647'), h=-1, data=[], note='';
    for(var i=0;i<rows.length;i++) if(rows[i].some(function(v){return /^Facility$/i.test(clean(v));})){h=i;break;}
    if(h>=0) for(var r=h+1;r<rows.length;r++){
      var t=rows[r].join(' ');
      if(/उपरोक्त संस्था के अलावा|^नोट/i.test(t)){note=t;continue;}
      if(/^(PHC|SHC|CHC)$/i.test(clean(rows[r][1]))) data.push(rows[r].slice(0,7));
    }
    var totals=[2,3,4,5,6].map(function(c){return data.reduce(function(a,r){return a+number(r[c]);},0);});
    var slide=pptx.addSlide();
    addHeader(slide,'🏅 NQAS Certification','NQAS CERTIFICATION DETAIL | FY 2026–27');
    addSectionTitle(slide,'NQAS CERTIFICATION SUMMARY — KHARSIA',0.45,1.00,7.0);
    addCard(slide,0.45,1.48,2.3,1.12,'TOTAL INSTITUTIONS',totals[0]);
    addCard(slide,2.95,1.48,2.3,1.12,'APPLICATIONS',totals[1]);
    addCard(slide,5.45,1.48,2.3,1.12,'EVALUATED',totals[2]);
    addCard(slide,7.95,1.48,2.3,1.12,'CERTIFIED',totals[3]);
    addCard(slide,10.45,1.48,2.3,1.12,'PENDING EVALUATION',totals[4]);
    addSectionTitle(slide,'FACILITY WISE NQAS STATUS',0.45,2.95,5.5);
    pptxTable(slide,['Sn','Facility','कुल संस्था','NQAS हेतु आवेदन','मूल्यांकन हो गया','कुल सर्टिफाईड','मूल्यांकन हेतु बाकी'],data,0.35,3.35,12.65,2.2,8.5);
    if(note) slide.addText('📝 '+note,{x:0.45,y:5.82,w:12.3,h:0.55,fontSize:10,bold:true,color:'92400E',fill:{color:'FFFBEB'},margin:0.12,fit:'shrink'});
  }

  function getSelectedPptxPrograms() {
    var defaults = ['Janani Portal','NCD','JAS Meeting','Ayushman Shivir','Wellness Activity','RBSK','Telemedicine Report','NRC Kharsia','Blindness Control','NQAS Certification'];
    var incoming = window.__pptxSelectedPrograms;
    if (Array.isArray(incoming) && incoming.length) return incoming.slice();
    return defaults.slice();
  }

  function isPptxSelected(name) {
    return getSelectedPptxPrograms().indexOf(name) >= 0;
  }

  async function addRCHPresentationOrdered(pptx, rows, sectors, total, moduleRanges) {
    var slide;
showProgress('Slides बन रही हैं', 55, 'Block Summary Dashboard तैयार हो रहा है...');
// 3. SECTOR WISE DATA — RCH 2.0 content slides in the selected order.
var rchStartPage = pptx.slides.length + 1;
showProgress('Slides बन रही हैं', 60, 'RCH 2.0 Sector Wise Data और color coding तैयार हो रही है...');

slide = pptx.addSlide();
addHeader(slide, 'Janani Portal (RCH 2.0)', 'SECTOR WISE DATA | FY 2026–27');
var chartRows = sectors.slice().sort(function (a, b) { return b.percent - a.percent; });

// Colored sector table
var tx = 0.45, ty = 1.18, tw = 12.45;
var rowH = 0.60;
var widths = [3.05, 1.78, 2.00, 1.68, 1.92, 2.02];
var headers = ['Sector', 'HMIS PW', 'RCH 2.0 PW', 'Achievement', 'Backlog', 'High Risk'];
var cx = tx;
headers.forEach(function(h, i) {
  slide.addShape('rect', {
    x:cx, y:ty, w:widths[i], h:rowH,
    fill:{color:'075985'}, line:{color:'FFFFFF', pt:1}
  });
  slide.addText(h, {
    x:cx+0.04, y:ty+0.12, w:widths[i]-0.08, h:0.32,
    fontSize:18, bold:true, color:'FFFFFF', align:'center', margin:0, fit:'shrink'
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
        x:cx+0.18, y:y+0.12, w:widths[i]-0.36, h:0.34,
        fontSize:18, bold:true, color:pctColor, align:'center', margin:0
      });
    } else if (i === 4) {
      slide.addShape('roundRect', {
        x:cx+0.25, y:y+0.075, w:widths[i]-0.5, h:0.33,
        fill:{color:backlogFill}, line:{color:backlogColor, pt:1}
      });
      slide.addText(String(v), {
        x:cx+0.20, y:y+0.12, w:widths[i]-0.40, h:0.34,
        fontSize:18, bold:true, color:backlogColor, align:'center', margin:0
      });
    } else {
      slide.addText(String(v), {
        x:cx+0.05, y:y+0.12, w:widths[i]-0.10, h:0.34,
        fontSize:i===0?18:17, bold:i===0, color:'172033',
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
      x:cx+0.05, y:totalY+0.12, w:widths[i]-0.10, h:0.30,
      fontSize:17, bold:true, color:'FFFFFF', align:'center', margin:0
    });
  } else if (i === 4) {
    var totalBacklogColor = total.backlog > 0 ? '86EFAC' : (total.backlog === 0 ? 'FDE68A' : 'FCA5A5');
    slide.addText(String(v), {
      x:cx+0.05, y:totalY+0.12, w:widths[i]-0.10, h:0.30,
      fontSize:16, bold:true, color:totalBacklogColor, align:'center', margin:0
    });
  } else {
    slide.addText(String(v), {
      x:cx+0.05, y:totalY+0.17, w:widths[i]-0.10, h:0.12,
      fontSize:14, bold:true, color:'FFFFFF',
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
  { title:'Sarwani', sectors:['sarwani'] },
  { title:'Turekela', sectors:['turekela'] },
  { title:'Sondka', sectors:['sondka'] },
  { title:'Binjkot', sectors:['binjkot'] }
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

  // Each sector uses: 1 sector band + facility rows + 1 TOTAL row.
  // Include both the sector band and TOTAL in the height calculation so
  // the last sector (Gorpar) never runs outside the slide.
  var lineCount = 1;
  orderedSectors.forEach(function(k) { lineCount += grouped[k].length + 2; });
  // Fit the complete facility-wise table inside the 7.5in slide.
  // Keep rows as large as possible while reserving room for header/footer.
  var availableH = 5.72;
  var rowH = Math.max(0.30, Math.min(0.50, availableH / Math.max(lineCount,1)));
  var sectionH = rowH;

  var cx = tx;
  headers.forEach(function(h,i) {
    slide.addShape('rect',{x:cx,y:ty,w:widths[i],h:rowH,fill:{color:'075985'},line:{color:'FFFFFF',pt:1}});
    slide.addText(h,{x:cx+0.03,y:ty+rowH*0.28,w:widths[i]-0.06,h:rowH*0.38,fontSize:Math.max(11,Math.min(15,rowH*34)),bold:true,color:'FFFFFF',align:'center',margin:0,fit:'shrink'});
    cx += widths[i];
  });

  var currentY = ty + rowH;
  orderedSectors.forEach(function(sectorName) {
    var sectorRows = grouped[sectorName] || [];
    // Sector band
    slide.addShape('rect',{x:tx,y:currentY,w:12.49,h:sectionH,fill:{color:'E0F2FE'},line:{color:'B6D7EA',pt:1}});
    slide.addText('SECTOR: ' + sectorName.toUpperCase(),{
      x:tx+0.08,y:currentY+rowH*0.26,w:12.25,h:rowH*0.38,
      fontSize:Math.max(10,Math.min(14,rowH*32)),bold:true,color:'075985',margin:0,fit:'shrink'
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
          slide.addText(String(v),{x:cx+0.18,y:currentY+rowH*0.31,w:widths[i]-0.36,h:rowH*0.30,fontSize:Math.max(11,Math.min(16,rowH*38)),bold:true,color:pctColor,align:'center',margin:0,fit:'shrink'});
        } else if(i===4) {
          slide.addShape('roundRect',{x:cx+0.18,y:currentY+rowH*0.13,w:widths[i]-0.36,h:rowH*0.70,fill:{color:backlogFill},line:{color:backlogColor,pt:0.8}});
          slide.addText(String(v),{x:cx+0.18,y:currentY+rowH*0.31,w:widths[i]-0.36,h:rowH*0.30,fontSize:Math.max(11,Math.min(16,rowH*38)),bold:true,color:backlogColor,align:'center',margin:0,fit:'shrink'});
        } else {
          slide.addText(String(v),{x:cx+0.04,y:currentY+rowH*0.31,w:widths[i]-0.08,h:rowH*0.30,fontSize:Math.max(11,Math.min(16,rowH*38)),bold:i===0,color:'172033',align:i===0?'left':'center',margin:0,fit:'shrink'});
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
      var fillColor = i === 3 ? (sectorPct >= 90 ? '16A34A' : (sectorPct >= 70 ? 'F59E0B' : 'DC2626')) : '0F766E';
      slide.addShape('rect',{x:cx,y:currentY,w:widths[i],h:rowH,fill:{color:'0F766E'},line:{color:'FFFFFF',pt:1.2}});
      slide.addText(String(v),{
        x:cx+0.04,y:currentY+rowH*0.29,w:widths[i]-0.08,h:rowH*0.34,
        fontSize:Math.max(11,Math.min(16,rowH*38)),bold:true,
        color:'FFFFFF',align:i===0?'left':'center',margin:0,fit:'shrink'
      });
      if (i === 3) {
        slide.addShape('roundRect',{x:cx+0.18,y:currentY+rowH*0.13,w:widths[i]-0.36,h:rowH*0.70,fill:{color:fillColor},line:{color:'FFFFFF',pt:0.8}});
        slide.addText(String(v),{x:cx+0.18,y:currentY+rowH*0.31,w:widths[i]-0.36,h:rowH*0.30,fontSize:Math.max(11,Math.min(16,rowH*38)),bold:true,color:'FFFFFF',align:'center',margin:0,fit:'shrink'});
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
  end: pptx.slides.length
});
    return moduleRanges;
  }

  async function generate() {
    var selectedPrograms = getSelectedPptxPrograms();
    if(!selectedPrograms.length){
      throw new Error('कम-से-कम एक Programme select करें।');
    }
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
          '<g><rect x="1197" y="706" width="116" height="76" rx="16" fill="#efffbf"/><text x="1255" y="750" fill="#3f6212">Blindness</text></g>' +
          '<g><rect x="1325" y="706" width="116" height="76" rx="16" fill="#fff1b8"/><text x="1383" y="750" fill="#854d0e">NQAS</text></g>' +
        '</g>' +
        '<rect y="820" width="1600" height="80" fill="#073b75"/>' +
        '<text x="70" y="870" font-family="Arial, Noto Sans, sans-serif" font-size="23" font-weight="800" fill="#ffffff">Health Department  |  District Raigarh  |  Chhattisgarh</text>' +
        '<text x="1355" y="870" text-anchor="end" font-family="Arial, Noto Sans, sans-serif" font-size="23" font-weight="900" fill="#ffd900">Monthly Performance Review</text><rect x="1435" y="835" width="120" height="43" rx="18" fill="#ffffff"/><text x="1495" y="864" text-anchor="middle" font-family="Arial, Noto Sans, sans-serif" font-size="20" font-weight="900" fill="#073b75">Page 01</text>' +
        '</svg>';

      var slide = pptx.addSlide();
      slide.background = { color:'F8FBFF' };
      slide.addImage({ data: svgDataUri(coverSvg), x:0, y:0, w:13.333, h:7.5 });

      // 2. INDEX — reserve slide 2 immediately.
      // This prevents the last generated module slide from ever being moved to page 2.
      var indexSlide = pptx.addSlide();
      indexSlide.background = { color:'F7FBFF' };

      var moduleRanges = [];
      var selectedOrder = getSelectedPptxPrograms();

      var allModuleDefs = {
        'NCD': {name:'NCD',icon:'❤️',gid:'1254412412'},
        'JAS Meeting': {name:'JAS Meeting',icon:'🤝',gid:'1018164338'},
        'Ayushman Shivir': {name:'Ayushman Shivir',icon:'🏕️',gid:'1262815420'},
        'Wellness Activity': {name:'Wellness Activity',icon:'🩺',gid:'447031017'},
        'Telemedicine Report': {name:'Telemedicine Report',icon:'🩻',gid:'318264987'},
        'RBSK': {name:'RBSK',icon:'👶',gid:null},
        'NRC Kharsia': {name:'NRC Kharsia',icon:'🏥',gid:'1010102020'},
        'Blindness Control': {name:'Blindness Control',icon:'👁️',gid:'1002009767'},
        'NQAS Certification': {name:'NQAS Certification',icon:'🏅',gid:'728123647'}
      };

      showProgress('Programme order', 55, 'Selected order के अनुसार PPT slides तैयार हो रही हैं...');

      for (var oi = 0; oi < selectedOrder.length; oi++) {
        var selectedName = selectedOrder[oi];
        var progress = 55 + Math.round((oi / Math.max(selectedOrder.length,1)) * 35);
        showProgress('Programme ' + (oi + 1) + '/' + selectedOrder.length, progress, selectedName + ' की slides तैयार हो रही हैं...');

        if (selectedName === 'Janani Portal') {
          await addRCHPresentationOrdered(pptx, rows, sectors, total, moduleRanges);
          continue;
        }

        var mod = allModuleDefs[selectedName];
        if (!mod) continue;

        var moduleStart = pptx.slides.length + 1;

        if (mod.name === 'NCD') {
          await addNCDPresentation(pptx);
        } else if (mod.name === 'JAS Meeting') {
          await addJASPresentation(pptx);
        } else if (mod.name === 'Ayushman Shivir') {
          await addAyushmanShivirPresentation(pptx);
        } else if (mod.name === 'Wellness Activity') {
          await addWellnessActivityPresentation(pptx);
        } else if (mod.name === 'Telemedicine Report') {
          await addTelemedicinePresentation(pptx);
        } else if (mod.name === 'RBSK') {
          await addRBSKPresentation(pptx);
        } else if (mod.name === 'NRC Kharsia') {
          await addNRCPresentation(pptx);
        } else if (mod.name === 'Blindness Control') {
          await addBlindnessPresentation(pptx);
        } else if (mod.name === 'NQAS Certification') {
          await addNQASPresentation(pptx);
        }

        var moduleEnd = pptx.slides.length;
        if (moduleEnd >= moduleStart) {
          moduleRanges.push({name:mod.name,start:moduleStart,end:moduleEnd});
        }
      }

      showProgress('Final analysis', 92, 'Overall Data Analysis slide तैयार हो रही है...');
      // OVERALL ANALYSIS / KEY OBSERVATIONS slide removed as requested.
      // Build the already-reserved slide 2 as the final Index.
      indexSlide.background = { color:'F7FBFF' };

      function pageRangeText(item) {
        return item.start === item.end
          ? 'Page ' + String(item.start).padStart(2,'0')
          : 'Page ' + String(item.start).padStart(2,'0') + '–' + String(item.end).padStart(2,'0');
      }

      var indexMeta = {
        'Janani Portal': {name:'Janani Portal', rangeName:'RCH 2.0', icon:'R', color:'#7C3AED', fill:'#EDE9FE'},
        'NCD': {name:'NCD', icon:'♥', color:'#0284C7', fill:'#E0F2FE'},
        'JAS Meeting': {name:'JAS Meeting', icon:'J', color:'#D97706', fill:'#FEF3C7'},
        'Ayushman Shivir': {name:'Ayushman Shivir', icon:'S', color:'#DB2777', fill:'#FCE7F3'},
        'Wellness Activity': {name:'Wellness Activity', icon:'W', color:'#9333EA', fill:'#F3E8FF'},
        'Telemedicine Report': {name:'Telemedicine Report', icon:'T', color:'#0F766E', fill:'#CCFBF1'},
        'RBSK': {name:'RBSK', icon:'R', color:'#4338CA', fill:'#EDE9FE'},
        'NRC Kharsia': {name:'NRC Kharsia', icon:'N', color:'#075985', fill:'#E0F2FE'},
        'Blindness Control': {name:'Blindness Control', icon:'B', color:'#4338CA', fill:'#EEF2FF'},
        'NQAS Certification': {name:'NQAS Certification', icon:'Q', color:'#0F766E', fill:'#CCFBF1'}
      };

      var desiredModules = getSelectedPptxPrograms().map(function(name){
        return indexMeta[name];
      }).filter(function(x){ return !!x; });

      var indexItems = desiredModules.map(function(dm) {
        var found = moduleRanges.find(function(r){ return r.name === (dm.rangeName || dm.name); });
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

      var itemCount=indexItems.length;
      var cols=itemCount<=3 ? 1 : 2;
      var rowsPerCol=cols===1 ? itemCount : Math.ceil(itemCount/2);
      var cardX=cols===1 ? 185 : 55;
      var cardGap=cols===1 ? 24 : 25;
      var cardW=cols===1 ? 1030 : 510;
      var topY=278;
      var bottomY=760;
      var maxH=53;
      var availableH=bottomY-topY;
      var rowStep=Math.min(82, Math.max(58, availableH/Math.max(rowsPerCol,1)));
      var cardH=Math.min(maxH,rowStep-7);
      var startY=topY;

      indexItems.forEach(function(item,i){
        var col = cols===1 ? 0 : (i < rowsPerCol ? 0 : 1);
        var row = cols===1 ? i : (i < rowsPerCol ? i : i-rowsPerCol);
        var x = cols===1 ? cardX : (col===0 ? 55 : 590);
        var y = startY + row*rowStep;
        var page = item.start ? pageRangeText(item) : 'Data Pending';
        svg += '<g filter="url(#shadow)">' +
          '<rect x="'+x+'" y="'+y+'" width="'+cardW+'" height="'+cardH+'" rx="14" fill="'+item.fill+'" stroke="'+item.color+'" stroke-opacity=".28" stroke-width="2"/>' +
          '<rect x="'+(x+9)+'" y="'+(y+8)+'" width="42" height="'+Math.max(30,cardH-16)+'" rx="10" fill="'+item.color+'"/>' +
          '<text x="'+(x+30)+'" y="'+(y+Math.max(22,cardH*0.61))+'" text-anchor="middle" font-family="Arial" font-size="16" font-weight="900" fill="#FFFFFF">'+String(i+1).padStart(2,'0')+'</text>' +
          '<circle cx="'+(x+76)+'" cy="'+(y+cardH/2)+'" r="18" fill="#FFFFFF" stroke="'+item.color+'" stroke-width="2"/>' +
          '<text x="'+(x+76)+'" y="'+(y+cardH/2+7)+'" text-anchor="middle" font-family="Arial" font-size="18" font-weight="900" fill="'+item.color+'">'+esc(item.icon)+'</text>' +
          '<text x="'+(x+108)+'" y="'+(y+cardH*0.64)+'" font-family="Arial, Noto Sans, sans-serif" font-size="'+(item.name.length>25?18:20)+'" font-weight="800" fill="#0F2D52">'+esc(item.name)+'</text>' +
          '<rect x="'+(x+cardW-115)+'" y="'+(y+9)+'" width="105" height="'+Math.max(30,cardH-18)+'" rx="11" fill="#FFFFFF" stroke="'+item.color+'" stroke-opacity=".30"/>' +
          '<text x="'+(x+cardW-62.5)+'" y="'+(y+cardH*0.60)+'" text-anchor="middle" font-family="Arial" font-size="12" font-weight="800" fill="'+(item.start?'#075985':'#C2410C')+'">'+esc(page)+'</text>' +
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

      // Index is already slide 2; no slide reordering is needed.
      // Final order: Cover → Index → RCH → NCD → JAS → Shivir → Activity → Telemedicine.

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
      window.__lastGeneratedPptxBlob = blob;
      var url = URL.createObjectURL(blob);
      var old = document.getElementById('pptxDownloadFallback');
      if (old) old.remove();

      if (window.__generatingPDF) {
        showProgress('PPTX तैयार है', 100, 'अब इसी exact PPTX layout से PDF तैयार हो रही है...');
        return blob;
      }

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