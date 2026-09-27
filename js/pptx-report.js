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
    addSectionTitle(slide,'SECTOR WISE DATA',0.34,0.94,4.2);

    // Fit the complete sector table on one slide with larger, readable fonts.
    var x=0.24, y=1.34;
    var w=[0.62,2.20,1.15,2.00,2.08,2.08,1.86];
    var headers=['SN','Sector','NIN','No of AAM Facility','Target till Aug 26','Achievement','%'];
    var headerH=0.70, dataH=0.57, totalH=0.60;
    var cx=x;

    headers.forEach(function(h,i){
      slide.addShape('rect',{
        x:cx,y:y,w:w[i],h:headerH,
        fill:{color:'075985'},line:{color:'FFFFFF',pt:1}
      });
      slide.addText(h,{
        x:cx+0.03,y:y+0.12,w:w[i]-0.06,h:headerH-0.18,
        fontSize:15,bold:true,color:'FFFFFF',
        align:'center',valign:'mid',margin:0,fit:'shrink'
      });
      cx+=w[i];
    });

    sectorRows.forEach(function(r,ri){
      var cy=y+headerH+(ri*dataH), vals=[
        r[0],r[1],r[2],number(r[3]),number(r[4]),number(r[5])
      ];
      cx=x;
      vals.forEach(function(v,ci){
        slide.addShape('rect',{
          x:cx,y:cy,w:w[ci],h:dataH,
          fill:{color:ri%2?'F8FBFF':'FFFFFF'},
          line:{color:'CBD5E1',pt:0.8}
        });
        slide.addText(String(v),{
          x:cx+0.04,y:cy+0.10,w:w[ci]-0.08,h:dataH-0.17,
          fontSize:14,bold:ci===1,
          color:'172033',align:ci===1?'left':'center',
          valign:'mid',margin:0,fit:'shrink'
        });
        cx+=w[ci];
      });

      slide.addShape('rect',{
        x:cx,y:cy,w:w[6],h:dataH,
        fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:0.8}
      });
      jasAddPctCell(slide,cx,cy,w[6],dataH,r[6]);
    });

    var target=sectorRows.reduce(function(s,r){return s+number(r[4]);},0);
    var ach=sectorRows.reduce(function(s,r){return s+number(r[5]);},0);
    var totalPct=target?ach/target*100:0;
    var ty=y+headerH+(sectorRows.length*dataH), tc=x;
    var tv=['TOTAL','',sectorRows.reduce(function(s,r){return s+number(r[3]);},0),target,ach,totalPct.toFixed(1)];
    var totalCols=[2.82,1.15,2.00,2.08,2.08,1.86];

    tv.forEach(function(v,ci){
      slide.addShape('rect',{
        x:tc,y:ty,w:totalCols[ci],h:totalH,
        fill:{color:'0F766E'},line:{color:'FFFFFF',pt:1}
      });
      if(ci===5){
        jasAddPctCell(slide,tc,ty,totalCols[ci],totalH,v);
      }else{
        slide.addText(String(v),{
          x:tc+0.04,y:ty+0.12,w:totalCols[ci]-0.08,h:totalH-0.18,
          fontSize:14,bold:true,color:'FFFFFF',
          align:ci===0?'left':'center',valign:'mid',
          margin:0,fit:'shrink'
        });
      }
      tc+=totalCols[ci];
    });

    slide.addText(
      'Performance: '+sectorRows.map(function(r){
        return clean(r[1])+' '+number(r[6])+'%';
      }).join('  •  '),
      {
        x:0.30,y:6.86,w:12.70,h:0.20,
        fontSize:9.5,bold:true,color:'475569',
        margin:0,fit:'shrink',align:'center'
      }
    );
    slide.addText('Source: live Google Sheet • '+title,{
      x:0.45,y:7.10,w:12.0,h:0.14,
      fontSize:7.5,color:'64748B',margin:0,fit:'shrink'
    });
    return slide;
  }

  function jasAddAllFacilitySlide(pptx,title,rows){
    var slide=pptx.addSlide();
    addHeader(slide,'🤝 JAS Meeting — Analysis Report','FACILITY MEETING ANALYSIS | FY 2026–27');
    addSectionTitle(slide,'JAS MEETING — ANALYSIS REPORT',0.34,0.94,6.2);

    var data=(rows||[]).filter(function(r){
      return clean(r[3]) && (jasLooksNumber(r[4]) || jasLooksNumber(r[5]));
    }).map(function(r){
      var target=jasLooksNumber(r[4])?number(r[4]):0;
      var held=jasLooksNumber(r[5])?number(r[5]):0;
      return {
        sector:clean(r[2]),
        facility:clean(r[3]),
        target:target,
        held:held,
        gap:Math.max(target-held,0)
      };
    });

    var groups={};
    data.forEach(function(r){
      var k=String(Math.round(r.held));
      if(!groups[k]) groups[k]=[];
      groups[k].push(r);
    });

    var keys=Object.keys(groups).map(Number).sort(function(a,b){return a-b;});
    var facilityCount=data.length;
    var totalTarget=data.reduce(function(s,r){return s+r.target;},0);
    var totalHeld=data.reduce(function(s,r){return s+r.held;},0);
    var totalGap=Math.max(totalTarget-totalHeld,0);
    var overall=totalTarget?totalHeld/totalTarget*100:0;
    var fullCount=(groups['5']||[]).length;
    var zeroCount=(groups['0']||[]).length;

    addCard(slide,0.34,1.28,2.35,0.98,'TOTAL FACILITIES',facilityCount);
    addCard(slide,2.82,1.28,2.35,0.98,'TARGET MEETINGS',totalTarget);
    addCard(slide,5.30,1.28,2.35,0.98,'MEETINGS HELD',totalHeld);
    addCard(slide,7.78,1.28,2.35,0.98,'GAP',totalGap);
    addCard(slide,10.26,1.28,2.73,0.98,'OVERALL %',overall.toFixed(1)+'%');

    addSectionTitle(slide,'MEETING-WISE FACILITY LIST',0.34,2.60,6.1);

    // Two-column meeting groups. Every meeting-count group shows the facility names.
    var leftKeys=keys.filter(function(k){return k<=2;});
    var rightKeys=keys.filter(function(k){return k>=3;});

    function drawGroupColumn(listKeys, x){
      var top=3.00;
      var available=3.78;
      var gapY=0.12;
      var n=listKeys.length||1;
      var boxH=(available-(Math.max(0,n-1)*gapY))/n;

      listKeys.forEach(function(k,idx){
        var list=groups[String(k)]||[];
        var y=top+idx*(boxH+gapY);
        var share=facilityCount?list.length/facilityCount*100:0;
        var titleColor=k===0?'DC2626':(k<5?'CA8A04':'16A34A');
        var soft=k===0?'FEE2E2':(k<5?'FEF3C7':'DCFCE7');

        slide.addShape('roundRect',{
          x:x,y:y,w:6.08,h:boxH,
          fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:0.9}
        });
        slide.addShape('rect',{
          x:x,y:y,w:1.06,h:boxH,
          fill:{color:soft},line:{color:soft}
        });
        slide.addText(String(k)+' meetings',{
          x:x+0.08,y:y+0.12,w:0.90,h:0.28,
          fontSize:13,bold:true,color:titleColor,
          align:'center',margin:0,fit:'shrink'
        });
        slide.addText(String(list.length)+' facilities\n'+share.toFixed(1)+'%',{
          x:x+0.08,y:y+0.47,w:0.90,h:Math.max(0.34,boxH-0.56),
          fontSize:9.5,bold:true,color:'475569',
          align:'center',valign:'mid',margin:0,fit:'shrink'
        });

        var names=list
          .slice()
          .sort(function(a,b){return a.facility.localeCompare(b.facility);})
          .map(function(r){return r.facility;})
          .join('  •  ');

        slide.addText(names||'No facility',{
          x:x+1.20,y:y+0.12,w:4.72,h:Math.max(0.35,boxH-0.22),
          fontSize:boxH>1.0?10:8.5,
          bold:true,color:'172033',
          align:'left',valign:'mid',
          margin:0.02,fit:'shrink',breakLine:false
        });
      });
    }

    drawGroupColumn(leftKeys,0.34);
    drawGroupColumn(rightKeys,6.72);

    addSectionTitle(slide,'KEY ANALYSIS / OBSERVATIONS',0.34,6.98,4.1);
    var bullets=[
      '• '+fullCount+' of '+facilityCount+' facilities ('+(facilityCount?fullCount/facilityCount*100:0).toFixed(1)+'%) have completed all 5 meetings.',
      '• '+(groups['4']||[]).length+' facilities are at 4 meetings and need 1 more meeting.',
      '• '+(groups['3']||[]).length+' at 3 meetings; '+(groups['2']||[]).length+' at 2; '+(groups['1']||[]).length+' at 1.',
      '• '+zeroCount+' facilities have recorded 0 meetings.',
      '• Overall achievement is '+overall.toFixed(1)+'%; total gap is '+totalGap+' meetings.'
    ];
    slide.addText(bullets.join('\n'),{
      x:4.25,y:6.90,w:8.45,h:0.38,
      fontSize:8.5,bold:true,color:'172033',
      margin:0.01,fit:'shrink',align:'left'
    });

    slide.addText('NIN number excluded • Meeting Held = Achievement column • Source: live Google Sheet • '+title,{
      x:0.45,y:7.28,w:12.0,h:0.10,
      fontSize:6.8,color:'64748B',margin:0,fit:'shrink',align:'center'
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
    // 1) Sector Wise — large font, full slide
    // 2) All Facility Wise — one slide, globally ranked highest to lowest
    jasAddSectorSlide(pptx,title,parsed.sector);

    jasAddAllFacilitySlide(pptx,title,parsed.facility);

    return {start:start,end:pptx.slides.length};
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

      // 2. INDEX — created after all content slides so page ranges are always exact.
      // The index slide is moved to position 2 after all modules are generated.
      var moduleRanges = [];
      var otherModules = [
        {name:'NCD',icon:'❤️',gid:'1254412412'},
        {name:'JAS Meeting',icon:'🤝',gid:'1018164338'},
        {name:'Health & Wellness Center',icon:'🏥',gid:'0'},
        {name:'Ayushman Shivir',icon:'🏕️',gid:'1262815420'},
        {name:'Wellness Activity',icon:'🩺',gid:'447031017'},
        {name:'RBSK',icon:'👶',gid:'1502752823'},
        {name:'NRC Kharsia',icon:'🏥',gid:'1010102020'},
        {name:'Blindness Control',icon:'👁️',gid:'1002009767'},
        {name:'NQAS Certification',icon:'🏅',gid:'728123647'},

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
        var rowH = Math.max(0.30, Math.min(0.46, availableH / Math.max(lineCount,1)));
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
                slide.addText(String(v),{x:cx+0.18,y:currentY+rowH*0.31,w:widths[i]-0.36,h:rowH*0.30,fontSize:Math.max(9,Math.min(14,rowH*34)),bold:true,color:pctColor,align:'center',margin:0,fit:'shrink'});
              } else if(i===4) {
                slide.addShape('roundRect',{x:cx+0.18,y:currentY+rowH*0.13,w:widths[i]-0.36,h:rowH*0.70,fill:{color:backlogFill},line:{color:backlogColor,pt:0.8}});
                slide.addText(String(v),{x:cx+0.18,y:currentY+rowH*0.31,w:widths[i]-0.36,h:rowH*0.30,fontSize:Math.max(9,Math.min(14,rowH*34)),bold:true,color:backlogColor,align:'center',margin:0,fit:'shrink'});
              } else {
                slide.addText(String(v),{x:cx+0.04,y:currentY+rowH*0.31,w:widths[i]-0.08,h:rowH*0.30,fontSize:Math.max(9,Math.min(14,rowH*34)),bold:i===0,color:'172033',align:i===0?'left':'center',margin:0,fit:'shrink'});
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
              fontSize:Math.max(9,Math.min(14,rowH*34)),bold:true,
              color:'FFFFFF',align:i===0?'left':'center',margin:0,fit:'shrink'
            });
            if (i === 3) {
              slide.addShape('roundRect',{x:cx+0.18,y:currentY+rowH*0.13,w:widths[i]-0.36,h:rowH*0.70,fill:{color:fillColor},line:{color:'FFFFFF',pt:0.8}});
              slide.addText(String(v),{x:cx+0.18,y:currentY+rowH*0.31,w:widths[i]-0.36,h:rowH*0.30,fontSize:Math.max(9,Math.min(14,rowH*34)),bold:true,color:'FFFFFF',align:'center',margin:0,fit:'shrink'});
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
        if (mod.name === 'NCD') {
          await addNCDPresentation(pptx);
        }
        else if (mod.name === 'JAS Meeting') {
          var jasRange = await addJASPresentation(pptx);
          moduleRanges.push({ name:'JAS Meeting', start:jasRange.start, end:jasRange.end });
        } else {
          var genericData = await queryGenericSheet(mod.gid);
          addGenericModuleSlide(pptx, mod, genericData, mi);
        }
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
      ];

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