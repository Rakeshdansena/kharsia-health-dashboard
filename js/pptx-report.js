/* Kharsia Health Dashboard — PPTX Generator
   Generates a meeting-ready 16:9 PowerPoint from the live rendered dashboard tables.
   No hard-coded programme numbers: values are read from the currently loaded Google Sheet data.
*/
(function(){
'use strict';

const MODULES=[
  'Ayushman Card','RCH 2.0','NCD','JAS Meeting','Health & Wellness Center',
  'Ayushman Shivir','Wellness Activity','RBSK','NRC Kharsia',
  'Blindness Control','NQAS Certification','Dialysis','NLEP'
];

function loadPptxLib(){
  return new Promise((resolve,reject)=>{
    if(window.PptxGenJS)return resolve();
    const s=document.createElement('script');
    s.src='https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
    s.onload=()=>resolve();
    s.onerror=()=>reject(new Error('PowerPoint library load failed'));
    document.head.appendChild(s);
  });
}

function text(x){return String(x==null?'':x).replace(/\s+/g,' ').trim();}
function num(x){
  const m=text(x).replace(/,/g,'').match(/-?\d+(?:\.\d+)?/);
  return m?Number(m[0]):NaN;
}
function esc(x){return text(x).replace(/[&<>]/g,'');}
function visible(el){return !!(el && el.offsetParent!==null && getComputedStyle(el).display!=='none');}

function findButtons(){
  return [...document.querySelectorAll('.menu-btn')].filter(b=>{
    const t=text(b.textContent);
    return MODULES.some(m=>t.toLowerCase().includes(m.toLowerCase()));
  });
}

function clickModule(name){
  const b=findButtons().find(x=>text(x.textContent).toLowerCase().includes(name.toLowerCase()));
  if(!b) return false;
  b.click(); return true;
}

function wait(ms){return new Promise(r=>setTimeout(r,ms));}

async function waitForModule(name){
  for(let i=0;i<40;i++){
    await wait(400);
    const title=text(document.getElementById('reportTitle')?.textContent);
    const area=document.getElementById('pdfArea');
    const body=text(area?.innerText);
    const ncd=area?.querySelector('#ncdModuleContainer');
    const tables=area?[...area.querySelectorAll('table')].filter(visible):[];
    if(name==='NCD' && ncd && ncd.querySelectorAll('table').length>=2) return;
    if(name!=='NCD' && (title || body.length>40 || tables.length)) return;
  }
}

function collectTables(){
  const area=document.getElementById('pdfArea')||document.body;
  const tables=[...area.querySelectorAll('table')].filter(visible);
  return tables.map((table,ti)=>{
    const rows=[...table.rows].map(tr=>[...tr.cells].map(c=>text(c.innerText)));
    return {rows,index:ti};
  }).filter(x=>x.rows.length);
}

function collectModule(name){
  const title=text(document.getElementById('reportTitle')?.innerText)||name;
  const ncd=document.getElementById('ncdModuleContainer');
  const heading=text(document.querySelector('.shivir-page-heading .subheading')?.innerText) ||
                (name==='NCD' ? text(ncd?.querySelector('.ncd-final-title')?.innerText) : '') || title;
  const status=text(document.getElementById('status')?.innerText);
  let tables=collectTables();
  if(name==='NCD' && ncd){
    const nt=[...ncd.querySelectorAll('table')].map((table,ti)=>({rows:[...table.rows].map(tr=>[...tr.cells].map(c=>text(c.innerText))),index:ti})).filter(x=>x.rows.length);
    if(nt.length) tables=nt;
  }
  return {name,title,heading,status,tables};
}

function analyse(mod){
  const facts=[];
  const allRows=mod.tables.flatMap(t=>t.rows);
  const dataRows=allRows.slice(1);
  let pct=[], nums=[];
  dataRows.forEach(r=>r.forEach(v=>{
    const s=text(v);
    if(/%$/.test(s)){const n=num(s);if(Number.isFinite(n))pct.push(n);}
    else {const n=num(s);if(Number.isFinite(n))nums.push(n);}
  }));
  if(pct.length){
    const avg=pct.reduce((a,b)=>a+b,0)/pct.length;
    facts.push('Average reported percentage: '+avg.toFixed(1)+'%.');
    facts.push('Observed percentage range: '+Math.min(...pct).toFixed(1)+'%–'+Math.max(...pct).toFixed(1)+'%.');
    const low=pct.filter(x=>x<50).length;
    if(low) facts.push(low+' reported percentage value(s) are below 50%.');
  }
  const totals=allRows.filter(r=>r.some(v=>/^total$|^कुल$|^योग$/i.test(text(v))));
  if(totals.length) facts.push('Block/facility total row is available in the rendered report.');
  if(nums.length) facts.push('Numeric indicators are available across '+dataRows.length+' rendered data rows.');
  if(!facts.length) facts.push('Report loaded; key points are based on the live rendered Google Sheet table.');
  return facts;
}

function extractMetrics(mod){
  const rows=mod.tables.flatMap(t=>t.rows);
  const header=rows[0]||[];
  const sample=rows.find(r=>r.some(v=>/^total$|^कुल$|^योग$/i.test(text(v))))||rows[1]||[];
  const out=[];
  header.forEach((h,i)=>{
    const raw=text(sample[i]);
    const n=num(raw);
    if(Number.isFinite(n) && text(h)) out.push({label:text(h),value:n,raw});
  });
  return out.slice(0,8);
}

function ayushmanMetrics(mod){
  const rows=mod.tables.flatMap(t=>t.rows);
  const totalRow=rows.find(r=>r.some(v=>/^total$|^कुल$|^योग$/i.test(text(v)))) || rows[rows.length-1] || rows[1] || [];
  const header=rows[0]||[];
  const out=[];
  header.forEach((h,i)=>{
    const label=text(h), raw=text(totalRow[i]), value=num(raw);
    if(label && Number.isFinite(value) && !/^s\.?\s*no|^sn$/i.test(label)) out.push({label,value,raw});
  });
  return out.slice(0,8);
}

function addAyushmanHeroSlide(pptx,mod){
  const slide=pptx.addSlide();
  slide.background={color:'075985'};
  slide.addShape(pptx.ShapeType.arc,{x:9.5,y:-1.15,w:4.5,h:4.5,line:{color:'14B8A6',pt:3}});
  slide.addText('AYUSHMAN CARD',{x:.65,y:.55,w:5.2,h:.55,fontSize:30,bold:true,color:'FFFFFF',margin:0,charSpacing:1});
  slide.addText('Programme Performance Snapshot',{x:.68,y:1.15,w:5.8,h:.35,fontSize:17,bold:true,color:'99F6E4',margin:0});
  slide.addText(mod.heading||mod.title||'आयुष्मान कार्ड रिपोर्ट',{x:.7,y:1.75,w:11.8,h:.55,fontSize:18,bold:true,color:'FFFFFF',margin:0,fit:'shrink'});
  slide.addText('Block Kharsia  •  District Raigarh  •  Chhattisgarh',{x:.7,y:2.35,w:9.5,h:.3,fontSize:12,color:'E2E8F0',margin:0});
  const metrics=ayushmanMetrics(mod);
  metrics.slice(0,6).forEach((m,i)=>{
    const col=i%3,row=Math.floor(i/3),x=.7+col*4.05,y=3.05+row*1.55;
    slide.addShape(pptx.ShapeType.roundRect,{x,y,w:3.7,h:1.2,rectRadius:.08,fill:{color:'FFFFFF'},line:{color:'D6E3EC',pt:1}});
    slide.addText(m.label,{x:x+.18,y:y+.16,w:3.3,h:.32,fontSize:11,bold:true,color:'0F766E',margin:0,fit:'shrink'});
    slide.addText(m.raw,{x:x+.18,y:y+.5,w:3.3,h:.5,fontSize:25,bold:true,color:'0B3558',margin:0,fit:'shrink'});
  });
  slide.addText('Live Google Sheet data • Generated from the current dashboard view',{x:.7,y:7.05,w:8,h:.2,fontSize:8,color:'CBD5E1',margin:0});
}

function addAyushmanVisualSlide(pptx,mod){
  const metrics=ayushmanMetrics(mod);
  const slide=pptx.addSlide();
  slide.background={color:'F4F7FB'};
  addTitle(slide,'Ayushman Card','VISUAL PERFORMANCE & KEY POINTS',true);
  if(metrics.length){
    try{
      slide.addChart(pptx.ChartType.doughnut,[{name:'Ayushman',labels:metrics.slice(0,6).map(m=>m.label),values:metrics.slice(0,6).map(m=>Math.max(0,m.value))}],{
        x:.55,y:1.55,w:5.25,h:4.7,holeSize:58,showLegend:true,legendPos:'b',
        legendFontFace:'Aptos',legendFontSize:10,chartColors:['0F766E','0EA5E9','F59E0B','8B5CF6','14B8A6','64748B'],
        showTitle:false,showValue:false,showCategoryName:false
      });
    }catch(e){}
  }
  slide.addText('KEY POINTS',{x:6.15,y:1.55,w:5.8,h:.4,fontSize:22,bold:true,color:'0F766E',margin:0});
  const facts=analyse(mod);
  slide.addText(facts.slice(0,5).map(x=>'• '+x).join('\\n'),{x:6.2,y:2.05,w:6.25,h:2.5,fontSize:16,color:'172033',bold:true,breakLine:false,margin:.03,fit:'shrink'});
  slide.addText('INDICATOR VALUES',{x:6.15,y:4.75,w:5.8,h:.35,fontSize:17,bold:true,color:'0F766E',margin:0});
  metrics.slice(0,6).forEach((m,i)=>{
    const y=5.08+i*.31;
    slide.addText(m.label,{x:6.15,y,w:4.45,h:.23,fontSize:9,bold:true,color:'0B3558',fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:.7},margin:.04,fit:'shrink'});
    slide.addText(m.raw,{x:10.6,y,w:1.8,h:.23,fontSize:10,bold:true,color:'0F766E',fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:.7},align:'center',margin:.03,fit:'shrink'});
  });
  slide.addText('Source: live Google Sheet data rendered in the dashboard.',{x:.55,y:7.22,w:8,h:.2,fontSize:7,color:'94A3B8',margin:0});
}

function addGraphSlide(pptx,mod){
  const metrics=extractMetrics(mod);
  if(!metrics.length) return;
  const slide=pptx.addSlide();
  slide.background={color:'F4F7FB'};
  addTitle(slide,mod.name,'VISUAL PERFORMANCE SUMMARY',true);
  const chartData=[{name:'Value',labels:metrics.map(m=>m.label),values:metrics.map(m=>m.value)}];
  try{
    slide.addChart(pptx.ChartType.bar,chartData,{
      x:.55,y:1.55,w:7.55,h:4.85,
      catAxisLabelFontFace:'Aptos',catAxisLabelFontSize:12,
      valAxisLabelFontFace:'Aptos',valAxisLabelFontSize:10,
      showLegend:false,showTitle:false,showValue:true,
      showCatName:false,showSerName:false,
      chartColors:['0F766E'],
      showValue:true,
      valGridLine:{color:'D6E3EC',pt:1},
      valAxisMinVal:0,
      showCatName:false,
      dataLabelPosition:'outEnd'
    });
  }catch(e){
    // If chart rendering is unavailable, keep the data visible as a large table.
    slide.addTable(metrics.map(m=>[{text:m.label,options:{fontSize:14,bold:true}},{text:m.raw,options:{fontSize:20,bold:true,align:'center'}}]),{
      x:.7,y:1.55,w:7.7,h:4.9,colW:[5.3,2.4],rowH:.55,
      border:{type:'solid',color:'CBD5E1',pt:1}
    });
  }
  slide.addText('KEY POINTS',{x:8.45,y:1.45,w:3.95,h:.38,fontSize:22,bold:true,color:'0F766E',margin:0});
  const facts=analyse(mod);
  slide.addText(facts.slice(0,5).map(x=>'• '+x).join('\n'),{
    x:8.45,y:1.95,w:3.95,h:3.65,fontSize:16,bold:true,color:'172033',
    breakLine:false,margin:.03,valign:'mid',fit:'shrink'
  });
  slide.addText('Values shown are taken from the dashboard at PPTX generation time.',{
    x:8.45,y:6.15,w:3.95,h:.45,fontSize:9,color:'64748B',margin:0
  });
}


function addTitle(slide,title,sub,hero){
  slide.addText(title,{x:.45,y:.25,w:12.0,h:.55,fontSize:hero?28:24,bold:true,color:'0B3558',margin:0});
  slide.addText(sub||'',{x:.48,y:.80,w:12.0,h:.28,fontSize:hero?11:10,bold:true,color:'0F766E',margin:0,charSpacing:1});
  slide.addText('',{x:.45,y:1.16,w:12.35,h:.02,line:{color:'14B8A6',pt:1.6},margin:0});
}

function addTableSlide(pptx,mod,table,idx){
  const slide=pptx.addSlide();
  slide.background={color:'F4F7FB'};
  addTitle(slide,mod.name,mod.heading||mod.title,true);
  const rows=table.rows.slice(0,24);
  const cols=Math.max(...rows.map(r=>r.length),1);
  const width=12.2/cols;
  const pptRows=rows.map((r,ri)=>r.map(v=>({text:v,options:{bold:ri===0||/^total$|^कुल$|^योग$/i.test(text(v)),fontSize:Math.max(6,Math.min(10,10-cols*.35)),color:ri===0?'FFFFFF':'172033',fill:ri===0?'075985':(/^total$|^कुल$|^योग$/i.test(text(v))?'DBEAFE':'FFFFFF'),align:'center',valign:'mid',margin:2}})));
  slide.addTable(pptRows,{x:.4,y:1.32,w:12.35,h:5.25,border:{type:'solid',color:'CBD5E1',pt:1},autoFit:false,colW:Array(cols).fill(width),rowH:.22,margin:2});
  const facts=analyse({tables:[table]});
  slide.addText(facts.map(x=>'• '+x).join('\n'),{x:.55,y:6.72,w:11.9,h:.62,fontSize:9,color:'334155',breakLine:false,margin:0.03});
  slide.addText('Source: Google Sheet · Kharsia Health Dashboard',{x:.55,y:7.25,w:7,h:.2,fontSize:7,color:'94A3B8',margin:0});
  return slide;
}

function addModuleAnalysis(pptx,mod){
  const slide=pptx.addSlide(); slide.background={color:'F4F7FB'};
  addTitle(slide,mod.name,'KEY FINDINGS & INDICATORS',true);
  const facts=analyse(mod);
  slide.addText('KEY OBSERVATIONS',{x:.55,y:1.48,w:4,h:.35,fontSize:18,bold:true,color:'0F766E',margin:0});
  slide.addText(facts.map(x=>'• '+x).join('\n'),{x:.65,y:1.95,w:5.8,h:2.2,fontSize:14,color:'172033',breakLine:false,margin:.02});
  const rows=mod.tables.flatMap(t=>t.rows);
  const header=rows[0]||[];
  const sample=rows.find(r=>r.some(v=>/^total$|^कुल$|^योग$/i.test(text(v))))||rows[1]||[];
  const metrics=[];
  header.forEach((h,i)=>{
    const n=num(sample[i]);
    if(Number.isFinite(n)) metrics.push([text(h)||('Indicator '+(i+1)),sample[i]]);
  });
  slide.addText('BLOCK INDICATORS',{x:6.55,y:1.48,w:4.8,h:.35,fontSize:18,bold:true,color:'0F766E',margin:0});
  if(metrics.length){
    const cards=metrics.slice(0,8);
    cards.forEach((m,i)=>{
      const col=i%2,row=Math.floor(i/2);
      const x=6.55+col*2.85,y=1.95+row*1.05;
      slide.addText(m[0]+'\n'+text(m[1]),{x,y,w:2.55,h:.78,fontSize:11,color:'172033',bold:true,fill:{color:'FFFFFF'},line:{color:'D6E3EC',pt:1.2},margin:.12,valign:'mid',breakLine:false});
    });
  } else slide.addText('No numeric indicator was detected in the rendered table.',{x:6.7,y:2,w:5,h:1,fontSize:12,color:'64748B'});
  slide.addText('Source: live Google Sheet data rendered in the dashboard.',{x:.55,y:7.25,w:11.5,h:.2,fontSize:7,color:'94A3B8',margin:0});
}

async function generate(){
  const dashboardBtn=document.getElementById('dashboardPptxBtn');
  const setBtn=(msg,busy)=>{ if(dashboardBtn){ dashboardBtn.disabled=!!busy; dashboardBtn.textContent=msg; dashboardBtn.style.opacity=busy?'0.75':'1'; } };
  try{
    setBtn('⏳ PPTX तैयार हो रहा है...',true);
    if(typeof google==='undefined' || !google.visualization){
      await new Promise((resolve,reject)=>{
        if(typeof google!=='undefined' && google.charts){
          try{ google.charts.load('current',{packages:['corechart','table']}); google.charts.setOnLoadCallback(resolve); }
          catch(e){ reject(e); }
        } else reject(new Error('Google Sheets service उपलब्ध नहीं है'));
      });
    }
    const query=new google.visualization.Query(
      'https://docs.google.com/spreadsheets/d/'+
      encodeURIComponent(SPREADSHEET_ID)+
      '/gviz/tq?gid=1044088930&headers=0'
    );
    query.setQuery('select *');
    query.send(async response=>{
      try{
        if(response.isError()){
          alert('Janani Portal data load नहीं हुआ: '+response.getMessage());
          return;
        }
        currentRCHData=parseRCHData(response.getDataTable());
        if(!currentRCHData || !currentRCHData.facilityRows || !currentRCHData.facilityRows.length){
          alert('Janani Portal में कोई facility data नहीं मिला।');
          return;
        }
        await generateRCHPPTX();
        setBtn('✅ PPTX तैयार — Download करें',false);
        setTimeout(()=>setBtn('📊 Generate PPTX',false),5000);
      }catch(e){
        console.error('PPTX generation error',e);
        setBtn('❌ PPTX में error — फिर प्रयास करें',false);
        alert('PPTX download नहीं हुआ: '+(e&&e.message?e.message:e));
      }
    });
  }catch(e){
    console.error(e);
    setBtn('❌ PPTX में error — फिर प्रयास करें',false);
    alert('PPTX download नहीं हुआ: '+(e&&e.message?e.message:e));
  }
}
window.generate=generate;

/* ============================================================
   RCH 2.0 — DEDICATED PPTX EXPORT
   Cover + Index + Block Summary + Sector Graph +
   Sector-wise Facility Data + Analysis
============================================================ */

function rchText(v){ return String(v==null?'':v).replace(/\s+/g,' ').trim(); }
function rchNum(v){ const n=Number(String(v==null?'':v).replace(/,/g,'')); return Number.isFinite(n)?n:0; }
function rchPct(v){ return Math.round(rchNum(v))+'%'; }

function rchAnalysis(row){
  const p=rchNum(row.percent), b=rchNum(row.backlog), h=rchNum(row.highRisk);
  const a=[];
  if(p>=90) a.push('Registration achievement is 90% or above.');
  else if(p>=70) a.push('Registration achievement is in the 70–89% range.');
  else a.push('Registration achievement is below 70%; facility-level gaps should be reviewed.');
  if(b<0) a.push('Backlog is negative.');
  else if(b===0) a.push('Backlog is zero.');
  else a.push('Positive backlog is recorded and should be followed up.');
  if(h>0) a.push('High-risk cases are reported in this sector.');
  return a;
}

function rchHeader(slide,title,sub){
  slide.background={color:'F7FAFC'};
  slide.addShape(pptx.ShapeType.rect,{x:0,y:0,w:13.333,h:.55,fill:{color:'075985'},line:{color:'075985'}});
  slide.addText(title,{x:.45,y:.72,w:12.3,h:.42,fontSize:22,bold:true,color:'0F172A',margin:0});
  slide.addText(sub||'',{x:.45,y:1.15,w:12.2,h:.28,fontSize:10,color:'475569',margin:0});
  slide.addText('Kharsia Health Dashboard | Block Kharsia | District Raigarh | Chhattisgarh',{x:.45,y:7.14,w:11.8,h:.18,fontSize:7.5,color:'64748B',align:'right',margin:0});
}

function rchCard(slide,x,y,w,h,label,value){
  slide.addShape(pptx.ShapeType.roundRect,{x,y,w,h,fill:{color:'FFFFFF'},line:{color:'CBD5E1',pt:1}});
  slide.addText(label,{x:x+.1,y:y+.12,w:w-.2,h:.22,fontSize:8.5,bold:true,color:'64748B',align:'center',margin:0,fit:'shrink'});
  slide.addText(String(value),{x:x+.1,y:y+.43,w:w-.2,h:.4,fontSize:19,bold:true,color:'075985',align:'center',margin:0,fit:'shrink'});
}

async function generateRCHPPTX(){
  if(!window.PptxGenJS){
    try{ await loadPptxLib(); }catch(e){ alert('PowerPoint library load नहीं हुई। Internet connection check करें।'); return; }
  }
  if(typeof currentRCHData==='undefined' || !currentRCHData || !currentRCHData.facilityRows || !currentRCHData.facilityRows.length){
    alert('Janani Portal (RCH 2.0) का data अभी load नहीं हुआ है। पहले Janani Portal खोलकर data load होने दें।');
    return;
  }

  const pptx=new window.PptxGenJS();
  pptx.layout='LAYOUT_WIDE';
  pptx.author='Kharsia Health Dashboard';
  pptx.company='Kharsia Health Dashboard';
  pptx.subject='RCH 2.0 PW Registration Performance';
  pptx.title='RCH 2.0 — Kharsia';
  pptx.lang='en-IN';

  const parsed=currentRCHData;
  const sectors=buildSectorData(parsed);
  const total=calculateRCHTotals(sectors);
  const facilities=parsed.facilityRows;
  const date=parsed.asOnDate||'Current Date';

  // 1. COVER
  let slide=pptx.addSlide();
  slide.background={color:'075985'};
  slide.addShape(pptx.ShapeType.rect,{x:0,y:0,w:13.333,h:7.5,fill:{color:'075985'},line:{color:'075985'}});
  slide.addText('KHARSIA HEALTH DASHBOARD',{x:.7,y:1.0,w:11.9,h:.35,fontSize:14,bold:true,color:'BAE6FD',align:'center',charSpacing:2,margin:0});
  slide.addText('RCH 2.0',{x:.7,y:1.75,w:11.9,h:.75,fontSize:40,bold:true,color:'FFFFFF',align:'center',margin:0});
  slide.addText('PW Registration Performance Report',{x:.7,y:2.6,w:11.9,h:.42,fontSize:21,color:'E0F2FE',align:'center',margin:0});
  slide.addText('FY 2026–27',{x:.7,y:3.2,w:11.9,h:.35,fontSize:16,bold:true,color:'FFFFFF',align:'center',margin:0});
  slide.addShape(pptx.ShapeType.roundRect,{x:4.05,y:4.15,w:5.25,h:1.05,fill:{color:'FFFFFF',transparency:8},line:{color:'FFFFFF',transparency:65,pt:1}});
  slide.addText('As On Date',{x:4.2,y:4.36,w:4.95,h:.22,fontSize:10,bold:true,color:'DBEAFE',align:'center',margin:0});
  slide.addText(date,{x:4.2,y:4.64,w:4.95,h:.36,fontSize:18,bold:true,color:'FFFFFF',align:'center',margin:0});
  slide.addText('Block Kharsia | District Raigarh | Chhattisgarh',{x:.7,y:6.55,w:11.9,h:.28,fontSize:12,color:'BAE6FD',align:'center',margin:0});

  // 2. INDEX
  slide=pptx.addSlide();
  rchHeader(slide,'INDEX','RCH 2.0 report structure');
  [
    ['01','Cover Page'],
    ['02','Index'],
    ['03','Block Summary Dashboard'],
    ['04','Sector Wise Data + Graph'],
    ['05 onward','Sector-wise Facility Data + Analysis'],
    ['Last','Overall Data Analysis']
  ].forEach((it,i)=>{
    const y=1.55+i*.78;
    slide.addShape(pptx.ShapeType.roundRect,{x:1,y,w:.95,h:.45,fill:{color:'0F766E'},line:{color:'0F766E'}});
    slide.addText(it[0],{x:1,y:y+.1,w:.95,h:.2,fontSize:9,bold:true,color:'FFFFFF',align:'center',margin:0});
    slide.addText(it[1],{x:2.2,y:y+.05,w:9.7,h:.28,fontSize:15,bold:i===2,color:'0F172A',margin:0});
  });
  slide.addText('All values are taken from the RCH 2.0 data currently loaded by the website.',{x:1,y:6.35,w:10.8,h:.3,fontSize:9,color:'64748B',margin:0});

  // 3. BLOCK SUMMARY
  slide=pptx.addSlide();
  rchHeader(slide,'Janani Portal (RCH 2.0) — Block Summary Dashboard',`PW Registration Detail FY 2026–27 | As On Date : ${date}`);
  [
    ['HMIS PW Registration',total.hmis],
    ['RCH 2.0 PW Registration',total.rch],
    ['Overall %',rchPct(total.percent)],
    ['Temporary Registration',total.temp],
    ['Backlog',total.backlog],
    ['High Risk',total.highRisk]
  ].forEach((c,i)=>rchCard(slide,.55+(i%3)*4.18,1.62+Math.floor(i/3)*1.22,3.82,.96,c[0],c[1]));
  const obs=[
    total.percent>=90?'Overall registration achievement is 90% or above.':total.percent>=70?'Overall registration achievement is in the 70–89% range.':'Overall registration achievement is below 70%.',
    `${sectors.length} sectors and ${facilities.length} facilities are included in the loaded RCH 2.0 data.`,
    total.backlog<0?'Overall backlog is negative.':total.backlog===0?'Overall backlog is zero.':'Overall backlog is positive.',
    total.highRisk>0?`${total.highRisk} high-risk cases are recorded in the loaded data.`:'No high-risk count is recorded in the loaded data.'
  ];
  slide.addText('KEY OBSERVATIONS',{x:.65,y:4.35,w:2.8,h:.25,fontSize:12,bold:true,color:'075985',margin:0});
  obs.forEach((t,i)=>slide.addText('• '+t,{x:.8,y:4.78+i*.45,w:11.6,h:.3,fontSize:12,color:'334155',margin:0}));

  // 4. SECTOR DATA + GRAPH
  slide=pptx.addSlide();
  rchHeader(slide,'Janani Portal (RCH 2.0) — Sector Wise Data with Graph',`As On Date : ${date}`);
  try{
    slide.addChart(pptx.ChartType.bar,[{name:'RCH 2.0 %',labels:sectors.map(r=>rchText(r.sector)),values:sectors.map(r=>rchNum(r.percent))}],{
      x:.45,y:1.5,w:5.05,h:5.0,showLegend:false,showValue:true,showTitle:false,
      catAxisLabelFontSize:8,valAxisLabelFontSize:8,valAxisMinVal:0,valAxisMaxVal:100,valAxisMajorUnit:20,
      chartColors:['0F766E'],dataLabelPosition:'outEnd'
    });
  }catch(e){}
  const sectorRows=[
    ['Sector','Facilities','HMIS PW','RCH PW','%','Backlog'],
    ...sectors.map(r=>[rchText(r.sector),r.facility,r.hmis,r.rch,rchPct(r.percent),r.backlog]),
    ['Total',total.facility,total.hmis,total.rch,rchPct(total.percent),total.backlog]
  ];
  slide.addTable(sectorRows.map((r,ri)=>r.map(v=>({text:String(v),options:{fontSize:8,bold:ri===0||ri===sectorRows.length-1,color:ri===0?'FFFFFF':'334155',fill:ri===0?'075985':ri===sectorRows.length-1?'DBEAFE':'FFFFFF',align:'center',margin:.04}})),{
    x:5.75,y:1.5,w:7.05,h:5.0,border:{type:'solid',color:'CBD5E1',pt:.7},colW:[1.8,.8,1.0,1.0,.65,.85],rowH:.33
  });

  // 5+. FACILITY-WISE BY SECTOR
  [...new Set(facilities.map(r=>r.sector))].forEach(sector=>{
    const fr=facilities.filter(r=>r.sector===sector);
    const sr=sectors.find(r=>r.sector===sector)||{facility:fr.length,hmis:0,rch:0,percent:0,temp:0,backlog:0,highRisk:0};
    for(let start=0;start<fr.length;start+=18){
      const chunk=fr.slice(start,start+18);
      slide=pptx.addSlide();
      rchHeader(slide,`Janani Portal (RCH 2.0) — ${sector} — Facility Wise`,`${start+1}–${start+chunk.length} of ${fr.length} facilities | As On Date : ${date}`);
      const rows=[
        ['Sn','Facility','HMIS PW','RCH PW','%','Temporary','Backlog','High Risk'],
        ...chunk.map((r,i)=>[r.sn||start+i+1,rchText(r.facility),r.hmis,r.rch,rchPct(r.hmis>0?r.rch/r.hmis*100:0),r.temp,r.backlog,r.highRisk])
      ];
      slide.addTable(rows.map((r,ri)=>r.map(v=>({text:String(v),options:{fontSize:7.3,bold:ri===0,color:ri===0?'FFFFFF':'334155',fill:ri===0?'0F766E':'FFFFFF',align:'center',margin:.035}})),{
        x:.45,y:1.5,w:12.4,h:4.2,border:{type:'solid',color:'CBD5E1',pt:.7},colW:[.45,3.25,1,1,.7,.95,.85,.9],rowH:.23
      });
      const high=fr.reduce((a,b)=>a.percent>b.percent?a:b,fr[0]);
      const low=fr.reduce((a,b)=>a.percent<b.percent?a:b,fr[0]);
      const notes=[
        `Sector achievement: ${rchPct(sr.percent)} | Facilities: ${sr.facility} | Sector backlog: ${sr.backlog}.`,
        `Highest facility achievement: ${rchText(high.facility)} (${rchPct(high.percent)}).`,
        `Lowest facility achievement: ${rchText(low.facility)} (${rchPct(low.percent)}).`,
        ...rchAnalysis(sr)
      ];
      slide.addText('SECTOR ANALYSIS',{x:.55,y:5.95,w:2.1,h:.24,fontSize:11,bold:true,color:'075985',margin:0});
      notes.slice(0,5).forEach((t,i)=>slide.addText('• '+t,{x:.7,y:6.25+i*.25,w:12,h:.2,fontSize:8.5,color:'334155',margin:0}));
    }
  });

  // FINAL ANALYSIS
  slide=pptx.addSlide();
  rchHeader(slide,'Janani Portal (RCH 2.0) — Overall Data Analysis',`Block Kharsia | As On Date : ${date}`);
  const maxS=sectors.reduce((a,b)=>a.percent>b.percent?a:b,sectors[0]);
  const minS=sectors.reduce((a,b)=>a.percent<b.percent?a:b,sectors[0]);
  [
    `Overall PW registration achievement: ${rchPct(total.percent)}.`,
    `Highest sector achievement in the loaded data: ${rchText(maxS?.sector)} (${rchPct(maxS?.percent)}).`,
    `Lowest sector achievement in the loaded data: ${rchText(minS?.sector)} (${rchPct(minS?.percent)}).`,
    `Total facilities reported: ${facilities.length}; temporary registrations: ${total.temp}.`,
    total.backlog<0?'Overall backlog is negative.':total.backlog===0?'Overall backlog is zero.':'Overall backlog is positive.',
    total.highRisk>0?`High-risk count in the loaded data: ${total.highRisk}.`:'No high-risk count is recorded in the loaded data.'
  ].forEach((t,i)=>slide.addText('• '+t,{x:.85,y:1.6+i*.68,w:11.5,h:.4,fontSize:14,color:'334155',margin:0}));
  slide.addText('Note: This analysis is descriptive and calculated from the data currently loaded by the website.',{x:.85,y:6.25,w:11.5,h:.3,fontSize:9,italic:true,color:'64748B',margin:0});

  const fileName=`Kharsia Health Progressive Report_${String(date).replace(/[\\/]/g,'-')}.pptx`;
  const blob=await pptx.write({outputType:'blob'});
  if(!(blob instanceof Blob) || blob.size<1000) throw new Error('PPTX file generate नहीं हुई या खाली है।');

  // Browser download fix:
  // The Google Sheet callback is asynchronous, so an automatic a.click()
  // can lose the original user-gesture and be blocked by the browser.
  // Keep a real download control visible so the user can click it directly.
  const url=URL.createObjectURL(blob);
  const host=document.getElementById('dashboardPptxBtn')?.parentElement || document.body;
  const old=document.getElementById('pptxDownloadFallback'); if(old) old.remove();

  const box=document.createElement('div');
  box.id='pptxDownloadFallback';
  box.style.cssText='margin:12px 0 0;display:flex;align-items:center;gap:10px;flex-wrap:wrap;';

  const a=document.createElement('a');
  a.href=url;
  a.download=fileName;
  a.textContent='⬇️ Download Kharsia Health Progressive Report';
  a.style.cssText='display:inline-block;padding:11px 18px;border-radius:8px;background:#15803d;color:#fff;font-weight:800;text-decoration:none;cursor:pointer;box-shadow:0 2px 6px rgba(0,0,0,.12);';

  const info=document.createElement('span');
  info.textContent='PPTX तैयार है — Download पर क्लिक करें';
  info.style.cssText='font-size:13px;font-weight:700;color:#15803d;';

  a.addEventListener('click',()=>setTimeout(()=>URL.revokeObjectURL(url),120000),{once:true});
  box.appendChild(a);
  box.appendChild(info);
  host.appendChild(box);
}

function mountRCHPPTXButton(){
  if(document.getElementById('rchPptxGenerateBtn')) return;
  const b=document.createElement('button');
  b.id='rchPptxGenerateBtn';
  b.textContent='📽️ Janani Portal PPTX';
  Object.assign(b.style,{position:'fixed',right:'22px',top:'64px',zIndex:99999,border:0,borderRadius:'9px',padding:'10px 14px',background:'#7c3aed',color:'#fff',fontWeight:'800',fontSize:'13px',cursor:'pointer',boxShadow:'0 4px 12px rgba(0,0,0,.2)',display:'none'});
  b.onclick=()=>{generateRCHPPTX().catch(e=>{console.error(e);alert('PPTX बनाने में समस्या हुई: '+(e?.message||e));});};
  document.body.appendChild(b);
  const update=()=>{
    const title=rchText(document.getElementById('reportTitle')?.innerText);
    const active=false;
    b.style.display=active?'block':'none';
  };
  update();
  setInterval(update,1200);
}

function mount(){
  if(document.getElementById('pptxGenerateBtn'))return;
  const btn=document.createElement('button');
  btn.id='pptxGenerateBtn';
  btn.textContent='📊 Generate PPTX';
  Object.assign(btn.style,{display:'none',position:'fixed',right:'22px',top:'18px',zIndex:99999,border:0,borderRadius:'9px',padding:'11px 15px',background:'#b91c1c',color:'#fff',fontWeight:'800',fontSize:'13px',cursor:'pointer',boxShadow:'0 4px 12px rgba(0,0,0,.2)'});
  btn.onclick=()=>generate().catch(e=>{console.error(e);alert('PPTX बनाने में समस्या हुई: '+(e&&e.message?e.message:e));});
  document.body.appendChild(btn);
  const p=document.createElement('div');
  p.id='pptxProgress';
  p.textContent='';
  Object.assign(p.style,{display:'none',position:'fixed',right:'22px',top:'64px',zIndex:99999,background:'#fff',border:'1px solid #cbd5e1',borderRadius:'8px',padding:'9px 12px',fontSize:'12px',color:'#334155',boxShadow:'0 4px 12px rgba(0,0,0,.15)',maxWidth:'380px'});
  document.body.appendChild(p);
  const sync=()=>{
    const dashboard=document.getElementById('dashboardPage');
    btn.style.display=dashboard&&dashboard.classList.contains('active')?'block':'none';
  };
  sync();
  setInterval(sync,500);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',()=>{mount();});else{mount();}
})();