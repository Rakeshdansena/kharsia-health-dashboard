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
    s.src='https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
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
  slide.addTable(metrics.slice(0,6).map(m=>[
    {text:m.label,options:{fontSize:11,bold:true,color:'0B3558',fill:'FFFFFF',margin:3}},
    {text:m.raw,options:{fontSize:14,bold:true,color:'0F766E',fill:'FFFFFF',align:'center',margin:3}}
  ]),{x:6.15,y:5.15,w:6.25,h:1.45,colW:[4.5,1.75],rowH:.24,border:{type:'solid',color:'CBD5E1',pt:1},margin:2});
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
  const btn=document.getElementById('pptxGenerateBtn');
  if(btn){btn.disabled=true;btn.textContent='⏳ PPTX बन रहा है...';}
  try{
    await loadPptxLib();
    const pptx=new window.PptxGenJS();
    window._kharsiaPptx=pptx;
    // Keep the active presentation available to helper functions.
    pptx.layout='LAYOUT_WIDE';
    pptx.author='Kharsia Health Dashboard';
    pptx.subject='Kharsia Health Programme Data Analysis';
    pptx.title='Kharsia Health Programme — Data Analysis';
    pptx.company='Kharsia Health Dashboard';
    pptx.lang='hi-IN';

    let progress=document.getElementById('pptxProgress');
    if(progress)progress.style.display='block';

    const modules=[];
    for(let i=0;i<MODULES.length;i++){
      if(progress)progress.textContent='📊 '+MODULES[i]+' data collect हो रहा है ('+(i+1)+'/'+MODULES.length+')';
      if(!clickModule(MODULES[i])) continue;
      await waitForModule(MODULES[i]);
      await wait(700);
      modules.push(collectModule(MODULES[i]));
    }

    let slide=pptx.addSlide(); slide.background={color:'075985'};
    slide.addText('KHARSIA HEALTH DASHBOARD',{x:.7,y:1.35,w:11.9,h:.7,fontSize:34,bold:true,color:'FFFFFF',align:'center',margin:0,charSpacing:1});
    slide.addText('HEALTH PROGRAMME PERFORMANCE REVIEW',{x:1.2,y:2.25,w:10.9,h:.5,fontSize:23,bold:true,color:'99F6E4',align:'center',margin:0});
    slide.addText('Block Kharsia  |  District Raigarh  |  Chhattisgarh',{x:1.2,y:3.0,w:10.9,h:.35,fontSize:15,color:'E2E8F0',align:'center',margin:0});
    slide.addText('DATA ANALYSIS PRESENTATION  •  '+new Date().toLocaleDateString('en-IN'),{x:3.4,y:6.45,w:5.2,h:.3,fontSize:10,bold:true,color:'CBD5E1',align:'center',margin:0});

    slide=pptx.addSlide(); slide.background={color:'F8FAFC'};
    addTitle(slide,'PROGRAMME COVERAGE','Modules loaded from the live dashboard',true);
    const summary=modules.map((m,i)=>[String(i+1),m.name,String(m.tables.length),m.heading||m.title]);
    slide.addTable([
      [{text:'S.No.',options:{bold:true,color:'FFFFFF',fill:'075985'}},{text:'Programme',options:{bold:true,color:'FFFFFF',fill:'075985'}},{text:'Tables',options:{bold:true,color:'FFFFFF',fill:'075985'}},{text:'Report heading / update',options:{bold:true,color:'FFFFFF',fill:'075985'}}],
      ...summary.map(r=>r.map((v,j)=>({text:v,options:{fontSize:9,fill:j===0?'DBEAFE':'FFFFFF',color:'172033',margin:2}})))
    ],{x:.45,y:1.35,w:12.2,h:5.6,colW:[.65,2.3,.8,8.45],border:{type:'solid',color:'CBD5E1',pt:1},rowH:.28});

    for(const mod of modules){
      if(mod.name==='Ayushman Card'){
        addAyushmanHeroSlide(pptx,mod);
        addAyushmanVisualSlide(pptx,mod);
        for(const table of mod.tables.slice(0,1)) addTableSlide(pptx,mod,table,0);
      }else{
        addGraphSlide(pptx,mod);
        addModuleAnalysis(pptx,mod);
        for(const table of mod.tables.slice(0,2)) addTableSlide(pptx,mod,table,0);
      }
    }

    slide=pptx.addSlide(); slide.background={color:'0B3558'};
    slide.addText('MEETING REVIEW & ACTION POINTS',{x:.7,y:.7,w:11.8,h:.5,fontSize:28,bold:true,color:'FFFFFF',align:'center',margin:0});
    slide.addText('• Module-wise figures and percentages are taken from the live dashboard at generation time.\n• Review low/zero percentage indicators, backlog values and programme-specific gaps from the corresponding module slides.\n• Use the dashboard Refresh button before generating the PPTX for the latest Google Sheet data.',{x:1.1,y:1.8,w:10.8,h:2.2,fontSize:16,color:'E2E8F0',breakLine:false,margin:.02});
    slide.addText('Generated automatically from Kharsia Health Dashboard',{x:1.1,y:6.55,w:10.8,h:.3,fontSize:10,color:'94A3B8',align:'center',margin:0});

    const file='Kharsia_Health_Data_Analysis_'+new Date().toISOString().slice(0,10)+'.pptx';
    await pptx.writeFile({fileName:file});
    if(progress){progress.textContent='✅ PPTX तैयार है — download शुरू हो गया।';setTimeout(()=>progress.style.display='none',5000);}
  }catch(e){
    console.error(e);
    alert('PPTX बनाने में समस्या: '+(e.message||e));
  }finally{
    if(btn){btn.disabled=false;btn.textContent='📊 Generate PPTX';}
  }
}

function mount(){
  if(document.getElementById('pptxGenerateBtn'))return;
  const btn=document.createElement('button');
  btn.id='pptxGenerateBtn'; btn.textContent='📊 Generate PPTX';
  Object.assign(btn.style,{position:'fixed',right:'22px',top:'18px',zIndex:99999,border:0,borderRadius:'9px',padding:'11px 15px',background:'#b91c1c',color:'#fff',fontWeight:'800',fontSize:'13px',cursor:'pointer',boxShadow:'0 4px 12px rgba(0,0,0,.2)'});
  btn.onclick=generate; document.body.appendChild(btn);
  const p=document.createElement('div');p.id='pptxProgress';p.textContent='';Object.assign(p.style,{display:'none',position:'fixed',right:'22px',top:'64px',zIndex:99999,background:'#fff',border:'1px solid #cbd5e1',borderRadius:'8px',padding:'9px 12px',fontSize:'12px',color:'#334155',boxShadow:'0 4px 12px rgba(0,0,0,.15)',maxWidth:'380px'});document.body.appendChild(p);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',mount);else mount();
})();