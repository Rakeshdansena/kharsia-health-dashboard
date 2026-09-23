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
  for(let i=0;i<24;i++){
    await wait(350);
    const title=text(document.getElementById('reportTitle')?.textContent);
    const body=text(document.getElementById('pdfArea')?.innerText);
    if(title || body.length>40) return;
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
  const heading=text(document.querySelector('.shivir-page-heading .subheading')?.innerText) ||
                text(document.querySelector('#ncdModuleContainer .ncd-final-title')?.innerText) || title;
  const status=text(document.getElementById('status')?.innerText);
  const tables=collectTables();
  return {name,title,heading,status,tables};
}

function analyse(mod){
  const facts=[];
  const allRows=mod.tables.flatMap(t=>t.rows);
  const dataRows=allRows.slice(1);
  let pct=[];
  let nums=[];
  dataRows.forEach(r=>r.forEach(v=>{
    if(/%$/.test(v)){const n=num(v);if(Number.isFinite(n))pct.push(n);}
    else {const n=num(v);if(Number.isFinite(n))nums.push(n);}
  }));
  if(pct.length){
    const avg=pct.reduce((a,b)=>a+b,0)/pct.length;
    const max=Math.max(...pct),min=Math.min(...pct);
    facts.push('Reported percentage values: '+pct.length+'; average '+avg.toFixed(1)+'%.');
    facts.push('Observed range: '+min.toFixed(1)+'% to '+max.toFixed(1)+'%.');
  }
  const totals=allRows.filter(r=>r.some(v=>/^total$|^कुल$|^योग$/i.test(text(v))));
  if(totals.length) facts.push('Total row detected in the report for block-level aggregation.');
  if(nums.length) facts.push('Numeric indicators available across '+dataRows.length+' rendered data rows.');
  if(!facts.length) facts.push('Report loaded; analysis is based on the live rendered Google Sheet table.');
  return facts;
}

function addTitle(slide,title,sub){
  slide.addText(title,{x:.45,y:.28,w:12.4,h:.48,fontSize:24,bold:true,color:'075985',margin:0});
  slide.addText(sub||'',{x:.48,y:.78,w:12.1,h:.32,fontSize:10,color:'64748B',margin:0});
  slide.addShape(window.PptxGenJS.ShapeType.line,{x:.45,y:1.13,w:12.35,h:0,line:{color:'0F766E',width:1.2}});
}

function addTableSlide(pptx,mod,table,idx){
  const slide=pptx.addSlide();
  slide.background={color:'F8FAFC'};
  addTitle(slide,mod.name,mod.heading||mod.title);
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
  const slide=pptx.addSlide(); slide.background={color:'F8FAFC'};
  addTitle(slide,mod.name,'Data analysis from live rendered report');
  const facts=analyse(mod);
  slide.addText('Key observations',{x:.55,y:1.42,w:3,h:.35,fontSize:17,bold:true,color:'075985',margin:0});
  slide.addText(facts.map(x=>'• '+x).join('\n'),{x:.65,y:1.95,w:5.8,h:2.2,fontSize:14,color:'172033',breakLine:false,margin:.02});
  const rows=mod.tables.flatMap(t=>t.rows);
  const header=rows[0]||[];
  const sample=rows.find(r=>r.some(v=>/^total$|^कुल$|^योग$/i.test(text(v))))||rows[1]||[];
  const metrics=[];
  header.forEach((h,i)=>{
    const n=num(sample[i]);
    if(Number.isFinite(n)) metrics.push([text(h)||('Indicator '+(i+1)),sample[i]]);
  });
  slide.addText('Available numeric indicators',{x:6.65,y:1.42,w:4.8,h:.35,fontSize:17,bold:true,color:'075985',margin:0});
  if(metrics.length){
    const cards=metrics.slice(0,8);
    cards.forEach((m,i)=>{
      const col=i%2,row=Math.floor(i/2);
      const x=6.65+col*2.85,y=1.95+row*1.02;
      slide.addShape(window.PptxGenJS.ShapeType.roundRect,{x,y,w:2.55,h:.78,rectRadius:.08,fill:{color:'FFFFFF'},line:{color:'D6E3EC',pt:1}});
      slide.addText(m[0],{x:x+.12,y:y+.1,w:2.3,h:.24,fontSize:8,color:'64748B',bold:true,margin:0});
      slide.addText(text(m[1]),{x:x+.12,y:y+.36,w:2.3,h:.28,fontSize:16,color:'0F766E',bold:true,margin:0});
    });
  } else slide.addText('No numeric indicator was detected in the rendered table.',{x:6.7,y:2,w:5,h:1,fontSize:12,color:'64748B'});
  slide.addText('Analysis is descriptive only and uses the values currently visible on the dashboard.',{x:.55,y:7.25,w:11.5,h:.2,fontSize:7,color:'94A3B8',margin:0});
}

async function generate(){
  const btn=document.getElementById('pptxGenerateBtn');
  if(btn){btn.disabled=true;btn.textContent='⏳ PPTX बन रहा है...';}
  try{
    await loadPptxLib();
    const pptx=new window.PptxGenJS();
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
    slide.addText('🏥 Kharsia Health Dashboard',{x:.7,y:1.45,w:11.9,h:.7,fontSize:30,bold:true,color:'FFFFFF',align:'center',margin:0});
    slide.addText('Health Programme Data Analysis',{x:1.2,y:2.35,w:10.9,h:.5,fontSize:22,color:'E0F2FE',align:'center',margin:0});
    slide.addText('Block Kharsia · District Raigarh · Chhattisgarh',{x:1.2,y:3.05,w:10.9,h:.35,fontSize:13,color:'D1FAE5',align:'center',margin:0});
    slide.addText(new Date().toLocaleDateString('en-IN'),{x:4.5,y:6.5,w:4,h:.3,fontSize:10,color:'CBD5E1',align:'center',margin:0});

    slide=pptx.addSlide(); slide.background={color:'F8FAFC'};
    addTitle(slide,'Programme Coverage Summary','Modules loaded from the live dashboard');
    const summary=modules.map((m,i)=>[String(i+1),m.name,String(m.tables.length),m.heading||m.title]);
    slide.addTable([
      [{text:'S.No.',options:{bold:true,color:'FFFFFF',fill:'075985'}},{text:'Programme',options:{bold:true,color:'FFFFFF',fill:'075985'}},{text:'Tables',options:{bold:true,color:'FFFFFF',fill:'075985'}},{text:'Report heading / update',options:{bold:true,color:'FFFFFF',fill:'075985'}}],
      ...summary.map(r=>r.map((v,j)=>({text:v,options:{fontSize:9,fill:j===0?'DBEAFE':'FFFFFF',color:'172033',margin:2}})))
    ],{x:.45,y:1.35,w:12.2,h:5.6,colW:[.65,2.3,.8,8.45],border:{type:'solid',color:'CBD5E1',pt:1},rowH:.28});

    for(const mod of modules){
      addModuleAnalysis(pptx,mod);
      for(const table of mod.tables.slice(0,2)) addTableSlide(pptx,mod,table,0);
    }

    slide=pptx.addSlide(); slide.background={color:'0F2942'};
    slide.addText('Meeting Notes / Action Review',{x:.7,y:.7,w:11.8,h:.5,fontSize:25,bold:true,color:'FFFFFF',align:'center',margin:0});
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