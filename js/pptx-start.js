/* Kharsia Health Dashboard — reliable PPTX start handler */
(function(){
  'use strict';

  function setStatus(msg, ok){
    let el=document.getElementById('pptxStartStatus');
    if(!el){
      el=document.createElement('div');
      el.id='pptxStartStatus';
      const host=document.getElementById('dashboardPptxBtn')?.parentElement || document.body;
      host.appendChild(el);
    }
    el.textContent=msg;
    el.style.cssText='margin:10px 0;padding:10px 14px;border-radius:8px;font-size:13px;font-weight:700;background:'+(ok?'#dcfce7':'#fff7ed')+';color:'+(ok?'#166534':'#9a3412')+';border:1px solid '+(ok?'#86efac':'#fdba74')+';';
  }

  function waitForGenerator(){
    return new Promise((resolve,reject)=>{
      if(typeof window.generate==='function') return resolve();

      const deadline=Date.now()+10000;
      const timer=setInterval(()=>{
        if(typeof window.generate==='function'){
          clearInterval(timer);
          resolve();
          return;
        }
        if(Date.now()>deadline){
          clearInterval(timer);
          reject(new Error('PPTX generator load नहीं हुआ।'));
        }
      },100);

      // Do NOT load pptx-report.js a second time. index.html/index_updated.html
      // already load it. A duplicate dynamic load could race with the first one.
      const report=document.getElementById('pptx-report-script');
      if(report){
        report.addEventListener('error',()=>{
          clearInterval(timer);
          reject(new Error('js/pptx-report.js load नहीं हुआ।'));
        },{once:true});
      }
    });
  }

  async function startPptxGeneration(){
    const btn=document.getElementById('dashboardPptxBtn');
    if(btn){
      btn.disabled=true;
      btn.textContent='⏳ PPTX तैयार हो रहा है...';
    }

    setStatus('PPTX process शुरू हो रहा है…',false);

    try{
      setStatus('PPTX generator तैयार हो रहा है…',false);
      await waitForGenerator();

      setStatus('Google Sheet data पढ़ा जा रहा है…',false);
      await window.generate();

      setStatus('✅ PPTX तैयार हो गया। नीचे Download button पर क्लिक करें।',true);
      if(btn){
        btn.disabled=false;
        btn.textContent='📊 Generate PPTX';
      }
    }catch(e){
      console.error('PPTX start error',e);
      setStatus('❌ '+(e?.message||e),false);
      if(btn){
        btn.disabled=false;
        btn.textContent='📊 Generate PPTX';
      }
      alert('PPTX शुरू नहीं हुआ: '+(e?.message||e));
    }
  }

  window.startPptxGeneration=startPptxGeneration;

  function bind(){
    const btn=document.getElementById('dashboardPptxBtn');
    if(!btn)return;
    btn.onclick=startPptxGeneration;
    btn.removeAttribute('onclick');
  }

  bind();
  document.addEventListener('DOMContentLoaded',bind);
  setInterval(bind,1000);
})();
