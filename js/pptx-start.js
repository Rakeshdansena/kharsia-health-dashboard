/* Kharsia Health Dashboard — PPTX bootstrap */
(function(){
  'use strict';

  const REPORT_URL='https://rakeshdansena.github.io/kharsia-health-dashboard/js/pptx-report.js?v=20260924-pptx-21';

  function setStatus(msg,ok){
    let el=document.getElementById('pptxStartStatus');
    if(!el){
      el=document.createElement('div');
      el.id='pptxStartStatus';
      const host=document.getElementById('dashboardPptxBtn')?.parentElement||document.body;
      host.appendChild(el);
    }
    el.textContent=msg;
    el.style.cssText='margin:10px 0;padding:10px 14px;border-radius:8px;font-size:13px;font-weight:700;background:'+(ok?'#dcfce7':'#fff7ed')+';color:'+(ok?'#166534':'#9a3412')+';border:1px solid '+(ok?'#86efac':'#fdba74')+';';
  }

  function loadGenerator(){
    return new Promise((resolve,reject)=>{
      if(typeof window.generate==='function') return resolve();

      const old=document.getElementById('pptx-report-script');
      if(old && old.dataset.pptxBootstrapTried==='1'){
        const timer=setInterval(()=>{
          if(typeof window.generate==='function'){
            clearInterval(timer); resolve();
          }
        },100);
        setTimeout(()=>{
          clearInterval(timer);
          if(typeof window.generate!=='function') reject(new Error('PPTX generator load नहीं हुआ।'));
        },12000);
        return;
      }

      const s=document.createElement('script');
      s.id='pptx-report-runtime-loader';
      s.src=REPORT_URL;
      s.async=false;
      s.dataset.pptxBootstrapTried='1';

      let done=false;
      const finish=()=>{if(done)return;done=true;clearTimeout(to);if(typeof window.generate==='function')resolve();else reject(new Error('PPTX generator script load हुआ, लेकिन generator function नहीं बना।'));};
      const fail=(msg)=>{if(done)return;done=true;clearTimeout(to);reject(new Error(msg));};
      const to=setTimeout(()=>fail('PPTX generator script load होने में समय लग रहा है।'),15000);

      s.onload=finish;
      s.onerror=()=>fail('js/pptx-report.js load नहीं हुआ।');
      document.head.appendChild(s);
    });
  }

  async function startPptxGeneration(){
    const btn=document.getElementById('dashboardPptxBtn');
    if(btn){btn.disabled=true;btn.textContent='⏳ PPTX तैयार हो रहा है...';}
    setStatus('PPTX generator load किया जा रहा है…',false);

    try{
      await loadGenerator();
      setStatus('Google Sheet data पढ़ा जा रहा है…',false);
      await window.generate();
      setStatus('✅ PPTX तैयार हो गया। नीचे Download button पर क्लिक करें।',true);
      if(btn){btn.disabled=false;btn.textContent='📊 Generate PPTX';}
    }catch(e){
      console.error('PPTX bootstrap error',e);
      setStatus('❌ '+(e?.message||e),false);
      if(btn){btn.disabled=false;btn.textContent='📊 Generate PPTX';}
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