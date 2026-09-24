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

  function loadReport(){
    return new Promise((resolve,reject)=>{
      if(typeof window.generate==='function') return resolve();
      const old=document.getElementById('pptx-report-runtime-loader');
      if(old){
        old.addEventListener('load',resolve,{once:true});
        old.addEventListener('error',()=>reject(new Error('PPTX report script load failed')),{once:true});
        return;
      }
      const s=document.createElement('script');
      s.id='pptx-report-runtime-loader';
      s.src='js/pptx-report.js?v=20260924-pptx-18';
      s.onload=()=>resolve();
      s.onerror=()=>reject(new Error('PPTX report script load failed'));
      document.head.appendChild(s);
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
      await loadReport();
      if(typeof window.generate!=='function') throw new Error('PPTX generator load नहीं हुआ।');
      setStatus('Google Sheet data पढ़ा जा रहा है…',false);
      await window.generate();
      setStatus('PPTX process पूरा हो गया। अगर नीचे Download button दिखे तो उस पर क्लिक करें।',true);
    }catch(e){
      console.error('PPTX start error',e);
      setStatus('❌ PPTX शुरू नहीं हुआ: '+(e?.message||e),false);
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