/* Independent report selector fallback. Kept separate from the large inline report/PDF code. */
(function(){
  'use strict';
  function openSelector(){
    var existing=document.getElementById('pptxProgrammeModalFixed');
    if(existing) existing.remove();
    var programs=[
      ['Janani Portal','👩‍🍼','Janani Portal (RCH 2.0)'],
      ['NCD','❤️','NCD'],
      ['JAS Meeting','🤝','JAS Meeting'],
      ['Ayushman Shivir','🏕️','Ayushman Shivir'],
      ['Wellness Activity','🩺','Wellness Activity'],
      ['RBSK','👶','RBSK — राष्ट्रीय बाल स्वास्थ्य कार्यक्रम'],
      ['Dialysis','💧','Dialysis — प्रधानमंत्री राष्ट्रीय डायलिसिस कार्यक्रम'],
      ['Telemedicine Report','🩻','Telemedicine Report'],
      ['NRC Kharsia','🏥','NRC Kharsia'],
      ['Blindness Control','👁️','Blindness Control — Month Wise Report'],
      ['NQAS Certification','🏅','NQAS Certification']
    ];
    var modal=document.createElement('div');
    modal.id='pptxProgrammeModalFixed';
    modal.style.cssText='position:fixed;inset:0;z-index:100000;background:rgba(15,23,42,.62);display:flex;align-items:center;justify-content:center;padding:16px;font-family:Arial,sans-serif';
    var card=document.createElement('div');
    card.style.cssText='width:min(720px,96vw);max-height:90vh;overflow:auto;background:#fff;border-radius:16px;padding:22px;box-shadow:0 20px 60px #0005;color:#0f172a';
    card.innerHTML='<h2 style="margin:0 0 8px">📄 Report Generation</h2><p style="margin:0 0 16px;color:#475569">रिपोर्ट में शामिल करने वाले Programme चुनें।</p>';
    var grid=document.createElement('div');
    grid.style.cssText='display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:10px';
    programs.forEach(function(p){
      var label=document.createElement('label');
      label.style.cssText='display:flex;gap:9px;align-items:center;padding:12px;border:1px solid #cbd5e1;border-radius:10px;cursor:pointer;font-weight:700';
      var cb=document.createElement('input'); cb.type='checkbox'; cb.value=p[0]; cb.checked=true; cb.className='pptx-fixed-check';
      var span=document.createElement('span'); span.textContent=p[1]+' '+p[2];
      label.appendChild(cb);label.appendChild(span);grid.appendChild(label);
    });
    card.appendChild(grid);
    var actions=document.createElement('div');
    actions.style.cssText='display:flex;justify-content:flex-end;gap:10px;flex-wrap:wrap;margin-top:20px';
    actions.innerHTML='<button type="button" id="pptxFixedCancel" style="padding:10px 16px;border:1px solid #94a3b8;border-radius:8px;background:#fff;cursor:pointer">Cancel</button><button type="button" id="pptxFixedGeneratePDF" style="padding:10px 16px;border:0;border-radius:8px;background:#0f766e;color:#fff;font-weight:800;cursor:pointer">📄 Generate PDF</button><button type="button" id="pptxFixedGenerate" style="padding:10px 16px;border:0;border-radius:8px;background:#b91c1c;color:#fff;font-weight:800;cursor:pointer">📊 Generate PPTX</button>';
    card.appendChild(actions);modal.appendChild(card);document.body.appendChild(modal);
    document.getElementById('pptxFixedCancel').onclick=function(){modal.remove();};
    document.getElementById('pptxFixedGeneratePDF').onclick=function(){
      var selected=Array.from(modal.querySelectorAll('.pptx-fixed-check:checked')).map(function(c){return c.value;});
      if(!selected.length){alert('कम-से-कम एक Programme select करें।');return;}
      window.__pptxSelectedPrograms=selected.slice();
      modal.remove();
      if(typeof window.runSelectedProgrammePDF==='function'){
        window.runSelectedProgrammePDF(selected);
      }else{
        alert('PDF Generator load नहीं हुआ। पेज को Ctrl+Shift+R से refresh करके फिर प्रयास करें।');
      }
    };
    document.getElementById('pptxFixedGenerate').onclick=function(){
      var selected=Array.from(modal.querySelectorAll('.pptx-fixed-check:checked')).map(function(c){return c.value;});
      if(!selected.length){alert('कम-से-कम एक Programme select करें।');return;}
      window.__pptxSelectedPrograms=selected;
      modal.remove();
      var btn=document.getElementById('dashboardPptxBtn');
      if(btn){btn.disabled=true;btn.textContent='⏳ PPTX Loading...';}
      function run(){
        if(typeof window.generatePPTX==='function'){window.generatePPTX();return;}
        if(btn){btn.disabled=false;btn.textContent='📄 Generate Report';}
        alert('PPTX Generator load नहीं हुआ। पेज को Ctrl+Shift+R से refresh करके फिर प्रयास करें।');
      }
      if(typeof window.generatePPTX==='function'){run();return;}
      var script=document.getElementById('pptx-report-script');
      if(script && script.dataset.loaded==='1'){run();return;}
      if(!script || !script.src || !script.src.includes('pptx-report-final-23.js')){
        script=document.createElement('script');
        script.id='pptx-report-script-fixed';
        script.src='js/pptx-report-final-23.js?v=20261012-report-button-fix-10';
        script.onload=run;
        script.onerror=function(){if(btn){btn.disabled=false;btn.textContent='📄 Generate Report';}alert('PPTX Generator file load नहीं हुई।');};
        document.head.appendChild(script);
      } else {
        var attempts=0;var timer=setInterval(function(){attempts++;if(typeof window.generatePPTX==='function'){clearInterval(timer);run();}else if(attempts>30){clearInterval(timer);if(btn){btn.disabled=false;btn.textContent='📄 Generate Report';}alert('PPTX Generator load नहीं हुआ।');}},200);
      }
    };
  }
  window.showPPTXProgrammeSelectorFixed=openSelector;
  document.addEventListener('DOMContentLoaded',function(){
    var btn=document.getElementById('dashboardPptxBtn');
    if(btn){btn.onclick=function(e){e.preventDefault();openSelector();};}
  });
  // The button can exist before this file loads.
  var btn=document.getElementById('dashboardPptxBtn');
  if(btn) btn.onclick=function(e){e.preventDefault();openSelector();};
})();