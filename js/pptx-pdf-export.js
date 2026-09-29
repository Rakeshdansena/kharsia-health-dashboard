/* Kharsia unified PPTX -> PDF exporter: both formats use the same generated PPTX bytes */
(function(){
'use strict';
function wait(ms){return new Promise(function(r){setTimeout(r,ms);});}

async function loadViewer(){
  if(window.__kharsiaPptxViewerModule) return window.__kharsiaPptxViewerModule;
  var urls=[
    'https://esm.sh/pptx-vanilla-viewer@3.9.0?bundle',
    'https://cdn.jsdelivr.net/npm/pptx-vanilla-viewer@3.9.0/+esm'
  ];
  var lastErr=null;
  for(var i=0;i<urls.length;i++){
    try{
      var mod=await Promise.race([
        import(urls[i]),
        new Promise(function(_,reject){setTimeout(function(){reject(new Error('Viewer CDN timeout'));},30000);})
      ]);
      if(mod && typeof mod.createPptxViewer==='function'){
        window.__kharsiaPptxViewerModule=mod;
        return mod;
      }
    }catch(e){lastErr=e;}
  }
  throw new Error('PDF renderer load नहीं हुआ। Internet connection/check करें। '+(lastErr&&lastErr.message?lastErr.message:''));
}

async function generatePdfFromExactPptx(selected){
  var button=document.getElementById('dashboardPptxBtn');
  var pdfButton=document.getElementById('dashboardPdfBtn');
  if(button) button.disabled=true;
  if(pdfButton){pdfButton.disabled=true;pdfButton.textContent='⏳ Report तैयार हो रही है...';}

  try{
    window.__pptxSelectedPrograms=(selected||[]).slice();
    window.__lastGeneratedPptxBlob=null;
    window.__generatingPDF=true;

    if(typeof window.generatePPTX!=='function') throw new Error('PPTX Generator load नहीं हुआ।');
    await window.generatePPTX();

    var pptxBlob=window.__lastGeneratedPptxBlob;
    if(!pptxBlob) throw new Error('PPTX तैयार नहीं हुई।');

    if(pdfButton) pdfButton.textContent='⏳ PDF renderer load हो रहा है...';
    if(typeof window.showProgress==='function') window.showProgress('PDF renderer load हो रहा है',100,'PPTX तैयार है; अब PDF renderer शुरू किया जा रहा है...');

    var mod=await loadViewer();
    var host=document.createElement('div');
    host.id='kharsiaPptxPdfRenderHost';
    host.style.cssText='position:fixed;left:-20000px;top:0;width:1600px;height:900px;overflow:hidden;background:#fff;z-index:999999;visibility:visible;';
    document.body.appendChild(host);

    var viewer=null;
    try{
      viewer=mod.createPptxViewer(host,{
        showToolbar:false,
        showThumbnails:false,
        showFormatToolbar:false,
        showInspector:false,
        editable:false,
        fileName:'Kharsia Health Progressive Report'
      });
      if(pdfButton) pdfButton.textContent='⏳ PPTX slides render हो रही हैं...';
      if(typeof window.showProgress==='function') window.showProgress('PPTX slides PDF के लिए render हो रही हैं',100,'PPTX viewer में slides load हो रही हैं...');
      await Promise.race([
        viewer.loadFile(await pptxBlob.arrayBuffer()),
        new Promise(function(_,reject){setTimeout(function(){reject(new Error('PPTX viewer load में 60 सेकंड से अधिक लग गया।'));},60000);})
      ]);
      await wait(1200);
      if(typeof viewer.exportPdf!=='function') throw new Error('इस browser में PPTX PDF export उपलब्ध नहीं है।');
      await viewer.exportPdf({
        onProgress:function(current,count){
          var done=Math.min(99,Math.round((current/count)*100));
          if(pdfButton) pdfButton.textContent='⏳ PDF Slide '+Math.min(current+1,count)+'/'+count;
          if(typeof window.showProgress==='function') window.showProgress('PDF Slide '+Math.min(current+1,count)+'/'+count,done,'हर slide उसी PPTX layout से PDF में बदली जा रही है...');
        }
      });
      if(typeof window.showProgress==='function') window.showProgress('PDF तैयार है',100,'PDF download शुरू हो रहा है...');
    }finally{
      try{if(viewer && viewer.destroy) viewer.destroy();}catch(e){}
      host.remove();
    }

    if(pdfButton) pdfButton.textContent='📄 Generate Report';
    alert('Report तैयार है। PDF उसी generated PPTX के slide layout और order से बनी है।');
  }catch(e){
    console.error('Unified PDF export error',e);
    alert('PDF नहीं बनी: '+(e&&e.message?e.message:e));
  }finally{
    window.__generatingPDF=false;
    if(button) button.disabled=false;
    if(pdfButton){pdfButton.disabled=false;pdfButton.textContent='📄 Generate Report';}
  }
}

function openUnifiedReportSelector(){
  var old=document.getElementById('unifiedReportSelector');
  if(old) old.remove();

  var programmes=[
    {key:'Janani Portal',icon:'👩‍🍼',label:'Janani Portal (RCH 2.0)'},
    {key:'NCD',icon:'❤️',label:'NCD'},
    {key:'JAS Meeting',icon:'🤝',label:'JAS Meeting'},
    {key:'Ayushman Shivir',icon:'🏕️',label:'Ayushman Shivir'},
    {key:'Wellness Activity',icon:'🩺',label:'Wellness Activity'},
    {key:'RBSK',icon:'👶',label:'RBSK — राष्ट्रीय बाल स्वास्थ्य कार्यक्रम'},
    {key:'Telemedicine Report',icon:'🩻',label:'Telemedicine Report'}
  ];
  var order=programmes.map(function(p){return p.key;});

  var modal=document.createElement('div');
  modal.id='unifiedReportSelector';
  modal.style.cssText='position:fixed;inset:0;z-index:2147483647;background:rgba(2,18,35,.72);display:flex;align-items:center;justify-content:center;padding:20px;box-sizing:border-box;font-family:Arial,sans-serif;';
  modal.innerHTML='<div style="width:min(760px,96vw);max-height:90vh;overflow:auto;background:#fff;border-radius:18px;box-shadow:0 25px 70px rgba(0,0,0,.35);padding:22px;box-sizing:border-box;">'+
    '<div style="display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #dbe5ef;padding-bottom:14px;">'+
      '<div><div style="font-size:25px;font-weight:900;color:#073b75;">📄 Report Generation</div><div style="font-size:14px;color:#475569;margin-top:5px;">Programme select करें और checkbox को जिस क्रम में click करेंगे, उसी क्रम में slides बनेंगी।</div></div>'+
      '<button id="unifiedClose" style="border:0;background:#fee2e2;color:#991b1b;border-radius:9px;padding:8px 12px;font-size:18px;font-weight:900;cursor:pointer;">✕</button>'+
    '</div>'+
    '<div style="display:flex;gap:8px;align-items:center;margin:15px 0;"><button id="unifiedAll" style="padding:8px 12px;border:1px solid #94a3b8;border-radius:8px;background:#f8fafc;cursor:pointer;">✓ Select All</button><button id="unifiedNone" style="padding:8px 12px;border:1px solid #94a3b8;border-radius:8px;background:#f8fafc;cursor:pointer;">Clear All</button><b id="unifiedCount" style="margin-left:auto;color:#075985;"></b></div>'+
    '<div id="unifiedList" style="display:grid;gap:9px;"></div>'+
    '<div style="display:flex;justify-content:flex-end;gap:10px;margin-top:18px;padding-top:15px;border-top:1px solid #dbe5ef;">'+
      '<button id="unifiedCancel" style="padding:11px 18px;border:1px solid #94a3b8;border-radius:9px;background:#fff;font-weight:800;cursor:pointer;">Cancel</button>'+
      '<button id="unifiedPdf" style="padding:11px 18px;border:0;border-radius:9px;background:#0f766e;color:#fff;font-weight:900;cursor:pointer;">📄 Generate PDF</button>'+
      '<button id="unifiedPptx" style="padding:11px 18px;border:0;border-radius:9px;background:#b91c1c;color:#fff;font-weight:900;cursor:pointer;">📊 Generate PPTX</button>'+
    '</div></div>';
  document.body.appendChild(modal);

  var list=modal.querySelector('#unifiedList');
  programmes.forEach(function(p){
    var row=document.createElement('label');
    row.style.cssText='display:flex;align-items:center;gap:12px;padding:13px 14px;border:2px solid #dbe5ef;border-radius:11px;background:#f8fbff;cursor:pointer;';
    row.innerHTML='<input type="checkbox" checked value="'+p.key+'" style="width:19px;height:19px;"><span class="ord" style="min-width:27px;height:27px;border-radius:50%;background:#075985;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:900;"></span><span style="font-size:16px;font-weight:800;color:#0f2d52;">'+p.icon+' '+p.label+'</span>';
    list.appendChild(row);
    row.querySelector('input').addEventListener('change',function(){
      if(this.checked){order=order.filter(function(k){return k!==p.key;});order.push(p.key);}
      else order=order.filter(function(k){return k!==p.key;});
      refresh();
    });
  });
  function refresh(){
    var checked=Array.from(list.querySelectorAll('input:checked')).map(function(x){return x.value;});
    order=order.filter(function(k){return checked.indexOf(k)>=0;});
    Array.from(list.querySelectorAll('input')).forEach(function(x){if(x.checked&&order.indexOf(x.value)<0)order.push(x.value);});
    Array.from(list.children).forEach(function(row){
      var x=row.querySelector('input'), o=row.querySelector('.ord');
      o.textContent=x.checked?String(order.indexOf(x.value)+1):'×';
      row.style.opacity=x.checked?'1':'.55';
    });
    modal.querySelector('#unifiedCount').textContent=order.length+' programme selected';
  }
  function selected(){refresh();return order.slice();}
  function close(){modal.remove();}
  modal.querySelector('#unifiedClose').onclick=close;
  modal.querySelector('#unifiedCancel').onclick=close;
  modal.querySelector('#unifiedAll').onclick=function(){order=programmes.map(function(p){return p.key;});list.querySelectorAll('input').forEach(function(x){x.checked=true;});refresh();};
  modal.querySelector('#unifiedNone').onclick=function(){order=[];list.querySelectorAll('input').forEach(function(x){x.checked=false;});refresh();};
  modal.querySelector('#unifiedPptx').onclick=function(){
    var s=selected(); if(!s.length){alert('कम-से-कम एक Programme select करें।');return;}
    window.__pptxSelectedPrograms=s; close();
    if(typeof window.generatePPTX!=='function'){alert('PPTX Generator load नहीं हुआ।');return;}
    window.generatePPTX();
  };
  modal.querySelector('#unifiedPdf').onclick=function(){
    var s=selected(); if(!s.length){alert('कम-से-कम एक Programme select करें।');return;}
    window.__pptxSelectedPrograms=s; close();
    generatePdfFromExactPptx(s);
  };
  refresh();
}
window.openUnifiedReportSelector=openUnifiedReportSelector;

window.runSelectedProgrammePDF=generatePdfFromExactPptx;
})();