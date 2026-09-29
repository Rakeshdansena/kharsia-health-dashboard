/* Kharsia unified PPTX -> PDF exporter: both formats use the same generated PPTX bytes */
(function(){
'use strict';
function wait(ms){return new Promise(function(r){setTimeout(r,ms);});}

async function loadViewer(){
  if(window.__kharsiaPptxViewerModule) return window.__kharsiaPptxViewerModule;
  var urls=[
    'https://cdn.jsdelivr.net/npm/pptx-vanilla-viewer@3.9.0/+esm',
    'https://esm.sh/pptx-vanilla-viewer@3.9.0?bundle'
  ];
  var lastErr=null;
  for(var i=0;i<urls.length;i++){
    try{
      var mod=await import(urls[i]);
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

    if(pdfButton) pdfButton.textContent='⏳ PDF slides तैयार हो रही हैं...';

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
      await viewer.loadFile(await pptxBlob.arrayBuffer());
      await wait(1500);
      if(typeof viewer.exportPdf!=='function') throw new Error('इस browser में PPTX PDF export उपलब्ध नहीं है।');
      await viewer.exportPdf({
        onProgress:function(current,count){
          if(pdfButton) pdfButton.textContent='⏳ PDF Slide '+(current+1)+'/'+count;
        }
      });
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
window.runSelectedProgrammePDF=generatePdfFromExactPptx;
})();