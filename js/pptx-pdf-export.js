/* Kharsia PPTX to PDF exporter */
(function(){
'use strict';
function wait(ms){return new Promise(function(r){setTimeout(r,ms);});}
function loadViewer(){
  if(window.__kharsiaPptxViewerModule) return Promise.resolve(window.__kharsiaPptxViewerModule);
  return import('https://esm.sh/pptx-vanilla-viewer@3.9.0?bundle').then(function(mod){
    if(!mod || typeof mod.createPptxViewer!=='function') throw new Error('PPTX PDF renderer load nahi hua.');
    window.__kharsiaPptxViewerModule=mod;
    return mod;
  });
}
async function generatePdfFromExactPptx(selected){
  var button=document.getElementById('dashboardPptxBtn');
  if(button){button.disabled=true;button.textContent='PDF तैयार हो रहा है...';}
  var originalCreateObjectURL=URL.createObjectURL;
  var pptxBlob=null;
  URL.createObjectURL=function(obj){
    try{
      if(obj instanceof Blob && (String(obj.type||'').indexOf('presentation')>=0 || obj.size>50000)) pptxBlob=obj;
    }catch(e){}
    return originalCreateObjectURL.call(URL,obj);
  };
  try{
    if(selected && selected.length) window.__pptxSelectedPrograms=selected.slice();
    if(typeof window.generatePPTX!=='function') throw new Error('PPTX Generator अभी load नहीं हुआ।');
    await window.generatePPTX();
    if(!pptxBlob) throw new Error('Generated PPTX bytes नहीं मिले।');
    var mod=await loadViewer();
    var host=document.createElement('div');
    host.style.cssText='position:fixed;left:-20000px;top:0;width:1600px;height:900px;overflow:hidden;opacity:0;pointer-events:none;z-index:-1;';
    document.body.appendChild(host);
    var viewer=null;
    try{
      viewer=mod.createPptxViewer(host,{showToolbar:false,showThumbnails:false,showFormatToolbar:false,showInspector:false,editable:false,fileName:'Kharsia Health Progressive Report'});
      await viewer.loadFile(pptxBlob);
      await wait(250);
      await viewer.exportPdf({onProgress:function(current,count){
        if(button) button.textContent='PDF: Slide '+(current+1)+'/'+count;
      }});
    }finally{
      try{if(viewer && viewer.destroy) viewer.destroy();}catch(e){}
      host.remove();
    }
    var old=document.getElementById('pptxDownloadFallback');
    if(old) old.remove();
    if(button){button.disabled=false;button.textContent='Generate PPTX';}
    alert('PDF तैयार है। यह उसी generated PPTX को PDF में export करके बनाया गया है।');
  }catch(e){
    console.error('PPTX PDF export error',e);
    if(button){button.disabled=false;button.textContent='Generate PPTX';}
    alert('PPTX से PDF नहीं बना: '+(e&&e.message?e.message:e));
  }finally{
    URL.createObjectURL=originalCreateObjectURL;
  }
}
window.runSelectedProgrammePDF=generatePdfFromExactPptx;
})();