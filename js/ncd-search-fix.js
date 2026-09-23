/* NCD SEARCH FIX — works with the actual NCD table rendered by index_updated.html */
(function(){
  'use strict';
  function bind(){
    const input=document.getElementById('searchBox');
    if(!input) return;
    if(input.dataset.ncdSearchBound!=='1'){
      input.dataset.ncdSearchBound='1';
      input.addEventListener('input',apply);
      input.addEventListener('keyup',apply);
    }
    apply();
  }
  function isNCD(){
    try{
      if(typeof window.currentReportIndex!=='undefined' && window.REPORTS && window.REPORTS[window.currentReportIndex]){
        return window.REPORTS[window.currentReportIndex].name==='NCD';
      }
    }catch(e){}
    const title=document.getElementById('reportTitle');
    return !!title && /NCD/i.test(title.textContent||'');
  }
  function apply(){
    if(!isNCD()) return;
    const input=document.getElementById('searchBox');
    if(!input) return;
    const q=String(input.value||'').toLowerCase().trim();
    document.querySelectorAll('#reportTable tbody tr').forEach(row=>{
      row.style.display=(!q || (row.innerText||'').toLowerCase().includes(q))?'':'none';
    });
  }
  window.__applyNCDSearch=apply;
  function run(){bind();}
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',run); else run();
  new MutationObserver(run).observe(document.documentElement,{childList:true,subtree:true});
  setInterval(run,1000);
})();