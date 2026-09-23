/* NCD SEARCH FIX — filters the custom NCD tables rendered by js/ncd.js */
(function(){
  'use strict';

  function bindNCDSearch(){
    const input = document.getElementById('searchBox');
    if(!input || input.dataset.ncdSearchBound === '1') return;
    input.dataset.ncdSearchBound = '1';

    function applySearch(){
      const value = String(input.value || '').toLowerCase().trim();
      const isNCD =
        window.currentReportIndex !== undefined &&
        window.REPORTS &&
        window.REPORTS[window.currentReportIndex] &&
        window.REPORTS[window.currentReportIndex].name === 'NCD';

      if(!isNCD) return;

      document.querySelectorAll('#ncdModuleContainer tbody tr').forEach(function(row){
        row.style.display =
          row.innerText.toLowerCase().includes(value) ? '' : 'none';
      });
    }

    input.addEventListener('input', applySearch);
    window.__applyNCDSearch = applySearch;
  }

  function run(){
    bindNCDSearch();
    if(window.__applyNCDSearch) window.__applyNCDSearch();
  }

  if(document.readyState === 'loading'){
    document.addEventListener('DOMContentLoaded', run);
  }else{
    run();
  }

  new MutationObserver(run).observe(document.documentElement,{childList:true,subtree:true});
})();