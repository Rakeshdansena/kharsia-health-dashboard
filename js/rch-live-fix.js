/* RCH 2.0 LIVE DATE FIX */
(function(){
'use strict';
const SID='1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0';
const GID='1044088930';
const FALLBACK='01-08-2026';
let lastDate='';

function clean(v){
  return String(v==null?'':v)
    .replace(/<br\s*\/?>(?=.)/gi,' ')
    .replace(/\s+/g,' ')
    .trim();
}
function extractDate(text){
  const s=clean(text);
  const patterns=[
    /as\s*on\s*date\s*[:\-]?\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})/i,
    /as\s*on\s*[:\-]?\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})/i,
    /(?:date|दिनांक)\s*[:\-]?\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})/i,
    /\b(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})\b/
  ];
  for(const p of patterns){
    const m=s.match(p);
    if(m) return m[1].replace(/\//g,'-');
  }
  return '';
}
function findDate(dt){
  if(!dt) return '';
  try{
    for(let c=0;c<dt.getNumberOfColumns();c++){
      const label=clean(dt.getColumnLabel(c));
      const d1=extractDate(label);
      if(d1) return d1;
      for(let r=0;r<dt.getNumberOfRows();r++){
        const d2=extractDate(clean(dt.getFormattedValue(r,c)));
        if(d2) return d2;
      }
    }
  }catch(e){}
  return '';
}
function applyDate(date){
  if(!date) return;
  lastDate=date;
  const text='RCH 2.0 PW Registration Detail FY 2026-27 | As On Date : '+date;
  document.querySelectorAll('.rch-page-date').forEach(el=>el.textContent=text);
}
function readSheet(){
  if(!window.google || !google.visualization || !google.visualization.Query) return;
  try{
    const url='https://docs.google.com/spreadsheets/d/'+SID+'/gviz/tq?gid='+encodeURIComponent(GID)+'&headers=0&tqx=out:html';
    const q=new google.visualization.Query(url);
    q.setQuery('select *');
    q.send(function(resp){
      try{
        if(resp && !resp.isError()){
          const d=findDate(resp.getDataTable());
          if(d) applyDate(d);
        }
      }catch(e){}
    });
  }catch(e){}
}
function patchRender(){
  if(typeof window.renderRCHTable!=='function' || window.renderRCHTable.__liveDateFix) return;
  const original=window.renderRCHTable;
  function patched(data){
    const result=original.apply(this,arguments);
    const d=findDate(data);
    applyDate(d || lastDate || FALLBACK);
    setTimeout(readSheet,100);
    return result;
  }
  patched.__liveDateFix=true;
  window.renderRCHTable=patched;
}
function boot(){
  patchRender();
  readSheet();
  if(lastDate) applyDate(lastDate);
}
[0,300,800,1500,3000,5000].forEach(t=>setTimeout(boot,t));
setInterval(boot,5000);
})();