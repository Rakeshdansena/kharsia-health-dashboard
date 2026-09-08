/* RCH 2.0 LIVE SHEET HEADING/DATE FIX
   Reads the RCH sheet title directly so changing the Google Sheet heading/date
   is reflected on the website automatically. Existing RCH tables/layout remain locked.
*/
(function(){
  'use strict';

  const SID='1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0';
  const GID='1044088930';
  const FALLBACK='01-08-2026';
  let lastDate='';
  let busy=false;

  function clean(v){
    return String(v==null?'':v)
      .replace(/<br\s*\/?>(?=.)/gi,' ')
      .replace(/\s+/g,' ')
      .trim();
  }

  function extractDate(text){
    const s=clean(text);
    let m=s.match(/(?:as\s*on\s*date|as\s*on|date)\s*[:\-]?\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})/i);
    if(m) return m[1].replace(/\//g,'-');
    m=s.match(/\b(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})\b/);
    return m ? m[1].replace(/\//g,'-') : '';
  }

  function findDate(dt){
    if(!dt) return '';
    const priority=[];
    const all=[];
    try{
      const cols=dt.getNumberOfColumns();
      for(let c=0;c<cols;c++){
        const x=clean(dt.getColumnLabel(c));
        if(x){ all.push(x); if(/rch|as\s*on|date|fy\s*20|till|upto/i.test(x)) priority.push(x); }
      }
    }catch(e){}
    try{
      for(let r=0;r<dt.getNumberOfRows();r++){
        const row=[];
        for(let c=0;c<dt.getNumberOfColumns();c++){
          try{ row.push(clean(dt.getFormattedValue(r,c))); }catch(e){ row.push(''); }
        }
        const x=row.filter(Boolean).join(' ');
        if(!x) continue;
        all.push(x);
        if(/rch|as\s*on|date|fy\s*20|till|upto/i.test(x)) priority.push(x);
      }
    }catch(e){}
    for(const x of priority){ const d=extractDate(x); if(d) return d; }
    for(const x of all){ const d=extractDate(x); if(d) return d; }
    return '';
  }

  function applyDate(date){
    if(!date) return;
    lastDate=date;
    const text='RCH 2.0 PW Registration Detail FY 2026-27 | As On Date : '+date;
    document.querySelectorAll('#rchContainer .rch-page-date').forEach(el=>{ el.textContent=text; });
    document.querySelectorAll('.rch-page-date').forEach(el=>{ el.textContent=text; });
  }

  function readSheet(){
    if(busy || !window.google || !google.visualization || !google.visualization.Query) return;
    busy=true;
    const url='https://docs.google.com/spreadsheets/d/'+SID+'/gviz/tq?gid='+encodeURIComponent(GID)+'&headers=0&tqx=out:html';
    const q=new google.visualization.Query(url);
    q.setQuery('select *');
    q.send(function(resp){
      busy=false;
      try{
        if(resp && !resp.isError()){
          const d=findDate(resp.getDataTable());
          if(d) applyDate(d);
        }
      }catch(e){ busy=false; }
    });
  }

  function patchParser(){
    if(typeof window.parseRCHData!=='function' || window.parseRCHData.__rchLiveHeading) return;
    const original=window.parseRCHData;
    function patched(data){
      const parsed=original.apply(this,arguments) || {};
      const d=findDate(data);
      if(d) parsed.asOnDate=d;
      else if(!parsed.asOnDate) parsed.asOnDate=lastDate||FALLBACK;
      return parsed;
    }
    patched.__rchLiveHeading=true;
    patched.__original=original;
    window.parseRCHData=patched;
  }

  function boot(){
    patchParser();
    readSheet();
    if(lastDate) applyDate(lastDate);
  }

  // Run repeatedly because the dashboard is an iframe and the RCH data/rendering
  // can finish before this fix script is injected.
  [0,250,750,1500,3000,5000].forEach(t=>setTimeout(boot,t));
  setInterval(boot,5000);

  // Also update immediately after RCH is opened.
  document.addEventListener('click',function(e){
    const b=e.target && e.target.closest ? e.target.closest('.menu-btn') : null;
    if(b && /rch\s*2\.0/i.test(b.textContent||'')) setTimeout(readSheet,300);
  },true);
})();
