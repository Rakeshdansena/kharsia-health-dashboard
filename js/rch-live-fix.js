/* RCH live date fix - keeps the existing RCH layout/tables locked */
(function(){
  'use strict';

  function clean(x){
    return String(x==null?'':x).replace(/<br\s*\/?>/gi,' ').replace(/\s+/g,' ').trim();
  }

  function extractDate(text){
    const s=clean(text);
    let m=s.match(/(?:as\s*on\s*date|as\s*on|date)\s*:?\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})/i);
    if(m)return m[1].replace(/\//g,'-');
    m=s.match(/\b(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})\b/);
    return m?m[1].replace(/\//g,'-'):'';
  }

  function sheetDate(data){
    if(!data)return '';
    const priority=[];
    const all=[];
    try{
      const cols=data.getNumberOfColumns();
      for(let c=0;c<cols;c++){
        const label=clean(data.getColumnLabel(c));
        if(label){all.push(label);if(/rch|as\s*on|date|fy\s*20|till|upto/i.test(label))priority.push(label)}
      }
    }catch(e){}
    try{
      const rows=data.getNumberOfRows(),cols=data.getNumberOfColumns();
      for(let r=0;r<rows;r++){
        const parts=[];
        for(let c=0;c<cols;c++){try{parts.push(clean(data.getFormattedValue(r,c)))}catch(e){}}
        const text=parts.filter(Boolean).join(' ');
        if(!text)continue;
        all.push(text);
        if(/rch|as\s*on|date|fy\s*20|till|upto/i.test(text))priority.push(text);
      }
    }catch(e){}
    for(const x of priority){const d=extractDate(x);if(d)return d}
    for(const x of all){const d=extractDate(x);if(d)return d}
    return '';
  }

  function install(){
    if(typeof window.parseRCHData!=='function')return false;
    if(window.parseRCHData.__liveDateFix)return true;
    const original=window.parseRCHData;
    function patched(data){
      const parsed=original.apply(this,arguments);
      const liveDate=sheetDate(data);
      if(liveDate)parsed.asOnDate=liveDate;
      return parsed;
    }
    patched.__liveDateFix=true;
    patched.__original=original;
    window.parseRCHData=patched;
    setTimeout(function(){
      try{
        if(typeof currentReportIndex!=='undefined' && REPORTS[currentReportIndex] && REPORTS[currentReportIndex].name==='RCH 2.0' && typeof currentData!=='undefined' && currentData && typeof window.renderRCHTable==='function'){
          window.renderRCHTable(currentData);
        }
      }catch(e){}
    },100);
    return true;
  }

  function boot(){if(install())return;setTimeout(boot,250)}
  boot();
  setInterval(install,1000);
})();
