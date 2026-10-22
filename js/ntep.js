/* NTEP MODULE — independent module; Ayushman Card and RCH 2.0 remain untouched. */
(function(){
'use strict';
var SOURCES={
 panchayat:{id:'1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0',gid:'976295329',label:'Panchayat Wise Report'},
 abhiyan:{id:'16rdSpS8htOHuL6RFm7u58uLjBBZ64uMPmmaZFWSQf5c',gid:'1906211315',label:'Abhiyan Report'}
};
var $=function(id){return document.getElementById(id)};
var clean=function(v){return String(v==null?'':v).replace(/\s+/g,' ').trim()};
var esc=function(v){return String(v==null?'':v).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]})};
var num=function(v){var n=parseFloat(clean(v).replace(/,/g,'').replace(/%/g,'').replace(/[₹]/g,''));return Number.isFinite(n)?n:null};
var activeTab='panchayat',cache={panchayat:null,abhiyan:null},loading={panchayat:false,abhiyan:false},shown={panchayat:null,abhiyan:null};

function addStyles(){
 if($('ntep-style'))return;
 var s=document.createElement('style');s.id='ntep-style';
 s.textContent=[
 '#ntepModule{font-family:Arial,"Noto Sans",sans-serif;color:#123047}',
 '#ntepModule .ntep-hero{padding:22px;border-radius:20px;color:#fff;background:radial-gradient(circle at 85% 5%,rgba(255,255,255,.28),transparent 30%),linear-gradient(125deg,#073b75,#087e8b 58%,#0f766e);box-shadow:0 14px 32px rgba(7,59,117,.20);margin-bottom:16px}',
 '#ntepModule .ntep-hero h2{font-size:25px;margin:0 0 7px;font-weight:900}#ntepModule .ntep-hero p{margin:0;opacity:.9;font-size:13px}',
 '#ntepModule .ntep-tabs{display:flex;flex-wrap:wrap;gap:10px;margin:0 0 14px}',
 '#ntepModule .ntep-tab{border:1px solid rgba(8,126,139,.22);padding:12px 16px;border-radius:14px;cursor:pointer;font-weight:850;color:#075985;background:linear-gradient(135deg,#fff,#effaff);box-shadow:0 6px 15px rgba(15,45,75,.07),inset 0 1px 0 #fff;transition:.18s}',
 '#ntepModule .ntep-tab.active{color:#fff;background:linear-gradient(135deg,#075985,#0f766e);border-color:#0f766e;box-shadow:0 9px 20px rgba(7,89,133,.25)}',
 '#ntepModule .ntep-summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(145px,1fr));gap:10px;margin:0 0 14px}',
 '#ntepModule .ntep-kpi{background:linear-gradient(145deg,rgba(255,255,255,.98),rgba(234,248,252,.92));border:1px solid #d8e8ef;border-radius:16px;padding:13px;box-shadow:0 8px 20px rgba(15,45,75,.06),inset 0 1px 0 #fff;min-width:0}',
 '#ntepModule .ntep-kpi span{display:block;font-size:11px;font-weight:800;color:#60788b;line-height:1.3}#ntepModule .ntep-kpi strong{display:block;font-size:24px;font-weight:950;color:#075985;margin-top:5px;overflow-wrap:anywhere}',
 '#ntepModule .ntep-toolbar{display:flex;gap:9px;flex-wrap:wrap;align-items:center;background:rgba(255,255,255,.78);border:1px solid #dce9f0;border-radius:15px;padding:12px;margin-bottom:12px}',
 '#ntepModule .ntep-toolbar input,#ntepModule .ntep-toolbar select{padding:10px 12px;border:1px solid #cbdce5;border-radius:10px;background:#fff;min-width:170px;flex:1;color:#17384f}',
 '#ntepModule .ntep-action{border:0;border-radius:10px;padding:10px 13px;color:#fff;font-weight:850;cursor:pointer;box-shadow:0 5px 12px rgba(15,45,75,.12);background:linear-gradient(135deg,#075985,#0f766e)}#ntepModule .ntep-action.excel{background:linear-gradient(135deg,#15803d,#16a34a)}#ntepModule .ntep-action.pdf{background:linear-gradient(135deg,#b91c1c,#e11d48)}',
 '#ntepModule .ntep-status{padding:11px 13px;border-radius:11px;margin-bottom:12px;background:#e0f2fe;color:#075985;font-weight:750;font-size:13px}#ntepModule .ntep-status.error{background:#fee2e2;color:#991b1b}',
 '#ntepModule .ntep-table-wrap{overflow:auto;background:#fff;border:1px solid #dbe7ee;border-radius:15px;box-shadow:0 8px 20px rgba(15,45,75,.05)}#ntepModule table{width:100%;border-collapse:separate;border-spacing:0;min-width:1000px;font-size:12px}',
 '#ntepModule th{position:sticky;top:0;background:linear-gradient(135deg,#073b75,#087e8b);color:#fff;text-align:center;padding:11px 9px;white-space:normal;line-height:1.3;cursor:pointer;border-right:1px solid rgba(255,255,255,.18)}#ntepModule td{padding:9px 8px;border-right:1px solid #e3edf2;border-bottom:1px solid #e3edf2;text-align:center;white-space:nowrap}',
 '#ntepModule td:first-child{text-align:left;font-weight:800}#ntepModule tbody tr:nth-child(even){background:#f4f9fc}#ntepModule tbody tr:hover{background:#e0f2fe}',
 '#ntepModule .ntep-pct-high{background:#dcfce7;color:#166534;font-weight:900}#ntepModule .ntep-pct-mid{background:#fef3c7;color:#92400e;font-weight:900}#ntepModule .ntep-pct-low{background:#fee2e2;color:#991b1b;font-weight:900}#ntepModule .ntep-total{background:#dbeafe!important;font-weight:950}#ntepModule .ntep-empty{padding:28px;text-align:center;color:#64748b}',
 '@media(max-width:650px){#ntepModule .ntep-hero{padding:17px}#ntepModule .ntep-hero h2{font-size:21px}#ntepModule .ntep-kpi strong{font-size:21px}#ntepModule .ntep-toolbar input,#ntepModule .ntep-toolbar select{min-width:100%}}',
 '@media print{#ntepModule .ntep-tabs,#ntepModule .ntep-toolbar,#ntepModule .ntep-status{display:none!important}#ntepModule .ntep-table-wrap{overflow:visible;border:0;box-shadow:none}#ntepModule table{min-width:0!important;font-size:7px!important}#ntepModule th,#ntepModule td{position:static;padding:3px!important;white-space:normal!important}}'
 ].join('\n');
 document.head.appendChild(s)
}
function ensureUI(){
 if($('ntepModule'))return;
 var area=$('pdfArea')||$('reportPage'),box=document.createElement('section');box.id='ntepModule';box.style.display='none';
 box.innerHTML='<div class="ntep-hero"><h2>🫁 NTEP · TB Programme</h2><p>National Tuberculosis Elimination Programme · Block Kharsia, District Raigarh</p></div>'+
 '<div class="ntep-tabs" role="tablist"><button class="ntep-tab active" data-tab="panchayat">🏘️ Panchayat Wise Report</button><button class="ntep-tab" data-tab="abhiyan">🎯 Abhiyan Report</button></div>'+
 '<div id="ntepSummary" class="ntep-summary"></div><div class="ntep-toolbar"><input id="ntepSearch" type="search" placeholder="🔎 पंचायत / PHC / SHC खोजें..." aria-label="Search report"><select id="ntepSectorFilter" aria-label="PHC filter"><option value="">सभी PHC</option></select><button id="ntepRefresh" class="ntep-action">🔄 Refresh</button><button id="ntepExcel" class="ntep-action excel">📊 Excel</button><button id="ntepPDF" class="ntep-action pdf">📄 PDF / Print</button></div>'+
 '<div id="ntepStatus" class="ntep-status">Google Sheets से डेटा लोड हो रहा है...</div><div class="ntep-table-wrap"><table id="ntepTable"><thead></thead><tbody></tbody></table></div>';
 area.appendChild(box);
 box.querySelectorAll('[data-tab]').forEach(function(b){b.addEventListener('click',function(){switchTab(b.dataset.tab)})});
 $('ntepSearch').addEventListener('input',renderTable);$('ntepSectorFilter').addEventListener('change',renderTable);
 $('ntepRefresh').addEventListener('click',function(){loadTab(activeTab,true)});
 $('ntepExcel').addEventListener('click',exportExcel);$('ntepPDF').addEventListener('click',function(){window.print()})
}
function switchTab(tab){activeTab=tab;$('ntepModule').querySelectorAll('[data-tab]').forEach(function(b){b.classList.toggle('active',b.dataset.tab===tab)});$('ntepSearch').value='';$('ntepSectorFilter').value='';$('ntepSearch').placeholder=tab==='panchayat'?'🔎 पंचायत / PHC / SHC खोजें...':'🔎 अभियान / सुविधा / विवरण खोजें...';loadTab(tab)}
function status(msg,error){var e=$('ntepStatus');if(e){e.textContent=msg;e.classList.toggle('error',!!error)}}
function parseCSV(text){
 var out=[],row=[],field='',quoted=false;text=String(text||'').replace(/^\uFEFF/,'');
 for(var i=0;i<text.length;i++){var c=text[i],n=text[i+1];if(c==='"'&&quoted&&n==='"'){field+='"';i++}else if(c==='"')quoted=!quoted;else if(c===','&&!quoted){row.push(field);field=''}else if((c==='\n'||c==='\r')&&!quoted){if(c==='\r'&&n==='\n')i++;row.push(field);if(row.some(function(x){return clean(x)}))out.push(row);row=[];field=''}else field+=c}
 if(field.length||row.length){row.push(field);if(row.some(function(x){return clean(x)}))out.push(row)}return out
}
async function fetchCSV(src){
 var url='https://docs.google.com/spreadsheets/d/'+src.id+'/export?format=csv&gid='+src.gid;
 var res=await fetch(url,{cache:'no-store'});if(!res.ok)throw new Error('Google Sheet नहीं खुली ('+res.status+'); Viewer sharing जाँचें।');
 var txt=await res.text();if(/<!doctype html|<html/i.test(txt.slice(0,300)))throw new Error('Sheet CSV में उपलब्ध नहीं है। Anyone with the link → Viewer जाँचें।');return parseCSV(txt)
}
function normalize(raw,tab){
 if(!raw||!raw.length)return {headers:[],rows:[],title:''};
 var best=0,score=-1;
 for(var i=0;i<Math.min(raw.length,12);i++){var line=raw[i].map(clean).join(' ').toLowerCase(),s=0;['ग्राम पंचायत','पंचायत','target','लक्ष्य','screening','जांच','जाँच','achievement','अभियान','notification','population','facility','report'].forEach(function(k){if(line.indexOf(k)>=0)s++});if(s>score){score=s;best=i}}
 var headers=(raw[best]||[]).map(function(v,i){return clean(v)||('Column '+(i+1))});
 var rows=raw.slice(best+1).filter(function(r){return r.some(function(v){return clean(v)})}).map(function(r){return headers.map(function(_,i){return clean(r[i]||'')})});
 var title=raw.slice(0,best).map(function(r){return r.map(clean).filter(Boolean).join(' · ')}).filter(Boolean).join(' | ');
 return {headers:headers,rows:rows,title:title}
}
function idx(headers,terms){return headers.findIndex(function(h){return terms.some(function(t){return clean(h).toLowerCase().indexOf(t)>=0})})}
function stats(d,tab){
 var h=d.headers,r=d.rows.filter(function(x){return !/^(total|कुल|योग)$/i.test(clean(x[0]))});
 var pct=idx(h,['प्रतिशत','%','screening %','achievement %']),target=idx(h,['लक्ष्य','target']),ach=idx(h,['संदेहास्पद मरीजों','जनवरी से सितम्बर','कुल जांच की संख्या','कुल जाँच की संख्या','जाँच की संख्या','screening','achievement','उपलब्धि']),pop=idx(h,['कुल जनसंख्या','population']),notif=idx(h,['टीबी नोटिफिकेशन','notification']);
 var total=r.find(function(x){return /^(total|कुल|योग)$/i.test(clean(x[0]))});
 var sum=function(i){if(i<0)return null;if(total&&num(total[i])!==null)return num(total[i]);return r.reduce(function(a,x){return a+(num(x[i])||0)},0)};
 var cards=[];
 if(tab==='panchayat'){
  if(pop>=0)cards.push(['कुल जनसंख्या',sum(pop).toLocaleString('en-IN')]);
  if(target>=0)cards.push(['कुल जाँच लक्ष्य',sum(target).toLocaleString('en-IN')]);
  if(ach>=0)cards.push(['कुल जाँच',sum(ach).toLocaleString('en-IN')]);
  if(target>=0&&ach>=0)cards.push(['लक्ष्य उपलब्धि',(sum(ach)/Math.max(1,sum(target))*100).toFixed(1)+'%']);
  if(notif>=0)cards.push(['टीबी नोटिफिकेशन',sum(notif).toLocaleString('en-IN')])
 }else{
  h.forEach(function(label,i){var vals=r.map(function(x){return num(x[i])}).filter(function(v){return v!==null});if(vals.length&&cards.length<5)cards.push([label,sum(i).toLocaleString('en-IN')])})
 }
 var rank=pct>=0?pct:ach,ranked=r.filter(function(x){return rank>=0&&num(x[rank])!==null}).slice().sort(function(a,b){return num(b[rank])-num(a[rank])});
 if(ranked.length){cards.push(['सबसे अधिक',clean(ranked[0][0])+' · '+clean(ranked[0][rank])]);cards.push(['सबसे कम',clean(ranked[ranked.length-1][0])+' · '+clean(ranked[ranked.length-1][rank])])}
 if(!cards.length)cards.push(['कुल रिकॉर्ड',String(r.length)]);
 return {cards:cards,pct:pct,target:target,ach:ach,rank:rank}
}
function loadTab(tab,force){
 if(loading[tab])return;if(cache[tab]&&!force){renderTable();return}
 loading[tab]=true;status('Google Sheets से '+SOURCES[tab].label+' लोड हो रही है…');
 fetchCSV(SOURCES[tab]).then(function(raw){var d=normalize(raw,tab);if(!d.headers.length||!d.rows.length)throw new Error('शीट में पढ़ने योग्य header या data नहीं मिला।');cache[tab]=d;shown[tab]=null;renderTable();status(SOURCES[tab].label+' लोड हो गई · '+d.rows.filter(function(r){return !/^(total|कुल|योग)$/i.test(clean(r[0]))}).length+' रिकॉर्ड')}).catch(function(e){status(e&&e.message?e.message:'डेटा लोड नहीं हो पाया।',true)}).finally(function(){loading[tab]=false})
}
function renderTable(){
 var d=cache[activeTab];if(!d)return;
 var search=clean($('ntepSearch').value).toLowerCase(),sector=$('ntepSectorFilter').value,s=stats(d,activeTab);
 $('ntepSummary').innerHTML=s.cards.map(function(c){return '<div class="ntep-kpi"><span>'+esc(c[0])+'</span><strong>'+esc(c[1])+'</strong></div>'}).join('');
 var sectorI=idx(d.headers,['link phc','phc','लिंक पी.एच.सी','लिंक पीएचसी','पी.एच.सी','पीएचसी','sector','block','विकासखण्ड']),sel=$('ntepSectorFilter'),old=sel.value;
 var options=Array.from(new Set(d.rows.map(function(r){return sectorI>=0?clean(r[sectorI]):''}).filter(Boolean))).sort();
 sel.innerHTML='<option value="">सभी PHC / Sector</option>'+options.map(function(v){return '<option value="'+esc(v)+'">'+esc(v)+'</option>'}).join('');sel.style.display=sectorI>=0?'':'none';if(options.indexOf(old)>=0)sel.value=old;
 var rows=d.rows.filter(function(r){if(/^total$/i.test(clean(r[0]))||/^(कुल|योग)$/i.test(clean(r[0])))return false;if(search&&r.join(' ').toLowerCase().indexOf(search)<0)return false;if(sector&&sectorI>=0&&clean(r[sectorI])!==sector)return false;return true});
 if(activeTab==='panchayat'&&s.rank>=0)rows.sort(function(a,b){return (num(b[s.rank])||0)-(num(a[s.rank])||0)});
 var thead=$('ntepTable').querySelector('thead'),tbody=$('ntepTable').querySelector('tbody');
 thead.innerHTML='<tr>'+d.headers.map(function(h,i){return '<th data-col="'+i+'" title="Sort">'+esc(h)+' ↕</th>'}).join('')+'</tr>';
 function draw(list,colour){tbody.innerHTML=list.map(function(r){return '<tr>'+d.headers.map(function(h,i){var v=num(r[i]),cl='';if(colour&&i===s.pct&&v!==null)cl=v>=80?'ntep-pct-high':v>=50?'ntep-pct-mid':'ntep-pct-low';return '<td class="'+cl+'">'+esc(r[i])+'</td>'}).join('')+'</tr>'}).join('')||'<tr><td colspan="'+d.headers.length+'" class="ntep-empty">कोई रिकॉर्ड नहीं मिला।</td></tr>'}
 draw(rows,true);
 thead.querySelectorAll('th').forEach(function(th){th.addEventListener('click',function(){var i=Number(th.dataset.col),asc=th.dataset.asc!=='true';th.dataset.asc=String(asc);rows.sort(function(a,b){var x=num(a[i]),y=num(b[i]),cmp=x!==null&&y!==null?x-y:clean(a[i]).localeCompare(clean(b[i]),'en',{numeric:true});return asc?cmp:-cmp});draw(rows,false)})});
 shown[activeTab]={headers:d.headers,rows:rows}
}
function exportExcel(){
 var d=shown[activeTab]||cache[activeTab];if(!d)return;
 if(window.XLSX){var ws=XLSX.utils.aoa_to_sheet([d.headers].concat(d.rows||[])),wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,activeTab==='panchayat'?'Panchayat Wise':'Abhiyan');XLSX.writeFile(wb,'NTEP_'+(activeTab==='panchayat'?'Panchayat_Wise':'Abhiyan_Report')+'.xlsx')}
 else{var csv=[d.headers].concat(d.rows||[]).map(function(r){return r.map(function(v){return '"'+String(v==null?'':v).replace(/"/g,'""')+'"'}).join(',')}).join('\r\n'),a=document.createElement('a');a.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));a.download='NTEP_Report.csv';a.click();URL.revokeObjectURL(a.href)}
}
function showNTEP(index){
 document.querySelectorAll('.page').forEach(function(p){p.classList.remove('active');p.style.display='none'});
 var rp=$('reportPage');rp.classList.add('active');rp.style.display='block';$('dashboardPage').classList.remove('active');$('dashboardPage').style.display='none';
 $('reportTitle').textContent='🫁 NTEP · TB Programme';$('pdfTitle').textContent='Kharsia Health Dashboard — NTEP';$('reportDescription').textContent='Panchayat Wise Report एवं Abhiyan Report · Live Google Sheets';$('searchBox').value='';
 document.querySelectorAll('#menuButtons .menu-btn').forEach(function(b,i){b.classList.toggle('active',i===index)});
 document.querySelectorAll('.report-actions .action-btn').forEach(function(b){b.style.display=b.classList.contains('back-btn')?'':'none'});$('searchBox').style.display='none';
 var box=document.querySelector('.table-box');if(box)box.style.display='none';
 ['rchContainer','nrcModule','blindnessModule','nqasModule','dialysisModule','telemedicineModule','jasMeetingModule','ncdModuleContainer'].forEach(function(id){var el=$(id);if(el)el.style.display='none'});
 var st=$('status');if(st)st.style.display='none';ensureUI();addStyles();$('ntepModule').style.display='block';switchTab(activeTab||'panchayat');window.scrollTo(0,0)
}
function init(){
 if(!Array.isArray(REPORTS)||typeof window.createMenu!=='function'||typeof window.openReport!=='function')return false;
 if(REPORTS.some(function(r){return r.name==='NTEP'}))return true;
 REPORTS.push({name:'NTEP',icon:'🫁',description:'TB Programme · Panchayat Wise एवं Abhiyan Report',ntep:true});
 var oldOpen=window.openReport;
 window.openReport=function(index){var report=REPORTS[index];if(report&&report.name==='NTEP'){showNTEP(index);return}var m=$('ntepModule');if(m)m.style.display='none';var st=$('status');if(st)st.style.display='';document.querySelectorAll('.report-actions .action-btn').forEach(function(b){b.style.display=''});var search=$('searchBox');if(search)search.style.display='';return oldOpen.apply(this,arguments)};
 var oldDashboard=window.showDashboard;
 window.showDashboard=function(){var m=$('ntepModule');if(m)m.style.display='none';var st=$('status');if(st)st.style.display='';document.querySelectorAll('.report-actions .action-btn').forEach(function(b){b.style.display=''});var s=$('searchBox');if(s)s.style.display='';return oldDashboard.apply(this,arguments)};
 window.createMenu();return true
}
var tries=0,timer=setInterval(function(){tries++;if(init()||tries>40)clearInterval(timer)},250);
})();