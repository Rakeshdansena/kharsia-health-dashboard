(function(){
'use strict';
var D={date:'30-09-2026',main:[],cat:[]};
var SPREADSHEET_ID='1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0';
var SHEET_GID='1502752823';
var liveLoaded=false;
function cellValue(c){return c && c.v!==null && c.v!==undefined ? (c.f!==undefined && c.f!==null ? c.f : c.v) : '';}
function clean(v){return String(v===undefined||v===null?'':v).replace(/\\s+/g,' ').trim();}
function numOrText(v){if(typeof v==='number')return v;var s=clean(v);if(!s)return '';var n=Number(s.replace(/,/g,''));return Number.isFinite(n)&&/^[-+]?[0-9,]+(\\.[0-9]+)?%?$/.test(s)?(s.indexOf('%')>=0?s:n):s;}
function loadLiveData(done){
 var cb='__rbskSheet_'+Date.now()+'_'+Math.floor(Math.random()*10000);
 var finished=false,timer=setTimeout(function(){finish(new Error('Google Sheet से डेटा लोड करने में समय लग गया।'))},15000);
 function finish(err){if(finished)return;finished=true;clearTimeout(timer);try{delete window[cb]}catch(e){}if(script&&script.parentNode)script.parentNode.removeChild(script);if(done)done(err);}
 window[cb]=function(resp){
  if(!resp||resp.status!=='ok'||!resp.table||!resp.table.rows){finish(new Error('Google Sheet का डेटा पढ़ा नहीं जा सका।'));return;}
  try{
   var rows=resp.table.rows.map(function(r){return (r.c||[]).map(cellValue);});
   var allText=rows.map(function(r){return r.map(clean).join(' ')}).join(' ');
   var dm=allText.match(/दिनांक\\s*([0-3]?\\d[-/]\\d{1,2}[-/]\\d{4})/);
   if(dm)D.date=dm[1].replace(/\\//g,'-');
   var mainRows=[];
   rows.forEach(function(r){
    var joined=r.map(clean).join(' ');
    var team=clean(r[1]);
    if(/^टीम\\s*ए$/i.test(team)||/^टीम\\s*A$/i.test(team))mainRows.push({n:1,t:'टीम ए',a:r.slice(2,18).map(numOrText)});
    else if(/^टीम\\s*बी$/i.test(team)||/^टीम\\s*B$/i.test(team))mainRows.push({n:2,t:'टीम बी',a:r.slice(2,18).map(numOrText)});
    else if((team==='योग'||team.toLowerCase()==='total') && r.length>=18)mainRows.push({n:'',t:'योग',a:r.slice(2,18).map(numOrText)});
   });
   if(mainRows.length>=2){
    var a=mainRows.find(function(x){return x.t==='टीम ए'}),b=mainRows.find(function(x){return x.t==='टीम बी'}),tot=mainRows.find(function(x){return x.t==='योग'});
    if(!tot&&a&&b){tot={n:'',t:'योग',a:a.a.map(function(v,i){var x=Number(String(v).replace('%','')),y=Number(String(b.a[i]).replace('%',''));if(i===2||i===5||i===8||i===11)return (x+y)+'%';return Number.isFinite(x)&&Number.isFinite(y)?x+y:''})};}
    D.main=[a,b,tot].filter(Boolean);
   }
   var catRows=[];
   rows.forEach(function(r){
    var id=clean(r[1]);
    if(id==='3863687A'||id==='3863687B'||id.toLowerCase()==='total'){
     var vals=r.slice(2).filter(function(v){return v!==''&&v!==null&&v!==undefined;}).slice(0,30).map(numOrText);
     while(vals.length<30)vals.push('');
     var groups=[];for(var i=0;i<6;i++)groups.push(vals.slice(i*5,(i+1)*5));
     catRows.push({n:id==='3863687A'?1:(id==='3863687B'?2:''),id:id,v:groups});
    }
   });
   if(catRows.length>=2){var ca=catRows.find(function(x){return x.id==='3863687A'}),cb2=catRows.find(function(x){return x.id==='3863687B'}),ct=catRows.find(function(x){return x.id.toLowerCase()==='total'});D.cat=[ca,cb2,ct].filter(Boolean);}
   if(!D.main.length||D.cat.length<2){finish(new Error('Sheet में RBSK की टीम/Category वाली पंक्तियाँ पहचान नहीं पाईं।'));return;}
   liveLoaded=true;finish(null);
  }catch(e){finish(e);}
 };
 var script=document.createElement('script');
 var url='https://docs.google.com/spreadsheets/d/'+SPREADSHEET_ID+'/gviz/tq?gid='+SHEET_GID+'&tqx=out:json;responseHandler:'+cb;
 script.src=url;script.onerror=function(){finish(new Error('Google Sheet उपलब्ध नहीं है। Sharing और इंटरनेट कनेक्शन जाँचें।'));};
 document.head.appendChild(script);
}


function css(){
if(document.getElementById('rbsk-css'))return;
var s=document.createElement('style');s.id='rbsk-css';
s.textContent='#rbskPage{display:none!important}#rbskPage.rbsk-visible{display:block!important}.rbsk-box{background:#fff;border-radius:14px;box-shadow:0 3px 12px rgba(0,0,0,.07);padding:16px;margin-bottom:16px}.rbsk-head{text-align:center;background:linear-gradient(135deg,#075985,#0f766e);color:#fff;border-radius:12px;padding:14px;margin-bottom:14px}.rbsk-head h2{margin:0;font-size:21px}.rbsk-head p{margin:5px 0 0;font-size:13px;font-weight:700}.rbsk-tabs{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}.rbsk-tab{border:1px solid #cbd5e1;background:#f8fafc;color:#334155;border-radius:9px;padding:10px 15px;font-weight:800;cursor:pointer}.rbsk-tab.active{background:#0f766e;color:#fff}.rbsk-pane{display:none}.rbsk-pane.active{display:block}.rbsk-title{text-align:center;color:#075985;font-size:18px;font-weight:900;margin:5px}.rbsk-sub{text-align:center;color:#64748b;font-size:12px;font-weight:700;margin-bottom:12px}.rbsk-scroll{overflow:auto;border:1px solid #dbe3ec;border-radius:10px}.rbsk-table{border-collapse:collapse;width:100%;min-width:1550px;font-size:12px}.rbsk-table.cat{min-width:2050px}.rbsk-cat-tabs{display:flex;gap:6px;flex-wrap:wrap;margin:8px 0 12px}.rbsk-cat-tab{padding:8px 14px;border:1px solid #075985;border-radius:7px;background:#fff;color:#075985;font-weight:800;cursor:pointer}.rbsk-cat-tab.active{background:#075985;color:#fff}.rbsk-section-title{text-align:center;color:#075985;font-size:15px;font-weight:900;margin:5px 0 8px}.rbsk-category-one{min-width:900px}.rbsk-main-table{min-width:1350px}.rbsk-main-table .rbsk-title-row th{background:#075985;color:#fff;font-size:16px;padding:10px}.rbsk-main-table .group th{background:#0f766e}.rbsk-main-table th{font-size:11px}.rbsk-main-table td{font-size:12px}.rbsk-table th,.rbsk-table td{border:1px solid #cbd5e1;padding:7px 5px;text-align:center;vertical-align:middle;white-space:normal}.rbsk-table th{background:#075985;color:#fff;font-weight:800}.rbsk-table .group th,.rbsk-table th.group{background:#0f766e}.rbsk-table tr:nth-child(even){background:#f8fafc}.rbsk-table tr.total{background:#dbeafe!important;font-weight:900}.rbsk-team{font-weight:900;color:#075985}.rbsk-note{margin-top:10px;padding:9px 12px;background:#f1f5f9;border-radius:8px;color:#475569;font-size:12px}.rbsk-empty{padding:25px;text-align:center;color:#64748b}';
document.head.appendChild(s)
}

function hideAllPages(){
document.querySelectorAll('.page').forEach(function(x){x.classList.remove('active');x.style.display='none'});
var d=document.getElementById('dashboardPage');if(d)d.style.display='none';
var r=document.getElementById('reportPage');if(r)r.style.display='none';
}

function showRBSK(p,b){
hideAllPages();
var dash=document.getElementById('dashboardBtn');if(dash)dash.classList.remove('active');
document.querySelectorAll('#menuButtons .menu-btn').forEach(function(x){x.classList.remove('active')});
b.classList.add('active');
p.classList.add('active','rbsk-visible');
p.style.display='block';
window.scrollTo(0,0);
}

function mainTable(){
var h='<div class="rbsk-box"><div class="rbsk-scroll"><table class="rbsk-table rbsk-main-table"><thead>'+
'<tr class="rbsk-title-row"><th colspan="18">राष्ट्रीय बाल स्वास्थ्य कार्यक्रम की जानकारी वित्तीय वर्ष 2026-27 दिनांक 31-08-2026 की स्थिति में</th></tr>'+
'<tr class="group"><th rowspan="2">क्र</th><th rowspan="2">टीम</th><th colspan="6">स्कूल जाॅच</th><th colspan="6">आगाॅनबाड़ी FIRST VISIT</th><th colspan="4">Reffer and Treatment</th></tr>'+
'<tr>'+
'<th>कुल स्कूल</th><th>स्कूल विजिट</th><th>प्रतिशत</th><th>कुल जाॅच हेतु लक्ष्य</th><th>कुल जाॅच</th><th>प्रतिशत</th>'+
'<th>कुल आगाॅनबाड़ी</th><th>आगाॅनबाड़ी विजिट</th><th>प्रतिशत</th><th>कुल जाॅच हेतु लक्ष्य</th><th>कुल जाॅच</th><th>प्रतिशत</th>'+
'<th>कुल बीमार बच्चो की संख्या</th><th>कुल जिनका ईलाज किया गया</th><th>कुल रिॅफर</th><th>कुल रिफर किये गये बच्चों का ईलाज किया गया</th>'+
'</tr></thead><tbody>';
D.main.forEach(function(r){
h+='<tr class="'+(r.t==='योग'?'total':'')+'"><td>'+r.n+'</td><td class="rbsk-team">'+r.t+'</td>';
r.a.slice(0,16).forEach(function(v){h+='<td>'+v.toLocaleString('en-IN')+'</td>'});
h+='</tr>';
});
return h+'</tbody></table></div></div>'
}
function cells(a){var s='';a.forEach(function(v){s+='<td>'+v+'</td>'});return s}

function catTable(){
var h='<div class="rbsk-box"><div class="rbsk-title">Team Wise Category Periodical Treatment Report</div><div class="rbsk-sub">Report from 01/04/2026 To 31/03/2027</div><div class="rbsk-cat-tabs">';
['A','B','C','D','E','All Category'].forEach(function(x,i){h+='<button class="rbsk-cat-tab '+(i===0?'active':'')+'" data-cat="'+x.replace(/ /g,'')+'">'+x+'</button>'});
h+='</div><div id="rbskCatContent"></div></div>';
setTimeout(function(){
var box=document.getElementById('rbskCatContent'),tabs=document.querySelectorAll('.rbsk-cat-tab');
function draw(key){
var idx={A:0,B:1,C:2,D:3,E:4,AllCategory:5}[key],title=key==='AllCategory'?'All Category (A+B+C+D+E)':key;
var q='<div class="rbsk-section-title">Category '+title+'</div><div class="rbsk-scroll"><table class="rbsk-table rbsk-category-one"><thead><tr class="group"><th rowspan="2">SrNo.</th><th rowspan="2">Team Id</th><th colspan="5">'+title+'</th></tr><tr><th>Identified</th><th>Under Treatment</th><th>Treatment Over</th><th>% of Treatment</th><th>Pending</th></tr></thead><tbody>';
D.cat.forEach(function(r){var a=r.v[idx];q+='<tr class="'+(r.id==='Total'?'total':'')+'"><td>'+r.n+'</td><td class="rbsk-team">'+(r.id==='Total'?'Total':'Team '+String.fromCharCode(64+r.n))+'</td>'+cells(a)+'</tr>'});
q+='</tbody></table></div>';box.innerHTML=q;
}
tabs.forEach(function(t){t.onclick=function(){tabs.forEach(function(x){x.classList.remove('active')});t.classList.add('active');draw(t.dataset.cat)}});
draw('A');
},0);
return h;
}
function build(){
if(document.getElementById('rbskPage'))return;
var c=document.querySelector('.content'),m=document.getElementById('menuButtons');if(!c||!m)return;
css();
var p=document.createElement('section');p.id='rbskPage';p.className='page';
p.innerHTML='<div class="rbsk-head"><h2>🧒 राष्ट्रीय बाल स्वास्थ्य कार्यक्रम (RBSK)</h2><p>वित्तीय वर्ष 2026-27 · दिनांक '+D.date+' की स्थिति में</p></div><div class="rbsk-tabs"><button class="rbsk-tab active" data-p="rbsk1">📊 Part 1 — राष्ट्रीय बाल स्वास्थ्य कार्यक्रम की जानकारी</button><button class="rbsk-tab" data-p="rbsk2">💊 Part 2 — Category Wise Periodical Treatment Report</button></div><div id="rbsk1" class="rbsk-pane active">'+mainTable()+'</div><div id="rbsk2" class="rbsk-pane">'+catTable()+'</div>';
c.appendChild(p);
var b=Array.from(m.querySelectorAll('.menu-btn')).find(function(x){return (x.textContent||'').indexOf('RBSK')>=0});
if(!b){b=document.createElement('button');b.className='menu-btn';b.textContent='🧒 RBSK';m.appendChild(b)}
b.id='rbskMenu';
var rbskHandler=function(e){if(e){e.preventDefault();e.stopPropagation();if(e.stopImmediatePropagation)e.stopImmediatePropagation()}showRBSK(p,b);setTimeout(function(){var r=document.getElementById('reportPage');if(r)r.style.display='none';var d=document.getElementById('dashboardPage');if(d)d.style.display='none';p.style.display='block'},30);return false};
b.onclick=rbskHandler;
b.addEventListener('click',rbskHandler,true);
p.querySelectorAll('.rbsk-tab').forEach(function(x){x.onclick=function(e){if(e){e.preventDefault();e.stopPropagation()}p.querySelectorAll('.rbsk-tab').forEach(function(y){y.classList.remove('active')});p.querySelectorAll('.rbsk-pane').forEach(function(y){y.classList.remove('active')});x.classList.add('active');var z=document.getElementById(x.dataset.p);z.classList.add('active')}})
}

function installRBSKGuard(){
var m=document.getElementById('menuButtons');if(!m||m.dataset.rbskGuard==='1')return;
m.dataset.rbskGuard='1';
m.addEventListener('click',function(e){
var el=e.target;
while(el&&el!==m&&!el.classList.contains('menu-btn'))el=el.parentElement;
if(!el||el===m)return;
if((el.textContent||'').indexOf('RBSK')<0)return;
var p=document.getElementById('rbskPage');
if(!p)return;
e.preventDefault();e.stopPropagation();if(e.stopImmediatePropagation)e.stopImmediatePropagation();
var buttons=m.querySelectorAll('.menu-btn');buttons.forEach(function(x){x.classList.remove('active')});el.classList.add('active');
hideAllPages();p.classList.add('active','rbsk-visible');p.style.display='block';
setTimeout(function(){var r=document.getElementById('reportPage');if(r){r.classList.remove('active');r.style.display='none'}var d=document.getElementById('dashboardPage');if(d){d.classList.remove('active');d.style.display='none'}p.style.display='block'},10);
window.scrollTo(0,0);
},true);
}
window.buildRBSK=build;
function init(){
 var m=document.getElementById('menuButtons');
 var hasRBSK=m&&Array.from(m.querySelectorAll('.menu-btn')).some(function(x){return (x.textContent||'').indexOf('RBSK')>=0});
 if(document.querySelector('.content')&&m&&hasRBSK){
  build();
  installRBSKGuard();
 }else{
  setTimeout(init,300);
 }
}
init();
function refreshRBSK(){
 loadLiveData(function(err){
  if(err){var head=document.querySelector('#rbskPage .rbsk-head p');if(head)head.textContent='लाइव डेटा लोड नहीं हुआ: '+err.message+' · कृपया Sheet की Viewer sharing जाँचें';return;}
  var p=document.getElementById('rbskPage');if(!p)return;
  var h=p.querySelector('.rbsk-head p');if(h)h.textContent='वित्तीय वर्ष 2026-27 · दिनांक '+D.date+' की स्थिति में';
  var one=document.getElementById('rbsk1');if(one)one.innerHTML=mainTable();
  var two=document.getElementById('rbsk2');if(two)two.innerHTML=catTable();
 });
}
var oldInit=init;
var rbskWait=setInterval(function(){if(document.getElementById('rbskPage')){clearInterval(rbskWait);refreshRBSK();}},400);
})();