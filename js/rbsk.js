(function(){
'use strict';
var D={date:'31-08-2026',main:[
{n:1,t:'टीम ए',a:[161,79,'49%',11381,4566,'40%',212,212,'100%',5322,5267,'99%',5322,5267,1000,895,106,105]},
{n:2,t:'टीम बी',a:[160,74,'46%',10036,3840,'38%',210,210,'100%',5526,5405,'98%',5526,5405,1419,1322,97,75]},
{n:'',t:'योग',a:[321,153,'48%',21417,8406,'39%',422,422,'100%',10848,10672,'98%',10848,10672,2419,2217,203,180]}],
cat:[
{n:1,id:'3863687A',v:[[2,1,1,'50%',1],[48,2,46,'96%',48],[16,3,13,'81%',16],[7,6,1,'14%',7],[4,0,0,'0%',4],[77,1,76,'99%',1]]},
{n:2,id:'3863687B',v:[[1,0,1,'100%',0],[49,7,42,'86%',49],[50,7,43,'86%',50],[5,2,3,'60%',5],[1,1,20,'20%',0],[106,17,89,'84%',17]]},
{n:'',id:'Total',v:[[3,1,2,'67%',1],[97,9,88,'91%',97],[66,10,56,'85%',66],[12,8,4,'33%',12],[5,1,8,'8%',0],[183,18,165,'90%',18]]}]};

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
var idx={A:0,AllCategory:1,B:2,C:3,D:4,E:5}[key],title=key==='AllCategory'?'All Category (A+B+C+D+E)':key;
var q='<div class="rbsk-section-title">Category '+title+'</div><div class="rbsk-scroll"><table class="rbsk-table rbsk-category-one"><thead><tr class="group"><th rowspan="2">SrNo.</th><th rowspan="2">Team Id</th><th colspan="5">'+title+'</th></tr><tr><th>Identified</th><th>Under Treatment</th><th>Treatment Over</th><th>% of Treatment</th><th>Pending</th></tr></thead><tbody>';
D.cat.forEach(function(r){var a=r.v[idx];q+='<tr class="'+(r.id==='Total'?'total':'')+'"><td>'+r.n+'</td><td class="rbsk-team">'+r.id+'</td>'+cells(a)+'</tr>'});
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
var b=document.createElement('button');b.className='menu-btn';b.textContent='🧒 RBSK';b.id='rbskMenu';m.appendChild(b);
b.onclick=function(e){if(e){e.preventDefault();e.stopPropagation()}showRBSK(p,b);setTimeout(function(){var r=document.getElementById('reportPage');if(r)r.style.display='none';var d=document.getElementById('dashboardPage');if(d)d.style.display='none';p.style.display='block'},20)};b.addEventListener('click',function(e){e.stopImmediatePropagation()},true);
p.querySelectorAll('.rbsk-tab').forEach(function(x){x.onclick=function(e){if(e){e.preventDefault();e.stopPropagation()}p.querySelectorAll('.rbsk-tab').forEach(function(y){y.classList.remove('active')});p.querySelectorAll('.rbsk-pane').forEach(function(y){y.classList.remove('active')});x.classList.add('active');var z=document.getElementById(x.dataset.p);z.classList.add('active')}})
}

function init(){if(document.querySelector('.content')&&document.getElementById('menuButtons'))build();else setTimeout(init,300)}
init();
})();