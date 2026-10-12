/* Global module/page isolation fix.
   Ensures only the selected dashboard page or report module is visible after switching. */
(function(){
'use strict';
var timer=null;
function txt(el){return (el&&el.textContent||'').replace(/\s+/g,' ').trim().toLowerCase();}
function selectedButton(){return document.querySelector('#menuButtons .menu-btn.active')||document.querySelector('.menu-btn.active');}
function hide(el){if(el)el.style.display='none';}
function show(el,display){if(el)el.style.display=display||'block';}
function isolate(){
 var btn=selectedButton(),label=txt(btn);
 var pages=Array.from(document.querySelectorAll('.page'));
 var rbsk=document.getElementById('rbskPage');
 var dashboard=document.getElementById('dashboardPage');
 var report=document.getElementById('reportPage');
 var dashboardVisible=!!(dashboard&&dashboard.classList.contains('active')&&dashboard.style.display!=='none');
 var isDashboard=dashboardVisible||!btn||btn.id==='dashboardBtn'||label.indexOf('dashboard')>=0;
 if(isDashboard){var db=document.querySelector('#menuButtons .menu-btn[id*=dashboard i]')||Array.from(document.querySelectorAll('#menuButtons .menu-btn')).find(function(x){return /dashboard/i.test(txt(x))});if(db){document.querySelectorAll('#menuButtons .menu-btn').forEach(function(x){x.classList.remove('active')});db.classList.add('active')}btn=db;label=txt(btn);}
 pages.forEach(function(p){if(p!==rbsk){p.classList.remove('active');hide(p);}});
 if(isDashboard){
  if(dashboard){dashboard.classList.add('active');show(dashboard);}
  hide(report);hide(rbsk);
  hide(document.getElementById('rchContainer'));
  hide(document.getElementById('telemedicineModule'));
  hide(document.getElementById('jasMeetingModule'));
  hide(document.getElementById('nrcModule'));
  hide(document.getElementById('blindnessModule'));
  hide(document.getElementById('nqasModule'));
  return;
 }
 if(dashboard){dashboard.classList.remove('active');hide(dashboard);}
 if(/rbsk/.test(label)){
  hide(report);hide(document.getElementById('rchContainer'));
  if(rbsk){rbsk.classList.add('active','rbsk-visible');show(rbsk);}
  return;
 }
 hide(rbsk);
 if(report){report.classList.add('active');show(report);}
 var tableBox=document.querySelector('.table-box');
 var rch=document.getElementById('rchContainer');
 var tele=document.getElementById('telemedicineModule');
 var jas=document.getElementById('jasMeetingModule');
 var nrc=document.getElementById('nrcModule');
 var blind=document.getElementById('blindnessModule');
 var nqas=document.getElementById('nqasModule');
 var wellness=document.getElementById('wellnessActivityModule');
 if(/wellness\s*activity/.test(label)){hide(tableBox);hide(rch);hide(tele);hide(jas);hide(nrc);hide(blind);hide(nqas);show(wellness);window.__wellnessActivityActive=true;return;}
 if(/janani portal/.test(label)){window.__wellnessActivityActive=false;if(wellness){hide(wellness);wellness.remove();}return;}
 window.__wellnessActivityActive=false;
 if(wellness){hide(wellness);wellness.remove();}
 [tableBox,rch,tele,jas,nrc,blind,nqas].forEach(hide);
 if(/rch 2/.test(label)){
  show(rch);return;
 }
 if(/telemedicine|e-?sanjeevani/.test(label)){show(tele);return;}
 if(/jas meeting/.test(label)){show(jas);return;}
 if(/nrc kharsia/.test(label)){show(nrc);return;}
 if(/blindness control/.test(label)){show(blind);return;}
 if(/nqas certification/.test(label)){show(nqas);return;}
 if(tableBox)show(tableBox);
}
function schedule(){if(timer)clearTimeout(timer);timer=setTimeout(isolate,180);setTimeout(isolate,650);}
function hook(){
 document.querySelectorAll('#menuButtons .menu-btn,.menu-btn,#dashboardBtn').forEach(function(b){
  if(b.dataset.globalPageIsolation==='1')return;
  b.dataset.globalPageIsolation='1';
  b.addEventListener('click',schedule,false);
 });
}
hook();
setInterval(function(){hook();},500);
document.addEventListener('click',function(e){
 if(e.target&&e.target.closest&&e.target.closest('#menuButtons .menu-btn,#dashboardBtn'))schedule();
},true);
})();