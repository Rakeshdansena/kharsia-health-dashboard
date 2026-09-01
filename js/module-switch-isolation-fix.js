/* Module switch isolation: keeps Ayushman Shivir and Wellness Activity independent. */
(function(){'use strict';
function text(el){return (el&&el.textContent||'').replace(/\s+/g,' ').trim()}
function buttons(){return Array.from(document.querySelectorAll('#menuButtons .menu-btn,.menu-btn'))}
function isWellnessButton(b){return /wellness\s*activity/i.test(text(b))}
function isShivirButton(b){return /ayushman\s*shivir/i.test(text(b))}
function findWellnessPage(){
 const nodes=Array.from(document.querySelectorAll('[id],[class]'));
 const hits=nodes.filter(el=>{const s=((el.id||'')+' '+(typeof el.className==='string'?el.className:'')).toLowerCase();return /wellness.*activity|activity.*wellness/.test(s)});
 return hits.find(el=>el.classList&&el.classList.contains('page'))||hits.find(el=>el.classList&&el.classList.contains('module'))||null;
}
function clearShivir(){document.querySelectorAll('.shivir-final').forEach(el=>{el.style.display='none';el.innerHTML=''})}
function setActive(btn){buttons().forEach(x=>x.classList.remove('active'));if(btn)btn.classList.add('active')}
function forceWellness(){const wb=buttons().find(isWellnessButton);if(wb)setActive(wb);clearShivir();const wp=findWellnessPage();if(wp){document.querySelectorAll('.page.active').forEach(p=>{if(p!==wp)p.classList.remove('active')});wp.classList.add('active');wp.style.display='block'}window.__ayushmanShivirActive=false}
function forceShivir(){const sb=buttons().find(isShivirButton);if(sb)setActive(sb);document.querySelectorAll('.page').forEach(p=>{if(/wellness.*activity|activity.*wellness/i.test((p.id||'')+' '+(typeof p.className==='string'?p.className:'')))p.classList.remove('active')});window.__ayushmanShivirActive=true}
function hook(){buttons().forEach(b=>{if(b.dataset.moduleIsolationFix)return;b.dataset.moduleIsolationFix='1';b.addEventListener('click',function(){if(isWellnessButton(this)){setTimeout(forceWellness,20);setTimeout(forceWellness,150);setTimeout(forceWellness,500);setTimeout(forceWellness,1200)}else if(isShivirButton(this)){setTimeout(forceShivir,20);setTimeout(forceShivir,150)}},true)})}
hook();setInterval(hook,300);
})();
