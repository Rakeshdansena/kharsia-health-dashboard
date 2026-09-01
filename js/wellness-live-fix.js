/* WELLNESS ACTIVITY LIVE FIX — Facility Wise + Sector Wise */
(function(){
'use strict';
const SID='1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0';
const GID='447031017';
const TITLE='Ayushman Arogya Mandir Wellness Activity FY 2026-27';
const clean=v=>String(v??'').trim();
const num=v=>{const n=Number(clean(v).replace(/,/g,'').replace('%',''));return Number.isFinite(n)?n:0};
const esc=v=>String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
function css(){if(document.getElementById('wellness-fix-css'))return;const s=document.createElement('style');s.id='wellness-fix-css';s.textContent=`
#wellnessActivityModule{display:block!important;width:100%;background:#fff;padding:12px;box-sizing:border-box;border-radius:12px}
.wellness-title{text-align:center;color:#075985;font-size:18px;font-weight:800;margin:2px 0 16px}
.wellness-section-title{text-align:left;color:#075985;font-size:15px;font-weight:800;margin:18px 0 8px}
.wellness-page{background:#fff;margin-bottom:18px;overflow:auto}
.wellness-table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:11px}
.wellness-table th,.wellness-table td{border:1px solid #334155;padding:7px 5px;text-align:center;vertical-align:middle;word-break:break-word}
.wellness-table th{background:#075985;color:#fff;font-weight:800}
.wellness-table td:nth-child(3),.wellness-table td:nth-child(4){text-align:left}
.wellness-total td{font-weight:800;background:#dbeafe}
.wellness-good{color:#15803d!important;font-weight:900}.wellness-low{color:#dc2626!important;font-weight:900}
.wellness-empty{padding:18px;text-align:center;color:#b91c1c;font-weight:700}
@media(max-width:700px){.wellness-table{font-size:9px}.wellness-table th,.wellness-table td{padding:5px 3px}}
`;document.head.appendChild(s)}
function getRows(dt){const a=[];for(let r=0;r<dt.getNumberOfRows();r++){const v=[];for(let c=0;c<7;c++)v.push(clean(dt.getFormattedValue(r,c)));a.push(v)}return a}
function isFacility(v){return /^\d{9,12}$/.test(v[1]) && !!v[2] && !!v[3] && num(v[4])>0}
function isSector(v){return !!v[2] && num(v[3])>0 && num(v[4])>0 && !/^\d{9,12}$/.test(v[1])}
function pctClass(p){return p>=100?'wellness-good':'wellness-low'}
function facilityTable(rows){let h='<div class="wellness-section-title">Facility Wise</div><div class="wellness-page"><table class="wellness-table"><thead><tr><th>Sn</th><th>NIN</th><th>Sector</th><th>Facility Name</th><th>Target Per Month /25</th><th>Total</th><th>%</th></tr></thead><tbody>';let t=0,a=0;rows.forEach(v=>{const T=num(v[4]),A=num(v[5]),P=num(v[6]);t+=T;a+=A;h+='<tr><td>'+esc(v[0])+'</td><td>'+esc(v[1])+'</td><td>'+esc(v[2])+'</td><td>'+esc(v[3])+'</td><td>'+T+'</td><td>'+A+'</td><td class="'+pctClass(P)+'">'+P+'%</td></tr>'});const p=t?Math.round(a/t*100):0;h+='<tr class="wellness-total"><td colspan="4">Total</td><td>'+t+'</td><td>'+a+'</td><td>'+p+'%</td></tr></tbody></table></div>';return h}
function sectorTable(rows){let h='<div class="wellness-section-title">Sector Wise</div><div class="wellness-page"><table class="wellness-table"><thead><tr><th>Sn</th><th>NIN</th><th>Sector</th><th>No of Facility</th><th>Target Per Month /25</th><th>Total</th><th>%</th></tr></thead><tbody>';let f=0,t=0,a=0;rows.forEach(v=>{const F=num(v[3]),T=num(v[4]),A=num(v[5]),P=num(v[6]);f+=F;t+=T;a+=A;h+='<tr><td>'+esc(v[0])+'</td><td>'+esc(v[1]||'-')+'</td><td>'+esc(v[2])+'</td><td>'+F+'</td><td>'+T+'</td><td>'+A+'</td><td class="'+pctClass(P)+'">'+P+'%</td></tr>'});const p=t?Math.round(a/t*100):0;h+='<tr class="wellness-total"><td colspan="3">Total</td><td>'+f+'</td><td>'+t+'</td><td>'+a+'</td><td>'+p+'%</td></tr></tbody></table></div>';return h}
function render(dt){css();const a=getRows(dt);const facilities=a.filter(isFacility);const sectors=a.filter(isSector);const rp=document.getElementById('reportPage');if(!rp)return;rp.querySelectorAll('.report-header,.status,.table-box,.report-actions,.footer').forEach(e=>e.style.display='none');let q=document.getElementById('wellnessActivityModule');if(!q){q=document.createElement('div');q.id='wellnessActivityModule';rp.appendChild(q)}let h='<div class="wellness-title">'+TITLE+'</div>';if(!facilities.length&&!sectors.length)h+='<div class="wellness-empty">Wellness Activity data नहीं मिला। Google Sheet से 7 columns पढ़े नहीं जा सके।</div>';else{if(facilities.length)h+=facilityTable(facilities);if(sectors.length)h+=sectorTable(sectors)}q.innerHTML=h}
function load(){if(!window.google?.visualization?.Query){setTimeout(load,300);return}const q=new google.visualization.Query('https://docs.google.com/spreadsheets/d/'+SID+'/gviz/tq?gid='+GID+'&headers=0');q.setQuery('select A,B,C,D,E,F,G');q.send(r=>{if(r.isError()){const qx=document.getElementById('wellnessActivityModule');if(qx)qx.innerHTML='<div class="wellness-title">'+TITLE+'</div><div class="wellness-empty">Google Sheet data load error: '+esc(r.getMessage())+'</div>';return}render(r.getDataTable())})}
function openWellness(){document.getElementById('dashboardPage')?.classList.remove('active');document.getElementById('reportPage')?.classList.add('active');load()}
function hook(){if(window.__wellnessHook)return;const old=window.openReport;if(typeof old!=='function')return;window.__wellnessHook=true;window.openReport=function(i){const b=document.querySelectorAll('#menuButtons .menu-btn')[i];if(b&&/wellness activity/i.test(b.textContent||'')){openWellness();return}return old.apply(this,arguments)};document.addEventListener('click',e=>{const b=e.target.closest('#menuButtons .menu-btn');if(b&&/wellness activity/i.test(b.textContent||'')){e.preventDefault();e.stopImmediatePropagation();openWellness()}},true)}
css();let tries=0;const timer=setInterval(()=>{hook();if(++tries>80)clearInterval(timer)},500);
})();