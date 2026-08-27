/* JAS LIVE FIX - locked six-column JAS layout */
(function(){
'use strict';
const SID='1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0',GID='1018164338';
const esc=v=>String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
const txt=v=>String(v??'').trim();
const num=v=>{let n=Number(String(v??'').replace(/,/g,'').replace('%','').trim());return Number.isFinite(n)?n:0};
const pc=v=>num(v)>=100?'green':'red';
function css(){if(document.getElementById('jas-live-css'))return;let s=document.createElement('style');s.id='jas-live-css';s.textContent=`
#jasMeetingModule.jas-live{width:100%}.jas-live .summary{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin:0 0 14px}.jas-live .card{border:1px solid #94a3b8;border-radius:10px;padding:10px;text-align:center;background:#fff}.jas-live .card b{display:block;font-size:20px}.jas-live .card span{font-size:11px;font-weight:700}.jas-live table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:12px;margin-bottom:16px}.jas-live th,.jas-live td{border:1px solid #334155!important;padding:7px 5px;text-align:center;vertical-align:middle;word-break:break-word}.jas-live th{background:#075985;color:#fff;font-weight:800}.jas-live .title{font-size:15px;font-weight:800;background:#e0f2fe;color:#075985}.jas-live .left{text-align:left}.jas-live .total td{font-weight:800;background:#dbeafe}.jas-live .green{color:#15803d!important;font-weight:900}.jas-live .red{color:#dc2626!important;font-weight:900}@media(max-width:800px){.jas-live .summary{grid-template-columns:repeat(2,1fr)}.jas-live table{font-size:10px}}
`;document.head.appendChild(s)}
function cell(dt,r,c){return txt(dt.getFormattedValue(r,c))}
function findHeader(dt,a,b,needle){for(let r=0;r<Math.min(dt.getNumberOfRows(),15);r++){let s='';for(let c=a;c<=b;c++)s+=' '+cell(dt,r,c).toLowerCase();if(s.includes(needle))return r}return -1}
function readFacility(dt){
  const headers=['SN','NIN','NAME','Target till July 26','Achievment','%'];
  let hr=findHeader(dt,0,5,'name');
  let start=hr>=0?hr+1:0;
  let rows=[];
  for(let r=start;r<dt.getNumberOfRows();r++){
    let v=[];for(let c=0;c<6;c++)v.push(cell(dt,r,c));
    let joined=v.join(' ').trim();
    if(!joined)continue;
    if(/jan arogya samiti|target.*july|name of facility/i.test(joined)&&!/^[0-9]+$/.test(v[0]))continue;
    if(v.every(x=>/^Column\s*\d+$/i.test(x)||x===''))continue;
    if(/^\d+$/.test(v[0])) rows.push(v);
    else if(/^total$|^योग$|^कुल$/i.test(v[0])) rows.push(v);
  }
  return {title:'Jan Arogya Samiti Meeting FY 2026-27 Till July 2026',headers,rows};
}
function table(b){
 if(!b||!b.rows.length)return '';
 let x='<table><thead><tr><th colspan="6" class="title">'+esc(b.title)+'</th></tr><tr>';
 b.headers.forEach(v=>x+='<th>'+esc(v)+'</th>');x+='</tr></thead><tbody>';
 let target=0,ach=0;
 b.rows.forEach(v=>{
   if(/^total$|^योग$|^कुल$/i.test(v[0]))return;
   target+=num(v[3]);ach+=num(v[4]);
   x+='<tr>';
   for(let i=0;i<6;i++){let z=v[i]||'';let cls='';if(i===2)cls='left';if(i===5)cls=pc(z);x+='<td class="'+cls+'">'+esc(z)+(i===5&&z&&!String(z).includes('%')?'%':'')+'</td>'}
   x+='</tr>';
 });
 let pct=target?Math.round(ach/target*100):0;
 x+='<tr class="total"><td>Total</td><td></td><td></td><td>'+target+'</td><td>'+ach+'</td><td>'+pct+'%</td></tr></tbody></table>';
 return {html:x,target,ach,pct};
}
function render(dt){
 css();let rp=document.getElementById('reportPage');if(!rp)return;
 rp.querySelectorAll('.report-header,.status,.table-box,.report-actions').forEach(e=>e.style.display='none');
 let q=document.getElementById('jasMeetingModule');if(!q){q=document.createElement('div');q.id='jasMeetingModule';q.className='jas-live';rp.appendChild(q)}
 let facility=readFacility(dt),ft=table(facility);if(!ft)return;
 q.innerHTML='<div class="summary"><div class="card"><b>'+facility.rows.filter(r=>/^\d+$/.test(r[0])).length+'</b><span>Total Facility</span></div><div class="card"><b>'+ft.target+'</b><span>Total Target</span></div><div class="card"><b>'+ft.ach+'</b><span>Total Achievement</span></div><div class="card"><b class="'+pc(ft.pct)+'">'+ft.pct+'%</b><span>Overall %</span></div><div class="card"><b>FY 2026-27</b><span>JAS Meeting</span></div></div>'+ft.html;
}
function load(){if(!window.google?.visualization?.Query)return setTimeout(load,300);let q=new google.visualization.Query('https://docs.google.com/spreadsheets/d/'+SID+'/gviz/tq?gid='+GID+'&headers=0');q.setQuery('select A,B,C,D,E,F');q.send(r=>{if(!r.isError())render(r.getDataTable());else console.error('JAS query',r.getMessage())})}
function isJAS(i){try{return window.REPORTS&&REPORTS[i]&&String(REPORTS[i].name).toLowerCase().includes('jas')}catch(e){return false}}
function hook(){if(window.__jasLiveLocked)return;const old=window.openReport;if(typeof old!=='function')return;window.__jasLiveLocked=true;window.openReport=function(i){if(isJAS(i)){document.getElementById('dashboardPage')?.classList.remove('active');document.getElementById('reportPage')?.classList.add('active');document.querySelectorAll('#menuButtons .menu-btn').forEach((b,k)=>b.classList.toggle('active',k===i));load();return}return old.apply(this,arguments)};document.addEventListener('click',e=>{let b=e.target.closest('.menu-btn');if(!b||!b.textContent.toLowerCase().includes('jas meeting'))return;e.preventDefault();e.stopImmediatePropagation();let i=Array.from(document.querySelectorAll('#menuButtons .menu-btn')).indexOf(b);if(i>=0)window.openReport(i)},true)}
css();let tries=0,tm=setInterval(()=>{hook();if(++tries>40)clearInterval(tm)},500);setTimeout(load,1200);
})();
