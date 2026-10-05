const DB='minex_rl_v9';
const APP_VERSION='26.1';
let chartLoader=null,xlsxLoader=null,bootstrapLoader=null;
const STORES=['cases','employees','volumetries','config','duplicateReviews'];
const EDITORS=['Claudia Yañez','Andres Prada','Diego Uribe'];
const state={page:'dashboard',cases:[],employees:[],duplicateReviews:[],catalogs:{},charts:{},filters:{},reportFilters:{},selectedCases:new Set(),selectedEmployees:new Set(),selectedDuplicates:new Set(),empSearch:'',followMode:'all',sort:'recent',importRows:[],importType:'cases',importMeta:null,currentEditor:localStorage.getItem('minex_editor')||'Andres Prada'};
const CASE_FIELDS=['ITEM','FECHA','MES','REQUERIMIENTO','EMP','DETALLE DE LA SOLICITUD','CEDULA','EMPLEADO / EMPRESA','ÁREA','CENTRO DE TRABAJO','SOLICITANTE','DESARROLLO DEL CASO','ÚTLIMO SEGUIMIENTO','OBSERVACIONES','PENDIENTE POR','FECHA DE SEGUIMIENTO','FECHA DE CIERRE','MES CIERRE','DIAS ACUMULADOS DEL PROCESO','ESTADO','VISIBLE_DASH','MATERNIDAD_DASH'];
const EMP_FIELDS=['EMPRESA CONTRATO','TIPO DOCUMENTO','DOCUMENTO','NOMBRE DEL EMPLEADO','CARGO','NIVEL EN LA ESTRUCTURA','JEFE INMEDIATO','SUB AREA','AREA / UNIDAD ORGANIZACIONAL','DIRECCION','GERENCIA','CLASIFICACION COSTO / GASTO','UBICACION','CLASIFICACION GENERAL','FECHA ANTIGÜEDAD','ULTIMA FECHA INGRESO','TELEFONO','CORREO'];
const DATE_FIELDS=['FECHA','FECHA DE SEGUIMIENTO','FECHA DE CIERRE','FECHA ANTIGÜEDAD','ULTIMA FECHA INGRESO'];
const esc=v=>String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]));
const fmt=n=>new Intl.NumberFormat('es-CO').format(Number(n)||0);
const todayISO=()=>{const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`};
function pad(n){return String(n).padStart(2,'0')}
function excelSerialToDate(n){if(typeof n!=='number'||!isFinite(n))return '';const d=new Date(Math.round((n-25569)*86400000));return isNaN(d)?'':`${d.getUTCFullYear()}-${pad(d.getUTCMonth()+1)}-${pad(d.getUTCDate())}`}
function parseDateValue(v){
 if(v===null||v===undefined||v==='')return '';
 if(v instanceof Date&&!isNaN(v))return `${v.getFullYear()}-${pad(v.getMonth()+1)}-${pad(v.getDate())}`;
 if(typeof v==='number')return excelSerialToDate(v)||String(v);
 const s=String(v).trim();if(!s)return '';
 let m=s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);if(m)return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
 m=s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);if(m)return `${m[3]}-${pad(m[2])}-${pad(m[1])}`;
 const d=new Date(s);return isNaN(d)?s:`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
}
function formatDate(v){const x=parseDateValue(v),m=String(x).match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?`${m[3]}/${m[2]}/${m[1]}`:String(v??'')}
function monthName(v){const m={'ENE':'Enero','FEB':'Febrero','MAR':'Marzo','ABR':'Abril','MAY':'Mayo','JUN':'Junio','JUL':'Julio','AGO':'Agosto','SEP':'Septiembre','OCT':'Octubre','NOV':'Noviembre','DIC':'Diciembre'};return m[String(v||'').toUpperCase()]||String(v||'')}
function norm(v){return String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toUpperCase()}
function canonicalField(k){const n=norm(k);const map={'CEDULA':'CEDULA','DOCUMENTO':'CEDULA','DOCUMENTO DE IDENTIDAD':'CEDULA','NUMERO DOCUMENTO':'CEDULA','NOMBRE DEL EMPLEADO':'NOMBRE DEL EMPLEADO','EMPLEADO':'EMPLEADO / EMPRESA','EMPLEADO / EMPRESA':'EMPLEADO / EMPRESA','AREA':'ÁREA','AREA / UNIDAD ORGANIZACIONAL':'AREA / UNIDAD ORGANIZACIONAL','EMPRESA':'EMP','EMPRESA CONTRATO':'EMPRESA CONTRATO','FECHA DE SEGUIMIENTO':'FECHA DE SEGUIMIENTO','FECHA SEGUIMIENTO':'FECHA DE SEGUIMIENTO','FECHA CIERRE':'FECHA DE CIERRE','DIAS ACUMULADOS':'DIAS ACUMULADOS DEL PROCESO','DIAS ACUMULADOS DEL PROCESO':'DIAS ACUMULADOS DEL PROCESO','ULTIMO SEGUIMIENTO':'ÚTLIMO SEGUIMIENTO','OBSERVACION':'OBSERVACIONES','OBSERVACIONES':'OBSERVACIONES','ESTADO':'ESTADO'};return map[n]||k}
function statusText(v){const n=norm(v);if(['1','TRUE','CERRADO','CERRADA','FINALIZADO','FINALIZADA'].includes(n))return 'CERRADO';if(['0','FALSE','ABIERTO','ABIERTA','EN SEGUIMIENTO','SEGUIMIENTO'].includes(n))return 'EN SEGUIMIENTO';return String(v??'').trim()||'SIN ESTADO'}
function badge(v){const s=statusText(v),c=s==='CERRADO'?'green':s==='EN SEGUIMIENTO'?'amber':'gray';return `<span class="badge ${c}">${esc(s)}</span>`}
function txDone(tx){return new Promise((r,j)=>{tx.oncomplete=r;tx.onerror=()=>j(tx.error||new Error('Error de almacenamiento'))})}
async function all(store){const db=await openDB();return new Promise((r,j)=>{const q=db.transaction(store).objectStore(store).getAll();q.onsuccess=()=>r(q.result);q.onerror=()=>j(q.error)})}
async function add(store,obj){const db=await openDB();return new Promise((r,j)=>{const q=db.transaction(store,'readwrite').objectStore(store).add(obj);q.onsuccess=()=>r(q.result);q.onerror=()=>j(q.error)})}
async function bulkAdd(store,rows,{clear=false}={}){if(!rows.length&&!clear)return;const db=await openDB();const tx=db.transaction(store,'readwrite'),os=tx.objectStore(store);if(clear)os.clear();for(const row of rows)os.add(row);await txDone(tx)}
async function bulkDelete(store,ids){if(!ids.length)return;const db=await openDB();const tx=db.transaction(store,'readwrite'),os=tx.objectStore(store);for(const id of ids)os.delete(id);await txDone(tx)}
async function put(store,obj){const db=await openDB();return new Promise((r,j)=>{const q=db.transaction(store,'readwrite').objectStore(store).put(obj);q.onsuccess=()=>r(q.result);q.onerror=()=>j(q.error)})}
async function del(store,id){const db=await openDB();return new Promise((r,j)=>{const q=db.transaction(store,'readwrite').objectStore(store).delete(id);q.onsuccess=r;q.onerror=()=>j(q.error)})}
async function clearStore(store){const db=await openDB();return new Promise((r,j)=>{const q=db.transaction(store,'readwrite').objectStore(store).clear();q.onsuccess=r;q.onerror=()=>j(q.error)})}
const rawAdd=add,rawPut=put,rawDel=del,rawBulkAdd=bulkAdd,rawBulkDelete=bulkDelete;
// MINEX usa IndexedDB como almacenamiento operativo local. GitHub es únicamente el medio de publicación del proyecto; no es una base de datos de MINEX.

async function storeCount(store){const db=await openDB();return new Promise((r,j)=>{const q=db.transaction(store).objectStore(store).count();q.onsuccess=()=>r(q.result);q.onerror=()=>j(q.error)})}
async function seed(){
 const cfg=await all('config');
 const initialized=cfg.some(x=>x.key==='initialized');
 if(initialized)return;
 const [c0,e0]=await Promise.all([storeCount('cases'),storeCount('employees')]);
 if(c0===0&&e0===0){
   const db=await openDB();
   const tx=db.transaction(['cases','employees','config'],'readwrite');
   const cs=tx.objectStore('cases'),es=tx.objectStore('employees'),co=tx.objectStore('config');
   for(const x of window.INIT_CASES||[])cs.add(cleanCase(x));
   for(const x of window.INIT_EMPLOYEES||[])es.add(cleanEmployee(x));
   co.add({key:'initialized',value:true});
   await txDone(tx);
 }else{
   await add('config',{key:'initialized',value:true});
 }
}
function cleanCase(x){const y={};for(const [k,v] of Object.entries(x||{})){y[k]=DATE_FIELDS.includes(k)?parseDateValue(v):v}if(y.ESTADO)y.ESTADO=statusText(y.ESTADO);if(y['DIAS ACUMULADOS DEL PROCESO']!==''&&y['DIAS ACUMULADOS DEL PROCESO']!=null)y['DIAS ACUMULADOS DEL PROCESO']=Number(y['DIAS ACUMULADOS DEL PROCESO'])||0;return y}
function nextCaseId(rows=state.cases){const nums=rows.map(x=>{const m=String(x?.ITEM??'').match(/\d+$/);return m?Number(m[0]):0}).filter(Number.isFinite);return String((nums.length?Math.max(...nums):rows.length)+1)}
function cleanEmployee(x){const y={};for(const k of EMP_FIELDS)y[k]=x[k]??x[k+' ']??'';for(const k of ['FECHA ANTIGÜEDAD','ULTIMA FECHA INGRESO'])y[k]=parseDateValue(y[k]);if(y.DOCUMENTO)y.DOCUMENTO=String(y.DOCUMENTO).trim();return y}
function loadScriptOnce(src, globalName, holder){
 if(window[globalName]) return Promise.resolve(window[globalName]);
 if(holder.promise) return holder.promise;
 holder.promise=new Promise((resolve,reject)=>{
   const tag=document.createElement('script');
   tag.src=src; tag.async=true;
   tag.onload=()=>window[globalName]?resolve(window[globalName]):reject(new Error(`No se pudo cargar ${src}`));
   tag.onerror=()=>reject(new Error(`No se pudo cargar ${src}`));
   document.head.appendChild(tag);
 });
 return holder.promise;
}
function ensureChart(){return loadScriptOnce('https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js','Chart',{get promise(){return chartLoader},set promise(v){chartLoader=v}})}
function ensureXLSX(){return loadScriptOnce('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js','XLSX',{get promise(){return xlsxLoader},set promise(v){xlsxLoader=v}})}
function ensureBootstrap(){return loadScriptOnce(`./data_bootstrap_min.js?v=${APP_VERSION}`,'__MINEX_DATA__',{get promise(){return bootstrapLoader},set promise(v){bootstrapLoader=v}})}
async function loadJson(path){
 try{
   const url=new URL(path,document.baseURI); url.searchParams.set('v',APP_VERSION);
   const r=await fetch(url.href,{cache:'force-cache'});
   if(r.ok)return await r.json();
 }catch(e){}
 try{
   await ensureBootstrap();
   const key=path.split('/').pop();
   if(window.__MINEX_DATA__&&window.__MINEX_DATA__[key]) return window.__MINEX_DATA__[key];
 }catch(e){}
 throw new Error(`No se pudo cargar ${path}. Verifica que la carpeta data esté publicada junto a index.html.`);
}
async function init(){
 try{
   const [c,e,cat]=await Promise.all([loadJson('./data/cases.json'),loadJson('./data/employees.json'),loadJson('./data/catalogs.json')]);
   window.INIT_CASES=c; window.INIT_EMPLOYEES=e; state.catalogs=cat||{};
   await seed(); await refresh();
 }catch(err){
   console.error(err);
   document.getElementById('app').innerHTML=`<div class="boot-error"><div class="card"><div class="brand-dark">MINEX</div><h1>No fue posible iniciar</h1><p>${esc(err.message)}</p><p class="muted">Verifica que index.html y la carpeta data estén publicados correctamente.</p><button class="btn primary" onclick="location.reload()">Reintentar</button></div></div>`;
 }
}
async function refresh(){const [cases,employees,duplicates]=await Promise.all([all('cases'),all('employees'),all('duplicateReviews')]);state.cases=cases;state.employees=employees;state.duplicateReviews=duplicates;render()}
function navItems(){return [['dashboard','▦','Dashboard'],['cases','▤','Casos'],['follow','◷','Seguimientos'],['upload','⇧','Cargar datos'],['reports','◫','Informes'],['employees','♙','Empleados'],['export','⇩','Exportar'],['settings','⚙','Configuración']]}
function layout(content){document.getElementById('app').innerHTML=`<div class="shell"><aside class="sidebar"><div class="brand"><span>MINEX</span><small>GESTIÓN DE RELACIONES LABORALES</small></div><div class="nav">${navItems().map(([id,ic,label])=>`<button class="${state.page===id?'active':''}" onclick="go('${id}')"><i>${ic}</i><span>${label}</span></button>`).join('')}</div><div class="side-foot"><div class="mini-brand">MINEX</div><small>GitHub Pages · MINEX</small></div></aside><main class="main"><div class="topbar"><div class="crumb">MINEX / ${pageLabel(state.page)}</div><div class="top-actions"><button class="top-search" onclick="go('cases')">⌕ <span>Buscar en la aplicación…</span><kbd>Ctrl K</kbd></button><div class="user-chip"><span class="avatar">${state.currentEditor.split(' ').map(x=>x[0]).slice(0,2).join('')}</span><div><strong>${esc(state.currentEditor)}</strong><small>Editor activo</small></div><select class="editor-mini" onchange="setEditor(this.value)">${EDITORS.map(e=>`<option ${e===state.currentEditor?'selected':''}>${esc(e)}</option>`).join('')}</select></div></div></div>${content}</main></div>`}
function pageLabel(p){return ({dashboard:'Dashboard',cases:'Casos',follow:'Seguimientos',upload:'Cargar datos',reports:'Informes',employees:'Empleados',export:'Exportar',settings:'Configuración'})[p]||p}
function header(title,sub,actions=''){return `<div class="page-head"><div><h1>${title}</h1><p>${sub}</p></div><div class="actions">${actions}</div></div>`}
async function go(p){state.page=p;state.filters={};await refresh()}
function unique(field,src=state.cases){return [...new Set(src.map(x=>String(x[field]??'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'))}
function opts(vals,selected=''){return `<option value="">Todos</option>${vals.map(v=>`<option value="${esc(v)}" ${String(v)===String(selected)?'selected':''}>${esc(v)}</option>`).join('')}`}
function kpi(label,value,sub,icon,cls=''){return `<div class="metric ${cls}"><div class="metric-top"><span>${icon}</span><small>${label}</small></div><strong>${value}</strong><p>${sub||''}</p></div>`}
function calcMetrics(){const cs=state.cases,today=todayISO(),end=weekFriday(today);const closed=cs.filter(x=>statusText(x.ESTADO)==='CERRADO').length,open=cs.length-closed;const overdue=cs.filter(x=>statusText(x.ESTADO)!=='CERRADO'&&x['FECHA DE SEGUIMIENTO']&&x['FECHA DE SEGUIMIENTO']<today).length,todayN=cs.filter(x=>statusText(x.ESTADO)!=='CERRADO'&&x['FECHA DE SEGUIMIENTO']===today).length,weekN=cs.filter(x=>statusText(x.ESTADO)!=='CERRADO'&&x['FECHA DE SEGUIMIENTO']>today&&x['FECHA DE SEGUIMIENTO']<=end).length;const avg=cs.length?Math.round(cs.reduce((a,x)=>a+(Number(x['DIAS ACUMULADOS DEL PROCESO'])||0),0)/cs.length):0;return {total:cs.length,closed,open,overdue,todayN,weekN,avg,rate:cs.length?Math.round(closed/cs.length*100):0,employees:state.employees.length}}
function dashboard(){const m=calcMetrics(),byMonth={},byArea={},byEmp={},byReq={},byResp={};state.cases.forEach(x=>{const mo=monthName(x.MES)||'Sin mes';byMonth[mo]=(byMonth[mo]||0)+1;const a=x['ÁREA']||'Sin área';byArea[a]=(byArea[a]||0)+1;const e=x.EMP||'Sin empresa';byEmp[e]=(byEmp[e]||0)+1;const r=x.REQUERIMIENTO||'Sin requerimiento';byReq[r]=(byReq[r]||0)+1;const p=x['PENDIENTE POR']||'Sin responsable';byResp[p]=(byResp[p]||0)+1});const noFollow=state.cases.filter(x=>statusText(x.ESTADO)!=='CERRADO'&&!x['FECHA DE SEGUIMIENTO']).length,dup=duplicateReport(state.cases,caseDuplicateKey).length;const recent=state.cases.slice().sort((a,b)=>String(b.FECHA).localeCompare(String(a.FECHA))).slice(0,8);layout(header('Dashboard','Visión ejecutiva de Relaciones Laborales',`<button class="btn" onclick="go('reports')">◫ Informes</button><button class="btn" onclick="go('upload')">⇧ Cargar datos</button><button class="btn primary" onclick="openCase()">＋ Nuevo caso</button>`)+`<div class="metric-grid dashboard-kpis">${kpi('Casos totales',fmt(m.total),'Base operativa','▤')}${kpi('En seguimiento',fmt(m.open),'Casos activos','◷','amber')}${kpi('Cerrados',fmt(m.closed),m.rate+'% de cierre','✓','green')}${kpi('Vencidos',fmt(m.overdue),'Requieren atención','!','red')}${kpi('Vencen hoy',fmt(m.todayN),'Seguimientos del día','◉','amber')}${kpi('Esta semana',fmt(m.weekN),'Lunes a viernes','◌','blue')}${kpi('Promedio días',fmt(m.avg),'Días acumulados','◴')}${kpi('Empleados',fmt(m.employees),'Base vigente','♙')}</div><div class="insight-strip"><div><span>Casos sin seguimiento</span><strong>${fmt(noFollow)}</strong><small>Abiertos sin fecha programada</small></div><div><span>Posibles duplicados</span><strong>${fmt(dup)}</strong><small>Revisar calidad de datos</small></div><div><span>Empresas activas</span><strong>${fmt(Object.keys(byEmp).length)}</strong><small>Con registros de casos</small></div><div><span>Áreas con gestión</span><strong>${fmt(Object.keys(byArea).length)}</strong><small>Distribución organizacional</small></div></div><div class="dash-grid"><section class="panel"><div class="panel-head"><div><h2>Evolución mensual</h2><p>Casos registrados por mes</p></div></div><div class="chart-lg"><canvas id="chartMonth"></canvas></div></section><section class="panel"><div class="panel-head"><div><h2>Estado de la gestión</h2><p>Distribución actual</p></div></div><div class="chart-lg"><canvas id="chartStatus"></canvas></div></section></div><div class="dash-grid"><section class="panel"><div class="panel-head"><div><h2>Gestión por área</h2><p>Áreas con mayor volumen</p></div><button class="link-btn" onclick="go('reports')">Ver informe →</button></div><div class="chart-md"><canvas id="chartArea"></canvas></div></section><section class="panel"><div class="panel-head"><div><h2>Requerimientos</h2><p>Distribución de solicitudes</p></div></div><div class="chart-md"><canvas id="chartReq"></canvas></div></section></div><div class="dash-grid"><section class="panel"><div class="panel-head"><div><h2>Carga por responsable</h2><p>Casos asignados o pendientes</p></div></div><div class="chart-md"><canvas id="chartResp"></canvas></div></section><section class="panel"><div class="panel-head"><div><h2>Casos por empresa</h2><p>Distribución organizacional</p></div></div><div class="chart-md"><canvas id="chartEmp"></canvas></div></section></div><section class="panel"><div class="panel-head"><div><h2>Casos recientes</h2><p>Últimos registros incorporados</p></div><button class="link-btn" onclick="go('cases')">Ver todos →</button></div>${simpleTable(recent)}</section>`);setTimeout(()=>drawDashboardCharts(byMonth,byArea,byReq,byResp,byEmp),0)}
function destroyCharts(){Object.values(state.charts).forEach(c=>c?.destroy());state.charts={}}
async function drawDashboardCharts(months,areas,req,resp,emp){try{await ensureChart()}catch(e){console.warn(e);return}destroyCharts();if(typeof Chart==='undefined')return;const colors=['#1d4ed8','#16a34a','#f59e0b','#dc2626','#7c3aed','#0f766e','#ea580c','#475569'];const mk=(id,type,labels,data,opts={})=>{const el=document.getElementById(id);if(!el)return;state.charts[id]=new Chart(el,{type,data:{labels,datasets:[{label:'Casos',data,backgroundColor:colors,borderColor:'#1d4ed8',borderWidth:2,borderRadius:6,tension:.35,fill:type==='line'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:type==='doughnut'}},scales:type==='doughnut'?{}:{y:{beginAtZero:true,ticks:{precision:0}},x:{grid:{display:false}}},...opts}})};const mkeys=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];mk('chartMonth','line',mkeys,mkeys.map(k=>months[k]||0));const mm=calcMetrics();mk('chartStatus','doughnut',['Cerrados','En seguimiento'],[mm.closed,mm.open]);const a=Object.entries(areas).sort((x,y)=>y[1]-x[1]).slice(0,10);mk('chartArea','bar',a.map(x=>x[0]),a.map(x=>x[1]),{indexAxis:'y'});const r=Object.entries(req).sort((x,y)=>y[1]-x[1]).slice(0,8);mk('chartReq','bar',r.map(x=>x[0]),r.map(x=>x[1]),{indexAxis:'y'});const p=Object.entries(resp).sort((x,y)=>y[1]-x[1]).slice(0,10);mk('chartResp','bar',p.map(x=>x[0]),p.map(x=>x[1]),{indexAxis:'y'});const e=Object.entries(emp).sort((x,y)=>y[1]-x[1]).slice(0,8);mk('chartEmp','bar',e.map(x=>x[0]),e.map(x=>x[1]),{indexAxis:'y'})}
function simpleTable(rows){if(!rows.length)return '<div class="empty">No hay registros.</div>';return `<div class="table-wrap"><table class="table"><thead><tr><th>ID</th><th>Fecha</th><th>Requerimiento</th><th>Empresa</th><th>Área</th><th>Empleado / Empresa</th><th>Estado</th><th></th></tr></thead><tbody>${rows.map(x=>`<tr><td>${esc(x.ITEM||x._id)}</td><td>${formatDate(x.FECHA)}</td><td>${esc(x.REQUERIMIENTO)}</td><td>${esc(x.EMP)}</td><td>${esc(x['ÁREA'])}</td><td>${esc(x['EMPLEADO / EMPRESA'])}</td><td>${badge(x.ESTADO)}</td><td><button class="icon-btn" onclick="editCase(${x._id})">✎</button></td></tr>`).join('')}</tbody></table></div>`}
function filterField(label,html){return `<div class="f-field"><label>${label}</label>${html}</div>`}
function selectFilter(key,label,values){return filterField(label,`<select onchange="setCaseFilter('${key}',this.value)">${opts(values,state.filters[key])}</select>`)}
function textFilter(key,label,placeholder='Filtrar…'){return filterField(label,`<input value="${esc(state.filters[key]||'')}" placeholder="${placeholder}" oninput="setCaseFilterText('${key}',this.value)">`)}
function dateFilter(key,label){return filterField(label,`<div class="date-pair"><input type="date" value="${esc(state.filters[key+'From']||'')}" onchange="setCaseFilter('${key}From',this.value)" title="Desde"><input type="date" value="${esc(state.filters[key+'To']||'')}" onchange="setCaseFilter('${key}To',this.value)" title="Hasta"></div>`)}
function casesPage(){const filtered=filteredCases();layout(header('Casos','Consulta, seguimiento, edición y gestión de casos',`<button class="btn" onclick="clearCaseFilters()">↺ Limpiar filtros</button><button class="btn" onclick="exportCasesFiltered()">⇩ Exportar</button><button class="btn" onclick="exportSelectedCases()">⇩ Seleccionados</button><button class="btn danger-outline" onclick="deleteSelectedCases()">🗑 Borrar seleccionados</button>${state.duplicateReviews.length?`<button class="btn" onclick="duplicateReviewPage()">⚠ Duplicados (${state.duplicateReviews.length})</button>`:''}<button class="btn primary" onclick="openCase()">＋ Nuevo caso</button>`)+`<section class="filter-card"><div class="filter-card-head"><div class="filter-title"><span class="filter-icon">⌕</span><div><strong>Filtros de búsqueda</strong><small>Combina uno o varios criterios para encontrar exactamente la información que necesitas</small></div></div><button class="link-btn" onclick="toggleAdvancedFilters()">⌃ Ocultar filtros</button></div><div class="filter-groups"><div class="filter-group"><div class="group-title blue">▤ Información general</div>${textFilter('ITEM','Item / ID','Buscar ID…')}${selectFilter('REQUERIMIENTO','Requerimiento',unique('REQUERIMIENTO'))}${selectFilter('ESTADO','Estado',['CERRADO','EN SEGUIMIENTO'])}${textFilter('OBSERVACIONES','Observaciones','Buscar observación…')}</div><div class="filter-group"><div class="group-title green">◷ Fechas</div>${dateFilter('FECHA','Fecha de radicación')}${dateFilter('FECHA DE SEGUIMIENTO','Fecha de seguimiento')}${dateFilter('FECHA DE CIERRE','Fecha de cierre')}</div><div class="filter-group"><div class="group-title purple">▦ Organización</div>${selectFilter('EMP','Empresa (EMP)',unique('EMP'))}${selectFilter('ÁREA','Área',unique('ÁREA'))}${selectFilter('CENTRO DE TRABAJO','Centro de trabajo',unique('CENTRO DE TRABAJO'))}${textFilter('SOLICITANTE','Solicitante','Buscar solicitante…')}</div><div class="filter-group"><div class="group-title orange">♙ Persona / caso</div>${textFilter('CEDULA','Cédula','Buscar cédula…')}${textFilter('EMPLEADO / EMPRESA','Empleado / Empresa','Buscar empleado o empresa…')}${textFilter('DETALLE DE LA SOLICITUD','Detalle de la solicitud','Buscar palabras clave…')}${textFilter('PENDIENTE POR','Pendiente por','Buscar responsable…')}<div class="f-label">Días acumulados</div><div class="date-pair"><input type="number" placeholder="Mín" value="${esc(state.filters.dmin||'')}" onchange="setCaseFilter('dmin',this.value)"><input type="number" placeholder="Máx" value="${esc(state.filters.dmax||'')}" onchange="setCaseFilter('dmax',this.value)"></div></div></div><div class="filter-foot"><span id="caseFilterCount">${fmt(filtered.length)} registros encontrados</span><input class="general-search" value="${esc(state.filters.q||'')}" placeholder="Buscar en cualquier columna…" oninput="setCaseFilterText('q',this.value)"></div></section><section id="casesResults" class="results-card"><div class="results-head"><div><strong>${fmt(filtered.length)} casos encontrados</strong><small>Mostrando resultados según los filtros aplicados</small></div><div class="results-tools"><select onchange="state.sort=this.value;casesPage()"><option value="recent" ${state.sort==='recent'?'selected':''}>Fecha (más reciente)</option><option value="old" ${state.sort==='old'?'selected':''}>Fecha (más antigua)</option><option value="follow" ${state.sort==='follow'?'selected':''}>Seguimiento próximo</option><option value="days" ${state.sort==='days'?'selected':''}>Más días acumulados</option></select></div></div>${casesTable(filtered)}</section>`)}
function toggleCaseSelection(id,checked){checked?state.selectedCases.add(id):state.selectedCases.delete(id);casesPage()}
function toggleAllCases(checked){filteredCases().forEach(x=>checked?state.selectedCases.add(x._id):state.selectedCases.delete(x._id));casesPage()}
function selectedCaseRows(){return state.cases.filter(x=>state.selectedCases.has(x._id))}
function exportSelectedCases(){const rows=selectedCaseRows();if(!rows.length)return toast('Selecciona al menos un caso');downloadRows(rows,'CASOS_SELECCIONADOS_RELACIONES_LABORALES')}
async function deleteSelectedCases(){const rows=selectedCaseRows();if(!rows.length)return toast('Selecciona al menos un caso');if(!confirm(`¿Eliminar ${rows.length} caso(s) seleccionado(s)? Esta acción no se puede deshacer.`))return;await bulkDelete('cases',rows.map(x=>x._id));state.selectedCases.clear();toast(`${rows.length} caso(s) eliminado(s)`);refresh()}

function filteredCases(){let rows=[...state.cases];const f=state.filters;for(const [k,v] of Object.entries(f)){if(!v||['q','dmin','dmax'].includes(k)||k.endsWith('From')||k.endsWith('To'))continue;rows=rows.filter(x=>norm(x[k]).includes(norm(v)))}if(f.q){const q=norm(f.q);rows=rows.filter(x=>Object.values(x).some(v=>norm(v).includes(q)))}if(f.dmin)rows=rows.filter(x=>(Number(x['DIAS ACUMULADOS DEL PROCESO'])||0)>=Number(f.dmin));if(f.dmax)rows=rows.filter(x=>(Number(x['DIAS ACUMULADOS DEL PROCESO'])||0)<=Number(f.dmax));for(const key of ['FECHA','FECHA DE SEGUIMIENTO','FECHA DE CIERRE']){if(f[key+'From'])rows=rows.filter(x=>parseDateValue(x[key])>=f[key+'From']);if(f[key+'To'])rows=rows.filter(x=>parseDateValue(x[key])<=f[key+'To'])}if(state.sort==='old')rows.sort((a,b)=>String(a.FECHA).localeCompare(String(b.FECHA)));else if(state.sort==='follow')rows.sort((a,b)=>String(a['FECHA DE SEGUIMIENTO']).localeCompare(String(b['FECHA DE SEGUIMIENTO'])));else if(state.sort==='days')rows.sort((a,b)=>(Number(b['DIAS ACUMULADOS DEL PROCESO'])||0)-(Number(a['DIAS ACUMULADOS DEL PROCESO'])||0));else rows.sort((a,b)=>String(b.FECHA).localeCompare(String(a.FECHA)));return rows}
function casesTable(rows){if(!rows.length)return '<div class="empty">No hay casos con los filtros seleccionados.</div>';return `<div class="table-wrap"><table class="table wide-table cases-main-table"><thead><tr><th class="select-col"><input type="checkbox" title="Seleccionar todos" onchange="toggleAllCases(this.checked)"></th><th>ID</th><th>Fecha</th><th>Requerimiento</th><th>EMP</th><th>Área</th><th class="detail-head">Detalle de la solicitud</th><th>Empleado / Empresa</th><th>Seguimiento</th><th>Días</th><th>Estado</th><th>Acciones</th></tr></thead><tbody>${rows.slice(0,1000).map(x=>`<tr><td class="select-col"><input type="checkbox" ${state.selectedCases.has(x._id)?'checked':''} onchange="toggleCaseSelection(${x._id},this.checked)"></td><td><strong>${esc(x.ITEM||x._id)}</strong></td><td>${formatDate(x.FECHA)}</td><td>${esc(x.REQUERIMIENTO)}</td><td>${esc(x.EMP)}</td><td>${esc(x['ÁREA'])}</td><td class="detail-cell" title="${esc(x['DETALLE DE LA SOLICITUD']||'')}">${esc(x['DETALLE DE LA SOLICITUD']||'—')}</td><td>${esc(x['EMPLEADO / EMPRESA'])}</td><td>${formatDate(x['FECHA DE SEGUIMIENTO'])||'—'}</td><td>${fmt(x['DIAS ACUMULADOS DEL PROCESO'])}</td><td>${badge(x.ESTADO)}</td><td><div class="row-actions"><button class="icon-btn" title="Ver / editar" onclick="editCase(${x._id})">✎</button><button class="icon-btn danger-icon" title="Eliminar" onclick="deleteCase(${x._id})">⌫</button></div></td></tr>`).join('')}</tbody></table></div>`}
function setCaseFilter(k,v){state.filters[k]=v;casesPage()}
function setCaseFilterText(k,v){state.filters[k]=v;updateCasesResults()}
function updateCasesResults(){
  const rows=filteredCases();
  const count=document.getElementById('caseFilterCount');
  if(count) count.textContent=`${fmt(rows.length)} registros encontrados`;
  const box=document.getElementById('casesResults');
  if(!box)return;
  box.innerHTML=`<div class="results-head"><div><strong>${fmt(rows.length)} casos encontrados</strong><small>Mostrando resultados según los filtros aplicados</small></div><div class="results-tools"><select onchange="state.sort=this.value;casesPage()"><option value="recent" ${state.sort==='recent'?'selected':''}>Fecha (más reciente)</option><option value="old" ${state.sort==='old'?'selected':''}>Fecha (más antigua)</option><option value="follow" ${state.sort==='follow'?'selected':''}>Seguimiento próximo</option><option value="days" ${state.sort==='days'?'selected':''}>Más días acumulados</option></select></div></div>${casesTable(rows)}`;
}
function clearCaseFilters(){state.filters={};casesPage()}
function toggleAdvancedFilters(){const g=document.querySelector('.filter-groups');if(g){g.classList.toggle('collapsed');document.querySelector('.filter-card .link-btn').textContent=g.classList.contains('collapsed')?'⌄ Mostrar filtros':'⌃ Ocultar filtros'}}
function openCase(existing=null){const x=existing||{};document.body.insertAdjacentHTML('beforeend',`<div class="modal-back" id="caseModal"><div class="modal xwide"><div class="modal-head"><div><h2>${existing?'Editar caso':'Nuevo caso'}</h2><p>${existing?'Actualiza información, fechas, observaciones y seguimiento.':'Registra una nueva gestión de Relaciones Laborales.'}</p></div><button class="icon-btn" onclick="closeModal('caseModal')">×</button></div><form id="caseForm">${caseForm(x)}<div class="modal-actions"><button type="button" class="btn" onclick="closeModal('caseModal')">Cancelar</button><button class="btn primary">${existing?'Guardar cambios':'Crear caso'}</button></div></form></div></div>`);bindCaseForm(existing?existing._id:null, existing?existing._reviewId:null)}
function editCase(id){const x=state.cases.find(y=>y._id===id);if(x)openCase(x)}
function fieldInput(name,label,value='',type='text',attrs=''){
  const id='f_'+String(name).replace(/[^A-Za-z0-9_]/g,'_');
  return `<div class="field"><label for="${esc(id)}">${esc(label)}</label><input id="${esc(id)}" type="${esc(type)}" name="${esc(name)}" value="${esc(value)}" ${attrs||''}></div>`;
}
function fieldTextarea(name,label,value=''){
  const id='t_'+String(name).replace(/[^A-Za-z0-9_]/g,'_');
  return `<div class="field"><label for="${esc(id)}">${esc(label)}</label><textarea id="${esc(id)}" name="${esc(name)}">${esc(value)}</textarea></div>`;
}
function fieldSelect(name,label,options=[],value=''){
  const id='s_'+String(name).replace(/[^A-Za-z0-9_]/g,'_');
  const opts=Array.from(new Set((options||[]).map(v=>String(v??'').trim()).filter(Boolean)));
  if(value!==undefined && value!==null && String(value).trim() && !opts.includes(String(value).trim())) opts.unshift(String(value).trim());
  return `<div class="field"><label for="${esc(id)}">${esc(label)}</label><select id="${esc(id)}" name="${esc(name)}"><option value="">Seleccionar...</option>${opts.map(v=>`<option value="${esc(v)}" ${String(v)===String(value??'')?'selected':''}>${esc(v)}</option>`).join('')}</select></div>`;
}
function caseForm(x){
 const isEdit=!!x._id;
 const hist=Array.isArray(x.SEGUIMIENTOS_HISTORIAL)?x.SEGUIMIENTOS_HISTORIAL:[];
 const last=hist[hist.length-1];
 const lastText=x['ÚTLIMO SEGUIMIENTO']||'';
 const editorOptions=EDITORS.map(e=>`<option value="${esc(e)}" ${e===state.currentEditor?'selected':''}>${esc(e)}</option>`).join('');
 return `<div class="formgrid corporate-form">
 <div class="form-section"><div class="form-section-title">Información general</div>${fieldInput('ITEM','ID (automático)',x.ITEM||nextCaseId(),'text','readonly tabindex="-1" title="Asignado automáticamente"')}${fieldInput('FECHA','Fecha',parseDateValue(x.FECHA),'date')}${fieldInput('MES','Mes',x.MES,'text')}${fieldSelect('REQUERIMIENTO','Requerimiento',state.catalogs.REQUERIMIENTO||unique('REQUERIMIENTO'),x.REQUERIMIENTO)}${fieldSelect('EMP','Empresa',state.catalogs.EMP||unique('EMP'),x.EMP)}${fieldSelect('ÁREA','Área',state.catalogs['ÁREA']||unique('ÁREA'),x['ÁREA'])}</div>
 <div class="form-section"><div class="form-section-title">Persona / organización</div>${fieldInput('CEDULA','Cédula',x.CEDULA,'text','onchange="autofillEmployee(this.value)"')}${fieldInput('EMPLEADO / EMPRESA','Empleado / Empresa',x['EMPLEADO / EMPRESA'],'text','id="caseEmployee"')}${fieldSelect('CENTRO DE TRABAJO','Centro de trabajo',unique('CENTRO DE TRABAJO'),x['CENTRO DE TRABAJO'])}${fieldInput('SOLICITANTE','Solicitante',x.SOLICITANTE,'text')}${fieldInput('PENDIENTE POR','Pendiente por',x['PENDIENTE POR'],'text')}</div>
 <div class="form-section full"><div class="form-section-title">Gestión y seguimiento</div>${fieldTextarea('DETALLE DE LA SOLICITUD','Detalle de la solicitud',x['DETALLE DE LA SOLICITUD'])}${fieldTextarea('DESARROLLO DEL CASO','Desarrollo del caso',x['DESARROLLO DEL CASO'])}
 <div class="follow-editor-card">
   <div class="follow-current"><div class="follow-current-head"><span>Último seguimiento</span><small>${last?`${esc(last.editor||'')} · ${formatDate(last.fecha)}`:'Sin seguimiento registrado'}</small></div><div class="follow-current-text">${lastText?esc(lastText):'Aún no hay un seguimiento registrado.'}</div></div>
   <div class="follow-new">
     <div class="follow-new-head"><div><strong>${isEdit?'Registrar / editar seguimiento':'Registrar seguimiento'}</strong><small>El sistema identifica automáticamente el cambio de día y conserva el histórico.</small></div></div>
     <div class="follow-new-grid"><div class="field"><label>Editor del seguimiento <span class="required">*</span></label><select name="EDITOR_SEGUIMIENTO" required><option value="">Seleccionar editor…</option>${editorOptions}</select></div><div class="field"><label>Nuevo comentario</label><textarea name="NUEVO_COMENTARIO" placeholder="Escribe el avance o comentario de esta gestión…"></textarea></div></div>
   </div>
 </div>

 <div class="threecol">${fieldTextarea('OBSERVACIONES','Observaciones',x.OBSERVACIONES)}${fieldTextarea('PRÓXIMO PASO','Próximo paso',x['PRÓXIMO PASO'])}${fieldInput('FECHA DE SEGUIMIENTO','Fecha de seguimiento',parseDateValue(x['FECHA DE SEGUIMIENTO']),'date')}</div>
 <div class="fourcol derived-case-fields"><div class="field"><label>Fecha de cierre</label><input type="date" name="FECHA DE CIERRE" value="${esc(parseDateValue(x['FECHA DE CIERRE']))}" onchange="updateCaseDerivedFields(this.form)" oninput="updateCaseDerivedFields(this.form)"></div><div class="field readonly-field"><label>Mes cierre <small>automático</small></label><input type="text" name="MES CIERRE" value="${esc(x['MES CIERRE']||'')}" readonly tabindex="-1"></div><div class="field readonly-field"><label>Días acumulados <small>automático</small></label><input type="number" name="DIAS ACUMULADOS DEL PROCESO" value="${esc(x['DIAS ACUMULADOS DEL PROCESO']||0)}" readonly tabindex="-1"></div><div class="field readonly-field"><label>Estado <small>automático</small></label><input type="text" name="ESTADO" value="${esc(statusText(x.ESTADO))}" readonly tabindex="-1"></div></div>
 <div class="fourcol">${fieldInput('VISIBLE_DASH','Visible dashboard',statusText(x.VISIBLE_DASH),'text')}${fieldInput('MATERNIDAD_DASH','Maternidad dashboard',statusText(x.MATERNIDAD_DASH),'text')}</div>
 </div></div>`
}
function autoMonthAbbr(iso){const x=parseDateValue(iso);if(!x)return '';const m=String(x).slice(5,7);return ({'01':'ene','02':'feb','03':'mar','04':'abr','05':'may','06':'jun','07':'jul','08':'ago','09':'sep','10':'oct','11':'nov','12':'dic'})[m]||''}
function inclusiveDays(start,end){const a=parseDateValue(start),b=parseDateValue(end);if(!a)return 0;const d1=new Date(a+'T12:00:00'),d2=new Date((b||todayISO())+'T12:00:00');const diff=Math.round((d2-d1)/86400000)+1;return Math.max(0,diff)}
function updateCaseDerivedFields(form){if(!form)return;const close=form.querySelector('[name="FECHA DE CIERRE"]')?.value||'';const start=form.querySelector('[name="FECHA"]')?.value||'';const month=form.querySelector('[name="MES CIERRE"]');const days=form.querySelector('[name="DIAS ACUMULADOS DEL PROCESO"]');const status=form.querySelector('[name="ESTADO"]');if(month)month.value=close?autoMonthAbbr(close):'';if(status)status.value=close?'CERRADO':'EN SEGUIMIENTO';if(days)days.value=inclusiveDays(start,close);}
function bindCaseForm(existingId, reviewId=null){
 const form=document.getElementById('caseForm');
 updateCaseDerivedFields(form);
 form.onsubmit=async e=>{
  e.preventDefault();
  const d=Object.fromEntries(new FormData(form).entries());
  const comment=(d.NUEVO_COMENTARIO||'').trim();
  const editor=(d.EDITOR_SEGUIMIENTO||'').trim();
  delete d.NUEVO_COMENTARIO; delete d.EDITOR_SEGUIMIENTO;
  for(const k of ['FECHA','FECHA DE SEGUIMIENTO','FECHA DE CIERRE']) d[k]=parseDateValue(d[k]);
  d.ESTADO=d['FECHA DE CIERRE']?'CERRADO':'EN SEGUIMIENTO';
  d['MES CIERRE']=d['FECHA DE CIERRE']?autoMonthAbbr(d['FECHA DE CIERRE']):'';
  d['DIAS ACUMULADOS DEL PROCESO']=inclusiveDays(d.FECHA,d['FECHA DE CIERRE']);
  const old=existingId?state.cases.find(x=>x._id===existingId):null;
  if(comment){
    if(!editor){toast('Selecciona quién registra el seguimiento');return}
    const history=Array.isArray(old?.SEGUIMIENTOS_HISTORIAL)?old.SEGUIMIENTOS_HISTORIAL.map(v=>({...v})):[];
    const today=todayISO();
    let last=history[history.length-1];
    if(!last && old && old['ÚTLIMO SEGUIMIENTO']){
      last={fecha:parseDateValue(old['FECHA DE SEGUIMIENTO'])||parseDateValue(old.FECHA)||today,editor:'Registro anterior',comentario:old['ÚTLIMO SEGUIMIENTO']};
      history.push(last);
    }
    if(last && last.fecha!==today){
      const archive=`${formatDate(last.fecha)} — ${last.editor||'Registro anterior'}: ${last.comentario||''}`;
      const currentDev=String(old?.['DESARROLLO DEL CASO']||d['DESARROLLO DEL CASO']||'').trim();
      const lines=currentDev?currentDev.split(/\n/).map(v=>v.trim()).filter(Boolean):[];
      if(!lines.some(v=>norm(v)===norm(archive))) lines.push(archive);
      d['DESARROLLO DEL CASO']=lines.join('\n');
      history.push({fecha:today,editor,comentario:comment});
    }else if(last && last.fecha===today){
      last.comentario = last.comentario ? `${last.comentario}\n[${editor}] ${comment}` : comment;
      last.editor = last.editor===editor ? editor : `${last.editor||''} / ${editor}`.replace(/^ \/ /,'');
      last.fecha=today;
    }else{
      history.push({fecha:today,editor,comentario:comment});
    }
    d.SEGUIMIENTOS_HISTORIAL=history;
    d['ÚTLIMO SEGUIMIENTO']=`[${editor} · ${formatDate(today)}] ${comment}`;
    d['EDITOR ÚLTIMO SEGUIMIENTO']=editor;
  }else if(old){
    d.SEGUIMIENTOS_HISTORIAL=old.SEGUIMIENTOS_HISTORIAL||[];
    d['EDITOR ÚLTIMO SEGUIMIENTO']=old['EDITOR ÚLTIMO SEGUIMIENTO']||'';
  }
  d['DIAS ACUMULADOS DEL PROCESO']=inclusiveDays(d.FECHA,d['FECHA DE CIERRE']);
  d['ESTADO']=d['FECHA DE CIERRE']?'CERRADO':'EN SEGUIMIENTO';
  d['MES CIERRE']=d['FECHA DE CIERRE']?autoMonthAbbr(d['FECHA DE CIERRE']):'';
  if(!d.FECHA)d.FECHA=todayISO();
  if(!existingId)d.ITEM=nextCaseId(state.cases);
  if(existingId){await put('cases',{...old,...d});toast('Caso actualizado correctamente')}
  else if(reviewId){delete d._reviewId;await add('cases',d);await del('duplicateReviews',reviewId);toast('Duplicado revisado e incorporado a la base maestra')}
  else{const dup=state.cases.find(x=>caseDuplicateKey(x)===caseDuplicateKey(d));if(dup&&!confirm(`Posible duplicado detectado (ID ${dup.ITEM||dup._id}). ¿Deseas crear de todas formas?`))return;await add('cases',d);toast('Caso creado correctamente')}
  closeModal('caseModal');await refresh()
 }
}
function setEditor(v){if(!EDITORS.includes(v))return;state.currentEditor=v;localStorage.setItem('minex_editor',v);toast(`Editor activo: ${v}`);render()}
function renderFollowHistory(x){const h=Array.isArray(x.SEGUIMIENTOS_HISTORIAL)?x.SEGUIMIENTOS_HISTORIAL:[];if(!h.length)return '<div class="history-empty">Sin historial de seguimientos registrados.</div>';return `<div class="history-box"><strong>Historial de seguimientos</strong>${h.slice().reverse().map(item=>`<div class="history-item"><div><strong>${esc(item.editor||'Sin editor')}</strong><small>${formatDate(item.fecha)}</small></div><p>${esc(item.comentario||'')}</p></div>`).join('')}</div>`}
function caseDuplicateKey(x){return [parseDateValue(x.FECHA),norm(x.EMP),norm(x.REQUERIMIENTO),norm(x.CEDULA||x['EMPLEADO / EMPRESA']),norm(x['DETALLE DE LA SOLICITUD'])].join('|')}
function autofillEmployee(doc){const e=state.employees.find(x=>norm(x.DOCUMENTO)===norm(doc));if(!e)return;const f=document.getElementById('caseEmployee');if(f)f.value=e['NOMBRE DEL EMPLEADO']||'';const area=document.querySelector('#caseModal select[name="ÁREA"]');if(area&&e['AREA / UNIDAD ORGANIZACIONAL']){const candidate=[...area.options].find(o=>norm(o.value)===norm(e['AREA / UNIDAD ORGANIZACIONAL']));if(candidate)area.value=candidate.value}toast('Datos del empleado encontrados')}
async function deleteCase(id){if(confirm('¿Eliminar este caso? Esta acción no se puede deshacer.')){await del('cases',id);toast('Caso eliminado');refresh()}}
function closeModal(id){document.getElementById(id)?.remove()}
function followPage(){const today=todayISO(),end=weekFriday(today),open=state.cases.filter(x=>statusText(x.ESTADO)!=='CERRADO'&&x['FECHA DE SEGUIMIENTO']);const overdue=open.filter(x=>x['FECHA DE SEGUIMIENTO']<today),todayRows=open.filter(x=>x['FECHA DE SEGUIMIENTO']===today),week=open.filter(x=>x['FECHA DE SEGUIMIENTO']>today&&x['FECHA DE SEGUIMIENTO']<=end);let sections='';if(state.followMode==='all'||state.followMode==='overdue')sections+=followSection('Vencidos',overdue,'danger','Casos con fecha de seguimiento anterior a hoy.');if(state.followMode==='all'||state.followMode==='today')sections+=followSection('Vencen hoy',todayRows,'warning','Casos que requieren gestión durante la jornada de hoy.');if(state.followMode==='all'||state.followMode==='week')sections+=followSection('Esta semana · lunes a viernes',week,'info','Seguimientos programados para el resto de la semana laboral.');layout(header('Seguimientos','Control de compromisos, vencimientos y próximas gestiones',`<button class="btn" onclick="go('cases')">Ver casos</button><button class="btn primary" onclick="openCase()">＋ Nuevo caso</button>`)+`<div class="follow-summary"><button class="follow-kpi neutral ${state.followMode==='all'?'selected':''}" onclick="state.followMode='all';followPage()"><span>Todos</span><strong>${fmt(overdue.length+todayRows.length+week.length)}</strong><small>Vista completa</small></button><button class="follow-kpi danger ${state.followMode==='overdue'?'selected':''}" onclick="state.followMode='overdue';followPage()"><span>Vencidos</span><strong>${fmt(overdue.length)}</strong><small>Requieren atención</small></button><button class="follow-kpi warning ${state.followMode==='today'?'selected':''}" onclick="state.followMode='today';followPage()"><span>Vencen hoy</span><strong>${fmt(todayRows.length)}</strong><small>${formatDate(today)}</small></button><button class="follow-kpi info ${state.followMode==='week'?'selected':''}" onclick="state.followMode='week';followPage()"><span>Esta semana</span><strong>${fmt(week.length)}</strong><small>Hasta ${formatDate(end)}</small></button></div>${sections}`)}
function weekFriday(iso){const d=new Date(iso+'T12:00:00');const day=d.getDay()||7;const diff=5-day;const f=new Date(d);f.setDate(d.getDate()+diff);return `${f.getFullYear()}-${pad(f.getMonth()+1)}-${pad(f.getDate())}`}
function followSection(title,rows,kind,sub){return `<section class="follow-section"><div class="follow-head"><div><h2>${title}</h2><p>${sub}</p></div><span class="count-pill ${kind}">${rows.length}</span></div>${rows.length?`<div class="table-wrap"><table class="table follow-table"><thead><tr><th>ID</th><th>Fecha de seguimiento</th><th>Días</th><th>Requerimiento</th><th>Empresa</th><th>Área</th><th>Detalle de la solicitud</th><th>Empleado / Empresa</th><th>Pendiente por</th><th>Último comentario / editor</th><th>Estado</th><th></th></tr></thead><tbody>${rows.sort((a,b)=>String(a['FECHA DE SEGUIMIENTO']).localeCompare(String(b['FECHA DE SEGUIMIENTO']))).map(x=>{const h=Array.isArray(x.SEGUIMIENTOS_HISTORIAL)?x.SEGUIMIENTOS_HISTORIAL:[];const last=h[h.length-1];return `<tr><td>${esc(x.ITEM||x._id)}</td><td><strong>${formatDate(x['FECHA DE SEGUIMIENTO'])}</strong></td><td>${fmt(x['DIAS ACUMULADOS DEL PROCESO'])}</td><td>${esc(x.REQUERIMIENTO)}</td><td>${esc(x.EMP)}</td><td>${esc(x['ÁREA'])}</td><td class="detail-cell" title="${esc(x['DETALLE DE LA SOLICITUD']||'')}">${esc(x['DETALLE DE LA SOLICITUD']||'—')}</td><td>${esc(x['EMPLEADO / EMPRESA'])}</td><td>${esc(x['PENDIENTE POR'])}</td><td>${last?`<strong>${esc(last.editor)}</strong><br><small>${esc(last.comentario)}</small>`:'—'}</td><td>${badge(x.ESTADO)}</td><td><button class="icon-btn" onclick="editCase(${x._id})">✎</button></td></tr>`}).join('')}</tbody></table></div>`:'<div class="empty">No hay seguimientos en esta categoría.</div>'}</section>`}
function uploadPage(){const meta=state.importMeta;layout(header('Cargar datos','Importación controlada de Excel y CSV con validación, duplicados y vista previa',`<button class="btn" onclick="${state.importType==='employees'?'downloadEmployeeTemplate()':'downloadCaseTemplate()'}">⇩ ${state.importType==='employees'?'Formato empleados':'Formato oficial de casos'}</button>`)+`<div class="upload-grid"><section class="panel"><div class="panel-head"><div><h2>1. Selecciona el tipo de información</h2><p>La estructura se valida antes de guardar.</p></div></div><div class="type-switch"><button class="${state.importType==='cases'?'selected':''}" onclick="state.importType='cases';uploadPage()">Casos</button><button class="${state.importType==='employees'?'selected':''}" onclick="state.importType='employees';uploadPage()">Empleados</button></div>${state.importType==='employees'?'<div class="notice warning-notice"><strong>Actualización de empleados:</strong> cada carga reemplazará completamente la base anterior y dejará únicamente la información del archivo más reciente. Se validarán documentos duplicados dentro del archivo antes de confirmar.</div>':''}<div class="drop" id="dropZone" onclick="document.getElementById('fileInput').click()"><div class="drop-icon">⇧</div><strong>Selecciona o arrastra tu archivo</strong><p>Excel .xlsx / .xls o CSV</p><input id="fileInput" type="file" accept=".xlsx,.xls,.csv" hidden onchange="previewImport(this.files[0])"></div><div class="notice"><strong>Flujo:</strong> seleccionar → analizar → detectar duplicados → revisar errores → confirmar importación.</div></section><section class="panel"><div class="panel-head"><div><h2>2. Resultado de validación</h2><p>${meta?`Archivo: ${esc(meta.name)}`:'Aún no se ha cargado un archivo.'}</p></div></div>${meta?importSummary(meta):'<div class="empty">La vista previa aparecerá aquí.</div>'}</section></div>${meta?importPreview():''}`)}
function importSummary(m){return `<div class="validation-grid"><div><span>Registros</span><strong>${fmt(m.total)}</strong></div><div class="ok"><span>Válidos</span><strong>${fmt(m.valid)}</strong></div><div class="warn"><span>Duplicados</span><strong>${fmt(m.duplicates)}</strong></div><div class="bad"><span>Errores</span><strong>${fmt(m.errors)}</strong></div></div><div class="progress"><span style="width:${m.total?Math.round(m.valid/m.total*100):0}%"></span></div>`}
function importPreview(){
 const rows=state.importRows||[];
 const valid=rows.filter(r=>r._importStatus==='ok').length;
 const hasErrors=rows.some(r=>r._importStatus==='error');
 const rawHeaders=Object.keys(rows[0]||{}).filter(k=>k!=='_importStatus');const headers=state.importType==='cases'?Array.from(new Set(['ITEM','FECHA','REQUERIMIENTO','EMP','ÁREA','DETALLE DE LA SOLICITUD','CEDULA','EMPLEADO / EMPRESA','FECHA DE SEGUIMIENTO','ESTADO'].filter(k=>rawHeaders.includes(k)))).slice(0,10):rawHeaders.slice(0,10);
 const body=rows.slice(0,30).map((r,i)=>`<tr><td>${i+1}</td><td>${r._importStatus==='duplicate-review'?'<span class="badge amber">Duplicado · revisión</span>':r._importStatus==='error'?'<span class="badge red">Error</span>':'<span class="badge green">Listo</span>'}</td>${headers.map(k=>`<td>${DATE_FIELDS.includes(k)?formatDate(r[k]):esc(r[k])}</td>`).join('')}</tr>`).join('');
 return `<section class="panel"><div class="panel-head"><div><h2>3. Vista previa</h2><p>Se muestran los primeros registros para revisión. Los duplicados se enviarán a una bandeja de revisión para decidir si se editan, se eliminan o se incorporan.</p></div><div class="actions"><button class="btn" onclick="state.importRows=[];state.importMeta=null;uploadPage()">Cancelar</button><button class="btn primary" ${valid===0?'disabled':''} onclick="confirmImport()">✓ Confirmar importación (${valid} válidos)</button></div></div>${hasErrors?'<div class="notice warning-notice"><strong>Hay registros con errores.</strong> Puedes continuar: los registros válidos se importarán y los que tengan error serán omitidos.</div>':''}<div class="table-wrap"><table class="table"><thead><tr><th>#</th><th>Resultado</th>${headers.map(k=>`<th>${esc(k)}</th>`).join('')}</tr></thead><tbody>${body}</tbody></table></div></section>`;
}

async function previewImport(file){if(!file)return;try{await ensureXLSX();const data=await file.arrayBuffer(),wb=XLSX.read(data,{type:'array',cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]],raw=XLSX.utils.sheet_to_json(ws,{defval:'',raw:true});if(!raw.length)throw new Error('El archivo no contiene registros.');const rows=raw.map(r=>state.importType==='employees'?mapEmployeeImportRow(r):mapImportRow(r));const existing=state.importType==='employees'?[]:(state.importType==='vols'?state.vols:state.cases);const keys=new Set(existing.map(x=>state.importType==='vols'?caseDuplicateKey(x):caseDuplicateKey(x)));let duplicates=0,errors=0;for(const r of rows){const k=state.importType==='employees'?employeeKey(r):caseDuplicateKey(r);if(k&&keys.has(k)){r._importStatus='duplicate-review';duplicates++}else{if(k)keys.add(k);r._importStatus='ok'}if(state.importType==='employees'&&!r.DOCUMENTO){r._importStatus='error';errors++}if(state.importType!=='employees'&&!r.FECHA){r._importStatus='error';errors++}}state.importRows=rows;state.importMeta={name:file.name,total:rows.length,valid:rows.filter(r=>r._importStatus==='ok').length,duplicates,errors,replacesEmployees:state.importType==='employees'};uploadPage()}catch(e){toast('No se pudo analizar el archivo: '+e.message)}
}
function mapEmployeeImportRow(r){const y={};const aliases={};EMP_FIELDS.forEach(f=>aliases[norm(f).replace(/\s+/g,' ')]=f);const extra={'DIRECCION':'DIRECCION','DIRECCION ':'DIRECCION','AREA/UNIDAD ORGANIZACIONAL':'AREA / UNIDAD ORGANIZACIONAL','EMPRESA CONTRATO':'EMPRESA CONTRATO','CLASIFICACION COSTO/GASTO':'CLASIFICACION COSTO / GASTO','ULTIMA FECHA DE INGRESO':'ULTIMA FECHA INGRESO','FECHA DE ANTIGUEDAD':'FECHA ANTIGÜEDAD'};Object.assign(aliases,extra);for(const [k,v] of Object.entries(r)){const n=norm(k).replace(/\s+/g,' ');const ck=aliases[n]||k.replace(/\u00a0/g,' ').trim();y[ck]=['FECHA ANTIGÜEDAD','ULTIMA FECHA INGRESO'].includes(ck)?parseDateValue(v):v}return cleanEmployee(y)}
function mapImportRow(r){const y={};for(const [k,v] of Object.entries(r)){const ck=canonicalField(k);y[ck]=DATE_FIELDS.includes(ck)?parseDateValue(v):v}if(y.ESTADO)y.ESTADO=statusText(y.ESTADO);if(y['DIAS ACUMULADOS DEL PROCESO'])y['DIAS ACUMULADOS DEL PROCESO']=Number(y['DIAS ACUMULADOS DEL PROCESO'])||0;return y}
function employeeKey(x){return norm(x.DOCUMENTO)}
async function confirmImport(){
 const good=state.importRows.filter(x=>x._importStatus==='ok').map(({_importStatus,...r})=>state.importType==='employees'?cleanEmployee(r):cleanCase(r));
 const dups=state.importRows.filter(x=>x._importStatus==='duplicate-review').map(({_importStatus,...r})=>cleanCase(r));
 if(state.importType==='employees'){if(!good.length){toast('No hay empleados válidos para importar');return}if(!confirm(`La carga contiene ${good.length} empleados válidos${state.importRows.length-good.length?` y ${state.importRows.length-good.length} registros duplicados o con error que no se reemplazarán automáticamente`:''}. Esta operación BORRARÁ la base actual de empleados y la reemplazará completamente por la información válida del archivo más reciente. ¿Deseas continuar?`))return;await bulkAdd('employees',good,{clear:true});toast(`${fmt(good.length)} empleados cargados. La base anterior fue reemplazada.`)}
 else if(state.importType==='cases'){let next=Number(nextCaseId(state.cases));for(const x of good){if(!String(x.ITEM||'').trim())x.ITEM=String(next++);}for(const x of dups){if(!String(x.ITEM||'').trim())x.ITEM=String(next++);}
 await bulkAdd('cases',good);
 await bulkAdd('duplicateReviews',dups.map(x=>({...x,reviewStatus:'PENDIENTE',detectedAt:todayISO(),detectedBy:state.currentEditor})));toast(`${fmt(good.length)} casos nuevos importados${dups.length?` y ${fmt(dups.length)} duplicados enviados a revisión`:''}`)}
 state.importRows=[];state.importMeta=null;await refresh()}
function downloadTemplate(cols,name,sample={}){ensureXLSX().then(()=>{const data=[Object.fromEntries(cols.map(k=>[k,sample[k]??'']))];const ws=XLSX.utils.json_to_sheet(data,{header:cols});ws['!cols']=cols.map(k=>({wch:Math.min(Math.max(k.length+3,14),34)}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'CARGA');const inst=[['INSTRUCCIONES'],['Complete la fila 2 y conserve exactamente los encabezados.'],['Las fechas deben diligenciarse como DD/MM/AAAA.'],['No elimine columnas obligatorias.']];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(inst),'INSTRUCCIONES');XLSX.writeFile(wb,`${name}.xlsx`);toast('Formato descargado')}).catch(e=>toast(e.message))}
function downloadEmployeeTemplate(){downloadTemplate(EMP_FIELDS,'FORMATO_CARGA_EMPLEADOS_MINEX',{'TIPO DOCUMENTO':'CC','EMPRESA CONTRATO':'MX','DOCUMENTO':'123456789','NOMBRE DEL EMPLEADO':'EJEMPLO','FECHA ANTIGÜEDAD':'01/01/2025','ULTIMA FECHA INGRESO':'01/01/2025'})}
function ensureJSZip(){
 if(window.JSZip)return Promise.resolve(window.JSZip);
 const srcs=['./jszip.min.js','./jszip_min.js','https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js'];
 return (async()=>{
  for(const src of srcs){
   try{return await loadScriptOnce(src,'JSZip',{promise:null})}catch(e){console.warn('JSZip no cargó desde',src)}
  }
  throw new Error('No se pudo cargar JSZip. Verifica que jszip.min.js esté publicado junto a index.html.');
 })();
}
function b64ToBytes(b64){const bin=atob(b64),bytes=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i);return bytes}
async function caseTemplateBytes(){
 for(const url of ['./assets/FORMATO_CARGA_CASOS_RELACIONES_LABORALES.xlsx','assets/FORMATO_CARGA_CASOS_RELACIONES_LABORALES.xlsx']){
  try{
   const r=await fetch(url,{cache:'no-store'});
   if(r.ok){const buf=await r.arrayBuffer();if(buf.byteLength)return new Uint8Array(buf)}
  }catch(e){console.warn('No se pudo cargar plantilla desde',url,e)}
 }
 if(!window.MINEX_CASE_TEMPLATE_BASE64){
  try{await loadScriptOnce('./data/case_template_base64.js','MINEX_CASE_TEMPLATE_BASE64',{promise:null})}catch(e){console.warn('Plantilla embebida no disponible')}
 }
 if(window.MINEX_CASE_TEMPLATE_BASE64){
  try{return b64ToBytes(window.MINEX_CASE_TEMPLATE_BASE64)}catch(e){console.warn('Plantilla embebida inválida:',e)}
 }
 throw new Error('La plantilla oficial no está disponible. Publica assets/FORMATO_CARGA_CASOS_RELACIONES_LABORALES.xlsx (copia de Libro3.xlsx) junto a index.html.');
}
async function downloadCaseTemplate(){
 try{
  const bytes=await caseTemplateBytes();
  const blob=new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='FORMATO_CARGA_CASOS_RELACIONES_LABORALES.xlsx';
  document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1500);
  toast('Formato oficial de casos descargado');
 }catch(e){console.error(e);toast('No se pudo descargar el formato oficial: '+e.message)}
}
function toggleDuplicateSelection(id,checked){checked?state.selectedDuplicates.add(id):state.selectedDuplicates.delete(id);duplicateReviewPage()}
function toggleAllDuplicates(checked){(state.duplicateReviews||[]).forEach(x=>checked?state.selectedDuplicates.add(x._id):state.selectedDuplicates.delete(x._id));duplicateReviewPage()}
function selectedDuplicateRows(){return (state.duplicateReviews||[]).filter(x=>state.selectedDuplicates.has(x._id))}
function exportSelectedDuplicates(){const rows=selectedDuplicateRows();if(!rows.length)return toast('Selecciona al menos un duplicado');downloadRows(rows,'DUPLICADOS_EN_REVISION_MINEX')}
async function deleteSelectedDuplicates(){const rows=selectedDuplicateRows();if(!rows.length)return toast('Selecciona al menos un duplicado');if(!confirm(`¿Eliminar ${rows.length} duplicado(s) de la bandeja de revisión?`))return;await bulkDelete('duplicateReviews',rows.map(x=>x._id));state.selectedDuplicates.clear();toast(`${rows.length} duplicado(s) eliminados`);await refresh()}
async function deleteDuplicate(id){if(!confirm('¿Eliminar este registro de la bandeja de revisión?'))return;await del('duplicateReviews',id);toast('Duplicado eliminado');await refresh()}
function editDuplicate(id){const x=state.duplicateReviews.find(y=>y._id===id);if(!x)return;openCase({...x,_id:undefined,_reviewId:id})}
async function approveDuplicate(id){const x=state.duplicateReviews.find(y=>y._id===id);if(!x)return;if(!confirm(`¿Incorporar el caso ${x.ITEM||x._id} a la base maestra?`))return;const d={...x};delete d._id;delete d.reviewStatus;delete d.detectedAt;delete d.detectedBy;delete d._reviewId;if(!String(d.ITEM||'').trim())d.ITEM=nextCaseId(state.cases);await add('cases',cleanCase(d));await del('duplicateReviews',id);toast('Caso incorporado a la base maestra');await refresh()}
function duplicateReviewPage(){const rows=state.duplicateReviews||[];layout(header('Revisión de duplicados','Casos detectados como potencialmente duplicados durante una carga',`<button class="btn" onclick="exportSelectedDuplicates()">⇩ Exportar seleccionados</button><button class="btn danger-outline" onclick="deleteSelectedDuplicates()">🗑 Borrar seleccionados</button><button class="btn" onclick="go('upload')">⇧ Nueva carga</button>`)+`<section class="panel"><div class="panel-head"><div><h2>${fmt(rows.length)} casos pendientes de revisión</h2><p>Revisa cada registro antes de incorporarlo a la base maestra.</p></div></div>${rows.length?`<div class="table-wrap"><table class="table wide-table"><thead><tr><th class="select-col"><input type="checkbox" title="Seleccionar todos" onchange="toggleAllDuplicates(this.checked)"></th><th>ID</th><th>Fecha</th><th>Requerimiento</th><th>Empresa</th><th>Área</th><th>Empleado</th><th>Detectado por</th><th>Acciones</th></tr></thead><tbody>${rows.map(x=>`<tr><td class="select-col"><input type="checkbox" ${state.selectedDuplicates.has(x._id)?'checked':''} onchange="toggleDuplicateSelection(${x._id},this.checked)"></td><td>${esc(x.ITEM||x._id)}</td><td>${formatDate(x.FECHA)}</td><td>${esc(x.REQUERIMIENTO)}</td><td>${esc(x.EMP)}</td><td>${esc(x['ÁREA'])}</td><td>${esc(x['EMPLEADO / EMPRESA'])}</td><td>${esc(x.detectedBy||'')}</td><td><button class="icon-btn" onclick="editDuplicate(${x._id})">✎</button><button class="icon-btn danger-icon" onclick="deleteDuplicate(${x._id})">⌫</button><button class="btn mini" onclick="approveDuplicate(${x._id})">✓ Incorporar</button></td></tr>`).join('')}</tbody></table></div>`:'<div class="empty">No hay duplicados pendientes de revisión.</div>'}</section>`)}
function filteredReportCases(){
 const f=state.reportFilters||{}; let rows=[...state.cases];
 if(f.EMP) rows=rows.filter(x=>norm(x.EMP)===norm(f.EMP));
 if(f['ÁREA']) rows=rows.filter(x=>norm(x['ÁREA'])===norm(f['ÁREA']));
 if(f.REQUERIMIENTO) rows=rows.filter(x=>norm(x.REQUERIMIENTO)===norm(f.REQUERIMIENTO));
 if(f.ESTADO) rows=rows.filter(x=>statusText(x.ESTADO)===f.ESTADO);
 if(f.MES) rows=rows.filter(x=>norm(monthName(x.MES))===norm(f.MES));
 if(f.FROM) rows=rows.filter(x=>parseDateValue(x.FECHA)>=f.FROM);
 if(f.TO) rows=rows.filter(x=>parseDateValue(x.FECHA)<=f.TO);
 return rows;
}
function metricsForRows(cs){const today=todayISO(),end=weekFriday(today);const closed=cs.filter(x=>statusText(x.ESTADO)==='CERRADO').length,open=cs.length-closed;const overdue=cs.filter(x=>statusText(x.ESTADO)!=='CERRADO'&&x['FECHA DE SEGUIMIENTO']&&x['FECHA DE SEGUIMIENTO']<today).length,todayN=cs.filter(x=>statusText(x.ESTADO)!=='CERRADO'&&x['FECHA DE SEGUIMIENTO']===today).length,weekN=cs.filter(x=>statusText(x.ESTADO)!=='CERRADO'&&x['FECHA DE SEGUIMIENTO']>today&&x['FECHA DE SEGUIMIENTO']<=end).length;const avg=cs.length?Math.round(cs.reduce((a,x)=>a+(Number(x['DIAS ACUMULADOS DEL PROCESO'])||0),0)/cs.length):0;return {total:cs.length,closed,open,overdue,todayN,weekN,avg,rate:cs.length?Math.round(closed/cs.length*100):0}}
function reportsPage(){
 const rows=filteredReportCases(),m=metricsForRows(rows),areaRows=groupStats('ÁREA',rows),empRows=groupStats('EMP',rows),reqRows=groupStats('REQUERIMIENTO',rows),respRows=groupStats('PENDIENTE POR',rows),monthRows=groupStats('MES',rows);
 const noDetail=rows.filter(x=>!String(x['DETALLE DE LA SOLICITUD']||'').trim()).length;
 const noFollow=rows.filter(x=>statusText(x.ESTADO)!=='CERRADO'&&!x['FECHA DE SEGUIMIENTO']).length;
 const dup=duplicateReport(rows,caseDuplicateKey).length;
 const topArea=areaRows[0],topReq=reqRows[0],topEmp=empRows[0];
 const months=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
 layout(header('Informes y análisis estadístico','Centro de análisis de Relaciones Laborales: indicadores, distribución, calidad y reportes descargables',`<button class="btn" onclick="printReport()">▣ Imprimir / PDF</button><button class="btn primary" onclick="exportAllReports()">⇩ Excel consolidado</button>`)+
 `<section class="filter-card report-filter-card"><div class="filter-card-head"><div class="filter-title"><span class="filter-icon">▦</span><div><strong>Filtros del análisis</strong><small>Todos los indicadores y reportes de esta página respetan estos filtros</small></div></div><button class="link-btn" onclick="state.reportFilters={};reportsPage()">↺ Limpiar</button></div><div class="filter-groups"><div class="filter-group"><div class="group-title blue">▦ Organización</div>${selectReportFilter('EMP','Empresa',unique('EMP'),state.reportFilters.EMP)}${selectReportFilter('ÁREA','Área',unique('ÁREA'),state.reportFilters['ÁREA'])}</div><div class="filter-group"><div class="group-title purple">▤ Gestión</div>${selectReportFilter('REQUERIMIENTO','Requerimiento',unique('REQUERIMIENTO'),state.reportFilters.REQUERIMIENTO)}${selectReportFilter('ESTADO','Estado',['CERRADO','EN SEGUIMIENTO'],state.reportFilters.ESTADO)}</div><div class="filter-group"><div class="group-title green">◷ Periodo</div>${selectReportFilter('MES','Mes',months,state.reportFilters.MES)}${dateReportFilter('FROM','Desde',state.reportFilters.FROM)}${dateReportFilter('TO','Hasta',state.reportFilters.TO)}</div></div><div class="filter-foot"><strong>${fmt(rows.length)} casos incluidos en el análisis</strong></div></section>`+
 `<div class="report-kpis report-kpis-v2"><div><span>Casos</span><strong>${fmt(m.total)}</strong><small>Según filtros</small></div><div><span>% de cierre</span><strong>${m.rate}%</strong><small>${fmt(m.closed)} cerrados</small></div><div><span>En seguimiento</span><strong>${fmt(m.open)}</strong><small>Gestión activa</small></div><div><span>Vencidos</span><strong>${fmt(m.overdue)}</strong><small>Atención prioritaria</small></div><div><span>Promedio días</span><strong>${fmt(m.avg)}</strong><small>Días acumulados</small></div><div><span>Esta semana</span><strong>${fmt(m.weekN)}</strong><small>Seguimientos próximos</small></div></div>`+
 `<div class="insight-strip report-insights"><div><span>Área con mayor volumen</span><strong>${esc(topArea?.name||'—')}</strong><small>${fmt(topArea?.count||0)} casos</small></div><div><span>Requerimiento principal</span><strong>${esc(topReq?.name||'—')}</strong><small>${fmt(topReq?.count||0)} casos</small></div><div><span>Empresa con mayor gestión</span><strong>${esc(topEmp?.name||'—')}</strong><small>${fmt(topEmp?.count||0)} casos</small></div><div><span>Casos sin seguimiento</span><strong>${fmt(noFollow)}</strong><small>Abiertos sin programación</small></div></div>`+
 `<div class="report-grid report-grid-v2">${reportCardData('Área',areaRows,'ÁREA')}${reportCardData('Empresa',empRows,'EMP')}${reportCardData('Requerimiento',reqRows,'REQUERIMIENTO')}${reportCardData('Responsable',respRows,'PENDIENTE POR')}${reportCardData('Mes',monthRows,'MES')}</div>`+
 `<div class="dash-grid"><section class="panel"><div class="panel-head"><div><h2>Distribución por estado</h2><p>Estado de los casos incluidos.</p></div></div><div class="stat-bars"><div><span>Cerrados</span><strong>${fmt(m.closed)} · ${m.rate}%</strong><i><b style="width:${m.rate}%"></b></i></div><div><span>En seguimiento</span><strong>${fmt(m.open)} · ${100-m.rate}%</strong><i><b class="amber-bar" style="width:${100-m.rate}%"></b></i></div></div></section><section class="panel"><div class="panel-head"><div><h2>Calidad de información</h2><p>Indicadores para depuración.</p></div></div><div class="quality-list"><div><span>Posibles duplicados</span><strong>${fmt(dup)}</strong></div><div><span>Sin detalle de solicitud</span><strong>${fmt(noDetail)}</strong></div><div><span>Sin fecha de seguimiento</span><strong>${fmt(noFollow)}</strong></div><div><span>Sin responsable</span><strong>${fmt(rows.filter(x=>!x['PENDIENTE POR']).length)}</strong></div></div></section></div>`+
 `<section class="panel report-detail"><div class="panel-head"><div><h2>Análisis por área</h2><p>Volumen, cierres, seguimiento, vencimientos y porcentaje de cierre.</p></div><button class="btn" onclick="exportGroupedReport('ÁREA')">⇩ Descargar</button></div>${reportDetailTable(areaRows)}</section>`+
 `<section class="panel report-detail"><div class="panel-head"><div><h2>Comportamiento mensual</h2><p>Distribución de casos por mes.</p></div><button class="btn" onclick="exportGroupedReport('MES')">⇩ Descargar</button></div>${reportDetailTable(monthRows)}</section>`+
 `<section class="panel report-detail"><div class="panel-head"><div><h2>Análisis por requerimiento</h2><p>Demanda de gestión y porcentaje de cierre.</p></div><button class="btn" onclick="exportGroupedReport('REQUERIMIENTO')">⇩ Descargar</button></div>${reportDetailTable(reqRows)}</section>`
 )
}
function selectReportFilter(k,l,vals,selected=''){return `<div class="f-field"><label>${l}</label><select onchange="state.reportFilters['${k}']=this.value;reportsPage()">${opts(vals,selected)}</select></div>`}
function dateReportFilter(k,l,v=''){return `<div class="f-field"><label>${l}</label><input type="date" value="${esc(v||'')}" onchange="state.reportFilters['${k}']=this.value;reportsPage()"></div>`}
function groupStats(group,source=state.cases){const map={};source.forEach(x=>{const k=group==='MES'?monthName(x[group]):(x[group]||'Sin dato');if(!map[k])map[k]={count:0,closed:0,open:0,overdue:0};map[k].count++;if(statusText(x.ESTADO)==='CERRADO')map[k].closed++;else map[k].open++;if(statusText(x.ESTADO)!=='CERRADO'&&x['FECHA DE SEGUIMIENTO']&&x['FECHA DE SEGUIMIENTO']<todayISO())map[k].overdue++});return Object.entries(map).map(([name,v])=>({name,...v,rate:v.count?Math.round(v.closed/v.count*100):0})).sort((a,b)=>b.count-a.count)}
function reportCardData(title,rows,group){return `<section class="panel report-card"><div class="panel-head"><div><h2>Por ${esc(title)}</h2><p>${fmt(rows.reduce((a,x)=>a+x.count,0))} casos</p></div><button class="link-btn" onclick="exportGroupedReport('${esc(group)}')">⇩ Excel</button></div><div class="report-list enhanced">${rows.slice(0,12).map(x=>`<div class="report-row"><div class="report-row-main"><span>${esc(x.name)}</span><small>${x.closed} cerrados · ${x.open} en seguimiento · ${x.overdue} vencidos · ${x.rate}% cierre</small></div><div class="bar"><i style="width:${Math.round(x.count/(rows[0]?.count||1)*100)}%"></i></div><strong>${fmt(x.count)}</strong></div>`).join('')}</div></section>`}
function reportDetailTable(rows){return `<div class="table-wrap"><table class="table"><thead><tr><th>Área</th><th>Casos</th><th>Cerrados</th><th>En seguimiento</th><th>Vencidos</th><th>% cierre</th></tr></thead><tbody>${rows.map(x=>`<tr><td><strong>${esc(x.name)}</strong></td><td>${fmt(x.count)}</td><td>${fmt(x.closed)}</td><td>${fmt(x.open)}</td><td>${fmt(x.overdue)}</td><td>${x.rate}%</td></tr>`).join('')}</tbody></table></div>`}
function reportCard(group){const map={};state.cases.forEach(x=>{const k=group==='MES'?monthName(x[group]):(x[group]||'Sin dato');map[k]=(map[k]||0)+1});const arr=Object.entries(map).sort((a,b)=>b[1]-a[1]);return `<section class="panel report-card"><div class="panel-head"><div><h2>Por ${group.toLowerCase()}</h2><p>${fmt(arr.reduce((a,x)=>a+x[1],0))} registros</p></div><button class="link-btn" onclick="exportGroupedReport('${esc(group)}')">⇩ Excel</button></div><div class="report-list">${arr.slice(0,12).map(([k,v])=>`<div><span>${esc(k)}</span><div class="bar"><i style="width:${Math.round(v/(arr[0]?.[1]||1)*100)}%"></i></div><strong>${fmt(v)}</strong></div>`).join('')}</div></section>`}
async function exportAllReports(){
 try{await ensureXLSX()}catch(e){return toast(e.message)}
 const wb=XLSX.utils.book_new(); const source=filteredReportCases();
 const groups=['ÁREA','EMP','REQUERIMIENTO','PENDIENTE POR','MES'];
 for(const g of groups){const rows=groupStats(g,source).map(x=>({CATEGORIA:x.name,CASOS:x.count,CERRADOS:x.closed,EN_SEGUIMIENTO:x.open,VENCIDOS:x.overdue,PORCENTAJE_CIERRE:x.rate}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),safeSheet(g))}
 const m=calcMetrics();
 const resumen=[
  {INDICADOR:'Casos totales',VALOR:m.total},
  {INDICADOR:'Casos cerrados',VALOR:m.closed},
  {INDICADOR:'Casos en seguimiento',VALOR:m.open},
  {INDICADOR:'Porcentaje de cierre',VALOR:m.rate+'%'},
  {INDICADOR:'Casos vencidos',VALOR:m.overdue},
  {INDICADOR:'Vencen hoy',VALOR:m.todayN},
  {INDICADOR:'Esta semana',VALOR:m.weekN},
  {INDICADOR:'Promedio días',VALOR:m.avg}
 ];
 XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(resumen),'RESUMEN');
 XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rowsForExport(source,CASE_FIELDS)),'DETALLE CASOS');
 XLSX.writeFile(wb,'INFORME_ESTADISTICO_RELACIONES_LABORALES.xlsx');toast('Informe estadístico consolidado descargado')
}
function exportGroupedReport(g){const rows=groupStats(g,filteredReportCases()).map(x=>({CATEGORIA:x.name,CASOS:x.count,CERRADOS:x.closed,EN_SEGUIMIENTO:x.open,VENCIDOS:x.overdue,PORCENTAJE_CIERRE:x.rate}));downloadRows(rows,'INFORME_'+safeSheet(g))}
function safeSheet(s){return norm(s).replace(/[^A-Z0-9]/g,'_').slice(0,28)||'REPORTE'}
function printReport(){window.print()}
function exportPage(){layout(header('Exportar','Descarga controlada de información, reportes y datos operativos')+`<div class="export-grid"><section class="panel export-card"><span class="export-icon">▤</span><h2>Casos</h2><p>Base completa o filtrada desde el módulo Casos.</p><button class="btn primary" onclick="exportCasesFiltered()">Descargar casos</button></section><section class="panel export-card"><span class="export-icon">♙</span><h2>Empleados</h2><p>Base administrativa organizada.</p><button class="btn primary" onclick="downloadRows(state.employees,'EMPLEADOS')">Descargar empleados</button></section><section class="panel export-card"><span class="export-icon">◫</span><h2>Informes</h2><p>Resumen por área, empresa, requerimiento y responsables.</p><button class="btn primary" onclick="exportAllReports()">Descargar informe</button></section></div>`)}
function pickForExport(x,cols){const o={};for(const k of cols){let v=x[k]??'';if(DATE_FIELDS.includes(k))v=parseDateValue(v)?formatDate(v):'';if(k==='ESTADO')v=statusText(v);if(['VISIBLE_DASH','MATERNIDAD_DASH'].includes(k))v=statusText(v);o[k]=v}return o}
function rowsForExport(rows,cols){return rows.map(x=>pickForExport(x,cols))}
async function downloadRows(rows,name){if(!rows?.length){toast('No hay datos para exportar');return}try{await ensureXLSX()}catch(e){return toast(e.message)}const cols=[...new Set(rows.flatMap(x=>Object.keys(x)).filter(k=>k!=='_id'))];const data=rowsForExport(rows,cols);const wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(data,{header:cols});XLSX.utils.book_append_sheet(wb,ws,'DATOS');XLSX.writeFile(wb,`${name}.xlsx`);toast('Archivo descargado')}
function exportCSV(rows){if(!rows.length){toast('No hay datos');return}const cols=[...new Set(rows.flatMap(x=>Object.keys(x)))];const lines=[cols.join(';'),...rows.map(r=>cols.map(c=>'"'+String(r[c]??'').replace(/"/g,'""')+'"').join(';'))];const blob=new Blob(['\ufeff'+lines.join('\n')],{type:'text/csv;charset=utf-8'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='MINEX_EXPORT.csv';a.click();URL.revokeObjectURL(a.href)}
async function ensureExcelJS(){
 if(window.ExcelJS)return window.ExcelJS;
 window.__exceljs=window.__exceljs||{promise:null};
 return loadScriptOnce('https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js','ExcelJS',window.__exceljs);
}
function excelExportValue(v,key){
 if(v==null||v==='')return null;
 if(DATE_FIELDS.includes(key)){const d=parseDateValue(v);return d||null;}
 if(key==='CEDULA'){const n=String(v).trim();return /^\d+$/.test(n)?Number(n):n;}
 if(key==='DIAS ACUMULADOS DEL PROCESO')return Number(v)||0;
 if(key==='ESTADO')return statusText(v);
 return v;
}
function excelColName(n){let s='';while(n>0){const r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26)}return s}
function xmlText(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;')}
function excelDateSerial(iso){if(!iso)return null;const m=String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)return null;const d=Date.UTC(Number(m[1]),Number(m[2])-1,Number(m[3]));return Math.round((d-Date.UTC(1899,11,30))/86400000)}
function replaceFormulaRow(formula,rowNum){return String(formula||'').replace(/\b([A-Z]{1,3})(\d+)\b/g,(m,col)=>`${col}${rowNum}`)}
function setCellValueXML(cell, value, key){
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 while(cell.firstChild)cell.removeChild(cell.firstChild);
 cell.removeAttribute('t');
 if(value===null||value===undefined||value==='')return;
 let v=value;
 if(DATE_FIELDS.includes(key)){const d=excelDateSerial(v);if(d!==null)v=d;}
 if(key==='CEDULA' && /^\d+$/.test(String(v).trim()))v=Number(v);
 if(typeof v==='number' && Number.isFinite(v)){
  const el=cell.ownerDocument.createElementNS(ns,'v');el.textContent=String(v);cell.appendChild(el);return;
 }
 cell.setAttribute('t','inlineStr');
 const is=cell.ownerDocument.createElementNS(ns,'is');const t=cell.ownerDocument.createElementNS(ns,'t');
 const text=String(v);if(/^\s|\s$/.test(text))t.setAttribute('xml:space','preserve');t.textContent=text;is.appendChild(t);cell.appendChild(is);
}
function cloneFormulaCellXML(cell,rowNum){
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 const f=cell.getElementsByTagNameNS(ns,'f')[0];
 if(!f)return;
 f.removeAttribute('t');f.removeAttribute('ref');f.removeAttribute('si');
 f.textContent=replaceFormulaRow(f.textContent,rowNum);
 const v=cell.getElementsByTagNameNS(ns,'v')[0];if(v)v.textContent='';
 cell.removeAttribute('t');
}
function colIndex(col){let n=0;for(const ch of col)n=n*26+ch.charCodeAt(0)-64;return n}
function ensureRowCells(row,headers,rowNum,ns){
 const colOf=c=>(c.getAttribute('r')||'').match(/^([A-Z]+)/)?.[1]||'A';
 const have=new Set(Array.from(row.getElementsByTagNameNS(ns,'c')).map(colOf));
 Object.keys(headers).forEach(col=>{
  if(have.has(col))return;
  const c=row.ownerDocument.createElementNS(ns,'c');c.setAttribute('r',`${col}${rowNum}`);
  const next=Array.from(row.getElementsByTagNameNS(ns,'c')).find(x=>colIndex(colOf(x))>colIndex(col));
  next?row.insertBefore(c,next):row.appendChild(c);
 });
}
async function exportCasesFiltered(){
 const rows=filteredCases();
 if(!rows.length)return toast('No hay casos para exportar con los filtros actuales');
 try{
  await ensureJSZip();
  const bytes=await caseTemplateBytes();
  const zip=await JSZip.loadAsync(bytes);
  const sheetFile=zip.file('xl/worksheets/sheet1.xml');
  const tableFile=zip.file('xl/tables/table1.xml');
  if(!sheetFile||!tableFile)throw new Error('La plantilla oficial no tiene la estructura esperada');
  const parser=new DOMParser();
  const serializer=new XMLSerializer();
  const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const xml=await sheetFile.async('string');
  const doc=parser.parseFromString(xml,'application/xml');
  const sheetData=doc.getElementsByTagNameNS(ns,'sheetData')[0];
  if(!sheetData)throw new Error('No se encontró sheetData en VOLUMETRIAS');
  const rowNodes=Array.from(sheetData.getElementsByTagNameNS(ns,'row'));
  const headerRow=rowNodes.find(r=>r.getAttribute('r')==='1');
  const templateRow=rowNodes.find(r=>r.getAttribute('r')==='2');
  if(!headerRow||!templateRow)throw new Error('La plantilla no contiene fila de encabezados y datos');
  const headers={};
  Array.from(headerRow.getElementsByTagNameNS(ns,'c')).forEach(c=>{const ref=c.getAttribute('r')||'';const col=ref.match(/^([A-Z]+)/)?.[1];if(!col)return;const v=c.getElementsByTagNameNS(ns,'v')[0];if(v){headers[col]=v.textContent;}});
  // Resolve shared-string headers from the workbook.
  const ssFile=zip.file('xl/sharedStrings.xml');
  if(ssFile){
   const ssDoc=parser.parseFromString(await ssFile.async('string'),'application/xml');
   const items=Array.from(ssDoc.getElementsByTagNameNS(ns,'si'));
   Object.keys(headers).forEach(col=>{const idx=Number(headers[col]);if(Number.isInteger(idx)&&items[idx]){headers[col]=Array.from(items[idx].getElementsByTagNameNS(ns,'t')).map(t=>t.textContent).join('')}});
  }
  const colByHeader={};Object.entries(headers).forEach(([col,h])=>{if(h)colByHeader[String(h).trim()]=col});
  const formulaCols=new Set(['ITEM','MES','MES CIERRE','DIAS ACUMULADOS DEL PROCESO','ESTADO','VISIBLE_DASH','MATERNIDAD_DASH']);
  // Remove every existing data row. The original workbook is used only as the formatting/formula blueprint.
  rowNodes.filter(r=>Number(r.getAttribute('r'))>=2).forEach(r=>r.parentNode.removeChild(r));
  rows.forEach((item,idx)=>{
   const rowNum=idx+2;const row=templateRow.cloneNode(true);row.setAttribute('r',String(rowNum));
   ensureRowCells(row,headers,rowNum,ns);
   Array.from(row.getElementsByTagNameNS(ns,'c')).forEach(cell=>{
    const ref=cell.getAttribute('r')||'';const col=ref.match(/^([A-Z]+)/)?.[1];if(!col)return;cell.setAttribute('r',`${col}${rowNum}`);
    const key=headers[col];
    if(formulaCols.has(key))cloneFormulaCellXML(cell,rowNum);
    else setCellValueXML(cell,item[key],key);
   });
   sheetData.appendChild(row);
  });
  const lastRow=Math.max(2,rows.length+1);
  // Keep table/filters synchronized with the actual exported rows.
  let tableXml=await tableFile.async('string');
  tableXml=tableXml.replace(/ref="A1:([A-Z]+)\d+"/g,(m,c)=>`ref="A1:${c}${lastRow}"`);
  zip.file('xl/tables/table1.xml',tableXml);
  const dim=doc.getElementsByTagNameNS(ns,'dimension')[0];if(dim){const lc=(dim.getAttribute('ref')||'').match(/:([A-Z]+)\d+$/)?.[1]||'AD';dim.setAttribute('ref',`A1:${lc}${lastRow}`)}
  zip.file('xl/worksheets/sheet1.xml',serializer.serializeToString(doc));
  // Force Excel to recalculate formulas on open while retaining the corporate workbook structure.
  const wbFile=zip.file('xl/workbook.xml');
  if(wbFile){let wbxml=await wbFile.async('string');if(/<calcPr\b/.test(wbxml))wbxml=wbxml.replace(/<calcPr([^>]*)\/>/, '<calcPr$1 fullCalcOnLoad="1" forceFullCalc="1" calcMode="auto"/>');else wbxml=wbxml.replace('</workbook>','<calcPr fullCalcOnLoad="1" forceFullCalc="1" calcMode="auto"/></workbook>');zip.file('xl/workbook.xml',wbxml)}
  const out=await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
  const a=document.createElement('a');a.href=URL.createObjectURL(out);a.download='CASOS_RELACIONES_LABORALES.xlsx';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),2000);
  toast(`Casos descargados correctamente · ${fmt(rows.length)} registros`);
 }catch(e){console.error(e);toast('No se pudo generar la descarga de casos: '+e.message)}
}
function exportVolFiltered(){downloadRows(rowsForVolExport(),'VOLUMETRIAS_FILTRADAS_MINEX')}
function toggleEmployeeSelection(id,checked){checked?state.selectedEmployees.add(id):state.selectedEmployees.delete(id);employeesPage()}
function toggleAllEmployees(checked){const q=norm(state.empSearch);state.employees.filter(x=>!q||Object.values(x).some(v=>norm(v).includes(q))).forEach(x=>checked?state.selectedEmployees.add(x._id):state.selectedEmployees.delete(x._id));employeesPage()}
function selectedEmployeeRows(){return state.employees.filter(x=>state.selectedEmployees.has(x._id))}
function exportSelectedEmployees(){const rows=selectedEmployeeRows();if(!rows.length)return toast('Selecciona al menos un empleado');downloadRows(rows,'EMPLEADOS_SELECCIONADOS')}
async function deleteSelectedEmployees(){const rows=selectedEmployeeRows();if(!rows.length)return toast('Selecciona al menos un empleado');if(!confirm(`¿Eliminar ${rows.length} empleado(s) seleccionado(s)?`))return;await bulkDelete('employees',rows.map(x=>x._id));state.selectedEmployees.clear();toast(`${rows.length} empleado(s) eliminado(s)`);refresh()}

function employeesPage(){const q=norm(state.empSearch);const rows=state.employees.filter(x=>!q||Object.values(x).some(v=>norm(v).includes(q)));layout(header('Empleados','Base administrativa para autocompletar y mantener información · la última carga reemplaza la anterior',`<button class="btn" onclick="downloadEmployeeTemplate()">⇩ Formato</button><button class="btn" onclick="state.importType='employees';go('upload')">⇧ Importar</button><button class="btn" onclick="exportSelectedEmployees()">⇩ Seleccionados</button><button class="btn danger-outline" onclick="deleteSelectedEmployees()">🗑 Borrar seleccionados</button><button class="btn primary" onclick="openEmployee()">＋ Agregar empleado</button>`)+`<section class="panel"><div class="employee-toolbar"><div><strong>${fmt(rows.length)}</strong> empleados encontrados · ${fmt(state.employees.length)} en base</div><input value="${esc(state.empSearch)}" placeholder="Buscar por documento, nombre, cargo, empresa…" oninput="state.empSearch=this.value;employeesPage()"></div>${employeesTable(rows)}</section>`)}
function employeesTable(rows){if(!rows.length)return '<div class="empty">No hay empleados.</div>';const cols=EMP_FIELDS;return `<div class="table-wrap employee-full-table"><table class="table wide-table"><thead><tr><th class="select-col"><input type="checkbox" title="Seleccionar todos" onchange="toggleAllEmployees(this.checked)"></th>${cols.map(k=>`<th>${esc(k)}</th>`).join('')}<th>ACCIONES</th></tr></thead><tbody>${rows.slice(0,1000).map(x=>`<tr><td class="select-col"><input type="checkbox" ${state.selectedEmployees.has(x._id)?'checked':''} onchange="toggleEmployeeSelection(${x._id},this.checked)"></td>${cols.map(k=>`<td>${DATE_FIELDS.includes(k)?esc(formatDate(x[k])):esc(x[k])}</td>`).join('')}<td><div class="row-actions"><button class="icon-btn" title="Editar empleado" onclick="openEmployee(${x._id})">✎</button><button class="icon-btn danger-icon" title="Eliminar empleado" onclick="deleteEmployee(${x._id})">⌫</button></div></td></tr>`).join('')}</tbody></table></div>`}
function openEmployee(existing=null){const x=existing||{};document.body.insertAdjacentHTML('beforeend',`<div class="modal-back" id="employeeModal"><div class="modal xwide"><div class="modal-head"><div><h2>${existing?'Editar empleado':'Agregar empleado'}</h2><p>La cédula/documento identifica un registro único.</p></div><button class="icon-btn" onclick="closeModal('employeeModal')">×</button></div><form id="employeeForm"><div class="formgrid">${EMP_FIELDS.map(k=>{const type=['FECHA ANTIGÜEDAD','ULTIMA FECHA INGRESO'].includes(k)?'date':'text';const val=['FECHA ANTIGÜEDAD','ULTIMA FECHA INGRESO'].includes(k)?parseDateValue(x[k]):x[k]??'';return `<div class="field"><label>${esc(k)}</label><input type="${type}" name="${esc(k)}" value="${esc(val)}"></div>`}).join('')}</div><div class="modal-actions"><button type="button" class="btn" onclick="closeModal('employeeModal')">Cancelar</button><button class="btn primary">${existing?'Guardar cambios':'Crear empleado'}</button></div></form></div></div>`);document.getElementById('employeeForm').onsubmit=async e=>{e.preventDefault();const d=Object.fromEntries(new FormData(e.target).entries());d.DOCUMENTO=String(d.DOCUMENTO||'').trim();if(!d.DOCUMENTO)return toast('El documento es obligatorio');for(const k of ['FECHA ANTIGÜEDAD','ULTIMA FECHA INGRESO'])d[k]=parseDateValue(d[k]);const dup=state.employees.find(y=>employeeKey(y)===employeeKey(d)&&(!existing||y._id!==existing._id));if(dup)return toast('Ya existe un empleado con ese documento');if(existing)await put('employees',{...existing,...d});else await add('employees',cleanEmployee(d));closeModal('employeeModal');toast(existing?'Empleado actualizado':'Empleado creado');refresh()}}
async function deleteEmployee(id){const x=state.employees.find(y=>y._id===id);if(x&&confirm(`¿Eliminar a ${x['NOMBRE DEL EMPLEADO']||x.DOCUMENTO}?`)){await del('employees',id);toast('Empleado eliminado');refresh()}}
function settingsPage(){
const dup=duplicateReport(state.cases,caseDuplicateKey), empDup=duplicateReport(state.employees,employeeKey);
layout(header('Configuración','Control de datos, calidad y almacenamiento local',`<button class="btn" onclick="restoreInitialData()">↻ Restaurar datos iniciales</button>`)+`
<div class="settings-grid">
<section class="panel">
<div class="panel-head"><div><h2>Calidad de datos</h2><p>Identifica inconsistencias antes de generar informes.</p></div></div>
<div class="quality-list">
<div><span>Casos duplicados</span><strong>${dup.length}</strong></div>
<div><span>Empleados duplicados</span><strong>${empDup.length}</strong></div>
<div><span>Casos sin fecha</span><strong>${state.cases.filter(x=>!x.FECHA).length}</strong></div>
<div><span>Casos sin seguimiento</span><strong>${state.cases.filter(x=>statusText(x.ESTADO)!=='CERRADO'&&!x['FECHA DE SEGUIMIENTO']).length}</strong></div>
<div><span>Casos sin responsable</span><strong>${state.cases.filter(x=>!x['PENDIENTE POR']).length}</strong></div>
</div>
</section>
<section class="panel">
<div class="panel-head"><div><h2>Almacenamiento local</h2><p>IndexedDB conserva una copia local para trabajar con rapidez y como respaldo de la sesión.</p></div></div>
<div class="notice">IndexedDB funciona como caché local. Cuando GitHub está conectado, los registros modificados se publican en el repositorio y pueden ser descargados por otros usuarios.</div>
</section>
<section class="panel danger-zone">
<div class="panel-head"><div><h2>Gestión de datos</h2><p>Acciones destructivas. Se solicita confirmación antes de eliminar información.</p></div></div>
<div class="data-actions">
<button class="btn danger-outline" onclick="clearDataStore('cases','casos')">Eliminar todos los casos</button>
<button class="btn danger-outline" onclick="clearDataStore('employees','empleados')">Eliminar todos los empleados</button>
<div class="notice">Las volumetrías se calculan directamente sobre la base de casos; no existe una segunda base para eliminar.</div>
<button class="btn danger" onclick="clearAllData()">Eliminar TODA la información</button>
</div>
<p class="muted small">Eliminar toda la información deja la aplicación vacía. No se restauran datos automáticamente. Si necesitas recuperar la información inicial, utiliza “Restaurar datos iniciales”.</p>
</section>
</div>`)}
function duplicateReport(rows,keyFn){const m=new Map();rows.forEach(x=>{const k=keyFn(x);if(!k)return;if(!m.has(k))m.set(k,[]);m.get(k).push(x)});return [...m.values()].filter(a=>a.length>1)}
async function restoreInitialData(){
if(!confirm('Esto reemplazará los datos operativos actuales por los datos iniciales incluidos con la aplicación. ¿Deseas continuar?'))return;
for(const s of STORES)await clearStore(s);
await seed();
toast('Datos iniciales restaurados');
await refresh();
}
async function clearDataStore(store,label){
if(!confirm(`Se eliminarán TODOS los ${label}. Esta acción no se puede deshacer. ¿Continuar?`))return;
await clearStore(store);
toast(`${label.charAt(0).toUpperCase()+label.slice(1)} eliminados`);
await refresh();
}
async function clearAllData(){
if(!confirm('ADVERTENCIA: se eliminarán CASOS, EMPLEADOS, VOLUMETRÍAS y configuraciones locales. La aplicación quedará vacía. ¿Deseas continuar?'))return;
if(!confirm('Última confirmación: ¿ELIMINAR TODA LA INFORMACIÓN? Esta acción no se puede deshacer.'))return;
for(const s of STORES)await clearStore(s);
await add('config',{key:'initialized',value:true});
state.cases=[];state.employees=[];state.vols=[];
toast('Toda la información fue eliminada');
await refresh();
}
function keySearch(){document.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){e.preventDefault();go('cases')}})}
function render(){destroyCharts();if(state.page==='dashboard')dashboard();else if(state.page==='cases')casesPage();else if(state.page==='follow')followPage();else if(state.page==='upload')uploadPage();else if(state.page==='reports')reportsPage();else if(state.page==='employees')employeesPage();else if(state.page==='export')exportPage();else settingsPage()}
function toast(t){const e=document.createElement('div');e.className='toast';e.textContent=t;document.body.appendChild(e);setTimeout(()=>e.remove(),2600)}
keySearch();init();
