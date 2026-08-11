const KEY='controloHorasV6';
const FILE_DB='controloHorasFicheiros';
function fileDb(){return new Promise((resolve,reject)=>{const req=indexedDB.open(FILE_DB,1);req.onupgradeneeded=()=>{const db=req.result;if(!db.objectStoreNames.contains('files'))db.createObjectStore('files')};req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)})}
async function storeFile(key,blob){const db=await fileDb();return new Promise((res,rej)=>{const tx=db.transaction('files','readwrite');tx.objectStore('files').put(blob,key);tx.oncomplete=res;tx.onerror=()=>rej(tx.error)})}
async function getStoredFile(key){const db=await fileDb();return new Promise((res,rej)=>{const req=db.transaction('files').objectStore('files').get(key);req.onsuccess=()=>res(req.result);req.onerror=()=>rej(req.error)})}
async function deleteStoredFile(key){const db=await fileDb();return new Promise((res,rej)=>{const tx=db.transaction('files','readwrite');tx.objectStore('files').delete(key);tx.oncomplete=res;tx.onerror=()=>rej(tx.error)})}
async function downloadStoredFile(key,name){const blob=await getStoredFile(key);if(!blob){alert('O ficheiro original não está disponível neste dispositivo.');return}const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),2000)}
async function openStoredSheet(key,name){
 try{
  const blob=await getStoredFile(key);
  if(!blob){alert('O ficheiro original não está disponível neste dispositivo.');return}
  if(!window.XLSX){alert('O leitor de Excel ainda não está disponível.');return}
  const wb=XLSX.read(await blob.arrayBuffer(),{type:'array',cellDates:true});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const sheetHtml=XLSX.utils.sheet_to_html(ws,{editable:false});
  let modal=document.getElementById('sheetViewerModal');
  if(!modal){
   modal=document.createElement('div');
   modal.id='sheetViewerModal';
   modal.style.cssText='position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.55);padding:12px;display:flex;align-items:stretch;justify-content:center';
   modal.innerHTML='<div style="background:#fff;border-radius:12px;width:min(1200px,100%);display:flex;flex-direction:column;overflow:hidden"><div style="display:flex;gap:8px;align-items:center;padding:10px;border-bottom:1px solid #ddd"><strong id="sheetViewerTitle" style="flex:1"></strong><button id="sheetViewerDownload">Descarregar</button><button id="sheetViewerClose">Fechar</button></div><div id="sheetViewerBody" style="overflow:auto;padding:10px;flex:1"></div></div>';
   document.body.appendChild(modal);
   document.getElementById('sheetViewerClose').onclick=()=>modal.remove();
  }
  document.getElementById('sheetViewerTitle').textContent=name;
  const body=document.getElementById('sheetViewerBody');
  body.innerHTML=sheetHtml;
  body.querySelectorAll('table').forEach(t=>t.style.cssText='border-collapse:collapse;min-width:900px');
  body.querySelectorAll('td,th').forEach(c=>c.style.cssText='border:1px solid #ccc;padding:4px 6px;white-space:nowrap');
  document.getElementById('sheetViewerDownload').onclick=()=>downloadStoredFile(key,name);
 }catch(err){
  alert('Não foi possível abrir esta folha dentro da aplicação.');
 }
}

const $=id=>document.getElementById(id);
const defaults={settings:{salarioBase:1500,taxaSS:11,valorHora:8.65,nomeUtilizador:'',estadoCivil:'solteiro',dependentes:0,ajudaDia:90,alojDia:65,refeicaoDia:10.46,inicioDespesas:'2027-01-01',saldoInicial:41,dataCorte:'2026-07-25',diasFeriasAnuais:22,primeiroAnoFerias:2027,m25:.25,m125:1.25,m1375:1.375,m150:1.5,m165:1.65},sheets:[],used:[],payments:[],expenses:[],receipts:[],closedMonths:[]};
let db=load(),locations=[],pendingImports=[];
(function repairSalarySettings(){
 let changed=false;
 const fixed={m25:0.25,m125:1.25,m1375:1.375,m150:1.5,m165:1.65};
 for(const [k,v] of Object.entries(fixed)){
  if(num(db.settings[k])!==v){db.settings[k]=v;changed=true}
 }
 if(num(db.settings.salarioBase)<=0||num(db.settings.salarioBase)>10000){db.settings.salarioBase=1500;changed=true}
 if(num(db.settings.valorHora)<=0||num(db.settings.valorHora)>100){db.settings.valorHora=8.65;changed=true}
 for(const p of db.payments||[]){
  if(num(p.net)>5000||num(p.net)<0){p.net=null;changed=true}
 }
 if(changed)localStorage.setItem(KEY,JSON.stringify(db));
})();
normalizeStoredSheetHours();
function clone(x){return JSON.parse(JSON.stringify(x))}
function load(){try{const o=JSON.parse(localStorage.getItem(KEY)||'{}');return {...clone(defaults),...o,settings:{...defaults.settings,...(o.settings||{})},sheets:Array.isArray(o.sheets)?o.sheets:[],used:Array.isArray(o.used)?o.used:[],payments:Array.isArray(o.payments)?o.payments:[],expenses:Array.isArray(o.expenses)?o.expenses:[],receipts:Array.isArray(o.receipts)?o.receipts:[],closedMonths:Array.isArray(o.closedMonths)?o.closedMonths:[]}}catch{return clone(defaults)}}
function save(){localStorage.setItem(KEY,JSON.stringify(db));render()}
function num(v){if(typeof v==='number')return Number.isFinite(v)?v:0;if(v===null||v===undefined||v==='')return 0;const x=Number(String(v).trim().replace(/\s/g,'').replace(/€/g,'').replace(/\.(?=\d{3}(?:\D|$))/g,'').replace(',','.'));return Number.isFinite(x)?x:0}
function hourValue(v){
 if(v instanceof Date&&!isNaN(v))return v.getHours()+v.getMinutes()/60+v.getSeconds()/3600;
 const n=num(v);
 return n>0&&n<1?n*24:n;
}
function normalizeStoredSheetHours(){
 let changed=false;
 for(const s of db.sheets){
  if(!Array.isArray(s.entries)||!s.entries.length||s.hoursVersion===2)continue;
  const keys=['normal','h25','h125','h1375','h150','h165'];
  const positive=s.entries.flatMap(e=>keys.map(k=>num(e[k]))).filter(v=>v>0);
  const normalTotal=s.entries.reduce((a,e)=>a+num(e.normal),0);
  const looksLikeExcelFractions=positive.length>0&&Math.max(...positive)<=1&&normalTotal<=s.entries.length*1.25;
  if(looksLikeExcelFractions){
   for(const e of s.entries)for(const k of keys)e[k]=hourValue(e[k]);
   for(const k of keys)s[k]=s.entries.reduce((a,e)=>a+num(e[k]),0);
   changed=true;
  }
  s.hoursVersion=2;
 }
 if(changed)localStorage.setItem(KEY,JSON.stringify(db));
}

function euro(v){return num(v).toLocaleString('pt-PT',{style:'currency',currency:'EUR'})}
function fmt(v){return num(v).toLocaleString('pt-PT',{maximumFractionDigits:2})}
function norm(v){return String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim()}
function uid(){return crypto.randomUUID?crypto.randomUUID():'id-'+Date.now()+'-'+Math.random().toString(16).slice(2)}
function message(t,type='ok'){const d=document.createElement('div');d.className='msg '+type;d.textContent=t;$('messages').prepend(d)}
function monthOf(d){return String(d||'').slice(0,7)}
function today(){return new Date().toISOString().slice(0,10)}
function currentMonth(){return today().slice(0,7)}
function isClosed(m){return db.closedMonths.includes(m)}
async function hashBuffer(buf){if(crypto?.subtle){const h=await crypto.subtle.digest('SHA-256',buf);return [...new Uint8Array(h)].map(x=>x.toString(16).padStart(2,'0')).join('')}return String(buf.byteLength)}
function excelDate(v){if(v instanceof Date&&!isNaN(v))return v.toISOString().slice(0,10);if(typeof v==='number'&&window.XLSX){const d=XLSX.SSF.parse_date_code(v);if(d)return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`}if(typeof v==='string'){const m=v.match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);if(m)return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`}return ''}
function cell(ws,a){return ws[a]?.v??''}

function calculateHoursFromRow(ws,r){
 const dayType=norm(cell(ws,`B${r}`));
 const shift=norm(cell(ws,`Q${r}`));
 let service=hourValue(cell(ws,`P${r}`));

 // Quando a fórmula da coluna P não tem resultado guardado, calcular pelas horas de início/fim.
 if(!service){
  const start=cell(ws,`C${r}`),end=cell(ws,`H${r}`);
  const toHours=v=>{
   if(v instanceof Date&&!isNaN(v))return v.getHours()+v.getMinutes()/60+v.getSeconds()/3600;
   if(typeof v==='number')return v*24;
   if(typeof v==='string'){
    const m=v.match(/(\d{1,2}):(\d{2})/);
    if(m)return Number(m[1])+Number(m[2])/60;
   }
   return null;
  };
  const a=toHours(start),b=toHours(end);
  if(a!==null&&b!==null){
   service=b-a;
   if(service<0)service+=24;
  }
 }
 service=Math.max(0,service);

 const weekend=dayType.includes('sabado')||dayType.includes('feriado');
 const sunday=dayType.includes('domingo');
 const night=shift.includes('noturno');

 let normal=0,h25=0,h125=0,h1375=0,h150=0,h165=0;
 if(night){
  if(!weekend&&!sunday)h25=Math.min(service,8);
  if(service>8){
   h150=service===8.5?.5:1;
   h165=Math.max(0,service-h25-h150);
  }
 }else{
  if(!weekend&&!sunday)normal=Math.min(service,8);
  if(!weekend&&!sunday&&service>8){
   h125=service===8.5?.5:1;
   h1375=Math.max(0,service-normal-h125);
  }
  if(weekend||sunday)h150=service;
 }
 return {service,normal,h25,h125,h1375,h150,h165};
}

function findSheet(wb){return wb.Sheets['MD-RH-04 - Registo de Horas']||wb.Sheets[wb.SheetNames.find(n=>norm(n).includes('registo de horas'))]||wb.Sheets[wb.SheetNames[0]]}
function parseTimesheet(file,buffer,hash){
 const wb=XLSX.read(buffer,{type:'array',cellDates:true}),ws=findSheet(wb),entries=[];

 // As linhas diárias servem apenas para datas e cobertura do subsídio de refeição.
 // Nunca são usadas para calcular totais de horas.
 for(let r=15;r<=22;r++){
  const date=excelDate(cell(ws,`A${r}`));
  if(!date)continue;
  entries.push({
   date,
   dayType:String(cell(ws,`B${r}`)||'')
  });
 }

 if(!entries.length)throw new Error('Não foram encontradas datas nas linhas 15 a 22.');

 const dataInicial=entries.map(x=>x.date).sort()[0];
 const dataFinal=entries.map(x=>x.date).sort().at(-1);
 const cliente=String(cell(ws,'C10')||'').trim();
 const local=String(cell(ws,'I10')||'').trim();
 const processo=String(cell(ws,'M10')||'').trim();
 const diasAjuda=num(cell(ws,'J26'));
 const diasAloj=num(cell(ws,'J27'));

 // REGRA ÚNICA: ler apenas a linha TOTAL da folha.
 // Estes valores já estão em horas decimais de 30 em 30 minutos:
 // 0,5 = 30 min; 1 = 1 h; 1,5 = 1 h 30 min.
 const normal=num(cell(ws,'I23'));
 const h25=num(cell(ws,'J23'));
 const h125=num(cell(ws,'K23'));
 const h1375=num(cell(ws,'L23'));
 const h150=num(cell(ws,'M23'));
 const h165=num(cell(ws,'N23'));

 const compGerada=entries.reduce((a,x)=>{
  const t=norm(x.dayType);
  return a+(t.includes('domingo')?1:(t.includes('sabado')||t.includes('feriado')?.5:0));
 },0);

 return {
  id:uid(),name:file.name,hash,cliente,local,processo,dataInicial,dataFinal,entries,
  normal,h25,h125,h1375,h150,h165,
  diasAjuda,diasAloj,
  ajudaUnit:db.settings.ajudaDia,
  alojUnit:db.settings.alojDia,
  compGerada,
  compConta:dataFinal>db.settings.dataCorte?compGerada:0,
  hoursVersion:5,
  totalsSource:'I23:N23',
  originalKey:'sheet-'+hash,
  imported:new Date().toISOString(),
  _file:file
 };
}
function overtimeValue(g){
 const valorHora=num(db.settings.valorHora);
 // 100% não entra aqui. Fatores fixos:
 // 25% = 0,25; 125% = 1,25; 137,5% = 1,375; 150% = 1,5; 165% = 1,65.
 return valorHora*(
  num(g.h25)*0.25+
  num(g.h125)*1.25+
  num(g.h1375)*1.375+
  num(g.h150)*1.5+
  num(g.h165)*1.65
 );
}
function irs2026(R){
 R=Math.max(0,num(R));
 const dep=Math.max(0,Math.floor(num(db.settings.dependentes)));
 const estado=db.settings.estadoCivil||'solteiro';

 let t=0,a=0,depAbate=0;

 // Tabela III - casado, único titular.
 if(estado==='casado1'){
  depAbate=42.86;
  if(R<=991){t=0;a=0}
  else if(R<=1042){t=.125;a=.125*2.6*(1372.15-R)}
  else if(R<=1108){t=.125;a=.125*1.35*(1677.85-R)}
  else if(R<=1119){t=.125;a=96.17}
  else if(R<=1432){t=.1272;a=98.64}
  else if(R<=1962){t=.157;a=141.32}
  else if(R<=2240){t=.1938;a=213.53}
  else if(R<=2773){t=.2277;a=289.47}
  else if(R<=3389){t=.257;a=370.72}
  else if(R<=5965){t=.2881;a=476.12}
  else if(R<=20265){t=.3843;a=1049.96}
  else{t=.4717;a=2821.13}
 }else{
  // Tabela I: não casado sem dependentes ou casado 2 titulares.
  // Tabela II: não casado com um ou mais dependentes.
  depAbate=(estado==='solteiro'&&dep>0)?34.29:21.43;
  if(R<=920){t=0;a=0;depAbate=0}
  else if(R<=1042){t=.125;a=.125*2.6*(1273.85-R)}
  else if(R<=1108){t=.157;a=.157*1.35*(1554.83-R)}
  else if(R<=1154){t=.157;a=94.71}
  else if(R<=1212){t=.212;a=158.18}
  else if(R<=1819){t=.241;a=193.33}
  else if(R<=2119){t=.311;a=320.66}
  else if(R<=2499){t=.349;a=401.19}
  else if(R<=3305){t=.3836;a=487.66}
  else if(R<=5547){t=.3969;a=531.62}
  else if(R<=20221){t=.4495;a=823.40}
  else{t=.4717;a=1272.31}
 }

 const v=Math.max(0,R*t-a-depAbate*dep);
 return{value:v,effective:R?v/R:0};
}
function sheetsForMonth(m){return db.sheets.filter(s=>monthOf(s.dataFinal)===m||s.entries?.some(e=>monthOf(e.date)===m))}
function monthGroup(m){
 const g={normal:0,h25:0,h125:0,h1375:0,h150:0,h165:0,travelDays:0,lodgingDays:0,travel:0,lodging:0,comp:0};

 for(const s of db.sheets){
  if(monthOf(s.dataFinal)!==m)continue;

  // Somar cada folha uma única vez, usando apenas os totais oficiais I23:N23.
  g.normal+=num(s.normal);
  g.h25+=num(s.h25);
  g.h125+=num(s.h125);
  g.h1375+=num(s.h1375);
  g.h150+=num(s.h150);
  g.h165+=num(s.h165);

  g.travelDays+=num(s.diasAjuda);
  g.lodgingDays+=num(s.diasAloj);
  g.travel+=num(s.diasAjuda)*num(s.ajudaUnit||db.settings.ajudaDia);
  g.lodging+=num(s.diasAloj)*num(s.alojUnit||db.settings.alojDia);
  g.comp+=num(s.compConta);
 }

 return g;
}
function previousMonth(m){
 const [y,mo]=m.split('-').map(Number);
 const d=new Date(y,mo-2,1);
 return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
}
function subsidyNet(m){
 // Subsídios pagos separadamente: férias em julho e Natal em dezembro.
 // Para salário base de 1.500 €: 11% SS + 11,2% IRS = 1.167 € líquidos.
 const mo=String(m||'').slice(5,7);
 if(mo!=='07'&&mo!=='12')return 0;
 const base=num(db.settings.salarioBase);
 if(Math.abs(base-1500)<0.01)return 1167;
 // Se o salário base for alterado, manter a mesma incidência percentual
 // até existir uma tabela específica configurada para o subsídio.
 return Math.max(0,base-(base*num(db.settings.taxaSS)/100)-(base*11.2/100));
}
function salaryMonth(m){
 const g=monthGroup(m);
 const extras=monthGroup(previousMonth(m));
 const grossHours=overtimeValue(extras);
 const gross=db.settings.salarioBase+grossHours,
 ss=gross*db.settings.taxaSS/100,
 irsBase=irs2026(db.settings.salarioBase),
 irsHours=grossHours*(irsBase.effective/2),
 regularNet=gross-ss-irsBase.value-irsHours;
 const correction=db.payments.find(x=>x.month===m);
 const baseNet=correction?.net??regularNet;
 const subsidy=subsidyNet(m);
 return{...g,grossHours,gross,ss,irs:irsBase.value+irsHours,regularNet:baseNet,subsidy,net:baseNet+subsidy};
}
function easterDate(y){const a=y%19,b=Math.floor(y/100),c=y%100,d=Math.floor(b/4),e=b%4,f=Math.floor((b+8)/25),g=Math.floor((b-f+1)/3),h=(19*a+b-d-g+15)%30,i=Math.floor(c/4),k=c%4,l=(32+2*e+2*i-h-k)%7,m=Math.floor((a+11*h+22*l)/451),month=Math.floor((h+l-7*m+114)/31),day=((h+l-7*m+114)%31)+1;return new Date(y,month-1,day)}
function isoLocal(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`}
function holidays(y){const fixed=['01-01','04-25','05-01','06-10','08-15','10-05','11-01','12-01','12-08','12-25'].map(x=>`${y}-${x}`);const e=easterDate(y),good=new Date(e);good.setDate(e.getDate()-2);const corpus=new Date(e);corpus.setDate(e.getDate()+60);return new Set([...fixed,isoLocal(good),isoLocal(e),isoLocal(corpus)])}
function coveredByLeave(date){return db.used.some(x=>date>=x.date&&date<=(x.endDate||x.date))}
function sheetDates(){
 const set=new Set();
 for(const s of db.sheets){
  // Dias explicitamente registados.
  (s.entries||[]).forEach(e=>{if(e.date)set.add(e.date)});
  // A folha semanal cobre todo o intervalo entre a primeira e a última data.
  if(s.dataInicial&&s.dataFinal){
   const start=new Date(s.dataInicial+'T12:00:00'),end=new Date(s.dataFinal+'T12:00:00');
   for(let d=new Date(start);d<=end;d.setDate(d.getDate()+1)){
    const dow=d.getDay();
    if(dow!==0&&dow!==6)set.add(isoLocal(d));
   }
  }
 }
 return set;
}
function mealDays(m){
 if(!m)return 0;
 const [y,mo]=m.split('-').map(Number),monthStart=new Date(y,mo-1,1),now=new Date();
 const currentStart=new Date(now.getFullYear(),now.getMonth(),1);
 if(monthStart>currentStart)return 0;
 const lastDay=m===currentMonth()?now.getDate():new Date(y,mo,0).getDate();
 const hol=holidays(y),worked=sheetDates();
 let n=0;
 for(let d=1;d<=lastDay;d++){
  const dt=new Date(y,mo-1,d),iso=isoLocal(dt),dow=dt.getDay();
  if(dow!==0&&dow!==6&&!hol.has(iso)&&!worked.has(iso)&&!coveredByLeave(iso))n++;
 }
 return n;
}
function annualLeaveAdded(){const y=new Date().getFullYear(),start=num(db.settings.primeiroAnoFerias);return y<start?0:(y-start+1)*num(db.settings.diasFeriasAnuais)}
function balance(){return num(db.settings.saldoInicial)+db.sheets.reduce((a,s)=>a+num(s.compConta),0)+annualLeaveAdded()-db.used.reduce((a,x)=>a+num(x.days),0)}
function expenseSummary(m){const ex=db.expenses.filter(x=>monthOf(x.date)===m),g=monthGroup(m),food=ex.reduce((a,x)=>a+num(x.food),0),sleep=ex.reduce((a,x)=>a+num(x.sleep),0);return{food,sleep,travel:g.travel,lodging:g.lodging,remain:g.travel+g.lodging-food-sleep}}
function allMonths(){const set=new Set(db.sheets.flatMap(s=>(s.entries||[]).map(e=>monthOf(e.date))).concat(db.payments.map(x=>x.month),db.expenses.map(x=>monthOf(x.date)),db.receipts.map(x=>x.month)));set.add(currentMonth());return [...set].filter(Boolean).sort()}
function monthStatus(m){const warnings=[];if(!sheetsForMonth(m).length)warnings.push('sem folhas');if(!db.receipts.some(r=>r.month===m))warnings.push('sem recibo');if(m>='2027-01'&&monthGroup(m).travel+monthGroup(m).lodging>0&&!db.expenses.some(x=>monthOf(x.date)===m))warnings.push('sem despesas');const r=db.receipts.find(x=>x.month===m),p=salaryMonth(m);if(r?.net&&Math.abs(num(r.net)-p.net)>5)warnings.push('diferença no recibo');return warnings}
function renderDashboard(){const m=$('dashMonth').value||currentMonth(),p=salaryMonth(m),e=expenseSummary(m),extra=p.h25+p.h125+p.h1375+p.h150+p.h165,totalHours=p.normal+extra;const allExtra=db.sheets.reduce((a,s)=>a+num(s.h25)+num(s.h125)+num(s.h1375)+num(s.h150)+num(s.h165),0);$('dashNet').textContent=euro(num(p.net)>=300&&num(p.net)<=5000?num(p.net):0);$('dashH100').textContent=fmt(p.normal)+' h';$('dashH25').textContent=fmt(p.h25)+' h';$('dashH125').textContent=fmt(p.h125)+' h';$('dashH1375').textContent=fmt(p.h1375)+' h';$('dashH150').textContent=fmt(p.h150)+' h';$('dashH165').textContent=fmt(p.h165)+' h';$('dashExtra').textContent=fmt(extra)+' h';$('dashTotalHours').textContent=fmt(totalHours)+' h';$('dashExtraAll').textContent=fmt(allExtra)+' h';const mealCount=mealDays(m);$('dashMeal').textContent=`${euro(mealCount*db.settings.refeicaoDia)} (${mealCount} dias sem folha)`;$('dashTravel').textContent=euro(p.travel);$('dashLodging').textContent=euro(p.lodging);$('dashBalance').textContent=fmt(balance())+' dias';$('dashExpenseBalance').textContent=euro(e.remain);const w=monthStatus(m),closed=isClosed(m);$('monthStatus').className='statusBox '+(w.length?'status-warn':'status-ok');$('monthStatus').innerHTML=`<strong>${closed?'🔒 Mês fechado':'🟢 Mês aberto'}</strong>${w.length?' · Falta: '+w.join(', '):' · Registos completos'}`;renderMonthly();renderAnnual()}
function renderMonthly(){
 let h='<tr><th>Mês</th><th>Líquido recibo</th><th>Horas a 100%</th><th>25%</th><th>125%</th><th>137,5%</th><th>150%</th><th>165%</th><th>Total extra</th><th>Total geral</th><th>Subs. refeição</th><th>Ajudas</th><th>Alojamento</th><th>Estado</th></tr>';
 for(const m of allMonths().reverse()){
  const p=salaryMonth(m);
  const extra=p.h25+p.h125+p.h1375+p.h150+p.h165;
  const totalHours=p.normal+extra;
  const w=monthStatus(m);
  h+=`<tr>
   <td>${m}</td>
   <td>${euro(num(p.net)>=300&&num(p.net)<=5000?num(p.net):0)}</td>
   <td><strong>${fmt(p.normal)} h</strong></td>
   <td>${fmt(p.h25)} h</td>
   <td>${fmt(p.h125)} h</td>
   <td>${fmt(p.h1375)} h</td>
   <td>${fmt(p.h150)} h</td>
   <td>${fmt(p.h165)} h</td>
   <td><strong>${fmt(extra)} h</strong></td>
   <td><strong>${fmt(totalHours)} h</strong></td>
   <td>${euro(mealDays(m)*db.settings.refeicaoDia)}</td>
   <td>${euro(p.travel)}</td>
   <td>${euro(p.lodging)}</td>
   <td>${isClosed(m)?'<span class="closedBadge">Fechado</span>':w.length?'<span class="closedBadge">'+w.join(', ')+'</span>':'<span class="openBadge">Completo</span>'}</td>
  </tr>`;
 }
 $('monthlyTable').innerHTML=h;
}
function renderAnnual(){
 const years={};
 for(const m of allMonths()){
  const y=m.slice(0,4),p=salaryMonth(m);
  years[y]??={net:0,normal:0,h25:0,h125:0,h1375:0,h150:0,h165:0,meal:0,travel:0,lodging:0};
  const x=years[y];
  x.net+=(num(p.net)>=300&&num(p.net)<=5000?num(p.net):0);x.normal+=p.normal;x.h25+=p.h25;x.h125+=p.h125;x.h1375+=p.h1375;x.h150+=p.h150;x.h165+=p.h165;
  x.meal+=mealDays(m)*db.settings.refeicaoDia;x.travel+=p.travel;x.lodging+=p.lodging;
 }
 let h='<tr><th>Ano</th><th>Líquido recibos</th><th>Horas a 100%</th><th>25%</th><th>125%</th><th>137,5%</th><th>150%</th><th>165%</th><th>Total extra</th><th>Total geral</th><th>Subs. refeição</th><th>Ajudas</th><th>Alojamento</th></tr>';
 for(const y of Object.keys(years).sort().reverse()){
  const x=years[y],extra=x.h25+x.h125+x.h1375+x.h150+x.h165,total=x.normal+extra;
  h+=`<tr>
   <td>${y}</td><td>${euro(x.net)}</td><td><strong>${fmt(x.normal)} h</strong></td>
   <td>${fmt(x.h25)} h</td><td>${fmt(x.h125)} h</td><td>${fmt(x.h1375)} h</td><td>${fmt(x.h150)} h</td><td>${fmt(x.h165)} h</td>
   <td><strong>${fmt(extra)} h</strong></td><td><strong>${fmt(total)} h</strong></td>
   <td>${euro(x.meal)}</td><td>${euro(x.travel)}</td><td>${euro(x.lodging)}</td>
  </tr>`;
 }
 $('annualTable').innerHTML=h;
}
function renderSheets(){const q=norm($('sheetSearch').value);let h='<tr><th>Ficheiro</th><th>Período</th><th>Cliente</th><th>Local</th><th>100%</th><th>Extra</th><th>Ajudas</th><th>Aloj.</th><th></th></tr>';for(const s of [...db.sheets].sort((a,b)=>b.dataFinal.localeCompare(a.dataFinal))){if(q&&!norm(`${s.name} ${s.cliente} ${s.local} ${s.processo}`).includes(q))continue;const ex=num(s.h25)+num(s.h125)+num(s.h1375)+num(s.h150)+num(s.h165);h+=`<tr><td>${s.name}</td><td>${s.dataInicial}<br>${s.dataFinal}</td><td>${s.cliente}</td><td>${s.local}</td><td>${fmt(s.normal)} h</td><td>${fmt(ex)} h</td><td>${fmt(s.diasAjuda)} dias</td><td>${fmt(s.diasAloj)} dias</td><td>${s.originalKey?`<button onclick="openStoredSheet('${s.originalKey}','${String(s.name).replace(/'/g,"&#39;")}')">Abrir</button> `:''}<button onclick="removeSheet('${s.id}')">Apagar</button></td></tr>`}$('sheetsTable').innerHTML=h}
function renderPayments(){
 const years=[...new Set(allMonths().map(m=>m.slice(0,4)))].sort().reverse(),
 sel=$('payYear').value||years[0]||String(new Date().getFullYear());
 $('payYear').innerHTML=years.map(y=>`<option ${y===sel?'selected':''}>${y}</option>`).join('');

 let sums={net:0,travel:0,lodging:0,meal:0,saved:0},
 h='<tr><th>Mês</th><th>Bruto</th><th>SS</th><th>IRS</th><th>Líquido recibo</th><th>Ajudas</th><th>Alojamento</th><th>Subs. refeição</th><th>Recibo</th></tr>';

 const months=allMonths().filter(x=>x.startsWith(sel)).reverse();
 for(const m of months){
  const p=salaryMonth(m),receipt=db.receipts.find(r=>r.month===m);
  sums.net+=p.net;sums.travel+=p.travel;sums.lodging+=p.lodging;sums.meal+=mealDays(m)*db.settings.refeicaoDia;

  // A poupança anual conta apenas meses onde foram efetivamente inseridas despesas.
  if(db.expenses.some(x=>monthOf(x.date)===m)){
   sums.saved+=expenseSummary(m).remain;
  }

  h+=`<tr><td>${m}</td><td>${euro(p.gross)}</td><td>- ${euro(p.ss)}</td><td>- ${euro(p.irs)}</td><td><strong>${euro(p.net)}</strong></td><td>${euro(p.travel)}</td><td>${euro(p.lodging)}</td><td>${euro(mealDays(m)*db.settings.refeicaoDia)}</td><td>${receipt?'✅':'⚠️ falta'}</td></tr>`;
 }

 $('yearNet').textContent=euro(sums.net);
 $('yearTravel').textContent=euro(sums.travel);
 $('yearLodging').textContent=euro(sums.lodging);
 $('yearMeal').textContent=euro(sums.meal);
 $('yearSaved').textContent=euro(sums.saved);

 const expenseMonths=[...new Set(db.expenses.map(x=>monthOf(x.date)).filter(m=>m&&m.startsWith(sel)))].sort();
 if(expenseMonths.length){
  const first=expenseMonths[0];
  const [y,mo]=first.split('-');
  const label=new Intl.DateTimeFormat('pt-PT',{month:'short',year:'numeric'}).format(new Date(+y,+mo-1,1));
  $('yearSavedLabel').textContent=`Total poupado nas despesas (desde ${label})`;
 }else{
  $('yearSavedLabel').textContent='Total poupado nas despesas';
 }

 $('paymentsTable').innerHTML=h;
}
function latestSheetMonth(){
 const months=db.sheets.flatMap(s=>Array.isArray(s.entries)&&s.entries.length?s.entries.map(e=>monthOf(e.date)):[monthOf(s.dataFinal)]).filter(Boolean).sort();
 return months.at(-1)||currentMonth();
}

function clientsPdfInfo(){
 const nome=String(db.settings.nomeUtilizador||'').trim();
 if(!nome){message('Indica primeiro o Nome do utilizador nas Definições e guarda.','err');return null}
 const m=$('clientMonth').value||latestSheetMonth();
 const table=$('clientsTable');
 const trs=[...table.querySelectorAll('tr')];
 if(trs.length<2){message('Não existem horas para partilhar neste mês.','err');return null}
 const [y,mo]=m.split('-');
 const mes=new Intl.DateTimeFormat('pt-PT',{month:'long',year:'numeric'}).format(new Date(+y,+mo-1,1));
 return {nome,m,mes,table};
}
function imageToDataUrl(src){
 return new Promise((resolve)=>{
  const img=new Image();
  img.onload=()=>{
   try{
    const c=document.createElement('canvas');
    c.width=img.naturalWidth||img.width;c.height=img.naturalHeight||img.height;
    c.getContext('2d').drawImage(img,0,0);
    resolve(c.toDataURL('image/png'));
   }catch{resolve(null)}
  };
  img.onerror=()=>resolve(null);
  img.src=src;
 });
}
async function buildClientsPdf(){
 const info=clientsPdfInfo(); if(!info)return null;
 if(!window.jspdf?.jsPDF){
  message('O gerador de PDF ainda não carregou. Confirma a ligação à Internet e tenta novamente.','err');
  return null;
 }
 const {jsPDF}=window.jspdf;
 const doc=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
 const brand=[11,94,117],dark=[7,63,80],soft=[233,245,247];

 // Cabeçalho colorido
 doc.setFillColor(...brand);doc.rect(0,0,297,31,'F');
 const logo=await imageToDataUrl('logotipo.png');
 if(logo){
  try{
   // O logótipo original é bastante horizontal. Mantém a proporção original,
   // em vez de o forçar para uma caixa 48 x 18 mm.
   const logoW=48;
   const logoRatio=1796/369;
   const logoH=logoW/logoRatio;
   const logoY=6+(18-logoH)/2;
   doc.addImage(logo,'PNG',12,logoY,logoW,logoH,undefined,'FAST');
  }catch{}
 }
 doc.setTextColor(255,255,255);
 doc.setFont('helvetica','bold');doc.setFontSize(17);
 doc.text('Horas por Cliente',68,13);
 doc.setFont('helvetica','normal');doc.setFontSize(10.5);
 doc.text(`Utilizador: ${info.nome}`,68,20);
 doc.text(`Mês: ${info.mes}`,68,26);

 // Converter a tabela visível de Clientes em dados para PDF,
 // respeitando colspan para que a linha TOTAL DO MÊS fique alinhada com as colunas.
 const trs=[...info.table.querySelectorAll('tr')];
 const expandRow=(tr)=>{
  const out=[];
  for(const c of tr.querySelectorAll('th,td')){
   const span=Math.max(1,Number(c.getAttribute('colspan')||1));
   out.push(c.innerText.trim());
   for(let i=1;i<span;i++)out.push('');
  }
  return out;
 };
 const head=expandRow(trs[0]);
 const body=trs.slice(1).map(expandRow);

 doc.autoTable({
  head:[head],
  body,
  startY:37,
  margin:{left:8,right:8},
  theme:'grid',
  styles:{font:'helvetica',fontSize:7.3,cellPadding:2.2,textColor:dark,lineColor:[213,228,232],lineWidth:.2,halign:'center',valign:'middle'},
  headStyles:{fillColor:soft,textColor:dark,fontStyle:'bold',lineColor:brand,lineWidth:.25},
  alternateRowStyles:{fillColor:[248,251,252]},
  columnStyles:{0:{halign:'left',cellWidth:35},1:{halign:'left',cellWidth:30}},
  didParseCell:(data)=>{
   const first=Array.isArray(data.row.raw)?String(data.row.raw[0]||''):'';
   if(data.section==='body' && first.includes('TOTAL DO MÊS')){
    data.cell.styles.fillColor=brand;
    data.cell.styles.textColor=[255,255,255];
    data.cell.styles.fontStyle='bold';
   }
  }
 });

 const y=Math.min(199,(doc.lastAutoTable?.finalY||40)+8);
 doc.setDrawColor(...brand);doc.line(8,y,289,y);
 doc.setTextColor(95,119,128);doc.setFontSize(7.5);doc.setFont('helvetica','normal');
 doc.text('Gerado pela aplicação Controlo Horas e Compensações',8,y+5);
 doc.text(`Data: ${new Intl.DateTimeFormat('pt-PT').format(new Date())}`,289,y+5,{align:'right'});

 return {doc,info};
}
async function downloadClientsPdf(){
 const built=await buildClientsPdf();if(!built)return;
 const safe=built.info.nome.replace(/[^\p{L}\p{N}]+/gu,'_');
 built.doc.save(`horas_clientes_${safe}_${built.info.m}.pdf`);
 message('PDF das horas criado.','ok');
}
async function shareClientsPdf(){
 const built=await buildClientsPdf();if(!built)return;
 const blob=built.doc.output('blob');
 const safe=built.info.nome.replace(/[^\p{L}\p{N}]+/gu,'_');
 const file=new File([blob],`horas_clientes_${safe}_${built.info.m}.pdf`,{type:'application/pdf'});

 // Primeiro tenta partilhar o PDF diretamente no telemóvel.
 try{
  if(window.isSecureContext && typeof navigator.share==='function' && (!navigator.canShare || navigator.canShare({files:[file]}))){
   await navigator.share({
    title:`Horas por cliente - ${built.info.nome} - ${built.info.mes}`,
    text:`Horas por cliente - ${built.info.nome} - ${built.info.mes}`,
    files:[file]
   });
   return;
  }
 }catch(e){
  if(e?.name==='AbortError')return;
 }

 // Se o browser não suportar ficheiros na partilha, guarda o PDF para o utilizador o enviar.
 const url=URL.createObjectURL(blob),a=document.createElement('a');
 a.href=url;a.download=file.name;document.body.appendChild(a);a.click();a.remove();
 setTimeout(()=>URL.revokeObjectURL(url),1000);
 message('O PDF foi guardado. O navegador deste dispositivo não permite enviá-lo diretamente pela partilha.','ok');
}

function renderClients(){
 const m=$('clientMonth').value||latestSheetMonth(),q=norm($('clientFilter').value),groups={};
 if(!$('clientMonth').value)$('clientMonth').value=m;

 // Na página Clientes, cada folha contribui apenas uma vez com os totais oficiais da folha.
 // Não somamos novamente as linhas diárias, porque isso podia duplicar ou distorcer as horas.
 for(const s of db.sheets){
  if(monthOf(s.dataFinal)!==m)continue;
  const key=`${s.cliente||'Sem cliente'}|||${s.local||'Sem local'}`;
  groups[key]??={client:s.cliente||'Sem cliente',local:s.local||'Sem local',normal:0,h25:0,h125:0,h1375:0,h150:0,h165:0,folhas:0};
  for(const k of ['normal','h25','h125','h1375','h150','h165'])groups[key][k]+=num(s[k]);
  groups[key].folhas++;
 }

 let h='<tr><th>Cliente</th><th>Local</th><th>Folhas</th><th>100%</th><th>25%</th><th>125%</th><th>137,5%</th><th>150%</th><th>165%</th><th>Total extra</th><th>Total geral</th></tr>';
 const rows=Object.values(groups)
  .filter(x=>!q||norm(x.client+' '+x.local).includes(q))
  .sort((a,b)=>a.client.localeCompare(b.client)||a.local.localeCompare(b.local));

 const total={normal:0,h25:0,h125:0,h1375:0,h150:0,h165:0,folhas:0};

 for(const x of rows){
  const extra=x.h25+x.h125+x.h1375+x.h150+x.h165;
  const totalHours=x.normal+extra;
  for(const k of Object.keys(total))total[k]+=num(x[k]);
  h+=`<tr>
   <td>${x.client}</td><td>${x.local}</td><td>${x.folhas}</td>
   <td><strong>${fmt(x.normal)}</strong></td>
   <td>${fmt(x.h25)}</td><td>${fmt(x.h125)}</td><td>${fmt(x.h1375)}</td><td>${fmt(x.h150)}</td><td>${fmt(x.h165)}</td>
   <td><strong>${fmt(extra)}</strong></td><td><strong>${fmt(totalHours)}</strong></td>
  </tr>`;
 }

 if(rows.length){
  const totalExtra=total.h25+total.h125+total.h1375+total.h150+total.h165;
  const grand=total.normal+totalExtra;
  h+=`<tr>
   <th colspan="2">TOTAL DO MÊS</th><th>${total.folhas}</th>
   <th>${fmt(total.normal)}</th><th>${fmt(total.h25)}</th><th>${fmt(total.h125)}</th><th>${fmt(total.h1375)}</th><th>${fmt(total.h150)}</th><th>${fmt(total.h165)}</th>
   <th>${fmt(totalExtra)}</th><th>${fmt(grand)}</th>
  </tr>`;
 }else{
  h+=`<tr><td colspan="11">Não existem folhas para ${m}. O último mês com folhas é ${latestSheetMonth()}.</td></tr>`;
 }

 $('clientsTable').innerHTML=h;
}
function renderExpenses(){const m=$('expenseMonth').value||currentMonth(),s=expenseSummary(m);$('expenseTravelReceived').textContent=euro(s.travel);$('expenseFoodSpent').textContent=euro(s.food);$('expenseLodgingReceived').textContent=euro(s.lodging);$('expenseSleepSpent').textContent=euro(s.sleep);$('expenseRemain').textContent=euro(s.remain);let h='<tr><th>Data</th><th>Alimentação</th><th>Dormida</th><th>Observação</th><th></th></tr>';for(const x of db.expenses.filter(x=>monthOf(x.date)===m).sort((a,b)=>b.date.localeCompare(a.date))){h+=`<tr><td>${x.date}</td><td>${euro(x.food)}</td><td>${euro(x.sleep)}</td><td>${x.note||''}</td><td><button onclick="removeExpense('${x.id}')">Apagar</button></td></tr>`}$('expensesTable').innerHTML=h}
function renderLeave(){const used=db.used.reduce((a,x)=>a+num(x.days),0);$('balanceCard').textContent=fmt(balance())+' dias';$('annualLeaveCard').textContent=fmt(annualLeaveAdded())+' dias';$('usedDaysCard').textContent=fmt(used)+' dias';let h='<tr><th>Tipo</th><th>Período</th><th>Dias</th><th>Descrição</th><th></th></tr>';for(const x of [...db.used].sort((a,b)=>b.date.localeCompare(a.date))){h+=`<tr><td>${x.type}</td><td>${x.date}${x.endDate&&x.endDate!==x.date?' a '+x.endDate:''}</td><td>${fmt(x.days)}</td><td>${x.desc||''}</td><td><button onclick="removeLeave('${x.id}')">Apagar</button></td></tr>`}$('leaveTable').innerHTML=h}
window.openWaze=function(lat,lon){
 const appUrl=`waze://?ll=${lat},${lon}&navigate=yes`;
 const webUrl=`https://waze.com/ul?ll=${lat},${lon}&navigate=yes`;
 const started=Date.now();
 window.location.href=appUrl;
 setTimeout(()=>{
  if(document.visibilityState==='visible'&&Date.now()-started<3000)window.location.href=webUrl;
 },1600);
};
function renderLocations(){
 const q=norm($('locationSearch').value),source=$('locationSource').value;
 const rows=locations.filter(x=>(!source||x.source===source)&&(!q||norm(`${x.name} ${x.address} ${x.source}`).includes(q))).slice(0,100);
 $('locationResults').innerHTML=rows.map(x=>`<article class="locationCard"><h3>${x.name}</h3><p><strong>${x.source}</strong></p><p>${x.address||'Sem morada'}</p>${x.lat&&x.lon?`<p>${x.lat.toFixed(6)}, ${x.lon.toFixed(6)}</p><div class="locationActions"><a target="_blank" href="https://www.google.com/maps/search/?api=1&query=${x.lat},${x.lon}"><button>Google Maps</button></a><button type="button" onclick="openWaze(${x.lat},${x.lon})">Waze</button></div>`:'<p>Coordenadas não disponíveis</p>'}</article>`).join('')||'<p>Sem resultados.</p>';
}

function parsePtNumber(s){return num(String(s||'').replace(/\s/g,''))}
function monthNameToNumber(name){const m={janeiro:'01',fevereiro:'02',marco:'03',março:'03',abril:'04',maio:'05',junho:'06',julho:'07',agosto:'08',setembro:'09',outubro:'10',novembro:'11',dezembro:'12'};return m[norm(name)]||''}
function parseReceiptText(text){
 const clean=text.replace(/\u00a0/g,' ').replace(/[ \t]+/g,' ');
 const mm=clean.match(/\b(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+(20\d{2})\b/i);
 const month=mm?`${mm[2]}-${monthNameToNumber(mm[1])}`:'';
 const line=(label)=>{const i=norm(clean).indexOf(norm(label));return i<0?'':clean.slice(i,i+180)};
 const money=(s)=>[...s.matchAll(/(\d{1,3}(?:[.\s]\d{3})*|\d+)[,.](\d{2})\s*€/g)].map(m=>parsePtNumber(m[0]));
 const baseVals=money(line('Ordenado Base')),base=baseVals.at(-1)??null;
 // Reconhece "Hora Extra" mesmo quando o PDF perde o símbolo € ou separa colunas.
 const extraMatch=clean.match(/Hora\s*Extra[\s\S]{0,180}/i);
 const extraPart=extraMatch?extraMatch[0]:'';
 const extraNums=[...extraPart.matchAll(/\b(\d+(?:[,.]\d+)?)\b/g)].map(m=>parsePtNumber(m[1]));
 const extraMoney=money(extraPart);
 const extraHours=extraNums.length?extraNums[0]:null;
 // Normalmente: quantidade, valor unitário, remuneração. Preferir o último valor numérico
 // quando não houver símbolo €, mas ignorar percentagens/códigos muito à frente.
 let extraValue=extraMoney.length?extraMoney.at(-1):null;
 if(extraValue==null&&extraNums.length>=3) extraValue=extraNums[2];
 if(extraValue==null&&extraNums.length>=2) extraValue=extraNums.at(-1);
 const ssVals=money(line('Segurança Social')),ss=ssVals.at(-1)??null;
 const irsParts=[...clean.matchAll(/Imposto\s*S\/Rendimento[\s\S]{0,120}?(\d+(?:[,.]\d{2}))\s*€/gi)].map(m=>parsePtNumber(m[1]));
 const irs=irsParts.length?irsParts.reduce((a,b)=>a+b,0):null;
 const amounts=money(clean);let subject=null,discounts=null,net=null;

 // REGRA PRIORITÁRIA PARA RECIBOS EM QUE OS CABEÇALHOS DOS TOTAIS
 // ESTÃO NUMA PÁGINA E OS VALORES NA PÁGINA SEGUINTE.
 // Exemplo real validado: 1694,63 € | 364,41 € | 1330,22 €.
 // "Total a Receber" é a coluna mais à direita, logo escolhemos o valor
 // monetário mais à direita na página seguinte quando a página anterior
 // contém os cabeçalhos Total Sujeito / Total de Descontos / Total a Receber.
 const posBlocks=[...clean.matchAll(/\[\[POS:(\d+)\]\]([\s\S]*?)\[\[\/POS\]\]/g)];
 const pagesOnly=clean.split(/\[\[PAGE:\d+\]\]/);
 if(/Total\s+Sujeito/i.test(clean) && /Total\s+(?:de\s+)?Descontos/i.test(clean) && /Total\s+a\s+Receber/i.test(clean)){
   for(const pb of posBlocks){
     const pageNo=Number(pb[1]);
     if(pageNo<2)continue;
     const vals=pb[2].split('||').map(s=>{
       const [x,y,...rest]=s.split('::');
       return {x:Number(x),y:Number(y),raw:rest.join('::'),v:parsePtNumber(rest.join('::'))};
     }).filter(o=>Number.isFinite(o.v) && o.v>=100);

     // No recibo real, os três totais aparecem na mesma faixa horizontal
     // e o Total a Receber é o mais à direita.
     if(vals.length>=3){
       const groups=[];
       for(const v of vals){
         let g=groups.find(g=>Math.abs(g.y-v.y)<18);
         if(!g){g={y:v.y,vals:[]};groups.push(g)}
         g.vals.push(v);
       }
       const totalsRow=groups.filter(g=>g.vals.length>=3).sort((a,b)=>b.y-a.y)[0];
       if(totalsRow){
         const rightmost=[...totalsRow.vals].sort((a,b)=>b.x-a.x)[0];
         if(rightmost){net=rightmost.v;break}
       }
     }
   }
 }

 const pageParts=clean.split(/\[\[PAGE:\d+\]\]/).map(x=>x.trim()).filter(Boolean);

 const readMoneyLoose=(s)=>[...s.matchAll(/(\d{1,3}(?:[.\s]\d{3})*|\d+)[,.](\d{2})\s*€?/g)]
   .map(m=>parsePtNumber(m[0]))
   .filter(v=>Number.isFinite(v));

 const labels=[/Total\s*a\s*Receber/i,/Total\s*L[ií]quido/i,/L[ií]quido\s*a\s*Receber/i];

 // A) Procurar o rótulo em cada página, apenas se a regra posicional ainda não encontrou o líquido.
 // Se o valor não estiver na mesma página, procurar no início da página seguinte.
 outer:
 if(net==null) for(let i=0;i<pageParts.length;i++){
  const page=pageParts[i];
  for(const rx of labels){
   const mm=rx.exec(page);
   if(!mm)continue;

   const after=page.slice(mm.index+mm[0].length);
   const samePage=readMoneyLoose(after).filter(v=>v>=100);
   if(samePage.length){
    net=samePage[0];
    break outer;
   }

   // Caso típico dos recibos de 2 páginas:
   // "Total a Receber" aparece no fim da página 1 e o valor só aparece na página 2.
   if(i+1<pageParts.length){
    const nextHead=pageParts[i+1].slice(0,700);
    const vals=readMoneyLoose(nextHead).filter(v=>v>=100);

    // Primeiro tentar reconhecer trio: total sujeito - descontos = líquido.
    let found=null;
    for(const a of vals)for(const b of vals)for(const c of vals){
     if(a>b && Math.abs((a-b)-c)<0.03){
      const expectedDiscounts=(ss!=null?ss:0)+(irs!=null?irs:0);
      const penalty=expectedDiscounts>0?Math.abs(b-expectedDiscounts):0;
      const cand={subject:a,discounts:b,net:c,penalty};
      if(!found || cand.penalty<found.penalty || (cand.penalty===found.penalty && cand.net>found.net))found=cand;
     }
    }
    if(found){
     subject=found.subject;discounts=found.discounts;net=found.net;
     break outer;
    }

    // Se não houver trio, usar o último valor monetário do topo da página seguinte,
    // pois nos recibos reais o Total a Receber aparece à direita depois dos outros totais.
    if(vals.length){
     net=vals.at(-1);
     break outer;
    }
   }
  }
 }

 // B) Compatibilidade com PDFs sem marcador de página ou texto linear.
 if(net==null){
  const directPatterns=[
   /Total\s*a\s*Receber[\s\S]{0,120}?(\d{1,3}(?:[.\s]\d{3})*|\d+)[,.](\d{2})\s*€?/i,
   /Total\s*L[ií]quido[\s\S]{0,120}?(\d{1,3}(?:[.\s]\d{3})*|\d+)[,.](\d{2})\s*€?/i,
   /L[ií]quido\s*a\s*Receber[\s\S]{0,120}?(\d{1,3}(?:[.\s]\d{3})*|\d+)[,.](\d{2})\s*€?/i
  ];
  for(const rx of directPatterns){
   const mm=clean.match(rx);
   if(mm){
    net=parsePtNumber(`${mm[1]},${mm[2]}`);
    break;
   }
  }
 }

 // C) Último fallback: relação total sujeito/remunerações - descontos = líquido.
 if(net==null){
  const expectedDiscounts=(ss!=null?ss:0)+(irs!=null?irs:0);
  const candidates=[];
  for(let i=0;i<amounts.length;i++){
   for(let j=0;j<amounts.length;j++){
    if(i===j)continue;
    for(let k=0;k<amounts.length;k++){
     if(k===i||k===j)continue;
     const a=amounts[i],b=amounts[j],c=amounts[k];
     if(a>b && c>=100 && Math.abs((a-b)-c)<0.03){
      const penalty=expectedDiscounts>0?Math.abs(b-expectedDiscounts):0;
      candidates.push({subject:a,discounts:b,net:c,penalty});
     }
    }
   }
  }
  if(candidates.length){
   candidates.sort((x,y)=>(x.penalty-y.penalty)||(y.net-x.net));
   ({subject,discounts,net}=candidates[0]);
  }
 }

 return {month,base,extraHours,extraValue,ss,irs,subject,discounts,net,text:clean};
}
async function extractReceipt(file){
 $('receiptProgress').textContent='A ler o recibo automaticamente…';
 let text='';

 if(file.type==='application/pdf'||file.name?.toLowerCase().endsWith('.pdf')){
  const pdfjs=await import('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs');
  pdfjs.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs';
  const pdf=await pdfjs.getDocument({data:await file.arrayBuffer()}).promise;

  for(let p=1;p<=pdf.numPages;p++){
   const page=await pdf.getPage(p),content=await page.getTextContent();
   const items=content.items.map(x=>({str:x.str,x:x.transform?.[4]??0,y:x.transform?.[5]??0}));
   const pageText=items.map(x=>x.str).join(' ');
   text+=`\n[[PAGE:${p}]]\n${pageText}\n`;

   // Guardar também os valores monetários com posição horizontal.
   // Isto é necessário nos recibos Ambicare: os cabeçalhos dos totais ficam na pág. 1
   // e os três valores correspondentes aparecem isolados na pág. 2.
   const positioned=items
     .filter(it=>/\d[\d.\s]*[,\.]\d{2}\s*€?/.test(it.str))
     .map(it=>`${it.x.toFixed(1)}::${it.y.toFixed(1)}::${it.str}`)
     .join('||');
   text+=`\n[[POS:${p}]]${positioned}[[/POS]]\n`;
  }

  // Alguns recibos PDF são essencialmente imagem e o PDF.js devolve pouco/nenhum texto.
  // Nesses casos fazemos OCR da primeira página, onde está a rubrica "Hora Extra".
  if(!/Hora\s*Extra/i.test(text) || text.replace(/\s+/g,'').length<80){
   $('receiptProgress').textContent='A ler o recibo por imagem…';
   const page=await pdf.getPage(1);
   const viewport=page.getViewport({scale:2});
   const canvas=document.createElement('canvas');
   const ctx=canvas.getContext('2d');
   canvas.width=Math.ceil(viewport.width);
   canvas.height=Math.ceil(viewport.height);
   await page.render({canvasContext:ctx,viewport}).promise;

   const result=await Tesseract.recognize(canvas,'por',{
    logger:m=>{
     if(m.status==='recognizing text')
      $('receiptProgress').textContent=`A ler recibo: ${Math.round((m.progress||0)*100)}%`;
    }
   });
   text+=' '+result.data.text;
  }
 }else{
  const result=await Tesseract.recognize(file,'por',{
   logger:m=>{
    $('receiptProgress').textContent=m.status==='recognizing text'
     ?`A ler fotografia: ${Math.round((m.progress||0)*100)}%`
     :'A preparar leitura…';
   }
  });
  text=result.data.text;
 }

 $('receiptProgress').textContent='';
 return parseReceiptText(text);
}
function refreshReceiptParsed(receipt){
 if(!receipt?.parsed)return receipt?.parsed;
 const old=receipt.parsed;
 if(old.extraValue==null&&old.text){
  const fresh=parseReceiptText(old.text);
  receipt.parsed={...old,...fresh,month:fresh.month||old.month||receipt.month};
  if(receipt.parsed.month)receipt.month=receipt.parsed.month;
 }
 return receipt.parsed;
}
async function ensureReceiptExtraValue(r){
 if(!r?.parsed)return false;
 if(r.parsed.extraValue!=null)return true;

 // Primeiro tenta novamente o texto já guardado.
 if(r.parsed.text){
  const fresh=parseReceiptText(r.parsed.text);
  r.parsed={...r.parsed,...fresh,month:fresh.month||r.parsed.month||r.month};
  if(r.parsed.extraValue!=null){
   save();
   return true;
  }
 }

 // Se o PDF antigo não tinha camada de texto, relê o ficheiro original por OCR.
 if(r.fileKey){
  const blob=await getStoredFile(r.fileKey);
  if(blob){
   $('receiptProgress').textContent='A reler o recibo para encontrar Hora Extra…';
   const file=new File([blob],r.name||'recibo.pdf',{type:blob.type||r.type||'application/pdf'});
   const fresh=await extractReceipt(file);
   r.parsed={...r.parsed,...fresh,month:fresh.month||r.parsed.month||r.month};
   if(r.parsed.month)r.month=r.parsed.month;
   save();
   return r.parsed.extraValue!=null;
  }
 }
 return false;
}

function comparisonRows(receipt){
 const p=salaryMonth(receipt.month);
 const x={label:'Total horas extra',expected:p.grossHours,actual:receipt.extraValue,unit:'€'};
 return [{...x,diff:x.actual==null?null:x.actual-x.expected,ok:x.actual!=null&&Math.abs(x.actual-x.expected)<0.05}];
}

function repairStoredReceiptNet(r){
 if(!r?.parsed?.text)return false;
 const fresh=parseReceiptText(r.parsed.text);
 if(fresh.net!=null && (r.parsed.net==null || Math.abs(num(fresh.net)-num(r.parsed.net))>0.001)){
  r.parsed={...r.parsed,...fresh,month:fresh.month||r.parsed.month||r.month};
  return true;
 }
 return false;
}

function receiptMonthSummary(month){
 const items=db.receipts.filter(r=>r.month===month&&r.parsed);
 const values=items.map(r=>r.parsed?.net).filter(v=>v!=null&&isFinite(Number(v))).map(Number);
 const actual=values.length?values.reduce((a,b)=>a+b,0):null;
 const expected=salaryMonth(month).net;
 const special=['07','12'].includes(String(month||'').slice(5,7));
 const complete=!special||items.length>=2;
 const diff=actual==null?null:actual-expected;
 return{items,actual,expected,diff,special,complete};
}
function renderReceipts(){
 let repaired=false;
 for(const r of db.receipts)if(repairStoredReceiptNet(r))repaired=true;
 if(repaired)localStorage.setItem(KEY,JSON.stringify(db));
 let h='<tr><th>Mês</th><th>Ficheiro</th><th>Líquido lido</th><th></th></tr>';
 for(const r of [...db.receipts].sort((a,b)=>(b.month||'').localeCompare(a.month||''))){
  h+=`<tr><td>${r.month||'—'}</td><td>${r.name}</td><td>${r.parsed?.net!=null?euro(r.parsed.net):'Não encontrado'}</td><td><button onclick="openReceipt('${r.id}')">Abrir</button> <button onclick="removeReceipt('${r.id}')">Apagar</button></td></tr>`;
 }
 $('receiptsTable').innerHTML=h;

 const months=[...new Set(db.receipts.map(r=>r.month).filter(Boolean))].sort().reverse();
 if(!months.length){$('receiptComparison').innerHTML='<p class="hint">Importa um recibo. A comparação será feita automaticamente.</p>';return}
 let c='<table><tr><th>Mês</th><th>Calculado pela aplicação</th><th>Total dos recibos</th><th>Diferença</th><th>Estado</th></tr>';
 for(const month of months){
  const s=receiptMonthSummary(month);
  const d=s.diff==null?'—':`${s.diff>0?'+':''}${euro(s.diff)}`;
  let state='⚠️ Total não encontrado';
  if(s.actual!=null&&!s.complete)state='⚠️ Falta 1 recibo do mês';
  else if(s.actual!=null&&Math.abs(s.diff)<0.05)state='✅ OK';
  else if(s.actual!=null)state='⚠️ Verificar';
  c+=`<tr><td>${month}</td><td>${euro(s.expected)}</td><td>${s.actual==null?'Não encontrado':euro(s.actual)}</td><td><strong>${d}</strong></td><td>${state}</td></tr>`;
 }
 c+='</table>';
 $('receiptComparison').innerHTML=c;
}
function renderSettings(){
 for(const el of $('settingsForm').elements)if(el.name)el.value=db.settings[el.name]??'';
 const nome=String(db.settings.nomeUtilizador||'').trim();
 const h=$('headerUserName');
 if(h)h.textContent=nome?`👤 ${nome}`:'';
}

async function processSharedBackupOnLaunch(){
 const qs=new URLSearchParams(location.search);
 if(qs.get('sharedBackupError')==='1'){
  message('O ficheiro partilhado não é um backup JSON válido. Nenhum dado foi alterado.','err');
  history.replaceState({},'',location.pathname);
  return;
 }
 if(qs.get('sharedBackup')!=='1')return;

 try{
  const sharedKey=new URL('./__shared_backup__.json',location.href).href;
  const c=await caches.open('controlo-horas-shared-backup-v1');
  const response=await c.match(sharedKey);
  if(!response)throw new Error('backup partilhado não encontrado');

  const raw=JSON.parse(await response.text());
  await c.delete(sharedKey);

  const incoming=raw?.data&&raw?.backupVersion?raw.data:raw;
  if(!incoming||typeof incoming!=='object'||!Array.isArray(incoming.sheets))throw new Error('formato inválido');

  const ok=confirm(
   `Backup recebido pela partilha.\n\n`+
   `Folhas: ${incoming.sheets.length}\n`+
   `Recibos: ${Array.isArray(incoming.receipts)?incoming.receipts.length:0}\n`+
   `Despesas: ${Array.isArray(incoming.expenses)?incoming.expenses.length:0}\n\n`+
   `Restaurar agora nesta aplicação?`
  );

  if(ok){
   db={
    ...clone(defaults),
    ...incoming,
    settings:{...defaults.settings,...(incoming.settings||{})},
    sheets:Array.isArray(incoming.sheets)?incoming.sheets:[],
    used:Array.isArray(incoming.used)?incoming.used:[],
    payments:Array.isArray(incoming.payments)?incoming.payments:[],
    expenses:Array.isArray(incoming.expenses)?incoming.expenses:[],
    receipts:Array.isArray(incoming.receipts)?incoming.receipts:[],
    closedMonths:Array.isArray(incoming.closedMonths)?incoming.closedMonths:[]
   };
   save();
   message(`Backup restaurado: ${db.sheets.length} folhas, ${db.receipts.length} recibos e restantes dados.`,'ok');
  }else{
   message('Restauro cancelado. Nenhum dado foi alterado.','ok');
  }
 }catch(err){
  message('Não foi possível ler o backup recebido. Nenhum dado foi alterado.','err');
 }finally{
  history.replaceState({},'',location.pathname);
 }
}

function render(){renderDashboard();renderSheets();renderPayments();renderClients();renderExpenses();renderLeave();renderLocations();renderReceipts();renderSettings()}
document.querySelectorAll('.tab').forEach(b=>b.onclick=()=>{document.querySelectorAll('.tab,.panel').forEach(x=>x.classList.remove('active'));b.classList.add('active');$(b.dataset.tab).classList.add('active')});
$('fileInput').onchange=async e=>{
 if(!window.XLSX){message('Não foi possível carregar o leitor de Excel. Confirma a ligação à Internet.','err');return}
 pendingImports=[];
 for(const f of [...e.target.files]){
  try{
   const buffer=await f.arrayBuffer(),hash=await hashBuffer(buffer);
   const item=parseTimesheet(f,buffer,hash);
   if(isClosed(monthOf(item.dataFinal)))throw new Error('o mês está fechado');
   const existing=db.sheets.find(s=>s.name.toLowerCase()===f.name.toLowerCase()||s.hash===hash);
   item.replaceId=existing?.id||null;
   pendingImports.push(item);
  }catch(err){message(`${f.name}: ${err.message}`,'err')}
 }
 showImportPreview();e.target.value='';
};
function showImportPreview(){if(!pendingImports.length)return;const box=$('importPreview');box.innerHTML=`<h2>Confirmar importação</h2>${pendingImports.map(x=>`<div class="previewItem"><strong>${x.name}</strong>${x.replaceId?' <span class="closedBadge">Substituir existente</span>':''}<br>${x.cliente} · ${x.local}<br>${x.dataInicial} a ${x.dataFinal}<br>TOTAL da folha → 100%: ${fmt(x.normal)} h · 25%: ${fmt(x.h25)} h · 125%: ${fmt(x.h125)} h · 137,5%: ${fmt(x.h1375)} h · 150%: ${fmt(x.h150)} h · 165%: ${fmt(x.h165)} h</div>`).join('')}<div class="toolbar"><button id="confirmImport">Confirmar</button><button id="cancelImport">Cancelar</button></div>`;box.classList.remove('hidden');$('confirmImport').onclick=async()=>{
 for(const item of pendingImports){
  if(item.replaceId){
   const old=db.sheets.find(s=>s.id===item.replaceId);
   if(old?.originalKey)await deleteStoredFile(old.originalKey);
   db.sheets=db.sheets.filter(s=>s.id!==item.replaceId);
  }
  if(item._file)await storeFile(item.originalKey,item._file);
  delete item._file;delete item.replaceId;
  db.sheets.push(item);
 }
 pendingImports=[];box.classList.add('hidden');save();
 message('Folhas importadas/substituídas com os totais oficiais da linha 23.','ok')
};$('cancelImport').onclick=()=>{pendingImports=[];box.classList.add('hidden')}}
$('expenseForm').onsubmit=e=>{e.preventDefault();if($('expenseDate').value<db.settings.inicioDespesas){alert(`Só são aceites despesas a partir de ${db.settings.inicioDespesas}.`);return}if(isClosed(monthOf($('expenseDate').value))){alert('Este mês está fechado.');return}const existing=db.expenses.find(x=>x.date===$('expenseDate').value),obj={id:existing?.id||uid(),date:$('expenseDate').value,food:num($('expenseFood').value),sleep:num($('expenseSleep').value),note:$('expenseNote').value};existing?Object.assign(existing,obj):db.expenses.push(obj);save();$('expenseFood').value=0;$('expenseSleep').value=0;$('expenseNote').value=''}
$('leaveForm').onsubmit=e=>{e.preventDefault();const obj={id:uid(),type:$('leaveType').value,date:$('leaveStart').value,endDate:$('leaveEnd').value,days:num($('leaveDays').value),desc:$('leaveNote').value};db.used.push(obj);save()}
function workingDays(a,b){let n=0;if(!a||!b)return n;for(let d=new Date(a+'T12:00:00'),end=new Date(b+'T12:00:00');d<=end;d.setDate(d.getDate()+1))if(d.getDay()!==0&&d.getDay()!==6)n++;return n}
$('leaveStart').onchange=()=>{if(!$('leaveEnd').value||$('leaveEnd').value<$('leaveStart').value)$('leaveEnd').value=$('leaveStart').value;$('leaveDays').value=workingDays($('leaveStart').value,$('leaveEnd').value)};$('leaveEnd').onchange=()=>$('leaveDays').value=workingDays($('leaveStart').value,$('leaveEnd').value);
$('receiptForm').onsubmit=async e=>{
 e.preventDefault();const f=$('receiptFile').files[0];if(!f)return;
 try{
  const parsed=await extractReceipt(f),month=parsed.month||$('receiptMonth').value;if(!month)throw new Error('Não foi possível identificar o mês.');parsed.month=month;
  const key='receipt-'+uid();await storeFile(key,f);
  const obj={id:uid(),month,name:f.name,type:f.type||f.name.split('.').pop(),fileKey:key,parsed};
  db.receipts.push(obj);save();$('receiptForm').reset();$('receiptProgress').textContent='Recibo lido e comparação atualizada automaticamente.';
 }catch(err){$('receiptProgress').textContent='Erro: '+err.message}
}
function fileToDataURL(f){return new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=rej;r.readAsDataURL(f)})}
window.openReceipt=async id=>{const r=db.receipts.find(x=>x.id===id);if(!r)return;if(r.fileKey){const blob=await getStoredFile(r.fileKey);if(blob)window.open(URL.createObjectURL(blob),'_blank')}else if(r.data)window.open(r.data,'_blank')};window.removeReceipt=async id=>{const r=db.receipts.find(x=>x.id===id);if(r?.fileKey)await deleteStoredFile(r.fileKey);db.receipts=db.receipts.filter(x=>x.id!==id);save()};window.removeSheet=async id=>{if(confirm('Apagar esta folha?')){const s=db.sheets.find(x=>x.id===id);if(s?.originalKey)await deleteStoredFile(s.originalKey);db.sheets=db.sheets.filter(x=>x.id!==id);save()}};window.removeExpense=id=>{db.expenses=db.expenses.filter(x=>x.id!==id);save()};window.removeLeave=id=>{db.used=db.used.filter(x=>x.id!==id);save()};
$('settingsForm').onsubmit=e=>{
 e.preventDefault();
 for(const el of e.target.elements){
  if(!el.name)continue;
  if(el.type==='number')db.settings[el.name]=num(el.value);
  else db.settings[el.name]=el.value;
 }
 db.settings.dependentes=Math.max(0,Math.floor(num(db.settings.dependentes)));
 db.sheets.forEach(s=>s.compConta=s.dataFinal>db.settings.dataCorte?s.compGerada:0);
 save();
 message('Definições guardadas.','ok');
};
async function buildBackupZip(){
 const now=new Date();
 const pad=n=>String(n).padStart(2,'0');
 const stamp=`${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}`;
 const payload={
  app:'Controlo Horas e Compensações',
  backupVersion:2,
  createdAt:now.toISOString(),
  appVersion:'3.9.6',
  data:db
 };
 if(typeof JSZip==='undefined')throw new Error('ZIP indisponível');
 const zip=new JSZip();
 zip.file('dados.json',JSON.stringify(payload,null,2));
 zip.file('LEIA-ME.txt',
`BACKUP COMPLETO — CONTROLO HORAS E COMPENSAÇÕES
Criado em: ${now.toLocaleString('pt-PT')}
Versão da aplicação: 3.9.6

Este ficheiro contém os dados guardados pela aplicação neste dispositivo.

PARA RESTAURAR:
1. Abra a aplicação Controlo Horas e Compensações.
2. Carregue em “Restaurar backup completo”.
3. Escolha este ficheiro ZIP.
4. Confirme a restauração.

Não é necessário extrair o ZIP.
`);
 const blob=await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
 return {blob,filename:`Backup_Controlo_Horas_${stamp}.zip`};
}
$('backupBtn').onclick=async()=>{
 try{
  const {blob,filename}=await buildBackupZip();
  const url=URL.createObjectURL(blob),a=document.createElement('a');
  a.href=url;a.download=filename;
  document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
  message(`Backup completo criado: ${db.sheets.length} folhas, ${db.receipts.length} recibos e restantes dados.`,'ok');
 }catch(err){
  message('Não foi possível criar o ZIP. Verifica a ligação à Internet e tenta novamente.','err');
 }
};
$('shareBackupBtn').onclick=async()=>{
 try{
  const now=new Date();
  const pad=n=>String(n).padStart(2,'0');
  const stamp=`${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}`;
  const payload={
   app:'Controlo Horas e Compensações',
   backupVersion:3,
   createdAt:now.toISOString(),
   appVersion:'3.9.8',
   data:db
  };
  const text=JSON.stringify(payload,null,2);
  const filename=`Backup_Controlo_Horas_${stamp}.json`;
  const blob=new Blob([text],{type:'application/json;charset=utf-8'});
  const file=new File([blob],filename,{type:'application/json'});

  // Mostra ao utilizador o que está efetivamente a sair desta instalação.
  const resumo=`${db.sheets.length} folhas · ${db.receipts.length} recibos · ${db.expenses.length} despesas`;
  const ok=confirm(`Backup preparado nesta aplicação:\n\n${resumo}\n\nContinuar para partilhar?`);
  if(!ok)return;

  if(window.isSecureContext && typeof navigator.share==='function'){
   try{
    await navigator.share({
     title:'Backup Controlo Horas',
     text:`Backup completo — ${resumo}`,
     files:[file]
    });
    message(`Backup partilhado: ${resumo}.`,'ok');
    return;
   }catch(err){
    if(err?.name==='AbortError')return;
   }
  }

  // Fallback para browsers que não partilham ficheiros.
  const url=URL.createObjectURL(blob),a=document.createElement('a');
  a.href=url;a.download=filename;
  document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
  message(`Backup guardado: ${resumo}.`,'ok');
 }catch(err){
  message('Não foi possível criar/partilhar o backup. Nenhum dado foi alterado.','err');
 }
};
$('restoreBtn').onclick=()=>$('restoreInput').click();
$('restoreInput').onchange=async()=>{
 const f=$('restoreInput').files?.[0];
 if(!f)return;
 try{
  let raw;
  if(f.name.toLowerCase().endsWith('.zip')){
   if(typeof JSZip==='undefined')throw new Error('ZIP indisponível');
   const zip=await JSZip.loadAsync(f);
   const entry=zip.file('dados.json');
   if(!entry)throw new Error('dados.json ausente');
   raw=JSON.parse(await entry.async('text'));
  }else{
   // Mantém compatibilidade com os backups JSON antigos.
   raw=JSON.parse(await f.text());
  }
  const incoming=raw?.data&&raw?.backupVersion?raw.data:raw;
  if(!incoming||typeof incoming!=='object'||!Array.isArray(incoming.sheets))throw new Error('Formato inválido');
  const ok=confirm(`Restaurar este backup completo?\n\nFolhas: ${incoming.sheets.length}\nRecibos: ${Array.isArray(incoming.receipts)?incoming.receipts.length:0}\nDespesas: ${Array.isArray(incoming.expenses)?incoming.expenses.length:0}\n\nOs dados atuais deste dispositivo serão substituídos pelos dados do backup.`);
  if(!ok){$('restoreInput').value='';return}
  db={
   ...clone(defaults),
   ...incoming,
   settings:{...defaults.settings,...(incoming.settings||{})},
   sheets:Array.isArray(incoming.sheets)?incoming.sheets:[],
   used:Array.isArray(incoming.used)?incoming.used:[],
   payments:Array.isArray(incoming.payments)?incoming.payments:[],
   expenses:Array.isArray(incoming.expenses)?incoming.expenses:[],
   receipts:Array.isArray(incoming.receipts)?incoming.receipts:[],
   closedMonths:Array.isArray(incoming.closedMonths)?incoming.closedMonths:[]
  };
  save();
  message(`Backup restaurado: ${db.sheets.length} folhas, ${db.receipts.length} recibos e restantes dados.`,'ok');
 }catch(err){
  message('Backup inválido ou danificado. Nenhum dado foi alterado.','err');
 }finally{$('restoreInput').value=''}
};
$('closeMonthBtn').onclick=()=>{const m=$('dashMonth').value;if(!m)return;if(!db.closedMonths.includes(m))db.closedMonths.push(m);save()};$('reopenMonthBtn').onclick=()=>{db.closedMonths=db.closedMonths.filter(x=>x!==$('dashMonth').value);save()};$('printMonthBtn').onclick=()=>window.print();
$('clearBtn').onclick=()=>{if(confirm('Apagar todos os dados?')){db=clone(defaults);save()}};
['dashMonth','clientMonth','expenseMonth'].forEach(id=>$(id).onchange=render);$('payYear').onchange=renderPayments;$('sheetSearch').oninput=renderSheets;$('clientFilter').oninput=renderClients;$('downloadClientsBtn').onclick=downloadClientsPdf;$('locationSearch').oninput=renderLocations;$('locationSource').onchange=renderLocations;
$('globalSearch').oninput=()=>{const q=$('globalSearch').value;if(q.length<2)return;const loc=locations.find(x=>norm(x.name).includes(norm(q))),sheet=db.sheets.find(x=>norm(`${x.cliente} ${x.local} ${x.name}`).includes(norm(q)));if(loc){document.querySelector('[data-tab="locais"]').click();$('locationSearch').value=q;renderLocations()}else if(sheet){document.querySelector('[data-tab="folhas"]').click();$('sheetSearch').value=q;renderSheets()}};
let installPrompt=null;window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();installPrompt=e;$('installBtn').hidden=false});$('installBtn').onclick=async()=>{if(installPrompt){installPrompt.prompt();await installPrompt.userChoice;installPrompt=null;$('installBtn').hidden=true}else alert('No iPhone: Partilhar → Adicionar ao ecrã principal.')};
function connection(){if(navigator.onLine){$('connectionStatus').textContent='';$('connectionStatus').className=''}else{$('connectionStatus').textContent='📴 Sem internet: os dados continuam guardados neste dispositivo.';$('connectionStatus').className='offline'}}window.addEventListener('online',connection);window.addEventListener('offline',connection);
fetch('locations.json').then(r=>r.json()).then(x=>{locations=x;const src=[...new Set(x.map(v=>v.source))].sort();$('locationSource').innerHTML='<option value="">Todas as entidades</option>'+src.map(s=>`<option>${s}</option>`).join('');renderLocations()}).catch(()=>{});
for(const id of ['dashMonth','expenseMonth','receiptMonth'])$(id).value=currentMonth();$('clientMonth').value=latestSheetMonth();$('expenseDate').value=today();$('leaveStart').value=today();$('leaveEnd').value=today();connection();render();processSharedBackupOnLaunch();
