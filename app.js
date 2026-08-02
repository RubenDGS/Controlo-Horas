const KEY='controloHorasV6';
const $=id=>document.getElementById(id);
const defaults={settings:{salarioBase:1500,taxaSS:11,valorHora:8.65,ajudaDia:90,alojDia:65,refeicaoDia:10.46,inicioDespesas:'2027-01-01',saldoInicial:41,dataCorte:'2026-07-25',diasFeriasAnuais:22,primeiroAnoFerias:2027,m25:.25,m125:1.25,m1375:1.375,m150:1.5,m165:1.65},sheets:[],used:[],payments:[],expenses:[],receipts:[],closedMonths:[]};
let db=load(),locations=[],pendingImports=[];
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
function findSheet(wb){return wb.Sheets['MD-RH-04 - Registo de Horas']||wb.Sheets[wb.SheetNames.find(n=>norm(n).includes('registo de horas'))]||wb.Sheets[wb.SheetNames[0]]}
function parseTimesheet(file,buffer,hash){
 const wb=XLSX.read(buffer,{type:'array',cellDates:true}),ws=findSheet(wb),entries=[];
 for(let r=15;r<=22;r++){const date=excelDate(cell(ws,`A${r}`));if(!date)continue;entries.push({date,dayType:String(cell(ws,`B${r}`)||''),normal:hourValue(cell(ws,`I${r}`)),h25:hourValue(cell(ws,`J${r}`)),h125:hourValue(cell(ws,`K${r}`)),h1375:hourValue(cell(ws,`L${r}`)),h150:hourValue(cell(ws,`M${r}`)),h165:hourValue(cell(ws,`N${r}`))})}
 if(!entries.length)throw new Error('Não foram encontradas datas nas linhas 15 a 22.');
 const dataInicial=entries.map(x=>x.date).sort()[0],dataFinal=entries.map(x=>x.date).sort().at(-1);
 const cliente=String(cell(ws,'C10')||'').trim(),local=String(cell(ws,'I10')||'').trim(),processo=String(cell(ws,'M10')||'').trim();
 const diasAjuda=num(cell(ws,'J26')),diasAloj=num(cell(ws,'J27'));
 const totals=entries.reduce((a,x)=>{for(const k of ['normal','h25','h125','h1375','h150','h165'])a[k]+=x[k];return a},{normal:0,h25:0,h125:0,h1375:0,h150:0,h165:0});
 const compGerada=entries.reduce((a,x)=>{const t=norm(x.dayType);return a+(t.includes('domingo')?1:(t.includes('sabado')||t.includes('feriado')?.5:0))},0);
 return {id:uid(),name:file.name,hash,cliente,local,processo,dataInicial,dataFinal,entries,...totals,diasAjuda,diasAloj,ajudaUnit:db.settings.ajudaDia,alojUnit:db.settings.alojDia,compGerada,compConta:dataFinal>db.settings.dataCorte?compGerada:0,hoursVersion:2,imported:new Date().toISOString()};
}
function overtimeValue(g){const s=db.settings;return s.valorHora*(num(g.h25)*s.m25+num(g.h125)*s.m125+num(g.h1375)*s.m1375+num(g.h150)*s.m150+num(g.h165)*s.m165)}
function irs2026(R){R=Math.max(0,num(R));let t=0,a=0;if(R<=920){}else if(R<=1042){t=.125;a=.125*2.6*(1273.85-R)}else if(R<=1108){t=.157;a=.157*1.35*(1554.83-R)}else if(R<=1154){t=.157;a=94.71}else if(R<=1212){t=.212;a=158.18}else if(R<=1819){t=.241;a=193.33}else if(R<=2119){t=.311;a=320.66}else if(R<=2499){t=.349;a=401.19}else if(R<=3305){t=.3836;a=487.66}else if(R<=5547){t=.3969;a=531.62}else if(R<=20221){t=.4495;a=823.4}else{t=.4717;a=1272.31}const v=Math.max(0,R*t-a);return{value:v,effective:R?v/R:0}}
function sheetsForMonth(m){return db.sheets.filter(s=>monthOf(s.dataFinal)===m||s.entries?.some(e=>monthOf(e.date)===m))}
function monthGroup(m){const g={normal:0,h25:0,h125:0,h1375:0,h150:0,h165:0,travelDays:0,lodgingDays:0,travel:0,lodging:0,comp:0};for(const s of sheetsForMonth(m)){for(const e of s.entries||[]){if(monthOf(e.date)!==m)continue;for(const k of ['normal','h25','h125','h1375','h150','h165'])g[k]+=num(e[k])}if(monthOf(s.dataFinal)===m){g.travelDays+=num(s.diasAjuda);g.lodgingDays+=num(s.diasAloj);g.travel+=num(s.diasAjuda)*num(s.ajudaUnit||db.settings.ajudaDia);g.lodging+=num(s.diasAloj)*num(s.alojUnit||db.settings.alojDia);g.comp+=num(s.compConta)}}return g}
function salaryMonth(m){const g=monthGroup(m),grossHours=overtimeValue(g),gross=db.settings.salarioBase+grossHours,ss=gross*db.settings.taxaSS/100,irsBase=irs2026(db.settings.salarioBase),irsHours=grossHours*(irsBase.effective/2),net=gross-ss-irsBase.value-irsHours;const correction=db.payments.find(x=>x.month===m);return{...g,grossHours,gross,ss,irs:irsBase.value+irsHours,net:correction?.net??net}}
function easterDate(y){const a=y%19,b=Math.floor(y/100),c=y%100,d=Math.floor(b/4),e=b%4,f=Math.floor((b+8)/25),g=Math.floor((b-f+1)/3),h=(19*a+b-d-g+15)%30,i=Math.floor(c/4),k=c%4,l=(32+2*e+2*i-h-k)%7,m=Math.floor((a+11*h+22*l)/451),month=Math.floor((h+l-7*m+114)/31),day=((h+l-7*m+114)%31)+1;return new Date(y,month-1,day)}
function isoLocal(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`}
function holidays(y){const fixed=['01-01','04-25','05-01','06-10','08-15','10-05','11-01','12-01','12-08','12-25'].map(x=>`${y}-${x}`);const e=easterDate(y),good=new Date(e);good.setDate(e.getDate()-2);const corpus=new Date(e);corpus.setDate(e.getDate()+60);return new Set([...fixed,isoLocal(good),isoLocal(e),isoLocal(corpus)])}
function coveredByLeave(date){return db.used.some(x=>date>=x.date&&date<=(x.endDate||x.date))}
function sheetDates(){const set=new Set();db.sheets.forEach(s=>(s.entries||[]).forEach(e=>set.add(e.date)));return set}
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
function renderDashboard(){const m=$('dashMonth').value||currentMonth(),p=salaryMonth(m),e=expenseSummary(m),extra=p.h25+p.h125+p.h1375+p.h150+p.h165;const allExtra=db.sheets.reduce((a,s)=>a+num(s.h25)+num(s.h125)+num(s.h1375)+num(s.h150)+num(s.h165),0);$('dashNet').textContent=euro(p.net);$('dashExtra').textContent=fmt(extra)+' h';$('dashExtraAll').textContent=fmt(allExtra)+' h';const mealCount=mealDays(m);$('dashMeal').textContent=`${euro(mealCount*db.settings.refeicaoDia)} (${mealCount} dias)`;$('dashTravel').textContent=euro(p.travel);$('dashLodging').textContent=euro(p.lodging);$('dashBalance').textContent=fmt(balance())+' dias';$('dashExpenseBalance').textContent=euro(e.remain);const w=monthStatus(m),closed=isClosed(m);$('monthStatus').className='statusBox '+(w.length?'status-warn':'status-ok');$('monthStatus').innerHTML=`<strong>${closed?'🔒 Mês fechado':'🟢 Mês aberto'}</strong>${w.length?' · Falta: '+w.join(', '):' · Registos completos'}`;renderMonthly();renderAnnual()}
function renderMonthly(){let h='<tr><th>Mês</th><th>Líquido recibo</th><th>Horas normais</th><th>Horas extra</th><th>Subs. refeição</th><th>Ajudas</th><th>Alojamento</th><th>Estado</th></tr>';for(const m of allMonths().reverse()){const p=salaryMonth(m),extra=p.h25+p.h125+p.h1375+p.h150+p.h165,w=monthStatus(m);h+=`<tr><td>${m}</td><td>${euro(p.net)}</td><td>${fmt(p.normal)} h</td><td>${fmt(extra)} h</td><td>${euro(mealDays(m)*db.settings.refeicaoDia)}</td><td>${euro(p.travel)}</td><td>${euro(p.lodging)}</td><td>${isClosed(m)?'<span class="closedBadge">Fechado</span>':w.length?'<span class="closedBadge">'+w.join(', ')+'</span>':'<span class="openBadge">Completo</span>'}</td></tr>`}$('monthlyTable').innerHTML=h}
function renderAnnual(){const years={};for(const m of allMonths()){const y=m.slice(0,4),p=salaryMonth(m);years[y]??={net:0,normal:0,extra:0,meal:0,travel:0,lodging:0};years[y].net+=p.net;years[y].normal+=p.normal;years[y].extra+=p.h25+p.h125+p.h1375+p.h150+p.h165;years[y].meal+=mealDays(m)*db.settings.refeicaoDia;years[y].travel+=p.travel;years[y].lodging+=p.lodging}let h='<tr><th>Ano</th><th>Líquido recibos</th><th>Horas normais</th><th>Horas extra</th><th>Subs. refeição</th><th>Ajudas</th><th>Alojamento</th></tr>';for(const y of Object.keys(years).sort().reverse()){const x=years[y];h+=`<tr><td>${y}</td><td>${euro(x.net)}</td><td>${fmt(x.normal)} h</td><td>${fmt(x.extra)} h</td><td>${euro(x.meal)}</td><td>${euro(x.travel)}</td><td>${euro(x.lodging)}</td></tr>`}$('annualTable').innerHTML=h}
function renderSheets(){const q=norm($('sheetSearch').value);let h='<tr><th>Ficheiro</th><th>Período</th><th>Cliente</th><th>Local</th><th>Normais</th><th>Extra</th><th>Ajudas</th><th>Aloj.</th><th></th></tr>';for(const s of [...db.sheets].sort((a,b)=>b.dataFinal.localeCompare(a.dataFinal))){if(q&&!norm(`${s.name} ${s.cliente} ${s.local} ${s.processo}`).includes(q))continue;const ex=num(s.h25)+num(s.h125)+num(s.h1375)+num(s.h150)+num(s.h165);h+=`<tr><td>${s.name}</td><td>${s.dataInicial}<br>${s.dataFinal}</td><td>${s.cliente}</td><td>${s.local}</td><td>${fmt(s.normal)} h</td><td>${fmt(ex)} h</td><td>${fmt(s.diasAjuda)} dias</td><td>${fmt(s.diasAloj)} dias</td><td><button onclick="removeSheet('${s.id}')">Apagar</button></td></tr>`}$('sheetsTable').innerHTML=h}
function renderPayments(){const years=[...new Set(allMonths().map(m=>m.slice(0,4)))].sort().reverse(),sel=$('payYear').value||years[0]||String(new Date().getFullYear());$('payYear').innerHTML=years.map(y=>`<option ${y===sel?'selected':''}>${y}</option>`).join('');let sums={net:0,travel:0,lodging:0,meal:0},h='<tr><th>Mês</th><th>Bruto</th><th>SS</th><th>IRS</th><th>Líquido recibo</th><th>Ajudas</th><th>Alojamento</th><th>Subs. refeição</th><th>Recibo</th></tr>';for(const m of allMonths().filter(x=>x.startsWith(sel)).reverse()){const p=salaryMonth(m),receipt=db.receipts.find(r=>r.month===m);sums.net+=p.net;sums.travel+=p.travel;sums.lodging+=p.lodging;sums.meal+=mealDays(m)*db.settings.refeicaoDia;h+=`<tr><td>${m}</td><td>${euro(p.gross)}</td><td>- ${euro(p.ss)}</td><td>- ${euro(p.irs)}</td><td><strong>${euro(p.net)}</strong></td><td>${euro(p.travel)}</td><td>${euro(p.lodging)}</td><td>${euro(mealDays(m)*db.settings.refeicaoDia)}</td><td>${receipt?'✅':'⚠️ falta'}</td></tr>`}$('yearNet').textContent=euro(sums.net);$('yearTravel').textContent=euro(sums.travel);$('yearLodging').textContent=euro(sums.lodging);$('yearMeal').textContent=euro(sums.meal);$('paymentsTable').innerHTML=h}
function latestSheetMonth(){
 const months=db.sheets.flatMap(s=>Array.isArray(s.entries)&&s.entries.length?s.entries.map(e=>monthOf(e.date)):[monthOf(s.dataFinal)]).filter(Boolean).sort();
 return months.at(-1)||currentMonth();
}
function renderClients(){
 const m=$('clientMonth').value||latestSheetMonth(),q=norm($('clientFilter').value),groups={};
 if(!$('clientMonth').value)$('clientMonth').value=m;
 for(const s of db.sheets){
  let matched=false;
  for(const e of s.entries||[]){
   if(monthOf(e.date)!==m)continue;
   matched=true;
   const key=`${s.cliente||'Sem cliente'}|||${s.local||'Sem local'}`;
   groups[key]??={client:s.cliente||'Sem cliente',local:s.local||'Sem local',normal:0,h25:0,h125:0,h1375:0,h150:0,h165:0};
   for(const k of ['normal','h25','h125','h1375','h150','h165'])groups[key][k]+=num(e[k]);
  }
  if(!matched&&monthOf(s.dataFinal)===m){
   const key=`${s.cliente||'Sem cliente'}|||${s.local||'Sem local'}`;
   groups[key]??={client:s.cliente||'Sem cliente',local:s.local||'Sem local',normal:0,h25:0,h125:0,h1375:0,h150:0,h165:0};
   for(const k of ['normal','h25','h125','h1375','h150','h165'])groups[key][k]+=num(s[k]);
  }
 }
 let h='<tr><th>Cliente</th><th>Local</th><th>Normais</th><th>Total extra</th><th>25%</th><th>125%</th><th>137,5%</th><th>150%</th><th>165%</th></tr>';
 const rows=Object.values(groups).filter(x=>!q||norm(x.client+' '+x.local).includes(q)).sort((a,b)=>a.client.localeCompare(b.client));
 for(const x of rows){
  const ex=x.h25+x.h125+x.h1375+x.h150+x.h165;
  h+=`<tr><td>${x.client}</td><td>${x.local}</td><td>${fmt(x.normal)}</td><td><strong>${fmt(ex)}</strong></td><td>${fmt(x.h25)}</td><td>${fmt(x.h125)}</td><td>${fmt(x.h1375)}</td><td>${fmt(x.h150)}</td><td>${fmt(x.h165)}</td></tr>`;
 }
 if(!rows.length)h+=`<tr><td colspan="9">Não existem folhas para ${m}. O último mês com folhas é ${latestSheetMonth()}.</td></tr>`;
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
function renderReceipts(){let h='<tr><th>Mês</th><th>Ficheiro</th><th>Formato</th><th>Líquido registado</th><th>Diferença estimativa</th><th></th></tr>';for(const r of [...db.receipts].sort((a,b)=>b.month.localeCompare(a.month))){const diff=r.net?num(r.net)-salaryMonth(r.month).net:null;h+=`<tr><td>${r.month}</td><td>${r.name}</td><td>${r.type}</td><td>${r.net?euro(r.net):'—'}</td><td>${diff===null?'—':euro(diff)}</td><td><button onclick="openReceipt('${r.id}')">Abrir</button> <button onclick="removeReceipt('${r.id}')">Apagar</button></td></tr>`}$('receiptsTable').innerHTML=h}
function renderSettings(){for(const el of $('settingsForm').elements)if(el.name)el.value=db.settings[el.name]??''}
function render(){renderDashboard();renderSheets();renderPayments();renderClients();renderExpenses();renderLeave();renderLocations();renderReceipts();renderSettings()}
document.querySelectorAll('.tab').forEach(b=>b.onclick=()=>{document.querySelectorAll('.tab,.panel').forEach(x=>x.classList.remove('active'));b.classList.add('active');$(b.dataset.tab).classList.add('active')});
$('fileInput').onchange=async e=>{if(!window.XLSX){message('Não foi possível carregar o leitor de Excel. Confirma a ligação à Internet.','err');return}pendingImports=[];for(const f of [...e.target.files]){try{if(db.sheets.some(s=>s.name.toLowerCase()===f.name.toLowerCase()))throw new Error('já existe um ficheiro com este nome');const buffer=await f.arrayBuffer(),hash=await hashBuffer(buffer);if(db.sheets.some(s=>s.hash===hash))throw new Error('conteúdo já importado');const item=parseTimesheet(f,buffer,hash);if(isClosed(monthOf(item.dataFinal)))throw new Error('o mês está fechado');pendingImports.push(item)}catch(err){message(`${f.name}: ${err.message}`,'err')}}showImportPreview();e.target.value=''};
function showImportPreview(){if(!pendingImports.length)return;const box=$('importPreview');box.innerHTML=`<h2>Confirmar importação</h2>${pendingImports.map(x=>`<div class="previewItem"><strong>${x.name}</strong><br>${x.cliente} · ${x.local}<br>${x.dataInicial} a ${x.dataFinal}<br>Normais: ${fmt(x.normal)} h · Extra: ${fmt(x.h25+x.h125+x.h1375+x.h150+x.h165)} h · Ajuda: ${x.diasAjuda} dias · Alojamento: ${x.diasAloj} dias</div>`).join('')}<div class="toolbar"><button id="confirmImport">Confirmar</button><button id="cancelImport">Cancelar</button></div>`;box.classList.remove('hidden');$('confirmImport').onclick=()=>{db.sheets.push(...pendingImports);pendingImports=[];box.classList.add('hidden');save();message('Folhas importadas.','ok')};$('cancelImport').onclick=()=>{pendingImports=[];box.classList.add('hidden')}}
$('expenseForm').onsubmit=e=>{e.preventDefault();if($('expenseDate').value<db.settings.inicioDespesas){alert(`Só são aceites despesas a partir de ${db.settings.inicioDespesas}.`);return}if(isClosed(monthOf($('expenseDate').value))){alert('Este mês está fechado.');return}const existing=db.expenses.find(x=>x.date===$('expenseDate').value),obj={id:existing?.id||uid(),date:$('expenseDate').value,food:num($('expenseFood').value),sleep:num($('expenseSleep').value),note:$('expenseNote').value};existing?Object.assign(existing,obj):db.expenses.push(obj);save();$('expenseFood').value=0;$('expenseSleep').value=0;$('expenseNote').value=''}
$('leaveForm').onsubmit=e=>{e.preventDefault();const obj={id:uid(),type:$('leaveType').value,date:$('leaveStart').value,endDate:$('leaveEnd').value,days:num($('leaveDays').value),desc:$('leaveNote').value};db.used.push(obj);save()}
function workingDays(a,b){let n=0;if(!a||!b)return n;for(let d=new Date(a+'T12:00:00'),end=new Date(b+'T12:00:00');d<=end;d.setDate(d.getDate()+1))if(d.getDay()!==0&&d.getDay()!==6)n++;return n}
$('leaveStart').onchange=()=>{if(!$('leaveEnd').value||$('leaveEnd').value<$('leaveStart').value)$('leaveEnd').value=$('leaveStart').value;$('leaveDays').value=workingDays($('leaveStart').value,$('leaveEnd').value)};$('leaveEnd').onchange=()=>$('leaveDays').value=workingDays($('leaveStart').value,$('leaveEnd').value);
$('receiptForm').onsubmit=async e=>{e.preventDefault();const f=$('receiptFile').files[0];if(!f)return;if(f.size>4*1024*1024&&!confirm('O ficheiro é grande e pode ocupar bastante espaço. Continuar?'))return;const data=await fileToDataURL(f),obj={id:uid(),month:$('receiptMonth').value,name:f.name,type:f.type||f.name.split('.').pop(),data,net:num($('receiptNet').value)||null,extra:num($('receiptExtra').value)||null};db.receipts=db.receipts.filter(x=>x.month!==obj.month);db.receipts.push(obj);save();$('receiptForm').reset()}
function fileToDataURL(f){return new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=rej;r.readAsDataURL(f)})}
window.openReceipt=id=>{const r=db.receipts.find(x=>x.id===id);if(r)window.open(r.data,'_blank')};window.removeReceipt=id=>{db.receipts=db.receipts.filter(x=>x.id!==id);save()};window.removeSheet=id=>{if(confirm('Apagar esta folha?')){db.sheets=db.sheets.filter(x=>x.id!==id);save()}};window.removeExpense=id=>{db.expenses=db.expenses.filter(x=>x.id!==id);save()};window.removeLeave=id=>{db.used=db.used.filter(x=>x.id!==id);save()};
$('settingsForm').onsubmit=e=>{e.preventDefault();for(const [k,v] of new FormData(e.target))db.settings[k]=k.includes('data')||k.includes('inicio')?v:num(v);db.sheets.forEach(s=>s.compConta=s.dataFinal>db.settings.dataCorte?s.compGerada:0);save();message('Definições guardadas.','ok')};
$('backupBtn').onclick=()=>{const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([JSON.stringify(db)],{type:'application/json'}));a.download=`controlo_horas_${today()}.json`;a.click();URL.revokeObjectURL(a.href)};
$('restoreBtn').onclick=()=>$('restoreInput').click();$('restoreInput').onchange=async()=>{try{db={...clone(defaults),...JSON.parse(await $('restoreInput').files[0].text())};save();message('Backup restaurado.','ok')}catch{message('Backup inválido.','err')}};
$('closeMonthBtn').onclick=()=>{const m=$('dashMonth').value;if(!m)return;if(!db.closedMonths.includes(m))db.closedMonths.push(m);save()};$('reopenMonthBtn').onclick=()=>{db.closedMonths=db.closedMonths.filter(x=>x!==$('dashMonth').value);save()};$('printMonthBtn').onclick=()=>window.print();
$('clearBtn').onclick=()=>{if(confirm('Apagar todos os dados?')){db=clone(defaults);save()}};
['dashMonth','clientMonth','expenseMonth'].forEach(id=>$(id).onchange=render);$('payYear').onchange=renderPayments;$('sheetSearch').oninput=renderSheets;$('clientFilter').oninput=renderClients;$('locationSearch').oninput=renderLocations;$('locationSource').onchange=renderLocations;
$('globalSearch').oninput=()=>{const q=$('globalSearch').value;if(q.length<2)return;const loc=locations.find(x=>norm(x.name).includes(norm(q))),sheet=db.sheets.find(x=>norm(`${x.cliente} ${x.local} ${x.name}`).includes(norm(q)));if(loc){document.querySelector('[data-tab="locais"]').click();$('locationSearch').value=q;renderLocations()}else if(sheet){document.querySelector('[data-tab="folhas"]').click();$('sheetSearch').value=q;renderSheets()}};
let installPrompt=null;window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();installPrompt=e;$('installBtn').hidden=false});$('installBtn').onclick=async()=>{if(installPrompt){installPrompt.prompt();await installPrompt.userChoice;installPrompt=null;$('installBtn').hidden=true}else alert('No iPhone: Partilhar → Adicionar ao ecrã principal.')};
function connection(){if(navigator.onLine){$('connectionStatus').textContent='';$('connectionStatus').className=''}else{$('connectionStatus').textContent='📴 Sem internet: os dados continuam guardados neste dispositivo.';$('connectionStatus').className='offline'}}window.addEventListener('online',connection);window.addEventListener('offline',connection);
fetch('locations.json').then(r=>r.json()).then(x=>{locations=x;const src=[...new Set(x.map(v=>v.source))].sort();$('locationSource').innerHTML='<option value="">Todas as entidades</option>'+src.map(s=>`<option>${s}</option>`).join('');renderLocations()}).catch(()=>{});
for(const id of ['dashMonth','expenseMonth','receiptMonth'])$(id).value=currentMonth();$('clientMonth').value=latestSheetMonth();$('expenseDate').value=today();$('leaveStart').value=today();$('leaveEnd').value=today();connection();render();
