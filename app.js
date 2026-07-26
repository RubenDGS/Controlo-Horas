const DBKEY='controloHorasV6';
const fileInput=document.getElementById('fileInput');
const backupBtn=document.getElementById('backupBtn');
const restoreBtn=document.getElementById('restoreBtn');
const restoreInput=document.getElementById('restoreInput');
const messages=document.getElementById('messages');
const monthlyTable=document.getElementById('monthlyTable');
const annualTable=document.getElementById('annualTable');
const sheetsTable=document.getElementById('sheetsTable');
const usedTable=document.getElementById('usedTable');
const payTable=document.getElementById('payTable');
const payYearFilter=document.getElementById('payYearFilter');
const yearNetCard=document.getElementById('yearNetCard');
const yearTravelCard=document.getElementById('yearTravelCard');
const yearLodgingCard=document.getElementById('yearLodgingCard');
const yearGlobalCard=document.getElementById('yearGlobalCard');
const search=document.getElementById('search');
const usedForm=document.getElementById('usedForm');
const usedDate=document.getElementById('usedDate');
const usedDays=document.getElementById('usedDays');
const usedEndDate=document.getElementById('usedEndDate');
const usedType=document.getElementById('usedType');
const saldoFeriasCard=document.getElementById('saldoFeriasCard');
const feriasAnuaisCard=document.getElementById('feriasAnuaisCard');
const diasUsadosCard=document.getElementById('diasUsadosCard');
const grossAutoCard=document.getElementById('grossAutoCard');
const ssAutoCard=document.getElementById('ssAutoCard');
const irsAutoCard=document.getElementById('irsAutoCard');
const monthAutoCard=document.getElementById('monthAutoCard');
const ajudaAutoCard=document.getElementById('ajudaAutoCard');
const alojAutoCard=document.getElementById('alojAutoCard');
const expensesAutoCard=document.getElementById('expensesAutoCard');
const usedDesc=document.getElementById('usedDesc');
const payForm=document.getElementById('payForm');
const payMonth=document.getElementById('payMonth');
const paySalary=document.getElementById('paySalary');
const payHours=document.getElementById('payHours');
const settingsForm=document.getElementById('settingsForm');
const clearBtn=document.getElementById('clearBtn');
const saldoCard=document.getElementById('saldoCard');
const horasCard=document.getElementById('horasCard');
const ajudasCard=document.getElementById('ajudasCard');
const alojCard=document.getElementById('alojCard');
const dashMonth=document.getElementById('dashMonth');
const dashNetCard=document.getElementById('dashNetCard');
const dashHoursMonthCard=document.getElementById('dashHoursMonthCard');
const dashTravelCard=document.getElementById('dashTravelCard');
const dashLodgingCard=document.getElementById('dashLodgingCard');

const defaults={
 settings:{salarioBase:1500,taxaSS:11,valorHora:8.65,ajudaDia:90,alojDia:65,saldoInicial:41,dataCorte:'2026-07-25',diasFeriasAnuais:22,primeiroAnoFerias:2027,m25:.25,m125:1.25,m1375:1.375,m150:1.5,m165:1.65},
 sheets:[], used:[], payments:[]
};
let db=load();

function fresh(){return JSON.parse(JSON.stringify(defaults))}
function load(){
  try{
    const old=JSON.parse(localStorage.getItem(DBKEY)||'{}');
    const x={...fresh(),...old};
    x.settings={...fresh().settings,...(old.settings||{})};
    x.sheets=Array.isArray(old.sheets)?old.sheets:[];
    x.used=Array.isArray(old.used)?old.used:[];
    x.payments=Array.isArray(old.payments)?old.payments:[];
    return x;
  }catch{return fresh()}
}
function save(){localStorage.setItem(DBKEY,JSON.stringify(db));render()}
function euro(v){return Number(v||0).toLocaleString('pt-PT',{style:'currency',currency:'EUR'})}
function num(v){
  if(typeof v==='number')return Number.isFinite(v)?v:0;
  if(typeof v==='string'){
    const clean=v.trim().replace(/\s/g,'').replace(/€/g,'').replace(/\.(?=\d{3}(?:\D|$))/g,'').replace(',','.');
    const x=Number(clean); return Number.isFinite(x)?x:0;
  }
  return 0;
}
function fmt(v){return Number(v||0).toLocaleString('pt-PT',{maximumFractionDigits:2})}
function msg(text,type='ok'){const d=document.createElement('div');d.className='msg '+type;d.textContent=text;messages.prepend(d)}
async function sha256(buf){
  if(globalThis.crypto && crypto.subtle){
    const h=await crypto.subtle.digest('SHA-256',buf);
    return [...new Uint8Array(h)].map(b=>b.toString(16).padStart(2,'0')).join('');
  }
  const a=new Uint8Array(buf); let h=2166136261;
  for(const b of a){h^=b;h=Math.imul(h,16777619)}
  return 'local-'+(h>>>0).toString(16)+'-'+a.length;
}
function excelDate(v){
  if(v instanceof Date && !isNaN(v))return v.toISOString().slice(0,10);
  if(typeof v==='number'){const d=XLSX.SSF.parse_date_code(v);if(d)return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`}
  if(typeof v==='string'){
    const m=v.match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
    if(m)return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  }
  return '';
}
function cell(ws,a){return ws[a]?.v??''}
function cleanText(v){return String(v??'').trim().replace(/\s+/g,' ')}
function norm(v){return cleanText(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase()}
function findSheet(wb){
  return wb.Sheets['MD-RH-04 - Registo de Horas']
      || wb.Sheets[wb.SheetNames.find(n=>norm(n).includes('registo de horas'))]
      || wb.Sheets[wb.SheetNames[0]];
}
function sheetMatrix(ws){
  const range=XLSX.utils.decode_range(ws['!ref']||'A1:A1');
  const rows=[];
  for(let r=range.s.r;r<=range.e.r;r++){
    const row=[];
    for(let c=range.s.c;c<=range.e.c;c++){
      const a=XLSX.utils.encode_cell({r,c});
      row.push(cell(ws,a));
    }
    rows.push(row);
  }
  return {rows,startRow:range.s.r,startCol:range.s.c};
}
function findLabelValue(ws, needles, maxRight=5){
  const range=XLSX.utils.decode_range(ws['!ref']||'A1:A1');
  for(let r=range.s.r;r<=range.e.r;r++){
    for(let c=range.s.c;c<=range.e.c;c++){
      const a=XLSX.utils.encode_cell({r,c});
      const text=norm(cell(ws,a));
      if(!text || !needles.some(n=>text.includes(norm(n))))continue;
      for(let k=1;k<=maxRight;k++){
        const b=XLSX.utils.encode_cell({r:r,c:c+k});
        const v=cell(ws,b);
        if(v!=='' && v!==null && v!==undefined && cleanText(v)!=='')return {value:v,label:cell(ws,a),address:b};
      }
    }
  }
  return {value:'',label:'',address:''};
}
function findDaysForLabel(ws, needles){
  const range=XLSX.utils.decode_range(ws['!ref']||'A1:A1');
  for(let r=range.s.r;r<=range.e.r;r++){
    for(let c=range.s.c;c<=range.e.c;c++){
      const a=XLSX.utils.encode_cell({r,c});
      const text=norm(cell(ws,a));
      if(!text || !needles.some(n=>text.includes(norm(n))))continue;

      // Primeiro procura números na mesma linha, até 8 colunas à direita.
      for(let k=1;k<=8;k++){
        const b=XLSX.utils.encode_cell({r:r,c:c+k});
        const v=cell(ws,b);
        if(v!=='' && v!==null && v!==undefined && !isNaN(num(v)) && cleanText(v)!==''){
          return {days:num(v),label:cleanText(cell(ws,a)),address:b};
        }
      }

      // Depois procura sob uma coluna intitulada "Dias" perto do rótulo.
      for(let rr=Math.max(range.s.r,r-3);rr<=r;rr++){
        for(let cc=c;cc<=Math.min(range.e.c,c+8);cc++){
          const h=XLSX.utils.encode_cell({r:rr,c:cc});
          if(norm(cell(ws,h))==='dias'){
            const b=XLSX.utils.encode_cell({r,c:cc});
            const v=cell(ws,b);
            if(cleanText(v)!=='' && !isNaN(num(v)))return {days:num(v),label:cleanText(cell(ws,a)),address:b};
          }
        }
      }
      return {days:0,label:cleanText(cell(ws,a)),address:''};
    }
  }
  return {days:0,label:'',address:''};
}
function amountFromLabel(label,fallback){
  const m=String(label||'').match(/(\d+(?:[.,]\d+)?)\s*€/);
  return m?num(m[1]):fallback;
}
function dateRows(ws){
  const dates=[]; let sab=0,dom=0;
  const range=XLSX.utils.decode_range(ws['!ref']||'A1:A1');
  for(let r=range.s.r;r<=range.e.r;r++){
    let rowDate='';
    for(let c=range.s.c;c<=Math.min(range.e.c,range.s.c+3);c++){
      rowDate=excelDate(cell(ws,XLSX.utils.encode_cell({r,c})));
      if(rowDate)break;
    }
    if(!rowDate)continue;
    dates.push(rowDate);
    let rowText='';
    for(let c=range.s.c;c<=Math.min(range.e.c,range.s.c+5);c++)rowText+=' '+norm(cell(ws,XLSX.utils.encode_cell({r,c})));
    if(rowText.includes('sabado')||rowText.includes('feriado'))sab++;
    if(rowText.includes('domingo'))dom++;
  }
  return {dates:[...new Set(dates)].sort(),sab,dom};
}
function parseFile(file,buffer,hash){
  const wb=XLSX.read(buffer,{type:'array',cellDates:true});
  const ws=findSheet(wb);
  const dr=dateRows(ws);
  if(!dr.dates.length)throw new Error('Não foi possível identificar as datas da folha.');
  const dataFinal=dr.dates.at(-1);
  const ajuda=findDaysForLabel(ws,['ajuda de custos','ajudas de custo']);
  const aloj=findDaysForLabel(ws,['alojamento à fatura','alojamento a fatura','alojamento']);
  const cliente=cleanText(cell(ws,'C10')||findLabelValue(ws,['cliente']).value);
  const local=cleanText(cell(ws,'I10')||findLabelValue(ws,['local']).value);
  const processo=cleanText(cell(ws,'M10')||findLabelValue(ws,['nº processo','n. processo','processo']).value);
  const compGerada=dr.sab*.5+dr.dom;
  const compConta=dataFinal>db.settings.dataCorte?compGerada:0;

  return{
    id:crypto.randomUUID(),name:file.name,hash,cliente,local,processo,
    dataInicial:dr.dates[0],dataFinal,
    h25:num(cell(ws,'J23')),h125:num(cell(ws,'K23')),h1375:num(cell(ws,'L23')),h150:num(cell(ws,'M23')),h165:num(cell(ws,'N23')),
    diasAjuda:ajuda.days,diasAloj:aloj.days,
    ajudaUnit:amountFromLabel(ajuda.label,db.settings.ajudaDia),
    alojUnit:db.settings.alojDia,
    ajudaCell:ajuda.address,alojCell:aloj.address,
    compGerada,compConta,imported:new Date().toISOString()
  };
}
fileInput.addEventListener('change',async e=>{
  if(typeof XLSX==='undefined'){msg('A biblioteca de leitura do Excel não carregou. Liga o computador à Internet e volta a abrir a aplicação.','err');return}
  const files=[...e.target.files];
  const duplicateNames=files.map(f=>f.name.toLowerCase()).filter((x,i,a)=>a.indexOf(x)!==i);
  if(duplicateNames.length){msg('Existem ficheiros com o mesmo nome na seleção: '+[...new Set(duplicateNames)].join(', '),'err');return}
  for(const file of files){
    try{
      if(db.sheets.some(s=>s.name.toLowerCase()===file.name.toLowerCase())){msg(`Já existe um ficheiro chamado ${file.name}.`,'warn');continue}
      const buffer=await file.arrayBuffer(); const hash=await sha256(buffer);
      if(db.sheets.some(s=>s.hash===hash)){msg(`${file.name}: conteúdo já importado com outro nome.`,'warn');continue}
      const s=parseFile(file,buffer,hash); db.sheets.push(s); await storeOriginalFile(s.id,file);
      msg(`${file.name}: ${s.diasAjuda} dia(s) de ajuda e ${s.diasAloj} dia(s) de alojamento.`, 'ok');
    }catch(err){msg(`${file.name}: ${err.message}`,'err')}
  }
  e.target.value=''; save();
});
document.querySelectorAll('.tab').forEach(b=>b.onclick=()=>{
  document.querySelectorAll('.tab,.panel').forEach(x=>x.classList.remove('active'));
  b.classList.add('active'); document.getElementById(b.dataset.tab).classList.add('active');
});

function annualLeaveAdded(referenceDate=new Date()){
  const currentYear=referenceDate.getFullYear();
  const startYear=Math.trunc(num(db.settings.primeiroAnoFerias)||2027);
  const daysPerYear=num(db.settings.diasFeriasAnuais)||0;
  if(currentYear<startYear)return 0;
  return (currentYear-startYear+1)*daysPerYear;
}
function businessDaysInclusive(start,end){
  if(!start||!end)return 0;
  const a=new Date(start+'T12:00:00');
  const b=new Date(end+'T12:00:00');
  if(isNaN(a)||isNaN(b)||b<a)return 0;
  let total=0;
  for(let d=new Date(a);d<=b;d.setDate(d.getDate()+1)){
    const day=d.getDay();
    if(day!==0&&day!==6)total++;
  }
  return total;
}
function updateCalculatedDays(){
  if(!usedDate.value||!usedEndDate.value)return;
  const result=businessDaysInclusive(usedDate.value,usedEndDate.value);
  if(result>0)usedDays.value=result;
}
usedDate.addEventListener('change',()=>{
  if(!usedEndDate.value||usedEndDate.value<usedDate.value)usedEndDate.value=usedDate.value;
  usedEndDate.min=usedDate.value;
  updateCalculatedDays();
});
usedEndDate.addEventListener('change',updateCalculatedDays);

usedForm.onsubmit=e=>{
  e.preventDefault();
  const days=num(usedDays.value);
  if(days<=0){msg('Indica um número de dias superior a zero.','err');return}
  db.used.push({
    id:crypto.randomUUID(),
    type:usedType.value,
    date:usedDate.value,
    endDate:usedEndDate.value||usedDate.value,
    days,
    desc:usedDesc.value
  });
  usedDesc.value='';
  usedEndDate.value=usedDate.value;
  usedDays.value=1;
  save();
  msg(`${days} dia(s) descontado(s) do saldo.`, 'ok');
};
payForm.onsubmit=e=>{
  e.preventDefault();
  const month=payMonth.value;
  const existing=db.payments.find(x=>x.month===month);
  const salaryValue=paySalary.value.trim()===''?null:num(paySalary.value);
  const hoursValue=payHours.value.trim()===''?null:num(payHours.value);
  if(salaryValue===null && hoursValue===null){
    if(existing)db.payments=db.payments.filter(x=>x.id!==existing.id);
    save(); msg(`O mês ${month} voltou ao cálculo automático.`,'ok'); return;
  }
  const record={id:existing?.id||crypto.randomUUID(),month,salary:salaryValue,hours:hoursValue};
  if(existing)Object.assign(existing,record);else db.payments.push(record);
  save(); msg(`Correção de ${month} guardada.`,'ok');
};
settingsForm.onsubmit=e=>{
  e.preventDefault(); const fd=new FormData(settingsForm);
  for(const [k,v] of fd)db.settings[k]=k==='dataCorte'?v:num(v);
  db.sheets.forEach(s=>s.compConta=s.dataFinal>db.settings.dataCorte?s.compGerada:0);
  save(); msg('Definições guardadas.','ok');
};
backupBtn.onclick=()=>{
  const blob=new Blob([JSON.stringify(db,null,2)],{type:'application/json'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download=`controlo_horas_backup_${new Date().toISOString().slice(0,10)}.json`; a.click();
  URL.revokeObjectURL(a.href);
};
restoreBtn.onclick=()=>restoreInput.click();
restoreInput.onchange=async()=>{
  try{db=JSON.parse(await restoreInput.files[0].text());save();msg('Backup restaurado.','ok')}
  catch{msg('Backup inválido.','err')}
};
clearBtn.onclick=()=>{
  if(confirm('Apagar todos os dados desta versão de teste?')){
    db=fresh(); localStorage.removeItem(DBKEY); save(); msg('Dados de teste apagados.','ok');
  }
};
search.oninput=renderSheets;


function overtimeValue(g){
 const s=db.settings;
 return num(s.valorHora)*(num(g.h25)*num(s.m25)+num(g.h125)*num(s.m125)+num(g.h1375)*num(s.m1375)+num(g.h150)*num(s.m150)+num(g.h165)*num(s.m165));
}
function irs2026(R){
 R=Math.max(0,num(R)); let t=0,a=0;
 if(R<=920){t=0;a=0}
 else if(R<=1042){t=.125;a=.125*2.60*(1273.85-R)}
 else if(R<=1108){t=.157;a=.157*1.35*(1554.83-R)}
 else if(R<=1154){t=.157;a=94.71}
 else if(R<=1212){t=.212;a=158.18}
 else if(R<=1819){t=.241;a=193.33}
 else if(R<=2119){t=.311;a=320.66}
 else if(R<=2499){t=.349;a=401.19}
 else if(R<=3305){t=.3836;a=487.66}
 else if(R<=5547){t=.3969;a=531.62}
 else if(R<=20221){t=.4495;a=823.40}
 else {t=.4717;a=1272.31}
 const v=Math.max(0,R*t-a);
 return {valor:v,taxaEfetiva:R?v/R:0};
}
function monthPayment(month,g){
 const c=db.payments.find(x=>x.month===month);
 const salario=num(db.settings.salarioBase);
 const horasAuto=overtimeValue(g);
 const horas=(c&&c.hours!==null&&c.hours!==''&&c.hours!==undefined)?num(c.hours):horasAuto;
 const bruto=salario+horas;
 const ss=bruto*(num(db.settings.taxaSS)/100);
 const irsSal=irs2026(salario);
 const irsHoras=horas*(irsSal.taxaEfetiva/2);
 const irs=irsSal.valor+irsHoras;
 const liquidoAuto=Math.max(0,bruto-ss-irs);
 const liquidoTrabalho=(c&&c.salary!==null&&c.salary!==''&&c.salary!==undefined)?num(c.salary):liquidoAuto;
 const ajudas=num(g.ajValue);
 const alojamento=num(g.alValue);
 const totalAjudas=ajudas+alojamento;
 return {salario,horas,bruto,ss,irsSalario:irsSal.valor,irsHoras,irs,liquidoTrabalho,ajudas,alojamento,totalAjudas,corrected:!!c};
}
function updatePaymentPreview(){
 if(!payMonth.value)return;
 const g=groupsByMonth()[payMonth.value]||{h25:0,h125:0,h1375:0,h150:0,h165:0,ajValue:0,alValue:0};
 const p=monthPayment(payMonth.value,g);
 grossAutoCard.textContent=euro(p.bruto);
 ssAutoCard.textContent='- '+euro(p.ss);
 irsAutoCard.textContent='- '+euro(p.irs);
 monthAutoCard.textContent=euro(p.liquidoTrabalho);
 ajudaAutoCard.textContent=euro(p.ajudas);
 alojAutoCard.textContent=euro(p.alojamento);
 expensesAutoCard.textContent=euro(p.totalAjudas);
}
if(payYearFilter)payYearFilter.addEventListener('change',renderPayments);
payMonth.addEventListener('change',()=>{
 const c=db.payments.find(x=>x.month===payMonth.value);
 paySalary.value=c?.salary??'';
 payHours.value=c?.hours??'';
 updatePaymentPreview();
});

function groupsByMonth(){
  const groups={};
  for(const x of db.sheets){
    const k=x.dataFinal.slice(0,7);
    groups[k]??={h25:0,h125:0,h1375:0,h150:0,h165:0,ajDays:0,ajValue:0,alDays:0,alValue:0,comp:0};
    const g=groups[k];
    for(const p of ['h25','h125','h1375','h150','h165'])g[p]+=num(x[p]);
    g.ajDays+=num(x.diasAjuda); g.ajValue+=num(x.diasAjuda)*num(x.ajudaUnit||db.settings.ajudaDia);
    g.alDays+=num(x.diasAloj); g.alValue+=num(x.diasAloj)*num(x.alojUnit||db.settings.alojDia);
    g.comp+=num(x.compConta);
  }
  for(const p of db.payments){
    groups[p.month]??={h25:0,h125:0,h1375:0,h150:0,h165:0,ajDays:0,ajValue:0,alDays:0,alValue:0,comp:0};
  }
  return groups;
}
function render(){
  const s=db.settings;
  for(const el of settingsForm.elements)if(el.name)el.value=s[el.name];
  usedDate.value ||= new Date().toISOString().slice(0,10);
  usedEndDate.value ||= usedDate.value;
  usedEndDate.min=usedDate.value;
  payMonth.value ||= new Date().toISOString().slice(0,7);

  const gained=db.sheets.reduce((a,x)=>a+num(x.compConta),0);
  const used=db.used.reduce((a,x)=>a+num(x.days),0);
  const annualAdded=annualLeaveAdded();
  const totalBalance=s.saldoInicial+gained+annualAdded-used;
  saldoCard.textContent=`${fmt(totalBalance)} dias`;
  if(saldoFeriasCard)saldoFeriasCard.textContent=`${fmt(totalBalance)} dias`;
  if(feriasAnuaisCard)feriasAnuaisCard.textContent=`${fmt(annualAdded)} dias`;
  if(diasUsadosCard)diasUsadosCard.textContent=`${fmt(used)} dias`;
  const totalH=db.sheets.reduce((a,x)=>a+num(x.h25)+num(x.h125)+num(x.h1375)+num(x.h150)+num(x.h165),0);
  horasCard.textContent=`${fmt(totalH)} h`;
  ajudasCard.textContent=euro(db.sheets.reduce((a,x)=>a+num(x.diasAjuda)*num(x.ajudaUnit||s.ajudaDia),0));
  alojCard.textContent=euro(db.sheets.reduce((a,x)=>a+num(x.diasAloj)*num(x.alojUnit||s.alojDia),0));
  renderMonthly(); renderAnnual(); renderSheets(); renderUsed(); renderPayments(); renderDashboard(); renderAnalyses(); renderCalendar();
}
function renderMonthly(){
  const groups=groupsByMonth();
  let html='<tr><th>Mês</th><th>Líquido do recibo</th><th>Dias ajuda</th><th>Ajudas de custo</th><th>Dias aloj.</th><th>Alojamento</th><th>Total ajudas + alojamento</th><th>Comp. nova</th></tr>';
  for(const k of Object.keys(groups).sort().reverse()){
    const g=groups[k],p=monthPayment(k,g);
    html+=`<tr><td>${k}</td><td><strong>${euro(p.liquidoTrabalho)}</strong>${p.corrected?'<br><small>corrigido</small>':''}</td><td>${fmt(g.ajDays)}</td><td>${euro(g.ajValue)}</td><td>${fmt(g.alDays)}</td><td>${euro(g.alValue)}</td><td><strong>${euro(p.totalAjudas)}</strong></td><td>${fmt(g.comp)}</td></tr>`;
  }
  monthlyTable.innerHTML=html;
}
function renderAnnual(){
  const months=groupsByMonth(), years={};
  for(const [m,g] of Object.entries(months)){
    const y=m.slice(0,4),p=monthPayment(m,g);
    years[y]??={liquido:0,ajDays:0,ajValue:0,alDays:0,alValue:0,comp:0};
    const a=years[y]; a.liquido+=p.liquidoTrabalho;a.ajDays+=g.ajDays;a.ajValue+=g.ajValue;a.alDays+=g.alDays;a.alValue+=g.alValue;a.comp+=g.comp;
  }
  let html='<tr><th>Ano</th><th>Total líquido dos recibos</th><th>Dias ajuda</th><th>Ajudas de custo</th><th>Dias aloj.</th><th>Alojamento</th><th>Total ajudas + alojamento</th><th>Comp. nova</th></tr>';
  for(const y of Object.keys(years).sort().reverse()){
    const a=years[y],expenses=a.ajValue+a.alValue;
    html+=`<tr><td>${y}</td><td><strong>${euro(a.liquido)}</strong></td><td>${fmt(a.ajDays)}</td><td>${euro(a.ajValue)}</td><td>${fmt(a.alDays)}</td><td>${euro(a.alValue)}</td><td><strong>${euro(expenses)}</strong></td><td>${fmt(a.comp)}</td></tr>`;
  }
  annualTable.innerHTML=html;
}
function renderSheets(){
  const q=norm(search.value||'');
  let html='<tr><th>Ficheiro</th><th>Período</th><th>Cliente</th><th>Local</th><th>Processo</th><th>25%</th><th>125%</th><th>137,5%</th><th>150%</th><th>165%</th><th>Dias ajuda</th><th>€/dia</th><th>Total ajuda</th><th>Dias aloj.</th><th>Total aloj.</th><th>Comp.</th><th></th></tr>';
  for(const x of [...db.sheets].sort((a,b)=>b.dataFinal.localeCompare(a.dataFinal))){
    if(!norm(`${x.name} ${x.cliente} ${x.local} ${x.processo} ${x.dataFinal}`).includes(q))continue;
    html+=`<tr>
      <td>${x.name}<br><button onclick="openOriginal('${x.id}',encodeURIComponent('${x.name}'))">Abrir original</button></td><td>${x.dataInicial}<br>${x.dataFinal}</td><td>${x.cliente}</td><td>${x.local}</td><td>${x.processo}</td>
      <td>${fmt(x.h25)}</td><td>${fmt(x.h125)}</td><td>${fmt(x.h1375)}</td><td>${fmt(x.h150)}</td><td>${fmt(x.h165)}</td>
      <td><input type="number" step="1" min="0" value="${x.diasAjuda}" onchange="edit('${x.id}','diasAjuda',this.value)"><small>${x.ajudaCell||'não detetada'}</small></td>
      <td><input type="number" step="0.01" min="0" value="${x.ajudaUnit||db.settings.ajudaDia}" onchange="edit('${x.id}','ajudaUnit',this.value)"></td>
      <td>${euro(num(x.diasAjuda)*num(x.ajudaUnit||db.settings.ajudaDia))}</td>
      <td><input type="number" step="1" min="0" value="${x.diasAloj}" onchange="edit('${x.id}','diasAloj',this.value)"><small>${x.alojCell||'não detetada'}</small></td>
      <td>${euro(num(x.diasAloj)*num(x.alojUnit||db.settings.alojDia))}</td>
      <td>${fmt(x.compConta)}</td><td><button onclick="removeSheet('${x.id}')">Apagar</button></td></tr>`;
  }
  sheetsTable.innerHTML=html;
}
window.edit=(id,k,v)=>{const x=db.sheets.find(x=>x.id===id);if(x){x[k]=num(v);save()}};
window.removeSheet=id=>{if(confirm('Apagar esta folha da aplicação?')){db.sheets=db.sheets.filter(x=>x.id!==id);save()}};
function renderUsed(){
  usedTable.innerHTML='<tr><th>Tipo</th><th>Período</th><th>Dias descontados</th><th>Descrição</th><th></th></tr>'+
  [...db.used].sort((a,b)=>b.date.localeCompare(a.date)).map(x=>`<tr><td>${x.type||'Registo'}</td><td>${x.date}${x.endDate&&x.endDate!==x.date?'<br>a '+x.endDate:''}</td><td>${fmt(x.days)}</td><td>${x.desc||''}</td><td><button onclick="removeUsed('${x.id}')">Apagar</button></td></tr>`).join('');
}
window.removeUsed=id=>{db.used=db.used.filter(x=>x.id!==id);save()};
function monthName(month){
 const [y,m]=month.split('-').map(Number);
 return new Date(y,m-1,1).toLocaleDateString('pt-PT',{month:'long',year:'numeric'}).replace(/^./,c=>c.toUpperCase());
}
function renderPayments(){
 const groups=groupsByMonth();
 const months=Object.keys(groups).sort().reverse();
 const years=[...new Set(months.map(m=>m.slice(0,4)))].sort().reverse();
 const currentYear=String(new Date().getFullYear());
 if(payYearFilter){
   const selected=payYearFilter.value|| (years.includes(currentYear)?currentYear:(years[0]||currentYear));
   payYearFilter.innerHTML=years.length?years.map(y=>`<option value="${y}"${y===selected?' selected':''}>${y}</option>`).join(''):`<option value="${currentYear}">${currentYear}</option>`;
 }
 const year=payYearFilter?.value||years[0]||currentYear;
 let totalNet=0,totalTravel=0,totalLodging=0;
 let html='<tr><th>Mês</th><th>Salário bruto</th><th>Horas extra brutas</th><th>Descontos</th><th>Líquido do recibo</th><th>Ajudas de custo</th><th>Alojamento</th><th>Total ajudas + alojamento</th><th>Total global</th></tr>';
 for(const m of months.filter(x=>x.startsWith(year))){
  const g=groups[m],p=monthPayment(m,g);
  const discounts=p.ss+p.irs;
  const global=p.liquidoTrabalho+p.totalAjudas;
  totalNet+=p.liquidoTrabalho; totalTravel+=p.ajudas; totalLodging+=p.alojamento;
  html+=`<tr><td><strong>${monthName(m)}</strong></td><td>${euro(p.salario)}</td><td>${euro(p.horas)}</td><td>- ${euro(discounts)}<br><small>SS ${euro(p.ss)} · IRS ${euro(p.irs)}</small></td><td><strong>${euro(p.liquidoTrabalho)}</strong>${p.corrected?'<br><small>corrigido manualmente</small>':''}</td><td>${euro(p.ajudas)}</td><td>${euro(p.alojamento)}</td><td><strong>${euro(p.totalAjudas)}</strong></td><td><strong>${euro(global)}</strong></td></tr>`;
 }
 if(!months.some(x=>x.startsWith(year)))html+='<tr><td colspan="9" class="emptyCell">Ainda não existem recebimentos registados neste ano.</td></tr>';
 payTable.innerHTML=html;
 if(yearNetCard)yearNetCard.textContent=euro(totalNet);
 if(yearTravelCard)yearTravelCard.textContent=euro(totalTravel);
 if(yearLodgingCard)yearLodgingCard.textContent=euro(totalLodging);
 if(yearGlobalCard)yearGlobalCard.textContent=euro(totalNet+totalTravel+totalLodging);
 updatePaymentPreview();
}
window.editPayment=id=>{
  const x=db.payments.find(x=>x.id===id);if(!x)return;
  payMonth.value=x.month;paySalary.value=x.salary??'';payHours.value=x.hours??'';updatePaymentPreview();
  document.querySelector('[data-tab="recebimentos"]').click();
};
window.removePayment=id=>{db.payments=db.payments.filter(x=>x.id!==id);save()};
render();


// Instalação PWA
let deferredInstallPrompt = null;
const installAppBtn = document.getElementById('installAppBtn');
const connectionStatus = document.getElementById('connectionStatus');

function updateConnectionStatus(){
  if(!connectionStatus) return;
  if(navigator.onLine){
    connectionStatus.textContent = '';
    connectionStatus.classList.remove('offline');
  }else{
    connectionStatus.textContent = '📴 Sem internet: a aplicação continua disponível e guarda os dados neste dispositivo.';
    connectionStatus.classList.add('offline');
  }
}
window.addEventListener('online', updateConnectionStatus);
window.addEventListener('offline', updateConnectionStatus);
updateConnectionStatus();

window.addEventListener('beforeinstallprompt', event => {
  event.preventDefault();
  deferredInstallPrompt = event;
  if(installAppBtn) installAppBtn.hidden = false;
});

if(installAppBtn){
  installAppBtn.addEventListener('click', async () => {
    if(!deferredInstallPrompt){
      alert('No iPhone, use Partilhar e escolha “Adicionar ao ecrã principal”.');
      return;
    }
    deferredInstallPrompt.prompt();
    await deferredInstallPrompt.userChoice;
    deferredInstallPrompt = null;
    installAppBtn.hidden = true;
  });
}

window.addEventListener('appinstalled', () => {
  deferredInstallPrompt = null;
  if(installAppBtn) installAppBtn.hidden = true;
});

// ===== Funcionalidades avançadas v2.0 =====
function openFilesDb(){
  return new Promise((resolve,reject)=>{
    const req=indexedDB.open('controloHorasFiles',1);
    req.onupgradeneeded=()=>req.result.createObjectStore('files');
    req.onsuccess=()=>resolve(req.result); req.onerror=()=>reject(req.error);
  });
}
async function storeOriginalFile(id,file){
  try{const d=await openFilesDb(); const tx=d.transaction('files','readwrite'); tx.objectStore('files').put(file,id)}catch(e){console.warn('Não foi possível arquivar o original',e)}
}
window.openOriginal=async(id,encodedName)=>{
  try{
    const d=await openFilesDb(); const tx=d.transaction('files','readonly');
    const req=tx.objectStore('files').get(id);
    req.onsuccess=()=>{
      if(!req.result){alert('O original desta folha não está guardado neste dispositivo. Volta a importá-lo para o arquivar.');return}
      const a=document.createElement('a');a.href=URL.createObjectURL(req.result);a.download=decodeURIComponent(encodedName);a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1500);
    };
  }catch{alert('Não foi possível abrir o ficheiro original.')}
};

function renderDashboard(){
  if(!dashMonth)return;
  dashMonth.value ||= new Date().toISOString().slice(0,7);
  const g=groupsByMonth()[dashMonth.value]||{h25:0,h125:0,h1375:0,h150:0,h165:0,ajValue:0,alValue:0};
  const p=monthPayment(dashMonth.value,g);
  const h=['h25','h125','h1375','h150','h165'].reduce((a,k)=>a+num(g[k]),0);
  dashNetCard.textContent=euro(p.liquidoTrabalho); dashHoursMonthCard.textContent=`${fmt(h)} h`;
  dashTravelCard.textContent=euro(p.ajudas); dashLodgingCard.textContent=euro(p.alojamento);
}
if(dashMonth)dashMonth.addEventListener('change',renderDashboard);

function drawBars(canvas,labels,values,title,suffix='€'){
  if(!canvas)return; const ctx=canvas.getContext('2d'); const ratio=devicePixelRatio||1;
  const w=Math.max(700,canvas.parentElement.clientWidth-24),h=240; canvas.width=w*ratio;canvas.height=h*ratio;canvas.style.width=w+'px';canvas.style.height=h+'px';ctx.scale(ratio,ratio);
  ctx.clearRect(0,0,w,h);ctx.font='13px sans-serif';ctx.fillText(title,12,18);
  const max=Math.max(1,...values), left=48,bottom=38,top=30,plotH=h-top-bottom,plotW=w-left-12,barW=plotW/Math.max(1,values.length);
  ctx.strokeStyle='#cbd8dd';ctx.beginPath();ctx.moveTo(left,top);ctx.lineTo(left,h-bottom);ctx.lineTo(w-12,h-bottom);ctx.stroke();
  values.forEach((v,i)=>{const bh=(v/max)*plotH;ctx.fillStyle='#0b5e75';ctx.fillRect(left+i*barW+barW*.18,h-bottom-bh,barW*.64,bh);ctx.fillStyle='#20343c';ctx.textAlign='center';ctx.fillText(labels[i],left+i*barW+barW/2,h-bottom+17);ctx.fillText(suffix==='€'?Math.round(v).toLocaleString('pt-PT'):fmt(v),left+i*barW+barW/2,Math.max(top+12,h-bottom-bh-5));});
}
function renderAnalyses(){
  const yearSel=document.getElementById('analysisYear'); if(!yearSel)return;
  const groups=groupsByMonth(),years=[...new Set(Object.keys(groups).map(x=>x.slice(0,4)))].sort().reverse(); const cy=String(new Date().getFullYear());
  const selected=yearSel.value||years[0]||cy;yearSel.innerHTML=(years.length?years:[cy]).map(y=>`<option ${y===selected?'selected':''}>${y}</option>`).join('');
  const year=yearSel.value; const months=Array.from({length:12},(_,i)=>`${year}-${String(i+1).padStart(2,'0')}`);
  const money=[],hours=[]; months.forEach(m=>{const g=groups[m]||{h25:0,h125:0,h1375:0,h150:0,h165:0,ajValue:0,alValue:0};const p=monthPayment(m,g);money.push(p.liquidoTrabalho+p.totalAjudas);hours.push(['h25','h125','h1375','h150','h165'].reduce((a,k)=>a+num(g[k]),0));});
  drawBars(document.getElementById('moneyChart'),months.map(m=>m.slice(5)),money,'Total recebido por mês','€');
  drawBars(document.getElementById('hoursChart'),months.map(m=>m.slice(5)),hours,'Horas extra por mês','h');
  const mode=document.getElementById('analysisGroup')?.value||'cliente',agg={};
  db.sheets.filter(x=>x.dataFinal.startsWith(year)).forEach(x=>{const key=(x[mode]||'Sem indicação').trim()||'Sem indicação';agg[key]??={folhas:0,h:0,aj:0,al:0};const a=agg[key];a.folhas++;a.h+=['h25','h125','h1375','h150','h165'].reduce((z,k)=>z+num(x[k]),0);a.aj+=num(x.diasAjuda)*num(x.ajudaUnit||db.settings.ajudaDia);a.al+=num(x.diasAloj)*num(x.alojUnit||db.settings.alojDia)});
  document.getElementById('groupTable').innerHTML='<tr><th>'+ (mode==='cliente'?'Cliente':'Local/obra') +'</th><th>Folhas</th><th>Horas extra</th><th>Ajudas</th><th>Alojamento</th></tr>'+Object.entries(agg).sort((a,b)=>b[1].h-a[1].h).map(([k,a])=>`<tr><td>${k}</td><td>${a.folhas}</td><td>${fmt(a.h)} h</td><td>${euro(a.aj)}</td><td>${euro(a.al)}</td></tr>`).join('');
  const annualLeave=annualLeaveAdded(new Date(Number(year),11,31)),used=db.used.filter(x=>x.date.startsWith(year)).reduce((a,x)=>a+num(x.days),0),comp=db.sheets.filter(x=>x.dataFinal.startsWith(year)).reduce((a,x)=>a+num(x.compConta),0);
  document.getElementById('fullAnnualTable').innerHTML=`<tr><th>Ano</th><th>Líquido recibos</th><th>Horas extra</th><th>Ajudas</th><th>Alojamento</th><th>Comp. gerada</th><th>Dias usados</th><th>Férias adicionadas acumuladas</th></tr><tr><td>${year}</td><td>${euro(money.reduce((a,v,i)=>a+(monthPayment(months[i],groups[months[i]]||{}).liquidoTrabalho),0))}</td><td>${fmt(hours.reduce((a,v)=>a+v,0))} h</td><td>${euro(months.reduce((a,m)=>a+num(groups[m]?.ajValue),0))}</td><td>${euro(months.reduce((a,m)=>a+num(groups[m]?.alValue),0))}</td><td>${fmt(comp)}</td><td>${fmt(used)}</td><td>${fmt(annualLeave)}</td></tr>`;
}
document.getElementById('analysisYear')?.addEventListener('change',renderAnalyses);document.getElementById('analysisGroup')?.addEventListener('change',renderAnalyses);window.addEventListener('resize',()=>setTimeout(renderAnalyses,100));

function renderCalendar(){
  const inp=document.getElementById('calendarMonth');if(!inp)return;inp.value ||= new Date().toISOString().slice(0,7);const [y,m]=inp.value.split('-').map(Number),first=new Date(y,m-1,1),days=new Date(y,m,0).getDate();let html='';const offset=(first.getDay()+6)%7;for(let i=0;i<offset;i++)html+='<div class="calendarDay empty"></div>';
  for(let d=1;d<=days;d++){const date=`${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;const sh=db.sheets.filter(x=>date>=x.dataInicial&&date<=x.dataFinal),u=db.used.filter(x=>date>=x.date&&date<=(x.endDate||x.date));let marks='';if(sh.length)marks+='🟢';if(sh.some(x=>x.diasAjuda>0))marks+='🔵';if(sh.some(x=>x.diasAloj>0))marks+='🟠';if(u.some(x=>x.type==='Férias'))marks+='🔴';if(u.some(x=>x.type!=='Férias'))marks+='🟣';html+=`<div class="calendarDay" data-date="${date}"><div class="n">${d}</div><div class="marks">${marks}</div></div>`}
  document.getElementById('calendarGrid').innerHTML=html;document.querySelectorAll('.calendarDay[data-date]').forEach(el=>el.onclick=()=>showCalendarDay(el.dataset.date));
}
function showCalendarDay(date){const sh=db.sheets.filter(x=>date>=x.dataInicial&&date<=x.dataFinal),u=db.used.filter(x=>date>=x.date&&date<=(x.endDate||x.date));document.getElementById('calendarDetail').innerHTML=`<strong>${new Date(date+'T12:00').toLocaleDateString('pt-PT',{dateStyle:'full'})}</strong><br>${sh.length?sh.map(x=>`Trabalho: ${x.cliente||'Sem cliente'} · ${x.local||'Sem local'} · ${x.name}`).join('<br>'):'Sem folha de trabalho'}${u.length?'<br>'+u.map(x=>`${x.type}: ${x.desc||''}`).join('<br>'):''}`}
document.getElementById('calendarMonth')?.addEventListener('change',renderCalendar);

async function extractPdfText(file){
  const pdfjs=await import('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs');pdfjs.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs';const pdf=await pdfjs.getDocument({data:await file.arrayBuffer()}).promise;let text='';for(let i=1;i<=pdf.numPages;i++){const page=await pdf.getPage(i),c=await page.getTextContent();text+=' '+c.items.map(x=>x.str).join(' ')}return text;
}
function moneyNear(text,words){for(const w of words){const i=text.toLowerCase().indexOf(w);if(i>=0){const part=text.slice(i,i+130),ms=[...part.matchAll(/\d{1,3}(?:[ .]\d{3})*(?:[,.]\d{2})/g)];if(ms.length)return num(ms.at(-1)[0])}}return null}
document.getElementById('checkReceiptBtn')?.addEventListener('click',async()=>{const file=document.getElementById('receiptFile').files[0],month=document.getElementById('receiptMonth').value||new Date().toISOString().slice(0,7),out=document.getElementById('receiptResult');if(!file){out.innerHTML='<p class="errText">Seleciona um recibo em PDF.</p>';return}out.innerHTML='<p>A analisar o recibo…</p>';try{const text=await extractPdfText(file);if(text.trim().length<30)throw new Error('O PDF parece ser uma imagem e não contém texto legível.');const g=groupsByMonth()[month]||{},p=monthPayment(month,g);const found={liquido:moneyNear(text,['líquido','liquido','valor líquido']),ss:moneyNear(text,['segurança social','seguranca social']),irs:moneyNear(text,['irs']),horas:moneyNear(text,['trabalho suplementar','horas extra'])};const rows=[['Líquido do recibo',p.liquidoTrabalho,found.liquido],['Segurança Social',p.ss,found.ss],['IRS',p.irs,found.irs],['Horas extra brutas',p.horas,found.horas]];out.innerHTML='<div class="resultCard"><h3>Resultado da comparação</h3><table>'+rows.map(([n,a,b])=>{const ok=b!==null&&Math.abs(a-b)<1;return `<tr><td>${n}</td><td>Aplicação: ${euro(a)}</td><td>Recibo: ${b===null?'não detetado':euro(b)}</td><td class="${ok?'okText':'warnText'}">${ok?'✓ Coincide':'⚠ Verificar'}</td></tr>`}).join('')+'</table></div>'}catch(e){out.innerHTML=`<p class="errText">${e.message}</p>`}});
