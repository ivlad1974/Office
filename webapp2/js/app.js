/* app.js — управление UI: заказы (localStorage), редактор листов, реестры, печать */
(function(){
const C=window.Calc, $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const LS='mest_webapp2';

/* ================= ХРАНИЛИЩЕ ================= */
let store={orders:[],r2:[],our:[],customers:[]};
function loadStore(){
  try{ const s=JSON.parse(localStorage.getItem(LS)); if(s&&s.orders) store=s; }catch(e){}
  if(!store.orders.length){
    // сид: текущий заказ из SEED + записи Реестра ЗАКАЗОВ (суммы) как список заказов
    store.orders=[seedOrder()];
    store.r2=SEED.r2.slice();
    store.our=SEED.our.slice();
    store.customers=SEED.customers.slice();
    save();
  }
}
function seedOrder(){
  return {id:'o-seed',created:new Date().toISOString(),
    meta:{objectName:SEED.meta.objectName,org:SEED.meta.org,customerList:SEED.meta.customerList,
      source:'',dogovorDate:SEED.meta.dogovorDate,stroyka:SEED.meta.stroyka,obekt:SEED.meta.obekt,
      dogovorNum:SEED.meta.dogovorNum,coefMat:1,coefWork:1,ks2date:'2026-10-09',
      dates:SEED.meta.dates,status:SEED.meta.status,ndsMode:'НДС 22%',perFrom:'2026-09-08',perTo:'2026-10-09'},
    rows:JSON.parse(JSON.stringify(SEED.rows)),vols:JSON.parse(JSON.stringify(SEED.vols))};
}
function newBlankOrder(){
  return {id:'o-'+Date.now(),created:new Date().toISOString(),
    meta:{objectName:'',org:store.our[0]?.list||'МЭС',customerList:'',source:'',dogovorDate:'',
      stroyka:'',obekt:'',dogovorNum:'',coefMat:1,coefWork:1,
      ks2date:new Date().toISOString().slice(0,10),dates:[],status:Array(9).fill('----'),
      ndsMode:'НДС 22%',perFrom:'',perTo:''},
    rows:[{t:null,n:'1.Новый раздел',u:null,q:null,p:null},{t:'Материалы',n:'',u:'шт',q:0,p:0}],
    vols:{}};
}
function save(){ localStorage.setItem(LS,JSON.stringify(store)); }
let cur=null; // открытый заказ

/* ================= НАВИГАЦИЯ ================= */
$$('.tab').forEach(b=>b.onclick=()=>{
  $$('.tab').forEach(x=>x.classList.remove('active')); b.classList.add('active');
  $$('.view').forEach(v=>v.classList.add('hidden'));
  $('#view-'+b.dataset.view).classList.remove('hidden');
  if(b.dataset.view==='reestr2') renderR2();
  if(b.dataset.view==='kontr') renderKontr();
  if(b.dataset.view==='orders') renderOrders();
});
$('#ed-tabs') && $$('.sheet-tab').forEach(b=>b.onclick=()=>{
  $$('.sheet-tab').forEach(x=>x.classList.remove('active')); b.classList.add('active');
  const keep=['materials','smeta','ks6','ks2','ks3','export'];
  keep.forEach(k=>$('#panel-'+k).classList.toggle('hidden',k!==b.dataset.sheet));
  renderSheet(b.dataset.sheet);
});
function openEditor(order){ cur=order; $('#view-orders').classList.add('hidden'); $('#view-editor').classList.remove('hidden');
  $$('.tab').forEach(x=>x.classList.toggle('active',x.dataset.view==='editor'));
  fillHeader(); renderSheet($('.sheet-tab.active').dataset.sheet||'materials'); }

/* ================= СПИСОК ЗАКАЗОВ ================= */
function renderOrders(){
  const q=$('#order-search').value.toLowerCase();
  const tb=$('#orders-table tbody'); tb.innerHTML='';
  store.orders.filter(o=>o.meta.objectName.toLowerCase().includes(q)).forEach((o,i)=>{
    C.calcMaterials(o); C.calcKS2(o);
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${i+1}</td><td>${esc(o.meta.objectName)}</td><td>${esc(o.meta.org)}</td>
      <td>${esc(o.meta.customerList)}</td><td>${o.meta.dogovorDate||''}</td>
      <td class="num">${C.money(o.totals.all)}</td><td class="num">${C.money(o.ks2.total)}</td>
      <td>${o.created.slice(0,10)}</td>
      <td><button class="btn mini" data-open="${o.id}">Открыть</button>
          <button class="btn mini danger" data-del="${o.id}">✕</button></td>`;
    tb.appendChild(tr);
  });
  tb.querySelectorAll('[data-open]').forEach(b=>b.onclick=()=>openEditor(store.orders.find(o=>o.id===b.dataset.open)));
  tb.querySelectorAll('[data-del]').forEach(b=>b.onclick=()=>{
    if(confirm('Удалить заказ?')){store.orders=store.orders.filter(o=>o.id!==b.dataset.del);save();renderOrders();}});
}
$('#order-search').oninput=renderOrders;
$('#btn-new-order').onclick=()=>{const o=newBlankOrder();store.orders.push(o);save();openEditor(o);};

/* ================= РЕДАКТОР: шапка ================= */
function fillHeader(){
  $('#ed-title').textContent='Заказ: '+(cur.meta.objectName||'(без названия)');
  $('#f-object').value=cur.meta.objectName||'';
  $('#f-org').innerHTML=store.our.map(o=>`<option ${o.list===cur.meta.org?'selected':''}>${esc(o.list)}</option>`).join('');
  $('#f-customer').innerHTML='<option value="">—</option>'+store.customers.map(c=>`<option ${c.list===cur.meta.customerList?'selected':''}>${esc(c.list)}</option>`).join('');
  $('#f-source').value=cur.meta.source||''; $('#f-ddate').value=cur.meta.dogovorDate||'';
  $('#f-stroyka').value=cur.meta.stroyka||''; $('#f-obekt').value=cur.meta.obekt||'';
  $('#f-dogovor').value=cur.meta.dogovorNum||'';
  $('#f-coefm').value=cur.meta.coefMat??1; $('#f-coefw').value=cur.meta.coefWork??1;
  $('#hdr-coefm').value=cur.meta.coefMat??1; $('#hdr-coefw').value=cur.meta.coefWork??1;
  $('#f-ks2date').value=cur.meta.ks2date||''; $('#f-per-from').value=cur.meta.perFrom||''; $('#f-per-to').value=cur.meta.perTo||'';
  bindH('#f-object','objectName');bindH('#f-org','org');bindH('#f-customer','customerList');
  bindH('#f-source','source');bindH('#f-ddate','dogovorDate');bindH('#f-stroyka','stroyka');
  bindH('#f-obekt','obekt');bindH('#f-dogovor','dogovorNum');bindH('#f-ks2date','ks2date');
  bindH('#f-per-from','perFrom');bindH('#f-per-to','perTo');
  $('#f-coefm').onchange=e=>{cur.meta.coefMat=C.num(e.target.value);$('#hdr-coefm').value=e.target.value;recalcAll();};
  $('#f-coefw').onchange=e=>{cur.meta.coefWork=C.num(e.target.value);$('#hdr-coefw').value=e.target.value;recalcAll();};
  $('#hdr-coefm').onchange=e=>{cur.meta.coefMat=C.num(e.target.value);$('#f-coefm').value=e.target.value;recalcAll();};
  $('#hdr-coefw').onchange=e=>{cur.meta.coefWork=C.num(e.target.value);$('#f-coefw').value=e.target.value;recalcAll();};
}
function bindH(sel,key){$(sel).onchange=e=>{cur.meta[key]=e.target.value; recalcAll();};}
function recalcAll(){ if(!cur)return; C.calcKS3(cur); ["materials","smeta","ks6"].forEach(n=>{const el=document.getElementById("panel-"+n);if(el&&!el.classList.contains("hidden"))renderSheet(n);}); save(); }  // автосохранение заказа
function renderCurrent(){ renderSheet($('.sheet-tab.active').dataset.sheet||'materials'); }

/* ================= ЛИСТЫ ================= */
function renderSheet(name){
  if(!cur) return;
  if(name==='materials') renderMaterials();
  if(name==='smeta') renderSmeta();
  if(name==='ks6'){ renderKs6(); renderObems(); }   // объемы закрытия — внутри панели КС-6, без дублей
  if(name==='ks2') renderKs2Doc();
  if(name==='ks3') renderKs3Doc();
  if(name==='export') renderExport();
}

/* ---- Материалы (редактируемая таблица) ---- */
function renderMaterials(){
  C.calcMaterials(cur);
  const tb=$('#mat-table tbody'); tb.innerHTML='';
  cur.rows.forEach((row,i)=>{
    const tr=document.createElement('tr');
    if(!row.t){ tr.className='section-row';
      tr.innerHTML=`<td></td><td></td><td colspan="8"><input class="sec-name" data-i="${i}" value="${esc(row.n||'')}"></td><td><button class="btn mini danger" data-rm="${i}">✕</button></td>`;
    } else {
      tr.innerHTML=`<td>${row.pp||''}</td><td>${esc(row.t)}</td>
        <td><input data-f="n" data-i="${i}" value="${esc(row.n||'')}"></td>
        <td><input data-f="u" data-i="${i}" value="${esc(row.u||'')}" size="4"></td>
        <td><input data-f="q" data-i="${i}" class="num" value="${row.q??''}"></td>
        <td><input data-f="p" data-i="${i}" class="num" value="${row.p??''}"></td>
        <td class="num">${C.money(row.sum)}</td><td class="num">${C.money(row.I)}</td><td class="num">${C.money(row.J)}</td>
        <td><input data-f="note" data-i="${i}" value="${esc(row.note||'')}" size="8"></td>
        <td><button class="btn mini danger" data-rm="${i}">✕</button></td>`;
    }
    tb.appendChild(tr);
  });
  $('#tot-mat').textContent=C.money(cur.totals.mat);
  $('#tot-work').textContent=C.money(cur.totals.work);
  $('#tot-all').textContent=C.money(cur.totals.all);
  $('#hdr-total').textContent=C.money(cur.totals.all);
  tb.querySelectorAll('input[data-f]').forEach(inp=>inp.onchange=e=>{
    const r=cur.rows[+e.target.dataset.i], f=e.target.dataset.f;
    r[f]=(f==='q'||f==='p')?C.num(e.target.value):e.target.value;
    recalcAll();
  });
  tb.querySelectorAll('.sec-name').forEach(inp=>inp.onchange=e=>{cur.rows[+e.target.dataset.i].n=e.target.value;renderCurrent();});
  tb.querySelectorAll('[data-rm]').forEach(b=>b.onclick=()=>{cur.rows.splice(+b.dataset.rm,1);renumberVols();recalcAll();});
}
$$('.add-row').forEach(b=>b.onclick=()=>{
  const kind=b.dataset.kind;
  if(kind==='header') cur.rows.push({t:null,n:'Новый заголовок',u:null,q:null,p:null});
  else cur.rows.push({t:kind,n:'',u:'шт',q:0,p:0});
  renumberVols(); recalcAll();
});
function renumberVols(){ /* позиции после удаления/добавления сдвигаются — пересчёт заново при рендере */ }

/* ---- Смета ---- */
function renderSmeta(){
  C.calcSmeta(cur);
  const t=cur.smetaTotals;
  $('#sm-total').textContent=C.money(t.total);
  $('#sm-done').textContent=C.money(t.doneNds);
  $('#sm-ytd').textContent=C.money(t.ytd);
  const tb=$('#smeta-table tbody'); tb.innerHTML='';
  cur.smeta.forEach(x=>{
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${x.pos}</td><td>${esc(x.name)}</td><td>${esc(x.u||'')}</td>
      <td class="num">${C.money(x.price)}</td><td class="num">${x.qty}</td>
      <td class="num">${C.money(x.costMat)}</td><td class="num">${C.money(x.costUsl)}</td>
      <td class="num">${x.doneAll}</td><td class="num">${C.money(x.doneSum)}</td>
      <td class="num">${x.rest}</td><td class="num">${C.money(x.restSum)}</td>`;
    tb.appendChild(tr);
  });
}

/* ---- Объемы закрытия (блоки под журналом КС-6) ---- */
function renderObems(){
  C.calcObems(cur);
  const box=$('#obems-blocks'); if(!box) return; box.innerHTML='';
  cur.obems.forEach(m=>{
    if(!m.rows.length && !m.date) return;
    const d=document.createElement('div'); d.className='obem-block';
    d.innerHTML=`<h4>${C.fmtMY(m.date)} <span class="tag ${m.status==='Закрытие'?'ok':''}">${m.status}</span> (${m.rows.length} поз.)</h4>
      <table class="grid compact"><thead><tr><th>№ позиции</th><th>Наименование</th><th>Объём</th><th>Остаток в смете</th></tr></thead>
      <tbody>${m.rows.map(r=>`<tr><td>${r.pos}</td><td>${esc(r.name)}</td><td class="num">${r.vol}</td><td class="num">${r.restQty}</td></tr>`).join('')}</tbody></table>`;
    box.appendChild(d);
  });
}

/* ---- КС-6 ----
   Выбор столбца редактирования и месяца/года — НЕ в отдельном поле, а прямо
   в шапке таблицы КС-6 (как строка дат F32:N32 в оригинальном листе Excel).
   Ровно ОДИН столбец может иметь статус «Закрытие», остальные — «----».
   В выбранном столбце дата отображается двумя выпадающими списками прямо
   в <th>: «Август» и «2026» (формат mmmm yyyy). Дублирующего блока настроек нет. */
function renderKs6(){
  C.calcSmeta(cur);
  if(!cur.meta.mmyy) cur.meta.mmyy=[];      // миграция старых заказов: даты -> {m,y}
  C.months(cur).forEach(m=>{ if(!cur.meta.mmyy[m.i]) cur.meta.mmyy[m.i]={m:m.date.getMonth(),y:m.date.getFullYear()}; });
  const months=C.months(cur);

  const onHeaderChange=e=>{
    const t=e.target, i=+t.dataset.hi; if(!i&&i!==0) return;
    if(t.hasAttribute('data-st')){                       // переключение статуса столбца
      if(!cur.meta.status||cur.meta.status.length!==9) cur.meta.status=Array(9).fill('----');
      if(t.value==='Закрытие') cur.meta.status.fill('----');   // ровно один «Закрытие»
      cur.meta.status[i]=t.value;
    } else {                                             // смена месяца / года в шапке
      const mm=cur.meta.mmyy[i];
      const m=t.hasAttribute('data-mm')?+t.value:mm.m;
      const y=t.hasAttribute('data-yy')?+t.value:mm.y;
      C.setMonthYear(cur,i,m,y);
    }
    recalcAll();
  };
  $('#ks6-table thead').onchange=onHeaderChange;

  const P=C.positions(cur);
  const thead=$('#ks6-table thead');
  thead.innerHTML=`<tr><th>№ п/п</th><th>№ поз.</th><th>Наименование</th><th>Ед.изм.</th><th>Кол-во</th>`+
    months.map(m=>{
      const act=m.status==='Закрытие';
      const inner = act
        ? `<select data-hi="${m.i}" data-mm title="Месяц закрытия">${C.MONTH_NAMES.map((n,mi)=>`<option value="${mi}"${mi===m.date.getMonth()?' selected':''}>${n}</option>`).join('')}</select>
           <select data-hi="${m.i}" data-yy title="Год закрытия">${Array.from({length:8},(_,k)=>2023+k).map(y=>`<option${y===m.date.getFullYear()?' selected':''}>${y}</option>`).join('')}</select>`
        : `${C.fmtMY(m.date)}`;
      return `<th class="hmonth${act?' unlock':''}"><div class="hdate">${inner}</div>
        <select class="hstatus" data-hi="${m.i}" data-st title="Редактировать можно только один столбец">
          <option value="----"${!act?' selected':''}>----</option>
          <option value="Закрытие"${act?' selected':''}>Закрытие</option>
        </select></th>`;
    }).join('')+`<th>Остаток</th></tr>`;
  const tb=$('#ks6-table tbody'); tb.innerHTML='';
  cur.rows.forEach((row,ri)=>{
    if(!row.t) return;
    const p=P.find(x=>x.idx===ri); if(!p) return;
    const v=cur.vols[p.pos]||{};
    let sum=0; months.forEach(m=>sum+=C.num(v['m'+m.i]));
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${p.pos}</td><td>${p.pos}</td><td>${esc(row.n)}</td><td>${esc(row.u||'')}</td><td class="num">${C.num(row.q)}</td>`+
      months.map(m=>`<td><input class="num vol ${m.status==='Закрытие'?'':'locked'}" data-pos="${p.pos}" data-m="${m.i}" value="${v['m'+m.i]??''}" ${m.status==='Закрытие'?'':'disabled'}></td>`).join('')+
      `<td class="num ${sum>C.num(row.q)?'over':''}">${sum>C.num(row.q)?'Превышение':C.num(row.q)-sum}</td>`;
    tb.appendChild(tr);
  });
  tb.querySelectorAll('.vol').forEach(inp=>inp.onchange=e=>{
    const pos=e.target.dataset.pos,m=e.target.dataset.m,val=C.num(e.target.value);
    cur.vols=pos in cur.vols?cur.vols:{...cur.vols};
    if(!cur.vols[pos]) cur.vols[pos]={};
    if(val>0) cur.vols[pos]['m'+m]=val; else delete cur.vols[pos]['m'+m];
    recalcAll();
  });
}

/* ---- КС-2 документ ---- */
function renderKs2Doc(){
  C.calcKS2(cur);
  const k=cur.ks2, cust=(store.customers.find(c=>c.list===cur.meta.customerList)||{});
  const our=store.our.find(o=>o.list===cur.meta.org)||{};
  let rowsHtml='';
  k.lines.forEach(l=>{
    if(l.type==='head') rowsHtml+=`<tr class="doc-head"><td colspan="5"><b>${esc(l.name)}</b></td><td></td><td></td></tr>`;
    else rowsHtml+=`<tr><td>${l.pp}</td><td>${l.pos}</td><td>${esc(l.name)}</td><td>${esc(l.u||'')}</td>
      <td class="num">${l.vol}</td><td class="num">${C.money(l.price)}</td><td class="num">${C.money(l.sum)}</td></tr>`;
  });
  $('#ks2-doc').innerHTML=`
  <div class="doc-head-line">${C.fmtD(C.parseD(cur.meta.ks2date))} &nbsp;&nbsp; Унифицированная форма № КС-2, ОКУД 0322005</div>
  <div>Заказчик (Генподрядчик) — ${esc(cust.name||cur.meta.customerList)}</div>
  <div>Подрядчик (Субподрядчик) — ${esc(our.name||cur.meta.org)}</div>
  <div>Стройка — ${esc(cur.meta.stroyka)}</div><div>Объект — ${esc(cur.meta.obekt)}</div>
  <div>Договор подряда: ${esc(String(cur.meta.dogovorNum).replace(/\n/g,' '))} от ${cur.meta.dogovorDate||''}</div>
  <h3 class="center">АКТ № ${k.docNum} О ПРИЕМКЕ ВЫПОЛНЕННЫХ РАБОТ</h3>
  <div>Отчётный период: с ${C.fmtD(k.perFrom)} по ${C.fmtD(k.perTo)} &nbsp;|&nbsp; ${esc(k.restFormula)}</div>
  <table class="grid doc-table"><thead><tr><th>п.п</th><th>поз.по калькуляции</th><th>Наименование работ</th><th>Ед.изм</th>
   <th>Объём</th><th>Цена руб.</th><th>Стоимость руб.</th></tr></thead><tbody>${rowsHtml}</tbody>
   <tfoot><tr><td colspan="6">Итого, в том числе:</td><td class="num">${C.money(k.total)}</td></tr>
   <tr><td colspan="6">Работы</td><td class="num">${C.money(k.work)}</td></tr>
   <tr><td colspan="6">Материалы</td><td class="num">${C.money(k.mat)}</td></tr>
   <tr><td colspan="6"><b>Всего по акту</b></td><td class="num"><b>${C.money(k.total)}</b></td></tr>
   <tr><td colspan="6">${k.ndsLabel}</td><td class="num">${C.money(k.nds)}</td></tr></tfoot></table>
  <div class="signs"><div>Сдал: ${esc(our.sign||'')} — ${esc(our.post||'')} ${esc(our.fio||'')}</div>
  <div>Принял: ${esc(cust.sign||cust.name||'')} — ${esc(cust.post||'')} ${esc(cust.fio||'')}</div></div>`;
}

/* ---- КС-3 документ ---- */
function renderKs3Doc(){
  C.calcKS3(cur);
  const x=cur.ks3, cust=(store.customers.find(c=>c.list===cur.meta.customerList)||{});
  const our=store.our.find(o=>o.list===cur.meta.org)||{};
  $('#ks3-doc').innerHTML=`
  <div class="doc-head-line">Форма № КС-3, ОКУД 0322001 &nbsp; ${C.fmtD(C.parseD(cur.meta.ks2date))}</div>
  <div>Инвестор: ${esc(cust.name||'')}</div><div>Подрядчик: ${esc(our.name||'')}</div>
  <div>Контракт: ${esc(String(cur.meta.dogovorNum).replace(/\n/g,' '))} от ${cur.meta.dogovorDate||''}</div>
  <h3 class="center">СПРАВКА О СТОИМОСТИ ВЫПОЛНЕННЫХ РАБОТ И ЗАТРАТ № ${x.docNum}</h3>
  <div>Отчётный период: с ${x.perFrom} по ${x.perTo}</div>
  <table class="grid doc-table"><thead><tr><th>№ п/п</th><th>Наименование</th><th>с начала проведения работ</th>
   <th>с начала года</th><th>в т.ч. за отчётный период</th></tr></thead><tbody>
   <tr><td>1</td><td>Всего работ и затрат</td><td class="num">${C.money(x.fromStart)}</td><td class="num">${C.money(x.ytd)}</td><td class="num">${C.money(x.forPeriod-x.nds)}</td></tr>
   <tr><td>1.1</td><td>${esc(String(cur.meta.obekt).slice(0,120))}</td><td class="num">${C.money(x.fromStart)}</td><td class="num">${C.money(x.ytd)}</td><td class="num">${C.money(x.forPeriod-x.nds)}</td></tr>
   <tr><td></td><td><b>Итого</b></td><td></td><td></td><td class="num"><b>${C.money(x.forPeriod-x.nds)}</b></td></tr>
   <tr><td></td><td>${x.labelNds}</td><td></td><td></td><td class="num">${C.money(x.nds)}</td></tr>
   <tr><td></td><td><b>Всего с учетом НДС</b></td><td></td><td></td><td class="num"><b>${C.money(x.totalWithNds)}</b></td></tr>
   </tbody></table>
  <div class="signs"><div>Подрядчик: ${esc(our.sign||'')} ${esc(our.fio||'')}</div>
  <div>Заказчик (генподрядчик): ${esc(cust.sign||cust.name||'')}</div></div>`;
}

/* ---- Экспорт ---- */
function renderExport(){
  C.calcKS3(cur);
  const row=C.exportRow(cur);
  const map=[['Источник Заказа','source'],['Дата договора','ddate'],['Имя ОБЪЕКТА в РЕЕСТР','object'],
   ['Название Заказчика','customer'],['Предмет договора','predmet'],['Адрес объекта','addrObj'],
   ['Закрытие на Сумму','closeSum'],['ИНН/КПП','inn'],['Стоимость контракта (ПЛАН)','planCost'],
   ['Стоимость контракта (ФАКТ)','factCost'],['Материалы (ПЛАН)','matPlan'],['Услуги (ПЛАН)','uslPlan'],
   ['Версия программы','version'],['Путь к файлу','path'],['ДатаЗаписи','dateRec'],['Пользователь','user']];
  $('#export-table tbody').innerHTML=map.map(([h,k])=>`<tr><td><b>${h}</b></td><td>${esc(String(row[k]??''))}</td></tr>`).join('');
}

/* ---- кнопки тулбара редактора ---- */
$('#btn-save-order').onclick=()=>{save();toast('Заказ сохранён в браузере (localStorage)');};
$('#btn-export-reestr').onclick=()=>{
  C.calcKS3(cur);
  const row=C.exportRow(cur);
  const ex=store.r2.find(r=>r.object===row.object);
  if(ex){Object.assign(ex,row);toast('Запись обновлена в Реестре 2 (аналог «Переписать?» → Да)');}
  else{store.r2.push(row);toast('Новая запись добавлена в Реестр 2');}
  save(); renderR2();
};
$('#btn-print-ks2').onclick=()=>printDoc('ks2-doc');
$('#btn-print-ks3').onclick=()=>printDoc('ks3-doc');
$('#btn-print-ks6').onclick=()=>{renderKs6();renderObems();printArea('#panel-ks6');};
$('#btn-download-zip').onclick=()=>{
  const blob=new Blob([JSON.stringify(cur,null,1)],{type:'application/json'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download='Заказ_'+(cur.meta.objectName||'new')+'.json';a.click();
};
$('#btn-close-order').onclick=()=>{$('#view-editor').classList.add('hidden');$('#view-orders').classList.remove('hidden');renderOrders();};
function switchSheet(k){$$('.sheet-tab').forEach(b=>b.classList.toggle('active',b.dataset.sheet===k));renderSheet(k);}
function printDoc(id){const w=window.open('','_blank');w.document.write('<html><head><title>Печать</title><link rel="stylesheet" href="css/style.css"></head><body class="print">'+$('#'+id).outerHTML+'</body></html>');w.document.close();w.print();}
function printArea(sel){const w=window.open('','_blank');w.document.write('<html><head><title>Печать</title><link rel="stylesheet" href="css/style.css"></head><body class="print">'+$(sel).innerHTML+'</body></html>');w.document.close();w.print();}

/* ================= РЕЕСТР 2 ================= */
const R2COLS=[['object','Имя ОБЪЕКТА'],['source','Источник'],['ddate','Дата договора'],['customer','Заказчик'],
 ['closeSum','Закрытие на сумму'],['planCost','Контракт ПЛАН'],['factCost','Контракт ФАКТ'],
 ['matPlan','Матер. ПЛАН'],['uslPlan','Усл. ПЛАН'],['inn','ИНН/КПП'],['kurator','Куратор'],
 ['version','Версия'],['dateRec','Дата записи'],['user','Польз.'],['path','Путь к файлу']];
function renderR2(){
  const q=($('#r2-search')?.value||'').toLowerCase();
  $('#r2-count').textContent=store.r2.length;
  $('#r2-table thead').innerHTML='<tr>'+R2COLS.map(c=>`<th>${c[1]}</th>`).join('')+'</tr>';
  const tb=$('#r2-table tbody');tb.innerHTML='';
  store.r2.filter(r=>!q||JSON.stringify(r).toLowerCase().includes(q)).slice(0,500).forEach(r=>{
    const tr=document.createElement('tr');
    tr.innerHTML=R2COLS.map(([k])=>{
      const v=r[k]??'';
      return `<td class="${typeof v==='number'?'num':''}" title="${esc(String(v))}">${typeof v==='number'?C.money(v):esc(String(v).slice(0,60))}</td>`;
    }).join('');
    tb.appendChild(tr);
  });
}
$('#r2-search')&&($('#r2-search').oninput=renderR2);
$('#btn-r2-add').onclick=()=>{store.r2.push({object:'Новый заказ',source:'',ddate:'',customer:'',closeSum:0,planCost:0,factCost:0,matPlan:0,uslPlan:0,inn:'',kurator:'',version:'web',dateRec:new Date().toISOString().slice(0,10),user:'web',path:''});save();renderR2();};

/* ================= КОНТРАГЕНТЫ ================= */
function renderKontr(){
  const otb=$('#our-table tbody');otb.innerHTML='';
  store.our.forEach((o,i)=>{
    const tr=document.createElement('tr');
    tr.innerHTML=['list','name','inn','kpp','okpo','addr','sign','post','fio','nds'].map(k=>`<td><input data-k="${k}" data-i="${i}" value="${esc(o[k]||'')}"></td>`).join('')+`<td><button class="btn mini danger" data-del-o="${i}">✕</button></td>`;
    otb.appendChild(tr);
  });
  otb.querySelectorAll('input').forEach(inp=>inp.onchange=e=>{store.our[+e.target.dataset.i][e.target.dataset.k]=e.target.value;save();});
  otb.querySelectorAll('[data-del-o]').forEach(b=>b.onclick=()=>{store.our.splice(+b.dataset.delO,1);save();renderKontr();});
  const ctb=$('#cust-table tbody');ctb.innerHTML='';
  store.customers.forEach((o,i)=>{
    const tr=document.createElement('tr');
    tr.innerHTML=['list','name','inn','kpp','okpo','addr','sign'].map(k=>`<td><input data-k="${k}" data-i="${i}" value="${esc(o[k]||'')}"></td>`).join('')+`<td><button class="btn mini danger" data-del-c="${i}">✕</button></td>`;
    ctb.appendChild(tr);
  });
  ctb.querySelectorAll('input').forEach(inp=>inp.onchange=e=>{store.customers[+e.target.dataset.i][e.target.dataset.k]=e.target.value;save();});
  ctb.querySelectorAll('[data-del-c]').forEach(b=>b.onclick=()=>{store.customers.splice(+b.dataset.delC,1);save();renderKontr();});
}
$('#btn-our-add').onclick=()=>{store.our.push({list:'Новая',name:'',inn:'',kpp:'',okpo:'',addr:'',sign:'',post:'',fio:'',nds:'НДС 22%'});save();renderKontr();};
$('#btn-cust-add').onclick=()=>{store.customers.push({list:'Новый заказчик',name:'',inn:'',kpp:'',okpo:'',addr:'',sign:''});save();renderKontr();};

/* ================= utils ================= */
function esc(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function toast(t){const el=$('#toast');el.textContent=t;el.classList.remove('hidden');setTimeout(()=>el.classList.add('hidden'),2500);}

loadStore(); renderOrders();
})();
