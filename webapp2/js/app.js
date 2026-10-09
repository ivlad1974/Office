/* app.js — управление UI: заказы (localStorage), редактор листов, реестры, печать */
(function(){
const C=window.Calc, $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const LS='mest_webapp2';
/* Подтверждение удаления: собственный диалог с контрастными темами;
   если window.confirm переопределён тестами — используем его. */
function askConfirm(msg){
  if(typeof window.__TEST_CONFIRM==='function') return Promise.resolve(window.__TEST_CONFIRM(msg));
  if(!document.body || !document.body.classList) return Promise.resolve(false);
  return new Promise(res=>{
    const ov=document.createElement('div'); ov.className='cfrm-overlay';
    ov.innerHTML=`<div class="cfrm-box"><div class="cfrm-msg"></div><div class="cfrm-btns">
      <button class="btn primary cfrm-ok">Да, удалить</button><button class="btn cfrm-no">Отмена</button></div></div>`;
    ov.querySelector('.cfrm-msg').textContent=msg;
    document.body.appendChild(ov);
    const done=v=>{ov.remove();res(v);};
    ov.querySelector('.cfrm-ok').onclick=()=>done(true);
    ov.querySelector('.cfrm-no').onclick=()=>done(false);
    ov.onclick=e=>{if(e.target===ov)done(false);};
    document.addEventListener('keydown',function h(e){if(e.key==='Escape'){done(false);document.removeEventListener('keydown',h);}});
    ov.querySelector('.cfrm-no').focus();
  });
}

/* ================= ХРАНИЛИЩЕ ================= */
let store={orders:[],r2:[],our:[],customers:[],dictMat:[],dictWork:[]};
/* Справочники материалов и работ (аналог листов-баз Excel «Материалы»/«Работы»).
   Заполняются один раз из SEED.rows (1178 строк исходной книги), далее редактируются
   пользователем на вкладке «Справочник». */
function buildDictsFromSeed(){
  const m={},w={};
  SEED.rows.forEach(r=>{
    if(r.t==='Материалы'){ const k=(r.n||'').trim().toLowerCase(); if(k&&!m[k]) m[k]={n:r.n,u:r.u||'',p:r.p||0}; }
    else if(r.t==='Работа'){ const k=(r.n||'').trim().toLowerCase(); if(k&&!w[k]) w[k]={n:r.n,u:r.u||'',p:r.p||0}; }
  });
  return {mat:Object.values(m).sort((a,b)=>a.n.localeCompare(b.n,'ru')),
          work:Object.values(w).sort((a,b)=>a.n.localeCompare(b.n,'ru'))};
}
function loadStore(){
  try{ const s=JSON.parse(localStorage.getItem(LS)); if(s&&s.orders) store=Object.assign(store,s); }catch(e){}
  if(!store.orders.length){
    // сид: текущий заказ из SEED + записи Реестра ЗАКАЗОВ (суммы) как список заказов
    store.orders=[seedOrder()];
    store.r2=SEED.r2.slice();
    store.our=SEED.our.slice();
    store.customers=SEED.customers.slice();
  }
  if(!store.dictMat.length||!store.dictWork.length){   // миграция хранилищ старых версий
    const d=buildDictsFromSeed();
    if(!store.dictMat.length) store.dictMat=d.mat;
    if(!store.dictWork.length) store.dictWork=d.work;
  }
  save();
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
let curSheet='materials';

/* ================= НАВИГАЦИЯ ================= */
$$('.tab').forEach(b=>b.onclick=()=>{
  $$('.tab').forEach(x=>x.classList.remove('active')); b.classList.add('active');
  $$('.view').forEach(v=>v.classList.add('hidden'));
  $('#view-'+b.dataset.view).classList.remove('hidden');
  if(b.dataset.view==='reestr2') renderR2();
  if(b.dataset.view==='kontr') renderKontr();
  if(b.dataset.view==='orders') renderOrders();
  if(b.dataset.view==='sprav') renderSprav();
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
  autoFitColumns('#orders-table');
  tb.querySelectorAll('[data-del]').forEach(b=>b.onclick=async()=>{
    if(await askConfirm('Удалить заказ?')){store.orders=store.orders.filter(o=>o.id!==b.dataset.del);save();renderOrders();}});
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

/* ---- Материалы / Работы (редактируемая таблица) ----
   Возможности:
   - при выборе позиции из всплывающего справочника (кнопка 🔎 в ячейке «Наименование»)
     либо при точном вводе наименования — Тип («Материалы»/«Работа»), Ед.изм. и Цена
     проставляются САМИ, в зависимости от того, из какого справочника выбрано (applyDictToRow);
   - подзаголовки при вводе строки; выбор типа «Подзаголовок» превращает строку в подзаголовок;
   - вставка строки МЕЖДУ существующими (кнопки слева △＋ Выше / ▽＋ Ниже и на любой строке);
   - копирование строки и группы строк (чекбокс + Shift+клик по диапазону), вставка в любое место;
   - удаление ВСЕГДА с подтверждением;
   - шапка таблицы прикреплена (sticky внутри контейнера прокрутки — не скрывается никогда);
   - ширина столбцов — авто по содержимому + ручное изменение за правую границу заголовка
     (двойной клик по границе — снова авто). */
let rowClipboard=[];   // буфер копирования строк заказа
let selRows=new Set(); // выделенные строки (для копирования группы)
let lastClickedRow=null; // для выделения диапазона через Shift
let focusRow=null;     // строка, выбранная кликом (ориентир для тулбара слева)
/* data-i в DOM всегда = актуальный индекс в cur.rows, т.к. таблица перерисовывается
   после любой вставки/удаления. rowsInDom() — страховка от устаревшего DOM. */
function rowsInDom(){ return $$('#mat-table tbody tr').map(tr=>{
  const el=tr.querySelector('[data-i]'); return el?+el.dataset.i:-1; }).filter(i=>i>=0); }
function dictLookup(name,t){
  const k=(name||'').trim().toLowerCase(); if(!k) return null;
  const src = t==='Работа'? store.dictWork : t==='Материалы'? store.dictMat : (store.dictMat.concat(store.dictWork));
  const hit=src.find(d=>d.n.trim().toLowerCase()===k);
  return hit||null;
}
/* Поиск по обоим справочникам с указанием источника (для всплывающего выбора) */
function dictFindAll(q,limit){
  const k=(q||'').trim().toLowerCase();
  const out=[];
  const push=(list,type)=>list.forEach((d,i)=>{
    if(k && !d.n.toLowerCase().includes(k)) return;
    out.push({d,type,i});
  });
  push(store.dictMat,'Материалы'); push(store.dictWork,'Работа');
  if(!k) out.sort((a,b)=>a.d.n.localeCompare(b.d.n,'ru'));   // без запроса — алфавит по объединению
  return out.slice(0,limit||300);
}
/* Автоопределение ТИПА по справочнику: наименование есть только в работах — «Работа»,
   только в материалах — «Материалы»; если не найдено — оставляем текущий тип. */
function dictTypeOf(k){
  const inM=store.dictMat.some(d=>d.n.trim().toLowerCase()===k);
  const inW=store.dictWork.some(d=>d.n.trim().toLowerCase()===k);
  if(inW&&!inM) return 'Работа';
  if(inM&&!inW) return 'Материалы';
  return null;
}
/* Проставить в строку данные из справочника.
   preferType — если пользователь выбрал позицию из конкретного справочника
   (всплывающее окно «🔎 Справочник»), тип берётся оттуда безусловно. */
function applyDictToRow(r,preferType){
  const k=(r.n||'').trim().toLowerCase(); if(!k) return false;
  let d=null,t=preferType||null;
  if(t){ d=dictLookup(r.n,t); }
  if(!d){ t=dictTypeOf(k)||r.t; d=dictLookup(r.n,t)||dictLookup(r.n,null); }
  if(d){
    if(t) r.t=t;
    r.u=d.u; r.p=d.p;
    toast(`Справочник «${r.t}»: ${String(d.n).slice(0,40)} — ед. ${d.u}, цена ${C.money(d.p)} проставлены автоматически`);
    return true;
  }
  return false;
}
/* ---- Всплывающий выбор позиции из справочника (Материалы / Работы) ----
   Открывается кнопкой 🔎 в ячейке «Наименование». Выбор из списка САМ
   проставляет Тип (по источнику выбора), Ед.изм. и Цена. */
let dictPopEl=null;
function closeDictPop(){ if(dictPopEl){dictPopEl.remove();dictPopEl=null;} }
document.addEventListener('mousedown',e=>{
  if(dictPopEl && !dictPopEl.contains(e.target)) closeDictPop();
});
window.addEventListener('scroll',closeDictPop,true);
function openDictPop(inp,i){
  if(!document.body||!cur||!cur.rows[i]) return;
  closeDictPop();
  const pop=document.createElement('div'); pop.className='dictpop';
  pop.innerHTML=`<div style="display:flex;gap:6px;align-items:center;margin-bottom:6px">
      <input class="dp-q" type="text" placeholder="Поиск по справочнику…" style="width:260px">
      <span class="dp-n tagline"></span></div>
    <div class="dp-list" style="max-height:300px;overflow:auto"></div>
    <div class="dp-hint">Клик по строке — выбрать позицию: Тип, ед.изм. и цена проставятся сами.</div>`;
  document.body.appendChild(pop);
  const r=inp.getBoundingClientRect();
  pop.style.left=Math.max(8,r.left)+'px';
  pop.style.top=(r.bottom+scrollY+4)+'px';
  dictPopEl=pop;
  const list=pop.querySelector('.dp-list'), cnt=pop.querySelector('.dp-n');
  const fill=q=>{
    const hits=dictFindAll(q,150);
    cnt.textContent=hits.length+' поз.';
    list.innerHTML=hits.map(h=>`<div class="dp-item" data-t="${h.type}" data-d="${h.i}"
       style="padding:4px 8px;border-bottom:1px solid var(--line);cursor:pointer;display:flex;gap:8px;align-items:center">
       <span class="tag">${h.type}</span><span style="flex:1">${esc(h.d.n)}</span>
       <span>${esc(h.d.u||'')}</span><b class="num" style="min-width:80px;text-align:right">${C.money(h.d.p)}</b></div>`).join('')
       ||'<div style="padding:8px;color:var(--hint)">Ничего не найдено</div>';
  };
  fill('');
  const qi=pop.querySelector('.dp-q'); qi.oninput=()=>fill(qi.value); qi.focus();
  list.onclick=e=>{
    const it=e.target.closest('.dp-item'); if(!it) return;
    const rrow=cur.rows[i];
    rrow.n = store[it.dataset.t==='Материалы'?'dictMat':'dictWork'][+it.dataset.d].n;
    applyDictToRow(rrow,it.dataset.t);              // Тип/ед./цена — сами, по выбору из справочника
    closeDictPop(); recalcAll();
  };
}
function renderMaterials(){
  C.calcMaterials(cur);
  const tb=$('#mat-table tbody'); tb.innerHTML='';
  cur.rows.forEach((row,i)=>{
    const tr=document.createElement('tr');
    tr.dataset.row=i;
    if(selRows.has(i)) tr.classList.add('sel');
    const insBtns=`<button class="btn mini ins" data-ins-before="${i}" title="Вставить строку выше">△＋</button>
                   <button class="btn mini ins" data-ins-after="${i}" title="Вставить строку ниже">▽＋</button>`;
    if(!row.t){ // заголовок раздела / подзаголовок
      tr.className='section-row'+(row.sub?' sub-row':'');
      tr.innerHTML=`<td><input type="checkbox" class="pick" data-pick="${i}" ${selRows.has(i)?'checked':''}></td>
        <td><select class="sec-type" data-i="${i}"><option${row.sub?'':' selected'}>Заголовок</option><option${row.sub?' selected':''}>Подзаголовок</option></select></td>
        <td colspan="7"><input class="sec-name${row.sub?' sub':''}" data-i="${i}" value="${esc(row.n||'')}" placeholder="Название раздела / подзаголовка"></td>
        <td>${insBtns}</td>
        <td><button class="btn mini" data-cp="${i}" title="Копировать группу/строку">⧉</button>
           <button class="btn mini danger" data-rm="${i}">✕</button></td>`;
    } else {
      tr.innerHTML=`<td><input type="checkbox" class="pick" data-pick="${i}" ${selRows.has(i)?'checked':''}></td>
        <td><select data-f="t" data-i="${i}">
            <option${row.t==='Материалы'?' selected':''}>Материалы</option>
            <option${row.t==='Работа'?' selected':''}>Работа</option>
            <option value="__sub">Подзаголовок</option></select></td>
        <td class="name-cell">
          <input data-f="n" data-i="${i}" value="${esc(row.n||'')}" list="dict-list-${row.t==='Работа'?'w':'m'}" placeholder="Наименование — выберите из справочника" title="Начните вводить или нажмите 🔎 — при выборе из справочника Тип, ед.изм. и цена проставятся сами">
          <button class="btn mini dict" data-dict="${i}" title="Выбрать позицию из справочника (Материалы / Работы)">🔎</button></td>
        <td><input data-f="u" data-i="${i}" value="${esc(row.u||'')}"></td>
        <td><input data-f="q" data-i="${i}" class="num" value="${row.q??''}"></td>
        <td><input data-f="p" data-i="${i}" class="num" value="${row.p??''}"></td>
        <td class="num">${C.money(row.sum)}</td><td class="num">${C.money(row.I)}</td><td class="num">${C.money(row.J)}</td>
        <td>${insBtns}</td>
        <td><button class="btn mini" data-cp="${i}" title="Копировать группу/строку">⧉</button>
           <button class="btn mini danger" data-rm="${i}">✕</button></td>`;
    }
    tb.appendChild(tr);
  });
  $('#tot-mat').textContent=C.money(cur.totals.mat);
  $('#tot-work').textContent=C.money(cur.totals.work);
  $('#tot-all').textContent=C.money(cur.totals.all);
  $('#hdr-total').textContent=C.money(cur.totals.all);
  $('#clip-info').textContent=rowClipboard.length?`Буфер: ${rowClipboard.length} стр.`:'';
  updateSelInfo();

  tb.onchange=e=>{                                   // делегирование: устойчиво к перерисовке DOM
    const el=e.target;
    if(el.classList.contains('sec-name')){ cur.rows[+el.dataset.i].n=el.value; return recalcAll(); }
    if(el.classList.contains('sec-type')){ cur.rows[+el.dataset.i].sub=el.value==='Подзаголовок'; return recalcAll(); }
    if(!el.dataset || el.dataset.f===undefined) return;
    const i=+el.dataset.i, r=cur.rows[i], f=el.dataset.f;
    if(f==='t'){ // смена типа; опция «Подзаголовок» превращает строку в подзаголовок
      if(el.value==='__sub'){ r.t=null; r.sub=true; return recalcAll(); }
      r.t=el.value;
    } else r[f]=(f==='q'||f==='p')?C.num(el.value):el.value;
    if(f==='n'){ applyDictToRow(r); }                // Тип/ед./цена проставляются сами
    recalcAll();
  };
  /* Кнопка 🔎 — всплывающий выбор позиции из справочника: Тип, ед.изм. и цена проставятся сами */
  tb.querySelectorAll('[data-dict]').forEach(b=>b.onclick=e=>{
    e.stopPropagation();
    const i=+b.dataset.dict, inp=b.closest('td').querySelector('input[data-f="n"]');
    openDictPop(inp,i);
  });
  /* Удаление — ВСЕГДА с подтверждением (кнопка ✕ на строке) */
  tb.querySelectorAll('[data-rm]').forEach(b=>b.onclick=async()=>{
    const i=+b.dataset.rm; if(rowsInDom()[i]!==i){selRows.clear();focusRow=null;return renderMaterials();}
    if(selRows.size>1 && selRows.has(i)){ // удаляем всю выделенную группу
      if(!await askConfirm(`Удалить выделенные строки (${selRows.size} шт.)?`)) return;
      cur.rows=cur.rows.filter((_,j)=>!selRows.has(j)); selRows.clear(); lastClickedRow=null; focusRow=null;
    } else {
      const nm=String(cur.rows[i].n||'позиция').slice(0,60);
      if(!await askConfirm(`Удалить строку №${i+1}: «${nm}»?`)) return;
      cur.rows.splice(i,1); selRows.clear(); lastClickedRow=null; focusRow=null;
    }
    renumberVols(); recalcAll();
  });
  tb.querySelectorAll('[data-pick]').forEach(cb=>cb.onchange=e=>{
    const i=+e.target.dataset.pick;
    e.target.checked?selRows.add(i):selRows.delete(i);
    lastClickedRow=i; focusRow=i;
    e.target.closest('tr').classList.toggle('sel',e.target.checked);
    updateSelInfo();
  });
  /* Выделение строк:
     - клик по свободной области строки — выделить её;
     - Shift+клик по строке ИЛИ по чекбоксу строки — выделить ДИАПАЗОН подряд
       от последней выбранной строки (как в Excel/проводнике);
     - Ctrl+клик — добавить/убрать строку из выделения */
  function paintSel(){
    $$('#mat-table tbody tr').forEach(x=>{const ri=+x.dataset.row; x.classList.toggle('sel',selRows.has(ri));
      const cb=x.querySelector('.pick'); if(cb) cb.checked=selRows.has(ri);});
    updateSelInfo();
  }
  tb.querySelectorAll('tr[data-row]').forEach(tr=>tr.onclick=e=>{
    if(e.target.matches('input,select,button,textarea')) return;
    const i=+tr.dataset.row;
    if(e.shiftKey && lastClickedRow!=null){
      const a=Math.min(lastClickedRow,i), b=Math.max(lastClickedRow,i);
      for(let j=a;j<=b;j++) selRows.add(j);
    } else if(e.ctrlKey||e.metaKey){
      selRows.has(i)?selRows.delete(i):selRows.add(i); lastClickedRow=i; focusRow=i;
    } else { selRows.clear(); selRows.add(i); lastClickedRow=i; focusRow=i; }
    paintSel();
  });
  /* Shift+клик по чекбоксу — диапазон (не даём браузеру переключать галку штатно) */
  tb.querySelectorAll('[data-pick]').forEach(cb=>cb.addEventListener('click',e=>{
    if(!(e.shiftKey && lastClickedRow!=null)) return;
    e.preventDefault();
    const i=+cb.dataset.pick;
    const a=Math.min(lastClickedRow,i), b=Math.max(lastClickedRow,i);
    for(let j=a;j<=b;j++) selRows.add(j);
    focusRow=b; paintSel();
  }));
  tb.querySelectorAll('[data-cp]').forEach(b=>b.onclick=copyAt);
  tb.querySelectorAll('[data-ins-before]').forEach(b=>b.onclick=()=>insertRow(+b.dataset.insBefore,'above'));
  tb.querySelectorAll('[data-ins-after]').forEach(b=>b.onclick=()=>insertRow(+b.dataset.insAfter,'below'));

  function copyAt(e){
    const i=+e.currentTarget.dataset.cp;
    if(rowsInDom()[i]!==i){selRows.clear();return renderMaterials();}
    let idxs = selRows.size ? [...selRows].sort((a,b)=>a-b) : [i];
    rowClipboard = idxs.map(j=>JSON.parse(JSON.stringify(cur.rows[j])));
    toast(`Скопировано строк: ${rowClipboard.length}${selRows.size?' (группа)':''}`);
    $('#clip-info').textContent=`Буфер: ${rowClipboard.length} стр.`;
  }
  function insertRow(anchor,where){
    if(rowsInDom()[anchor]!==anchor){selRows.clear();renderMaterials();return toast('Таблица обновилась — нажмите ещё раз');}
    pasteAt(where==='above'?anchor:anchor+1);
  }
  autoFitColumns('#mat-table');
}
function updateSelInfo(){
  const c=$('#sel-count'); if(c) c.textContent=selRows.size?`Выделено: ${selRows.size}`:'';
  const b=$('#btn-clear-sel'); if(b) b.classList.toggle('hidden',!selRows.size);
}
function selectedIdxs(){ return [...selRows].sort((a,b)=>a-b); }
/* Удаление — ВСЕГДА с подтверждением (и кнопка ✕ на строке, и тулбар слева) */
async function deleteSelectionWithConfirm(){
  if(!selRows.size){ // ничего не выделено — удаляем строку под курсором выделения (последний клик), тоже с подтверждением
    if(focusRow==null||!cur.rows[focusRow]) return toast('Выделите строки (клик / Shift+клик / чекбокс)');
    const nm=String(cur.rows[focusRow].n||'позиция').slice(0,60);
    if(!await askConfirm(`Удалить строку №${focusRow+1}: «${nm}»?`)) return;
    cur.rows.splice(focusRow,1); selRows.clear(); lastClickedRow=null; focusRow=null;
    renumberVols(); recalcAll(); return;
  }
  if(!await askConfirm(`Удалить выделенные строки (${selRows.size} шт.)?`)) return;
  cur.rows=cur.rows.filter((_,j)=>!selRows.has(j)); selRows.clear(); lastClickedRow=null; focusRow=null;
  renumberVols(); recalcAll();
}
/* тулбар СЛЕВА: вставка/удаление (работают от строки под курсором/выделения) */
function toolbarAnchor(){ return selRows.size? Math.min(...selRows) : (focusRow!=null?focusRow:(lastClickedRow!=null?lastClickedRow:cur.rows.length-1)); }
$('#btn-ins-above').onclick=()=>pasteAt(Math.max(0,toolbarAnchor()));
$('#btn-ins-below').onclick=()=>pasteAt(Math.min(cur.rows.length,toolbarAnchor()+1));
$('#btn-del-rows').onclick=deleteSelectionWithConfirm;
$('#btn-clear-sel').onclick=()=>{selRows.clear();lastClickedRow=null;focusRow=null;updateSelInfo();renderMaterials();};
function pasteAt(at){
  if(rowClipboard.length){                       // вставка из буфера (строка или группа) в любое место
    cur.rows.splice(at,0,...rowClipboard.map(r=>JSON.parse(JSON.stringify(r))));
    toast(`Вставлено строк: ${rowClipboard.length}`);
  } else {
    cur.rows.splice(at,0,{t:'Материалы',n:'',u:'шт',q:0,p:0});
    toast('Вставлена новая строка');
  }
  renumberVols(); recalcAll();
}
$$('.add-row').forEach(b=>b.onclick=()=>{
  const kind=b.dataset.kind;
  if(kind==='header') cur.rows.push({t:null,n:'Новый заголовок',sub:false,u:null,q:null,p:null});
  else if(kind==='sub') cur.rows.push({t:null,n:'Новый подзаголовок',sub:true,u:null,q:null,p:null});
  else cur.rows.push({t:kind,n:'',u:'шт',q:0,p:0});
  renumberVols(); recalcAll();
});
document.getElementById('btn-cut-rows').addEventListener('click',()=>{
  if(!selRows.size) return toast('Выделите строки (клик / Shift+клик / чекбокс)');
  rowClipboard=selectedIdxs().map(j=>JSON.parse(JSON.stringify(cur.rows[j])));
  cur.rows=cur.rows.filter((_,j)=>!selRows.has(j)); selRows.clear(); lastClickedRow=null; renumberVols(); recalcAll();
  toast(`Вырезано строк: ${rowClipboard.length}`); });
$('#btn-copy-rows').onclick=()=>{
  if(!selRows.size) return toast('Выделите строки чекбоксами или нажмите ⧉ на строке');
  const i=Math.min(...selRows);
  if(rowsInDom()[i]!==i){selRows.clear();return renderMaterials();}
  $(`#mat-table [data-cp="${i}"]`)?.click(); };
document.getElementById('btn-paste-rows').addEventListener('click',()=>{
  if(!rowClipboard.length) return toast('Буфер пуст');
  cur.rows.push(...rowClipboard.map(r=>JSON.parse(JSON.stringify(r)))); renumberVols(); recalcAll();
  toast(`Вставлено в конец: ${rowClipboard.length} стр. (или △＋/▽＋ — в любое место)`); });
function renumberVols(){ /* позиции после удаления/добавления сдвигаются — пересчёт заново при рендере */ }

/* ---- Автоширина + ручное изменение ширины столбцов ----
   Ширина колонки = ширине самого длинного содержимого (включая значения в полях
   ввода и заголовки). Реализация через <colgroup><col> + table-layout:fixed,
   поэтому ширина гарантированно применяется и не «плывёт».
   Пользователь может вручную тянуть за правую границу заголовка (col-grip);
   двойной клик по границе — вернуть автоширину. Работает для всех таблиц приложения. */
const MINW=36, MAXW=560;
/* Ширина текста в px текущим шрифтом элемента (или 13px Segoe UI для ячеек) */
function textW(txt,el){
  const probe=textW._p||(textW._p=document.createElement('span'));
  if(!probe.parentNode) document.body.appendChild(probe);
  probe.style.cssText='position:absolute;left:-9999px;top:-9999px;visibility:hidden;white-space:nowrap';
  const cs=el?getComputedStyle(el):null;
  probe.style.fontFamily=cs?cs.fontFamily:'"Segoe UI",Arial,sans-serif';
  probe.style.fontSize=cs?cs.fontSize:'13px';
  probe.style.fontWeight=cs?cs.fontWeight:'400';
  probe.textContent=txt||'';
  return probe.offsetWidth;
}
function measureColWidth(table,ci){
  let max=0;
  /* заголовок — не должен обрезаться никогда */
  const th=table.querySelector('thead tr')?.children[ci];
  if(th) max=Math.max(max,textW(th.textContent.trim(),th)+28);
  [...table.querySelectorAll('tr')].forEach(tr=>{
    if(tr.closest('thead')) return;
    const td=tr.children[ci]; if(!td||td.closest('thead')) return;
    if(td.colSpan>1) return;                       // объединённые ячейки не участвуют в замерах
    /* поля ввода/селекты — по их текущему значению */
    td.querySelectorAll('input,select').forEach(el=>{
      if(el.type==='checkbox') return;
      const add=el.tagName==='SELECT'?30:(el.type==='number'?34:24);
      max=Math.max(max,textW(el.value??'',el)+add);
    });
    /* текст ячейки (в т.ч. кнопок) */
    const t=(td.textContent||'').trim();
    if(t) max=Math.max(max,textW(t,td)+18);
    td.querySelectorAll('button').forEach(b=>{ max=Math.max(max,textW(b.textContent.trim(),b)+22); });
    /* ячейка «Наименование + кнопка 🔎»: поле и кнопка должны помещаться вместе */
    if(td.classList.contains('name-cell')){
      const inp=td.querySelector('input');
      if(inp) max=Math.max(max,textW(inp.value??'',inp)+24+textW('🔎',td.querySelector('.dict'))+30);
    }
  });
  return Math.min(Math.max(max+8,MINW),MAXW);
}
function autoFitColumns(sel){
  const table=document.querySelector(sel); if(!table||!table.querySelector('thead tr')) return;
  const headRow=table.querySelector('thead tr');
  const nCols=headRow.children.length;
  if(!table.colTouched) table.colTouched=Array(nCols).fill(false); // колонки, заданные вручную
  let cg=table.querySelector('colgroup');
  if(!cg||cg.children.length!==nCols){
    if(cg) cg.remove();
    cg=document.createElement('colgroup');
    for(let ci=0;ci<nCols;ci++) cg.appendChild(document.createElement('col'));
    table.insertBefore(cg,table.firstChild);
  }
  const widths=[];
  for(let ci=0;ci<nCols;ci++){
    if(!table.colTouched[ci]) widths[ci]=measureColWidth(table,ci);
    else widths[ci]=parseInt(cg.children[ci].style.width,10)||MINW;
    cg.children[ci].style.width=widths[ci]+'px';
  }
  /* суммарная ширина таблицы = сумма колонок; узкие таблицы растягиваем на всю ширину */
  const sum=widths.reduce((a,b)=>a+b,0);
  const wrap=table.closest('.scroll-wrap,.scroll-x');
  const avail=(wrap?wrap.clientWidth:0)||sum;
  table.style.tableLayout='fixed';
  table.style.width=Math.max(sum,avail)+'px';
  table.style.maxWidth='none';
  ensureResizers(table,cg,nCols);
}
function ensureResizers(table,cg,nCols){
  if(table._resized) return; table._resized=true;
  table.querySelectorAll('thead th').forEach((th,idx)=>{
    th.style.position='relative';
    const g=document.createElement('div'); g.className='col-grip'; g.title='Потяните — изменить ширину; двойной клик — автоширина';
    th.appendChild(g);
    g.addEventListener('mousedown',e=>{
      e.preventDefault(); e.stopPropagation();
      const startX=e.clientX, startW=th.getBoundingClientRect().width;
      document.body.style.cursor='col-resize';
      const mv=ev=>{
        const w=Math.max(MINW,startW+ev.clientX-startX);
        table.colTouched[idx]=true;
        cg.children[idx].style.width=w+'px';
        let s=0; for(const c of cg.children) s+=parseInt(c.style.width,10)||MINW;
        const wrap=table.closest('.scroll-wrap,.scroll-x');
        table.style.width=Math.max(s,(wrap?wrap.clientWidth:0)||s)+'px';
      };
      const up=()=>{document.removeEventListener('mousemove',mv);document.removeEventListener('mouseup',up);document.body.style.cursor='';};
      document.addEventListener('mousemove',mv); document.addEventListener('mouseup',up);
    });
    g.addEventListener('dblclick',e=>{ e.stopPropagation();
      table.colTouched[idx]=false;
      autoFitColumns('#'+table.id);
    });
  });
}


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
  autoFitColumns('#smeta-table');
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
    } else {                       // смена месяца / года в шапке (только в «Закрытии»)
      const mm=cur.meta.mmyy[i];
      const m=t.hasAttribute('data-mm')?+t.value:mm.m;
      const y=t.hasAttribute('data-yy')?+t.value:mm.y;
      C.setMonthYear(cur,i,m,y);   // при смене МЕСЯЦА последующие столбцы сдвигаются авто по порядку
      renderKs6();                 // перерисовать шапку: каскад виден сразу, без потери фокуса на году
      C.calcKS3(cur); renderSmeta(); renderObems(); save();
      return;
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
  autoFitColumns('#ks6-table');
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
function switchSheet(k){curSheet=k;$$('.sheet-tab').forEach(b=>b.classList.toggle('active',b.dataset.sheet===k));renderSheet(k);}
/* печать активной таблицы (Материалы/Смета/КС-6) — отдельная кнопка */
const PRINT_MAP={materials:'#panel-materials',smeta:'#panel-smeta',ks6:'#panel-ks6',ks2:'#panel-ks2',ks3:'#panel-ks3',export:'#panel-export'};
function printCurrent(){printArea(PRINT_MAP[curSheet]||'#panel-materials');}
$('#btn-print-current') && ($('#btn-print-current').onclick=printCurrent);
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
  autoFitColumns('#r2-table');
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
  otb.querySelectorAll('[data-del-o]').forEach(b=>b.onclick=async()=>{
    const o=store.our[+b.dataset.delO];
    if(!await askConfirm(`Удалить организацию «${o.list||o.name}»?`)) return;
    store.our.splice(+b.dataset.delO,1);save();renderKontr();});
  autoFitColumns('#our-table');
  const ctb=$('#cust-table tbody');ctb.innerHTML='';
  store.customers.forEach((o,i)=>{
    const tr=document.createElement('tr');
    tr.innerHTML=['list','name','inn','kpp','okpo','addr','sign'].map(k=>`<td><input data-k="${k}" data-i="${i}" value="${esc(o[k]||'')}"></td>`).join('')+`<td><button class="btn mini danger" data-del-c="${i}">✕</button></td>`;
    ctb.appendChild(tr);
  });
  ctb.querySelectorAll('input').forEach(inp=>inp.onchange=e=>{store.customers[+e.target.dataset.i][e.target.dataset.k]=e.target.value;save();});
  ctb.querySelectorAll('[data-del-c]').forEach(b=>b.onclick=async()=>{
    const c=store.customers[+b.dataset.delC];
    if(!await askConfirm(`Удалить заказчика «${c.list||c.name}»?`)) return;
    store.customers.splice(+b.dataset.delC,1);save();renderKontr();});
  autoFitColumns('#cust-table');
}
$('#btn-our-add').onclick=()=>{store.our.push({list:'Новая',name:'',inn:'',kpp:'',okpo:'',addr:'',sign:'',post:'',fio:'',nds:'НДС 22%'});save();renderKontr();};
$('#btn-cust-add').onclick=()=>{store.customers.push({list:'Новый заказчик',name:'',inn:'',kpp:'',okpo:'',addr:'',sign:''});save();renderKontr();};

/* ================= СПРАВОЧНИК МАТЕРИАЛОВ И РАБОТ =================
   Отдельная вкладка: две редактируемые таблицы — Материалы и Работы.
   Используется редактором заказа: при вводе наименования строки автоподстановка
   ед.изм. и цены из справочника (dictLookup). Изначально заполняется из данных
   исходной Excel-книги (1178 строк), далее полностью под контролем пользователя. */
let spravFilter={m:'',w:''};
function renderSprav(){
  const defs=[['m','dictMat','Материалы'],['w','dictWork','Работы']];
  defs.forEach(([s,key,label])=>{
    const q=(spravFilter[s]||'').toLowerCase();
    const tb=$('#sprav-'+s+'-tbody'); if(!tb) return; tb.innerHTML='';
    let shown=0;
    store[key].forEach((d,i)=>{
      if(q && !d.n.toLowerCase().includes(q)) return;
      shown++;
      const tr=document.createElement('tr');
      tr.innerHTML=`<td>${i+1}</td>
        <td><input data-k="n" data-i="${i}" value="${esc(d.n)}"></td>
        <td><input data-k="u" data-i="${i}" value="${esc(d.u||'')}" size="6"></td>
        <td><input data-k="p" data-i="${i}" class="num" value="${d.p??''}"></td>
        <td><button class="btn mini" data-addorder="${s}" data-i="${i}" title="Добавить в открытый заказ">→ Заказ</button>
            <button class="btn mini danger" data-del="${i}">✕</button></td>`;
      tb.appendChild(tr);
    });
    $('#sprav-'+s+'-count').textContent=`${shown} / ${store[key].length}`;
    autoFitColumns('#sprav-'+s+'-table');
    tb.querySelectorAll('input').forEach(inp=>inp.onchange=e=>{
      const d=store[key][+e.target.dataset.i];
      d[e.target.dataset.k]=(e.target.dataset.k==='p')?C.num(e.target.value):e.target.value;
      save();
    });
    tb.querySelectorAll('[data-del]').forEach(b=>b.onclick=async()=>{
      const d=store[key][+b.dataset.del];
      if(!await askConfirm(`Удалить из справочника: «${String(d.n).slice(0,80)}»?`)) return;
      store[key].splice(+b.dataset.del,1);save();renderSprav();});
    tb.querySelectorAll('[data-addorder]').forEach(b=>b.onclick=()=>{
      if(!cur) return toast('Сначала откройте заказ (вкладка Заказы → Открыть)');
      const d=store[key][+b.dataset.i];
      cur.rows.push({t:s==='m'?'Материалы':'Работа',n:d.n,u:d.u,q:1,p:d.p});
      recalcAll(); switchSheet('materials'); toast(`Добавлено в заказ: ${String(d.n).slice(0,50)}`);
    });
  });
}
$('#btn-sprav-add-m').onclick=()=>{store.dictMat.unshift({n:'Новый материал',u:'шт',p:0});save();renderSprav();};
$('#btn-sprav-add-w').onclick=()=>{store.dictWork.unshift({n:'Новая работа',u:'шт',p:0});save();renderSprav();};
$('#sprav-search-m').oninput=e=>{spravFilter.m=e.target.value.toLowerCase();renderSprav();};
$('#sprav-search-w').oninput=e=>{spravFilter.w=e.target.value.toLowerCase();renderSprav();};
$('#btn-sprav-reset').onclick=()=>{
  if(!confirm('Перезаполнить справочники из данных исходной Excel-книги? Ваши правки будут потеряны.'))return;
  const d=buildDictsFromSeed(); store.dictMat=d.mat; store.dictWork=d.work; save(); renderSprav();
  toast('Справочники восстановлены из оригинала');
};

/* ================= utils ================= */
function esc(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
/* datalist для автоподстановки наименований в редакторе (первые 300 записей —
   чтобы не тормозить DOM; полный поиск — через dictLookup по точному совпадению) */
function fillDatalists(){
  const dl=m=>m.slice(0,300).map(d=>`<option value="${esc(d.n)}">`).join('');
  $('#dict-list-m').innerHTML=dl(store.dictMat);
  $('#dict-list-w').innerHTML=dl(store.dictWork);
}
function toast(t){const el=$('#toast');el.textContent=t;el.classList.remove('hidden');setTimeout(()=>el.classList.add('hidden'),2500);}

/* ================= ТЕМЫ (светлая/тёмная/синяя/сепия) ================= */
const THEMES={light:'Светлая',dark:'Тёмная',blue:'Синяя',sepia:'Сепия'};
function applyTheme(t){
  document.body.dataset.theme=t;
  const sel=$('#theme-select'); if(sel) sel.value=t;
  try{localStorage.setItem('mes_theme',t);}catch(e){}
}
$('#theme-select').onchange=e=>applyTheme(e.target.value);
let savedTheme='light'; try{savedTheme=localStorage.getItem('mes_theme')||'light';}catch(e){}
if(!THEMES[savedTheme]) savedTheme='light';
applyTheme(savedTheme);

/* offset «прилипания» шапок таблиц = реальная высота верхней панели (обновляется при ресайзе),
   чтобы шапки никогда не скрывались при прокрутке */
function updateTopOffset(){
  const tb=document.querySelector('.topbar'); if(!tb) return;
  document.documentElement.style.setProperty('--top-h',tb.offsetHeight+'px');
}
window.addEventListener('resize',updateTopOffset);
updateTopOffset();

loadStore(); renderOrders(); fillDatalists();
/* отладка/тесты: доступ к внутреннему состоянию */
window.__APP={get cur(){return cur},set cur(v){cur=v},store,rowClipboard:()=>rowClipboard,selRows};
})();
