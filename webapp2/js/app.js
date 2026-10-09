/* app.js — управление UI: заказы (localStorage), редактор листов, реестры, печать */
(function(){
const C=window.Calc, $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const LS='mest_webapp2';
/* Подтверждение удаления: собственный диалог с контрастными темами;
   если window.confirm переопределён тестами — используем его. */
function askConfirm(msg){
  if(typeof window.__TEST_CONFIRM==='function') return window.__TEST_CONFIRM(msg);
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
  invalidateDictIdx();
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
let _saveT=null;
function saveNow(){
  /* лёгкий кэш сумм открытого заказа для списка заказов (полный каскад здесь не гоняется) */
  if(cur){ try{ C.calcMaterials(cur); cur._tot={all:cur.totals.all, ks2:(cur.ks2&&cur.ks2.total)||0}; }catch(e){} }
  localStorage.setItem(LS,JSON.stringify(store));
}
/* сериализация+запись ~1 МБ JSON на КАЖДОЕ изменение ячейки подтормаживала UI —
   автосохранение дебаунсим; критические места (печать/экспорт/закрытие) зовут saveNow() */
function save(){ clearTimeout(_saveT); _saveT=setTimeout(saveNow,400); }
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
    /* суммы берём из сохранённого кэша расчёта (order._tot) — полный пересчёт
       КАЖДОГО заказа при каждой отрисовке списка был главным источником тормозов */
    const t=o._tot||{};
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${i+1}</td><td>${esc(o.meta.objectName)}</td><td>${esc(o.meta.org)}</td>
      <td>${esc(o.meta.customerList)}</td><td>${o.meta.dogovorDate||''}</td>
      <td class="num">${C.money(t.all||0)}</td><td class="num">${C.money(t.ks2||0)}</td>
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
/* поиск по списку заказов — с дебаунсом (не перерисовывать на каждый символ) */
let _ordT=null; $('#order-search').oninput=()=>{clearTimeout(_ordT);_ordT=setTimeout(renderOrders,150);};
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
/* Лёгкий пересчёт: только видимый лист (Материалы/Смета/КС-6) + автосохранение.
   Раньше на каждое изменение ячейки гнался весь каскад КС-3→КС-2→Объемы→Смета и
   перерисовывались все панели — на 1178 строках это «вешало» интерфейс. */
function recalcAll(){ if(!cur)return;
  const vis=['materials','smeta','ks6'].find(n=>{const el=document.getElementById('panel-'+n);return el&&!el.classList.contains('hidden');});
  if(vis) renderSheet(vis); else C.calcMaterials(cur);
  save();
}
/* Полный каскад — только когда он реально нужен: печать/экспорт/переключение на
   листы КС-2/КС-3/Экспорт (см. renderSheet). */
function fullRecalc(){ if(cur) C.calcKS3(cur); }
function renderCurrent(){ renderSheet($('.sheet-tab.active').dataset.sheet||'materials'); }

/* ================= ЛИСТЫ ================= */
function renderSheet(name){
  if(!cur) return;
  if(name==='ks2'||name==='ks3'||name==='export') fullRecalc();   // тяжёлые документы — считаем по требованию
  if(name==='materials') renderMaterials();
  if(name==='smeta') renderSmeta();
  if(name==='ks6'){ renderKs6(); renderObems(); }   // объемы закрытия — внутри панели КС-6, без дублей
  if(name==='ks2') renderKs2Doc();
  if(name==='ks3') renderKs3Doc();
  if(name==='export') renderExport();
}

/* ---- Материалы (редактируемая таблица) ----
   Возможности:
   - при выборе позиции из справочника (datalist) или точном вводе наименования
     Тип / Ед.изм. / Цена проставляются САМИ (applyDictToRow);
   - подзаголовки при вводе строки; выбор типа «Подзаголовок» превращает строку в подзаголовок;
   - вставка строки МЕЖДУ существующими (кнопки слева △＋ Выше / ▽＋ Ниже и на любой строке);
   - копирование строки и группы строк (чекбокс + Shift+клик по диапазону), вставка в любое место;
   - удаление ВСЕГДА с подтверждением (confirm);
   - шапка таблицы прикреплена (sticky), ширина столбцов — авто по содержимому
     и перетаскивание мышью за границу заголовка (двойной клик — снова авто). */
let rowClipboard=[];   // буфер копирования строк заказа
let selRows=new Set(); // выделенные строки (для копирования группы)
let lastClickedRow=null; // для выделения диапазона через Shift
/* data-i в DOM всегда = актуальный индекс в cur.rows, т.к. таблица перерисовывается
   после любой вставки/удаления. rowsInDom() — страховка от устаревшего DOM. */
function rowsInDom(){ return $$('#mat-table tbody tr').map(tr=>{
  const el=tr.querySelector('[data-i]'); return el?+el.dataset.i:-1; }).filter(i=>i>=0); }
function dictLookup(name,t){
  const k=(name||'').trim().toLowerCase(); if(!k) return null;
  const ix=getDictIdx();
  if(t==='Работа') return ix.w.get(k)||null;
  if(t==='Материалы') return ix.m.get(k)||null;
  return ix.m.get(k)||ix.w.get(k)||null;
}
/* Индексы справочников по нижнему регистру наименования — O(1) вместо полного
   перебора массивов (в т.ч. concat двух массивов на каждую отредактированную строку). */
const dictIdx={m:null,w:null};
function getDictIdx(){
  if(!dictIdx.m||!dictIdx.w){
    dictIdx.m=new Map(store.dictMat.map(d=>[d.n.trim().toLowerCase(),d]));
    dictIdx.w=new Map(store.dictWork.map(d=>[d.n.trim().toLowerCase(),d]));
  }
  return dictIdx;
}
function invalidateDictIdx(){ dictIdx.m=dictIdx.w=null; }
/* Автоопределение ТИПА по справочнику: наименование есть только в работах — «Работа»,
   только в материалах — «Материалы»; если не найдено — оставляем текущий тип. */
function dictTypeOf(k){
  const ix=getDictIdx();
  const inM=ix.m.has(k), inW=ix.w.has(k);
  if(inW&&!inM) return 'Работа';
  if(inM&&!inW) return 'Материалы';
  return null;
}
function applyDictToRow(r){
  const k=(r.n||'').trim().toLowerCase(); if(!k) return false;
  const t=dictTypeOf(k); if(t) r.t=t;
  const d=dictLookup(r.n,r.t)||dictLookup(r.n,null);
  if(d){ r.u=d.u; r.p=d.p; toast(`Справочник: ${r.t}, ед. ${d.u}, цена ${C.money(d.p)} — проставлено автоматически`); return true; }
  return false;
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
        <td><input data-f="n" data-i="${i}" value="${esc(row.n||'')}" list="dict-list-${row.t==='Работа'?'w':'m'}" title="Выберите позицию из справочника — Тип, ед.изм. и цена проставятся сами"></td>
        <td><input data-f="u" data-i="${i}" value="${esc(row.u||'')}" size="4"></td>
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
  /* Удаление — ВСЕГДА с подтверждением */
  tb.querySelectorAll('[data-rm]').forEach(b=>b.onclick=async()=>{
    const i=+b.dataset.rm; if(rowsInDom()[i]!==i){selRows.clear();return renderMaterials();}
    if(selRows.size>1 && selRows.has(i)){ // удаляем всю выделенную группу
      if(!await askConfirm(`Удалить выделенные строки (${selRows.size} шт.)?`)) return;
      cur.rows=cur.rows.filter((_,j)=>!selRows.has(j)); selRows.clear(); lastClickedRow=null;
    } else {
      const nm=String(cur.rows[i].n||'позиция').slice(0,60);
      if(!await askConfirm(`Удалить строку №${i+1}: «${nm}»?`)) return;
      cur.rows.splice(i,1); selRows.clear(); lastClickedRow=null;
    }
    renumberVols(); recalcAll();
  });
  tb.querySelectorAll('[data-pick]').forEach(cb=>cb.onchange=e=>{
    const i=+e.target.dataset.pick;
    e.target.checked?selRows.add(i):selRows.delete(i);
    lastClickedRow=i;
    e.target.closest('tr').classList.toggle('sel',e.target.checked);
    updateSelInfo();
  });
  /* Клик по строке + Shift+клик — выделение ДИАПАЗОНА подряд (как в проводнике);
     Ctrl+клик — добавить/убрать строку из выделения */
  tb.querySelectorAll('tr[data-row]').forEach(tr=>tr.onclick=e=>{
    if(e.target.matches('input,select,button,textarea')) return;
    const i=+tr.dataset.row;
    if(e.shiftKey && lastClickedRow!=null){
      const a=Math.min(lastClickedRow,i), b=Math.max(lastClickedRow,i);
      for(let j=a;j<=b;j++) selRows.add(j);
    } else if(e.ctrlKey||e.metaKey){
      selRows.has(i)?selRows.delete(i):selRows.add(i); lastClickedRow=i;
    } else { selRows.clear(); selRows.add(i); lastClickedRow=i; }
    $$('#mat-table tbody tr').forEach(x=>{const ri=+x.dataset.row; x.classList.toggle('sel',selRows.has(ri));
      const cb=x.querySelector('.pick'); if(cb) cb.checked=selRows.has(ri);});
    updateSelInfo();
  });
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
async function deleteSelectionWithConfirm(){
  if(!selRows.size) return toast('Выделите строки (клик / Shift+клик / чекбокс)');
  if(!await askConfirm(`Удалить выделенные строки (${selRows.size} шт.)?`)) return;
  cur.rows=cur.rows.filter((_,j)=>!selRows.has(j)); selRows.clear(); lastClickedRow=null;
  renumberVols(); recalcAll();
}
/* тулбар слева: вставка/удаление */
function toolbarAnchor(){ return selRows.size? Math.min(...selRows) : (lastClickedRow!=null?lastClickedRow:cur.rows.length-1); }
$('#btn-ins-above').onclick=()=>pasteAt(Math.max(0,toolbarAnchor()));
$('#btn-ins-below').onclick=()=>pasteAt(Math.min(cur.rows.length,toolbarAnchor()+1));
$('#btn-del-rows').onclick=deleteSelectionWithConfirm;
$('#btn-clear-sel').onclick=()=>{selRows.clear();lastClickedRow=null;updateSelInfo();renderMaterials();};
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
   Колонки получают ширину по содержимому (autoFitColumns), затем пользователь
   может тянуть за правую границу <th> (col-grip) или двойным кликом по границе
   вернуть автоширину. Работает для всех таблиц приложения.

   ПРОИЗВОДИТЕЛЬНОСТЬ: раньше ширина каждой колонки измерялась перебором ВСЕХ
   строк таблицы с созданием скрытого <span> на КАЖДУЮ ячейку и вставкой его в
   document.body (cols × rows перерасчётов макета — «Материалы» открывались
   очень долго). Теперь измерения выполняются ОДНИМ пакетом: все probe-элементы
   собираются во fragment, одна вставка в DOM, одно чтение offsetWidth, один
   removeChild. Дополнительно: ограничиваем число измеряемых строк (первые
   MAX_MEASURE_ROWS) и кэшируем результат измерения колонок таблицы.

   ПОВЕДЕНИЕ ПРИ ИЗМЕНЕНИИ ШИРИНЫ: таблица больше НЕ растягивается на всю ширину
   окна (width:auto вместо width:100%), а остальные столбцы НЕ подстраиваются
   (table-layout:fixed) — тянем один столбец, все остальные остаются как есть. */
const MAX_MEASURE_ROWS=60;                    // достаточно для оценки автоширины
function measureColWidths(table){
  const headRow=table.querySelector('thead tr'); if(!headRow) return [];
  const nCols=headRow.children.length;
  /* границы колонок с учётом colspan (заголовок раздела в «Материалах» = colspan 7) */
  const colCells=[];                          // colCells[ci] = [{cell,isHead}]
  for(let ci=0;ci<nCols;ci++) colCells.push([]);
  const rows=[...table.querySelectorAll('tr')].slice(0,MAX_MEASURE_ROWS);
  rows.forEach(tr=>{
    let ci=0;
    [...tr.children].forEach(cell=>{
      const span=Math.max(1,+cell.getAttribute('colspan')||1);
      for(let k=0;k<span&&ci<nCols;k++,ci++){
        if(colCells[ci].length<MAX_MEASURE_ROWS) colCells[ci].push({cell,isHead:!!cell.closest('thead')});
      }
    });
  });
  /* один оффскрин-контейнер, одна вставка в документ, одно чтение ширин */
  const host=document.createElement('div');
  host.style.cssText='position:absolute;left:-99999px;top:0;visibility:hidden;white-space:nowrap;font-family:"Segoe UI",Arial,sans-serif';
  const compact=table.classList.contains('compact');
  const groups=[];                            // группы span'ов по колонкам (flex-колонка)
  for(let ci=0;ci<nCols;ci++){
    const g=document.createElement('div');
    g.style.cssText='display:flex;flex-direction:column;align-items:flex-start;width:max-content';
    colCells[ci].forEach(({cell,isHead})=>{
      const inp=cell.querySelector('input,select');
      const text=((inp&&inp.value)?inp.value:(cell.textContent||'')).trim().slice(0,200);
      const span=document.createElement('span');
      span.style.cssText='box-sizing:border-box;white-space:nowrap;padding:'+(compact?'2px 6px':'4px 8px')+
        ';font-size:'+(compact?'12px':'13px')+';font-weight:'+(isHead?'600':'400');
      span.textContent=text;
      g.appendChild(span);
    });
    groups.push(g); host.appendChild(g);
  }
  document.body.appendChild(host);            // единственная вставка в DOM
  const widths=[];
  for(let ci=0;ci<nCols;ci++){
    let max=0;
    groups[ci].childNodes.forEach(s=>{const w=s.offsetWidth;if(w>max)max=w;});
    widths[ci]=Math.min(Math.max(max+6,42),520);
  }
  host.remove();
  return widths;
}
function autoFitColumns(sel){
  const table=document.querySelector(sel); if(!table||!table.querySelector('thead tr')) return;
  const headRow=table.querySelector('thead tr');
  const nCols=headRow.children.length;
  /* сброс кэша при изменении числа строк/столбцов (таблица перерисована) */
  const sig=nCols+'|'+table.querySelectorAll('tbody tr').length;
  if(table._fitSig!==sig){ table._fitCache=null; table._fitSig=sig; }
  if(!table._fitCache) table._fitCache=measureColWidths(table);
  for(let ci=0;ci<nCols;ci++){
    const th=headRow.children[ci];
    if(th._userWidth) continue;               // не трогаем колонки, которые менял пользователь
    th.style.width=table._fitCache[ci]+'px';
  }
  ensureResizers(table);
}
function ensureResizers(table){
  if(table._resized) return; table._resized=true;
  table.querySelectorAll('thead th').forEach(th=>{
    const g=document.createElement('div'); g.className='col-grip';
    th.appendChild(g);
    g.addEventListener('mousedown',e=>{
      e.preventDefault(); e.stopPropagation();
      const startX=e.clientX, startW=th.offsetWidth;
      const mv=ev=>{ th.style.width=Math.max(36,startW+ev.clientX-startX)+'px'; th._userWidth=true; };
      const up=()=>{document.removeEventListener('mousemove',mv);document.removeEventListener('mouseup',up);};
      document.addEventListener('mousemove',mv); document.addEventListener('mouseup',up);
    });
    g.addEventListener('dblclick',e=>{ e.stopPropagation();
      table._fitCache=null;                   // содержимое могло измениться — перемерить
      const ci=[...th.parentNode.children].indexOf(th);
      if(!table._fitCache) table._fitCache=measureColWidths(table);
      th._userWidth=false; th.style.width=table._fitCache[ci]+'px';
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
      C.calcSmeta(cur); renderSmeta(); renderObems(); save();
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
    /* только Смета + Объемы: полный каскад КС-2/КС-3 — по требованию (печать/экспорт/их листы) */
    C.calcSmeta(cur); renderSmeta(); renderObems(); save();
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
$('#btn-save-order').onclick=()=>{saveNow();toast('Заказ сохранён в браузере (localStorage)');};
$('#btn-export-reestr').onclick=()=>{
  fullRecalc(); if(cur.ks2&&cur._tot) cur._tot.ks2=cur.ks2.total;
  const row=C.exportRow(cur);
  const ex=store.r2.find(r=>r.object===row.object);
  if(ex){Object.assign(ex,row);toast('Запись обновлена в Реестре 2 (аналог «Переписать?» → Да)');}
  else{store.r2.push(row);toast('Новая запись добавлена в Реестр 2');}
  saveNow(); renderR2();
};
$('#btn-print-ks2').onclick=()=>{fullRecalc();if(cur._tot&&cur.ks2)cur._tot.ks2=cur.ks2.total;save();printDoc('ks2-doc');};
$('#btn-print-ks3').onclick=()=>{fullRecalc();save();printDoc('ks3-doc');};
$('#btn-print-ks6').onclick=()=>{renderKs6();renderObems();printArea('#panel-ks6');};
$('#btn-download-zip').onclick=()=>{
  const blob=new Blob([JSON.stringify(cur,null,1)],{type:'application/json'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download='Заказ_'+(cur.meta.objectName||'new')+'.json';a.click();
};
$('#btn-close-order').onclick=()=>{saveNow();$('#view-editor').classList.add('hidden');$('#view-orders').classList.remove('hidden');renderOrders();};
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
      invalidateDictIdx(); save();
    });
    tb.querySelectorAll('[data-del]').forEach(b=>b.onclick=async()=>{
      const d=store[key][+b.dataset.del];
      if(!await askConfirm(`Удалить из справочника: «${String(d.n).slice(0,80)}»?`)) return;
      store[key].splice(+b.dataset.del,1);invalidateDictIdx();save();renderSprav();});
    tb.querySelectorAll('[data-addorder]').forEach(b=>b.onclick=()=>{
      if(!cur) return toast('Сначала откройте заказ (вкладка Заказы → Открыть)');
      const d=store[key][+b.dataset.i];
      cur.rows.push({t:s==='m'?'Материалы':'Работа',n:d.n,u:d.u,q:1,p:d.p});
      recalcAll(); switchSheet('materials'); toast(`Добавлено в заказ: ${String(d.n).slice(0,50)}`);
    });
  });
}
$('#btn-sprav-add-m').onclick=()=>{store.dictMat.unshift({n:'Новый материал',u:'шт',p:0});invalidateDictIdx();save();renderSprav();};
$('#btn-sprav-add-w').onclick=()=>{store.dictWork.unshift({n:'Новая работа',u:'шт',p:0});invalidateDictIdx();save();renderSprav();};
$('#sprav-search-m').oninput=e=>{spravFilter.m=e.target.value.toLowerCase();renderSprav();};
$('#sprav-search-w').oninput=e=>{spravFilter.w=e.target.value.toLowerCase();renderSprav();};
$('#btn-sprav-reset').onclick=()=>{
  if(!confirm('Перезаполнить справочники из данных исходной Excel-книги? Ваши правки будут потеряны.'))return;
  const d=buildDictsFromSeed(); store.dictMat=d.mat; store.dictWork=d.work; invalidateDictIdx(); save(); renderSprav();
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

loadStore(); renderOrders(); fillDatalists();
/* отладка/тесты: доступ к внутреннему состоянию */
window.__APP={get cur(){return cur},set cur(v){cur=v},store,rowClipboard:()=>rowClipboard,selRows};
})();
