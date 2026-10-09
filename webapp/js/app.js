/* ============================================================
   WEB-копия связки Excel-программ:
   1) "Заказ_...xlsm"  — Расчет и Документация (листы: Материалы, Смета,
      КС-6, КС-2, КС-3, Контрагенты, Экспорт + макросы Module1/2/4/5/6, Лист2/5...)
   2) "Реестр 2.xlsm"  — лист РасчетДокументация (база заказов)
   3) "Реестр ЗАКАЗОВ 2.xlsm" — ПереченьЗаказов (внешние ссылки на Реестр 2,
      макросы: двойной клик = открыть заказ, ↑/↓ = перемещение заказа)

   Здесь VBA-логика и формулы воспроизведены на JavaScript.
   Хранение — localStorage (аналог файлов на сетевом диске \\Nas).
   ============================================================ */
'use strict';

/* ---------- Утилиты: аналоги Excel-функций ---------- */
const num = v => { const n = parseFloat(String(v).replace(',', '.')); return isNaN(n) ? 0 : n; };
const round = (v, d = 2) => { const m = Math.pow(10, d); return Math.round((v + Number.EPSILON * 10) * m) / m; };
// Excel ROUNDUP — «от нуля» вверх
const roundup = (v, d = 2) => { const m = Math.pow(10, d); return (v >= 0 ? Math.ceil(v * m - 1e-9) : Math.floor(v * m + 1e-9)) / m; };
const fmt = (v, d = 2) => v == null || isNaN(v) ? '' : Number(v).toLocaleString('ru-RU', { minimumFractionDigits: d, maximumFractionDigits: d });
const fmt0 = v => fmt(v, 0);
const eomonth = dt => new Date(dt.getFullYear(), dt.getMonth() + 1, 0);
const addDays = (dt, n) => { const d = new Date(dt); d.setDate(d.getDate() + n); return d; };
const iso = dt => dt ? (dt instanceof Date ? dt.toISOString().slice(0, 10) : String(dt).slice(0, 10)) : '';
const ruDate = s => { if (!s) return ''; const d = new Date(s); return isNaN(d) ? String(s) : d.toLocaleDateString('ru-RU'); };
const el = (tag, attrs = {}, ...kids) => {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k.startsWith('on')) e.addEventListener(k.slice(2), v); else if (k === 'class') e.className = v; else e.setAttribute(k, v);
  }
  for (const k of kids.flat()) e.append(k instanceof Node ? k : document.createTextNode(k ?? ''));
  return e;
};
function toast(msg, ms = 3500) {
  const t = document.getElementById('toast');
  t.textContent = msg; t.classList.remove('hidden');
  clearTimeout(t._h); t._h = setTimeout(() => t.classList.add('hidden'), ms);
}

/* ---------- Числительное прописью (аналог именованных диапазонов
   n_1..n_5, n0x, n1x, тыс, мил, послПользователь из книги «Заказ») ---------- */
function plural(n, one, few, many) { n = Math.abs(n) % 100; const d = n % 10;
  if (n > 10 && n < 20) return many; if (d > 1 && d < 5) return few; if (d === 1) return one; return many; }
function numWords(amount, currency = ['рубль', 'рубля', 'рублей'], kop = ['копейка', 'копейки', 'копеек']) {
  const units = ['', 'один ', 'два ', 'три ', 'четыре ', 'пять ', 'шесть ', 'семь ', 'восемь ', 'девять '];
  const teens = ['десять ', 'одиннадцать ', 'двенадцать ', 'тринадцать ', 'четырнадцать ', 'пятнадцать ', 'шестнадцать ', 'семнадцать ', 'восемнадцать ', 'девятнадцать '];
  const tens = ['', '', 'двадцать ', 'тридцать ', 'сорок ', 'пятьдесят ', 'шестьдесят ', 'семьдесят ', 'восемьдесят ', 'девяносто '];
  const hundreds = ['', 'сто ', 'двести ', 'триста ', 'четыреста ', 'пятьсот ', 'шестьсот ', 'семьсот ', 'восемьсот ', 'девятьсот '];
  function trip(n, fem) {
    let s = hundreds[Math.floor(n / 100)] || '';
    n %= 100;
    if (n >= 10 && n < 20) s += teens[n - 10];
    else { s += tens[Math.floor(n / 10)] || ''; s += (fem && n % 10 === 1 ? 'одна ' : fem && n % 10 === 2 ? 'две ' : units[n % 10] || ''); }
    return s.trim();
  }
  let rub = Math.floor(amount), kopV = Math.round((amount - rub) * 100);
  if (kopV === 100) { rub++; kopV = 0; }
  const parts = [];
  const classes = [[Math.floor(rub / 1e9), ['миллиард', 'миллиарда', 'миллиардов'], false],
                   [Math.floor(rub / 1e6) % 1000, ['миллион', 'миллиона', 'миллионов'], false],
                   [Math.floor(rub / 1e3) % 1000, ['тысяча', 'тысячи', 'тысяч'], true]];
  let out = '';
  for (const [v, forms, fem] of classes) if (v) out += trip(v, fem) + ' ' + plural(v, ...forms) + ' ';
  out += (rub % 1000 ? trip(rub % 1000, false) + ' ' + plural(rub % 1000, ...currency) + ' ' : (rub ? '' : plural(0, ...currency) + ' '));
  out = out.trim().replace(/^./, c => c.toUpperCase());
  return `${out} ${String(kopV).padStart(2, '0')} ${plural(kopV, ...kop)}`;
}

/* ---------- Справочник контрагентов (лист «Контрагенты») ----------
   Индексы колонок исходного листа A..T (0..19):
   1(B)=ключ списка, 2(C)=наименование, 3(D)=ИНН, 4(E)=КПП, 5(F)=ОКПО,
   6(G)=адрес, 7(H)=наименование в подписи, 8(I)=должность, 9(J)=ФИО,
   10(K)=НДС, 13(N)=строка для КС-6, 14(O)=подпись */
let KONS = [];
function konsByListKey(key) { return KONS.find(k => String(k[1]) === String(key)); }        // B10:T78 → col B
function ourOrgs() { return KONS.slice(0, 5).filter(k => k[1] && k[1] !== 0); }              // B3:T7 — наши организации
function customerList() { return KONS.filter(k => k[1] && k[1] !== 0 && !ourOrgs().includes(k)); }

/* ---------- Модель состояния ---------- */
const LS_ORDER = 'webzakaz.order.v1', LS_REESTR = 'webzakaz.reestr.v1';

function defaultOrder() {
  return {
    objectName: ORDER_META.objectName, org: ORDER_META.org, version: ORDER_META.version,
    coefMat: num(ORDER_META.coefMat), coefWork: num(ORDER_META.coefWork),
    ks2Date: ORDER_META.ks2date, manualFrom: ORDER_META.manual_from, manualTo: ORDER_META.manual_to,
    customer: ORDER_META.customerDropdown, stroyka: ORDER_META.stroyka, obekt: ORDER_META.obekt,
    predmet: KS6_PREDMET, dogovorNum: ORDER_META.dogovorNum, dogovorDate: ORDER_META.dogovorDate,
    periods: ORDER_META.ks6_dates.map((d, i) => ({ date: d, status: ORDER_META.ks6_status[i] || '----' })),
    rows: MATERIALS_RAW.map(m => ({ type: m.t, name: m.n, unit: m.u, qty: num(m.q), price: num(m.p), closed: {} }))
  };
}
const KS6_PREDMET = 'выполнение комплекса работ по по монтажу внутренних сетей электроснабжения и электроосвещения. на объекте:\n- "Сосновский муниципальный район Челябинской области. Квартал №3. Многоквартирный жилой дом со встроенно-пристроенными нежилыми помещениями №3.1 Этап 1" (кадастровый номер земельного участка 74:19:1203001:7337)';

let order = loadJSON(LS_ORDER, null) || defaultOrder();
let reestr = loadJSON(LS_REESTR, null) || REESTR_ORDERS.map(o => ({ ...o }));
function loadJSON(k, def) { try { const s = localStorage.getItem(k); return s ? JSON.parse(s) : def; } catch (e) { return def; } }
function saveState() { localStorage.setItem(LS_ORDER, JSON.stringify(order)); localStorage.setItem(LS_REESTR, JSON.stringify(reestr)); }

/* ---------- Ядро расчётов (формулы листов) ---------- */

// Лист «Материалы»: L8 = IF(type="Работа",ROUND(G*coefWork,2),0)+IF(type="Материалы",ROUND(G*coefMat,2),0)
function rowPriceL(r) {
  if (r.type === 'Работа') return round(r.price * order.coefWork, 2);
  if (r.type === 'Материалы') return round(r.price * order.coefMat, 2);
  return 0;
}
// H8 = ROUNDUP(F*G,2) — сумма с НДС по прайсу; I/J — суммы с коэфф. по типу
function rowSums(r) {
  const L = rowPriceL(r);
  return {
    h: r.qty && r.price ? roundup(r.qty * r.price, 2) : 0,
    i: r.type === 'Материалы' ? round(L * r.qty, 2) : 0,
    j: r.type === 'Работа' ? round(L * r.qty, 2) : 0,
    l: L
  };
}
function materialsTotals() {
  let i = 0, j = 0;
  for (const r of order.rows) { if (!isSection(r) && num(r.qty)) { const s = rowSums(r); i += s.i; j += s.j; } }
  return { mat: round(i, 2), work: round(j, 2, ), total: round(i + j, 2) };
}
const isSection = r => !num(r.qty) && r.name && (!r.type || r.type === '0' || String(r.type) === '0');

// Позиция «№ по калькуляции» = индекс строки в массиве rows (аналог позиции В в Смета/КС-6)
function posIndex(i) { return i; }

// Лист «Объемы закрытия»: по каждому периоду — пары (№ позиции, объем) из КС-6
function closureColumns() {
  return order.periods.map((p, pi) => {
    const pairs = [];
    order.rows.forEach((r, ri) => {
      const v = num(r.closed?.[pi]);
      if (v !== 0 && !isSection(r)) pairs.push({ pos: ri, vol: v });
    });
    return { ...p, pi, pairs };
  });
}

// Лист «Смета»: J,K (выполнено с начала), N,O (остаток), Q..AR (по периодам)
function smetaRows() {
  const cols = closureColumns();
  return order.rows.map((r, ri) => {
    const E = rowPriceL(r), F = num(r.qty);
    const perPeriod = cols.map(c => c.pairs.find(p => p.pos === ri)?.vol ?? 0);
    const doneVol = perPeriod.reduce((a, b) => a + b, 0);
    const doneSum = round(doneVol * E, 2);
    const restVol = F - doneVol, restSum = round(restVol * E, 2);
    return { r, ri, E, F, perPeriod, doneVol, doneSum, restVol, restSum, isSec: isSection(r) };
  });
}
function smetaTotals() {
  const rows = smetaRows();
  const T = { calc: 0, doneSum: 0, restSum: 0, perPeriodSum: order.periods.map(() => 0), perPeriodVol: order.periods.map(() => 0) };
  for (const s of rows) {
    if (s.isSec) continue;
    T.calc += round(s.F * s.E, 2);
    T.doneSum += s.doneSum; T.restSum += s.restSum;
    s.perPeriod.forEach((v, i) => { T.perPeriodSum[i] += round(v * s.E, 2); T.perPeriodVol[i] += v; });
  }
  T.calc = round(T.calc, 2); T.doneSum = round(T.doneSum, 2); T.restSum = round(T.restSum, 2);
  T.perPeriodSum = T.perPeriodSum.map(v => round(v, 2));
  // G3 = SUM(стоимость) — сметная стоимость в тыс.руб. для КС-6 H25
  return T;
}

// НДС: лист КС-2 A2 = VLOOKUP(org, Контрагенты!B3:T7,10) → «НДС 22%»/«НДС 5%»
function currentNds() {
  const k = konsByListKey(order.org);
  return k ? String(k[10] || 'НДС 22%') : 'НДС 22%';
}
// Формула Смета!Q3: безНДС = сНДС − ROUND(сНДС*rate/(100+rate),2)
function stripNds(sumWithNds) {
  const rate = currentNds() === 'НДС 5%' ? 5 : 22;
  return round(round(sumWithNds, 2) - round(sumWithNds * rate / (100 + rate), 2), 2);
}

// Даты столбцов КС-2 S3:S14 (аналог «Объемы закрытия»!A1..AK1 = КС-6!F32..N32)
function ks2ColumnsDates() { return order.periods.map(p => p.date ? new Date(p.date) : null); }

// Строки акта КС-2 (лист КС-2, строки 26..2002):
// F26 = Σ IF(Дата_КС2 = дата столбца, объем периода, 0); H = ROUND(F*G,2)
// R/S — разнесение работа/материалы (Смета!H/G), I/K — признак печати («ЗАГ» — заголовки разделов)
function ks2ActRows() {
  const dateKS2 = new Date(order.ks2Date);
  const cols = ks2ColumnsDates();
  const sm = smetaRows();
  const out = [];
  sm.forEach(s => {
    const F = s.perPeriod.reduce((acc, v, i) => acc + (cols[i] && +cols[i] === +dateKS2 ? v : 0), 0);
    const G = s.E;
    const H = round(F * G, 2);
    const workPart = F > 0 && s.r.type === 'Работа' ? H : 0;
    const matPart = F > 0 && s.r.type === 'Материалы' ? H : 0;
    if (s.isSec) { out.push({ kind: 'ЗАГ', name: s.r.name }); return; }
    if (roundup(F, 2) !== 0) out.push({ kind: 'row', pos: s.ri + 1, name: s.r.name, unit: s.r.unit, F, G, H, workPart, matPart });
  });
  // нумерация Б26: =IF(F>0, MAX(пред. номера)+1, ...)
  let max = 0;
  for (const o of out) if (o.kind === 'row') o.B = ++max;
  // ИТОГИ: H2003 = Σ H; R2003=Σ работа; S2003=Σ материалы
  const tot = {
    all: round(out.reduce((a, o) => a + (o.H || 0), 0), 2),
    work: round(out.reduce((a, o) => a + (o.workPart || 0), 0), 2),
    mat: round(out.reduce((a, o) => a + (o.matPart || 0), 0), 0)
  };
  tot.nds = round(currentNds() === 'НДС 5%' ? tot.all * 5 / 105 : tot.all * 22 / 122, 2);
  // H2006 = ROUND(H2003,2) где H2003 — уже без НДС? В книге H2006=ROUND(H2003), H2003=SUM всего;
  // «Всего по акту» = сумма строк; НДС выделяется из неё.
  return { rows: out, tot };
}

// Отчётный период КС-2: G20 = IF(W19=0, DATE(YEAR(A1),MONTH(A1),1), W19); H20 = IF(W20=0, EOMONTH(A1,0), W20)
function ks2Period() {
  const d = new Date(order.ks2Date);
  const from = order.manualFrom ? new Date(order.manualFrom) : new Date(d.getFullYear(), d.getMonth(), 1);
  const to = order.manualTo ? new Date(order.manualTo) : eomonth(d);
  return { from, to };
}

/* ---------- Навигация ---------- */
document.querySelectorAll('#tabs button').forEach(b => b.onclick = () => {
  document.querySelectorAll('#tabs button').forEach(x => x.classList.toggle('active', x === b));
  document.querySelectorAll('main>.tab').forEach(t => t.classList.add('hidden'));
  document.getElementById('tab-' + b.dataset.tab).classList.remove('hidden');
  if (b.dataset.tab === 'reestr2') renderReestr();
  if (b.dataset.tab === 'perechen') renderPerechen();
});
document.querySelectorAll('#zakaz-subtabs button').forEach(b => b.onclick = () => {
  document.querySelectorAll('#zakaz-subtabs button').forEach(x => x.classList.toggle('active', x === b));
  document.querySelectorAll('#tab-zakaz .sub').forEach(t => t.classList.add('hidden'));
  document.getElementById('sub-' + b.dataset.sub).classList.remove('hidden');
  renderAll();
});

/* ---------- Рендер: ВВОД / Материалы ---------- */
function fillSelect(sel, items, value, onChange) {
  sel.innerHTML = '';
  for (const it of items) sel.append(el('option', { value: it.v }, it.l));
  sel.value = value;
  sel.onchange = onChange;
}
function renderHeaderForm() {
  document.getElementById('f-objectName').value = order.objectName;
  document.getElementById('f-objectName').oninput = e => { order.objectName = e.target.value; saveState(); renderExport(); };
  fillSelect(document.getElementById('f-org'), ourOrgs().map(k => ({ v: k[1], l: `${k[1]} — ${String(k[2]).trim()}` })), order.org,
    e => { order.org = e.target.value; saveState(); renderAll(); });
  document.getElementById('f-coefMat').value = order.coefMat;
  document.getElementById('f-coefMat').oninput = e => { order.coefMat = num(e.target.value); saveState(); renderAll(); };
  document.getElementById('f-coefWork').value = order.coefWork;
  document.getElementById('f-coefWork').oninput = e => { order.coefWork = num(e.target.value); saveState(); renderAll(); };
  document.getElementById('f-ks2date').value = iso(order.ks2Date);
  document.getElementById('f-ks2date').onchange = e => { order.ks2Date = e.target.value; saveState(); renderAll();
    // аналог Worksheet_Change по Дата_КС2: автофильтр Таблицы1 по Условия_фильтра (M5:M6 => печать >0)
  };
  document.getElementById('f-version').value = order.version;

  const t = materialsTotals();
  document.getElementById('mat-totals').innerHTML = `
    <div class="t">Материалы (Σ I6)<b>${fmt(t.mat)}</b></div>
    <div class="t">Работа (Σ J6)<b>${fmt(t.work)}</b></div>
    <div class="t">Итого (Σ K6)<b>${fmt(t.total)}</b></div>
    <div class="t">Ставка НДС (А2 КС-2)<b>${currentNds()}</b></div>`;
}

function renderMaterialsTable() {
  const tb = document.getElementById('tbl-materials');
  tb.innerHTML = '';
  const head = el('tr', {}, ...['№', 'Тип', 'Наименование', 'Ед.', 'Кол-во (F)', 'Цена прайс (G)', 'Сумма H (ROUNDUP F×G)',
    'Цена с коэфф (L)', 'Материалы (I)', 'Работа (J)', ''].map(h => el('th', {}, h))));
  tb.append(head);
  let counter = 0;
  order.rows.forEach((r, i) => {
    const s = rowSums(r);
    if (!isSection(r) && num(r.qty)) counter++;
    const tr = el('tr', { class: isSection(r) ? 'section-row' : '' },
      el('td', { class: 'rownum' }, String(counter || '')),
      tdSel(['', 'Материалы', 'Работа', '0'], r.type, v => { r.type = v; afterChange(); }),
      tdInp(r.name, v => { r.name = v; afterChange(false); }, true),
      tdInp(r.unit, v => { r.unit = v; afterChange(false); }),
      tdInp(r.qty, v => { r.qty = num(v); afterChange(); }, false, 'num'),
      tdInp(r.price, v => { r.price = num(v); afterChange(); }, false, 'num'),
      el('td', { class: 'num' }, s.h ? fmt(s.h) : ''),
      el('td', { class: 'num' }, s.l ? fmt(s.l) : ''),
      el('td', { class: 'num' }, s.i ? fmt(s.i) : ''),
      el('td', { class: 'num' }, s.j ? fmt(s.j) : ''),
      el('td', {}, el('button', { class: 'mini', onclick: () => { order.rows.splice(i, 1); afterChange(); } }, '✕')));
    tb.append(tr);
  });
}
function tdInp(v, set, wide = false, cls = '') {
  const inp = el('input', { value: v ?? '', class: cls, style: wide ? 'min-width:340px' : '' });
  inp.addEventListener('change', e => set(e.target.value));
  return el('td', { class: cls }, inp);
}
function tdSel(opts, v, set) {
  const s = el('select', {});
  for (const o of opts) s.append(el('option', { value: o }, o));
  s.value = v ?? '';
  s.addEventListener('change', e => set(e.target.value));
  return el('td', {}, s);
}
let _raf = 0;
function afterChange(full = true) {
  saveState();
  cancelAnimationFrame(_raf);
  _raf = requestAnimationFrame(() => { full ? renderAll() : lightRender(); });
}
/* ---------- Рендер: КС-6 (шапка) — вызывается из lightRender ---------- */
function renderKs6All() { renderKs6Header(); renderKs6Periods(); renderKs6Table(); }
function lightRender() { renderHeaderForm(); renderMaterialsTable(); renderKs6All(); renderSmeta(); renderKs2(); renderKs3(); renderExport(); }
function renderAll() { renderKontr(); lightRender(); }

/* ---------- Рендер: Контрагенты ---------- */
function renderKontr() {
  const tb = document.getElementById('tbl-kontr');
  tb.innerHTML = '';
  tb.append(el('tr', {}, ...['Ключ (B)', 'Наименование (C)', 'ИНН (D)', 'КПП (E)', 'ОКПО (F)', 'Адрес (G)', 'Подпись (H)', 'Должность (I)', 'ФИО (J)', 'НДС (K)'].map(h => el('th', {}, h)))));
  KONS.forEach((k, idx) => {
    tb.append(el('tr', {},
      tdInp(k[1], v => { k[1] = v; saveState(); }),
      tdInp(k[2], v => { k[2] = v; saveState(); }, true),
      tdInp(k[3], v => { k[3] = v; saveState(); }),
      tdInp(k[4], v => { k[4] = v; saveState(); }),
      tdInp(k[5], v => { k[5] = v; saveState(); }),
      tdInp(k[6], v => { k[6] = v; saveState(); }, true),
      tdInp(k[7], v => { k[7] = v; saveState(); }),
      tdInp(k[8], v => { k[8] = v; saveState(); }),
      tdInp(k[9], v => { k[9] = v; saveState(); }),
      tdInp(k[10], v => { k[10] = v; saveState(); })));
  });
}
document.getElementById('btn-add-kontr').onclick = () => {
  KONS.push([0, 'Новый', 'Новое наименование', '', '', '', '', '', '', '', 'НДС 22%', '', '', '', '', '', '', '', '', '']);
  saveState(); renderAll();
};
document.getElementById('btn-add-row').onclick = () => {
  order.rows.push({ type: 'Материалы', name: '', unit: 'шт', qty: 0, price: 0, closed: {} });
  afterChange();
};
document.getElementById('btn-clear-filters').onclick = () => {
  order.rows.forEach(r => r.closed = {}); afterChange(); toast('Фильтры закрытия очищены (объемы КС-6 обнулены)');
};

/* ---------- Рендер: КС-6 ---------- */
function renderKs6Header() {
  const cust = customerList().map(k => ({ v: k[1], l: String(k[1]) }));
  if (!cust.find(c => c.v === order.customer)) cust.push({ v: order.customer, l: String(order.customer) });
  fillSelect(document.getElementById('f-customer'), cust, order.customer, e => { order.customer = e.target.value; saveState(); renderKs2(); renderKs3(); });
  const bindTA = (id, key) => { const t = document.getElementById(id); t.value = order[key]; t.oninput = e => { order[key] = e.target.value; saveState(); }; };
  bindTA('f-stroyka', 'stroyka'); bindTA('f-obekt', 'obekt'); bindTA('f-predmet', 'predmet'); bindTA('f-dogovorNum', 'dogovorNum');
  const dd = document.getElementById('f-dogovorDate'); dd.value = iso(order.dogovorDate); dd.onchange = e => { order.dogovorDate = e.target.value; saveState(); };
  const T = smetaTotals();
  document.getElementById('ks6-smeta-cost').innerHTML =
    `<span>Сметная (договорная) стоимость: <b>${fmt(T.calc / 1000, 3)} тыс. руб.</b> (=Смета!G3/1000)</span>
     <span>Заказчик (строка C3): <b>${customerFull()}</b></span>
     <span>Подрядчик (C5): <b>${contractorFull()}</b></span>`;
}
function customerFull() {
  const k = konsByListKey(order.customer);
  // КС-6!C3 = CONCAT через VLOOKUP(P3, Контрагенты!B10:T78,13) → колонка M13? фактически готовая строка N:
  return k ? (k[13] && k[13] !== 0 ? String(k[13]).trim() : `${String(k[2]).trim()}, ИНН ${k[3]}, ${k[6]}`) : String(order.customer);
}
function contractorFull() {
  const k = konsByListKey(order.org);
  return k ? String(k[13] && k[13] !== 0 ? k[13] : k[2]).trim() : String(order.org);
}
function renderKs6Periods() {
  const box = document.getElementById('ks6-periods');
  box.innerHTML = '';
  const grid = el('div', { class: 'grid2' });
  order.periods.forEach((p, i) => {
    const locked = p.status !== 'Закрытие';
    grid.append(el('label', {}, `Период ${i + 1} (столбец ${'FGHIJKLMN'[i]})`,
      (() => { const d = el('input', { type: 'date', value: iso(p.date) });
        d.onchange = e => { p.date = e.target.value;
          // каскад дат: след. дата = +31 день (G32=F32+31) — только если следующая не задана вручную
          saveState(); renderAll(); }; return d; })(),
      (() => { const s = el('select', {});
        for (const o of ['----', 'Закрытие']) s.append(el('option', { value: o }, o));
        s.value = p.status;
        // Аналог макроса Лист5.Worksheet_Change: при «Закрытие» — столбец разблокирован для ввода
        s.onchange = e => { p.status = e.target.value;
          if (p.status !== 'Закрытие') order.rows.forEach(r => delete r.closed[i]);
          saveState(); renderAll(); toast(p.status === 'Закрытие' ? `Столбец ${'FGHIJKLMN'[i]} открыт для ввода (как «Закрытие» в F30:N30)` : `Столбец ${'FGHIJKLMN'[i]} закрыт для ввода`); };
        return s; })()));
  });
  box.append(grid);
}
function renderKs6Table() {
  const tb = document.getElementById('tbl-ks6');
  tb.innerHTML = '';
  const headCells = ['№ поз. (B)', 'Наименование (C)', 'Ед. (D)', 'Кол-во (E)'];
  order.periods.forEach((p, i) => headCells.push(`Объём ${iso(p.date)}${p.status === 'Закрытие' ? ' ✎' : ' 🔒'}`));
  headCells.push('Остаток (Q)');
  tb.append(el('tr', {}, ...headCells.map(h => el('th', {}, h))));
  order.rows.forEach((r, ri) => {
    if (isSection(r)) { tb.append(el('tr', { class: 'section-row' }, el('td', {}, ''), el('td', { colspan: String(headCells.length - 1) }, String(r.name)))); return; }
    if (!num(r.qty) && !r.name) return;
    const sumClosed = order.periods.map((_, i) => num(r.closed?.[i])).reduce((a, b) => a + b, 0);
    const rest = num(r.qty) - sumClosed;
    const tr = el('tr', {},
      el('td', { class: 'rownum' }, String(ri)),
      el('td', { style: 'max-width:420px;white-space:normal' }, String(r.name || '')),
      el('td', {}, String(r.unit || '')),
      el('td', { class: 'num' }, fmt0(num(r.qty))));
    order.periods.forEach((p, i) => {
      const editable = p.status === 'Закрытие';
      const inp = el('input', { type: 'number', value: r.closed?.[i] ?? '', disabled: !editable ? '' : null, class: 'num', title: editable ? '' : 'Статус «----»: ввод заблокирован (макрос защиты)' });
      if (editable) inp.addEventListener('change', e => { r.closed[i] = num(e.target.value); afterChange(); });
      tr.append(el('td', {}, inp));
    });
    // Q34 = IF(SUM(F:N)>E,"Превышение",E-SUM)
    tr.append(el('td', { class: 'num', style: rest < 0 ? 'color:#b03a2e;font-weight:700' : '' }, rest < 0 ? 'Превышение' : fmt0(rest)));
    tb.append(tr);
  });
}

/* ---------- Рендер: Смета ---------- */
function renderSmeta() {
  const tb = document.getElementById('tbl-smeta'); tb.innerHTML = '';
  const T = smetaTotals();
  const heads = ['№ (A)', 'поз.калк. (B)', 'Наименование (C)', 'Ед. (D)', 'Цена (E)', 'Кол-во (F)', 'Стоимость (G+H)'];
  order.periods.forEach((p, i) => heads.push(`Закрыв. ${iso(p.date)}: кол-во/сумма`));
  heads.push('Выполнено с начала (J/K)', 'Остаток (N/O)');
  tb.append(el('tr', {}, ...heads.map(h => el('th', {}, h))));
  for (const s of smetaRows()) {
    if (s.isSec) { tb.append(el('tr', { class: 'section-row' }, el('td', { colspan: String(heads.length) }, String(s.r.name)))); continue; }
    if (!num(s.F) && !s.r.name) continue;
    const tr = el('tr', {},
      el('td', { class: 'rownum' }, String(s.ri)), el('td', { class: 'rownum' }, String(s.ri)),
      el('td', { style: 'max-width:380px;white-space:normal' }, String(s.r.name)),
      el('td', {}, String(s.r.unit)), el('td', { class: 'num' }, fmt(s.E)), el('td', { class: 'num' }, fmt0(s.F)),
      el('td', { class: 'num' }, fmt(round(s.F * s.E, 2))));
    s.perPeriod.forEach(v => tr.append(el('td', { class: 'num' }, v ? `${fmt0(v)} / ${fmt(round(v * s.E, 2))}` : '')));
    tr.append(el('td', { class: 'num' }, `${fmt0(s.doneVol)} / ${fmt(s.doneSum)}`));
    tr.append(el('td', { class: 'num', style: s.restVol < 0 ? 'color:#b03a2e' : '' }, `${fmt0(s.restVol)} / ${fmt(s.restSum)}`));
    tb.append(tr);
  }
  document.getElementById('smeta-head').innerHTML =
    `<span>Смета на сумму (G3): <b>${fmt(T.calc)}</b></span>
     <span>Выполнено с начала (J2, с НДС): <b>${fmt(T.doneSum)}</b></span>
     <span>Без НДС (K3): <b>${fmt(stripNds(T.doneSum))}</b></span>
     <span>Остаток (N2): <b>${fmt(T.restSum)}</b></span>
     <span>Без НДС (O3): <b>${fmt(stripNds(T.restSum))}</b></span>`;
}

/* ---------- Рендер: документ КС-2 ---------- */
function renderKs2() {
  const host = document.getElementById('ks2-doc');
  const act = ks2ActRows();
  const per = ks2Period();
  const cust = customerFull(), contr = contractorFull();
  const orgK = konsByListKey(order.org);
  const sended = orgK ? `${orgK[8]} ${String(orgK[9]).trim()}` : '';
  const custK = konsByListKey(order.customer);
  const accepted = custK ? `${custK[8]} ${String(custK[9]).trim()}` : '';
  const docNum = ks2DocNum();
  const rowsHtml = act.rows.map(o => o.kind === 'ЗАГ'
    ? `<tr class="section-row"><td colspan="6">${esc(o.name)}</td></tr>`
    : `<tr><td>${o.B}</td><td>${esc(o.name)}</td><td>${esc(o.unit)}</td><td class="num">${fmt(o.F, 3)}</td><td class="num">${fmt(o.G)}</td><td class="num">${fmt(o.H)}</td></tr>`).join('');
  host.innerHTML = `
  <p class="small">Унифицированная форма № КС-2<br>Утверждена постановлением Госкомстата России от 11 ноября 1999 г. № 100</p>
  <table style="width:auto;margin-bottom:8px"><tr><td class="small">Форма по ОКУД</td><td>0322005</td></tr></table>
  <p><b>Дата составления:</b> ${ruDate(order.ks2Date)} &nbsp; (${esc(String(order.objectName))})<br>
  ${esc(String(order.objectName))}: ${esc(contractorFull())}<br>
  Заказчик (Генподрядчик) — ${esc(cust)}<br>
  Подрядчик (Субподрядчик) — ${esc(contr)}<br>
  Стройка — ${esc(String(order.stroyka))}<br>
  Объект — ${esc(String(order.obekt))}<br>
  Договор подряда (контракт) № ${esc(String(order.dogovorNum)).replace(/\n/g, ' ')} от ${ruDate(order.dogovorDate)}</p>
  <table style="width:auto;margin:6px 0">
    <tr><td class="small">Номер документа</td><td>${docNum}</td>
        <td class="small">Дата составления</td><td>${ruDate(order.ks2Date)}</td>
        <td class="small">Отчётный период</td><td>с ${ruDate(per.from)} по ${ruDate(per.to)}</td></tr>
  </table>
  <h2>АКТ О ПРИЕМКЕ ВЫПОЛНЕННЫХ РАБОТ</h2>
  <table><thead><tr><th>№ п/п</th><th>Наименование работ</th><th>Ед.изм</th><th>Объём</th><th>Цена, руб.</th><th>Стоимость, руб.</th></tr></thead>
  <tbody>${rowsHtml}</tbody>
  <tfoot>
    <tr><td colspan="5"><b>Работы</b></td><td class="num"><b>${fmt(act.tot.work)}</b></td></tr>
    <tr><td colspan="5"><b>Материалы</b></td><td class="num"><b>${fmt(act.tot.mat)}</b></td></tr>
    <tr><td colspan="5"><b>Всего по акту</b></td><td class="num"><b>${fmt(act.tot.all)}</b></td></tr>
    <tr><td colspan="5"><b>${currentNds() === 'НДС 5%' ? 'В т.ч. НДС 5%' : 'В т.ч. НДС 22%'}</b></td><td class="num"><b>${fmt(act.tot.nds)}</b></td></tr>
  </tfoot></table>
  <p style="margin-top:8px"><i>${esc(numWords(act.tot.all))}</i></p>
  <div class="sign"><div><b>Сдал:</b> ${esc(sended)}<br><span class="small">(должность, подпись, расшифровка) &nbsp; М.П.</span></div>
  <div><b>Принял:</b> ${esc(accepted)}<br><span class="small">(должность, подпись, расшифровка) &nbsp; М.П.</span></div></div>
  <p class="small">Запрет печати (аналог имени «Запрет_печати», Workbook_BeforePrint): включён — печать из Excel-шаблона блокируется, пока не сформирован PDF макросом.</p>`;
}
function ks2DocNum() {
  // E20 = COUNTIFS(Смета!Q4:AY4,"<="&Дата_КС2, Смета!Q5:AY5,">0") — номер документа = число закрытых периодов с суммой>0
  const d = new Date(order.ks2Date);
  const T = smetaTotals();
  return order.periods.filter((p, i) => p.date && new Date(p.date) <= d && T.perPeriodSum[i] > 0).length;
}
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/* ---------- Рендер: КС-3 ---------- */
function renderKs3() {
  const host = document.getElementById('ks3-doc');
  const act = ks2ActRows();
  const T = smetaTotals();
  const ndsLabel = currentNds() === 'НДС 5%' ? 'Сумма НДС 5%' : 'Сумма НДС 22%';
  // F22 = IF(КС-2!A2="Без НДС", Смета!J2, Смета!K3) — «с начала проведения работ»
  const totalAll = round(act.tot.all, 2);
  const withNdsTotal = T.doneSum;               // с начала работ (с НДС)
  const noNds = stripNds(withNdsTotal);
  host.innerHTML = `
  <p class="small">Унифицированная форма № КС-3, Форма по ОКУД 0322001</p>
  <p><b>СПРАВКА О СТОИМОСТИ ВЫПОЛНЕННЫХ РАБОТ И ЗАТРАТ</b><br>
  Инвестор: ${esc(customerFull())}<br>
  Подрядчик: ${esc(contractorFull())}<br>
  Контракт № ${esc(String(order.dogovorNum)).replace(/\n/g, ' ')} от ${ruDate(order.dogovorDate)}<br>
  Номер документа: ${ks2DocNum()} от ${ruDate(order.ks2Date)}, отчётный период с ${ruDate(ks2Period().from)} по ${ruDate(ks2Period().to)}</p>
  <table><thead><tr><th>№</th><th>Наименование</th><th>С начала проведения работ, руб.</th><th>В том числе за отчётный период, руб.</th></tr></thead>
  <tbody>
    <tr><td>1</td><td>Всего работ и затрат</td><td class="num">${fmt(withNdsTotal)}</td><td class="num">${fmt(totalAll)}</td></tr>
    <tr><td>1.1</td><td>${esc(String(order.predmet).split('\n')[0])}</td><td class="num">${fmt(withNdsTotal)}</td><td class="num">${fmt(totalAll)}</td></tr>
    <tr><td colspan="2"><b>Итого</b></td><td class="num"><b>${fmt(noNds)}</b></td><td class="num"><b>${fmt(stripNds(totalAll))}</b></td></tr>
    <tr><td colspan="2"><b>${ndsLabel}</b></td><td class="num">${fmt(round(withNdsTotal - noNds, 2))}</td><td class="num">${fmt(round(totalAll - stripNds(totalAll), 2))}</td></tr>
    <tr><td colspan="2"><b>Всего с учетом НДС</b></td><td class="num"><b>${fmt(withNdsTotal)}</b></td><td class="num"><b>${fmt(totalAll)}</b></td></tr>
  </tbody></table>
  <div class="sign"><div>Заказчик (генподрядчик) ____________</div><div>Подрядчик (субподрядчик) ____________</div></div>`;
}

/* ---------- Экспорт → Реестр 2 (лист «Экспорт», макросы Module1/4/5) ---------- */
function exportRow() {
  const T = smetaTotals(), act = ks2ActRows();
  const now = new Date();
  const user = (navigator.userAgent.match(/\(([^)]*)\)/) || [])[1]?.split(';')[0] || 'WebUser';
  // Порядок = лист Экспорт A3:AE3 (29 полей), вставляется в Реестр начиная с колонки E
  return {
    source: 'WEB', date_contract: order.dogovorDate, object: order.objectName, legalForm: '',
    customer: customerFull(), leaderPost: '', leaderFIO: '', predmet: String(order.predmet),
    addressObj: String(order.obekt), addressLegal: '', close_sum: act.tot.all, innkpp: '',
    curator: '', plan_cost: T.calc, fact_cost: T.doneSum,
    mat_plan: materialsTotals().mat, mat_fact: '', usl_plan: materialsTotals().work, usl_fact: '',
    nr_plan: '', nr_fact: '', version: order.version, adm_fact: '', profit_plan: '', profit_fact: '',
    zp_plan: '', zp_fact: '', prem_plan: '', prem_fact: '', revenue: '',
    path: '\\\\Nas\\Shared Docs\\АРХИВ ЗАКАЗОВ 2\\' + order.objectName + '\\Заказ_' + order.objectName + '.xlsm',
    time_rec: now.toLocaleTimeString('ru-RU'), date_rec: now.toISOString().slice(0, 10),
    user: user, computer: 'WEB', employee: user,
    folder: '\\\\Nas\\Shared Docs\\АРХИВ ЗАКАЗОВ 2\\' + order.objectName,
    docType: 'РАСЧЕТ', templateVer: order.version
  };
}
const EXPORT_FIELDS = [
  ['source', 'Источник Заказа'], ['date_contract', 'Дата договора'], ['object', 'Имя ОБЪЕКТА в РЕЕСТР'],
  ['legalForm', 'Правовая форма'], ['customer', 'Название Заказчика'], ['leaderPost', 'Должность руководителя'],
  ['leaderFIO', 'ФИО руководителя'], ['predmet', 'Предмет договора'], ['addressObj', 'Адрес объекта'],
  ['addressLegal', 'Адрес Юридический'], ['close_sum', 'Закрытие на Сумму'], ['innkpp', 'ИНН/КПП'],
  ['curator', 'Куратор'], ['plan_cost', 'Стоимость контракта (ПЛАН)'], ['fact_cost', 'Стоимость контракта (ФАКТ)'],
  ['mat_plan', 'Материалы (ПЛАН)'], ['mat_fact', 'Материалы (ФАКТ)'], ['usl_plan', 'Услуги (ПЛАН)'],
  ['usl_fact', 'Услуги (ФАКТ)'], ['nr_plan', 'НР (ПЛАН)'], ['nr_fact', 'НР (ФАКТ)'],
  ['version', 'Версия программы'], ['adm_fact', 'Адм. расходы (ФАКТ)'], ['profit_plan', 'Прибыль (ПЛАН)'],
  ['profit_fact', 'Прибыль (ФАКТ)'], ['zp_plan', 'ЗП Менеджер (ПЛАН)'], ['zp_fact', 'ЗП Менеджер (ФАКТ)'],
  ['prem_plan', 'Премия ФОТ (ПЛАН)'], ['prem_fact', 'Премия ФОТ (ФАКТ)'], ['revenue', 'Выручка'],
  ['path', 'Путь к файлу'], ['time_rec', 'ВремяЗаписи'], ['date_rec', 'ДатаЗаписи'],
  ['user', 'Пользователь'], ['computer', 'Компьютер'], ['employee', 'Сотрудник'], ['folder', 'Папка хранения']
];
function renderExport() {
  const tb = document.getElementById('tbl-export'); tb.innerHTML = '';
  const row = exportRow();
  tb.append(el('tr', {}, el('th', {}, 'Поле'), el('th', {}, 'Значение (строка «Экспорт» A3:AE3 → вставка в Реестр!E:…)')));
  for (const [k, label] of EXPORT_FIELDS) {
    const v = row[k];
    tb.append(el('tr', {}, el('td', { style: 'font-weight:600' }, label),
      el('td', { style: 'white-space:normal;max-width:600px' }, typeof v === 'number' ? fmt(v) : esc(v))));
  }
}
/* Макрос «СохранениеНовогоЗаказа»: проверка дубликата по колонке G (object),
   перезапись или новая строка в конец (Строка = Реестр!A1 + 11 → index = count + позиция 11) */
document.getElementById('btn-save-to-reestr').onclick = () => {
  const log = document.getElementById('export-log');
  log.textContent = '';
  const write = () => {
    if (!String(order.objectName).trim()) { log.textContent += '❌ Нет Названия Заказа (ИмяОБЪЕКТАвРЕЕСТР) — сохранение отменено, как в MsgBox макроса.\n'; return; }
    const rec = exportRow();
    const dupIdx = reestr.findIndex(r => String(r.object) === String(rec.object));
    if (dupIdx >= 0) {
      const yes = confirm('Есть такой ЗАКАЗ. Переписать?');   // аналог MsgBox vbYesNo
      if (!yes) { log.textContent += '⏹ Отменено пользователем (Reestr не изменён).\n'; return; }
      reestr[dupIdx] = { ...reestr[dupIdx], ...rec };
      log.textContent += `♻ Заказ найден в реестре (строка ${dupIdx + 11}) — перезаписан.\n`;
    } else {
      reestr.push(rec);
      log.textContent += `➕ Новая запись добавлена в строку ${reestr.length + 10 + 1} (Лист Excel: A1(${reestr.length})+11).\n`;
    }
    saveState();
    log.textContent += `💾 Workbooks("Реестр 2.xlsm").Save → Close\n✔ Изменения внесены в Реестр 2.\n`;
    toast('Изменения внесены в Реестр 2.');
    renderReestr(); renderPerechen();
  };
  write();
};
document.getElementById('btn-pdf-ks2').onclick = () => {
  // Аналог Сохранить_в_PDF_КС2: снимает «Запрет_печати», печатает КС-6 (скрытым шрифтом F30:N30) и КС-2
  document.querySelectorAll('.tab,.sub').forEach(x => x.classList.add('hidden'));
  const k2 = document.getElementById('tab-zakaz'); k2.classList.remove('hidden');
  document.querySelectorAll('#tab-zakaz .sub').forEach(x => x.classList.add('hidden'));
  document.getElementById('sub-ks2').classList.remove('hidden');
  window.print();
  document.getElementById('tab-zakaz').classList.remove('hidden');
  document.querySelectorAll('#tab-zakaz .sub').forEach(x => x.classList.add('hidden'));
  document.getElementById('sub-ks2').classList.remove('hidden');
  toast('PDF формируется через диалог печати браузера (Сохранить как PDF). Путь в Excel: \\\\Nas\\Shared Docs\\ЗАКРЫТИЕ\\…');
};
document.getElementById('btn-export-xls').onclick = () => {
  // Аналог «Экспорт_в_Exell»: значения без формул, фильтр «Печать=1»
  const act = ks2ActRows();
  let csv = '\uFEFF' + '№;Наименование;Ед.изм;Объем;Цена;Стоимость\n';
  for (const o of act.rows) csv += o.kind === 'ЗАГ' ? `;${csvq(o.name)};;;;;\n` : `${o.B};${csvq(o.name)};${o.unit};${o.F};${o.G};${o.H}\n`;
  download('Экспорт_КС2_' + order.objectName + '.csv', csv);
};
const csvq = s => `"${String(s ?? '').replace(/"/g, '""')}"`;
function download(name, text) {
  const a = el('a', { href: URL.createObjectURL(new Blob([text], { type: 'text/plain;charset=utf-8' })), download: name });
  document.body.append(a); a.click(); a.remove();
}

/* ---------- Вкладка «Реестр 2» ---------- */
const REESTR_COLS = [['source', 'Источник'], ['date_contract', 'Дата договора'], ['object', 'Имя ОБЪЕКТА'],
  ['customer', 'Заказчик'], ['close_sum', 'Закрытие на Сумму'], ['plan_cost', 'План стор.'], ['fact_cost', 'Факт стор.'],
  ['mat_plan', 'Мат план'], ['usl_plan', 'Усл план'], ['version', 'Версия'], ['path', 'Путь к файлу'],
  ['date_rec', 'Дата записи'], ['user', 'Пользователь'], ['employee', 'Сотрудник'], ['folder', 'Папка']];
function renderReestr() {
  const q = (document.getElementById('reestr-search').value || '').toLowerCase();
  const tb = document.getElementById('tbl-reestr'); tb.innerHTML = '';
  tb.append(el('tr', {}, el('th', {}, 'Строка'), ...REESTR_COLS.map(c => el('th', {}, c[1]))));
  let shown = 0;
  reestr.forEach((r, i) => {
    if (q && !`${r.object} ${r.customer}`.toLowerCase().includes(q)) return;
    shown++;
    const tr = el('tr', {}, el('td', { class: 'rownum' }, String(i + 11)));
    for (const [k] of REESTR_COLS) {
      const v = r[k];
      tr.append(el('td', { style: 'white-space:normal;max-width:260px' }, typeof v === 'number' ? fmt(v) : esc(v)));
    }
    tb.append(tr);
  });
  document.getElementById('reestr-count').textContent = `${shown} из ${reestr.length}`;
  // Аналог A1 = SUM(A11:A5000) — счётчик непустых заказов
}
document.getElementById('reestr-search').oninput = renderReestr;
document.getElementById('btn-reestr-export').onclick = () => {
  let csv = '\uFEFF' + REESTR_COLS.map(c => csvq(c[1])).join(';') + '\n';
  for (const r of reestr) csv += REESTR_COLS.map(([k]) => csvq(typeof r[k] === 'number' ? r[k] : String(r[k] ?? ''))).join(';') + '\n';
  download('Реестр_2.csv', csv);
};
document.getElementById('btn-reestr-reset').onclick = () => {
  if (!confirm('Вернуть Реестр 2 к исходным данным из файла «Реестр 2.xlsm»?')) return;
  reestr = REESTR_ORDERS.map(o => ({ ...o })); saveState(); renderReestr(); renderPerechen(); toast('Реестр восстановлен из исходного снимка');
};

/* ---------- Вкладка «Перечень заказов» (Реестр ЗАКАЗОВ 2.xlsm) ----------
   Формулы: A=[2]O, B=IF(C>0,"-","Нет файла"), C=[2]G, D=[2]S, E=[2]AI, F=[2]Z, G=[2]O */
function renderPerechen() {
  const tb = document.getElementById('tbl-perechen'); tb.innerHTML = '';
  tb.append(el('tr', {}, el('th', {}, 'Стр.'), el('th', {}, 'Сумма (A)'), el('th', {}, 'Статус (B)'),
    el('th', {}, 'ОБЪЕКТ (C)'), el('th', {}, 'Сумма (D)'), el('th', {}, 'Путь (E)'), el('th', {}, 'Версия (F)'),
    el('th', {}, 'Сумма КС-6 (G)'), el('th', {}, '↑'), el('th', {}, '↓')));
  reestr.forEach((r, i) => {
    const hasFile = r.path && String(r.path).trim() && String(r.path) !== '0';
    const tr = el('tr', { ondblclick: () => openOrderFromReestr(i) },
      el('td', { class: 'rownum' }, String(i + 11)),
      el('td', { class: 'num' }, r.close_sum ? fmt(r.close_sum) : ''),
      el('td', {}, hasFile ? '-' : 'Нет файла'),
      el('td', { style: 'white-space:normal;max-width:340px' }, esc(r.object)),
      el('td', { class: 'num' }, typeof r.fact_cost === 'number' ? fmt(r.fact_cost) : esc(r.fact_cost)),
      el('td', { style: 'white-space:normal;max-width:380px;font-size:11px' }, esc(r.path)),
      el('td', {}, esc(r.version)),
      el('td', { class: 'num' }, r.close_sum ? fmt(r.close_sum) : '0'),
      el('td', { class: 'updown', onclick: ev => { ev.stopPropagation(); moveOrder(i, -1); } }, '↑'),
      el('td', { class: 'updown', onclick: ev => { ev.stopPropagation(); moveOrder(i, +1); } }, '↓'));
    tb.append(tr);
  });
}
// Аналог Worksheet_SelectionChange col L/M: Cut строки D:BA и Insert на ±1
function moveOrder(i, dir) {
  const j = i + dir;
  if (j < 0 || j >= reestr.length) return;
  [reestr[i], reestr[j]] = [reestr[j], reestr[i]];
  saveState(); renderPerechen(); renderReestr();
  toast(`Заказ перемещён ${dir < 0 ? 'вверх' : 'вниз'} (аналог Cut/Insert в Реестр 2)`);
}
// Аналог Worksheet_BeforeDoubleClick: открыть заказ (или предложить новый шаблон)
function openOrderFromReestr(i) {
  const r = reestr[i];
  const hasFile = r.path && String(r.path).trim() && String(r.path) !== '0';
  if (!hasFile) {
    if (confirm('Нет файла. Добавить новый Заказ (открыть шаблон)?')) newOrderTemplate();
    return;
  }
  if (!confirm(`Открыть? - ${r.object}`)) return;
  // «Открываем» заказ: подставляем известные данные реестра в текущую форму
  order.objectName = r.object || order.objectName;
  if (r.customer) order.customer = guessCustomerKey(r.customer) || order.customer;
  if (r.date_contract) order.dogovorDate = String(r.date_contract).slice(0, 10);
  if (r.version) order.version = r.version;
  saveState();
  document.querySelector('#tabs button[data-tab="zakaz"]').click();
  renderAll();
  toast('Заказ загружен из Реестра (в веб-версии доступны поля реестра; полная начинка — в файле на NAS)');
}
function guessCustomerKey(name) {
  const k = KONS.find(x => String(x[2]).trim() && String(name).includes(String(x[2]).trim().slice(0, 15)));
  return k ? k[1] : null;
}
function newOrderTemplate() {
  order = defaultOrder(); order.objectName = 'Новый заказ'; order.rows = [];
  saveState(); renderAll();
  document.querySelector('#tabs button[data-tab="zakaz"]').click();
  toast('Создан заказ по шаблону (аналог открытия «Расчет и Документация Шаблон 2.xlsm»)');
}

/* ---------- Инициализация ---------- */
KONS = KONTRAGENTS_RAW.map(r => [...r]);
// В исходном листе «Контрагенты» строка 8 — разделитель-метка «Заказчики»; убираем служебные пустые строки
KONS = KONS.filter(k => k[1] && String(k[1]) !== '0' && String(k[1]) !== 'Заказчики');
document.getElementById('userbox').textContent = `Пользователь: ${localStorage.getItem('webzakaz.user') || 'WebUser'} | ${new Date().toLocaleDateString('ru-RU')} | ${location.protocol === 'file:' ? 'локальный режим' : 'web'}`;
try { renderAll(); renderReestr(); renderPerechen(); }
catch (e) { console.error(e); toast('Ошибка инициализации: ' + e.message, 10000); }
window.addEventListener('error', ev => console.error(ev.message));
