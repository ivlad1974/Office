/* ============================================================
   calc.js — полный перенос формул из Excel-книг в JS
   Формулы-источники (из xlsm):
   Материалы:  H=F*G ; I=IF(B="Материалы",H*Kм,0) ; J=IF(B="Работа",H*Kr,0) ; L=I+J (цена с коэфф.)
               строки A8..A2002; итоги I6=SUM(I8:I2002), J6, K6=I6+J6
   Смета:      G3 = SUM(стоимость) ; J2 = SUMIF(даты<=Дата_КС2, суммы без НДС)
               N2 = SUMIFS(... год = YEAR(Дата_КС2))
               Q3 = Q5 - ROUND(Q5*22/122)  или *5/105 (НДС 22% / 5%) — без НДС по месяцам
               J8 = Σ объёмов закрытия по позиции ; K8 = J8*E8 ; N8 = F8-J8 ; O8 = N8*E8
   Объемы закрытия: FILTER позиций КС-6 с ненулевым объёмом + VLOOKUP(№поз → остаток Сметы col M)
   КС-2:  F26 = объём позиции для выбранной Дата_КС2 (сумма колонок месяцев, чья дата = Дата_КС2)
          H26 = ROUND(F26*G26,2); B26 нумерация IF(F>0,MAX+1,...,"ЗАГ");
          R26 работа = IF(Smeta.H>0,H26,0); S26 материалы = IF(Smeta.G>0,H26,0)
          H2003=ΣH; H2004=ΣR; H2005=ΣS; H2006=ROUND(H2003,2);
          H2007=ROUND(IF(НДС5%,H2006*5/105,H2006*22/122),2)
          E20 номер документа = COUNTIFS(даты<=Дата_КС2, суммы>0)
          G20/H20 отчётный период: manual или DATE(YEAR,MONTH,1)/EOMONTH
   КС-3:  F22 = сумма без НДС с начала работ; H22 с начала года;
          I22 = I26-I25; I25 = НДС из КС-2; I26 = Всего по акту КС-2
   ============================================================ */
(function(){
const C = {};
const r2 = x => Math.round((x+Number.EPSILON)*100)/100;   // Excel ROUND (2 знака)
C.round2 = r2;
C.num = v => { const n = parseFloat(String(v).replace(',','.')); return isNaN(n)?0:n; };

/* --- календарные помощники (Excel DATE/EOMONTH/YEAR) --- */
C.parseD = s => { if(!s) return null; const d=new Date(s+'T00:00:00'); return isNaN(d)?null:d; };
C.addDays = (d,n)=>{const x=new Date(d); x.setDate(x.getDate()+n); return x;};
C.eomonth = d => new Date(d.getFullYear(), d.getMonth()+1, 0);
C.sameDate = (a,b)=> a&&b && a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate();
C.fmtD = d => d ? ('0'+d.getDate()).slice(-2)+'.'+('0'+(d.getMonth()+1)).slice(-2)+'.'+d.getFullYear() : '';
C.MONTHS_RU=['январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь'];
C.money = x => (isFinite(x)?x:0).toLocaleString('ru-RU',{minimumFractionDigits:2,maximumFractionDigits:2});

/* ---------- МАТЕРИАЛЫ: расчёт строк и итогов ----------
   order.rows: [{t:'Материалы'|'Работа'|null(header), n,u,q,p,note}] */
C.calcMaterials = function(order){
  const Km=C.num(order.meta.coefMat??1), Kr=C.num(order.meta.coefWork??1);
  let pos=0, sumMat=0, sumWork=0;
  order.rows.forEach(row=>{
    const q=C.num(row.q), p=C.num(row.p);
    row.sum = (q&&p)? r2(q*p) : (row.t? r2(q*p):0);         // H = F*G
    row.I = (row.t==='Материалы') ? r2(row.sum*Km) : 0;     // Σ матер с коэфф
    row.J = (row.t==='Работа')    ? r2(row.sum*Kr) : 0;     // Σ работа с коэфф
    row.L = (row.t==='Материалы') ? r2(p*Km) : (row.t==='Работа'? r2(p*Kr):0); // цена с коэфф
    if(row.t) pos++;
    row.pp = pos;                                            // A=IF(F>0,COUNTIF(">0"))
    sumMat+=row.I; sumWork+=row.J;
  });
  order.totals={mat:r2(sumMat),work:r2(sumWork),all:r2(sumMat+sumWork)};
};

/* ---------- Нумерация позиций для Сметы/КС-6 ----------
   позиция = порядковый номер строки-предмет (t!=null), как в Excel B8=Материалы!A8 */
C.positions = function(order){
  const arr=[];
  order.rows.forEach((row,i)=>{ if(row.t) arr.push({idx:i,pos:arr.length+1,row}); });
  return arr; // pos начинается с 1
};

/* ---------- СМЕТА (без НДС) + месячные колонки закрытия ---------- */
C.SMETA_NDS = order => (order.meta.ndsMode==='НДС 5%') ? 5/105 : 22/122;

C.calcSmeta = function(order){
  C.calcMaterials(order);
  const ks2date=C.parseD(order.meta.ks2date);
  const months=C.months(order);                    // [{i,date,status,volsKey}]
  const P=C.positions(order);
  const frac=C.SMETA_NDS(order);
  const sm=P.map(({pos,row})=>{
    const priceNds=row.L||0;                        // E = Материалы!L (цена с коэфф, но без НДС ниже)
    // В Excel Смета!E = Материалы!L (полная цена). Стоимость G/H = I/J материалов (с коэфф, с НДС)
    const costMat=row.I, costUsl=row.J;             // G,H (руб с НДС)
    const doneQty={}, monthSum={};                  // по индексам месяцев
    let doneAll=0;
    months.forEach(m=>{
      const v=C.num((order.vols[pos]||{})['m'+m.i]);
      m.volByPos=m.m||('m'+m.i);
      doneAll+=v;
      monthSum[m.i]=v*(priceNds||0);
    });
    const qty=C.num(row.q);
    return {pos,name:row.n,u:row.u,price:priceNds,qty,costMat,costUsl,
            doneAll, doneSum:r2(doneAll*(priceNds||0)),
            rest:Math.max(qty-doneAll,0), restSum:r2(Math.max(qty-doneAll,0)*(priceNds||0)),
            monthVol:monthSum};
  });
  order.smeta=sm;
  // Итоги Смета: G3=SUM(G8:H..) стоимость; J2/K3/N2/O3 — выполнено (в ценах сметы без НДС через вычитание НДС)
  const totalCost=r2(sm.reduce((s,x)=>s+x.costMat+x.costUsl,0));       // Смета!G3 (с НДС)
  let doneToKs2=0, ytdToKs2=0;
  sm.forEach(x=>{
    // "выполнено на момент печати": сумма объёмов закрытий с датой <= Дата_КС2 × цена
    let vol=0, volYtd=0;
    months.forEach(m=>{
      if(m.date && ks2date && m.date<=ks2date){
        const v=C.num((order.vols[x.pos]||{})['m'+m.i]);
        vol+=v;
        if(m.date.getFullYear()===ks2date.getFullYear()) volYtd+=v;
      }
    });
    doneToKs2+=vol*x.price; ytdToKs2+=volYtd*x.price;
  });
  order.smetaTotals={
    total:totalCost,                                   // G3
    doneNds:r2(doneToKs2),                             // J2 (в ценах с НДС → ниже без НДС)
    doneNoNds:r2(doneToKs2 - r2(doneToKs2*frac)),      // ≈ J2 при "Без НДС" режиме
    ytd:r2(ytdToKs2)
  };
  return months;
};

/* ---------- Месяцы КС-6: даты авто (F32=Дата1, далее +31), статус ---- */
C.months = function(order){
  const res=[];
  const base=C.parseD(order.meta.ks2date);
  for(let i=0;i<9;i++){
    let d=C.parseD((order.meta.dates||[])[i]);
    if(!d && base) d=C.addDays(base,i*31);            // G32=F32+31, H32=G32+31 ...
    res.push({i,date:d,status:(order.meta.status||[])[i]==='Закрытие'?'Закрытие':'----'});
  }
  return res;
};

/* ---------- ОБЪЕМЫ ЗАКРЫТИЯ (аналог листа «Объемы закрытия») ---------- */
C.calcObems = function(order){
  C.calcSmeta(order);
  const months=C.months(order);
  const byPos={}; order.smeta.forEach(x=>byPos[x.pos]=x);
  order.obems=months.map(m=>{
    const rows=[];
    Object.keys(order.vols).forEach(k=>{
      const v=C.num(order.vols[k]['m'+m.i]);
      if(v>0){ const sp=byPos[k];
        rows.push({pos:+k,name:sp?sp.name:'',vol:v,restQty:sp?sp.rest:0}); }
    });
    return {...m,rows};
  });
  return order.obems;
};

/* ---------- КС-2 ---------- */
C.calcKS2 = function(order){
  C.calcObems(order);
  const ks2date=C.parseD(order.meta.ks2date);
  const months=C.months(order);
  // активный месяц: точное совпадение с Дата_КС2, иначе последний закрытый месяц <= Дата_КС2
  // (в Excel печатник перед расчётом ставит Дату закрытия = Дату КС-2)
  let actIdx=-1;
  months.forEach(m=>{ if(C.sameDate(m.date,ks2date)) actIdx=m.i; });
  if(actIdx<0){
    months.forEach(m=>{ if(m.status==='Закрытие' && m.date && ks2date && m.date<=ks2date) actIdx=m.i; });
    if(actIdx>=0) months[actIdx].date=ks2date;   // синхронизируем дату закрытия с датой акта
  }
  const frac = (order.meta.ndsMode==='НДС 5%')?5/105:22/122;
  const ndsLabel = order.meta.ndsMode==='НДС 5%'?'В т.ч. НДС  5%':'В т.ч. НДС 22%';

  // строки акта: заголовки (разделы) + позиции с объёмом >0 в активном месяце
  const lines=[]; let counter=0;
  order.rows.forEach((row,ri)=>{
    if(!row.t){ // потенциальный заголовок раздела
      const nextItems=order.rows.slice(ri+1);
      lines.push({type:'head',name:row.n}); return;
    }
    const pos=C.positions(order).find(p=>p.idx===ri);
    if(!pos) return;
    const v = actIdx>=0 ? C.num((order.vols[pos.pos]||{})['m'+actIdx]) : 0;
    if(v>0){
      counter++;
      const price=row.L||0;
      const sum=r2(v*price);
      lines.push({type:'item',pp:counter,pos:pos.pos,name:row.n,u:row.u,vol:v,price,sum,
                  work: row.t==='Работа'?sum:0, mat: row.t==='Материалы'?sum:0});
    }
  });
  // удалить пустые заголовки (под которыми нет строк в акте) — как IF(C26=0,SUM...) в Excel
  const out=[];
  for(let i=0;i<lines.length;i++){
    if(lines[i].type==='head'){
      let j=i+1, has=false;
      while(j<lines.length && lines[j].type==='head') j++;
      has = j<lines.length && lines[j].type==='item';
      if(has) out.push(lines[i]);
    } else out.push(lines[i]);
  }
  const sumAll=r2(out.filter(l=>l.type==='item').reduce((s,l)=>s+l.sum,0));
  // разделение Работа/Материалы — как в Excel по признаку стоимости Сметы (G>0 → материалы, H>0 → работы),
  // позиция может иметь обе составляющие одновременно
  const smap={}; (order.smeta||[]).forEach(x=>smap[x.pos]=x);
  const sumWork=r2(out.filter(l=>l.type==='item').reduce((s,l)=>{const x=smap[l.pos];return s+(x&&x.costUsl>0?l.sum:0)},0));
  const sumMat =r2(out.filter(l=>l.type==='item').reduce((s,l)=>{const x=smap[l.pos];return s+(x&&x.costMat>0?l.sum:0)},0));
  const nds=r2(sumAll*frac);
  // Номер документа E20 = COUNTIFS(даты<=Дата_КС2, суммы>0)
  let docNum=0;
  months.forEach(m=>{
    if(m.date && ks2date && m.date<=ks2date){
      const any=Object.keys(order.vols).some(k=>C.num(order.vols[k]['m'+m.i])>0);
      if(any) docNum++;
    }
  });
  // Отчётный период с/по (V19/W19 вручную, иначе 1 число / EOMONTH)
  const mf=C.parseD(order.meta.perFrom), mt=C.parseD(order.meta.perTo);
  const perFrom = mf || (ks2date? new Date(ks2date.getFullYear(),ks2date.getMonth(),1):null);
  const perTo   = mt || (ks2date? C.eomonth(ks2date):null);
  order.ks2={lines:out,total:sumAll,work:sumWork,mat:sumMat,nds,ndsLabel,docNum,
             perFrom,perTo,actIdx,year:ks2date?ks2date.getFullYear():'',
             monthName:ks2date?C.MONTHS_RU[ks2date.getMonth()]:'',
             restFormula:'Остаток по смете: '+C.money((order.smetaTotals.total||0)-(order.smetaTotals.doneNds||0))+' р.'};
  return order.ks2;
};

/* ---------- КС-3 ---------- */
C.calcKS3 = function(order){
  C.calcKS2(order);
  const t=order.smetaTotals, k=order.ks2;
  const noNds=(order.meta.ndsMode==='НДС 5%')?5/105:22/122;
  order.ks3={
    docNum:k.docNum, date:C.fmtD(C.parseD(order.meta.ks2date)),
    perFrom:C.fmtD(k.perFrom), perTo:C.fmtD(k.perTo),
    totalNds:t.total,                       // всего по смете (с НДС)
    fromStart:t.doneNds,                    // выполнено с начала работ
    ytd:t.ytd,                              // с начала года
    forPeriod:r2(k.total),                  // за отчётный период (акт КС-2)
    nds:k.nds, totalWithNds:k.total,
    labelNds:(order.meta.ndsMode==='НДС 5%'?'Сумма НДС 5%':'Сумма НДС 22%'),
    totalNoNds:r2(t.total - r2(t.total*noNds))
  };
  return order.ks3;
};

/* ---------- ЭКСПОРТ в Реестр 2 (строка «Экспорт») ---------- */
C.exportRow = function(order){
  const cust=(SEED.customers.find(c=>c.list===order.meta.customerList)||{});
  return {
    source:order.meta.source||'', ddate:order.meta.dogovorDate||'', object:order.meta.objectName||'',
    form:cust.form||'', customer:cust.name||order.meta.customerList||'', post:cust.post||'', fio:cust.fio||'',
    predmet:order.meta.obekt||'', addrObj:order.meta.stroyka||'', addrYur:cust.addr||'',
    closeSum:order.ks2?order.ks2.total:0, inn:[cust.inn,cust.kpp].filter(Boolean).join('/'),
    kurator:'', planCost:order.totals?order.totals.all:0, factCost:order.totals?order.totals.all:0,
    matPlan:order.totals?order.totals.mat:0, matFact:order.totals?order.totals.mat:0,
    uslPlan:order.totals?order.totals.work:0, uslFact:order.totals?order.totals.work:0,
    version:'V 3.2 web', path:'webapp://orders/'+encodeURIComponent(order.meta.objectName||''),
    dateRec:new Date().toISOString().slice(0,10), user:'web', employee:''
  };
};

window.Calc=C;
})();
