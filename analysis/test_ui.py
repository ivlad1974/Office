import time, json
from playwright.sync_api import sync_playwright

URL='http://localhost:8765/index.html'
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1200,'height':800})
    errors=[]
    pg.on('pageerror', lambda e: errors.append(str(e)))
    pg.goto(URL); pg.wait_for_load_state('networkidle')

    # 1) открыть заказ — замер времени
    t0=time.time()
    pg.click('#orders-table [data-open]')
    pg.wait_for_selector('#panel-materials:not(.hidden)')
    t_open=time.time()-t0
    print('open order sec:', round(t_open,3))

    rows=pg.eval_on_selector_all('#mat-table tbody tr','els=>els.length')
    print('materials rows:', rows)

    # таблица не должна быть растянута на всё окно (если колонок много и они шире окна — норм; проверим width:auto поведение иначе)
    tw=pg.eval_on_selector('#mat-table','t=>t.offsetWidth')
    print('table width:', tw)

    # 2) sticky шапка: скролл внутри .scroll-wrap — thead остаётся видимым
    wrap=pg.query_selector('#panel-materials .scroll-wrap')
    box=wrap.bounding_box()
    first_th_top_before=pg.eval_on_selector('#mat-table thead th','e=>e.getBoundingClientRect().top')
    wrap.evaluate('el=>el.scrollTop=400')
    pg.wait_for_timeout(100)
    first_th_top_after=pg.eval_on_selector('#mat-table thead th','e=>e.getBoundingClientRect().top')
    scrolled=wrap.evaluate('el=>el.scrollTop')
    print('sticky: top before/after =', round(first_th_top_before), round(first_th_top_after), 'scrollTop=', scrolled)
    assert abs(first_th_top_after-first_th_top_before)<2 and scrolled>0, 'thead not sticky'

    # 3) ресайз колонки: остальные столбцы не меняют ширину
    widths_before=pg.eval_on_selector_all('#mat-table thead th','els=>els.map(e=>e.style.width||e.offsetWidth+"")')
    grip=pg.query_selector('#mat-table thead th:nth-child(3) .col-grip')
    gb=grip.bounding_box()
    pg.mouse.move(gb['x']+gb['width']/2, gb['y']+gb['height']/2)
    pg.mouse.down()
    pg.mouse.move(gb['x']+120, gb['y']+gb['height']/2, steps=5)
    pg.mouse.up()
    pg.wait_for_timeout(100)
    widths_after=pg.eval_on_selector_all('#mat-table thead th','els=>els.map(e=>e.style.width||e.offsetWidth+"")')
    print('before:', widths_before[:6]); print('after :', widths_after[:6])
    changed=[i for i,(a,c) in enumerate(zip(widths_before,widths_after)) if a!=c]
    print('changed cols:', changed)
    assert changed==[2], f'other columns changed: {changed}'

    # 4) редактирование ячейки не должно лабить (recalcAll)
    t0=time.time()
    inp=pg.query_selector('#mat-table tbody tr:nth-child(3) input[data-f="q"]')
    inp.click(); inp.fill('7'); 
    pg.wait_for_timeout(50)
    inp.press('Tab')
    pg.wait_for_timeout(300)
    print('edit+recalc sec:', round(time.time()-t0,3))

    # 5) справочник (1178 строк) рендер + автоширина
    t0=time.time()
    pg.click('.tab[data-view="sprav"]')
    pg.wait_for_selector('#sprav-m-tbody tr')
    print('sprav render sec:', round(time.time()-t0,3))
    mrows=pg.eval_on_selector_all('#sprav-m-tbody tr','e=>e.length')
    print('sprav material rows:', mrows)

    # 6) реестр 2
    t0=time.time()
    pg.click('.tab[data-view="reestr2"]')
    pg.wait_for_timeout(200)
    print('r2 render sec:', round(time.time()-t0,3))

    # 7) КС-6 лист
    pg.click('.tab[data-view="editor"]')
    t0=time.time()
    pg.click('.sheet-tab[data-sheet="ks6"]')
    pg.wait_for_timeout(200)
    print('ks6 render sec:', round(time.time()-t0,3))

    print('PAGE ERRORS:', errors)
    assert not errors, errors
    b.close()
print('ALL OK')
