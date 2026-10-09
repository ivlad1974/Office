import time
from playwright.sync_api import sync_playwright

URL='http://localhost:8765/index.html'
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1200,'height':800})
    errors=[]; pg.on('pageerror', lambda e: errors.append(str(e)))
    pg.goto(URL); pg.wait_for_load_state('networkidle')

    t0=time.time(); pg.click('#orders-table [data-open]'); pg.wait_for_timeout(300)
    print('open order sec:', round(time.time()-t0,3))

    rows=pg.eval_on_selector_all('#mat-table tbody tr','e=>e.length'); print('materials rows:',rows)

    # sticky header inside scroll-wrap
    wrap=pg.query_selector('#panel-materials .scroll-wrap')
    before=pg.eval_on_selector('#mat-table thead th','e=>e.getBoundingClientRect().top')
    wrap.evaluate('el=>el.scrollTop=500'); pg.wait_for_timeout(100)
    after=pg.eval_on_selector('#mat-table thead th','e=>e.getBoundingClientRect().top')
    sc=wrap.evaluate('el=>el.scrollTop')
    print('sticky top before/after:',round(before),round(after),'scrollTop',sc)
    assert abs(after-before)<2 and sc>0, 'NOT STICKY'

    # resize col 3 — only it changes
    wb=pg.eval_on_selector_all('#mat-table thead th','els=>els.map(e=>e.style.width)')
    g=pg.query_selector('#mat-table thead th:nth-child(3) .col-grip'); gb=g.bounding_box()
    pg.mouse.move(gb['x']+3,gb['y']+gb['height']/2); pg.mouse.down(); pg.mouse.move(gb['x']+150,gb['y']+gb['height']/2,steps=4); pg.mouse.up()
    wa=pg.eval_on_selector_all('#mat-table thead th','els=>els.map(e=>e.style.width)')
    ch=[i for i,(a,c) in enumerate(zip(wb,wa)) if a!=c]; print('resize changed cols:',ch, wa[2])
    assert ch==[2], ch

    # table not stretched to window: check orders view small table
    pg.click('#btn-close-order'); pg.wait_for_timeout(400)
    tw=pg.eval_on_selector('#orders-table','t=>({w:t.offsetWidth, vw:window.innerWidth})')
    print('orders table width vs viewport:',tw)
    assert tw['w'] < tw['vw'], 'table stretched to full window?'

    # totals visible in orders list (from cache)
    tot=pg.eval_on_selector('#orders-table tbody tr td:nth-child(6)','e=>e.textContent')
    print('order total cell:',repr(tot))
    assert tot.strip()!='' , 'totals empty'

    # edit qty -> fast recalc
    pg.click('#orders-table [data-open]'); pg.wait_for_timeout(200)
    t0=time.time()
    inp=pg.query_selector('#mat-table tbody tr:nth-child(3) input[data-f="q"]')
    inp.click(); inp.fill('7'); inp.press('Tab'); pg.wait_for_timeout(600)
    print('edit+recalc sec:',round(time.time()-t0,3))

    # dictionary autofill still works: type known material name into an empty row
    # find a known dict name
    name=pg.evaluate("__APP.store.dictMat.find(d=>d.n && d.n.length>3).n")
    ins=pg.query_selector('#mat-table tbody tr:first-child [data-ins-after]')
    ins.click(); pg.wait_for_timeout(700)
    t0=time.time()
    ni=pg.query_selector('#mat-table tbody tr:nth-child(2) input[data-f="n"]')
    ni.click(); ni.fill(name); ni.dispatch_event('change'); pg.wait_for_timeout(600)
    print('name set + autofill sec:',round(time.time()-t0,3))
    u=pg.eval_on_selector('#mat-table tbody tr:nth-child(2) input[data-f="u"]','e=>e.value')
    print('unit autofilled:',repr(u))

    # sprav tab
    t0=time.time(); pg.click('.tab[data-view="sprav"]'); pg.wait_for_selector('#sprav-m-tbody tr'); pg.wait_for_timeout(200)
    print('sprav render sec:',round(time.time()-t0,3))
    # r2 sticky + horizontal scroll
    pg.click('.tab[data-view="reestr2"]'); pg.wait_for_timeout(300)
    w2=pg.query_selector('#view-reestr .scroll-wrap, #view-reestr2 .scroll-wrap')
    # ks6
    pg.click('.tab[data-view="editor"]'); pg.wait_for_timeout(200)
    t0=time.time(); pg.click('.sheet-tab[data-sheet="ks6"]'); pg.wait_for_timeout(300)
    print('ks6 switch sec:',round(time.time()-t0,3))
    # edit volume
    vol=pg.query_selector('#ks6-table tbody input.vol:not([disabled])')
    if vol:
        t0=time.time(); vol.click(); vol.fill('1'); vol.press('Tab'); pg.wait_for_timeout(500)
        print('ks6 vol edit sec:',round(time.time()-t0,3))
    # smeta sheet
    t0=time.time(); pg.click('.sheet-tab[data-sheet="smeta"]'); pg.wait_for_timeout(300)
    print('smeta switch sec:',round(time.time()-t0,3))
    # ks2 doc sheet (full recalc on demand)
    t0=time.time(); pg.click('.sheet-tab[data-sheet="ks2"]'); pg.wait_for_timeout(500)
    print('ks2 switch sec:',round(time.time()-t0,3))
    has_doc=pg.eval_on_selector('#ks2-doc','e=>e.textContent.includes("АКТ")')
    print('ks2 doc ok:',has_doc)

    print('PAGE ERRORS:',errors); assert not errors
    b.close()
print('ALL OK')
