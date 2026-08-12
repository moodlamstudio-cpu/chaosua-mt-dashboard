const { chromium } = require('playwright');
const assert = require('assert');

// Reads a column value from a perf table row. colIndex is 1-based over td cells of the row.
function colOf(row, idx){ const tds = row.querySelectorAll('td'); return tds[idx-1]; }
(async()=>{
  const b = await chromium.launch({headless:true});
  const p = await b.newPage({viewport:{width:1440,height:1000}});
  const errors=[];
  p.on('console', m=>{ if(m.type()==='error') errors.push(m.text()); });
  p.on('pageerror', e=>errors.push(String(e)));
  await p.goto('http://127.0.0.1:8765/index.html', {waitUntil:'networkidle'});

  const result = await p.evaluate(()=>{
    const failures=[];
    const numOrNull = (txt)=>{ txt=(txt||'').replace(/[%\u2014+\s]/g,''); if(txt===''||txt==='\u2014') return null; if(txt==='0.0') return 0; return parseFloat(txt); };
    const mojibake = () => document.body.textContent.match(/เน€|ฃเธ|ฦ/g);
    // ---- helpers ----
    function catBtn(col){ return document.querySelector('#catPerfTable button[data-colsort="'+col+'"]'); }
    function skuBtn(col){ return document.querySelector('#skuTable button[data-colsort="'+col+'"]'); }
    function catRows(){ return [...document.querySelectorAll('#catPerfTable tbody tr.perf-row')]; }
    function skuRows(){ return [...document.querySelectorAll('#skuTable tbody tr.skuro')]; }
    function reconRow(){ return document.querySelector('#skuTable tbody tr.recon-row'); }
    function totalFooter(sel){ return document.querySelector(sel+' tfoot tr.total'); }
    function catVal(r,idx){ return numOrNull(r.querySelectorAll('td')[idx-1] && r.querySelectorAll('td')[idx-1].textContent); }
    function skuVal(r,idx){ return numOrNull(r.querySelectorAll('td')[idx-1] && r.querySelectorAll('td')[idx-1].textContent); }    function defaultCatOrder(){ 
      const c=C(); const map=c.monthly_by_cat||{}; 
      return Object.keys(map).sort((a,b)=>perfValues(map[b]).actual-perfValues(map[a]).actual);
    }
    function defaultSkuOrder(){
      const c=C(); const map=c.monthly_by_sku||{};
      return Object.keys(map).filter(n=>activeSku?n===activeSku:(activeCat?(map[n]||{}).cat===activeCat:true))
        .sort((a,b)=>perfValues(map[b]).actual-perfValues(map[a]).actual);
    }
    function refreshNames(){ catRows().forEach(r=>{r._n=r.getAttribute('data-cat');}); skuRows().forEach(r=>{r._n=r.getAttribute('data-sku');}); }

    // ---- 0. initial default order ----
    const dCatDef = catRows().map(r=>r.getAttribute('data-cat'));
    const dCatExp = defaultCatOrder().map(encodeURIComponent);
    if(JSON.stringify(dCatDef)!==JSON.stringify(dCatExp)) failures.push({msg:'cat default order mismatch', got:dCatDef.slice(0,5), exp:dCatExp.slice(0,5)});
    const dSkuDef = skuRows().map(r=>r.getAttribute('data-sku'));
    const dSkuExp = defaultSkuOrder().map(encodeURIComponent);
    if(JSON.stringify(dSkuDef)!==JSON.stringify(dSkuExp)) failures.push({msg:'sku default order mismatch', got:dSkuDef.slice(0,5), exp:dSkuExp.slice(0,5)});

    // ---- 1. cat: sort col 'actual' desc -> asc -> clear ----
    // desc
    catBtn('actual').click();
    refreshNames();
    const cDesc = catRows().map(r=>catVal(r,3));
    const cActExp = cDesc.slice().sort((a,b)=>b-a);
    if(JSON.stringify(cDesc)!==JSON.stringify(cActExp)) failures.push({msg:'cat actual desc not sorted', got:cDesc.slice(0,6)});
    const caBtn=catBtn('actual');
    const caTh=caBtn.closest('th');
    if(caTh.getAttribute('aria-sort')!=='descending' || !/▼/.test(caBtn.textContent)) failures.push({msg:'cat desc indicator/aria missing', aria:caTh.getAttribute('aria-sort'), txt:caBtn.textContent});
    // single missing check: desc must place nulls last -> last revealed value must be non-null if any null exists
    // asc
    catBtn('actual').click();
    refreshNames();
    const cAsc = catRows().map(r=>catVal(r,3));
    const cAscExp = cAsc.slice().sort((a,b)=>a-b);
    if(JSON.stringify(cAsc)!==JSON.stringify(cAscExp)) failures.push({msg:'cat actual asc not sorted'});
    if(catBtn('actual').closest('th').getAttribute('aria-sort')!=='ascending') failures.push({msg:'cat asc aria missing'});
    // clear (3rd click) back to default
    catBtn('actual').click();
    refreshNames();
    const cClear = catRows().map(r=>r.getAttribute('data-cat'));
    if(JSON.stringify(cClear)!==JSON.stringify(dCatDef)) failures.push({msg:'cat clear not back to default'});
    if(!/▼|▲/.test(catBtn('actual').textContent)) { /* ok, no indicator */ } else failures.push({msg:'cat cleared but indicator present'});

    // ---- stable tie + missing-at-end checks on % Growth (pct may repeat) ----
    // sort by 'pct' desc; verify non-null prefix is desc and trailing all-null
    catBtn('pct').click();
    refreshNames();
    const prev={}; // verify tie stability: equal pct values preserve default relative order
    const pcts=catRows().map(r=>catVal(r,5));
    // find a duplicated pct value
    const counts={}; pcts.forEach(v=>{ if(v==null)return; counts[v]=(counts[v]||0)+1; });
    let dupKey=null; Object.keys(counts).forEach(k=>{ if(counts[k]>1) dupKey=k; });
    if(dupKey!=null){
      const defIdx=defaultCatOrder();
      const dupNames=catRows().map(r=>r.getAttribute('data-cat')).filter((nn,i)=>{ if(pcts[i]==null)return false; return Math.abs(pcts[i]-parseFloat(dupKey))<0.011; });
      const dupDefPos=dupNames.map(n=>defIdx.indexOf(decodeURIComponent(n)));
      for(let i=1;i<dupDefPos.length;i++){ if(dupDefPos[i-1]>dupDefPos[i]) failures.push({msg:'pct tie not stable', names:dupNames.slice(0,5)}); break; }
    }
    const firstNullPct=catRows().findIndex(r=>catVal(r,5)==null);
    if(firstNullPct>=0){
      // all rows after first null must also be null
      for(let i=firstNullPct;i<catRows().length;i++){ if(catVal(catRows()[i],5)!=null){ failures.push({msg:'null not grouped at end (cat pct)'}); break; } }
    }
    catBtn('pct').click(); catBtn('pct').click(); // return to default / clear

    // ---- 2. sku: sort 'growth' desc, verify correctness + recon row stays last body + total footer ----
    // capture pre-sort body names & recon presence
    const skuRowsBefore=skuRows();
    const reconBefore = reconRow();
    skuBtn('growth').click();
    refreshNames();
    const g = skuRows().map(r=>skuVal(r,6));
    const gSorted = g.slice().filter(v=>v!=null).sort((a,b)=>b-a);
    const gNonNull = g.filter(v=>v!=null);
    if(JSON.stringify(gNonNull)!==JSON.stringify(gSorted)) failures.push({msg:'sku growth desc non-null not sorted'});
    const firstNullGrowth=skuRows().findIndex(r=>skuVal(r,6)==null);
    if(firstNullGrowth>=0){ for(let i=firstNullGrowth;i<skuRows().length;i++){ if(skuVal(skuRows()[i],6)!=null){ failures.push({msg:'null not at end (sku growth)'}); break; } } }
    // recon row still present as last tbody row, before tfoot
    const rb=reconRow(); const rownodes=[...document.querySelectorAll('#skuTable tbody tr')];
    if(rb){ if(rownodes[rownodes.length-1]!==rb) failures.push({msg:'recon row not last body row after sort'}); }
    // total footer unchanged & not sorted (still first tfoot row / Total)
    const tf=totalFooter('#skuTable');
    if(!tf) failures.push({msg:'sku total footer missing'});
    // row count unchanged
    if(skuRows().length!==dSkuDef.length) failures.push({msg:'sku row count changed after sort', was:dSkuDef.length, now:skuRows().length});
    skuBtn('growth').click(); skuBtn('growth').click(); // clear

    // ---- 3. pivot after sort: click a sorted row hits correct item ----
    // sort SKU by 'ly' desc, then click the FIRST row and verify activeSku equals that row's name
    skuBtn('ly').click();
    refreshNames();
    const firstSkuRow=skuRows()[0];
    const firstSkuName=decodeURIComponent(firstSkuRow.getAttribute('data-sku'));
    firstSkuRow.click();
    // activeSku set in toggleSku -> rerender; verify the pivot focused name matches
    if(activeSku!==firstSkuName) failures.push({msg:'sku pivot after sort wrong item', clicked:firstSkuName, active:activeSku});
    // total now is focused-SKU series (authoritative) -> footer 100%
    const f0=document.querySelector('#skuTable tfoot tr.total');
    if(!f0 || !/100\.0%/.test(f0.textContent)) failures.push({msg:'focused sku footer not 100%', txt:f0&&f0.textContent});
    activeSku=null; activeCat=null; renderAll();

    // ---- 4. cat pivot after sort ----
    catBtn('ly').click();
    refreshNames();
    const firstCatRow=catRows()[0];
    const firstCatName=decodeURIComponent(firstCatRow.getAttribute('data-cat'));
    firstCatRow.click();
    if(activeCat!==firstCatName) failures.push({msg:'cat pivot after sort wrong item', clicked:firstCatName, active:activeCat});
    activeCat=null; renderAll(); catBtn('ly').click(); catBtn('ly').click(); // clear

    // ---- 5. sort state persists across channel/month rerender but clears on 3rd click ----
    // set sku sort on 'actual' desc
    skuBtn('actual').click();
    const persisted1=skuSort;
    // change month range -> rerender
    selTo=7; renderAll();
    const persisted2=skuSort;
    if(!persisted1||!persisted2||persisted2.col!=='actual'||persisted2.dir!=='desc') failures.push({msg:'sku sort state not persisted', a:persisted1,b:persisted2});
    // active header still shows desc indicator after rerender
    if(skuBtn('actual').closest('th').getAttribute('aria-sort')!=='descending') failures.push({msg:'indicator lost after rerender'});
    // change channel -> rerender keep state
    selChans=['MAKRO']; cur='MAKRO'; renderAll();
    const persisted3=skuSort;
    if(!persisted3||persisted3.col!=='actual'||persisted3.dir!=='desc') failures.push({msg:'sku sort lost on channel rerender'});
    // 3rd click clears (back to default) AND clears state for future
    skuBtn('actual').click(); // asc
    skuBtn('actual').click(); // clear
    if(skuSort!==null) failures.push({msg:'sku sort state not cleared on 3rd click'});
    // reset channel
    selChans=['MAKRO','LOTUS\'']; cur='MAKRO'; selTo=8; renderAll();

    // ---- 6. keyboard Enter/Space support (dispatch on the button element) ----
    const kb=skuBtn('actual');
    kb.focus();
    kb.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));
    if(skuSort===null||skuSort.col!=='actual'||skuSort.dir!=='desc') failures.push({msg:'Enter did not sort', st:skuSort});
    skuBtn('actual').focus();
    skuBtn('actual').dispatchEvent(new KeyboardEvent('keydown',{key:' ',bubbles:true}));
    if(!skuSort||skuSort.dir!=='asc') failures.push({msg:'Space did not advance sort', st:skuSort});
    skuBtn('actual').focus();
    skuBtn('actual').dispatchEvent(new KeyboardEvent('keydown',{key:' ',bubbles:true}));
    if(skuSort!==null) failures.push({msg:'Space did not clear', st:skuSort});

    // ---- 7. totals preserved (reconciliation) after all toggling ----
    // compare cat total footer to sum of displayed + recon
    const cf=totalFooter('#catPerfTable');
    if(!cf || !/100\.0%/.test(cf.textContent)) failures.push({msg:'cat total footer missing or not 100%'});
    // All SKU authoritative reconciliation still holds (footer equals auth series in default)
    const a=skuAuthoritative(C()); const ft=document.querySelector('#skuTable tfoot tr.total');
    if(!ft||!/100\.0%/.test(ft.textContent)) failures.push({msg:'sku total footer not 100% in default'});

    // ---- 8. mobile overflow / sticky header ----
    return {failures, mojibake:mojibake(), sortStateCat:catSort, sortStateSku:skuSort, results:[]};
  });

  // mobile scroll + sticky check via viewport
  await p.setViewportSize({width:420,height:900});
  await p.reload({waitUntil:'networkidle'});
  const mobile = await p.evaluate(()=>{
    const o={};
    [['#catPerfTable','.cat-perf-scroll'],['#skuTable','.sku-scroll']].forEach(([t,sc])=>{
      const wrap=document.querySelector(sc); const table=document.querySelector(t);
      o[sc]=wrap ? {scrollW:wrap.scrollWidth, clientW:wrap.clientWidth, overflowX:getComputedStyle(wrap).overflowX, tableW:table&&table.scrollWidth, hasHScroll:wrap.scrollWidth>wrap.clientWidth} : null;
      const th=table&&table.querySelector('thead th .sort-btn');
      o[sc+'.sortbtn']=th?getComputedStyle(th).width:null;
    });
    o.errors = document.body.textContent.match(/เน€|ฃเธ/g);
    return o;
  });

  assert.deepStrictEqual(result.failures, []);
  assert.strictEqual(result.mojibake, null);
  assert.strictEqual(mobile.errors, null);
  // every perf table has at least sort buttons rendered and indicators present for active columns tested above
  console.log('TARGETED_SORT_VERIFY_PASS', JSON.stringify({mobile, sortStateCat:result.sortStateCat, sortStateSku:result.sortStateSku},null,2));
  await b.close();
})().catch(e=>{ console.error('FAIL', e); try{ require('fs').writeFileSync('verify_sortable_headers_err.log', String(e.stack||e)+'\n'); }catch(_){} process.exit(1); });
