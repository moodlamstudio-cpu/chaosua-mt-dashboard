const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.goto(process.env.DASHBOARD_URL || 'http://127.0.0.1:1477/index.html', { waitUntil: 'networkidle' });
  await page.waitForFunction(() => document.querySelectorAll('#catPerfTable tbody .perf-row').length > 0);
  const result = await page.evaluate(() => {
    const catWrap = document.querySelector('#catPerfTable').parentElement;
    const skuWrap = document.querySelector('#skuTable').parentElement;
    const cards = [...document.querySelectorAll('.card')];
    const catCard = cards.find(x => x.textContent.includes('Category Performance'));
    const skuCard = cards.find(x => x.textContent.includes('All SKU Performance'));
    const values = table => [...table.querySelectorAll('tfoot td')].map(td => td.textContent.trim());
    return {
      categoryRows: document.querySelectorAll('#catPerfTable tbody .perf-row').length,
      categoryMapCount: Object.keys(C().monthly_by_cat || {}).length,
      catOverflowY: getComputedStyle(catWrap).overflowY,
      catMaxHeight: getComputedStyle(catWrap).maxHeight,
      skuMaxHeight: getComputedStyle(skuWrap).maxHeight,
      catWidth: catCard.getBoundingClientRect().width,
      skuWidth: skuCard.getBoundingClientRect().width,
      wrapWidth: document.querySelector('.wrap').getBoundingClientRect().width,
      catTotal: values(document.querySelector('#catPerfTable')),
      skuTotal: values(document.querySelector('#skuTable')),
      totalClickable: !!document.querySelector('tfoot .perf-row')
    };
  });
  const assertions = {
    all_categories_visible: result.categoryRows === result.categoryMapCount,
    category_has_no_vertical_scroll: result.catOverflowY !== 'scroll' && result.catMaxHeight === 'none',
    sku_has_bounded_vertical_scroll: parseFloat(result.skuMaxHeight) > 0,
    cards_are_full_width_rows: result.catWidth > result.wrapWidth * .9 && result.skuWidth > result.wrapWidth * .9,
    category_total_is_100: result.catTotal.at(-1) === '100.0%',
    sku_total_is_100: result.skuTotal.at(-1) === '100.0%',
    totals_are_not_clickable: !result.totalClickable,
    no_js_errors: errors.length === 0
  };
  await page.screenshot({ path: 'performance_layout_desktop.png', fullPage: true });
  await page.setViewportSize({ width: 390, height: 844 });
  await page.screenshot({ path: 'performance_layout_mobile.png', fullPage: true });
  console.log(JSON.stringify({ result, assertions, errors }, null, 2));
  await browser.close();
  if (Object.values(assertions).some(x => !x)) process.exit(1);
})();
