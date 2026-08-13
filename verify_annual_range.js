const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.goto(process.env.DASHBOARD_URL || 'http://127.0.0.1:1477/index.html', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction(() => typeof charts !== 'undefined' && charts.cAnnual);
  const result = await page.evaluate(() => {
    const rangeSumOf = (history, y, f, t) => { const h = history[y] || {}; let s = 0; for (let m = f; m <= t; m++) s += Number(h[String(m)] || 0); return s; };
    selChans = ['MAKRO', "LOTUS'"]; cur = 'MAKRO'; kpiPick = null;
    activeCat = null; activeSku = null; selFrom = 1; selTo = 8; renderAll();
    const totalJanAug = charts.cAnnual.data.datasets[0].data.slice();
    const years = charts.cAnnual.data.labels.slice();
    const totalJanAugExpected = years.map(y => rangeSumOf(C().history, y, 1, 8));
    // selecting a single specific month must reflect that month in the total graph
    selFrom = 3; selTo = 3; renderAll();
    const totalMar = charts.cAnnual.data.datasets[0].data.slice();
    const totalMarExpected = years.map(y => rangeSumOf(C().history, y, 3, 3));
    activeCat = '\u0e02\u0e49\u0e32\u0e27\u0e15\u0e31\u0e07'; selFrom = 1; selTo = 8; renderAll();
    const categoryJanAug = charts.cAnnual.data.datasets[0].data.slice();
    const categoryExpected = years.map(y => {
      const h = C().history_by_cat[activeCat][y] || {};
      return Array.from({ length: 8 }, (_, i) => Number(h[String(i + 1)] || 0)).reduce((a, b) => a + b, 0);
    });
    selFrom = 3; selTo = 5; renderAll();
    const categoryMarMay = charts.cAnnual.data.datasets[0].data.slice();
    const sku = Object.keys(C().monthly_by_sku).find(x => C().monthly_by_sku[x] && C().history_by_sku[x]);
    activeSku = sku; renderAll();
    const skuMarMay = charts.cAnnual.data.datasets[0].data.slice();
    const skuExpected = years.map(y => {
      const h = C().history_by_sku[sku][y] || {};
      return [3, 4, 5].map(m => Number(h[String(m)] || 0)).reduce((a, b) => a + b, 0);
    });
    return { years, totalJanAug, totalJanAugExpected, totalMar, totalMarExpected, categoryJanAug, categoryExpected, categoryMarMay, skuMarMay, skuExpected };
  });
  const close = (a, b) => a.length === b.length && a.every((x, i) => Math.abs(x - b[i]) < 0.001);
  const y2025 = result.years.indexOf('2025');
  const y2023 = result.years.indexOf('2023');
  const y2024 = result.years.indexOf('2024');
  const assertions = {
    total_uses_selected_range: close(result.totalJanAug, result.totalJanAugExpected),
    total_single_month_specific: close(result.totalMar, result.totalMarExpected),
    total_changes_with_month_filter: !close(result.totalJanAug, result.totalMar),
    total_2025_jan_aug_below_full_year: result.totalJanAug[y2025] < 245.64,
    total_2023_visible_in_multi_channel: y2023 >= 0 && result.totalJanAug[y2023] > 0 && result.totalMar[y2023] > 0,
    total_2024_visible_in_multi_channel: y2024 >= 0 && result.totalJanAug[y2024] > 0 && result.totalMar[y2024] > 0,
    category_uses_selected_range: close(result.categoryJanAug, result.categoryExpected),
    category_2025_jan_aug_is_78_1: Math.abs(result.categoryJanAug[y2025] - 78.09) < 0.011,
    category_changes_with_month_filter: !close(result.categoryJanAug, result.categoryMarMay),
    sku_uses_selected_range: close(result.skuMarMay, result.skuExpected),
    no_js_errors: errors.length === 0
  };
  console.log(JSON.stringify({ result, assertions, errors }, null, 2));
  await browser.close();
  if (Object.values(assertions).some(x => !x)) process.exit(1);
})().catch(error => { console.error(error); process.exit(1); });
