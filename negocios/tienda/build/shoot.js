const { chromium } = require('playwright'); const fs = require('fs'); const path = require('path');
(async () => {
  const b = await chromium.launch(); const ctx = await b.newContext({ viewport: { width: 1500, height: 1125 }, deviceScaleFactor: 2 });
  const root = path.resolve('../productos');
  for (const slug of fs.readdirSync(root)) {
    const src = path.join(root, slug, 'images', 'src'); if (!fs.existsSync(src)) continue;
    for (const f of fs.readdirSync(src).filter(x => x.endsWith('.html')).sort()) {
      const p = await ctx.newPage(); await p.goto('file://' + path.join(src, f)); await p.evaluate(() => document.fonts.ready); await p.waitForTimeout(300);
      await p.screenshot({ path: path.join(root, slug, 'images', f.replace('.html', '.png')) }); await p.close();
    }
  }
  await b.close(); console.log('done');
})();
