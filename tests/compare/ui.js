/* Page checks for Deck Compare. Run check.js first, it draws the decks.

   Needs   npm packages playwright and jszip (JSZip is served locally so
           the check works offline).
   Run     node tests/compare/ui.js
   Shots   land in $WORK/shots */
'use strict';
const path = require('path'), fs = require('fs'), os = require('os'), http = require('http');
function need(name) {
  try { return require(name); } catch (e) {}
  if (process.env.NODE_MODULES_DIR) return require(path.join(process.env.NODE_MODULES_DIR, name));
  throw new Error('Cannot find ' + name);
}
const { chromium } = need('playwright');
const jszipFile = require.resolve(path.join(process.env.NODE_MODULES_DIR || '', 'jszip/dist/jszip.min.js'));
const root = path.resolve(__dirname, '..', '..');
const work = process.env.WORK || path.join(os.tmpdir(), 'deck-compare-checks');
const shots = path.join(work, 'shots');
fs.mkdirSync(shots, { recursive: true });
const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml' };

const server = http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split('?')[0]);
  if (p === '/') p = '/index.html';
  const file = path.join(root, p);
  if (!file.startsWith(root) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream' });
  fs.createReadStream(file).pipe(res);
});

let fails = 0;
function ok(c, m) { console.log((c ? '  ok   ' : '  FAIL ') + m); if (!c) fails++; }

server.listen(8913, '127.0.0.1', async () => {
  const browser = await chromium.launch();
  try {
    for (const vp of [{ width: 1400, height: 1000, tag: 'desktop' }, { width: 390, height: 844, tag: 'phone' }]) {
      const page = await browser.newPage({ viewport: { width: vp.width, height: vp.height } });
      const errs = [];
      page.on('pageerror', e => errs.push(e.message));
      await page.route(/cdnjs\.cloudflare\.com\/ajax\/libs\/jszip/, r => r.fulfill({ path: jszipFile, contentType: 'text/javascript' }));
      await page.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
      await page.goto('http://127.0.0.1:8913/deck-compare.html');
      // both decks into A at once, in the wrong order, to check the name sort
      await page.setInputFiles('#file-a', [path.join(work, 'v2.pptx'), path.join(work, 'v1.pptx')]);
      await page.waitForSelector('#rows .srow', { timeout: 20000 });
      console.log(vp.tag);
      ok(/v1\.pptx/.test(await page.textContent('#drop-a')) && /v2\.pptx/.test(await page.textContent('#drop-b')), 'two files sorted into A and B');
      const visible = await page.$$eval('#rows .srow:not(.off)', els => els.length);
      ok(visible === 8, 'eight slide rows shown (' + visible + ')');
      ok((await page.textContent('#rows')).includes('"yaer" → "year"'), 'typo fix listed');
      await page.screenshot({ path: path.join(shots, vp.tag + '-all.png'), fullPage: true });
      await page.click('.fbtn[data-f="critical"]');
      const crit = await page.$$eval('#rows .srow:not(.off)', els => els.map(e => e.querySelector('.ttl').textContent));
      ok(crit.includes('Sizzle reel') && crit.includes('Agenda') && !crit.includes('Annual Kick-off 2026'), 'critical filter keeps video and hidden slide, drops the date edit');
      await page.click('.fbtn[data-f="all"]');
      await page.click('#expand');
      await page.screenshot({ path: path.join(shots, vp.tag + '-expanded.png'), fullPage: true });
      const wide = await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth + 1);
      ok(!wide, 'no sideways scroll');
      await page.click('#swap');
      await page.waitForFunction(() => /v2\.pptx/.test(document.getElementById('drop-a').textContent) && document.querySelector('#rows .srow'));
      ok((await page.textContent('#rows')).includes('"year" → "yaer"'), 'swap reverses the report');
      ok(!errs.length, 'no page errors ' + errs.join(' | '));
      await page.close();
    }
    const home = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
    await home.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
    await home.goto('http://127.0.0.1:8913/');
    const tile = await home.$('.compare-tool');
    ok(!!tile, 'tile on the launchpad');
    await home.$eval('.other-tools', el => el.scrollIntoView());
    await (await home.$('.other-tools')).screenshot({ path: path.join(shots, 'home-tiles.png') });
    const heights = await home.$$eval('.tools-grid .tool-link', els => [...new Set(els.map(e => Math.round(e.getBoundingClientRect().height)))]);
    ok(heights.length === 1, 'tiles all the same height');
  } finally {
    await browser.close();
    server.close();
  }
  console.log(fails ? fails + ' failed' : 'all passed');
  process.exit(fails ? 1 : 0);
});
