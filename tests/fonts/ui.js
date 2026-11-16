/* Page checks for Font Extractor. Run check.js first, it draws the decks.

   Needs   npm packages playwright and jszip (JSZip is served locally so
           the check works offline).
   Run     node tests/fonts/ui.js
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
const work = process.env.WORK || path.join(os.tmpdir(), 'font-extractor-checks');
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

server.listen(8914, '127.0.0.1', async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  try {
    for (const vp of [{ width: 1400, height: 1000, tag: 'desktop' }, { width: 390, height: 844, tag: 'phone' }]) {
      const page = await browser.newPage({ viewport: { width: vp.width, height: vp.height }, acceptDownloads: true });
      const errs = [];
      page.on('pageerror', e => errs.push(e.message));
      await page.route(/cdnjs\.cloudflare\.com\/ajax\/libs\/jszip/, r => r.fulfill({ path: jszipFile, contentType: 'text/javascript' }));
      await page.route(/fonts\.(googleapis|gstatic)\.com|cloudflareinsights/, r => r.abort());
      await page.goto('http://127.0.0.1:8914/font-extractor.html');
      console.log(vp.tag);

      await page.setInputFiles('#file', path.join(work, 'subset.pptx'));
      await page.waitForSelector('.card', { timeout: 20000 });
      await page.waitForFunction(() => !document.querySelector('.sample .ph') || [...document.querySelectorAll('.sample .ph')].every(e => !/Loading/.test(e.textContent)));
      const pills = await page.$$eval('.card .card-head .pill', els => els.map(e => e.textContent));
      ok(pills.length === 5, 'five font cards (' + pills.length + ')');
      ok(pills.filter(p => p === 'Subset').length === 3 && pills.includes('Check') && pills.includes('Whole font'), 'three subsets, one check, one whole: ' + pills.join(', '));
      ok(/3 subsets/.test(await page.textContent('.deck-counts')), 'deck bar counts the subsets');
      ok(/only a subset/.test(await page.textContent('.deck-bar')), 'subset warning shown');
      ok((await page.textContent('.card.subset .chars')).includes('26HWdelor'), 'subset card lists the characters it holds');
      ok(/Same file as the Regular slot/.test(await page.textContent('.cards')), 'duplicate slot called out');
      const drawn = await page.$$eval('.sample .big', els => els.length);
      ok(drawn === 5, 'all five previews drawn in their own font (' + drawn + ')');
      await page.screenshot({ path: path.join(shots, vp.tag + '-subset.png'), fullPage: true });

      let [dl] = await Promise.all([page.waitForEvent('download'), page.click('.card.subset .dl-font')]);
      ok(/-SUBSET\.ttf$/.test(dl.suggestedFilename()), 'subset downloads as ' + dl.suggestedFilename());
      const ttf = fs.readFileSync(await dl.path());
      ok(ttf.readUInt32BE(0) === 0x00010000, 'download is a TrueType file');
      [dl] = await Promise.all([page.waitForEvent('download'), page.click('.dl-zip')]);
      ok(dl.suggestedFilename() === 'subset fonts.zip', 'zip named ' + dl.suggestedFilename());
      const z = await need('jszip').loadAsync(fs.readFileSync(await dl.path()));
      ok(Object.keys(z.files).length === 4, 'zip holds four files, the duplicate slot once: ' + Object.keys(z.files).join(', '));
      [dl] = await Promise.all([page.waitForEvent('download'), page.click('.dl-strip')]);
      ok(dl.suggestedFilename() === 'subset (fonts removed).pptx', 'stripped deck named ' + dl.suggestedFilename());
      const sz = await need('jszip').loadAsync(fs.readFileSync(await dl.path()));
      ok(!Object.keys(sz.files).some(n => /fonts\//.test(n)), 'stripped deck has no font parts');

      await page.setInputFiles('#file', [path.join(work, 'full.pptx'), path.join(work, 'none.pptx')]);
      await page.waitForFunction(() => document.querySelectorAll('.deck').length === 3 && !document.querySelector('.loading-note'));
      ok(/No embedded fonts/.test(await page.textContent('#results')), 'deck with no fonts says so');
      ok((await page.$$eval('.card.full', els => els.length)) === 3, 'whole fonts shown green');
      const wide = await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth + 1);
      ok(!wide, 'no sideways scroll');
      await page.screenshot({ path: path.join(shots, vp.tag + '-all.png'), fullPage: true });

      if (process.env.REAL_DECKS) {
        const real = fs.readdirSync(process.env.REAL_DECKS).filter(n => /\.pptx$/i.test(n)).map(n => path.join(process.env.REAL_DECKS, n));
        await page.reload();
        await page.setInputFiles('#file', real);
        await page.waitForFunction(n => document.querySelectorAll('.deck').length === n && !document.querySelector('.loading-note'), real.length);
        await page.waitForTimeout(500);
        ok(!(await page.$('.card.failed')), 'real PowerPoint decks unwrap in the page');
        await page.screenshot({ path: path.join(shots, vp.tag + '-real.png'), fullPage: true });
      }
      ok(!errs.length, 'no page errors ' + errs.join(' | '));
      await page.close();
    }
    const home = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
    await home.route(/fonts\.(googleapis|gstatic)\.com|cloudflareinsights/, r => r.abort());
    await home.goto('http://127.0.0.1:8914/');
    const tile = await home.$('.font-tool');
    ok(!!tile && /Font Extractor/.test(await tile.textContent()), 'home page has the Font Extractor tile');
    await tile.scrollIntoViewIfNeeded();
    await home.screenshot({ path: path.join(shots, 'home.png') });
  } catch (e) { console.error(e); fails++; }
  await browser.close();
  server.close();
  console.log(fails ? fails + ' FAILED' : 'all passed');
  process.exit(fails ? 1 : 0);
});
