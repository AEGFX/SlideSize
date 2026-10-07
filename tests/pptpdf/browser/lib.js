/* Shared parts of the PowerPoint to PDF browser checks. See run.js. */
'use strict';
const fs = require('fs'), os = require('os'), path = require('path'), http = require('http'), crypto = require('crypto');
const { spawnSync } = require('child_process');

const root = path.resolve(__dirname, '..', '..', '..');
const decks = path.join(root, 'tests', 'pptpdf', 'decks');
const deps = process.env.PPTPDF_DEPS || path.join(root, 'tests', 'pptpdf', 'node_modules');
const work = process.env.WORK || path.join(os.tmpdir(), 'slidesize-pptpdf-checks');
const Inspect = require(path.join(root, 'pptpdf-inspect.js'));
const Core = require(path.join(root, 'pptpdf-core.js'));
const Pdf = require(path.join(root, 'pptpdf-pdf.js'));
fs.mkdirSync(path.join(work, 'cache'), { recursive: true });
fs.mkdirSync(path.join(work, 'shots'), { recursive: true });

function need(mod) {
  for (const base of [deps, path.join(root, 'node_modules')]) {
    try { return require(path.join(base, mod)); } catch (e) { /* next */ }
  }
  return require(mod);
}
function pdfjsFile(rel) {
  for (const base of [deps, path.join(root, 'node_modules')]) {
    const f = path.join(base, 'pdfjs-dist', rel);
    if (fs.existsSync(f)) return f;
  }
  return null;
}

const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json' };
function serve(port) {
  return new Promise(resolve => {
    const server = http.createServer((req, res) => {
      let p = decodeURIComponent(req.url.split('?')[0]);
      if (p === '/') p = '/index.html';
      const file = path.join(root, p);
      if (!file.startsWith(root) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); res.end('not found'); return; }
      res.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-store' });
      fs.createReadStream(file).pipe(res);
    });
    server.listen(port, '127.0.0.1', () => resolve(server));
  });
}

/* The stand-in engine. LibreOffice makes the PDF and poppler makes the "PowerPoint" pictures of the slides
   from that PDF, so the page compares two different renderers, as it does with real PowerPoint. */
async function convert(b64, name, optionsJson) {
  const o = JSON.parse(optionsJson), buf = Buffer.from(b64, 'base64');
  const key = crypto.createHash('sha1').update(buf).update(JSON.stringify([o.includeHidden, o.range, o.output, o.pdfa, o.tags, !!o.reference, o.reference && o.reference.longEdge])).digest('hex');
  const cached = path.join(work, 'cache', key + '.json');
  if (fs.existsSync(cached)) return fs.readFileSync(cached, 'utf8');
  const dir = fs.mkdtempSync(path.join(work, 'lo-')), ext = path.extname(name) || '.pptx', input = path.join(dir, 'deck' + ext);
  fs.writeFileSync(input, buf);
  let out;
  try {
    const ins = await Inspect.inspect(new Blob([buf]), name);
    if (!ins.ok) throw Object.assign(new Error(ins.error.message), { code: ins.error.code === 'password' ? 'password' : 'open-failed' });
    const n = ins.slideCount, hidden = ins.hiddenSlides || [];
    const from = o.range ? Math.max(1, o.range[0]) : 1, to = o.range ? Math.min(n, o.range[1]) : n, map = [];
    for (let i = from; i <= to; i++) if (o.includeHidden || !hidden.includes(i)) map.push(i);
    const f = { ExportHiddenSlides: { type: 'boolean', value: String(!!o.includeHidden) }, UseTaggedPDF: { type: 'boolean', value: String(!!o.tags) } };
    if (o.range) f.PageRange = { type: 'string', value: from + '-' + to };
    if (o.pdfa) f.SelectPdfVersion = { type: 'long', value: '1' };
    if (o.output === 'notes') { f.ExportNotesPages = { type: 'boolean', value: 'true' }; f.ExportOnlyNotesPages = { type: 'boolean', value: 'true' }; }
    const r = spawnSync('soffice', ['-env:UserInstallation=file://' + path.join(work, 'lo-profile'), '--headless', '--convert-to', 'pdf:impress_pdf_Export:' + JSON.stringify(f), '--outdir', dir, input], { encoding: 'utf8', timeout: 180000 });
    const pdf = path.join(dir, 'deck.pdf');
    if (!fs.existsSync(pdf)) throw Object.assign(new Error('LibreOffice made no PDF. ' + (r.stderr || '')), { code: 'export-failed' });
    const refs = {};
    if (o.reference && o.output !== 'notes') {
      spawnSync('pdftoppm', ['-png', '-scale-to', String(o.reference.longEdge || 1280), pdf, path.join(dir, 'ref')], { timeout: 180000 });
      const pngs = fs.readdirSync(dir).filter(x => /^ref-\d+\.png$/.test(x)).sort((a, b) => parseInt(a.slice(4), 10) - parseInt(b.slice(4), 10));
      pngs.forEach((p, i) => { if (map[i] != null) refs[map[i]] = fs.readFileSync(path.join(dir, p)).toString('base64'); });
    }
    out = JSON.stringify({ ok: true, pdf: fs.readFileSync(pdf).toString('base64'), pageMap: o.output === 'notes' ? map : map, refs,
      facts: { slides: n, hidden, slideWidthPt: ins.widthPt, slideHeightPt: ins.heightPt, fonts: null, readOnly: true } });
  } catch (e) {
    out = JSON.stringify({ ok: false, error: { code: e.code || 'convert-failed', message: e.message } });
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
  fs.writeFileSync(cached, out);
  return out;
}

/* Gives the page a private browser folder in place of the two pickers it cannot show when headless. */
const SHIM = `
(function () {
  var root = null;
  async function dir(name) { root = root || await navigator.storage.getDirectory(); return root.getDirectoryHandle(name, { create: true }); }
  window.__picked = [];
  if (!window.__noFolderAccess) {
    window.showDirectoryPicker = async function (o) { var n = (o && o.id === 'slidesize-pptpdf-src') ? (window.__srcDir || 'src') : 'out'; window.__picked.push(n); return dir(n); };
    window.showSaveFilePicker = async function (o) { var d = await dir('saved'); return d.getFileHandle(o.suggestedName, { create: true }); };
    window.showOpenFilePicker = async function () { var d = await dir(window.__srcDir || 'src'), out = []; for await (var e of d.entries()) if (e[1].kind === 'file') out.push(e[1]); return out; };
  } else { delete window.showDirectoryPicker; window.showDirectoryPicker = undefined; window.showSaveFilePicker = undefined; window.showOpenFilePicker = undefined; }
  window.__opfs = {
    async put(p, b64) { var parts = p.split('/'), d = await dir(parts[0]); for (var i = 1; i < parts.length - 1; i++) d = await d.getDirectoryHandle(parts[i], { create: true });
      var h = await d.getFileHandle(parts[parts.length - 1], { create: true }), w = await h.createWritable(), b = atob(b64), u = new Uint8Array(b.length);
      for (var k = 0; k < b.length; k++) u[k] = b.charCodeAt(k); await w.write(u); await w.close(); },
    async get(p) { try { var parts = p.split('/'), d = await dir(parts[0]); for (var i = 1; i < parts.length - 1; i++) d = await d.getDirectoryHandle(parts[i]);
      var f = await (await d.getFileHandle(parts[parts.length - 1])).getFile(), u = new Uint8Array(await f.arrayBuffer()), s = '';
      for (var k = 0; k < u.length; k += 32768) s += String.fromCharCode.apply(null, u.subarray(k, k + 32768)); return btoa(s); } catch (e) { return null; } },
    async list(p) { try { var parts = p.split('/').filter(Boolean), d = await dir(parts[0]); for (var i = 1; i < parts.length; i++) d = await d.getDirectoryHandle(parts[i]);
      var out = []; for await (var e of d.entries()) out.push(e[0] + (e[1].kind === 'directory' ? '/' : '')); return out.sort(); } catch (e) { return null; } },
    async wipe() { root = root || await navigator.storage.getDirectory(); for await (var e of root.entries()) await root.removeEntry(e[0], { recursive: true }); }
  };
})();`;

async function open(browser, base, o) {
  o = o || {};
  const ua = o.platform === 'macos'
    ? 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36'
    : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36';
  const ctx = o.context || await browser.newContext({ userAgent: ua, viewport: { width: 1400, height: 1000 }, acceptDownloads: true });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(String(e)));
  page.on('console', m => { if (m.type() === 'error' && !/Failed to load resource|net::ERR/.test(m.text())) errors.push(m.text()); });
  await page.addInitScript(p => { Object.defineProperty(navigator, 'platform', { get: () => p === 'macos' ? 'MacIntel' : 'Win32' });
    Object.defineProperty(navigator, 'userAgentData', { get: () => ({ platform: p === 'macos' ? 'macOS' : 'Windows', mobile: false }) }); }, o.platform || 'windows');
  if (o.noFolderAccess) await page.addInitScript(() => { window.__noFolderAccess = true; });
  await page.addInitScript(SHIM);
  await ctx.route('https://fonts.googleapis.com/**', r => r.abort());
  await ctx.route('https://fonts.gstatic.com/**', r => r.abort());
  await ctx.route('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/**', r => {
    const rel = r.request().url().split('/3.11.174/')[1].split('?')[0];
    const f = pdfjsFile(rel.startsWith('cmaps/') || rel.startsWith('standard_fonts/') ? rel : 'build/' + rel);
    if (!f || !fs.existsSync(f)) return r.fulfill({ status: 404, body: '' });
    r.fulfill({ status: 200, body: fs.readFileSync(f), contentType: /\.js$/.test(f) ? 'text/javascript' : 'application/octet-stream' });
  });
  if (!o.context) await ctx.exposeFunction('__convert', convert);
  await page.goto(base + '/ppt-pdf.html' + (o.query || ''));
  if (o.wipe !== false && !o.context) { await page.evaluate(() => window.__opfs.wipe()); await page.evaluate(() => indexedDB.deleteDatabase('slidesize-pptpdf')); await page.evaluate(() => localStorage.clear()); await page.reload(); }
  if (o.mock !== false) await page.addScriptTag({ path: path.join(__dirname, 'mock-helper.js') });
  if (o.platform === 'macos' && o.mock !== false) await page.evaluate(() => { const c = window.__mock.caps; window.__mock.platform = 'macos'; ['pdfa', 'notes', 'quality', 'bitmapText', 'tags', 'docProps', 'markup'].forEach(k => { c[k] = false; }); });
  return { page, ctx, errors };
}

async function putDecks(page, names, into) {
  for (const n of names) {
    const src = Array.isArray(n) ? n[0] : n, dst = Array.isArray(n) ? n[1] : n;
    await page.evaluate(([p, b]) => window.__opfs.put(p, b), [(into || 'src') + '/' + dst, fs.readFileSync(path.join(decks, src)).toString('base64')]);
  }
}
async function getFile(page, p) { const b = await page.evaluate(x => window.__opfs.get(x), p); return b === null ? null : Buffer.from(b, 'base64'); }
async function list(page, p) { return page.evaluate(x => window.__opfs.list(x), p); }
const items = page => page.evaluate(() => window.__pptpdf.batch.items.map(i => ({ id: i.id, name: i.name, status: i.status, outName: i.outName, hasPdf: i.hasPdf, attempts: i.attempts,
  state: (i.status === 'done' || i.status === 'failed' || i.status === 'cancelled') ? window.PptPdfCore.classify(i) : null,
  issues: (i.issues || []).map(x => ({ code: x.code, severity: x.severity, certainty: x.certainty, slide: x.slide, page: x.page, review: x.review, accepted: x.accepted, description: x.description })),
  checks: i.checks, inspect: i.inspect && { ok: i.inspect.ok, slideCount: i.inspect.slideCount, widthPt: i.inspect.widthPt, hidden: i.inspect.hiddenSlides, error: i.inspect.error } })));
async function waitChecked(page) { await page.waitForFunction(() => window.__pptpdf.checking === 0 && window.__pptpdf.batch.items.length > 0 && window.__pptpdf.batch.items.every(i => i.status !== 'new' && i.status !== 'checking'), null, { timeout: 60000 }); }
async function waitRun(page, ms) { await page.waitForFunction(() => window.__pptpdf.running === false && window.__pptpdf.batch.items.some(i => i.status === 'done' || i.status === 'failed' || i.status === 'cancelled'), null, { timeout: ms || 300000 }); }
async function waitHelper(page) { await page.waitForFunction(() => { const b = document.getElementById('light-ppt'); return b && b.classList.contains('ok'); }, null, { timeout: 30000 }); }

module.exports = { root, decks, work, need, serve, open, convert, putDecks, getFile, list, items, waitChecked, waitRun, waitHelper, Core, Pdf, Inspect };
