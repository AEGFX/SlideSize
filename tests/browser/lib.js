/* Shared helpers for the browser checks. See run.js. */
function need(name) {
  try { return require(name); } catch (e) {}
  const extra = process.env.NODE_MODULES_DIR;
  if (extra) { try { return require(require('path').join(extra, name)); } catch (e) {} }
  throw new Error('Cannot find "' + name + '". Run npm i ' + name + ' or set NODE_MODULES_DIR to a folder that has it.');
}
const { chromium } = need('playwright');
const fs = require('fs'), path = require('path');
const ROOT = process.env.BASE || 'http://localhost:8912';
const BASE = ROOT + '/photo-montage.html';
async function open(opts = {}) {
  const browser = await chromium.launch({ headless: true, args: opts.args || [] });
  const page = await browser.newPage({ viewport: opts.viewport || { width: 1500, height: 950 } });
  const logs = [];
  page.on('console', m => { if ((m.type() === 'error' || m.type() === 'warning') && !/ERR_TUNNEL|fonts\.g/.test(m.text())) logs.push(m.type() + ': ' + m.text()); });
  page.on('pageerror', e => logs.push('PAGEERROR: ' + e.message + '\n' + (e.stack || '').split('\n').slice(0, 5).join('\n')));
  await page.goto(BASE + (opts.query || ''));
  await page.waitForFunction(() => window.SlideshowApp && !SlideshowApp.state().probing, null, { timeout: 20000 });
  return { browser, page, logs };
}
function listPhotos(dir, filter) { return fs.readdirSync(dir).filter(f => !filter || filter(f)).sort().map(f => path.join(dir, f)); }
async function addPhotos(page, files) {
  await page.setInputFiles('#file-input', files);
  await page.waitForFunction(() => !SlideshowApp.state().imp.running, null, { timeout: 300000 });
  await page.waitForTimeout(150);
}
// mutate project state then refresh; fn runs in the page with (P, photos)
async function tweak(page, fn, arg) {
  await page.evaluate(([src, a]) => { const s = SlideshowApp.state(); (new Function('P', 'photos', 'arg', 'return (' + src + ')(P, photos, arg)'))(s.P, s.photos, a); SlideshowApp.refresh(); }, [fn.toString(), arg]);
  await page.waitForFunction(() => !SlideshowApp.state().probing, null, { timeout: 20000 });
  await page.waitForTimeout(60);
}
async function planInfo(page) {
  return page.evaluate(() => { const s = SlideshowApp.state(), p = s.plan; return { ok: p.ok, frames: p.frames, seconds: p.seconds, counts: p.counts, errors: p.errors.map(e => e.msg), notes: p.notes, warnings: p.warnings, order: p.segs.map(x => x.name), segs: p.segs.map(x => ({ name: x.name, start: x.start, slot: x.slot, trans: x.trans, transIn: x.transIn, kind: x.kind, mode: x.mode })), excluded: p.classes.filter(c => !c.included).map(c => ({ id: c.id, reason: c.reason, auto: c.auto })), format: s.P.exp.format }; });
}
async function exportTo(page, outPath, timeout = 600000) {
  await page.waitForFunction(() => !document.getElementById('btn-export').disabled, null, { timeout: 20000 });
  await page.click('#btn-export');
  await page.waitForFunction(() => !SlideshowApp.state().ui.exporting && document.getElementById('result').children.length > 0, null, { timeout });
  const failed = await page.evaluate(() => { const l = document.getElementById('download-link'); return l ? null : document.getElementById('result').innerText; });
  if (failed) throw new Error('export did not produce a download: ' + failed);
  // wait for the checks to finish
  await page.waitForFunction(() => { const d = document.getElementById('checks'); return d && /opens here/i.test(d.textContent); }, null, { timeout: 30000 });
  const b64 = await page.evaluate(async () => {
    const buf = new Uint8Array(await (await fetch(document.getElementById('download-link').href)).arrayBuffer());
    let s = ''; const CH = 0x8000;
    for (let i = 0; i < buf.length; i += CH) s += String.fromCharCode.apply(null, buf.subarray(i, i + CH));
    return btoa(s);
  });
  fs.writeFileSync(outPath, Buffer.from(b64, 'base64'));
  const info = await page.evaluate(() => ({ name: document.getElementById('download-link').download, result: document.getElementById('result').innerText, mode: SlideshowApp.state().ui.mode }));
  return info;
}
module.exports = { open, listPhotos, addPhotos, tweak, planInfo, exportTo, chromium, need, ROOT };
