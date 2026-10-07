/* A stand-in for the local helper, used only by the browser checks.

   A headless browser cannot show a folder picker and the test machine has
   no PowerPoint, so the checks give the page a private browser folder to
   work in and run this in place of the helper. It speaks the same file
   protocol as the real helper scripts. The PDFs come from LibreOffice on
   the test machine, through window.__convert.

   It calls itself Microsoft PowerPoint, because the page refuses anything
   else and the page's own handling of a real result is what is being
   checked. It is loaded by the checks only. No page on the site loads it. */
(function () {
'use strict';
var M = window.__mock = {
  on: true, platform: 'windows', engine: 'Microsoft PowerPoint', version: '16.0', installed: true, verapdf: false,
  caps: { pdfa: true, notes: true, quality: true, bitmapText: true, tags: true, docProps: true, markup: true, hidden: true, range: true, reference: true, placeholders: true },
  fonts: ['Arial', 'Calibri', 'Carlito'], userPresentations: 0,
  delay: 0, stall: {}, timeoutMs: 0, fail: {}, tamper: {}, dropRef: false, pdfaVerdict: null, ignorePdfa: false,
  converted: {}, tickets: [], current: null, stopped: false, seq: 0, abort: null, beats: 0
};
var sleep = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
async function text(dir, name) { try { return await (await (await dir.getFileHandle(name)).getFile()).text(); } catch (e) { return null; } }
async function write(dir, name, data) { var h = await dir.getFileHandle(name, { create: true }), w = await h.createWritable(); await w.write(data); await w.close(); }
function b64(bytes) { var s = ''; for (var i = 0; i < bytes.length; i += 32768) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 32768)); return btoa(s); }
function unb64(s) { var b = atob(s), u = new Uint8Array(b.length); for (var i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; }

async function finish(work, id, res) {
  await write(await work.getDirectoryHandle('done', { create: true }), id + '.json', JSON.stringify(Object.assign({ protocol: 1, id: id, finishedAt: new Date().toISOString(),
    engine: { name: M.engine, version: M.version, build: 'test', platform: M.platform } }, res)));
  M.current = null;
}

async function run(out, work, t) {
  var id = t.id, started = Date.now();
  M.tickets.push(t);
  if (t.task === 'validate') { await sleep(50); return finish(work, id, { ok: true, status: 'done', pdfa: M.pdfaVerdict }); }
  var inDir = await work.getDirectoryHandle('in'), name = t.input.replace(/^in\//, '');
  if (M.stall[id]) {
    for (;;) {
      await sleep(50);
      if (M.abort === id) { M.abort = null; await inDir.removeEntry(name).catch(function () {}); return finish(work, id, { ok: false, status: 'cancelled', error: { code: 'cancelled', message: 'Skipped from the page while it was converting.' } }); }
      if (M.timeoutMs && Date.now() - started > M.timeoutMs) { await inDir.removeEntry(name).catch(function () {}); return finish(work, id, { ok: false, status: 'timeout', timeoutSec: t.timeoutSec, error: { code: 'timeout', message: 'PowerPoint did not finish.' } }); }
    }
  }
  if (M.delay) await sleep(M.delay);
  var bytes = new Uint8Array(await (await (await inDir.getFileHandle(name)).getFile()).arrayBuffer());
  await inDir.removeEntry(name);
  if (M.fail[id]) return finish(work, id, { ok: false, status: 'failed', error: { code: M.fail[id], message: 'The stand-in was told to fail this file.' } });
  var opts = JSON.parse(JSON.stringify(t.options));
  if (M.ignorePdfa) opts.pdfa = false;
  var r = JSON.parse(await window.__convert(b64(bytes), name, JSON.stringify(opts)));
  if (!r.ok) return finish(work, id, { ok: false, status: 'failed', error: r.error });
  var target = t.pdf.indexOf('../') === 0 ? { dir: out, name: t.pdf.slice(3) } : { dir: await work.getDirectoryHandle('out', { create: true }), name: t.pdf.slice(4) };
  var exists = true; try { await target.dir.getFileHandle(target.name); } catch (e) { exists = false; }
  if (exists && !t.overwrite) return finish(work, id, { ok: false, status: 'failed', error: { code: 'exists', message: 'The output file already exists: ' + target.name } });
  await write(target.dir, target.name, unb64(r.pdf));
  var ref = null;
  if (t.options.reference && r.refs && !M.dropRef) {
    var refRoot = await work.getDirectoryHandle('ref', { create: true });
    await refRoot.removeEntry(id, { recursive: true }).catch(function () {});
    var refDir = await refRoot.getDirectoryHandle(id, { create: true }), slides = Object.keys(r.refs).map(Number).sort(function (a, b) { return a - b; });
    if (t.options.reference.mode === 'sample' && slides.length > 5) { var pick = []; for (var k = 0; k < 5; k++) pick.push(slides[Math.floor(k * (slides.length - 1) / 4 + 0.5)]); slides = pick.filter(function (v, i, a) { return a.indexOf(v) === i; }); }
    for (var i = 0; i < slides.length; i++) {
      var png = unb64(r.refs[slides[i]]);
      if (M.tamper[id] === slides[i]) png = await tamper(png);
      await write(refDir, 'slide-' + ('0000' + slides[i]).slice(-4) + '.png', png);
    }
    ref = { dir: 'ref/' + id, slides: slides, pattern: 'slide-%04d.png' };
  }
  M.converted[id] = (M.converted[id] || 0) + 1;
  await finish(work, id, { ok: true, status: 'done', startedAt: new Date(started).toISOString(),
    facts: r.facts, export: { method: 'ExportAsFixedFormat2', pageMap: r.pageMap, notApplied: [], removedSlides: [],
      placeholders: (t.options.placeholders || []).map(function (p) { return { slide: p.slide, ok: true }; }) },
    pdf: { bytes: unb64(r.pdf).length }, reference: ref, timing: { openMs: 5, exportMs: Date.now() - started, totalMs: Date.now() - started },
    pdfa: t.validateNow ? M.pdfaVerdict : undefined });
}

/* draws a block over part of the picture, as if something on the slide had gone missing */
async function tamper(png) {
  var bmp = await createImageBitmap(new Blob([png], { type: 'image/png' })), c = document.createElement('canvas');
  c.width = bmp.width; c.height = bmp.height;
  var x = c.getContext('2d'); x.drawImage(bmp, 0, 0); x.fillStyle = '#103070'; x.fillRect(c.width * 0.55, c.height * 0.3, c.width * 0.3, c.height * 0.4);
  var blob = await new Promise(function (r) { c.toBlob(r, 'image/png'); });
  return new Uint8Array(await blob.arrayBuffer());
}

(async function loop() {
  var root = await navigator.storage.getDirectory(), fontsFor = null;
  for (;;) {
    await sleep(120);
    if (!M.on) continue;
    var out, work;
    try { out = await root.getDirectoryHandle('out'); work = await out.getDirectoryHandle('_slidesize'); } catch (e) { continue; }
    try {
      var c = JSON.parse(await text(work, 'control.json') || 'null');
      if (c && c.seq > M.seq) { M.seq = c.seq; if (c.stop) M.stopped = true; if (c.abort) M.abort = c.abort; if (c.retry) M.blocked = null; }
      await write(work, 'helper.json', JSON.stringify({ protocol: 1, helper: '1.0.0', platform: M.platform, os: 'test', pid: 1, heartbeat: new Date().toISOString(),
        state: M.stopped ? 'stopped' : M.blocked ? 'blocked' : M.current ? 'working' : 'idle', message: M.blocked || '',
        powerpoint: { installed: M.installed, version: M.version, build: 'test', runningBefore: false, userPresentations: M.userPresentations },
        caps: M.caps, verapdf: M.verapdf, engine: M.engine, current: M.current, converted: 0, failed: 0 }));
      M.beats++;
      if (fontsFor !== M.fonts) { await write(work, 'fonts.json', JSON.stringify({ platform: M.platform, families: M.fonts })); fontsFor = M.fonts; }
      if (M.stopped || M.blocked || M.current || !M.installed) continue;
      var q = await work.getDirectoryHandle('queue'), names = [];
      for await (var e of q.entries()) if (/\.json$/.test(e[0])) names.push(e[0]);
      names.sort();
      if (!names.length) continue;
      var t = JSON.parse(await text(q, names[0]));
      await q.removeEntry(names[0]);
      M.current = { id: t.id, phase: 'exporting', since: new Date().toISOString() };
      run(out, work, t).catch(function (err) { finish(work, t.id, { ok: false, status: 'failed', error: { code: 'convert-failed', message: String(err && err.message || err) } }); });
    } catch (err) { /* a file was mid write. Try again on the next beat */ }
  }
})();
})();
