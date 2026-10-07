/* ============================================================
   PowerPoint to PDF: the page  (pptpdf-app.js)

   Select, check, settings and destination, convert, review, save.

   The page never converts anything itself. It copies one deck at
   a time into the work folder, writes a job ticket, and waits for
   the local helper to hand back a PDF made by PowerPoint. Then it
   opens that PDF and checks it.

   Memory stays flat however long the batch is. One deck is copied
   at a time as a stream, one PDF is open at a time, previews are
   drawn when asked for and kept in a small cache, and the record
   of the batch is written to disk after every file.

   AEGFX / SlideSize
   ============================================================ */
(function () {
'use strict';

var C = window.PptPdfCore, I = window.PptPdfInspect, P = window.PptPdfPdf;
var $ = function (id) { return document.getElementById(id); };
var esc = C.esc;

var PDFJS = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/';
var HELPER_FILES = {
  windows: [{ name: 'Start-SlideSize-Helper.cmd', crlf: true }, { name: 'slidesize-helper.ps1', crlf: true }],
  macos: [{ name: 'Start-SlideSize-Helper.command', exec: true }]
};
var LOOKAHEAD = 2;                 /* decks copied or queued ahead of PowerPoint, never more */
var COMPARE_W = 384;
var BLOCKING = { password: 1, empty: 1, corrupt: 1, unsupported: 1, 'no-slides': 1, 'remote-links': 1, 'source-missing': 1 };
var IN_FLIGHT = { staging: 1, queued: 1, converting: 1, verifying: 1 };
var LS_SETTINGS = 'slidesize.pptpdf.settings', LS_PRESETS = 'slidesize.pptpdf.presets';

var S = {
  platform: C.detectPlatform(navigator),
  support: C.browserSupport(window),
  batch: null,
  src: new Map(),               /* item id -> FileSystemFileHandle or File */
  out: null,                    /* { dir, work, sub: {queue, in, out, done, ref}, name, existing, ours } */
  helperFiles: 'none', helperFilesError: '',
  helper: null, fonts: null, helperReadAt: 0,
  running: false, stopping: false, runSet: null, stagingNow: false, releaseLock: null,
  verifyChain: Promise.resolve(), saveChain: Promise.resolve(), checkChain: Promise.resolve(),
  expanded: {}, selected: {}, resExpanded: {}, filter: 'all', sev: 'all',
  decisions: {}, presets: [], thumb: null, controlSeq: Date.now(), pendingFolderBatch: null,
  dirty: {}, activity: '', verifying: '', checking: 0, viewer: null
};
window.__pptpdf = S;             /* for the browser checks */

/* ------------------------------------------------------------------ *
 * Small tools                                                         *
 * ------------------------------------------------------------------ */

function flag(kind, title, body, sub, buttons) {
  return '<div class="flag ' + kind + '"><div>' + (title ? '<b>' + esc(title) + '</b> ' : '') + (body || '') +
    (sub ? '<span class="sub">' + sub + '</span>' : '') + (buttons ? '<div class="btn-row">' + buttons + '</div>' : '') + '</div></div>';
}
function setSeg(el, value) {
  Array.prototype.forEach.call(el.querySelectorAll('button'), function (b) { b.classList.toggle('active', b.dataset.v === String(value)); });
}
function onSeg(el, fn) {
  el.addEventListener('click', function (e) { var b = e.target.closest('button'); if (b && !b.disabled) fn(b.dataset.v); });
}
function download(blob, name) {
  var url = URL.createObjectURL(blob), a = document.createElement('a');
  a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(function () { URL.revokeObjectURL(url); }, 120000);
}
function loadScript(src) {
  return new Promise(function (res, rej) {
    var t = document.createElement('script');
    t.src = src; t.onload = res; t.onerror = function () { rej(new Error('Could not load ' + src)); };
    document.head.appendChild(t);
  });
}
function pause(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
function settingsOf(item) { return C.settingsFor(item, S.batch.settings); }
function helperFresh() { return !!(S.helper && (Date.now() - Date.parse(S.helper.heartbeat || 0)) < C.HEARTBEAT_STALE_MS && S.helper.state !== 'stopped'); }
function enginePlatform() { return helperFresh() && (S.helper.platform === 'windows' || S.helper.platform === 'macos') ? S.helper.platform : S.platform; }
function caps() { return helperFresh() ? S.helper.caps : null; }
function resolved(item) { return C.resolveSettings(item ? settingsOf(item) : S.batch.settings, enginePlatform(), caps()); }
function stamp(d) {
  d = d || new Date();
  function p(n) { return (n < 10 ? '0' : '') + n; }
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + p(d.getMinutes());
}

/* ------------------------------------------------------------------ *
 * Disk                                                                *
 * ------------------------------------------------------------------ */

async function readText(dir, name) {
  try { return await (await (await dir.getFileHandle(name)).getFile()).text(); } catch (e) { return null; }
}
async function readJson(dir, name) {
  var t = await readText(dir, name);
  if (!t) return null;
  try { return JSON.parse(t); } catch (e) { return null; }     /* half written. Read again on the next tick */
}
async function writeFile(dir, name, data) {
  var h = await dir.getFileHandle(name, { create: true }), w = await h.createWritable();
  try { await w.write(data); await w.close(); } catch (e) { try { await w.abort(); } catch (x) { /* already closed */ } throw e; }
}
async function removeEntry(dir, name, recursive) {
  try { await dir.removeEntry(name, { recursive: !!recursive }); return true; } catch (e) { return false; }
}
async function hasFile(dir, name) {
  try { await dir.getFileHandle(name); return true; } catch (e) { return false; }
}
async function listNames(dir) {
  var out = [];
  for await (var entry of dir.entries()) out.push({ name: entry[0], kind: entry[1].kind, handle: entry[1] });
  return out;
}
async function emptyDir(dir) {
  var n = 0;
  for (var e of await listNames(dir)) { if (await removeEntry(dir, e.name, true)) n++; }
  return n;
}

/* A tiny key and value store for the handles that let a batch be resumed. */
var idb = (function () {
  var open = null;
  function db() {
    if (open) return open;
    open = new Promise(function (res, rej) {
      var r = indexedDB.open('slidesize-pptpdf', 1);
      r.onupgradeneeded = function () { r.result.createObjectStore('kv'); };
      r.onsuccess = function () { res(r.result); };
      r.onerror = function () { rej(r.error); };
    });
    return open;
  }
  function tx(mode, fn) {
    return db().then(function (d) {
      return new Promise(function (res, rej) {
        var t = d.transaction('kv', mode), req = fn(t.objectStore('kv'));
        t.oncomplete = function () { res(req && req.result); };
        t.onerror = function () { rej(t.error); };
      });
    });
  }
  return {
    get: function (k) { return tx('readonly', function (s) { return s.get(k); }).catch(function () { return undefined; }); },
    set: function (k, v) { return tx('readwrite', function (s) { return s.put(v, k); }).catch(function () {}); },
    del: function (k) { return tx('readwrite', function (s) { return s.delete(k); }).catch(function () {}); }
  };
})();

/* ---------- a store-only zip writer, the same idea as the one in PDF to Slides ---------- */
var CRC_TABLE = (function () {
  var t = new Uint32Array(256), c, n, k;
  for (n = 0; n < 256; n++) { c = n; for (k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; }
  return t;
})();
function crcUpdate(c, buf) { for (var i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return c; }
var TE = new TextEncoder();

function ZipWriter(sink) {
  this.sink = sink; this.offset = 0; this.entries = [];
  var d = new Date();
  this.dosTime = ((d.getHours() & 31) << 11) | ((d.getMinutes() & 63) << 5) | ((d.getSeconds() / 2) & 31);
  this.dosDate = (((d.getFullYear() - 1980) & 127) << 9) | (((d.getMonth() + 1) & 15) << 5) | (d.getDate() & 31);
}
ZipWriter.prototype._w = function (u8) { this.offset += u8.length; return this.sink.write(u8); };
ZipWriter.prototype._head = function (name, crc, size, mode) {
  var nb = TE.encode(name), head = new Uint8Array(30 + nb.length), dv = new DataView(head.buffer);
  dv.setUint32(0, 0x04034b50, true); dv.setUint16(4, 20, true); dv.setUint16(6, 0x0800, true); dv.setUint16(8, 0, true);
  dv.setUint16(10, this.dosTime, true); dv.setUint16(12, this.dosDate, true);
  dv.setUint32(14, crc, true); dv.setUint32(18, size, true); dv.setUint32(22, size, true);
  dv.setUint16(26, nb.length, true); dv.setUint16(28, 0, true); head.set(nb, 30);
  this.entries.push({ name: nb, crc: crc, size: size, offset: this.offset, mode: mode || 0 });
  return head;
};
ZipWriter.prototype.add = async function (name, data, mode) {
  var body = typeof data === 'string' ? TE.encode(data) : data;
  await this._w(this._head(name, (crcUpdate(0xFFFFFFFF, body) ^ 0xFFFFFFFF) >>> 0, body.length, mode));
  await this._w(body);
};
/* a file on disk goes through in two streamed passes, one for the checksum and one for the bytes */
ZipWriter.prototype.addBlob = async function (name, blob) {
  if (blob.size > 0xFFFFFFF0) throw new Error(name + ' is too large for a zip made in the browser.');
  var c = 0xFFFFFFFF, rd = blob.stream().getReader(), r;
  for (;;) { r = await rd.read(); if (r.done) break; c = crcUpdate(c, r.value); }
  await this._w(this._head(name, (c ^ 0xFFFFFFFF) >>> 0, blob.size));
  rd = blob.stream().getReader();
  for (;;) { r = await rd.read(); if (r.done) break; await this._w(r.value); }
};
ZipWriter.prototype.finish = async function () {
  var start = this.offset, i, e, rec, dv;
  if (start > 0xFFFFFFF0) throw new Error('The zip is larger than 4 GB, which a zip made in the browser cannot hold.');
  for (i = 0; i < this.entries.length; i++) {
    e = this.entries[i]; rec = new Uint8Array(46 + e.name.length); dv = new DataView(rec.buffer);
    dv.setUint32(0, 0x02014b50, true); dv.setUint16(4, e.mode ? 0x0314 : 20, true); dv.setUint16(6, 20, true); dv.setUint16(8, 0x0800, true);
    dv.setUint16(12, this.dosTime, true); dv.setUint16(14, this.dosDate, true);
    dv.setUint32(16, e.crc, true); dv.setUint32(20, e.size, true); dv.setUint32(24, e.size, true);
    dv.setUint16(28, e.name.length, true);
    if (e.mode) dv.setUint32(38, (e.mode << 16) >>> 0, true);     /* unix permissions, so the Mac script stays runnable */
    dv.setUint32(42, e.offset, true); rec.set(e.name, 46);
    await this._w(rec);
  }
  var eocd = new Uint8Array(22), ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, this.entries.length, true); ev.setUint16(10, this.entries.length, true);
  ev.setUint32(12, this.offset - start, true); ev.setUint32(16, start, true);
  await this._w(eocd);
  await this.sink.close();
};
function blobSink() {
  var parts = [];
  return { write: function (u8) { parts.push(u8); return Promise.resolve(); }, close: function () { return Promise.resolve(); }, result: function () { return new Blob(parts); } };
}

/* ------------------------------------------------------------------ *
 * Environment                                                         *
 * ------------------------------------------------------------------ */

function renderEnv() {
  var h = '';
  if (!S.support.ok) h += flag('bad', 'This browser cannot run the converter.', esc(S.support.message),
    'The page needs to read and write a folder on this computer, which Chrome and Edge allow and Safari and Firefox do not. You can still add files here to check them.');
  if (S.platform === 'other') h += flag('bad', 'PowerPoint conversion runs on Windows and macOS.',
    'This device is neither, so there is no PowerPoint here to do the conversion. Open this page on the computer that has PowerPoint.');
  else h += flag('info', '', 'This computer looks like <b>' + esc(C.platformName(S.platform)) + '</b>. ' +
    (S.platform === 'macos' ? 'PowerPoint for Mac can convert, leave out hidden slides and keep to a slide range. It cannot write PDF/A or notes pages and has no export options. Those settings are switched off below with the reason.'
      : 'PowerPoint for Windows supports every setting below.'),
    'The helper reports what the installed PowerPoint can really do once it is running, and that replaces this guess.');
  $('env-msgs').innerHTML = h;

  var plat = enginePlatform(), rows = [], k;
  rows.push(['Browser', 'Chrome or Edge on a desktop computer. The page reads and writes the output folder directly, which Safari and Firefox do not allow.']);
  rows.push(['Helper', 'Needed on both platforms. A page cannot start PowerPoint. On Windows it is a PowerShell script started by double-click. On macOS it is a shell script started from Terminal. PCs managed to block PowerShell scripts cannot run it.']);
  for (k in C.FEATURES) {
    var a = C.featureAvailable(k, plat, caps());
    if (!a.ok) rows.push([C.FEATURES[k].label, a.reason]);
  }
  rows.push(['PDF/A', 'Shown as Passed only when veraPDF, a free validator, is installed on this computer and has checked the finished file. Without it the result is Failed when a rule is plainly broken and Not verified otherwise.']);
  rows.push(['Fidelity check', 'Compares each PDF page, drawn by the browser, with PowerPoint\'s own picture of the slide. It finds missing or moved content. It cannot prove a page is right, and it is not run on notes pages.']);
  rows.push(['Open presentations', 'Presentations you have open in PowerPoint are left alone. While they are open on Windows a file that stalls PowerPoint cannot be cleared automatically, because the helper will not close a PowerPoint that is showing your work.']);
  if (plat === 'macos') rows.push(['macOS', 'PowerPoint shows each presentation briefly while it converts it. When slides are left out and PowerPoint will not skip them, they are removed from the temporary copy, and slide numbers printed on later slides count from the slides that remain.']);
  rows.push(['Large PDFs', 'A PDF above ' + C.formatBytes(C.PDF_EDIT_MAX_BYTES) + ' is not opened for checking or resizing. It is kept as PowerPoint wrote it and reported as not verified.']);
  rows.push(['.ppt files', 'Only slide size, slide count and password protection are read before conversion. Hidden slides, fonts and media are reported by PowerPoint afterwards.']);
  rows.push(['Linked files', 'A deck that links to pictures or media on the internet or a network share is held back until Allow linked files from the network is turned on.']);
  $('limits').innerHTML = rows.map(function (r) { return '<dt>' + esc(r[0]) + '</dt><dd style="font-family:var(--sans);font-size:12px;color:var(--label)">' + esc(r[1]) + '</dd>'; }).join('');
}

/* ------------------------------------------------------------------ *
 * 1 Select files                                                      *
 * ------------------------------------------------------------------ */

function wanted(name) { return /\.(pptx|ppt)$/i.test(name) && !/^(~\$|\._)/.test(name); }

async function walk(dirHandle, prefix, out, depth) {
  if (depth > 8) return;
  for await (var entry of dirHandle.entries()) {
    var name = entry[0], h = entry[1];
    if (h.kind === 'file') { if (wanted(name)) out.push({ handle: h, folder: prefix }); }
    else if (name !== C.WORK_DIR && name.charAt(0) !== '.') await walk(h, prefix ? prefix + '/' + name : name, out, depth + 1);
  }
}

async function addSources(list, skippedOther) {
  var added = 0, dup = 0, coll = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
  list.sort(function (a, b) {
    return coll.compare((a.folder || '') + '/' + ((a.handle || a.file).name), (b.folder || '') + '/' + ((b.handle || b.file).name));
  });
  for (var i = 0; i < list.length; i++) {
    var src = list[i].handle || list[i].file, file;
    try { file = list[i].file || await list[i].handle.getFile(); } catch (e) { continue; }
    var f = { name: file.name, size: file.size, lastModified: file.lastModified, folder: list[i].folder || '' };
    if (S.batch.items.some(function (it) { return C.sameFile(it, f) && (it.folder || '') === f.folder; })) { dup++; continue; }
    var item = C.addItem(S.batch, f);
    S.src.set(item.id, src);
    added++;
    queueCheck(item);
  }
  var msgs = [];
  if (dup) msgs.push(flag('info', '', C.plural(dup, 'file') + ' already in the list ' + (dup === 1 ? 'was' : 'were') + ' not added again.'));
  if (skippedOther) msgs.push(flag('info', '', C.plural(skippedOther, 'file') + ' that ' + (skippedOther === 1 ? 'is' : 'are') + ' not .pptx or .ppt ' + (skippedOther === 1 ? 'was' : 'were') + ' left out.'));
  if (!added && !dup && !skippedOther) msgs.push(flag('warn', '', 'No .pptx or .ppt files were found there.'));
  $('select-msgs').innerHTML = msgs.join('');
  renderAll();
  storeHandles();
}

async function pickFiles() {
  if (typeof window.showOpenFilePicker === 'function') {
    try {
      var hs = await window.showOpenFilePicker({ multiple: true, id: 'slidesize-pptpdf-src',
        types: [{ description: 'PowerPoint presentations', accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'], 'application/vnd.ms-powerpoint': ['.ppt'] } }] });
      await addSources(hs.filter(function (h) { return wanted(h.name); }).map(function (h) { return { handle: h }; }),
        hs.filter(function (h) { return !wanted(h.name); }).length);
    } catch (e) { if (e && e.name !== 'AbortError') throw e; }
  } else $('file-input').click();
}

async function pickFolder() {
  if (typeof window.showDirectoryPicker !== 'function') { $('select-msgs').innerHTML = flag('bad', '', 'This browser cannot open a folder. Use Add files.'); return; }
  try {
    var d = await window.showDirectoryPicker({ mode: 'read', id: 'slidesize-pptpdf-src' }), found = [];
    await walk(d, d.name, found, 0);
    await addSources(found, 0);
  } catch (e) { if (e && e.name !== 'AbortError') throw e; }
}

async function onDrop(e) {
  var items = e.dataTransfer && e.dataTransfer.items, found = [], other = 0, pending = [];
  if (items && items.length && typeof items[0].getAsFileSystemHandle === 'function') {
    for (var i = 0; i < items.length; i++) if (items[i].kind === 'file') pending.push(items[i].getAsFileSystemHandle());
    var hs = await Promise.all(pending);
    for (var k = 0; k < hs.length; k++) {
      var h = hs[k];
      if (!h) continue;
      if (h.kind === 'directory') await walk(h, h.name, found, 0);
      else if (wanted(h.name)) found.push({ handle: h });
      else other++;
    }
  } else if (e.dataTransfer && e.dataTransfer.files) {
    Array.prototype.forEach.call(e.dataTransfer.files, function (f) { if (wanted(f.name)) found.push({ file: f }); else other++; });
  }
  await addSources(found, other);
}

function clearList() {
  if (S.running) return;
  var settings = S.batch.settings, preset = S.batch.preset;
  S.batch = C.newBatch({ platform: S.platform, settings: settings, preset: preset });
  S.src.clear(); S.expanded = {}; S.selected = {}; S.resExpanded = {}; S.decisions = {}; S.thumb = null;
  if (S.out) { S.out.ours = {}; }
  $('select-msgs').innerHTML = '';
  idb.del('last');
  renderAll();
}

/* ------------------------------------------------------------------ *
 * 2 Check presentations                                               *
 * ------------------------------------------------------------------ */

async function sourceFile(item) {
  var src = S.src.get(item.id);
  if (!src) throw new Error('Access to the original file has to be confirmed again.');
  if (typeof src.getFile !== 'function') return src;
  if (typeof src.queryPermission === 'function') {
    var p = await src.queryPermission({ mode: 'read' });
    if (p !== 'granted') throw new Error('Access to the original file has to be confirmed again.');
  }
  return src.getFile();
}

async function posterIsBlank(ins, part) {
  try {
    var bytes = await ins.readBytes(part, 24 * 1024 * 1024);
    var bmp = await createImageBitmap(new Blob([bytes]));
    var w = 64, h = Math.max(1, Math.round(64 * bmp.height / bmp.width));
    var cv = document.createElement('canvas'); cv.width = w; cv.height = h;
    var cx = cv.getContext('2d', { willReadFrequently: true });
    cx.drawImage(bmp, 0, 0, w, h); bmp.close();
    return C.looksBlank(cx.getImageData(0, 0, w, h).data, w, h).blank;
  } catch (e) { return false; }
}

function recomputeIssues(item) {
  if (!item.inspect) return;
  var s = resolved(item).effective;
  var pre = C.preflightIssues(item, item.inspect, s, enginePlatform());
  var later = (item.issues || []).filter(function (i) { return i.stage !== 'check'; });
  var kept = {};
  (item.issues || []).forEach(function (i) { if (i.accepted) kept[i.id] = i.accepted; });
  item.issues = C.mergeIssues([pre, later]);
  item.issues.forEach(function (i) { if (kept[i.id]) i.accepted = kept[i.id]; });
  if (item.status === 'ready' || item.status === 'blocked' || item.status === 'checking' || item.status === 'new') {
    item.status = item.issues.some(function (i) { return BLOCKING[i.code] && i.severity === 'error'; }) ? 'blocked' : 'ready';
  }
}

function queueCheck(item) {
  S.checking++;
  S.checkChain = S.checkChain.then(async function () {
    item.status = 'checking'; mark('files');
    try {
      var file = await sourceFile(item), ins = await I.inspect(file, item.name);
      if (ins.ok && ins.slides) {
        for (var i = 0; i < ins.slides.length; i++) {
          var sl = ins.slides[i];
          for (var k = 0; sl && sl.media && k < sl.media.length; k++) {
            var m = sl.media[k];
            if (m.type === 'video' && m.posterPart) m.posterBlank = await posterIsBlank(ins, m.posterPart);
          }
        }
        if (!S.thumb && ins.thumbnailPart) {
          try {
            var bmp = await createImageBitmap(new Blob([await ins.readBytes(ins.thumbnailPart, 8 * 1024 * 1024)]));
            S.thumb = { id: item.id, bmp: bmp };
          } catch (e) { /* no preview picture, the outline is drawn */ }
        }
      }
      item.inspect = C.slimInspect(ins);
    } catch (e) {
      item.inspect = { ok: false, kind: item.ext, error: { code: 'source-missing', message: e.message || String(e) } };
    }
    item.status = 'new';
    recomputeIssues(item);
  }).catch(function () {}).then(function () { S.checking--; mark('files', 'plan', 'start', 'size'); });
}

function issueCounts(item, stage) {
  var c = { error: 0, warning: 0, info: 0 };
  (item.issues || []).forEach(function (i) { if (!stage || i.stage === stage) c[i.severity]++; });
  return c;
}

function chipsFor(item) {
  var ins = item.inspect, out = [];
  if (!ins) return '';
  if (!ins.ok) return '<span class="chip bad">' + esc(({ password: 'Password protected', 'not-zip': 'Damaged', 'zip-damaged': 'Damaged', 'cfb-damaged': 'Damaged', empty: 'Empty', 'source-missing': 'Access needed' })[ins.error.code] || 'Not a presentation') + '</span>';
  if (ins.kind === 'ppt') out.push('<span class="chip dim">.ppt, limited check</span>');
  if (ins.hiddenSlides && ins.hiddenSlides.length) out.push('<span class="chip dim">' + ins.hiddenSlides.length + ' hidden</span>');
  if (ins.counts) {
    if (ins.counts.video) out.push('<span class="chip warn">' + ins.counts.video + ' video</span>');
    if (ins.counts.audio) out.push('<span class="chip dim">' + ins.counts.audio + ' audio</span>');
    if (ins.counts.animated) out.push('<span class="chip dim">' + ins.counts.animated + ' animated</span>');
    if (ins.counts.notes) out.push('<span class="chip dim">' + ins.counts.notes + ' with notes</span>');
  }
  if (ins.externalLinks && ins.externalLinks.length) out.push('<span class="chip ' + (ins.externalLinks.some(function (l) { return l.remote; }) ? 'bad' : 'warn') + '">' + ins.externalLinks.length + ' linked</span>');
  var c = issueCounts(item, 'check');
  if (c.error) out.push('<span class="chip bad">' + C.plural(c.error, 'error') + '</span>');
  if (c.warning) out.push('<span class="chip warn">' + C.plural(c.warning, 'warning') + '</span>');
  if (!out.length) out.push('<span class="chip good">Nothing to flag</span>');
  return out.join('');
}

function statusChip(item) {
  if (item.status === 'done' || item.status === 'failed' || item.status === 'cancelled') {
    var st = C.classify(item); return '<span class="st ' + st + '">' + esc(C.STATES[st]) + '</span>';
  }
  var live = IN_FLIGHT[item.status] || item.status === 'checking';
  var label = item.status === 'blocked' ? (item.inspect && item.inspect.ok ? 'Held back' : 'Cannot convert') : item.status === 'ready' ? 'Ready' : C.statusLabel(item.status);
  return '<span class="st ' + (live ? 'live' : item.status === 'blocked' ? 'failed' : 'wait') + '">' + (item.status === 'checking' ? '<span class="spinner"></span>' : '') + esc(label) + '</span>';
}

function issueHtml(item, i, withAccept) {
  var where = i.slide != null ? 'Slide ' + i.slide + (i.page != null ? '<br>PDF page ' + i.page : '') : i.page != null ? 'PDF page ' + i.page : 'Whole file';
  var btn = '';
  if (withAccept && (i.review || i.severity === 'error') && !i.accepted && i.code !== 'engine-unknown') btn = '<button type="button" class="btn small" data-accept="' + esc(item.id + '|' + i.id) + '">Mark as reviewed</button>';
  if (withAccept && i.code === 'visual-diff') btn = '<button type="button" class="btn small" data-view="' + esc(item.id + '|' + i.page) + '">Side by side</button>' + btn;
  return '<div class="issue ' + i.severity + (i.accepted ? ' accepted' : '') + '"><div class="where"><b>' + C.SEVERITY[i.severity] + '</b>' + where + '<br>' + C.CERTAINTY[i.certainty] + '</div>' +
    '<div><div class="what">' + esc(i.description) + (i.accepted ? ' <span class="chip good">Reviewed</span>' : i.review ? ' <span class="chip warn">Needs review</span>' : '') + '</div>' +
    (i.impact ? '<div class="more"><b>Likely impact</b> ' + esc(i.impact) + '</div>' : '') + (i.fix ? '<div class="more"><b>Suggested fix</b> ' + esc(i.fix) + '</div>' : '') + '</div>' +
    '<div class="btn-row">' + btn + '</div></div>';
}

function overrideHtml(item) {
  var o = item.overrides || {}, b = S.batch.settings, r = resolved(item), dis = r.disabled;
  function sel(key, cur, opts) {
    return '<select data-ov="' + key + '" data-id="' + item.id + '">' + opts.map(function (x) { return '<option value="' + x[0] + '"' + (String(cur) === String(x[0]) ? ' selected' : '') + (x[2] ? ' disabled' : '') + '>' + esc(x[1]) + '</option>'; }).join('') + '</select>';
  }
  var size = o.size ? o.size.preset : '', fit = o.size ? o.size.fit : 'fit';
  var locked = S.running || IN_FLIGHT[item.status];
  return '<div class="sub-head">Settings for this file only</div>' +
    '<div class="grid4"' + (locked ? ' style="opacity:.5;pointer-events:none"' : '') + '>' +
    '<div class="field"><span class="fld">Page size</span>' + sel('size', size, [['', 'Same as the batch']].concat(C.SIZE_PRESETS.filter(function (p) { return p.id !== 'custom'; }).map(function (p) { return [p.id, p.name]; }))) + '</div>' +
    '<div class="field"><span class="fld">When the shape differs</span>' + sel('fit', fit, [['fit', 'Fit, add margins'], ['fill', 'Fill, crop']]) + '</div>' +
    '<div class="field"><span class="fld">Hidden slides</span>' + sel('hidden', o.hidden === undefined ? '' : o.hidden ? '1' : '0', [['', 'Same as the batch'], ['1', 'Include'], ['0', 'Leave out']]) + '</div>' +
    '<div class="field"><span class="fld">Fidelity check</span>' + sel('fidelity', o.fidelity || '', [['', 'Same as the batch'], ['all', 'Every slide'], ['sample', 'Sample'], ['off', 'Off']]) + '</div>' +
    '<div class="field"><span class="fld">Slides, first to last</span><input type="text" data-ov="range" data-id="' + item.id + '" placeholder="All, or 2-9" value="' + (o.range ? o.range.from + '-' + o.range.to : '') + '"></div>' +
    '<div class="field"><span class="fld">Pages</span>' + sel('output', o.output || '', [['', 'Same as the batch'], ['slides', 'Slides only'], ['notes', 'Notes pages', !!dis.notes]]) + '</div>' +
    '<div class="field"><span class="fld">Format</span>' + sel('pdfa', o.pdfa === undefined ? '' : o.pdfa ? '1' : '0', [['', 'Same as the batch'], ['0', 'Standard PDF'], ['1', 'PDF/A', !!dis.pdfa]]) + '</div>' +
    '<div class="field"><span class="fld">Time limit (minutes)</span><input type="number" min="1" max="120" data-ov="timeout" data-id="' + item.id + '" placeholder="' + Math.round(b.timeoutSec / 60) + '" value="' + (o.timeoutSec ? Math.round(o.timeoutSec / 60) : '') + '"></div>' +
    '</div>' + (item.overrides ? '<div class="btn-row"><button type="button" class="btn small" data-ovreset="' + item.id + '">Use the batch settings</button><span class="hint">This file has settings of its own.</span></div>' : '') +
    (r.errors.length ? '<div class="msgs">' + r.errors.map(function (e) { return flag('bad', '', esc(e)); }).join('') + '</div>' : '');
}

function renderFiles() {
  var items = S.batch.items, body = $('files-body');
  $('p-check').classList.toggle('idle', !items.length);
  $('btn-clear').disabled = !items.length || S.running;
  if (!items.length) {
    $('check-count').textContent = 'No files yet'; $('check-progress').textContent = '';
    body.innerHTML = '<tr><td colspan="6" class="empty">Add presentations and each one is read here. Slide count, slide size, hidden slides, fonts, video, links and anything a PDF cannot show.</td></tr>';
    return;
  }
  var total = items.reduce(function (a, it) { return a + (it.size || 0); }, 0);
  $('check-count').textContent = C.plural(items.length, 'file') + ', ' + C.formatBytes(total);
  $('check-progress').innerHTML = S.checking ? '<span class="spinner"></span>Checking, ' + S.checking + ' to go' : '';
  body.innerHTML = items.map(function (it) {
    var ins = it.inspect, ok = ins && ins.ok;
    var row = '<tr class="main" data-row="' + it.id + '"><td class="name">' + esc(it.name) + (it.folder ? '<small>' + esc(it.folder) + '</small>' : '') + (it.overrides ? '<small>Own settings</small>' : '') + '</td>' +
      '<td class="num">' + esc(C.formatBytes(it.size)) + '</td>' +
      '<td class="num">' + (ok && ins.slideCount != null ? ins.slideCount : '') + '</td>' +
      '<td class="num">' + (ok && ins.widthPt ? esc(C.sizeLabel(ins.widthPt, ins.heightPt)) + '<br><span style="color:var(--muted)">' + esc(C.aspectLabel(ins.widthPt, ins.heightPt)) + '</span>' : '') + '</td>' +
      '<td><div class="chips">' + chipsFor(it) + '</div></td><td>' + statusChip(it) + '</td></tr>';
    if (S.expanded[it.id]) {
      var pre = (it.issues || []).filter(function (i) { return i.stage === 'check'; });
      row += '<tr class="detail"><td colspan="6">' +
        (ok && ins.fonts && ins.fonts.length ? '<p class="hint" style="margin-bottom:8px"><b>Fonts named in the deck</b> ' + esc(ins.fonts.join(', ')) +
          (ins.embeddedFonts && ins.embeddedFonts.length ? '. <b>Embedded in the deck</b> ' + esc(ins.embeddedFonts.join(', ')) : '') +
          '. Whether each one is on this computer is checked once the helper is running.</p>' : '') +
        (pre.length ? '<div class="issues">' + pre.map(function (i) { return issueHtml(it, i, false); }).join('') + '</div>' : '<p class="hint">Nothing to flag before conversion.</p>') +
        (ok ? overrideHtml(it) : '') +
        '<div class="btn-row" style="margin-top:10px"><button type="button" class="btn small" data-remove="' + it.id + '"' + (S.running ? ' disabled' : '') + '>Remove from the list</button></div></td></tr>';
    }
    return row;
  }).join('');
}

function applyOverride(id, key, value) {
  var item = S.batch.items.filter(function (i) { return i.id === id; })[0];
  if (!item || S.running) return;
  var o = item.overrides ? JSON.parse(JSON.stringify(item.overrides)) : {};
  if (key === 'size') { if (value === '') delete o.size; else o.size = Object.assign({ orientation: 'auto', fit: 'fit', margin: 'none' }, o.size || {}, { preset: value }); }
  else if (key === 'fit') { if (o.size) o.size.fit = value; }
  else if (key === 'hidden' || key === 'pdfa') { if (value === '') delete o[key]; else o[key] = value === '1'; }
  else if (key === 'fidelity' || key === 'output') { if (value === '') delete o[key]; else o[key] = value; }
  else if (key === 'range') {
    var m = /^\s*(\d+)\s*(?:-|to|–)\s*(\d+)\s*$/i.exec(value), one = /^\s*(\d+)\s*$/.exec(value);
    if (m) o.range = { from: +m[1], to: +m[2] }; else if (one) o.range = { from: +one[1], to: +one[1] }; else delete o.range;
  } else if (key === 'timeout') { var t = Math.round(+value); if (t >= 1) o.timeoutSec = t * 60; else delete o.timeoutSec; }
  item.overrides = Object.keys(o).length ? o : null;
  recomputeIssues(item);
  mark('files', 'plan', 'start');
}

/* ------------------------------------------------------------------ *
 * 3 Settings                                                          *
 * ------------------------------------------------------------------ */

function loadPresets() {
  var mine = [];
  try { mine = JSON.parse(localStorage.getItem(LS_PRESETS) || '[]'); } catch (e) { mine = []; }
  S.presets = C.BUILTIN_PRESETS.concat(Array.isArray(mine) ? mine.filter(function (p) { return p && p.id && p.name && p.settings; }) : []);
}
function savePresets() {
  try { localStorage.setItem(LS_PRESETS, JSON.stringify(S.presets.filter(function (p) { return !p.builtin; }))); } catch (e) { /* private window */ }
}
function renderPresetList() {
  var sel = $('preset');
  sel.innerHTML = S.presets.map(function (p) { return '<option value="' + esc(p.id) + '">' + esc(p.name) + '</option>'; }).join('') + '<option value="">Changed, not saved</option>';
  sel.value = S.batch.preset || '';
  var cur = S.presets.filter(function (p) { return p.id === S.batch.preset; })[0];
  $('btn-preset-delete').disabled = !cur || !!cur.builtin;
}

function settingsChanged(fromPreset) {
  if (!fromPreset) S.batch.preset = '';
  try { localStorage.setItem(LS_SETTINGS, JSON.stringify({ settings: S.batch.settings, preset: S.batch.preset })); } catch (e) { /* private window */ }
  S.batch.items.forEach(recomputeIssues);
  renderSettings(); mark('files', 'plan', 'start', 'size'); flush();
}

function setDisabled(key, labelId, inputs, reason) {
  var why = $('why-' + key);
  if (why) { why.hidden = !reason; why.textContent = reason || ''; }
  if (labelId && $(labelId)) $(labelId).classList.toggle('off', !!reason);
  inputs.forEach(function (el) { if (el) el.disabled = !!reason || S.running; });
}

function renderSettings() {
  var s = S.batch.settings, r = resolved(null), e = r.effective, d = r.disabled, lock = S.running;
  $('p-settings').classList.toggle('idle', !S.batch.items.length);
  renderPresetList();
  if (!$('seg-size').children.length) $('seg-size').innerHTML = C.SIZE_PRESETS.map(function (p) { return '<button type="button" data-v="' + p.id + '" title="' + esc(p.note || '') + '">' + esc(p.name) + '</button>'; }).join('');
  ['seg-size', 'seg-fit', 'seg-margin', 'seg-range', 'seg-output', 'seg-pdfa', 'seg-quality', 'seg-fidelity'].forEach(function (id) {
    Array.prototype.forEach.call($(id).querySelectorAll('button'), function (b) { b.disabled = lock; });
  });

  setSeg($('seg-size'), s.size.preset);
  $('size-custom').hidden = s.size.preset !== 'custom';
  $('size-options').hidden = s.size.preset === 'original';
  if (document.activeElement !== $('custom-w')) $('custom-w').value = C.trimNum(C.ptToCm(s.size.customW), 2);
  if (document.activeElement !== $('custom-h')) $('custom-h').value = C.trimNum(C.ptToCm(s.size.customH), 2);
  $('orientation').value = s.size.orientation;
  $('orientation').disabled = s.size.preset === 'custom' || lock;
  setSeg($('seg-fit'), s.size.fit); setSeg($('seg-margin'), s.size.margin);
  $('size-hint').textContent = s.size.preset === 'original'
    ? 'Each PDF keeps the slide size of its deck. Choose a page size to put every deck on the same page.'
    : 'The PDF is resized after PowerPoint has made it, by scaling and centring each page. The presentation is not changed and nothing is ever stretched.';

  $('opt-hidden').checked = !!e.hidden; setDisabled('hidden', 'lab-hidden', [$('opt-hidden')], d.hidden);
  setSeg($('seg-range'), s.range ? 'range' : 'all'); $('range-fields').hidden = !s.range;
  if (s.range) { if (document.activeElement !== $('range-from')) $('range-from').value = s.range.from; if (document.activeElement !== $('range-to')) $('range-to').value = s.range.to; }
  setDisabled('range', null, Array.prototype.slice.call($('seg-range').querySelectorAll('button')), d.range);
  setSeg($('seg-output'), e.output);
  setDisabled('notes', null, [$('seg-output').querySelector('[data-v="notes"]')], d.notes);
  setSeg($('seg-pdfa'), e.pdfa ? '1' : '0');
  setDisabled('pdfa', null, [$('seg-pdfa').querySelector('[data-v="1"]')], d.pdfa);
  setSeg($('seg-quality'), e.quality);
  setDisabled('quality', null, [$('seg-quality').querySelector('[data-v="minimum"]')], d.quality);
  setSeg($('seg-fidelity'), e.fidelity);
  setDisabled('reference', null, Array.prototype.slice.call($('seg-fidelity').querySelectorAll('button')), d.reference);

  [['bitmapText', 'bitmapText'], ['tags', 'tags'], ['docProps', 'docProps'], ['markup', 'markup'], ['placeholders', 'placeholders']].forEach(function (p) {
    $('opt-' + p[0]).checked = !!e[p[1]];
    setDisabled(p[0], 'lab-' + p[0], [$('opt-' + p[0])], d[p[1]]);
  });
  $('opt-links').checked = s.links !== 'remove'; $('opt-links').disabled = lock;
  $('opt-metadata').checked = e.metadata === 'strip'; setDisabled('metadata', 'lab-metadata', [$('opt-metadata')], d.metadata);
  $('opt-remote').checked = !!s.allowRemoteLinks; $('opt-remote').disabled = lock;
  $('opt-keeprefs').checked = !!s.keepReferences; $('opt-keeprefs').disabled = lock;
  $('sensitivity').value = s.sensitivity; $('sensitivity').disabled = lock;
  if (document.activeElement !== $('timeout')) $('timeout').value = Math.round(s.timeoutSec / 60);
  $('timeout').disabled = lock;
  ['custom-w', 'custom-h', 'range-from', 'range-to', 'preset', 'btn-preset-save'].forEach(function (id) { $(id).disabled = lock; });

  var notes = r.errors.map(function (x) { return flag('bad', '', esc(x)); }).concat(r.notes.map(function (n) { return flag(n.level === 'warn' ? 'warn' : 'info', '', esc(n.text)); }));
  $('setting-notes').innerHTML = notes.join('');
}

/* Draws the first deck's slide on the chosen page, so a crop or a margin is seen before anything is converted. */
function renderSize() {
  var cv = $('size-canvas'), cx = cv.getContext('2d'), W = cv.width, H = cv.height;
  var item = S.batch.items.filter(function (it) { return it.inspect && it.inspect.ok && it.inspect.widthPt; })[0];
  cx.fillStyle = '#111'; cx.fillRect(0, 0, W, H);
  if (!item) { $('size-readout').textContent = 'Add a presentation to see how its slides sit on the page.'; return; }
  var s = resolved(null).effective, sw = item.inspect.widthPt, sh = item.inspect.heightPt;
  var t = C.targetSize(sw, sh, s.size), tf = t.changed ? C.fitTransform(sw, sh, t.w, t.h, s.size.fit) : { scale: 1, tx: 0, ty: 0, cropX: 0, cropY: 0, marginX: 0, marginY: 0, cropped: false };
  /* everything the slide and the page cover, fitted into the canvas */
  var minX = Math.min(0, tf.tx), minY = Math.min(0, tf.ty), maxX = Math.max(t.w, tf.tx + tf.scale * sw), maxY = Math.max(t.h, tf.ty + tf.scale * sh);
  var k = Math.min((W - 40) / (maxX - minX), (H - 40) / (maxY - minY));
  var ox = (W - k * (maxX - minX)) / 2 - k * minX, oy = (H - k * (maxY - minY)) / 2 - k * minY;
  function X(x) { return ox + k * x; }
  function Y(y) { return oy + k * (t.h - y); }             /* PDF y runs up, the canvas runs down */
  var px = X(0), py = Y(t.h), pw = k * t.w, ph = k * t.h;
  var sx = X(tf.tx), sy = Y(tf.ty + tf.scale * sh), ssw = k * tf.scale * sw, ssh = k * tf.scale * sh;
  cx.fillStyle = s.size.margin === 'black' ? '#000' : '#fff'; cx.fillRect(px, py, pw, ph);
  function slide(alpha) {
    cx.globalAlpha = alpha;
    if (S.thumb && S.thumb.id === item.id) cx.drawImage(S.thumb.bmp, sx, sy, ssw, ssh);
    else {
      cx.fillStyle = '#3a3a3a'; cx.fillRect(sx, sy, ssw, ssh);
      cx.fillStyle = '#bbb'; cx.font = '13px DM Sans, sans-serif'; cx.textAlign = 'center';
      cx.fillText('Slide ' + C.aspectLabel(sw, sh), sx + ssw / 2, sy + ssh / 2 + 4);
    }
    cx.globalAlpha = 1;
  }
  slide(0.28);                                               /* the part that falls off the page, shown faint */
  cx.save(); cx.beginPath(); cx.rect(px, py, pw, ph); cx.clip(); slide(1); cx.restore();
  if (tf.cropped) { cx.strokeStyle = '#E7302A'; cx.setLineDash([5, 4]); cx.lineWidth = 1.5; cx.strokeRect(sx, sy, ssw, ssh); cx.setLineDash([]); }
  cx.strokeStyle = '#E86FA6'; cx.lineWidth = 2; cx.strokeRect(px, py, pw, ph);

  var text = 'Page ' + C.sizeLabel(t.w, t.h) + ', ' + t.orientation + '. ';
  if (!t.changed) text += 'Same as the slide. ';
  else if (tf.cropped) text += 'Slide scaled to ' + C.trimNum(tf.scale * 100, 1) + '% and cropped by ' +
    (tf.cropX ? C.trimNum(tf.cropX * 100, 1) + '% at the left and at the right' : C.trimNum(tf.cropY * 100, 1) + '% at the top and at the bottom') + '. The dashed line is what is lost. ';
  else if (tf.marginX || tf.marginY) text += 'Slide scaled to ' + C.trimNum(tf.scale * 100, 1) + '% with margins of ' +
    (tf.marginX ? C.trimNum(C.ptToCm(tf.marginX), 2) + ' cm left and right' : C.trimNum(C.ptToCm(tf.marginY), 2) + ' cm top and bottom') +
    (s.size.margin === 'none' ? ', left unpainted, which shows as white in most readers. ' : ', painted ' + s.size.margin + '. ');
  else text += 'Slide scaled to ' + C.trimNum(tf.scale * 100, 1) + '%, same shape. ';
  text += 'Shown for ' + item.name + (S.thumb && S.thumb.id === item.id ? ', using the preview picture stored in the file.' : '.');
  $('size-readout').textContent = text;
}

/* ------------------------------------------------------------------ *
 * Output folder and helper                                            *
 * ------------------------------------------------------------------ */

async function openOutput(dir) {
  var work = await dir.getDirectoryHandle(C.WORK_DIR, { create: true }), sub = {};
  for (var n of ['queue', 'in', 'out', 'done', 'ref']) sub[n] = await work.getDirectoryHandle(n, { create: true });
  await writeFile(work, 'marker.json', JSON.stringify({ protocol: C.PROTOCOL, createdBy: 'slidesize.com PowerPoint to PDF', batch: S.batch.id }));
  var existing = [];
  for (var e of await listNames(dir)) if (e.kind === 'file' && /\.pdf$/i.test(e.name)) existing.push(e.name);
  S.out = { dir: dir, work: work, sub: sub, name: dir.name, existing: existing, ours: (S.out && S.out.dirName === dir.name && S.out.ours) || {}, dirName: dir.name };
  S.helper = null; S.fonts = null;
}

async function writeHelperFiles() {
  var list = HELPER_FILES[S.platform];
  if (!S.out || !list) { S.helperFiles = 'none'; return; }
  try {
    for (var i = 0; i < list.length; i++) {
      var r = await fetch('helper/' + list[i].name, { cache: 'no-store' });
      if (!r.ok) throw new Error('The helper file ' + list[i].name + ' could not be fetched from this site.');
      var text = await r.text();
      if (list[i].crlf) text = text.replace(/\r?\n/g, '\r\n');
      await writeFile(S.out.work, list[i].name, text);
    }
    S.helperFiles = 'written'; S.helperFilesError = '';
  } catch (e) {
    /* some browsers and security tools refuse to let a page write a script. The download is the fallback. */
    S.helperFiles = 'failed'; S.helperFilesError = e && e.message ? e.message : String(e);
  }
}

async function downloadHelper(ev) {
  ev.preventDefault();
  var sink = blobSink(), zip = new ZipWriter(sink), any = false;
  for (var plat of ['windows', 'macos']) {
    for (var f of HELPER_FILES[plat]) {
      var r = await fetch('helper/' + f.name, { cache: 'no-store' });
      if (!r.ok) continue;
      var text = await r.text();
      if (f.crlf) text = text.replace(/\r?\n/g, '\r\n');
      await zip.add('SlideSize helper/' + f.name, text, f.exec ? 0o100755 : 0o100644); any = true;
    }
  }
  if (!any) { $('helper-msgs').innerHTML = flag('bad', '', 'The helper files could not be fetched from this site.'); return; }
  await zip.add('SlideSize helper/READ ME.txt', 'SlideSize PowerPoint to PDF helper\r\n\r\nCopy the file for your system into the _slidesize folder inside the output folder you chose on the page, then start it as the page describes.\r\n\r\nWindows  Start-SlideSize-Helper.cmd and slidesize-helper.ps1, both together\r\nmacOS    Start-SlideSize-Helper.command\r\n\r\nNothing is installed. Close the helper window to stop it.\r\n', 0o100644);
  await zip.finish();
  download(sink.result(), 'SlideSize helper.zip');
}

async function chooseFolder() {
  if (!S.support.ok) return;
  var dir;
  try { dir = await window.showDirectoryPicker({ mode: 'readwrite', id: 'slidesize-pptpdf-out' }); }
  catch (e) { if (e && e.name !== 'AbortError') $('helper-msgs').innerHTML = flag('bad', '', esc(e.message || String(e))); return; }
  await useFolder(dir, false);
}

async function useFolder(dir, forResume) {
  var saved = null;
  try {
    var w = await dir.getDirectoryHandle(C.WORK_DIR);
    saved = await readJson(w, 'batch.json');
  } catch (e) { saved = null; }
  await openOutput(dir);
  await writeHelperFiles();
  $('setup').hidden = false;
  S.pendingFolderBatch = null;
  if (saved && saved.version === 1 && saved.id !== S.batch.id) {
    var sum = C.resumeSummary(C.restoreBatch(saved));
    if (forResume || (sum.left > 0 && !S.batch.items.length)) { await resumeFrom(saved); return; }
    if (sum.left > 0) S.pendingFolderBatch = saved;
  }
  storeHandles();
  await pollHelper();
  renderAll();
}

function storeHandles() {
  if (!S.out) return;
  var sources = [];
  S.src.forEach(function (h, id) { if (h && typeof h.getFile === 'function') sources.push({ id: id, handle: h }); });
  idb.set('last', { dir: S.out.dir, batchId: S.batch.id, name: S.out.name, sources: sources, at: Date.now() });
}

async function pollHelper() {
  if (!S.out) return;
  var h = await readJson(S.out.work, 'helper.json');
  if (h) { S.helper = h; S.helperReadAt = Date.now(); }
  if (helperFresh() && !S.fonts) {
    var f = await readJson(S.out.work, 'fonts.json');
    if (f && Array.isArray(f.families)) S.fonts = f.families;
  }
  if (S.helper && S.batch) S.batch.helper = { helper: S.helper.helper, platform: S.helper.platform, os: S.helper.os, powerpoint: S.helper.powerpoint, verapdf: !!S.helper.verapdf };
}

async function sendControl(o) {
  if (!S.out) return;
  S.controlSeq = Math.max(S.controlSeq + 1, Date.now());
  await writeFile(S.out.work, 'control.json', JSON.stringify(Object.assign({ seq: S.controlSeq, stop: false, abort: null, retry: false }, o)));
}

function setupSteps() {
  var w = '<code>' + C.WORK_DIR + '</code>';
  if (S.platform === 'windows') return '<ol>' +
    '<li>In File Explorer open the output folder <b>' + esc(S.out.name) + '</b>, then the ' + w + ' folder inside it.</li>' +
    '<li>Double-click <code>Start-SlideSize-Helper.cmd</code>. If Windows warns that the file came from the internet, choose <b>Run</b>, or <b>More info</b> then <b>Run anyway</b>.</li>' +
    '<li>A small window opens and says it is waiting. Leave it open. This page finds it within a few seconds.</li>' +
    '</ol><p class="hint">Nothing is installed. The helper is a PowerShell script you can read. Closing its window stops it. A work PC that blocks PowerShell scripts cannot run it.</p>';
  if (S.platform === 'macos') return '<ol>' +
    '<li>Open <b>Terminal</b>. Press Command and Space, type Terminal and press Return.</li>' +
    '<li>Type <code>sh</code> and a space. Do not press Return yet.</li>' +
    '<li>In Finder open the output folder <b>' + esc(S.out.name) + '</b>, then ' + w + ', and drag <code>Start-SlideSize-Helper.command</code> into the Terminal window. Now press Return.</li>' +
    '<li>macOS asks whether Terminal may control Microsoft PowerPoint, and may ask about the folder. Allow both. PowerPoint may also ask once to be granted access to the folder.</li>' +
    '<li>Leave the Terminal window open. This page finds the helper within a few seconds.</li>' +
    '</ol><p class="hint">Nothing is installed. The helper is a shell script you can read. It is started from Terminal because macOS will not run a script that a browser has written when it is double-clicked. Closing the Terminal window stops it.</p>';
  return '<p class="hint">The helper runs on Windows and macOS.</p>';
}

function renderSetup() {
  var has = !!S.out;
  $('folder-name').textContent = has ? S.out.name : '';
  $('setup').hidden = !has;
  $('btn-folder').textContent = has ? 'Change the output folder' : 'Choose the output folder';
  $('btn-folder').disabled = !S.support.ok || S.running;
  $('btn-open-batch').disabled = !S.support.ok || S.running;
  if (!has) return;
  var st = C.helperStatus(S.helper, Date.now());
  function light(id, cls, text) { var el = $(id); el.className = 'light ' + cls; el.querySelector('span').textContent = text; }
  light('light-folder', 'ok', S.out.name);
  light('light-files', S.helperFiles === 'written' ? 'ok' : S.helperFiles === 'failed' ? 'bad' : 'wait',
    S.helperFiles === 'written' ? 'Written into ' + C.WORK_DIR : S.helperFiles === 'failed' ? 'Could not be written' : 'Not written');
  var running = st.state !== 'absent' && st.state !== 'stale' && st.state !== 'stopped' && st.state !== 'mismatch';
  light('light-helper', running ? 'ok' : st.state === 'mismatch' ? 'bad' : 'wait', running ? 'Running, version ' + (S.helper.helper || '') : st.state === 'stopped' ? 'Stopped' : st.state === 'stale' ? 'Not heard from' : st.state === 'mismatch' ? 'Wrong version' : 'Waiting for you to start it');
  var ppt = S.helper && S.helper.powerpoint;
  light('light-ppt', st.state === 'ready' ? 'ok' : st.state === 'no-powerpoint' || st.state === 'blocked' ? 'bad' : 'wait',
    st.state === 'ready' ? 'PowerPoint ' + (ppt.version || '') : st.state === 'no-powerpoint' ? 'Not installed' : st.state === 'blocked' ? 'Not responding' : st.state === 'starting' ? 'Starting' : 'Not checked yet');
  $('setup-steps').innerHTML = st.state === 'ready' ? '' : setupSteps();
  var msgs = '';
  if (S.helperFiles === 'failed') msgs += flag('warn', 'The browser would not let this page write the helper into the folder.', esc(S.helperFilesError),
    'Press Download the helper, unzip it, and copy the file for this system into the ' + C.WORK_DIR + ' folder of the output folder. Then start it as described above.');
  msgs += flag(st.level === 'ok' ? 'ok' : st.level === 'bad' ? 'bad' : 'info', '', esc(st.text),
    st.state === 'ready' && S.helper.powerpoint.userPresentations ? 'Close your own presentations before a long batch if you can. With them open, a file that stalls PowerPoint has to be cleared by hand.' : '');
  if (st.state === 'ready' && S.helper.engine && S.helper.engine !== C.ENGINE_NAME) msgs += flag('bad', 'This helper is not using PowerPoint.',
    'It reports its engine as ' + esc(S.helper.engine) + '. Results from it are marked as failed.');
  if (S.pendingFolderBatch) {
    var sum = C.resumeSummary(C.restoreBatch(S.pendingFolderBatch));
    msgs += flag('warn', 'This folder holds an unfinished batch.', sum.done + ' of ' + sum.total + ' files were finished.',
      'Resuming it replaces the list above. Starting a new batch here leaves its PDFs in place and replaces its record.',
      '<button type="button" class="btn small" id="btn-resume-folder">Resume that batch</button>');
  }
  $('helper-msgs').innerHTML = msgs;
  $('btn-helper-retry').hidden = st.state !== 'blocked';
  $('btn-helper-stop').hidden = !running || S.running;
  $('btn-write-helper').disabled = S.running;
}

/* ---------- what will be made ---------- */

function namePlan() {
  var items = S.batch.items.filter(function (it) { return it.status === 'ready' || it.status === 'blocked' && it.inspect && it.inspect.ok; });
  var existing = S.out ? S.out.existing.filter(function (n) { return !S.out.ours[n.toLowerCase()]; }) : [];
  /* names already given to finished files in this batch stay taken */
  var done = S.batch.items.filter(function (it) { return it.hasPdf && it.outName && items.indexOf(it) < 0; }).map(function (it) { return it.outName; });
  var plan = C.planNames(items, existing.concat(done), S.decisions);
  /* a file being converted again replaces its own earlier PDF */
  items.forEach(function (it) {
    /* so does a file that was interrupted after its name was settled. Anything under that name is this batch's own. */
    if ((it.hasPdf || it.interrupted || it.claimed) && it.outName) plan[it.id] = { outName: it.outName, renamed: false, reason: '', exists: false, decision: null, blocked: false, own: true, replaces: !!it.hasPdf };
  });
  return { items: items, plan: plan };
}

function renderPlan() {
  var np = namePlan(), body = $('plan-body');
  if (!np.items.length) { body.innerHTML = '<tr><td colspan="6" class="empty">Nothing to convert yet.</td></tr>'; return; }
  body.innerHTML = np.items.map(function (it) {
    var r = resolved(it), s = r.effective, p = C.planOutput(it.inspect, s), n = np.plan[it.id], notes = [];
    if (it.status === 'blocked') notes.push('<span class="chip bad">Held back</span>');
    if (n.renamed) notes.push(esc(n.reason));
    if (n.replaces) notes.push('Replaces its own earlier PDF.');
    if (n.exists) notes.push('<span style="color:var(--amber)">A file with this name is already in the folder.</span> <select data-decide="' + it.id + '" style="width:auto;padding:4px 6px;font-size:11px"' + (S.running ? ' disabled' : '') + '>' +
      [['', 'Choose'], ['overwrite', 'Overwrite it'], ['rename', 'Keep both'], ['skip', 'Skip this file']].map(function (o) { return '<option value="' + o[0] + '"' + ((n.decision || '') === o[0] ? ' selected' : '') + '>' + o[1] + '</option>'; }).join('') + '</select>');
    if (p.transform && p.transform.cropped) notes.push('<span style="color:var(--amber)">Cropped by ' + C.trimNum((p.transform.cropX || p.transform.cropY) * 100, 1) + '% on two sides.</span>');
    if (p.pages === 0) notes.push('<span style="color:#ff6b66">No slides are left with these settings.</span>');
    if (r.errors.length) notes.push('<span style="color:#ff6b66">' + esc(r.errors[0]) + '</span>');
    return '<tr><td class="name">' + esc(it.name) + '</td><td class="name">' + (n.decision === 'skip' ? '<span style="color:var(--muted)">Skipped</span>' : esc(n.outName)) + '</td>' +
      '<td class="num">' + (p.pages == null ? 'Known after opening' : p.pages) + '</td>' +
      '<td class="num">' + (p.known ? esc(C.sizeLabel(p.w, p.h)) + (p.changed ? '<br><span style="color:var(--muted)">from ' + esc(C.sizeLabel(p.srcW, p.srcH)) + '</span>' : '') : p.notes ? 'Notes page size, set in the deck' : 'Known after opening') + '</td>' +
      '<td>' + (p.known ? esc(p.orientation) : '') + '</td><td>' + notes.join(' ') + '</td></tr>';
  }).join('');
}

/* ------------------------------------------------------------------ *
 * 4 Convert                                                           *
 * ------------------------------------------------------------------ */

function startBlockers() {
  var out = [], np = namePlan(), ready = np.items.filter(function (it) { return it.status === 'ready'; });
  if (!S.support.ok) out.push('This browser cannot run the converter. Use Chrome or Edge.');
  if (S.platform === 'other') out.push('PowerPoint conversion needs a Windows PC or a Mac.');
  if (S.checking) out.push('Files are still being checked.');
  if (!ready.length) out.push(S.batch.items.length ? 'No file is ready to convert.' : 'Add presentations in section 1.');
  if (!S.out) out.push('Choose the output folder in section 3.');
  else {
    var st = C.helperStatus(S.helper, Date.now());
    if (st.state !== 'ready') out.push(st.state === 'absent' || st.state === 'stale' || st.state === 'stopped' ? 'Start the helper as described in section 3.' : st.text);
  }
  if (resolved(null).errors.length) out.push(resolved(null).errors[0]);
  var undecided = ready.filter(function (it) { return np.plan[it.id].blocked; }).length;
  if (undecided) out.push(C.plural(undecided, 'output name') + ' already ' + (undecided === 1 ? 'exists' : 'exist') + ' in the folder. Choose what to do in the list above.');
  var bad = ready.filter(function (it) { return resolved(it).errors.length; }).length;
  if (bad) out.push(C.plural(bad, 'file') + ' ' + (bad === 1 ? 'has' : 'have') + ' settings of ' + (bad === 1 ? 'its' : 'their') + ' own that are not valid.');
  return out;
}

function renderStart() {
  var blockers = startBlockers(), np = namePlan();
  var ready = np.items.filter(function (it) { return it.status === 'ready' && np.plan[it.id].decision !== 'skip'; }).length;
  var held = S.batch.items.filter(function (it) { return it.status === 'blocked'; }).length;
  $('p-convert').classList.toggle('idle', !S.running && blockers.length > 0);
  $('btn-start').disabled = S.running || blockers.length > 0;
  $('btn-start').textContent = S.running ? 'Converting' : ready ? 'Convert ' + C.plural(ready, 'presentation') : 'Convert';
  $('btn-start').hidden = S.running;
  var msgs = S.running ? '' : blockers.map(function (b) { return flag('info', '', esc(b)); }).join('');
  var linked = S.batch.items.filter(function (it) { return it.status === 'blocked' && it.inspect && it.inspect.ok && (it.issues || []).some(function (i) { return i.code === 'remote-links'; }); }).length;
  if (!S.running && linked) msgs += flag('warn', C.plural(linked, 'presentation') + ' ' + (linked === 1 ? 'links' : 'link') + ' to files on the internet or a network share.',
    'PowerPoint fetches those files as soon as it opens the deck, so ' + (linked === 1 ? 'it is' : 'they are') + ' held back until you allow it.', '',
    '<button type="button" class="btn small" id="btn-allow-remote">Allow linked files for this batch</button>');
  if (!S.running && !blockers.length && held - linked > 0) msgs += flag('warn', '', C.plural(held - linked, 'file') + ' cannot be converted. ' + (held - linked === 1 ? 'It is' : 'They are') + ' marked in section 2 with the reason.');
  $('start-msgs').innerHTML = msgs;
  $('progress').hidden = !S.running;
}

function renderProgress() {
  if (!S.running) return;
  var items = S.batch.items.filter(function (it) { return S.runSet[it.id]; });
  var done = items.filter(function (it) { return it.status === 'done' || it.status === 'failed' || it.status === 'cancelled'; });
  var sum = C.summarize(done);
  $('progress-stage').textContent = S.stopping ? 'Stopping after the file in hand' : 'Converting';
  $('progress-detail').textContent = done.length + ' of ' + items.length + ' done. ' + sum.completed + ' completed, ' + sum.warnings + ' with warnings, ' + sum.review + ' to review, ' + sum.failed + ' failed';
  $('progress-fill').style.width = (items.length ? Math.round(done.length / items.length * 100) : 0) + '%';
  var cur = items.filter(function (it) { return it.status === 'converting'; })[0], lines = [];
  var st = C.helperStatus(S.helper, Date.now());
  if (st.state !== 'ready') lines.push('<span style="color:var(--amber)">' + esc(st.text) + ' The batch waits.</span>');
  if (cur) {
    var since = S.helper && S.helper.current && S.helper.current.id === cur.id ? Date.now() - Date.parse(S.helper.current.since) : 0;
    var phase = ({ opening: 'opening it', exporting: 'writing the PDF', reference: 'saving pictures of the slides', validating: 'validating PDF/A', closing: 'closing it' })[cur.phase] || 'working';
    lines.push('PowerPoint  ' + esc(cur.name) + ', ' + phase + ', ' + C.formatDuration(since));
  }
  var staging = items.filter(function (it) { return it.status === 'staging'; })[0];
  if (staging) lines.push('Copying     ' + esc(staging.name));
  if (S.verifying) lines.push('Checking    ' + esc(S.verifying));
  var queued = items.filter(function (it) { return it.status === 'queued'; }).length, waiting = items.filter(function (it) { return it.status === 'ready'; }).length;
  lines.push('Queued ' + queued + ', waiting ' + waiting);
  $('progress-now').innerHTML = lines.join('<br>');
  $('btn-stop').disabled = S.stopping;
  $('btn-skip').disabled = !cur;
}

async function startRun(ids) {
  if (S.running) return;
  var np = namePlan(), set = {}, n = 0;
  np.items.forEach(function (it) {
    if (it.status !== 'ready' || (ids && ids.indexOf(it.id) < 0)) return;
    var p = np.plan[it.id];
    if (p.decision === 'skip') { it.status = 'skipped'; return; }
    it.outName = p.outName; it.overwrite = !!p.own || p.decision === 'overwrite';
    it.claimed = true;                      /* from here on, a file of this name in the folder is this batch's own */
    set[it.id] = true; n++;
  });
  if (!n) { renderAll(); return; }
  S.runSet = set; S.running = true; S.stopping = false; S.verifyChain = Promise.resolve();
  if (navigator.locks && navigator.locks.request) {
    /* held for the length of the run. It stops a second tab running the same batch and tells the browser this tab is busy. */
    navigator.locks.request('slidesize-pptpdf-run', { ifAvailable: true }, function (lock) {
      if (!lock) return null;
      return new Promise(function (res) { S.releaseLock = res; });
    });
  }
  S.batch.platform = enginePlatform();
  S.batch.reportName = S.batch.reportName || 'SlideSize report ' + stamp(new Date(S.batch.createdAt));
  storeHandles();
  await saveAll();
  renderAll();
}

async function stage(item) {
  S.stagingNow = true; item.status = 'staging'; item.phase = ''; mark('files', 'progress');
  try {
    var sub = S.out.sub, s = resolved(item).effective;
    for (var d of ['queue', 'done']) await removeEntry(sub[d], item.id + '.json');
    await removeEntry(sub.done, item.id + 'v.json');
    await removeEntry(sub.ref, item.id, true);
    var file = await sourceFile(item);
    var ticket = C.buildTicket(item, s);
    var name = ticket.input.replace(/^in\//, '');
    var h = await sub.in.getFileHandle(name, { create: true }), w = await h.createWritable();
    await file.stream().pipeTo(w);                     /* a stream, so a 2 GB deck never sits in memory */
    item.result = null; item.checks = null; item.timing = null; item.hasPdf = false; item.pdfBytes = 0;
    item.issues = (item.issues || []).filter(function (i) { return i.stage === 'check'; });
    item.attempts = (item.attempts || 0) + 1; item.settingsUsed = s; delete item.interrupted;
    await writeFile(sub.queue, item.id + '.json', JSON.stringify(ticket));
    item.status = 'queued';
  } catch (e) {
    item.status = 'failed'; item.hasPdf = false;
    item.issues = C.mergeIssues([(item.issues || []).filter(function (i) { return i.stage === 'check'; }), [C.issue('source-unreadable', { severity: 'error', stage: 'convert',
      description: 'The presentation could not be copied for conversion. ' + (e && e.message ? e.message : String(e)),
      impact: 'It was not converted. The rest of the batch carried on.',
      fix: 'Check the file still exists and the disk has room, then retry it. If access has to be confirmed again, add the file again.' })]]);
    await saveAll();
  } finally { S.stagingNow = false; mark('files', 'progress', 'results'); }
}

async function pump() {
  var items = S.batch.items, sub = S.out.sub, ready = helperFresh() && C.helperStatus(S.helper, Date.now()).state === 'ready';
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    if (it.status !== 'queued' && it.status !== 'converting') continue;
    var r = await readJson(sub.done, it.id + '.json');
    if (r && r.id === it.id) {
      /* the result file stays on disk until the PDF has been checked, so a tab closed in between loses nothing */
      it.result = r; it.status = 'verifying'; it.phase = '';
      queueVerify(it);
    } else if (S.helper && S.helper.current && S.helper.current.id === it.id) { it.status = 'converting'; it.phase = S.helper.current.phase; }
  }
  if (!S.stopping && ready && !S.stagingNow) {
    var ahead = items.filter(function (x) { return x.status === 'staging' || x.status === 'queued' || x.status === 'converting'; }).length;
    if (ahead < LOOKAHEAD) {
      var next = items.filter(function (x) { return x.status === 'ready' && S.runSet[x.id]; })[0];
      if (next) stage(next);
    }
  }
  var busy = items.some(function (x) { return IN_FLIGHT[x.status] || (!S.stopping && x.status === 'ready' && S.runSet[x.id]); });
  if (!busy && !S.stagingNow) await finishRun();
  mark('progress', 'files', 'results');
}

async function finishRun() {
  await S.verifyChain;
  S.running = false; S.stopping = false; S.verifying = '';
  if (S.releaseLock) { S.releaseLock(); S.releaseLock = null; }
  await saveAll();
  renderAll();
}

async function stopAfterCurrent() {
  if (!S.running) return;
  S.stopping = true;
  var helperAlive = helperFresh();
  for (var it of S.batch.items) {
    /* a ticket the helper has not picked up yet is taken back */
    if (it.status === 'queued' && !(S.helper && S.helper.current && S.helper.current.id === it.id)) {
      if (await removeEntry(S.out.sub.queue, it.id + '.json') || !helperAlive) {
        await removeEntry(S.out.sub.in, 'slidesize-' + it.id + '.' + it.ext);
        it.status = 'ready';
      }
    } else if ((it.status === 'converting' || it.status === 'queued') && !helperAlive) {
      it.status = 'ready'; it.interrupted = true;       /* nothing is running it. It is converted again next time */
    }
  }
  mark('progress', 'files');
}

async function skipCurrent() {
  var cur = S.batch.items.filter(function (it) { return it.status === 'converting'; })[0];
  if (cur) await sendControl({ abort: cur.id });
}

/* ------------------------------------------------------------------ *
 * Checking the finished PDF                                           *
 * ------------------------------------------------------------------ */

var pdfjsReady = null;
function ensurePdfjs() {
  if (pdfjsReady) return pdfjsReady;
  pdfjsReady = loadScript(PDFJS + 'pdf.min.js').then(function () { window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS + 'pdf.worker.min.js'; });
  pdfjsReady.catch(function () { pdfjsReady = null; });
  return pdfjsReady;
}
function openPdfjs(bytes) {
  return window.pdfjsLib.getDocument({ data: bytes, cMapUrl: PDFJS + 'cmaps/', cMapPacked: true, standardFontDataUrl: PDFJS + 'standard_fonts/' }).promise;
}
async function renderPdfPage(doc, n, longEdge) {
  var page = await doc.getPage(n), base = page.getViewport({ scale: 1 });
  var vp = page.getViewport({ scale: longEdge / Math.max(base.width, base.height) });
  var cv = document.createElement('canvas'); cv.width = Math.max(1, Math.round(vp.width)); cv.height = Math.max(1, Math.round(vp.height));
  var cx = cv.getContext('2d', { alpha: false }); cx.fillStyle = '#fff'; cx.fillRect(0, 0, cv.width, cv.height);
  await page.render({ canvasContext: cx, viewport: vp }).promise;
  page.cleanup();
  return cv;
}
function regionData(src, sw, sh, win, w, h) {
  var cv = document.createElement('canvas'); cv.width = w; cv.height = h;
  var cx = cv.getContext('2d', { willReadFrequently: true });
  cx.fillStyle = '#fff'; cx.fillRect(0, 0, w, h);
  cx.imageSmoothingQuality = 'high';
  cx.drawImage(src, win.x * sw, win.y * sh, win.w * sw, win.h * sh, 0, 0, w, h);
  return cx.getImageData(0, 0, w, h);
}
function refName(slide) { return 'slide-' + ('0000' + slide).slice(-4) + '.png'; }
var COMPARE_METHOD = 'Each compared PDF page was drawn in the browser with pdf.js and set against the picture PowerPoint saved of the same slide, both at ' + COMPARE_W +
  ' pixels wide and cut into 8 pixel blocks. Edge smoothing, gradients and shadows differ between the two programs, so small differences are ignored. A flagged page is a suspicion and a clean page is not proof.';

function windowsFor(tp) {
  return tp && tp.scale ? C.compareWindows(tp.srcW, tp.srcH, tp.w, tp.h, { scale: tp.scale, tx: tp.tx, ty: tp.ty }) : { ref: { x: 0, y: 0, w: 1, h: 1 }, pdf: { x: 0, y: 0, w: 1, h: 1 } };
}

async function comparePair(bmp, pdfCanvas, tp, sensitivity) {
  var win = windowsFor(tp);
  var aspect = (bmp.height * win.ref.h) / (bmp.width * win.ref.w), w = COMPARE_W, h = Math.max(8, Math.round(w * aspect));
  var a = regionData(bmp, bmp.width, bmp.height, win.ref, w, h), b = regionData(pdfCanvas, pdfCanvas.width, pdfCanvas.height, win.pdf, w, h);
  var r = C.compareImages(a.data, b.data, w, h, sensitivity);
  r.w = w; r.h = h; r.window = win;
  return r;
}

async function compareItem(item, res, bytes, info, tinfo, s) {
  if (s.fidelity === 'off') return null;
  if (s.output === 'notes') return { performed: false, reason: 'Notes pages are a different layout from the slide, so they are not compared.' };
  var ref = res.reference, map = res.export && res.export.pageMap;
  if (!ref || !ref.slides || !ref.slides.length) return { performed: false, reason: (ref && ref.note) || 'PowerPoint did not save pictures of the slides.' };
  if (!map) return { performed: false, reason: 'The helper could not say which slide is on which page.' };
  if (info.pageCount !== map.length) return { performed: false, reason: 'The page count does not match the slides, so pages cannot be matched to slides.' };
  try { await ensurePdfjs(); } catch (e) { return { performed: false, reason: 'The PDF drawing library could not be loaded. ' + e.message }; }
  var doc, dir, out = { performed: true, method: COMPARE_METHOD, sensitivity: s.sensitivity, pages: [], skipped: [], totalPages: map.length };
  try { doc = await openPdfjs(bytes); dir = await S.out.sub.ref.getDirectoryHandle(item.id); }
  catch (e) { return { performed: false, reason: 'The PDF or the reference pictures could not be opened. ' + (e.message || e) }; }
  try {
    for (var i = 0; i < ref.slides.length; i++) {
      var slide = ref.slides[i], page = map.indexOf(slide) + 1;
      if (!page) continue;
      S.verifying = item.name + ', comparing page ' + page + ' of ' + map.length; mark('progress');
      try {
        var bmp = await createImageBitmap(await (await dir.getFileHandle(refName(slide))).getFile());
        var cv = await renderPdfPage(doc, page, 640);
        var r = await comparePair(bmp, cv, tinfo && tinfo.pages ? tinfo.pages[page - 1] : null, s.sensitivity);
        bmp.close(); cv.width = cv.height = 0;
        out.pages.push({ slide: slide, page: page, changedShare: +r.changedShare.toFixed(4), verdict: r.verdict });
      } catch (e) { out.skipped.push({ slide: slide, page: page, reason: 'Its picture could not be read. ' + (e.message || e) }); }
      if (i % 4 === 3) await pause(0);
    }
  } finally { try { await doc.destroy(); } catch (e) { /* already gone */ } }
  /* the pictures stay on disk only where there is something to look at */
  if (!s.keepReferences) {
    for (var k = 0; k < out.pages.length; k++) if (out.pages[k].verdict !== 'different') await removeEntry(dir, refName(out.pages[k].slide));
  }
  out.kept = s.keepReferences ? 'all' : 'flagged';
  if (!s.keepReferences && !out.pages.some(function (p) { return p.verdict === 'different'; })) await removeEntry(S.out.sub.ref, item.id, true);
  return out;
}

function mediaHandling(item, res, s) {
  var out = [], ins = item.inspect, map = res.export && res.export.pageMap, ph = {};
  ((res.export && res.export.placeholders) || []).forEach(function (p) { ph[p.slide] = p.ok ? 'ok' : 'failed'; });
  if (ins && ins.slides) ins.slides.forEach(function (sl) {
    (sl && sl.media || []).forEach(function (m) {
      var how;
      if (map && map.indexOf(sl.n) < 0) how = 'not in the PDF, the slide is excluded';
      else if (m.type === 'audio') how = 'no sound in a PDF';
      else if (m.posterPart && !m.posterBlank) how = 'poster frame kept at the same position and size, no playback';
      else if (ph[sl.n] === 'ok') how = 'no usable poster frame, marked placeholder drawn';
      else if (ph[sl.n] === 'failed') how = 'no usable poster frame, placeholder could not be drawn';
      else how = 'no usable poster frame, drawn as PowerPoint shows it';
      out.push({ slide: sl.n, type: m.type, name: m.name, linked: !!m.linked, handling: how });
    });
  });
  ((res.facts && res.facts.media) || []).forEach(function (m) { out.push({ slide: m.slide, type: m.type, name: m.name, handling: m.type === 'audio' ? 'no sound in a PDF' : 'poster frame, no playback' }); });
  return out;
}

async function waitForResult(id, ms) {
  var end = Date.now() + ms;
  while (Date.now() < end) {
    var r = await readJson(S.out.sub.done, id + '.json');
    if (r && r.id === id) { await removeEntry(S.out.sub.done, id + '.json'); return r; }
    if (!helperFresh()) await pollHelper();
    await pause(500);
  }
  return null;
}

async function verifyItem(item) {
  var res = item.result || {}, s = item.settingsUsed || resolved(item).effective, plat = enginePlatform();
  var pre = (item.issues || []).filter(function (i) { return i.stage === 'check'; });
  S.verifying = item.name; mark('progress');
  item.timing = res.timing || null;
  if (!res.ok) {
    item.status = res.status === 'cancelled' ? 'cancelled' : 'failed'; item.hasPdf = false;
    item.issues = C.mergeIssues([pre, [C.helperFailureIssue(res)]]);
    return;
  }
  if (res.engine && res.engine.name !== C.ENGINE_NAME) {
    item.status = 'failed'; item.hasPdf = false;
    item.issues = C.mergeIssues([pre, C.resultIssues(item, { settings: s, result: res, platform: plat })]);
    return;
  }
  var facts = res.facts || {}, plan;
  var known = { ok: true, slideCount: facts.slides, hiddenSlides: facts.hidden || (item.inspect && item.inspect.hiddenSlides) || null,
                widthPt: facts.slideWidthPt || (item.inspect && item.inspect.widthPt), heightPt: facts.slideHeightPt || (item.inspect && item.inspect.heightPt) };
  plan = C.planOutput(known, s);
  var edits = [], ops = {};
  if (s.size.preset !== 'original') edits.push('page size');
  if (s.links === 'remove') { edits.push('link removal'); ops.removeLinks = true; }
  if (s.metadata === 'strip') { edits.push('metadata removal'); ops.stripMetadata = true; }
  var ctx = { settings: s, platform: plat, plan: plan, result: res }, cmp = null, pdfa = null, info = null, tinfo = null;
  try {
    var fromWork = edits.length > 0, handle;
    if (fromWork) {
      try { handle = await S.out.sub.out.getFileHandle(item.id + '.pdf'); }
      catch (e) {
        /* the edit was finished before an interruption. The file in the folder is already the edited one. */
        handle = await S.out.dir.getFileHandle(item.outName); fromWork = false;
      }
    } else handle = await S.out.dir.getFileHandle(item.outName);
    var file = await handle.getFile();
    item.pdfBytes = file.size;
    if (file.size > C.PDF_EDIT_MAX_BYTES) {
      if (fromWork) {
        /* too large to edit here. It still goes into the output folder, as PowerPoint wrote it. */
        var dst = await S.out.dir.getFileHandle(item.outName, { create: true }), dw = await dst.createWritable();
        await file.stream().pipeTo(dw);
        await removeEntry(S.out.sub.out, item.id + '.pdf');
        ctx.editsSkipped = edits;
      }
      ctx.pdfSkipped = file.size; item.hasPdf = true;
    } else {
      var bytes = new Uint8Array(await file.arrayBuffer());
      info = await P.inspect(bytes);
      if (fromWork) {
        if (s.size.preset !== 'original') {
          var t = C.targetSize(info.pages[0].w, info.pages[0].h, s.size);
          if (t.changed) ops.resize = { w: t.w, h: t.h, mode: s.size.fit, margin: s.size.margin };
        }
        if (ops.resize || ops.removeLinks || ops.stripMetadata) {
          tinfo = await P.transform(bytes, ops);
          bytes = tinfo.bytes;
          ctx.edits = { linksRemoved: tinfo.linksRemoved, metadataStripped: tinfo.metadataStripped, resized: tinfo.resized };
          if (ops.resize) {
            var p0 = tinfo.pages[0];
            ctx.transform = p0 && p0.scale ? C.fitTransform(p0.srcW, p0.srcH, p0.w, p0.h, s.size.fit) : null;
            ctx.plan = Object.assign({}, plan, { known: true, w: ops.resize.w, h: ops.resize.h });
          }
        }
        await writeFile(S.out.dir, item.outName, bytes);
        await removeEntry(S.out.sub.out, item.id + '.pdf');
        item.pdfBytes = bytes.length;
        info = await P.inspect(bytes);                 /* the finished file is the one that is checked */
      } else if (s.output === 'slides' && plan.known) {
        ctx.plan = plan;
      }
      item.hasPdf = true;
      ctx.pdf = info;
      ctx.fonts = C.fontFindings((item.inspect && item.inspect.fonts) || [], S.fonts, (item.inspect && item.inspect.embeddedFonts) || [], facts.fonts || null, info.fonts, s, plat, (item.inspect && item.inspect.slideFonts) || null);
      if (s.pdfa) {
        pdfa = P.pdfaChecks(info);
        var v = res.pdfa;
        if (!v && S.helper && S.helper.verapdf && fromWork) {
          /* the file changed after PowerPoint wrote it, so the validator is asked about the finished file */
          await writeFile(S.out.sub.queue, item.id + 'v.json', JSON.stringify({ protocol: C.PROTOCOL, id: item.id + 'v', task: 'validate', pdf: '../' + item.outName, sourceName: item.outName, timeoutSec: 600, options: {} }));
          var vr = await waitForResult(item.id + 'v', 660000);
          v = vr && vr.pdfa;
        }
        if (v && (v.result === 'passed' || v.result === 'failed')) { pdfa.result = v.result; pdfa.validator = v.validator; pdfa.profile = v.profile || pdfa.profile; pdfa.detail = v.detail || ''; }
        else if (v) pdfa.reason = 'veraPDF is installed but did not give a verdict. ' + (v.detail || '');
        else pdfa.reason = S.helper && S.helper.verapdf ? 'veraPDF did not answer.' : 'A complete check needs veraPDF, a free validator, and it was not found on this computer.';
      }
      cmp = await compareItem(item, res, bytes, info, tinfo, s);
      bytes = null;
    }
  } catch (e) {
    if (!info) ctx.pdfError = e && e.message ? e.message : String(e);
    else ctx.verifyError = e && e.message ? e.message : String(e);
    try { item.hasPdf = item.hasPdf || await hasFile(S.out.dir, item.outName); } catch (x) { /* unknown */ }
  }
  var lists = [pre, C.resultIssues(item, ctx), C.pdfaIssues(pdfa), C.compareIssues(cmp, s)];
  if (ctx.verifyError) lists.push([C.issue('verify-failed', { severity: 'warning', certainty: 'unverified', stage: 'verify',
    description: 'Some checks on the PDF stopped early. ' + ctx.verifyError, impact: 'The PDF is in the folder but was not fully checked.', fix: 'Open it and look through it.' })]);
  item.issues = C.mergeIssues(lists);
  var all = []; for (var n = 1; n <= (facts.slides || 0); n++) all.push(n);
  var map = res.export && res.export.pageMap;
  item.checks = {
    pageCount: info ? info.pageCount : null, pageSize: info ? { w: info.pages[0].w, h: info.pages[0].h } : null, tagged: info ? info.tagged : null,
    excludedSlides: map ? all.filter(function (x) { return map.indexOf(x) < 0; }) : [],
    fonts: ctx.fonts ? { missing: ctx.fonts.missing, restricted: ctx.fonts.restricted, notEmbedded: ctx.fonts.notEmbedded, absent: ctx.fonts.absent, substitutes: ctx.fonts.substitutes,
                         inPdf: info.fonts.map(function (f) { return f.name.replace(/^[A-Z]{6}\+/, '') + (f.embedded ? '' : ' (not embedded)'); }) } : null,
    media: mediaHandling(item, res, s), fidelity: cmp, pdfa: pdfa, edits: ctx.edits || null,
    transform: tinfo && tinfo.pages ? tinfo.pages.map(function (p) { return p.scale ? { srcW: p.srcW, srcH: p.srcH, w: p.w, h: p.h, scale: p.scale, tx: p.tx, ty: p.ty } : null; }) : null
  };
  item.result = { ok: res.ok, status: res.status, engine: res.engine, facts: { slides: facts.slides, hidden: facts.hidden, slideWidthPt: facts.slideWidthPt, slideHeightPt: facts.slideHeightPt },
                  export: res.export, reference: res.reference, timing: res.timing, startedAt: res.startedAt, finishedAt: res.finishedAt };
  item.status = item.hasPdf ? 'done' : 'failed';
  if (item.hasPdf && S.out) { S.out.ours[item.outName.toLowerCase()] = true; if (S.out.existing.indexOf(item.outName) < 0) S.out.existing.push(item.outName); }
  if (!cmp || cmp.performed === false) await removeEntry(S.out.sub.ref, item.id, true);
}

function queueVerify(item) {
  S.verifyChain = S.verifyChain.then(function () { return verifyItem(item); }).catch(function (e) {
    item.status = 'failed';
    item.issues = C.mergeIssues([(item.issues || []).filter(function (i) { return i.stage === 'check'; }), [C.issue('verify-crashed', { severity: 'error', certainty: 'unverified', stage: 'verify',
      description: 'The check of this file stopped with an error. ' + (e && e.message ? e.message : e), impact: 'Its result is unknown.', fix: 'Retry it.' })]]);
  }).then(function () { S.verifying = ''; return saveAll(); })
    .then(function () { return removeEntry(S.out.sub.done, item.id + '.json'); })
    .then(function () { mark('files', 'results', 'progress', 'plan', 'save'); });
}

/* ------------------------------------------------------------------ *
 * Saving the record and the reports                                   *
 * ------------------------------------------------------------------ */

function saveAll() {
  if (!S.out) return Promise.resolve();
  S.saveChain = S.saveChain.then(async function () {
    var b = S.batch, name = b.reportName || 'SlideSize report ' + stamp(new Date(b.createdAt));
    await writeFile(S.out.work, 'batch.json', C.serializeBatch(b));
    await writeFile(S.out.dir, name + '.html', C.reportHtml(b));
    await writeFile(S.out.dir, name + '.csv', C.reportCsv(b));
    await writeFile(S.out.dir, name + '.json', C.reportJson(b));
  }).catch(function (e) { S.saveError = e && e.message ? e.message : String(e); });
  return S.saveChain;
}

/* ------------------------------------------------------------------ *
 * 5 Results                                                           *
 * ------------------------------------------------------------------ */

function finished() { return S.batch.items.filter(function (it) { return it.status === 'done' || it.status === 'failed' || it.status === 'cancelled'; }); }

function renderResults() {
  var done = finished(), sum = C.summarize(S.batch.items);
  $('p-results').classList.toggle('idle', !done.length);
  $('tiles').innerHTML = [['Finished', sum.finished, ''], ['Completed', sum.completed, 'good'], ['With warnings', sum.warnings, 'warn'], ['Needs review', sum.review, 'rev'], ['Failed', sum.failed, 'bad'], ['Not converted', sum.pending + sum.skipped + sum.blocked, '']]
    .map(function (t) { return '<div class="tile ' + t[2] + '"><b>' + t[1] + '</b><span>' + t[0] + '</span></div>'; }).join('');
  var body = $('results-body');
  var rows = done.filter(function (it) { return S.filter === 'all' || C.classify(it) === S.filter; });
  if (!done.length) { body.innerHTML = '<tr><td colspan="7" class="empty">Results appear here as each file finishes.</td></tr>'; $('btn-retry').disabled = true; return; }
  if (!rows.length) { body.innerHTML = '<tr><td colspan="7" class="empty">No file has this result.</td></tr>'; }
  else body.innerHTML = rows.map(function (it) {
    var st = C.classify(it), c = issueCounts(it), open = (it.issues || []).filter(function (i) { return !i.accepted && (i.review || i.severity === 'error'); }).length;
    var row = '<tr class="main" data-res="' + it.id + '"><td><input type="checkbox" data-pick="' + it.id + '"' + (S.selected[it.id] ? ' checked' : '') + ' aria-label="Select ' + esc(it.name) + '"></td>' +
      '<td class="name">' + esc(it.name) + (it.hasPdf ? '<small>' + esc(it.outName) + ', ' + esc(C.formatBytes(it.pdfBytes)) + '</small>' : '<small>No PDF</small>') + '</td>' +
      '<td><span class="st ' + st + '">' + esc(C.STATES[st]) + '</span></td>' +
      '<td class="num">' + (it.checks && it.checks.pageCount != null ? it.checks.pageCount : '') + '</td>' +
      '<td><div class="chips">' + (c.error ? '<span class="chip bad">' + C.plural(c.error, 'error') + '</span>' : '') + (c.warning ? '<span class="chip warn">' + C.plural(c.warning, 'warning') + '</span>' : '') +
        (c.info ? '<span class="chip dim">' + C.plural(c.info, 'note') + '</span>' : '') + (open ? '<span class="chip warn">' + open + ' to review</span>' : '') + (!c.error && !c.warning && !c.info ? '<span class="chip good">None</span>' : '') + '</div></td>' +
      '<td class="num">' + esc(it.timing ? C.formatDuration(it.timing.totalMs) : '') + '</td>' +
      '<td><div class="btn-row">' + (it.hasPdf ? '<button type="button" class="btn small" data-open="' + it.id + '">Open PDF</button><button type="button" class="btn small" data-save="' + it.id + '">Save a copy</button>' : '') +
        (it.hasPdf && it.checks && it.checks.pageCount ? '<button type="button" class="btn small" data-view="' + it.id + '|1">Side by side</button>' : '') + '</div></td></tr>';
    if (S.resExpanded[it.id]) {
      var list = (it.issues || []).filter(function (i) { return S.sev === 'all' || i.severity === 'error' || (S.sev === 'warning' && i.severity === 'warning'); });
      var ck = it.checks || {}, s = it.settingsUsed || settingsOf(it), res = it.result || {};
      row += '<tr class="detail"><td colspan="7"><dl class="kv" style="margin-bottom:10px">' +
        '<dt>Settings</dt><dd style="font-family:var(--sans)">' + esc(C.describeSettings(s).join('. ')) + '.</dd>' +
        (res.engine ? '<dt>Engine</dt><dd>' + esc(res.engine.name + ' ' + (res.engine.version || '') + (res.engine.build ? ' build ' + res.engine.build : '')) + (res.export && res.export.method ? ', ' + esc(res.export.method) : '') + '</dd>' : '') +
        (ck.pageSize ? '<dt>PDF</dt><dd>' + C.plural(ck.pageCount, 'page') + ', ' + esc(C.sizeLabel(ck.pageSize.w, ck.pageSize.h)) + ', ' + (ck.tagged ? 'tagged' : 'not tagged') + '</dd>' : '') +
        (ck.excludedSlides && ck.excludedSlides.length ? '<dt>Excluded slides</dt><dd>' + esc(C.listSlides(ck.excludedSlides, 30)) + '</dd>' : '') +
        (ck.fonts ? '<dt>Fonts in the PDF</dt><dd style="font-family:var(--sans)">' + esc(ck.fonts.inPdf.join(', ') || 'None') + '</dd>' : '') +
        '<dt>Fidelity check</dt><dd style="font-family:var(--sans)">' + esc(C.fidelityLabel(ck.fidelity, s)) + '</dd>' +
        '<dt>PDF/A</dt><dd style="font-family:var(--sans)">' + esc(C.pdfaLabel(ck.pdfa, s)) + (ck.pdfa && ck.pdfa.validator ? ', ' + esc(ck.pdfa.validator) : '') + '</dd>' +
        '</dl>' + (list.length ? '<div class="issues">' + list.map(function (i) { return issueHtml(it, i, true); }).join('') + '</div>' : '<p class="hint">No issues at this level.</p>') +
        (open ? '<div class="btn-row" style="margin-top:10px"><button type="button" class="btn small" data-acceptall="' + it.id + '">Mark everything on this file as reviewed</button></div>' : '') + '</td></tr>';
    }
    return row;
  }).join('');
  var picked = done.filter(function (it) { return S.selected[it.id]; }).length;
  $('btn-retry').disabled = S.running || !picked;
  $('btn-retry').textContent = picked ? 'Retry ' + C.plural(picked, 'selected file') : 'Retry selected';
  $('res-all').checked = rows.length > 0 && rows.every(function (it) { return S.selected[it.id]; });
}

async function acceptIssues(id, issueId) {
  var item = S.batch.items.filter(function (i) { return i.id === id; })[0];
  if (!item) return;
  var now = new Date().toISOString();
  item.issues.forEach(function (i) { if ((issueId === null && (i.review || i.severity === 'error') && i.code !== 'engine-unknown') || i.id === issueId) i.accepted = i.accepted || now; });
  mark('results', 'files');
  await saveAll();
}

async function acceptSlide(id, slide) {
  var item = S.batch.items.filter(function (i) { return i.id === id; })[0];
  if (!item) return;
  var now = new Date().toISOString();
  item.issues.forEach(function (i) { if (i.slide === slide && (i.review || i.severity === 'error')) i.accepted = i.accepted || now; });
  mark('results', 'files');
  await saveAll();
}

async function openPdf(id, saveCopy) {
  var item = S.batch.items.filter(function (i) { return i.id === id; })[0];
  if (!item || !item.hasPdf) return;
  try {
    var file = await (await S.out.dir.getFileHandle(item.outName)).getFile();
    if (saveCopy) { download(file, item.outName); return; }
    var url = URL.createObjectURL(new Blob([file], { type: 'application/pdf' }));
    window.open(url, '_blank', 'noopener');
    setTimeout(function () { URL.revokeObjectURL(url); }, 600000);
  } catch (e) { $('save-msgs').innerHTML = flag('bad', '', 'The PDF could not be opened from the folder. ' + esc(e.message || String(e))); }
}

async function retrySelected() {
  if (S.running) return;
  var ids = [];
  finished().forEach(function (it) {
    if (!S.selected[it.id]) return;
    it.status = 'ready'; delete S.selected[it.id];
    it.issues = (it.issues || []).filter(function (i) { return i.stage === 'check'; });
    recomputeIssues(it);
    if (it.status === 'ready') ids.push(it.id);
  });
  if (ids.length && !startBlockers().length) await startRun(ids);
  else renderAll();
}

/* ---------- side by side ---------- */

var viewCache = new C.LRU(8, function (v) { if (v && v.close) v.close(); else if (v && v.width) v.width = v.height = 0; });
var viewDoc = { id: null, doc: null };

async function viewerDoc(item) {
  if (viewDoc.id === item.id && viewDoc.doc) return viewDoc.doc;
  if (viewDoc.doc) { try { await viewDoc.doc.destroy(); } catch (e) { /* gone */ } viewDoc.doc = null; }
  await ensurePdfjs();
  var file = await (await S.out.dir.getFileHandle(item.outName)).getFile();
  viewDoc.doc = await openPdfjs(new Uint8Array(await file.arrayBuffer())); viewDoc.id = item.id;
  return viewDoc.doc;
}

async function showViewer(id, page) {
  var item = S.batch.items.filter(function (i) { return i.id === id; })[0];
  if (!item || !item.hasPdf) return;
  var pages = (item.checks && item.checks.pageCount) || 1;
  page = C.clamp(page, 1, pages);
  S.viewer = { id: id, page: page };
  $('viewer').hidden = false;
  var map = item.result && item.result.export && item.result.export.pageMap, slide = map ? map[page - 1] : null;
  $('viewer-title').textContent = item.name + ', PDF page ' + page + ' of ' + pages + (slide ? ', slide ' + slide : '');
  $('viewer-prev').disabled = page <= 1; $('viewer-next').disabled = page >= pages;
  var flaggedPages = (item.issues || []).filter(function (i) { return i.code === 'visual-diff' && !i.accepted; }).map(function (i) { return i.page; });
  $('viewer-flagged').disabled = !flaggedPages.length;
  $('viewer-accept').disabled = slide == null || !(item.issues || []).some(function (i) { return i.slide === slide && !i.accepted && (i.review || i.severity === 'error'); });
  var cmp = item.checks && item.checks.fidelity, rec = cmp && cmp.pages ? cmp.pages.filter(function (p) { return p.page === page; })[0] : null;
  ['viewer-ref', 'viewer-pdf', 'viewer-diff'].forEach(function (x) { $(x).innerHTML = '<div class="none"><span><span class="spinner"></span>Drawing</span></div>'; });
  $('viewer-msgs').innerHTML = '';
  var token = S.viewer;
  try {
    var key = id + ':' + page, cv = viewCache.get(key + ':pdf');
    if (!cv) { cv = await renderPdfPage(await viewerDoc(item), page, 1100); viewCache.set(key + ':pdf', cv); }
    if (S.viewer !== token) return;
    function show(target, source) { var c = document.createElement('canvas'); c.width = source.width; c.height = source.height; c.getContext('2d').drawImage(source, 0, 0); $(target).innerHTML = ''; $(target).appendChild(c); }
    show('viewer-pdf', cv);
    var bmp = null;
    if (slide != null) {
      bmp = viewCache.get(key + ':ref');
      if (!bmp) {
        try { bmp = await createImageBitmap(await (await (await S.out.sub.ref.getDirectoryHandle(item.id)).getFileHandle(refName(slide))).getFile()); viewCache.set(key + ':ref', bmp); }
        catch (e) { bmp = null; }
      }
    }
    if (S.viewer !== token) return;
    if (!bmp) {
      var why = !cmp ? 'The fidelity check was off for this file, so PowerPoint saved no picture of the slide.'
        : cmp.performed === false ? 'The comparison was not run. ' + (cmp.reason || '')
        : rec && rec.verdict !== 'different' ? 'No difference above the threshold was found on this page, so the picture was not kept. Turn on Keep every reference picture under Advanced to keep them all.'
        : 'This page was not one of those compared.';
      $('viewer-ref').innerHTML = '<div class="none">' + esc(why) + '</div>';
      $('viewer-diff').innerHTML = '<div class="none">Nothing to compare with.</div>';
    } else {
      show('viewer-ref', bmp);
      var tp = item.checks.transform ? item.checks.transform[page - 1] : null;
      var r = await comparePair(bmp, cv, tp, (item.settingsUsed || S.batch.settings).sensitivity);
      var win = r.window.pdf, d = document.createElement('canvas'); d.width = cv.width; d.height = cv.height;
      var dx = d.getContext('2d'); dx.drawImage(cv, 0, 0); dx.fillStyle = 'rgba(255,255,255,0.72)'; dx.fillRect(0, 0, d.width, d.height);
      var bw = win.w * d.width / r.blocksX, bh = win.h * d.height / r.blocksY;
      dx.fillStyle = 'rgba(231,48,42,0.55)';
      for (var y = 0; y < r.blocksY; y++) for (var x = 0; x < r.blocksX; x++) if (r.flags[y * r.blocksX + x]) dx.fillRect(win.x * d.width + x * bw, win.y * d.height + y * bh, Math.ceil(bw), Math.ceil(bh));
      $('viewer-diff').innerHTML = ''; $('viewer-diff').appendChild(d);
      $('viewer-msgs').innerHTML = flag(r.verdict === 'different' ? 'warn' : 'ok', '', r.verdict === 'different'
        ? C.trimNum(r.changedShare * 100, 1) + '% of the compared area differs by more than the threshold. The red blocks show where. Suspected, not confirmed.'
        : 'No difference above the threshold on this page (' + C.trimNum(r.changedShare * 100, 1) + '% of blocks). This is not proof that the page is right.');
    }
  } catch (e) {
    $('viewer-msgs').innerHTML = flag('bad', '', 'The page could not be drawn. ' + esc(e && e.message ? e.message : String(e)));
  }
}

function closeViewer() { S.viewer = null; $('viewer').hidden = true; }

/* ------------------------------------------------------------------ *
 * 6 Save and reports                                                  *
 * ------------------------------------------------------------------ */

function renderSave() {
  var done = finished(), has = done.length > 0, withPdf = done.filter(function (it) { return it.hasPdf; }).length;
  $('p-save').classList.toggle('idle', !has);
  ['btn-report-open', 'btn-report-html', 'btn-report-csv', 'btn-report-json'].forEach(function (id) { $(id).disabled = !has; });
  $('btn-zip').disabled = !withPdf || S.running || typeof window.showSaveFilePicker !== 'function';
  $('btn-clear-temp').disabled = !S.out || S.running;
  $('btn-remove-work').disabled = !S.out || S.running;
  $('save-where').textContent = S.out
    ? C.plural(withPdf, 'PDF') + ' in the folder ' + S.out.name + '. The report is rewritten there after every file, as ' + (S.batch.reportName || 'SlideSize report') + ' in .html, .csv and .json.'
    : 'Finished PDFs are written into the output folder as they are made. The report files are rewritten there after every file.';
  if (S.saveError) $('save-msgs').innerHTML = flag('bad', 'The record of the batch could not be written.', esc(S.saveError), 'Check the output folder is still there and the disk is not full.');
}

function reportBlob(kind) {
  var b = S.batch;
  if (kind === 'csv') return new Blob([C.reportCsv(b)], { type: 'text/csv' });
  if (kind === 'json') return new Blob([C.reportJson(b)], { type: 'application/json' });
  return new Blob([C.reportHtml(b)], { type: 'text/html' });
}

async function saveZip() {
  var handle;
  try { handle = await window.showSaveFilePicker({ suggestedName: (S.batch.reportName || 'SlideSize PDFs') + '.zip', types: [{ description: 'Zip archive', accept: { 'application/zip': ['.zip'] } }] }); }
  catch (e) { return; }
  var w = await handle.createWritable(), zip = new ZipWriter({ write: function (u8) { return w.write(u8); }, close: function () { return w.close(); } });
  try {
    var n = 0, name = S.batch.reportName || 'SlideSize report';
    for (var it of finished()) {
      if (!it.hasPdf) continue;
      $('save-msgs').innerHTML = flag('info', '', '<span class="spinner"></span>Adding ' + esc(it.outName));
      await zip.addBlob(it.outName, await (await S.out.dir.getFileHandle(it.outName)).getFile()); n++;
    }
    await zip.add(name + '.html', C.reportHtml(S.batch)); await zip.add(name + '.csv', C.reportCsv(S.batch)); await zip.add(name + '.json', C.reportJson(S.batch));
    await zip.finish();
    $('save-msgs').innerHTML = flag('ok', '', C.plural(n, 'PDF') + ' and the three report files saved as one zip. Warnings stay in the reports inside it.');
  } catch (e) {
    try { await w.abort(); } catch (x) { /* closed */ }
    $('save-msgs').innerHTML = flag('bad', '', 'The zip could not be written. ' + esc(e && e.message ? e.message : String(e)));
  }
}

async function clearTemp() {
  if (!S.out || S.running) return;
  var n = 0;
  for (var d of ['in', 'out', 'ref', 'done', 'queue']) n += await emptyDir(S.out.sub[d]);
  viewCache.clear();
  $('save-msgs').innerHTML = flag('ok', '', n ? C.plural(n, 'temporary item') + ' removed. Copies of presentations, unfinished PDFs and reference pictures are gone. The PDFs, the reports and the batch record are untouched.'
    : 'There were no temporary files to remove.', 'Side by side views of flagged slides no longer have PowerPoint\'s picture to show.');
}

async function removeWork() {
  if (!S.out || S.running) return;
  await sendControl({ stop: true });
  for (var i = 0; i < 20 && helperFresh(); i++) { await pause(400); await pollHelper(); }
  var ok = await removeEntry(S.out.dir, C.WORK_DIR, true);
  idb.del('last');
  var name = S.out.name;
  S.out = null; S.helper = null; S.helperFiles = 'none';
  renderAll();
  $('save-msgs').innerHTML = ok ? flag('ok', '', 'The helper, the temporary files and the batch record have been removed from ' + esc(name) + '. The PDFs and reports are still there.')
    : flag('warn', '', 'The ' + C.WORK_DIR + ' folder could not be removed, most likely because the helper is still running. Close its window and delete the folder by hand.');
}

/* ------------------------------------------------------------------ *
 * Resuming                                                            *
 * ------------------------------------------------------------------ */

async function resumeFrom(saved) {
  var b = C.restoreBatch(saved), old = S.src;
  var last = await idb.get('last'), handles = {};
  if (last && last.batchId === b.id) (last.sources || []).forEach(function (x) { handles[x.id] = x.handle; });
  S.batch = b; S.src = new Map(); S.expanded = {}; S.selected = {}; S.resExpanded = {}; S.decisions = {};
  S.pendingFolderBatch = null;
  if (S.out) { S.out.ours = {}; b.items.forEach(function (it) { if (it.hasPdf && it.outName) S.out.ours[it.outName.toLowerCase()] = true; }); }
  for (var it of b.items) {
    var h = handles[it.id] || old.get(it.id);
    if (h && typeof h.queryPermission === 'function') {
      try { if (await h.queryPermission({ mode: 'read' }) !== 'granted') await h.requestPermission({ mode: 'read' }); } catch (e) { /* asked for again below */ }
    }
    if (h) S.src.set(it.id, h);
    /* what the helper did while the page was away */
    if (S.out && it.status === 'ready' && (it.interrupted || it.claimed)) {
      var r = await readJson(S.out.sub.done, it.id + '.json');
      if (r && r.id === it.id) { it.result = r; it.status = 'verifying'; }
      else if (await hasFile(S.out.sub.queue, it.id + '.json')) it.status = 'queued';
    }
  }
  try { localStorage.setItem(LS_SETTINGS, JSON.stringify({ settings: b.settings, preset: b.preset })); } catch (e) { /* private window */ }
  await pollHelper();
  b.items.forEach(function (x) { if (x.status === 'ready' || x.status === 'blocked' || x.status === 'new') recomputeIssues(x); });
  storeHandles();
  renderAll();
  renderResume();
  var live = b.items.filter(function (x) { return IN_FLIGHT[x.status]; });
  if (live.length) {
    S.runSet = {}; b.items.forEach(function (x) { if (IN_FLIGHT[x.status] || x.status === 'ready') S.runSet[x.id] = true; });
    S.running = true; S.stopping = true;            /* finish what is in flight, then wait to be told to carry on */
    live.forEach(function (x) { if (x.status === 'verifying') queueVerify(x); });
    renderAll();
  }
}

async function needsAccess() {
  var out = [];
  for (var it of S.batch.items) {
    if (it.status !== 'ready' && it.status !== 'new' && it.status !== 'blocked') continue;
    var h = S.src.get(it.id), ok = !!h;
    if (h && typeof h.queryPermission === 'function') { try { ok = await h.queryPermission({ mode: 'read' }) === 'granted'; } catch (e) { ok = false; } }
    if (!ok) out.push(it);
  }
  return out;
}

async function renderResume() {
  var el = $('resume-msgs'), h = '';
  if (!S.batch.items.length) {
    var last = await idb.get('last');
    if (last && last.dir && S.support.ok) h = flag('warn', 'A batch was started here before.', 'Its record is in the folder ' + esc(last.name || '') + '.',
      'Resume it to carry on where it stopped. Files that were finished are not converted again. The browser asks you to confirm access to the folder.',
      '<button type="button" class="btn small" id="btn-resume-last">Resume</button><button type="button" class="btn small" id="btn-forget-last">Forget it</button>');
  } else {
    var missing = await needsAccess();
    if (missing.length) h = flag('warn', 'Access to ' + C.plural(missing.length, 'original file') + ' has to be confirmed again.',
      'The browser does not let a page keep reading files after it has been closed, unless you allowed it on every visit.',
      'Select the same files or their folder again. They are matched by name, size and date. Nothing already converted is touched.',
      '<button type="button" class="btn small" id="btn-reselect-files">Select the files again</button><button type="button" class="btn small" id="btn-reselect-folder">Select their folder</button>');
  }
  el.innerHTML = h;
}

async function reselect(list) {
  var matched = 0;
  for (var x of list) {
    var file; try { file = x.file || await x.handle.getFile(); } catch (e) { continue; }
    var f = { name: file.name, size: file.size, lastModified: file.lastModified };
    S.batch.items.forEach(function (it) {
      if (C.sameFile(it, f)) {
        S.src.set(it.id, x.handle || x.file); matched++;
        if (it.inspect && !it.inspect.ok && it.inspect.error.code === 'source-missing') { it.status = 'new'; queueCheck(it); }
      }
    });
  }
  S.batch.items.forEach(function (it) { if (it.status === 'ready' || it.status === 'blocked') recomputeIssues(it); });
  storeHandles(); renderAll(); await renderResume();
  $('select-msgs').innerHTML = flag(matched ? 'ok' : 'warn', '', matched ? 'Access confirmed for ' + C.plural(matched, 'file') + '.' : 'None of those files match the batch. They are matched by name, size and date.');
}

/* ------------------------------------------------------------------ *
 * Rendering and the ticker                                            *
 * ------------------------------------------------------------------ */

var RENDER = { files: renderFiles, plan: renderPlan, start: renderStart, progress: renderProgress, results: renderResults, save: renderSave, setup: renderSetup, size: renderSize, settings: renderSettings, env: renderEnv };
var flushQueued = false;
function mark() {
  for (var i = 0; i < arguments.length; i++) S.dirty[arguments[i]] = true;
  if (!flushQueued) { flushQueued = true; setTimeout(function () { flushQueued = false; flush(); }, 0); }
}
function flush() {
  var d = S.dirty; S.dirty = {};
  Object.keys(d).forEach(function (k) { try { RENDER[k](); } catch (e) { console.error(e); } });
}
function renderAll() { Object.keys(RENDER).forEach(function (k) { S.dirty[k] = true; }); flush(); }

var ticking = false, lastCaps = '';
async function tick() {
  if (ticking) return;
  ticking = true;
  try {
    if (S.out) {
      await pollHelper();
      var sig = JSON.stringify([helperFresh(), S.helper && S.helper.state, S.helper && S.helper.caps, S.helper && S.helper.powerpoint, !!S.fonts]);
      if (sig !== lastCaps) {
        lastCaps = sig;
        S.batch.items.forEach(function (it) { if (it.status === 'ready' || it.status === 'blocked') recomputeIssues(it); });
        mark('settings', 'env', 'plan', 'files');
      }
      mark('setup', 'start');
      if (S.running) await pump();
    }
    flush();
  } catch (e) { console.error(e); }
  finally { ticking = false; }
}

function startTicker() {
  /* timers in a background tab are slowed to a crawl. A worker's are not, so the beat comes from one. */
  try {
    var w = new Worker(URL.createObjectURL(new Blob(['setInterval(function(){postMessage(0)},700)'], { type: 'text/javascript' })));
    w.onmessage = tick;
    w.onerror = function () { setInterval(tick, 700); };
  } catch (e) { setInterval(tick, 700); }
}

/* ------------------------------------------------------------------ *
 * Wiring                                                              *
 * ------------------------------------------------------------------ */

function setSetting(fn) { if (S.running) return; fn(S.batch.settings); settingsChanged(false); }

function wire() {
  $('btn-add-files').addEventListener('click', pickFiles);
  $('btn-add-folder').addEventListener('click', pickFolder);
  $('btn-clear').addEventListener('click', clearList);
  $('drop').addEventListener('click', pickFiles);
  $('file-input').addEventListener('change', function () {
    var fs = Array.prototype.slice.call(this.files);
    addSources(fs.filter(function (f) { return wanted(f.name); }).map(function (f) { return { file: f }; }), fs.filter(function (f) { return !wanted(f.name); }).length);
    this.value = '';
  });
  ['dragenter', 'dragover'].forEach(function (ev) { document.addEventListener(ev, function (e) { e.preventDefault(); $('drop').classList.add('drag'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { document.addEventListener(ev, function (e) { e.preventDefault(); if (ev === 'drop' || e.target === document.documentElement || !e.relatedTarget) $('drop').classList.remove('drag'); }); });
  document.addEventListener('drop', function (e) { if (!S.running) onDrop(e); });

  $('files-body').addEventListener('click', function (e) {
    var rm = e.target.closest('[data-remove]'), reset = e.target.closest('[data-ovreset]'), row = e.target.closest('tr.main');
    if (rm) {
      if (S.running) return;
      var id = rm.dataset.remove;
      S.batch.items = S.batch.items.filter(function (i) { return i.id !== id; }); S.src.delete(id);
      if (S.thumb && S.thumb.id === id) S.thumb = null;
      storeHandles(); renderAll();
    } else if (reset) {
      var it = S.batch.items.filter(function (i) { return i.id === reset.dataset.ovreset; })[0];
      if (it && !S.running) { it.overrides = null; recomputeIssues(it); renderAll(); }
    } else if (row && !e.target.closest('select, input, button')) { S.expanded[row.dataset.row] = !S.expanded[row.dataset.row]; renderFiles(); }
  });
  $('files-body').addEventListener('change', function (e) {
    var el = e.target.closest('[data-ov]');
    if (el) { applyOverride(el.dataset.id, el.dataset.ov, el.value); flush(); }
  });

  $('preset').addEventListener('change', function () {
    var p = S.presets.filter(function (x) { return x.id === $('preset').value; })[0];
    if (!p || S.running) return;
    var keep = S.batch.settings.size;
    S.batch.settings = C.defaultSettings(p.settings); S.batch.preset = p.id;
    /* a preset that does not set a custom size leaves the one already typed, or brought from the calculator */
    if (!p.settings.size || p.settings.size.customW == null) { S.batch.settings.size.customW = keep.customW; S.batch.settings.size.customH = keep.customH; }
    settingsChanged(true);
  });
  $('btn-preset-save').addEventListener('click', function () {
    var name = (window.prompt('Name for this preset') || '').trim();
    if (!name) return;
    var id = 'user-' + name.toLowerCase().replace(/[^a-z0-9]+/g, '-');
    S.presets = S.presets.filter(function (p) { return p.id !== id; });
    S.presets.push({ id: id, name: name, settings: JSON.parse(JSON.stringify(S.batch.settings)) });
    savePresets(); S.batch.preset = id; settingsChanged(true);
  });
  $('btn-preset-delete').addEventListener('click', function () {
    S.presets = S.presets.filter(function (p) { return p.builtin || p.id !== S.batch.preset; });
    savePresets(); S.batch.preset = ''; settingsChanged(true);
  });

  onSeg($('seg-size'), function (v) { setSetting(function (s) { s.size.preset = v; }); });
  onSeg($('seg-fit'), function (v) { setSetting(function (s) { s.size.fit = v; }); });
  onSeg($('seg-margin'), function (v) { setSetting(function (s) { s.size.margin = v; }); });
  $('orientation').addEventListener('change', function () { var v = this.value; setSetting(function (s) { s.size.orientation = v; }); });
  ['custom-w', 'custom-h'].forEach(function (id) {
    $(id).addEventListener('input', function () {
      var v = parseFloat(String(this.value).replace(',', '.')), key = id === 'custom-w' ? 'customW' : 'customH';
      this.classList.toggle('bad', !(v > 0));
      if (v > 0) setSetting(function (s) { s.size[key] = C.cmToPt(v); });
    });
  });
  $('opt-hidden').addEventListener('change', function () { var v = this.checked; setSetting(function (s) { s.hidden = v; }); });
  onSeg($('seg-range'), function (v) { setSetting(function (s) { s.range = v === 'range' ? { from: Math.max(1, +$('range-from').value || 1), to: Math.max(1, +$('range-to').value || 1) } : null; }); });
  ['range-from', 'range-to'].forEach(function (id) {
    $(id).addEventListener('input', function () { setSetting(function (s) { if (s.range) s.range = { from: +$('range-from').value, to: +$('range-to').value }; }); });
  });
  onSeg($('seg-output'), function (v) { setSetting(function (s) { s.output = v; }); });
  onSeg($('seg-pdfa'), function (v) { setSetting(function (s) { s.pdfa = v === '1'; }); });
  onSeg($('seg-quality'), function (v) { setSetting(function (s) { s.quality = v; }); });
  onSeg($('seg-fidelity'), function (v) { setSetting(function (s) { s.fidelity = v; }); });
  [['opt-bitmapText', 'bitmapText'], ['opt-tags', 'tags'], ['opt-docProps', 'docProps'], ['opt-markup', 'markup'], ['opt-placeholders', 'placeholders'], ['opt-remote', 'allowRemoteLinks'], ['opt-keeprefs', 'keepReferences']].forEach(function (p) {
    $(p[0]).addEventListener('change', function () { var v = this.checked; setSetting(function (s) { s[p[1]] = v; }); });
  });
  $('opt-links').addEventListener('change', function () { var v = this.checked; setSetting(function (s) { s.links = v ? 'keep' : 'remove'; }); });
  $('opt-metadata').addEventListener('change', function () { var v = this.checked; setSetting(function (s) { s.metadata = v ? 'strip' : 'keep'; }); });
  $('sensitivity').addEventListener('change', function () { var v = this.value; setSetting(function (s) { s.sensitivity = v; }); });
  $('timeout').addEventListener('input', function () { var v = Math.round(+this.value); if (v >= 1) setSetting(function (s) { s.timeoutSec = v * 60; }); });

  $('btn-folder').addEventListener('click', chooseFolder);
  $('btn-open-batch').addEventListener('click', async function () {
    try { var dir = await window.showDirectoryPicker({ mode: 'readwrite', id: 'slidesize-pptpdf-out' }); await useFolder(dir, true); if (!S.batch.items.length) $('helper-msgs').innerHTML += flag('info', '', 'That folder holds no batch record.'); }
    catch (e) { /* cancelled */ }
  });
  $('btn-write-helper').addEventListener('click', async function () { await writeHelperFiles(); renderSetup(); });
  $('btn-dl-helper').addEventListener('click', downloadHelper);
  $('btn-helper-retry').addEventListener('click', function () { sendControl({ retry: true }); });
  $('btn-helper-stop').addEventListener('click', function () { sendControl({ stop: true }); });
  $('helper-msgs').addEventListener('click', function (e) { if (e.target.id === 'btn-resume-folder' && S.pendingFolderBatch) resumeFrom(S.pendingFolderBatch); });
  $('plan-body').addEventListener('change', function (e) {
    var el = e.target.closest('[data-decide]');
    if (el) { if (el.value) S.decisions[el.dataset.decide] = el.value; else delete S.decisions[el.dataset.decide]; mark('plan', 'start'); flush(); }
  });

  $('btn-start').addEventListener('click', function () { startRun(null); });
  $('start-msgs').addEventListener('click', function (e) { if (e.target.id === 'btn-allow-remote') setSetting(function (s) { s.allowRemoteLinks = true; }); });
  $('btn-stop').addEventListener('click', stopAfterCurrent);
  $('btn-skip').addEventListener('click', skipCurrent);

  onSeg($('seg-filter'), function (v) { S.filter = v; setSeg($('seg-filter'), v); renderResults(); });
  $('sev-filter').addEventListener('change', function () { S.sev = this.value; renderResults(); });
  $('res-all').addEventListener('change', function () {
    var on = this.checked;
    finished().forEach(function (it) { if (S.filter === 'all' || C.classify(it) === S.filter) S.selected[it.id] = on; });
    renderResults();
  });
  $('results-body').addEventListener('click', function (e) {
    var t = e.target, a;
    if ((a = t.closest('[data-open]'))) openPdf(a.dataset.open, false);
    else if ((a = t.closest('[data-save]'))) openPdf(a.dataset.save, true);
    else if ((a = t.closest('[data-view]'))) { var p = a.dataset.view.split('|'); showViewer(p[0], +p[1] || 1); }
    else if ((a = t.closest('[data-accept]'))) { var q = a.dataset.accept.split('|'); acceptIssues(q[0], q.slice(1).join('|')).then(flush); }
    else if ((a = t.closest('[data-acceptall]'))) acceptIssues(a.dataset.acceptall, null).then(flush);
    else if ((a = t.closest('[data-pick]'))) { S.selected[a.dataset.pick] = a.checked; renderResults(); }
    else if ((a = t.closest('tr.main'))) { S.resExpanded[a.dataset.res] = !S.resExpanded[a.dataset.res]; renderResults(); }
  });
  $('btn-retry').addEventListener('click', retrySelected);

  $('viewer-close').addEventListener('click', closeViewer);
  $('viewer').addEventListener('click', function (e) { if (e.target === $('viewer')) closeViewer(); });
  document.addEventListener('keydown', function (e) {
    if (!S.viewer) return;
    if (e.key === 'Escape') closeViewer();
    else if (e.key === 'ArrowRight') showViewer(S.viewer.id, S.viewer.page + 1);
    else if (e.key === 'ArrowLeft') showViewer(S.viewer.id, S.viewer.page - 1);
  });
  $('viewer-prev').addEventListener('click', function () { showViewer(S.viewer.id, S.viewer.page - 1); });
  $('viewer-next').addEventListener('click', function () { showViewer(S.viewer.id, S.viewer.page + 1); });
  $('viewer-flagged').addEventListener('click', function () {
    var item = S.batch.items.filter(function (i) { return i.id === S.viewer.id; })[0];
    var pages = (item.issues || []).filter(function (i) { return i.code === 'visual-diff' && !i.accepted; }).map(function (i) { return i.page; }).sort(function (a, b) { return a - b; });
    var next = pages.filter(function (p) { return p > S.viewer.page; })[0] || pages[0];
    if (next) showViewer(S.viewer.id, next);
  });
  $('viewer-accept').addEventListener('click', async function () {
    var item = S.batch.items.filter(function (i) { return i.id === S.viewer.id; })[0];
    var map = item.result && item.result.export && item.result.export.pageMap;
    if (map) { await acceptSlide(item.id, map[S.viewer.page - 1]); flush(); showViewer(S.viewer.id, S.viewer.page); }
  });

  $('btn-report-open').addEventListener('click', function () { var u = URL.createObjectURL(reportBlob('html')); window.open(u, '_blank', 'noopener'); setTimeout(function () { URL.revokeObjectURL(u); }, 600000); });
  $('btn-report-html').addEventListener('click', function () { download(reportBlob('html'), (S.batch.reportName || 'SlideSize report') + '.html'); });
  $('btn-report-csv').addEventListener('click', function () { download(reportBlob('csv'), (S.batch.reportName || 'SlideSize report') + '.csv'); });
  $('btn-report-json').addEventListener('click', function () { download(reportBlob('json'), (S.batch.reportName || 'SlideSize report') + '.json'); });
  $('btn-zip').addEventListener('click', saveZip);
  $('btn-clear-temp').addEventListener('click', clearTemp);
  $('btn-remove-work').addEventListener('click', removeWork);

  $('resume-msgs').addEventListener('click', async function (e) {
    var id = e.target.id;
    if (id === 'btn-forget-last') { await idb.del('last'); renderResume(); }
    else if (id === 'btn-resume-last') {
      var last = await idb.get('last');
      try {
        if (await last.dir.requestPermission({ mode: 'readwrite' }) !== 'granted') throw new Error('Access to the folder was not allowed.');
        await useFolder(last.dir, true);
      } catch (err) { $('resume-msgs').innerHTML = flag('bad', '', 'The batch could not be resumed. ' + esc(err.message || String(err)), 'Use Open a folder with an unfinished batch in section 3 and pick the output folder.'); }
    } else if (id === 'btn-reselect-files') {
      try { var hs = await window.showOpenFilePicker({ multiple: true, id: 'slidesize-pptpdf-src' }); await reselect(hs.map(function (h) { return { handle: h }; })); } catch (err) { /* cancelled */ }
    } else if (id === 'btn-reselect-folder') {
      try { var d = await window.showDirectoryPicker({ mode: 'read', id: 'slidesize-pptpdf-src' }), found = []; await walk(d, d.name, found, 0); await reselect(found); } catch (err) { /* cancelled */ }
    }
  });

  window.addEventListener('beforeunload', function (e) { if (S.running) { e.preventDefault(); e.returnValue = ''; } });
}

function init() {
  var saved = null;
  try { saved = JSON.parse(localStorage.getItem(LS_SETTINGS) || 'null'); } catch (e) { saved = null; }
  S.batch = C.newBatch({ platform: S.platform, settings: saved && saved.settings, preset: saved ? saved.preset : 'standard' });
  loadPresets();

  /* a size carried over from the calculator fills the custom size, at PowerPoint's 19.05 cm slide height */
  var q = new URLSearchParams(location.search), qw = parseFloat(q.get('w')), qh = parseFloat(q.get('h'));
  if (qw > 0 && qh > 0) {
    S.batch.settings.size.customH = 540; S.batch.settings.size.customW = +(540 * qw / qh).toFixed(3);
    $('custom-note').textContent = 'Filled in from the calculator, ' + qw + ' x ' + qh + ' pixels. Choose Custom to use it.';
  }
  wire();
  renderAll();
  renderResume();
  startTicker();
}

init();
})();
