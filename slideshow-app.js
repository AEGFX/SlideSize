/* ============================================================
   Slideshow Builder page logic

   Holds the project, turns it into a plan with SlideshowCore,
   draws the preview with SlideshowRender and hands exports to
   SlideshowExport. Photos never leave the browser.

   AEGFX / SlideSize
   ============================================================ */
(function () {
'use strict';

var Core = window.SlideshowCore, Render = window.SlideshowRender, Exp = window.SlideshowExport;
function $(id) { return document.getElementById(id); }

var PROXY_EDGE = 1600;        /* long edge of the preview copy kept for each photo */
var THUMB_EDGE = 260;
var PV_EDGE = 1280;           /* long edge of the preview canvas */
var PV_EDGE_FULL = 1920;      /* and when the preview is full screen */
var PV_CACHE = 10;            /* decoded preview photos held at once */
var IMPORT_PARALLEL = 3;
var DISK_AUTO_BYTES = 1e9;    /* above this, "decide by size" writes straight to disk */
var MI_SRC = 'media-inspect.js';
var IMAGE_EXT = /\.(jpe?g|png|webp|avif|gif|bmp|heic|heif|tiff?|jfif)$/i;

var env = Exp.environment();

/* ------------------------------------------------------------------ *
 * Project state                                                       *
 * ------------------------------------------------------------------ */

var P = {
  output: { w: 1920, h: 1080, fps: 30 },
  lock: false, lockRatio: 16 / 9,
  loop: true,
  timing: { mode: 'perPhoto', perPhoto: 5, total: 60, transition: 1, transitionType: 'crossfade' },
  motion: { enabled: true, style: 'mixed', intensity: 'gentle', variation: true, seed: Core.newSeed() },
  framing: { exclude: true, threshold: Core.DEFAULT_THRESHOLD, mode: 'fill', bg: '#000000' },
  exp: { format: '', userPicked: false, autoNote: '', quality: 'high', bitrateMbps: null, bitrateBad: false,
         bitrateMode: 'variable', keySeconds: 2, keyBad: false, hw: 'no-preference', toDisk: 'auto' }
};

var photos = [];              /* every photo in the project, in playing order */
var byId = new Map();
var idSeq = 0, importSeq = 0;
var selId = null;
var plan = null;
var ui = {
  frame: 0, playing: false, repeat: true, exporting: false, starting: false, cancelling: false,
  edTarget: 'focus', edUrl: '', edUrlFor: null,
  orderMode: 'import', orderSeed: 0, exportCtl: null, resultUrl: '', mode: ''
};
var imp = { queue: [], active: 0, total: 0, done: 0, added: 0, failed: [], dups: [], notImages: [], running: false, report: null };
var probes = {}, probeKey = '', probing = true, probeTimer = 0;

function snapshot() {
  var list = [];
  for (var i = 0; i < photos.length; i++) {
    var p = photos[i];
    if (!p.ready) continue;
    list.push({ id: p.id, name: p.name, iw: p.iw, ih: p.ih, opaque: p.opaque, include: p.include,
                framing: p.framing, motion: p.motion, focus: p.focus, duration: p.duration, transition: p.transition });
  }
  return { output: { w: P.output.w, h: P.output.h, fps: P.output.fps }, loop: P.loop,
           timing: P.timing, motion: P.motion, framing: P.framing, photos: list };
}

function classOf(id) { for (var i = 0; i < plan.classes.length; i++) if (plan.classes[i].id === id) return plan.classes[i]; return null; }
function segOf(id) { for (var i = 0; i < plan.segs.length; i++) if (plan.segs[i].id === id) return plan.segs[i]; return null; }
function dimsOk() { return !Core.validateDims(P.output.w, P.output.h); }
function nf(n) { return Number(n).toLocaleString('en-GB'); }
function secs(frames) { return Core.trimNum(frames / plan.fps, 2); }

/* ------------------------------------------------------------------ *
 * Small DOM helpers                                                   *
 * ------------------------------------------------------------------ */

function el(tag, cls, text) {
  var e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}

/* items: [{ kind: 'bad'|'warn'|'info'|'ok', text, sub, actions: [{ label, fn, solid }] }] */
function setMsgs(box, items) {
  box.textContent = '';
  (items || []).forEach(function (it) {
    var f = el('div', 'flag ' + (it.kind || 'info'));
    var body = el('div');
    body.appendChild(el('span', null, it.text));
    if (it.sub) body.appendChild(el('span', 'sub', it.sub));
    if (it.node) body.appendChild(it.node);
    if (it.actions && it.actions.length) {
      var row = el('div', 'btn-row');
      it.actions.forEach(function (a) {
        var b = el('button', 'btn' + (a.solid ? ' solid' : ''), a.label);
        b.type = 'button';
        b.addEventListener('click', a.fn);
        row.appendChild(b);
      });
      body.appendChild(row);
    }
    f.appendChild(body);
    box.appendChild(f);
  });
}

function setSeg(box, value) {
  Array.prototype.forEach.call(box.querySelectorAll('button'), function (b) {
    var on = b.getAttribute('data-v') === String(value);
    b.classList.toggle('active', on);
    b.setAttribute('aria-checked', on ? 'true' : 'false');
  });
}
function bindSeg(box, fn) {
  box.addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (!b || b.disabled) return;
    fn(b.getAttribute('data-v'));
  });
}
function setVal(input, value) { if (document.activeElement !== input) input.value = value; }
function say(text) { var l = $('live'); l.textContent = ''; setTimeout(function () { l.textContent = text; }, 30); }

/* Whole number of pixels from a field. Simple sums are allowed, the same
   as on the calculator: 1920*2 or 4320/2. */
function parsePixels(str) {
  var c = String(str).trim();
  if (!c || !/^[\d\s+\-*/.()]+$/.test(c)) return NaN;
  try {
    var r = Function('"use strict";return(' + c + ')')();
    return (typeof r === 'number' && isFinite(r)) ? r : NaN;
  } catch (e) { return NaN; }
}
function parseNumber(str) {
  var c = String(str).trim().replace(',', '.');
  return /^\d*\.?\d+$/.test(c) || /^\d+\.$/.test(c) ? parseFloat(c) : NaN;
}

/* ------------------------------------------------------------------ *
 * Refresh: one pass from state to screen                              *
 * ------------------------------------------------------------------ */

var refreshQueued = false;
function refreshSoon() {
  if (refreshQueued) return;
  refreshQueued = true;
  requestAnimationFrame(function () { refreshQueued = false; refresh(); });
}

function refresh() {
  plan = Core.buildPlan(snapshot());
  if (ui.frame >= plan.frames) ui.frame = Math.max(0, plan.frames - 1);
  renderTray();
  renderExcluded();
  renderSelection();
  renderOutput();
  renderTiming();
  renderMotion();
  renderFit();
  renderResolve();
  renderStage();
  renderTransport();
  drawStrip();
  scheduleProbe();
  renderFormat();
  renderExport();
  requestDraw();
}

/* ------------------------------------------------------------------ *
 * Import                                                              *
 * ------------------------------------------------------------------ */

function fileKey(f) { return f.name + '|' + f.size + '|' + f.lastModified; }

function addFiles(fileList, allowDuplicates) {
  if (ui.exporting) return;
  var files = Array.prototype.slice.call(fileList || []);
  if (!files.length) return;
  /* Natural name order, so IMG_2 comes before IMG_10 whatever order the OS hands them over in. */
  files.sort(function (a, b) { return a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }); });
  if (!imp.running) { imp.total = 0; imp.done = 0; imp.added = 0; imp.failed = []; imp.dups = []; imp.notImages = []; imp.report = null; }
  var have = {};
  photos.forEach(function (p) { have[p.key] = true; });
  files.forEach(function (f) {
    var looksImage = (f.type && f.type.indexOf('image/') === 0) || IMAGE_EXT.test(f.name);
    if (!looksImage) { imp.notImages.push(f.name); return; }
    var key = fileKey(f);
    if (have[key] && !allowDuplicates) { imp.dups.push(f); return; }
    have[key] = true;
    var p = {
      id: 'p' + (++idSeq), key: key, file: f, name: f.name, importIndex: ++importSeq, ready: false,
      iw: 0, ih: 0, opaque: /jpe?g|jfif/i.test(f.type) || /\.(jpe?g|jfif)$/i.test(f.name),
      proxy: null, blur: null, thumbUrl: '',
      include: 'auto', framing: null, motion: { kind: 'auto' }, focus: null, duration: null, transition: null
    };
    photos.push(p);
    byId.set(p.id, p);
    imp.queue.push(p);
    imp.total++;
  });
  if (ui.orderMode === 'random' && imp.queue.length) ui.orderMode = 'custom';
  imp.running = imp.queue.length > 0 || imp.active > 0;
  if (!imp.running) finishImport();
  pumpImport();
  renderImport();
  refresh();
}

function pumpImport() {
  while (imp.active < IMPORT_PARALLEL && imp.queue.length) {
    (function (p) {
      imp.active++;
      ingest(p).then(function () {
        if (byId.has(p.id)) { p.ready = true; imp.added++; }
        else releasePhoto(p);
      }, function (err) {
        imp.failed.push({ name: p.name, why: whyUnreadable(p.file, err) });
        dropPhoto(p.id);
      }).then(function () {
        imp.active--; imp.done++;
        if (!imp.queue.length && !imp.active) finishImport();
        renderImport();
        refreshSoon();
        pumpImport();
      });
    })(imp.queue.shift());
  }
}

function whyUnreadable(file, err) {
  if (/\.(heic|heif)$/i.test(file.name) || /hei[cf]/i.test(file.type)) return 'HEIC is not readable in this browser. Export it as JPEG first.';
  if (/\.tiff?$/i.test(file.name)) return 'TIFF is not readable in this browser. Export it as JPEG or PNG first.';
  if (/\.svg$/i.test(file.name)) return 'SVG is not a photo format this tool reads.';
  if (!file.size) return 'The file is empty.';
  return 'Could not be decoded. The format is not supported here or the file is damaged.';
}

function canvasBlob(canvas, type, quality) {
  if (canvas.convertToBlob) return canvas.convertToBlob({ type: type, quality: quality });
  return new Promise(function (res, rej) { canvas.toBlob(function (b) { b ? res(b) : rej(new Error('toBlob failed')); }, type, quality); });
}

/* Decode once at full size to learn the true, correctly oriented
   dimensions, then keep only small copies: a proxy for the preview, a
   thumbnail for the tray and a tiny blurred image for backgrounds. The
   full size pixels are let go straight away. */
async function ingest(p) {
  var bmp = await createImageBitmap(p.file, { imageOrientation: 'from-image' });
  try {
    if (!bmp.width || !bmp.height) throw new Error('empty image');
    p.iw = bmp.width; p.ih = bmp.height;
    var s = Math.min(1, PROXY_EDGE / Math.max(p.iw, p.ih));
    var pw = Math.max(1, Math.round(p.iw * s)), ph = Math.max(1, Math.round(p.ih * s));
    var c = Render.makeCanvas(pw, ph), ctx = c.getContext('2d');
    ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(bmp, 0, 0, pw, ph);
    var ts = Math.min(1, THUMB_EDGE / Math.max(pw, ph));
    var tw = Math.max(1, Math.round(pw * ts)), th = Math.max(1, Math.round(ph * ts));
    var tc = Render.makeCanvas(tw, th), tctx = tc.getContext('2d', { willReadFrequently: true });
    tctx.imageSmoothingEnabled = true; tctx.imageSmoothingQuality = 'high';
    tctx.drawImage(c, 0, 0, tw, th);
    /* A PNG or WebP may or may not have see-through areas. Look, so that the
       preview copy keeps them exactly when the original has them. */
    if (!p.opaque) {
      var px = tctx.getImageData(0, 0, tw, th).data, solid = true;
      for (var i = 3; i < px.length; i += 4) if (px[i] < 250) { solid = false; break; }
      p.opaque = solid;
    }
    p.proxy = p.opaque ? await canvasBlob(c, 'image/jpeg', 0.9) : await canvasBlob(c, 'image/png');
    p.blur = Render.makeBlurData(c, pw, ph);
    p.thumbUrl = URL.createObjectURL(await canvasBlob(tc, p.opaque ? 'image/jpeg' : 'image/png', 0.82));
    c.width = c.height = 0; tc.width = tc.height = 0;
  } finally { bmp.close(); }
}

function finishImport() {
  imp.running = false;
  imp.report = { added: imp.added, failed: imp.failed.slice(), dups: imp.dups.slice(), notImages: imp.notImages.slice() };
}

var importClearTimer = 0;
function renderImport() {
  var box = $('import-status'), items = [];
  clearTimeout(importClearTimer);
  if (imp.running) {
    items.push({ kind: 'info', text: 'Reading photos, ' + imp.done + ' of ' + imp.total + ' done.' });
  } else if (imp.report) {
    var r = imp.report, lines = [], parts = [];
    parts.push(r.added ? 'Added ' + r.added + ' photo' + (r.added === 1 ? '' : 's') + '.' : 'No photos were added.');
    if (r.failed.length) {
      parts.push(r.failed.length + ' file' + (r.failed.length === 1 ? '' : 's') + ' could not be read and ' + (r.failed.length === 1 ? 'was' : 'were') + ' left out. The rest carried on.');
      r.failed.forEach(function (f) { lines.push(f.name + '. ' + f.why); });
    }
    if (r.notImages.length) {
      parts.push(r.notImages.length + ' file' + (r.notImages.length === 1 ? ' is' : 's are') + ' not an image and ' + (r.notImages.length === 1 ? 'was' : 'were') + ' skipped.');
      r.notImages.forEach(function (n) { lines.push(n + '. Not an image.'); });
    }
    if (r.dups.length) {
      parts.push(r.dups.length + ' duplicate' + (r.dups.length === 1 ? '' : 's') + ' skipped, already in the project.');
      r.dups.forEach(function (f) { lines.push(f.name + '. Already in the project.'); });
    }
    var clean = !lines.length, acts = [];
    if (r.dups.length) {
      var dups = r.dups;
      acts.push({ label: 'Add the duplicate' + (dups.length === 1 ? '' : 's') + ' anyway', fn: function () { addFiles(dups, true); } });
    }
    if (!clean) acts.push({ label: 'Dismiss', fn: function () { imp.report = null; renderImport(); } });
    items.push({ kind: r.failed.length ? 'bad' : (clean ? 'ok' : 'warn'), text: parts.join(' '), node: clean ? null : listNode(lines), actions: acts });
    /* a clean import needs no reading, so it clears itself */
    if (clean) importClearTimer = setTimeout(function () { if (imp.report === r) { imp.report = null; renderImport(); } }, 6000);
  }
  setMsgs(box, items);
}

function listNode(lines) {
  var max = 8, d = el('span', 'sub');
  d.textContent = lines.slice(0, max).join('\n') + (lines.length > max ? '\nand ' + (lines.length - max) + ' more' : '');
  d.style.whiteSpace = 'pre-line';
  return d;
}

function releasePhoto(p) {
  if (p.thumbUrl) { try { URL.revokeObjectURL(p.thumbUrl); } catch (e) {} p.thumbUrl = ''; }
  p.proxy = null; p.blur = null; p.file = null;
  pvEvict(p.id);
}

function dropPhoto(id) {
  var p = byId.get(id);
  if (!p) return;
  byId.delete(id);
  var i = photos.indexOf(p);
  if (i >= 0) photos.splice(i, 1);
  var q = imp.queue.indexOf(p);
  if (q >= 0) { imp.queue.splice(q, 1); imp.total--; }
  releasePhoto(p);
  if (selId === id) selId = null;
}

function clearProject() {
  stopPlayback();
  imp.queue.length = 0;
  photos.slice().forEach(function (p) { dropPhoto(p.id); });
  imp.report = null; imp.total = imp.done = imp.added = 0;
  ui.frame = 0; ui.orderMode = 'import'; importSeq = 0;
  clearResult();
  renderImport();
  refresh();
  say('Project cleared.');
}

/* ------------------------------------------------------------------ *
 * Ordering                                                            *
 * ------------------------------------------------------------------ */

function movePhoto(id, toIncludedIndex) {
  var p = byId.get(id), seg = segOf(id);
  if (!p || !seg) return;
  var n = plan.segs.length;
  toIncludedIndex = Core.clamp(toIncludedIndex, 0, n - 1);
  if (toIncludedIndex === seg.index) return;
  var anchor = byId.get(plan.segs[toIncludedIndex].id);
  var after = toIncludedIndex > seg.index;
  photos.splice(photos.indexOf(p), 1);
  photos.splice(photos.indexOf(anchor) + (after ? 1 : 0), 0, p);
  ui.orderMode = 'custom';
  refresh();
  say(p.name + ' moved to position ' + (toIncludedIndex + 1) + ' of ' + n + '.');
}

function movePhotoNextTo(id, targetId, after) {
  var p = byId.get(id), t = byId.get(targetId);
  if (!p || !t || p === t) return;
  photos.splice(photos.indexOf(p), 1);
  photos.splice(photos.indexOf(t) + (after ? 1 : 0), 0, p);
  ui.orderMode = 'custom';
  refresh();
  var seg = segOf(id);
  if (seg) say(p.name + ' moved to position ' + (seg.index + 1) + ' of ' + plan.segs.length + '.');
}

/* A new seed gives a new order. The order itself is then stored, so
   nothing reshuffles until this is pressed again. */
function randomiseOrder() {
  var seed = Core.newSeed();
  var base = photos.slice().sort(function (a, b) { return a.importIndex - b.importIndex; });
  photos = Core.seededShuffle(base, seed);
  ui.orderMode = 'random'; ui.orderSeed = seed;
  refresh();
  say('Order randomised.');
}

function restoreImportOrder() {
  photos.sort(function (a, b) { return a.importIndex - b.importIndex; });
  ui.orderMode = 'import';
  refresh();
  say('Import order restored.');
}

function inImportOrder() {
  for (var i = 1; i < photos.length; i++) if (photos[i].importIndex < photos[i - 1].importIndex) return false;
  return true;
}

/* ------------------------------------------------------------------ *
 * Tray                                                                *
 * ------------------------------------------------------------------ */

var thumbEls = new Map();

function hasOverrides(p) {
  return p.include !== 'auto' || !!p.framing || (p.motion && p.motion.kind !== 'auto') || !!p.focus || p.duration != null || !!p.transition;
}

function makeThumb(p) {
  var d = el('div', 'thumb');
  d.setAttribute('data-id', p.id);
  d.setAttribute('role', 'option');
  return d;
}

function fillThumb(d, p) {
  d.textContent = '';
  d.classList.toggle('loading', !p.ready);
  if (!p.ready) {
    d.draggable = false;
    d.style.width = '';
    d.appendChild(el('span', 'spinner'));
    d.setAttribute('aria-label', 'Reading ' + p.name);
    return;
  }
  d.draggable = true;
  d.style.width = Math.round(70 * Core.clamp(p.iw / p.ih, 0.6, 2.4)) + 'px';
  var img = el('img');
  img.src = p.thumbUrl; img.alt = ''; img.draggable = false; img.loading = 'lazy';
  d.appendChild(img);
  d.appendChild(el('span', 'num'));
  d.title = p.name;
}

function renderTray() {
  var tray = $('tray'), want = [], idx = {};
  plan.segs.forEach(function (s) { idx[s.id] = s.index; });
  photos.forEach(function (p) { if (!p.ready || idx[p.id] != null) want.push(p); });
  var keep = {};
  want.forEach(function (p) { keep[p.id] = true; });
  thumbEls.forEach(function (d, id) { if (!keep[id]) { d.remove(); thumbEls.delete(id); } });
  var hasFocusable = false;
  want.forEach(function (p, i) {
    var d = thumbEls.get(p.id);
    if (!d) { d = makeThumb(p); d._ready = null; thumbEls.set(p.id, d); }
    if (d._ready !== p.ready) { fillThumb(d, p); d._ready = p.ready; }
    if (tray.children[i] !== d) tray.insertBefore(d, tray.children[i] || null);
    if (p.ready) {
      var on = p.id === selId, n = idx[p.id] + 1;
      d.classList.toggle('sel', on);
      d.setAttribute('aria-selected', on ? 'true' : 'false');
      d.setAttribute('aria-label', 'Photo ' + n + ' of ' + plan.segs.length + ', ' + p.name);
      d.tabIndex = on ? 0 : -1;
      if (on) hasFocusable = true;
      d.querySelector('.num').textContent = n;
      var dot = d.querySelector('.dot'), need = hasOverrides(p);
      if (need && !dot) { dot = el('span', 'dot'); dot.title = 'Has its own settings'; d.appendChild(dot); }
      if (!need && dot) dot.remove();
    }
  });
  if (!hasFocusable) {
    var first = tray.querySelector('.thumb:not(.loading)');
    if (first) first.tabIndex = 0;
  }

  var n = plan.counts.included, ex = plan.counts.excluded, any = photos.length > 0;
  var chip = $('count-chip');
  chip.textContent = any ? (n + ' in video' + (ex ? ', ' + ex + ' excluded' : '')) : 'none yet';
  chip.className = 'chip' + (any ? (ex ? ' warn' : '') : ' dim');
  $('btn-random').disabled = ui.exporting || photos.length < 2;
  $('btn-restore').disabled = ui.exporting || photos.length < 2 || inImportOrder();
  $('btn-clear').disabled = ui.exporting || !any;
  $('btn-add').disabled = ui.exporting;
  var note = $('order-note');
  note.hidden = !n;
  if (n) {
    var order = inImportOrder() ? 'import order, sorted by file name'
      : (ui.orderMode === 'random' ? 'random, seed ' + ('0000000' + ui.orderSeed.toString(16).toUpperCase()).slice(-8) + '. It stays as it is until you randomise again'
      : 'your own order');
    note.textContent = 'Playing order is ' + order + '. Drag a photo to move it, or select it and use Earlier and Later. Alt with the arrow keys does the same.';
  }
}

function selectPhoto(id, focusIt) {
  selId = id && byId.has(id) ? id : null;
  ui.edTarget = 'focus';
  renderTray(); renderExcluded(); renderSelection(); drawStrip();
  if (focusIt && selId && thumbEls.get(selId)) thumbEls.get(selId).focus();
}

function bindTray() {
  var tray = $('tray'), dragId = null, target = null;
  function clearMarks() {
    Array.prototype.forEach.call(tray.querySelectorAll('.drop-before,.drop-after'), function (n) { n.classList.remove('drop-before', 'drop-after'); });
  }
  tray.addEventListener('click', function (e) {
    var t = e.target.closest('.thumb');
    if (!t || t.classList.contains('loading')) return;
    selectPhoto(t.getAttribute('data-id'), true);
  });
  tray.addEventListener('dblclick', function (e) {
    var t = e.target.closest('.thumb');
    if (t && !t.classList.contains('loading')) showInPreview(t.getAttribute('data-id'));
  });
  tray.addEventListener('keydown', function (e) {
    var t = e.target.closest('.thumb');
    if (!t) return;
    var id = t.getAttribute('data-id'), seg = segOf(id);
    if (!seg) return;
    var n = plan.segs.length, to = -1;
    if (e.key === 'ArrowRight' || e.key === 'ArrowDown') to = seg.index + 1;
    else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') to = seg.index - 1;
    else if (e.key === 'Home') to = 0;
    else if (e.key === 'End') to = n - 1;
    else if (e.key === 'Enter') { e.preventDefault(); showInPreview(id); return; }
    else if (e.key === ' ') { e.preventDefault(); e.stopPropagation(); selectPhoto(id, true); return; }
    else if ((e.key === 'Delete' || e.key === 'Backspace') && !ui.exporting) { e.preventDefault(); removeSelected(id); return; }
    else return;
    e.preventDefault();
    to = Core.clamp(to, 0, n - 1);
    if (e.altKey) {
      if (!ui.exporting) { movePhoto(id, to); selectPhoto(id, true); }
    } else {
      selectPhoto(plan.segs[to].id, true);
    }
  });
  tray.addEventListener('dragstart', function (e) {
    var t = e.target.closest('.thumb');
    if (!t || ui.exporting || t.classList.contains('loading')) { e.preventDefault(); return; }
    dragId = t.getAttribute('data-id');
    e.dataTransfer.effectAllowed = 'move';
    try { e.dataTransfer.setData('text/plain', dragId); } catch (err) {}
    t.classList.add('dragging');
  });
  tray.addEventListener('dragover', function (e) {
    if (!dragId) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    var t = e.target.closest('.thumb');
    clearMarks(); target = null;
    if (!t || t.getAttribute('data-id') === dragId || t.classList.contains('loading')) return;
    var r = t.getBoundingClientRect(), after = e.clientX > r.left + r.width / 2;
    t.classList.add(after ? 'drop-after' : 'drop-before');
    target = { id: t.getAttribute('data-id'), after: after };
  });
  tray.addEventListener('drop', function (e) {
    if (!dragId) return;
    e.preventDefault();
    var id = dragId, t = target;
    finish();
    if (t) { movePhotoNextTo(id, t.id, t.after); selectPhoto(id, false); }
  });
  tray.addEventListener('dragend', finish);
  function finish() {
    clearMarks();
    var d = dragId && thumbEls.get(dragId);
    if (d) d.classList.remove('dragging');
    dragId = null; target = null;
  }
}

function removeSelected(id) {
  var p = byId.get(id || selId);
  if (!p) return;
  var seg = segOf(p.id), nextId = null;
  if (seg && plan.segs.length > 1) nextId = plan.segs[seg.index + 1 < plan.segs.length ? seg.index + 1 : seg.index - 1].id;
  var name = p.name;
  dropPhoto(p.id);
  selId = nextId;
  refresh();
  if (nextId && thumbEls.get(nextId)) thumbEls.get(nextId).focus();
  say(name + ' removed from the project. The file on disk is untouched.');
}

/* ------------------------------------------------------------------ *
 * Excluded photos                                                     *
 * ------------------------------------------------------------------ */

function renderExcluded() {
  var box = $('excluded-box'), list = $('excluded-list');
  var out = plan.classes.filter(function (c) { return !c.included; });
  box.hidden = !out.length;
  $('excluded-title').textContent = 'Excluded photos (' + out.length + ')';
  list.textContent = '';
  out.forEach(function (c) {
    var p = byId.get(c.id);
    if (!p) return;
    var row = el('div', 'ex-item' + (p.id === selId ? ' sel' : ''));
    var img = el('img');
    img.src = p.thumbUrl; img.alt = ''; img.loading = 'lazy';
    img.addEventListener('click', function () { selectPhoto(p.id, false); });
    row.appendChild(img);
    var mid = el('div');
    mid.appendChild(el('div', 'ex-name', p.name));
    mid.appendChild(el('div', 'ex-why' + (c.auto ? '' : ' user'), c.reason));
    row.appendChild(mid);
    var btns = el('div', 'btn-row');
    function add(label, fn) {
      var b = el('button', 'btn', label);
      b.type = 'button'; b.disabled = ui.exporting;
      b.addEventListener('click', fn);
      btns.appendChild(b);
    }
    if (c.auto) {
      add('Include anyway', function () { p.include = 'in'; refresh(); say(p.name + ' included, cropped to fill.'); });
      add('Include over blur', function () { p.include = 'in'; p.framing = 'blur'; refresh(); say(p.name + ' included whole over a blurred background.'); });
    } else {
      add('Put back', function () { p.include = 'auto'; refresh(); say(p.name + ' is back to automatic.'); });
    }
    add('Remove', function () { removeSelected(p.id); });
    row.appendChild(btns);
    list.appendChild(row);
  });
}

/* Nothing left to show: say why and offer the ways out. */
function renderResolve() {
  var box = $('resolve-box'), ready = photos.filter(function (p) { return p.ready; });
  if (!ready.length || plan.counts.included > 0) { setMsgs(box, []); return; }
  var autoOut = plan.classes.filter(function (c) { return !c.included && c.auto; }).length;
  var items = [];
  if (autoOut) {
    items.push({
      kind: 'warn',
      text: 'No photos are left in the video. ' + (autoOut === ready.length ? 'All ' + autoOut : autoOut + ' of ' + ready.length) +
        ' would lose too much to cropping on a ' + P.output.w + ' x ' + P.output.h + ' screen.',
      sub: 'Pick whichever suits the photos. Nothing is deleted either way.',
      actions: [
        { label: 'Show them whole over blur', solid: true, fn: function () { P.framing.exclude = false; P.framing.mode = 'blur'; refresh(); } },
        { label: 'Crop them all to fill', fn: function () { P.framing.exclude = false; P.framing.mode = 'fill'; refresh(); } },
        { label: 'Allow more cropping', fn: function () { $('adv-fit').open = true; $('threshold').focus(); } }
      ]
    });
  } else {
    items.push({
      kind: 'warn', text: 'No photos are left in the video. Every photo has been excluded by hand.',
      actions: [{ label: 'Put them all back', solid: true, fn: function () { photos.forEach(function (p) { if (p.include === 'out') p.include = 'auto'; }); refresh(); } }]
    });
  }
  setMsgs(box, items);
}

/* ------------------------------------------------------------------ *
 * Selected photo and its advanced controls                            *
 * ------------------------------------------------------------------ */

function renderSelection() {
  var p = selId && byId.get(selId);
  var bar = $('sel-bar'), adv = $('adv-photo');
  if (!p || !p.ready) {
    selId = null;
    bar.hidden = true; adv.hidden = true;
    if (ui.edUrl) { URL.revokeObjectURL(ui.edUrl); ui.edUrl = ''; ui.edUrlFor = null; }
    return;
  }
  var seg = segOf(p.id), cls = classOf(p.id), n = plan.segs.length;
  bar.hidden = false; adv.hidden = false;
  $('sel-name').textContent = p.name;
  var meta = nf(p.iw) + ' x ' + nf(p.ih) + ' px, ' + Core.aspectLabel(p.iw, p.ih) + '. ';
  if (seg) {
    meta += 'Photo ' + (seg.index + 1) + ' of ' + n + ', on screen for ' + secs(seg.slot) + ' s' +
      (seg.trans ? ' including a ' + secs(seg.trans) + ' s crossfade out' : '') + '.';
    if (cls && cls.mode === 'fill') meta += ' At least ' + Math.round(cls.visible * 100) + '% of it stays visible.';
  } else meta += 'Not in the video. ' + (cls ? cls.reason : '');
  $('sel-meta').textContent = meta;

  var lock = ui.exporting;
  $('mv-first').disabled = $('mv-prev').disabled = lock || !seg || seg.index === 0;
  $('mv-last').disabled = $('mv-next').disabled = lock || !seg || seg.index === n - 1;
  $('sel-show').disabled = !seg;
  $('sel-toggle').disabled = lock;
  $('sel-toggle').textContent = seg ? 'Exclude' : 'Include';
  $('sel-remove').disabled = lock;

  $('ph-include').value = p.include;
  $('ph-framing').value = p.framing || '';
  $('ph-motion').value = Core.motionKindFor(p);
  setVal($('ph-duration'), p.duration == null ? '' : Core.trimNum(p.duration));
  $('ph-trans-type').value = (p.transition && p.transition.type) || '';
  setVal($('ph-trans-secs'), p.transition && p.transition.seconds != null ? Core.trimNum(p.transition.seconds) : '');
  $('ph-trans-secs').disabled = lock || ((p.transition && p.transition.type) || P.timing.transitionType) === 'cut';
  ['ph-include', 'ph-framing', 'ph-motion', 'ph-duration', 'ph-trans-type', 'ph-reset'].forEach(function (id) { $(id).disabled = lock; });

  var msgs = [];
  plan.errors.forEach(function (e) { if (e.field === 'photo' && e.id === p.id) msgs.push({ kind: 'bad', text: e.msg }); });
  if (seg && !P.loop && seg.index === n - 1 && n > 1) msgs.push({ kind: 'info', text: 'This is the last photo and looping is off, so it has no transition out. The video ends on it cleanly.' });
  if (seg && !P.motion.enabled && Core.motionKindFor(p) !== 'auto') msgs.push({ kind: 'info', text: 'Motion is switched off for the whole project, so this photo stays still.' });
  renderEditor(p, seg, cls, msgs);
  setMsgs($('ph-msgs'), msgs);
}

function placeBox(box, r) {
  box.style.left = (-r.x / r.w * 100) + '%';
  box.style.top = (-r.y / r.h * 100) + '%';
  box.style.width = (100 / r.w) + '%';
  box.style.height = (100 / r.h) + '%';
}

function renderEditor(p, seg, cls, msgs) {
  var ed = $('editor');
  if (!seg) {
    ed.hidden = true;
    msgs.push({ kind: 'info', text: 'Framing can be adjusted once the photo is in the video.' });
    return;
  }
  if (seg.mode !== 'fill') {
    ed.hidden = true;
    msgs.push({ kind: 'info', text: 'This photo is shown whole, so there is no crop to adjust. Push, pull and still apply. Pans and custom framing need Crop to fill.' });
    return;
  }
  ed.hidden = false;
  if (ui.edUrlFor !== p.id) {
    if (ui.edUrl) URL.revokeObjectURL(ui.edUrl);
    ui.edUrl = URL.createObjectURL(p.proxy);
    ui.edUrlFor = p.id;
    $('ed-img').src = ui.edUrl;
  }
  var custom = Core.motionKindFor(p) === 'custom';
  if (!custom && ui.edTarget !== 'focus') ui.edTarget = 'focus';
  var af = P.output.w / P.output.h, ai = p.iw / p.ih;
  placeBox($('ed-start'), seg.a);
  placeBox($('ed-end'), seg.b);
  $('ed-start').classList.toggle('live', ui.edTarget === 'a');
  $('ed-end').classList.toggle('live', ui.edTarget === 'b');
  $('ed-end').hidden = !seg.moving;
  $('ed-start').querySelector('span').textContent = seg.moving ? 'Start' : 'Frame';
  var focus = p.focus || Core.DEFAULT_FOCUS, fm = $('ed-focus');
  fm.hidden = custom;
  fm.style.left = (focus.x * 100) + '%';
  fm.style.top = (focus.y * 100) + '%';

  setSeg($('ed-target'), ui.edTarget);
  Array.prototype.forEach.call($('ed-target').querySelectorAll('button'), function (b) {
    var v = b.getAttribute('data-v');
    b.disabled = ui.exporting || (custom ? v === 'focus' : v !== 'focus');
  });
  var cur = ui.edTarget === 'focus' ? { cx: focus.x, cy: focus.y, z: 1 } : Core.clampCrop(ai, af, p.motion[ui.edTarget]);
  setVal($('ed-x'), Core.trimNum(cur.cx * 100, 1));
  setVal($('ed-y'), Core.trimNum(cur.cy * 100, 1));
  $('ed-x').disabled = $('ed-y').disabled = ui.exporting;
  $('ed-zoom-field').hidden = ui.edTarget === 'focus';
  if (ui.edTarget !== 'focus') {
    setVal($('ed-zoom'), Math.round(cur.z * 100));
    $('ed-zoom-val').textContent = Math.round(cur.z * 100) + '%';
    $('ed-zoom').disabled = ui.exporting;
  }
  $('ed-hint').textContent = custom
    ? 'The dashed frame is where the move starts and the green frame is where it ends, crossfades included. Pick one above, then click or drag on the photo to place it. Frames cannot leave the photo, so no empty edge can show.'
    : 'The amber ring is the focus point. Crops are placed around it and pushes hold it steady. Click the photo to move it onto the subject. The frames show the move this photo will make.';
}

function ensureCustom(p) {
  var seg = segOf(p.id), af = P.output.w / P.output.h, ai = p.iw / p.ih, a, b;
  if (seg && seg.mode === 'fill') {
    a = Core.rectToCrop(seg.a, ai, af);
    b = Core.rectToCrop(seg.b, ai, af);
  } else {
    var f = p.focus || Core.DEFAULT_FOCUS;
    a = Core.clampCrop(ai, af, { cx: f.x, cy: f.y, z: 1 });
    b = Core.clampCrop(ai, af, { cx: f.x, cy: f.y, z: 1.08 });
  }
  p.motion = { kind: 'custom', a: a, b: b };
}

function editorApply(u, v, z) {
  var p = selId && byId.get(selId);
  if (!p || ui.exporting) return;
  if (ui.edTarget === 'focus') {
    var f = p.focus || Core.DEFAULT_FOCUS;
    p.focus = { x: Core.clamp(u == null ? f.x : u, 0, 1), y: Core.clamp(v == null ? f.y : v, 0, 1) };
  } else if (p.motion && p.motion.kind === 'custom') {
    var af = P.output.w / P.output.h, ai = p.iw / p.ih, c = p.motion[ui.edTarget];
    p.motion[ui.edTarget] = Core.clampCrop(ai, af, { cx: u == null ? c.cx : u, cy: v == null ? c.cy : v, z: z == null ? c.z : z });
  }
  refreshSoon();
}

function bindSelection() {
  $('mv-first').addEventListener('click', function () { if (selId) movePhoto(selId, 0); });
  $('mv-prev').addEventListener('click', function () { var s = selId && segOf(selId); if (s) movePhoto(selId, s.index - 1); });
  $('mv-next').addEventListener('click', function () { var s = selId && segOf(selId); if (s) movePhoto(selId, s.index + 1); });
  $('mv-last').addEventListener('click', function () { if (selId) movePhoto(selId, plan.segs.length - 1); });
  $('sel-show').addEventListener('click', function () { if (selId) showInPreview(selId); });
  $('sel-toggle').addEventListener('click', function () {
    var p = selId && byId.get(selId);
    if (!p) return;
    var inNow = !!segOf(p.id), cls = classOf(p.id);
    if (inNow) p.include = 'out';
    else p.include = (cls && cls.unsuitable && P.framing.exclude) ? 'in' : 'auto';
    refresh();
    say(p.name + (inNow ? ' excluded. It stays in the project.' : ' included.'));
  });
  $('sel-remove').addEventListener('click', function () { removeSelected(); });

  $('ph-include').addEventListener('change', function () { var p = byId.get(selId); if (p) { p.include = this.value; refresh(); } });
  $('ph-framing').addEventListener('change', function () { var p = byId.get(selId); if (p) { p.framing = this.value || null; refresh(); } });
  $('ph-motion').addEventListener('change', function () {
    var p = byId.get(selId);
    if (!p) return;
    if (this.value === 'custom') { ensureCustom(p); ui.edTarget = 'a'; }
    else { p.motion = { kind: this.value }; ui.edTarget = 'focus'; }
    refresh();
  });
  $('ph-duration').addEventListener('input', function () {
    var p = byId.get(selId);
    if (!p) return;
    var raw = this.value.trim(), v = Core.parseDuration(raw);
    this.classList.toggle('bad', raw !== '' && !(v > 0));
    if (raw === '') p.duration = null; else if (v > 0) p.duration = v; else return;
    refreshSoon();
  });
  function transChanged() {
    var p = byId.get(selId);
    if (!p) return;
    var type = $('ph-trans-type').value, raw = $('ph-trans-secs').value.trim(), v = parseNumber(raw);
    $('ph-trans-secs').classList.toggle('bad', raw !== '' && !(v >= 0));
    var s = raw === '' || !(v >= 0) ? null : v;
    p.transition = (type || s != null) ? { type: type || null, seconds: s } : null;
    refreshSoon();
  }
  $('ph-trans-type').addEventListener('change', transChanged);
  $('ph-trans-secs').addEventListener('input', transChanged);
  $('ph-reset').addEventListener('click', function () {
    var p = byId.get(selId);
    if (!p) return;
    p.include = 'auto'; p.framing = null; p.motion = { kind: 'auto' }; p.focus = null; p.duration = null; p.transition = null;
    ui.edTarget = 'focus';
    $('ph-duration').value = ''; $('ph-trans-secs').value = '';
    $('ph-duration').classList.remove('bad'); $('ph-trans-secs').classList.remove('bad');
    refresh();
    say(p.name + ' reset to automatic.');
  });

  bindSeg($('ed-target'), function (v) { ui.edTarget = v; renderSelection(); });
  var view = $('ed-view'), dragging = false;
  function point(e) {
    var r = view.getBoundingClientRect();
    return { u: Core.clamp((e.clientX - r.left) / r.width, 0, 1), v: Core.clamp((e.clientY - r.top) / r.height, 0, 1) };
  }
  view.addEventListener('pointerdown', function (e) {
    if (ui.exporting) return;
    dragging = true;
    try { view.setPointerCapture(e.pointerId); } catch (err) {}
    var pt = point(e); editorApply(pt.u, pt.v, null);
  });
  view.addEventListener('pointermove', function (e) { if (dragging) { var pt = point(e); editorApply(pt.u, pt.v, null); } });
  function end() { dragging = false; }
  view.addEventListener('pointerup', end);
  view.addEventListener('pointercancel', end);
  view.addEventListener('keydown', function (e) {
    var step = e.shiftKey ? 0.05 : 0.01, dx = 0, dy = 0;
    if (e.key === 'ArrowLeft') dx = -step; else if (e.key === 'ArrowRight') dx = step;
    else if (e.key === 'ArrowUp') dy = -step; else if (e.key === 'ArrowDown') dy = step;
    else return;
    e.preventDefault();
    var p = byId.get(selId);
    if (!p) return;
    var cur = ui.edTarget === 'focus' ? (p.focus || Core.DEFAULT_FOCUS) : null;
    if (cur) editorApply(cur.x + dx, cur.y + dy, null);
    else { var c = p.motion[ui.edTarget]; editorApply(c.cx + dx, c.cy + dy, null); }
  });
  $('ed-x').addEventListener('input', function () { var v = parseNumber(this.value); this.classList.toggle('bad', !(v >= 0 && v <= 100)); if (v >= 0 && v <= 100) editorApply(v / 100, null, null); });
  $('ed-y').addEventListener('input', function () { var v = parseNumber(this.value); this.classList.toggle('bad', !(v >= 0 && v <= 100)); if (v >= 0 && v <= 100) editorApply(null, v / 100, null); });
  $('ed-zoom').addEventListener('input', function () { editorApply(null, null, parseFloat(this.value) / 100); });
}

/* ------------------------------------------------------------------ *
 * Preview                                                             *
 * ------------------------------------------------------------------ */

var cv = $('preview');
var pv = { ctx: cv.getContext('2d', { alpha: false }), scratchCanvas: null, scratchCtx: null,
           cache: new Map(), pending: {}, tick: 0, baseTime: 0, baseFrame: 0, raf: 0, drawQueued: false };

function previewSize() {
  var W = plan.W, H = plan.H;
  var edge = document.fullscreenElement ? PV_EDGE_FULL : PV_EDGE;
  var s = Math.min(1, edge / Math.max(W, H));
  return { w: Math.max(2, Math.round(W * s)), h: Math.max(2, Math.round(H * s)), reduced: s < 1 };
}

function pvScratch() {
  if (!pv.scratchCtx || pv.scratchCanvas.width !== cv.width || pv.scratchCanvas.height !== cv.height) {
    pv.scratchCanvas = Render.makeCanvas(cv.width, cv.height);
    pv.scratchCtx = pv.scratchCanvas.getContext('2d', { alpha: false });
  }
  return pv.scratchCtx;
}

function pvAsset(seg) {
  var a = pv.cache.get(seg.id);
  if (a) { a.at = ++pv.tick; return a; }
  pvLoad(seg.id);
  return null;
}

function pvLoad(id) {
  if (pv.cache.has(id) || pv.pending[id]) return;
  var p = byId.get(id);
  if (!p || !p.proxy) return;
  pv.pending[id] = true;
  createImageBitmap(p.proxy).then(function (bmp) {
    delete pv.pending[id];
    var still = byId.get(id);
    if (!still || !still.proxy) { bmp.close(); return; }
    pv.cache.set(id, { bmp: bmp, blur: Render.blurSource(still.blur), at: ++pv.tick });
    pvTrim();
    requestDraw();
  }, function () { delete pv.pending[id]; });
}

function pvEvict(id) {
  var a = pv.cache.get(id);
  if (!a) return;
  try { a.bmp.close(); } catch (e) {}
  if (a.blur) a.blur.width = a.blur.height = 0;
  pv.cache.delete(id);
}

function pvTrim() {
  while (pv.cache.size > PV_CACHE) {
    var oldest = null, at = Infinity;
    pv.cache.forEach(function (a, id) { if (a.at < at) { at = a.at; oldest = id; } });
    if (oldest == null) break;
    pvEvict(oldest);
  }
}

function preloadAround(frame) {
  var n = plan.segs.length;
  if (!n) return;
  var i = Core.locate(plan, frame);
  for (var d = 0; d <= 2 && d < n; d++) {
    var j = i + d;
    if (j >= n) { if (!(plan.loop || ui.repeat)) break; j %= n; }
    pvLoad(plan.segs[j].id);
  }
}

function frameReady(frame) {
  var layers = Core.stateAt(plan, frame);
  for (var i = 0; i < layers.length; i++) if (!pv.cache.has(layers[i].seg.id)) return false;
  return true;
}

function requestDraw() {
  if (pv.drawQueued) return;
  pv.drawQueued = true;
  requestAnimationFrame(function () { pv.drawQueued = false; draw(); });
}

function draw() {
  if (!plan || !plan.segs.length || !plan.frames || cv.hidden) { $('stage-busy').hidden = true; return; }
  var size = previewSize();
  if (cv.width !== size.w || cv.height !== size.h) { cv.width = size.w; cv.height = size.h; fitCanvas(); }
  var missing = Render.renderFrame(pv.ctx, cv.width, cv.height, plan, ui.frame, pvAsset, pvScratch);
  $('stage-busy').hidden = !missing.length;
  preloadAround(ui.frame);
}

/* Size the canvas element to the largest box of the right shape that fits the stage. */
function fitCanvas() {
  if (!plan || !plan.W || cv.hidden) return;
  var st = $('stage'), bw = st.clientWidth, bh = st.clientHeight, ar = plan.W / plan.H;
  if (!bw || !bh) return;
  var w = bw, h = bw / ar;
  if (h > bh) { h = bh; w = bh * ar; }
  cv.style.width = Math.floor(w) + 'px';
  cv.style.height = Math.floor(h) + 'px';
}

function renderStage() {
  var any = photos.length > 0, ready = photos.some(function (p) { return p.ready; });
  var playable = plan.segs.length > 0 && plan.frames > 0 && dimsOk();
  $('drop').hidden = any;
  cv.hidden = !playable;
  if (dimsOk()) $('stage').style.aspectRatio = P.output.w + ' / ' + P.output.h;
  var msg = $('stage-msg');
  msg.hidden = !any || playable;
  if (!msg.hidden) {
    msg.textContent = '';
    if (!ready) { msg.appendChild(el('b', null, 'Reading photos')); msg.appendChild(el('span', null, 'They appear here as soon as the first one is decoded.')); }
    else if (!dimsOk()) { msg.appendChild(el('b', null, 'The output size is not valid')); msg.appendChild(el('span', null, 'Correct the width and height to see the preview.')); }
    else { msg.appendChild(el('b', null, 'No photos are left in the video')); msg.appendChild(el('span', null, 'See the options under Photos below.')); }
  }
  $('transport').hidden = !playable;
  var note = $('preview-note');
  note.hidden = !playable;
  if (playable) {
    var size = previewSize();
    note.textContent = '';
    if (size.reduced) {
      note.appendChild(document.createTextNode('Preview drawn at ' + size.w + ' x ' + size.h + ' from reduced copies of the photos, for speed. '));
      note.appendChild(el('b', null, 'The export renders every frame at ' + plan.W + ' x ' + plan.H + ' from the originals.'));
    } else {
      note.appendChild(document.createTextNode('Preview drawn at the full ' + size.w + ' x ' + size.h + ' from reduced copies of the photos. The export uses the originals.'));
    }
    note.appendChild(document.createTextNode(' Same framing, motion and timing in both.'));
    fitCanvas();
  }
  if (!playable) stopPlayback();
  var logo = $('logo-mark');
  if (dimsOk()) logo.style.width = Core.clamp(Math.round(24 * P.output.w / P.output.h), 14, 72) + 'px';
}

function setPlayIcon() {
  $('ico-play').innerHTML = ui.playing ? '<path d="M7 4v16M17 4v16"/>' : '<path d="M7 4l13 8-13 8V4z"/>';
  $('btn-play').setAttribute('aria-label', ui.playing ? 'Pause preview' : 'Play preview');
}

function renderTransport() {
  var N = plan.frames, sc = $('scrub');
  sc.max = Math.max(0, N - 1);
  sc.value = ui.frame;
  var tc = $('timecode');
  tc.textContent = '';
  tc.appendChild(document.createTextNode(Core.formatTime(ui.frame, plan.fps) + ' / ' + Core.formatTime(N, plan.fps)));
  tc.appendChild(el('small', null, 'frame ' + nf(ui.frame + 1) + ' of ' + nf(N)));
  $('btn-seam').disabled = !P.loop || plan.segs.length < 1;
  $('btn-seam').textContent = P.loop ? 'Check loop point' : 'Loop is off';
  setPlayIcon();
}

function updatePlayhead() {
  $('scrub').value = ui.frame;
  var tc = $('timecode');
  tc.firstChild.nodeValue = Core.formatTime(ui.frame, plan.fps) + ' / ' + Core.formatTime(plan.frames, plan.fps);
  tc.lastChild.textContent = 'frame ' + nf(ui.frame + 1) + ' of ' + nf(plan.frames);
}

function drawStrip() {
  var c = $('strip'), N = plan.frames;
  var w = Math.max(10, Math.round(c.clientWidth * (window.devicePixelRatio || 1)));
  if (c.width !== w) c.width = w;
  var g = c.getContext('2d');
  g.clearRect(0, 0, c.width, c.height);
  if (!N) return;
  plan.segs.forEach(function (s, i) {
    var x0 = s.start / N * w, x1 = (s.start + s.slot) / N * w;
    g.fillStyle = s.id === selId ? '#A6CE4E' : (i % 2 ? '#3a3a3a' : '#4a4a4a');
    g.fillRect(x0, 0, Math.max(1, x1 - x0 - (plan.segs.length < 200 ? 1 : 0)), c.height);
    if (s.trans > 0 && s.id !== selId) {
      var xt = (s.start + s.slot - s.trans) / N * w;
      g.fillStyle = 'rgba(166,206,78,0.35)';
      g.fillRect(xt, 0, x1 - xt, c.height);
    }
  });
}

function startPlayback() {
  if (!plan.segs.length || ui.exporting) return;
  if (!ui.repeat && ui.frame >= plan.frames - 1) ui.frame = 0;
  ui.playing = true;
  pv.baseTime = performance.now(); pv.baseFrame = ui.frame;
  setPlayIcon();
  cancelAnimationFrame(pv.raf);
  pv.raf = requestAnimationFrame(tick);
}
function stopPlayback() {
  ui.playing = false;
  cancelAnimationFrame(pv.raf);
  setPlayIcon();
}
function seek(frame) {
  ui.frame = Core.clamp(Math.round(frame), 0, Math.max(0, plan.frames - 1));
  pv.baseTime = performance.now(); pv.baseFrame = ui.frame;
  updatePlayhead();
  requestDraw();
}

/* The preview steps through the same frame numbers the export renders.
   The clock only decides which frame is due, never what is in it. */
function tick(now) {
  if (!ui.playing) return;
  var N = plan.frames;
  if (!N) { stopPlayback(); return; }
  var k = pv.baseFrame + Math.floor((now - pv.baseTime) * plan.fps / 1000);
  if (k >= N) {
    if (ui.repeat) { k %= N; pv.baseFrame = k; pv.baseTime = now; }
    else { ui.frame = N - 1; updatePlayhead(); draw(); stopPlayback(); return; }
  }
  if (k !== ui.frame) {
    if (frameReady(k)) { ui.frame = k; updatePlayhead(); draw(); }
    else {
      /* wait for the photo instead of skipping past it */
      Core.stateAt(plan, k).forEach(function (l) { pvLoad(l.seg.id); });
      $('stage-busy').hidden = false;
      pv.baseFrame = ui.frame; pv.baseTime = now;
    }
  }
  pv.raf = requestAnimationFrame(tick);
}

function showInPreview(id) {
  var f = Core.frameOfPhoto(plan, id);
  if (f < 0) return;
  seek(f);
}

function bindPreview() {
  $('btn-play').addEventListener('click', function () { ui.playing ? stopPlayback() : startPlayback(); });
  $('btn-restart').addEventListener('click', function () { seek(0); });
  $('scrub').addEventListener('input', function () { seek(parseInt(this.value, 10) || 0); });
  $('loop-preview').addEventListener('change', function () { ui.repeat = this.checked; });
  $('btn-seam').addEventListener('click', function () {
    if (!P.loop || !plan.frames) return;
    var back = Math.min(Math.round(3 * plan.fps), Math.floor(plan.frames / 2));
    ui.repeat = true; $('loop-preview').checked = true;
    seek(plan.frames - back);
    startPlayback();
  });
  $('btn-full').addEventListener('click', function () {
    var panel = $('preview-panel');
    if (document.fullscreenElement) document.exitFullscreen();
    else if (panel.requestFullscreen) panel.requestFullscreen().catch(function () {});
  });
  document.addEventListener('fullscreenchange', function () { renderStage(); requestDraw(); });
  if (window.ResizeObserver) new ResizeObserver(function () { fitCanvas(); drawStrip(); }).observe($('stage'));
  window.addEventListener('resize', function () { fitCanvas(); drawStrip(); });
  document.addEventListener('keydown', function (e) {
    if (e.key !== ' ' || e.defaultPrevented) return;
    var t = e.target, tag = t && t.tagName;
    if (tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA' || tag === 'BUTTON' || tag === 'SUMMARY' || (t && t.isContentEditable)) return;
    if (!plan.segs.length) return;
    e.preventDefault();
    ui.playing ? stopPlayback() : startPlayback();
  });
}

/* ------------------------------------------------------------------ *
 * Settings panels                                                     *
 * ------------------------------------------------------------------ */

function renderOutput() {
  var o = P.output, ok = dimsOk();
  var pre = ok ? Core.findPreset(o.w, o.h) : null;
  $('preset').value = pre ? pre.id : 'custom';
  setVal($('out-w'), isFinite(o.w) ? String(o.w) : $('out-w').value);
  setVal($('out-h'), isFinite(o.h) ? String(o.h) : $('out-h').value);
  var err = Core.validateDims(o.w, o.h);
  $('out-w').classList.toggle('bad', !!err);
  $('out-h').classList.toggle('bad', !!err);
  setMsgs($('size-msgs'), err ? [{ kind: 'bad', text: err }] : []);
  var chip = $('aspect-chip');
  chip.textContent = ok ? Core.aspectLabel(o.w, o.h) + ' ' + Core.shapeOf(o.w / o.h) : 'size not valid';
  chip.className = 'chip' + (ok ? '' : ' warn');
  setSeg($('fps-seg'), o.fps);
  var lock = $('lock');
  lock.setAttribute('aria-pressed', P.lock ? 'true' : 'false');
  lock.classList.toggle('solid', P.lock);
  lock.title = P.lock ? 'Aspect ratio locked. Click to unlock.' : 'Lock the aspect ratio';
  $('ico-lock').innerHTML = P.lock
    ? '<rect x="5" y="11" width="14" height="9" rx="1"/><path d="M8 11V7a4 4 0 018 0v4"/>'
    : '<rect x="5" y="11" width="14" height="9" rx="1"/><path d="M8 11V7a4 4 0 017-2.6"/>';
}

function bindOutput() {
  var sel = $('preset');
  Core.PRESETS.forEach(function (p) {
    var o = el('option', null, p.name + '  ' + p.w + ' x ' + p.h);
    o.value = p.id;
    sel.appendChild(o);
  });
  var c = el('option', null, 'Custom size');
  c.value = 'custom';
  sel.appendChild(c);
  sel.addEventListener('change', function () {
    var id = this.value;
    if (id === 'custom') { $('out-w').focus(); $('out-w').select(); return; }
    for (var i = 0; i < Core.PRESETS.length; i++) if (Core.PRESETS[i].id === id) {
      P.output.w = Core.PRESETS[i].w; P.output.h = Core.PRESETS[i].h;
      P.lockRatio = P.output.w / P.output.h;
    }
    $('out-w').value = P.output.w; $('out-h').value = P.output.h;
    refresh();
  });
  function onSize(which) {
    return function () {
      var v = parsePixels(this.value);
      P.output[which] = v;
      /* With the lock on, the other side follows and shows the new value at once. */
      if (P.lock && isFinite(v) && v > 0 && P.lockRatio > 0) {
        if (which === 'w') { P.output.h = Math.max(1, Math.round(v / P.lockRatio)); $('out-h').value = P.output.h; }
        else { P.output.w = Math.max(1, Math.round(v * P.lockRatio)); $('out-w').value = P.output.w; }
      }
      refreshSoon();
    };
  }
  $('out-w').addEventListener('input', onSize('w'));
  $('out-h').addEventListener('input', onSize('h'));
  ['out-w', 'out-h'].forEach(function (id) {
    $(id).addEventListener('change', function () { var k = id === 'out-w' ? 'w' : 'h'; if (isFinite(P.output[k])) this.value = P.output[k]; });
  });
  $('lock').addEventListener('click', function () {
    P.lock = !P.lock;
    if (P.lock && dimsOk()) P.lockRatio = P.output.w / P.output.h;
    renderOutput();
  });
  var fs = $('fps-seg');
  Core.FPS_CHOICES.forEach(function (f) {
    var b = el('button', null, String(f));
    b.type = 'button'; b.setAttribute('data-v', f); b.setAttribute('role', 'radio');
    fs.appendChild(b);
  });
  bindSeg(fs, function (v) { P.output.fps = parseInt(v, 10); refresh(); });
}

function renderTiming() {
  var t = P.timing, n = plan.segs.length, fps = plan.fps;
  setSeg($('timing-mode'), t.mode);
  $('per-photo-field').hidden = t.mode !== 'perPhoto';
  $('total-field').hidden = t.mode !== 'total';
  $('trans-type').value = t.transitionType;
  $('trans-field').hidden = t.transitionType === 'cut';
  $('loop').checked = P.loop;
  $('dur-chip').textContent = n ? Core.formatTime(plan.frames, fps) : '0:00.00';

  var bad = {};
  plan.errors.forEach(function (e) { bad[e.field] = true; });
  $('per-photo').classList.toggle('bad', !!bad.perPhoto);
  $('total').classList.toggle('bad', !!bad.total);
  $('trans').classList.toggle('bad', !!bad.transition);

  var ex = '';
  if (n === 1) {
    ex = 'One photo, on screen for ' + Core.formatTime(plan.frames, fps) + '.' +
      (P.loop ? ' Its motion goes out and comes back within that time, so the loop has no reset.' : '');
  } else if (n) {
    var tm = plan.timing, slot = tm.perPhotoFrames, x = tm.transitionFrames;
    if (t.mode === 'perPhoto') {
      ex = x > 0
        ? 'A photo’s time includes its crossfade into the next one, so nothing is added on top. That is ' + secs(slot - x) + ' s on its own, then a ' + secs(x) + ' s crossfade. ' + n + ' photos make ' + Core.formatTime(plan.frames, fps) + '.'
        : 'Photos cut straight from one to the next. ' + n + ' photos make ' + Core.formatTime(plan.frames, fps) + '.';
    } else {
      ex = 'The total is shared equally between ' + n + ' photos, ' + (tm.uniform ? '' : 'about ') + secs(slot) + ' s each' +
        (x > 0 ? ' including a ' + secs(x) + ' s crossfade into the next one.' : ', cut to cut.');
    }
    if (!tm.uniform && t.mode === 'perPhoto') ex += ' Some photos have their own times.';
    ex += P.loop ? ' The last photo’s transition leads back into the first and is counted in the length.'
                 : ' The video starts on the first photo and ends on the last with no fade.';
  } else ex = 'A photo’s time includes its crossfade into the next one, so the video length is simply the photo times added up.';
  $('timing-explain').textContent = ex;

  var items = [];
  plan.errors.forEach(function (e) { if (e.field === 'perPhoto' || e.field === 'total' || e.field === 'transition' || e.field === 'photo') items.push({ kind: 'bad', text: e.msg }); });
  plan.warnings.forEach(function (w) { items.push({ kind: 'warn', text: w }); });
  plan.notes.forEach(function (w) { items.push({ kind: 'info', text: w }); });
  setMsgs($('timing-msgs'), items);
}

function bindTiming() {
  bindSeg($('timing-mode'), function (v) {
    if (v === P.timing.mode) return;
    /* carry the current length across so switching mode does not change the video */
    if (plan.segs.length && plan.ok) {
      if (v === 'total') { P.timing.total = plan.frames / plan.fps; $('total').value = formatSecondsField(P.timing.total); }
      else if (plan.timing.uniform) { P.timing.perPhoto = plan.timing.perPhotoFrames / plan.fps; $('per-photo').value = Core.trimNum(P.timing.perPhoto); }
    }
    P.timing.mode = v;
    refresh();
  });
  $('per-photo').addEventListener('input', function () { P.timing.perPhoto = Core.parseDuration(this.value); refreshSoon(); });
  $('total').addEventListener('input', function () { P.timing.total = Core.parseDuration(this.value); refreshSoon(); });
  $('trans').addEventListener('input', function () { P.timing.transition = parseNumber(this.value); refreshSoon(); });
  $('trans-type').addEventListener('change', function () { P.timing.transitionType = this.value; refresh(); });
  $('loop').addEventListener('change', function () { P.loop = this.checked; refresh(); });
}

function formatSecondsField(s) {
  if (s < 60) return Core.trimNum(s);
  var m = Math.floor(s / 60), r = s - m * 60;
  return m + ':' + (r < 10 ? '0' : '') + Core.trimNum(r);
}

function renderMotion() {
  var m = P.motion;
  $('motion-on').checked = m.enabled;
  $('motion-style').value = m.style;
  setSeg($('intensity'), m.intensity);
  $('motion-vary').checked = m.variation;
  var off = !m.enabled;
  $('motion-style').disabled = off;
  $('motion-vary').disabled = off;
  $('motion-reseed').disabled = off || !m.variation;
  Array.prototype.forEach.call($('intensity').querySelectorAll('button'), function (b) { b.disabled = off; });
}

function bindMotion() {
  $('motion-on').addEventListener('change', function () { P.motion.enabled = this.checked; refresh(); });
  $('motion-style').addEventListener('change', function () { P.motion.style = this.value; refresh(); });
  bindSeg($('intensity'), function (v) { P.motion.intensity = v; refresh(); });
  $('motion-vary').addEventListener('change', function () { P.motion.variation = this.checked; refresh(); });
  $('motion-reseed').addEventListener('click', function () { P.motion.seed = Core.newSeed(); refresh(); say('New set of moves chosen.'); });
}

function renderFit() {
  var f = P.framing, n = plan.counts.included, ex = plan.counts.excluded, ready = n + ex;
  $('exclude').checked = f.exclude;
  $('fit-mode-field').hidden = f.exclude;
  setSeg($('fit-mode'), f.mode);
  var anyFit = (!f.exclude && f.mode === 'fit') || photos.some(function (p) { return p.framing === 'fit'; });
  $('bg-field').hidden = !anyFit;
  $('bg').value = f.bg;
  $('bg-val').textContent = f.bg.toUpperCase();
  setVal($('threshold'), Math.round(f.threshold * 100));
  $('threshold-val').textContent = Math.round(f.threshold * 100) + '%';
  $('threshold').disabled = !f.exclude;
  var autoOut = plan.classes.filter(function (c) { return !c.included && c.auto; }).length;
  var chip = $('fit-chip');
  chip.textContent = !ready ? '' : (autoOut ? autoOut + ' excluded' : 'all included');
  chip.className = 'chip' + (autoOut ? ' warn' : ' dim');
  chip.hidden = !ready;
  var os = dimsOk() ? Core.shapeOf(P.output.w / P.output.h) : 'landscape';
  $('exclude-sub').textContent = os === 'landscape' ? 'Portrait photos are normally left out of a landscape video. A slightly different shape is fine.'
    : (os === 'portrait' ? 'Landscape photos are normally left out of a portrait video. A slightly different shape is fine.'
    : 'Photos much wider or taller than the square are left out. A slightly different shape is fine.');
  var t;
  if (f.exclude) {
    t = 'Photos that stay are cropped to fill the screen. ' + (ready ? (autoOut ? autoOut + ' of ' + ready + ' are excluded right now and listed under Photos, where you can put any of them back.' : 'None are excluded right now.') : '');
  } else if (f.mode === 'fill') t = 'Every photo is cropped to fill the screen, however much that removes. Nothing is stretched.';
  else if (f.mode === 'fit') t = 'Every photo is shown whole on the background colour. Nothing is cropped or stretched.';
  else t = 'Every photo is shown whole over a blurred, darkened copy of itself. Nothing is cropped or stretched.';
  $('fit-explain').textContent = t;
}

function bindFit() {
  $('exclude').addEventListener('change', function () { P.framing.exclude = this.checked; refresh(); });
  bindSeg($('fit-mode'), function (v) { P.framing.mode = v; refresh(); });
  $('bg').addEventListener('input', function () { P.framing.bg = this.value; refreshSoon(); });
  $('threshold').addEventListener('input', function () { P.framing.threshold = parseInt(this.value, 10) / 100; refreshSoon(); });
}

/* ------------------------------------------------------------------ *
 * Format and capability detection                                     *
 * ------------------------------------------------------------------ */

var FORMATS = [];
Core.CODEC_ORDER.forEach(function (fam) {
  Core.CODECS[fam].containers.forEach(function (con) {
    FORMATS.push({ value: fam + '/' + con, family: fam, container: con, label: Core.CODECS[fam].label + ' in ' + Core.CONTAINERS[con].label });
  });
});
var PLAYBACK_NOTE = {
  avc: 'Plays in PowerPoint, Keynote, QLab, VLC and most show playback software.',
  vp9: 'Plays in browsers and VLC. PowerPoint on Windows and most show playback software will not play it.',
  av1: 'Needs a recent player. PowerPoint and most show playback software will not play it.',
  hevc: 'Playback is patchy on Windows show machines. Test it on the machine that will run the show.'
};

function currentFormat() {
  for (var i = 0; i < FORMATS.length; i++) if (FORMATS[i].value === P.exp.format) return FORMATS[i];
  return null;
}

function bitrateFor(family) {
  if (P.exp.bitrateMbps) return Math.round(P.exp.bitrateMbps * 1e6);
  return Core.suggestBitrate(family, P.output.w, P.output.h, P.output.fps, P.exp.quality);
}

function probeSettings(family) {
  return { W: P.output.w, H: P.output.h, fps: P.output.fps, bitrate: bitrateFor(family), bitrateMode: P.exp.bitrateMode, hw: P.exp.hw };
}

/* Re-ask the browser whenever anything that affects the answer changes. */
function scheduleProbe() {
  var key = [P.output.w, P.output.h, P.output.fps, P.exp.quality, P.exp.bitrateMbps, P.exp.bitrateMode, P.exp.hw].join('|');
  if (key === probeKey) return;
  probeKey = key;
  probing = true;
  clearTimeout(probeTimer);
  probeTimer = setTimeout(runProbe, 140);
}

function runProbe() {
  var key = probeKey, res = {};
  if (!dimsOk()) { probes = {}; probing = false; renderFormat(); renderExport(); return; }
  var chain = Promise.resolve();
  Core.CODEC_ORDER.forEach(function (fam) {
    chain = chain.then(function () { return Exp.probe(fam, probeSettings(fam)); }).then(function (r) { res[fam] = r; });
  });
  chain.then(function () {
    if (key !== probeKey) return;
    probes = res; probing = false;
    pickDefaultFormat();
    renderFormat(); renderExport();
  }, function () {
    if (key !== probeKey) return;
    probes = {}; probing = false;
    renderFormat(); renderExport();
  });
}

/* Only ever fills an empty choice. Once a format is selected, by the
   page or by the user, it is never swapped for another behind their back. */
function pickDefaultFormat() {
  if (P.exp.format) return;
  for (var i = 0; i < FORMATS.length; i++) {
    var r = probes[FORMATS[i].family];
    if (r && r.supported) {
      P.exp.format = FORMATS[i].value;
      P.exp.autoNote = FORMATS[i].family === 'avc' ? '' : ((probes.avc && probes.avc.reason) || 'H.264 is not available here.');
      return;
    }
  }
}

function renderFormat() {
  var sel = $('format'), fmt = currentFormat(), items = [];
  sel.textContent = '';
  var known = Object.keys(probes).length > 0;
  if (!known && probing && env.canEncode) {
    var o0 = el('option', null, 'Checking this browser');
    o0.value = ''; sel.appendChild(o0);
  } else {
    if (!fmt) { var oe = el('option', null, 'Nothing available'); oe.value = ''; sel.appendChild(oe); }
    FORMATS.forEach(function (f) {
      var r = probes[f.family], ok = !!(r && r.supported);
      var o = el('option', null, f.label + (ok ? '' : '  (not available)'));
      o.value = f.value;
      o.disabled = !ok && f.value !== P.exp.format;
      sel.appendChild(o);
    });
    sel.value = fmt ? fmt.value : '';
  }
  setSeg($('quality'), P.exp.quality);
  Array.prototype.forEach.call($('quality').querySelectorAll('button'), function (b) { b.disabled = !!P.exp.bitrateMbps; });

  if (!env.canEncode) {
    items.push({ kind: 'bad', text: env.reason });
  } else if (fmt && known) {
    var r = probes[fmt.family];
    if (r && r.supported) {
      if (P.exp.autoNote && !P.exp.userPicked) {
        items.push({ kind: 'warn', text: 'H.264 in MP4 is the usual choice for show playback but this browser cannot encode it here, so ' + fmt.label + ' is selected instead.', sub: P.exp.autoNote });
      }
      items.push({ kind: 'info', text: r.candidate.label + ', 8 bit 4:2:0, ' + Core.trimNum(r.config.bitrate / 1e6, 1) + ' Mbit/s' + (P.exp.bitrateMbps ? ' set by hand' : '') + '.', sub: PLAYBACK_NOTE[fmt.family] });
      if (fmt.family === 'avc' && r.config.bitrate > 60e6) {
        items.push({ kind: 'warn', text: 'That is above the 60 Mbit/s the VT Inspector flags for PowerPoint on Windows. Fine for a media server, heavy for a laptop.' });
      }
    } else if (r) {
      var acts = [];
      if (r.oddDims) {
        var ew = P.output.w + (P.output.w % 2), eh = P.output.h + (P.output.h % 2);
        acts.push({ label: 'Use ' + ew + ' x ' + eh, solid: true, fn: function () {
          P.output.w = ew; P.output.h = eh; $('out-w').value = ew; $('out-h').value = eh; P.lockRatio = ew / eh; refresh();
        } });
      }
      FORMATS.forEach(function (f) {
        var pr = probes[f.family];
        if (pr && pr.supported && acts.length < 3) acts.push({ label: 'Use ' + f.label, fn: function () { P.exp.format = f.value; P.exp.userPicked = true; P.exp.autoNote = ''; refresh(); } });
      });
      items.push({ kind: 'bad', text: r.reason, sub: 'Nothing has been changed for you. Pick one of these or adjust the settings.', actions: acts });
    }
  } else if (known && !fmt) {
    var why = (probes.avc && probes.avc.reason) || 'No codec is available.';
    items.push({ kind: 'bad', text: 'This browser cannot encode any of the formats at these settings.', sub: why });
  }
  setMsgs($('format-msgs'), items);

  var fam = fmt ? fmt.family : 'avc';
  $('bitrate').placeholder = dimsOk() ? 'auto ' + Core.trimNum(Core.suggestBitrate(fam, P.output.w, P.output.h, P.output.fps, P.exp.quality) / 1e6, 1) : 'auto';
  $('bitrate').classList.toggle('bad', P.exp.bitrateBad);
  $('key-secs').classList.toggle('bad', P.exp.keyBad);
  var adv = [];
  if (P.exp.bitrateMbps) adv.push({ kind: 'info', text: 'A bitrate set by hand replaces the quality preset. Clear the field to go back to the preset.' });
  if (P.exp.bitrateBad) adv.push({ kind: 'bad', text: 'Enter a bitrate between 0.5 and 800 Mbit/s, or leave it empty.' });
  if (P.exp.keyBad) adv.push({ kind: 'bad', text: 'Enter a key frame spacing between 0.1 and 10 seconds.' });
  if (!env.canStreamToDisk) adv.push({ kind: 'info', text: 'This browser cannot write a file straight to disk, so the video is built in memory and then downloaded. Chrome and Edge on a desktop can.' });
  adv.push({ kind: 'info', text: 'ProRes, HAP, DNxHR and NotchLC cannot be encoded by a browser. If the media server needs one of them, export H.264 at Maximum and transcode that.' });
  setMsgs($('adv-export-msgs'), adv);
  $('to-disk').disabled = !env.canStreamToDisk;
  if (!env.canStreamToDisk) $('to-disk').value = 'off';
}

function bindFormat() {
  $('format').addEventListener('change', function () {
    if (!this.value) return;
    P.exp.format = this.value; P.exp.userPicked = true; P.exp.autoNote = '';
    refresh();
  });
  bindSeg($('quality'), function (v) { P.exp.quality = v; refresh(); });
  $('bitrate').addEventListener('input', function () {
    var raw = this.value.trim(), v = parseNumber(raw);
    P.exp.bitrateBad = raw !== '' && !(v >= 0.5 && v <= 800);
    P.exp.bitrateMbps = raw === '' || P.exp.bitrateBad ? null : v;
    refreshSoon();
  });
  $('bitrate-mode').addEventListener('change', function () { P.exp.bitrateMode = this.value; refresh(); });
  $('key-secs').addEventListener('input', function () {
    var v = parseNumber(this.value);
    P.exp.keyBad = !(v >= 0.1 && v <= 10);
    if (!P.exp.keyBad) P.exp.keySeconds = v;
    refreshSoon();
  });
  $('hw').addEventListener('change', function () { P.exp.hw = this.value; refresh(); });
  $('to-disk').addEventListener('change', function () { P.exp.toDisk = this.value; refresh(); });
}

/* ------------------------------------------------------------------ *
 * Summary, warnings and the export button                             *
 * ------------------------------------------------------------------ */

function exportEstimate() {
  var fmt = currentFormat(), r = fmt && probes[fmt.family];
  var bitrate = r && r.supported ? r.config.bitrate : (fmt ? bitrateFor(fmt.family) : 0);
  var maxPx = 0;
  plan.segs.forEach(function (s) { maxPx = Math.max(maxPx, s.iw * s.ih); });
  var bytes = Core.estimateBytes(bitrate, plan.frames / plan.fps);
  var toDisk = env.canStreamToDisk && (P.exp.toDisk === 'on' || (P.exp.toDisk === 'auto' && bytes > DISK_AUTO_BYTES));
  var res = Core.assessResources({
    W: plan.W, H: plan.H, fps: plan.fps, frames: plan.frames, photoCount: plan.segs.length, maxPhotoPixels: maxPx,
    bitrate: bitrate, toDisk: toDisk, canStream: env.canStreamToDisk, deviceMemory: env.deviceMemory, cores: env.cores
  });
  return { bitrate: bitrate, bytes: bytes, toDisk: toDisk, res: res, supported: !!(r && r.supported) };
}

function exportBlocks() {
  var blocks = [];
  if (!env.canEncode) { blocks.push(env.reason); return blocks; }
  if (imp.running) blocks.push('Still reading photos. Export is available as soon as they are in.');
  plan.errors.forEach(function (e) { blocks.push(e.msg); });
  if (!plan.segs.length) return blocks;
  var fmt = currentFormat();
  if (probing) blocks.push('Checking what this browser can encode at these settings.');
  else if (!fmt) blocks.push('No video format is available in this browser at these settings.');
  else if (!probes[fmt.family] || !probes[fmt.family].supported) blocks.push((probes[fmt.family] && probes[fmt.family].reason) || 'The selected format is not available.');
  if (P.exp.bitrateBad) blocks.push('The bitrate under Advanced is not valid.');
  if (P.exp.keyBad) blocks.push('The key frame spacing under Advanced is not valid.');
  return blocks;
}

function addRow(dl, k, v, small) {
  dl.appendChild(el('dt', null, k));
  var dd = el('dd', null, v);
  if (small) { dd.appendChild(document.createTextNode(' ')); dd.appendChild(el('small', null, small)); }
  dl.appendChild(dd);
}

function renderExport() {
  var dl = $('summary'), fmt = currentFormat(), n = plan.counts.included, ex = plan.counts.excluded;
  dl.textContent = '';
  var est = n && plan.frames ? exportEstimate() : null;
  addRow(dl, 'Photos', n + ' in video', ex ? ex + ' excluded' : (photos.length ? 'none excluded' : ''));
  if (dimsOk()) addRow(dl, 'Output', P.output.w + ' x ' + P.output.h, Core.aspectLabel(P.output.w, P.output.h) + ', ' + P.output.fps + ' fps');
  else addRow(dl, 'Output', 'not valid');
  addRow(dl, 'Duration', n ? Core.formatTime(plan.frames, plan.fps) : '0:00.00',
    n ? nf(plan.frames) + ' frames exactly' + (P.loop ? ', seamless loop' : ', plays once') : '');
  var r = fmt && probes[fmt.family];
  addRow(dl, 'Format', fmt ? fmt.label : (probing && env.canEncode ? 'checking' : 'none'),
    r && r.supported ? r.candidate.label.replace(/^[^ ]+ /, '') + ', ' + Core.trimNum(r.config.bitrate / 1e6, 1) + ' Mbit/s' : '');
  if (est && est.supported) {
    addRow(dl, 'File size', 'about ' + Core.formatBytes(est.bytes), 'if the encoder holds its bitrate. Slideshows often come in smaller.');
    addRow(dl, 'Written', est.toDisk ? 'straight to disk' : 'in memory', est.toDisk ? 'you choose where when the export starts' : 'then offered as a download');
    addRow(dl, 'Memory', 'roughly ' + Core.formatBytes(est.res.memBytes), 'working estimate, not a limit');
  }

  var line = $('dock-line');
  line.textContent = '';
  if (n && dimsOk()) {
    [n + ' photo' + (n === 1 ? '' : 's'), P.output.w + ' x ' + P.output.h, P.output.fps + ' fps', Core.formatTime(plan.frames, plan.fps),
     fmt ? fmt.label : '', est && est.supported ? 'about ' + Core.formatBytes(est.bytes) : ''].forEach(function (t, i) {
      if (!t) return;
      if (line.childNodes.length) line.appendChild(document.createTextNode('  \u00b7  '));
      line.appendChild(el('b', null, t));
    });
  } else line.textContent = photos.length ? 'Nothing to export yet.' : 'Add photos to begin.';

  var items = [], blocks = exportBlocks();
  if (photos.length || !env.canEncode) blocks.forEach(function (b) { items.push({ kind: 'bad', text: b }); });
  if (est && !blocks.length) {
    est.res.risks.forEach(function (w) { items.push({ kind: 'warn', text: w, sub: 'This is an estimate of how heavy the job is, not a prediction. You can export anyway.' }); });
    if (est.toDisk) items.push({ kind: 'info', text: 'You will be asked where to save the file before the export starts, so it can be written straight to disk.' });
  }
  if (ui.exporting) items = [];
  setMsgs($('export-msgs'), items);
  var btn = $('btn-export');
  btn.disabled = ui.exporting || ui.starting || blocks.length > 0;
  btn.hidden = ui.exporting;
  $('progress').hidden = !ui.exporting;
}

/* ------------------------------------------------------------------ *
 * Export                                                              *
 * ------------------------------------------------------------------ */

function lockUi(on) { $('settings').disabled = on; }

function clearResult() {
  if (ui.resultUrl) { try { URL.revokeObjectURL(ui.resultUrl); } catch (e) {} ui.resultUrl = ''; }
  $('result').textContent = '';
  $('dock').classList.remove('settled');
}

/* While there is an outcome to read, the dock stops floating so it cannot sit on top of the settings. */
function settleDock() {
  var d = $('dock');
  d.classList.add('settled');
  if (d.scrollIntoView) d.scrollIntoView({ block: 'nearest' });
}

function fmtSpan(ms) {
  var s = Math.max(0, Math.round(ms / 1000));
  if (s < 60) return s + ' s';
  var m = Math.floor(s / 60);
  return m + ' min ' + (s % 60) + ' s';
}

async function startExport() {
  if (ui.exporting || ui.starting || exportBlocks().length) return;
  var fmt = currentFormat(), probe = probes[fmt.family], est = exportEstimate();
  var con = Core.CONTAINERS[fmt.container];
  var name = Core.outputFilename(plan, fmt.container);
  var handle = null, diskNote = '';
  if (est.toDisk) {
    /* the save dialog is open for as long as the user likes, so hold the button meanwhile */
    ui.starting = true;
    $('btn-export').disabled = true;
    try {
      var accept = {}; accept[con.mime] = ['.' + con.ext];
      handle = await window.showSaveFilePicker({ suggestedName: name, types: [{ description: con.label + ' video', accept: accept }] });
      name = handle.name;
    } catch (e) {
      handle = null;
      if (e && e.name === 'AbortError') { ui.starting = false; renderExport(); return; }   /* they closed the dialog, so do nothing */
      diskNote = 'The save dialog could not be opened, so the video is being built in memory and will be offered as a download. ' + ((e && e.message) || '');
    }
    ui.starting = false;
    if (exportBlocks().length) { renderExport(); return; }
  }
  stopPlayback();
  clearResult();
  if (diskNote) setMsgs($('result'), [{ kind: 'warn', text: diskNote }]);

  /* Freeze exactly what the preview is showing. Later edits cannot reach a running export. */
  var job = {
    plan: { W: plan.W, H: plan.H, fps: plan.fps, loop: plan.loop, frames: plan.frames, segs: JSON.parse(JSON.stringify(plan.segs)) },
    photos: {},
    family: fmt.family, container: fmt.container,
    encoder: probe.config,
    keyFrames: Math.max(1, Math.round(P.exp.keySeconds * plan.fps)),
    target: handle ? { kind: 'file', handle: handle } : { kind: 'memory' }
  };
  plan.segs.forEach(function (s) { var p = byId.get(s.id); job.photos[s.id] = { file: p.file, blur: s.mode === 'blur' ? p.blur : null }; });
  var expect = { W: plan.W, H: plan.H, fps: plan.fps, frames: plan.frames, loop: plan.loop, label: fmt.label, family: fmt.family,
                 container: fmt.container, mime: con.mime, name: name, profile: probe.candidate.label, bitrate: probe.config.bitrate };

  ui.exporting = true; ui.cancelling = false; ui.mode = '';
  lockUi(true);
  showProgress({ stage: 'prepare' });
  refresh();
  var began = Date.now();
  var launchOpts = { forceMainThread: /[?&]export=page(&|$)/.test(window.location.search) };
  ui.exportCtl = Exp.launch(job, {
    mode: function (m) { ui.mode = m; },
    progress: showProgress,
    done: function (result) {
      endExport();
      var file;
      if (result.kind === 'memory') file = Promise.resolve(new Blob([result.buffer], { type: con.mime }));
      else file = handle.getFile();
      file.then(function (blob) { showResult(blob, result, expect, Date.now() - began, !!handle); },
                function (err) { showFailure(new Error('The file was written but could not be reopened for checking. ' + (err && err.message || ''))); });
    },
    cancelled: function () {
      endExport();
      if (handle && handle.remove) handle.remove().catch(function () {});
      setMsgs($('result'), [{ kind: 'info', text: 'Export cancelled. No file was kept and the project is exactly as it was.',
        actions: [{ label: 'Dismiss', fn: clearResult }] }]);
      settleDock();
      say('Export cancelled.');
    },
    error: function (err) {
      endExport();
      if (handle && handle.remove) handle.remove().catch(function () {});
      showFailure(err);
    }
  }, launchOpts);
}

function endExport() {
  ui.exporting = false; ui.cancelling = false; ui.exportCtl = null;
  lockUi(false);
  refresh();
}

var STAGE_TEXT = { prepare: 'Preparing the encoder', encode: 'Rendering and encoding', finalize: 'Finishing the file' };

function showProgress(info) {
  var bar = $('progress-bar'), fill = $('progress-fill');
  $('progress-stage').textContent = ui.cancelling ? 'Cancelling' : (STAGE_TEXT[info.stage] || 'Working');
  if (info.stage === 'encode' && info.frames) {
    bar.classList.remove('busy');
    fill.style.width = (info.frame / info.frames * 100).toFixed(2) + '%';
    var t = 'frame ' + nf(info.frame) + ' of ' + nf(info.frames);
    /* Speed and time left come from frames actually finished, and only once there is enough to go on. */
    if (info.elapsedMs > 3000 && info.frame > 10) {
      var rate = info.frame / (info.elapsedMs / 1000);
      t += ', ' + Core.trimNum(rate, 1) + ' fps';
      if (info.frame < info.frames) t += ', about ' + fmtSpan((info.frames - info.frame) / rate * 1000) + ' left';
    }
    $('progress-detail').textContent = t;
  } else {
    bar.classList.add('busy');
    fill.style.width = '';
    $('progress-detail').textContent = info.stage === 'finalize' ? 'writing the index' : '';
  }
}

function showFailure(err) {
  var msg = (err && err.message) || String(err);
  setMsgs($('result'), [{
    kind: 'bad', text: 'Export failed. ' + msg,
    sub: 'This is what actually went wrong, not an estimate. The project is untouched, so you can change a setting and try again. A lower resolution or frame rate, another format, or Prefer software under Advanced are the usual fixes.',
    actions: [{ label: 'Dismiss', fn: clearResult }]
  }]);
  settleDock();
  say('Export failed.');
}

function showResult(blob, result, expect, tookMs, onDisk) {
  var box = $('result');
  box.textContent = '';
  var f = el('div', 'flag ok'), body = el('div');
  body.appendChild(el('span', null, 'Export complete. ' + expect.name + ', ' + Core.formatBytes(blob.size) + ', made in ' + fmtSpan(tookMs) + '.'));
  var row = el('div', 'btn-row');
  if (onDisk) {
    body.appendChild(el('span', 'sub', 'Saved where you chose. Nothing else to download.'));
  } else {
    ui.resultUrl = URL.createObjectURL(blob);
    var a = el('a', 'btn solid', 'Download video');
    a.href = ui.resultUrl; a.download = expect.name; a.id = 'download-link';
    row.appendChild(a);
    var x = el('button', 'btn', 'Discard');
    x.type = 'button';
    x.title = 'Frees the memory the finished video is using. Download it first.';
    x.addEventListener('click', function () { clearResult(); say('Finished video discarded.'); });
    row.appendChild(x);
  }
  body.appendChild(row);
  f.appendChild(body);
  box.appendChild(f);
  var checks = el('dl', 'kv');
  checks.id = 'checks';
  var wrap = el('div', 'flag info'), wb = el('div');
  wb.appendChild(el('span', null, 'Checking the finished file'));
  wb.appendChild(checks);
  wrap.appendChild(wb);
  box.appendChild(wrap);
  settleDock();
  say('Export complete.');
  verifyOutput(blob, result, expect, checks, wb.firstChild);
}

function checkRow(dl, label, value, state) {
  dl.appendChild(el('dt', null, label));
  var dd = el('dd', state === 'good' ? 'good' : (state === 'poor' ? 'poor' : (state === 'soso' ? 'soso' : '')), value);
  dl.appendChild(dd);
}

function loadScript(src) {
  return new Promise(function (res, rej) {
    var s = document.createElement('script');
    s.src = src; s.onload = res; s.onerror = function () { rej(new Error('could not load ' + src)); };
    document.head.appendChild(s);
  });
}

/* Open the file in a video element and read back what it reports. */
function probePlayback(blob) {
  return new Promise(function (resolve) {
    var v = document.createElement('video'), url = URL.createObjectURL(blob), done = false;
    function fin(res) {
      if (done) return;
      done = true;
      clearTimeout(timer);
      v.removeAttribute('src');
      try { v.load(); } catch (e) {}
      URL.revokeObjectURL(url);
      resolve(res);
    }
    var timer = setTimeout(function () { fin(null); }, 10000);
    v.muted = true; v.preload = 'auto';
    v.addEventListener('loadeddata', function () { fin({ w: v.videoWidth, h: v.videoHeight, duration: v.duration }); });
    v.addEventListener('error', function () { fin(null); });
    v.src = url;
  });
}

/* Read the finished file back and compare it with what was asked for.
   MP4 goes through the same inspector the VT Inspector uses. */
async function verifyOutput(blob, result, expect, dl, heading) {
  var problems = 0;
  function row(label, value, ok) { if (ok === false) problems++; checkRow(dl, label, value, ok === true ? 'good' : (ok === false ? 'poor' : '')); }
  row('Frames written', nf(result.frames) + ' of ' + nf(expect.frames), result.frames === expect.frames);
  row('Encoder used', (result.codec || 'not reported') + ', ' + result.colour + ' colour', null);

  if (expect.container === 'mp4') {
    try {
      if (!window.MI) await loadScript(MI_SRC);
      var info = await window.MI.inspectFile(new File([blob], expect.name, { type: expect.mime }));
      var v = info.video && info.video[0];
      if (!info.ok || !v) { row('Container', 'could not be parsed', false); }
      else {
        var w = (v.cropped && v.cropped.width) || v.coded.width, h = (v.cropped && v.cropped.height) || v.coded.height;
        /* The measured rate, not the inspector's snapped label. Its snapping
           reads an exact 30 as 29.97, which would be a false alarm here. */
        var fr = v.frameRate || {}, fps = fr.fps || fr.avgFps;
        row('Codec in file', v.codec.name + (v.profile ? ' ' + v.profile : '') + (v.level ? ' L' + v.level : ''), null);
        row('Picture', w + ' x ' + h, w === expect.W && h === expect.H);
        row('Frame rate', (fps ? Core.trimNum(fps, 3) : '?') + ' fps ' + (fr.mode || ''), !!fps && Math.abs(fps - expect.fps) < 0.01 && (fr.mode || 'CFR') === 'CFR');
        row('Frames in file', nf(v.sampleCount), v.sampleCount === expect.frames);
        var dur = v.duration != null ? v.duration : info.duration;
        row('Duration', Core.trimNum(dur, 3) + ' s', Math.abs(dur - expect.frames / expect.fps) < 0.5 / expect.fps + 0.002);
        row('Index at front', info.faststart ? 'yes' : 'no', info.faststart ? true : null);
        if (v.colour) row('Colour', v.colour.matrix + ' matrix' + (v.colour.fullRange ? ', full range' : ', video range'), v.colour.matrix === result.colour);
        try {
          var risk = window.MI.assess(info, { target: 'powerpoint-win' });
          checkRow(dl, 'VT Inspector', risk.summary, risk.verdict === 'green' ? 'good' : (risk.verdict === 'red' ? 'poor' : 'soso'));
        } catch (e) {}
      }
    } catch (e) {
      row('Container', 'not inspected, ' + ((e && e.message) || e), null);
    }
  }

  var pb = await probePlayback(blob);
  if (!pb) {
    checkRow(dl, 'Opens here', 'this browser could not play it back, so playback is unchecked', 'soso');
  } else {
    row('Opens here', pb.w + ' x ' + pb.h + ', ' + Core.trimNum(pb.duration, 3) + ' s',
      pb.w === expect.W && pb.h === expect.H && Math.abs(pb.duration - expect.frames / expect.fps) < 1.5 / expect.fps + 0.005);
  }
  heading.textContent = problems ? 'The finished file does not match the settings in ' + problems + ' place' + (problems === 1 ? '' : 's') + '. Check before using it.'
    : 'Finished file checked against the settings';
  heading.parentNode.parentNode.className = 'flag ' + (problems ? 'bad' : 'info');
}

function bindExport() {
  $('btn-export').addEventListener('click', function () { startExport(); });
  $('btn-cancel').addEventListener('click', function () {
    if (!ui.exportCtl || ui.cancelling) return;
    ui.cancelling = true;
    $('progress-stage').textContent = 'Cancelling';
    ui.exportCtl.cancel();
  });
  window.addEventListener('beforeunload', function (e) {
    if (!ui.exporting) return;
    e.preventDefault();
    e.returnValue = '';
  });
}

/* ------------------------------------------------------------------ *
 * Page level wiring                                                   *
 * ------------------------------------------------------------------ */

function bindFiles() {
  var input = $('file-input'), stage = $('stage');
  function browse() { if (!ui.exporting) input.click(); }
  $('drop').addEventListener('click', browse);
  $('btn-add').addEventListener('click', browse);
  input.addEventListener('change', function () { addFiles(this.files); this.value = ''; });

  var depth = 0;
  function isFiles(e) { var t = e.dataTransfer && e.dataTransfer.types; return !!t && Array.prototype.indexOf.call(t, 'Files') >= 0; }
  document.addEventListener('dragenter', function (e) { if (!isFiles(e)) return; depth++; stage.classList.add('drag'); });
  document.addEventListener('dragleave', function (e) { if (!isFiles(e)) return; depth = Math.max(0, depth - 1); if (!depth) stage.classList.remove('drag'); });
  document.addEventListener('dragover', function (e) { if (isFiles(e)) { e.preventDefault(); e.dataTransfer.dropEffect = ui.exporting ? 'none' : 'copy'; } });
  document.addEventListener('drop', function (e) {
    if (!isFiles(e)) return;
    e.preventDefault();
    depth = 0; stage.classList.remove('drag');
    addFiles(e.dataTransfer.files);
  });

  $('btn-random').addEventListener('click', randomiseOrder);
  $('btn-restore').addEventListener('click', restoreImportOrder);

  /* Clearing wipes the whole project, so it takes a second press. */
  var clr = $('btn-clear'), armTimer = 0;
  function disarm() { clr.classList.remove('armed'); clr.textContent = 'Clear project'; clearTimeout(armTimer); }
  clr.addEventListener('click', function () {
    if (!clr.classList.contains('armed')) {
      clr.classList.add('armed');
      clr.textContent = 'Press again to clear';
      armTimer = setTimeout(disarm, 4000);
      return;
    }
    disarm();
    clearProject();
  });
  clr.addEventListener('blur', disarm);
}

/* The calculator can hand its size over in the link, the same way its share links work. */
function readUrl() {
  var q;
  try { q = new URLSearchParams(window.location.search); } catch (e) { return; }
  var w = parseInt(q.get('w'), 10), h = parseInt(q.get('h'), 10), fps = parseFloat(q.get('fps'));
  if (String(w) === q.get('w') && String(h) === q.get('h') && !Core.validateDims(w, h)) {
    P.output.w = w; P.output.h = h; P.lockRatio = w / h;
    $('out-w').value = w; $('out-h').value = h;
  }
  if (Core.FPS_CHOICES.indexOf(fps) >= 0) P.output.fps = fps;
}

function init() {
  bindFiles(); bindTray(); bindSelection(); bindPreview();
  bindOutput(); bindTiming(); bindMotion(); bindFit(); bindFormat(); bindExport();
  readUrl();
  refresh();
}

/* Handles for the test scripts. Not used by the page itself. */
window.SlideshowApp = {
  state: function () { return { P: P, plan: plan, photos: photos, ui: ui, imp: imp, probes: probes, probing: probing, selId: selId, env: env }; },
  addFiles: addFiles, refresh: refresh, seek: seek, select: selectPhoto, draw: draw,
  frameReady: frameReady, randomiseOrder: randomiseOrder, startExport: startExport
};

init();
})();
