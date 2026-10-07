/* ============================================================
   PowerPoint to PDF core  (PptPdfCore)

   Everything the batch converter decides without touching the
   DOM, the disk or PowerPoint: which settings a platform can
   honour, output names, page geometry for resized output, the
   job ticket handed to the local helper, what counts as an
   issue and how sure we are of it, the result state of each
   file, the reports, and the picture comparison.

   Pure functions, so the same file runs in the page and in Node
   for the tests.

   Three words are used with care throughout.
     confirmed   read from the file or reported by PowerPoint
     suspected   inferred, likely but not proven
     unverified  a check that could not be run

   AEGFX / SlideSize
   ============================================================ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PptPdfCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
'use strict';

var PROTOCOL = 1;
var HELPER_VERSION = '1.0.0';
var ENGINE_NAME = 'Microsoft PowerPoint';
var WORK_DIR = '_slidesize';
var PT_PER_IN = 72, PT_PER_CM = 72 / 2.54;
var PPT_MAX_PT = 56 * 72;                         /* PowerPoint's largest slide side */
var PDF_MAX_PT = 14400;                           /* a PDF page side stops at 200 inches */
var SOURCE_WARN_BYTES = 2 * 1024 * 1024 * 1024;   /* decks above this are outside what has been tested */
var PDF_EDIT_MAX_BYTES = 300 * 1024 * 1024;       /* PDFs above this are not opened for inspection or resizing */
var HEARTBEAT_STALE_MS = 12000;

/* ------------------------------------------------------------------ *
 * Small helpers                                                       *
 * ------------------------------------------------------------------ */

function clamp(v, lo, hi) { return Math.min(hi, Math.max(lo, v)); }
function gcd(a, b) { a = Math.abs(Math.round(a)); b = Math.abs(Math.round(b)); while (b) { var t = b; b = a % b; a = t; } return a; }
function trimNum(n, dp) { return String(+Number(n).toFixed(dp === undefined ? 2 : dp)); }

function aspectLabel(w, h) {
  if (!(w > 0 && h > 0)) return '';
  var known = [[16, 9], [4, 3], [16, 10], [3, 2], [1, 1], [21, 9], [32, 9], [5, 4], [9, 16], [3, 4], [2, 3]];
  for (var i = 0; i < known.length; i++) {
    if (Math.abs(w / h - known[i][0] / known[i][1]) < 0.004) return known[i][0] + ':' + known[i][1];
  }
  var rw = Math.round(w * 100), rh = Math.round(h * 100), g = gcd(rw, rh) || 1, a = rw / g, b = rh / g;
  if (a <= 40 && b <= 40) return a + ':' + b;
  return w >= h ? trimNum(w / h, 2) + ':1' : '1:' + trimNum(h / w, 2);
}

function ptToCm(pt) { return pt / PT_PER_CM; }
function ptToIn(pt) { return pt / PT_PER_IN; }
function cmToPt(cm) { return cm * PT_PER_CM; }
function sizeLabel(w, h) {
  return trimNum(ptToCm(w), 2) + ' x ' + trimNum(ptToCm(h), 2) + ' cm';
}
function orientationOf(w, h) { return Math.abs(w - h) < 0.01 ? 'square' : w > h ? 'landscape' : 'portrait'; }

function formatBytes(b) {
  if (b == null || isNaN(b)) return '';
  if (b < 1024) return b + ' B';
  if (b < 1048576) return (b / 1024).toFixed(0) + ' KB';
  if (b < 1073741824) return (b / 1048576).toFixed(1) + ' MB';
  return (b / 1073741824).toFixed(2) + ' GB';
}

function formatDuration(ms) {
  if (ms == null || isNaN(ms)) return '';
  var s = ms / 1000;
  if (s < 10) return trimNum(s, 1) + ' s';
  if (s < 60) return Math.round(s) + ' s';
  var m = Math.floor(s / 60), r = Math.round(s - m * 60);
  if (m < 60) return m + ' min ' + (r < 10 ? '0' : '') + r + ' s';
  return Math.floor(m / 60) + ' h ' + (m % 60) + ' min';
}

function plural(n, one, many) { return n + ' ' + (n === 1 ? one : (many || one + 's')); }

function listSlides(nums, max) {
  if (!nums || !nums.length) return '';
  var a = nums.slice().sort(function (x, y) { return x - y; }), out = [], i = 0;
  while (i < a.length) {
    var j = i; while (j + 1 < a.length && a[j + 1] === a[j] + 1) j++;
    out.push(j > i + 1 ? a[i] + ' to ' + a[j] : j === i + 1 ? a[i] + ', ' + a[j] : String(a[i]));
    i = j + 1;
  }
  max = max || 12;
  return out.length > max ? out.slice(0, max).join(', ') + ' and more' : out.join(', ');
}

function esc(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/* ------------------------------------------------------------------ *
 * Platform and capability                                             *
 * ------------------------------------------------------------------ */

function detectPlatform(nav) {
  nav = nav || {};
  var p = ((nav.userAgentData && nav.userAgentData.platform) || nav.platform || '').toLowerCase();
  var ua = (nav.userAgent || '').toLowerCase();
  if (/android|iphone|ipad|ipod/.test(ua) || (nav.userAgentData && nav.userAgentData.mobile)) return 'other';
  if (p.indexOf('win') === 0 || /windows nt/.test(ua)) return 'windows';
  if (p.indexOf('mac') === 0 || /mac os x|macintosh/.test(ua)) return (nav.maxTouchPoints > 1 && !/chrome|edg\//.test(ua)) ? 'other' : 'macos';
  return 'other';
}

function platformName(p) { return p === 'windows' ? 'Windows' : p === 'macos' ? 'macOS' : 'this system'; }

/* What the page itself needs from the browser. */
function browserSupport(win) {
  win = win || {};
  var missing = [];
  if (typeof win.showDirectoryPicker !== 'function') missing.push('folder access (the File System Access API)');
  if (typeof win.DecompressionStream !== 'function') missing.push('stream decompression');
  if (win.isSecureContext === false) missing.push('a secure page address (https or localhost)');
  return {
    ok: missing.length === 0, missing: missing,
    message: missing.length ? 'This browser cannot run the converter. It is missing ' + missing.join(' and ') +
      '. Use a current Chrome or Edge on a Windows PC or a Mac.' : ''
  };
}

/* Settings that depend on what PowerPoint offers on each platform. A helper
   reports its own capability list once it is running and that list wins.
   The reasons are shown beside the disabled control. */
var FEATURES = {
  pdfa: {
    label: 'PDF/A',
    windows: true, macos: false,
    why: { macos: 'PowerPoint for Mac cannot write PDF/A. Convert on a Windows PC when the file must be PDF/A.' }
  },
  notes: {
    label: 'Notes pages',
    windows: true, macos: false,
    why: { macos: 'PowerPoint for Mac only makes notes pages through its Print dialog, which the helper cannot drive. Convert on a Windows PC for notes pages.' }
  },
  quality: {
    label: 'Image compression',
    windows: true, macos: false,
    why: { macos: 'PowerPoint for Mac exports PDF at one quality and offers no choice.' }
  },
  bitmapText: {
    label: 'Bitmap text when fonts cannot be embedded',
    windows: true, macos: false,
    why: { macos: 'PowerPoint for Mac decides this itself and offers no switch.' }
  },
  tags: {
    label: 'Accessibility tags',
    windows: true, macos: false,
    why: { macos: 'PowerPoint for Mac offers no switch for structure tags when a script exports the PDF.' }
  },
  docProps: {
    label: 'Document properties',
    windows: true, macos: false,
    why: { macos: 'PowerPoint for Mac offers no switch. Use Strip metadata to remove the properties afterwards.' }
  },
  markup: {
    label: 'Comments and ink',
    windows: true, macos: false,
    why: { macos: 'PowerPoint for Mac does not write comments or ink into the PDF and offers no switch.' }
  },
  hidden: { label: 'Hidden slides', windows: true, macos: true, why: {} },
  range: { label: 'Slide range', windows: true, macos: true, why: {} },
  reference: { label: 'Reference renders for the fidelity check', windows: true, macos: true, why: {} },
  placeholders: { label: 'Video placeholders', windows: true, macos: true, why: {} }
};

function featureAvailable(key, platform, caps) {
  var f = FEATURES[key];
  if (!f) return { ok: true, reason: '' };
  if (platform !== 'windows' && platform !== 'macos') {
    return { ok: false, reason: 'PowerPoint conversion runs on Windows and macOS only.' };
  }
  if (caps && Object.prototype.hasOwnProperty.call(caps, key)) {
    if (caps[key]) return { ok: true, reason: '' };
    return { ok: false, reason: (caps.reasons && caps.reasons[key]) || f.why[platform] ||
      'The helper on this computer reports that ' + f.label.toLowerCase() + ' is not available with the installed PowerPoint.' };
  }
  return f[platform] ? { ok: true, reason: '' } : { ok: false, reason: f.why[platform] || 'Not available on ' + platformName(platform) + '.' };
}

/* ------------------------------------------------------------------ *
 * Settings                                                            *
 * ------------------------------------------------------------------ */

var SIZE_PRESETS = [
  { id: 'original', name: 'Original' },
  { id: '16:9', name: '16:9', w: 960, h: 540, note: '33.87 x 19.05 cm' },
  { id: '4:3', name: '4:3', w: 720, h: 540, note: '25.4 x 19.05 cm' },
  { id: 'a4', name: 'A4', w: 841.8898, h: 595.2756, note: '29.7 x 21 cm' },
  { id: 'letter', name: 'Letter', w: 792, h: 612, note: '11 x 8.5 in' },
  { id: 'custom', name: 'Custom' }
];

var DEFAULTS = {
  size: { preset: 'original', orientation: 'auto', fit: 'fit', margin: 'none', customW: 960, customH: 540 },
  pdfa: false,
  hidden: false,                 /* include hidden slides */
  range: null,                   /* { from, to } in slide numbers, or null for all */
  output: 'slides',              /* slides | notes */
  quality: 'standard',           /* standard | minimum */
  bitmapText: true,
  tags: true,
  docProps: true,
  markup: false,
  links: 'keep',                 /* keep | remove */
  metadata: 'keep',              /* keep | strip */
  fidelity: 'all',               /* all | sample | off */
  sensitivity: 'normal',         /* low | normal | high */
  referenceLongEdge: 1280,
  keepReferences: false,
  placeholders: true,
  allowRemoteLinks: false,
  timeoutSec: 600
};

/* Presets that ship with the page. Users add their own in local storage. */
var BUILTIN_PRESETS = [
  { id: 'standard', name: 'Standard', builtin: true, settings: {} },
  { id: 'small', name: 'Small file', builtin: true, settings: { quality: 'minimum', fidelity: 'sample' } },
  { id: 'archive', name: 'Archive PDF/A', builtin: true, settings: { pdfa: true, bitmapText: true, tags: true, docProps: true } },
  { id: 'handout', name: 'A4 handout with hidden slides', builtin: true,
    settings: { size: { preset: 'a4', orientation: 'auto', fit: 'fit', margin: 'white' }, hidden: true } }
];

function clone(o) { return JSON.parse(JSON.stringify(o)); }

function mergeSettings(base, over) {
  var out = clone(base);
  if (!over) return out;
  Object.keys(over).forEach(function (k) {
    if (!Object.prototype.hasOwnProperty.call(DEFAULTS, k)) return;
    if (k === 'size' && over.size) Object.keys(over.size).forEach(function (s) { out.size[s] = over.size[s]; });
    else out[k] = over[k] === null ? null : clone(over[k]);
  });
  return out;
}

function defaultSettings(over) { return mergeSettings(DEFAULTS, over); }

/* Turns what was asked for into what this platform will do.
   Returns { effective, disabled: {key: reason}, notes: [{key, level, text}], errors: [text] } */
function resolveSettings(settings, platform, caps) {
  var s = mergeSettings(DEFAULTS, settings), disabled = {}, notes = [], errors = [];
  function gate(key, off) {
    var a = featureAvailable(key, platform, caps);
    if (!a.ok) { disabled[key] = a.reason; off(); }
  }
  gate('pdfa', function () { s.pdfa = false; });
  gate('notes', function () { s.output = 'slides'; });
  gate('quality', function () { s.quality = 'standard'; });
  gate('bitmapText', function () { s.bitmapText = DEFAULTS.bitmapText; });
  gate('tags', function () { s.tags = DEFAULTS.tags; });
  gate('docProps', function () { s.docProps = DEFAULTS.docProps; });
  gate('markup', function () { s.markup = false; });
  gate('hidden', function () { s.hidden = false; });
  gate('range', function () { s.range = null; });
  gate('reference', function () { s.fidelity = 'off'; });
  gate('placeholders', function () { s.placeholders = false; });

  if (s.range) {
    var f = Math.round(+s.range.from), t = Math.round(+s.range.to);
    if (!(f >= 1) || !(t >= f)) errors.push('The slide range needs a first slide of 1 or more and a last slide that is not before it.');
    else s.range = { from: f, to: t };
  }
  if (s.size.preset === 'custom') {
    var w = +s.size.customW, h = +s.size.customH;
    if (!(w >= 72 && h >= 72)) errors.push('The custom page size needs a width and a height of at least 2.54 cm.');
    else if (w > PDF_MAX_PT || h > PDF_MAX_PT) errors.push('A PDF page cannot be larger than 508 cm on a side.');
  }
  if (!(s.timeoutSec >= 30)) s.timeoutSec = 30;
  if (s.timeoutSec > 7200) s.timeoutSec = 7200;

  var edits = [];
  if (s.size.preset !== 'original') edits.push('resizes the pages');
  if (s.links === 'remove') edits.push('removes the links');
  if (s.pdfa) {
    if (s.metadata === 'strip') {
      disabled.metadata = 'PDF/A requires the metadata to be present, so it cannot be stripped from a PDF/A file.';
      s.metadata = 'keep';
    }
    if (!s.bitmapText) notes.push({ key: 'pdfa', level: 'warn',
      text: 'PDF/A needs every font embedded. With bitmap text off, a font that cannot be embedded will make the file fail validation.' });
    if (!s.tags) notes.push({ key: 'pdfa', level: 'warn',
      text: 'PowerPoint writes PDF/A level A, which requires structure tags. Without tags the file is likely to fail validation.' });
    if (edits.length) notes.push({ key: 'pdfa', level: 'warn',
      text: 'SlideSize ' + edits.join(' and ') + ' after PowerPoint has written the file. The PDF/A checks run on the finished file, so a change that breaks conformance is reported.' });
    notes.push({ key: 'pdfa', level: 'info',
      text: 'PowerPoint offers one PDF/A setting, ISO 19005-1. The part and level it actually wrote are read back from each file.' });
  } else if (s.metadata === 'strip') edits.push('strips the metadata');
  if (s.output === 'notes') {
    if (s.fidelity !== 'off') notes.push({ key: 'fidelity', level: 'info',
      text: 'The picture comparison is not run on notes pages. A notes page is a different layout from the slide it shows.' });
    if (s.size.preset !== 'original') notes.push({ key: 'size', level: 'info',
      text: 'With notes pages the resize applies to the whole notes page, not to the slide picture inside it.' });
  }
  if (!s.bitmapText && !disabled.bitmapText) notes.push({ key: 'bitmapText', level: 'info',
    text: 'With bitmap text off, text in a font that cannot be embedded stays as text but only displays correctly on a computer that has the font.' });
  if (s.tags === false && !disabled.tags) notes.push({ key: 'tags', level: 'info',
    text: 'Without accessibility tags a screen reader has no reading order to follow.' });
  if (platform === 'macos' && s.range) notes.push({ key: 'range', level: 'info',
    text: 'On macOS the helper leaves slides out by removing them from its temporary copy when PowerPoint will not skip them. Slide numbers printed on later slides then count from the slides that remain.' });
  return { effective: s, disabled: disabled, notes: notes, errors: errors };
}

function settingsFor(item, batchSettings) {
  return item && item.overrides ? mergeSettings(batchSettings, item.overrides) : mergeSettings(DEFAULTS, batchSettings);
}

/* A short line describing settings that differ from the defaults. */
function describeSettings(s) {
  var out = [];
  var size = SIZE_PRESETS.filter(function (p) { return p.id === s.size.preset; })[0];
  if (s.size.preset === 'original') out.push('Original slide size');
  else out.push((s.size.preset === 'custom' ? sizeLabel(s.size.customW, s.size.customH) : size.name) + ', ' +
    (s.size.fit === 'fill' ? 'filled and cropped' : 'fitted with margins'));
  out.push(s.pdfa ? 'PDF/A' : 'Standard PDF');
  out.push(s.output === 'notes' ? 'Notes pages' : 'Slides only');
  out.push(s.hidden ? 'Hidden slides included' : 'Hidden slides excluded');
  out.push(s.range ? 'Slides ' + s.range.from + ' to ' + s.range.to : 'All slides');
  out.push(s.quality === 'minimum' ? 'Minimum size images' : 'High quality images');
  if (s.bitmapText) out.push('Bitmap text when a font cannot be embedded');
  if (!s.tags) out.push('No accessibility tags');
  if (!s.docProps) out.push('No document properties');
  if (s.markup) out.push('Comments and ink included');
  if (s.links === 'remove') out.push('Links removed');
  if (s.metadata === 'strip') out.push('Metadata stripped');
  out.push('Fidelity check ' + (s.fidelity === 'all' ? 'on every slide' : s.fidelity === 'sample' ? 'on a sample' : 'off'));
  return out;
}

/* ------------------------------------------------------------------ *
 * Page geometry                                                       *
 * ------------------------------------------------------------------ */

function targetSize(srcW, srcH, size) {
  size = size || DEFAULTS.size;
  if (size.preset === 'original' || !(srcW > 0 && srcH > 0)) {
    return { w: srcW, h: srcH, changed: false, orientation: orientationOf(srcW, srcH) };
  }
  var p = SIZE_PRESETS.filter(function (x) { return x.id === size.preset; })[0];
  var a = size.preset === 'custom' ? +size.customW : p.w, b = size.preset === 'custom' ? +size.customH : p.h;
  var want = size.orientation === 'auto' || !size.orientation ? (srcW >= srcH ? 'landscape' : 'portrait') : size.orientation;
  var long = Math.max(a, b), short = Math.min(a, b);
  var w = want === 'portrait' ? short : long, h = want === 'portrait' ? long : short;
  if (size.preset === 'custom' && size.orientation === 'auto') { w = a; h = b; }
  return { w: w, h: h, changed: Math.abs(w - srcW) > 0.05 || Math.abs(h - srcH) > 0.05, orientation: orientationOf(w, h) };
}

/* Where a slide of srcW x srcH lands on a page of dstW x dstH.
   One scale for both axes, always. Nothing is ever stretched.
   tx and ty are in PDF space, origin at the bottom left. */
function fitTransform(srcW, srcH, dstW, dstH, mode) {
  var sx = dstW / srcW, sy = dstH / srcH;
  var s = mode === 'fill' ? Math.max(sx, sy) : Math.min(sx, sy);
  var tx = (dstW - s * srcW) / 2, ty = (dstH - s * srcH) / 2;
  var eps = 0.01;
  var cropX = tx < -eps ? (-tx / s) / srcW : 0, cropY = ty < -eps ? (-ty / s) / srcH : 0;
  return {
    scale: s, tx: tx, ty: ty,
    cropX: cropX, cropY: cropY,                         /* share of the slide cut from each side */
    marginX: tx > eps ? tx : 0, marginY: ty > eps ? ty : 0,
    cropped: cropX > 0.0005 || cropY > 0.0005,
    exact: Math.abs(sx - sy) < 1e-4,
    visibleShare: (1 - 2 * cropX) * (1 - 2 * cropY)
  };
}

/* The parts of the two pictures that show the same thing, as shares of each.
   y runs down from the top here, as it does in a canvas. */
function compareWindows(srcW, srcH, dstW, dstH, t) {
  if (!t) return { ref: { x: 0, y: 0, w: 1, h: 1 }, pdf: { x: 0, y: 0, w: 1, h: 1 } };
  var x0 = Math.max(0, t.tx), x1 = Math.min(dstW, t.tx + t.scale * srcW);
  var y0 = Math.max(0, t.ty), y1 = Math.min(dstH, t.ty + t.scale * srcH);
  var pdf = { x: x0 / dstW, y: (dstH - y1) / dstH, w: (x1 - x0) / dstW, h: (y1 - y0) / dstH };
  var rx0 = (x0 - t.tx) / t.scale, rx1 = (x1 - t.tx) / t.scale, ry0 = (y0 - t.ty) / t.scale, ry1 = (y1 - t.ty) / t.scale;
  var ref = { x: rx0 / srcW, y: (srcH - ry1) / srcH, w: (rx1 - rx0) / srcW, h: (ry1 - ry0) / srcH };
  return { ref: ref, pdf: pdf };
}

/* Which slides should end up in the PDF, by slide number.
   Returns null when the hidden slides are not known yet (a legacy .ppt before PowerPoint has opened it). */
function expectedSlides(facts, settings) {
  if (!facts || !(facts.slideCount >= 0)) return null;
  if (!facts.hiddenSlides && !settings.hidden) return null;
  var hidden = {}, out = [];
  (facts.hiddenSlides || []).forEach(function (n) { hidden[n] = true; });
  var from = settings.range ? settings.range.from : 1, to = settings.range ? Math.min(settings.range.to, facts.slideCount) : facts.slideCount;
  for (var n = from; n <= to; n++) if (settings.hidden || !hidden[n]) out.push(n);
  return out;
}

/* What the output will be, shown before anything is converted. */
function planOutput(facts, settings) {
  var s = settings, slides = expectedSlides(facts, s);
  var srcW = facts && facts.widthPt, srcH = facts && facts.heightPt, out = { pages: slides ? slides.length : null, slides: slides };
  if (!(srcW > 0 && srcH > 0)) return Object.assign(out, { known: false });
  if (s.output === 'notes') {
    /* a notes page has its own size, set in the deck. It is only known once PowerPoint has made it. */
    return Object.assign(out, { known: false, notes: true, w: null, h: null });
  }
  var t = targetSize(srcW, srcH, s.size);
  var tf = t.changed ? fitTransform(srcW, srcH, t.w, t.h, s.size.fit) : null;
  return Object.assign(out, { known: true, w: t.w, h: t.h, changed: t.changed, orientation: t.orientation, transform: tf,
    srcW: srcW, srcH: srcH, srcOrientation: orientationOf(srcW, srcH) });
}

/* ------------------------------------------------------------------ *
 * Output names                                                        *
 * ------------------------------------------------------------------ */

var RESERVED = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i;

function safeBase(name) {
  var b = String(name || '').replace(/\.[A-Za-z0-9]{1,5}$/, '');
  b = b.replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '_').replace(/[. ]+$/g, '').replace(/^\s+/, '');
  if (!b) b = 'presentation';
  if (RESERVED.test(b)) b = '_' + b;
  if (b.length > 180) b = b.slice(0, 180);
  return b;
}

/* items: [{ id, name, folder }]. existing: names already in the output folder.
   decisions: { id: 'overwrite' | 'rename' | 'skip' } for names that already exist.
   Returns { id: { outName, renamed, reason, exists, decision, blocked } } */
function planNames(items, existing, decisions) {
  var taken = Object.create(null), there = Object.create(null), out = {};
  (existing || []).forEach(function (n) { there[String(n).toLowerCase()] = true; });
  decisions = decisions || {};
  function free(base, suffixes) {
    var i, cand;
    for (i = 0; i < suffixes.length; i++) {
      cand = base + suffixes[i] + '.pdf';
      if (!taken[cand.toLowerCase()]) return { name: cand, suffix: suffixes[i] };
    }
    for (i = 2; i < 10000; i++) {
      cand = base + ' (' + i + ').pdf';
      if (!taken[cand.toLowerCase()]) return { name: cand, suffix: ' (' + i + ')' };
    }
    return { name: base + ' (' + Date.now() + ').pdf', suffix: '' };
  }
  items.forEach(function (it) {
    var base = safeBase(it.name), ext = (/\.([A-Za-z0-9]+)$/.exec(it.name) || [])[1] || '';
    var plain = base + '.pdf', r = { outName: plain, renamed: false, reason: '', exists: false, decision: null, blocked: false };
    if (taken[plain.toLowerCase()]) {
      var hints = [];
      if (it.folder) hints.push(' (' + safeBase(String(it.folder).split(/[\\/]/).filter(Boolean).pop() + '.x') + ')');
      if (ext) hints.push(' (' + ext.toLowerCase() + ')');
      var f = free(base, hints);
      r.outName = f.name; r.renamed = true;
      r.reason = 'Another file in this batch is also called ' + base + '. This one is saved as ' + f.name + '.';
    }
    if (there[r.outName.toLowerCase()]) {
      r.exists = true;
      var d = decisions[it.id] || null;
      r.decision = d;
      if (d === 'rename') {
        var was = r.outName, k = 2, cand;
        do { cand = safeBase(was) + ' (' + k++ + ').pdf'; } while (taken[cand.toLowerCase()] || there[cand.toLowerCase()]);
        r.outName = cand; r.renamed = true;
        r.reason = was + ' is already in the output folder. This one is saved as ' + cand + '.';
      } else if (d !== 'overwrite' && d !== 'skip') r.blocked = true;
    }
    if (r.decision !== 'skip') taken[r.outName.toLowerCase()] = true;
    out[it.id] = r;
  });
  return out;
}

/* ------------------------------------------------------------------ *
 * Issues                                                              *
 * ------------------------------------------------------------------ */

var SEVERITY = { error: 'Error', warning: 'Warning', info: 'Note' };
var CERTAINTY = { confirmed: 'Confirmed', suspected: 'Suspected', unverified: 'Not verified' };

function issue(code, o) {
  o = o || {};
  return {
    code: code,
    severity: o.severity || 'warning',
    certainty: o.certainty || 'confirmed',
    slide: o.slide == null ? null : o.slide,
    page: o.page == null ? null : o.page,
    description: o.description || '',
    impact: o.impact || '',
    fix: o.fix || '',
    review: !!o.review,
    stage: o.stage || 'check',
    accepted: null
  };
}

function numberIssues(list) {
  var seen = {};
  list.forEach(function (i) {
    var k = i.code + '@' + (i.slide == null ? 'deck' : i.slide);
    seen[k] = (seen[k] || 0) + 1;
    i.id = k + (seen[k] > 1 ? '#' + seen[k] : '');
  });
  return list;
}

function readErrorIssue(err) {
  var c = (err && err.code) || 'unreadable', m = (err && err.message) || 'The file could not be read.';
  if (c === 'password') return issue('password', { severity: 'error',
    description: 'The file is password protected.',
    impact: 'PowerPoint cannot open it without the password, so it will not be converted.',
    fix: 'Open it in PowerPoint, remove the password under File, Info, Protect Presentation, save it and add it again.' });
  if (c === 'empty') return issue('empty', { severity: 'error', description: 'The file is empty, 0 bytes.',
    impact: 'There is nothing to convert.', fix: 'Copy the file again from its source.' });
  if (c === 'not-zip' || c === 'zip-damaged' || c === 'cfb-damaged') return issue('corrupt', { severity: 'error',
    description: 'The file is damaged. ' + m,
    impact: 'PowerPoint is unlikely to open it. It is skipped.',
    fix: 'Copy the file again from its source. A download or a copy to a USB stick that stopped early looks like this.' });
  return issue('unsupported', { severity: 'error', description: m,
    impact: 'It is not a PowerPoint presentation, so it is skipped.',
    fix: 'Check the file opens in PowerPoint. Save it as .pptx from PowerPoint and add that.' });
}

/* Issues that can be found before PowerPoint opens the file. */
function preflightIssues(file, ins, settings, platform) {
  var out = [], s = settings;
  if (!ins.ok) return numberIssues([readErrorIssue(ins.error)]);

  if (file.size > SOURCE_WARN_BYTES) out.push(issue('too-large', { severity: 'warning', certainty: 'unverified',
    description: 'The file is ' + formatBytes(file.size) + ', above the 2 GB that has been tested.',
    impact: 'It may take a long time to copy and PowerPoint may run out of memory opening it.',
    fix: 'Convert it in a batch of its own, or compress the media in PowerPoint first.' }));
  if (ins.extMismatch) out.push(issue('ext-mismatch', { severity: 'warning',
    description: 'The file name ends in .' + (ins.kind === 'ppt' ? 'pptx' : 'ppt') + ' but the contents are ' + (ins.kind === 'ppt' ? 'the older .ppt format' : 'a .pptx') + '.',
    impact: 'PowerPoint may refuse it or ask a question when it opens it.',
    fix: 'Open it in PowerPoint and save it again as .pptx.' }));
  if (ins.kind === 'ppt') out.push(issue('legacy-format', { severity: 'info', certainty: 'unverified',
    description: 'This is a PowerPoint 97-2003 file. Its hidden slides, fonts, media and links cannot be read before conversion.',
    impact: 'Those checks run after PowerPoint has opened it and appear in the results.',
    fix: 'Nothing to do. Save as .pptx first if you want the full check up front.' }));
  if (ins.unread && ins.unread.length) out.push(issue('unread-parts', { severity: 'warning', certainty: 'unverified',
    description: plural(ins.unread.length, 'part') + ' of the file could not be read for checking.',
    impact: 'Some checks are incomplete for this file. PowerPoint may still convert it.',
    fix: 'Open the file in PowerPoint to see whether it reports damage.' }));
  if (ins.slideCount === 0) out.push(issue('no-slides', { severity: 'error',
    description: 'The presentation has no slides.', impact: 'There is nothing to convert.', fix: 'Check this is the right file.' }));
  if (ins.writeProtected) out.push(issue('write-protected', { severity: 'info',
    description: 'The file has a password to modify. It opens read only, which is all the conversion needs.',
    impact: 'None expected.', fix: '' }));
  if (ins.macro) out.push(issue('macro', { severity: 'info',
    description: 'The file contains macros. The helper opens it with macros disabled.', impact: 'None for a PDF.', fix: '' }));
  if (ins.kind !== 'pptx') return numberIssues(out);

  var hidden = ins.hiddenSlides || [];
  if (hidden.length) out.push(issue('hidden-slides', { severity: 'info',
    description: plural(hidden.length, 'hidden slide') + ' (' + listSlides(hidden) + ').',
    impact: s.hidden ? 'They are included in the PDF because Include hidden slides is on.' : 'They are left out of the PDF.',
    fix: s.hidden ? 'Turn Include hidden slides off to leave them out.' : 'Turn on Include hidden slides to keep them.' }));

  var remote = (ins.externalLinks || []).filter(function (l) { return l.remote; });
  var local = (ins.externalLinks || []).filter(function (l) { return !l.remote; });
  if (remote.length) {
    var rs = remote.map(function (l) { return l.slide; }).filter(function (v, i, a) { return a.indexOf(v) === i; });
    out.push(issue('remote-links', { severity: s.allowRemoteLinks ? 'warning' : 'error', stage: 'check',
      description: plural(remote.length, 'picture or media file') + ' on slide ' + listSlides(rs) + ' ' + (remote.length === 1 ? 'is' : 'are') + ' linked from the internet or a network share, not stored in the deck.',
      impact: s.allowRemoteLinks
        ? 'PowerPoint will fetch ' + (remote.length === 1 ? 'it' : 'them') + ' when it opens the deck, because Allow linked files from the network is on. What it fetches can differ from what the author saw.'
        : 'PowerPoint fetches linked files as soon as it opens a deck. This deck is held back so that nothing is fetched without your say.',
      fix: s.allowRemoteLinks ? 'Embed the pictures in PowerPoint to make the deck self contained.'
        : 'Turn on Allow linked files from the network to convert it, or embed the pictures in PowerPoint and add the deck again.' }));
  }
  if (local.length) {
    var ls = local.map(function (l) { return l.slide; }).filter(function (v, i, a) { return a.indexOf(v) === i; });
    out.push(issue('local-links', { severity: 'warning', certainty: 'unverified',
      description: plural(local.length, 'picture or media file') + ' on slide ' + listSlides(ls) + ' ' + (local.length === 1 ? 'is' : 'are') + ' linked to a file path, not stored in the deck.',
      impact: 'If the path does not exist on this computer PowerPoint shows its last saved copy or an empty box.',
      fix: 'Check those slides in the PDF. Embedding the files in PowerPoint removes the doubt.' }));
  }

  (ins.slides || []).forEach(function (sl) {
    if (!sl || sl.unreadable) return;
    (sl.media || []).forEach(function (m) {
      if (m.type === 'video') {
        var poster = m.posterBlank ? 'blank' : m.posterPart ? 'ok' : 'none';
        if (poster === 'ok') out.push(issue('video', { slide: sl.n, severity: 'warning',
          description: 'Video "' + m.name + '"' + (m.linked ? ', linked not embedded' : '') + '. The PDF shows its poster frame at the same position and size.',
          impact: 'The video does not play in a PDF. Anyone reading it sees a still picture.',
          fix: 'Send the video file alongside the PDF if it matters.' }));
        else out.push(issue('video-no-poster', { slide: sl.n, severity: 'warning', review: true,
          certainty: poster === 'blank' ? 'suspected' : 'confirmed',
          description: 'Video "' + m.name + '" ' + (poster === 'blank' ? 'has a poster frame that is a single flat colour, most often a black first frame.' : 'has no poster frame stored in the deck.'),
          impact: s.placeholders && m.box ? 'A clearly marked placeholder is drawn over it in the PDF so the gap is obvious. The video does not play in a PDF.'
            : m.box ? 'The PDF shows the blank frame. Placeholders are turned off.'
            : 'The video is inside a group, so the helper cannot place a placeholder over it. The PDF shows whatever PowerPoint draws.',
          fix: 'In PowerPoint choose Video Format, Poster Frame, and pick a frame or a picture. Then add the deck again.' }));
      } else {
        out.push(issue('audio', { slide: sl.n, severity: 'info',
          description: 'Audio "' + m.name + '"' + (m.linked ? ', linked not embedded' : '') + '.',
          impact: 'A PDF has no sound. The speaker icon is drawn if it is visible on the slide.', fix: '' }));
      }
    });
    if (sl.anim && sl.anim.total) {
      var risky = sl.anim.overlapping > 0 || sl.anim.exit > 0;
      if (risky) out.push(issue('animation-overlap', { slide: sl.n, severity: 'warning', certainty: 'suspected', review: true,
        description: plural(sl.anim.total, 'animation') + ', including ' +
          (sl.anim.exit ? plural(sl.anim.exit, 'exit effect') : '') + (sl.anim.exit && sl.anim.overlapping ? ' and ' : '') +
          (sl.anim.overlapping ? plural(sl.anim.overlapping, 'pair') + ' of animated shapes that overlap' : '') + '.',
        impact: 'The PDF shows every shape at once, as the slide looks in the editor. Content meant to appear one after another, or to leave, is stacked and may hide or contradict what is beneath it.',
        fix: 'Check this slide in the PDF. If it reads wrongly, split the builds across separate slides in PowerPoint.' }));
      else out.push(issue('animation', { slide: sl.n, severity: 'info',
        description: plural(sl.anim.total, 'animation') + '.',
        impact: 'The PDF shows the slide with everything visible, as it looks in the editor.', fix: '' }));
    }
    if (sl.transition && /morph/i.test(sl.transition.type)) out.push(issue('morph', { slide: sl.n, severity: 'info',
      description: 'Morph transition.', impact: 'The movement between this slide and the one before is not shown. Each slide is a separate still page.', fix: '' }));
    if (sl.controls) out.push(issue('activex', { slide: sl.n, severity: 'warning', certainty: 'unverified',
      description: plural(sl.controls, 'ActiveX control') + '.',
      impact: 'Controls are drawn as PowerPoint last saw them, or not at all on macOS.', fix: 'Check this slide in the PDF.' }));
    if (sl.model3d) out.push(issue('model3d', { slide: sl.n, severity: 'info',
      description: plural(sl.model3d, '3D model') + '.', impact: 'The PDF shows the model from its saved angle as a flat picture.', fix: '' }));
    if (sl.ole) out.push(issue('ole', { slide: sl.n, severity: 'info',
      description: plural(sl.ole, 'embedded object') + ' such as a worksheet or document.',
      impact: 'The PDF shows the object\'s saved picture. Its contents cannot be opened from the PDF.', fix: '' }));
  });
  var trans = (ins.counts && ins.counts.transitions) || 0;
  if (trans) out.push(issue('transitions', { severity: 'info',
    description: plural(trans, 'slide') + ' with a transition.', impact: 'Transitions are not shown in a PDF.', fix: '' }));
  if (s.output !== 'notes' && ins.counts && ins.counts.notes) out.push(issue('notes-present', { severity: 'info',
    description: plural(ins.counts.notes, 'slide') + ' with speaker notes.',
    impact: 'Notes are not in the PDF because the output is slides only.',
    fix: featureAvailable('notes', platform).ok ? 'Choose Notes pages to include them.' : '' }));
  return numberIssues(out);
}

/* ------------------------------------------------------------------ *
 * Fonts                                                               *
 * ------------------------------------------------------------------ */

function normFont(name) {
  return String(name || '').replace(/^[A-Z]{6}\+/, '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function fontMatches(deckName, pdfName) {
  var d = normFont(deckName), p = normFont(pdfName);
  if (!d || !p) return false;
  return p === d || p.indexOf(d) === 0 || (d.indexOf(p) === 0 && p.length >= 4);
}

/* deckFonts: names the deck uses. installed: family names on this computer, or null when unknown.
   deckEmbedded: fonts stored inside the deck. pptFonts: PowerPoint's own list [{name, embeddable, embedded}] or null.
   pdfFonts: [{name, embedded, type3}] read back from the PDF, or null when the PDF was not inspected. */
function fontFindings(deckFonts, installed, deckEmbedded, pptFonts, pdfFonts, settings, platform, namedOnSlides) {
  var out = [], inst = null, emb = {}, ppt = {}, onSlides = null;
  /* a font set only on a master or in the theme is often never used, so its absence from the PDF says nothing */
  if (namedOnSlides) { onSlides = {}; namedOnSlides.forEach(function (n) { onSlides[normFont(n)] = true; }); }
  if (installed) { inst = {}; installed.forEach(function (n) { inst[normFont(n)] = true; }); }
  (deckEmbedded || []).forEach(function (n) { emb[normFont(n)] = true; });
  (pptFonts || []).forEach(function (f) { ppt[normFont(f.name)] = f; });
  function isInstalled(name) {
    if (!inst) return null;
    var n = normFont(name);
    if (inst[n]) return true;
    for (var k in inst) if (k.indexOf(n) === 0 || (n.indexOf(k) === 0 && k.length >= 4)) return true;
    return false;
  }
  var names = (deckFonts || []).slice();
  (pptFonts || []).forEach(function (f) { if (!names.some(function (n) { return normFont(n) === normFont(f.name); })) names.push(f.name); });

  var matchedPdf = {}, missing = [], restricted = [], notEmbedded = [], absent = [];
  names.forEach(function (name) {
    var inPdf = pdfFonts ? pdfFonts.filter(function (p) { return fontMatches(name, p.name); }) : null;
    if (inPdf && inPdf.length) {
      inPdf.forEach(function (p) { matchedPdf[p.name] = true; });
      if (inPdf.some(function (p) { return !p.embedded && !p.type3; })) notEmbedded.push(name);
      return;
    }
    var have = isInstalled(name), inDeck = !!emb[normFont(name)], pf = ppt[normFont(name)];
    if (have === false && !(inDeck && platform === 'windows')) missing.push(name);
    else if (pf && pf.embeddable === false) restricted.push(name);
    else if (pdfFonts && (!onSlides || onSlides[normFont(name)] || pf)) absent.push(name);
  });
  var extras = pdfFonts ? pdfFonts.filter(function (p) { return !matchedPdf[p.name]; }).map(function (p) { return p.name.replace(/^[A-Z]{6}\+/, ''); })
    .filter(function (v, i, a) { return a.indexOf(v) === i; }) : [];

  if (missing.length) out.push(issue('font-missing', { severity: 'warning', stage: 'convert', review: true,
    certainty: pdfFonts ? 'confirmed' : 'suspected',
    description: (missing.length === 1 ? 'The font ' : 'The fonts ') + missing.join(', ') + (missing.length === 1 ? ' is' : ' are') +
      ' used by the deck but not installed on this computer' + (pdfFonts ? ' and ' + (missing.length === 1 ? 'does' : 'do') + ' not appear in the PDF.' : '.'),
    impact: 'PowerPoint drew that text in a substitute font, so line breaks, spacing and the look of the text can differ from the original. Turning text into a bitmap does not help. A bitmap of the wrong font is still the wrong font.' +
      (extras.length ? ' Fonts in the PDF that the deck does not name, and so the likely substitutes, are ' + extras.slice(0, 6).join(', ') + '.' : ''),
    fix: 'Install the font on this computer and convert the deck again, or ask the author for a copy saved with the fonts embedded.' }));
  if (restricted.length) out.push(issue('font-restricted', { severity: 'warning', stage: 'convert',
    description: (restricted.length === 1 ? 'The font ' : 'The fonts ') + restricted.join(', ') + (restricted.length === 1 ? ' is' : ' are') +
      ' installed, but the font licence does not allow embedding, as reported by PowerPoint.',
    impact: settings.bitmapText ? 'Text in ' + (restricted.length === 1 ? 'that font' : 'those fonts') + ' was written into the PDF as a picture. It looks right but cannot be searched, selected or read by a screen reader.'
      : 'Text in ' + (restricted.length === 1 ? 'that font' : 'those fonts') + ' stays as text without the font inside the PDF. It only displays correctly on a computer that has the font.',
    fix: 'Use a font that allows embedding, or accept the trade. The font is present, so this is a licence limit and not a missing font.' }));
  if (notEmbedded.length) out.push(issue('font-not-embedded', { severity: 'warning', stage: 'verify',
    description: (notEmbedded.length === 1 ? 'The font ' : 'The fonts ') + notEmbedded.join(', ') + (notEmbedded.length === 1 ? ' is' : ' are') +
      ' named in the PDF but not embedded in it.',
    impact: 'The PDF only displays that text correctly on a computer that has the font. Elsewhere a substitute is drawn.',
    fix: 'Turn on Bitmap text when fonts cannot be embedded, or use a font that allows embedding.' }));
  if (absent.length) out.push(issue('font-absent', { severity: 'info', stage: 'verify', certainty: 'unverified',
    description: (absent.length === 1 ? 'The font ' : 'The fonts ') + absent.join(', ') + (absent.length === 1 ? ' is' : ' are') +
      ' named in the deck and installed, but not found in the PDF.',
    impact: 'Either no exported slide uses ' + (absent.length === 1 ? 'it' : 'them') + ', which is common for fonts set on a master, or the text was written as a picture. This check cannot tell which.',
    fix: 'If text in that font matters, zoom in on it in the PDF. Text that will not select has been turned into a picture.' }));
  return { issues: out, missing: missing, restricted: restricted, notEmbedded: notEmbedded, absent: absent, substitutes: missing.length ? extras : [] };
}

/* ------------------------------------------------------------------ *
 * Job ticket                                                          *
 * ------------------------------------------------------------------ */

function ticketId(n) { return 'f' + ('0000' + n).slice(-5); }

function placeholdersFor(ins, settings) {
  if (!settings.placeholders || !ins || !ins.slides) return [];
  var out = [];
  ins.slides.forEach(function (sl) {
    (sl && sl.media || []).forEach(function (m) {
      if (m.type !== 'video' || !m.box) return;
      if (m.posterPart && !m.posterBlank) return;
      out.push({ slide: sl.n, left: +m.box.left.toFixed(2), top: +m.box.top.toFixed(2), width: +m.box.width.toFixed(2), height: +m.box.height.toFixed(2),
                 label: 'VIDEO  ' + (m.name || 'no name') + '\nNo poster frame. This is a placeholder added by SlideSize.' });
    });
  });
  return out;
}

/* item: { id, name, ext, inspect, outName, overwrite }. Paths are relative to the work folder. */
function buildTicket(item, settings, opts) {
  opts = opts || {};
  var s = settings, ins = item.inspect || {};
  var resize = s.size.preset !== 'original' || s.links === 'remove' || s.metadata === 'strip';
  return {
    protocol: PROTOCOL, id: item.id, sourceName: item.name,
    input: 'in/slidesize-' + item.id + '.' + (item.ext || 'pptx'),
    /* a file the page still has to edit is written inside the work folder first */
    pdf: resize ? 'out/' + item.id + '.pdf' : '../' + item.outName,
    overwrite: resize ? true : !!item.overwrite,
    mayPrompt: !!(ins.writeProtected) || ins.kind === 'ppt',
    timeoutSec: s.timeoutSec,
    options: {
      intent: s.quality === 'minimum' ? 'screen' : 'print',
      includeHidden: !!s.hidden,
      range: s.range ? [s.range.from, s.range.to] : null,
      output: s.output,
      bitmapText: !!s.bitmapText, tags: !!s.tags, docProps: !!s.docProps, markup: !!s.markup, pdfa: !!s.pdfa,
      reference: s.fidelity !== 'off' && s.output === 'slides'
        ? { mode: s.fidelity, longEdge: s.referenceLongEdge, slides: null } : null,
      placeholders: placeholdersFor(ins, s),
      validatePdfa: !!s.pdfa
    },
    /* PDF/A is validated on the finished file. When the page still has edits to make, it asks again afterwards. */
    validateNow: !!s.pdfa && !resize,
    collect: { media: ins.kind === 'ppt' },
    created: opts.now || new Date().toISOString()
  };
}

/* Up to five slides spread through the deck, first and last included. */
function sampleSlides(slides, max) {
  max = max || 5;
  if (!slides || slides.length <= max) return (slides || []).slice();
  var out = [];
  for (var i = 0; i < max; i++) out.push(slides[Math.round(i * (slides.length - 1) / (max - 1))]);
  return out.filter(function (v, i, a) { return a.indexOf(v) === i; });
}

/* ------------------------------------------------------------------ *
 * Helper state                                                        *
 * ------------------------------------------------------------------ */

/* h is the parsed helper.json or null. Returns { state, text, level } for the setup panel. */
function helperStatus(h, now, expectPlatform) {
  if (!h) return { state: 'absent', level: 'wait', text: 'The helper is not running yet.' };
  var age = now - Date.parse(h.heartbeat || 0);
  if (h.protocol !== PROTOCOL) return { state: 'mismatch', level: 'bad',
    text: 'The helper in this folder is a different version from this page. Write the helper files again and restart it.' };
  if (h.state === 'stopped') return { state: 'stopped', level: 'wait', text: 'The helper has been stopped. Start it again to continue.' };
  if (!(age < HEARTBEAT_STALE_MS)) return { state: 'stale', level: 'wait',
    text: 'The helper was last heard from ' + formatDuration(Math.max(0, age)) + ' ago. Its window may have been closed. Start it again.' };
  if (!h.powerpoint || !h.powerpoint.installed) return { state: 'no-powerpoint', level: 'bad',
    text: 'The helper is running but cannot find Microsoft PowerPoint on this computer. Microsoft PowerPoint must be installed for PowerPoint-based conversion.' };
  if (h.state === 'blocked') return { state: 'blocked', level: 'bad', text: h.message || 'PowerPoint is not responding to the helper.' };
  if (h.state === 'starting') return { state: 'starting', level: 'wait', text: h.message || 'The helper is starting PowerPoint.' };
  return { state: 'ready', level: 'ok',
    text: 'Connected to PowerPoint ' + (h.powerpoint.version || '') + ' on ' + platformName(h.platform) +
      (h.powerpoint.userPresentations ? '. ' + plural(h.powerpoint.userPresentations, 'presentation') + ' of yours ' + (h.powerpoint.userPresentations === 1 ? 'is' : 'are') + ' open and will be left alone' : '') + '.' };
}

/* ------------------------------------------------------------------ *
 * Results                                                             *
 * ------------------------------------------------------------------ */

var STATES = {
  completed: 'Completed',
  warnings: 'Completed with warnings',
  review: 'Needs review',
  failed: 'Failed'
};

function helperFailureIssue(res) {
  var e = (res && res.error) || {}, c = e.code || res.status || 'failed';
  var msg = e.message ? ' PowerPoint said: ' + e.message : '';
  if (c === 'timeout' || res.status === 'timeout') return issue('timeout', { severity: 'error', stage: 'convert',
    description: 'PowerPoint did not finish this file within the time limit of ' + formatDuration((res.timeoutSec || 0) * 1000) + '.',
    impact: 'No PDF was made. The rest of the batch carried on.',
    fix: 'Retry it with a longer time limit under Advanced. If it stalls again, open it in PowerPoint to see whether it asks a question or reports damage.' });
  if (c === 'cancelled' || res.status === 'cancelled') return issue('cancelled', { severity: 'error', stage: 'convert',
    description: 'This file was skipped while it was converting.', impact: 'No PDF was made.', fix: 'Retry it when ready.' });
  if (c === 'password') return issue('password', { severity: 'error', stage: 'convert',
    description: 'PowerPoint could not open the file because it is password protected.' + msg,
    impact: 'No PDF was made.', fix: 'Remove the password in PowerPoint and add the file again.' });
  if (c === 'exists') return issue('exists', { severity: 'error', stage: 'convert',
    description: 'A file with the output name appeared in the folder after the batch started.',
    impact: 'Nothing was overwritten and no PDF was made.', fix: 'Retry and choose whether to overwrite or keep both.' });
  if (c === 'open-failed') return issue('open-failed', { severity: 'error', stage: 'convert',
    description: 'PowerPoint could not open the file.' + msg,
    impact: 'No PDF was made.', fix: 'Open the file in PowerPoint by hand. If it offers to repair it, let it, save the result and add that.' });
  return issue('convert-failed', { severity: 'error', stage: 'convert',
    description: 'PowerPoint opened the file but the export failed.' + msg,
    impact: 'No PDF was made.', fix: 'Retry it. If it fails again, try without PDF/A or with the fidelity check off, to find which step fails.' });
}

/* Issues that come from comparing what was asked for with what PowerPoint reported and what the PDF contains.
   ctx: { settings, platform, plan, result (from the helper), pdf (inspectPdf output or null), pdfError, transform, fonts, installedFonts } */
function resultIssues(item, ctx) {
  var out = [], s = ctx.settings, res = ctx.result || {}, pdf = ctx.pdf, ins = item.inspect || {};
  if (res.engine && res.engine.name !== ENGINE_NAME) {
    out.push(issue('engine-unknown', { severity: 'error', stage: 'convert',
      description: 'The helper reports that this file was converted by "' + (res.engine.name || 'an unnamed program') + '", not by Microsoft PowerPoint.',
      impact: 'The result is not treated as a PowerPoint conversion and is marked as failed.',
      fix: 'Write the helper files again from this page and restart the helper.' }));
    return numberIssues(out);
  }
  var facts = res.facts || {}, exp = res.export || {};

  if (ins.ok && ins.slideCount != null && facts.slides != null && ins.slideCount !== facts.slides) out.push(issue('slide-count-differs', {
    severity: 'warning', stage: 'convert',
    description: 'PowerPoint counts ' + plural(facts.slides, 'slide') + '. The check before conversion counted ' + ins.slideCount + '.',
    impact: 'PowerPoint may have repaired or dropped part of the file when it opened it.',
    fix: 'Open the file in PowerPoint and compare the slide count with what the author expects.' }));

  if (exp.method && exp.method !== 'ExportAsFixedFormat2' && exp.method !== 'ExportAsFixedFormat' && exp.method !== 'save as PDF') out.push(issue('fallback-export', {
    severity: 'warning', stage: 'convert',
    description: 'PowerPoint\'s detailed export failed, so the helper used ' + exp.method + '.' + (exp.fallbackReason ? ' The error was: ' + exp.fallbackReason : ''),
    impact: 'Only PowerPoint\'s default export settings were applied to this file.', fix: 'Retry it. If it happens every time, update PowerPoint.' }));
  (exp.notApplied || []).forEach(function (n) {
    out.push(issue('option-not-applied', { severity: 'warning', stage: 'convert',
      description: 'The setting "' + ((FEATURES[n.option] && FEATURES[n.option].label) || n.option) + '" was not applied. ' + (n.reason || ''),
      impact: 'The PDF was made without it.', fix: 'Check whether the PDF is acceptable without it.' }));
  });
  if (exp.removedSlides && exp.removedSlides.length && (ins.slides || []).some(function (sl) { return sl && sl.slideNumberField; })) out.push(issue('renumbered', {
    severity: 'warning', stage: 'convert',
    description: 'To leave out slide ' + listSlides(exp.removedSlides) + ' the helper removed ' + (exp.removedSlides.length === 1 ? 'it' : 'them') + ' from its temporary copy, because PowerPoint for Mac would not skip ' + (exp.removedSlides.length === 1 ? 'it' : 'them') + '.',
    impact: 'Slide numbers printed on the slides after the first removed one are lower in the PDF than in the deck.',
    fix: 'Convert on a Windows PC to keep the printed numbers, or include all slides.' }));
  (exp.placeholders || []).forEach(function (p) {
    if (p.ok === false) out.push(issue('placeholder-failed', { slide: p.slide, severity: 'warning', stage: 'convert', review: true,
      description: 'The video placeholder could not be added to this slide. ' + (p.error || ''),
      impact: 'The PDF shows the video\'s blank poster frame with nothing to mark it.', fix: 'Set a poster frame in PowerPoint.' }));
  });
  (res.pathChecks || []).forEach(function (pc) {
    if (pc.exists === false) out.push(issue('link-missing', { slide: pc.slide, severity: 'warning', stage: 'convert', review: true,
      description: 'A linked file was not found on this computer: ' + pc.target,
      impact: 'PowerPoint drew its last saved copy of the picture if the deck holds one, or an empty box if not.',
      fix: 'Check this slide in the PDF. Embed the file in PowerPoint to remove the doubt.' }));
  });

  /* facts PowerPoint knew and the pre-check did not (legacy .ppt) */
  if (ins.kind === 'ppt' && facts.hidden && facts.hidden.length) out.push(issue('hidden-slides', { severity: 'info', stage: 'convert',
    description: plural(facts.hidden.length, 'hidden slide') + ' (' + listSlides(facts.hidden) + '), reported by PowerPoint.',
    impact: s.hidden ? 'They are included in the PDF.' : 'They are left out of the PDF.', fix: '' }));
  if (ins.kind === 'ppt' && facts.media && facts.media.length) {
    facts.media.forEach(function (m) {
      out.push(issue(m.type === 'audio' ? 'audio' : 'video', { slide: m.slide, severity: m.type === 'audio' ? 'info' : 'warning', stage: 'convert',
        description: (m.type === 'audio' ? 'Audio' : 'Video') + ' "' + (m.name || '') + '", reported by PowerPoint.',
        impact: m.type === 'audio' ? 'A PDF has no sound.' : 'The video does not play in a PDF. The page shows its poster frame.', fix: '' }));
    });
  }

  if (ctx.pdfError) {
    out.push(issue('pdf-unreadable', { severity: 'error', stage: 'verify',
      description: 'PowerPoint reported success but the PDF could not be opened to check it. ' + ctx.pdfError,
      impact: 'The file may be damaged. It is kept so you can try it in a PDF reader.',
      fix: 'Open the PDF. If it is damaged, retry the conversion.' }));
    return numberIssues(out);
  }
  if (ctx.pdfSkipped) {
    out.push(issue('verify-skipped', { severity: 'warning', certainty: 'unverified', stage: 'verify',
      description: 'The PDF is ' + formatBytes(ctx.pdfSkipped) + ', too large to open in the browser for checking. Page count, page size, fonts and PDF/A were not checked.',
      impact: 'The file is exactly as PowerPoint wrote it. Nothing about it has been verified here.',
      fix: 'Open it in a PDF reader and check the page count yourself.' }));
    if (ctx.editsSkipped && ctx.editsSkipped.length) out.push(issue('edits-skipped', { severity: 'error', stage: 'verify',
      description: 'The PDF is too large to edit in the browser, so these settings were not applied: ' + ctx.editsSkipped.join(', ') + '.',
      impact: 'The PDF has PowerPoint\'s original pages.', fix: 'Reduce the image quality or split the deck, then convert again.' }));
    return numberIssues(out);
  }
  if (!pdf) return numberIssues(out);

  /* page count */
  var want = exp.pageMap ? exp.pageMap.length : (ctx.plan && ctx.plan.pages);
  if (want != null && pdf.pageCount !== want) out.push(issue('page-count', { severity: 'error', stage: 'verify',
    description: 'The PDF has ' + plural(pdf.pageCount, 'page') + ' where ' + want + ' ' + (want === 1 ? 'was' : 'were') + ' expected from the slides and settings.',
    impact: 'Slides are missing or extra. Page numbers in this report may not line up with slide numbers.',
    fix: 'Open the PDF and the deck side by side and count. Retry the conversion if slides are missing.' }));
  else if (want == null) out.push(issue('page-count-unverified', { severity: 'info', certainty: 'unverified', stage: 'verify',
    description: 'The PDF has ' + plural(pdf.pageCount, 'page') + '. The expected count was not known, so it was not compared.', impact: '', fix: '' }));

  /* page size */
  var tgt = ctx.plan && ctx.plan.known ? { w: ctx.plan.w, h: ctx.plan.h } : null;
  if (s.output === 'slides' && tgt && pdf.pages.length) {
    var off = [], tol = 0.75;   /* PowerPoint rounds page boxes to the nearest point or so */
    pdf.pages.forEach(function (p, i) { if (Math.abs(p.w - tgt.w) > tol || Math.abs(p.h - tgt.h) > tol) off.push(i + 1); });
    if (off.length) out.push(issue('page-size', { page: off[0], severity: 'error', stage: 'verify',
      description: plural(off.length, 'page') + ' (' + listSlides(off) + ') ' + (off.length === 1 ? 'is' : 'are') + ' not the expected ' + sizeLabel(tgt.w, tgt.h) +
        '. Page ' + off[0] + ' is ' + sizeLabel(pdf.pages[off[0] - 1].w, pdf.pages[off[0] - 1].h) + '.',
      impact: 'The output does not match the size that was chosen.', fix: 'Retry the conversion. If it repeats, report it with this deck.' }));
  }
  if (ctx.transform && ctx.transform.cropped) out.push(issue('cropped', { severity: 'warning', stage: 'verify', review: true,
    description: 'Fill crops ' + (ctx.transform.cropX ? trimNum(ctx.transform.cropX * 100, 1) + '% from the left and from the right' : '') +
      (ctx.transform.cropX && ctx.transform.cropY ? ' and ' : '') +
      (ctx.transform.cropY ? trimNum(ctx.transform.cropY * 100, 1) + '% from the top and from the bottom' : '') + ' of every slide.',
    impact: 'Anything in those strips is off the page. It is still inside the file, hidden beyond the page edge.',
    fix: 'Choose Fit with margins to keep the whole slide.' }));

  /* fonts */
  if (ctx.fonts) ctx.fonts.issues.forEach(function (i) { out.push(i); });

  /* what the edits did */
  if (ctx.edits) {
    if (ctx.edits.linksRemoved != null) out.push(issue('links-removed', { severity: 'info', stage: 'verify',
      description: plural(ctx.edits.linksRemoved, 'link') + ' removed from the PDF.', impact: '', fix: '' }));
    if (ctx.edits.metadataStripped) out.push(issue('metadata-stripped', { severity: 'info', stage: 'verify',
      description: 'Title, author, subject, keywords and the XMP metadata were removed from the PDF.', impact: '', fix: '' }));
  }
  if (!s.tags && pdf.tagged) out.push(issue('tags-present', { severity: 'info', stage: 'verify',
    description: 'Accessibility tags were turned off but the PDF still contains a structure tree.', impact: '', fix: '' }));
  if (s.tags && pdf.tagged === false && ctx.platform === 'windows') out.push(issue('tags-missing', { severity: 'warning', stage: 'verify',
    description: 'Accessibility tags were requested but the PDF has no structure tree.',
    impact: 'A screen reader has no reading order to follow.', fix: 'Retry the conversion.' }));
  return numberIssues(out);
}

function pdfaIssues(pdfa) {
  if (!pdfa) return [];
  var out = [], claim = pdfa.claim ? 'PDF/A-' + pdfa.claim.part + String(pdfa.claim.conformance || '').toLowerCase() : null;
  if (pdfa.result === 'passed') out.push(issue('pdfa-passed', { severity: 'info', stage: 'verify',
    description: 'PDF/A validation passed. ' + (pdfa.validator || 'The validator') + ' checked the file against ' + (pdfa.profile || claim) + '.', impact: '', fix: '' }));
  else if (pdfa.result === 'failed') {
    var fails = (pdfa.checks || []).filter(function (c) { return c.result === 'fail'; });
    out.push(issue('pdfa-failed', { severity: 'error', stage: 'verify',
      description: 'PDF/A validation failed. ' + (pdfa.validator
        ? pdfa.validator + ' reports the file does not conform to ' + (pdfa.profile || claim || 'PDF/A') + '.' + (pdfa.detail ? ' ' + pdfa.detail : '')
        : fails.length ? fails.map(function (c) { return c.label + ': ' + c.detail; }).join(' ') : 'The file does not declare itself as PDF/A.'),
      impact: 'The file must not be described as PDF/A. It is still a usable PDF.',
      fix: 'Convert again with bitmap text and accessibility tags on and without resizing or link removal. If it still fails, the deck contains something PDF/A does not allow.' }));
  } else out.push(issue('pdfa-unverified', { severity: 'warning', certainty: 'unverified', stage: 'verify',
    description: 'PDF/A was requested and the file declares itself as ' + (claim || 'PDF/A') + ', but full validation was not run. ' +
      (pdfa.reason || 'A complete check needs veraPDF, a free validator, installed on this computer.') +
      ' The ' + ((pdfa.checks || []).length) + ' structural checks SlideSize runs itself found no problem.',
    impact: 'The file probably conforms. That is not the same as verified. Do not state that it is validated PDF/A on the strength of this.',
    fix: 'Install veraPDF from verapdf.org and convert again, or run the file through the validator your archive uses.' }));
  return numberIssues(out);
}

function compareIssues(cmp, settings) {
  if (!cmp) return [];
  var out = [];
  if (cmp.performed === false) {
    if (settings.fidelity !== 'off') out.push(issue('compare-not-performed', { severity: 'info', certainty: 'unverified', stage: 'verify',
      description: 'The picture comparison was not run. ' + (cmp.reason || ''),
      impact: 'Visual fidelity for this file has not been checked by SlideSize.', fix: 'Look through the PDF yourself.' }));
    return numberIssues(out);
  }
  (cmp.pages || []).forEach(function (p) {
    if (p.verdict !== 'different') return;
    out.push(issue('visual-diff', { slide: p.slide, page: p.page, severity: 'warning', certainty: 'suspected', review: true, stage: 'verify',
      description: 'The PDF page differs from PowerPoint\'s own picture of the slide in ' + trimNum(p.changedShare * 100, 1) + '% of the compared area.',
      impact: 'Something may have moved, changed font, lost an effect or gone missing. It can also be a harmless difference in how two programs smooth edges, gradients and shadows.',
      fix: 'Open the side by side view for this slide. Accept it if the difference does not matter.' }));
  });
  (cmp.skipped || []).forEach(function (p) {
    out.push(issue('compare-skipped', { slide: p.slide, page: p.page, severity: 'info', certainty: 'unverified', stage: 'verify',
      description: 'This slide was not compared. ' + (p.reason || ''), impact: '', fix: '' }));
  });
  var n = (cmp.pages || []).length, total = cmp.totalPages || n;
  if (n && n < total) out.push(issue('compare-sampled', { severity: 'info', certainty: 'unverified', stage: 'verify',
    description: plural(n, 'page') + ' of ' + total + ' compared. The others were not checked.', impact: '', fix: 'Choose Every slide for a full comparison.' }));
  return numberIssues(out);
}

function mergeIssues(lists) {
  var out = [];
  lists.forEach(function (l) { (l || []).forEach(function (i) { out.push(i); }); });
  return numberIssues(out);
}

/* The four result states. A PDF that exists can always be opened and saved, whatever the state. */
function classify(item) {
  if (item.status === 'failed' || item.status === 'cancelled' || (item.status === 'done' && !item.hasPdf)) return 'failed';
  var issues = item.issues || [], open = issues.filter(function (i) { return !i.accepted; });
  if (open.some(function (i) { return i.severity === 'error' || i.review; })) return 'review';
  if (issues.some(function (i) { return i.severity === 'warning' || i.severity === 'error' || i.review; })) return 'warnings';
  return 'completed';
}

function summarize(items) {
  var out = { total: items.length, completed: 0, warnings: 0, review: 0, failed: 0, pending: 0, skipped: 0, blocked: 0,
              errors: 0, warningsCount: 0, notes: 0, pages: 0, ms: 0, bytes: 0 };
  items.forEach(function (it) {
    if (it.status === 'skipped') { out.skipped++; return; }
    if (it.status === 'blocked') { out.blocked++; return; }
    if (it.status !== 'done' && it.status !== 'failed' && it.status !== 'cancelled') { out.pending++; return; }
    out[classify(it)]++;
    (it.issues || []).forEach(function (i) { if (i.severity === 'error') out.errors++; else if (i.severity === 'warning') out.warningsCount++; else out.notes++; });
    out.pages += (it.checks && it.checks.pageCount) || 0;
    out.ms += (it.timing && it.timing.totalMs) || 0;
    out.bytes += it.pdfBytes || 0;
  });
  out.finished = out.completed + out.warnings + out.review + out.failed;
  return out;
}

/* ------------------------------------------------------------------ *
 * Batch record, saved after every file                                *
 * ------------------------------------------------------------------ */

var ACTIVE = { staging: 1, queued: 1, converting: 1, verifying: 1 };

function newBatch(o) {
  o = o || {};
  var now = o.now || new Date().toISOString();
  return { version: 1, id: o.id || 'b' + Date.parse(now).toString(36), createdAt: now, updatedAt: now,
           platform: o.platform || 'other', settings: mergeSettings(DEFAULTS, o.settings), preset: o.preset || 'standard',
           helper: null, items: [], nextId: 1 };
}

function addItem(batch, f) {
  var id = ticketId(batch.nextId++);
  var item = { id: id, name: f.name, folder: f.folder || '', size: f.size, lastModified: f.lastModified || 0,
               ext: (/\.([A-Za-z0-9]+)$/.exec(f.name) || [])[1] ? (/\.([A-Za-z0-9]+)$/.exec(f.name))[1].toLowerCase() : 'pptx',
               status: 'new', inspect: null, overrides: null, outName: null, overwrite: false, nameDecision: null,
               issues: [], result: null, checks: null, timing: null, attempts: 0, hasPdf: false, pdfBytes: 0 };
  batch.items.push(item);
  return item;
}

function sameFile(item, f) {
  return item.name === f.name && item.size === f.size && (!item.lastModified || !f.lastModified || item.lastModified === f.lastModified);
}

/* Only what is needed to carry on later is kept. The slide by slide detail is rebuilt by checking the file again. */
function slimInspect(ins) {
  if (!ins) return null;
  if (!ins.ok) return { ok: false, kind: ins.kind, error: ins.error };
  return {
    ok: true, kind: ins.kind, slideCount: ins.slideCount, widthPt: ins.widthPt, heightPt: ins.heightPt,
    hiddenSlides: ins.hiddenSlides, fonts: ins.fonts, slideFonts: ins.slideFonts || null, embeddedFonts: ins.embeddedFonts, counts: ins.counts,
    externalLinks: ins.externalLinks, writeProtected: ins.writeProtected, macro: ins.macro, extMismatch: ins.extMismatch,
    thumbnailPart: ins.thumbnailPart, unread: ins.unread,
    slides: ins.slides ? ins.slides.map(function (sl) {
      if (!sl || sl.unreadable) return { n: sl && sl.n, unreadable: true };
      return { n: sl.n, hidden: sl.hidden, slideNumberField: sl.slideNumberField, hasNotes: sl.hasNotes,
               media: (sl.media || []).map(function (m) { return { type: m.type, name: m.name, linked: m.linked, box: m.box, posterPart: m.posterPart, posterBlank: !!m.posterBlank, inGroup: m.inGroup }; }),
               anim: sl.anim && sl.anim.total ? sl.anim : null, transition: sl.transition, ole: sl.ole, controls: sl.controls, model3d: sl.model3d };
    }) : null
  };
}

function serializeBatch(batch, now) {
  var b = clone(batch);
  b.updatedAt = now || new Date().toISOString();
  return JSON.stringify(b);
}

/* Reads a saved batch back. Anything that was in flight goes back to waiting,
   so an interrupted file is converted again and a finished one is not. */
function restoreBatch(json) {
  var b = typeof json === 'string' ? JSON.parse(json) : clone(json);
  if (!b || b.version !== 1 || !Array.isArray(b.items)) throw new Error('This is not a SlideSize batch record, or it comes from a newer version.');
  b.settings = mergeSettings(DEFAULTS, b.settings);
  b.items.forEach(function (it) {
    if (ACTIVE[it.status]) { it.status = 'ready'; it.interrupted = true; }
    if (!it.issues) it.issues = [];
  });
  return b;
}

function resumeSummary(batch) {
  var s = summarize(batch.items);
  var left = batch.items.filter(function (it) { return it.status === 'ready' || it.status === 'new' || it.status === 'blocked'; }).length;
  return { done: s.finished, left: left, total: batch.items.length, interrupted: batch.items.filter(function (it) { return it.interrupted; }).length };
}

/* ------------------------------------------------------------------ *
 * A small cache with a hard ceiling, for previews                     *
 * ------------------------------------------------------------------ */

function LRU(max, dispose) {
  this.max = Math.max(1, max); this.dispose = dispose || null; this.map = new Map();
}
LRU.prototype.get = function (k) {
  if (!this.map.has(k)) return undefined;
  var v = this.map.get(k); this.map.delete(k); this.map.set(k, v); return v;
};
LRU.prototype.set = function (k, v) {
  if (this.map.has(k)) { var old = this.map.get(k); this.map.delete(k); if (this.dispose && old !== v) this.dispose(old, k); }
  this.map.set(k, v);
  while (this.map.size > this.max) {
    var first = this.map.keys().next().value, ov = this.map.get(first);
    this.map.delete(first); if (this.dispose) this.dispose(ov, first);
  }
};
LRU.prototype.clear = function () {
  var self = this;
  if (this.dispose) this.map.forEach(function (v, k) { self.dispose(v, k); });
  this.map.clear();
};
Object.defineProperty(LRU.prototype, 'size', { get: function () { return this.map.size; } });

/* ------------------------------------------------------------------ *
 * Picture comparison                                                  *
 * ------------------------------------------------------------------ */

var SENSITIVITY = {
  low:    { mean: 26, abs: 60, share: 0.04,  minBlocks: 6 },
  normal: { mean: 16, abs: 44, share: 0.012, minBlocks: 3 },
  high:   { mean: 10, abs: 30, share: 0.004, minBlocks: 1 }
};

/* a and b are RGBA pixel arrays of the same w x h. The picture is cut into blocks.
   A block is flagged when its average brightness differs (something is there in one
   and not the other) or when its pixels differ a lot on average (something moved or
   changed shape). Small differences in edge smoothing do neither. */
function compareImages(a, b, w, h, sensitivity, block) {
  var t = SENSITIVITY[sensitivity] || SENSITIVITY.normal;
  block = block || 8;
  var bx = Math.ceil(w / block), by = Math.ceil(h / block), flags = new Uint8Array(bx * by), n = 0, worst = 0, sum = 0;
  for (var gy = 0; gy < by; gy++) {
    for (var gx = 0; gx < bx; gx++) {
      var x0 = gx * block, y0 = gy * block, x1 = Math.min(w, x0 + block), y1 = Math.min(h, y0 + block);
      var sa = 0, sb = 0, sd = 0, c = 0;
      for (var y = y0; y < y1; y++) {
        var o = (y * w + x0) * 4;
        for (var x = x0; x < x1; x++, o += 4) {
          var la = 0.299 * a[o] + 0.587 * a[o + 1] + 0.114 * a[o + 2];
          var lb = 0.299 * b[o] + 0.587 * b[o + 1] + 0.114 * b[o + 2];
          var dc = (Math.abs(a[o] - b[o]) + Math.abs(a[o + 1] - b[o + 1]) + Math.abs(a[o + 2] - b[o + 2])) / 3;
          sa += la; sb += lb; sd += dc; c++;
        }
      }
      var dm = Math.abs(sa - sb) / c, da = sd / c;
      sum += da;
      if (da > worst) worst = da;
      if (dm > t.mean || da > t.abs) { flags[gy * bx + gx] = 1; n++; }
    }
  }
  var share = n / (bx * by);
  return { blocksX: bx, blocksY: by, block: block, flags: flags, flagged: n, changedShare: share, worst: worst, meanDelta: sum / (bx * by),
           verdict: share >= t.share && n >= t.minBlocks ? 'different' : 'similar' };
}

/* A poster frame that is one flat colour tells the reader nothing. */
function looksBlank(rgba, w, h) {
  var n = w * h, sum = 0, sq = 0, step = Math.max(1, Math.floor(n / 20000));
  var c = 0;
  for (var i = 0; i < n; i += step) {
    var o = i * 4, l = 0.299 * rgba[o] + 0.587 * rgba[o + 1] + 0.114 * rgba[o + 2];
    sum += l; sq += l * l; c++;
  }
  var mean = sum / c, sd = Math.sqrt(Math.max(0, sq / c - mean * mean));
  return { blank: sd < 3, mean: mean, deviation: sd, colour: mean < 24 ? 'black' : mean > 232 ? 'white' : 'flat' };
}

/* ------------------------------------------------------------------ *
 * Reports                                                             *
 * ------------------------------------------------------------------ */

function reportData(batch, opts) {
  opts = opts || {};
  var items = batch.items.map(function (it) {
    var s = settingsFor(it, batch.settings), res = it.result || {};
    var state = (it.status === 'done' || it.status === 'failed' || it.status === 'cancelled') ? classify(it) : null;
    return {
      id: it.id, file: it.name, folder: it.folder || '', sizeBytes: it.size, output: it.hasPdf ? it.outName : null,
      status: it.status, state: state, stateLabel: state ? STATES[state] : statusLabel(it.status),
      attempts: it.attempts || 0,
      slides: (it.inspect && it.inspect.slideCount != null ? it.inspect.slideCount : (res.facts && res.facts.slides)) || null,
      slideSizePt: it.inspect && it.inspect.widthPt ? { w: it.inspect.widthPt, h: it.inspect.heightPt } : null,
      settings: s, settingsSummary: describeSettings(s), hasOverrides: !!it.overrides,
      engine: res.engine || null,
      timing: it.timing || null,
      excludedSlides: it.checks ? it.checks.excludedSlides || [] : [],
      pageMap: res.export ? res.export.pageMap || null : null,
      exportMethod: res.export ? res.export.method || null : null,
      fonts: it.checks ? it.checks.fonts || null : null,
      media: it.checks ? it.checks.media || [] : [],
      fidelity: it.checks ? it.checks.fidelity || null : null,
      pdfa: it.checks ? it.checks.pdfa || null : null,
      pdf: it.checks ? { pages: it.checks.pageCount, sizePt: it.checks.pageSize, bytes: it.pdfBytes, tagged: it.checks.tagged } : null,
      edits: it.checks ? it.checks.edits || null : null,
      issues: (it.issues || []).map(function (i) {
        return { id: i.id, code: i.code, file: it.name, slide: i.slide, page: i.page, severity: i.severity, certainty: i.certainty,
                 description: i.description, impact: i.impact, fix: i.fix, needsReview: !!i.review, accepted: i.accepted || null };
      })
    };
  });
  return {
    generator: 'SlideSize PowerPoint to PDF', protocol: PROTOCOL, batchId: batch.id,
    createdAt: batch.createdAt, updatedAt: batch.updatedAt, reportWrittenAt: opts.now || new Date().toISOString(),
    platform: batch.platform, helper: batch.helper || null,
    settings: batch.settings, settingsSummary: describeSettings(batch.settings),
    summary: summarize(batch.items),
    notes: [
      'A completed conversion, a matching page count and a clean picture comparison are evidence, not a guarantee of fidelity.',
      'Confirmed means read from the file or reported by PowerPoint. Suspected means inferred. Not verified means the check could not be run.',
      'All processing happened on the computer that ran the batch. Nothing was uploaded.'
    ],
    files: items
  };
}

function statusLabel(st) {
  return ({ 'new': 'Not checked', checking: 'Checking', ready: 'Waiting', blocked: 'Held back', skipped: 'Skipped', staging: 'Copying',
            queued: 'Queued', converting: 'Converting', verifying: 'Checking the PDF', done: 'Done', failed: 'Failed', cancelled: 'Cancelled' })[st] || st;
}

function reportJson(batch, opts) { return JSON.stringify(reportData(batch, opts), null, 2); }

function csvCell(v) {
  var s = v == null ? '' : String(v);
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;         /* stop a spreadsheet reading a file name as a formula */
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function reportCsv(batch, opts) {
  var d = reportData(batch, opts);
  var head = ['File', 'Output PDF', 'Result', 'Slide', 'PDF page', 'Severity', 'Certainty', 'Issue', 'Description', 'Likely impact', 'Suggested fix',
              'Needs review', 'Accepted', 'Engine', 'Engine version', 'Processing seconds', 'PDF pages', 'PDF/A', 'Fidelity check'];
  var rows = [head];
  d.files.forEach(function (f) {
    var tail = [f.engine ? f.engine.name : '', f.engine ? f.engine.version : '', f.timing && f.timing.totalMs != null ? trimNum(f.timing.totalMs / 1000, 1) : '',
                f.pdf && f.pdf.pages != null ? f.pdf.pages : '', pdfaLabel(f.pdfa, f.settings), fidelityLabel(f.fidelity, f.settings)];
    var base = [f.file, f.output || '', f.stateLabel];
    if (!f.issues.length) rows.push(base.concat(['', '', '', '', '', '', '', '', '', '']).concat(tail));
    f.issues.forEach(function (i) {
      rows.push(base.concat([i.slide == null ? '' : i.slide, i.page == null ? '' : i.page, SEVERITY[i.severity], CERTAINTY[i.certainty], i.code,
        i.description, i.impact, i.fix, i.needsReview ? 'yes' : '', i.accepted || '']).concat(tail));
    });
  });
  return '﻿' + rows.map(function (r) { return r.map(csvCell).join(','); }).join('\r\n') + '\r\n';
}

function pdfaLabel(p, s) {
  if (!s || !s.pdfa) return 'Not requested';
  if (!p) return 'Not verified';
  return p.result === 'passed' ? 'Passed' : p.result === 'failed' ? 'Failed' : 'Not verified';
}

function fidelityLabel(f, s) {
  if (!f) return s && s.fidelity === 'off' ? 'Not performed, turned off' : 'Not performed';
  if (f.performed === false) return 'Not performed';
  var n = (f.pages || []).length, d = (f.pages || []).filter(function (p) { return p.verdict === 'different'; }).length;
  return plural(n, 'page') + ' of ' + (f.totalPages || n) + ' compared, ' + (d ? d + ' with a suspected difference' : 'no difference above the threshold');
}

function reportHtml(batch, opts) {
  var d = reportData(batch, opts), s = d.summary, h = [];
  function row(k, v) { return '<tr><th>' + esc(k) + '</th><td>' + v + '</td></tr>'; }
  h.push('<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">');
  h.push('<title>PowerPoint to PDF report</title><style>');
  h.push('body{font:14px/1.5 -apple-system,Segoe UI,Helvetica,Arial,sans-serif;color:#1a1a1a;background:#fff;margin:0;padding:28px;max-width:1100px}');
  h.push('h1{font-size:22px;margin:0 0 4px}h2{font-size:16px;margin:30px 0 8px;padding-top:14px;border-top:2px solid #1a1a1a}h3{font-size:13px;margin:16px 0 6px;text-transform:uppercase;letter-spacing:.5px;color:#555}');
  h.push('p{margin:6px 0}.muted{color:#666}.tiles{display:flex;flex-wrap:wrap;gap:10px;margin:14px 0}.tile{border:1px solid #ccc;border-radius:3px;padding:8px 14px;min-width:120px}.tile b{display:block;font-size:22px}');
  h.push('table{border-collapse:collapse;width:100%;margin:6px 0 10px}th,td{text-align:left;vertical-align:top;padding:5px 8px;border-bottom:1px solid #e2e2e2;font-size:13px}th{font-weight:600;white-space:nowrap}');
  h.push('table.kv th{width:190px;color:#555}.st{display:inline-block;padding:1px 8px;border-radius:3px;font-weight:600;font-size:12px;border:1px solid}');
  h.push('.completed{color:#1b7a4b;border-color:#1b7a4b}.warnings{color:#8a6100;border-color:#8a6100}.review{color:#a14a00;border-color:#a14a00;background:#fff4e8}.failed{color:#b3160f;border-color:#b3160f;background:#fdecea}.none{color:#666;border-color:#bbb}');
  h.push('.sev-error{color:#b3160f;font-weight:600}.sev-warning{color:#8a6100;font-weight:600}.sev-info{color:#555}.acc{color:#1b7a4b}');
  h.push('ul{margin:4px 0 4px 18px;padding:0}@media print{body{padding:0}h2{break-before:auto}tr{break-inside:avoid}}');
  h.push('</style></head><body>');
  h.push('<h1>PowerPoint to PDF report</h1>');
  h.push('<p class="muted">Batch started ' + esc(d.createdAt) + '. Report written ' + esc(d.reportWrittenAt) + '. ' +
    (d.helper && d.helper.powerpoint ? 'Converted by Microsoft PowerPoint ' + esc(d.helper.powerpoint.version || '') + (d.helper.powerpoint.build ? ' build ' + esc(d.helper.powerpoint.build) : '') + ' on ' + esc(platformName(d.platform)) + '.' : '') + '</p>');
  h.push('<div class="tiles">');
  [['Files', s.total], ['Completed', s.completed], ['With warnings', s.warnings], ['Need review', s.review], ['Failed', s.failed], ['Not converted', s.pending + s.skipped + s.blocked]]
    .forEach(function (t) { h.push('<div class="tile"><b>' + t[1] + '</b>' + esc(t[0]) + '</div>'); });
  h.push('</div>');
  h.push('<p>' + d.notes.map(esc).join(' ') + '</p>');
  h.push('<h3>Batch settings</h3><ul>' + d.settingsSummary.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul>');

  h.push('<h2>Files</h2><table><tr><th>File</th><th>Result</th><th>Output</th><th>Slides</th><th>PDF pages</th><th>Issues</th><th>Time</th></tr>');
  d.files.forEach(function (f) {
    var open = f.issues.filter(function (i) { return i.severity !== 'info'; }).length;
    h.push('<tr><td><a href="#' + f.id + '">' + esc(f.file) + '</a></td><td><span class="st ' + (f.state || 'none') + '">' + esc(f.stateLabel) + '</span></td><td>' + esc(f.output || '') +
      '</td><td>' + (f.slides == null ? '' : f.slides) + '</td><td>' + (f.pdf && f.pdf.pages != null ? f.pdf.pages : '') + '</td><td>' + (open || '') +
      '</td><td>' + esc(f.timing ? formatDuration(f.timing.totalMs) : '') + '</td></tr>');
  });
  h.push('</table>');

  d.files.forEach(function (f) {
    h.push('<h2 id="' + f.id + '">' + esc(f.file) + ' <span class="st ' + (f.state || 'none') + '">' + esc(f.stateLabel) + '</span></h2>');
    h.push('<table class="kv">');
    h.push(row('Output', f.output ? esc(f.output) : 'No PDF'));
    if (f.slideSizePt) h.push(row('Original slide', esc(sizeLabel(f.slideSizePt.w, f.slideSizePt.h)) + ', ' + esc(aspectLabel(f.slideSizePt.w, f.slideSizePt.h)) + ', ' + plural(f.slides || 0, 'slide')));
    if (f.pdf && f.pdf.pages != null) h.push(row('PDF', plural(f.pdf.pages, 'page') + (f.pdf.sizePt ? ', ' + esc(sizeLabel(f.pdf.sizePt.w, f.pdf.sizePt.h)) : '') + (f.pdf.bytes ? ', ' + esc(formatBytes(f.pdf.bytes)) : '')));
    h.push(row('Settings', (f.hasOverrides ? '<em>Own settings for this file.</em> ' : '') + esc(f.settingsSummary.join('. ')) + '.'));
    if (f.engine) h.push(row('Engine', esc(f.engine.name) + ' ' + esc(f.engine.version || '') + (f.engine.build ? ' build ' + esc(f.engine.build) : '') + (f.exportMethod ? ', ' + esc(f.exportMethod) : '')));
    if (f.timing) h.push(row('Processing time', esc(formatDuration(f.timing.totalMs)) + (f.timing.exportMs != null ? ', of which export ' + esc(formatDuration(f.timing.exportMs)) : '') + (f.attempts > 1 ? ', attempt ' + f.attempts : '')));
    h.push(row('Excluded slides', f.excludedSlides.length ? esc(listSlides(f.excludedSlides, 40)) : 'None'));
    if (f.fonts) {
      h.push(row('Font substitutions', f.fonts.missing && f.fonts.missing.length
        ? esc(f.fonts.missing.join(', ')) + ' not installed.' + (f.fonts.substitutes && f.fonts.substitutes.length ? ' Likely substitutes in the PDF ' + esc(f.fonts.substitutes.join(', ')) + ' (suspected).' : '')
        : 'None found'));
      h.push(row('Rasterised text', f.fonts.restricted && f.fonts.restricted.length
        ? esc(f.fonts.restricted.join(', ')) + (f.settings.bitmapText ? ' written as bitmap because the licence forbids embedding.' : ' not embeddable, left as text.')
        : f.fonts.absent && f.fonts.absent.length ? 'Not verified for ' + esc(f.fonts.absent.join(', ')) : 'None found'));
    }
    h.push(row('Media handling', f.media && f.media.length ? '<ul>' + f.media.map(function (m) {
      return '<li>Slide ' + m.slide + ', ' + esc(m.type) + ' ' + esc(m.name || '') + ', ' + esc(m.handling) + '</li>'; }).join('') + '</ul>' : 'No video or audio found'));
    h.push(row('Fidelity check', esc(fidelityLabel(f.fidelity, f.settings)) + (f.fidelity && f.fidelity.method ? '. ' + esc(f.fidelity.method) : '')));
    h.push(row('PDF/A validation', esc(pdfaLabel(f.pdfa, f.settings)) + (f.pdfa && f.pdfa.claim ? '. File declares PDF/A-' + esc(f.pdfa.claim.part + String(f.pdfa.claim.conformance || '').toLowerCase()) : '') +
      (f.pdfa && f.pdfa.validator ? '. Validator ' + esc(f.pdfa.validator) : f.pdfa ? '. No full validator was run' : '')));
    h.push('</table>');
    if (f.issues.length) {
      h.push('<table><tr><th>Slide</th><th>PDF page</th><th>Severity</th><th>Certainty</th><th>Description</th><th>Likely impact</th><th>Suggested fix</th></tr>');
      f.issues.forEach(function (i) {
        h.push('<tr><td>' + (i.slide == null ? '' : i.slide) + '</td><td>' + (i.page == null ? '' : i.page) + '</td><td class="sev-' + i.severity + '">' + SEVERITY[i.severity] +
          '</td><td>' + CERTAINTY[i.certainty] + '</td><td>' + esc(i.description) + (i.accepted ? ' <span class="acc">Reviewed and accepted ' + esc(i.accepted) + '.</span>' : i.needsReview ? ' <b>Needs review.</b>' : '') +
          '</td><td>' + esc(i.impact) + '</td><td>' + esc(i.fix) + '</td></tr>');
      });
      h.push('</table>');
    } else h.push('<p class="muted">No issues recorded.</p>');
  });
  h.push('<p class="muted">Made by SlideSize, slidesize.com. Conversion by Microsoft PowerPoint on the computer that ran the batch.</p></body></html>');
  return h.join('\n');
}

return {
  PROTOCOL: PROTOCOL, HELPER_VERSION: HELPER_VERSION, ENGINE_NAME: ENGINE_NAME, WORK_DIR: WORK_DIR,
  PT_PER_IN: PT_PER_IN, PT_PER_CM: PT_PER_CM, PPT_MAX_PT: PPT_MAX_PT, PDF_MAX_PT: PDF_MAX_PT,
  SOURCE_WARN_BYTES: SOURCE_WARN_BYTES, PDF_EDIT_MAX_BYTES: PDF_EDIT_MAX_BYTES, HEARTBEAT_STALE_MS: HEARTBEAT_STALE_MS,
  FEATURES: FEATURES, SIZE_PRESETS: SIZE_PRESETS, DEFAULTS: DEFAULTS, BUILTIN_PRESETS: BUILTIN_PRESETS,
  STATES: STATES, SEVERITY: SEVERITY, CERTAINTY: CERTAINTY, SENSITIVITY: SENSITIVITY,
  clamp: clamp, gcd: gcd, trimNum: trimNum, aspectLabel: aspectLabel, ptToCm: ptToCm, ptToIn: ptToIn, cmToPt: cmToPt,
  sizeLabel: sizeLabel, orientationOf: orientationOf, formatBytes: formatBytes, formatDuration: formatDuration,
  plural: plural, listSlides: listSlides, esc: esc,
  detectPlatform: detectPlatform, platformName: platformName, browserSupport: browserSupport, featureAvailable: featureAvailable,
  defaultSettings: defaultSettings, mergeSettings: mergeSettings, resolveSettings: resolveSettings, settingsFor: settingsFor, describeSettings: describeSettings,
  targetSize: targetSize, fitTransform: fitTransform, compareWindows: compareWindows, expectedSlides: expectedSlides, planOutput: planOutput,
  safeBase: safeBase, planNames: planNames,
  issue: issue, numberIssues: numberIssues, preflightIssues: preflightIssues, resultIssues: resultIssues, pdfaIssues: pdfaIssues,
  compareIssues: compareIssues, mergeIssues: mergeIssues, helperFailureIssue: helperFailureIssue,
  normFont: normFont, fontMatches: fontMatches, fontFindings: fontFindings,
  ticketId: ticketId, placeholdersFor: placeholdersFor, buildTicket: buildTicket, sampleSlides: sampleSlides,
  helperStatus: helperStatus, classify: classify, summarize: summarize,
  newBatch: newBatch, addItem: addItem, sameFile: sameFile, slimInspect: slimInspect, serializeBatch: serializeBatch, restoreBatch: restoreBatch, resumeSummary: resumeSummary,
  LRU: LRU, compareImages: compareImages, looksBlank: looksBlank,
  reportData: reportData, reportJson: reportJson, reportCsv: reportCsv, reportHtml: reportHtml, statusLabel: statusLabel,
  pdfaLabel: pdfaLabel, fidelityLabel: fidelityLabel
};
});
