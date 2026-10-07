/* ============================================================
   PowerPoint to PDF: reading and editing the finished PDF
   (PptPdfPdf)

   Three jobs, all on the user's computer.

   inspect      Opens the PDF PowerPoint wrote and reads back what
                is really in it: page count, page sizes, which
                fonts are embedded, whether it is tagged, what it
                declares about PDF/A.

   pdfaChecks   A handful of structural PDF/A rules that can be
                checked here. They can prove a file is NOT PDF/A.
                They cannot prove that it is. Only a full
                validator can, so the best this returns on its
                own is "not verified".

   transform    Resizes pages by scaling and centring the page
                content on a new page box. One scale for both
                axes, so nothing is stretched. The text, vectors,
                tags and links PowerPoint wrote stay as they are.
                Also removes links and strips metadata.

   Uses pdf-lib (MIT, Andrew Dillon), vendored as pdf-lib.min.js.
   No DOM access. Runs in the page and in Node.

   AEGFX / SlideSize
   ============================================================ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(function () { return require('./pdf-lib.min.js'); });
  else root.PptPdfPdf = factory(function () { return root.PDFLib; });
})(typeof self !== 'undefined' ? self : this, function (getLib) {
'use strict';

function lib() {
  var L = getLib();
  if (!L) throw new Error('The PDF library has not loaded.');
  return L;
}

function latin1(bytes, start, end) {
  var s = '', e = Math.min(bytes.length, end);
  for (var i = start; i < e; i += 8192) s += String.fromCharCode.apply(null, bytes.subarray(i, Math.min(e, i + 8192)));
  return s;
}

function num(o) { return o && typeof o.asNumber === 'function' ? o.asNumber() : null; }
function nameOf(o) { return o && typeof o.asString === 'function' ? o.asString().replace(/^\//, '') : null; }
function textOf(o) {
  if (!o) return null;
  try { if (typeof o.decodeText === 'function') return o.decodeText(); } catch (e) { /* fall through */ }
  return typeof o.asString === 'function' ? o.asString() : null;
}

/* Follows a reference to the object behind it. Never throws. */
function deref(ctx, o) {
  var L = lib(), guard = 0;
  while (o instanceof L.PDFRef && guard++ < 16) o = ctx.lookup(o);
  return o;
}
function dictGet(ctx, dict, key) {
  var L = lib();
  if (!(dict instanceof L.PDFDict)) return undefined;
  return deref(ctx, dict.get(L.PDFName.of(key)));
}
function streamDict(o) {
  var L = lib();
  if (o instanceof L.PDFDict) return o;
  if (o && o.dict instanceof L.PDFDict) return o.dict;
  return null;
}

/* ------------------------------------------------------------------ *
 * Inspect                                                             *
 * ------------------------------------------------------------------ */

async function load(bytes) {
  var L = lib();
  return L.PDFDocument.load(bytes, { updateMetadata: false, throwOnInvalidObject: false });
}

function pageBox(page) {
  var b;
  try { b = page.getCropBox(); } catch (e) { b = page.getMediaBox(); }
  var rot = 0;
  try { rot = ((page.getRotation().angle % 360) + 360) % 360; } catch (e) { rot = 0; }
  var w = Math.abs(b.width), h = Math.abs(b.height);
  return rot === 90 || rot === 270 ? { w: h, h: w, rotate: rot, x: b.x, y: b.y } : { w: w, h: h, rotate: rot, x: b.x, y: b.y };
}

function collectFontsFromResources(ctx, res, fonts, state, depth) {
  var L = lib();
  if (!(res instanceof L.PDFDict) || depth > 6) return;
  var fd = dictGet(ctx, res, 'Font');
  if (fd instanceof L.PDFDict) {
    fd.entries().forEach(function (pair) {
      var ref = pair[1], key = ref instanceof L.PDFRef ? ref.toString() : null;
      if (key && state.seenFont[key]) return;
      if (key) state.seenFont[key] = true;
      var f = deref(ctx, ref);
      if (!(f instanceof L.PDFDict)) return;
      var subtype = nameOf(f.get(L.PDFName.of('Subtype'))) || '';
      var base = nameOf(f.get(L.PDFName.of('BaseFont'))) || nameOf(f.get(L.PDFName.of('Name'))) || '(unnamed)';
      var desc = dictGet(ctx, f, 'FontDescriptor');
      if (subtype === 'Type0') {
        var kids = dictGet(ctx, f, 'DescendantFonts');
        var kid = kids instanceof L.PDFArray && kids.size() ? deref(ctx, kids.get(0)) : null;
        if (kid instanceof L.PDFDict) desc = dictGet(ctx, kid, 'FontDescriptor');
      }
      var embedded = subtype === 'Type3';
      if (desc instanceof L.PDFDict) {
        embedded = embedded || !!(desc.get(L.PDFName.of('FontFile')) || desc.get(L.PDFName.of('FontFile2')) || desc.get(L.PDFName.of('FontFile3')));
      }
      var clean = base.replace(/#([0-9A-Fa-f]{2})/g, function (m, h) { return String.fromCharCode(parseInt(h, 16)); });
      fonts.push({ name: clean, subtype: subtype, embedded: embedded, type3: subtype === 'Type3', subset: /^[A-Z]{6}\+/.test(clean) });
    });
  }
  var gs = dictGet(ctx, res, 'ExtGState');
  if (gs instanceof L.PDFDict) {
    gs.entries().forEach(function (pair) {
      var g = deref(ctx, pair[1]);
      if (!(g instanceof L.PDFDict)) return;
      var CA = num(g.get(L.PDFName.of('CA'))), ca = num(g.get(L.PDFName.of('ca')));
      var sm = deref(ctx, g.get(L.PDFName.of('SMask'))), bm = nameOf(deref(ctx, g.get(L.PDFName.of('BM'))));
      if ((CA !== null && CA < 1) || (ca !== null && ca < 1)) state.transparency = true;
      if (sm && nameOf(sm) !== 'None') state.transparency = true;
      if (bm && bm !== 'Normal' && bm !== 'Compatible') state.transparency = true;
    });
  }
  var xo = dictGet(ctx, res, 'XObject');
  if (xo instanceof L.PDFDict) {
    xo.entries().forEach(function (pair) {
      var ref = pair[1], key = ref instanceof L.PDFRef ? ref.toString() : null;
      if (key && state.seenX[key]) return;
      if (key) state.seenX[key] = true;
      var x = streamDict(deref(ctx, ref));
      if (!x) return;
      var st = nameOf(x.get(L.PDFName.of('Subtype')));
      if (st === 'Image') { if (x.get(L.PDFName.of('SMask'))) state.transparency = true; state.images++; }
      else if (st === 'Form') {
        var grp = dictGet(ctx, x, 'Group');
        if (grp instanceof L.PDFDict && nameOf(grp.get(L.PDFName.of('S'))) === 'Transparency') state.transparency = true;
        collectFontsFromResources(ctx, dictGet(ctx, x, 'Resources'), fonts, state, depth + 1);
      }
    });
  }
}

function readXmp(ctx, catalog) {
  var L = lib(), md = dictGet(ctx, catalog, 'Metadata');
  if (!md) return null;
  try {
    var bytes = md instanceof L.PDFRawStream ? L.decodePDFRawStream(md).decode() : (md.getContents ? md.getContents() : null);
    if (!bytes) return null;
    return new TextDecoder('utf-8').decode(bytes);
  } catch (e) { return null; }
}

function pdfaClaim(xmp) {
  if (!xmp) return null;
  var part = /pdfaid:part\s*=\s*["'](\d)["']/.exec(xmp) || /<pdfaid:part>\s*(\d)\s*<\/pdfaid:part>/.exec(xmp);
  if (!part) return null;
  var conf = /pdfaid:conformance\s*=\s*["']([A-Za-z])["']/.exec(xmp) || /<pdfaid:conformance>\s*([A-Za-z])\s*<\/pdfaid:conformance>/.exec(xmp);
  return { part: +part[1], conformance: conf ? conf[1].toUpperCase() : '' };
}

/* Returns the facts, or throws an Error whose .code is 'encrypted' or 'damaged'. */
async function inspect(bytes) {
  var L = lib(), doc;
  var header = latin1(bytes, 0, 16), hm = /%PDF-(\d\.\d)/.exec(header);
  if (!hm) { var e0 = new Error('The file does not start like a PDF.'); e0.code = 'damaged'; throw e0; }
  try { doc = await load(bytes); }
  catch (e) {
    var err = new Error(/encrypt/i.test(String(e && e.message)) ? 'The PDF is encrypted.' : 'The PDF could not be parsed. ' + (e && e.message ? e.message : ''));
    err.code = /encrypt/i.test(String(e && e.message)) ? 'encrypted' : 'damaged';
    throw err;
  }
  var ctx = doc.context, catalog = doc.catalog, pages = doc.getPages();
  if (!pages.length) { var e1 = new Error('The PDF has no pages.'); e1.code = 'damaged'; throw e1; }

  var out = { version: hm[1], pageCount: pages.length, pages: [], fonts: [], links: 0, annotations: 0, images: 0,
              tagged: false, transparency: false, javascript: false, embeddedFiles: false, outputIntents: 0,
              hasId: !!(ctx.trailerInfo && ctx.trailerInfo.ID), xmp: false, claim: null, info: {}, objectStreams: false, lang: null };
  var state = { seenFont: {}, seenX: {}, transparency: false, images: 0 };

  pages.forEach(function (p) {
    out.pages.push(pageBox(p));
    collectFontsFromResources(ctx, p.node.Resources(), out.fonts, state, 0);
    var grp = dictGet(ctx, p.node, 'Group');
    if (grp instanceof L.PDFDict && nameOf(grp.get(L.PDFName.of('S'))) === 'Transparency') state.transparency = true;
    var annots = p.node.Annots();
    if (annots) for (var i = 0; i < annots.size(); i++) {
      var a = deref(ctx, annots.get(i));
      if (!(a instanceof L.PDFDict)) continue;
      out.annotations++;
      if (nameOf(a.get(L.PDFName.of('Subtype'))) === 'Link') out.links++;
      var act = dictGet(ctx, a, 'A');
      if (act instanceof L.PDFDict && nameOf(act.get(L.PDFName.of('S'))) === 'JavaScript') out.javascript = true;
    }
  });
  out.transparency = state.transparency; out.images = state.images;

  var mark = dictGet(ctx, catalog, 'MarkInfo');
  var marked = mark instanceof L.PDFDict ? mark.get(L.PDFName.of('Marked')) : null;
  out.tagged = !!(dictGet(ctx, catalog, 'StructTreeRoot') && marked && String(marked) === 'true');
  out.hasStructTree = !!dictGet(ctx, catalog, 'StructTreeRoot');
  out.lang = textOf(dictGet(ctx, catalog, 'Lang'));
  var oi = dictGet(ctx, catalog, 'OutputIntents');
  out.outputIntents = oi instanceof L.PDFArray ? oi.size() : 0;
  var names = dictGet(ctx, catalog, 'Names');
  if (names instanceof L.PDFDict) {
    if (names.get(L.PDFName.of('JavaScript'))) out.javascript = true;
    if (names.get(L.PDFName.of('EmbeddedFiles'))) out.embeddedFiles = true;
  }
  var oa = dictGet(ctx, catalog, 'OpenAction');
  if (oa instanceof L.PDFDict && nameOf(oa.get(L.PDFName.of('S'))) === 'JavaScript') out.javascript = true;

  var xmp = readXmp(ctx, catalog);
  out.xmp = !!xmp; out.claim = pdfaClaim(xmp);
  ['Title', 'Author', 'Subject', 'Keywords', 'Creator', 'Producer'].forEach(function (k) {
    var v;
    try { v = doc['get' + k](); } catch (e) { v = undefined; }
    if (v) out.info[k.toLowerCase()] = String(v);
  });

  /* object streams and cross-reference streams are a PDF 1.5 feature that PDF/A-1 forbids */
  var tail = latin1(bytes, Math.max(0, bytes.length - 4096), bytes.length);
  out.objectStreams = /\/Type\s*\/XRef\b/.test(tail) || /\/Type\s*\/ObjStm\b/.test(latin1(bytes, 0, Math.min(bytes.length, 262144)));

  var uniform = out.pages.every(function (p) { return Math.abs(p.w - out.pages[0].w) < 0.5 && Math.abs(p.h - out.pages[0].h) < 0.5; });
  out.uniform = uniform;
  out.unembeddedFonts = out.fonts.filter(function (f) { return !f.embedded; }).map(function (f) { return f.name; });
  return out;
}

/* ------------------------------------------------------------------ *
 * PDF/A structural checks                                             *
 * ------------------------------------------------------------------ */

/* info is what inspect() returned. requested says PDF/A was asked for.
   The result is 'failed' when a rule is broken and 'not-verified' otherwise. Never 'passed'. */
function pdfaChecks(info) {
  var c = info.claim, checks = [];
  function add(id, label, ok, detail, applies) {
    checks.push({ id: id, label: label, result: applies === false ? 'n/a' : ok ? 'pass' : 'fail', detail: detail });
  }
  add('claim', 'Declares PDF/A in its metadata', !!c, c ? 'Declares PDF/A-' + c.part + c.conformance.toLowerCase() + '.' : 'The XMP metadata has no PDF/A identification, so the file does not claim to be PDF/A at all.');
  add('fonts', 'Every font is embedded', info.unembeddedFonts.length === 0,
    info.unembeddedFonts.length ? 'Not embedded: ' + info.unembeddedFonts.slice(0, 8).join(', ') + '.' : plural(info.fonts.length, 'font') + ', all embedded.');
  add('id', 'Has a file identifier', info.hasId, info.hasId ? 'Present.' : 'The trailer has no ID entry.');
  add('javascript', 'No JavaScript', !info.javascript, info.javascript ? 'The file contains JavaScript.' : 'None found.');
  var part = c ? c.part : 1;
  add('transparency', 'No transparency', !info.transparency,
    info.transparency ? 'The file uses transparency, which PDF/A-1 does not allow.' : 'None found.', part === 1);
  add('embedded-files', 'No embedded files', !info.embeddedFiles,
    info.embeddedFiles ? 'The file carries attachments, which PDF/A-1 does not allow.' : 'None found.', part === 1);
  add('object-streams', 'No object streams', !info.objectStreams,
    info.objectStreams ? 'The file uses object or cross-reference streams, which PDF/A-1 does not allow.' : 'None found.', part === 1);
  add('tagged', 'Tagged, for conformance level A', info.tagged,
    info.tagged ? 'Marked and has a structure tree.' : 'Level A needs a structure tree and the file has none or is not marked.', !!c && c.conformance === 'A');
  var failed = checks.filter(function (x) { return x.result === 'fail'; });
  return { claim: c, checks: checks, result: failed.length ? 'failed' : 'not-verified', validator: null,
           profile: c ? 'PDF/A-' + c.part + c.conformance.toLowerCase() : null };
}

function plural(n, w) { return n + ' ' + w + (n === 1 ? '' : 's'); }

/* ------------------------------------------------------------------ *
 * Transform                                                           *
 * ------------------------------------------------------------------ */

function fit(srcW, srcH, dstW, dstH, mode) {
  var sx = dstW / srcW, sy = dstH / srcH, s = mode === 'fill' ? Math.max(sx, sy) : Math.min(sx, sy);
  return { scale: s, tx: (dstW - s * srcW) / 2, ty: (dstH - s * srcH) / 2 };
}

function mapPoint(t, x, y) { return [t.scale * (x - t.ox) + t.tx, t.scale * (y - t.oy) + t.ty]; }

function mapNumbers(ctx, arr, t, pairs) {
  var L = lib();
  if (!(arr instanceof L.PDFArray)) return false;
  for (var i = 0; i + 1 < arr.size() && i < pairs * 2; i += 2) {
    var x = num(deref(ctx, arr.get(i))), y = num(deref(ctx, arr.get(i + 1)));
    if (x === null || y === null) return false;
    var p = mapPoint(t, x, y);
    arr.set(i, L.PDFNumber.of(+p[0].toFixed(3))); arr.set(i + 1, L.PDFNumber.of(+p[1].toFixed(3)));
  }
  return true;
}

function mapDest(ctx, dest, byPage) {
  var L = lib();
  if (!(dest instanceof L.PDFArray) || dest.size() < 2) return;
  var target = dest.get(0), t = target instanceof L.PDFRef ? byPage[target.toString()] : null;
  if (!t) return;
  var kind = nameOf(dest.get(1));
  function setX(i) { var v = num(deref(ctx, dest.get(i))); if (v !== null) dest.set(i, L.PDFNumber.of(+(t.scale * (v - t.ox) + t.tx).toFixed(3))); }
  function setY(i) { var v = num(deref(ctx, dest.get(i))); if (v !== null) dest.set(i, L.PDFNumber.of(+(t.scale * (v - t.oy) + t.ty).toFixed(3))); }
  if (kind === 'XYZ') { if (dest.size() > 2) setX(2); if (dest.size() > 3) setY(3); }
  else if (kind === 'FitH' || kind === 'FitBH') { if (dest.size() > 2) setY(2); }
  else if (kind === 'FitV' || kind === 'FitBV') { if (dest.size() > 2) setX(2); }
  else if (kind === 'FitR' && dest.size() > 5) { setX(2); setY(3); setX(4); setY(5); }
}

/* ops: {
     resize: { w, h, mode: 'fit' | 'fill', margin: 'none' | 'white' | 'black' }   page size in points
     removeLinks: true
     stripMetadata: true
   }
   Returns { bytes, pages: [{ srcW, srcH, w, h, scale, tx, ty, skipped }], linksRemoved, metadataStripped, resized }. */
async function transform(bytes, ops) {
  var L = lib(), doc = await load(bytes), ctx = doc.context, pages = doc.getPages();
  var out = { pages: [], linksRemoved: null, metadataStripped: false, resized: 0, skipped: [] };
  var byPage = {};

  if (ops.resize) {
    var W = +ops.resize.w, H = +ops.resize.h;
    if (!(W > 0 && H > 0)) throw new Error('The page size to resize to is not valid.');
    pages.forEach(function (page, idx) {
      var b = pageBox(page);
      if (b.rotate !== 0) { out.pages.push({ srcW: b.w, srcH: b.h, skipped: 'rotated page' }); out.skipped.push(idx + 1); return; }
      var f = fit(b.w, b.h, W, H, ops.resize.mode), t = { scale: f.scale, tx: f.tx, ty: f.ty, ox: b.x, oy: b.y };
      byPage[page.ref.toString()] = t;
      var same = Math.abs(b.w - W) < 0.01 && Math.abs(b.h - H) < 0.01 && Math.abs(b.x) < 0.01 && Math.abs(b.y) < 0.01;
      out.pages.push({ srcW: b.w, srcH: b.h, w: W, h: H, scale: f.scale, tx: f.tx, ty: f.ty });
      if (same) return;

      /* one matrix in front of everything PowerPoint drew, and the page box changed around it */
      var start = [L.pushGraphicsState()];
      if (ops.resize.margin === 'white' || ops.resize.margin === 'black') {
        var v = ops.resize.margin === 'white' ? 1 : 0;
        start.push(L.setFillingRgbColor(v, v, v), L.rectangle(0, 0, W, H), L.fill());
      }
      start.push(L.pushGraphicsState(), L.concatTransformationMatrix(f.scale, 0, 0, f.scale, f.tx - f.scale * b.x, f.ty - f.scale * b.y));
      page.node.normalize();
      var startRef = ctx.register(ctx.contentStream(start));
      var endRef = ctx.register(ctx.contentStream([L.popGraphicsState(), L.popGraphicsState()]));
      page.node.wrapContentStreams(startRef, endRef);
      page.setMediaBox(0, 0, W, H);
      ['CropBox', 'BleedBox', 'TrimBox', 'ArtBox'].forEach(function (k) { page.node.delete(L.PDFName.of(k)); });
      out.resized++;

      var annots = page.node.Annots();
      if (annots) for (var i = 0; i < annots.size(); i++) {
        var a = deref(ctx, annots.get(i));
        if (!(a instanceof L.PDFDict)) continue;
        var rect = dictGet(ctx, a, 'Rect');
        if (rect instanceof L.PDFArray && rect.size() === 4) mapNumbers(ctx, rect, t, 2);
        var quad = dictGet(ctx, a, 'QuadPoints');
        if (quad instanceof L.PDFArray) mapNumbers(ctx, quad, t, quad.size() / 2);
      }
    });
    /* link targets inside other pages move with those pages */
    pages.forEach(function (page) {
      var annots = page.node.Annots();
      if (!annots) return;
      for (var i = 0; i < annots.size(); i++) {
        var a = deref(ctx, annots.get(i));
        if (!(a instanceof L.PDFDict)) continue;
        mapDest(ctx, dictGet(ctx, a, 'Dest'), byPage);
        var act = dictGet(ctx, a, 'A');
        if (act instanceof L.PDFDict) mapDest(ctx, dictGet(ctx, act, 'D'), byPage);
      }
    });
  }

  if (ops.removeLinks) {
    var removed = 0;
    pages.forEach(function (page) {
      var annots = page.node.Annots();
      if (!annots) return;
      for (var i = annots.size() - 1; i >= 0; i--) {
        var raw = annots.get(i), a = deref(ctx, raw);
        if (!(a instanceof L.PDFDict) || nameOf(a.get(L.PDFName.of('Subtype'))) !== 'Link') continue;
        annots.remove(i);
        if (raw instanceof L.PDFRef) ctx.delete(raw);       /* the address goes out of the file, not just off the page */
        removed++;
      }
      if (annots.size() === 0) page.node.delete(L.PDFName.of('Annots'));
    });
    out.linksRemoved = removed;
  }

  if (ops.stripMetadata) {
    var infoRef = ctx.trailerInfo.Info, info = deref(ctx, infoRef);
    if (info instanceof L.PDFDict) info.keys().forEach(function (k) { info.delete(k); });
    var mdRef = doc.catalog.get(L.PDFName.of('Metadata'));
    doc.catalog.delete(L.PDFName.of('Metadata'));
    if (mdRef instanceof L.PDFRef) ctx.delete(mdRef);
    out.metadataStripped = true;
  }

  out.bytes = await doc.save({ useObjectStreams: false, addDefaultPage: false, updateFieldAppearances: false });
  return out;
}

return { inspect: inspect, pdfaChecks: pdfaChecks, pdfaClaim: pdfaClaim, transform: transform, fit: fit };
});
