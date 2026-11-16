/* font-extract.js — engine for the Font Extractor page
 *
 * PowerPoint stores embedded fonts as ppt/fonts/fontN.fntdata. Each one is
 * an Embedded OpenType (EOT) container, and PowerPoint's own are squeezed
 * with MicroType Express (MTX). This file
 *
 *   - lists the embedded fonts in a deck and which slot each one fills
 *     (regular, bold, italic, bold italic)
 *   - unwraps each one back to a plain .ttf or .otf
 *   - reads the recovered font to work out whether it is the whole font
 *     or only a subset of its characters
 *   - writes a copy of the deck with the embedded fonts taken out
 *
 * No DOM in here, so the same file runs in the page and under node for the
 * checks in tests/fonts. MTX decompression comes from mtx-decompressor
 * (mtx-decompressor.min.js, MPL-2.0), passed in or picked up as the
 * MTXDecompressor global.
 */
(function (root, factory) {
if (typeof module === 'object' && module.exports) module.exports = factory();
else root.FontExtract = factory();
})(typeof self !== 'undefined' ? self : this, function () {
'use strict';

var FX = {};

/* ---------- small helpers ---------- */

function u16(b, o) { return (b[o] << 8) | b[o + 1]; }
function u32(b, o) { return ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0; }
function u16le(b, o) { return b[o] | (b[o + 1] << 8); }
function u32le(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }
function tag(b, o) { return String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]); }

function xmlDecode(s) {
return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
.replace(/&apos;/g, "'").replace(/&#x([0-9a-f]+);/gi, function (m, h) { return String.fromCodePoint(parseInt(h, 16)); })
.replace(/&#(\d+);/g, function (m, d) { return String.fromCodePoint(parseInt(d, 10)); })
.replace(/&amp;/g, '&');
}
function attr(el, name) {
var m = new RegExp('(?:^|\\s)' + name + '\\s*=\\s*("([^"]*)"|\'([^\']*)\')').exec(el);
return m ? xmlDecode(m[2] != null ? m[2] : m[3]) : null;
}

FX.SLOTS = ['regular', 'bold', 'italic', 'boldItalic'];
FX.SLOT_LABEL = { regular: 'Regular', bold: 'Bold', italic: 'Italic', boldItalic: 'Bold Italic' };

/* ---------- what kind of file is this ---------- */

function isSfnt(b, o) {
o = o || 0;
if (b.length < o + 12) return false;
var t = u32(b, o);
return t === 0x00010000 || t === 0x4F54544F /* OTTO */ || t === 0x74727565 /* true */ || t === 0x74746366 /* ttcf */;
}
/* EOT keeps the magic number 0x504C at byte 34 of the header. */
function isEot(b) { return b.length > 82 && b[34] === 0x4C && b[35] === 0x50; }

/* ---------- EOT header ----------
   Read by hand so the subset flag and names are there even when the font
   data itself will not decompress. Layout per the W3C EOT submission. */
FX.readEotHeader = function (b) {
if (!isEot(b)) return null;
var h = {
eotSize: u32le(b, 0), fontDataSize: u32le(b, 4), version: u32le(b, 8), flags: u32le(b, 12),
italic: b[27] !== 0, weight: u32le(b, 28), fsType: u16le(b, 32)
};
h.subset = (h.flags & 0x00000001) !== 0;
h.compressed = (h.flags & 0x00000004) !== 0;
h.encrypted = (h.flags & 0x10000000) !== 0;
/* Names start after the fixed 82 bytes: size(2) + UTF-16LE string, with
   2 bytes of padding in front of each one after the first. */
var o = 82, names = [];
try {
for (var i = 0; i < 4 && o + 2 <= b.length; i++) {
var n = u16le(b, o); o += 2;
var s = '';
for (var k = 0; k + 1 < n && o + k + 1 < b.length; k += 2) {
var c = u16le(b, o + k);
if (c) s += String.fromCharCode(c);
}
names.push(s);
o += n + 2;
}
} catch (e) {}
h.familyName = names[0] || ''; h.styleName = names[1] || '';
h.versionName = names[2] || ''; h.fullName = names[3] || '';
return h;
};

/* ---------- unwrap one embedded font part ----------
   opts.mtx   the mtx-decompressor module (defaults to the browser global)
   opts.guid  relationship or file name GUID, for the obfuscated kind
   Returns { container, sfnt, eot, error } */
FX.unwrap = function (bytes, opts) {
opts = opts || {};
var out = { container: 'unknown', sfnt: null, eot: null, error: null };
if (!(bytes instanceof Uint8Array)) bytes = new Uint8Array(bytes);

if (isSfnt(bytes)) { out.container = 'plain'; out.sfnt = bytes; return out; }

if (isEot(bytes)) {
out.container = 'eot';
out.eot = FX.readEotHeader(bytes);
var mtx = opts.mtx || (typeof MTXDecompressor !== 'undefined' ? MTXDecompressor : null);
if (!mtx) { out.error = 'MTX decompressor not loaded'; return out; }
try {
var f = mtx.eotToTtf(bytes);
if (!isSfnt(f)) throw new Error('decoded data is not a font');
out.sfnt = f;
} catch (e) { out.error = (e && e.message) || String(e); }
return out;
}

/* Obfuscated font (the Word way, rare in decks): the first 32 bytes are
   XORed with a GUID taken from the relationship or the file name. */
var g = opts.guid && /([0-9a-f]{8})-?([0-9a-f]{4})-?([0-9a-f]{4})-?([0-9a-f]{4})-?([0-9a-f]{12})/i.exec(opts.guid);
if (g && bytes.length > 32) {
var hex = g.slice(1).join(''), key = [];
for (var i = 0; i < 16; i++) key.push(parseInt(hex.substr(i * 2, 2), 16));
var copy = new Uint8Array(bytes);
for (var j = 0; j < 16; j++) { copy[j] ^= key[15 - j]; copy[j + 16] ^= key[15 - j]; }
if (isSfnt(copy)) { out.container = 'obfuscated'; out.sfnt = copy; return out; }
}

out.error = 'Not an EOT or OpenType file';
return out;
};

/* ---------- read a recovered font ---------- */

function tableDir(b) {
var base = 0;
if (u32(b, 0) === 0x74746366) base = u32(b, 12); /* collection: first font */
var n = u16(b, base + 4), t = {};
for (var i = 0; i < n; i++) {
var o = base + 12 + i * 16;
if (o + 16 > b.length) break;
t[tag(b, o)] = { off: u32(b, o + 8), len: u32(b, o + 12) };
}
return { flavor: u32(b, base), collection: base !== 0, tables: t };
}

function readNames(b, t) {
var out = {};
if (!t || t.off + 6 > b.length) return out;
var count = u16(b, t.off + 2), strOff = t.off + u16(b, t.off + 4);
var best = {};
for (var i = 0; i < count; i++) {
var r = t.off + 6 + i * 12;
if (r + 12 > b.length) break;
var pid = u16(b, r), lang = u16(b, r + 4), nid = u16(b, r + 6), len = u16(b, r + 8), off = strOff + u16(b, r + 10);
if (nid > 17 || off + len > b.length) continue;
var score = pid === 3 ? (lang === 0x409 ? 4 : 3) : pid === 0 ? 2 : pid === 1 ? 1 : 0;
if (!score || (best[nid] || 0) >= score) continue;
var s = '', k;
if (pid === 1) { for (k = 0; k < len; k++) s += String.fromCharCode(b[off + k]); }
else { for (k = 0; k + 1 < len; k += 2) s += String.fromCharCode(u16(b, off + k)); }
s = s.replace(/\u0000/g, '').trim();
if (s) { best[nid] = score; out[nid] = s; }
}
return out;
}

/* Returns a Map of code point -> glyph id from the best Unicode subtable. */
function readCmap(b, t) {
var map = new Map();
if (!t || t.off + 4 > b.length) return map;
var n = u16(b, t.off + 2), pick = null, pickScore = -1, symbol = false;
for (var i = 0; i < n; i++) {
var r = t.off + 4 + i * 8;
if (r + 8 > b.length) break;
var pid = u16(b, r), eid = u16(b, r + 2), off = t.off + u32(b, r + 4);
if (off + 4 > b.length) continue;
var fmt = u16(b, off);
if (fmt !== 4 && fmt !== 12 && fmt !== 0 && fmt !== 6) continue;
var s = (pid === 3 && eid === 10) ? 6 : (pid === 0 && fmt === 12) ? 5 : (pid === 3 && eid === 1) ? 4 :
(pid === 0) ? 3 : (pid === 3 && eid === 0) ? 2 : (pid === 1 && eid === 0) ? 1 : 0;
if (s > pickScore) { pickScore = s; pick = { off: off, fmt: fmt }; symbol = (pid === 3 && eid === 0); }
}
if (!pick) return map;
var o = pick.off, k, c;
if (pick.fmt === 0) {
for (c = 0; c < 256 && o + 6 + c < b.length; c++) if (b[o + 6 + c]) map.set(c, b[o + 6 + c]);
} else if (pick.fmt === 6) {
var first = u16(b, o + 6), cnt = u16(b, o + 8);
for (k = 0; k < cnt && o + 12 + k * 2 <= b.length; k++) { var g6 = u16(b, o + 10 + k * 2); if (g6) map.set(first + k, g6); }
} else if (pick.fmt === 4) {
var segX2 = u16(b, o + 6), endO = o + 14, startO = endO + segX2 + 2, deltaO = startO + segX2, rangeO = deltaO + segX2;
for (k = 0; k < segX2; k += 2) {
var end = u16(b, endO + k), start = u16(b, startO + k), delta = u16(b, deltaO + k), ro = u16(b, rangeO + k);
if (start === 0xFFFF) break;
for (c = start; c <= end && c < 0xFFFF; c++) {
var gid;
if (ro === 0) gid = (c + delta) & 0xFFFF;
else {
var p = rangeO + k + ro + (c - start) * 2;
if (p + 2 > b.length) break;
gid = u16(b, p);
if (gid) gid = (gid + delta) & 0xFFFF;
}
if (gid) map.set(c, gid);
}
}
} else if (pick.fmt === 12) {
var groups = u32(b, o + 12);
for (k = 0; k < groups; k++) {
var go = o + 16 + k * 12;
if (go + 12 > b.length) break;
var sc = u32(b, go), ec = u32(b, go + 4), sg = u32(b, go + 8);
if (ec - sc > 0x20000) continue;
for (c = sc; c <= ec; c++) map.set(c, sg + (c - sc));
}
}
/* Symbol fonts (Wingdings and friends) map into F020..F0FF. Mirror them
   down so the characters people actually type are found. */
if (symbol) {
var extra = [];
map.forEach(function (g, cp) { if (cp >= 0xF020 && cp <= 0xF0FF) extra.push([cp - 0xF000, g]); });
extra.forEach(function (e) { if (!map.has(e[0])) map.set(e[0], e[1]); });
}
return map;
}

FX.inspect = function (b) {
var d = tableDir(b), T = d.tables, info = {
format: d.flavor === 0x4F54544F ? 'otf' : 'ttf',
collection: d.collection, size: b.length,
numGlyphs: 0, outlineGlyphs: null, emptyGlyphs: null,
names: {}, fsType: null, chars: [], symbol: false
};
info.ext = d.collection ? 'ttc' : info.format;
if (T.maxp && T.maxp.off + 6 <= b.length) info.numGlyphs = u16(b, T.maxp.off + 4);
if (T['OS/2'] && T['OS/2'].off + 10 <= b.length) info.fsType = u16(b, T['OS/2'].off + 8);
var nm = readNames(b, T.name);
info.names = { family: nm[16] || nm[1] || '', style: nm[17] || nm[2] || '', full: nm[4] || '', postscript: nm[6] || '', version: nm[5] || '' };

/* Which glyphs actually have an outline. TrueType only: a subset made by
   PowerPoint can keep every glyph slot but leave the unused ones empty. */
var hasOutline = null;
if (T.glyf && T.loca && T.head && T.head.off + 52 <= b.length && info.numGlyphs) {
var long = u16(b, T.head.off + 50) === 1, lo = T.loca.off, n = info.numGlyphs;
if (lo + (n + 1) * (long ? 4 : 2) <= b.length) {
hasOutline = new Uint8Array(n);
var empty = 0;
for (var i = 0; i < n; i++) {
var a = long ? u32(b, lo + i * 4) : u16(b, lo + i * 2);
var z = long ? u32(b, lo + i * 4 + 4) : u16(b, lo + i * 2 + 2);
if (z > a) hasOutline[i] = 1; else empty++;
}
info.emptyGlyphs = empty; info.outlineGlyphs = n - empty;
}
}

var cmap = readCmap(b, T.cmap), chars = [];
cmap.forEach(function (gid, cp) {
if (gid >= info.numGlyphs && info.numGlyphs) return;
var blank = cp === 0x20 || cp === 0xA0 || cp === 0x09 || cp === 0x0D || cp === 0x0A || (cp >= 0x2000 && cp <= 0x200F) || cp === 0x202F || cp === 0x3000 || cp === 0xFEFF || cp === 0xAD;
if (hasOutline && !hasOutline[gid] && !blank) return; /* mapped, but nothing to draw */
chars.push(cp);
});
chars.sort(function (x, y) { return x - y; });
info.chars = chars;
return info;
};

/* ---------- full font or subset ---------- */

var UPPER = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', LOWER = 'abcdefghijklmnopqrstuvwxyz', DIGITS = '0123456789';

FX.verdict = function (info, eot) {
var v = { level: 'full', reasons: [], missing: '', latin: 0 };
if (!info) { v.level = 'unknown'; return v; }
var have = {};
info.chars.forEach(function (c) { if (c < 128) have[String.fromCharCode(c)] = 1; });
function miss(set) { return set.split('').filter(function (c) { return !have[c]; }).join(''); }
var mU = miss(UPPER), mL = miss(LOWER), mD = miss(DIGITS);
v.missing = mU + mL + mD;
v.latin = 62 - v.missing.length;

var hard = false;
if (eot && eot.subset) { hard = true; v.reasons.push('The font is marked as a subset in its header'); }
if (info.emptyGlyphs != null && info.numGlyphs >= 20 && info.emptyGlyphs / info.numGlyphs > 0.25) {
hard = true;
v.reasons.push(info.emptyGlyphs + ' of ' + info.numGlyphs + ' glyphs have had their outlines removed');
}
if (v.latin > 0 && v.missing.length) {
/* A whole class missing (no lowercase at all, no digits at all) can be how
   the font was drawn. Odd letters missing cannot. */
var whole = (mU.length === 0 || mU.length === 26) && (mL.length === 0 || mL.length === 26) && (mD.length === 0 || mD.length === 10);
if (!whole) { hard = true; v.reasons.push('Missing basic characters'); }
else if (!hard) {
v.level = 'check';
v.reasons.push('Has no ' + [mU.length ? 'capitals' : '', mL.length ? 'lowercase' : '', mD.length ? 'digits' : ''].filter(Boolean).join(' or ') + ', which may be how the font was drawn');
}
}
if (!hard && v.latin === 0 && info.chars.length < 60) {
v.level = 'check';
v.reasons.push('No Latin letters and only ' + info.chars.length + ' characters, so either a small symbol font or a subset');
}
if (hard) v.level = 'subset';
return v;
};

/* ---------- embedding permission (OS/2 fsType) ---------- */
FX.permission = function (fsType) {
if (fsType == null) return { key: 'unknown', label: 'Not stated' };
if (fsType & 0x0002) return { key: 'restricted', label: 'Restricted, no embedding' };
if (fsType & 0x0008) return { key: 'editable', label: 'Editable embedding' };
if (fsType & 0x0004) return { key: 'preview', label: 'Preview and print only' };
return { key: 'installable', label: 'Installable' };
};

/* ---------- file name for a recovered font ---------- */
FX.fileName = function (font) {
var i = font.info, base = '';
if (i) base = i.names.postscript || [i.names.family, i.names.style].filter(Boolean).join('-');
if (!base) base = [font.typeface, FX.SLOT_LABEL[font.slot] || ''].filter(Boolean).join('-');
if (!base) base = (font.path || 'font').split('/').pop().replace(/\.[^.]+$/, '');
base = base.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '').replace(/\s+/g, '').slice(0, 80) || 'font';
if (font.verdict && font.verdict.level === 'subset') base += '-SUBSET';
return base + '.' + (i ? i.ext : 'ttf');
};

/* Cheap content fingerprint, to spot the same file sitting in two slots. */
function fingerprint(b) {
var h = 2166136261;
for (var i = 0; i < b.length; i++) { h ^= b[i]; h = Math.imul(h, 16777619); }
return (h >>> 0).toString(16) + ':' + b.length;
}

/* ---------- read a deck ----------
   zip is a loaded JSZip. Returns a promise of
   { fonts: [...], embedFlag, hasFonts } */
var REL_FONT = /\/relationships\/font$/;

function resolveTarget(target) {
if (target.charAt(0) === '/') return target.slice(1);
var parts = ('ppt/' + target).split('/'), out = [];
parts.forEach(function (p) { if (p === '..') out.pop(); else if (p && p !== '.') out.push(p); });
return out.join('/');
}

FX.readDeck = function (zip, opts) {
opts = opts || {};
var presFile = zip.file('ppt/presentation.xml');
if (!presFile) return Promise.reject(new Error('Not a PowerPoint deck, ppt/presentation.xml is missing'));
var relsFile = zip.file('ppt/_rels/presentation.xml.rels');
return Promise.all([presFile.async('string'), relsFile ? relsFile.async('string') : '']).then(function (r) {
var pres = r[0], rels = r[1];
var rootEl = (/<(?:\w+:)?presentation\b[^>]*>/.exec(pres) || [''])[0];
/* saveSubsetFonts is not read: PowerPoint writes it as 1 on most decks
   whether or not anything is embedded, so it says nothing about the fonts. */
var deck = {
fonts: [],
embedFlag: /^(1|true)$/.test(attr(rootEl, 'embedTrueTypeFonts') || '')
};

var relMap = {};
(rels.match(/<Relationship\b[^>]*>/g) || []).forEach(function (el) {
var id = attr(el, 'Id'), type = attr(el, 'Type') || '', target = attr(el, 'Target');
if (id && target && REL_FONT.test(type)) relMap[id] = resolveTarget(target);
});

var listed = {}, jobs = [];
var lst = /<(?:\w+:)?embeddedFontLst\b[^>]*>([\s\S]*?)<\/(?:\w+:)?embeddedFontLst>/.exec(pres);
if (lst) {
var re = /<(?:\w+:)?embeddedFont\b[^>]*>([\s\S]*?)<\/(?:\w+:)?embeddedFont>/g, m;
while ((m = re.exec(lst[1]))) {
var body = m[1];
var typeface = attr((/<(?:\w+:)?font\b[^>]*>/.exec(body) || [''])[0], 'typeface') || '';
FX.SLOTS.forEach(function (slot) {
var el = new RegExp('<(?:\\w+:)?' + slot + '\\b[^>]*>').exec(body);
if (!el) return;
var rid = attr(el[0], '(?:\\w+:)?id');
var path = rid && relMap[rid];
jobs.push({ typeface: typeface, slot: slot, rid: rid, path: path || null });
if (path) listed[path] = 1;
});
}
}
/* Font parts sitting in the package that the font list does not mention. */
zip.forEach(function (p, f) {
if (f.dir || listed[p]) return;
if (/^ppt\/fonts\/[^/]+$/i.test(p) || /\.(fntdata|odttf)$/i.test(p)) jobs.push({ typeface: '', slot: '', rid: null, path: p, orphan: true });
});

return Promise.all(jobs.map(function (j) {
var f = j.path && zip.file(j.path);
if (!f) { j.error = 'The font file is missing from the deck'; j.missingFile = true; return j; }
return f.async('uint8array').then(function (bytes) { return FX.readFont(bytes, j, opts); });
})).then(function (fonts) {
/* Same recovered file in more than one slot, e.g. a Light weight filling
   both Regular and Bold. Point the later ones at the first. */
var seen = {};
fonts.forEach(function (f) {
if (!f.sfnt) return;
var k = fingerprint(f.sfnt);
if (seen[k]) f.sameAs = seen[k]; else seen[k] = f;
});
deck.fonts = fonts;
deck.hasFonts = fonts.length > 0;
return deck;
});
});
};

/* Unwrap and inspect one font. j carries typeface/slot/path when known. */
FX.readFont = function (bytes, j, opts) {
j = j || {};
var guid = (j.rid || '') + ' ' + (j.path || '');
var u = FX.unwrap(bytes, { mtx: opts && opts.mtx, guid: guid });
j.storedSize = bytes.length;
j.container = u.container; j.eot = u.eot; j.sfnt = u.sfnt;
if (u.error) j.error = u.error;
if (u.sfnt) {
try { j.info = FX.inspect(u.sfnt); }
catch (e) { j.error = 'Recovered, but the font tables could not be read'; }
}
j.verdict = FX.verdict(j.info || null, u.eot);
if (!j.typeface) j.typeface = (j.info && j.info.names.family) || (u.eot && u.eot.familyName) || '';
j.fileName = FX.fileName(j);
return j;
};

/* ---------- deck copy with the embedded fonts removed ----------
   Works on the loaded JSZip in place and resolves to it, so the caller
   picks the output type. Load a fresh zip if the original is still needed. */
FX.stripDeck = function (zip) {
var P = 'ppt/presentation.xml', R = 'ppt/_rels/presentation.xml.rels', C = '[Content_Types].xml';
var presFile = zip.file(P);
if (!presFile) return Promise.reject(new Error('Not a PowerPoint deck'));
return Promise.all([presFile.async('string'), zip.file(R) ? zip.file(R).async('string') : null, zip.file(C) ? zip.file(C).async('string') : null]).then(function (r) {
var pres = r[0], rels = r[1], types = r[2], removed = [];

pres = pres.replace(/<(?:\w+:)?embeddedFontLst\b[^>]*>[\s\S]*?<\/(?:\w+:)?embeddedFontLst>/, '')
.replace(/<(?:\w+:)?embeddedFontLst\b[^>]*\/>/, '');
/* Stop PowerPoint embedding them again on the next save. */
pres = pres.replace(/<(?:\w+:)?presentation\b[^>]*>/, function (el) {
return el.replace(/\s+embedTrueTypeFonts\s*=\s*("[^"]*"|'[^']*')/, '');
});
zip.file(P, pres, { createFolders: false });

if (rels != null) {
rels = rels.replace(/<Relationship\b[^>]*?(?:\/>|>\s*<\/Relationship>)/g, function (el) {
if (!REL_FONT.test(attr(el, 'Type') || '')) return el;
var t = attr(el, 'Target');
if (t && removed.indexOf(resolveTarget(t)) < 0) removed.push(resolveTarget(t));
return '';
});
zip.file(R, rels, { createFolders: false });
}
zip.forEach(function (p, f) { if (!f.dir && /^ppt\/fonts\//i.test(p) && removed.indexOf(p) < 0) removed.push(p); });
removed.forEach(function (p) { zip.remove(p); });
if (zip.folder(/^ppt\/fonts\/$/).length) zip.remove('ppt/fonts');

if (types != null) {
types = types.replace(/<Override\b[^>]*?(?:\/>|>\s*<\/Override>)/g, function (el) {
var pn = attr(el, 'PartName') || '';
return removed.indexOf(pn.replace(/^\//, '')) >= 0 ? '' : el;
});
zip.file(C, types, { createFolders: false });
}
return { zip: zip, removed: removed };
});
};

FX.isSfnt = isSfnt;
FX.isEot = isEot;
return FX;
});
