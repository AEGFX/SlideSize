/* deck-compare.js — the engine behind Deck Compare
 *
 * Reads two .pptx files (already opened with JSZip) and works out what
 * changed between them, slide by slide. No DOM work here, so the same
 * file runs in the page and in node for the checks.
 *
 *   DeckCompare.parse(zip, { name, size, parseXml })   -> Promise<deck>
 *   DeckCompare.compare(deckA, deckB)                  -> report
 *   DeckCompare.wordDiff(a, b)                         -> [[op, text], ...]
 *
 * How slides are paired
 *   Every slide has a permanent ID in ppt/presentation.xml (p:sldId id)
 *   that survives reordering, so that is tried first. IDs restart at 256
 *   in every new deck though, so two unrelated decks share IDs by chance.
 *   If most ID pairs look alike the decks share a history and IDs are
 *   trusted outright. Otherwise an ID pair has to look reasonably alike to
 *   count. Whatever is left is paired on content (words, media, layout)
 *   and marked as a probable match.
 *
 *   Moved slides are the ones outside the longest run that kept its
 *   order, so one inserted slide does not flag everything after it.
 */
(function (root, factory) {
if (typeof module === 'object' && module.exports) module.exports = factory();
else root.DeckCompare = factory();
})(typeof self !== 'undefined' ? self : this, function () {
'use strict';

var R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
var EMU_CM = 360000;
var GEOM_TOL = 0.02 * EMU_CM;   /* ignore nudges under 0.2 mm */

/* ---------- small XML helpers (localName based, prefix agnostic) ---------- */
function kids(el, name) {
var out = [];
if (!el) return out;
for (var c = el.firstChild; c; c = c.nextSibling) {
if (c.nodeType === 1 && (!name || c.localName === name)) out.push(c);
}
return out;
}
function kid(el, name) {
if (!el) return null;
for (var c = el.firstChild; c; c = c.nextSibling) {
if (c.nodeType === 1 && c.localName === name) return c;
}
return null;
}
function path(el) {
for (var i = 1; i < arguments.length && el; i++) el = kid(el, arguments[i]);
return el || null;
}
function desc(el, name, out) {
out = out || [];
if (!el) return out;
for (var c = el.firstChild; c; c = c.nextSibling) {
if (c.nodeType !== 1) continue;
if (c.localName === name) out.push(c);
desc(c, name, out);
}
return out;
}
function firstDesc(el, name) {
if (!el) return null;
for (var c = el.firstChild; c; c = c.nextSibling) {
if (c.nodeType !== 1) continue;
if (c.localName === name) return c;
var d = firstDesc(c, name);
if (d) return d;
}
return null;
}
function at(el, name) {
if (!el) return null;
var v = el.getAttribute(name);
return v === '' || v == null ? null : v;
}
function rAt(el, name) {
if (!el) return null;
var v = el.getAttributeNS ? el.getAttributeNS(R_NS, name) : null;
if (!v) v = el.getAttribute('r:' + name);
return v || null;
}
function num(v, d) { var n = parseFloat(v); return isNaN(n) ? (d || 0) : n; }

/* canonical signature of an element, ignoring attribute order and
   save-to-save noise attributes */
var NOISE_ATTR = { dirty: 1, err: 1, lang: 1, altLang: 1, smtClean: 1, smtId: 1, noProof: 1, bmk: 1, spc_ignore: 1 };
function canon(el) {
if (!el) return '';
if (el.nodeType === 3) { var t = el.nodeValue.replace(/\s+/g, ' ').trim(); return t ? '"' + t + '"' : ''; }
if (el.nodeType !== 1) return '';
var a = [];
for (var i = 0; i < el.attributes.length; i++) {
var x = el.attributes[i];
if (NOISE_ATTR[x.localName] || /^xmlns/.test(x.name)) continue;
a.push(x.localName + '=' + x.value);
}
a.sort();
var s = '<' + el.localName + (a.length ? ' ' + a.join(' ') : '') + '>';
for (var c = el.firstChild; c; c = c.nextSibling) s += canon(c);
return s + '</>';
}

/* ---------- paths and relationships ---------- */
function dirOf(p) { var i = p.lastIndexOf('/'); return i < 0 ? '' : p.slice(0, i + 1); }
function resolve(base, target) {
if (target.charAt(0) === '/') return target.slice(1);
var out = [];
(dirOf(base) + target).split('/').forEach(function (s) {
if (s === '..') out.pop();
else if (s !== '.' && s !== '') out.push(s);
});
return out.join('/');
}
function relsPathOf(p) { var d = dirOf(p); return d + '_rels/' + p.slice(d.length) + '.rels'; }
function relKind(type) { var i = (type || '').lastIndexOf('/'); return i < 0 ? type : type.slice(i + 1); }

/* ---------- hashing ---------- */
function hex(buf) {
var b = new Uint8Array(buf), s = '';
for (var i = 0; i < b.length; i++) s += (b[i] < 16 ? '0' : '') + b[i].toString(16);
return s;
}
function fnv(bytes) {
/* fallback when crypto.subtle is missing (plain http pages) */
var h1 = 0x811c9dc5, h2 = 0x01000193 ^ bytes.length;
for (var i = 0; i < bytes.length; i++) {
h1 ^= bytes[i]; h1 = Math.imul(h1, 16777619);
h2 ^= bytes[(bytes.length - 1 - i)]; h2 = Math.imul(h2, 2246822519);
}
return ('00000000' + (h1 >>> 0).toString(16)).slice(-8) + ('00000000' + (h2 >>> 0).toString(16)).slice(-8) + bytes.length.toString(16);
}
function digest(bytes) {
var c = (typeof crypto !== 'undefined' && crypto.subtle) ? crypto : null;
if (!c) return Promise.resolve(fnv(bytes));
return c.subtle.digest('SHA-1', bytes).then(function (b) { return hex(b).slice(0, 20); }, function () { return fnv(bytes); });
}

/* ---------- deck context ---------- */
function Ctx(zip, parseXml) {
this.zip = zip;
this.parseXml = parseXml;
this.xmlCache = {};
this.relsCache = {};
this.hashCache = {};
this.sizeCache = {};
}
Ctx.prototype.text = function (p) {
var f = this.zip.file(p);
return f ? f.async('string') : Promise.resolve(null);
};
Ctx.prototype.xml = function (p) {
if (!this.xmlCache[p]) {
var self = this;
this.xmlCache[p] = this.text(p).then(function (s) {
if (s == null) return null;
try {
var d = self.parseXml(s);
if (!d || !d.documentElement || d.getElementsByTagName('parsererror').length) return null;
return d;
} catch (e) { return null; }
});
}
return this.xmlCache[p];
};
Ctx.prototype.rels = function (p) {
if (!this.relsCache[p]) {
var self = this;
this.relsCache[p] = this.xml(relsPathOf(p)).then(function (d) {
var map = {};
if (!d) return map;
kids(d.documentElement, 'Relationship').forEach(function (r) {
var ext = at(r, 'TargetMode') === 'External';
var tgt = at(r, 'Target') || '';
map[at(r, 'Id')] = { type: relKind(at(r, 'Type')), external: ext, target: ext ? tgt : resolve(p, tgt) };
});
return map;
});
}
return this.relsCache[p];
};
Ctx.prototype.hash = function (p) {
if (!this.hashCache[p]) {
var self = this, f = this.zip.file(p);
/* the zip directory already holds a CRC-32 and size for every file, which
   is plenty to tell media apart and saves inflating a 2 GB video */
var d = f && f._data;
if (d && typeof d.crc32 === 'number' && d.uncompressedSize > 0) {
this.sizeCache[p] = d.uncompressedSize;
this.hashCache[p] = Promise.resolve('crc' + (d.crc32 >>> 0).toString(16) + ':' + d.uncompressedSize.toString(16));
return this.hashCache[p];
}
this.hashCache[p] = f ?f.async('uint8array').then(function (b) { self.sizeCache[p] = b.length; return digest(b); }) : Promise.resolve(null);
}
return this.hashCache[p];
};
/* hash of an XML part with noise stripped, for charts and SmartArt where
   a resave rewrites ids but nothing visible */
Ctx.prototype.xmlHash = function (p) {
var key = 'x:' + p;
if (!this.hashCache[key]) {
this.hashCache[key] = this.xml(p).then(function (d) { return d ? fnv(utf8(canon(d.documentElement))) : null; });
}
return this.hashCache[key];
};
function utf8(s) {
if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(s);
var out = []; for (var i = 0; i < s.length; i++) out.push(s.charCodeAt(i) & 255); return out;
}
function firstRel(rels, kind) {
for (var k in rels) if (rels[k].type === kind) return rels[k];
return null;
}

/* ---------- text and formatting ---------- */
function clean(s) { return String(s || '').replace(/ /g, ' ').replace(/[ \t]+/g, ' '); }

function runFmt(rPr, themeFonts) {
if (!rPr) return '';
var parts = [];
['sz', 'b', 'i', 'u', 'strike', 'baseline', 'cap'].forEach(function (k) { var v = at(rPr, k); if (v != null) parts.push(k + v); });
var latin = kid(rPr, 'latin');
if (latin) parts.push('f' + fontName(at(latin, 'typeface'), themeFonts));
var fill = kid(rPr, 'solidFill');
if (fill) parts.push('c' + colourOf(fill));
return parts.join(',');
}
function colourOf(fill) {
var c = fill && fill.firstChild;
while (c && c.nodeType !== 1) c = c.nextSibling;
if (!c) return '';
var v = at(c, 'val') || at(c, 'lastClr') || '';
var mods = kids(c).map(function (m) { return m.localName + (at(m, 'val') || ''); }).join('');
return (c.localName === 'srgbClr' ? '#' + v : v) + (mods ? '~' + mods : '');
}
function fontName(tf, themeFonts) {
if (!tf) return '';
if (tf.charAt(0) === '+' && themeFonts) {
if (/^\+mj/.test(tf)) return themeFonts.major || 'Theme headings';
if (/^\+mn/.test(tf)) return themeFonts.minor || 'Theme body';
}
return tf;
}

/* reads a txBody into text plus a per-character formatting signature.
   Per character so run splits from spellcheck or resaves do not count. */
function readText(txBody, themeFonts, fonts) {
var paras = [], fmt = [];
kids(txBody, 'p').forEach(function (p) {
var line = '', lf = [];
var pPr = kid(p, 'pPr');
var pSig = pPr ? ['algn', 'lvl', 'indent', 'marL'].map(function (k) { return at(pPr, k) || ''; }).join('|') +
(kid(pPr, 'buNone') ? 'n' : kid(pPr, 'buChar') ? 'c' + at(kid(pPr, 'buChar'), 'char') : kid(pPr, 'buAutoNum') ? 'a' : '') : '';
kids(p).forEach(function (r) {
var t = '', rPr = kid(r, 'rPr');
if (r.localName === 'r') t = textOf(kid(r, 't'));
else if (r.localName === 'br') t = '\n';
else if (r.localName === 'fld') {
/* slide number fields change when a slide moves, which is not an edit */
t = /slidenum/i.test(at(r, 'type') || '') ? '#' : textOf(kid(r, 't'));
}
else return;
t = clean(t);
if (rPr) collectFonts(rPr, themeFonts, fonts);
var f = runFmt(rPr, themeFonts);
line += t;
for (var i = 0; i < t.length; i++) lf.push(f);
});
paras.push(line);
fmt.push(pSig + ':' + compress(lf));
});
var text = paras.join('\n').replace(/[ ]+\n/g, '\n').replace(/\n+$/, '').trim();
return { text: text, fmt: fmt.join('\n') };
}
function textOf(el) { return el ? (el.textContent || '') : ''; }
function compress(arr) {
var out = [], last = null, n = 0;
arr.forEach(function (v) { if (v === last) n++; else { if (n) out.push(n + '*' + last); last = v; n = 1; } });
if (n) out.push(n + '*' + last);
return out.join(';');
}
function collectFonts(rPr, themeFonts, fonts) {
['latin', 'ea', 'cs', 'sym'].forEach(function (k) {
var e = kid(rPr, k);
var tf = e && at(e, 'typeface');
if (tf) { var n = fontName(tf, themeFonts); if (n) fonts[n] = 1; }
});
}

/* distinct formatting values across a shape, used to describe what changed */
function fmtValues(fmt) {
var v = { size: {}, font: {}, colour: {}, bold: {}, italic: {}, underline: {} };
(fmt || '').split(/[\n;]/).forEach(function (chunk) {
var star = chunk.indexOf('*');
if (star < 0) return;
chunk.slice(star + 1).split(',').forEach(function (p) {
if (/^sz/.test(p)) v.size[(num(p.slice(2)) / 100) + ' pt'] = 1;
else if (/^f/.test(p)) v.font[p.slice(1)] = 1;
else if (/^c/.test(p)) v.colour[p.slice(1).split('~')[0]] = 1;
else if (/^b/.test(p)) v.bold[p.slice(1) === '1' ? 'bold' : 'not bold'] = 1;
else if (/^i/.test(p)) v.italic[p.slice(1) === '1' ? 'italic' : 'not italic'] = 1;
else if (/^u/.test(p)) v.underline[p.slice(1) === 'none' ? 'no underline' : 'underline'] = 1;
});
});
return v;
}

/* ---------- geometry with group transforms ---------- */
function identity() { return { ox: 0, oy: 0, sx: 1, sy: 1 }; }
function applyXf(t, x, y, w, h) { return { x: t.ox + x * t.sx, y: t.oy + y * t.sy, w: w * t.sx, h: h * t.sy }; }
function readXfrm(xfrm) {
if (!xfrm) return null;
var off = kid(xfrm, 'off'), ext = kid(xfrm, 'ext');
return {
x: num(at(off, 'x')), y: num(at(off, 'y')),
w: num(at(ext, 'cx')), h: num(at(ext, 'cy')),
rot: num(at(xfrm, 'rot')) / 60000,
flip: (at(xfrm, 'flipH') === '1' ? 'H' : '') + (at(xfrm, 'flipV') === '1' ? 'V' : ''),
chOff: kid(xfrm, 'chOff'), chExt: kid(xfrm, 'chExt')
};
}

/* ---------- slide parsing ---------- */
async function readTheme(ctx, layoutPath) {
var key = 'theme:' + layoutPath;
if (ctx.hashCache[key]) return ctx.hashCache[key];
ctx.hashCache[key] = (async function () {
var fonts = { major: '', minor: '' }, master = '';
if (!layoutPath) return { fonts: fonts, master: master };
var lr = await ctx.rels(layoutPath);
var m = firstRel(lr, 'slideMaster');
if (!m) return { fonts: fonts, master: master };
master = m.target;
var mr = await ctx.rels(m.target);
var th = firstRel(mr, 'theme');
var td = th ? await ctx.xml(th.target) : null;
if (td) {
var maj = firstDesc(td.documentElement, 'majorFont'), min = firstDesc(td.documentElement, 'minorFont');
fonts.major = at(kid(maj, 'latin'), 'typeface') || '';
fonts.minor = at(kid(min, 'latin'), 'typeface') || '';
}
return { fonts: fonts, master: master };
})();
return ctx.hashCache[key];
}

async function parseSlide(ctx, slidePath, sldId, index) {
var doc = await ctx.xml(slidePath);
var rels = await ctx.rels(slidePath);
var s = {
id: sldId, index: index, num: index + 1, path: slidePath,
hidden: false, layout: '', layoutPath: '', master: '',
shapes: [], notes: '', transition: null, anims: 0, bg: '', fonts: {}, title: ''
};
if (!doc) { s.broken = true; return s; }
var sld = doc.documentElement;
s.hidden = at(sld, 'show') === '0';

var lay = firstRel(rels, 'slideLayout');
var themeFonts = null;
if (lay) {
s.layoutPath = lay.target;
var ld = await ctx.xml(lay.target);
var lc = ld && kid(ld.documentElement, 'cSld');
s.layout = (lc && at(lc, 'name')) || lay.target.replace(/^.*\//, '');
var th = await readTheme(ctx, lay.target);
themeFonts = th.fonts; s.master = th.master;
}

var cSld = kid(sld, 'cSld');
var bg = kid(cSld, 'bg');
if (bg) {
var bgSig = canon(bg);
var blips = desc(bg, 'blip');
for (var b = 0; b < blips.length; b++) {
var rel = rels[rAt(blips[b], 'embed')];
if (rel && !rel.external) bgSig += '#' + (await ctx.hash(rel.target));
}
s.bg = fnv(utf8(bgSig));
}

/* timing first, so video settings can be attached to their shapes */
var media = readTiming(kid(sld, 'timing'));
s.anims = media.anims;

var shapes = [];
await walkTree(ctx, kid(cSld, 'spTree'), identity(), rels, themeFonts, s.fonts, shapes, media.bySpid, '');
s.shapes = shapes;

/* transition, which can sit inside mc:AlternateContent */
var tr = null;
for (var c = sld.firstChild; c; c = c.nextSibling) {
if (c.nodeType !== 1) continue;
if (c.localName === 'transition') { tr = c; break; }
if (c.localName === 'AlternateContent') {
var ch = kid(c, 'Choice') || kid(c, 'Fallback');
var t2 = ch && kid(ch, 'transition');
if (t2) { tr = t2; break; }
}
}
if (tr) {
var effect = kids(tr).filter(function (e) { return e.localName !== 'sndAc' && e.localName !== 'extLst'; })[0];
var dur = at(tr, 'dur') || (tr.getAttribute('p14:dur')) || null;
s.transition = {
effect: effect ? effect.localName : 'none',
detail: effect ? canon(effect) : '',
speed: dur ? (num(dur) / 1000) + ' s' : (at(tr, 'spd') || ''),
advClick: at(tr, 'advClick') !== '0',
advTm: at(tr, 'advTm') != null ? num(at(tr, 'advTm')) / 1000 : null
};
}

/* notes */
var nr = firstRel(rels, 'notesSlide');
if (nr) {
var nd = await ctx.xml(nr.target);
if (nd) {
var parts = [];
desc(nd.documentElement, 'sp').forEach(function (sp) {
var ph = firstDesc(kid(sp, 'nvSpPr'), 'ph');
if (ph && at(ph, 'type') === 'body') parts.push(readText(kid(sp, 'txBody'), themeFonts, {}).text);
});
s.notes = parts.join('\n').trim();
}
}

/* title, for the report */
var tShape = shapes.filter(function (x) { return x.ph === 'title' || x.ph === 'ctrTitle'; })[0] ||
shapes.filter(function (x) { return x.text; })[0];
s.title = tShape ? tShape.text.split('\n')[0].slice(0, 120) : '';
return s;
}

function readTiming(timing) {
var out = { anims: 0, bySpid: {} };
if (!timing) return out;
function rec(spid) { return out.bySpid[spid] || (out.bySpid[spid] = {}); }
desc(timing, 'cTn').forEach(function (cTn) {
var pc = at(cTn, 'presetClass');
if (pc && pc !== 'mediacall' && pc !== 'verb') out.anims++;
});
['video', 'audio'].forEach(function (k) {
desc(timing, k).forEach(function (v) {
var node = kid(v, 'cMediaNode');
var tgt = firstDesc(node, 'spTgt');
var spid = tgt && at(tgt, 'spid');
if (!spid) return;
var r = rec(spid);
var vol = at(node, 'vol');
r.volume = vol != null ? Math.round(num(vol) / 1000) + '%' : '100%';
r.muted = at(node, 'mute') === '1';
var cTn = kid(node, 'cTn');
r.loop = !!(cTn && at(cTn, 'repeatCount') === 'indefinite');
r.fullScreen = at(v, 'fullScrn') === '1';
r.hideWhenStopped = at(node, 'showWhenStopped') === '0';
r.acrossSlides = num(at(node, 'numSld'), 1) > 1;
});
});
desc(timing, 'cmd').forEach(function (cmd) {
var c = at(cmd, 'cmd') || '';
if (!/^playFrom|^togglePause/.test(c)) return;
var tgt = firstDesc(cmd, 'spTgt');
var spid = tgt && at(tgt, 'spid');
if (!spid) return;
/* the trigger is the nodeType on the effect's own cTn, walking up */
var trig = null, p = cmd.parentNode;
while (p && p.nodeType === 1 && !trig) {
if (p.localName === 'cTn' && at(p, 'nodeType')) trig = at(p, 'nodeType');
if (p.localName === 'seq' || p.localName === 'timing') break;
p = p.parentNode;
}
var r = rec(spid);
if (/^playFrom/.test(c) && !r.start) {
r.start = trig === 'clickEffect' ? 'On click' : trig === 'afterEffect' || trig === 'withEffect' ? 'Automatically' : trig === 'interactiveSeq' ? 'When clicked on' : 'In sequence';
}
});
return out;
}

async function walkTree(ctx, tree, xf, rels, themeFonts, fonts, out, media, group) {
var list = kids(tree);
for (var i = 0; i < list.length; i++) {
var el = list[i], k = el.localName;
if (k === 'AlternateContent') {
var choice = kid(el, 'Choice') || kid(el, 'Fallback');
if (choice) await walkTree(ctx, choice, xf, rels, themeFonts, fonts, out, media, group);
continue;
}
if (['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp', 'contentPart'].indexOf(k) < 0) continue;
var nv = kids(el).filter(function (c) { return /^nv/.test(c.localName); })[0];
var cNvPr = kid(nv, 'cNvPr');
var nvPr = kid(nv, 'nvPr');
var ph = kid(nvPr, 'ph');
var shp = {
sid: at(cNvPr, 'id') || '', name: at(cNvPr, 'name') || k, alt: at(cNvPr, 'descr') || '',
hidden: at(cNvPr, 'hidden') === '1',
kind: 'shape', ph: ph ? (at(ph, 'type') || 'body') : '', phIdx: ph ? (at(ph, 'idx') || '') : '',
group: group, text: '', fmt: '', x: 0, y: 0, w: 0, h: 0, rot: 0, flip: '', hasGeom: false,
mediaHash: null, mediaName: '', link: '', posterHash: null, crop: '', partHash: null, fill: '', video: null
};
var spPr = kid(el, 'spPr') || kid(el, 'grpSpPr');
var xfrm = k === 'graphicFrame' ? kid(el, 'xfrm') : kid(spPr, 'xfrm');
var g = readXfrm(xfrm);
if (g) {
var p = applyXf(xf, g.x, g.y, g.w, g.h);
shp.x = p.x; shp.y = p.y; shp.w = p.w; shp.h = p.h; shp.rot = g.rot; shp.flip = g.flip; shp.hasGeom = true;
}

if (k === 'grpSp') {
/* children use the group's own coordinate space */
var t = xf;
if (g && g.chOff && g.chExt) {
var cx = num(at(g.chExt, 'cx')) || 1, cy = num(at(g.chExt, 'cy')) || 1;
var sx = g.w / cx, sy = g.h / cy;
t = { ox: xf.ox + (g.x - num(at(g.chOff, 'x')) * sx) * xf.sx, oy: xf.oy + (g.y - num(at(g.chOff, 'y')) * sy) * xf.sy, sx: xf.sx * sx, sy: xf.sy * sy };
}
await walkTree(ctx, el, t, rels, themeFonts, fonts, out, media, (group ? group + ' / ' : '') + shp.name);
continue;
}

if (spPr) {
var geom = kid(spPr, 'prstGeom');
var f = kid(spPr, 'solidFill') || kid(spPr, 'gradFill') || kid(spPr, 'blipFill') || kid(spPr, 'noFill');
shp.fill = (geom ? at(geom, 'prst') : '') + '|' + (f ? (f.localName === 'solidFill' ? colourOf(f) : f.localName) : '') + '|' + canon(kid(spPr, 'ln'));
}

var tx = kid(el, 'txBody');
if (tx) { var rt = readText(tx, themeFonts, fonts); shp.text = rt.text; shp.fmt = rt.fmt; }

if (k === 'cxnSp') shp.kind = 'connector';
if (k === 'pic') {
shp.kind = 'picture';
var blipFill = kid(el, 'blipFill');
var blip = kid(blipFill, 'blip');
var emb = blip && rels[rAt(blip, 'embed')];
if (emb && !emb.external) { shp.mediaHash = await ctx.hash(emb.target); shp.mediaName = emb.target.replace(/^.*\//, ''); }
else if (blip && rels[rAt(blip, 'link')]) { shp.link = rels[rAt(blip, 'link')].target; shp.mediaHash = 'link:' + shp.link; }
var src = kid(blipFill, 'srcRect');
if (src) shp.crop = ['l', 't', 'r', 'b'].map(function (x) { return Math.round(num(at(src, x)) / 1000); }).join(',');
/* video and audio live on a p:pic */
var vf = kid(nvPr, 'videoFile'), af = kid(nvPr, 'audioFile');
var p14 = firstDesc(kid(nvPr, 'extLst'), 'media');
if (vf || af || p14) {
shp.kind = vf ? 'video' : af ? 'audio' : (/^audio/.test(String(emb && emb.target)) ? 'audio' : 'video');
shp.posterHash = shp.mediaHash;
shp.mediaHash = null; shp.mediaName = '';
var mrel = (p14 && rels[rAt(p14, 'embed')]) || null;
var lrel = (p14 && rels[rAt(p14, 'link')]) || ((vf || af) && rels[rAt(vf || af, 'link')]) || null;
if (mrel && !mrel.external) { shp.mediaHash = await ctx.hash(mrel.target); shp.mediaName = mrel.target.replace(/^.*\//, ''); }
else if (lrel && lrel.external) { shp.link = lrel.target; shp.mediaHash = 'link:' + lrel.target; shp.mediaName = decodeURIComponent(lrel.target.replace(/^.*[\/\\]/, '')); }
else if (lrel) { shp.mediaHash = await ctx.hash(lrel.target); shp.mediaName = lrel.target.replace(/^.*\//, ''); }
var v = media[shp.sid] || {};
var trim = firstDesc(p14, 'trim');
shp.video = {
start: v.start || 'No play trigger',
muted: !!v.muted, volume: v.volume || '100%', loop: !!v.loop, fullScreen: !!v.fullScreen,
hideWhenStopped: !!v.hideWhenStopped, acrossSlides: !!v.acrossSlides,
trimStart: trim && at(trim, 'st') ? num(at(trim, 'st')) / 1000 : 0,
trimEnd: trim && at(trim, 'end') ? num(at(trim, 'end')) / 1000 : 0
};
}
}
if (k === 'graphicFrame') {
var gd = path(el, 'graphic', 'graphicData');
var uri = at(gd, 'uri') || '';
var tbl = kid(gd, 'tbl');
if (tbl) {
shp.kind = 'table';
var rows = [], tf = [];
kids(tbl, 'tr').forEach(function (tr) {
var cells = [];
kids(tr, 'tc').forEach(function (tc) {
var r = readText(kid(tc, 'txBody'), themeFonts, fonts);
cells.push(r.text.replace(/\n/g, ' '));
tf.push(r.fmt);
});
rows.push(cells.join(' | '));
});
shp.text = rows.join('\n'); shp.fmt = tf.join('\n');
} else if (/chart/.test(uri)) {
shp.kind = 'chart';
var cr = rels[rAt(firstDesc(gd, 'chart'), 'id')];
if (cr) shp.partHash = await ctx.xmlHash(cr.target);
} else if (/diagram/.test(uri)) {
shp.kind = 'SmartArt';
var ri = firstDesc(gd, 'relIds');
var dm = ri && rels[rAt(ri, 'dm')];
if (dm) {
shp.partHash = await ctx.xmlHash(dm.target);
var dd = await ctx.xml(dm.target);
if (dd) shp.text = desc(dd.documentElement, 't').map(textOf).filter(Boolean).join('\n');
}
} else if (/ole/.test(uri)) {
shp.kind = 'embedded object';
var ob = firstDesc(gd, 'oleObj');
var orel = ob && rels[rAt(ob, 'id')];
if (orel && !orel.external) shp.partHash = await ctx.hash(orel.target);
} else {
shp.kind = 'graphic';
shp.partHash = fnv(utf8(canon(gd)));
}
}
if (k === 'contentPart') shp.kind = 'ink';
out.push(shp);
}
}

/* ---------- deck parsing ---------- */
async function parse(zip, opts) {
opts = opts || {};
var parseXml = opts.parseXml || function (s) { return new DOMParser().parseFromString(s, 'application/xml'); };
var ctx = new Ctx(zip, parseXml);
var deck = { name: opts.name || 'deck.pptx', size: opts.size || 0, slides: [], width: 0, height: 0, embeddedFonts: [], layouts: {}, savedBy: '', savedAt: '', app: '', warnings: [] };
var pres = await ctx.xml('ppt/presentation.xml');
if (!pres) throw new Error(deck.name + ' is not a readable .pptx (no presentation.xml). Older .ppt files need saving as .pptx first.');
var pe = pres.documentElement;
var sz = kid(pe, 'sldSz');
deck.width = num(at(sz, 'cx')); deck.height = num(at(sz, 'cy'));
deck.embeddedFonts = kids(kid(pe, 'embeddedFontLst'), 'embeddedFont').map(function (f) { return at(kid(f, 'font'), 'typeface'); }).filter(Boolean).sort();
var prels = await ctx.rels('ppt/presentation.xml');
var ids = kids(kid(pe, 'sldIdLst'), 'sldId');
for (var i = 0; i < ids.length; i++) {
var rel = prels[rAt(ids[i], 'id')];
if (!rel) { deck.warnings.push('Slide ' + (i + 1) + ' points at a missing part'); continue; }
var s = await parseSlide(ctx, rel.target, at(ids[i], 'id'), deck.slides.length);
deck.slides.push(s);
if (s.layout) deck.layouts[s.layout] = (deck.layouts[s.layout] || 0) + 1;
}
var core = await ctx.xml('docProps/core.xml');
if (core) {
deck.savedBy = textOf(firstDesc(core.documentElement, 'lastModifiedBy')).trim();
deck.savedAt = textOf(firstDesc(core.documentElement, 'modified')).trim();
}
var app = await ctx.xml('docProps/app.xml');
if (app) {
var an = textOf(firstDesc(app.documentElement, 'Application')).trim();
var av = textOf(firstDesc(app.documentElement, 'AppVersion')).trim();
deck.app = an + (av ? ' ' + av : '');
}
/* every slide master, layout and their themes, so template swaps show
   up. Notes and handout masters have themes too, those are left out. */
var masters = {};
var names = Object.keys(zip.files).filter(function (n) { return /^ppt\/(slideMasters|slideLayouts)\/[^\/]+\.xml$/.test(n); });
for (var sm = 0; sm < names.length; sm++) {
if (!/slideMasters/.test(names[sm])) continue;
var thr = firstRel(await ctx.rels(names[sm]), 'theme');
if (thr && names.indexOf(thr.target) < 0) names.push(thr.target);
}
names.sort();
for (var m = 0; m < names.length; m++) masters[names[m]] = await ctx.xmlHash(names[m]);
deck.masters = masters;
deck.masterSig = fnv(utf8(names.map(function (n) { return masters[n]; }).join('|')));
return deck;
}

/* ---------- similarity and pairing ---------- */
var WORD_RE = /[A-Za-z0-9À-ɏͰ-ϿЀ-ӿ぀-ヿ一-鿿]+/g;
function words(s) { return (String(s).toLowerCase().match(WORD_RE) || []); }
function slideTokens(s) {
if (s._tok) return s._tok;
var w = {}, m = {};
s.shapes.forEach(function (x) {
words(x.text).forEach(function (t) { w[t] = 1; });
if (x.mediaHash) m[x.mediaHash] = 1;
if (x.partHash) m[x.partHash] = 1;
});
s._tok = { w: w, m: m, nw: Object.keys(w).length, nm: Object.keys(m).length };
return s._tok;
}
function jac(a, b) {
var ka = Object.keys(a), kb = Object.keys(b);
if (!ka.length && !kb.length) return 1;
var inter = 0;
ka.forEach(function (k) { if (b[k]) inter++; });
return inter / (ka.length + kb.length - inter);
}
function similarity(a, b) {
var ta = slideTokens(a), tb = slideTokens(b);
var hasText = ta.nw || tb.nw, hasMedia = ta.nm || tb.nm;
var sw = jac(ta.w, tb.w), sm = jac(ta.m, tb.m), sl = a.layout === b.layout ? 1 : 0;
if (hasText && hasMedia) return 0.6 * sw + 0.3 * sm + 0.1 * sl;
if (hasText) return 0.85 * sw + 0.15 * sl;
if (hasMedia) return 0.85 * sm + 0.15 * sl;
/* nothing to go on but the layout and shape count */
var ca = a.shapes.length, cb = b.shapes.length;
return 0.5 * sl + 0.3 * (Math.min(ca, cb) + 1) / (Math.max(ca, cb) + 1);
}

function lis(seq) {
/* indices (into seq) of one longest increasing subsequence */
var tails = [], prev = new Array(seq.length);
for (var i = 0; i < seq.length; i++) {
var lo = 0, hi = tails.length;
while (lo < hi) { var mid = (lo + hi) >> 1; if (seq[tails[mid]] < seq[i]) lo = mid + 1; else hi = mid; }
prev[i] = lo > 0 ? tails[lo - 1] : -1;
tails[lo] = i;
}
var out = {}, k = tails.length ? tails[tails.length - 1] : -1;
while (k >= 0) { out[k] = 1; k = prev[k]; }
return out;
}

function pairSlides(A, B) {
var pairs = [], usedA = {}, usedB = {};
var byIdB = {};
B.slides.forEach(function (s) { if (s.id) byIdB[s.id] = s; });
var idPairs = [];
A.slides.forEach(function (a) { var b = a.id && byIdB[a.id]; if (b) idPairs.push([a, b, similarity(a, b)]); });
var alike = idPairs.filter(function (p) { return p[2] >= 0.5; }).length;
var trustIds = idPairs.length > 0 && alike / idPairs.length >= 0.4;
idPairs.forEach(function (p) {
if (trustIds || p[2] >= 0.35) {
pairs.push({ a: p[0], b: p[1], match: 'id', sim: p[2] });
usedA[p[0].index] = 1; usedB[p[1].index] = 1;
}
});
/* content match whatever is left */
var cand = [];
A.slides.forEach(function (a) {
if (usedA[a.index]) return;
B.slides.forEach(function (b) {
if (usedB[b.index]) return;
var sim = similarity(a, b);
var ta = slideTokens(a), tb = slideTokens(b);
if (!ta.nw && !tb.nw && !ta.nm && !tb.nm) return;   /* blank slides only pair on ID */
if (sim >= 0.45) cand.push([a, b, sim]);
});
});
cand.sort(function (x, y) {
return y[2] - x[2] || Math.abs(x[0].index - x[1].index) - Math.abs(y[0].index - y[1].index);
});
cand.forEach(function (c) {
if (usedA[c[0].index] || usedB[c[1].index]) return;
pairs.push({ a: c[0], b: c[1], match: c[2] >= 0.999 ? 'content-exact' : 'content', sim: c[2] });
usedA[c[0].index] = 1; usedB[c[1].index] = 1;
});
return { pairs: pairs, usedA: usedA, usedB: usedB, trustIds: trustIds, idPairs: idPairs.length };
}

/* ---------- word diff ---------- */
/* words keep their inner dots, commas, slashes and colons, so 8.1m, 1,200,
   14/01 and 09:30 change as one token */
function tokens(s) { return String(s).match(/\s+|[A-Za-z0-9À-ɏ'’]+(?:[.,:\/\-][A-Za-z0-9À-ɏ]+)*|[^\sA-Za-z0-9À-ɏ'’]/g) || []; }
function wordDiff(a, b) {
var x = tokens(a), y = tokens(b);
/* trim the common head and tail so the DP only sees the edited middle */
var pre = 0;
while (pre < x.length && pre < y.length && x[pre] === y[pre]) pre++;
var suf = 0;
while (suf < x.length - pre && suf < y.length - pre && x[x.length - 1 - suf] === y[y.length - 1 - suf]) suf++;
var xm = x.slice(pre, x.length - suf), ym = y.slice(pre, y.length - suf);
var ops = [];
if (pre) ops.push(['=', x.slice(0, pre).join('')]);
if (xm.length * ym.length > 1500000) {
if (xm.length) ops.push(['-', xm.join('')]);
if (ym.length) ops.push(['+', ym.join('')]);
} else {
var n = xm.length, m = ym.length, w = m + 1;
var L = new Uint16Array((n + 1) * w);
for (var i = n - 1; i >= 0; i--) for (var j = m - 1; j >= 0; j--) {
L[i * w + j] = xm[i] === ym[j] ? L[(i + 1) * w + j + 1] + 1 : Math.max(L[(i + 1) * w + j], L[i * w + j + 1]);
}
var i2 = 0, j2 = 0;
while (i2 < n && j2 < m) {
if (xm[i2] === ym[j2]) { push(ops, '=', xm[i2]); i2++; j2++; }
else if (L[(i2 + 1) * w + j2] >= L[i2 * w + j2 + 1]) { push(ops, '-', xm[i2]); i2++; }
else { push(ops, '+', ym[j2]); j2++; }
}
while (i2 < n) push(ops, '-', xm[i2++]);
while (j2 < m) push(ops, '+', ym[j2++]);
}
if (suf) push(ops, '=', x.slice(x.length - suf).join(''));
return tidy(ops);
}
function push(ops, op, t) { var l = ops[ops.length - 1]; if (l && l[0] === op) l[1] += t; else ops.push([op, t]); }
/* fold whitespace-only equal runs sitting between edits into the edits,
   so "a b c" -> "x y z" reads as one change and not three */
function tidy(ops) {
var out = [];
for (var i = 0; i < ops.length; i++) {
var o = ops[i];
if (o[0] === '=' && /^\s+$/.test(o[1]) && out.length && out[out.length - 1][0] !== '=' && ops[i + 1] && ops[i + 1][0] !== '=') {
var dels = [], adds = [];
/* pull back the edit block just before */
while (out.length && out[out.length - 1][0] !== '=') { var p = out.pop(); (p[0] === '-' ? dels : adds).unshift(p[1]); }
dels.push(o[1]); adds.push(o[1]);
while (ops[i + 1] && ops[i + 1][0] !== '=') { i++; (ops[i][0] === '-' ? dels : adds).push(ops[i][1]); }
if (dels.join('')) out.push(['-', dels.join('')]);
if (adds.join('')) out.push(['+', adds.join('')]);
continue;
}
var l = out[out.length - 1];
if (l && l[0] === o[0]) l[1] += o[1]; else out.push([o[0], o[1]]);
}
/* order each edit block as delete then insert */
var res = [];
for (var k = 0; k < out.length; k++) {
if (out[k][0] === '=') { res.push(out[k]); continue; }
var d = '', a = '';
while (k < out.length && out[k][0] !== '=') { if (out[k][0] === '-') d += out[k][1]; else a += out[k][1]; k++; }
k--;
if (d) res.push(['-', d]);
if (a) res.push(['+', a]);
}
return res;
}
function diffSummary(ops) {
var edits = ops.filter(function (o) { return o[0] !== '='; });
var dels = ops.filter(function (o) { return o[0] === '-'; }), adds = ops.filter(function (o) { return o[0] === '+'; });
function q(s) { s = s.replace(/\s+/g, ' ').trim(); return '"' + (s.length > 60 ? s.slice(0, 57) + '...' : s) + '"'; }
if (edits.length <= 2 && dels.length <= 1 && adds.length <= 1) {
var d = dels[0] && dels[0][1].trim(), a = adds[0] && adds[0][1].trim();
if (d && a) return q(d) + ' → ' + q(a);
if (a) return 'added ' + q(a);
if (d) return 'removed ' + q(d);
}
var nd = 0, na = 0;
dels.forEach(function (o) { nd += words(o[1]).length; });
adds.forEach(function (o) { na += words(o[1]).length; });
return 'edited, ' + (nd ? nd + ' word' + (nd === 1 ? '' : 's') + ' out' : '') + (nd && na ? ', ' : '') + (na ? na + ' word' + (na === 1 ? '' : 's') + ' in' : '') + (!nd && !na ? 'punctuation or spacing only' : '');
}

/* ---------- shape pairing and diff ---------- */
function shapeLabel(x) {
var base;
if (x.ph === 'title' || x.ph === 'ctrTitle') base = 'Title';
else if (x.ph === 'subTitle') base = 'Subtitle';
else if (x.ph && x.ph !== 'body' && x.ph !== 'obj') base = ({ dt: 'Date', ftr: 'Footer', sldNum: 'Slide number', pic: 'Picture placeholder', chart: 'Chart placeholder', tbl: 'Table placeholder', media: 'Media placeholder' })[x.ph] || x.name;
else base = x.name;
return (x.group ? x.group + ' / ' : '') + base;
}
function kindWord(x) {
return ({ shape: x.text ? 'text box' : 'shape', picture: 'picture', video: 'video', audio: 'audio', table: 'table', chart: 'chart', SmartArt: 'SmartArt', connector: 'line', 'embedded object': 'embedded object', graphic: 'graphic', ink: 'ink' })[x.kind] || x.kind;
}
function snippet(t, n) { t = String(t || '').replace(/\s+/g, ' ').trim(); n = n || 50; return t.length > n ? t.slice(0, n - 3) + '...' : t; }
function cm(v) { return (Math.round(v / EMU_CM * 100) / 100).toFixed(2).replace(/\.?0+$/, '') ; }
function cmd(v) { var s = cm(Math.abs(v)); return s === '' ? '0' : s; }

function pairShapes(sa, sb) {
var pairs = [], ua = {}, ub = {};
function take(i, j, how) { pairs.push([sa[i], sb[j], how]); ua[i] = 1; ub[j] = 1; }
function compatible(x, y) {
if (x.kind === y.kind) return true;
var visual = { picture: 1, video: 1, audio: 1 };
return visual[x.kind] && visual[y.kind];
}
/* 1. same shape id. Placeholders are also checked by their slot. */
sa.forEach(function (x, i) {
for (var j = 0; j < sb.length; j++) {
if (ub[j]) continue;
var y = sb[j];
if (x.sid && x.sid === y.sid && compatible(x, y) && x.group === y.group) { take(i, j, 'id'); return; }
}
});
/* 2. same placeholder slot */
sa.forEach(function (x, i) {
if (ua[i] || !x.ph) return;
for (var j = 0; j < sb.length; j++) {
if (ub[j]) continue;
if (sb[j].ph === x.ph && sb[j].phIdx === x.phIdx) { take(i, j, 'ph'); return; }
}
});
/* 3. same media, or alike text */
var cand = [];
sa.forEach(function (x, i) {
if (ua[i]) return;
sb.forEach(function (y, j) {
if (ub[j] || !compatible(x, y)) return;
var s = 0;
if (x.mediaHash && x.mediaHash === y.mediaHash) s = 1;
else if (x.partHash && x.partHash === y.partHash) s = 1;
else if (x.text && y.text) {
var wa = {}, wb = {};
words(x.text).forEach(function (t) { wa[t] = 1; });
words(y.text).forEach(function (t) { wb[t] = 1; });
s = jac(wa, wb) * 0.9;
}
if (x.name === y.name) s += 0.15;
if (s >= 0.5) cand.push([i, j, s]);
});
});
cand.sort(function (p, q) { return q[2] - p[2]; });
cand.forEach(function (c) { if (!ua[c[0]] && !ub[c[1]]) take(c[0], c[1], 'content'); });
/* 4. one leftover each of the same kind and name, e.g. a replaced picture */
sa.forEach(function (x, i) {
if (ua[i]) return;
var same = [];
sb.forEach(function (y, j) { if (!ub[j] && compatible(x, y) && y.name === x.name) same.push(j); });
if (same.length === 1) take(i, same[0], 'name');
});
return {
pairs: pairs,
removed: sa.filter(function (x, i) { return !ua[i]; }),
added: sb.filter(function (x, j) { return !ub[j]; })
};
}

function C(cat, text, extra) {
var c = { cat: cat, text: text };
if (extra) for (var k in extra) c[k] = extra[k];
return c;
}

function diffShape(x, y, out) {
var label = shapeLabel(y);
/* text */
if (x.text !== y.text) {
if (x.text.replace(/\s+/g, ' ').trim() === y.text.replace(/\s+/g, ' ').trim()) {
out.push(C('format', label + ' line breaks or spacing changed'));
} else {
var ops = wordDiff(x.text, y.text);
out.push(C('text', label + ' text ' + diffSummary(ops), { diff: ops }));
}
} else if (x.fmt !== y.fmt && x.text) {
out.push(C('format', label + ' formatting changed' + describeFmt(x.fmt, y.fmt)));
}
/* media */
if (x.kind !== y.kind) out.push(C('media', label + ' changed from ' + kindWord(x) + ' to ' + kindWord(y)));
if (x.mediaHash !== y.mediaHash && (x.mediaHash || y.mediaHash)) {
var what = y.kind === 'picture' ? 'picture replaced' : y.kind + ' file replaced';
var detail = '';
if (x.link || y.link) detail = ' (linked, ' + (x.mediaName || '?') + ' → ' + (y.mediaName || '?') + ')';
else if (y.kind !== 'picture' && x.mediaName && y.mediaName && x.mediaName !== y.mediaName) detail = ' (' + x.mediaName + ' → ' + y.mediaName + ')';
out.push(C('media', label + ' ' + what + detail, { major: y.kind !== 'picture' }));
}
if (x.posterHash !== y.posterHash && x.mediaHash === y.mediaHash && (x.kind === 'video')) {
out.push(C('media', label + ' poster frame changed'));
}
if (x.crop !== y.crop && y.kind === 'picture') out.push(C('layout', label + ' crop changed'));
if (x.partHash !== y.partHash && (x.partHash || y.partHash)) {
out.push(C('media', label + ' ' + kindWord(y) + ' content changed', { major: false }));
}
/* playback */
if (x.video && y.video) {
var v = x.video, w = y.video, p = [];
if (v.start !== w.start) p.push('starts ' + lower(v.start) + ' → ' + lower(w.start));
if (v.muted !== w.muted) p.push(w.muted ? 'now muted' : 'now unmuted');
if (v.volume !== w.volume) p.push('volume ' + v.volume + ' → ' + w.volume);
if (v.loop !== w.loop) p.push(w.loop ? 'now loops' : 'no longer loops');
if (v.fullScreen !== w.fullScreen) p.push(w.fullScreen ? 'now plays full screen' : 'no longer full screen');
if (v.hideWhenStopped !== w.hideWhenStopped) p.push(w.hideWhenStopped ? 'now hides when not playing' : 'now shows when not playing');
if (v.acrossSlides !== w.acrossSlides) p.push(w.acrossSlides ? 'now plays across slides' : 'no longer plays across slides');
if (v.trimStart !== w.trimStart || v.trimEnd !== w.trimEnd) p.push('trim ' + trimText(v) + ' → ' + trimText(w));
if (p.length) out.push(C('playback', label + ' ' + p.join(', '), { major: true }));
}
/* geometry */
if (x.hasGeom && y.hasGeom) {
var dx = y.x - x.x, dy = y.y - x.y, dw = y.w - x.w, dh = y.h - x.h;
var g = [];
if (Math.abs(dx) >= GEOM_TOL || Math.abs(dy) >= GEOM_TOL) {
var mv = [];
if (Math.abs(dx) >= GEOM_TOL) mv.push(cmd(dx) + ' cm ' + (dx > 0 ? 'right' : 'left'));
if (Math.abs(dy) >= GEOM_TOL) mv.push(cmd(dy) + ' cm ' + (dy > 0 ? 'down' : 'up'));
g.push('moved ' + mv.join(', '));
}
if (Math.abs(dw) >= GEOM_TOL || Math.abs(dh) >= GEOM_TOL) g.push('resized ' + cm(x.w) + ' × ' + cm(x.h) + ' → ' + cm(y.w) + ' × ' + cm(y.h) + ' cm');
if (Math.abs(x.rot - y.rot) >= 0.5) g.push('rotated ' + round1(x.rot) + '° → ' + round1(y.rot) + '°');
if (x.flip !== y.flip) g.push('flipped');
if (g.length) out.push(C('layout', label + ' ' + g.join(', ')));
}
if (x.fill !== y.fill && x.kind === 'shape') out.push(C('format', label + ' fill, outline or shape type changed'));
if (x.hidden !== y.hidden) out.push(C('layout', label + (y.hidden ? ' hidden in the selection pane' : ' unhidden in the selection pane')));
if (x.alt !== y.alt && (x.alt || y.alt)) out.push(C('notes', label + ' alt text changed'));
}
function lower(s) { return String(s).charAt(0).toLowerCase() + String(s).slice(1); }
function round1(v) { return Math.round(v * 10) / 10; }
function trimText(v) { return (v.trimStart || v.trimEnd) ? fmtSec(v.trimStart) + ' to -' + fmtSec(v.trimEnd) : 'none'; }
function fmtSec(s) { return (Math.round(s * 100) / 100) + ' s'; }
function describeFmt(a, b) {
var va = fmtValues(a), vb = fmtValues(b), bits = [];
var names = { size: 'size', font: 'font', colour: 'colour', bold: 'bold', italic: 'italic', underline: 'underline' };
Object.keys(names).forEach(function (k) {
var ka = Object.keys(va[k]).sort().join(', '), kb = Object.keys(vb[k]).sort().join(', ');
if (ka !== kb) bits.push(names[k] + ' ' + (ka || 'from layout') + ' → ' + (kb || 'from layout'));
});
return bits.length ? ' (' + bits.join('; ') + ')' : ' (alignment, bullets or indents)';
}

function diffSlide(a, b) {
var out = [];
if (a.hidden !== b.hidden) out.push(C('playback', b.hidden ? 'Slide is now hidden' : 'Slide is no longer hidden', { major: true }));
if (a.layout !== b.layout) out.push(C('layout', 'Layout ' + q(a.layout) + ' → ' + q(b.layout)));
if (a.bg !== b.bg) out.push(C('format', a.bg && b.bg ? 'Background changed' : b.bg ? 'Background override added' : 'Background override removed'));
var ta = a.transition, tb = b.transition;
/* an empty p:transition and no transition at all are the same thing */
var tsa = ta && ta.effect !== 'none' ? ta.effect + ' ' + ta.speed + ' ' + ta.detail : 'none';
var tsb = tb && tb.effect !== 'none' ? tb.effect + ' ' + tb.speed + ' ' + tb.detail : 'none';
if (tsa !== tsb) out.push(C('playback', 'Transition ' + transText(ta) + ' → ' + transText(tb)));
var aa = advText(ta), ab = advText(tb);
if (aa !== ab) out.push(C('playback', 'Advance ' + aa + ' → ' + ab, { major: true }));
if (a.anims !== b.anims) out.push(C('playback', 'Animations ' + a.anims + ' → ' + b.anims));
var sp = pairShapes(a.shapes, b.shapes);
sp.removed.forEach(function (x) {
out.push(C(catFor(x), 'Removed ' + kindWord(x) + ' ' + q(shapeLabel(x)) + (x.text ? ', ' + q(snippet(x.text)) : x.mediaName ? ', ' + x.mediaName : ''), { major: x.kind === 'video' || x.kind === 'audio' }));
});
sp.added.forEach(function (y) {
out.push(C(catFor(y), 'Added ' + kindWord(y) + ' ' + q(shapeLabel(y)) + (y.text ? ', ' + q(snippet(y.text)) : y.mediaName ? ', ' + y.mediaName : ''), { major: y.kind === 'video' || y.kind === 'audio' }));
});
sp.pairs.forEach(function (p) { diffShape(p[0], p[1], out); });
if (a.notes !== b.notes) {
if (a.notes.replace(/\s+/g, ' ').trim() === b.notes.replace(/\s+/g, ' ').trim()) out.push(C('notes', 'Speaker notes spacing changed'));
else {
var ops = wordDiff(a.notes, b.notes);
out.push(C('notes', 'Speaker notes ' + (!a.notes ? 'added' : !b.notes ? 'removed' : diffSummary(ops)), { diff: ops }));
}
}
var fa = Object.keys(a.fonts).sort(), fb = Object.keys(b.fonts).sort();
var fAdd = fb.filter(function (f) { return fa.indexOf(f) < 0; }), fDel = fa.filter(function (f) { return fb.indexOf(f) < 0; });
if (fAdd.length) out.push(C('format', 'Font' + (fAdd.length > 1 ? 's' : '') + ' now used ' + fAdd.join(', ')));
if (fDel.length) out.push(C('format', 'Font' + (fDel.length > 1 ? 's' : '') + ' no longer used ' + fDel.join(', ')));
return out;
}
function catFor(x) { return x.kind === 'video' || x.kind === 'audio' || x.kind === 'picture' ? 'media' : x.text ? 'text' : 'layout'; }
function q(s) { return '"' + s + '"'; }
function transText(t) {
if (!t || t.effect === 'none') return 'none';
return t.effect.replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase() + (t.speed ? ' ' + t.speed : '');
}
function advText(t) {
if (!t) return 'on click';
var bits = [];
if (t.advClick) bits.push('on click');
if (t.advTm != null) bits.push('after ' + t.advTm + ' s');
return bits.join(' or ') || 'never (no click, no timer)';
}

/* ---------- deck level ---------- */
function diffDeck(A, B) {
var out = [];
if (Math.abs(A.width - B.width) >= GEOM_TOL / 2 || Math.abs(A.height - B.height) >= GEOM_TOL / 2) {
out.push(C('layout', 'Slide size ' + cm(A.width) + ' × ' + cm(A.height) + ' → ' + cm(B.width) + ' × ' + cm(B.height) + ' cm', { major: true }));
}
var ea = A.embeddedFonts.join(', '), eb = B.embeddedFonts.join(', ');
if (ea !== eb) out.push(C('format', 'Embedded fonts ' + (ea || 'none') + ' → ' + (eb || 'none'), { major: true }));
if (A.masterSig !== B.masterSig) {
var ka = Object.keys(A.masters), kb = Object.keys(B.masters);
var changed = ka.filter(function (k) { return B.masters[k] && B.masters[k] !== A.masters[k]; }).length;
var added = kb.filter(function (k) { return !A.masters[k]; }).length;
var removed = ka.filter(function (k) { return !B.masters[k]; }).length;
var bits = [];
if (changed) bits.push(changed + ' changed');
if (added) bits.push(added + ' added');
if (removed) bits.push(removed + ' removed');
out.push(C('format', 'Template (masters, layouts, theme) ' + bits.join(', ')));
}
var la = Object.keys(A.layouts), lb = Object.keys(B.layouts);
var lAdd = lb.filter(function (l) { return la.indexOf(l) < 0; }), lDel = la.filter(function (l) { return lb.indexOf(l) < 0; });
if (lAdd.length) out.push(C('layout', 'Layouts now in use ' + lAdd.map(q).join(', ')));
if (lDel.length) out.push(C('layout', 'Layouts no longer used ' + lDel.map(q).join(', ')));
return out;
}

/* ---------- the report ---------- */
function compare(A, B) {
var pr = pairSlides(A, B);
var rows = [];
/* moved = matched slides outside the longest run that kept its order */
var byA = pr.pairs.slice().sort(function (x, y) { return x.a.index - y.a.index; });
var keep = lis(byA.map(function (p) { return p.b.index; }));
byA.forEach(function (p, i) {
var changes = diffSlide(p.a, p.b);
rows.push({
status: changes.length ? 'changed' : 'unchanged',
match: p.match, sim: p.sim, a: p.a, b: p.b, moved: !keep[i] && p.a.index !== p.b.index,
changes: changes
});
});
A.slides.forEach(function (a) {
if (!pr.usedA[a.index]) rows.push({ status: 'deleted', a: a, b: null, moved: false, changes: [] });
});
B.slides.forEach(function (b) {
if (!pr.usedB[b.index]) rows.push({ status: 'added', a: null, b: b, moved: false, changes: [] });
});
/* order by the new deck, with deleted slides placed after whatever
   followed their predecessor in the old deck */
var bOfA = {};
pr.pairs.forEach(function (p) { bOfA[p.a.index] = p.b.index; });
rows.forEach(function (r) {
if (r.b) { r.key = r.b.index; return; }
var k = -1;
for (var i = r.a.index - 1; i >= 0; i--) if (bOfA[i] != null) { k = bOfA[i]; break; }
r.key = k + 0.5 + r.a.index / 100000;
});
rows.sort(function (x, y) { return x.key - y.key; });
rows.forEach(function (r) {
if (r.moved && r.status === 'unchanged') r.status = 'moved';
r.cats = {};
r.changes.forEach(function (c) { r.cats[c.cat] = 1; });
if (r.moved) r.cats.order = 1;
if (r.status === 'added' || r.status === 'deleted') r.cats.slides = 1;
});
var counts = { changed: 0, unchanged: 0, added: 0, deleted: 0, moved: 0, probable: 0 };
rows.forEach(function (r) {
counts[r.status === 'moved' ? 'unchanged' : r.status]++;
if (r.moved) counts.moved++;
if (r.match === 'content') counts.probable++;
});
return {
a: summary(A), b: summary(B), deck: diffDeck(A, B), rows: rows, counts: counts,
idsTrusted: pr.trustIds, idPairs: pr.idPairs
};
}
function summary(D) {
return { name: D.name, size: D.size, slides: D.slides.length, hidden: D.slides.filter(function (s) { return s.hidden; }).length, width: D.width, height: D.height, savedBy: D.savedBy, savedAt: D.savedAt, app: D.app, warnings: D.warnings };
}

/* ---------- plain text export ---------- */
function toText(rep, opts) {
opts = opts || {};
var L = [];
L.push('DECK COMPARE');
L.push('A (older)  ' + rep.a.name + '  ' + rep.a.slides + ' slides' + (rep.a.savedAt ? '  saved ' + rep.a.savedAt.replace('T', ' ').replace('Z', '') + (rep.a.savedBy ? ' by ' + rep.a.savedBy : '') : ''));
L.push('B (newer)  ' + rep.b.name + '  ' + rep.b.slides + ' slides' + (rep.b.savedAt ? '  saved ' + rep.b.savedAt.replace('T', ' ').replace('Z', '') + (rep.b.savedBy ? ' by ' + rep.b.savedBy : '') : ''));
var c = rep.counts;
L.push(c.changed + ' changed, ' + c.added + ' added, ' + c.deleted + ' deleted, ' + c.moved + ' moved, ' + c.unchanged + ' unchanged');
L.push('');
if (rep.deck.length) {
L.push('WHOLE DECK');
rep.deck.forEach(function (d) { L.push('  - ' + d.text); });
L.push('');
}
rep.rows.forEach(function (r) {
if (!opts.all && (r.status === 'unchanged')) return;
L.push(slot(r) + '  ' + r.status.toUpperCase() + (r.moved && r.status !== 'moved' ? ', MOVED' : '') + (r.match === 'content' ? ' (probable match)' : '') + '  ' + (titleOf(r) || '(untitled)'));
r.changes.forEach(function (ch) { L.push('    - ' + ch.text); });
});
return L.join('\r\n') + '\r\n';
}
function slot(r) {
if (r.a && r.b) return r.a.num === r.b.num ? 'Slide ' + r.b.num : 'Slide ' + r.a.num + ' → ' + r.b.num;
if (r.b) return 'Slide ' + r.b.num + ' (new)';
return 'Old slide ' + r.a.num;
}
function titleOf(r) { return (r.b && r.b.title) || (r.a && r.a.title) || ''; }
function toCSV(rep) {
var rows = [['A slide', 'B slide', 'Status', 'Moved', 'Match', 'Title', 'Category', 'Change']];
rep.deck.forEach(function (d) { rows.push(['', '', 'deck', '', '', '', d.cat, d.text]); });
rep.rows.forEach(function (r) {
var base = [r.a ? r.a.num : '', r.b ? r.b.num : '', r.status, r.moved ? 'yes' : '', r.match || '', titleOf(r)];
if (!r.changes.length) rows.push(base.concat(['', '']));
r.changes.forEach(function (ch) { rows.push(base.concat([ch.cat, ch.text])); });
});
return rows.map(function (row) {
return row.map(function (v) { v = String(v); return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(',');
}).join('\r\n') + '\r\n';
}

return {
parse: parse, compare: compare, wordDiff: wordDiff, diffSummary: diffSummary,
toText: toText, toCSV: toCSV, slotLabel: slot, titleOf: titleOf,
_similarity: similarity
};
});
