/* font-extractor-ui.js — page wiring for Font Extractor.
   The reading and unwrapping lives in font-extract.js. */
(function () {
'use strict';

var FX = window.FontExtract, SS = window.SS;
var esc = SS.esc, fmtBytes = SS.fmtBytes;
var results = document.getElementById('results');
var drop = document.getElementById('drop'), input = document.getElementById('file');
var faceSeq = 0;

var DECK_RE = /\.(pptx|ppsx|potx|pptm|ppsm|potm)$/i;
var LOOSE_RE = /\.(fntdata|eot|odttf|ttf|otf)$/i;

/* ---------- drop zone ---------- */
['dragenter', 'dragover'].forEach(function (ev) {
drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('drag'); });
});
['dragleave', 'drop'].forEach(function (ev) {
drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('drag'); });
});
drop.addEventListener('drop', function (e) {
if (e.dataTransfer && e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
});
input.addEventListener('change', function () {
if (input.files.length) handleFiles(input.files);
input.value = '';
});
/* A deck dropped anywhere on the page should not navigate away. */
window.addEventListener('dragover', function (e) { e.preventDefault(); });
window.addEventListener('drop', function (e) { e.preventDefault(); });

function handleFiles(list) {
if (SS.hideHandoffOverlay) SS.hideHandoffOverlay();
var files = Array.prototype.slice.call(list);
var decks = files.filter(function (f) { return DECK_RE.test(f.name); });
var loose = files.filter(function (f) { return !DECK_RE.test(f.name) && LOOSE_RE.test(f.name); });
var other = files.length - decks.length - loose.length;
if (other && !decks.length && !loose.length) {
SS.showToast(/\.ppt$/i.test(files[0].name) ? 'Old .ppt files are not supported, save as .pptx first' : 'Drop a .pptx deck or a .fntdata file');
return;
}
decks.forEach(readDeckFile);
if (loose.length) readLoose(loose);
}

function section() {
var el = document.createElement('div');
el.className = 'deck';
results.insertBefore(el, results.firstChild);
return el;
}

/* ---------- a deck ---------- */
function readDeckFile(file) {
var el = section();
el.innerHTML = '<div class="loading-note"><span class="spinner"></span>Reading ' + esc(file.name) + '</div>';
var bytes;
file.arrayBuffer().then(function (ab) {
bytes = ab;
return SS.loadJSZip();
}).then(function (JSZip) {
return JSZip.loadAsync(bytes);
}).then(function (zip) {
return FX.readDeck(zip);
}).then(function (deck) {
renderDeck(el, { name: file.name, fonts: deck.fonts, bytes: bytes, isDeck: true });
}).catch(function (e) {
var msg = (e && e.message) || String(e);
if (/encrypted|central directory|zip/i.test(msg)) msg = 'This file could not be opened. If the deck is password protected, remove the password and try again.';
el.innerHTML = '<div class="deck-bar bad"><div class="deck-top"><span class="deck-name">' + esc(file.name) + '</span></div>' +
'<div class="deck-hint warn">' + esc(msg) + '</div></div>';
});
}

/* ---------- loose font files ---------- */
function readLoose(files) {
var el = section();
Promise.all(files.map(function (f) {
return f.arrayBuffer().then(function (ab) {
return FX.readFont(new Uint8Array(ab), { typeface: '', slot: '', path: f.name, loose: true });
});
})).then(function (fonts) {
renderDeck(el, { name: files.length === 1 ? files[0].name : files.length + ' font files', fonts: fonts, isDeck: false });
});
}

/* ---------- render ---------- */
function plural(n, one, many) { return n + ' ' + (n === 1 ? one : many); }
function baseName(n) { return n.replace(/\.[^.]+$/, ''); }

function renderDeck(el, d) {
var fonts = d.fonts;
if (!fonts.length) {
el.innerHTML = '<div class="deck-bar"><div class="deck-top"><span class="deck-name">' + esc(d.name) + '</span>' +
'<span class="deck-counts"><span class="pill">No embedded fonts</span></span></div>' +
'<div class="deck-hint">This deck was saved without embedded fonts, so there is nothing to pull out. Any font it uses has to be installed on the machine that plays it.</div></div>';
return;
}
var got = fonts.filter(function (f) { return f.sfnt; });
var unique = got.filter(function (f) { return !f.sameAs; });
var nSub = fonts.filter(function (f) { return f.verdict.level === 'subset'; }).length;
var nChk = fonts.filter(function (f) { return f.verdict.level === 'check'; }).length;
var nBad = fonts.length - got.length;
var nFull = got.length - nSub - nChk;

/* Two different fonts can come out with the same name. Keep downloads apart. */
var used = {};
unique.forEach(function (f) {
var n = f.fileName, i = 2;
while (used[n.toLowerCase()]) n = f.fileName.replace(/(\.[^.]+)$/, '-' + (i++) + '$1');
used[n.toLowerCase()] = 1; f.outName = n;
});
got.forEach(function (f) { if (f.sameAs) f.outName = f.sameAs.outName; });

var h = '<div class="deck-bar' + (nSub || nBad ? ' bad' : '') + '"><div class="deck-top"><span class="deck-name">' + esc(d.name) + '</span><span class="deck-counts">' +
'<span class="pill">' + plural(fonts.length, 'embedded font', 'embedded fonts') + '</span>' +
(nFull ? '<span class="pill full">' + nFull + ' whole</span>' : '') +
(nSub ? '<span class="pill subset">' + plural(nSub, 'subset', 'subsets') + '</span>' : '') +
(nChk ? '<span class="pill check">' + nChk + ' to check</span>' : '') +
(nBad ? '<span class="pill subset">' + nBad + ' unreadable</span>' : '') +
'</span></div>';
if (nSub) {
h += '<div class="deck-hint warn"><b>' + (nSub === 1 ? 'One font is' : nSub + ' fonts are') + ' only a subset.</b> ' +
'A subset holds just the characters that were on the slides when the deck was saved. Type anything new and PowerPoint swaps in another font. ' +
'Best practice would be to get the full font from the client before the show.</div>';
}
h += '<div class="deck-btns">' +
(unique.length > 1 ? '<button class="btn primary dl-zip">Download all ' + unique.length + ' fonts (.zip)</button>' : '') +
(d.isDeck ? '<button class="btn dl-strip">Download deck without embedded fonts</button>' : '') +
'</div>' +
'<div class="deck-hint">Fonts stay under their own licence once they are out of the deck. Check you are covered before installing one on a show machine.</div>' +
'</div><div class="cards"></div>';
el.innerHTML = h;

var cards = el.querySelector('.cards');
fonts.forEach(function (f) { cards.appendChild(card(f)); });

var zipBtn = el.querySelector('.dl-zip');
if (zipBtn) zipBtn.addEventListener('click', function () {
zipBtn.disabled = true;
SS.loadJSZip().then(function (JSZip) {
var z = new JSZip();
unique.forEach(function (f) { z.file(f.outName, f.sfnt); });
return z.generateAsync({ type: 'blob', compression: 'DEFLATE' });
}).then(function (blob) {
SS.download(blob, baseName(d.name) + ' fonts.zip', 'application/zip');
}).catch(function () { SS.showToast('Could not build the zip'); }).then(function () { zipBtn.disabled = false; });
});

var stripBtn = el.querySelector('.dl-strip');
if (stripBtn) stripBtn.addEventListener('click', function () {
var old = stripBtn.textContent;
stripBtn.disabled = true; stripBtn.innerHTML = '<span class="spinner"></span>Building';
SS.loadJSZip().then(function (JSZip) {
return JSZip.loadAsync(d.bytes);
}).then(function (zip) {
return FX.stripDeck(zip);
}).then(function (r) {
return r.zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
}).then(function (blob) {
var ext = (/\.[^.]+$/.exec(d.name) || ['.pptx'])[0];
SS.download(blob, baseName(d.name) + ' (fonts removed)' + ext);
SS.showToast('Saved without embedded fonts, ' + fmtBytes(blob.size));
}).catch(function () { SS.showToast('Could not rebuild the deck'); }).then(function () {
stripBtn.disabled = false; stripBtn.textContent = old;
});
});
}

function row(k, v) { return '<div class="r"><span class="k">' + k + '</span><span class="v">' + v + '</span></div>'; }

function charList(info) {
var s = '';
for (var i = 0; i < info.chars.length && s.length < 400; i++) {
var c = info.chars[i];
if (c <= 32 || c === 0xA0 || (c >= 0x7F && c < 0xA0) || (c >= 0x2000 && c <= 0x200F)) continue;
if (c >= 0xE000 && c <= 0xF8FF) continue;
s += String.fromCodePoint(c);
}
return s;
}

function card(f) {
var v = f.verdict, info = f.info, lvl = f.sfnt ? v.level : 'failed';
var el = document.createElement('div');
el.className = 'card ' + lvl;
var pill = lvl === 'full' ? '<span class="pill full">Whole font</span>' :
lvl === 'subset' ? '<span class="pill subset">Subset</span>' :
lvl === 'check' ? '<span class="pill check">Check</span>' : '<span class="pill subset">Unreadable</span>';
var title = f.typeface || (info && info.names.family) || f.path.split('/').pop();
var slot = f.slot ? FX.SLOT_LABEL[f.slot] + ' slot' : f.orphan ? 'In the file but not in the font list' : f.loose ? esc(f.path) : '';

var h = '<div class="card-head"><div class="card-title"><span class="face">' + esc(title) + '</span>' +
(slot ? '<span class="slot">' + slot + '</span>' : '') + '</div>' + pill + '</div><div class="card-body">';

if (!f.sfnt) {
h += '<div class="meta">' + row('Stored as', esc((f.path || '').split('/').pop()) + (f.storedSize ? ' &middot; ' + fmtBytes(f.storedSize) : '')) + '</div>' +
'<div class="note bad">' + esc(f.error || 'Could not be read') + '.</div></div>';
el.innerHTML = h;
return el;
}

var n = info ? info.names : {};
var wrap = f.container === 'eot' ? (f.eot.compressed ? 'MTX compressed EOT' : f.eot.encrypted ? 'scrambled EOT' : 'EOT') :
f.container === 'obfuscated' ? 'obfuscated font' : 'plain font';
h += '<div class="sample"><div class="ph">Loading preview</div></div><div class="meta">';
if (info) {
h += row('Font', esc(n.full || [n.family, n.style].join(' ').trim() || 'No name'));
h += row('File', info.ext.toUpperCase() + ' &middot; ' + fmtBytes(f.sfnt.length) + (f.loose && f.container === 'plain' ? '' : ' &middot; was ' + fmtBytes(f.storedSize) + ' ' + wrap));
h += row('Glyphs', (info.outlineGlyphs != null ? info.outlineGlyphs + ' drawn of ' + info.numGlyphs : info.numGlyphs) + ' &middot; ' + plural(info.chars.length, 'character', 'characters'));
h += row('Embedding', esc(FX.permission(info.fsType).label));
}
h += '</div>';

if (lvl === 'subset') {
h += '<div class="note bad">' + esc(v.reasons.join('. ')) + '. Installing this will not give you the full font.</div>';
if (info) {
var cl = charList(info);
h += '<div class="chars"><span class="lbl">' + (cl ? 'Characters it holds' : 'No printable characters') + '</span>' + esc(cl) +
(v.missing ? '<span class="lbl" style="margin-top:6px">Missing from A to Z, a to z and 0 to 9</span><span class="miss">' + esc(v.missing) + '</span>' : '') + '</div>';
}
} else if (lvl === 'check') {
h += '<div class="note">' + esc(v.reasons.join('. ')) + '.</div>';
} else {
h += '<div class="note ok">Nothing points to a subset. The basic alphabet is complete and the glyphs are drawn.</div>';
}
if (f.sameAs) {
h += '<div class="note">Same file as the ' + esc(FX.SLOT_LABEL[f.sameAs.slot] || 'other') + ' slot. The deck has no separate font for this style, so PowerPoint fakes it.</div>';
}
if (f.error) h += '<div class="note">' + esc(f.error) + '.</div>';
h += '<div class="actions"><button class="btn' + (lvl === 'full' ? ' primary' : '') + ' dl-font">Download ' + esc(f.outName || f.fileName) + '</button></div></div>';
el.innerHTML = h;

el.querySelector('.dl-font').addEventListener('click', function () {
SS.download(f.sfnt, f.outName || f.fileName, info && info.format === 'otf' ? 'font/otf' : 'font/ttf');
});
preview(el.querySelector('.sample'), f);
return el;
}

/* Draw a sample in the recovered font itself. */
function preview(box, f) {
var info = f.info, text, small;
if (f.verdict.level === 'subset' && info) {
text = charList(info).slice(0, 48); small = '';
} else if (info && f.verdict.latin === 0) {
text = charList(info).slice(0, 24); small = '';
} else {
text = 'Hamburgefonstiv'; small = 'ABCDEFGHIJKLM abcdefghijklm 0123456789';
}
function fail() { box.innerHTML = '<div class="ph">No preview for this font</div>'; }
if (!window.FontFace || !text) { fail(); return; }
var fam = 'fx-face-' + (++faceSeq);
var buf = f.sfnt.buffer.slice(f.sfnt.byteOffset, f.sfnt.byteOffset + f.sfnt.byteLength);
var style = info && /italic|oblique/i.test(info.names.style) ? 'italic' : 'normal';
try {
new FontFace(fam, buf, { style: style }).load().then(function (face) {
document.fonts.add(face);
var css = 'font-family:\'' + fam + '\';font-style:' + style;
box.innerHTML = '<div class="big" style="' + css + '">' + esc(text) + '</div>' + (small ? '<div class="small" style="' + css + '">' + esc(small) + '</div>' : '');
}, fail);
} catch (e) { fail(); }
}

/* Pick up a deck dropped on the Font Extractor tile on the launchpad */
if (SS.consumeHandoff) SS.consumeHandoff(handleFiles);

/* Hook for the page checks */
window.__fontExtractorHandle = handleFiles;
})();
