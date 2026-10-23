/* script-read.js — shared script reader for the prompting tools (Script Comparison, Script Format)
 *
 * window.ScriptRead
 *   extract(file)              -> promise of { name, kind, text, fmt, fmtKeys, flow, modified }
 *                                 text is plain text, \n for line breaks, \n\n between paragraphs.
 *                                 fmt is an Int32Array the same length as text giving each
 *                                 character's index into fmtKeys, or null when the format has
 *                                 no run formatting we read (TXT, PDF, RTF and so on).
 *   fromText(text, name)       -> the same shape for pasted text
 *   paragraphs(src, joinLines) -> [{ text, words: [{ w, f }] }]  f is the word's format key
 *   describeFmt(a, b)          -> "highlight yellow → green, bold on"
 *   fmtStyle(key)              -> inline CSS to show a word's highlight and weight
 *   canon(s)                   -> folds curly quotes, dashes, ellipsis and odd spaces
 *
 * Formatting is read for DOCX (direct run formatting, highlight and shading) and
 * PPTX (run properties). Text colour, highlight, bold, italic and underline.
 */
(function(){
'use strict';

const PDFJS = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/';
const loaded = {};
function loadScript(src){
  if (loaded[src]) return loaded[src];
  return loaded[src] = new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = src; s.onload = res;
    s.onerror = () => { delete loaded[src]; rej(new Error("Couldn't load a helper library. Check your internet connection and try again.")); };
    document.head.appendChild(s);
  });
}
const getZip = () => window.SS && SS.loadJSZip ? SS.loadJSZip() : loadScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js').then(() => window.JSZip);
async function getPdf(){
  await loadScript(PDFJS + 'pdf.min.js');
  await loadScript(PDFJS + 'pdf.worker.min.js');
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS + 'pdf.worker.min.js';
  return window.pdfjsLib;
}

const xml = s => new DOMParser().parseFromString(s, 'application/xml');
const byTag = (n, t) => Array.from(n.getElementsByTagName(t));
function decodeText(buf){
  let t = new TextDecoder('utf-8').decode(buf);
  if ((t.match(/�/g) || []).length > 3) t = new TextDecoder('windows-1252').decode(buf);
  return t.replace(/^﻿/, '').replace(/\r\n?/g, '\n');
}
function canon(s){
  return s.replace(/[‘’‚′]/g, "'").replace(/[“”„″]/g, '"')
    .replace(/[–—−]/g, '-').replace(/…/g, '...').replace(/[   ]/g, ' ');
}

/* ---------- colour names ---------- */
const NAMED = {
  'FFFF00':'yellow','00FF00':'bright green','00FFFF':'turquoise','FF00FF':'pink','0000FF':'blue','FF0000':'red',
  '000080':'dark blue','008080':'teal','008000':'green','800080':'violet','800000':'dark red','808000':'dark yellow',
  '808080':'grey','C0C0C0':'light grey','000000':'black','FFFFFF':'white'
};
const DOCX_HL = { yellow:'FFFF00', green:'00FF00', cyan:'00FFFF', magenta:'FF00FF', blue:'0000FF', red:'FF0000',
  darkBlue:'000080', darkCyan:'008080', darkGreen:'008000', darkMagenta:'800080', darkRed:'800000',
  darkYellow:'808000', darkGray:'808080', lightGray:'C0C0C0', black:'000000', white:'FFFFFF' };
const PALETTE = [
  ['black',0,0,0],['white',255,255,255],['grey',128,128,128],['light grey',200,200,200],['dark grey',64,64,64],
  ['red',220,30,30],['dark red',128,0,0],['orange',245,140,30],['yellow',250,230,40],['gold',200,160,40],
  ['green',40,160,60],['bright green',60,230,60],['dark green',0,90,30],['teal',0,128,128],['turquoise',40,220,220],
  ['light blue',120,180,240],['blue',30,80,220],['dark blue',0,0,128],['purple',120,50,170],['violet',150,60,200],
  ['pink',245,120,190],['magenta',230,40,200],['brown',120,70,30]
];
function nameHex(hex){
  if (!hex) return '';
  if (hex.indexOf('scheme:') === 0) return 'theme ' + hex.slice(7);
  hex = hex.toUpperCase();
  if (NAMED[hex]) return NAMED[hex];
  const r = parseInt(hex.slice(0,2),16), g = parseInt(hex.slice(2,4),16), b = parseInt(hex.slice(4,6),16);
  let best = 'colour', bd = 1e9;
  for (const [n, R, G, B] of PALETTE){ const d = (r-R)**2 + (g-G)**2 + (b-B)**2; if (d < bd){ bd = d; best = n; } }
  return best;
}

/* ---------- format keys ---------- */
// A key is a stable string, '' for plain text: "c=FF0000;h=FFFF00;b;i;u"
function fmtKey(o){
  if (!o) return '';
  const p = [];
  if (o.c) p.push('c=' + o.c);
  if (o.h) p.push('h=' + o.h);
  if (o.b) p.push('b'); if (o.i) p.push('i'); if (o.u) p.push('u');
  return p.join(';');
}
function parseKey(k){
  const o = { c:'', h:'', b:false, i:false, u:false };
  if (!k) return o;
  // a word spread over several runs has keys joined with ' + '; use the first for display
  k.split(' + ')[0].split(';').forEach(part => {
    if (part.indexOf('c=') === 0) o.c = part.slice(2);
    else if (part.indexOf('h=') === 0) o.h = part.slice(2);
    else if (part === 'b') o.b = true; else if (part === 'i') o.i = true; else if (part === 'u') o.u = true;
  });
  return o;
}
function colourChange(label, x, y){
  if (x === y) return '';
  const nx = nameHex(x), ny = nameHex(y);
  const show = (n, h) => (nx === ny && h && h.indexOf('scheme:') < 0) ? n + ' #' + h : n;
  if (!x) return label + ' added (' + show(ny, y) + ')';
  if (!y) return label + ' removed (was ' + show(nx, x) + ')';
  return label + ' ' + show(nx, x) + ' → ' + show(ny, y);
}
function describeFmt(ka, kb){
  if (ka === kb) return '';
  const a = parseKey(ka), b = parseKey(kb), out = [];
  const h = colourChange('highlight', a.h, b.h); if (h) out.push(h);
  const c = colourChange('text colour', a.c, b.c); if (c) out.push(c);
  if (a.b !== b.b) out.push(b.b ? 'bold on' : 'bold off');
  if (a.i !== b.i) out.push(b.i ? 'italic on' : 'italic off');
  if (a.u !== b.u) out.push(b.u ? 'underline on' : 'underline off');
  if (!out.length) out.push('formatting changed');
  return out.join(', ');
}
function fmtStyle(k){
  if (!k) return '';
  const o = parseKey(k), s = [];
  if (o.h){ s.push('background:#' + o.h); const r=parseInt(o.h.slice(0,2),16),g=parseInt(o.h.slice(2,4),16),b=parseInt(o.h.slice(4,6),16); s.push('color:' + ((r*299+g*587+b*114)/1000 > 128 ? '#111' : '#fff')); }
  if (o.b) s.push('font-weight:700');
  if (o.i) s.push('font-style:italic');
  if (o.u) s.push('text-decoration:underline');
  return s.join(';');
}

/* ---------- run builder ---------- */
function Builder(){ this.parts = []; this.ids = []; this.keys = ['']; this.index = { '': 0 }; this.any = false; }
Builder.prototype.add = function(t, key){
  if (!t) return;
  key = key || '';
  if (key) this.any = true;
  let id = this.index[key];
  if (id === undefined){ id = this.keys.length; this.keys.push(key); this.index[key] = id; }
  this.parts.push(t); this.ids.push([id, t.length]);
};
Builder.prototype.done = function(){
  const text = this.parts.join('');
  if (!this.any) return { text, fmt: null, fmtKeys: null };
  const fmt = new Int32Array(text.length); let p = 0;
  for (const [id, n] of this.ids){ fmt.fill(id, p, p + n); p += n; }
  return { text, fmt, fmtKeys: this.keys };
};

/* ---------- DOCX ---------- */
function attr(el, name){ return el ? el.getAttribute(name) : null; }
function onOff(el){ if (!el) return false; const v = attr(el, 'w:val'); return v === null || !/^(0|false|off|none)$/i.test(v); }
function docxRunFmt(r){
  const pr = r.getElementsByTagName('w:rPr')[0];
  if (!pr) return '';
  const o = {};
  const col = attr(pr.getElementsByTagName('w:color')[0], 'w:val');
  if (col && col !== 'auto' && col !== '000000') o.c = col.toUpperCase();
  const hl = attr(pr.getElementsByTagName('w:highlight')[0], 'w:val');
  if (hl && hl !== 'none') o.h = DOCX_HL[hl] || hl.toUpperCase();
  if (!o.h){ const fill = attr(pr.getElementsByTagName('w:shd')[0], 'w:fill'); if (fill && fill !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(fill) && fill.toUpperCase() !== 'FFFFFF') o.h = fill.toUpperCase(); }
  o.b = onOff(pr.getElementsByTagName('w:b')[0]);
  o.i = onOff(pr.getElementsByTagName('w:i')[0]);
  const u = pr.getElementsByTagName('w:u')[0]; o.u = !!u && attr(u, 'w:val') !== 'none';
  return fmtKey(o);
}
function readDocx(x){
  const body = x.getElementsByTagName('w:body')[0] || x;
  const B = new Builder(); let first = true;
  for (const p of byTag(body, 'w:p')){
    if (!first) B.add('\n\n'); first = false;
    (function walk(n, f){
      for (const c of n.childNodes){
        if (c.nodeType !== 1) continue;
        const tn = c.tagName;
        if (tn === 'w:r') walk(c, docxRunFmt(c));
        else if (tn === 'w:t') B.add(c.textContent, f);
        else if (tn === 'w:tab') B.add('\t', f);
        else if (tn === 'w:br' || tn === 'w:cr') B.add('\n', f);
        else if (tn === 'w:noBreakHyphen') B.add('-', f);
        else if (tn === 'w:del' || tn === 'w:delText' || tn === 'w:instrText' || tn === 'w:p' || tn === 'w:rPr' || tn === 'w:pPr') continue;
        else walk(c, f);
      }
    })(p, '');
  }
  return B.done();
}

/* ---------- PPTX ---------- */
function pptColour(el){
  if (!el) return '';
  const s = el.getElementsByTagName('a:srgbClr')[0]; if (s) return (s.getAttribute('val') || '').toUpperCase();
  const t = el.getElementsByTagName('a:schemeClr')[0]; if (t) return 'scheme:' + t.getAttribute('val');
  return '';
}
function pptRunFmt(r){
  const pr = r.getElementsByTagName('a:rPr')[0];
  if (!pr) return '';
  const o = {};
  const fill = Array.from(pr.childNodes).find(n => n.tagName === 'a:solidFill');
  const c = pptColour(fill); if (c && c !== '000000') o.c = c;
  const h = pptColour(pr.getElementsByTagName('a:highlight')[0]); if (h) o.h = h.indexOf('scheme:') === 0 ? 'FFFF00' : h;
  o.b = pr.getAttribute('b') === '1'; o.i = pr.getAttribute('i') === '1';
  const u = pr.getAttribute('u'); o.u = !!u && u !== 'none';
  return fmtKey(o);
}
async function readPptx(buf){
  const JSZip = await getZip(); const zip = await JSZip.loadAsync(buf);
  let slides = Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort((a, b) => +a.match(/\d+/)[0] - +b.match(/\d+/)[0]);
  try {
    const pres = xml(await zip.file('ppt/presentation.xml').async('string'));
    const rels = xml(await zip.file('ppt/_rels/presentation.xml.rels').async('string'));
    const map = {}; byTag(rels, 'Relationship').forEach(r => map[r.getAttribute('Id')] = r.getAttribute('Target'));
    const ordered = byTag(pres, 'p:sldId').map(s => 'ppt/' + String(map[s.getAttribute('r:id')]).replace(/^\//, '').replace(/^ppt\//, '')).filter(f => zip.file(f));
    if (ordered.length) slides = ordered;
  } catch (e) {}
  const B = new Builder(); let first = true;
  const sep = () => { if (!first) B.add('\n\n'); first = false; };
  const addPara = p => {
    if (!p.textContent.trim()) return;
    sep();
    for (const c of p.childNodes){
      if (c.nodeType !== 1) continue;
      if (c.tagName === 'a:r' || c.tagName === 'a:fld'){ const t = c.getElementsByTagName('a:t')[0]; if (t) B.add(t.textContent, pptRunFmt(c)); }
      else if (c.tagName === 'a:br') B.add('\n');
    }
  };
  for (let i = 0; i < slides.length; i++){
    const f = slides[i];
    const x = xml(await zip.file(f).async('string'));
    sep(); B.add('SLIDE ' + (i + 1));
    byTag(x, 'a:p').forEach(addPara);
    const relF = f.replace('slides/', 'slides/_rels/') + '.rels';
    if (!zip.file(relF)) continue;
    const r = xml(await zip.file(relF).async('string'));
    const nr = byTag(r, 'Relationship').find(e => /notesSlide$/.test(e.getAttribute('Type')));
    if (!nr) continue;
    const nf = 'ppt/' + nr.getAttribute('Target').replace(/^\.\.\//, '');
    if (!zip.file(nf)) continue;
    const nx = xml(await zip.file(nf).async('string'));
    const bodies = byTag(nx, 'p:sp').filter(sp => { const ph = sp.getElementsByTagName('p:ph')[0]; return ph && ph.getAttribute('type') === 'body'; });
    if (bodies.some(sp => sp.textContent.trim())){
      sep(); B.add('NOTES (slide ' + (i + 1) + ')');
      bodies.forEach(sp => byTag(sp, 'a:p').forEach(addPara));
    }
  }
  return B.done();
}

/* ---------- PDF ---------- */
async function readPdf(buf){
  const lib = await getPdf();
  const pdf = await lib.getDocument({ data: new Uint8Array(buf) }).promise;
  const paras = []; let cur = '';
  for (let p = 1; p <= pdf.numPages; p++){
    const page = await pdf.getPage(p); const tc = await page.getTextContent();
    const lines = []; let line = '', lastY = null, lastH = 12;
    for (const it of tc.items){
      if (!('str' in it)) continue;
      const y = it.transform[5]; const h = Math.abs(it.transform[3]) || lastH;
      if (lastY !== null && Math.abs(y - lastY) > h * 0.5){ lines.push({ t: line, y: lastY }); line = ''; }
      line += it.str; lastY = y; lastH = h;
      if (it.hasEOL){ lines.push({ t: line, y }); line = ''; lastY = null; }
    }
    if (line) lines.push({ t: line, y: lastY });
    const L = lines.filter(l => l.t.trim());
    const gaps = []; for (let i = 1; i < L.length; i++) gaps.push(Math.abs(L[i-1].y - L[i].y));
    const med = gaps.length ? gaps.slice().sort((a, b) => a - b)[Math.floor(gaps.length / 2)] : 12;
    for (let i = 0; i < L.length; i++){
      if (i > 0){ const g = Math.abs(L[i-1].y - L[i].y); if (g > med * 1.45 || g < 0.1){ if (cur.trim()) paras.push(cur); cur = ''; } }
      cur += (cur ? ' ' : '') + L[i].t.trim();
    }
    if (cur.trim() && /[.!?:"')\]]$/.test(cur.trim())){ paras.push(cur); cur = ''; }
  }
  if (cur.trim()) paras.push(cur);
  return paras.join('\n\n');
}

/* ---------- RTF, HTML, ODT, FDX ---------- */
function rtfToText(s){
  const skipDest = new Set(['fonttbl','colortbl','stylesheet','info','pict','header','footer','headerl','headerr','footerl','footerr','themedata','datastore','latentstyles','xmlnstbl','listtable','listoverridetable','rsidtbl','generator','mmathPr','object','fldinst','filetbl','revtbl','listtext','pntext','pntxtb','pntxta']);
  let out = '', i = 0, depth = 0, skipAt = -1, uc = 1, pendingSkip = 0; const n = s.length;
  while (i < n){
    const c = s[i];
    if (c === '{'){ depth++; i++; continue; }
    if (c === '}'){ if (skipAt === depth) skipAt = -1; depth--; i++; continue; }
    if (c === '\\'){
      const d = s[i+1];
      if (d === '\\' || d === '{' || d === '}'){ if (skipAt < 0) out += d; i += 2; continue; }
      if (d === '*'){ if (skipAt < 0) skipAt = depth; i += 2; continue; }
      if (d === "'"){ const h = parseInt(s.substr(i+2, 2), 16); if (pendingSkip > 0) pendingSkip--; else if (skipAt < 0) out += String.fromCharCode(h); i += 4; continue; }
      if (d === '~'){ if (skipAt < 0) out += ' '; i += 2; continue; }
      if (d === '-' || d === '_'){ if (skipAt < 0 && d === '_') out += '-'; i += 2; continue; }
      if (d === '\n' || d === '\r'){ if (skipAt < 0) out += '\n\n'; i += 2; continue; }
      const m = /^([a-zA-Z]+)(-?\d+)? ?/.exec(s.slice(i+1, i+40));
      if (!m){ i += 2; continue; }
      i += 1 + m[0].length; const w = m[1], p = m[2];
      if (skipDest.has(w)){ if (skipAt < 0) skipAt = depth; continue; }
      if (skipAt >= 0) continue;
      if (w === 'par' || w === 'sect' || w === 'page') out += '\n\n';
      else if (w === 'line') out += '\n';
      else if (w === 'tab' || w === 'cell') out += '\t';
      else if (w === 'row') out += '\n';
      else if (w === 'uc') uc = +p || 1;
      else if (w === 'u'){ let v = +p; if (v < 0) v += 65536; out += String.fromCharCode(v); pendingSkip = uc; }
      else if (w === 'emdash' || w === 'endash') out += '-';
      else if (w === 'lquote' || w === 'rquote') out += "'";
      else if (w === 'ldblquote' || w === 'rdblquote') out += '"';
      continue;
    }
    if (c === '\n' || c === '\r'){ i++; continue; }
    if (skipAt < 0){ if (pendingSkip > 0) pendingSkip--; else out += c; }
    i++;
  }
  return out;
}
function htmlToText(s){
  const doc = new DOMParser().parseFromString(s, 'text/html');
  doc.querySelectorAll('script,style,noscript,head').forEach(e => e.remove());
  const block = /^(P|DIV|H[1-6]|LI|TR|BLOCKQUOTE|SECTION|ARTICLE|HEADER|FOOTER|PRE|TABLE|UL|OL|DT|DD)$/;
  let out = '';
  (function walk(n){
    for (const c of n.childNodes){
      if (c.nodeType === 3) out += c.nodeValue.replace(/\s+/g, ' ');
      else if (c.nodeType === 1){
        if (c.tagName === 'BR'){ out += '\n'; continue; }
        const b = block.test(c.tagName); if (b) out += '\n\n'; walk(c); if (b) out += '\n\n';
      }
    }
  })(doc.body || doc);
  return out;
}
function odtText(x){
  const ps = byTag(x, '*').filter(e => e.tagName === 'text:p' || e.tagName === 'text:h');
  return ps.map(p => { let t = ''; (function w(n){ for (const c of n.childNodes){ if (c.nodeType === 3) t += c.nodeValue; else if (c.nodeType === 1){ if (c.tagName === 'text:s') t += ' '.repeat(+c.getAttribute('text:c') || 1); else if (c.tagName === 'text:tab') t += '\t'; else if (c.tagName === 'text:line-break') t += '\n'; else if (c.tagName !== 'text:note') w(c); } } })(p); return t; }).join('\n\n');
}

/* ---------- entry points ---------- */
async function extract(file){
  const name = file.name; const ext = (name.split('.').pop() || '').toLowerCase();
  const buf = await file.arrayBuffer();
  let r = null, text = '', kind = ext.toUpperCase(), flow = false;
  if (ext === 'docx' || ext === 'docm' || ext === 'dotx'){
    const JSZip = await getZip(); const z = await JSZip.loadAsync(buf);
    r = readDocx(xml(await z.file('word/document.xml').async('string'))); kind = 'Word';
  } else if (ext === 'pptx' || ext === 'pptm' || ext === 'ppsx'){ r = await readPptx(buf); kind = 'PowerPoint'; }
  else if (ext === 'pdf'){ text = await readPdf(buf); kind = 'PDF'; flow = true; }
  else if (ext === 'odt'){ const JSZip = await getZip(); const z = await JSZip.loadAsync(buf); text = odtText(xml(await z.file('content.xml').async('string'))); kind = 'OpenDocument'; }
  else if (ext === 'rtf'){ text = rtfToText(decodeText(buf)); kind = 'RTF'; }
  else if (ext === 'fdx'){
    const x = xml(decodeText(buf));
    text = byTag(x, 'Paragraph').map(p => { const ty = p.getAttribute('Type') || ''; const t = byTag(p, 'Text').map(e => e.textContent).join(''); return ty === 'Character' || ty === 'Scene Heading' ? t.toUpperCase() : t; }).join('\n\n'); kind = 'Final Draft';
  } else if (ext === 'html' || ext === 'htm'){ text = htmlToText(decodeText(buf)); kind = 'HTML'; }
  else if (ext === 'doc' || ext === 'ppt' || ext === 'pages' || ext === 'key'){
    throw new Error(name + " is an older or Apple format this page can't read. Save it as ." + (ext === 'ppt' || ext === 'key' ? 'pptx' : 'docx') + ' or PDF and drop it again.');
  } else {
    text = decodeText(buf); kind = ext ? ext.toUpperCase() : 'Text';
    if (/\0/.test(text.slice(0, 2000))) throw new Error(name + " doesn't look like a text file. Try saving it as .txt, .docx or PDF.");
  }
  if (!r) r = { text: text.replace(/\r\n?/g, '\n'), fmt: null, fmtKeys: null };
  if (!r.text.trim()) throw new Error("Couldn't find any text in " + name + ". If it's a scanned PDF it has no real text to read.");
  return { name, kind, text: r.text, fmt: r.fmt, fmtKeys: r.fmtKeys, flow, modified: file.lastModified };
}
function fromText(text, name){
  return { name: name || 'Pasted text', kind: 'Pasted', text: String(text).replace(/\r\n?/g, '\n'), fmt: null, fmtKeys: null, flow: false, modified: Date.now() };
}

/* Split into paragraphs of words, keeping each word's format key.
   joinLines (or a PDF) treats single line breaks as spaces and only blank lines as paragraph ends. */
function paragraphs(src, joinLines){
  const t = src.text, fmt = src.fmt, keys = src.fmtKeys, join = joinLines || src.flow;
  const out = []; let words = [], w = '', wf = null;
  const endWord = () => {
    if (!w) return;
    let key = '';
    if (wf){ const u = Array.from(new Set(wf)).map(i => keys[i]); key = u.length === 1 ? u[0] : u.join(' + '); }
    words.push({ w, f: key }); w = ''; wf = null;
  };
  const endPara = () => { endWord(); if (words.length) out.push({ words, text: words.map(x => x.w).join(' ') }); words = []; };
  for (let i = 0; i < t.length; i++){
    const c = t[i];
    if (c === '\n'){
      if (!join){ endPara(); continue; }
      let j = i + 1, nl = 1;
      while (j < t.length && /\s/.test(t[j])){ if (t[j] === '\n') nl++; j++; }
      if (nl >= 2){ endPara(); i = j - 1; } else endWord();
      continue;
    }
    if (/\s/.test(c)){ endWord(); continue; }
    w += c;
    if (fmt){ (wf || (wf = [])).push(fmt[i]); }
  }
  endPara();
  return out;
}

window.ScriptRead = { extract, fromText, paragraphs, describeFmt, fmtStyle, parseKey, canon, nameHex, DOCX_HL };
})();
