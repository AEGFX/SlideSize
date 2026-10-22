/* script-format.js — Script Format for SlideSize
 *
 * Puts every sentence of a script on its own paragraph for the prompter.
 * Sentence ends are decided in context rather than at every full stop, so
 * U.S.A., Dr. Patel, 9 a.m., e.g., J. K. Rowling and No. 10 stay intact.
 * Uncertain calls are flagged for a quick check and every call can be
 * flipped with a click. Formatting from Word and PowerPoint (highlights,
 * colours, bold) is kept in the .docx export.
 */
(function(){
'use strict';
const $ = id => document.getElementById(id);
const SR = window.ScriptRead;
const esc = s => s.replace(/[&<>"]/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;' }[c]));

const state = { src: null, example: true, flips: new Map(), model: null, check: -1 };

/* ---------- sentence rules ---------- */
const TITLES = new Set(('mr mrs ms mx dr prof sr jr st mt rev revd hon gen col lt capt sgt cpl maj adm brig cmdr gov sen rep pres fr messrs mme mlle supt insp det cllr sir dame lord').split(' '));
const BEFORE_NUMBER = new Set(('no nos p pp vol vols fig figs art ch sec para approx ca c tel ext op rm apt jan feb mar apr jun jul aug sep sept oct nov dec').split(' '));
const NEVER_END = new Set(('e.g i.e vs v cf viz approx ca esp incl excl dept est fig figs vol vols ch pp ft yr yrs hr hrs min mins sec secs no nos tel ext jan feb mar apr jun jul aug sep sept oct nov dec mon tue tues wed thu thur thurs fri sat sun').split(' '));
const CAN_END = new Set(('etc inc ltd co corp plc llc al bros a.m p.m am pm u.s u.k u.s.a e.u u.n d.c').split(' '));
const STARTERS = new Set(('the a an this that these those it its i i\'m i\'ll i\'ve i\'d we we\'re we\'ll you your you\'re he she they their there here our my his her so but and or now then today tonight tomorrow yesterday please thank thanks ladies welcome let let\'s what when where why how who whose which if in on at as for with after before since also however finally first next over together every all not yes of to from by one two three each many some is are was were do does did can will would could should may might our let\'s good great well okay ok oh after during while once meanwhile again'.replace(/\s+/g, ' ')).split(' '));

const END_RE = /([.!?…]+)(["'”’)\]]*)$/;
const lead = w => w.replace(/^["'“‘(\[]+/, '');
const isUpperStart = w => /^\p{Lu}/u.test(lead(w));
const isLowerStart = w => /^\p{Ll}/u.test(lead(w));
const isDigitStart = w => /^\d/.test(lead(w));
const isStarter = w => STARTERS.has(lead(w).toLowerCase().replace(/[^\p{L}']/gu, ''));

// Decide whether a sentence ends after word w (next is the following word).
// Returns { split, sure, why } or null when this is not a candidate at all.
function decide(w, next, first){
  const cw = SR.canon(w);
  const m = END_RE.exec(cw);
  if (!m) return null;
  const punct = m[1];
  const core = cw.slice(0, cw.length - m[0].length);
  const base = core.replace(/^["'(\[]+/, '').toLowerCase();
  if (isLowerStart(next)) return { split: false, sure: true, why: 'next word starts lower case' };
  if (/[!?]/.test(punct)) return { split: true, sure: true, why: 'question or exclamation' };
  if (/\.\.\.|…/.test(punct)) return { split: isUpperStart(next), sure: false, why: 'ellipsis' };
  // plain full stop from here
  if (first && /^\d+$/.test(base)) return { split: false, sure: true, why: 'list number' };
  if (isDigitStart(next)){
    if (BEFORE_NUMBER.has(base)) return { split: false, sure: true, why: 'abbreviation before a number' };
    return { split: true, sure: !/^\d+$/.test(base), why: 'next sentence starts with a number' };
  }
  if (!isUpperStart(next) && !/^["'“‘(\[]/.test(next)) return { split: false, sure: false, why: 'no capital after the stop' };
  if (TITLES.has(base)) return { split: false, sure: true, why: 'title before a name' };
  if (NEVER_END.has(base)) return { split: false, sure: true, why: 'abbreviation' };
  const dotted = /^(\p{L}\.)+\p{L}$/u.test(base);           // u.s.a  a.m  e.u
  if (CAN_END.has(base) || dotted){
    const s = isStarter(next);
    return { split: s, sure: false, why: s ? 'abbreviation, but the next word usually starts a sentence' : 'abbreviation' };
  }
  if (/^\p{Lu}$/u.test(core.replace(/^["'(\[]+/, ''))){      // single initial, J. K. Rowling
    const s = isStarter(next);
    return { split: s, sure: false, why: s ? 'single letter, next word usually starts a sentence' : 'initial before a name' };
  }
  if (/^[\w.+-]+@[\w-]+\.[\w.]+$/.test(core) || /\.(com|org|net|co\.uk|io)$/i.test(core)) return { split: true, sure: false, why: 'web or email address' };
  return { split: true, sure: true, why: 'full stop' };
}

/* ---------- build the formatted model ---------- */
/* Speaker names and roles followed by : or ;  ANDREW:  VOICE OF GOD;  DR. PATEL:
   In capitals they are found anywhere in a paragraph. At the start of a paragraph a
   short Title Case name also counts (Sarah:  Voice of God:). */
const CAPS_WORD = /^[(\["'“]*[\p{Lu}0-9][\p{Lu}0-9.'’&\-/()]*$/u;
const hasLetter = w => /\p{L}/u.test(w);
function findLabels(words){
  const starts = new Map();   // start index -> end index
  for (let i = 0; i < words.length; i++){
    const w = SR.canon(words[i].w);
    const m = /^(.*?)[:;]$/.exec(w);
    if (!m || !m[1] || !hasLetter(m[1])) continue;
    let start = -1;
    if (CAPS_WORD.test(m[1])){
      start = i;
      // walk back over the rest of a multi word name in capitals
      while (start > 0 && i - start < 4){
        const prev = SR.canon(words[start - 1].w);
        if (words[start - 1].slide || !CAPS_WORD.test(prev) || !hasLetter(prev) || /[.!?:;,]$/.test(prev) && !/^\p{Lu}{1,4}\.$/u.test(prev)) break;
        start--;
      }
    } else if (i <= 3 && /^\p{Lu}/u.test(m[1])){
      const name = words.slice(0, i + 1).map(x => x.w).join(' ');
      const ok = words.slice(0, i).every(x => /^(\p{Lu}|of$|the$|and$|for$|de$|van$)/u.test(x.w));
      if (ok && name.length <= 30) start = 0;
    }
    if (start >= 0 && i < words.length - 1) starts.set(start, i);
  }
  return starts;
}

/* Slide cues: SLIDE 3, Slide 12a, [SLIDE 4], slides 5-7, NEXT SLIDE, SLIDE CHANGE */
const SLIDE_RE = /[\[(]?\b(?:(?:next|new|change)\s+slides?\b|slides?\s*(?:#|no\.?|number)?\s*\d+[a-z]?(?:\s*(?:-|–|to|&|and)\s*\d+[a-z]?)?|slides?\s+(?:change|up)\b)[\])]?/gi;
function markSlides(words){
  let text = '', starts = [];
  words.forEach((w, i) => { if (i) text += ' '; starts.push(text.length); text += w.w; });
  let m; SLIDE_RE.lastIndex = 0;
  while ((m = SLIDE_RE.exec(text))){
    const a = m.index, b = a + m[0].length;
    words.forEach((w, i) => { const s = starts[i], e = s + w.w.length; if (s < b && e > a) w.slide = true; });
  }
}
function build(){
  const src = state.src;
  if (!src){ state.model = null; return; }
  const join = $('optJoin').checked, speakers = $('optSpk').checked;
  const paras = SR.paragraphs(src, join);
  const model = [];   // paragraphs of lines, each line an array of words; candidates listed separately
  let cand = 0;
  paras.forEach((p, pi) => {
    const lines = [];
    let line = [];
    const words = p.words.map(w => ({ w: w.w, f: w.f }));
    markSlides(words);
    const labels = speakers ? findLabels(words) : new Map();
    const labelEnds = new Set(labels.values());
    for (let i = 0; i < words.length; i++){
      const w = words[i];
      line.push({ w: w.w, f: w.f, slide: w.slide });
      const id = 'p' + pi + 'w' + i;
      if (speakers && i < words.length - 1 && (labels.has(i + 1) || labelEnds.has(i))){
        // break before a speaker name and again straight after its colon; a click undoes it
        const why = labels.has(i + 1) ? 'speaker name follows' : 'speaker name';
        if (!state.flips.has(id)){ line[line.length - 1].cut = { id, split: true, sure: true, why, spk: true }; lines.push(line); line = []; continue; }
        line[line.length - 1].cut = { id, split: false, sure: true, why, spk: true, flipped: true };
        continue;
      }
      if (i === words.length - 1) break;
      const d = decide(w.w, words[i + 1].w, i === 0);
      if (!d) continue;
      const flipped = state.flips.has(id);
      const split = flipped ? !d.split : d.split;
      line[line.length - 1].cut = { id, split, sure: d.sure || flipped, auto: d.split, why: d.why, flipped };
      cand++;
      if (split){ lines.push(line); line = []; }
    }
    if (line.length) lines.push(line);
    model.push(lines);
  });
  state.model = model;
}

/* ---------- slide cue colour ---------- */
const slideOn = () => $('optSlide').checked;
function slideStyle(){
  const c = $('slideColour').value, hl = $('slideMode').value === 'highlight';
  if (!hl) return 'color:' + c + ';font-weight:700';
  const r = parseInt(c.slice(1,3),16), g = parseInt(c.slice(3,5),16), b = parseInt(c.slice(5,7),16);
  return 'background:' + c + ';color:' + ((r*299+g*587+b*114)/1000 > 128 ? '#111' : '#fff') + ';font-weight:700;border-radius:2px;padding:0 2px';
}

/* ---------- render ---------- */
function render(){
  build();
  const out = $('result'), stats = $('stats');
  if (!state.model){
    out.innerHTML = '<div class="empty-state"><strong>Waiting for a script</strong>Drop a file or paste text above.</div>';
    stats.innerHTML = ''; updateNav(); return;
  }
  const showHl = !!state.src.fmt;
  let html = '', n = 0, toCheck = 0, flips = 0;
  state.model.forEach((lines, pi) => {
    if (pi > 0) html += '<div class="pgap" aria-hidden="true"></div>';
    lines.forEach((line, li) => {
      n++;
      let t = '';
      line.forEach((w, wi) => {
        const st = w.slide && slideOn() ? slideStyle() : (showHl ? SR.fmtStyle(w.f) : '');
        t += (wi ? ' ' : '') + (st ? '<span class="w" style="' + st + '">' + esc(w.w) + '</span>' : esc(w.w));
        const c = w.cut;
        if (!c) return;
        if (!c.sure) toCheck++;
        if (c.flipped) flips++;
        const cls = 'cut' + (c.split ? ' is-split' : ' is-joined') + (c.sure ? '' : ' unsure') + (c.flipped ? ' flipped' : '');
        const label = c.split ? 'New line here. Click to join with the next sentence' : 'Kept together. Click to start a new line here';
        t += '<button type="button" class="' + cls + '" data-id="' + c.id + '" title="' + esc(label + ' (' + c.why + ')') + '" aria-label="' + esc(label) + '">' + (c.split ? '&crarr;' : '&middot;') + '</button>';
      });
      html += '<div class="line"><span class="n">' + n + '</span><div class="t">' + t + '</div></div>';
    });
  });
  out.innerHTML = html;
  stats.innerHTML = '<span class="chip big">' + n + ' line' + (n === 1 ? '' : 's') + '</span>'
    + '<span class="chip">' + state.model.length + ' paragraph' + (state.model.length === 1 ? '' : 's') + ' in</span>'
    + (toCheck ? '<span class="chip warn">' + toCheck + ' to check</span>' : '<span class="chip ok">Nothing to check</span>')
    + (flips ? '<span class="chip">' + flips + ' changed by you</span>' : '');
  state.check = -1;
  updateNav();
}
function updateNav(){
  const list = $('result').querySelectorAll('.cut.unsure');
  $('count').textContent = (state.check < 0 ? '–' : state.check + 1) + ' / ' + list.length;
  $('prev').disabled = $('next').disabled = !list.length;
  ['copy', 'dlTxt', 'dlDocx'].forEach(id => $(id).disabled = !state.model);
}
function go(dir){
  const list = Array.from($('result').querySelectorAll('.cut.unsure'));
  if (!list.length) return;
  list.forEach(b => b.closest('.line').classList.remove('current'));
  state.check = state.check < 0 ? (dir > 0 ? 0 : list.length - 1) : (state.check + dir + list.length) % list.length;
  const el = list[state.check];
  el.closest('.line').classList.add('current');
  el.scrollIntoView({ block: 'center', behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' });
  el.focus({ preventScroll: true });
  updateNav();
}
$('result').addEventListener('click', e => {
  const b = e.target.closest('.cut');
  if (!b) return;
  const id = b.dataset.id;
  if (state.flips.has(id)) state.flips.delete(id); else state.flips.set(id, true);
  const y = b.getBoundingClientRect().top;
  render();
  const nb = $('result').querySelector('.cut[data-id="' + id + '"]');
  if (nb) window.scrollBy(0, nb.getBoundingClientRect().top - y);
});

/* ---------- output ---------- */
function outputLines(){
  const gap = $('optGap').value, keepPara = $('optPara').checked;
  const out = [];
  state.model.forEach((lines, pi) => {
    if (pi > 0 && keepPara) out.push(null);
    lines.forEach(line => out.push(line));
  });
  // null marks an extra blank line between original paragraphs
  const res = [];
  out.forEach((l, i) => {
    if (l === null){ res.push([]); return; }
    if (res.length && gap === 'blank' && res[res.length - 1].length) res.push([]);
    res.push(l);
  });
  return res;
}
function asText(){ return outputLines().map(l => l.map(w => w.w).join(' ')).join('\n') + '\n'; }

function xmlEsc(s){ return s.replace(/[&<>"]/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;' }[c])); }
function rPr(key){
  if (!key) return '';
  if (key === 'slide'){ const c = $('slideColour').value.slice(1).toUpperCase(); return '<w:rPr><w:b/>' + ($('slideMode').value === 'highlight' ? '<w:shd w:val="clear" w:color="auto" w:fill="' + c + '"/>' : '<w:color w:val="' + c + '"/>') + '</w:rPr>'; }
  const o = SR.parseKey(key); let x = '';
  if (o.b) x += '<w:b/>';
  if (o.i) x += '<w:i/>';
  if (o.u) x += '<w:u w:val="single"/>';
  if (o.c && /^[0-9A-F]{6}$/i.test(o.c)) x += '<w:color w:val="' + o.c + '"/>';
  if (o.h){
    const name = Object.keys(SR.DOCX_HL).find(k => SR.DOCX_HL[k] === o.h.toUpperCase());
    x += name ? '<w:highlight w:val="' + name + '"/>' : '<w:shd w:val="clear" w:color="auto" w:fill="' + o.h + '"/>';
  }
  return x ? '<w:rPr>' + x + '</w:rPr>' : '';
}
async function asDocx(){
  const JSZip = await SS.loadJSZip();
  const body = outputLines().map(line => {
    if (!line.length) return '<w:p/>';
    const runs = []; let cur = null;
    line.forEach((w, i) => {
      const k = (w.slide && slideOn()) ? 'slide' : (w.f || '');
      const text = (i ? ' ' : '') + w.w;
      if (cur && cur.k === k) cur.t += text;
      else { if (cur && i) { /* space belongs to the earlier run */ } cur = { k, t: text }; runs.push(cur); }
    });
    return '<w:p>' + runs.map(r => '<w:r>' + rPr(r.k) + '<w:t xml:space="preserve">' + xmlEsc(r.t) + '</w:t></w:r>').join('') + '</w:p>';
  }).join('');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('word/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>');
  zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}
const baseName = () => (state.src && state.src.name || 'script').replace(/\.[^.]+$/, '').replace(/[^\w.-]+/g, '-');

$('copy').addEventListener('click', async () => {
  const t = asText();
  try { await navigator.clipboard.writeText(t); SS.showToast('Script copied'); }
  catch (e) {
    const ta = document.createElement('textarea'); ta.value = t; ta.style.position = 'fixed'; ta.style.opacity = '0'; document.body.appendChild(ta); ta.select();
    let ok = false; try { ok = document.execCommand('copy'); } catch (_) {}
    ta.remove(); SS.showToast(ok ? 'Script copied' : 'Copy was blocked, use the TXT download');
  }
});
$('dlTxt').addEventListener('click', () => SS.download(asText(), baseName() + '-formatted.txt', 'text/plain;charset=utf-8'));
$('dlDocx').addEventListener('click', async () => { try { SS.download(await asDocx(), baseName() + '-formatted.docx'); } catch (e) { setNotice(esc(e.message || String(e)), true); } });

/* ---------- input ---------- */
function setNotice(html, err){ $('notice').innerHTML = html ? '<div class="notice' + (err ? ' err' : '') + '">' + html + '</div>' : ''; }
function showSlot(){
  const s = state.src, nm = $('name');
  $('slot').classList.toggle('filled', !!s);
  if (s){ nm.textContent = s.name; nm.classList.remove('empty'); $('meta').textContent = s.kind + ' · ' + s.text.split(/\s+/).filter(Boolean).length.toLocaleString('en-GB') + ' words' + (s.fmt ? ' · formatting kept' : ''); }
  else { nm.textContent = 'Drop a script here'; nm.classList.add('empty'); $('meta').textContent = ''; }
}
function setSource(src){
  state.src = src; state.flips = new Map(); state.example = false;
  // Hard wrapped text (no blank lines but many short lines) reads better with lines joined
  if (src && !src.fmt){ const t = src.text; $('optJoin').checked = /\n\s*\n/.test(t); }
  else $('optJoin').checked = false;
  showSlot(); render();
}
async function loadFile(file){
  if (!file) return;
  document.body.classList.add('busy'); setNotice('Reading ' + esc(file.name) + '…');
  try { const src = await SR.extract(file); closePaste(); setNotice(''); setSource(src); }
  catch (e) { setNotice(esc(e.message || String(e)), true); }
  finally { document.body.classList.remove('busy'); }
}
function closePaste(){ $('paste').hidden = true; $('paste').value = ''; }
const slot = $('slot');
slot.addEventListener('dragover', e => { e.preventDefault(); slot.classList.add('over'); });
slot.addEventListener('dragleave', () => slot.classList.remove('over'));
slot.addEventListener('drop', e => { e.preventDefault(); e.stopPropagation(); slot.classList.remove('over'); loadFile(e.dataTransfer.files[0]); });
document.addEventListener('dragover', e => e.preventDefault());
document.addEventListener('drop', e => { e.preventDefault(); if (e.dataTransfer.files.length) loadFile(e.dataTransfer.files[0]); });
$('file').addEventListener('change', () => { loadFile($('file').files[0]); $('file').value = ''; });
$('pick').addEventListener('click', () => $('file').click());
$('pasteBtn').addEventListener('click', () => { const ta = $('paste'); ta.hidden = false; if (state.example) ta.value = ''; ta.focus(); });
let deb;
$('paste').addEventListener('input', () => { clearTimeout(deb); deb = setTimeout(() => { const v = $('paste').value; setNotice(''); setSource(v.trim() ? SR.fromText(v) : null); }, 250); });
['optJoin', 'optSpk'].forEach(id => $(id).addEventListener('change', () => { state.flips = new Map(); render(); }));
['optSlide', 'slideColour', 'slideMode'].forEach(id => $(id).addEventListener('input', () => { saveSlidePrefs(); render(); }));
function saveSlidePrefs(){ try { localStorage.setItem('ss-script-format-slide', JSON.stringify({ on: $('optSlide').checked, c: $('slideColour').value, m: $('slideMode').value })); } catch (e) {} }
try { const p = JSON.parse(localStorage.getItem('ss-script-format-slide') || 'null'); if (p){ $('optSlide').checked = !!p.on; if (/^#[0-9a-f]{6}$/i.test(p.c)) $('slideColour').value = p.c; if (p.m) $('slideMode').value = p.m; } } catch (e) {}
$('prev').addEventListener('click', () => go(-1));
$('next').addEventListener('click', () => go(1));
document.addEventListener('keydown', e => {
  if (e.target.closest('textarea,input,select')) return;
  if (e.key === 'n' || e.key === 'j'){ go(1); e.preventDefault(); }
  else if (e.key === 'p' || e.key === 'k'){ go(-1); e.preventDefault(); }
});

/* ---------- example ---------- */
const EX = `SARAH: Good evening and welcome to the launch of the new U.S.A. programme. [SLIDE 1] It's 7 p.m. here in London and 2 p.m. in New York. Tonight we'll hear from Dr. Anita Patel, our Chief Scientist. She joined us from No. 10 in 2019, e.g. working on the U.K. space strategy. Are you ready? Let's begin. NEXT SLIDE ANITA; Thank you Sarah. Slide 2 shows where we started.

VOICE OF GOD: Please welcome to the stage J. K. Rowling's favourite illustrator, Mr. Tom Hughes. He has worked with Apple Inc. and the B.B.C. for over 20 years. His new book is out on Oct. 14 at 9 a.m. The queue starts at the front doors...

SARAH: Thank you, Tom! And finally, a huge thank you to everyone at Acme Ltd. We couldn't have done it without you.`;
state.src = Object.assign(SR.fromText(EX, 'example-script.txt'), { kind: 'Example' });
showSlot(); render();
setNotice('<span class="tag">Example</span> A sample script is loaded. Dotted markers are places it chose not to break, amber ones are worth a check. Drop or paste your own to replace it.');

SS.consumeHandoff(files => loadFile(files[0]));
})();
