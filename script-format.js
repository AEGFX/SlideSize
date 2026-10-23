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
        if (words[start - 1].cue || !CAPS_WORD.test(prev) || !hasLetter(prev) || /[.!?:;,]$/.test(prev) && !/^\p{Lu}{1,4}\.$/u.test(prev)) break;
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

/* ---------- cues: slide cues, directions and speaker names ----------
   Each word can carry w.cue = 'slide' | 'dir' | 'name', coloured from the options.
   Slide cues count only when they are cues rather than part of a sentence:
   in brackets [SLIDE 1], in capitals SLIDE 2 / NEXT SLIDE, or on a line of their own.
   "Slide 2 shows where we started" is left alone. */
const SLIDE_RE = /[\[(]?\b(?:(?:next|new|change)\s+slides?\b|slides?\s*(?:#|no\.?|number)?\s*\d+[a-z]?(?:\s*(?:-|–|to|&|and)\s*\d+[a-z]?)?|slides?\s+(?:change|up)\b)[\])]?/gi;
/* Directions: anything in square brackets, parentheses that hold capitals or a stage
   direction word (pause), and stand-alone capital words like PAUSE or BEAT. */
const DIR_WORDS = 'pause|beat|beats|applause|laugh|laughs|laughter|breath|breathe|smile|smiles|music|sting|play|playout|walk|walks|walk-on|walk-off|exit|exits|enter|enters|stand|sit|look|looks|gesture|gestures|turn|turns|vt|video|cue|cues|wait|slowly|slow|emphasis|stress|ad-lib|ad lib|adlib|aside|off mic|on mic|lights|blackout|stinger|bumper|lectern|reveal|award|winner';
const PAREN_DIR_RE = new RegExp('^\\(\\s*(?:' + DIR_WORDS + ')\\b', 'i');
const CAPS_DIR_RE = new RegExp('^(?:' + DIR_WORDS.toUpperCase() + ')[.!,]*$');
function spanText(words){
  let text = '', starts = [];
  words.forEach((w, i) => { if (i) text += ' '; starts.push(text.length); text += w.w; });
  return { text, starts };
}
function markRange(words, starts, a, b, cue){
  words.forEach((w, i) => { const s = starts[i], e = s + w.w.length; if (s < b && e > a) w.cue = cue; });
}
function markCues(words){
  const { text, starts } = spanText(words);
  let m;
  // directions in brackets
  const BR = /\[[^\]]{1,80}\]|\([^)]{1,80}\)/g;
  while ((m = BR.exec(text))){
    const t = m[0];
    const inner = t.slice(1, -1);
    const isDir = t[0] === '[' || (inner === inner.toUpperCase() && /\p{Lu}/u.test(inner)) || PAREN_DIR_RE.test(t);
    if (isDir) markRange(words, starts, m.index, m.index + t.length, 'dir');
  }
  words.forEach(w => { if (!w.cue && CAPS_DIR_RE.test(w.w.replace(/^[(\[]+|[)\]]+$/g, ''))) w.cue = 'dir'; });
  // slide cues
  SLIDE_RE.lastIndex = 0;
  while ((m = SLIDE_RE.exec(text))){
    const t = m[0];
    const bracketed = /^[\[(]/.test(t) || /[\])]$/.test(t) || text[m.index - 1] === '[' || text[m.index - 1] === '(';
    const caps = /\p{Lu}/u.test(t) && t === t.toUpperCase();
    const alone = text.slice(0, m.index).trim() === '' && /^[\s.:!\-–—]*$/.test(text.slice(m.index + t.length));
    if (bracketed || caps || alone) markRange(words, starts, m.index, m.index + t.length, 'slide');
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
    markCues(words);
    const labels = findLabels(words);
    for (const [a, b] of labels) for (let k = a; k <= b; k++) words[k].cue = 'name';
    if (!speakers) labels.clear();
    const labelEnds = new Set(labels.values());
    for (let i = 0; i < words.length; i++){
      const w = words[i];
      line.push({ w: w.w, f: w.f, cue: w.cue });
      const id = 'p' + pi + 'w' + i;
      if (speakers && i < words.length - 1 && (labels.has(i + 1) || labelEnds.has(i))){
        // break before a speaker name and again straight after its colon; a click undoes it
        const why = labels.has(i + 1) ? 'speaker name follows' : 'speaker name';
        if (!state.flips.has(id)){
          line[line.length - 1].cut = { id, split: true, sure: true, why, spk: true };
          // after a name the break is a line return, not a new paragraph
          if (labelEnds.has(i)) line.soft = true;
          lines.push(line); line = []; continue;
        }
        line[line.length - 1].cut = { id, split: false, sure: true, why, spk: true, flipped: true };
        continue;
      }
      if (i === words.length - 1) break;
      let d = decide(w.w, words[i + 1].w, i === 0);
      if (!d) continue;
      // a bracketed cue after a full stop starts the next line, as in  ...9 a.m.  [PAUSE] The queue...
      if (!d.split && /^[\[(]/.test(words[i + 1].w) && words[i + 1].cue && !/[,;:]$/.test(w.w)) d = { split: true, sure: true, why: 'cue starts the next line' };
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

/* ---------- cue colours ---------- */
const CUES = ['slide', 'dir', 'name'];
const cueId = (c, part) => 'cue-' + c + '-' + part;
const cueOn = c => !!c && $(cueId(c, 'on')).checked;
const cueColour = c => $(cueId(c, 'col')).value;
const cueHl = c => $(cueId(c, 'mode')).value === 'highlight';
function inkFor(hex){ const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16); return (r*299+g*587+b*114)/1000 > 128 ? '#111' : '#fff'; }
function cueStyle(c){
  const col = cueColour(c);
  return cueHl(c) ? 'background:' + col + ';color:' + inkFor(col) + ';font-weight:700;border-radius:2px;padding:0 2px' : 'color:' + col + ';font-weight:700';
}
const CUE_KEY = 'ss-script-format-cues';
function saveCues(){
  const o = {}; CUES.forEach(c => o[c] = { on: cueOn(c), m: $(cueId(c, 'mode')).value, col: cueColour(c) });
  try { localStorage.setItem(CUE_KEY, JSON.stringify(o)); } catch (e) {}
}
function loadCues(){
  try {
    const o = JSON.parse(localStorage.getItem(CUE_KEY) || 'null'); if (!o) return;
    CUES.forEach(c => { const v = o[c]; if (!v) return; $(cueId(c, 'on')).checked = !!v.on; if (v.m) $(cueId(c, 'mode')).value = v.m; if (/^#[0-9a-f]{6}$/i.test(v.col)) $(cueId(c, 'col')).value = v.col; });
  } catch (e) {}
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
    if (pi > 0 && $('optPara').checked) html += '<div class="pgap" aria-hidden="true"></div>';
    lines.forEach((line, li) => {
      n++;
      let t = '';
      line.forEach((w, wi) => {
        const st = cueOn(w.cue) ? cueStyle(w.cue) : (showHl ? SR.fmtStyle(w.f) : '');
        t += (wi ? ' ' : '') + (st ? '<span class="w" style="' + st + '">' + esc(w.w) + '</span>' : esc(w.w));
        const c = w.cut;
        if (!c) return;
        if (!c.sure) toCheck++;
        if (c.flipped) flips++;
        const cls = 'cut' + (c.split ? ' is-split' : ' is-joined') + (c.sure ? '' : ' unsure') + (c.flipped ? ' flipped' : '');
        const label = c.split ? 'New line here. Click to join with the next sentence' : 'Kept together. Click to start a new line here';
        t += '<button type="button" class="' + cls + '" data-id="' + c.id + '" title="' + esc(label + ' (' + c.why + ')') + '" aria-label="' + esc(label) + '">' + (c.split ? '&crarr;' : '&middot;') + '</button>';
      });
      html += '<div class="line' + (line.soft ? ' soft' : '') + '"><span class="n">' + n + '</span><div class="t">' + t + '</div></div>';
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

/* ---------- output ----------
   Each output line is { words, soft }. soft means it ends in a line return
   (after a speaker name) rather than a paragraph return. Blank lines are { words: [] }. */
function outputLines(){
  const blank = $('optGap').value === 'blank', keepPara = $('optPara').checked;
  const res = [];
  state.model.forEach((lines, pi) => {
    if (pi > 0 && keepPara) res.push({ words: [] });
    lines.forEach(line => {
      const prev = res[res.length - 1];
      if (blank && prev && prev.words.length && !prev.soft) res.push({ words: [] });
      res.push({ words: line, soft: !!line.soft });
    });
  });
  return res;
}
function asText(){ return outputLines().map(l => l.words.map(w => w.w).join(' ')).join('\n') + '\n'; }

function xmlEsc(s){ return s.replace(/[&<>"]/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;' }[c])); }
function rPr(key){
  if (!key) return '';
  if (key.indexOf('cue:') === 0){ const c = key.slice(4), col = cueColour(c).slice(1).toUpperCase(); return '<w:rPr><w:b/>' + (cueHl(c) ? '<w:shd w:val="clear" w:color="auto" w:fill="' + col + '"/>' : '<w:color w:val="' + col + '"/>') + '</w:rPr>'; }
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
  const runsXml = line => {
    const runs = []; let cur = null;
    line.forEach((w, i) => {
      const k = cueOn(w.cue) ? 'cue:' + w.cue : (w.f || '');
      const text = (i ? ' ' : '') + w.w;
      if (cur && cur.k === k) cur.t += text; else { cur = { k, t: text }; runs.push(cur); }
    });
    return runs.map(r => '<w:r>' + rPr(r.k) + '<w:t xml:space="preserve">' + xmlEsc(r.t) + '</w:t></w:r>').join('');
  };
  // soft lines join the next line in the same Word paragraph with a line break (Shift+Enter)
  let body = '', open = '';
  outputLines().forEach(l => {
    if (!l.words.length){ if (open){ body += '<w:p>' + open + '</w:p>'; open = ''; } body += '<w:p/>'; return; }
    open += runsXml(l.words);
    if (l.soft) open += '<w:r><w:br/></w:r>';
    else { body += '<w:p>' + open + '</w:p>'; open = ''; }
  });
  if (open) body += '<w:p>' + open + '</w:p>';
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
loadCues();
CUES.forEach(c => ['on', 'mode', 'col'].forEach(part => $(cueId(c, part)).addEventListener('input', () => { saveCues(); render(); })));
['optGap', 'optPara'].forEach(id => $(id).addEventListener('input', render));
$('prev').addEventListener('click', () => go(-1));
$('next').addEventListener('click', () => go(1));
document.addEventListener('keydown', e => {
  if (e.target.closest('textarea,input,select')) return;
  if (e.key === 'n' || e.key === 'j'){ go(1); e.preventDefault(); }
  else if (e.key === 'p' || e.key === 'k'){ go(-1); e.preventDefault(); }
});

/* ---------- example ---------- */
const EX = `SARAH: Good evening and welcome to the launch of the new U.S.A. programme. [SLIDE 1] It's 7 p.m. here in London and 2 p.m. in New York. Tonight we'll hear from Dr. Anita Patel, our Chief Scientist. She joined us from No. 10 in 2019, e.g. working on the U.K. space strategy. Are you ready? Let's begin. NEXT SLIDE ANITA; Thank you Sarah. Slide 2 shows where we started.

VOICE OF GOD: Please welcome to the stage J. K. Rowling's favourite illustrator, Mr. Tom Hughes. He has worked with Apple Inc. and the B.B.C. for over 20 years. His new book is out on Oct. 14 at 9 a.m. [PAUSE] The queue starts at the front doors...

SARAH: Thank you, Tom! And finally, a huge thank you to everyone at Acme Ltd. We couldn't have done it without you.`;
state.src = Object.assign(SR.fromText(EX, 'example-script.txt'), { kind: 'Example' });
showSlot(); render();
setNotice('<span class="tag">Example</span> A sample script is loaded. Dotted markers are places it chose not to break, amber ones are worth a check. Drop or paste your own to replace it.');

SS.consumeHandoff(files => loadFile(files[0]));
})();
