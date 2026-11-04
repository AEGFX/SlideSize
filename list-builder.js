/* list-builder.js, the List Builder engine
 *
 * Turns a spreadsheet of names into a PowerPoint deck of lists, N names a
 * slide, with every name box a real placeholder that inherits from a LIST
 * layout on one master. A JavaScript port of the NAMEDROP Flow engine
 * (engine_flow.py, rules.py, source.py) plus a pptx writer that builds the
 * package by hand with JSZip, because no library writes true inherited
 * placeholders.
 *
 * Loads in a browser as window.ListBuilder and in Node through
 * module.exports, so the tests and the headless build script share this
 * exact code with the page. Nothing here touches the DOM. JSZip and SheetJS
 * are passed in by the caller.
 *
 * Sections
 *   1  rules      sorting, particles, Mc and Mac, accents, natural numeric, title case
 *   2  source     cell cleaning, header detection, CSV parsing, SheetJS grid
 *   3  mapping    column guesses and record building
 *   4  validate   warnings and the three automatic fixes
 *   5  engine     chunking and the slide plan
 *   6  geometry   where the placeholders go on the layout
 *   7  pptx       the package writer
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ListBuilder = factory();
}(typeof self !== 'undefined' ? self : this, function () {
'use strict';

var LB = { version: '1.0.0', namedropVersion: '0.1.0' };

/* ======================================================================
   1  rules  (rules.py)
   ====================================================================== */

var PARTICLES = [
  'van der', 'van den', 'van de', 'von der', 'von dem', 'de la', 'de los', 'de las',
  'del la', 'da la', "in 't", 'in het', 'op de', 'op den', "van 't", 'ter', 'ten',
  'van', 'von', 'de', 'del', 'della', 'di', 'da', 'du', 'des', 'le', 'la', 'den', 'der',
  'te', 'af', 'av', 'zu', 'zum', 'zur', 'el', 'al', 'bin', 'ibn', "d'", "l'", "o'",
  'st', 'st.', 'san', 'santa'
];
var PARTICLE_SET = {};
PARTICLES.forEach(function (p) { PARTICLE_SET[p] = true; });

var HONORIFICS = toSet('dr dr. prof prof. professor sir dame lord lady mr mr. mrs mrs. ms ms. miss mx rev rev. hon hon. capt capt. col col. major lt lt. cllr baroness baron');
var SUFFIXES = toSet('jr jr. sr sr. ii iii iv mbe obe cbe kbe dbe phd ph.d ph.d. md esq esq. frs qc kc cfa cpa frsa fca ma ba msc bsc mba dphil dds rn');
var ROMAN = toSet('ii iii iv vi vii viii');
var CAP_SUFFIX = toSet('mbe obe cbe kbe dbe phd qc kc md');

function toSet(words) {
  var s = {};
  words.split(/\s+/).forEach(function (w) { if (w) s[w] = true; });
  return s;
}

var ACCENT_SPECIAL = { '\u00f8': 'o', '\u00d8': 'O', '\u00df': 'ss', '\u00e6': 'ae', '\u00c6': 'AE', '\u0153': 'oe', '\u0152': 'OE',
  '\u0142': 'l', '\u0141': 'L', '\u0111': 'd', '\u0110': 'D', '\u00fe': 'th', '\u00de': 'Th', '\u00f0': 'd', '\u00d0': 'D' };

/* M\u00fcller to Muller, \u00d8degaard to Odegaard. */
function foldAccents(s) {
  var out = '';
  for (var i = 0; i < s.length; i++) {
    var ch = s[i];
    out += ACCENT_SPECIAL[ch] !== undefined ? ACCENT_SPECIAL[ch] : ch;
  }
  return out.normalize('NFKD').replace(/\p{M}/gu, '');
}

function isLowerStr(s) { return s !== s.toUpperCase() && s === s.toLowerCase(); }
function isUpperStr(s) { return s !== s.toLowerCase() && s === s.toUpperCase(); }

/* 'van der Berg' -> ['van der', 'Berg']. A particle is only recognised when
   lower case or when the whole surname shares one case, so 'De Niro' stays
   one name. */
function splitParticles(surname) {
  var s = String(surname || '').trim();
  if (!s) return ['', ''];
  var words = s.split(' ');
  if (words.length < 2) return ['', s];
  var uniform = isLowerStr(s) || isUpperStr(s);
  for (var n = 3; n >= 1; n--) {
    if (words.length > n) {
      var head = words.slice(0, n).join(' ');
      if (PARTICLE_SET[head.toLowerCase()] && (uniform || head === head.toLowerCase())) {
        return [head, words.slice(n).join(' ')];
      }
    }
  }
  return ['', s];
}

function hasParticle(surname) { return splitParticles(surname)[0] !== ''; }

/* Mc and Mac read as the same prefix, McDonald near MacDonald. */
function normaliseMc(core) {
  return core.replace(/^(mc|mac)([a-z])/i, function (m, a, b) { return 'mac' + b; });
}

function sortSettings(o) {
  o = o || {};
  return {
    particles: o.particles === 'dutch' ? 'dutch' : 'uk',
    mcMacTogether: !!o.mcMacTogether,
    caseInsensitive: o.caseInsensitive === undefined ? true : !!o.caseInsensitive,
    foldAccents: o.foldAccents === undefined ? true : !!o.foldAccents,
    naturalNumeric: o.naturalNumeric === undefined ? true : !!o.naturalNumeric
  };
}

/* The surname as it should sort. UK keeps the particle in front, Dutch moves it after. */
function surnameSortForm(surname, settings) {
  settings = sortSettings(settings);
  var pc = splitParticles(surname), particle = pc[0], core = pc[1];
  var form;
  if (settings.particles === 'dutch' && particle) form = core + ' ' + particle;
  else form = String(surname).trim();
  if (settings.mcMacTogether) {
    if (particle && settings.particles === 'dutch') form = normaliseMc(core) + ' ' + particle;
    else if (!particle) form = normaliseMc(form);
  }
  return form;
}

/* 'Table 2' before 'Table 10'. Numbers compare as numbers, text as text. */
function naturalKey(s) {
  var parts = String(s).split(/(\d+)/);
  var key = [];
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i];
    if (p === '') continue;
    if (/^\d+$/.test(p)) key.push([0, Number(p), '']);
    else key.push([1, 0, p]);
  }
  return key;
}

function textKey(s, settings) {
  settings = sortSettings(settings);
  if (settings.foldAccents) s = foldAccents(s);
  if (settings.caseInsensitive) s = s.toLowerCase();
  s = s.trim();
  if (settings.naturalNumeric) return naturalKey(s);
  return [[1, 0, s]];
}

/* Key for any cell value. Blanks sort last. Numbers before text. */
function valueKey(value, settings) {
  if (value === null || value === undefined || value === '') return [2, []];
  if (typeof value === 'boolean') return [1, [[0, value ? 1 : 0, '']]];
  if (typeof value === 'number') return [0, [[0, value, '']]];
  return [1, textKey(String(value), settings)];
}

/* Python tuple comparison on nested arrays of numbers and strings. */
function compareKeys(a, b) {
  if (Array.isArray(a) && Array.isArray(b)) {
    var n = Math.min(a.length, b.length);
    for (var i = 0; i < n; i++) {
      var c = compareKeys(a[i], b[i]);
      if (c) return c;
    }
    return a.length - b.length;
  }
  if (typeof a === 'number' && typeof b === 'number') return a < b ? -1 : a > b ? 1 : 0;
  a = String(a); b = String(b);
  return a < b ? -1 : a > b ? 1 : 0;
}

/* A sort rule is a list of keys. A string key '-last_name' is descending. */
function sortRule(keys, settings) {
  var out = [];
  (keys || []).forEach(function (item) {
    if (typeof item === 'string') {
      out.push({ field: item.replace(/^[-+]/, ''), descending: item.charAt(0) === '-', kind: 'text' });
    } else {
      out.push({ field: item.field, descending: !!item.descending, kind: item.kind || 'text' });
    }
  });
  return { keys: out, settings: sortSettings(settings) };
}

function recordKey(record, rule) {
  return rule.keys.map(function (key) {
    var v = record[key.field];
    var k;
    if (key.kind === 'surname' || key.field === 'last_name' || key.field === 'surname') {
      k = valueKey(v ? surnameSortForm(String(v), rule.settings) : null, rule.settings);
    } else {
      k = valueKey(v, rule.settings);
    }
    return { k: k, desc: key.descending };
  });
}

/* Stable sort of records by the rule. Source order is the final tie-break. */
function sortRecords(records, rule) {
  var recs = records.slice();
  if (!rule.keys.length) return recs;
  var keyed = recs.map(function (r, i) { return { r: r, i: i, k: recordKey(r, rule) }; });
  keyed.sort(function (x, y) {
    for (var j = 0; j < x.k.length; j++) {
      var c = compareKeys(x.k[j].k, y.k[j].k);
      if (c) return x.k[j].desc ? -c : c;
    }
    return x.i - y.i;
  });
  return keyed.map(function (e) { return e.r; });
}

function capitalize(w) { return w ? w.charAt(0).toUpperCase() + w.slice(1).toLowerCase() : w; }

/* JOHN SMITH -> John Smith. ANNA VAN DER BERG -> Anna van der Berg. MCDONALD -> McDonald. */
function smartTitleCase(name) {
  function capWord(w) {
    if (!w) return w;
    var lw = w.toLowerCase();
    if (ROMAN[lw]) return w.toUpperCase();
    if (SUFFIXES[lw] && CAP_SUFFIX[lw]) return w.length <= 3 ? w.toUpperCase() : 'PhD';
    var m = /^(mc|mac)([a-z])(.*)$/.exec(lw);
    if (m && lw.length > 4 && m[1] === 'mc') return 'Mc' + m[2].toUpperCase() + m[3];
    if (m && m[1] === 'mac' && lw.length > 5 && ['mack', 'mace', 'macon', 'macey'].indexOf(lw) < 0) return 'Mac' + m[2].toUpperCase() + m[3];
    var ap = lw.indexOf("'");
    if (ap === 1 || ap === 2) return capitalize(lw.slice(0, ap)) + "'" + capitalize(lw.slice(ap + 1));
    return lw.charAt(0).toUpperCase() + lw.slice(1);
  }
  var words = String(name).trim().split(' ');
  return words.map(function (w, i) {
    var lw = w.toLowerCase();
    if (i > 0 && PARTICLE_SET[lw]) return lw;
    return w.split('-').map(capWord).join('-');
  }).join(' ');
}

function letters(s) { return String(s).match(/\p{L}/gu) || []; }
function isAllCaps(s) {
  var l = letters(s);
  return l.length >= 2 && l.every(function (c) { return c === c.toUpperCase() && c !== c.toLowerCase(); });
}
function isAllLower(s) {
  var l = letters(s);
  return l.length >= 2 && l.every(function (c) { return c === c.toLowerCase() && c !== c.toUpperCase(); });
}

/* Remove honorifics from the front and suffixes from the end. Returns [name, removed]. */
function stripHonorifics(name) {
  var words = String(name).trim().split(/\s+/).filter(Boolean);
  var removed = [];
  function isHon(w) {
    var lw = w.toLowerCase();
    return HONORIFICS[lw.replace(/\.+$/, '') + (w.endsWith('.') ? '.' : '')] || HONORIFICS[lw];
  }
  while (words.length && isHon(words[0])) removed.push(words.shift());
  while (words.length && (SUFFIXES[words[words.length - 1].toLowerCase()] || SUFFIXES[words[words.length - 1].toLowerCase().replace(/,+$/, '')])) removed.push(words.pop());
  if (words.length && words[words.length - 1].endsWith(',')) words[words.length - 1] = words[words.length - 1].replace(/,+$/, '');
  return [words.join(' '), removed];
}

function isHyphenated(surname) {
  var s = String(surname).trim();
  return s.replace(/^-+|-+$/g, '').indexOf('-') >= 0 || (s.indexOf(' ') >= 0 && !hasParticle(s));
}

/* One cell holding a whole name. 'Anna van der Berg' -> ['Anna', 'van der Berg'].
   The surname is the last word plus any lower case particle words in front
   of it, so 'Mary Jane Smith' gives ['Mary Jane', 'Smith']. Honorifics and
   suffixes are dropped first. */
function splitFullName(full) {
  var cleaned = stripHonorifics(full)[0];
  var words = cleaned.split(' ').filter(Boolean);
  if (words.length < 2) return [cleaned, ''];
  var i = words.length - 1;
  while (i > 1 && PARTICLE_SET[words[i - 1].toLowerCase()] && words[i - 1] === words[i - 1].toLowerCase()) i--;
  return [words.slice(0, i).join(' '), words.slice(i).join(' ')];
}

/* ======================================================================
   2  source  (source.py)
   ====================================================================== */

var ZERO_WIDTH = ['\u200b', '\u200c', '\u200d', '\u2060', '\ufeff'];
var NBSP = '\u00a0';
var LINE_BREAKS = ['\r', '\n', '\u000b', '\u000c', '\u2028', '\u2029'];

function pad2(n) { return (n < 10 ? '0' : '') + n; }

function toText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return String(v);
  if (v instanceof Date) {
    var d = v.getFullYear() + '-' + pad2(v.getMonth() + 1) + '-' + pad2(v.getDate());
    if (v.getHours() === 0 && v.getMinutes() === 0 && v.getSeconds() === 0) return d;
    return d + ' ' + pad2(v.getHours()) + ':' + pad2(v.getMinutes());
  }
  return String(v);
}

/* Clean one string. Returns [cleaned, notes]. */
function cleanText(raw) {
  var notes = [];
  var s = raw;
  var zw = 0, bom = 0;
  ZERO_WIDTH.forEach(function (ch) {
    var c = s.split(ch).length - 1;
    if (ch === '\ufeff') bom += c; else zw += c;
  });
  if (bom) notes.push('byte order mark');
  if (zw) notes.push('zero width character');
  if (bom || zw) ZERO_WIDTH.forEach(function (ch) { s = s.split(ch).join(''); });
  if (s.indexOf(NBSP) >= 0) { notes.push('non-breaking space'); s = s.split(NBSP).join(' '); }
  if (LINE_BREAKS.some(function (ch) { return s.indexOf(ch) >= 0; })) {
    notes.push('line break');
    LINE_BREAKS.forEach(function (ch) { s = s.split(ch).join(' '); });
  }
  if (s.indexOf('\t') >= 0) { notes.push('tab'); s = s.split('\t').join(' '); }
  var stripped = s.replace(/^[ \t\n\r\f\v]+|[ \t\n\r\f\v]+$/g, '');
  if (stripped !== s) {
    if (s.charAt(0) === ' ') notes.push('leading space');
    if (s.charAt(s.length - 1) === ' ') notes.push('trailing space');
    s = stripped;
  }
  var collapsed = s.replace(/ {2,}/g, ' ');
  if (collapsed !== s) { if (!notes.length) notes.push('double spaces'); s = collapsed; }
  return [s, notes];
}

function cleanCell(raw, formulaNoCache) {
  if (typeof raw === 'string') {
    var r = cleanText(raw);
    return { value: r[0] === '' ? null : r[0], raw: raw, stripped: r[1], formulaNoCache: !!formulaNoCache };
  }
  return { value: raw === undefined ? null : raw, raw: raw, stripped: [], formulaNoCache: !!formulaNoCache };
}

function isBlankCell(c) { return c.value === null || c.value === undefined || c.value === ''; }

var NUMERIC_TEXT = /^-?\d+([.,]\d+)?$/;
var DATE_TEXT = /^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}$/;
var TIME_TEXT_SIG = /^\d{1,2}[:.]\d{2}$/;

function isEmptyVal(v) { return v === null || v === undefined || (typeof v === 'string' && v.trim() === ''); }
function isText(v) { return typeof v === 'string' && v.trim() !== ''; }

function signature(values) {
  return values.map(function (v) {
    if (isEmptyVal(v)) return '-';
    if (typeof v === 'string') {
      var t = v.trim();
      if (NUMERIC_TEXT.test(t)) return 'n';
      if (DATE_TEXT.test(t) || TIME_TEXT_SIG.test(t)) return 'd';
      return 't';
    }
    if (v instanceof Date) return 'd';
    if (typeof v === 'number') return 'n';
    return '?';
  }).join('');
}

/* True when any non-empty text in row i repeats in the same column below it. */
function recursBelow(grid, i, look) {
  look = look || 50;
  var row = grid[i];
  for (var c = 0; c < row.length; c++) {
    var v = row[c];
    if (isEmptyVal(v)) continue;
    var t = String(v).trim().toLowerCase();
    for (var r = i + 1; r < Math.min(grid.length, i + 1 + look); r++) {
      var rr = grid[r];
      if (c < rr.length && !isEmptyVal(rr[c]) && String(rr[c]).trim().toLowerCase() === t) return true;
    }
  }
  return false;
}

/* Find the header row in a grid of raw values. Returns {index, candidates}, 0-based. */
function detectHeaderRow(grid) {
  if (!grid || !grid.length) return { index: 0, candidates: [] };
  var width = 0;
  grid.forEach(function (r) { if (r.length > width) width = r.length; });
  var counts = grid.map(function (r) { return r.filter(function (v) { return !isEmptyVal(v); }).length; });
  var widest = Math.max.apply(null, counts.concat([0]));
  var candidates = [];
  for (var i = 0; i < grid.length; i++) {
    var n = counts[i];
    if (n === 0 || widest === 0) continue;
    var texts = grid[i].filter(function (v) { return !isEmptyVal(v); });
    if (!texts.every(function (v) { return isText(v) && !NUMERIC_TEXT.test(String(v).trim()); })) continue;
    if (n < Math.max(2, Math.round(widest * 0.6))) continue;
    var vals = texts.map(function (v) { return String(v).trim(); });
    var distinct = {};
    vals.forEach(function (v) { distinct[v] = true; });
    if (Object.keys(distinct).length !== vals.length) continue;
    if (recursBelow(grid, i)) continue;
    candidates.push(i);
    if (candidates.length >= 10) break;
  }
  if (!candidates.length) return { index: 0, candidates: [] };
  for (var a = 0; a < candidates.length; a++) {
    var ci = candidates[a];
    var j = ci + 1;
    while (j < grid.length && counts[j] === 0) j++;
    if (j >= grid.length) continue;
    if (signature(grid[j].slice(0, width)) !== signature(grid[ci].slice(0, width))) return { index: ci, candidates: candidates };
  }
  return { index: candidates[0], candidates: candidates };
}

function uniqueHeaders(rawHeaders) {
  var seen = {}, out = [];
  rawHeaders.forEach(function (h, i) {
    var name = typeof h === 'string' ? cleanText(h)[0] : toText(h);
    if (!name) name = 'Column ' + (i + 1);
    var base = name;
    if (seen[name]) { seen[name] += 1; name = base + ' (' + seen[base] + ')'; }
    else seen[name] = 1;
    out.push(name);
  });
  return out;
}

/* Build a table from a grid. opts: hidden (bool per row), formulaFlags, headerRow (0-based or null),
   includeHidden, firstRowNumber (1). */
function buildTable(grid, opts) {
  opts = opts || {};
  var first = opts.firstRowNumber || 1;
  var det = detectHeaderRow(grid);
  var idx = (opts.headerRow === null || opts.headerRow === undefined) ? det.index : opts.headerRow;
  if (idx < 0 || idx >= grid.length) throw new Error('Header row ' + (idx + first) + ' is outside the sheet');
  var headers = uniqueHeaders(grid[idx]);
  var hidden = opts.hidden || [];
  var flags = opts.formulaFlags || null;
  var rows = [], skipped = 0;
  for (var r = idx + 1; r < grid.length; r++) {
    var rawRow = grid[r];
    if (hidden[r] && !opts.includeHidden) { skipped++; continue; }
    var cells = {}, any = false;
    for (var c = 0; c < headers.length; c++) {
      var raw = c < rawRow.length ? rawRow[c] : null;
      var flag = !!(flags && flags[r] && flags[r][c]);
      var cell = cleanCell(raw, flag);
      if (!isBlankCell(cell) || cell.formulaNoCache) any = true;
      cells[headers[c]] = cell;
    }
    if (!any) continue;
    rows.push({ number: first + r, cells: cells, hidden: !!hidden[r] });
  }
  return {
    headers: headers, rows: rows,
    headerRow: idx, headerCandidates: det.candidates,
    hiddenRowsSkipped: skipped, rowCount: grid.length
  };
}

/* ---------- CSV ---------- */

var CP1252_HIGH = [0x20AC, 0x81, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x8D, 0x017D, 0x8F,
  0x90, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x9D, 0x017E, 0x0178];

function decodeCp1252(bytes) {
  var out = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i];
    out += String.fromCharCode(b >= 0x80 && b < 0xA0 ? CP1252_HIGH[b - 0x80] : b);
  }
  return out;
}

/* UTF-8 with BOM, then UTF-8, then Windows-1252 as the fallback. Returns {text, encoding}. */
function decodeBytes(bytes) {
  if (!(bytes instanceof Uint8Array)) bytes = new Uint8Array(bytes);
  var TD = typeof TextDecoder !== 'undefined' ? TextDecoder : null;
  if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
    return { text: new TD('utf-8').decode(bytes.subarray(3)), encoding: 'utf-8-sig' };
  }
  try {
    return { text: new TD('utf-8', { fatal: true }).decode(bytes), encoding: 'utf-8' };
  } catch (e) {
    var text;
    try { text = new TD('windows-1252').decode(bytes); } catch (e2) { text = decodeCp1252(bytes); }
    return { text: text, encoding: 'cp1252' };
  }
}

/* Pick comma, semicolon or tab from the first 4 KB. */
function sniffDelimiter(text) {
  var sample = text.slice(0, 4096);
  var lines = sample.split(/\r\n|\r|\n/).filter(function (l) { return l.trim() !== ''; }).slice(0, 20);
  var best = ',', bestScore = -1;
  [',', ';', '\t'].forEach(function (d) {
    var counts = lines.map(function (l) {
      var n = 0, q = false;
      for (var i = 0; i < l.length; i++) {
        var ch = l[i];
        if (ch === '"') q = !q;
        else if (ch === d && !q) n++;
      }
      return n;
    });
    if (!counts.length) return;
    var min = Math.min.apply(null, counts), max = Math.max.apply(null, counts);
    if (max === 0) return;
    var score = min === max ? max * 10 : min;
    if (score > bestScore) { bestScore = score; best = d; }
  });
  return best;
}

function parseCsvText(text, delim) {
  delim = delim || sniffDelimiter(text);
  var rows = [], row = [], field = '', i = 0, q = false, n = text.length;
  while (i < n) {
    var ch = text[i];
    if (q) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
        q = false; i++; continue;
      }
      field += ch; i++; continue;
    }
    if (ch === '"') { q = true; i++; continue; }
    if (ch === delim) { row.push(field); field = ''; i++; continue; }
    if (ch === '\r' || ch === '\n') {
      row.push(field); field = '';
      rows.push(row); row = [];
      if (ch === '\r' && text[i + 1] === '\n') i++;
      i++; continue;
    }
    field += ch; i++;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows.map(function (r) { return r.map(function (v) { return v === '' ? null : v; }); });
}

/* Bytes of a CSV file to a table. */
function readCsv(bytes, opts) {
  opts = opts || {};
  var dec = decodeBytes(bytes);
  var grid = parseCsvText(dec.text, opts.delimiter);
  var table = buildTable(grid, { headerRow: opts.headerRow, includeHidden: true });
  table.encoding = dec.encoding;
  table.grid = grid;
  return table;
}

/* ---------- SheetJS ---------- */

/* Raw grid, hidden flags and formula flags from a SheetJS worksheet. */
function gridFromSheet(ws, XLSX) {
  var out = { grid: [], hidden: [], formulaFlags: [] };
  if (!ws || !ws['!ref']) return out;
  var range = XLSX.utils.decode_range(ws['!ref']);
  var rowInfo = ws['!rows'] || [];
  for (var r = range.s.r; r <= range.e.r; r++) {
    var vals = [], flags = [];
    for (var c = range.s.c; c <= range.e.c; c++) {
      var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
      var v = null, f = false;
      if (cell) {
        if (cell.t === 'e') v = null;
        else if (cell.v === undefined || cell.v === null) v = null;
        else if (cell.t === 'd' && cell.v instanceof Date) v = cell.v;
        else if (cell.t === 'n') v = cell.v;
        else if (cell.t === 'b') v = !!cell.v;
        else v = String(cell.v);
        if (cell.f && v === null) f = true;
      }
      vals.push(v); flags.push(f);
    }
    out.grid.push(vals);
    out.formulaFlags.push(flags);
    out.hidden.push(!!(rowInfo[r] && rowInfo[r].hidden));
  }
  return out;
}

function readSheet(ws, XLSX, opts) {
  opts = opts || {};
  var g = gridFromSheet(ws, XLSX);
  var table = buildTable(g.grid, { hidden: g.hidden, formulaFlags: g.formulaFlags, headerRow: opts.headerRow, includeHidden: !!opts.includeHidden });
  table.grid = g.grid;
  table.hiddenFlags = g.hidden;
  table.formulaFlags = g.formulaFlags;
  return table;
}

/* ======================================================================
   3  mapping  (mapping.py)
   ====================================================================== */

var SYNONYMS = {
  first_name: ['first name', 'firstname', 'first', 'forename', 'given name', 'given', 'name 1', 'voornaam'],
  last_name: ['last name', 'lastname', 'surname', 'last', 'family name', 'family', 'second name', 'achternaam'],
  full_name: ['full name', 'name', 'guest name', 'guest', 'speaker name', 'speaker', 'attendee', 'delegate', 'delegate name', 'attendee name'],
  table: ['table', 'table number', 'table no', 'table #', 'tbl', 'table name', 'seat table', 'tafel']
};

function normHeader(h) { return String(h).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }

/* Propose columns for first, last, full and table from the headers, then by content. */
function proposeMapping(table) {
  var headers = {};
  table.headers.forEach(function (h) { if (headers[normHeader(h)] === undefined) headers[normHeader(h)] = h; });
  var used = {};
  function byHeader(f) {
    var cands = [f.replace(/_/g, ' ')].concat(SYNONYMS[f] || []);
    for (var i = 0; i < cands.length; i++) {
      var h = headers[normHeader(cands[i])];
      if (h !== undefined && !used[h]) return h;
    }
    return null;
  }
  /* second pass, a header that contains a synonym as whole words, Guest First Name, Delegate Surname */
  function byContains(f) {
    var cands = [f.replace(/_/g, ' ')].concat(SYNONYMS[f] || []).map(normHeader).filter(function (c) { return c.length > 2; });
    var best = null;
    table.headers.forEach(function (h) {
      if (used[h]) return;
      var nh = ' ' + normHeader(h) + ' ';
      if (cands.some(function (c) { return nh.indexOf(' ' + c + ' ') >= 0; })) {
        if (f === 'full_name' && /\b(first|last|sur|family|given|fore)/.test(nh)) return;
        if (best === null || h.length < best.length) best = h;
      }
    });
    return best;
  }
  var p = { first: null, last: null, full: null, table: null };
  p.first = byHeader('first_name') || byContains('first_name'); if (p.first) used[p.first] = true;
  p.last = byHeader('last_name') || byContains('last_name'); if (p.last) used[p.last] = true;
  if (!p.first || !p.last) {
    var full = byHeader('full_name') || byContains('full_name');
    if (full) { p.full = full; used[full] = true; p.first = null; p.last = null; }
  }
  p.table = byHeader('table') || byContains('table'); if (p.table) used[p.table] = true;
  if (!p.table) {
    for (var i = 0; i < table.headers.length; i++) {
      var h = table.headers[i];
      if (used[h]) continue;
      var cells = table.rows.map(function (r) { return r.cells[h]; }).filter(function (c) { return c && !isBlankCell(c); });
      if (!cells.length) continue;
      var ints = cells.filter(function (c) { return typeof c.value === 'number' && Number.isInteger(c.value) || (typeof c.value === 'string' && /^\d+$/.test(c.value)); });
      var maxv = 0;
      ints.forEach(function (c) { var n = Number(c.value); if (n > maxv) maxv = n; });
      if (ints.length / cells.length > 0.8 && maxv < 500) { p.table = h; used[h] = true; break; }
    }
  }
  return p;
}

/* Numeric text becomes a number so CSV and Excel sort the same way. */
function numberish(v) {
  if (typeof v === 'string' && /^\d+$/.test(v.trim())) return Number(v.trim());
  if (typeof v === 'number' && Number.isInteger(v)) return v;
  return v;
}

/* Records from table rows and a mapping {first, last, full, table, extra}. */
function buildRecords(table, mapping) {
  var recs = [];
  table.rows.forEach(function (row) {
    function cell(h) { return h ? (row.cells[h] || { value: null, stripped: [] }) : null; }
    var rec = { row: row.number, first_name: '', last_name: '', table: null, extra: '', cells: {}, columns: {} };
    if (mapping.full) {
      var fc = cell(mapping.full);
      var full = toText(fc.value);
      var split = splitFullName(full);
      rec.first_name = split[0]; rec.last_name = split[1];
      rec.full_source = full;
      rec.cells.first_name = fc; rec.cells.last_name = fc;
      rec.columns.first_name = mapping.full; rec.columns.last_name = mapping.full;
    } else {
      var a = cell(mapping.first), b = cell(mapping.last);
      rec.first_name = a ? toText(a.value) : '';
      rec.last_name = b ? toText(b.value) : '';
      rec.cells.first_name = a; rec.cells.last_name = b;
      rec.columns.first_name = mapping.first; rec.columns.last_name = mapping.last;
    }
    if (mapping.table) {
      var tc = cell(mapping.table);
      var tv = tc.value instanceof Date ? toText(tc.value) : tc.value;
      rec.table = numberish(tv);
      rec.cells.table = tc; rec.columns.table = mapping.table;
    }
    if (mapping.extra) {
      var ec = cell(mapping.extra);
      rec.extra = toText(ec.value instanceof Date ? toText(ec.value) : ec.value);
      rec.cells.extra = ec; rec.columns.extra = mapping.extra;
    }
    recs.push(rec);
  });
  recs.forEach(composeName);
  return recs;
}

function composeName(rec) {
  rec.name = [rec.first_name, rec.last_name].filter(Boolean).join(' ');
}

function personLabel(rec) {
  return rec.name || ('row ' + rec.row);
}

/* ======================================================================
   4  validate  (validate.py, the rules that matter for a list)
   ====================================================================== */

var COMMON_FIRST_NAMES = toSet('aaron abigail adam adrian aidan aimee alan alastair albert alec alex alexander alexandra alfie alice alison alistair amanda amber amelia amy andrea andrew andy angela angus anita ann anna anne annie anthony antonia archie arthur ashley barbara barry beatrice ben benjamin bernard beth bethany bill billy bob brian bridget bruce caitlin callum cameron carl carol caroline carrie catherine cathy charles charlie charlotte chloe chris christian christine christopher claire clare clive colin connor craig dan daniel danielle danny darren dave david dawn dean debbie deborah declan denise dennis derek diane dominic donald donna dorothy douglas duncan dylan ed eddie edward eleanor elizabeth ella ellen ellie elliot emily emma eric erin ethan eva eve evelyn ewan felix finlay fiona florence frances francis frank fred freddie freya gail gareth gary gavin gemma george georgia georgina gerald gillian gordon grace graham grant greg gregory guy hannah harriet harry harvey hayley hazel heather helen henry holly howard hugh hugo ian imogen irene isaac isabel isabella isla ivan jack jackie jacob jacqueline jade jake james jamie jan jane janet janice jason jasmine jay jean jeff jeffrey jemma jenna jennifer jenny jeremy jess jessica jill jim joan joanna joanne jodie joe joel john jon jonathan jordan joseph josephine josh joshua joyce judith julia julian julie justin karen kate katherine kathleen kathryn katie kay keith kelly ken kenneth kerry kevin kieran kim kimberley kirsty kyle laura lauren lawrence leah lee leo leon leslie lewis liam lily linda lisa liz lloyd logan lorraine louis louise lucas lucy luke lydia lynn madeleine maggie malcolm marc marcus margaret maria marie marion mark martin mary matt matthew maureen max megan melanie melissa michael michelle mike molly morgan nancy naomi natalie natasha nathan neil niall nicholas nick nicola nicole nigel noah norman oliver olivia oscar owen paige pam pamela pat patricia patrick paul paula pauline penny pete peter phil philip philippa phoebe poppy rachel ralph ray raymond rebecca reece rhys richard rob robert robin roger ronald rory rosie ross roy ruby russell ruth ryan sally sam samantha samuel sandra sara sarah scott sean shane sharon sheila shirley sian simon sophia sophie stacey stephanie stephen steve steven stewart stuart sue susan suzanne sylvia tara ted teresa terry theo thomas tim timothy tina toby tom tony tracy trevor valerie vanessa vicky victoria vincent wayne wendy will william yvonne zara zoe ahmed ali amir anil arjun ayesha deepak fatima hassan imran kavita mohammed muhammad neha omar priya rahul raj ravi sanjay sunil vikram yusuf anders anja anke annika astrid bart bas birgit bram carsten daan dirk elke erik femke floris frans freek frits gerrit hans heike henk hendrik ingrid jaap jan jeroen johan joost jorg jurgen kai karin kees klaas lars lotte maarten marieke marije marjolein mark matthijs merel michiel niels nils pieter rik rob roel ruud sander sanne sjoerd sven thijs tim wouter wim willem alessandro alessandra andrea antonio carlos carmen chiara diego elena enrique fabio federico francesca francesco giovanni giulia javier jorge jose juan laura lorenzo luca luis manuel marco maria mario marta matteo miguel pablo paolo pedro rafael roberto rosa sergio silvia simone sofia stefano valentina antoine aurelie benoit camille cecile celine claude clement delphine elodie emmanuel fabien francois guillaume helene isabelle jacques jean jerome julien laurent lucie marion mathieu michel nicolas olivier pascal philippe pierre sebastien sophie stephane sylvie thierry thomas valerie vincent yann');

var COMMON_SURNAMES = toSet('smith jones williams taylor brown davies evans wilson thomas roberts johnson lewis walker robinson wood thompson white watson jackson wright green harris cooper king lee martin clarke james morgan hughes edwards hill moore clark harrison scott young morris hall ward turner carter phillips mitchell patel adams campbell anderson allen cook bailey parker miller davis murphy price bell baker griffiths kelly simpson marshall collins bennett cox richardson fox gray rose chapman hunt robertson shaw reynolds lloyd ellis harvey stevens fisher knight hudson barnes holmes palmer mills webb day rogers dixon mason pearson kaur singh khan ahmed ali hussain begum sharma mehta fischer weber schulz becker hoffmann rossi russo ferrari esposito garcia martinez lopez fernandez dubois lefebvre moreau bernard jansen bakker visser smit meijer mulder bos vos peters dekker nguyen tran kim park tanaka sato okafor mensah osei owen george henry grant graham gordon howard douglas murray stewart hamilton ross russell wallace dean foster gibson hunter lawrence');

function q(v) { return "'" + toText(v) + "'"; }
function isCommonFirst(s) { return !!COMMON_FIRST_NAMES[foldAccents(String(s).trim().toLowerCase())]; }
var NON_LATIN = /[^\u0000-\u00ff]/;

/* Validate and fix records in place. Returns {records, skipped, warnings}.
   warnings are {level: 'error'|'warning'|'info', row, text, rule}. Rows with no
   name at all are dropped, nothing else blocks. The fixes applied here match
   the NAMEDROP reference projects, a full name in one cell is split, ALL CAPS
   or all lower names in a title case list are title cased and honorifics are
   stripped. The sort and the deck see the fixed values. */
function validate(records, opts) {
  opts = opts || {};
  var settings = sortSettings(opts.sortSettings);
  var warnings = [];
  var fixNames = opts.fixNames === undefined ? true : !!opts.fixNames;
  var hasFull = !!opts.fullNameColumn;
  function warn(level, rule, rec, text) { warnings.push({ level: level, rule: rule, row: rec ? rec.row : null, text: text }); }
  function col(rec, f) { return rec.columns[f] || f; }

  records.forEach(function (rec) {
    /* stripped whitespace, info */
    ['first_name', 'last_name', 'table', 'extra'].forEach(function (f) {
      var c = rec.cells[f];
      if (c && c.stripped && c.stripped.length && (f !== 'last_name' || !hasFull)) {
        warn('info', 'stripped', rec, 'Row ' + rec.row + ', ' + col(rec, f) + ' had ' + c.stripped.join(', ') + ' in ' + q(visible(c.raw)) + '. Cleaned to ' + q(toText(c.value)) + '.');
      }
      if (c && c.formulaNoCache) {
        warn('warning', 'formula_no_cache', rec, 'Row ' + rec.row + ', ' + col(rec, f) + ' holds a formula with no saved result, so it reads as blank. Open the sheet in Excel and save it.');
      }
    });
  });

  /* full name in one cell */
  if (!hasFull) {
    records.forEach(function (rec) {
      var first = rec.first_name, last = rec.last_name;
      if (first && !last && first.trim().indexOf(' ') >= 0) {
        var parts = first.trim().split(' ');
        var cleaned = stripHonorifics(first)[0];
        parts = cleaned ? cleaned.split(' ') : parts;
        if (parts.length >= 2) {
          var nf = parts[0], nl = parts.slice(1).join(' ');
          if (fixNames) { rec.first_name = nf; rec.last_name = nl; }
          warn('warning', 'full_name_in_one_field', rec, 'Row ' + rec.row + ', ' + col(rec, 'last_name') + ' is blank but ' + col(rec, 'first_name') + ' holds ' + q(first) + '. Looks like a full name in one cell. ' + (fixNames ? 'Split into ' + q(nf) + ' and ' + q(nl) + '.' : 'Left as it is.'));
        }
      } else if (last && !first && last.trim().indexOf(' ') >= 0) {
        var parts2 = last.trim().split(' ');
        if (parts2.length >= 2 && !hasParticle(last)) {
          var nf2 = parts2[0], nl2 = parts2.slice(1).join(' ');
          if (fixNames) { rec.first_name = nf2; rec.last_name = nl2; }
          warn('warning', 'full_name_in_one_field', rec, 'Row ' + rec.row + ', ' + col(rec, 'first_name') + ' is blank but ' + col(rec, 'last_name') + ' holds ' + q(last) + '. Looks like a full name in one cell. ' + (fixNames ? 'Split into ' + q(nf2) + ' and ' + q(nl2) + '.' : 'Left as it is.'));
        }
      }
    });
  }

  /* swapped names, warning only */
  records.forEach(function (rec) {
    var first = rec.first_name, last = rec.last_name;
    if (!first || !last || first.indexOf(' ') >= 0 || last.indexOf(' ') >= 0) return;
    if (isCommonFirst(last) && !isCommonFirst(first) && COMMON_SURNAMES[foldAccents(first.toLowerCase())]) {
      warn('warning', 'swapped_names', rec, 'Row ' + rec.row + ', ' + col(rec, 'first_name') + ' ' + q(first) + ' and ' + col(rec, 'last_name') + ' ' + q(last) + ' look swapped. ' + last + ' is a common first name, ' + first + ' is a surname. Check the sheet.');
    }
  });

  /* honorifics, fixed */
  records.forEach(function (rec) {
    ['first_name', 'last_name'].forEach(function (f) {
      var val = rec[f];
      if (!val) return;
      var r = stripHonorifics(val);
      if (r[1].length && r[0]) {
        if (fixNames) rec[f] = r[0];
        warn('warning', 'honorific_in_name', rec, 'Row ' + rec.row + ', ' + col(rec, f) + ' ' + q(val) + ' contains ' + r[1].map(q).join(', ') + ', which usually does not belong on a slide. ' + (fixNames ? 'Removed, leaving ' + q(r[0]) + '.' : 'Left as it is.'));
      }
    });
  });

  /* case, fixed when the list is otherwise title case */
  ['first_name', 'last_name'].forEach(function (f) {
    var values = records.map(function (r) { return r[f]; }).filter(Boolean);
    if (values.length < 3) return;
    var caps = values.filter(isAllCaps).length, lower = values.filter(isAllLower).length;
    if (caps > values.length * 0.5 || lower > values.length * 0.5) return;
    records.forEach(function (rec) {
      var val = rec[f];
      if (!val) return;
      if (isAllCaps(val) || isAllLower(val)) {
        var how = isAllCaps(val) ? 'ALL CAPS' : 'all lower case';
        var fixed = smartTitleCase(val);
        if (fixNames) rec[f] = fixed;
        warn('warning', 'name_case', rec, 'Row ' + rec.row + ', ' + col(rec, f) + ' ' + q(val) + ' is ' + how + ' in a list that is otherwise title case. ' + (fixNames ? 'Changed to ' + q(fixed) + '.' : 'Left as it is.'));
      }
    });
  });

  records.forEach(composeName);

  /* blank names, both blank drops the row */
  var kept = [], skipped = [];
  records.forEach(function (rec) {
    var first = rec.first_name, last = rec.last_name;
    if (!first && !last) {
      skipped.push(rec);
      warn('error', 'blank_name', rec, 'Row ' + rec.row + ' has no name. Nobody to put on a slide, so the row is skipped and everyone after it moves up one place.');
      return;
    }
    if (!hasFull) {
      if (!first && last && last.indexOf(' ') < 0) warn('warning', 'blank_first_name', rec, 'Row ' + rec.row + ', ' + col(rec, 'first_name') + ' is blank. Only ' + q(last) + ' will show.');
      else if (!last && first && first.indexOf(' ') < 0) warn('warning', 'blank_last_name', rec, 'Row ' + rec.row + ', ' + col(rec, 'last_name') + ' is blank. Only ' + q(first) + ' will show.');
    } else if (!last) {
      warn('warning', 'blank_last_name', rec, 'Row ' + rec.row + ', ' + q(first) + ' is a single word, so it sorts as a surname with no first name.');
    }
    kept.push(rec);
  });

  /* duplicates across the whole list */
  var seen = {};
  kept.forEach(function (rec) {
    var key = foldAccents(rec.first_name).toLowerCase().trim() + '|' + foldAccents(rec.last_name).toLowerCase().trim();
    (seen[key] = seen[key] || []).push(rec);
  });
  Object.keys(seen).forEach(function (k) {
    var g = seen[k];
    if (g.length < 2) return;
    var rows = g.map(function (r) { return r.row; }).join(', ');
    g.slice(1).forEach(function (rec) {
      warn('warning', 'duplicate_person', rec, 'Row ' + rec.row + ', ' + personLabel(rec) + ' appears more than once (rows ' + rows + '). The name shows twice. Delete one row if it is the same person.');
    });
  });

  /* particles and hyphens, info */
  var particleRecs = kept.filter(function (r) { return r.last_name && hasParticle(r.last_name); });
  if (particleRecs.length) {
    var uk = settings.particles === 'uk';
    var ex = particleRecs.slice(0, 3).map(function (r) { return q(r.last_name); }).join(', ');
    var more = particleRecs.length > 3 ? ' and ' + (particleRecs.length - 3) + ' more' : '';
    warn('info', 'surname_particle', null, particleRecs.length + ' surname' + (particleRecs.length > 1 ? 's have' : ' has') + ' a particle, ' + ex + more + '. Sorting ' + (uk ? 'UK style under the particle, so van der Berg is under V. Switch Particles to Dutch to sort under B.' : 'Dutch style under the main name, so van der Berg is under B. Switch Particles to UK to sort under V.'));
  }
  kept.forEach(function (rec) {
    var last = rec.last_name;
    if (last && !hasParticle(last) && (last.replace(/^-+|-+$/g, '').indexOf('-') >= 0 || last.trim().indexOf(' ') >= 0)) {
      warn('info', 'double_barrelled', rec, 'Row ' + rec.row + ', ' + col(rec, 'last_name') + ' ' + q(last) + ' is ' + (last.indexOf('-') >= 0 ? 'hyphenated' : 'double-barrelled') + '. Kept together as one surname.');
    }
  });

  /* characters outside Latin-1 */
  kept.forEach(function (rec) {
    ['first_name', 'last_name', 'extra'].forEach(function (f) {
      var val = toText(rec[f]);
      var m = NON_LATIN.exec(val);
      if (m) warn('warning', 'non_latin_character', rec, 'Row ' + rec.row + ', ' + col(rec, f) + ' ' + q(val) + ' has the character ' + q(m[0]) + ', which many brand fonts do not have. Check it in PowerPoint.');
    });
  });

  /* table, blank or non-numeric or out of range */
  if (opts.tableColumn) {
    var vals = kept.map(function (r) { return r.table; });
    var nonblank = vals.filter(function (v) { return v !== null && v !== undefined && v !== ''; });
    var numeric = nonblank.filter(function (v) { return typeof v === 'number' || (typeof v === 'string' && /^\d+$/.test(v.trim())); });
    var mostlyNumeric = numeric.length > nonblank.length * 0.6;
    kept.forEach(function (rec) {
      var v = rec.table, text = toText(v);
      if (!text) warn('warning', 'blank_table', rec, 'Row ' + rec.row + ', ' + col(rec, 'table') + ' is blank for ' + personLabel(rec) + '. They still appear in the list, with an empty table cell, and sort after everyone with a value.');
      else if (mostlyNumeric && !(typeof v === 'number' || /^\d+$/.test(text))) warn('warning', 'non_numeric_table', rec, 'Row ' + rec.row + ', ' + col(rec, 'table') + ' is ' + q(text) + ' for ' + personLabel(rec) + ' where the other rows hold numbers. It sorts as text after the numbers, so it lands at the end of the deck.');
    });
    var nums = kept.map(function (r) { return toText(r.table); }).filter(function (t) { return /^\d+$/.test(t); }).map(Number);
    if (nums.length >= 5) {
      var sorted = nums.slice().sort(function (a, b) { return a - b; });
      var cutoff = sorted[sorted.length - 1];
      for (var i = sorted.length - 1; i > 0; i--) {
        if (sorted[i] - sorted[i - 1] > 10 && sorted[i] > 2 * sorted[i - 1]) cutoff = sorted[i - 1];
      }
      if (cutoff !== sorted[sorted.length - 1]) {
        kept.forEach(function (rec) {
          var t = toText(rec.table);
          if (/^\d+$/.test(t) && Number(t) > cutoff) warn('warning', 'table_out_of_range', rec, 'Row ' + rec.row + ', ' + col(rec, 'table') + ' is ' + t + ' for ' + personLabel(rec) + ' but the other tables stop at ' + cutoff + '. Probably a typo. Check the sheet, otherwise they land at the end.');
        });
      }
    }
  }

  var order = { error: 0, warning: 1, info: 2 };
  warnings.sort(function (a, b) { return order[a.level] - order[b.level] || (a.row || 0) - (b.row || 0); });
  return { records: kept, skipped: skipped, warnings: warnings };
}

/* Draw odd characters so a user can see them. */
function visible(s) {
  return String(s == null ? '' : s).replace(/\u00a0/g, '\u2423').replace(/[\u200b\u200c\u200d\u2060\ufeff]/g, '\u00b7').replace(/\t/g, '\u2192').replace(/[\r\n\u000b\u000c\u2028\u2029]/g, '\u21b5').replace(/^ | $/g, '\u2423');
}

/* ======================================================================
   5  engine  (engine_flow.py)
   ====================================================================== */

/* How many rows go on each slide. Without balance every slide is full except the last. */
function chunkSizes(total, perSlide, balance) {
  if (total <= 0) return [];
  perSlide = Math.max(1, perSlide);
  var nSlides = Math.ceil(total / perSlide);
  var sizes = [];
  if (!balance) {
    for (var i = 0; i < nSlides - 1; i++) sizes.push(perSlide);
    sizes.push(total - perSlide * (nSlides - 1));
    return sizes;
  }
  var base = Math.floor(total / nSlides), extra = total % nSlides;
  for (var j = 0; j < nSlides; j++) sizes.push(j < extra ? base + 1 : base);
  return sizes;
}

function chunkSizesFirst(total, first, rest, balance) {
  if (total <= 0) return [];
  if (first === rest) return chunkSizes(total, rest, balance);
  if (total <= first) return [total];
  return [first].concat(chunkSizes(total - first, rest, balance));
}

function flowSettings(o) {
  o = o || {};
  return {
    listFields: o.listFields || ['name', 'table'],
    rowsPerSlide: o.rowsPerSlide || 30,
    sort: o.sort || ['last_name', 'first_name'],
    breakField: o.breakField || null,
    balance: !!o.balance,
    rangeField: o.rangeField || null,
    headerField: o.headerField || null,
    sortSettings: sortSettings(o.sortSettings)
  };
}

/* Chunk into [{key, rows}] slides. key is null without a break field. */
function chunkRecords(recs, settings, breakFirstRows) {
  settings = flowSettings(settings);
  var out = [], pos, sizes;
  if (!settings.breakField) {
    sizes = chunkSizes(recs.length, settings.rowsPerSlide, settings.balance);
    pos = 0;
    sizes.forEach(function (size) { out.push({ key: null, rows: recs.slice(pos, pos + size) }); pos += size; });
    return out;
  }
  var runs = [];
  recs.forEach(function (rec) {
    var val = toText(rec[settings.breakField]);
    if (runs.length && runs[runs.length - 1].key === val) runs[runs.length - 1].rows.push(rec);
    else runs.push({ key: val, rows: [rec] });
  });
  var first = breakFirstRows || settings.rowsPerSlide;
  runs.forEach(function (run) {
    var s = chunkSizesFirst(run.rows.length, first, settings.rowsPerSlide, settings.balance);
    var p = 0;
    s.forEach(function (size) { out.push({ key: run.key, rows: run.rows.slice(p, p + size) }); p += size; });
  });
  return out;
}

function sortFlow(recs, settings) {
  settings = flowSettings(settings);
  return sortRecords(recs, sortRule(settings.sort, settings.sortSettings));
}

/* Header text for a break value. A bare number reads as Table 5. */
function headerLabel(value) {
  var t = toText(value);
  if (!t) return '';
  if (/^\d+[A-Za-z]?$/.test(t)) return 'Table ' + t;
  return t;
}

/* Plan the deck. Returns {slides, warnings}. Each slide has index, layout,
   key, rows (records) and values keyed by placeholder name. */
function planFlow(recs, settings, skipped) {
  settings = flowSettings(settings);
  var plan = { slides: [], warnings: [] };
  var ordered = sortFlow(recs, settings);
  var gaps = {};
  if (skipped && skipped.length) {
    var everyone = sortFlow(recs.concat(skipped), settings);
    var pending = [];
    everyone.forEach(function (r) {
      if (skipped.indexOf(r) >= 0) pending.push(r);
      else if (pending.length) { gaps[r.row] = pending; pending = []; }
    });
    if (pending.length) gaps[-1] = pending;
  }
  var chunks = chunkRecords(ordered, settings);
  var total = chunks.length;
  var rangeField = settings.rangeField || settings.listFields[0];
  var headerField = settings.headerField || settings.breakField;
  var lastKey = {};
  chunks.forEach(function (chunk, i0) {
    var i = i0 + 1;
    var startsBreak = settings.breakField !== null && chunk.key !== lastKey;
    lastKey = chunk.key;
    var layout = (startsBreak && settings.breakField) ? 'LIST BREAK' : 'LIST';
    var sp = { index: i, layout: layout, key: chunk.key, rows: chunk.rows, values: {}, warnings: [] };
    chunk.rows.forEach(function (rec, n0) {
      settings.listFields.forEach(function (f) { sp.values['row.' + (n0 + 1) + '.' + f] = toText(rec[f]); });
    });
    var missing = [];
    chunk.rows.forEach(function (m) { if (gaps[m.row]) missing = missing.concat(gaps[m.row]); });
    if (i === total && gaps[-1]) missing = missing.concat(gaps[-1]);
    missing.forEach(function (r) { sp.warnings.push('Row ' + r.row + ' was skipped for having no name and would have been on slide ' + i + '. Everyone after it has moved up one place.'); });
    if (headerField) {
      sp.values.header = chunk.rows.length ? headerLabel(chunk.rows[0][headerField]) : '';
      if (settings.breakField && !sp.values.header) {
        var names = chunk.rows.slice(0, 3).map(personLabel).join(', ');
        var more = chunk.rows.length > 3 ? ' and ' + (chunk.rows.length - 3) + ' more' : '';
        var rows = chunk.rows.map(function (r) { return r.row; });
        var lo = Math.min.apply(null, rows), hi = Math.max.apply(null, rows);
        sp.warnings.push('Table is blank for ' + names + more + ' (' + (lo === hi ? 'row ' + lo : 'rows ' + lo + ' to ' + hi) + '), so slide ' + i + ' has an empty header.');
      }
    }
    if (rangeField && chunk.rows.length) {
      /* a person with no surname still needs to show at the end of a range */
      var firstRec = chunk.rows[0], lastRec = chunk.rows[chunk.rows.length - 1];
      sp.values['range.start'] = toText(firstRec[rangeField]) || toText(firstRec.name);
      sp.values['range.end'] = toText(lastRec[rangeField]) || toText(lastRec.name);
    }
    sp.values['page.number'] = String(i);
    sp.values['page.count'] = String(total);
    plan.slides.push(sp);
    plan.warnings = plan.warnings.concat(sp.warnings);
  });
  return plan;
}

/* ======================================================================
   6  geometry
   ====================================================================== */

var EMU_IN = 914400;
var SLIDE_W = 12192000, SLIDE_H = 6858000;   /* 13.333 x 7.5 in */
function inch(v) { return Math.round(v * EMU_IN); }

/* Where everything sits on the layout. cfg: perSlide, fields (list of row
   field names), title, header, range, page. Returns boxes in EMU. */
function layoutGeometry(cfg) {
  var N = Math.max(1, cfg.perSlide | 0);
  var top = (cfg.title || cfg.header || cfg.range) ? 1.45 : 0.7;
  var bottom = cfg.page ? 6.8 : 7.0;
  var margin = 0.6, gap = 0.3;
  var cols = Math.min(4, Math.ceil(N / 15));
  var rowsPerCol = Math.ceil(N / cols);
  var avail = bottom - top;
  var rowH = Math.min(0.45, avail / rowsPerCol);
  var fontPt = Math.max(8, Math.min(20, Math.floor(rowH * 72 * 0.55)));
  var colW = (13.333 - 2 * margin - gap * (cols - 1)) / cols;
  var hasTable = cfg.fields.indexOf('table') >= 0;
  var hasExtra = cfg.fields.indexOf('extra') >= 0;
  var separate = cfg.fields.indexOf('first_name') >= 0;
  /* width from the longest value, roughly 0.55 em a character at this size */
  var em = fontPt / 72;
  var tableW = hasTable ? Math.min(colW * 0.35, Math.max(0.7, (cfg.tableChars || 3) * em * 0.6 + 0.2)) : 0;
  var extraW = hasExtra ? Math.min(colW * 0.4, Math.max(1.2, (cfg.extraChars || 12) * em * 0.55 + 0.2)) : 0;
  var nameW = colW - tableW - extraW - (hasTable || hasExtra ? 0.1 : 0);
  var rows = [];
  for (var n = 0; n < N; n++) {
    var c = Math.floor(n / rowsPerCol), r = n % rowsPerCol;
    var x = margin + c * (colW + gap), y = top + r * rowH;
    var boxes = {};
    if (separate) {
      var fw = Math.min(nameW * 0.42, Math.max(1.4, (cfg.firstChars || 10) * em * 0.6 + 0.3));
      boxes.first_name = { x: x, y: y, w: fw, h: rowH, align: 'l' };
      boxes.last_name = { x: x + fw, y: y, w: nameW - fw, h: rowH, align: 'l' };
    } else {
      boxes.name = { x: x, y: y, w: nameW, h: rowH, align: 'l' };
    }
    var nx = x + nameW + 0.1;
    if (hasExtra) { boxes.extra = { x: nx, y: y, w: extraW, h: rowH, align: 'l' }; nx += extraW; }
    if (hasTable) boxes.table = { x: nx, y: y, w: tableW, h: rowH, align: 'r' };
    rows.push(boxes);
  }
  var g = { cols: cols, rowsPerCol: rowsPerCol, rowH: rowH, fontPt: fontPt, rows: rows, top: top, bottom: bottom, margin: margin, colW: colW, extras: {} };
  if (cfg.title) g.extras.title = { x: margin, y: 0.3, w: 8.5, h: 0.55, pt: 24, bold: true, align: 'l' };
  if (cfg.header) g.extras.header = { x: margin, y: 0.85, w: 8.5, h: 0.5, pt: 20, bold: true, align: 'l' };
  if (cfg.range) {
    g.extras['range.start'] = { x: 9.3, y: 0.85, w: 1.7, h: 0.5, pt: 14, align: 'r', grey: true };
    g.extras['range.to'] = { x: 11.0, y: 0.85, w: 0.4, h: 0.5, pt: 14, align: 'ctr', grey: true, text: 'to' };
    g.extras['range.end'] = { x: 11.4, y: 0.85, w: 1.7, h: 0.5, pt: 14, align: 'l', grey: true };
  }
  if (cfg.page) {
    g.extras['page.number'] = { x: 10.9, y: 6.95, w: 0.7, h: 0.35, pt: 11, align: 'r', grey: true };
    g.extras['page.of'] = { x: 11.6, y: 6.95, w: 0.4, h: 0.35, pt: 11, align: 'ctr', grey: true, text: 'of' };
    g.extras['page.count'] = { x: 12.0, y: 6.95, w: 0.75, h: 0.35, pt: 11, align: 'l', grey: true };
  }
  return g;
}

/* ======================================================================
   7  pptx writer
   ====================================================================== */

function esc(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
/* XML 1.0 cannot carry control characters, drop them rather than corrupt the file. */
function cleanXmlText(s) { return String(s == null ? '' : s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, ''); }

var NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
var XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

function xfrm(b) {
  return '<a:xfrm><a:off x="' + inch(b.x) + '" y="' + inch(b.y) + '"/><a:ext cx="' + inch(b.w) + '" cy="' + inch(b.h) + '"/></a:xfrm>';
}

/* A placeholder on a layout. Carries the position, the text style and the
   prompt text. Slides reference it by idx and inherit all of this. */
function layoutPlaceholder(id, name, type, idx, box, style) {
  var ph = type === 'title' ? '<p:ph type="title" hasCustomPrompt="1"/>' : '<p:ph type="body" idx="' + idx + '" hasCustomPrompt="1"/>';
  var wrap = style.wrap === false ? 'none' : 'square';
  var rPr = '<a:defRPr sz="' + (style.pt * 100) + '"' + (style.bold ? ' b="1"' : '') + '>' + (style.grey ? '<a:solidFill><a:srgbClr val="7F7F7F"/></a:solidFill>' : '') + '</a:defRPr>';
  return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + esc(name) + '"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>' + ph + '</p:nvPr></p:nvSpPr>'
    + '<p:spPr>' + xfrm(box) + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
    + '<p:txBody><a:bodyPr wrap="' + wrap + '" lIns="45720" tIns="18288" rIns="45720" bIns="18288" anchor="ctr"><a:normAutofit/></a:bodyPr>'
    + '<a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="' + (style.align || 'l') + '"><a:buNone/>' + rPr + '</a:lvl1pPr></a:lstStyle>'
    + '<a:p><a:r><a:rPr lang="en-GB"/><a:t>' + esc(name) + '</a:t></a:r></a:p></p:txBody></p:sp>';
}

/* A plain text box, used for the static 'to' and 'of' on the layout and the manifest on the master. */
function textBox(id, name, box, text, style, hidden) {
  var rPr = '<a:rPr lang="en-GB" sz="' + (style.pt * 100) + '"' + (style.bold ? ' b="1"' : '') + '>' + (style.grey ? '<a:solidFill><a:srgbClr val="7F7F7F"/></a:solidFill>' : '') + '</a:rPr>';
  return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + esc(name) + '"' + (hidden ? ' hidden="1"' : '') + '/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
    + '<p:spPr>' + xfrm(box) + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
    + '<p:txBody><a:bodyPr wrap="square" lIns="45720" tIns="18288" rIns="45720" bIns="18288" anchor="ctr"' + (style.autofit ? '><a:spAutoFit/></a:bodyPr>' : '/>')
    + '<a:lstStyle/><a:p><a:pPr algn="' + (style.align || 'l') + '"/><a:r>' + rPr + '<a:t>' + esc(cleanXmlText(text)) + '</a:t></a:r></a:p></p:txBody></p:sp>';
}

/* A placeholder on a slide. No position, no run properties, only the text.
   Everything else comes from the layout. */
function slidePlaceholder(id, name, type, idx, text) {
  var ph = type === 'title' ? '<p:ph type="title"/>' : '<p:ph type="body" idx="' + idx + '"/>';
  return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + esc(name) + '"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>' + ph + '</p:nvPr></p:nvSpPr>'
    + '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-GB" dirty="0"/><a:t>' + esc(cleanXmlText(text)) + '</a:t></a:r></a:p></p:txBody></p:sp>';
}

var SPTREE_HEAD = '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

/* The placeholder list for a layout, in order, with idx numbers. Shared by
   LIST and LIST BREAK so a slide moved between them keeps its text. */
function placeholderList(geom, cfg, withHeader) {
  var list = [], idx = 10;
  if (cfg.title) list.push({ name: 'title', type: 'title', idx: null, box: geom.extras.title, style: geom.extras.title });
  if (withHeader && geom.extras.header) list.push({ name: 'header', type: 'body', idx: idx++, box: geom.extras.header, style: geom.extras.header });
  else idx++;
  if (cfg.range) {
    list.push({ name: 'range.start', type: 'body', idx: idx++, box: geom.extras['range.start'], style: geom.extras['range.start'] });
    list.push({ name: 'range.end', type: 'body', idx: idx++, box: geom.extras['range.end'], style: geom.extras['range.end'] });
  }
  if (cfg.page) {
    list.push({ name: 'page.number', type: 'body', idx: idx++, box: geom.extras['page.number'], style: geom.extras['page.number'] });
    list.push({ name: 'page.count', type: 'body', idx: idx++, box: geom.extras['page.count'], style: geom.extras['page.count'] });
  }
  idx = 100;
  geom.rows.forEach(function (boxes, n0) {
    cfg.fields.forEach(function (f) {
      var b = boxes[f];
      if (!b) return;
      list.push({ name: 'row.' + (n0 + 1) + '.' + f, type: 'body', idx: idx++, box: b, style: { pt: geom.fontPt, align: b.align, wrap: false } });
    });
  });
  return list;
}

function layoutXml(name, phs, geom, cfg) {
  var id = 2, parts = [];
  phs.forEach(function (p) { parts.push(layoutPlaceholder(id++, p.name, p.type, p.idx, p.box, p.style)); });
  if (cfg.range) parts.push(textBox(id++, 'range.to', geom.extras['range.to'], 'to', geom.extras['range.to']));
  if (cfg.page) parts.push(textBox(id++, 'page.of', geom.extras['page.of'], 'of', geom.extras['page.of']));
  return XML_HEAD + '<p:sldLayout ' + NS + ' preserve="1" userDrawn="1"><p:cSld name="' + esc(name) + '">' + SPTREE_HEAD + parts.join('') + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
}

function slideXml(sp, phs) {
  var id = 2, parts = [];
  phs.forEach(function (p) {
    var v = sp.values[p.name];
    if (p.name === 'title') v = sp.title;
    if (v === undefined || v === null) return;
    parts.push(slidePlaceholder(id++, p.name, p.type, p.idx, v));
  });
  return XML_HEAD + '<p:sld ' + NS + '><p:cSld>' + SPTREE_HEAD + parts.join('') + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
}

function masterXml(layoutCount, manifest, font) {
  var layoutIds = '';
  for (var i = 0; i < layoutCount; i++) layoutIds += '<p:sldLayoutId id="' + (2147483649 + i) + '" r:id="rId' + (i + 1) + '"/>';
  var manifestBox = textBox(2, 'NAMEDROP.MANIFEST', { x: -6.5, y: 0.3, w: 6, h: 2 }, manifest, { pt: 8, autofit: true }, true);
  function lvl(tag, sz, bold) {
    return '<a:' + tag + ' marL="0" indent="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:buNone/><a:defRPr sz="' + sz + '"' + (bold ? ' b="1"' : '') + ' kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:' + tag + '>';
  }
  var body = '';
  for (var l = 1; l <= 9; l++) body += lvl('lvl' + l + 'pPr', 1400, false);
  var other = '<a:defPPr><a:defRPr lang="en-GB"/></a:defPPr>';
  for (var m = 1; m <= 9; m++) other += lvl('lvl' + m + 'pPr', 1800, false);
  return XML_HEAD + '<p:sldMaster ' + NS + '><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
    + SPTREE_HEAD + manifestBox + '</p:spTree></p:cSld>'
    + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    + '<p:sldLayoutIdLst>' + layoutIds + '</p:sldLayoutIdLst>'
    + '<p:txStyles><p:titleStyle>' + lvl('lvl1pPr', 2400, true) + '</p:titleStyle><p:bodyStyle>' + body + '</p:bodyStyle><p:otherStyle>' + other + '</p:otherStyle></p:txStyles></p:sldMaster>';
}

function themeXml(font) {
  var f = esc(font || 'Calibri');
  function fonts(tag) { return '<a:' + tag + '><a:latin typeface="' + f + '"/><a:ea typeface=""/><a:cs typeface=""/></a:' + tag + '>'; }
  return XML_HEAD + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="List Builder"><a:themeElements>'
    + '<a:clrScheme name="List Builder"><a:dk1><a:srgbClr val="1A1A1A"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
    + '<a:accent1><a:srgbClr val="E7302A"/></a:accent1><a:accent2><a:srgbClr val="4A78B5"/></a:accent2><a:accent3><a:srgbClr val="A6CE4E"/></a:accent3><a:accent4><a:srgbClr val="E0A020"/></a:accent4><a:accent5><a:srgbClr val="9B7BD4"/></a:accent5><a:accent6><a:srgbClr val="7FD1AE"/></a:accent6>'
    + '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
    + '<a:fontScheme name="List Builder">' + fonts('majorFont') + fonts('minorFont') + '</a:fontScheme>'
    + '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
    + '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
    + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    + '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
    + '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
}

function rels(list) {
  return XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + list.map(function (r) { return '<Relationship Id="' + r.id + '" Type="' + (r.fullType || 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + r.type) + '" Target="' + r.target + '"/>'; }).join('')
    + '</Relationships>';
}

function isoNow(d) { return d.toISOString().replace(/\.\d+Z$/, 'Z'); }

/* The NAMEDROP manifest, the same keys scaffold.py writes plus the generator. */
function manifestJson(cfg, listFields, extras) {
  var placeholders = [];
  for (var n = 1; n <= cfg.perSlide; n++) listFields.forEach(function (f) { placeholders.push('row.' + n + '.' + f); });
  placeholders = placeholders.concat(extras);
  var m = {
    namedrop: LB.namedropVersion,
    engine: 'flow',
    max_per_slide: cfg.perSlide,
    size: '16:9',
    slide_mm: [338.7, 190.5],
    extras: extras,
    budgets: {},
    placeholders: placeholders,
    list_fields: listFields,
    generator: 'SlideSize List Builder ' + LB.version
  };
  return JSON.stringify(m, null, 1);
}

/* The geometry inputs for a plan, with the longest value per field so the
   boxes are wide enough. Shared by the page preview and the writer. */
function geometryConfig(plan, cfg) {
  var longest = { table: 0, extra: 0, first_name: 0 };
  plan.slides.forEach(function (sp) {
    sp.rows.forEach(function (rec) {
      Object.keys(longest).forEach(function (f) { var n = toText(rec[f]).length; if (n > longest[f]) longest[f] = n; });
    });
  });
  return { perSlide: cfg.perSlide, fields: cfg.fields, title: !!cfg.title, header: !!cfg.breakOn, range: !!cfg.range, page: !!cfg.page,
    tableChars: longest.table, extraChars: longest.extra, firstChars: longest.first_name };
}

/* Build the deck. plan from planFlow, cfg {perSlide, fields, title (text or
   ''), range, page, breakOn, font, docTitle, created}. Returns the JSZip
   instance, caller picks the output type. */
function buildDeck(plan, cfg, JSZip) {
  var fields = cfg.fields.slice();
  var hasHeader = !!cfg.breakOn;
  var geomCfg = geometryConfig(plan, cfg);
  var geom = layoutGeometry(geomCfg);
  var layouts = [{ name: 'LIST', phs: placeholderList(geom, geomCfg, false) }];
  if (hasHeader) layouts.push({ name: 'LIST BREAK', phs: placeholderList(geom, geomCfg, true) });
  var extras = [];
  if (cfg.title) extras.push('title');
  if (hasHeader) extras.push('header');
  if (cfg.range) extras.push('range.start', 'range.end');
  if (cfg.page) extras.push('page.number', 'page.count');

  var zip = new JSZip();
  var ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
    '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>',
    '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>',
    '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>',
    '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'];

  /* layouts */
  var masterRels = [];
  layouts.forEach(function (lay, i) {
    var n = i + 1;
    zip.file('ppt/slideLayouts/slideLayout' + n + '.xml', layoutXml(lay.name, lay.phs, geom, geomCfg));
    zip.file('ppt/slideLayouts/_rels/slideLayout' + n + '.xml.rels', rels([{ id: 'rId1', type: 'slideMaster', target: '../slideMasters/slideMaster1.xml' }]));
    ct.push('<Override PartName="/ppt/slideLayouts/slideLayout' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>');
    masterRels.push({ id: 'rId' + n, type: 'slideLayout', target: '../slideLayouts/slideLayout' + n + '.xml' });
  });
  masterRels.push({ id: 'rId' + (layouts.length + 1), type: 'theme', target: '../theme/theme1.xml' });
  zip.file('ppt/slideMasters/slideMaster1.xml', masterXml(layouts.length, manifestJson(cfg, fields, extras), cfg.font));
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels(masterRels));
  zip.file('ppt/theme/theme1.xml', themeXml(cfg.font));

  /* slides */
  var presRels = [{ id: 'rId1', type: 'slideMaster', target: 'slideMasters/slideMaster1.xml' }, { id: 'rId2', type: 'theme', target: 'theme/theme1.xml' },
    { id: 'rId3', type: 'presProps', target: 'presProps.xml' }, { id: 'rId4', type: 'viewProps', target: 'viewProps.xml' }, { id: 'rId5', type: 'tableStyles', target: 'tableStyles.xml' }];
  var sldIds = '';
  plan.slides.forEach(function (sp, i) {
    var n = i + 1;
    var lay = layouts[sp.layout === 'LIST BREAK' && hasHeader ? 1 : 0];
    var layIndex = layouts.indexOf(lay) + 1;
    sp.title = cfg.title || undefined;
    zip.file('ppt/slides/slide' + n + '.xml', slideXml(sp, lay.phs));
    zip.file('ppt/slides/_rels/slide' + n + '.xml.rels', rels([{ id: 'rId1', type: 'slideLayout', target: '../slideLayouts/slideLayout' + layIndex + '.xml' }]));
    ct.push('<Override PartName="/ppt/slides/slide' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>');
    var rid = 'rId' + (5 + n);
    presRels.push({ id: rid, type: 'slide', target: 'slides/slide' + n + '.xml' });
    sldIds += '<p:sldId id="' + (255 + n) + '" r:id="' + rid + '"/>';
  });

  zip.file('ppt/presentation.xml', XML_HEAD + '<p:presentation ' + NS + ' saveSubsetFonts="1" autoCompressPictures="0">'
    + '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>' + sldIds + '</p:sldIdLst>'
    + '<p:sldSz cx="' + SLIDE_W + '" cy="' + SLIDE_H + '"/><p:notesSz cx="6858000" cy="9144000"/>'
    + '<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-GB"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle>'
    + '</p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', rels(presRels));
  zip.file('ppt/presProps.xml', XML_HEAD + '<p:presentationPr ' + NS + '/>');
  zip.file('ppt/viewProps.xml', XML_HEAD + '<p:viewPr ' + NS + '><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>');
  zip.file('ppt/tableStyles.xml', XML_HEAD + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>');

  var when = isoNow(cfg.created || new Date());
  zip.file('docProps/core.xml', XML_HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    + '<dc:title>' + esc(cfg.docTitle || 'List') + '</dc:title><dc:creator>SlideSize List Builder</dc:creator><cp:lastModifiedBy>SlideSize List Builder</cp:lastModifiedBy>'
    + '<dcterms:created xsi:type="dcterms:W3CDTF">' + when + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + when + '</dcterms:modified></cp:coreProperties>');
  zip.file('docProps/app.xml', XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>SlideSize List Builder</Application><Slides>' + plan.slides.length + '</Slides><PresentationFormat>Widescreen</PresentationFormat></Properties>');
  zip.file('[Content_Types].xml', XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + ct.join('') + '</Types>');
  zip.file('_rels/.rels', rels([{ id: 'rId1', type: 'officeDocument', target: 'ppt/presentation.xml' },
    { id: 'rId2', fullType: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', target: 'docProps/core.xml' },
    { id: 'rId3', type: 'extended-properties', target: 'docProps/app.xml' }]));
  return zip;
}

/* ======================================================================
   One call from source to plan, shared by the page and the headless build.
   ====================================================================== */

/* opts: mapping {first, last, full, table, extra}, perSlide, sort
   ('surname'|'first'|'table'), breakOn, particles ('uk'|'dutch'), nameFormat
   ('full'|'separate'), range, page, fixNames, balance. */
function prepare(table, opts) {
  var mapping = opts.mapping;
  var recs = buildRecords(table, mapping);
  var v = validate(recs, { sortSettings: { particles: opts.particles }, fullNameColumn: !!mapping.full, tableColumn: !!mapping.table, fixNames: opts.fixNames });
  var sort;
  if (opts.sort === 'first') sort = ['first_name', 'last_name'];
  else if (opts.sort === 'table') sort = ['table', 'last_name', 'first_name'];
  else sort = ['last_name', 'first_name'];
  if (opts.breakOn && sort[0] !== 'table') sort = ['table'].concat(sort);
  var fields = opts.nameFormat === 'separate' ? ['first_name', 'last_name'] : ['name'];
  if (mapping.extra) fields.push('extra');
  if (mapping.table) fields.push('table');
  var settings = {
    listFields: fields, rowsPerSlide: opts.perSlide, sort: sort,
    breakField: opts.breakOn && mapping.table ? 'table' : null,
    balance: !!opts.balance, rangeField: 'last_name',
    sortSettings: { particles: opts.particles }
  };
  var plan = planFlow(v.records, settings, v.skipped);
  return { records: v.records, skipped: v.skipped, warnings: v.warnings, plan: plan, settings: settings, fields: fields };
}

function stamp(d) {
  d = d || new Date();
  return d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) + '_' + pad2(d.getHours()) + pad2(d.getMinutes());
}

function outputName(sourceName, d) {
  var base = String(sourceName || 'list').replace(/\.[^.]+$/, '').replace(/[\\/:*?"<>|]+/g, '-').trim() || 'list';
  return base + '_lists_' + stamp(d) + '.pptx';
}

LB.PARTICLES = PARTICLES;
LB.foldAccents = foldAccents;
LB.splitParticles = splitParticles;
LB.hasParticle = hasParticle;
LB.normaliseMc = normaliseMc;
LB.sortSettings = sortSettings;
LB.surnameSortForm = surnameSortForm;
LB.naturalKey = naturalKey;
LB.textKey = textKey;
LB.valueKey = valueKey;
LB.compareKeys = compareKeys;
LB.sortRule = sortRule;
LB.sortRecords = sortRecords;
LB.smartTitleCase = smartTitleCase;
LB.isAllCaps = isAllCaps;
LB.isAllLower = isAllLower;
LB.stripHonorifics = stripHonorifics;
LB.isHyphenated = isHyphenated;
LB.splitFullName = splitFullName;
LB.toText = toText;
LB.cleanText = cleanText;
LB.cleanCell = cleanCell;
LB.detectHeaderRow = detectHeaderRow;
LB.buildTable = buildTable;
LB.decodeBytes = decodeBytes;
LB.sniffDelimiter = sniffDelimiter;
LB.parseCsvText = parseCsvText;
LB.readCsv = readCsv;
LB.gridFromSheet = gridFromSheet;
LB.readSheet = readSheet;
LB.proposeMapping = proposeMapping;
LB.buildRecords = buildRecords;
LB.validate = validate;
LB.visible = visible;
LB.chunkSizes = chunkSizes;
LB.chunkSizesFirst = chunkSizesFirst;
LB.chunkRecords = chunkRecords;
LB.sortFlow = sortFlow;
LB.planFlow = planFlow;
LB.headerLabel = headerLabel;
LB.layoutGeometry = layoutGeometry;
LB.geometryConfig = geometryConfig;
LB.placeholderList = placeholderList;
LB.buildDeck = buildDeck;
LB.prepare = prepare;
LB.outputName = outputName;
LB.manifestJson = manifestJson;

return LB;
}));
