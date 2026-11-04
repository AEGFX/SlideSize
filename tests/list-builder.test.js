/* List Builder tests. The cases are copied from the NAMEDROP reference tests
   (test_rules.py, test_engines.py, test_source.py) so the JavaScript engine
   can be held to the same behaviour as the Python.

   Run from the repository root with:  node --test tests/list-builder.test.js
   The browser runner tests/list-builder.html loads this same file. */
'use strict';
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    const test = require('node:test');
    const assert = require('node:assert/strict');
    const LB = require('../list-builder.js');
    let JSZip = null;
    try { JSZip = require('jszip'); } catch (e) { /* optional, the build test is skipped without it */ }
    factory(test, assert, LB, JSZip);
  } else {
    root.listBuilderTests = factory;
  }
}(typeof self !== 'undefined' ? self : this, function (test, assert, LB, JSZip) {

function names(recs) { return recs.map(r => r.last_name); }
function recs(list) { return list.map(n => ({ last_name: n })); }

/* ---------- rules.py ---------- */

test('split particles', () => {
  assert.deepEqual(LB.splitParticles('van der Berg'), ['van der', 'Berg']);
  assert.deepEqual(LB.splitParticles('van den Broek'), ['van den', 'Broek']);
  assert.deepEqual(LB.splitParticles('de Jong'), ['de', 'Jong']);
  assert.deepEqual(LB.splitParticles('Smith'), ['', 'Smith']);
  assert.deepEqual(LB.splitParticles('De Niro'), ['', 'De Niro']);
  assert.deepEqual(LB.splitParticles('VAN DER BERG'), ['VAN DER', 'BERG']);
  assert.equal(LB.hasParticle('von Trapp'), true);
  assert.equal(LB.hasParticle('Hamilton-Price'), false);
});

test('UK particles sort under the particle', () => {
  const rule = LB.sortRule(['last_name'], { particles: 'uk' });
  assert.deepEqual(names(LB.sortRecords(recs(['Zimmer', 'van der Berg', 'Adams', 'de Jong']), rule)), ['Adams', 'de Jong', 'van der Berg', 'Zimmer']);
});

test('Dutch particles sort under the core', () => {
  const rule = LB.sortRule(['last_name'], { particles: 'dutch' });
  assert.deepEqual(names(LB.sortRecords(recs(['Zimmer', 'van der Berg', 'Adams', 'de Jong', 'Bakker']), rule)), ['Adams', 'Bakker', 'van der Berg', 'de Jong', 'Zimmer']);
});

test('Mc and Mac setting', () => {
  const list = ['MacGregor', 'Mason', 'McGregor', 'Mabbott'];
  assert.deepEqual(names(LB.sortRecords(recs(list), LB.sortRule(['last_name'], { mcMacTogether: false }))), ['Mabbott', 'MacGregor', 'Mason', 'McGregor']);
  assert.deepEqual(names(LB.sortRecords(recs(list), LB.sortRule(['last_name'], { mcMacTogether: true }))), ['Mabbott', 'MacGregor', 'McGregor', 'Mason']);
});

test('surname sort form', () => {
  const s = { particles: 'dutch', mcMacTogether: true };
  assert.equal(LB.surnameSortForm('van der Berg', s), 'Berg van der');
  assert.equal(LB.surnameSortForm('McDonald', s), 'macDonald');
});

test('accent folding', () => {
  assert.equal(LB.foldAccents('M\u00fcller'), 'Muller');
  assert.equal(LB.foldAccents('\u00d8degaard'), 'Odegaard');
  assert.equal(LB.foldAccents('\u0141ukasz'), 'Lukasz');
  assert.equal(LB.foldAccents('Soci\u00e9t\u00e9 G\u00e9n\u00e9rale'), 'Societe Generale');
  assert.deepEqual(names(LB.sortRecords(recs(['Mulder', 'M\u00fcller', 'Mullen']), LB.sortRule(['last_name']))), ['Mulder', 'Mullen', 'M\u00fcller']);
});

test('case insensitive', () => {
  assert.deepEqual(names(LB.sortRecords(recs(['smith', 'ADAMS', 'Jones']), LB.sortRule(['last_name']))), ['ADAMS', 'Jones', 'smith']);
});

test('natural numeric', () => {
  assert.ok(LB.compareKeys(LB.naturalKey('Table 2'), LB.naturalKey('Table 10')) < 0);
  const r = ['Table 10', 'Table 2', 'Table 1', 'Table 2A'].map(t => ({ table: t }));
  assert.deepEqual(LB.sortRecords(r, LB.sortRule(['table'])).map(x => x.table), ['Table 1', 'Table 2', 'Table 2A', 'Table 10']);
});

test('numbers before text and blanks last', () => {
  const r = ['TBC', 12, null, '12A', 3].map(t => ({ table: t }));
  assert.deepEqual(LB.sortRecords(r, LB.sortRule(['table'])).map(x => x.table), [3, 12, '12A', 'TBC', null]);
});

test('tie break and descending', () => {
  const r = [{ last_name: 'Smith', first_name: 'Zoe' }, { last_name: 'Smith', first_name: 'Adam' }, { last_name: 'Adams', first_name: 'Bob' }];
  assert.deepEqual(LB.sortRecords(r, LB.sortRule(['last_name', 'first_name'])).map(x => x.first_name), ['Bob', 'Adam', 'Zoe']);
  assert.deepEqual(LB.sortRecords(r, LB.sortRule(['-last_name', 'first_name'])).map(x => x.first_name), ['Adam', 'Zoe', 'Bob']);
});

test('sort is stable on source order', () => {
  const r = [0, 1, 2, 3, 4].map(i => ({ last_name: 'Smith', n: i }));
  assert.deepEqual(LB.sortRecords(r, LB.sortRule(['last_name'])).map(x => x.n), [0, 1, 2, 3, 4]);
});

test('smart title case', () => {
  const cases = [['JOHN SMITH', 'John Smith'], ['ANNA VAN DER BERG', 'Anna van der Berg'], ['MCDONALD', 'McDonald'], ['MACGREGOR', 'MacGregor'],
    ["O'NEILL", "O'Neill"], ['SMITH-JONES', 'Smith-Jones'], ['liam', 'Liam'], ["o'brien", "O'Brien"], ['MASON', 'Mason']];
  cases.forEach(c => assert.equal(LB.smartTitleCase(c[0]), c[1]));
});

test('is all caps', () => {
  assert.equal(LB.isAllCaps('MARTIN'), true);
  assert.equal(LB.isAllCaps('Martin'), false);
  assert.equal(LB.isAllCaps('A'), false);
});

test('strip honorifics', () => {
  assert.deepEqual(LB.stripHonorifics('Dr Priya'), ['Priya', ['Dr']]);
  assert.deepEqual(LB.stripHonorifics('Hudson Jr'), ['Hudson', ['Jr']]);
  assert.deepEqual(LB.stripHonorifics('Prof. Sir James Smith OBE'), ['James Smith', ['Prof.', 'Sir', 'OBE']]);
  assert.deepEqual(LB.stripHonorifics('Smith'), ['Smith', []]);
});

test('split a full name', () => {
  assert.deepEqual(LB.splitFullName('Hannah Fischer'), ['Hannah', 'Fischer']);
  assert.deepEqual(LB.splitFullName('Anna van der Berg'), ['Anna', 'van der Berg']);
  assert.deepEqual(LB.splitFullName('Mary Jane Smith'), ['Mary Jane', 'Smith']);
  assert.deepEqual(LB.splitFullName('Dr Amara Okafor'), ['Amara', 'Okafor']);
  assert.deepEqual(LB.splitFullName('Madonna'), ['Madonna', '']);
});

/* ---------- source.py ---------- */

test('clean text', () => {
  assert.deepEqual(LB.cleanText('Lars '), ['Lars', ['trailing space']]);
  assert.deepEqual(LB.cleanText(' Nielsen'), ['Nielsen', ['leading space']]);
  assert.deepEqual(LB.cleanText('Kai\u00a0'), ['Kai', ['non-breaking space', 'trailing space']]);
  assert.deepEqual(LB.cleanText('a\u200bb'), ['ab', ['zero width character']]);
  assert.deepEqual(LB.cleanText('two  spaces'), ['two spaces', ['double spaces']]);
  assert.deepEqual(LB.cleanText('line\nbreak'), ['line break', ['line break']]);
});

test('header row detection', () => {
  assert.equal(LB.detectHeaderRow([['First Name', 'Last Name', 'Table'], ['Ann', 'Lee', 1], ['Bob', 'Day', 2]]).index, 0);
  /* a title above the header */
  assert.equal(LB.detectHeaderRow([['Guest list', null, null], [null, null, null], ['First Name', 'Last Name', 'Table'], ['Ann', 'Lee', 1]]).index, 2);
  /* all text sheet, the data row repeats a value below so it is never the header */
  assert.equal(LB.detectHeaderRow([['Name', 'Role'], ['Ann Lee', 'Speaker'], ['Bob Day', 'Speaker']]).index, 0);
  /* CSV numbers are text but count as numbers */
  assert.equal(LB.detectHeaderRow([['First', 'Last', 'Table'], ['Ann', 'Lee', '1']]).index, 0);
});

test('csv parsing, quotes, embedded newlines, delimiter sniffing', () => {
  const g = LB.parseCsvText('a,b\n"x, y","line\nbreak"\n"say ""hi""",2\n');
  assert.deepEqual(g, [['a', 'b'], ['x, y', 'line\nbreak'], ['say "hi"', '2']]);
  assert.equal(LB.sniffDelimiter('a;b;c\n1;2;3\n'), ';');
  assert.equal(LB.sniffDelimiter('a\tb\n1\t2\n'), '\t');
  assert.equal(LB.sniffDelimiter('a,b\n1,2\n'), ',');
});

test('csv decoding, BOM and Windows-1252 fallback', () => {
  const bom = new Uint8Array([0xEF, 0xBB, 0xBF, 0x61, 0x2C, 0x62]);
  assert.equal(LB.decodeBytes(bom).encoding, 'utf-8-sig');
  assert.equal(LB.decodeBytes(bom).text, 'a,b');
  const cp = new Uint8Array([0xD8, 0x64, 0x65]);   /* \u00d8de in cp1252, invalid UTF-8 */
  const d = LB.decodeBytes(cp);
  assert.equal(d.encoding, 'cp1252');
  assert.equal(d.text, '\u00d8de');
});

test('table from a grid skips empty rows and names unnamed columns', () => {
  const t = LB.buildTable([['First', null, 'First'], ['Ann', 'x', '1'], [null, null, null], ['Bob', null, '2']]);
  assert.deepEqual(t.headers, ['First', 'Column 2', 'First (2)']);
  assert.deepEqual(t.rows.map(r => r.number), [2, 4]);
});

/* ---------- mapping.py ---------- */

test('column proposal by header and by content', () => {
  const t = LB.buildTable([['Forename', 'Surname', 'Seat'], ['Ann', 'Lee', 12], ['Bob', 'Day', 3], ['Cy', 'Fox', 7], ['Di', 'Ray', 1], ['Ed', 'Kim', 9]]);
  const p = LB.proposeMapping(t);
  assert.deepEqual([p.first, p.last, p.table, p.full], ['Forename', 'Surname', 'Seat', null]);
  const t2 = LB.buildTable([['Guest name', 'Company'], ['Ann Lee', 'Acme'], ['Bob Day', 'Zed']]);
  const p2 = LB.proposeMapping(t2);
  assert.equal(p2.full, 'Guest name');
});

/* ---------- engine_flow.py ---------- */

test('chunk sizes', () => {
  const cases = [[0, 30, false, []], [10, 30, false, [10]], [30, 30, false, [30]], [31, 30, false, [30, 1]], [63, 30, false, [30, 30, 3]],
    [63, 30, true, [21, 21, 21]], [64, 30, true, [22, 21, 21]], [31, 30, true, [16, 15]], [300, 30, true, Array(10).fill(30)]];
  cases.forEach(c => assert.deepEqual(LB.chunkSizes(c[0], c[1], c[2]), c[3], JSON.stringify(c)));
});

function rec(row, v) { return Object.assign({ row, first_name: '', last_name: 'n' + row, table: null }, v || {}); }

test('chunk without break', () => {
  const r = [1, 2, 3, 4, 5, 6, 7].map(i => rec(i));
  const chunks = LB.chunkRecords(r, { rowsPerSlide: 3 });
  assert.deepEqual(chunks.map(c => [c.key, c.rows.map(x => x.row)]), [[null, [1, 2, 3]], [null, [4, 5, 6]], [null, [7]]]);
});

test('chunk with break field', () => {
  const r = [rec(1, { table: 1 }), rec(2, { table: 1 }), rec(3, { table: 2 }), rec(4, { table: 2 }), rec(5, { table: 2 }), rec(6, { table: 2 }), rec(7, { table: 3 })];
  const chunks = LB.chunkRecords(r, { rowsPerSlide: 3, breakField: 'table' });
  assert.deepEqual(chunks.map(c => [c.key, c.rows.map(x => x.row)]), [['1', [1, 2]], ['2', [3, 4, 5]], ['2', [6]], ['3', [7]]]);
});

test('chunk with break and balance', () => {
  const r = [1, 2, 3, 4].map(i => rec(i, { table: 1 }));
  assert.deepEqual(LB.chunkRecords(r, { rowsPerSlide: 3, breakField: 'table', balance: true }).map(c => c.rows.length), [2, 2]);
});

test('chunk with a break layout holding fewer rows', () => {
  const r = [1, 2, 3, 4, 5, 6, 7].map(i => rec(i, { table: 1 }));
  assert.deepEqual(LB.chunkRecords(r, { rowsPerSlide: 3, breakField: 'table' }, 2).map(c => c.rows.length), [2, 3, 2]);
});

test('default sort is surname then first name', () => {
  const r = [rec(1, { first_name: 'Zoe', last_name: 'Smith' }), rec(2, { first_name: 'Adam', last_name: 'Smith' }), rec(3, { first_name: 'B', last_name: 'Adams' })];
  assert.deepEqual(LB.sortFlow(r, {}).map(x => x.row), [3, 2, 1]);
});

test('plan fills rows, range and page', () => {
  const r = ['Adams', 'Bell', 'Cox', 'Day', 'Fox'].map((n, i) => ({ row: i + 1, first_name: 'F' + (i + 1), last_name: n, table: i + 1, name: 'F' + (i + 1) + ' ' + n }));
  const plan = LB.planFlow(r, { listFields: ['first_name', 'last_name', 'table'], rowsPerSlide: 3, rangeField: 'last_name' });
  assert.equal(plan.slides.length, 2);
  const [a, b] = plan.slides;
  assert.equal(a.values['row.1.last_name'], 'Adams');
  assert.equal(a.values['row.3.last_name'], 'Cox');
  assert.equal('row.4.last_name' in a.values, false);
  assert.deepEqual([a.values['range.start'], a.values['range.end']], ['Adams', 'Cox']);
  assert.deepEqual([b.values['range.start'], b.values['range.end']], ['Day', 'Fox']);
  assert.deepEqual([a.values['page.number'], a.values['page.count']], ['1', '2']);
  assert.deepEqual([b.values['page.number'], b.values['page.count']], ['2', '2']);
  assert.equal(b.values['row.2.last_name'], 'Fox');
  assert.equal('row.3.last_name' in b.values, false);
});

test('plan with a break uses the break layout and a header', () => {
  const r = [rec(1, { last_name: 'A', table: 1 }), rec(2, { last_name: 'B', table: 1 }), rec(3, { last_name: 'C', table: 1 }), rec(4, { last_name: 'D', table: 1 }), rec(5, { last_name: 'E', table: 2 })];
  const plan = LB.planFlow(r, { listFields: ['last_name', 'table'], rowsPerSlide: 3, sort: ['table', 'last_name'], breakField: 'table' });
  assert.deepEqual(plan.slides.map(s => [s.layout, s.values.header]), [['LIST BREAK', 'Table 1'], ['LIST', 'Table 1'], ['LIST BREAK', 'Table 2']]);
  assert.equal(plan.slides[1].values['row.1.last_name'], 'D');
});

test('plan with balance', () => {
  const r = Array.from({ length: 63 }, (_, i) => rec(i + 1, { last_name: 'n' + String(i + 1).padStart(3, '0') }));
  const plan = LB.planFlow(r, { listFields: ['last_name'], rowsPerSlide: 30, balance: true });
  assert.deepEqual(plan.slides.map(s => s.rows.length), [21, 21, 21]);
});

test('a skipped row is reported on the slide that would have held it', () => {
  const r = ['Adams', 'Bell', 'Cox'].map((n, i) => rec(i + 1, { last_name: n }));
  const skipped = [rec(9, { last_name: 'Bird' })];
  const plan = LB.planFlow(r, { listFields: ['last_name'], rowsPerSlide: 2 }, skipped);
  assert.equal(plan.slides.length, 2);
  assert.match(plan.slides[1].warnings[0], /Row 9/);
});

/* ---------- validate.py, the fixes and warnings ---------- */

test('validation fixes and warnings on dirty rows', () => {
  const t = LB.buildTable([['First Name', 'Last Name', 'Table'],
    ['Hannah Fischer', null, 6], ['DAVID', 'GRIFFITHS', 5], ['Dr Amara', 'Okafor', 11], ['Charles', 'Hudson Jr', 17],
    ['Walker', 'Sophie', 12], ['Rosie', 'Bell', 'TBC'], ['Nadia', 'Reynolds', null], [null, null, 9],
    ['Joshua', 'Bennett', 3], ['Joshua', 'Bennett', 7], ['Mei', null, 21], ['Omar', 'Ali', 112],
    ['Ann', 'Lee', 1], ['Bob', 'Day', 2], ['Cy', 'Fox', 3], ['Di', 'Ray', 4], ['Ed', 'Kim', 5], ['Fay', 'Ng', 6]]);
  const recs = LB.buildRecords(t, { first: 'First Name', last: 'Last Name', table: 'Table' });
  const v = LB.validate(recs, { tableColumn: true });
  const by = {};
  v.records.forEach(r => { by[r.row] = r; });
  assert.deepEqual([by[2].first_name, by[2].last_name], ['Hannah', 'Fischer']);
  assert.deepEqual([by[3].first_name, by[3].last_name], ['David', 'Griffiths']);
  assert.equal(by[4].first_name, 'Amara');
  assert.equal(by[5].last_name, 'Hudson');
  assert.equal(by[6].first_name, 'Walker');            /* swapped names are reported, not changed */
  assert.deepEqual(v.skipped.map(r => r.row), [9]);
  const rules = v.warnings.map(w => w.rule);
  ['full_name_in_one_field', 'name_case', 'honorific_in_name', 'swapped_names', 'non_numeric_table', 'blank_table', 'blank_name',
    'duplicate_person', 'blank_last_name', 'table_out_of_range'].forEach(r => assert.ok(rules.includes(r), 'expected ' + r));
});

/* ---------- layout and the pptx package ---------- */

test('geometry derives columns from names per slide', () => {
  assert.equal(LB.layoutGeometry({ perSlide: 30, fields: ['name', 'table'] }).cols, 2);
  assert.equal(LB.layoutGeometry({ perSlide: 12, fields: ['name'] }).cols, 1);
  assert.equal(LB.layoutGeometry({ perSlide: 45, fields: ['name'] }).cols, 3);
  assert.equal(LB.layoutGeometry({ perSlide: 100, fields: ['name'] }).cols, 4);
  const g = LB.layoutGeometry({ perSlide: 30, fields: ['first_name', 'last_name', 'table'], title: true, range: true, page: true });
  assert.equal(g.rows.length, 30);
  assert.ok(g.rows[0].first_name && g.rows[0].last_name && g.rows[0].table);
  assert.ok(g.extras.title && g.extras['range.start'] && g.extras['page.count']);
});

test('placeholder list names and idx are stable across LIST and LIST BREAK', () => {
  const cfg = { perSlide: 3, fields: ['name', 'table'], title: true, header: true, range: true, page: true };
  const g = LB.layoutGeometry(cfg);
  const a = LB.placeholderList(g, cfg, false), b = LB.placeholderList(g, cfg, true);
  assert.deepEqual(a.map(p => p.name), ['title', 'range.start', 'range.end', 'page.number', 'page.count', 'row.1.name', 'row.1.table', 'row.2.name', 'row.2.table', 'row.3.name', 'row.3.table']);
  assert.equal(b[1].name, 'header');
  const idxA = {}, idxB = {};
  a.forEach(p => { idxA[p.name] = p.idx; }); b.forEach(p => { idxB[p.name] = p.idx; });
  Object.keys(idxA).forEach(n => assert.equal(idxA[n], idxB[n], n));
});

test('output name comes from the source name plus a timestamp', () => {
  assert.equal(LB.outputName('Guests 2026.xlsx', new Date(2026, 10, 4, 9, 5)), 'Guests 2026_lists_20261104_0905.pptx');
});

test('the pptx holds inherited placeholders and the manifest', { skip: !JSZip && 'jszip not installed' }, async () => {
  const r = ['Adams', 'Bell', 'Cox', 'Day', 'Fox'].map((n, i) => ({ row: i + 1, first_name: 'F' + (i + 1), last_name: n, table: i + 1, name: 'F' + (i + 1) + ' ' + n }));
  const plan = LB.planFlow(r, { listFields: ['name', 'table'], rowsPerSlide: 3, rangeField: 'last_name' });
  const zip = LB.buildDeck(plan, { perSlide: 3, fields: ['name', 'table'], title: 'Guests', range: true, page: true, breakOn: false, created: new Date('2026-11-04T12:00:00Z') }, JSZip);
  const files = Object.keys(zip.files);
  assert.ok(files.includes('ppt/slideLayouts/slideLayout1.xml'));
  assert.ok(!files.includes('ppt/slideLayouts/slideLayout2.xml'));
  assert.ok(files.includes('ppt/slides/slide2.xml'));
  assert.ok(!files.includes('ppt/slides/slide3.xml'));
  const lay = await zip.file('ppt/slideLayouts/slideLayout1.xml').async('string');
  const sl = await zip.file('ppt/slides/slide1.xml').async('string');
  const master = await zip.file('ppt/slideMasters/slideMaster1.xml').async('string');
  assert.match(lay, /<p:cSld name="LIST">/);
  assert.equal((lay.match(/<p:ph type="body" idx="\d+"/g) || []).length, 10);
  assert.match(lay, /name="row\.1\.name"/);
  assert.equal((sl.match(/<p:ph type="body" idx="\d+"/g) || []).length, 10);
  assert.equal((sl.match(/<a:xfrm/g) || []).length, 1, 'only the group xfrm on the slide');
  assert.equal((sl.match(/ sz="/g) || []).length, 0, 'no run sizes on the slide');
  assert.equal((sl.match(/<p:spPr\/>/g) || []).length, 11);
  assert.match(sl, /<p:ph type="title"\/>/);
  assert.match(master, /name="NAMEDROP\.MANIFEST" hidden="1"/);
  assert.match(master, /&quot;engine&quot;: &quot;flow&quot;/);
  /* every slide idx exists on the layout with the same name */
  const layIdx = {};
  lay.replace(/name="([^"]+)"\/><p:cNvSpPr><a:spLocks noGrp="1"\/><\/p:cNvSpPr><p:nvPr><p:ph type="body" idx="(\d+)"/g, (m, n, i) => { layIdx[i] = n; });
  sl.replace(/name="([^"]+)"\/><p:cNvSpPr><a:spLocks noGrp="1"\/><\/p:cNvSpPr><p:nvPr><p:ph type="body" idx="(\d+)"/g, (m, n, i) => { assert.equal(layIdx[i], n); });
  const sl2 = await zip.file('ppt/slides/slide2.xml').async('string');
  assert.ok(sl2.includes('row.2.name') && !sl2.includes('row.3.name'), 'unused rows are not written');
});

}));
