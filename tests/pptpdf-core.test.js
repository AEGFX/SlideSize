/* PowerPoint to PDF tests. No dependencies.
   Run from the repository root with:  node --test

   This file covers the parts with no disk, browser or PowerPoint in them:
   settings and what each platform can honour, page geometry, output names,
   the deck reader, the PDF reader and editor, issues and result states,
   resuming, reports and the picture comparison.                          */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'), path = require('path');
const C = require('../pptpdf-core.js');
const I = require('../pptpdf-inspect.js');
const P = require('../pptpdf-pdf.js');

const DECKS = path.join(__dirname, 'pptpdf', 'decks');
const PDFS = path.join(__dirname, 'pptpdf', 'pdfs');
const deck = name => new Blob([fs.readFileSync(path.join(DECKS, name))]);
const pdf = name => new Uint8Array(fs.readFileSync(path.join(PDFS, name)));
const near = (a, b, tol, msg) => assert.ok(Math.abs(a - b) <= (tol || 0.01), (msg || '') + ' expected ' + b + ' got ' + a);
const codes = list => list.map(i => i.code);

/* ---------- platform and settings ---------- */

test('platform is read from the browser, and phones are not treated as computers', () => {
  assert.equal(C.detectPlatform({ userAgentData: { platform: 'Windows' } }), 'windows');
  assert.equal(C.detectPlatform({ userAgentData: { platform: 'macOS' } }), 'macos');
  assert.equal(C.detectPlatform({ platform: 'MacIntel', userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) Chrome/140' }), 'macos');
  assert.equal(C.detectPlatform({ platform: 'Linux x86_64', userAgent: 'X11; Linux' }), 'other');
  assert.equal(C.detectPlatform({ platform: 'iPhone', userAgent: 'iPhone' }), 'other');
  assert.equal(C.detectPlatform({ platform: 'MacIntel', maxTouchPoints: 5, userAgent: 'Mozilla/5.0 (Macintosh) Safari' }), 'other', 'an iPad that says it is a Mac');
  assert.equal(C.detectPlatform({ userAgentData: { platform: 'Android', mobile: true } }), 'other');
});

test('a browser without folder access is reported, not worked around', () => {
  assert.equal(C.browserSupport({ showDirectoryPicker() {}, DecompressionStream() {}, isSecureContext: true }).ok, true);
  const s = C.browserSupport({ DecompressionStream() {}, isSecureContext: true });
  assert.equal(s.ok, false);
  assert.match(s.message, /Chrome or Edge/);
});

test('defaults follow the brief', () => {
  const s = C.defaultSettings();
  assert.equal(s.size.preset, 'original');
  assert.equal(s.pdfa, false);
  assert.equal(s.hidden, false, 'hidden slides are left out by default');
  assert.equal(s.range, null);
  assert.equal(s.output, 'slides');
  assert.equal(s.quality, 'standard', 'high image quality by default');
  assert.equal(s.allowRemoteLinks, false, 'linked files are not fetched without being asked');
});

test('Windows offers every setting', () => {
  const asked = { pdfa: true, output: 'notes', quality: 'minimum', bitmapText: false, tags: false, docProps: false, markup: true, hidden: true, range: { from: 2, to: 9 } };
  const r = C.resolveSettings(asked, 'windows');
  assert.deepEqual(r.disabled, {});
  assert.equal(r.effective.pdfa, true);
  assert.equal(r.effective.output, 'notes');
  assert.equal(r.effective.quality, 'minimum');
  assert.deepEqual(r.effective.range, { from: 2, to: 9 });
});

test('macOS switches off what PowerPoint for Mac cannot do, each with a reason', () => {
  const asked = { pdfa: true, output: 'notes', quality: 'minimum', bitmapText: false, tags: false, docProps: false, markup: true, hidden: true, range: { from: 2, to: 9 } };
  const r = C.resolveSettings(asked, 'macos');
  for (const k of ['pdfa', 'notes', 'quality', 'bitmapText', 'tags', 'docProps', 'markup']) {
    assert.ok(r.disabled[k] && r.disabled[k].length > 20, k + ' has a reason');
  }
  assert.equal(r.effective.pdfa, false);
  assert.equal(r.effective.output, 'slides');
  assert.equal(r.effective.quality, 'standard');
  assert.equal(r.effective.markup, false);
  assert.equal(r.effective.hidden, true, 'hidden slides still work on macOS');
  assert.deepEqual(r.effective.range, { from: 2, to: 9 }, 'so does a slide range');
  assert.ok(r.notes.some(n => /count from the slides that remain/.test(n.text)), 'the renumbering is explained before conversion');
});

test('what the running helper reports replaces the guess', () => {
  const caps = { pdfa: false, notes: true, reasons: { pdfa: 'This PowerPoint is too old for PDF/A.' } };
  const r = C.resolveSettings({ pdfa: true, output: 'notes' }, 'macos', caps);
  assert.equal(r.disabled.pdfa, 'This PowerPoint is too old for PDF/A.');
  assert.equal(r.disabled.notes, undefined, 'the helper said notes pages are available');
  assert.equal(r.effective.output, 'notes');
});

test('nothing converts on a system with no PowerPoint', () => {
  const r = C.resolveSettings({ hidden: true }, 'other');
  assert.match(r.disabled.hidden, /Windows and macOS/);
});

test('settings that do not go together are explained before conversion', () => {
  let r = C.resolveSettings({ pdfa: true, metadata: 'strip' }, 'windows');
  assert.match(r.disabled.metadata, /PDF\/A requires the metadata/);
  assert.equal(r.effective.metadata, 'keep');

  r = C.resolveSettings({ pdfa: true, bitmapText: false, tags: false, size: { preset: 'a4' }, links: 'remove' }, 'windows');
  const text = r.notes.map(n => n.text).join(' | ');
  assert.match(text, /needs every font embedded/);
  assert.match(text, /requires structure tags/);
  assert.match(text, /resizes the pages and removes the links after PowerPoint has written the file/);
  assert.match(text, /part and level it actually wrote are read back/);

  r = C.resolveSettings({ output: 'notes', fidelity: 'all', size: { preset: 'a4' } }, 'windows');
  assert.ok(r.notes.some(n => /not run on notes pages/.test(n.text)));
  assert.ok(r.notes.some(n => /whole notes page/.test(n.text)));
});

test('a slide range and a custom size are validated', () => {
  assert.equal(C.resolveSettings({ range: { from: 5, to: 2 } }, 'windows').errors.length, 1);
  assert.equal(C.resolveSettings({ range: { from: 0, to: 2 } }, 'windows').errors.length, 1);
  assert.equal(C.resolveSettings({ range: { from: 2, to: 2 } }, 'windows').errors.length, 0);
  assert.equal(C.resolveSettings({ size: { preset: 'custom', customW: 10, customH: 500 } }, 'windows').errors.length, 1);
  assert.equal(C.resolveSettings({ size: { preset: 'custom', customW: 20000, customH: 500 } }, 'windows').errors.length, 1);
  assert.equal(C.resolveSettings({ timeoutSec: 1 }, 'windows').effective.timeoutSec, 30);
});

test('one file can have settings of its own without touching the batch', () => {
  const batch = C.defaultSettings({ hidden: false });
  const item = { overrides: { hidden: true, size: { preset: 'a4' } } };
  const s = C.settingsFor(item, batch);
  assert.equal(s.hidden, true);
  assert.equal(s.size.preset, 'a4');
  assert.equal(s.size.fit, 'fit', 'the rest of the size settings are kept');
  assert.equal(batch.hidden, false);
  assert.equal(C.settingsFor({ overrides: null }, batch).hidden, false);
});

test('built in presets resolve to valid settings', () => {
  for (const p of C.BUILTIN_PRESETS) {
    const r = C.resolveSettings(C.defaultSettings(p.settings), 'windows');
    assert.deepEqual(r.errors, [], p.name);
  }
  assert.equal(C.defaultSettings(C.BUILTIN_PRESETS.find(p => p.id === 'archive').settings).pdfa, true);
});

/* ---------- geometry ---------- */

test('the original size is kept unless a page size is chosen', () => {
  const t = C.targetSize(960, 540, { preset: 'original' });
  assert.deepEqual([t.w, t.h, t.changed], [960, 540, false]);
});

test('standard page sizes follow the orientation of the deck', () => {
  let t = C.targetSize(960, 540, { preset: 'a4', orientation: 'auto' });
  near(t.w, 841.89); near(t.h, 595.28); assert.equal(t.orientation, 'landscape');
  t = C.targetSize(595, 842, { preset: 'a4', orientation: 'auto' });
  near(t.w, 595.28); near(t.h, 841.89); assert.equal(t.orientation, 'portrait');
  t = C.targetSize(595, 842, { preset: 'letter', orientation: 'landscape' });
  assert.deepEqual([t.w, t.h], [792, 612]);
  t = C.targetSize(960, 540, { preset: '4:3', orientation: 'auto' });
  assert.deepEqual([t.w, t.h, t.changed], [720, 540, true]);
  t = C.targetSize(960, 540, { preset: '16:9', orientation: 'auto' });
  assert.equal(t.changed, false, '16:9 on a 16:9 deck changes nothing');
  t = C.targetSize(960, 540, { preset: 'custom', orientation: 'auto', customW: 400, customH: 800 });
  assert.deepEqual([t.w, t.h], [400, 800], 'a custom size is used as typed');
});

test('fit adds margins, fill crops, and neither ever stretches', () => {
  const shapes = [[960, 540], [720, 540], [1920, 540], [595, 842], [540, 540]];
  const pages = [[960, 540], [720, 540], [841.89, 595.28], [612, 792], [300, 900]];
  for (const [sw, sh] of shapes) for (const [pw, ph] of pages) for (const mode of ['fit', 'fill']) {
    const t = C.fitTransform(sw, sh, pw, ph, mode);
    /* one scale for both axes is the whole guarantee. The picture sits centred. */
    near(t.tx + t.scale * sw / 2, pw / 2, 1e-6);
    near(t.ty + t.scale * sh / 2, ph / 2, 1e-6);
    if (mode === 'fit') {
      assert.ok(t.scale * sw <= pw + 1e-6 && t.scale * sh <= ph + 1e-6, 'fit stays inside the page');
      assert.equal(t.cropped, false);
    } else {
      assert.ok(t.scale * sw >= pw - 1e-6 && t.scale * sh >= ph - 1e-6, 'fill covers the page');
      assert.equal(t.cropped, !t.exact);
    }
  }
  const crop = C.fitTransform(960, 540, 720, 540, 'fill');
  near(crop.scale, 1); near(crop.cropX, 0.125); assert.equal(crop.cropY, 0);
  near(crop.visibleShare, 0.75);
  const fit = C.fitTransform(960, 540, 720, 540, 'fit');
  near(fit.scale, 0.75); near(fit.marginY, 67.5); assert.equal(fit.marginX, 0);
});

test('the compared windows line up the same content in both pictures', () => {
  let t = C.fitTransform(960, 540, 720, 540, 'fit');
  let w = C.compareWindows(960, 540, 720, 540, t);
  assert.deepEqual(w.ref, { x: 0, y: 0, w: 1, h: 1 }, 'with fit the whole slide is visible');
  near(w.pdf.y, 0.125); near(w.pdf.h, 0.75); near(w.pdf.x, 0); near(w.pdf.w, 1);
  t = C.fitTransform(960, 540, 720, 540, 'fill');
  w = C.compareWindows(960, 540, 720, 540, t);
  near(w.ref.x, 0.125); near(w.ref.w, 0.75); near(w.ref.y, 0); near(w.ref.h, 1);
  near(w.pdf.x, 0); near(w.pdf.w, 1); near(w.pdf.h, 1);
  assert.deepEqual(C.compareWindows(1, 1, 1, 1, null).pdf, { x: 0, y: 0, w: 1, h: 1 });
});

test('expected pages come from the slides, the hidden slides and the range', () => {
  const facts = { slideCount: 8, hiddenSlides: [3, 7], widthPt: 960, heightPt: 540 };
  assert.deepEqual(C.expectedSlides(facts, C.defaultSettings()), [1, 2, 4, 5, 6, 8]);
  assert.deepEqual(C.expectedSlides(facts, C.defaultSettings({ hidden: true })), [1, 2, 3, 4, 5, 6, 7, 8]);
  assert.deepEqual(C.expectedSlides(facts, C.defaultSettings({ range: { from: 2, to: 5 } })), [2, 4, 5]);
  assert.deepEqual(C.expectedSlides(facts, C.defaultSettings({ range: { from: 6, to: 99 }, hidden: true })), [6, 7, 8]);
  assert.equal(C.expectedSlides({ slideCount: 3, hiddenSlides: null }, C.defaultSettings()), null, 'a .ppt does not say which slides are hidden until PowerPoint opens it');
  assert.deepEqual(C.expectedSlides({ slideCount: 3, hiddenSlides: null }, C.defaultSettings({ hidden: true })), [1, 2, 3]);
});

test('the plan shown before conversion has the size, orientation and page count', () => {
  const facts = { slideCount: 8, hiddenSlides: [3, 7], widthPt: 960, heightPt: 540 };
  let p = C.planOutput(facts, C.defaultSettings());
  assert.deepEqual([p.pages, p.w, p.h, p.orientation, p.changed], [6, 960, 540, 'landscape', false]);
  p = C.planOutput(facts, C.defaultSettings({ size: { preset: '4:3', fit: 'fill' } }));
  assert.equal(p.transform.cropped, true);
  near(p.transform.cropX, 0.125);
  p = C.planOutput(facts, C.defaultSettings({ output: 'notes' }));
  assert.equal(p.known, false, 'a notes page has its own size, known only once made');
  assert.equal(p.pages, 6);
});

/* ---------- names ---------- */

test('file names are made safe for both systems', () => {
  assert.equal(C.safeBase('Opening Keynote v3.pptx'), 'Opening Keynote v3');
  assert.equal(C.safeBase('a/b:c*?.pptx'), 'a_b_c_');
  assert.equal(C.safeBase('CON.ppt'), '_CON');
  assert.equal(C.safeBase('trailing dots....pptx'), 'trailing dots');
  assert.equal(C.safeBase('.pptx'), 'presentation');
  assert.equal(C.safeBase('v1.2 final.pptx'), 'v1.2 final');
});

test('duplicate names inside a batch are resolved in a way that says why', () => {
  const items = [
    { id: 'a', name: 'Deck.pptx', folder: 'day1' }, { id: 'b', name: 'Deck.pptx', folder: 'show/day2' },
    { id: 'c', name: 'Deck.ppt', folder: '' }, { id: 'd', name: 'deck.PPTX', folder: 'day1' }, { id: 'e', name: 'Other.pptx' }];
  const p = C.planNames(items, [], {});
  assert.equal(p.a.outName, 'Deck.pdf');
  assert.equal(p.b.outName, 'Deck (day2).pdf', 'the folder tells them apart');
  assert.equal(p.c.outName, 'Deck (ppt).pdf', 'or the file type');
  assert.equal(p.d.outName, 'deck (day1).pdf');
  assert.equal(p.e.outName, 'Other.pdf');
  assert.match(p.b.reason, /Another file in this batch is also called Deck/);
  const names = Object.values(p).map(x => x.outName.toLowerCase());
  assert.equal(new Set(names).size, names.length, 'no two outputs share a name, whatever the case');
});

test('an existing PDF is never overwritten without a decision', () => {
  const items = [{ id: 'a', name: 'Deck.pptx' }, { id: 'b', name: 'New.pptx' }];
  let p = C.planNames(items, ['deck.pdf', 'Unrelated.pdf'], {});
  assert.equal(p.a.exists, true); assert.equal(p.a.blocked, true, 'conversion waits for an answer');
  assert.equal(p.b.exists, false); assert.equal(p.b.blocked, false);
  p = C.planNames(items, ['deck.pdf'], { a: 'overwrite' });
  assert.deepEqual([p.a.outName, p.a.blocked, p.a.decision], ['Deck.pdf', false, 'overwrite']);
  p = C.planNames(items, ['deck.pdf', 'Deck (2).pdf'], { a: 'rename' });
  assert.equal(p.a.outName, 'Deck (3).pdf', 'keep both finds a free name');
  assert.equal(p.a.blocked, false);
  p = C.planNames(items, ['deck.pdf'], { a: 'skip' });
  assert.equal(p.a.decision, 'skip');
});

/* ---------- reading presentations ---------- */

test('reads a .pptx without loading its media', async () => {
  const r = await I.inspect(deck('Kitchen sink 16x9.pptx'), 'Kitchen sink 16x9.pptx');
  assert.equal(r.ok, true);
  assert.equal(r.kind, 'pptx');
  assert.equal(r.slideCount, 7);
  near(r.widthPt, 960, 0.1); assert.equal(r.heightPt, 540);
  assert.deepEqual(r.hiddenSlides, [3]);
  assert.ok(r.fonts.includes('Gotham Light') && r.fonts.includes('Montserrat') && r.fonts.includes('Carlito'));
  assert.ok(r.slideFonts.includes('Gotham Light'));
  assert.ok(!r.slideFonts.includes('Arial'), 'a font only on the master is not counted as used on a slide');
  assert.deepEqual(r.counts, { video: 2, audio: 0, notes: 1, transitions: 1, animated: 1, ole: 0, controls: 0, model3d: 0, linkedMedia: 0 });
  assert.equal(r.slides[1].notesWords, 15);
  assert.equal(r.slides[1].transition.type, 'fade');
  assert.equal(r.slides[1].hyperlinks, 1);
  assert.equal(r.slides[0].slideNumberField, true);
  assert.equal(r.slides[4].slideNumberField, false);
});

test('finds video, its poster frame and where it sits on the slide', async () => {
  const r = await I.inspect(deck('Kitchen sink 16x9.pptx'), 'k.pptx');
  const v = r.slides[3].media[0];
  assert.equal(v.type, 'video'); assert.equal(v.linked, false);
  assert.match(v.target, /^ppt\/media\/.*\.mp4$/);
  assert.match(v.posterPart, /^ppt\/media\/.*\.png$/);
  assert.deepEqual(v.box, { left: 180, top: 122.4, width: 576, height: 324 });
  const poster = await r.readBytes(v.posterPart);
  assert.deepEqual([...poster.subarray(1, 4)], [0x50, 0x4e, 0x47], 'the poster frame can be read on its own');
  assert.notEqual(r.slides[4].media[0].posterPart, v.posterPart, 'the second video has a different poster');
});

test('finds animations that stack content, and links to the internet', async () => {
  const r = await I.inspect(deck('Kitchen sink 16x9.pptx'), 'k.pptx');
  assert.deepEqual(r.slides[5].anim, { entrance: 1, exit: 1, emphasis: 0, path: 0, total: 2, overlapping: 1, exitTargets: 1 });
  assert.deepEqual(r.externalLinks, [{ slide: 7, kind: 'image', target: 'https://example.invalid/logo.png', remote: true, hasEmbeddedCopy: true }]);
  assert.equal(I.isRemote('https://x/y.png'), true);
  assert.equal(I.isRemote('\\\\server\\share\\a.png'), true);
  assert.equal(I.isRemote('file://server/share/a.png'), true);
  assert.equal(I.isRemote('file:///C:/a.png'), false);
  assert.equal(I.isRemote('../media/a.png'), false);
});

test('reads mixed slide sizes', async () => {
  const sizes = {};
  for (const n of ['Classic 4x3.pptx', 'Poster A4 portrait.pptx', 'Wide blend 32x9.pptx']) {
    const r = await I.inspect(deck(n), n);
    sizes[n] = [Math.round(r.widthPt), Math.round(r.heightPt), r.slideCount, r.hiddenSlides.length];
  }
  assert.deepEqual(sizes, { 'Classic 4x3.pptx': [720, 540, 3, 0], 'Poster A4 portrait.pptx': [595, 842, 2, 0], 'Wide blend 32x9.pptx': [1920, 540, 2, 1] });
});

test('reads what it can from a legacy .ppt and says what it cannot', async () => {
  const r = await I.inspect(deck('Legacy 97-2003.ppt'), 'Legacy 97-2003.ppt');
  assert.equal(r.ok, true); assert.equal(r.kind, 'ppt');
  assert.deepEqual([r.slideCount, r.widthPt, r.heightPt], [3, 720, 540]);
  assert.equal(r.hiddenSlides, null, 'hidden slides are not known before conversion');
  const issues = C.preflightIssues({ size: 1 }, r, C.defaultSettings(), 'windows');
  assert.deepEqual(codes(issues), ['legacy-format']);
  assert.equal(issues[0].certainty, 'unverified');
});

test('corrupt, password protected, empty and wrong files are told apart', async () => {
  const expect = { 'Password protected.pptx': 'password', 'Cut short.pptx': 'not-zip', 'Empty.pptx': 'empty', 'Not a deck.pptx': 'unsupported', 'A zip, not a deck.pptx': 'not-pptx' };
  for (const n of Object.keys(expect)) {
    const r = await I.inspect(deck(n), n);
    assert.equal(r.ok, false, n);
    assert.equal(r.error.code, expect[n], n);
    const issues = C.preflightIssues({ size: 1 }, r, C.defaultSettings(), 'windows');
    assert.equal(issues.length, 1); assert.equal(issues[0].severity, 'error');
    assert.ok(issues[0].description && issues[0].impact && issues[0].fix, n + ' has a description, an impact and a fix');
  }
  const pw = C.preflightIssues({ size: 1 }, await I.inspect(deck('Password protected.pptx'), 'p.pptx'), C.defaultSettings(), 'windows');
  assert.equal(pw[0].code, 'password');
  const cut = C.preflightIssues({ size: 1 }, await I.inspect(deck('Cut short.pptx'), 'c.pptx'), C.defaultSettings(), 'windows');
  assert.equal(cut[0].code, 'corrupt');
});

test('a wrong extension is noticed', async () => {
  const r = await I.inspect(deck('Legacy 97-2003.ppt'), 'renamed.pptx');
  assert.equal(r.ok, true); assert.equal(r.extMismatch, true);
  assert.ok(codes(C.preflightIssues({ size: 1 }, r, C.defaultSettings(), 'windows')).includes('ext-mismatch'));
});

test('only the parts that are needed are read from the file', async () => {
  const buf = fs.readFileSync(path.join(DECKS, 'Kitchen sink 16x9.pptx'));
  let read = 0;
  const spy = { size: buf.length, slice(a, b) { read += Math.max(0, Math.min(b, buf.length) - a); return new Blob([buf.subarray(a, b)]); } };
  const zip = await I.openZip(spy);
  assert.ok(zip.names.includes('ppt/presentation.xml'));
  const before = read;
  await zip.text('ppt/presentation.xml');
  assert.ok(read - before < 8192, 'one small XML part costs a few kilobytes');
  const media = zip.names.filter(n => /\.mp4$/.test(n))[0];
  assert.ok(zip.entry(media).usize > 0);
});

/* ---------- issues before conversion ---------- */

async function kitchenIssues(settings, platform) {
  const ins = await I.inspect(deck('Kitchen sink 16x9.pptx'), 'k.pptx');
  ins.slides[4].media[0].posterBlank = true;                  /* the page finds this by looking at the picture */
  const slim = C.slimInspect(ins);
  return { slim, issues: C.preflightIssues({ size: 84000, name: 'k.pptx' }, slim, C.defaultSettings(settings), platform || 'windows') };
}

test('every issue names its slide, how sure it is, the impact and a fix', async () => {
  const { issues } = await kitchenIssues({ allowRemoteLinks: true });
  const by = Object.fromEntries(issues.map(i => [i.code, i]));
  assert.deepEqual(codes(issues).sort(), ['animation-overlap', 'hidden-slides', 'notes-present', 'remote-links', 'transitions', 'video', 'video-no-poster'].sort());
  assert.equal(by.video.slide, 4); assert.equal(by.video.certainty, 'confirmed');
  assert.match(by.video.description, /poster frame at the same position and size/);
  assert.match(by.video.impact, /does not play/);
  assert.equal(by['video-no-poster'].slide, 5); assert.equal(by['video-no-poster'].certainty, 'suspected'); assert.equal(by['video-no-poster'].review, true);
  assert.match(by['video-no-poster'].impact, /clearly marked placeholder/);
  assert.equal(by['animation-overlap'].slide, 6); assert.equal(by['animation-overlap'].certainty, 'suspected'); assert.equal(by['animation-overlap'].review, true);
  assert.match(by['animation-overlap'].impact, /every shape at once/);
  assert.match(by['hidden-slides'].description, /1 hidden slide \(3\)/);
  assert.match(by['hidden-slides'].impact, /left out/);
  for (const i of issues) {
    assert.ok(['error', 'warning', 'info'].includes(i.severity) && ['confirmed', 'suspected', 'unverified'].includes(i.certainty), i.code);
    assert.ok(i.id, i.code + ' has an id');
    if (i.severity !== 'info') assert.ok(i.impact && i.fix, i.code + ' has an impact and a fix');
  }
});

test('a deck that links to the internet is held back until it is allowed', async () => {
  let r = await kitchenIssues({});
  let link = r.issues.find(i => i.code === 'remote-links');
  assert.equal(link.severity, 'error');
  assert.match(link.impact, /held back so that nothing is fetched/);
  r = await kitchenIssues({ allowRemoteLinks: true });
  link = r.issues.find(i => i.code === 'remote-links');
  assert.equal(link.severity, 'warning');
  assert.match(link.impact, /will fetch it/);
});

test('issues follow the settings', async () => {
  let r = await kitchenIssues({ hidden: true, output: 'notes', placeholders: false, allowRemoteLinks: true });
  assert.match(r.issues.find(i => i.code === 'hidden-slides').impact, /included in the PDF/);
  assert.ok(!codes(r.issues).includes('notes-present'), 'no note about missing notes when notes pages are on');
  assert.match(r.issues.find(i => i.code === 'video-no-poster').impact, /Placeholders are turned off/);
  r = await kitchenIssues({ allowRemoteLinks: true }, 'macos');
  assert.equal(r.issues.find(i => i.code === 'notes-present').fix, '', 'no advice to use a setting this platform does not have');
});

test('placeholders are asked for only where a video has no usable poster frame', async () => {
  const { slim } = await kitchenIssues({});
  const ph = C.placeholdersFor(slim, C.defaultSettings());
  assert.equal(ph.length, 1);
  assert.deepEqual([ph[0].slide, ph[0].left, ph[0].top, ph[0].width, ph[0].height], [5, 180, 122.4, 576, 324]);
  assert.match(ph[0].label, /placeholder added by SlideSize/);
  assert.deepEqual(C.placeholdersFor(slim, C.defaultSettings({ placeholders: false })), []);
});

test('a flat poster frame is recognised and a real one is not', () => {
  const w = 32, h = 18, flat = new Uint8ClampedArray(w * h * 4), busy = new Uint8ClampedArray(w * h * 4);
  for (let i = 0; i < w * h; i++) { flat.set([2, 2, 3, 255], i * 4); busy.set([(i * 7) % 255, (i * 13) % 255, 90, 255], i * 4); }
  assert.deepEqual([C.looksBlank(flat, w, h).blank, C.looksBlank(flat, w, h).colour], [true, 'black']);
  assert.equal(C.looksBlank(busy, w, h).blank, false);
});

/* ---------- the job ticket ---------- */

test('the ticket carries the settings to the helper', () => {
  const item = { id: 'f00007', name: 'Deck.pptx', ext: 'pptx', outName: 'Deck.pdf', overwrite: false, inspect: { ok: true, kind: 'pptx', slides: [] } };
  let t = C.buildTicket(item, C.defaultSettings());
  assert.equal(t.protocol, C.PROTOCOL);
  assert.equal(t.input, 'in/slidesize-f00007.pptx', 'the copy has a name of its own, so nothing of the user\'s can be mistaken for it');
  assert.equal(t.pdf, '../Deck.pdf');
  assert.equal(t.overwrite, false);
  assert.deepEqual(t.options, { intent: 'print', includeHidden: false, range: null, output: 'slides', bitmapText: true, tags: true, docProps: true, markup: false, pdfa: false,
    reference: { mode: 'all', longEdge: 1280, slides: null }, placeholders: [], validatePdfa: false });
  t = C.buildTicket(item, C.defaultSettings({ quality: 'minimum', hidden: true, range: { from: 2, to: 4 }, output: 'notes', pdfa: true, fidelity: 'all' }));
  assert.equal(t.options.intent, 'screen');
  assert.deepEqual(t.options.range, [2, 4]);
  assert.equal(t.options.reference, null, 'no slide pictures for notes pages');
  assert.equal(t.validateNow, true);
});

test('a PDF the page still has to edit goes to the work folder first', () => {
  const item = { id: 'f00001', name: 'Deck.pptx', ext: 'pptx', outName: 'Deck.pdf', overwrite: false, inspect: { ok: true, kind: 'pptx' } };
  for (const s of [{ size: { preset: 'a4' } }, { links: 'remove' }, { metadata: 'strip' }]) {
    const t = C.buildTicket(item, C.defaultSettings(s));
    assert.equal(t.pdf, 'out/f00001.pdf');
    assert.equal(t.overwrite, true);
  }
  assert.equal(C.buildTicket(item, C.defaultSettings({ pdfa: true, size: { preset: 'a4' } })).validateNow, false, 'PDF/A is validated after the edit, not before');
});

test('a sample is five slides spread through the deck', () => {
  assert.deepEqual(C.sampleSlides([1, 2, 3]), [1, 2, 3]);
  assert.deepEqual(C.sampleSlides(Array.from({ length: 21 }, (_, i) => i + 1)), [1, 6, 11, 16, 21]);
  assert.deepEqual(C.sampleSlides([2, 4, 5, 6, 8, 9, 12]), [2, 5, 6, 9, 12]);
});

/* ---------- the helper as the page sees it ---------- */

test('the helper state is described in plain words', () => {
  const now = Date.parse('2026-10-07T12:00:10Z');
  const h = { protocol: 1, heartbeat: '2026-10-07T12:00:08Z', state: 'idle', platform: 'windows', powerpoint: { installed: true, version: '16.0', userPresentations: 2 } };
  assert.equal(C.helperStatus(null, now).state, 'absent');
  assert.equal(C.helperStatus(h, now).state, 'ready');
  assert.match(C.helperStatus(h, now).text, /2 presentations of yours are open and will be left alone/);
  assert.equal(C.helperStatus(Object.assign({}, h, { heartbeat: '2026-10-07T11:59:00Z' }), now).state, 'stale');
  assert.equal(C.helperStatus(Object.assign({}, h, { protocol: 2 }), now).state, 'mismatch');
  assert.equal(C.helperStatus(Object.assign({}, h, { state: 'stopped' }), now).state, 'stopped');
  assert.equal(C.helperStatus(Object.assign({}, h, { state: 'blocked', message: 'PowerPoint is not responding.' }), now).text, 'PowerPoint is not responding.');
  const none = C.helperStatus(Object.assign({}, h, { powerpoint: { installed: false } }), now);
  assert.equal(none.state, 'no-powerpoint');
  assert.match(none.text, /Microsoft PowerPoint must be installed for PowerPoint-based conversion\./);
});

/* ---------- fonts ---------- */

test('a missing font and an embedding restriction are different findings', () => {
  const deckFonts = ['Gotham Light', 'Locked Sans', 'Carlito', 'Open Sans'];
  const installed = ['Carlito', 'Locked Sans', 'Open Sans', 'Arial'];
  const pptFonts = [{ name: 'Gotham Light', embeddable: true }, { name: 'Locked Sans', embeddable: false }, { name: 'Carlito', embeddable: true }];
  const pdfFonts = [{ name: 'ABCDEF+Carlito-Bold', embedded: true }, { name: 'BCDEFG+ArialMT', embedded: true }, { name: 'Open Sans', embedded: false }];
  const f = C.fontFindings(deckFonts, installed, [], pptFonts, pdfFonts, C.defaultSettings(), 'windows', deckFonts);
  assert.deepEqual(f.missing, ['Gotham Light']);
  assert.deepEqual(f.restricted, ['Locked Sans']);
  assert.deepEqual(f.notEmbedded, ['Open Sans']);
  assert.deepEqual(f.substitutes, ['ArialMT'], 'a font in the PDF that the deck never names is the likely substitute');
  const by = Object.fromEntries(f.issues.map(i => [i.code, i]));
  assert.equal(by['font-missing'].certainty, 'confirmed');
  assert.match(by['font-missing'].impact, /A bitmap of the wrong font is still the wrong font/, 'rasterising is not offered as a cure for a missing font');
  assert.match(by['font-missing'].impact, /likely substitutes, are ArialMT/);
  assert.match(by['font-restricted'].description, /licence does not allow embedding, as reported by PowerPoint/);
  assert.match(by['font-restricted'].impact, /written into the PDF as a picture/);
  assert.match(by['font-restricted'].fix, /licence limit and not a missing font/);
  assert.match(by['font-not-embedded'].impact, /only displays that text correctly on a computer that has the font/);

  const off = C.fontFindings(deckFonts, installed, [], pptFonts, pdfFonts, C.defaultSettings({ bitmapText: false }), 'windows', deckFonts);
  assert.match(off.issues.find(i => i.code === 'font-restricted').impact, /stays as text without the font inside the PDF/);
});

test('font names are matched across the ways PowerPoint and PDFs write them', () => {
  assert.equal(C.fontMatches('Gotham Light', 'AAAAAB+Gotham-Light'), true);
  assert.equal(C.fontMatches('Calibri', 'ABCDEE+Calibri,Bold'), true);
  assert.equal(C.fontMatches('Times New Roman', 'TimesNewRomanPSMT'), true);
  assert.equal(C.fontMatches('Helvetica Neue', 'HelveticaNeue-Light'), true);
  assert.equal(C.fontMatches('Gotham Light', 'Gotham-Book'), false);
  assert.equal(C.fontMatches('Arial', 'Inter-Regular'), false);
});

test('font findings say less when less is known', () => {
  /* no installed list, as before the helper connects: nothing is called missing */
  let f = C.fontFindings(['Gotham Light'], null, [], null, [{ name: 'Inter', embedded: true }], C.defaultSettings(), 'macos', ['Gotham Light']);
  assert.deepEqual(f.missing, []);
  assert.equal(f.issues[0].code, 'font-absent'); assert.equal(f.issues[0].certainty, 'unverified');
  /* no PDF read back: a missing font is only suspected */
  f = C.fontFindings(['Gotham Light'], ['Arial'], [], null, null, C.defaultSettings(), 'windows', null);
  assert.equal(f.issues[0].code, 'font-missing'); assert.equal(f.issues[0].certainty, 'suspected');
  /* a font embedded in the deck is available to PowerPoint on Windows */
  f = C.fontFindings(['Brand Sans'], ['Arial'], ['Brand Sans'], null, [{ name: 'Arial', embedded: true }], C.defaultSettings(), 'windows', ['Brand Sans']);
  assert.deepEqual(f.missing, []);
  /* a font named only on a master and absent from the PDF is not worth a note */
  f = C.fontFindings(['Arial', 'Carlito'], ['Arial', 'Carlito'], [], null, [{ name: 'Carlito', embedded: true }], C.defaultSettings(), 'windows', ['Carlito']);
  assert.deepEqual(f.issues, []);
});

/* ---------- reading and editing PDFs ---------- */

test('reads back what is really in a PDF', async () => {
  const k = await P.inspect(pdf('kitchen.pdf'));
  assert.equal(k.pageCount, 6);
  near(k.pages[0].w, 959.98, 0.01); assert.equal(k.pages[0].h, 540);
  assert.equal(k.uniform, true);
  assert.deepEqual(k.fonts.map(f => [f.name.replace(/^[A-Z]{6}\+/, ''), f.embedded]).sort(), [['Carlito-Regular', true], ['Inter-Regular', true]]);
  assert.equal(k.links, 1);
  assert.equal(k.tagged, true);
  assert.equal(k.claim, null, 'an ordinary PDF makes no PDF/A claim');
  assert.match(k.info.producer, /LibreOffice/);
});

test('a damaged or encrypted PDF is reported, not guessed at', async () => {
  await assert.rejects(P.inspect(new TextEncoder().encode('not a pdf at all')), e => e.code === 'damaged');
  await assert.rejects(P.inspect(pdf('classic.pdf').subarray(0, 300)), e => e.code === 'damaged');
});

test('PDF/A is never called passed on the strength of a setting or a claim', async () => {
  const plain = P.pdfaChecks(await P.inspect(pdf('classic.pdf')));
  assert.equal(plain.result, 'failed', 'a file that does not even claim PDF/A has failed');
  assert.equal(plain.checks.find(c => c.id === 'claim').result, 'fail');

  const a = P.pdfaChecks(await P.inspect(pdf('classic-pdfa.pdf')));
  assert.deepEqual(a.claim, { part: 1, conformance: 'B' }, 'the part and level come from the file, not from the setting');
  assert.equal(a.result, 'not-verified', 'structural checks alone never amount to passed');
  assert.ok(a.checks.every(c => c.result !== 'fail'));
  assert.equal(a.checks.find(c => c.id === 'tagged').result, 'n/a', 'tags are only required at level A');

  let issues = C.pdfaIssues(a);
  assert.equal(issues[0].code, 'pdfa-unverified'); assert.equal(issues[0].certainty, 'unverified'); assert.equal(issues[0].severity, 'warning');
  assert.match(issues[0].description, /declares itself as PDF\/A-1b, but full validation was not run/);
  assert.match(issues[0].impact, /That is not the same as verified/);
  assert.equal(C.pdfaLabel(a, { pdfa: true }), 'Not verified');
  assert.equal(C.pdfaLabel(null, { pdfa: true }), 'Not verified');
  assert.equal(C.pdfaLabel(null, { pdfa: false }), 'Not requested');

  /* only a validator that ran can produce passed */
  const passed = Object.assign({}, a, { result: 'passed', validator: 'veraPDF 1.26.2', profile: 'PDF/A-1B validation profile' });
  issues = C.pdfaIssues(passed);
  assert.equal(issues[0].code, 'pdfa-passed'); assert.equal(issues[0].severity, 'info');
  assert.match(issues[0].description, /veraPDF 1\.26\.2 checked the file/);
  assert.equal(C.pdfaLabel(passed, { pdfa: true }), 'Passed');

  issues = C.pdfaIssues(Object.assign({}, a, { result: 'failed', validator: 'veraPDF 1.26.2', detail: '3 failed rules.' }));
  assert.equal(issues[0].code, 'pdfa-failed'); assert.equal(issues[0].severity, 'error');
  assert.match(issues[0].impact, /must not be described as PDF\/A/);
  assert.match(C.pdfaIssues(plain)[0].description, /no PDF\/A identification/);
});

test('editing a PDF/A file is caught by the checks that follow', async () => {
  const edited = await P.transform(pdf('classic-pdfa.pdf'), { stripMetadata: true });
  const after = P.pdfaChecks(await P.inspect(edited.bytes));
  assert.equal(after.result, 'failed');
  assert.equal(after.claim, null);
});

test('resizing scales and centres, keeps text, fonts, tags and links, and never stretches', async () => {
  const src = pdf('kitchen.pdf'), before = await P.inspect(src);
  for (const [w, h, mode] of [[841.8898, 595.2756, 'fit'], [720, 540, 'fill'], [612, 792, 'fit'], [540, 540, 'fill']]) {
    const r = await P.transform(src, { resize: { w, h, mode, margin: 'none' } });
    const after = await P.inspect(r.bytes);
    assert.equal(after.pageCount, before.pageCount);
    after.pages.forEach(p => { near(p.w, w, 0.01); near(p.h, h, 0.01); assert.equal(p.x, 0); assert.equal(p.y, 0); });
    assert.deepEqual(after.fonts.map(f => f.name).sort(), before.fonts.map(f => f.name).sort(), 'the same fonts, still embedded');
    assert.ok(after.fonts.every(f => f.embedded));
    assert.equal(after.tagged, true, 'the structure tree is untouched');
    assert.equal(after.links, 1);
    const t = C.fitTransform(before.pages[0].w, before.pages[0].h, w, h, mode);
    near(r.pages[0].scale, t.scale, 1e-9); near(r.pages[0].tx, t.tx, 1e-6); near(r.pages[0].ty, t.ty, 1e-6);
    assert.equal(r.resized, before.pageCount);
  }
});

test('the scale applied to a resized page is the same on both axes', async () => {
  const L = require('../pdf-lib.min.js');
  const r = await P.transform(pdf('kitchen.pdf'), { resize: { w: 720, h: 540, mode: 'fit', margin: 'black' } });
  const doc = await L.PDFDocument.load(r.bytes), page = doc.getPages()[0];
  const first = doc.context.lookup(page.node.Contents().get(0));
  const text = new TextDecoder().decode(L.decodePDFRawStream(first).decode());
  const m = /([\d.]+) 0 0 ([\d.]+) ([-\d.]+) ([-\d.]+) cm/.exec(text);
  assert.ok(m, 'the page content is wrapped in one transform: ' + text);
  assert.equal(m[1], m[2], 'horizontal and vertical scale are identical');
  near(+m[1], 720 / 959.9811, 1e-4);
  assert.match(text, /0 0 0 rg/, 'the margin is painted black when asked');
  assert.match(text, /0 0 720 540 re/);
});

test('a link moves with the page content when the page is resized', async () => {
  const L = require('../pdf-lib.min.js');
  const rectOf = async bytes => {
    const doc = await L.PDFDocument.load(bytes);
    for (const p of doc.getPages()) {
      const an = p.node.Annots(); if (!an) continue;
      for (let i = 0; i < an.size(); i++) { const a = doc.context.lookup(an.get(i)); if (String(a.get(L.PDFName.of('Subtype'))) === '/Link') return a.lookup(L.PDFName.of('Rect')).asArray().map(n => n.asNumber()); }
    }
  };
  const src = pdf('kitchen.pdf'), a = await rectOf(src);
  const r = await P.transform(src, { resize: { w: 720, h: 540, mode: 'fit', margin: 'none' } });
  const b = await rectOf(r.bytes), t = r.pages[0];
  near(b[0], a[0] * t.scale + t.tx, 0.01); near(b[1], a[1] * t.scale + t.ty, 0.01);
  near(b[2], a[2] * t.scale + t.tx, 0.01); near(b[3], a[3] * t.scale + t.ty, 0.01);
});

test('removing links and metadata takes them out of the file, not just out of sight', async () => {
  const src = pdf('kitchen.pdf');
  assert.ok(Buffer.from(src).includes('https://slidesize.com/'));
  const r = await P.transform(src, { removeLinks: true, stripMetadata: true });
  assert.equal(r.linksRemoved, 1); assert.equal(r.metadataStripped, true);
  const after = await P.inspect(r.bytes);
  assert.equal(after.links, 0);
  assert.deepEqual(after.info, {});
  assert.ok(!Buffer.from(r.bytes).includes('https://slidesize.com/'), 'the address is gone from the bytes');
  assert.ok(!Buffer.from(r.bytes).includes('LibreOffice'), 'so is the producer');
  assert.equal(after.pageCount, 6);
});

test('a page that is already the right size is left alone', async () => {
  const r = await P.transform(pdf('classic.pdf'), { resize: { w: 720, h: 540, mode: 'fit', margin: 'white' } });
  assert.equal(r.resized, 0);
});

/* ---------- results ---------- */

const PPT = { name: C.ENGINE_NAME, version: '16.0', build: '17328', platform: 'windows' };
function goodResult(over) {
  return Object.assign({ ok: true, status: 'done', engine: PPT, facts: { slides: 7, hidden: [3], slideWidthPt: 960, slideHeightPt: 540, fonts: [] },
    export: { method: 'ExportAsFixedFormat2', pageMap: [1, 2, 4, 5, 6, 7], notApplied: [], placeholders: [], removedSlides: [] }, timing: { totalMs: 8200, exportMs: 5100 } }, over || {});
}
async function kitchenCtx(over) {
  const info = await P.inspect(pdf('kitchen.pdf'));
  const s = C.defaultSettings(over && over.settings);
  const plan = C.planOutput({ slideCount: 7, hiddenSlides: [3], widthPt: 959.98, heightPt: 540 }, s);
  return Object.assign({ settings: s, platform: 'windows', plan, result: goodResult(), pdf: info }, over || {}, { settings: s });
}
const kItem = { name: 'Kitchen.pptx', inspect: { ok: true, kind: 'pptx', slideCount: 7, slides: [{ n: 1, slideNumberField: true }] } };

test('a clean result raises nothing', async () => {
  assert.deepEqual(C.resultIssues(kItem, await kitchenCtx()), []);
});

test('a wrong page count is an error, however successful the conversion looked', async () => {
  const ctx = await kitchenCtx();
  ctx.result = goodResult({ export: { method: 'ExportAsFixedFormat2', pageMap: [1, 2, 3, 4, 5, 6, 7], notApplied: [] } });
  const issues = C.resultIssues(kItem, ctx);
  assert.deepEqual(codes(issues), ['page-count']);
  assert.equal(issues[0].severity, 'error');
  assert.match(issues[0].description, /6 pages where 7 were expected/);
});

test('an unknown page count is said to be unverified, not assumed right', async () => {
  const ctx = await kitchenCtx();
  ctx.result = goodResult({ export: { method: 'save as PDF', pageMap: null, notApplied: [] } });
  ctx.plan = { pages: null };
  const issues = C.resultIssues(kItem, ctx);
  assert.deepEqual(codes(issues), ['page-count-unverified']);
  assert.equal(issues[0].certainty, 'unverified');
});

test('a wrong page size is an error', async () => {
  const ctx = await kitchenCtx({ settings: { size: { preset: 'a4' } } });
  const issues = C.resultIssues(kItem, ctx);
  assert.deepEqual(codes(issues), ['page-size']);
  assert.match(issues[0].description, /not the expected 29\.7 x 21 cm/);
});

test('cropping is flagged for review', async () => {
  const ctx = await kitchenCtx();
  ctx.transform = C.fitTransform(960, 540, 720, 540, 'fill');
  ctx.plan = { known: true, w: 959.98, h: 540, pages: 6 };
  const crop = C.resultIssues(kItem, ctx).find(i => i.code === 'cropped');
  assert.equal(crop.review, true);
  assert.match(crop.description, /12\.5% from the left and from the right/);
});

test('settings PowerPoint did not apply are reported', async () => {
  const ctx = await kitchenCtx();
  ctx.result = goodResult({ export: { method: 'Save As PDF', fallbackReason: 'Type mismatch', pageMap: [1, 2, 4, 5, 6, 7], notApplied: [{ option: 'tags', reason: 'Save As PDF uses PowerPoint defaults.' }] } });
  const issues = C.resultIssues(kItem, ctx);
  assert.deepEqual(codes(issues), ['fallback-export', 'option-not-applied']);
  assert.match(issues[0].description, /used Save As PDF\. The error was: Type mismatch/);
  assert.match(issues[1].description, /"Accessibility tags" was not applied/);
});

test('slides removed on macOS are reported with the renumbering they cause', async () => {
  const ctx = await kitchenCtx({ platform: 'macos' });
  ctx.result = goodResult({ engine: Object.assign({}, PPT, { platform: 'macos' }), export: { method: 'save as PDF', pageMap: [1, 2, 4, 5, 6, 7], removedSlides: [3], notApplied: [] } });
  const issues = C.resultIssues(kItem, ctx);
  assert.deepEqual(codes(issues), ['renumbered']);
  assert.match(issues[0].impact, /lower in the PDF than in the deck/);
  /* no slide number fields, nothing to renumber, nothing to report */
  assert.deepEqual(C.resultIssues({ inspect: { ok: true, kind: 'pptx', slideCount: 7, slides: [{ n: 1, slideNumberField: false }] } }, ctx), []);
});

test('a result that did not come from PowerPoint is refused outright', () => {
  const issues = C.resultIssues(kItem, { settings: C.defaultSettings(), result: goodResult({ engine: { name: 'LibreOffice', version: '24' } }), platform: 'windows' });
  assert.deepEqual(codes(issues), ['engine-unknown']);
  assert.match(issues[0].description, /"LibreOffice", not by Microsoft PowerPoint/);
});

test('a PDF too large to open is reported as not verified, and unapplied edits as an error', () => {
  const issues = C.resultIssues(kItem, { settings: C.defaultSettings(), result: goodResult(), pdfSkipped: 400 * 1024 * 1024, editsSkipped: ['page size'], platform: 'windows' });
  assert.deepEqual(codes(issues), ['verify-skipped', 'edits-skipped']);
  assert.equal(issues[0].certainty, 'unverified');
  assert.equal(issues[1].severity, 'error');
});

test('helper failures become plain issues', () => {
  const t = C.helperFailureIssue({ ok: false, status: 'timeout', timeoutSec: 600 });
  assert.equal(t.code, 'timeout'); assert.match(t.description, /time limit of 10 min 00 s/); assert.match(t.impact, /rest of the batch carried on/);
  assert.equal(C.helperFailureIssue({ ok: false, status: 'cancelled' }).code, 'cancelled');
  assert.equal(C.helperFailureIssue({ ok: false, status: 'failed', error: { code: 'password', message: 'x' } }).code, 'password');
  assert.equal(C.helperFailureIssue({ ok: false, status: 'failed', error: { code: 'exists' } }).code, 'exists');
  assert.match(C.helperFailureIssue({ ok: false, status: 'failed', error: { code: 'open-failed', message: 'PowerPoint found a problem with content.' } }).description, /PowerPoint said: PowerPoint found a problem/);
  assert.equal(C.helperFailureIssue({ ok: false, status: 'failed', error: { code: 'whatever' } }).code, 'convert-failed');
});

test('the four result states', () => {
  const warn = C.issue('video', { severity: 'warning' }), rev = C.issue('visual-diff', { severity: 'warning', certainty: 'suspected', review: true });
  const err = C.issue('page-count', { severity: 'error' }), note = C.issue('transitions', { severity: 'info' });
  const it = issues => ({ status: 'done', hasPdf: true, issues: C.numberIssues(issues) });
  assert.equal(C.classify(it([])), 'completed');
  assert.equal(C.classify(it([note])), 'completed', 'notes do not change the result');
  assert.equal(C.classify(it([warn, note])), 'warnings');
  assert.equal(C.classify(it([warn, rev])), 'review');
  assert.equal(C.classify(it([err])), 'review', 'a PDF with a confirmed error still exists and can be opened');
  assert.equal(C.classify({ status: 'failed', hasPdf: false, issues: [] }), 'failed');
  assert.equal(C.classify({ status: 'cancelled', hasPdf: false, issues: [] }), 'failed');
  assert.equal(C.classify({ status: 'done', hasPdf: false, issues: [] }), 'failed');
  assert.deepEqual(Object.values(C.STATES), ['Completed', 'Completed with warnings', 'Needs review', 'Failed']);
});

test('marking a slide as reviewed keeps the warning attached', () => {
  const item = { status: 'done', hasPdf: true, issues: C.numberIssues([C.issue('visual-diff', { slide: 4, page: 3, severity: 'warning', certainty: 'suspected', review: true })]) };
  assert.equal(C.classify(item), 'review');
  item.issues[0].accepted = '2026-10-07T12:00:00Z';
  assert.equal(C.classify(item), 'warnings', 'accepted, and still not shown as clean');
  assert.equal(item.issues.length, 1);
});

test('picture comparison issues say suspected, and say when nothing was compared', () => {
  const cmp = { performed: true, totalPages: 6, pages: [{ slide: 1, page: 1, changedShare: 0, verdict: 'similar' }, { slide: 5, page: 4, changedShare: 0.083, verdict: 'different' }], skipped: [{ slide: 6, page: 5, reason: 'Its picture could not be read.' }] };
  const issues = C.compareIssues(cmp, C.defaultSettings());
  assert.deepEqual(codes(issues), ['visual-diff', 'compare-skipped', 'compare-sampled']);
  assert.deepEqual([issues[0].slide, issues[0].page, issues[0].certainty, issues[0].review], [5, 4, 'suspected', true]);
  assert.match(issues[0].description, /8\.3% of the compared area/);
  assert.match(issues[0].impact, /can also be a harmless difference/);
  assert.match(issues[2].description, /2 pages of 6 compared\. The others were not checked/);
  const none = C.compareIssues({ performed: false, reason: 'PowerPoint did not save pictures of the slides.' }, C.defaultSettings());
  assert.equal(none[0].code, 'compare-not-performed'); assert.equal(none[0].certainty, 'unverified');
  assert.deepEqual(C.compareIssues(null, C.defaultSettings({ fidelity: 'off' })), []);
  assert.equal(C.fidelityLabel(null, { fidelity: 'off' }), 'Not performed, turned off');
  assert.equal(C.fidelityLabel(cmp, {}), '2 pages of 6 compared, 1 with a suspected difference');
  assert.equal(C.fidelityLabel({ performed: true, pages: [{ verdict: 'similar' }], totalPages: 1 }, {}), '1 page of 1 compared, no difference above the threshold');
});

/* ---------- picture comparison ---------- */

function picture(w, h, draw) {
  const d = new Uint8ClampedArray(w * h * 4).fill(255);
  const set = (x, y, v) => { if (x >= 0 && y >= 0 && x < w && y < h) { const o = (y * w + x) * 4; d[o] = d[o + 1] = d[o + 2] = v; } };
  draw((x0, y0, x1, y1, v) => { for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) set(x, y, v); }, set);
  return d;
}

test('identical pictures are similar', () => {
  const a = picture(384, 216, r => { r(40, 40, 200, 120, 30); r(220, 60, 340, 180, 160); });
  const c = C.compareImages(a, a, 384, 216, 'normal');
  assert.deepEqual([c.flagged, c.changedShare, c.verdict], [0, 0, 'similar']);
});

test('soft edges and faint tone shifts do not raise a flag', () => {
  const a = picture(384, 216, r => { r(40, 40, 200, 120, 30); });
  /* the same shape with a one pixel soft edge and a slightly different grey, as two renderers would draw it */
  const b = picture(384, 216, r => { r(40, 40, 200, 120, 38); r(39, 40, 40, 120, 150); r(200, 40, 201, 120, 150); r(40, 39, 200, 40, 150); r(40, 120, 200, 121, 150); });
  assert.equal(C.compareImages(a, b, 384, 216, 'normal').verdict, 'similar');
});

test('a missing object is flagged, and the flags say where', () => {
  const a = picture(384, 216, r => { r(40, 40, 200, 120, 30); r(240, 80, 340, 180, 60); });
  const b = picture(384, 216, r => { r(40, 40, 200, 120, 30); });
  const c = C.compareImages(a, b, 384, 216, 'normal');
  assert.equal(c.verdict, 'different');
  assert.ok(c.changedShare > 0.05);
  const at = (x, y) => c.flags[Math.floor(y / 8) * c.blocksX + Math.floor(x / 8)];
  assert.equal(at(290, 130), 1, 'flagged where the object was');
  assert.equal(at(100, 80), 0, 'not flagged where nothing changed');
  assert.equal(at(10, 200), 0);
});

test('content that has moved is flagged', () => {
  const a = picture(384, 216, r => { for (let i = 0; i < 8; i++) r(40, 30 + i * 18, 300, 38 + i * 18, 20); });
  const b = picture(384, 216, r => { for (let i = 0; i < 8; i++) r(40, 39 + i * 18, 300, 47 + i * 18, 20); });
  assert.equal(C.compareImages(a, b, 384, 216, 'normal').verdict, 'different', 'lines of text that have reflowed down');
});

test('sensitivity changes how much it takes', () => {
  const a = picture(384, 216, r => { r(40, 40, 200, 120, 30); });
  const b = picture(384, 216, r => { r(40, 40, 200, 120, 30); r(300, 150, 316, 166, 40); });      /* one small extra mark */
  assert.equal(C.compareImages(a, b, 384, 216, 'low').verdict, 'similar');
  assert.equal(C.compareImages(a, b, 384, 216, 'high').verdict, 'different');
});

/* ---------- the batch record and resuming ---------- */

function sampleBatch() {
  const b = C.newBatch({ platform: 'windows', now: '2026-10-07T09:00:00.000Z', settings: { hidden: true } });
  const names = ['A.pptx', 'B.pptx', 'C.pptx', 'D.pptx', 'E.ppt', 'F.pptx'];
  const st = ['done', 'done', 'converting', 'queued', 'ready', 'blocked'];
  names.forEach((n, i) => {
    const it = C.addItem(b, { name: n, size: 1000 * (i + 1), lastModified: 1700000000000 + i });
    it.status = st[i]; it.inspect = { ok: true, kind: 'pptx', slideCount: 3, widthPt: 960, heightPt: 540, hiddenSlides: [] };
    if (st[i] === 'done') { it.hasPdf = true; it.outName = n.replace(/\.\w+$/, '.pdf'); it.pdfBytes = 5000; it.timing = { totalMs: 4000, exportMs: 3000 };
      it.result = { engine: PPT, export: { method: 'ExportAsFixedFormat2', pageMap: [1, 2, 3] } }; it.checks = { pageCount: 3, pageSize: { w: 960, h: 540 }, excludedSlides: [], media: [], fonts: { missing: [], restricted: [], absent: [], substitutes: [] } }; }
  });
  b.items[1].issues = C.numberIssues([C.issue('video', { slide: 2, severity: 'warning', description: 'Video "a, b".mp4', impact: 'Does not play.', fix: 'Send the file.' }),
    C.issue('visual-diff', { slide: 3, page: 3, severity: 'warning', certainty: 'suspected', review: true, description: 'Differs.', impact: 'Maybe.', fix: 'Look.' })]);
  return b;
}

test('item ids are stable, ordered and safe as file names', () => {
  const b = sampleBatch();
  assert.deepEqual(b.items.map(i => i.id), ['f00001', 'f00002', 'f00003', 'f00004', 'f00005', 'f00006']);
  assert.equal(b.items[4].ext, 'ppt');
  assert.ok(b.items.every(i => /^[a-z0-9]+$/.test(i.id)));
});

test('an interrupted batch resumes without redoing finished files', () => {
  const saved = C.serializeBatch(sampleBatch(), '2026-10-07T09:30:00.000Z');
  const b = C.restoreBatch(saved);
  assert.deepEqual(b.items.map(i => i.status), ['done', 'done', 'ready', 'ready', 'ready', 'blocked'], 'what was in flight goes back to waiting, what was done stays done');
  assert.deepEqual(b.items.map(i => !!i.interrupted), [false, false, true, true, false, false]);
  assert.equal(b.items[0].hasPdf, true); assert.equal(b.items[0].outName, 'A.pdf');
  assert.equal(b.settings.hidden, true, 'the settings come back with it');
  assert.deepEqual(C.resumeSummary(b), { done: 2, left: 4, total: 6, interrupted: 2 });
  assert.equal(b.nextId, 7, 'new files keep counting where the batch left off');
  assert.throws(() => C.restoreBatch('{"version":9,"items":[]}'), /not a SlideSize batch record/);
});

test('files are matched again by name, size and date when access has to be confirmed', () => {
  const it = { name: 'A.pptx', size: 1000, lastModified: 1700000000000 };
  assert.equal(C.sameFile(it, { name: 'A.pptx', size: 1000, lastModified: 1700000000000 }), true);
  assert.equal(C.sameFile(it, { name: 'A.pptx', size: 1001, lastModified: 1700000000000 }), false);
  assert.equal(C.sameFile(it, { name: 'A.pptx', size: 1000, lastModified: 1700000009999 }), false, 'an edited file is not the same file');
  assert.equal(C.sameFile(it, { name: 'a.pptx', size: 1000, lastModified: 1700000000000 }), false);
});

test('the saved record of a file is small', async () => {
  const ins = await I.inspect(deck('Kitchen sink 16x9.pptx'), 'k.pptx');
  const slim = C.slimInspect(ins);
  assert.ok(JSON.stringify(slim).length < 6000, 'a few kilobytes a file, whatever the deck weighs');
  assert.equal(typeof slim.readBytes, 'undefined', 'no handle on the file is kept');
  assert.equal(slim.slides[3].media[0].box.width, 576);
});

test('the preview cache has a hard ceiling and frees what it drops', () => {
  const freed = [], c = new C.LRU(3, (v, k) => freed.push(k));
  ['a', 'b', 'c'].forEach(k => c.set(k, k.toUpperCase()));
  assert.equal(c.get('a'), 'A');                 /* a is now the most recently used */
  c.set('d', 'D'); c.set('e', 'E');
  assert.equal(c.size, 3);
  assert.deepEqual(freed, ['b', 'c']);
  assert.equal(c.get('b'), undefined); assert.equal(c.get('a'), 'A');
  for (let i = 0; i < 500; i++) c.set('k' + i, i);
  assert.equal(c.size, 3, 'five hundred previews later it still holds three');
  c.clear(); assert.equal(c.size, 0);
});

test('the summary counts each file once', () => {
  const s = C.summarize(sampleBatch().items);
  assert.deepEqual([s.total, s.finished, s.completed, s.warnings, s.review, s.failed, s.pending, s.blocked], [6, 2, 1, 0, 1, 0, 3, 1]);
  assert.equal(s.pages, 6); assert.equal(s.ms, 8000);
});

/* ---------- reports ---------- */

test('the JSON report records settings, engine, timing and every issue in full', () => {
  const d = JSON.parse(C.reportJson(sampleBatch(), { now: '2026-10-07T09:31:00.000Z' }));
  assert.equal(d.files.length, 6);
  assert.equal(d.summary.review, 1);
  const f = d.files[1];
  assert.deepEqual([f.file, f.output, f.state, f.stateLabel], ['B.pptx', 'B.pdf', 'review', 'Needs review']);
  assert.deepEqual(f.engine, PPT);
  assert.equal(f.timing.totalMs, 4000);
  assert.equal(f.settings.hidden, true);
  assert.ok(f.settingsSummary.includes('Hidden slides included'));
  assert.deepEqual(Object.keys(f.issues[1]).sort(), ['accepted', 'certainty', 'code', 'description', 'file', 'fix', 'id', 'impact', 'needsReview', 'page', 'severity', 'slide'].sort());
  assert.deepEqual([f.issues[1].file, f.issues[1].slide, f.issues[1].page, f.issues[1].severity, f.issues[1].certainty], ['B.pptx', 3, 3, 'warning', 'suspected']);
  assert.equal(d.files[2].output, null, 'a file with no PDF claims none');
  assert.equal(d.files[2].stateLabel, 'Converting');
  assert.ok(d.notes.some(n => /evidence, not a guarantee/.test(n)));
});

test('the CSV has one row per issue and survives awkward text', () => {
  const b = sampleBatch();
  b.items[0].name = '=HYPERLINK("x")';
  const csv = C.reportCsv(b), lines = csv.replace(/^\ufeff/, '').trim().split('\r\n');
  assert.equal(lines.length, 1 + 1 + 2 + 4, 'a header, one row for the clean file, two for the file with two issues, one each for the rest');
  assert.match(lines[0], /^File,Output PDF,Result,Slide,PDF page,Severity,Certainty,Issue,Description,Likely impact,Suggested fix/);
  assert.ok(lines[1].startsWith('"\'=HYPERLINK(""x"")"'), 'a file name cannot run as a formula');
  assert.match(lines[2], /"Video ""a, b""\.mp4"/);
  assert.match(lines[3], /B\.pptx,B\.pdf,Needs review,3,3,Warning,Suspected,visual-diff/);
});

test('the HTML report is self contained and escapes file names', () => {
  const b = sampleBatch();
  b.items[0].name = '<script>alert(1)</script>.pptx';
  b.items[1].issues[1].accepted = '2026-10-07T09:20:00Z';
  const html = C.reportHtml(b);
  assert.ok(!html.includes('<script>alert(1)'), 'no markup from a file name');
  assert.ok(html.includes('&lt;script&gt;alert(1)&lt;/script&gt;.pptx'));
  assert.ok(!/<link|<script|src=|href="http/.test(html), 'nothing is fetched when the report is opened');
  for (const need of ['Completed', 'Completed with warnings', 'Excluded slides', 'Font substitutions', 'Rasterised text', 'Media handling', 'Fidelity check', 'PDF/A validation',
    'Processing time', 'Microsoft PowerPoint 16.0 build 17328', 'ExportAsFixedFormat2', 'Reviewed and accepted 2026-10-07T09:20:00Z', 'Likely impact', 'Suggested fix', 'evidence, not a guarantee']) {
    assert.ok(html.includes(need), 'the report mentions ' + need);
  }
});

test('reports can be written before anything has finished', () => {
  const b = C.newBatch({ platform: 'macos' });
  C.addItem(b, { name: 'A.pptx', size: 10 });
  assert.doesNotThrow(() => { C.reportHtml(b); C.reportCsv(b); C.reportJson(b); });
  assert.equal(JSON.parse(C.reportJson(b)).files[0].stateLabel, 'Not checked');
});

/* ---------- formatting ---------- */

test('sizes, times and slide lists read naturally', () => {
  assert.equal(C.sizeLabel(960, 540), '33.87 x 19.05 cm');
  assert.equal(C.aspectLabel(960, 540), '16:9'); assert.equal(C.aspectLabel(959.98, 540), '16:9'); assert.equal(C.aspectLabel(1920, 540), '32:9');
  assert.equal(C.aspectLabel(595.28, 841.89), '1:1.41');
  assert.equal(C.formatDuration(8200), '8.2 s'); assert.equal(C.formatDuration(75000), '1 min 15 s'); assert.equal(C.formatDuration(3720000), '1 h 2 min');
  assert.equal(C.formatBytes(1536), '2 KB'); assert.equal(C.formatBytes(2.5 * 1024 * 1024 * 1024), '2.50 GB');
  assert.equal(C.listSlides([7, 1, 2, 3, 9, 10]), '1 to 3, 7, 9, 10');
  assert.equal(C.plural(1, 'slide'), '1 slide'); assert.equal(C.plural(2, 'slide'), '2 slides');
});
