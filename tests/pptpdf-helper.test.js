/* PowerPoint to PDF helper checks.

   These run the two helper scripts for real, as the page would, by writing
   job tickets into a folder and reading the results back. PowerPoint itself
   is replaced by a stand-in (tests/pptpdf/fake-engine.ps1 for Windows,
   tests/pptpdf/fake-osascript.js for macOS), so what is checked here is the
   queue, the time limits, cancelling, recovery, resuming and the result
   files. Whether PowerPoint is driven correctly can only be checked on a
   computer that has PowerPoint.

   The Windows helper needs PowerShell. Set PWSH to its path, or have pwsh
   or powershell on the path. Without it those checks are skipped.
   The macOS helper needs a POSIX sh, so its checks are skipped on Windows.

   Run from the repository root with:  node --test                         */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'), os = require('os'), path = require('path');
const { spawn, spawnSync } = require('child_process');
const Core = require('../pptpdf-core.js');
const Pdf = require('../pptpdf-pdf.js');

const HELPERS = path.join(__dirname, '..', 'helper');
const FAKES = path.join(__dirname, 'pptpdf');
const FAKE_NAME = 'SlideSize test double (not PowerPoint)';

function findPwsh() {
  const cands = [process.env.PWSH, 'pwsh', 'powershell'].filter(Boolean);
  for (const c of cands) {
    const r = spawnSync(c, ['-NoProfile', '-Command', '$PSVersionTable.PSVersion.Major'], { encoding: 'utf8' });
    if (r.status === 0) return c;
  }
  return null;
}
const PWSH = findPwsh();
const KINDS = [];
if (PWSH) KINDS.push('windows');
if (process.platform !== 'win32') KINDS.push('macos');

const sleep = ms => new Promise(r => setTimeout(r, ms));
const readJson = f => { try { return JSON.parse(fs.readFileSync(f, 'utf8')); } catch (e) { return null; } };

async function until(fn, ms, what) {
  const end = Date.now() + ms;
  for (;;) {
    const v = await fn();
    if (v) return v;
    if (Date.now() > end) throw new Error('Timed out waiting for ' + what);
    await sleep(100);
  }
}

class Rig {
  constructor(kind, env) {
    this.kind = kind;
    this.out = fs.mkdtempSync(path.join(os.tmpdir(), 'slidesize-helper-'));
    this.root = path.join(this.out, '_slidesize');
    this.state = fs.mkdtempSync(path.join(os.tmpdir(), 'slidesize-fake-'));
    for (const d of ['queue', 'in', 'active']) fs.mkdirSync(path.join(this.root, d), { recursive: true });
    fs.writeFileSync(path.join(this.root, 'marker.json'), JSON.stringify({ protocol: Core.PROTOCOL }));
    this.env = env || {};
    this.seq = 0; this.n = 0; this.proc = null; this.log = '';
  }
  start() {
    const env = Object.assign({}, process.env, { SLIDESIZE_FAKE_STATE: this.state }, this.env);
    if (this.kind === 'windows') {
      env.SLIDESIZE_HELPER_TEST_ENGINE = path.join(FAKES, 'fake-engine.ps1');
      this.proc = spawn(PWSH, ['-NoProfile', '-File', path.join(HELPERS, 'slidesize-helper.ps1'), '-Root', this.root], { env });
    } else {
      env.SLIDESIZE_HELPER_OSASCRIPT = path.join(FAKES, 'fake-osascript.js');
      env.SLIDESIZE_HELPER_ROOT = this.root;
      env.SLIDESIZE_HELPER_TEST_INSTALLED = this.env.SLIDESIZE_FAKE_NOT_INSTALLED === '1' ? '' : '1';
      env.SLIDESIZE_HELPER_TEST_ENGINE_NAME = FAKE_NAME;
      this.proc = spawn('sh', [path.join(HELPERS, 'Start-SlideSize-Helper.command')], { env });
    }
    this.exited = new Promise(r => this.proc.on('exit', code => { this.code = code; r(code); }));
    this.proc.stdout.on('data', d => { this.log += d; });
    this.proc.stderr.on('data', d => { this.log += d; });
    return this;
  }
  helper() { return readJson(path.join(this.root, 'helper.json')); }
  async ready() {
    try { return await until(() => { const h = this.helper(); return h && h.state === 'idle' && h.pid === this.proc.pid ? h : null; }, 30000, 'the helper to be ready'); }
    catch (e) { throw new Error(e.message + '\n' + this.log + '\n' + JSON.stringify(this.helper())); }
  }
  /* writes the temporary copy and the ticket, exactly as the page does */
  submit(deck, settings, over) {
    const id = Core.ticketId(++this.n);
    const item = Object.assign({ id, name: 'Deck ' + this.n + '.pptx', ext: 'pptx', outName: 'Deck ' + this.n + '.pdf', overwrite: false,
                                 inspect: { ok: true, kind: 'pptx', slides: [] } }, (over && over.item) || {});
    const t = Core.buildTicket(item, Core.defaultSettings(settings));
    Object.assign(t, (over && over.ticket) || {});
    fs.writeFileSync(path.join(this.root, t.input), JSON.stringify(deck));
    const tmp = path.join(this.root, 'queue', id + '.json.crswap');
    fs.writeFileSync(tmp, JSON.stringify(t));
    fs.renameSync(tmp, path.join(this.root, 'queue', id + '.json'));
    return { id, ticket: t, item };
  }
  done(id, ms) {
    return until(() => readJson(path.join(this.root, 'done', id + '.json')), ms || 30000, 'the result of ' + id)
      .catch(e => { throw new Error(e.message + '\n' + this.log + '\n' + JSON.stringify(this.helper())); });
  }
  control(o) { fs.writeFileSync(path.join(this.root, 'control.json'), JSON.stringify(Object.assign({ seq: ++this.seq, stop: false, abort: null, retry: false }, o))); }
  async stop() {
    if (!this.proc || this.code !== undefined) return;
    this.control({ stop: true });
    const t = setTimeout(() => { try { this.proc.kill('SIGKILL'); } catch (e) { /* gone */ } }, 20000);
    await this.exited; clearTimeout(t);
  }
  clean() { for (const d of [this.out, this.state]) fs.rmSync(d, { recursive: true, force: true }); }
}

async function withRig(kind, env, fn) {
  const rig = new Rig(kind, env);
  try { rig.start(); await fn(rig); }
  finally { await rig.stop(); rig.clean(); }
}

async function pages(file) { return (await Pdf.inspect(new Uint8Array(fs.readFileSync(file)))).pageCount; }

if (!KINDS.length) test('helper checks', { skip: 'No PowerShell and no POSIX shell on this machine.' }, () => {});
if (!PWSH) test('Windows helper', { skip: 'PowerShell was not found. Set PWSH to its path to run these.' }, () => {});

for (const kind of KINDS) {
  const T = (name, fn) => test(kind + ' helper: ' + name, { timeout: 120000 }, fn);

  T('reports itself, its engine and what it can do', () => withRig(kind, {}, async rig => {
    const h = await rig.ready();
    assert.equal(h.protocol, Core.PROTOCOL);
    assert.equal(h.helper, Core.HELPER_VERSION);
    assert.equal(h.powerpoint.installed, true);
    assert.equal(h.engine, FAKE_NAME);
    assert.equal(h.caps.hidden, true);
    assert.equal(h.caps.reference, true);
    assert.equal(h.caps.pdfa, kind === 'windows');
    assert.equal(h.caps.notes, kind === 'windows');
    assert.equal(Core.helperStatus(h, Date.now()).state, 'ready');
    const fonts = await until(() => readJson(path.join(rig.root, 'fonts.json')), 10000, 'the font list');
    assert.ok(fonts.families.includes('Carlito'));
  }));

  T('says so when PowerPoint is not installed', () => withRig(kind, { SLIDESIZE_FAKE_NOT_INSTALLED: '1' }, async rig => {
    const h = await until(() => { const x = rig.helper(); return x && x.pid === rig.proc.pid ? x : null; }, 30000, 'the first heartbeat');
    assert.equal(h.powerpoint.installed, false);
    assert.equal(Core.helperStatus(h, Date.now()).state, 'no-powerpoint');
    assert.match(Core.helperStatus(h, Date.now()).text, /Microsoft PowerPoint must be installed for PowerPoint-based conversion/);
  }));

  T('converts one file, leaves hidden slides out and cleans up its copy', () => withRig(kind, {}, async rig => {
    await rig.ready();
    const j = rig.submit({ slides: 6, hidden: [3] }, { fidelity: 'all' });
    const r = await rig.done(j.id);
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.equal(r.status, 'done');
    assert.equal(r.engine.name, FAKE_NAME);
    assert.equal(r.facts.slides, 6);
    assert.deepEqual(r.facts.hidden, [3]);
    assert.deepEqual(r.export.pageMap, [1, 2, 4, 5, 6]);
    assert.equal(await pages(path.join(rig.out, 'Deck 1.pdf')), 5);
    assert.equal(fs.existsSync(path.join(rig.root, j.ticket.input)), false, 'the temporary copy is deleted');
    assert.equal(fs.existsSync(path.join(rig.root, 'queue', j.id + '.json')), false);
    assert.equal(fs.existsSync(path.join(rig.root, 'active', j.id + '.json')), false);
    assert.deepEqual(r.reference.slides, [1, 2, 4, 5, 6]);
    for (const n of r.reference.slides) assert.ok(fs.existsSync(path.join(rig.root, 'ref', j.id, 'slide-' + String(n).padStart(4, '0') + '.png')));
    assert.ok(r.timing.totalMs >= 0);
    assert.equal(r.pdf.bytes, fs.statSync(path.join(rig.out, 'Deck 1.pdf')).size);
  }));

  T('includes hidden slides and keeps to a slide range when asked', () => withRig(kind, {}, async rig => {
    await rig.ready();
    const a = rig.submit({ slides: 6, hidden: [3] }, { hidden: true, fidelity: 'off' });
    const b = rig.submit({ slides: 8, hidden: [3, 7] }, { range: { from: 2, to: 5 }, fidelity: 'off' });
    const c = rig.submit({ slides: 8, hidden: [3, 7] }, { range: { from: 2, to: 50 }, hidden: true, fidelity: 'off' });
    const ra = await rig.done(a.id), rb = await rig.done(b.id), rc = await rig.done(c.id);
    assert.deepEqual(ra.export.pageMap, [1, 2, 3, 4, 5, 6]);
    assert.equal(await pages(path.join(rig.out, 'Deck 1.pdf')), 6);
    assert.deepEqual(rb.export.pageMap, [2, 4, 5]);
    assert.equal(await pages(path.join(rig.out, 'Deck 2.pdf')), 3);
    assert.deepEqual(rc.export.pageMap, [2, 3, 4, 5, 6, 7, 8]);
    assert.equal(await pages(path.join(rig.out, 'Deck 3.pdf')), 7);
    assert.equal(ra.reference, null);
  }));

  T('samples the reference pictures when asked', () => withRig(kind, {}, async rig => {
    await rig.ready();
    const j = rig.submit({ slides: 21 }, { fidelity: 'sample' });
    const r = await rig.done(j.id);
    assert.deepEqual(r.reference.slides, Core.sampleSlides(r.export.pageMap));
    assert.deepEqual(r.reference.slides, [1, 6, 11, 16, 21]);
    /* seven slides puts two of the five picks exactly half way between slides. Every part has to round the same way. */
    const k = rig.submit({ slides: 7 }, { fidelity: 'sample' });
    const rk = await rig.done(k.id);
    assert.deepEqual(rk.reference.slides, Core.sampleSlides(rk.export.pageMap));
    assert.deepEqual(rk.reference.slides, [1, 3, 4, 6, 7]);
  }));

  T('writes a PDF the page still has to edit into the work folder', () => withRig(kind, {}, async rig => {
    await rig.ready();
    const j = rig.submit({ slides: 2 }, { size: { preset: 'a4' }, fidelity: 'off' });
    assert.equal(j.ticket.pdf, 'out/' + j.id + '.pdf');
    const r = await rig.done(j.id);
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.equal(await pages(path.join(rig.root, 'out', j.id + '.pdf')), 2);
    assert.equal(fs.existsSync(path.join(rig.out, 'Deck 1.pdf')), false);
  }));

  T('never overwrites an existing PDF unless told to', () => withRig(kind, {}, async rig => {
    await rig.ready();
    const target = path.join(rig.out, 'Deck 1.pdf');
    fs.writeFileSync(target, 'somebody else\'s file');
    const a = rig.submit({ slides: 2 }, { fidelity: 'off' });
    const ra = await rig.done(a.id);
    assert.equal(ra.ok, false);
    assert.equal(ra.error.code, 'exists');
    assert.equal(fs.readFileSync(target, 'utf8'), 'somebody else\'s file');
    const b = rig.submit({ slides: 2 }, { fidelity: 'off' }, { item: { outName: 'Deck 1.pdf', overwrite: true } });
    const rb = await rig.done(b.id);
    assert.equal(rb.ok, true, JSON.stringify(rb));
    assert.equal(await pages(target), 2);
  }));

  T('a stalled file times out and the queue carries on', () => withRig(kind, {}, async rig => {
    await rig.ready();
    const a = rig.submit({ slides: 3, stall: true }, { fidelity: 'off' }, { ticket: { timeoutSec: 3 } });
    const b = rig.submit({ slides: 4 }, { fidelity: 'off' });
    const ra = await rig.done(a.id, 40000);
    assert.equal(ra.ok, false);
    assert.equal(ra.status, 'timeout');
    assert.equal(Core.helperFailureIssue(ra).code, 'timeout');
    assert.equal(fs.existsSync(path.join(rig.out, 'Deck 1.pdf')), false);
    const rb = await rig.done(b.id, 60000);
    assert.equal(rb.ok, true, JSON.stringify(rb));
    assert.equal(await pages(path.join(rig.out, 'Deck 2.pdf')), 4);
    const h = await until(() => { const x = rig.helper(); return x && x.converted === 1 && x.failed === 1 ? x : null; }, 10000, 'the counts');
    assert.equal(h.failed, 1);
  }));

  T('a file that fails to open is reported and the next one converts', () => withRig(kind, {}, async rig => {
    await rig.ready();
    const a = rig.submit({ slides: 3, fail: 'The file is damaged and could not be repaired.' }, { fidelity: 'off' });
    const b = rig.submit({ slides: 1 }, { fidelity: 'off' });
    const ra = await rig.done(a.id), rb = await rig.done(b.id);
    assert.equal(ra.ok, false);
    assert.equal(ra.error.code, 'open-failed');
    assert.match(ra.error.message, /damaged/);
    assert.equal(fs.existsSync(path.join(rig.root, a.ticket.input)), false, 'the copy of a failed file is deleted too');
    assert.equal(rb.ok, true);
  }));

  T('the page can skip the file in hand', () => withRig(kind, {}, async rig => {
    await rig.ready();
    const a = rig.submit({ slides: 3, stall: true }, { fidelity: 'off' });
    const b = rig.submit({ slides: 2 }, { fidelity: 'off' });
    await until(() => { const h = rig.helper(); return h && h.current && h.current.id === a.id; }, 20000, 'the first file to start');
    rig.control({ abort: a.id });
    const ra = await rig.done(a.id, 30000);
    assert.equal(ra.status, 'cancelled');
    const rb = await rig.done(b.id, 60000);
    assert.equal(rb.ok, true, JSON.stringify(rb));
  }));

  T('a ticket that points outside the output folder is refused', () => withRig(kind, {}, async rig => {
    await rig.ready();
    const outside = path.join(path.dirname(rig.out), 'slidesize-escape-' + process.pid + '.pdf');
    const a = rig.submit({ slides: 2 }, { fidelity: 'off' }, { ticket: { pdf: '../../' + path.basename(outside) } });
    const ra = await rig.done(a.id);
    assert.equal(ra.ok, false);
    assert.equal(ra.error.code, 'bad-ticket');
    assert.equal(fs.existsSync(outside), false);
    const b = rig.submit({ slides: 2 }, { fidelity: 'off' }, { ticket: { input: '../Deck 9.pptx' } });
    const rb = await rig.done(b.id);
    assert.equal(rb.error.code, 'bad-ticket');
  }));

  T('stops when asked, and a second start picks up a ticket left half done', async () => {
    const rig = new Rig(kind, {});
    try {
      rig.start();
      await rig.ready();
      await rig.stop();
      assert.equal(rig.helper().state, 'stopped');
      assert.equal(Core.helperStatus(rig.helper(), Date.now()).state, 'stopped');
      assert.equal(fs.existsSync(path.join(rig.root, 'helper.lock')), false);

      /* as if the helper had been closed in the middle of a file */
      const j = rig.submit({ slides: 5 }, { fidelity: 'off' });
      fs.renameSync(path.join(rig.root, 'queue', j.id + '.json'), path.join(rig.root, 'active', j.id + '.json'));
      rig.code = undefined; rig.start();
      const r = await rig.done(j.id);
      assert.equal(r.ok, true, JSON.stringify(r));
      assert.equal(await pages(path.join(rig.out, 'Deck 1.pdf')), 5);
    } finally { await rig.stop(); rig.clean(); }
  });

  T('only one helper runs per folder', () => withRig(kind, {}, async rig => {
    await rig.ready();
    const second = new Rig(kind, {});
    second.clean();
    second.out = rig.out; second.root = rig.root; second.state = rig.state;
    second.start();
    const code = await second.exited;
    assert.notEqual(code, 0);
    assert.match(second.log, /already running/);
    const j = rig.submit({ slides: 1 }, { fidelity: 'off' });
    assert.equal((await rig.done(j.id)).ok, true, 'the first helper is unaffected');
  }));

  T('reports a PowerPoint that will not start instead of spinning', () => withRig(kind, { SLIDESIZE_FAKE_START_FAILS: '1' }, async rig => {
    const h = await until(() => { const x = rig.helper(); return x && x.state === 'blocked' ? x : null; }, 30000, 'the blocked state');
    assert.match(h.message, /could not be started/);
    assert.equal(Core.helperStatus(h, Date.now()).state, 'blocked');
  }));

  T('the page rejects a result that did not come from PowerPoint', () => withRig(kind, {}, async rig => {
    await rig.ready();
    const j = rig.submit({ slides: 2 }, { fidelity: 'off' });
    const r = await rig.done(j.id);
    const issues = Core.resultIssues(j.item, { settings: Core.defaultSettings(), result: r, platform: kind });
    assert.equal(issues.length, 1);
    assert.equal(issues[0].code, 'engine-unknown');
    assert.equal(issues[0].severity, 'error');
  }));
}

if (KINDS.includes('macos')) {
  const T = (name, fn) => test('macos helper: ' + name, { timeout: 120000 }, fn);

  T('when PowerPoint exports hidden slides anyway, removes them from the copy and says so', () =>
    withRig('macos', { SLIDESIZE_FAKE_EXPORTS_HIDDEN: '1' }, async rig => {
      await rig.ready();
      const a = rig.submit({ slides: 6, hidden: [3, 5] }, { fidelity: 'all' });
      const ra = await rig.done(a.id);
      assert.equal(ra.ok, true, JSON.stringify(ra));
      assert.deepEqual(ra.export.pageMap, [1, 2, 4, 6]);
      assert.deepEqual(ra.export.removedSlides, [3, 5]);
      assert.equal(ra.export.hiddenMethod, 'delete');
      assert.equal(await pages(path.join(rig.out, 'Deck 1.pdf')), 4);
      /* the pictures were taken before anything was removed, so they keep their slide numbers */
      assert.deepEqual(ra.reference.slides, [1, 2, 4, 6]);
      assert.equal(fs.readFileSync(path.join(rig.root, 'ref', a.id, 'slide-0006.png'), 'utf8'), 'picture of original slide 6');
      const issues = Core.resultIssues({ inspect: { ok: true, kind: 'pptx', slideCount: 6, slides: [{ n: 1, slideNumberField: true }] } },
        { settings: Core.defaultSettings(), result: Object.assign({}, ra, { engine: { name: Core.ENGINE_NAME } }), platform: 'macos' });
      assert.ok(issues.some(i => i.code === 'renumbered'), 'the renumbering is reported');

      const b = rig.submit({ slides: 6, hidden: [3, 5] }, { hidden: true, fidelity: 'off' });
      const rb = await rig.done(b.id);
      assert.deepEqual(rb.export.pageMap, [1, 2, 3, 4, 5, 6]);
      assert.deepEqual(rb.export.removedSlides, []);
    }));

  T('when PowerPoint skips hidden slides, a slide range is done by hiding and nothing is removed', () =>
    withRig('macos', {}, async rig => {
      await rig.ready();
      const a = rig.submit({ slides: 8, hidden: [3] }, { range: { from: 2, to: 6 }, hidden: true, fidelity: 'off' });
      const ra = await rig.done(a.id);
      assert.deepEqual(ra.export.pageMap, [2, 3, 4, 5, 6]);
      assert.deepEqual(ra.export.removedSlides, []);
      assert.equal(ra.export.hiddenMethod, 'hide');
      assert.equal(await pages(path.join(rig.out, 'Deck 1.pdf')), 5);
    }));

  T('says when it could not find out which slides are hidden', () =>
    withRig('macos', { SLIDESIZE_FAKE_NO_HIDDEN_TERM: '1' }, async rig => {
      await rig.ready();
      const a = rig.submit({ slides: 4, hidden: [2] }, { fidelity: 'off' });
      const ra = await rig.done(a.id);
      assert.equal(ra.ok, true, JSON.stringify(ra));
      assert.equal(ra.facts.hidden, null);
      assert.equal(ra.export.pageMap, null, 'the page map is not claimed when it is not known');
      assert.equal(ra.export.notApplied[0].option, 'hidden');
      assert.equal(ra.export.pages, 3);
    }));

  T('adds video placeholders and reports one it could not add', () =>
    withRig('macos', {}, async rig => {
      await rig.ready();
      const ins = { ok: true, kind: 'pptx', slides: [
        { n: 2, media: [{ type: 'video', name: 'Sting.mp4', box: { left: 10.4, top: 20, width: 300, height: 200 }, posterPart: null }] },
        { n: 9, media: [{ type: 'video', name: 'Gone "quoted".mp4', box: { left: 0, top: 0, width: 10, height: 10 }, posterPart: 'x.png', posterBlank: true }] }] };
      const a = rig.submit({ slides: 3 }, { fidelity: 'off' }, { item: { inspect: ins } });
      assert.equal(a.ticket.options.placeholders.length, 2);
      const ra = await rig.done(a.id);
      assert.equal(ra.ok, true, JSON.stringify(ra));
      assert.deepEqual(ra.export.placeholders.map(p => [p.slide, p.ok]), [[2, true], [9, false]]);
      assert.match(ra.export.placeholders[1].error, /slide 9/);
    }));
}

/* ---------- the helper files themselves ---------- */

test('the Windows helper is plain ASCII and keeps to Windows PowerShell 5.1', () => {
  const ps = fs.readFileSync(path.join(HELPERS, 'slidesize-helper.ps1'), 'utf8');
  assert.ok(!/[^\x00-\x7F]/.test(ps), 'PowerShell 5.1 misreads a script with no byte order mark as soon as it holds anything but ASCII');
  const code = ps.split('\n').filter(l => !/^\s*#/.test(l)).join('\n');
  for (const [re, what] of [[/\?\?/, 'the ?? operator'], [/\?\./, 'the ?. operator'], [/&&|\|\|/, 'pipeline chain operators'], [/-AsHashtable/, 'ConvertFrom-Json -AsHashtable'],
    [/\$Is(Windows|Linux|MacOS)\b/, '$IsWindows'], [/ForEach-Object\s+-Parallel/, 'ForEach-Object -Parallel'], [/\bclean\s*\{/, 'clean blocks']]) {
    assert.ok(!re.test(code), 'the script uses ' + what + ', which PowerShell 5.1 does not have');
  }
  assert.match(ps, /\$HelperVersion = '1\.0\.0'/);
  assert.equal(Core.HELPER_VERSION, '1.0.0');
});

test('the Windows launcher has Windows line endings and the Mac helper does not', () => {
  const cmd = fs.readFileSync(path.join(HELPERS, 'Start-SlideSize-Helper.cmd'), 'latin1');
  assert.ok(/\r\n/.test(cmd) && !/[^\r]\n/.test(cmd));
  assert.match(cmd, /powershell\.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0slidesize-helper\.ps1"/);
  const mac = fs.readFileSync(path.join(HELPERS, 'Start-SlideSize-Helper.command'), 'latin1');
  assert.ok(!mac.includes('\r'));
  assert.ok(mac.startsWith('#!/bin/sh\n'));
  assert.match(mac, /HELPER_VERSION="1\.0\.0"/);
});

test('both helpers name the real engine unless a stand-in says otherwise', () => {
  assert.match(fs.readFileSync(path.join(HELPERS, 'slidesize-helper.ps1'), 'utf8'), /function Engine-Name \{ return 'Microsoft PowerPoint' \}/);
  assert.match(fs.readFileSync(path.join(HELPERS, 'Start-SlideSize-Helper.command'), 'utf8'), /ENGINE_NAME="\$\{SLIDESIZE_HELPER_TEST_ENGINE_NAME:-Microsoft PowerPoint\}"/);
  assert.equal(Core.ENGINE_NAME, 'Microsoft PowerPoint');
});
