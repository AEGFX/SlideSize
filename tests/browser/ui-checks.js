const L = require('./lib.js'), A = require('./analyse.js'); const S = process.env.WORK; const fs = require('fs');
const results = [];
function check(name, ok, detail) { results.push({ name, ok: !!ok, detail: detail || '' }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  [' + detail + ']' : '')); }
const only = process.argv[2] ? new RegExp(process.argv[2]) : null;
const want = n => !only || only.test(n);
const text = (page, sel) => page.evaluate(s => { const e = document.querySelector(s); return e ? e.textContent : null; }, sel);
const st = (page, fn) => page.evaluate(src => (new Function('s', 'return (' + src + ')(s)'))(SlideshowApp.state()), fn.toString());

(async () => {
  const land = L.listPhotos(S + '/photos/set40', f => /land/.test(f));
  const mixed = L.listPhotos(S + '/photos/set12');

  /* ---------- F: filtering through the interface ---------- */
  if (want('F')) {
    const { browser, page, logs } = await L.open();
    await L.addPhotos(page, mixed);
    let plan = await L.planInfo(page);
    const rep = await st(page, s => ({ added: s.imp.report.added, failed: s.imp.report.failed.map(f => f.name), notImages: s.imp.report.notImages }));
    check('F1 unreadable files are reported and the valid ones still load', rep.added === 15 && rep.failed.join() === 'broken.jpg,phone.heic' && rep.notImages.join() === 'notes.txt' && plan.counts.total === 15, JSON.stringify(rep));
    check('F2 the import report names each problem file', /broken\.jpg\. Could not be decoded/.test(await text(page, '#import-status')) && /phone\.heic\. HEIC is not readable/.test(await text(page, '#import-status')));
    const exif = await st(page, s => { const p = s.photos.find(x => /exif/.test(x.name)); return [p.iw, p.ih]; });
    check('F3 orientation metadata is honoured', exif[0] === 2000 && exif[1] === 3000, 'stored 3000 x 2000 with rotation, read as ' + exif.join(' x '));
    check('F4 16:9 keeps 3:2, 4:3, 5:4 and 16:9, drops portrait, square and panorama', plan.counts.included === 10 && plan.excluded.length === 5 && plan.order.every(n => /land|small|alpha|webp/.test(n)), plan.order.length + ' in');
    // same files again: duplicates reported, not silently dropped or doubled
    await L.addPhotos(page, mixed.filter(f => /IMG_00[12]/.test(f)));
    const dup = await text(page, '#import-status');
    check('F5 duplicates are reported with a way to add them anyway', /2 duplicates skipped/.test(dup) && (await L.planInfo(page)).counts.total === 15, dup.slice(0, 90));
    await page.click('#import-status button:has-text("anyway")');
    await page.waitForFunction(() => !SlideshowApp.state().imp.running);
    await page.waitForTimeout(200);
    check('F6 duplicates can be added on request', (await L.planInfo(page)).counts.total === 17);
    // portrait output through the preset menu
    await page.selectOption('#preset', 'p1080');
    await page.waitForTimeout(250);
    plan = await L.planInfo(page);
    check('F7 portrait output flips the decision', plan.order.every(n => /port|exif/.test(n)) && plan.counts.included === 3 && plan.excluded.some(e => /^Landscape photo, too much cropping for this portrait output/.test(e.reason)), plan.order.join(','));
    await page.selectOption('#preset', '1080p');
    await page.waitForTimeout(250);
    // restore one excluded photo from the excluded tray, no reimport
    await page.evaluate(() => { document.getElementById('excluded-box').open = true; });
    const firstEx = await text(page, '#excluded-list .ex-item .ex-name');
    await page.click('#excluded-list .ex-item button:has-text("Include anyway")');
    await page.waitForTimeout(250);
    plan = await L.planInfo(page);
    check('F8 an excluded photo is restored from the tray without reimporting', plan.order.includes(firstEx) && plan.counts.total === 17, firstEx);
    // and it survives a change of output size as the user's own choice
    await page.selectOption('#preset', '2160p'); await page.waitForTimeout(250);
    check('F9 the manual choice survives a size change', (await L.planInfo(page)).order.includes(firstEx));
    // threshold slider
    await page.evaluate(() => { const t = document.getElementById('threshold'); t.value = 45; t.dispatchEvent(new Event('input', { bubbles: true })); });
    await page.waitForTimeout(250);
    const more = await L.planInfo(page);
    check('F10 lowering the threshold lets the square photo in', more.order.some(n => /square/.test(n)) && more.counts.included > plan.counts.included);
    // filter off
    await page.uncheck('#exclude'); await page.waitForTimeout(250);
    const off = await L.planInfo(page);
    check('F11 filter off uses every photo and offers Fill, Fit and Blurred background', off.counts.excluded === 0 && await page.isVisible('#fit-mode button[data-v="blur"]'));
    // nothing left: ultra wide
    await page.check('#exclude');
    await page.evaluate(() => { const s = SlideshowApp.state(); s.photos.forEach(p => { p.include = 'auto'; }); s.P.framing.threshold = 0.55; SlideshowApp.refresh(); });
    await page.fill('#out-w', '5760'); await page.fill('#out-h', '1080'); await page.waitForTimeout(300);
    const none = await L.planInfo(page);
    const box = await text(page, '#resolve-box');
    check('F12 when nothing is left the page says why and offers ways out', none.counts.included === 0 && /No photos are left in the video/.test(box) && await page.isDisabled('#btn-export') && /Show them whole over blur/.test(box), box.slice(0, 110));
    await page.screenshot({ path: S + '/shots/F-none-left.png' });
    await page.click('#resolve-box button:has-text("Show them whole over blur")'); await page.waitForTimeout(300);
    const fixed = await L.planInfo(page);
    check('F13 one click resolves it', fixed.counts.included === 17 && fixed.segs.every(s => s.mode === 'blur'));
    check('F14 no console errors', logs.length === 0, logs.join(' | '));
    await browser.close();
  }

  /* ---------- G: ordering ---------- */
  if (want('G')) {
    const { browser, page, logs } = await L.open();
    await L.addPhotos(page, land.slice(0, 8));
    const o0 = (await L.planInfo(page)).order;
    // keyboard: select the third photo, move it two places later with Alt + arrow
    await page.click('#tray .thumb:nth-child(3)');
    await page.keyboard.press('Alt+ArrowRight'); await page.keyboard.press('Alt+ArrowRight'); await page.waitForTimeout(150);
    const o1 = (await L.planInfo(page)).order;
    const exp1 = o0.slice(); exp1.splice(4, 0, exp1.splice(2, 1)[0]);
    check('G1 keyboard reorder, Alt with arrow keys', JSON.stringify(o1) === JSON.stringify(exp1));
    // buttons
    await page.click('#mv-first'); await page.waitForTimeout(150);
    const o2 = (await L.planInfo(page)).order;
    check('G2 button reorder, move to the start', o2[0] === o0[2] && o2.length === 8);
    await page.click('#mv-next'); await page.waitForTimeout(150);
    check('G3 button reorder, later', (await L.planInfo(page)).order[1] === o0[2]);
    // drag and drop: drag the last photo onto the first
    const before = (await L.planInfo(page)).order;
    await page.dragAndDrop('#tray .thumb:nth-child(8)', '#tray .thumb:nth-child(1)', { targetPosition: { x: 4, y: 20 } });
    await page.waitForTimeout(200);
    const o3 = (await L.planInfo(page)).order;
    check('G4 drag to reorder', o3[0] === before[7] && o3.slice(1).join() === before.slice(0, 7).join(), o3.map(n => n.slice(4, 7)).join(','));
    // manual order kept through unrelated changes
    await page.click('#fps-seg button[data-v="50"]'); await page.uncheck('#loop'); await page.check('#loop'); await page.waitForTimeout(150);
    check('G5 manual order is preserved through other changes', (await L.planInfo(page)).order.join() === o3.join());
    check('G6 restore import order', await (async () => { await page.click('#btn-restore'); await page.waitForTimeout(150); return (await L.planInfo(page)).order.join() === o0.join(); })());
    // randomise twice gives two orders, each stable
    await page.click('#btn-random'); await page.waitForTimeout(100); const r1 = (await L.planInfo(page)).order.join();
    await page.click('#btn-play'); await page.waitForTimeout(700); await page.click('#btn-play');
    const r1b = (await L.planInfo(page)).order.join();
    const moves1 = await st(page, s => s.plan.segs.map(x => x.kind + x.a.x.toFixed(6)).join());
    await page.click('#btn-random'); await page.waitForTimeout(100); const r2 = (await L.planInfo(page)).order.join();
    check('G7 randomise order is explicit and stable until pressed again', r1 === r1b && r1 !== o0.join() && r2 !== r1);
    const seedBefore = await st(page, s => s.P.motion.seed);
    await page.click('#motion-reseed'); await page.waitForTimeout(100);
    const after = await st(page, s => ({ seed: s.P.motion.seed, order: s.plan.segs.map(x => x.name).join() }));
    check('G8 new motion variation leaves the photo order alone', after.seed !== seedBefore && after.order === r2);
    // remove selected, clear project needs two presses
    await page.click('#tray .thumb:nth-child(2)'); const victim = (await L.planInfo(page)).order[1];
    await page.click('#sel-remove'); await page.waitForTimeout(150);
    const afterRm = await L.planInfo(page);
    check('G9 remove selected photo', afterRm.counts.total === 7 && !afterRm.order.includes(victim));
    await page.click('#btn-clear'); await page.waitForTimeout(100);
    const armed = (await L.planInfo(page)).counts.total;
    await page.click('#btn-clear'); await page.waitForTimeout(200);
    check('G10 clear project asks once more, then clears', armed === 7 && (await L.planInfo(page)).counts.total === 0 && await page.isVisible('#drop'));
    check('G11 no console errors', logs.length === 0, logs.join(' | '));
    await browser.close();
  }

  /* ---------- H: timing through the interface ---------- */
  if (want('H')) {
    const { browser, page, logs } = await L.open();
    await L.addPhotos(page, land.slice(0, 7));
    await page.fill('#per-photo', '4'); await page.fill('#trans', '0.8'); await page.waitForTimeout(200);
    let plan = await L.planInfo(page);
    check('H1 per photo mode: 7 x 4 s at 30 fps', plan.frames === 840 && plan.segs.every(s => s.slot === 120 && s.trans === 24));
    await page.click('#timing-mode button[data-v="total"]'); await page.waitForTimeout(150);
    check('H2 switching to total keeps the length', (await L.planInfo(page)).frames === 840 && (await page.inputValue('#total')) === '28');
    await page.fill('#total', '1:30'); await page.waitForTimeout(200);
    plan = await L.planInfo(page);
    check('H3 total duration 1:30 gives exactly 2700 frames', plan.frames === 2700 && plan.segs.reduce((a, s) => a + s.slot, 0) === 2700 && (await text(page, '#dur-chip')) === '1:30.00');
    await page.click('#fps-seg button[data-v="25"]'); await page.fill('#total', '10.02'); await page.waitForTimeout(200);
    plan = await L.planInfo(page);
    check('H4 rounding to whole frames is explained', plan.frames === 251 && /250\.50 frames at 25 fps\. Using 251 frames/.test(await text(page, '#timing-msgs')));
    await page.fill('#total', '1'); await page.waitForTimeout(200);
    check('H5 an impossible total blocks export with the reason', /Too short for 7 photos/.test(await text(page, '#timing-msgs')) && await page.isDisabled('#btn-export') && /Too short for 7 photos/.test(await text(page, '#export-msgs')));
    await page.fill('#total', 'abc'); await page.waitForTimeout(200);
    check('H6 invalid input is flagged at the field, no dialog', (await page.getAttribute('#total', 'class')).includes('bad') && /Enter a total duration/.test(await text(page, '#timing-msgs')));
    // per photo override
    await page.click('#timing-mode button[data-v="perPhoto"]'); await page.click('#fps-seg button[data-v="30"]'); await page.fill('#per-photo', '4'); await page.waitForTimeout(150);
    await page.click('#tray .thumb:nth-child(2)');
    await page.evaluate(() => { document.getElementById('adv-photo').open = true; });
    await page.fill('#ph-duration', '10'); await page.waitForTimeout(200);
    plan = await L.planInfo(page);
    check('H7 a per photo duration changes the total', plan.frames === 6 * 120 + 300 && plan.segs[1].slot === 300);
    await page.selectOption('#ph-trans-type', 'cut'); await page.waitForTimeout(150);
    plan = await L.planInfo(page);
    check('H8 a per photo transition override', plan.segs[1].trans === 0 && plan.segs[2].transIn === 0 && plan.segs[0].trans === 24 && plan.frames === 6 * 120 + 300);
    await page.fill('#trans', '0'); await page.waitForTimeout(150);
    check('H9 zero second transition means cuts everywhere', (await L.planInfo(page)).segs.every(s => s.trans === 0));
    await page.fill('#trans', '9'); await page.waitForTimeout(150);
    plan = await L.planInfo(page);
    check('H10 a transition longer than half a photo is shortened and reported', plan.segs[0].trans === 60 && /crossfade was shortened/.test(await text(page, '#timing-msgs')) && plan.frames === 6 * 120 + 300);
    await page.click('#ph-reset'); await page.waitForTimeout(150);
    check('H11 reset to automatic', (await L.planInfo(page)).frames === 840 && (await page.inputValue('#ph-duration')) === '');
    // motion controls
    await page.selectOption('#ph-motion', 'pan-left'); await page.waitForTimeout(150);
    const k1 = await st(page, s => { const g = s.plan.segs[1]; return [g.kind, g.a.x < g.b.x]; });
    check('H12 per photo directional pan', k1[0] === 'pan' && k1[1] === true);
    await page.selectOption('#ph-motion', 'custom'); await page.waitForTimeout(150);
    await page.evaluate(() => { const z = document.getElementById('ed-zoom'); z.value = 160; z.dispatchEvent(new Event('input', { bubbles: true })); }); await page.waitForTimeout(150);
    await page.evaluate(() => document.getElementById('ed-view').scrollIntoView({ block: 'center' }));
    const box = await page.locator('#ed-view').boundingBox();
    await page.mouse.click(box.x + box.width * 0.3, box.y + box.height * 0.4); await page.waitForTimeout(150);
    const cst = await st(page, s => { const p = s.photos[1], g = s.plan.segs[1]; return { a: p.motion.a, kind: g.kind, aw: g.a.w }; });
    check('H13 visual editor sets a custom start frame', cst.kind === 'custom' && Math.abs(cst.a.z - 1.6) < 1e-9 && Math.abs(cst.a.cx - 0.3125) < 0.02 && Math.abs(cst.a.cy - 0.4) < 0.02 && Math.abs(cst.aw - 1.6) < 1e-9, JSON.stringify(cst.a));
    await page.screenshot({ path: S + '/shots/H-editor.png' });
    await page.uncheck('#motion-on'); await page.waitForTimeout(150);
    check('H14 motion off holds every photo still', await st(page, s => s.plan.segs.every(g => !g.moving)));
    check('H15 no console errors', logs.length === 0, logs.join(' | '));
    await browser.close();
  }

  /* ---------- I: unsupported settings are caught before export ---------- */
  if (want('I')) {
    const { browser, page, logs } = await L.open();
    await L.addPhotos(page, land.slice(0, 3));
    const fmt = await st(page, s => ({ format: s.P.exp.format, avc: s.probes.avc.supported, note: s.P.exp.autoNote }));
    if (fmt.avc) {
      check('I1 H.264 in MP4 is available here and is the default', fmt.format === 'avc/mp4' && !fmt.note);
    } else {
      const msg = await text(page, '#format-msgs');
      check('I1 this browser has no H.264 encoder, and the page says so instead of pretending', fmt.avc === false && fmt.format === 'vp9/webm' && /this browser cannot encode it here, so VP9 in WebM is selected instead/.test(msg));
      const opts = await page.evaluate(() => Array.from(document.getElementById('format').options).map(o => o.textContent + (o.disabled ? ' [disabled]' : '')));
      check('I2 unsupported formats are listed but cannot be picked', opts.some(o => /H\.264 in MP4 .*not available.*disabled/.test(o)) && opts.some(o => o === 'VP9 in WebM'), opts.join(' | '));
      // force the situation "selected format stopped being available"
      await L.tweak(page, P => { P.exp.format = 'avc/mp4'; P.exp.userPicked = true; });
      const blocked = await text(page, '#export-msgs');
      const stay = await st(page, s => s.P.exp.format);
      check('I3 an unavailable selection blocks export and is not swapped silently', await page.isDisabled('#btn-export') && stay === 'avc/mp4' && /H\.264 cannot be encoded by this browser at 1920 x 1080, 30 fps/.test(blocked) && /Nothing has been changed for you/.test(await text(page, '#format-msgs')));
      await page.click('#format-msgs button:has-text("Use VP9 in WebM")'); await page.waitForTimeout(300);
      check('I4 the offered alternative is one click', (await st(page, s => s.P.exp.format)) === 'vp9/webm' && !(await page.isDisabled('#btn-export')));
    }
    // odd size with H.264: specific reason and a suggested even size (uses the pure rule, shown before any probe)
    await L.tweak(page, P => { P.exp.format = 'avc/mp4'; P.output.w = 1365; P.output.h = 767; });
    const odd = await text(page, '#format-msgs');
    check('I5 odd size with H.264 gives the exact reason and offers the even size, without resizing by itself', /needs an even width and height\. 1365 x 767 has an odd side/.test(odd) && /Use 1366 x 768/.test(odd) && (await st(page, s => s.P.output.w)) === 1365);
    // invalid size
    await page.fill('#out-w', '0'); await page.waitForTimeout(200);
    check('I6 invalid size is flagged at the field and blocks export', /at least 16 px/.test(await text(page, '#size-msgs')) && await page.isDisabled('#btn-export'));
    await page.fill('#out-w', '1920*2'); await page.fill('#out-h', '1080'); await page.waitForTimeout(250);
    check('I7 size fields accept the same sums as the calculator', (await st(page, s => s.P.output.w)) === 3840 && (await text(page, '#aspect-chip')).startsWith('32:9'));
    // aspect lock
    await page.fill('#out-w', '1920'); await page.fill('#out-h', '1080'); await page.click('#lock'); await page.fill('#out-w', '2560'); await page.waitForTimeout(200);
    check('I8 aspect lock moves the other side and shows it', (await page.inputValue('#out-h')) === '1440' && (await page.inputValue('#preset')) === '1440p');
    // demanding settings: warning, not a block
    await L.tweak(page, P => { P.exp.format = 'vp9/webm'; P.output.w = 3840; P.output.h = 2160; P.output.fps = 60; P.timing.mode = 'total'; P.timing.total = 900; });
    const warn = await text(page, '#export-msgs');
    check('I9 heavy settings raise a practical warning but stay exportable', /This export is demanding, 3840 x 2160 at 60 fps/.test(warn) && /not a prediction/.test(warn) && !(await page.isDisabled('#btn-export')), warn.slice(0, 160));
    await page.screenshot({ path: S + '/shots/I-warning.png' });
    // advanced validation
    await page.evaluate(() => { document.getElementById('adv-export').open = true; });
    await page.fill('#bitrate', '9999'); await page.waitForTimeout(200);
    check('I10 invalid bitrate is flagged and blocks export', /between 0\.5 and 800/.test(await text(page, '#adv-export-msgs')) && await page.isDisabled('#btn-export'));
    await page.fill('#bitrate', '25'); await page.waitForTimeout(400);
    check('I11 a bitrate set by hand is used', (await st(page, s => s.probes.vp9.config.bitrate)) === 25e6 && /25 Mbit\/s set by hand/.test(await text(page, '#format-msgs')));
    check('I12 no console errors', logs.length === 0, logs.join(' | '));
    await browser.close();
  }

  /* ---------- J: cancel, then export again ---------- */
  if (want('J')) {
    const { browser, page, logs } = await L.open();
    await L.addPhotos(page, land.slice(0, 6));
    await L.tweak(page, P => { P.output.w = 1920; P.output.h = 1080; P.timing.perPhoto = 6; });
    const before = await L.planInfo(page);
    await page.click('#btn-export');
    await page.waitForFunction(() => /frame \d+ of/.test(document.getElementById('progress-detail').textContent), null, { timeout: 60000 });
    await page.waitForTimeout(1200);
    const mid = await text(page, '#progress-detail');
    const lockedDuring = await page.evaluate(() => document.getElementById('settings').disabled && document.getElementById('btn-add').disabled);
    await page.screenshot({ path: S + '/shots/J-exporting.png' });
    const tC = Date.now();
    await page.click('#btn-cancel');
    await page.waitForFunction(() => !SlideshowApp.state().ui.exporting, null, { timeout: 15000 });
    const dt = Date.now() - tC;
    const after = await L.planInfo(page);
    const msg = await text(page, '#result');
    check('J1 progress shows real frame counts while exporting, settings locked', /frame [\d,]+ of 1,080/.test(mid) && lockedDuring, mid);
    check('J2 cancel stops promptly and keeps the project', dt < 6000 && /Export cancelled/.test(msg) && after.counts.total === 6 && after.order.join() === before.order.join() && !(await page.isDisabled('#btn-export')), `stopped in ${dt} ms`);
    check('J3 nothing is left behind after cancel', !(await page.$('#download-link')) && !(await page.evaluate(() => document.getElementById('settings').disabled)));
    await L.tweak(page, P => { P.timing.perPhoto = 1; P.output.w = 1280; P.output.h = 720; });
    const out = S + '/out/J_after_cancel.webm'; await L.exportTo(page, out);
    check('J4 a new export after cancelling works', A.probe(out).frames === 180);
    check('J5 no console errors', logs.length === 0, logs.join(' | '));
    await browser.close();
  }

  /* ---------- K: export on the page itself when a worker cannot be used ---------- */
  if (want('K')) {
    const { browser, page, logs } = await L.open({ query: '?export=page' });
    await L.addPhotos(page, land.slice(0, 3));
    await L.tweak(page, P => { P.output.w = 1280; P.output.h = 720; P.timing.perPhoto = 1; });
    const out = S + '/out/K_page.webm'; const info = await L.exportTo(page, out);
    const pr = A.probe(out);
    check('K1 fallback export without a worker', info.mode === 'page' && pr.frames === 90 && pr.matrix === 'bt709', info.mode + ' ' + pr.frames);
    check('K2 no console errors', logs.length === 0, logs.join(' | '));
    await browser.close();
  }

  /* ---------- M: straight to disk ---------- */
  if (want('M')) {
    const { browser, page, logs } = await L.open();
    // stand in for the save dialog with a real file handle from the origin private file system
    await page.evaluate(() => { window.showSaveFilePicker = async (o) => { const root = await navigator.storage.getDirectory(); window.__picked = o; return root.getFileHandle(o.suggestedName, { create: true }); }; });
    await L.addPhotos(page, land.slice(0, 4));
    for (const fmt of ['av1/mp4', 'vp9/webm']) {
      await L.tweak(page, (P, ph, f) => { P.output.w = 1280; P.output.h = 720; P.timing.perPhoto = 1; P.exp.toDisk = 'on'; P.exp.format = f; P.exp.userPicked = true; }, fmt);
      const willAsk = /asked where to save/.test(await text(page, '#export-msgs'));
      await page.click('#btn-export');
      await page.waitForFunction(() => SlideshowApp.state().ui.exporting, null, { timeout: 10000 });
      await page.waitForFunction(() => !SlideshowApp.state().ui.exporting && /Export complete|Export failed/.test(document.getElementById('result').textContent), null, { timeout: 120000 });
      await page.waitForFunction(() => { const d = document.getElementById('checks'); return d && /opens here/i.test(d.textContent); }, null, { timeout: 30000 });
      const res = await text(page, '#result');
      const ext = fmt.split('/')[1];
      const b64 = await page.evaluate(async (name) => { const root = await navigator.storage.getDirectory(); const f = await (await root.getFileHandle(name)).getFile(); const u = new Uint8Array(await f.arrayBuffer()); let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }, `slideshow_1280x720_30fps_loop.${ext}`);
      const out = S + '/out/M_disk.' + ext; fs.writeFileSync(out, Buffer.from(b64, 'base64'));
      const pr = A.probe(out);
      const moovFirst = ext !== 'mp4' || fs.readFileSync(out).subarray(0, 4096).includes(Buffer.from('moov'));
      check('M straight to disk, ' + fmt, willAsk && /Saved where you chose/.test(res) && !/does not match/.test(res) && pr.frames === 120 && Math.abs(pr.duration - 4) < 0.003 && moovFirst, `${pr.codec} ${pr.frames} frames ${pr.duration}s index first ${moovFirst} | ` + res.replace(/\s+/g, ' ').slice(0, 150));
    }
    // cancelling the save dialog starts nothing
    await page.evaluate(() => { window.showSaveFilePicker = async () => { const e = new Error('closed'); e.name = 'AbortError'; throw e; }; });
    await page.click('#btn-export'); await page.waitForTimeout(400);
    check('M closing the save dialog starts nothing', !(await st(page, s => s.ui.exporting)));
    check('M no console errors', logs.length === 0, logs.join(' | '));
    await browser.close();
  }

  /* ---------- N: navigation ---------- */
  if (want('N')) {
    const { browser, page, logs } = await L.open();
    await page.goto(L.ROOT + '/index.html');
    const link = page.locator('#slideshow-link');
    check('N1 homepage has a clearly labelled entry', /Slideshow Builder/.test(await link.textContent()) && (await link.getAttribute('href')) === '/slideshow.html');
    await page.screenshot({ path: S + '/shots/N-home.png', fullPage: true });
    await link.click(); await page.waitForFunction(() => window.SlideshowApp);
    check('N2 the entry opens the builder', /\/slideshow\.html$/.test(page.url()) && (await page.title()).startsWith('Slideshow Builder'));
    await page.click('header .header-by a'); await page.waitForLoadState();
    check('N3 the builder links straight back to SlideSize', page.url() === L.ROOT + '/' && /Slide Size Calculator/.test(await page.title()), page.url());
    await page.fill('#px-w', '3840'); await page.fill('#px-h', '1080'); await page.selectOption('#fps', '50'); await page.click('.calc-btn');
    await page.click('#slideshow-link'); await page.waitForFunction(() => window.SlideshowApp && !SlideshowApp.state().probing);
    const got = await page.evaluate(() => { const o = SlideshowApp.state().P.output; return [o.w, o.h, o.fps, document.getElementById('out-w').value, document.getElementById('preset').value]; });
    check('N4 the calculator hands its size and frame rate over', got.join() === '3840,1080,50,3840,dw1080', page.url() + ' ' + got.join());
    await page.goto(L.ROOT + '/slideshow.html?w=1024&h=768&fps=23.976'); await page.waitForFunction(() => window.SlideshowApp);
    const g2 = await page.evaluate(() => { const o = SlideshowApp.state().P.output; return [o.w, o.h, o.fps]; });
    check('N5 direct links with a size work, unsupported rates fall back to the default', g2.join() === '1024,768,30');
    await page.goto(L.ROOT + '/slideshow.html?w=abc&h=-1'); await page.waitForFunction(() => window.SlideshowApp);
    check('N6 a bad link is ignored, defaults used', (await page.evaluate(() => { const o = SlideshowApp.state().P.output; return [o.w, o.h, o.fps].join(); })) === '1920,1080,30');
    check('N7 no console errors', logs.length === 0, logs.join(' | '));
    await browser.close();
  }

  const failed = results.filter(r => !r.ok);
  console.log(`\n${results.length - failed.length} of ${results.length} checks passed`);
  if (failed.length) process.exit(1);
})().catch(e => { console.error('FAIL', e); process.exit(1); });
