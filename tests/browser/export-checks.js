const L = require('./lib.js'), A = require('./analyse.js'); const S = process.env.WORK; const fs = require('fs');
const results = [];
function check(name, ok, detail) { results.push({ name, ok: !!ok, detail: detail || '' }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  [' + detail + ']' : '')); }
const fx = v => (typeof v === 'number' ? v.toFixed(2) : v);
async function previewRgb(page, k, w, h) {
  await page.evaluate(async (k) => { SlideshowApp.seek(k); for (let i = 0; i < 400 && !SlideshowApp.frameReady(k); i++) await new Promise(r => setTimeout(r, 15)); SlideshowApp.draw(); }, k);
  const url = await page.evaluate(() => document.getElementById('preview').toDataURL('image/png'));
  return A.pngToRgb(Buffer.from(url.split(',')[1], 'base64'), w, h);
}
async function setFormat(page, f) { await L.tweak(page, (P, ph, f) => { P.exp.format = f; P.exp.userPicked = true; }, f); }
const only = process.argv[2] ? new RegExp(process.argv[2]) : null;
const want = n => !only || only.test(n);

(async () => {
  const land = L.listPhotos(S + '/photos/set40', f => /land/.test(f));
  const all40 = L.listPhotos(S + '/photos/set40');
  const mixed = L.listPhotos(S + '/photos/set12');

  /* ---------- A: 40 photos, randomised, loop: import, preview, export, metadata, seam ---------- */
  if (want('A')) {
    const { browser, page, logs } = await L.open();
    const t0 = Date.now();
    await L.addPhotos(page, all40);
    let plan = await L.planInfo(page);
    check('A1 40 photos import', plan.counts.total === 40, `total ${plan.counts.total}, in ${plan.counts.included}, out ${plan.counts.excluded}, ${Date.now() - t0} ms`);
    check('A2 landscape output excludes portrait and square by default', plan.excluded.length > 0 && plan.excluded.every(e => /Portrait photo|Square photo|much wider/.test(e.reason)) && !plan.order.some(n => /port|square/.test(n)), plan.excluded.length + ' excluded');
    await L.tweak(page, P => { P.output.w = 1280; P.output.h = 720; P.timing.perPhoto = 1.5; P.timing.transition = 0.5; });
    const before = (await L.planInfo(page)).order;
    await page.click('#btn-random');
    plan = await L.planInfo(page);
    check('A3 randomise changes the order', JSON.stringify(plan.order) !== JSON.stringify(before) && plan.order.slice().sort().join() === before.slice().sort().join());
    const kinds1 = plan.segs.map(s => s.kind).join();
    // preview stability: same frame twice, with playback and a settings round trip in between
    const N = plan.frames, probeFrames = [0, 37, Math.floor(N / 2) + 3, N - 5];
    const snap1 = []; for (const k of probeFrames) snap1.push(await previewRgb(page, k, 320, 180));
    await page.click('#btn-play'); await page.waitForTimeout(1200); await page.click('#btn-play');
    await L.tweak(page, P => { P.exp.quality = 'standard'; }); await L.tweak(page, P => { P.exp.quality = 'high'; P.output.fps = 25; }); await L.tweak(page, P => { P.output.fps = 30; });
    const plan2 = await L.planInfo(page);
    const snap2 = []; for (const k of probeFrames) snap2.push(await previewRgb(page, k, 320, 180));
    check('A4 order and moves are unchanged by replaying or by unrelated settings', JSON.stringify(plan2.order) === JSON.stringify(plan.order) && plan2.segs.map(s => s.kind).join() === kinds1 && snap1.every((b, i) => A.mad(b, snap2[i]) === 0));
    // lossless look at the loop point in the preview renderer (1280 x 720 is full size here)
    const ring = []; for (const k of [N - 4, N - 3, N - 2, N - 1, 0, 1, 2, 3]) ring.push(await previewRgb(page, k, 640, 360));
    const steps = ring.slice(1).map((f, i) => A.mad(ring[i], f));
    check('A5 rendered frames either side of the loop point are neighbours, not copies and not a jump', steps[3] > 0.01 && steps[3] < Math.max(steps[2], steps[4]) * 2.5 + 0.2 && steps[3] < 3, 'steps ' + steps.map(fx).join(' ') + ' (wrap is the 4th)');
    // export twice
    const out1 = S + '/out/A_loop_1.webm', out2 = S + '/out/A_loop_2.webm';
    const tE = Date.now(); const info = await L.exportTo(page, out1); const took = Date.now() - tE;
    await L.exportTo(page, out2);
    const pr = A.probe(out1);
    check('A6 export of a typical project completes in a worker', info.mode === 'worker' && pr.frames === N, `${N} frames in ${took} ms, ${(N / took * 1000).toFixed(0)} fps, ${(pr.size / 1e6).toFixed(1)} MB`);
    check('A7 metadata matches the settings', pr.codec === 'vp9' && pr.w === 1280 && pr.h === 720 && pr.rate === '30/1' && pr.avg === '30/1' && Math.abs(pr.duration - N / 30) < 0.002 && /webm/.test(pr.format), JSON.stringify(pr));
    check('A8 colour is BT.709, video range, for an HD picture', pr.matrix === 'bt709' && pr.primaries === 'bt709' && pr.range === 'tv' && pr.pix === 'yuv420p');
    const f1 = A.rgbFrames(out1, 320, 180), f2 = A.rgbFrames(out2, 320, 180);
    let worst = 0; for (let k = 0; k < f1.length; k += 9) worst = Math.max(worst, A.mad(f1[k], f2[k]));
    check('A9 exporting again gives the same slideshow', f1.length === f2.length && worst < 1.0, 'largest frame difference between the two exports ' + fx(worst));
    // preview vs export at matching frames
    let pw = 0; const cmp = [];
    for (const k of [0, 20, 44, 100, Math.floor(N / 2), N - 40, N - 8, N - 1]) { const p = await previewRgb(page, k, 320, 180); const d = A.mad(p, f1[k]); cmp.push(k + ':' + fx(d)); pw = Math.max(pw, d); }
    check('A10 preview and export show the same picture at the same frame', pw < 4, 'mean difference per channel, 0 to 255: ' + cmp.join(' '));
    // seam in the encoded file
    const d = A.stepDiffs(f1), hold = A.median(d), fadePeak = Math.max(...d);
    const kf = []; for (let k = 30; k < N; k += 30) kf.push(d[k - 1]);
    check('A11 encoded loop point: no cut, no flash, no repeated frame', d[N - 1] > 0.05 && d[N - 1] < fadePeak && d[N - 1] < 8, `wrap step ${fx(d[N - 1])}, typical step ${fx(hold)}, largest mid crossfade step ${fx(fadePeak)}, other key frame steps up to ${fx(Math.max(...kf))}`);
    // what a hard cut would measure, for scale
    const cutSize = A.mad(f1[10], f1[10 + 45]);
    check('A12 for scale, two different photos differ far more than the loop step', cutSize > d[N - 1] * 5, `different photos ${fx(cutSize)} vs wrap ${fx(d[N - 1])}`);
    check('A13 no console errors', logs.length === 0, logs.join(' | '));
    fs.writeFileSync(S + '/out/A_steps.json', JSON.stringify({ N, d: d.map(v => +v.toFixed(3)) }));
    await browser.close();
  }

  /* ---------- B: one photo loop ---------- */
  if (want('B')) {
    const { browser, page, logs } = await L.open();
    await L.addPhotos(page, [land[0]]);
    await L.tweak(page, P => { P.output.w = 1280; P.output.h = 720; P.timing.perPhoto = 6; P.motion.intensity = 'moderate'; });
    const plan = await L.planInfo(page), N = plan.frames;
    const out = S + '/out/B_one.webm'; await L.exportTo(page, out);
    const f = A.rgbFrames(out, 640, 360), d = A.stepDiffs(f);
    const far = A.mad(f[0], f[Math.floor(N / 2)]);
    const interior = d.slice(0, N - 1);
    check('B1 one photo loop: 180 frames, real motion', f.length === 180 && far > 1.5, `frame 0 vs halfway ${fx(far)}`);
    check('B2 one photo loop: no reset at the loop point', d[N - 1] < Math.max(...interior) * 1.3 + 0.2 && d[N - 1] < far * 0.25, `wrap step ${fx(d[N - 1])}, largest step elsewhere ${fx(Math.max(...interior))}, typical ${fx(A.median(d))}`);
    const ring = []; for (const k of [N - 3, N - 2, N - 1, 0, 1, 2]) ring.push(await previewRgb(page, k, 1280, 720));
    const st = ring.slice(1).map((x, i) => A.mad(ring[i], x));
    check('B3 one photo loop, lossless render: the turn is smooth and frames differ', st.every(v => v > 0) && st[2] <= Math.max(...st) + 1e-9 && Math.max(...st) < 0.6, 'steps ' + st.map(v => v.toFixed(4)).join(' '));
    check('B4 no console errors', logs.length === 0, logs.join(' | '));
    await browser.close();
  }

  /* ---------- C: two photo loop, then the same pair not looping ---------- */
  if (want('C')) {
    const { browser, page, logs } = await L.open();
    await L.addPhotos(page, [land[0], land[1]]);
    await L.tweak(page, P => { P.output.w = 1280; P.output.h = 720; P.timing.perPhoto = 3; P.timing.transition = 1; });
    let plan = await L.planInfo(page), N = plan.frames;
    check('C1 two photo loop plan has both transitions', plan.segs.length === 2 && plan.segs.every(s => s.trans === 30 && s.transIn === 30) && N === 180);
    const out = S + '/out/C_two_loop.webm'; await L.exportTo(page, out);
    const f = A.rgbFrames(out, 320, 180), d = A.stepDiffs(f);
    const fade1 = d.slice(60, 90), fade2 = d.slice(150, 180);
    const m1 = Math.max(...fade1), m2 = Math.max(...fade2);
    check('C2 closing transition is a real crossfade like the first one', Math.abs(m1 - m2) / m1 < 0.35 && m2 > A.median(d) * 2, `peak step first fade ${fx(m1)}, closing fade ${fx(m2)}`);
    check('C3 loop point is continuous', d[N - 1] > 0.02 && d[N - 1] < m2 * 0.8, `wrap step ${fx(d[N - 1])}`);
    check('C4 frame 0 is the first photo, last frame is nearly the first photo too', A.mad(f[0], f[N - 1]) < 4 && A.mad(f[0], f[100]) > 15, `first vs last ${fx(A.mad(f[0], f[N - 1]))}, first vs second photo ${fx(A.mad(f[0], f[100]))}`);
    // not looping
    await page.uncheck('#loop');
    plan = await L.planInfo(page);
    check('C5 loop off: no closing transition in the plan, same length', plan.segs[1].trans === 0 && plan.segs[0].transIn === 0 && plan.frames === 180);
    const out2 = S + '/out/C_two_once.webm'; await L.exportTo(page, out2);
    const g = A.rgbFrames(out2, 320, 180), e = A.stepDiffs(g);
    const tail = e.slice(150, 179);
    check('C6 loop off: the video ends on the last photo with no fade', Math.max(...tail) < m1 * 0.4 && A.mad(g[179], g[165]) < 6 && A.mad(g[179], g[0]) > 15, `largest step in the last second ${fx(Math.max(...tail))}, last vs first frame ${fx(A.mad(g[179], g[0]))}`);
    const p0 = await previewRgb(page, 0, 320, 180), pl = await previewRgb(page, 179, 320, 180);
    check('C7 loop off: first and last frames match the preview', A.mad(p0, g[0]) < 4 && A.mad(pl, g[179]) < 4, `${fx(A.mad(p0, g[0]))} ${fx(A.mad(pl, g[179]))}`);
    check('C8 loop off: file name says so', (await page.evaluate(() => document.getElementById('download-link').download)) === 'slideshow_1280x720_30fps.webm');
    check('C9 no console errors', logs.length === 0, logs.join(' | '));
    await browser.close();
  }

  /* ---------- D: sizes, rates and formats ---------- */
  if (want('D')) {
    const { browser, page, logs } = await L.open();
    await L.addPhotos(page, land.slice(0, 3));
    await L.tweak(page, P => { P.timing.perPhoto = 1; P.timing.transition = 0.4; });
    for (const [w, h, fps, fmt, tag] of [[1366, 768, 60, 'av1/mp4', 'custom size at 60 fps, AV1 in MP4'], [1365, 767, 50, 'vp9/webm', 'odd size at 50 fps, VP9'], [3840, 1080, 25, 'vp9/webm', 'double wide at 25 fps'], [1080, 1920, 24, 'av1/webm', 'portrait at 24 fps, AV1 in WebM'], [640, 480, 30, 'av1/mp4', 'small 4:3']]) {
      await page.fill('#out-w', String(w)); await page.fill('#out-h', String(h));
      await page.click(`#fps-seg button[data-v="${fps}"]`);
      await L.tweak(page, (P) => { P.framing.exclude = false; P.framing.mode = 'fill'; });
      await setFormat(page, fmt);
      const plan = await L.planInfo(page);
      const out = S + `/out/D_${w}x${h}_${fps}.${fmt.split('/')[1]}`;
      const info = await L.exportTo(page, out);
      const pr = A.probe(out);
      const hd = w >= 1280 || h > 576;
      check('D ' + tag, pr.w === w && pr.h === h && pr.rate === fps + '/1' && pr.frames === plan.frames && plan.frames === 3 * fps && Math.abs(pr.duration - 3) < 0.003 && pr.matrix === (hd ? 'bt709' : 'smpte170m') && pr.codec === fmt.split('/')[0],
        `${pr.codec} ${pr.w}x${pr.h} ${pr.rate} ${pr.frames} frames ${pr.duration}s ${pr.matrix}` + (/does not match/.test(info.result) ? ' PAGE CHECK FAILED: ' + info.result.replace(/\n/g, ' / ') : ''));
      if (/does not match/.test(info.result)) check('D page self check ' + tag, false, info.result.replace(/\n/g, ' / '));
    }
    check('D no console errors', logs.length === 0, logs.join(' | '));
    await browser.close();
  }

  /* ---------- E: fit modes, alpha, EXIF, blur ---------- */
  if (want('E')) {
    const { browser, page, logs } = await L.open();
    await L.addPhotos(page, mixed.filter(f => /alpha|exif|IMG_001|IMG_004|IMG_006|IMG_007|webp/.test(f)));
    await L.tweak(page, P => { P.output.w = 1280; P.output.h = 720; P.timing.perPhoto = 1.2; P.timing.transition = 0.6; P.framing.exclude = false; P.framing.mode = 'blur'; });
    let plan = await L.planInfo(page);
    check('E1 filter off, blurred background: every photo is in', plan.counts.included === 7 && plan.segs.every(s => s.mode === 'blur'));
    const out = S + '/out/E_blur.webm'; await L.exportTo(page, out);
    const f = A.rgbFrames(out, 320, 180);
    let pw = 0; for (let k = 5; k < plan.frames; k += 17) pw = Math.max(pw, A.mad(await previewRgb(page, k, 320, 180), f[k]));
    check('E2 blurred background: preview and export agree, mid transition included', pw < 4.5, 'largest difference ' + fx(pw));
    await page.screenshot({ path: S + '/shots/E-blur.png' });
    await L.tweak(page, P => { P.framing.mode = 'fit'; P.framing.bg = '#203040'; });
    const out2 = S + '/out/E_fit.webm'; await L.exportTo(page, out2);
    const g = A.rgbFrames(out2, 320, 180);
    plan = await L.planInfo(page);
    const portrait = plan.segs.find(s => /port|exif/.test(s.name));
    const fr = g[portrait.start + 10];
    const corner = [fr[0], fr[1], fr[2]];
    check('E3 fit mode: bars are the chosen background colour', Math.abs(corner[0] - 0x20) < 6 && Math.abs(corner[1] - 0x30) < 6 && Math.abs(corner[2] - 0x40) < 6, 'top left pixel ' + corner.join(','));
    pw = 0; for (let k = 5; k < plan.frames; k += 17) pw = Math.max(pw, A.mad(await previewRgb(page, k, 320, 180), g[k]));
    check('E4 fit mode: preview and export agree', pw < 4.5, 'largest difference ' + fx(pw));
    check('E5 no console errors', logs.length === 0, logs.join(' | '));
    await browser.close();
  }

  const failed = results.filter(r => !r.ok);
  console.log(`\n${results.length - failed.length} of ${results.length} checks passed`);
  if (failed.length) process.exit(1);
})().catch(e => { console.error('FAIL', e); process.exit(1); });
