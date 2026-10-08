/* Photo Montage tests. No dependencies.
   Run from the repository root with:  node --test                     */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const Core = require('../slideshow-core.js');
const Render = require('../slideshow-render.js');
const Exp = require('../slideshow-export.js');

/* ---------- helpers ---------- */

const SHAPES = { land32: [6000, 4000], land43: [4000, 3000], land169: [3840, 2160], land1610: [1920, 1200],
                 port23: [4000, 6000], port34: [3000, 4000], square: [3000, 3000], pano: [9000, 3000], land54: [2500, 2000] };

function photo(id, shape, extra) {
  const s = Array.isArray(shape) ? shape : SHAPES[shape];
  return Object.assign({ id, name: id + '.jpg', iw: s[0], ih: s[1], opaque: true, include: 'auto', framing: null,
                         motion: { kind: 'auto' }, focus: null, duration: null, transition: null }, extra || {});
}
function project(photos, o) {
  o = o || {};
  return {
    output: Object.assign({ w: 1920, h: 1080, fps: 30 }, o.output),
    loop: o.loop === undefined ? true : o.loop,
    timing: Object.assign({ mode: 'perPhoto', perPhoto: 5, total: 60, transition: 1, transitionType: 'crossfade' }, o.timing),
    motion: Object.assign({ enabled: true, style: 'mixed', intensity: 'gentle', variation: true, seed: 12345 }, o.motion),
    framing: Object.assign({ exclude: true, threshold: 0.55, mode: 'fill', bg: '#000000' }, o.framing),
    photos
  };
}
function landscapes(n, shape) { return Array.from({ length: n }, (_, i) => photo('p' + i, shape || ['land32', 'land43', 'land169', 'land1610'][i % 4])); }
function included(plan) { return plan.segs.map(s => s.id); }
function sumSlots(plan) { return plan.segs.reduce((a, s) => a + s.slot, 0); }
function rectDelta(a, b) { return Math.abs(a.x - b.x) + Math.abs(a.y - b.y) + Math.abs(a.w - b.w) + Math.abs(a.h - b.h); }
function layerOf(layers, id) { return layers.find(l => l.seg.id === id); }

/* A stand-in 2D context that records what was drawn. */
function fakeCtx() {
  const calls = [];
  const ctx = {
    calls, globalAlpha: 1, fillStyle: '', canvas: { tag: 'canvas' },
    setTransform() {}, fillRect(x, y, w, h) { calls.push({ op: 'fill', style: ctx.fillStyle, alpha: ctx.globalAlpha, x, y, w, h }); },
    drawImage(img, x, y, w, h) { calls.push({ op: 'draw', img, alpha: ctx.globalAlpha, x, y, w, h }); }
  };
  return ctx;
}
function renderCalls(plan, frame, W, H) {
  const ctx = fakeCtx(), scratch = fakeCtx();
  scratch.canvas = { tag: 'scratch' };
  const missing = Render.renderFrame(ctx, W, H, plan, frame, seg => ({ bmp: { id: seg.id }, blur: { blurOf: seg.id } }), () => scratch);
  return { main: ctx.calls, scratch: scratch.calls, missing };
}

/* ================================================================== *
 * Timing                                                              *
 * ================================================================== */

test('per photo mode: length is the photo times added up, crossfade included', () => {
  const plan = Core.buildPlan(project(landscapes(12)));
  assert.equal(plan.ok, true);
  assert.equal(plan.frames, 12 * 150);
  assert.equal(plan.seconds, 60);
  assert.equal(sumSlots(plan), plan.frames);
  plan.segs.forEach(s => { assert.equal(s.slot, 150); assert.equal(s.trans, 30); assert.equal(s.transIn, 30); assert.equal(s.life, 180); });
});

test('transition length never changes the total', () => {
  const a = Core.buildPlan(project(landscapes(9), { timing: { transition: 0 } }));
  const b = Core.buildPlan(project(landscapes(9), { timing: { transition: 1 } }));
  const c = Core.buildPlan(project(landscapes(9), { timing: { transition: 2.4 } }));
  assert.equal(a.frames, 9 * 150);
  assert.equal(b.frames, a.frames);
  assert.equal(c.frames, a.frames);
});

test('total duration mode: exact frame count at every frame rate', () => {
  for (const fps of Core.FPS_CHOICES) {
    for (const n of [1, 2, 7, 13, 40, 100]) {
      const plan = Core.buildPlan(project(landscapes(n), { output: { fps }, timing: { mode: 'total', total: 90 } }));
      assert.equal(plan.ok, true, `fps ${fps} n ${n}`);
      assert.equal(plan.frames, 90 * fps);
      assert.equal(sumSlots(plan), plan.frames);
      const slots = plan.segs.map(s => s.slot);
      assert.ok(Math.max(...slots) - Math.min(...slots) <= 1, 'slots differ by at most one frame');
    }
  }
});

test('total duration mode: closing transition is inside the total when looping', () => {
  const plan = Core.buildPlan(project(landscapes(6), { timing: { mode: 'total', total: 30 } }));
  assert.equal(plan.frames, 900);
  const last = plan.segs[5];
  assert.equal(last.trans, 30, 'last photo dissolves into the first');
  assert.equal(last.start + last.slot, plan.frames, 'and that dissolve ends exactly at the end of the file');
});

test('durations that are not a whole number of frames are rounded and explained', () => {
  const total = Core.buildPlan(project(landscapes(4), { output: { fps: 25 }, timing: { mode: 'total', total: 10.02 } }));
  assert.equal(total.frames, 251);
  assert.ok(total.notes.some(n => /250\.50 frames/.test(n) && /251 frames/.test(n)), total.notes.join(' | '));
  const per = Core.buildPlan(project(landscapes(4), { output: { fps: 24 }, timing: { perPhoto: 2.55 } }));
  assert.equal(per.segs[0].slot, 61);
  assert.ok(per.notes.some(n => /61\.20 frames/.test(n)), per.notes.join(' | '));
  const exact = Core.buildPlan(project(landscapes(4), { output: { fps: 24 }, timing: { perPhoto: 5 } }));
  assert.equal(exact.notes.length, 0);
});

test('an uneven split is reported', () => {
  const plan = Core.buildPlan(project(landscapes(7), { timing: { mode: 'total', total: 10 } }));
  assert.equal(plan.frames, 300);
  assert.ok(plan.notes.some(n => /does not divide evenly/.test(n)));
});

test('per photo overrides change the calculated total', () => {
  const ps = landscapes(5);
  ps[2].duration = 12;
  const plan = Core.buildPlan(project(ps));
  assert.equal(plan.frames, 4 * 150 + 360);
  assert.equal(plan.segs[2].slot, 360);
  assert.equal(plan.segs[3].start, 150 + 150 + 360);
});

test('in total mode an override is fixed and the others share what is left', () => {
  const ps = landscapes(5);
  ps[0].duration = 20;
  const plan = Core.buildPlan(project(ps, { timing: { mode: 'total', total: 60 } }));
  assert.equal(plan.frames, 1800);
  assert.equal(plan.segs[0].slot, 600);
  plan.segs.slice(1).forEach(s => assert.equal(s.slot, 300));
});

test('a transition cannot take more than half of a photo and says so', () => {
  const plan = Core.buildPlan(project(landscapes(6), { timing: { perPhoto: 1, transition: 2 } }));
  assert.equal(plan.ok, true);
  plan.segs.forEach(s => { assert.equal(s.slot, 30); assert.equal(s.trans, 15); });
  assert.ok(plan.notes.some(n => /crossfade was shortened on every photo/.test(n)));
  assert.equal(plan.frames, 180);
});

test('a zero length transition is a cut', () => {
  const plan = Core.buildPlan(project(landscapes(4), { timing: { transition: 0 } }));
  plan.segs.forEach(s => { assert.equal(s.trans, 0); assert.equal(s.transType, 'cut'); });
  for (let k = 0; k < plan.frames; k++) assert.equal(Core.stateAt(plan, k).length, 1);
  const cut = Core.buildPlan(project(landscapes(4), { timing: { transitionType: 'cut', transition: 1 } }));
  cut.segs.forEach(s => assert.equal(s.trans, 0));
});

test('per photo transition override', () => {
  const ps = landscapes(4);
  ps[1].transition = { type: 'cut', seconds: null };
  ps[2].transition = { type: null, seconds: 2 };
  const plan = Core.buildPlan(project(ps));
  assert.deepEqual(plan.segs.map(s => s.trans), [30, 0, 60, 30]);
  assert.equal(plan.segs[2].transIn, 0, 'photo 3 arrives by a cut');
  assert.equal(plan.segs[3].transIn, 60);
  assert.equal(plan.frames, 600);
});

test('impossible and very short totals are rejected with a reason', () => {
  const tooShort = Core.buildPlan(project(landscapes(50), { timing: { mode: 'total', total: 5 } }));
  assert.equal(tooShort.ok, false);
  assert.ok(tooShort.errors.some(e => e.field === 'total' && /Too short for 50 photos/.test(e.msg)));
  const zero = Core.buildPlan(project(landscapes(3), { timing: { mode: 'total', total: 0 } }));
  assert.equal(zero.ok, false);
  const nan = Core.buildPlan(project(landscapes(3), { timing: { perPhoto: NaN } }));
  assert.equal(nan.ok, false);
  assert.ok(nan.errors.some(e => e.field === 'perPhoto'));
  const fast = Core.buildPlan(project(landscapes(3), { timing: { perPhoto: 1 } }));
  assert.equal(fast.ok, true);
  assert.ok(fast.warnings.some(w => /quick for a large screen/.test(w)));
});

test('timing follows photos being excluded, removed, added and reordered', () => {
  const ps = landscapes(8);
  const base = Core.buildPlan(project(ps));
  assert.equal(base.frames, 1200);
  ps[3].include = 'out';
  assert.equal(Core.buildPlan(project(ps)).frames, 1050);
  const fewer = ps.filter((_, i) => i !== 0);
  assert.equal(Core.buildPlan(project(fewer)).frames, 900);
  const more = ps.concat([photo('new', 'land32')]);
  assert.equal(Core.buildPlan(project(more)).frames, 1200);
  // total mode holds the total and shares it out again
  const t1 = Core.buildPlan(project(ps, { timing: { mode: 'total', total: 70 } }));
  assert.equal(t1.frames, 2100);
  assert.equal(t1.segs.length, 7);
  assert.equal(t1.segs[0].slot, 300);
  // reordering keeps every start consistent with the new order
  const rev = ps.slice().reverse();
  const r = Core.buildPlan(project(rev));
  assert.deepEqual(included(r), rev.filter(p => p.include !== 'out').map(p => p.id));
  r.segs.forEach((s, i) => assert.equal(s.start, i * 150));
});

test('non looping export starts and ends cleanly with no closing transition', () => {
  const plan = Core.buildPlan(project(landscapes(5), { loop: false }));
  assert.equal(plan.frames, 750);
  assert.equal(plan.segs[4].trans, 0, 'no transition out of the last photo');
  assert.equal(plan.segs[0].transIn, 0, 'no transition into the first photo');
  const first = Core.stateAt(plan, 0), last = Core.stateAt(plan, plan.frames - 1);
  assert.equal(first.length, 1); assert.equal(first[0].seg.id, 'p0'); assert.equal(first[0].alpha, 1);
  assert.equal(last.length, 1); assert.equal(last[0].seg.id, 'p4'); assert.equal(last[0].alpha, 1);
  for (let k = plan.segs[4].start; k < plan.frames; k++) assert.equal(Core.stateAt(plan, k).length, 1, 'frame ' + k);
  assert.equal(first[0].phase, 0, 'first photo starts at the start of its move');
  // frames outside the range clamp instead of wrapping
  assert.equal(Core.wrapFrame(plan, plan.frames + 10), plan.frames - 1);
});

/* ================================================================== *
 * Suitability and exclusion                                           *
 * ================================================================== */

test('landscape output excludes portrait photos, with the reason', () => {
  const ps = [photo('l', 'land32'), photo('p', 'port23'), photo('q', 'port34'), photo('s', 'square')];
  const plan = Core.buildPlan(project(ps));
  assert.deepEqual(included(plan), ['l']);
  const why = plan.classes.find(c => c.id === 'p');
  assert.equal(why.included, false); assert.equal(why.auto, true);
  assert.match(why.reason, /^Portrait photo, too much cropping for this landscape output \(\d+% would stay visible, minimum 55%\)\.$/);
  assert.match(plan.classes.find(c => c.id === 's').reason, /^Square photo, too much cropping for this landscape output once its motion zooms in \(49%/);
});

test('portrait output excludes landscape photos', () => {
  const ps = [photo('l', 'land32'), photo('m', 'land169'), photo('p', 'port23'), photo('q', 'port34')];
  const plan = Core.buildPlan(project(ps, { output: { w: 1080, h: 1920 } }));
  assert.deepEqual(included(plan), ['p', 'q']);
  assert.match(plan.classes[0].reason, /^Landscape photo, too much cropping for this portrait output/);
});

test('slight ratio differences are accepted', () => {
  const ps = [photo('a', 'land32'), photo('b', 'land43'), photo('c', 'land1610'), photo('d', 'land54'), photo('e', [1921, 1080]), photo('f', [2048, 1080])];
  const plan = Core.buildPlan(project(ps));
  assert.equal(plan.counts.excluded, 0, JSON.stringify(plan.classes.filter(c => !c.included)));
  // and the common 4:3 phone photo still passes at the strongest motion setting
  const strong = Core.buildPlan(project([photo('b', 'land43')], { motion: { intensity: 'moderate' } }));
  assert.equal(strong.counts.included, 1);
});

test('the rule is the visible share of the photo, so shape alone does not decide', () => {
  // a panorama is landscape like the output, but far too wide for it
  const plan = Core.buildPlan(project([photo('pano', 'pano')]));
  assert.equal(plan.counts.included, 0);
  assert.match(plan.classes[0].reason, /^Photo is much wider than the output, too much cropping once its motion zooms in/);
  assert.match(Core.classify(photo('w', [8000, 2000]), project([])).reason, /^Photo is much wider than the output, too much cropping \(/);
  // the same panorama suits an ultra wide screen
  const wide = Core.buildPlan(project([photo('pano', 'pano')], { output: { w: 5760, h: 1920 } }));
  assert.equal(wide.counts.included, 1);
  // visible share matches the geometry
  const c = Core.classify(photo('x', 'land43'), project([], { motion: { enabled: false } }));
  assert.ok(Math.abs(c.visible - (4 / 3) / (16 / 9)) < 1e-12);
});

test('planned motion counts toward the crop, not just the static fit', () => {
  // 5:4 in 16:9 keeps 70.3% when still. Gentle motion zooms 7%, moderate 12%.
  const p = photo('x', 'land54');
  const at = (thr, motion) => Core.classify(p, project([p], { framing: { threshold: thr }, motion })).included;
  assert.equal(at(0.60, { enabled: false }), true);
  assert.equal(at(0.60, { intensity: 'gentle' }), true);     // 70.3 / 1.07^2 = 61.4
  assert.equal(at(0.60, { intensity: 'moderate' }), false);  // 70.3 / 1.12^2 = 56.0
  const vis = Core.classify(p, project([p], { motion: { intensity: 'moderate' } })).visible;
  assert.ok(Math.abs(vis - (1.25 / (16 / 9)) / (1.12 * 1.12)) < 1e-12);
  // a photo held still is judged without the motion allowance
  const still = photo('y', 'land54', { motion: { kind: 'still' } });
  assert.equal(Core.classify(still, project([still], { framing: { threshold: 0.60 }, motion: { intensity: 'moderate' } })).included, true);
});

test('the real motion never crops tighter than the suitability test assumed', () => {
  for (const intensity of Object.keys(Core.INTENSITY)) {
    for (const style of Core.MOTION_STYLES) {
      const ps = Object.keys(SHAPES).map((k, i) => photo('p' + i, k, { include: 'in' }));
      const proj = project(ps, { motion: { intensity, style }, timing: { perPhoto: 30 } });
      const plan = Core.buildPlan(proj);
      plan.segs.forEach(seg => {
        const cls = plan.classes.find(c => c.id === seg.id);
        for (let i = 0; i <= 20; i++) {
          const share = Core.visibleShare(Core.rectAt(seg, i / 20));
          assert.ok(share >= cls.visible - 1e-9, `${intensity} ${style} ${seg.id}: ${share} < ${cls.visible}`);
        }
      });
    }
  }
});

test('changing the output size re-evaluates every photo', () => {
  const ps = [photo('l', 'land32'), photo('p', 'port23'), photo('s', 'square')];
  assert.deepEqual(included(Core.buildPlan(project(ps))), ['l']);
  assert.deepEqual(included(Core.buildPlan(project(ps, { output: { w: 1080, h: 1920 } }))), ['p']);
  assert.deepEqual(included(Core.buildPlan(project(ps, { output: { w: 1440, h: 1080 } }))), ['l', 's']);
  // 3:2 either way round keeps 58% of itself in a square, which clears the default 55%
  assert.deepEqual(included(Core.buildPlan(project(ps, { output: { w: 1080, h: 1080 } }))), ['l', 'p', 's']);
  assert.deepEqual(included(Core.buildPlan(project(ps, { output: { w: 1080, h: 1080 }, framing: { threshold: 0.7 } }))), ['s']);
});

test('exclusion is reversible and manual choices survive a size change', () => {
  const ps = [photo('l', 'land32'), photo('p', 'port23')];
  assert.deepEqual(included(Core.buildPlan(project(ps))), ['l']);
  ps[1].include = 'in';
  const kept = Core.buildPlan(project(ps));
  assert.deepEqual(included(kept), ['l', 'p']);
  assert.equal(kept.classes[1].auto, false);
  assert.match(kept.classes[1].reason, /^Kept by you\./);
  ps[0].include = 'out';
  // user decisions hold when the output changes shape
  const flipped = Core.buildPlan(project(ps, { output: { w: 1080, h: 1920 } }));
  assert.deepEqual(included(flipped), ['p']);
  assert.equal(flipped.classes[0].reason, 'Excluded by you.');
  ps[0].include = 'auto'; ps[1].include = 'auto';
  assert.deepEqual(included(Core.buildPlan(project(ps))), ['l']);
});

test('framing set by hand is never excluded for cropping, and motion only failures say so', () => {
  const tight = photo('x', 'land169', { motion: { kind: 'custom', a: { cx: 0.5, cy: 0.5, z: 1 }, b: { cx: 0.5, cy: 0.5, z: 2.5 } } });
  const c = Core.classify(tight, project([tight]));
  assert.equal(c.included, true);
  assert.ok(c.visible < 0.2, 'the heavy crop is still reported honestly');
  const exact = photo('y', 'land169');
  const strict = Core.classify(exact, project([exact], { framing: { threshold: 0.95 } }));
  assert.equal(strict.included, false);
  assert.match(strict.reason, /^Same shape as the output, but the motion zoom alone takes it past the limit \(87% would stay visible, minimum 95%\)\.$/);
  assert.equal(Core.classify(exact, project([exact], { framing: { threshold: 0.95 }, motion: { enabled: false } })).included, true);
});

test('the threshold is adjustable', () => {
  const ps = [photo('s', 'square')];
  assert.equal(Core.buildPlan(project(ps)).counts.included, 0);
  assert.equal(Core.buildPlan(project(ps, { framing: { threshold: 0.45 } })).counts.included, 1);
  assert.equal(Core.buildPlan(project([photo('a', 'land32')], { framing: { threshold: 0.9 } })).counts.included, 0);
});

test('with the filter off every photo is used, in the chosen fit mode', () => {
  const ps = [photo('l', 'land32'), photo('p', 'port23'), photo('s', 'square'), photo('w', 'pano')];
  for (const mode of ['fill', 'fit', 'blur']) {
    const plan = Core.buildPlan(project(ps, { framing: { exclude: false, mode } }));
    assert.equal(plan.counts.included, 4);
    plan.segs.forEach(s => assert.equal(s.mode, mode));
  }
});

test('a photo with its own fit mode is never excluded for cropping', () => {
  const ps = [photo('p', 'port23', { framing: 'blur' }), photo('q', 'port23', { framing: 'fit' }), photo('r', 'port23')];
  const plan = Core.buildPlan(project(ps));
  assert.deepEqual(included(plan), ['p', 'q']);
  assert.deepEqual(plan.segs.map(s => s.mode), ['blur', 'fit']);
});

test('when nothing is left the plan says so instead of failing oddly', () => {
  const plan = Core.buildPlan(project([photo('p', 'port23'), photo('q', 'port34')]));
  assert.equal(plan.ok, false);
  assert.equal(plan.frames, 0);
  assert.ok(plan.errors.some(e => /No photos are left/.test(e.msg)));
  assert.deepEqual(Core.stateAt(plan, 0), []);
});

/* ================================================================== *
 * Motion geometry                                                     *
 * ================================================================== */

test('images are never stretched', () => {
  for (const mode of ['fill', 'fit', 'blur']) {
    const ps = Object.keys(SHAPES).map((k, i) => photo('p' + i, k));
    for (const out of [[1920, 1080], [1080, 1920], [1080, 1080], [3840, 1080], [1024, 768]]) {
      const plan = Core.buildPlan(project(ps, { output: { w: out[0], h: out[1] }, framing: { exclude: false, mode } }));
      const af = out[0] / out[1];
      plan.segs.forEach(seg => {
        for (const p of [0, 0.37, 1]) {
          const r = Core.rectAt(seg, p);
          const drawnAspect = (r.w * out[0]) / (r.h * out[1]);
          assert.ok(Math.abs(drawnAspect / (seg.iw / seg.ih) - 1) < 1e-9, `${mode} ${seg.id} ${out}`);
        }
      });
      assert.ok(af > 0);
    }
  }
});

test('every position along every move is valid, not only the two ends', () => {
  const rnd = Core.mulberry32(99);
  let checked = 0;
  for (let trial = 0; trial < 300; trial++) {
    const iw = 400 + Math.floor(rnd() * 6000), ih = 400 + Math.floor(rnd() * 6000);
    const ow = 2 * (160 + Math.floor(rnd() * 2000)), oh = 2 * (160 + Math.floor(rnd() * 1200));
    const mode = ['fill', 'fit', 'blur'][trial % 3];
    const kind = Core.PHOTO_MOTIONS[trial % Core.PHOTO_MOTIONS.length];
    const p = photo('x', [iw, ih], { include: 'in', focus: { x: rnd(), y: rnd() } });
    p.motion = kind === 'custom'
      ? { kind, a: { cx: rnd() * 1.4 - 0.2, cy: rnd() * 1.4 - 0.2, z: 0.5 + rnd() * 3 }, b: { cx: rnd() * 1.4 - 0.2, cy: rnd() * 1.4 - 0.2, z: 0.5 + rnd() * 3 } }
      : { kind };
    const proj = project([p, photo('y', 'land32', { include: 'in' })], {
      output: { w: ow, h: oh }, framing: { exclude: false, mode },
      motion: { intensity: ['subtle', 'gentle', 'moderate'][trial % 3], style: Core.MOTION_STYLES[trial % 5], seed: trial, variation: trial % 2 === 0 },
      timing: { perPhoto: 0.5 + rnd() * 20 }
    });
    const plan = Core.buildPlan(proj);
    assert.equal(plan.ok, true);
    const seg = plan.segs[0];
    for (let i = 0; i <= 50; i++) {
      const r = Core.rectAt(seg, i / 50);
      assert.ok(Core.rectValid(r, seg.mode === 'fill' ? 'fill' : 'fit', 1e-9), `trial ${trial} ${mode} ${kind} p=${i / 50} ${JSON.stringify(r)}`);
      checked++;
    }
  }
  assert.ok(checked > 10000);
});

test('fill never shows an empty edge and fit never crops, frame by frame', () => {
  const ps = Object.keys(SHAPES).map((k, i) => photo('p' + i, k));
  for (const mode of ['fill', 'fit']) {
    const plan = Core.buildPlan(project(ps, { framing: { exclude: false, mode }, timing: { perPhoto: 2 } }));
    for (let k = 0; k < plan.frames; k++) {
      Core.stateAt(plan, k).forEach(l => assert.ok(Core.rectValid(l.rect, mode, 1e-9), `${mode} frame ${k}`));
    }
  }
});

test('motion is restrained and slows down for short photos instead of speeding up', () => {
  const long = Core.buildPlan(project(landscapes(6, 'land169'), { timing: { perPhoto: 10 }, motion: { style: 'push' } }));
  const short = Core.buildPlan(project(landscapes(6, 'land169'), { timing: { perPhoto: 1.5 }, motion: { style: 'push' } }));
  const zoomOf = s => Math.max(s.a.w, s.b.w) / Math.min(s.a.w, s.b.w) - 1;
  long.segs.forEach(s => assert.ok(Math.abs(zoomOf(s) - 0.07) < 1e-9));
  short.segs.forEach(s => assert.ok(zoomOf(s) < 0.035 && zoomOf(s) > 0));
  // per second rate never exceeds the intensity cap
  for (const plan of [long, short]) plan.segs.forEach(s => assert.ok(zoomOf(s) / (s.life / plan.fps) <= Core.INTENSITY.gentle.zoomRate + 1e-9));
});

test('a push holds the focus point still on screen', () => {
  const p = photo('x', 'land169', { focus: { x: 0.7, y: 0.3 }, motion: { kind: 'push' } });
  const plan = Core.buildPlan(project([p, photo('y', 'land169')]));
  const seg = plan.segs[0];
  const at = r => ({ x: r.x + 0.7 * r.w, y: r.y + 0.3 * r.h });
  const a = at(Core.rectAt(seg, 0)), b = at(Core.rectAt(seg, 1)), mid = at(Core.rectAt(seg, 0.5));
  assert.ok(Math.abs(a.x - b.x) < 1e-9 && Math.abs(a.y - b.y) < 1e-9);
  assert.ok(Math.abs(a.x - mid.x) < 1e-9 && Math.abs(a.y - mid.y) < 1e-9);
});

test('directional pans go the way they say and keep a constant scale', () => {
  const dir = { 'pan-left': ['x', -1], 'pan-right': ['x', 1], 'pan-up': ['y', -1], 'pan-down': ['y', 1] };
  for (const kind of Object.keys(dir)) {
    const p = photo('x', 'land32', { motion: { kind } });
    const seg = Core.buildPlan(project([p, photo('y', 'land32')])).segs[0];
    const a = Core.rectToCrop(seg.a, 1.5, 16 / 9), b = Core.rectToCrop(seg.b, 1.5, 16 / 9);
    const [axis, sign] = dir[kind];
    const d = axis === 'x' ? b.cx - a.cx : b.cy - a.cy, other = axis === 'x' ? b.cy - a.cy : b.cx - a.cx;
    assert.ok(d * sign > 1e-4, kind + ' moved ' + d);
    assert.ok(Math.abs(other) < 1e-12);
    assert.ok(Math.abs(seg.a.w - seg.b.w) < 1e-12);
  }
});

test('zoom runs at a constant apparent speed', () => {
  const p = photo('x', 'land169', { motion: { kind: 'custom', a: { cx: 0.5, cy: 0.5, z: 1 }, b: { cx: 0.5, cy: 0.5, z: 2 } } });
  const seg = Core.buildPlan(project([p, photo('y', 'land169')])).segs[0];
  const ratios = [];
  for (let i = 0; i < 10; i++) ratios.push(Core.rectAt(seg, (i + 1) / 10).w / Core.rectAt(seg, i / 10).w);
  ratios.forEach(r => assert.ok(Math.abs(r - ratios[0]) < 1e-12));
  assert.ok(Math.abs(ratios[0] - Math.pow(2, 0.1)) < 1e-12);
});

test('custom framing is clamped inside the photo', () => {
  const c = Core.clampCrop(1.5, 16 / 9, { cx: -3, cy: 9, z: 0.2 });
  assert.equal(c.z, 1);
  assert.ok(Core.rectValid(Core.cropRect(1.5, 16 / 9, c.cx, c.cy, c.z), 'fill'));
  const back = Core.rectToCrop(Core.cropRect(1.5, 16 / 9, 0.4, 0.55, 1.6), 1.5, 16 / 9);
  assert.ok(Math.abs(back.cx - 0.4) < 1e-12 && Math.abs(back.cy - 0.55) < 1e-12 && Math.abs(back.z - 1.6) < 1e-12);
});

test('motion off means every photo is still', () => {
  const plan = Core.buildPlan(project(landscapes(5), { motion: { enabled: false } }));
  plan.segs.forEach(s => { assert.equal(s.moving, false); assert.equal(rectDelta(s.a, s.b), 0); });
});

/* ================================================================== *
 * Ordering and determinism                                            *
 * ================================================================== */

test('a seeded shuffle is repeatable and is a real reordering', () => {
  const ids = Array.from({ length: 60 }, (_, i) => 'p' + i);
  const a = Core.seededShuffle(ids, 0xC0FFEE), b = Core.seededShuffle(ids, 0xC0FFEE), c = Core.seededShuffle(ids, 0xC0FFEF);
  assert.deepEqual(a, b);
  assert.notDeepEqual(a, ids);
  assert.notDeepEqual(a, c);
  assert.deepEqual(a.slice().sort(), ids.slice().sort(), 'nothing lost, nothing repeated');
  assert.deepEqual(ids, Array.from({ length: 60 }, (_, i) => 'p' + i), 'input untouched');
});

test('the same project always gives the same plan', () => {
  const build = () => Core.buildPlan(project(Core.seededShuffle(landscapes(40), 777)));
  const a = build(), b = build();
  assert.deepEqual(a.segs, b.segs);
  for (const k of [0, 17, 449, 1234, a.frames - 1]) assert.deepEqual(Core.stateAt(a, k), Core.stateAt(b, k));
});

test('order and motion are randomised independently', () => {
  const ps = landscapes(30);
  const base = Core.buildPlan(project(ps, { motion: { seed: 1 } }));
  const newMotion = Core.buildPlan(project(ps, { motion: { seed: 2 } }));
  assert.deepEqual(included(newMotion), included(base), 'a new motion seed leaves the order alone');
  assert.notDeepEqual(newMotion.segs.map(s => s.a), base.segs.map(s => s.a), 'but changes the moves');
  const shuffled = Core.seededShuffle(ps, 4242);
  const reordered = Core.buildPlan(project(shuffled, { motion: { seed: 1 } }));
  assert.deepEqual(included(reordered), shuffled.map(p => p.id));
  assert.notDeepEqual(included(reordered), included(base));
  assert.deepEqual(reordered.segs.map(s => s.kind), base.segs.map(s => s.kind), 'the sequence of moves is unchanged by reordering');
});

test('manual order is kept exactly', () => {
  const ps = landscapes(10);
  const moved = ps.slice();
  moved.splice(7, 0, moved.splice(2, 1)[0]);
  assert.deepEqual(included(Core.buildPlan(project(moved))), ['p0', 'p1', 'p3', 'p4', 'p5', 'p6', 'p7', 'p2', 'p8', 'p9']);
});

test('automatic moves avoid erratic repetition', () => {
  for (let seed = 1; seed <= 50; seed++) {
    const m = Core.assignMoves(100, 'mixed', true, seed).map(x => x.kind);
    for (let i = 2; i < m.length; i++) assert.ok(!(m[i] === m[i - 1] && m[i] === m[i - 2]), 'three alike in a row, seed ' + seed);
    for (let i = 1; i < m.length; i++) assert.ok(!(m[i] === 'pan' && m[i - 1] === 'pan'), 'two pans in a row, seed ' + seed);
    assert.ok(new Set(m).size === 3);
    const pans = Core.assignMoves(100, 'mixed', true, seed).filter(x => x.kind === 'pan').map(x => x.dir);
    for (let i = 1; i < pans.length; i++) assert.equal(pans[i], -pans[i - 1], 'pan direction alternates');
  }
});

/* ================================================================== *
 * Seamless loop                                                       *
 * ================================================================== */

test('the timeline is periodic: frame N is frame 0', () => {
  const plan = Core.buildPlan(project(Core.seededShuffle(landscapes(25), 5)));
  const N = plan.frames;
  for (const k of [0, 1, 29, 150, 1111, N - 31, N - 1]) {
    assert.deepEqual(Core.stateAt(plan, k + N), Core.stateAt(plan, k));
    assert.deepEqual(Core.stateAt(plan, k - N), Core.stateAt(plan, k));
  }
});

test('the closing transition is a real transition and ends at the file boundary', () => {
  const plan = Core.buildPlan(project(landscapes(8)));
  const N = plan.frames, X = plan.segs[7].trans;
  assert.equal(X, 30);
  const firstId = plan.segs[0].id, lastId = plan.segs[7].id;
  let prevAlpha = 0;
  for (let j = 0; j < X; j++) {
    const layers = Core.stateAt(plan, N - X + j);
    assert.equal(layers.length, 2);
    assert.equal(layers[0].seg.id, lastId, 'last photo underneath');
    assert.equal(layers[1].seg.id, firstId, 'first photo fading in on top');
    assert.equal(layers[0].alpha, 1);
    assert.ok(layers[1].alpha > prevAlpha && layers[1].alpha < 1, 'fade strictly between 0 and 1 and rising');
    prevAlpha = layers[1].alpha;
  }
  assert.equal(Core.stateAt(plan, N - X - 1).length, 1, 'frame before the transition is the last photo alone');
  const zero = Core.stateAt(plan, 0);
  assert.equal(zero.length, 1);
  assert.equal(zero[0].seg.id, firstId, 'frame 0 is the first photo, fully on screen');
});

test('the closing transition matches every other transition, sample for sample', () => {
  const plan = Core.buildPlan(project(landscapes(8, 'land169'), { motion: { style: 'push', variation: false } }));
  const N = plan.frames, X = 30;
  const alphas = at => Array.from({ length: X }, (_, j) => Core.stateAt(plan, at + j)[1].alpha);
  assert.deepEqual(alphas(N - X), alphas(plan.segs[1].start - X));
  assert.deepEqual(alphas(N - X), alphas(plan.segs[4].start - X));
});

test('no frame is repeated across the loop point', () => {
  const plan = Core.buildPlan(project(landscapes(8)));
  const N = plan.frames;
  const last = Core.stateAt(plan, N - 1), first = Core.stateAt(plan, 0), second = Core.stateAt(plan, 1);
  assert.notDeepEqual(last, first);
  const lastTop = layerOf(last, plan.segs[0].id);
  assert.ok(lastTop.alpha < 1, 'the last frame is still mid fade, so it cannot equal frame 0');
  assert.ok(rectDelta(lastTop.rect, first[0].rect) > 0, 'and the photo has moved between them');
  assert.ok(rectDelta(first[0].rect, second[0].rect) > 0);
});

test('motion carries across the loop point with no reset and no change of speed', () => {
  for (const [n, seed] of [[2, 1], [3, 2], [30, 3], [100, 4]]) {
    const ps = Core.seededShuffle(landscapes(n), seed);
    const plan = Core.buildPlan(project(ps, { motion: { seed } }));
    const N = plan.frames, id = plan.segs[0].id;
    // photo 0 is visible from the start of the closing transition right through its own slot
    const frames = [];
    for (let k = N - plan.segs[n - 1].trans; k < N + 60; k++) frames.push(layerOf(Core.stateAt(plan, k), id));
    frames.forEach(l => assert.ok(l, 'photo 0 is on screen on both sides of the boundary'));
    const phases = frames.map(l => l.phase);
    for (let i = 1; i < phases.length; i++) {
      assert.ok(Math.abs((phases[i] - phases[i - 1]) - 1 / plan.segs[0].life) < 1e-12, `n=${n}: phase advances by exactly one frame at index ${i}`);
    }
    // step sizes in the drawn rectangle are continuous across the boundary
    const steps = [];
    for (let i = 1; i < frames.length; i++) steps.push(rectDelta(frames[i].rect, frames[i - 1].rect));
    const boundary = plan.segs[n - 1].trans - 1;          // step from frame N-1 to frame 0
    const before = steps[boundary - 1], across = steps[boundary], after = steps[boundary + 1];
    if (plan.segs[0].moving) {
      assert.ok(across > 0);
      assert.ok(Math.abs(across - before) / across < 0.01 && Math.abs(across - after) / across < 0.01,
        `n=${n}: step across the loop point ${across} vs ${before} and ${after}`);
    }
  }
});

test('across the whole loop no visible photo ever jumps', () => {
  const ps = Core.seededShuffle(landscapes(40), 2024);
  const plan = Core.buildPlan(project(ps, { timing: { perPhoto: 3 } }));
  const N = plan.frames;
  let prev = Core.stateAt(plan, N - 1), worst = 0, worstAlpha = 0;
  for (let k = 0; k < N; k++) {
    const cur = Core.stateAt(plan, k);
    cur.forEach(l => {
      const was = layerOf(prev, l.seg.id);
      if (!was) { assert.ok(l.alpha < 0.05, `frame ${k}: a photo appears at alpha ${l.alpha}`); return; }
      worst = Math.max(worst, rectDelta(l.rect, was.rect));
      worstAlpha = Math.max(worstAlpha, Math.abs(l.alpha - was.alpha));
    });
    prev.forEach(l => { if (!layerOf(cur, l.seg.id)) assert.ok(cur[0].alpha === 1 && cur.length >= 1, `frame ${k}: a photo vanished without being covered`); });
    prev = cur;
  }
  assert.ok(worst < 0.0025, 'largest per frame move in frame units: ' + worst);
  assert.ok(worstAlpha < 0.06, 'largest per frame fade step: ' + worstAlpha);
});

test('one photo loop: the move goes out and comes back with a smooth turn, never a reset', () => {
  const plan = Core.buildPlan(project([photo('only', 'land32')], { timing: { perPhoto: 10 } }));
  const N = plan.frames, seg = plan.segs[0];
  assert.equal(N, 300);
  assert.equal(seg.periodic, true);
  assert.equal(seg.trans, 0, 'a photo does not dissolve into itself');
  const r = k => Core.stateAt(plan, k)[0].rect;
  for (let k = 0; k < N; k++) assert.equal(Core.stateAt(plan, k).length, 1);
  assert.deepEqual(r(N), r(0), 'periodic');
  assert.ok(rectDelta(r(N / 2), r(0)) > 0.02, 'it really moves');
  // symmetric about the boundary, which is what makes the turn smooth
  for (const d of [1, 2, 7, 40]) assert.ok(rectDelta(r(N - d), r(d)) < 1e-9);
  let maxStep = 0, maxAccel = 0, prevStep = rectDelta(r(0), r(N - 1));
  for (let k = 0; k < N; k++) {
    const step = rectDelta(r(k + 1), r(k));
    maxStep = Math.max(maxStep, step);
    maxAccel = Math.max(maxAccel, Math.abs(step - prevStep));
    prevStep = step;
  }
  const seamStep = rectDelta(r(0), r(N - 1));
  assert.ok(seamStep > 0, 'frame N-1 and frame 0 are different frames');
  assert.ok(seamStep < maxStep * 0.05, 'and the move is at its slowest there, not jumping');
  assert.ok(maxAccel < maxStep * 0.05, 'no sudden change of speed anywhere, the loop point included');
  // every position valid
  for (let k = 0; k < N; k++) assert.ok(Core.rectValid(r(k), 'fill', 1e-9));
});

test('one photo, not looping: a single one way move', () => {
  const plan = Core.buildPlan(project([photo('only', 'land32')], { loop: false }));
  assert.equal(plan.segs[0].periodic, false);
  assert.equal(Core.stateAt(plan, 0)[0].phase, 0);
});

test('two photo loop: both transitions exist and both photos move through them', () => {
  const plan = Core.buildPlan(project([photo('a', 'land32'), photo('b', 'land43')], { timing: { perPhoto: 4 } }));
  const N = plan.frames;
  assert.equal(N, 240);
  assert.deepEqual(plan.segs.map(s => [s.trans, s.transIn, s.life]), [[30, 30, 150], [30, 30, 150]]);
  const mid = Core.stateAt(plan, 119), end = Core.stateAt(plan, 239);
  assert.deepEqual(mid.map(l => l.seg.id), ['a', 'b']);
  assert.deepEqual(end.map(l => l.seg.id), ['b', 'a']);
  // a's phase: fades in over the last 30 frames, then continues from frame 0
  assert.ok(Math.abs(layerOf(end, 'a').phase - 29 / 150) < 1e-12);
  assert.ok(Math.abs(Core.stateAt(plan, 0)[0].phase - 30 / 150) < 1e-12);
  // b is hidden when its move restarts, so the restart is never seen
  assert.equal(layerOf(Core.stateAt(plan, 0), 'b'), undefined);
});

test('a cut at the loop point is reported, not passed off as a dissolve', () => {
  const plan = Core.buildPlan(project(landscapes(4), { timing: { transitionType: 'cut' } }));
  assert.ok(plan.notes.some(n => /loop point is a cut/.test(n)));
  const dissolve = Core.buildPlan(project(landscapes(4)));
  assert.ok(!dissolve.notes.some(n => /loop point is a cut/.test(n)));
  const once = Core.buildPlan(project(landscapes(4), { loop: false, timing: { transitionType: 'cut' } }));
  assert.ok(!once.notes.some(n => /loop point is a cut/.test(n)));
});

/* ================================================================== *
 * Rendering                                                           *
 * ================================================================== */

test('renderer draws what the timeline says, with fractional positions kept', () => {
  const plan = Core.buildPlan(project(landscapes(5)));
  const W = 1920, H = 1080, k = 137;
  const out = renderCalls(plan, k, W, H);
  const layers = Core.stateAt(plan, k);
  const draws = out.main.filter(c => c.op === 'draw');
  assert.equal(draws.length, layers.length);
  layers.forEach((l, i) => {
    assert.deepEqual(draws[i].img, { id: l.seg.id });
    assert.equal(draws[i].x, l.rect.x * W); assert.equal(draws[i].y, l.rect.y * H);
    assert.equal(draws[i].w, l.rect.w * W); assert.equal(draws[i].h, l.rect.h * H);
    assert.equal(draws[i].alpha, l.alpha);
  });
  assert.equal(out.main[0].op, 'fill', 'frame is cleared to black first');
  assert.ok(draws.some(d => d.w !== Math.round(d.w) || d.x !== Math.round(d.x)), 'sub pixel positions are not rounded away');
});

test('the frames either side of the loop point render as neighbours', () => {
  const plan = Core.buildPlan(project(Core.seededShuffle(landscapes(30), 8), { motion: { seed: 8 } }));
  const N = plan.frames, W = 1920, H = 1080, firstId = plan.segs[0].id;
  const drawOf = k => renderCalls(plan, k, W, H).main.filter(c => c.op === 'draw' && c.img.id === firstId)[0];
  const a = drawOf(N - 2), b = drawOf(N - 1), c = drawOf(0), d = drawOf(1);
  [a, b, c, d].forEach(x => assert.ok(x, 'first photo drawn in all four frames'));
  const step = (p, q) => Math.hypot(q.x - p.x, q.y - p.y, q.w - p.w, q.h - p.h);
  const s1 = step(a, b), s2 = step(b, c), s3 = step(c, d);
  assert.ok(s2 > 0, 'not a repeated frame');
  assert.ok(Math.abs(s2 - s1) / s2 < 0.01 && Math.abs(s2 - s3) / s2 < 0.01, `pixel steps ${s1} ${s2} ${s3}`);
  assert.ok(s2 < 2, 'about a pixel per frame at 1080p, so fractional positioning matters: ' + s2);
  assert.ok(b.alpha < 1 && b.alpha > 0.99, 'last frame is the tail of the fade: ' + b.alpha);
  assert.equal(c.alpha, 1);
});

test('preview and export render the same composition at any size', () => {
  const ps = [photo('a', 'land32'), photo('b', 'port23', { include: 'in', framing: 'blur' }), photo('c', 'land43', { opaque: false }), photo('d', 'square', { include: 'in', framing: 'fit' })];
  const plan = Core.buildPlan(project(ps, { timing: { perPhoto: 2, transition: 0.5 } }));
  for (let k = 0; k < plan.frames; k += 7) {
    const big = renderCalls(plan, k, 3840, 2160), small = renderCalls(plan, k, 960, 540);
    for (const part of ['main', 'scratch']) {
      assert.equal(big[part].length, small[part].length, `frame ${k} ${part}`);
      big[part].forEach((c, i) => {
        const s = small[part][i];
        assert.equal(c.op, s.op); assert.equal(c.alpha, s.alpha); assert.deepEqual(c.img, s.img);
        for (const f of ['x', 'y', 'w', 'h']) if (c[f] !== undefined) assert.ok(Math.abs(c[f] / 4 - s[f]) < 1e-9, `frame ${k} ${f}`);
      });
    }
  }
});

test('layers that do not fill the frame are built whole before they fade in', () => {
  const ps = [photo('a', 'land32'), photo('b', 'port23', { include: 'in', framing: 'blur' })];
  const plan = Core.buildPlan(project(ps, { timing: { perPhoto: 2, transition: 1 } }));
  const k = plan.segs[1].start - 10;      // b is fading in over a
  const layers = Core.stateAt(plan, k);
  assert.deepEqual(layers.map(l => l.seg.id), ['a', 'b']);
  const out = renderCalls(plan, k, 1920, 1080);
  const top = out.main[out.main.length - 1];
  assert.deepEqual(top.img, { tag: 'scratch' }, 'the finished layer is composited in one draw');
  assert.equal(top.alpha, layers[1].alpha);
  assert.ok(out.scratch.some(c => c.op === 'draw' && c.img.blurOf === 'b'), 'blurred background drawn on the scratch layer');
  assert.ok(out.scratch.some(c => c.op === 'draw' && c.img.id === 'b'));
  out.scratch.filter(c => c.op === 'draw').forEach(c => assert.equal(c.alpha, 1));
  // an opaque photo that fills the frame is blended directly
  const direct = renderCalls(Core.buildPlan(project(landscapes(3))), 140, 1920, 1080);
  assert.equal(direct.scratch.length, 0);
});

test('a photo that is not decoded yet is reported, not drawn wrong', () => {
  const plan = Core.buildPlan(project(landscapes(3)));
  const ctx = fakeCtx();
  const missing = Render.renderFrame(ctx, 640, 360, plan, 0, () => null, () => fakeCtx());
  assert.deepEqual(missing, [plan.segs[0].id]);
  assert.equal(ctx.calls.filter(c => c.op === 'draw').length, 0);
});

test('box blur keeps flat colour flat and spreads a point', () => {
  const w = 16, h = 12, flat = new Uint8ClampedArray(w * h * 4).fill(200);
  Render.boxBlur(flat, w, h, 3, 3);
  assert.ok(flat.every(v => v === 200));
  const dot = new Uint8ClampedArray(w * h * 4);
  dot[(6 * w + 8) * 4] = 255;
  Render.boxBlur(dot, w, h, 2, 2);
  assert.ok(dot[(6 * w + 8) * 4] < 60 && dot[(6 * w + 8) * 4] > 0);
  assert.ok(dot[(6 * w + 6) * 4] > 0 && dot[(4 * w + 8) * 4] > 0);
});

/* ================================================================== *
 * Codecs, colour and output                                           *
 * ================================================================== */

test('H.264 levels match the standard for common sizes', () => {
  const lv = (w, h, fps, mbps) => Core.avcLevel(w, h, fps, mbps * 1e6, 1.25).name;
  assert.equal(lv(1280, 720, 30, 8), '3.1');
  assert.equal(lv(1920, 1080, 30, 14), '4.0');
  assert.equal(lv(1920, 1080, 60, 20), '4.2');
  assert.equal(lv(3840, 2160, 30, 45), '5.1');
  assert.equal(lv(3840, 2160, 60, 70), '5.2');
  assert.equal(lv(3840, 1080, 30, 30), '5.0');
  assert.equal(lv(7680, 1080, 30, 60), '5.1');
  assert.equal(Core.avcLevel(16384, 8192, 60, 50e6, 1.25), null);
  assert.equal(Core.codecCandidates('avc', 1920, 1080, 30, 14e6).list[0].codec, 'avc1.640028');
  assert.equal(Core.codecCandidates('avc', 3840, 2160, 60, 70e6).list[0].codec, 'avc1.640034');
  assert.equal(Core.codecCandidates('vp9', 1920, 1080, 30, 10e6).list[0].codec, 'vp09.00.40.08');
  assert.equal(Core.codecCandidates('av1', 1920, 1080, 30, 8e6).list[0].codec, 'av01.0.08M.08');
  assert.equal(Core.codecCandidates('hevc', 1920, 1080, 30, 9e6).list[0].codec, 'hvc1.1.6.L120.B0');
});

test('odd sizes are refused for H.264 with the reason, and allowed for VP9 and AV1', () => {
  const r = Core.codecCandidates('avc', 1365, 768, 30, 8e6);
  assert.equal(r.list.length, 0);
  assert.equal(r.oddDims, true);
  assert.match(r.reason, /even width and height/);
  assert.ok(Core.codecCandidates('vp9', 1365, 767, 30, 8e6).list.length);
  assert.ok(Core.codecCandidates('av1', 1365, 767, 30, 8e6).list.length);
});

test('suggested bitrate rises with size, rate and quality, and stays sane', () => {
  const b = (w, h, fps, q) => Core.suggestBitrate('avc', w, h, fps, q);
  assert.equal(b(1920, 1080, 30, 'high'), 25e6);
  assert.equal(Core.suggestBitrate('avc', 7680, 4320, 60, 'max'), Core.BITRATE_MAX * 1e6);
  assert.equal(Core.proResBitrate(3, 1920, 1080, 29.97), 220e6);
  assert.ok(Math.abs(Core.proResBitrate(3, 3840, 2160, 25) / 1e6 - 220 * 4 * 25 / 29.97) < 1);
  assert.ok(b(3840, 2160, 30, 'high') > b(1920, 1080, 30, 'high') * 3 && b(3840, 2160, 30, 'high') < b(1920, 1080, 30, 'high') * 4);
  assert.ok(b(1920, 1080, 60, 'high') > b(1920, 1080, 30, 'high') && b(1920, 1080, 60, 'high') < b(1920, 1080, 30, 'high') * 2);
  assert.ok(b(1920, 1080, 30, 'standard') < b(1920, 1080, 30, 'high') && b(1920, 1080, 30, 'high') < b(1920, 1080, 30, 'max'));
  assert.ok(Core.suggestBitrate('av1', 1920, 1080, 30, 'high') < b(1920, 1080, 30, 'high'));
  assert.ok(Math.abs(Core.estimateBytes(8e6, 100) - 101e6) < 1);
});

test('colour conversion uses BT.709 for HD sizes and BT.601 for small ones', () => {
  const px = (r, g, b) => { const a = new Uint8Array(16); for (let i = 0; i < 4; i++) { a[i * 4] = r; a[i * 4 + 1] = g; a[i * 4 + 2] = b; a[i * 4 + 3] = 255; } return a; };
  // the converter is built for a frame size, so use a 2 x 2 stand-in and pick the matrix through isHdSize
  assert.equal(Exp.isHdSize(1920, 1080), true);
  assert.equal(Exp.isHdSize(1024, 768), true);
  assert.equal(Exp.isHdSize(720, 576), false);
  assert.equal(Exp.isHdSize(640, 360), false);
  const hd = Exp.createYuvConverter(1280, 2), sd = Exp.createYuvConverter(2, 2);
  assert.deepEqual(hd.colorSpace, { primaries: 'bt709', transfer: 'bt709', matrix: 'bt709', fullRange: false });
  assert.equal(sd.colorSpace.matrix, 'smpte170m');
  const run = (c, W, r, g, b) => { const src = new Uint8Array(W * 2 * 4); for (let i = 0; i < W * 2; i++) { src[i * 4] = r; src[i * 4 + 1] = g; src[i * 4 + 2] = b; } const o = c.convert(src, W * 4, false); return [o[0], o[W * 2], o[W * 2 + ((W + 1) >> 1)]]; };
  assert.deepEqual(run(hd, 1280, 255, 255, 255), [235, 128, 128]);
  assert.deepEqual(run(hd, 1280, 0, 0, 0), [16, 128, 128]);
  assert.deepEqual(run(hd, 1280, 255, 0, 0), [63, 102, 240]);     // BT.709 red
  assert.deepEqual(run(sd, 2, 255, 0, 0), [81, 90, 240]);         // BT.601 red
  assert.deepEqual(run(hd, 1280, 0, 0, 255), [32, 240, 118]);     // BT.709 blue
  // BGRA input gives the same answer as RGBA
  const src = px(200, 40, 30), bgr = px(30, 40, 200);
  assert.deepEqual(Array.from(sd.convert(src, 8, false)), Array.from(Exp.createYuvConverter(2, 2).convert(bgr, 8, true)));
  // odd sizes produce correctly sized planes
  const odd = Exp.createYuvConverter(5, 3);
  assert.equal(odd.buffer.length, 15 + 2 * 3 * 2);
  odd.convert(new Uint8Array(5 * 3 * 4).fill(128), 20, false);
  assert.ok(Array.from(odd.buffer.subarray(15)).every(v => v === 128));
});

test('size validation and presets', () => {
  assert.equal(Core.validateDims(1920, 1080), null);
  assert.equal(Core.validateDims(7680, 1080), null);
  assert.match(Core.validateDims(8, 1080), /at least/);
  assert.match(Core.validateDims(20000, 1080), /16384/);
  assert.match(Core.validateDims(1920.5, 1080), /whole pixels/);
  assert.match(Core.validateDims(NaN, 1080), /Enter a width/);
  const bad = Core.buildPlan(project(landscapes(3), { output: { w: NaN, h: 1080 } }));
  assert.equal(bad.ok, false);
  const names = Core.PRESETS.map(p => p.w + 'x' + p.h);
  for (const need of ['1280x720', '1920x1080', '2560x1440', '3840x2160', '1080x1920', '2160x3840', '1080x1080', '1024x768', '1920x1200']) assert.ok(names.includes(need), need);
  assert.equal(Core.findPreset(1920, 1080).id, '1080p');
  assert.equal(Core.findPreset(1921, 1080), null);
  assert.deepEqual(Core.FPS_CHOICES, [24, 25, 30, 50, 60]);
});

test('time and duration text', () => {
  assert.equal(Core.formatTime(0, 30), '0:00.00');
  assert.equal(Core.formatTime(4500, 30), '2:30.00');
  assert.equal(Core.formatTime(151, 30), '0:05.03');
  assert.equal(Core.parseDuration('90'), 90);
  assert.equal(Core.parseDuration('1:30'), 90);
  assert.equal(Core.parseDuration('1:30.5'), 90.5);
  assert.equal(Core.parseDuration('0:01:30'), 90);
  assert.equal(Core.parseDuration('2,5'), 2.5);
  assert.ok(Number.isNaN(Core.parseDuration('abc')));
  assert.ok(Number.isNaN(Core.parseDuration('')));
  assert.equal(Core.aspectLabel(1920, 1080), '16:9 (1.78:1)');
  assert.equal(Core.aspectLabel(1366, 768), '1.78:1');
  assert.equal(Core.outputFilename({ W: 1920, H: 1080, fps: 30, loop: true }, 'mp4'), 'montage_1920x1080_30fps_loop.mp4');
});

test('resource warnings are practical and never claim to know the limit', () => {
  const light = Core.assessResources({ W: 1920, H: 1080, fps: 30, frames: 4500, photoCount: 30, maxPhotoPixels: 24e6, bitrate: 14e6, toDisk: false, canStream: true, deviceMemory: 8 });
  assert.deepEqual(light.risks, []);
  const heavy = Core.assessResources({ W: 3840, H: 2160, fps: 60, frames: 54000, photoCount: 180, maxPhotoPixels: 24e6, bitrate: 70e6, toDisk: false, canStream: true, deviceMemory: 8 });
  assert.ok(heavy.risks.some(r => /^This export is demanding, 3840 x 2160 at 60 fps with 180 photos\. Reducing the frame rate or resolution may improve reliability\.$/.test(r)), heavy.risks.join(' | '));
  assert.ok(heavy.risks.some(r => /held in memory/.test(r)));
  assert.ok(!heavy.risks.some(r => /will fail|cannot|crash/i.test(r.replace('This browser cannot save straight to disk.', ''))));
  const disk = Core.assessResources({ W: 3840, H: 2160, fps: 60, frames: 54000, photoCount: 180, maxPhotoPixels: 24e6, bitrate: 70e6, toDisk: true, canStream: true, deviceMemory: 8 });
  assert.ok(!disk.risks.some(r => /held in memory/.test(r)));
  assert.ok(disk.memBytes < heavy.memBytes);
});
