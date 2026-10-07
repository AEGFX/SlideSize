/* ============================================================
   Slideshow Builder core  (SlideshowCore)
   Timeline, motion, crop suitability, ordering and codec maths.

   Pure functions with no DOM access, so the same file runs in the
   page, in the export worker and in Node for the tests.

   The timeline built here is the single source of truth. The
   preview and the exporter both ask stateAt(plan, frame) what to
   draw, so they cannot drift apart.

   AEGFX / SlideSize
   ============================================================ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SlideshowCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
'use strict';

/* ------------------------------------------------------------------ *
 * Constants                                                           *
 * ------------------------------------------------------------------ */

var FPS_CHOICES = [24, 25, 30, 50, 60];

/* The first four mirror the presets on the SlideSize calculator. */
var PRESETS = [
  { id: '720p',   name: '720p HD',          w: 1280, h: 720 },
  { id: '1080p',  name: '1080p FHD',        w: 1920, h: 1080 },
  { id: '1440p',  name: '1440p QHD',        w: 2560, h: 1440 },
  { id: '2160p',  name: '4K UHD',           w: 3840, h: 2160 },
  { id: 'p1080',  name: 'Portrait 1080p',   w: 1080, h: 1920 },
  { id: 'p2160',  name: 'Portrait 4K',      w: 2160, h: 3840 },
  { id: 'sq1080', name: 'Square',           w: 1080, h: 1080 },
  { id: 'sq2160', name: 'Square 4K',        w: 2160, h: 2160 },
  { id: 'xga',    name: '4:3 XGA',          w: 1024, h: 768 },
  { id: '4x3',    name: '4:3 1440 x 1080',  w: 1440, h: 1080 },
  { id: 'wxga',   name: '16:10 WXGA',       w: 1280, h: 800 },
  { id: 'wuxga',  name: '16:10 WUXGA',      w: 1920, h: 1200 },
  { id: 'dw1080', name: '32:9 double wide', w: 3840, h: 1080 }
];

var MIN_DIM = 16;
var MAX_DIM = 16384;                 /* Chrome canvas edge limit */
var MAX_AREA = 16384 * 16384;        /* Chrome canvas area limit */

var MIN_SLOT_SECONDS = 0.25;         /* hard floor for one photo */
var FAST_SLOT_SECONDS = 1.5;         /* soft warning below this */
var MAX_TRANSITION_SHARE = 0.5;      /* a transition may use at most half a photo's time */

var DEFAULT_THRESHOLD = 0.55;        /* minimum visible share of a photo, see classify() */
var DEFAULT_FOCUS = { x: 0.5, y: 0.45 };

/* zoom and pan are the most a photo travels over its whole time on
   screen. zoomRate and panRate cap the speed per second, so short
   photos move less instead of moving faster. */
var INTENSITY = {
  subtle:   { zoom: 0.04, pan: 0.025, zoomRate: 0.008, panRate: 0.005 },
  gentle:   { zoom: 0.07, pan: 0.04,  zoomRate: 0.014, panRate: 0.008 },
  moderate: { zoom: 0.12, pan: 0.07,  zoomRate: 0.024, panRate: 0.014 }
};

var MOTION_STYLES = ['mixed', 'push', 'pull', 'pushpull', 'pan'];
var PHOTO_MOTIONS = ['auto', 'push', 'pull', 'pan-left', 'pan-right', 'pan-up', 'pan-down', 'still', 'custom'];
var FRAMING_MODES = ['fill', 'fit', 'blur'];

/* ------------------------------------------------------------------ *
 * Small helpers                                                       *
 * ------------------------------------------------------------------ */

function clamp(v, lo, hi) { return v < lo ? lo : (v > hi ? hi : v); }
function lerp(a, b, t) { return a + (b - a) * t; }
function gcd(a, b) { a = Math.round(Math.abs(a)); b = Math.round(Math.abs(b)); while (b) { var t = b; b = a % b; a = t; } return a || 1; }
function smoothstep(q) { q = clamp(q, 0, 1); return q * q * (3 - 2 * q); }
function trimNum(v, dp) { return String(parseFloat(Number(v).toFixed(dp == null ? 3 : dp))); }

function aspectLabel(w, h) {
  var g = gcd(w, h);
  var a = Math.round(w) / g, b = Math.round(h) / g;
  var dec = (w / h).toFixed(2) + ':1';
  /* 1366 x 768 reduces to 683:384, which tells nobody anything */
  if (a > 64 || b > 64) return dec;
  return a + ':' + b + ' (' + dec + ')';
}

function shapeOf(aspect) { return aspect > 1.05 ? 'landscape' : (aspect < 0.95 ? 'portrait' : 'square'); }

/* Frames to a clock string. Always exact, never rounded to a neighbour frame. */
function formatTime(frames, fps) {
  var totalMs = Math.round(frames * 1000 / fps);
  var m = Math.floor(totalMs / 60000);
  var s = Math.floor((totalMs % 60000) / 1000);
  var cs = Math.floor((totalMs % 1000) / 10);
  return m + ':' + (s < 10 ? '0' : '') + s + '.' + (cs < 10 ? '0' : '') + cs;
}

/* Accepts 90, 90.5, 1:30, 1:30.5, 0:01:30 */
function parseDuration(str) {
  if (typeof str === 'number') return isFinite(str) ? str : NaN;
  var s = String(str == null ? '' : str).trim().replace(',', '.');
  if (!s) return NaN;
  if (/^\d+(\.\d+)?$/.test(s)) return parseFloat(s);
  var parts = s.split(':');
  if (parts.length < 2 || parts.length > 3) return NaN;
  var total = 0;
  for (var i = 0; i < parts.length; i++) {
    var last = i === parts.length - 1;
    if (!(last ? /^\d+(\.\d+)?$/ : /^\d+$/).test(parts[i])) return NaN;
    total = total * 60 + parseFloat(parts[i]);
  }
  return total;
}

function formatBytes(b) {
  if (!isFinite(b)) return '-';
  if (b >= 1e9) return (b / 1e9).toFixed(2) + ' GB';
  if (b >= 1e6) return (b / 1e6).toFixed(b >= 1e8 ? 0 : 1) + ' MB';
  if (b >= 1e3) return Math.round(b / 1e3) + ' KB';
  return Math.round(b) + ' B';
}

/* ------------------------------------------------------------------ *
 * Seeded randomness                                                   *
 * ------------------------------------------------------------------ */

function mulberry32(seed) {
  var a = seed >>> 0;
  return function () {
    a = (a + 0x6D2B79F5) >>> 0;
    var t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function newSeed() {
  if (typeof crypto !== 'undefined' && crypto.getRandomValues) {
    var a = new Uint32Array(1);
    crypto.getRandomValues(a);
    return a[0] >>> 0;
  }
  return (Math.floor(Math.random() * 4294967296)) >>> 0;
}

/* Fisher-Yates driven by the seed. Same input and seed, same output. */
function seededShuffle(list, seed) {
  var out = list.slice();
  var rnd = mulberry32(seed);
  for (var i = out.length - 1; i > 0; i--) {
    var j = Math.floor(rnd() * (i + 1));
    var t = out[i]; out[i] = out[j]; out[j] = t;
  }
  return out;
}

/* ------------------------------------------------------------------ *
 * Output size                                                         *
 * ------------------------------------------------------------------ */

function validateDims(w, h) {
  if (!isFinite(w) || !isFinite(h)) return 'Enter a width and a height in pixels.';
  if (Math.round(w) !== w || Math.round(h) !== h) return 'Width and height must be whole pixels.';
  if (w < MIN_DIM || h < MIN_DIM) return 'Width and height must be at least ' + MIN_DIM + ' px.';
  if (w > MAX_DIM || h > MAX_DIM) return 'The browser canvas stops at ' + MAX_DIM + ' px per side.';
  if (w * h > MAX_AREA) return 'That is more pixels than the browser canvas can hold.';
  return null;
}

function findPreset(w, h) {
  for (var i = 0; i < PRESETS.length; i++) if (PRESETS[i].w === w && PRESETS[i].h === h) return PRESETS[i];
  return null;
}

/* ------------------------------------------------------------------ *
 * Geometry                                                            *
 *                                                                     *
 * A photo's placement is the rectangle it is drawn into, measured in  *
 * frame units where the frame is 0..1 on both axes. So w = 1.2 means  *
 * the photo is drawn 1.2 frame widths wide.                           *
 *                                                                     *
 *   fill  the rectangle covers the frame      x<=0, y<=0, x+w>=1 ...  *
 *   fit   the rectangle sits inside the frame x>=0, y>=0, x+w<=1 ...  *
 *                                                                     *
 * Both conditions describe convex sets, and every motion here moves   *
 * along the straight line between a valid start and a valid end, so   *
 * every in-between position is valid too. rectValid() lets the tests  *
 * check that rather than trust it.                                    *
 * ------------------------------------------------------------------ */

function coverBase(ai, af) {
  var w0 = Math.max(1, ai / af);
  return { w0: w0, h0: w0 * af / ai };
}

function containBase(ai, af) {
  var w1 = Math.min(1, ai / af);
  return { w1: w1, h1: w1 * af / ai };
}

/* Crop described the way a person thinks about it: centre of the
   visible window in photo units (0..1) and a zoom where 1 is the
   loosest crop that still fills the frame. Returns a valid fill rect. */
function cropRect(ai, af, cx, cy, z) {
  var base = coverBase(ai, af);
  z = Math.max(1, isFinite(z) ? z : 1);
  var w = base.w0 * z, h = base.h0 * z;
  var hx = 0.5 / w, hy = 0.5 / h;
  cx = clamp(isFinite(cx) ? cx : 0.5, hx, 1 - hx);
  cy = clamp(isFinite(cy) ? cy : 0.5, hy, 1 - hy);
  return { x: 0.5 - cx * w, y: 0.5 - cy * h, w: w, h: h };
}

function rectToCrop(rect, ai, af) {
  var base = coverBase(ai, af);
  return { cx: (0.5 - rect.x) / rect.w, cy: (0.5 - rect.y) / rect.h, z: rect.w / base.w0 };
}

/* Normalise a crop so the editor can show exactly what will render. */
function clampCrop(ai, af, crop) {
  return rectToCrop(cropRect(ai, af, crop.cx, crop.cy, crop.z), ai, af);
}

/* Whole photo visible, centred. scale 1 touches the frame edges. */
function fitRect(ai, af, scale) {
  var base = containBase(ai, af);
  scale = clamp(isFinite(scale) ? scale : 1, 0.05, 1);
  var w = base.w1 * scale, h = base.h1 * scale;
  return { x: (1 - w) / 2, y: (1 - h) / 2, w: w, h: h };
}

function rectValid(rect, mode, eps) {
  eps = eps == null ? 1e-9 : eps;
  if (!(rect.w > 0 && rect.h > 0)) return false;
  if (mode === 'fill') {
    return rect.x <= eps && rect.y <= eps && rect.x + rect.w >= 1 - eps && rect.y + rect.h >= 1 - eps;
  }
  return rect.x >= -eps && rect.y >= -eps && rect.x + rect.w <= 1 + eps && rect.y + rect.h <= 1 + eps;
}

/* Share of the photo's area that is on screen for a fill rect. */
function visibleShare(rect) { return 1 / (rect.w * rect.h); }

/* Position along a segment's motion path. p runs 0..1 over the photo's
   whole time on screen. Scale is interpolated geometrically so a zoom
   looks like a constant speed however large it is, and position follows
   scale along the straight line from a to b. */
function rectAt(seg, p) {
  if (seg.periodic) p = 0.5 - 0.5 * Math.cos(2 * Math.PI * p);
  var a = seg.a, b = seg.b, s = p;
  if (Math.abs(b.w - a.w) > 1e-12) {
    var w = a.w * Math.pow(b.w / a.w, p);
    s = (w - a.w) / (b.w - a.w);
  }
  return {
    x: lerp(a.x, b.x, s), y: lerp(a.y, b.y, s),
    w: lerp(a.w, b.w, s), h: lerp(a.h, b.h, s)
  };
}

/* ------------------------------------------------------------------ *
 * Suitability                                                         *
 * ------------------------------------------------------------------ */

function framingModeFor(photo, project) {
  if (photo.framing && FRAMING_MODES.indexOf(photo.framing) >= 0) return photo.framing;
  if (project.framing.exclude) return 'fill';
  return FRAMING_MODES.indexOf(project.framing.mode) >= 0 ? project.framing.mode : 'fill';
}

function motionKindFor(photo) {
  var k = photo.motion && photo.motion.kind;
  return PHOTO_MOTIONS.indexOf(k) >= 0 ? k : 'auto';
}

/* The tightest zoom this photo's motion can reach. Deliberately based on
   the intensity setting alone and not on timing, so that excluding a
   photo never depends on how long the others are shown. */
function worstZoom(photo, project, af) {
  var m = project.motion, kind = motionKindFor(photo);
  if (kind === 'custom' && photo.motion.a && photo.motion.b) {
    var ai = photo.iw / photo.ih;
    var za = clampCrop(ai, af, photo.motion.a).z;
    if (!m.enabled) return za;
    return Math.max(za, clampCrop(ai, af, photo.motion.b).z);
  }
  if (!m.enabled || kind === 'still') return 1;
  return 1 + (INTENSITY[m.intensity] || INTENSITY.gentle).zoom;
}

/* Decide whether a photo is in the video and why.
   The rule: crop the correctly oriented photo to fill the frame, then
   tighten by the most the planned motion will zoom. If less than
   `threshold` of the photo's area is still on screen, it is unsuitable. */
function classify(photo, project) {
  var out = project.output, af = out.w / out.h, ai = photo.iw / photo.ih;
  var mode = framingModeFor(photo, project);
  var threshold = clamp(project.framing.threshold == null ? DEFAULT_THRESHOLD : project.framing.threshold, 0.05, 1);
  var z = worstZoom(photo, project, af);
  var still = Math.min(ai, af) / Math.max(ai, af);
  var visible = mode === 'fill' ? still / (z * z) : 1;
  var res = { id: photo.id, mode: mode, visible: visible, included: true, auto: true, unsuitable: false, reason: '' };
  /* Framing set by hand is the user's own decision about the crop, so it is never second guessed. */
  var chosen = motionKindFor(photo) === 'custom' && photo.motion.a && photo.motion.b;
  var unsuitable = mode === 'fill' && !chosen && visible + 1e-9 < threshold;
  res.unsuitable = unsuitable;
  if (unsuitable) res.reason = cropReason(ai, af, visible, threshold, still);

  if (photo.include === 'out') {
    res.included = false; res.auto = false;
    res.reason = 'Excluded by you.';
  } else if (photo.include === 'in') {
    res.auto = false;
    if (unsuitable) res.reason = 'Kept by you. ' + res.reason;
  } else if (project.framing.exclude && unsuitable) {
    res.included = false;
  }
  return res;
}

function cropReason(ai, af, visible, threshold, still) {
  var ps = shapeOf(ai), os = shapeOf(af), rel = ai / af, lead;
  var motionOnly = still + 1e-9 >= threshold;     /* it would pass if it were held still */
  if (rel > 0.95 && rel < 1.05) {
    lead = 'Same shape as the output, but the motion zoom alone takes it past the limit';
  } else {
    if (ps !== os) lead = ps.charAt(0).toUpperCase() + ps.slice(1) + ' photo, too much cropping for this ' + os + ' output';
    else lead = ai > af ? 'Photo is much wider than the output, too much cropping' : 'Photo is much taller than the output, too much cropping';
    if (motionOnly) lead += ' once its motion zooms in';
  }
  return lead + ' (' + Math.round(visible * 100) + '% would stay visible, minimum ' + Math.round(threshold * 100) + '%).';
}

/* ------------------------------------------------------------------ *
 * Automatic motion                                                    *
 * ------------------------------------------------------------------ */

/* One move per included photo. With variation on, the order of moves,
   pan directions and a small focus drift come from the seed. With it
   off the pattern is strictly regular. Either way the result depends
   only on (count, style, variation, seed), so preview and export agree. */
function assignMoves(n, style, variation, seed) {
  var rnd = mulberry32((seed >>> 0) ^ 0x9E3779B9);
  var moves = [], panSign = variation && rnd() < 0.5 ? -1 : 1;
  var flip = variation && rnd() < 0.5 ? 1 : 0;
  var prev = '', prevPrev = '';
  for (var i = 0; i < n; i++) {
    var kind;
    if (style === 'push' || style === 'pull' || style === 'pan') kind = style;
    else if (style === 'pushpull') kind = ((i + flip) % 2 === 0) ? 'push' : 'pull';
    else if (!variation) kind = ['push', 'pull', 'pan'][i % 3];
    else {
      /* Weighted pick, never a pan after a pan and never three alike in a row. */
      for (var tries = 0; tries < 12; tries++) {
        var r = rnd();
        kind = r < 0.4 ? 'push' : (r < 0.75 ? 'pull' : 'pan');
        if (kind === 'pan' && prev === 'pan') continue;
        if (kind === prev && kind === prevPrev) continue;
        break;
      }
      if (kind === 'pan' && prev === 'pan') kind = 'push';
      if (kind === prev && kind === prevPrev) kind = kind === 'push' ? 'pull' : 'push';
    }
    var mv = { kind: kind, dir: 1, jx: 0, jy: 0 };
    if (kind === 'pan') { mv.dir = panSign; panSign = -panSign; }
    if (variation) { mv.jx = rnd() * 2 - 1; mv.jy = rnd() * 2 - 1; }
    moves.push(mv);
    prevPrev = prev; prev = kind;
  }
  return moves;
}

function panTravel(af, axis, pan) {
  /* pan is a share of the frame's longer side, so a sideways pan and a
     vertical pan cover the same distance in pixels */
  var d = axis === 'x' ? (af >= 1 ? pan : pan / af) : (af >= 1 ? pan * af : pan);
  return Math.min(d, 0.25);
}

/* Start and end rectangles for one photo.
   spec  { kind, dir, axis, jx, jy }
   amp   { zoom, pan } already capped for this photo's time on screen */
function motionRects(ai, af, mode, spec, focus, amp) {
  var kind = spec.kind, a, b, t;
  if (mode !== 'fill') {
    /* Whole photo stays visible. Pans make no sense here, so they become pushes. */
    if (kind === 'still') { a = fitRect(ai, af, 1); return { a: a, b: a }; }
    a = fitRect(ai, af, 1 / (1 + amp.zoom));
    b = fitRect(ai, af, 1);
    if (kind === 'pull') { t = a; a = b; b = t; }
    return { a: a, b: b };
  }
  var fx = focus.x, fy = focus.y;
  if (kind === 'still') { a = cropRect(ai, af, fx, fy, 1); return { a: a, b: a }; }
  if (kind === 'pan') {
    var base = coverBase(ai, af);
    var axis = spec.axis;
    if (!axis) {
      /* Prefer the axis where the photo already overhangs the frame, so the
         pan costs no extra cropping. Otherwise pan along the longer side. */
      var freeX = base.w0 - 1 >= panTravel(af, 'x', amp.pan) - 1e-9;
      var freeY = base.h0 - 1 >= panTravel(af, 'y', amp.pan) - 1e-9;
      axis = freeX ? 'x' : (freeY ? 'y' : (af >= 1 ? 'x' : 'y'));
    }
    var d = panTravel(af, axis, amp.pan);
    var b0 = axis === 'x' ? base.w0 : base.h0;
    var z = Math.max(1, (1 + d) / b0), zMax = 1 + amp.zoom;
    if (z > zMax) { z = zMax; d = Math.max(0, b0 * z - 1); }
    var size = b0 * z, half = 0.5 / size, dc = d / size;
    var mid = clamp(axis === 'x' ? fx : fy, half + dc / 2, 1 - half - dc / 2);
    var c0 = mid - spec.dir * dc / 2, c1 = mid + spec.dir * dc / 2;
    a = cropRect(ai, af, axis === 'x' ? c0 : fx, axis === 'y' ? c0 : fy, z);
    b = cropRect(ai, af, axis === 'x' ? c1 : fx, axis === 'y' ? c1 : fy, z);
    return { a: a, b: b };
  }
  /* push or pull: loosest crop at one end, a modest zoom at the other.
     The zoom is centred on the focus point, so whatever sits there holds
     its place on screen while the edges tighten around it. The jitter adds
     a slight drift, at most half the zoom travel, so no two moves match. */
  var zEnd = 1 + amp.zoom;
  var bz = coverBase(ai, af);
  var jx = (spec.jx || 0) * 0.5 * amp.zoom / (bz.w0 * zEnd);
  var jy = (spec.jy || 0) * 0.5 * amp.zoom / (bz.h0 * zEnd);
  a = cropRect(ai, af, fx, fy, 1);
  var c1 = rectToCrop(a, ai, af);
  b = cropRect(ai, af, fx - (fx - c1.cx) / zEnd + jx, fy - (fy - c1.cy) / zEnd + jy, zEnd);
  if (kind === 'pull') { t = a; a = b; b = t; }
  return { a: a, b: b };
}

/* ------------------------------------------------------------------ *
 * Timeline                                                            *
 *                                                                     *
 * Everything is in whole frames.                                      *
 *                                                                     *
 * Each included photo owns one slot. A slot starts on the frame where *
 * its photo is fully on screen and ends on the frame where the next   *
 * photo is fully on screen. The transition into the next photo is the *
 * last `trans` frames of the slot, so it is part of the photo's       *
 * duration and never added on top. The video length is therefore the  *
 * plain sum of the slots.                                             *
 *                                                                     *
 * Looping: the timeline is periodic with period N = total frames. The *
 * last slot's transition leads into photo 0, exactly like every other *
 * transition. Frame N would be frame 0 again and is never rendered.   *
 * Photo 0's motion starts during that closing transition and carries  *
 * on across the file boundary, so the last and first encoded frames   *
 * are neighbours on one continuous move.                              *
 *                                                                     *
 * Not looping: the last photo has no transition out and the first has *
 * none in, so the file starts and ends cleanly.                       *
 * ------------------------------------------------------------------ */

function buildPlan(project) {
  var out = project.output, W = out.w, H = out.h, fps = out.fps;
  var errors = [], notes = [], warnings = [];
  var dimErr = validateDims(W, H);
  if (dimErr) errors.push({ field: 'size', msg: dimErr });
  if (FPS_CHOICES.indexOf(fps) < 0) { errors.push({ field: 'fps', msg: 'Choose 24, 25, 30, 50 or 60 fps.' }); fps = 30; }
  var af = dimErr ? 16 / 9 : W / H;
  var loop = !!project.loop;
  var t = project.timing, m = project.motion;

  /* 1. who is in */
  var classes = [], inc = [];
  for (var i = 0; i < project.photos.length; i++) {
    var c = classify(project.photos[i], project);
    classes.push(c);
    if (c.included) inc.push(i);
  }
  var n = inc.length;
  var plan = {
    ok: false, errors: errors, notes: notes, warnings: warnings,
    W: W, H: H, fps: fps, loop: loop, frames: 0, seconds: 0,
    segs: [], classes: classes,
    counts: { total: project.photos.length, included: n, excluded: project.photos.length - n },
    timing: { perPhotoFrames: 0, transitionFrames: 0, uniform: true }
  };
  if (!project.photos.length) { errors.push({ field: 'photos', msg: 'Add some photos to begin.' }); return plan; }
  if (!n) { errors.push({ field: 'photos', msg: 'No photos are left in the video.' }); return plan; }

  /* 2. slot lengths */
  var minSlot = Math.max(2, Math.ceil(MIN_SLOT_SECONDS * fps));
  var slots = new Array(n), free = [], fixed = 0, k, p, frames;
  for (k = 0; k < n; k++) {
    p = project.photos[inc[k]];
    if (p.duration != null && isFinite(p.duration)) {
      frames = Math.round(p.duration * fps);
      if (frames < minSlot) {
        errors.push({ field: 'photo', id: p.id, msg: p.name + ' is set to ' + trimNum(p.duration) + ' s. The shortest a photo can be is ' + trimNum(minSlot / fps) + ' s.' });
        frames = minSlot;
      }
      slots[k] = frames; fixed += frames;
    } else free.push(k);
  }
  if (t.mode === 'total') {
    var reqT = t.total * fps, totalFrames = Math.round(reqT);
    if (!(t.total > 0) || !isFinite(reqT)) {
      errors.push({ field: 'total', msg: 'Enter a total duration, for example 90 or 1:30.' });
      totalFrames = n * Math.round(5 * fps);
    } else if (Math.abs(reqT - totalFrames) > 1e-6) {
      notes.push(trimNum(t.total) + ' s is ' + reqT.toFixed(2) + ' frames at ' + fps + ' fps. Using ' + totalFrames + ' frames, ' + trimNum(totalFrames / fps) + ' s.');
    }
    if (!free.length) {
      notes.push('Every photo has its own duration, so the total duration setting has no effect.');
    } else {
      var rem = totalFrames - fixed;
      if (rem < free.length * minSlot) {
        var need = (fixed + free.length * minSlot) / fps;
        errors.push({ field: 'total', msg: 'Too short for ' + n + ' photos. At ' + fps + ' fps the minimum is ' + trimNum(need, 2) + ' s.' });
        rem = free.length * minSlot;
      }
      /* spread any leftover frames evenly through the run instead of piling them at one end */
      for (k = 0; k < free.length; k++) {
        slots[free[k]] = Math.floor((k + 1) * rem / free.length) - Math.floor(k * rem / free.length);
      }
      if (rem % free.length !== 0) {
        var lo = Math.floor(rem / free.length);
        notes.push('The total does not divide evenly. Photos get ' + lo + ' or ' + (lo + 1) + ' frames, ' + trimNum(lo / fps) + ' to ' + trimNum((lo + 1) / fps) + ' s.');
      }
    }
  } else {
    var reqP = t.perPhoto * fps, per = Math.round(reqP);
    if (!(t.perPhoto > 0) || !isFinite(reqP)) {
      errors.push({ field: 'perPhoto', msg: 'Enter how many seconds each photo is shown.' });
      per = Math.round(5 * fps);
    } else if (per < minSlot) {
      errors.push({ field: 'perPhoto', msg: 'The shortest a photo can be is ' + trimNum(minSlot / fps) + ' s.' });
      per = minSlot;
    } else if (Math.abs(reqP - per) > 1e-6) {
      notes.push(trimNum(t.perPhoto) + ' s is ' + reqP.toFixed(2) + ' frames at ' + fps + ' fps. Using ' + per + ' frames, ' + trimNum(per / fps) + ' s per photo.');
    }
    for (k = 0; k < free.length; k++) slots[free[k]] = per;
  }

  /* 3. transitions, each one the tail of its own slot */
  if (!(t.transition >= 0) || !isFinite(t.transition)) errors.push({ field: 'transition', msg: 'Enter a transition time in seconds, or 0 for a cut.' });
  var trans = new Array(n), types = new Array(n), clamped = 0, maxClampTo = 0;
  for (k = 0; k < n; k++) {
    p = project.photos[inc[k]];
    var ov = p.transition || null;
    var type = (ov && ov.type) || t.transitionType || 'crossfade';
    var secs = ov && ov.seconds != null && isFinite(ov.seconds) ? ov.seconds : t.transition;
    var x = type === 'cut' || !(secs > 0) ? 0 : Math.round(secs * fps);
    var cap = Math.floor(slots[k] * MAX_TRANSITION_SHARE);
    if (x > cap) { x = cap; clamped++; maxClampTo = Math.max(maxClampTo, cap); }
    if (n === 1) x = 0;                       /* a photo cannot dissolve into itself */
    if (!loop && k === n - 1) x = 0;          /* clean ending when not looping */
    trans[k] = x;
    types[k] = x > 0 ? 'crossfade' : 'cut';
  }
  if (clamped) {
    notes.push('The crossfade was shortened on ' + (clamped === n ? 'every photo' : clamped + ' of ' + n + ' photos') +
      ' because it cannot take more than half of a photo’s time. Longest now ' + trimNum(maxClampTo / fps) + ' s.');
  }

  /* 4. segments */
  var moves = assignMoves(n, MOTION_STYLES.indexOf(m.style) >= 0 ? m.style : 'mixed', !!m.variation, m.seed >>> 0);
  var amp0 = INTENSITY[m.intensity] || INTENSITY.gentle;
  var start = 0, uniform = true, shortest = Infinity;
  for (k = 0; k < n; k++) {
    p = project.photos[inc[k]];
    var cls = classes[inc[k]];
    var prevIdx = (k - 1 + n) % n;
    var transIn = n === 1 ? 0 : ((loop || k > 0) ? trans[prevIdx] : 0);
    var life = transIn + slots[k];
    var ai = p.iw / p.ih;
    var periodic = false;
    var kind = motionKindFor(p);
    var spec, rects;
    var focus = p.focus && isFinite(p.focus.x) && isFinite(p.focus.y)
      ? { x: clamp(p.focus.x, 0, 1), y: clamp(p.focus.y, 0, 1) } : DEFAULT_FOCUS;

    /* how far this photo may travel: the intensity limit, or less if it is on screen briefly */
    var travelSecs = life / fps;
    if (n === 1 && loop) travelSecs /= 2;     /* out and back in one loop */
    var amp = {
      zoom: Math.min(amp0.zoom, amp0.zoomRate * travelSecs),
      pan: Math.min(amp0.pan, amp0.panRate * travelSecs)
    };

    if (kind === 'custom' && p.motion.a && p.motion.b && cls.mode === 'fill') {
      var ca = p.motion.a, cb = m.enabled ? p.motion.b : p.motion.a;
      rects = { a: cropRect(ai, af, ca.cx, ca.cy, ca.z), b: cropRect(ai, af, cb.cx, cb.cy, cb.z) };
      spec = { kind: 'custom' };
    } else {
      if (!m.enabled || kind === 'still') spec = { kind: 'still' };
      else if (kind === 'auto' || kind === 'custom') spec = moves[k];
      else if (kind === 'push' || kind === 'pull') spec = { kind: kind, jx: 0, jy: 0 };
      else spec = {
        kind: 'pan',
        axis: (kind === 'pan-left' || kind === 'pan-right') ? 'x' : 'y',
        dir: (kind === 'pan-right' || kind === 'pan-down') ? 1 : -1
      };
      rects = motionRects(ai, af, cls.mode, spec, focus, amp);
    }
    var moving = Math.abs(rects.a.x - rects.b.x) + Math.abs(rects.a.y - rects.b.y) + Math.abs(rects.a.w - rects.b.w) > 1e-12;
    if (n === 1 && loop && moving) periodic = true;

    plan.segs.push({
      id: p.id, name: p.name, iw: p.iw, ih: p.ih, opaque: !!p.opaque,
      index: k, start: start, slot: slots[k],
      trans: trans[k], transType: types[k], transIn: transIn, life: life,
      mode: cls.mode, bg: (project.framing.bg || '#000000'),
      a: rects.a, b: rects.b, periodic: periodic, moving: moving, kind: spec.kind
    });
    start += slots[k];
    if (slots[k] !== slots[0] || trans[k] !== trans[0]) uniform = false;
    if (slots[k] < shortest) shortest = slots[k];
  }
  plan.frames = start;
  plan.seconds = start / fps;
  plan.timing = { perPhotoFrames: slots[0], transitionFrames: trans[0], uniform: uniform, shortestFrames: shortest };

  if (shortest / fps < FAST_SLOT_SECONDS) {
    warnings.push('Photos change every ' + trimNum(shortest / fps, 2) + ' s at the fastest. That is quick for a large screen.');
  }
  if (loop && n > 1 && trans[n - 1] === 0) {
    notes.push('The loop point is a cut from the last photo to the first, the same as a cut anywhere else in the sequence. Use a crossfade there if you want the loop point to dissolve.');
  }
  if (loop && n === 1 && !plan.segs[0].moving) {
    notes.push('One still photo with no motion. Every frame is identical, so it loops by definition.');
  }
  plan.ok = errors.length === 0;
  return plan;
}

function locate(plan, frame) {
  var segs = plan.segs, lo = 0, hi = segs.length - 1;
  while (lo < hi) {
    var mid = (lo + hi + 1) >> 1;
    if (segs[mid].start <= frame) lo = mid; else hi = mid - 1;
  }
  return lo;
}

/* Wrap or clamp any integer to a real frame index. */
function wrapFrame(plan, frame) {
  var N = plan.frames;
  if (!N) return 0;
  if (plan.loop) return ((frame % N) + N) % N;
  return clamp(frame, 0, N - 1);
}

/* What is on screen at a frame. Returns one or two layers, bottom first.
   Every value is derived from the frame index alone. */
function stateAt(plan, frame) {
  var k = wrapFrame(plan, Math.round(frame));
  var segs = plan.segs, n = segs.length;
  if (!n) return [];
  var i = locate(plan, k), s = segs[i];
  var local = k - s.start;
  var layers = [{ seg: s, alpha: 1, rect: rectAt(s, (local + s.transIn) / s.life), phase: (local + s.transIn) / s.life }];
  var holdEnd = s.slot - s.trans;
  if (s.trans > 0 && local >= holdEnd) {
    /* j runs 0..trans-1. The fade never sits at exactly 0 or 1 on a
       transition frame, so no frame repeats the one before or after it. */
    var j = local - holdEnd;
    var nx = segs[(i + 1) % n];
    layers.push({ seg: nx, alpha: smoothstep((j + 1) / (s.trans + 1)), rect: rectAt(nx, j / nx.life), phase: j / nx.life });
  }
  return layers;
}

/* The frame at which a photo is first fully on screen. */
function frameOfPhoto(plan, id) {
  for (var i = 0; i < plan.segs.length; i++) if (plan.segs[i].id === id) return plan.segs[i].start;
  return -1;
}

/* ------------------------------------------------------------------ *
 * Codecs                                                              *
 * ------------------------------------------------------------------ */

var CODECS = {
  avc:  { label: 'H.264',        containers: ['mp4'],         evenOnly: true },
  hevc: { label: 'H.265 (HEVC)', containers: ['mp4'],         evenOnly: true },
  vp9:  { label: 'VP9',          containers: ['webm'],        evenOnly: false },
  av1:  { label: 'AV1',          containers: ['mp4', 'webm'], evenOnly: false }
};
var CODEC_ORDER = ['avc', 'vp9', 'av1', 'hevc'];
var CONTAINERS = {
  mp4:  { label: 'MP4',  ext: 'mp4',  mime: 'video/mp4' },
  webm: { label: 'WebM', ext: 'webm', mime: 'video/webm' }
};

/* [name, idc, MaxMBPS, MaxFS, MaxBR kbit/s for Baseline and Main] */
var AVC_LEVELS = [
  ['3.1', 31, 108000, 3600, 14000],
  ['3.2', 32, 216000, 5120, 20000],
  ['4.0', 40, 245760, 8192, 20000],
  ['4.1', 41, 245760, 8192, 50000],
  ['4.2', 42, 522240, 8704, 50000],
  ['5.0', 50, 589824, 22080, 135000],
  ['5.1', 51, 983040, 36864, 240000],
  ['5.2', 52, 2073600, 36864, 240000],
  ['6.0', 60, 4177920, 139264, 240000],
  ['6.1', 61, 8355840, 139264, 480000],
  ['6.2', 62, 16711680, 139264, 800000]
];
var AVC_PROFILES = [
  { id: 'high',     label: 'High',                 hex: '6400', brScale: 1.25 },
  { id: 'main',     label: 'Main',                 hex: '4d00', brScale: 1 },
  { id: 'baseline', label: 'Constrained Baseline', hex: '42e0', brScale: 1 }
];

function avcLevel(w, h, fps, bitrate, brScale) {
  var mbw = Math.ceil(w / 16), mbh = Math.ceil(h / 16), fs = mbw * mbh, mbps = fs * fps;
  for (var i = 0; i < AVC_LEVELS.length; i++) {
    var L = AVC_LEVELS[i], side = Math.sqrt(8 * L[3]);
    if (fs <= L[3] && mbps <= L[2] && mbw <= side && mbh <= side && bitrate <= L[4] * 1000 * (brScale || 1)) return { name: L[0], idc: L[1] };
  }
  return null;
}

/* [name, idc, MaxLumaPs, MaxLumaSr, MaxBR main tier, MaxBR high tier] kbit/s */
var HEVC_LEVELS = [
  ['3.1', 93, 983040, 33177600, 10000, 0],
  ['4', 120, 2228224, 66846720, 12000, 30000],
  ['4.1', 123, 2228224, 133693440, 20000, 50000],
  ['5', 150, 8912896, 267386880, 25000, 100000],
  ['5.1', 153, 8912896, 534773760, 40000, 160000],
  ['5.2', 156, 8912896, 1069547520, 60000, 240000],
  ['6', 180, 35651584, 1069547520, 60000, 240000],
  ['6.1', 183, 35651584, 2139095040, 120000, 480000],
  ['6.2', 186, 35651584, 4278190080, 240000, 800000]
];
function hevcLevel(w, h, fps, bitrate) {
  var ps = w * h, sr = ps * fps;
  for (var i = 0; i < HEVC_LEVELS.length; i++) {
    var L = HEVC_LEVELS[i], side = Math.sqrt(8 * L[2]);
    if (ps > L[2] || sr > L[3] || w > side || h > side) continue;
    if (bitrate <= L[4] * 1000) return { name: L[0], idc: L[1], tier: 'L' };
    if (L[5] && bitrate <= L[5] * 1000) return { name: L[0], idc: L[1], tier: 'H' };
  }
  return null;
}

/* [name, id, MaxLumaSampleRate, MaxLumaPictureSize, MaxBitrate kbit/s, MaxDimension] */
var VP9_LEVELS = [
  ['3', 30, 20736000, 552960, 7200, 2048],
  ['3.1', 31, 36864000, 983040, 12000, 2752],
  ['4', 40, 83558400, 2228224, 18000, 4160],
  ['4.1', 41, 160432128, 2228224, 30000, 4160],
  ['5', 50, 311951360, 8912896, 60000, 8384],
  ['5.1', 51, 588251136, 8912896, 120000, 8384],
  ['5.2', 52, 1176502272, 8912896, 180000, 8384],
  ['6', 60, 1176502272, 35651584, 180000, 16832],
  ['6.1', 61, 2353004544, 35651584, 240000, 16832],
  ['6.2', 62, 4706009088, 35651584, 480000, 16832]
];
function vp9Level(w, h, fps, bitrate) {
  var ps = w * h, sr = ps * fps;
  for (var i = 0; i < VP9_LEVELS.length; i++) {
    var L = VP9_LEVELS[i];
    if (ps <= L[3] && sr <= L[2] && bitrate <= L[4] * 1000 && w <= L[5] && h <= L[5]) return { name: L[0], id: L[1] };
  }
  return null;
}

/* [name, seq_level_idx, MaxPicSize, MaxHSize, MaxVSize, MaxDisplayRate, Main Mbit/s, High Mbit/s] */
var AV1_LEVELS = [
  ['3.0', 4, 665856, 4352, 2448, 19975680, 6, 0],
  ['3.1', 5, 1065024, 5504, 3096, 31950720, 10, 0],
  ['4.0', 8, 2359296, 6144, 3456, 70778880, 12, 30],
  ['4.1', 9, 2359296, 6144, 3456, 141557760, 20, 50],
  ['5.0', 12, 8912896, 8192, 4352, 267386880, 30, 100],
  ['5.1', 13, 8912896, 8192, 4352, 534773760, 40, 160],
  ['5.2', 14, 8912896, 8192, 4352, 1069547520, 60, 240],
  ['6.0', 16, 35651584, 16384, 8704, 1069547520, 60, 240],
  ['6.1', 17, 35651584, 16384, 8704, 2139095040, 100, 480],
  ['6.2', 18, 35651584, 16384, 8704, 4278190080, 160, 800]
];
function av1Level(w, h, fps, bitrate) {
  var ps = w * h, dr = ps * fps;
  for (var i = 0; i < AV1_LEVELS.length; i++) {
    var L = AV1_LEVELS[i];
    if (ps > L[2] || w > L[3] || h > L[4] || dr > L[5]) continue;
    if (bitrate <= L[6] * 1e6) return { name: L[0], idx: L[1], tier: 'M' };
    if (L[7] && bitrate <= L[7] * 1e6) return { name: L[0], idx: L[1], tier: 'H' };
  }
  return null;
}

/* Codec strings worth asking the browser about, best first. Every one
   is a legal profile and level for these exact settings. An empty list
   comes with the reason. */
function codecCandidates(family, w, h, fps, bitrate) {
  var info = CODECS[family], list = [], L, i;
  if (!info) return { list: list, reason: 'Unknown codec.' };
  if (info.evenOnly && (w % 2 || h % 2)) {
    return { list: list, reason: info.label + ' needs an even width and height. ' + w + ' x ' + h + ' has an odd side.', oddDims: true };
  }
  if (family === 'avc') {
    for (i = 0; i < AVC_PROFILES.length; i++) {
      var P = AVC_PROFILES[i];
      L = avcLevel(w, h, fps, bitrate, P.brScale);
      if (L) list.push({ codec: 'avc1.' + P.hex + (L.idc < 16 ? '0' : '') + L.idc.toString(16), label: 'H.264 ' + P.label + ' L' + L.name, profile: P.id, level: L.name });
    }
    if (!list.length) return { list: list, reason: 'This size, frame rate and bitrate are beyond the largest H.264 level (6.2).' };
  } else if (family === 'hevc') {
    L = hevcLevel(w, h, fps, bitrate);
    if (!L) return { list: list, reason: 'This size, frame rate and bitrate are beyond the largest H.265 level.' };
    list.push({ codec: 'hvc1.1.6.' + L.tier + L.idc + '.B0', label: 'H.265 Main L' + L.name, profile: 'main', level: L.name });
    list.push({ codec: 'hev1.1.6.' + L.tier + L.idc + '.B0', label: 'H.265 Main L' + L.name, profile: 'main', level: L.name });
  } else if (family === 'vp9') {
    L = vp9Level(w, h, fps, bitrate);
    if (!L) return { list: list, reason: 'This size, frame rate and bitrate are beyond the largest VP9 level.' };
    list.push({ codec: 'vp09.00.' + L.id + '.08', label: 'VP9 Profile 0 L' + L.name, profile: '0', level: L.name });
  } else if (family === 'av1') {
    L = av1Level(w, h, fps, bitrate);
    if (!L) return { list: list, reason: 'This size, frame rate and bitrate are beyond the largest AV1 level.' };
    list.push({ codec: 'av01.0.' + (L.idx < 10 ? '0' : '') + L.idx + L.tier + '.08', label: 'AV1 Main L' + L.name, profile: 'main', level: L.name });
  }
  return { list: list, reason: '' };
}

/* Mbit/s at 1920 x 1080, 30 fps. Scaled for other sizes and rates below. */
var QUALITY = {
  standard: { label: 'Standard', avc: 8,  hevc: 5.5, vp9: 5.5, av1: 4.5 },
  high:     { label: 'High',     avc: 14, hevc: 9,   vp9: 10,  av1: 8 },
  max:      { label: 'Maximum',  avc: 28, hevc: 18,  vp9: 20,  av1: 16 }
};

/* More pixels need more bits, but not in proportion, and doubling the
   frame rate of slow moving stills needs well under double. */
function suggestBitrate(family, w, h, fps, quality) {
  var q = QUALITY[quality] || QUALITY.high;
  var mbps = (q[family] || q.avc) * Math.pow((w * h) / (1920 * 1080), 0.85) * Math.pow(fps / 30, 0.6);
  mbps = clamp(mbps, 0.5, 400);
  return Math.round(mbps * 10) * 100000;
}

function estimateBytes(bitrate, seconds) { return bitrate * seconds / 8 * 1.01; }

function outputFilename(plan, containerId) {
  var c = CONTAINERS[containerId] || CONTAINERS.mp4;
  return 'slideshow_' + plan.W + 'x' + plan.H + '_' + plan.fps + 'fps' + (plan.loop ? '_loop' : '') + '.' + c.ext;
}

/* ------------------------------------------------------------------ *
 * Resource estimate                                                   *
 *                                                                     *
 * Rough arithmetic to decide when a warning is worth showing. The     *
 * browser does not say how much memory it will allow, so nothing here *
 * predicts failure. It only flags settings that are heavy.            *
 * ------------------------------------------------------------------ */

function assessResources(o) {
  var px = o.W * o.H, risks = [];
  var srcPx = Math.min(o.maxPhotoPixels || px, px * 1.4);
  var frameMem = px * 4 * 2          /* canvas and transition scratch */
               + px * 4 * 2          /* frames in flight to the encoder */
               + px * 1.5 * 6;       /* the encoder's own queue, 4:2:0 */
  var sourceMem = srcPx * 4 * 3 + (o.maxPhotoPixels || 0) * 4;
  var estBytes = estimateBytes(o.bitrate || 0, o.frames / o.fps);
  var outMem = o.toDisk ? 32e6 : estBytes * 2;   /* buffer grows by doubling, then is copied once */
  var mem = frameMem + sourceMem + outMem;

  var heavy = [];
  if (px * o.fps > 3840 * 2160 * 30 * 1.01) heavy.push('rate');
  if (o.frames > 54000) heavy.push('frames');
  if (o.photoCount > 300) heavy.push('photos');
  if (heavy.length) {
    var what = o.W + ' x ' + o.H + ' at ' + o.fps + ' fps';
    if (heavy.indexOf('frames') >= 0) what += ', ' + o.frames.toLocaleString('en-GB') + ' frames';
    what += ' with ' + o.photoCount + ' photo' + (o.photoCount === 1 ? '' : 's');
    var fix = heavy.indexOf('rate') >= 0 ? 'Reducing the frame rate or resolution may improve reliability.'
            : (heavy.indexOf('frames') >= 0 ? 'A shorter video or a lower frame rate may improve reliability.'
            : 'Fewer photos may improve reliability.');
    risks.push('This export is demanding, ' + what + '. ' + fix);
  }
  if (!o.toDisk && estBytes > 1.5e9) {
    risks.push('The finished file is held in memory until you download it, and it is estimated at ' + formatBytes(estBytes) + '. ' +
      (o.canStream ? 'Turn on saving straight to disk in the advanced export settings.' : 'A lower quality or a shorter video is safer. This browser cannot save straight to disk.'));
  }
  if ((o.maxPhotoPixels || 0) > 60e6) {
    risks.push('The largest photo is ' + Math.round(o.maxPhotoPixels / 1e6) + ' megapixels. Each one is decoded in full before it is scaled down, which is slow and memory hungry.');
  }
  if (o.deviceMemory && mem > o.deviceMemory * 1e9 * 0.5) {
    risks.push('Estimated working memory is about ' + formatBytes(mem) + ' and this browser reports roughly ' + o.deviceMemory + ' GB on the machine' +
      (o.deviceMemory >= 8 ? ' (it never reports more than 8).' : '.') + ' Close other tabs before exporting.');
  }
  return { memBytes: mem, estBytes: estBytes, risks: risks };
}

return {
  FPS_CHOICES: FPS_CHOICES, PRESETS: PRESETS, INTENSITY: INTENSITY, MOTION_STYLES: MOTION_STYLES,
  PHOTO_MOTIONS: PHOTO_MOTIONS, FRAMING_MODES: FRAMING_MODES, CODECS: CODECS, CODEC_ORDER: CODEC_ORDER,
  CONTAINERS: CONTAINERS, QUALITY: QUALITY, DEFAULT_THRESHOLD: DEFAULT_THRESHOLD, DEFAULT_FOCUS: DEFAULT_FOCUS,
  MIN_SLOT_SECONDS: MIN_SLOT_SECONDS, MAX_TRANSITION_SHARE: MAX_TRANSITION_SHARE, MIN_DIM: MIN_DIM, MAX_DIM: MAX_DIM,
  clamp: clamp, lerp: lerp, gcd: gcd, smoothstep: smoothstep, trimNum: trimNum,
  aspectLabel: aspectLabel, shapeOf: shapeOf, formatTime: formatTime, parseDuration: parseDuration, formatBytes: formatBytes,
  mulberry32: mulberry32, newSeed: newSeed, seededShuffle: seededShuffle,
  validateDims: validateDims, findPreset: findPreset,
  coverBase: coverBase, containBase: containBase, cropRect: cropRect, rectToCrop: rectToCrop, clampCrop: clampCrop,
  fitRect: fitRect, rectValid: rectValid, visibleShare: visibleShare, rectAt: rectAt,
  framingModeFor: framingModeFor, motionKindFor: motionKindFor, worstZoom: worstZoom, classify: classify,
  assignMoves: assignMoves, motionRects: motionRects,
  buildPlan: buildPlan, locate: locate, wrapFrame: wrapFrame, stateAt: stateAt, frameOfPhoto: frameOfPhoto,
  avcLevel: avcLevel, hevcLevel: hevcLevel, vp9Level: vp9Level, av1Level: av1Level,
  codecCandidates: codecCandidates, suggestBitrate: suggestBitrate, estimateBytes: estimateBytes,
  outputFilename: outputFilename, assessResources: assessResources
};
});
