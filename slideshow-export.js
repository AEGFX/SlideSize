/* ============================================================
   Photo Montage export  (SlideshowExport)

   Renders every frame of a plan in order and feeds it to the
   browser's WebCodecs encoder, then muxes the result to MP4 or
   WebM. Nothing is recorded in real time, so a slow machine takes
   longer but never drops or mistimes a frame.

   run() works in a worker or on the main thread. launch() prefers
   the worker and falls back to the page when a worker cannot be
   used, for example when the page is opened from a file.

   Needs slideshow-core.js and slideshow-render.js loaded first.
   mp4-muxer.js or webm-muxer.js is loaded on demand.

   AEGFX / SlideSize
   ============================================================ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(root, require('./slideshow-core.js'), require('./slideshow-render.js'));
  else root.SlideshowExport = factory(root, root.SlideshowCore, root.SlideshowRender);
})(typeof self !== 'undefined' ? self : globalThis, function (root, Core, Render) {
'use strict';

var MAX_ENCODE_QUEUE = 3;     /* frames waiting in the encoder before we stop feeding it */
var YIELD_EVERY_MS = 30;      /* let cancel messages and repaints through */
var PROGRESS_EVERY_MS = 120;
var LIBS = { mp4: 'mp4-muxer.js', webm: 'webm-muxer.js', prores: 'prores-encoder-parallel.min.js' };

function inWorker() { return typeof importScripts === 'function' && typeof document === 'undefined'; }

function abortError() {
  var e = new Error('Export cancelled.');
  e.name = 'AbortError';
  return e;
}

/* ------------------------------------------------------------------ *
 * What can this browser do                                            *
 * ------------------------------------------------------------------ */

function environment() {
  var hasEncoder = typeof VideoEncoder !== 'undefined' && typeof VideoFrame !== 'undefined';
  var secure = typeof isSecureContext === 'undefined' ? true : isSecureContext;
  var reason = '';
  if (!hasEncoder) {
    reason = secure
      ? 'This browser has no WebCodecs video encoder. Use a current Chrome or Edge on a desktop computer.'
      : 'Video encoding only works on a secure page. Open this page over https or from localhost.';
  }
  return {
    canEncode: hasEncoder,
    reason: reason,
    canStreamToDisk: typeof root.showSaveFilePicker === 'function',
    deviceMemory: (typeof navigator !== 'undefined' && navigator.deviceMemory) || 0,
    cores: (typeof navigator !== 'undefined' && navigator.hardwareConcurrency) || 0
  };
}

function encoderConfig(codec, family, o, hw) {
  var cfg = {
    codec: codec, width: o.W, height: o.H,
    bitrate: o.bitrate, framerate: o.fps,
    latencyMode: 'quality',
    bitrateMode: o.bitrateMode || 'variable',
    hardwareAcceleration: hw || 'no-preference'
  };
  /* length-prefixed samples with the parameter sets out of band, which is what MP4 stores */
  if (family === 'avc') cfg.avc = { format: 'avc' };
  if (family === 'hevc') cfg.hevc = { format: 'hevc' };
  return cfg;
}

async function firstSupported(family, list, o, hw) {
  for (var i = 0; i < list.length; i++) {
    var cfg = encoderConfig(list[i].codec, family, o, hw);
    try {
      var res = await VideoEncoder.isConfigSupported(cfg);
      if (res && res.supported) return { config: cfg, candidate: list[i] };
    } catch (e) { /* a config the browser cannot even parse is simply not supported */ }
  }
  return null;
}

/* Ask the browser about one codec with the real size, frame rate and
   bitrate. A general "Chrome supports H.264" is not good enough: the
   answer changes with resolution and with the machine's hardware.
   o = { W, H, fps, bitrate, bitrateMode, hw } */
async function probe(family, o) {
  var info = Core.CODECS[family];
  var out = { family: family, supported: false, reason: '', config: null, candidate: null, oddDims: false };
  if (family === 'prores') return probeProRes(o, out);
  var env = environment();
  if (!env.canEncode) { out.reason = env.reason; return out; }
  var cand = Core.codecCandidates(family, o.W, o.H, o.fps, o.bitrate);
  if (!cand.list.length) { out.reason = cand.reason; out.oddDims = !!cand.oddDims; return out; }
  var hit = await firstSupported(family, cand.list, o, o.hw || 'no-preference');
  if (hit) { out.supported = true; out.config = hit.config; out.candidate = hit.candidate; return out; }

  var plain = (o.bitrateMode || 'variable') === 'variable' && (o.hw || 'no-preference') === 'no-preference';
  if (!plain) {
    var base = await firstSupported(family, cand.list, { W: o.W, H: o.H, fps: o.fps, bitrate: o.bitrate, bitrateMode: 'variable' }, 'no-preference');
    if (base) {
      out.reason = info.label + ' is available at this size, but not with the chosen bitrate mode or encoder preference. Reset those under Advanced.';
      return out;
    }
  }
  out.reason = info.label + ' cannot be encoded by this browser at ' + o.W + ' x ' + o.H + ', ' + o.fps + ' fps.';
  return out;
}

/* ProRes is encoded in WebAssembly rather than by the browser, so the
   question is only whether WebAssembly and workers are there and the
   size is within what the encoder handles. */
function probeProRes(o, out) {
  if (typeof WebAssembly !== 'object' || typeof Worker === 'undefined') {
    out.reason = 'ProRes is encoded with WebAssembly in workers, and this browser has neither.';
    return out;
  }
  if (o.W > 8192 || o.H > 8192) {
    out.reason = 'The ProRes encoder here stops at 8192 px a side.';
    return out;
  }
  var p = Core.PRORES_PROFILES[o.proresProfile == null ? 3 : o.proresProfile];
  out.supported = true;
  out.config = { codec: 'prores', bitrate: Core.proResBitrate(p.id, o.W, o.H, o.fps) };
  out.candidate = { label: 'Apple ProRes ' + p.label, profile: p.id };
  return out;
}

/* ------------------------------------------------------------------ *
 * Helpers for the render loop                                         *
 * ------------------------------------------------------------------ */

function loadLib(url) {
  if (inWorker()) { importScripts(url); return Promise.resolve(); }
  return new Promise(function (resolve, reject) {
    var s = document.createElement('script');
    s.src = url;
    s.onload = function () { resolve(); };
    s.onerror = function () { reject(new Error('Could not load ' + url + '. Check the connection and try again.')); };
    document.head.appendChild(s);
  });
}

/* A turn of the event loop that is not subject to timer clamping. */
var yieldChannel = null, yieldWaiting = [];
function yieldTask() {
  if (typeof MessageChannel === 'undefined') return new Promise(function (r) { setTimeout(r, 0); });
  if (!yieldChannel) {
    yieldChannel = new MessageChannel();
    yieldChannel.port1.onmessage = function () { var r = yieldWaiting.shift(); if (r) r(); };
  }
  return new Promise(function (r) { yieldWaiting.push(r); yieldChannel.port2.postMessage(0); });
}

function waitForEncoder(encoder) {
  return new Promise(function (resolve) {
    var done = false;
    function fin() {
      if (done) return;
      done = true;
      encoder.removeEventListener('dequeue', fin);
      clearTimeout(timer);
      resolve();
    }
    encoder.addEventListener('dequeue', fin);
    var timer = setTimeout(fin, 100);
  });
}

function releaseAsset(a) {
  if (!a) return;
  try { if (a.bmp && a.bmp.close) a.bmp.close(); } catch (e) {}
  if (a.blur) { try { a.blur.width = a.blur.height = 0; } catch (e) {} }
}

/* Decode one photo for export, no larger than it will ever be drawn. */
async function decodePhoto(seg, rec, W) {
  if (!rec || !rec.file) throw new Error('The file for ' + seg.name + ' is no longer available.');
  var full;
  try {
    full = await createImageBitmap(rec.file, { imageOrientation: 'from-image' });
  } catch (e) {
    throw new Error('Could not decode ' + seg.name + ' for export. The file may have moved or changed since it was added.');
  }
  var aspect = full.width / full.height, expected = seg.iw / seg.ih;
  if (Math.abs(aspect / expected - 1) > 0.01) {
    full.close();
    throw new Error(seg.name + ' decoded at a different shape than when it was added. Remove it and add it again.');
  }
  var bmp = full;
  var needW = Math.ceil(Math.max(seg.a.w, seg.b.w) * W);
  if (needW >= 1 && full.width > needW * 1.02) {
    var needH = Math.max(1, Math.round(needW * full.height / full.width));
    try {
      bmp = await createImageBitmap(full, { resizeWidth: needW, resizeHeight: needH, resizeQuality: 'high' });
      full.close();
    } catch (e) { bmp = full; }
  }
  return { bmp: bmp, blur: seg.mode === 'blur' ? Render.blurSource(rec.blur) : null };
}

/* ------------------------------------------------------------------ *
 * Colour                                                              *
 *                                                                     *
 * Left to itself the browser turns canvas pixels into video with the  *
 * BT.601 matrix and tags the file to match, whatever the size. That   *
 * is the standard definition matrix. On an HD picture the VT          *
 * Inspector flags it, because players that ignore the tag assume      *
 * BT.709 and shift the colour. So the conversion is done here, with   *
 * the matrix a player expects for the picture size, and the encoder   *
 * is handed finished 4:2:0 frames tagged to match.                    *
 * ------------------------------------------------------------------ */

function isHdSize(W, H) { return W >= 1280 || H > 576; }

function createYuvConverter(W, H) {
  var hd = isHdSize(W, H);
  var kr = hd ? 0.2126 : 0.299, kb = hd ? 0.0722 : 0.114, kg = 1 - kr - kb;
  var S = 65536, ys = 219 / 255 * S, cs = 224 / 255 * S;
  var yr = Math.round(kr * ys), yg = Math.round(kg * ys), yb = Math.round(kb * ys);
  var ur = Math.round(-0.5 * kr / (1 - kb) * cs), ug = Math.round(-0.5 * kg / (1 - kb) * cs), ub = Math.round(0.5 * cs);
  var vr = Math.round(0.5 * cs), vg = Math.round(-0.5 * kg / (1 - kr) * cs), vb = Math.round(-0.5 * kb / (1 - kr) * cs);
  var yOff = Math.round(16.5 * S), cOff = Math.round(128.5 * 4 * S);
  var cw = (W + 1) >> 1, ch = (H + 1) >> 1;
  var out = new Uint8Array(W * H + 2 * cw * ch);
  var name = hd ? 'bt709' : 'smpte170m';
  return {
    hd: hd,
    buffer: out,
    colorSpace: { primaries: name, transfer: name, matrix: name, fullRange: false },
    /* src is 8 bit RGBA or BGRA rows of `stride` bytes */
    convert: function (src, stride, bgr) {
      var ro = bgr ? 2 : 0, bo = bgr ? 0 : 2;
      var x, y, i, o = 0;
      for (y = 0; y < H; y++) {
        i = y * stride;
        for (x = 0; x < W; x++, i += 4) out[o++] = (yr * src[i + ro] + yg * src[i + 1] + yb * src[i + bo] + yOff) >> 16;
      }
      /* chroma is the average of each 2 x 2 block, edge pixels repeated on odd sizes */
      var uo = W * H, vo = uo + cw * ch;
      for (y = 0; y < ch; y++) {
        var r0 = (y * 2) * stride, r1 = Math.min(H - 1, y * 2 + 1) * stride;
        for (x = 0; x < cw; x++) {
          var a = x * 8, b = Math.min(W - 1, x * 2 + 1) * 4;
          var i00 = r0 + a, i01 = r0 + b, i10 = r1 + a, i11 = r1 + b;
          var r = src[i00 + ro] + src[i01 + ro] + src[i10 + ro] + src[i11 + ro];
          var g = src[i00 + 1] + src[i01 + 1] + src[i10 + 1] + src[i11 + 1];
          var bl = src[i00 + bo] + src[i01 + bo] + src[i10 + bo] + src[i11 + bo];
          out[uo++] = (ur * r + ug * g + ub * bl + cOff) >> 18;
          out[vo++] = (vr * r + vg * g + vb * bl + cOff) >> 18;
        }
      }
      return out;
    }
  };
}

/* Pull the canvas pixels back out. A VideoFrame copy reads the picture
   straight from where the browser drew it, GPU included, without turning
   the canvas into a slow one the way repeated getImageData calls do. */
async function readCanvas(canvas, ctx, W, H, state) {
  if (!state.useImageData) {
    var vf = null;
    try {
      vf = new VideoFrame(canvas, { timestamp: 0 });
      var fmt = vf.format;
      if ((fmt === 'RGBA' || fmt === 'RGBX' || fmt === 'BGRA' || fmt === 'BGRX') && vf.codedWidth === W && vf.codedHeight === H) {
        var size = vf.allocationSize();
        if (!state.rgba || state.rgba.length < size) state.rgba = new Uint8Array(size);
        var layout = await vf.copyTo(state.rgba);
        return { data: state.rgba, stride: (layout && layout[0] && layout[0].stride) || W * 4, offset: (layout && layout[0] && layout[0].offset) || 0, bgr: fmt.charAt(0) === 'B' };
      }
      state.useImageData = true;
    } catch (e) {
      state.useImageData = true;
    } finally {
      if (vf) vf.close();
    }
  }
  var img = ctx.getImageData(0, 0, W, H);
  return { data: img.data, stride: W * 4, offset: 0, bgr: false };
}

/* ------------------------------------------------------------------ *
 * Sinks: where rendered frames go                                     *
 *                                                                     *
 * Both take a finished canvas per frame and hand back a result at the *
 * end. The render loop in run() does not care which one it feeds.     *
 * ------------------------------------------------------------------ */

/* WebCodecs encoder into mp4-muxer or webm-muxer. */
async function webCodecsSink(job, W, H, fps, N, stream, state) {
  if (typeof VideoEncoder === 'undefined') throw new Error('This browser has no WebCodecs video encoder.');
  var libName = job.container === 'mp4' ? 'Mp4Muxer' : 'WebMMuxer';
  if (!root[libName]) await loadLib(LIBS[job.container]);
  var Lib = root[libName];
  if (!Lib) throw new Error('The ' + job.container.toUpperCase() + ' writer did not load.');
  var target = stream ? new Lib.FileSystemWritableFileStreamTarget(stream) : new Lib.ArrayBufferTarget();
  var muxer = job.container === 'mp4'
    ? new Lib.Muxer({
        target: target,
        /* No frameRate here on purpose. The default 57600 timescale divides evenly
           by every frame rate on offer, so each frame gets the same exact duration. */
        video: { codec: job.family, width: W, height: H },
        /* index at the front of the file either way, so players can start at once */
        fastStart: stream ? { expectedVideoChunks: N + 16 } : 'in-memory',
        firstTimestampBehavior: 'strict'
      })
    : new Lib.Muxer({
        target: target,
        video: { codec: job.family === 'av1' ? 'V_AV1' : 'V_VP9', width: W, height: H, frameRate: fps },
        type: 'webm',
        firstTimestampBehavior: 'strict'
      });
  var chunks = 0, codec = '';
  var encoder = new VideoEncoder({
    output: function (chunk, meta) {
      try {
        if (meta && meta.decoderConfig && meta.decoderConfig.codec) codec = meta.decoderConfig.codec;
        muxer.addVideoChunk(chunk, meta);
        chunks++;
        state.written += chunk.byteLength;
      } catch (e) { state.failure = state.failure || e; }
    },
    error: function (e) {
      state.failure = state.failure || new Error('The encoder stopped. ' + (e && e.message ? e.message : e));
    }
  });
  encoder.configure(job.encoder);
  var frameUs = 1e6 / fps, keyEvery = Math.max(1, job.keyFrames | 0);
  var yuv = createYuvConverter(W, H), grab = { rgba: null, useImageData: false };

  return {
    add: async function (canvas, ctx, k) {
      var px = await readCanvas(canvas, ctx, W, H, grab);
      state.check();
      yuv.convert(px.offset ? px.data.subarray(px.offset) : px.data, px.stride, px.bgr);
      /* Frame k sits at exactly k / fps. Nothing here reads a clock. */
      var frame = new VideoFrame(yuv.buffer, {
        format: 'I420', codedWidth: W, codedHeight: H,
        timestamp: Math.round(k * frameUs), duration: Math.round(frameUs),
        colorSpace: yuv.colorSpace
      });
      try { encoder.encode(frame, { keyFrame: k % keyEvery === 0 }); }
      finally { frame.close(); }
      /* Backpressure: never let more than a few frames pile up in the encoder. */
      while (encoder.encodeQueueSize > MAX_ENCODE_QUEUE) { await waitForEncoder(encoder); state.check(); }
    },
    finish: async function () {
      await encoder.flush();
      state.check();
      encoder.close();
      if (chunks !== N) throw new Error('The encoder returned ' + chunks + ' frames for ' + N + ' sent. The file was not written.');
      muxer.finalize();
      var out = { frames: chunks, codec: codec || job.encoder.codec, colour: yuv.hd ? 'BT.709' : 'BT.601', readback: grab.useImageData ? 'canvas' : 'frame copy' };
      if (!stream) out.buffer = target.buffer;
      return out;
    },
    close: function () { if (encoder.state !== 'closed') { try { encoder.close(); } catch (e) {} } }
  };
}

/* Apple ProRes through prores-wasm-encoder. ProRes is intra frame only,
   so the library spreads frames across its own workers and puts them
   back in order. Each finished frame becomes a Blob part, which the
   browser can page out to disk, so a long export never needs one huge
   buffer. The file is assembled from those parts at the end. */
async function proResSink(job, W, H, fps, N, stream, state) {
  if (typeof WebAssembly !== 'object') throw new Error('This browser cannot run WebAssembly, which the ProRes encoder needs.');
  if (!root.ProResParallel) await loadLib(LIBS.prores);
  var PR = root.ProResParallel;
  if (!PR) throw new Error('The ProRes encoder did not load.');
  var parts = [], frames = 0;
  var pool = await PR.createProResEncoderPool({
    width: W, height: H, frameRate: fps, profile: job.proresProfile,
    onFrameData: function (chunk) {
      frames++;
      state.written += chunk.byteLength;
      parts.push(new Blob([chunk]));
    }
  });
  return {
    add: async function (canvas) {
      await pool.addFrameFromCanvas(canvas);
      state.check();
    },
    finish: async function () {
      var ends = await pool.finalizeStreaming();
      state.check();
      if (frames !== N) throw new Error('The ProRes encoder returned ' + frames + ' frames for ' + N + ' sent. The file was not written.');
      var file = new Blob([ends.header].concat(parts, [ends.moov]), { type: 'video/quicktime' });
      parts = [];
      var out = { frames: frames, codec: 'ProRes ' + PRORES_NAMES[job.proresProfile], colour: 'BT.709', readback: 'canvas' };
      if (stream) await file.stream().pipeTo(stream, { preventClose: true });
      else out.blob = file;
      return out;
    },
    close: function () { try { pool.destroy(); } catch (e) {} }
  };
}

var PRORES_NAMES = ['422 Proxy', '422 LT', '422', '422 HQ'];

/* ------------------------------------------------------------------ *
 * The export                                                          *
 *                                                                     *
 * job = {                                                             *
 *   plan,                      from SlideshowCore.buildPlan           *
 *   photos: { id: { file, blur } },                                   *
 *   family, container,         'avc' 'hevc' 'vp9' 'av1' or 'prores'   *
 *   encoder,                   VideoEncoder config from probe()       *
 *   proresProfile,             0 to 3, for ProRes                     *
 *   keyFrames,                 key frame every this many frames       *
 *   target: { kind: 'memory' } | { kind: 'file', handle }             *
 * }                                                                   *
 * hooks = { progress(info), isCancelled() }                           *
 * ------------------------------------------------------------------ */

async function run(job, hooks) {
  var plan = job.plan, W = plan.W, H = plan.H, fps = plan.fps, N = plan.frames;
  var stream = null, sink = null;
  var canvas = null, ctx = null, scratchCanvas = null, scratchCtx = null;
  var cache = new Map();
  var began = Date.now();
  var state = {
    failure: null, written: 0,
    check: function () {
      if (state.failure) throw state.failure;
      if (hooks.isCancelled && hooks.isCancelled()) throw abortError();
    }
  };
  function getScratch() {
    if (!scratchCtx) {
      scratchCanvas = Render.makeCanvas(W, H);
      scratchCtx = scratchCanvas.getContext('2d', { alpha: false });
    }
    return scratchCtx;
  }

  try {
    hooks.progress({ stage: 'prepare' });
    if (!(N > 0)) throw new Error('There is nothing to export.');
    canvas = Render.makeCanvas(W, H);
    ctx = canvas.getContext('2d', { alpha: false });
    if (!ctx) throw new Error('The browser could not create a ' + W + ' x ' + H + ' canvas.');
    if (job.target && job.target.kind === 'file') stream = await job.target.handle.createWritable();
    sink = job.family === 'prores'
      ? await proResSink(job, W, H, fps, N, stream, state)
      : await webCodecsSink(job, W, H, fps, N, stream, state);

    var lastYield = Date.now(), lastProgress = 0;
    for (var k = 0; k < N; k++) {
      state.check();
      var layers = Core.stateAt(plan, k), need = {}, i;
      for (i = 0; i < layers.length; i++) need[layers[i].seg.id] = true;
      /* Only the photos on screen right now are held decoded. */
      cache.forEach(function (asset, id) { if (!need[id]) { releaseAsset(asset); cache.delete(id); } });
      for (i = 0; i < layers.length; i++) {
        var seg = layers[i].seg;
        if (!cache.has(seg.id)) {
          cache.set(seg.id, await decodePhoto(seg, job.photos[seg.id], W));
          state.check();
        }
      }
      Render.renderFrame(ctx, W, H, plan, k, function (s) { return cache.get(s.id); }, getScratch);
      await sink.add(canvas, ctx, k);

      var now = Date.now();
      if (now - lastProgress >= PROGRESS_EVERY_MS || k === N - 1) {
        lastProgress = now;
        hooks.progress({ stage: 'encode', frame: k + 1, frames: N, elapsedMs: now - began, written: state.written });
      }
      if (now - lastYield >= YIELD_EVERY_MS) { await yieldTask(); lastYield = Date.now(); }
    }

    hooks.progress({ stage: 'finalize' });
    var result = await sink.finish();
    result.elapsedMs = Date.now() - began;
    if (stream) {
      await stream.close();
      stream = null;
      result.kind = 'file';
    } else {
      result.kind = 'memory';
    }
    return result;
  } catch (err) {
    if (stream) { try { await stream.abort(); } catch (e) {} stream = null; }
    throw err;
  } finally {
    if (sink) sink.close();
    cache.forEach(releaseAsset);
    cache.clear();
    if (canvas) { try { canvas.width = canvas.height = 0; } catch (e) {} }
    if (scratchCanvas) { try { scratchCanvas.width = scratchCanvas.height = 0; } catch (e) {} }
  }
}

/* ------------------------------------------------------------------ *
 * Start an export from the page                                       *
 *                                                                     *
 * cb = { progress(info), done(result), cancelled(), error(err),       *
 *        mode(name) }                                                 *
 * Returns { cancel() }.                                               *
 * ------------------------------------------------------------------ */

function launch(job, cb, opts) {
  opts = opts || {};
  var cancelled = false, worker = null, settled = false, killTimer = null;

  function finish(fn, arg) {
    if (settled) return;
    settled = true;
    if (killTimer) clearTimeout(killTimer);
    if (worker) { try { worker.terminate(); } catch (e) {} worker = null; }
    fn(arg);
  }
  function onMain() {
    if (cb.mode) cb.mode('page');
    run(job, { progress: cb.progress, isCancelled: function () { return cancelled; } }).then(
      function (res) { finish(cb.done, res); },
      function (err) { if (err && err.name === 'AbortError') finish(cb.cancelled); else finish(cb.error, err); }
    );
  }

  var canWorker = !opts.forceMainThread && typeof Worker !== 'undefined' && typeof OffscreenCanvas !== 'undefined';
  if (canWorker) {
    try { worker = new Worker(opts.workerUrl || 'slideshow-worker.js'); } catch (e) { worker = null; }
  }
  if (!worker) { onMain(); }
  else {
    var heard = false;
    worker.onmessage = function (ev) {
      var m = ev.data || {};
      heard = true;
      if (m.type === 'ready') { if (cb.mode) cb.mode('worker'); }
      else if (m.type === 'unsupported') { try { worker.terminate(); } catch (e) {} worker = null; onMain(); }
      else if (m.type === 'progress') cb.progress(m.info);
      else if (m.type === 'done') finish(cb.done, m.result);
      else if (m.type === 'cancelled') finish(cb.cancelled);
      else if (m.type === 'error') { var e = new Error(m.message); e.name = m.name || 'Error'; finish(cb.error, e); }
    };
    worker.onerror = function (ev) {
      if (ev && ev.preventDefault) ev.preventDefault();
      if (!heard) {
        /* the worker script itself did not start, so do the work here instead */
        try { worker.terminate(); } catch (e) {}
        worker = null;
        onMain();
      } else {
        finish(cb.error, new Error('The export worker failed. ' + ((ev && ev.message) || '')));
      }
    };
    worker.postMessage({ type: 'start', job: job });
  }

  return {
    cancel: function () {
      if (settled || cancelled) return;
      cancelled = true;
      if (worker) {
        try { worker.postMessage({ type: 'cancel' }); } catch (e) {}
        /* a worker stuck inside a long decode still gets stopped */
        killTimer = setTimeout(function () { finish(cb.cancelled); }, 4000);
      }
    }
  };
}

return {
  LIBS: LIBS, environment: environment, encoderConfig: encoderConfig, probe: probe,
  createYuvConverter: createYuvConverter, isHdSize: isHdSize,
  run: run, launch: launch
};
});
