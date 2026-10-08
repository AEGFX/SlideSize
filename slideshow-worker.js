/* ============================================================
   Photo Montage export worker

   Runs the whole render and encode loop off the page so the
   interface stays responsive and a background tab does not slow
   the export to a crawl.

   AEGFX / SlideSize
   ============================================================ */
'use strict';
importScripts('slideshow-core.js', 'slideshow-render.js', 'slideshow-export.js');

var cancelled = false;

self.onmessage = function (ev) {
  var m = ev.data || {};
  if (m.type === 'cancel') { cancelled = true; return; }
  if (m.type !== 'start') return;

  var canDraw = false;
  try { canDraw = typeof OffscreenCanvas !== 'undefined' && !!new OffscreenCanvas(2, 2).getContext('2d'); } catch (e) {}
  var needsEncoder = !m.job || m.job.family !== 'prores';
  if ((needsEncoder && (typeof VideoEncoder === 'undefined' || typeof VideoFrame === 'undefined')) || !canDraw) {
    self.postMessage({ type: 'unsupported' });
    return;
  }

  cancelled = false;
  self.postMessage({ type: 'ready' });
  SlideshowExport.run(m.job, {
    progress: function (info) { self.postMessage({ type: 'progress', info: info }); },
    isCancelled: function () { return cancelled; }
  }).then(function (result) {
    self.postMessage({ type: 'done', result: result }, result.buffer ? [result.buffer] : []);
  }, function (err) {
    if (err && err.name === 'AbortError') self.postMessage({ type: 'cancelled' });
    else self.postMessage({ type: 'error', name: (err && err.name) || 'Error', message: String((err && err.message) || err) });
  });
};
