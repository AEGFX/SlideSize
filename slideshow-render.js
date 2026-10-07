/* ============================================================
   Slideshow Builder renderer  (SlideshowRender)

   Draws one frame of a plan onto a 2D canvas context. The preview
   calls it with a small canvas and proxy images, the exporter with
   a full size canvas and full resolution images. All geometry is
   in frame units, so both produce the same composition.

   Needs slideshow-core.js loaded first.

   AEGFX / SlideSize
   ============================================================ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./slideshow-core.js'));
  else root.SlideshowRender = factory(root.SlideshowCore);
})(typeof self !== 'undefined' ? self : this, function (Core) {
'use strict';

var BLUR_EDGE = 64;        /* long edge of the tiny image behind "blurred background" */
var BLUR_OVERSCAN = 1.12;  /* drawn a little large so the soft edges fall outside the frame */
var BLUR_DIM = 0.38;       /* darkened so the photo in front reads clearly */

/* Draw one photo as a complete layer: background first when the photo
   does not fill the frame, then the photo at its rectangle. Coordinates
   stay fractional on purpose. A slow move shifts the image by a fraction
   of a pixel per frame and rounding here is what makes cheap slideshows
   judder. */
function drawPhoto(ctx, W, H, layer, asset) {
  var seg = layer.seg, r = layer.rect;
  if (seg.mode !== 'fill') {
    if (seg.mode === 'blur' && asset.blur) {
      var base = Core.coverBase(seg.iw / seg.ih, W / H);
      var bw = base.w0 * BLUR_OVERSCAN * W, bh = base.h0 * BLUR_OVERSCAN * H;
      ctx.drawImage(asset.blur, (W - bw) / 2, (H - bh) / 2, bw, bh);
      ctx.fillStyle = 'rgba(0,0,0,' + BLUR_DIM + ')';
      ctx.fillRect(0, 0, W, H);
    } else {
      ctx.fillStyle = seg.bg || '#000000';
      ctx.fillRect(0, 0, W, H);
    }
  }
  ctx.drawImage(asset.bmp, r.x * W, r.y * H, r.w * W, r.h * H);
}

/* Render plan frame `frame` into ctx (W x H).
   getAsset(seg) returns { bmp, blur } or null if not decoded yet.
   getScratch() returns a second W x H 2D context, created on demand.
   Returns the ids of any photos that were not ready. */
function renderFrame(ctx, W, H, plan, frame, getAsset, getScratch) {
  var layers = Core.stateAt(plan, frame), missing = [];
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.globalAlpha = 1;
  ctx.globalCompositeOperation = 'source-over';
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.fillStyle = '#000000';
  ctx.fillRect(0, 0, W, H);
  for (var i = 0; i < layers.length; i++) {
    var L = layers[i], asset = getAsset(L.seg);
    if (!asset) { missing.push(L.seg.id); continue; }
    if (i === 0) {
      drawPhoto(ctx, W, H, L, asset);
    } else if (L.seg.mode === 'fill' && L.seg.opaque) {
      /* An opaque photo covering the whole frame can be blended in one draw. */
      ctx.globalAlpha = L.alpha;
      drawPhoto(ctx, W, H, L, asset);
      ctx.globalAlpha = 1;
    } else {
      /* Anything with a background or transparency is built whole on the
         scratch canvas first. Fading its parts in one by one would let the
         outgoing photo show through twice. */
      var s = getScratch();
      s.setTransform(1, 0, 0, 1, 0, 0);
      s.globalAlpha = 1;
      s.globalCompositeOperation = 'source-over';
      s.imageSmoothingEnabled = true;
      s.imageSmoothingQuality = 'high';
      s.fillStyle = '#000000';
      s.fillRect(0, 0, W, H);
      drawPhoto(s, W, H, L, asset);
      ctx.globalAlpha = L.alpha;
      ctx.drawImage(s.canvas, 0, 0);
      ctx.globalAlpha = 1;
    }
  }
  return missing;
}

/* Three passes of a box blur come close to a Gaussian. Done in plain
   arithmetic so every browser gives the same result, with or without
   canvas filter support. */
function boxBlur(px, w, h, radius, passes) {
  var tmp = new Uint8ClampedArray(px.length), win = radius * 2 + 1;
  for (var p = 0; p < passes; p++) {
    blurPass(px, tmp, w, h, radius, win, true);
    blurPass(tmp, px, w, h, radius, win, false);
  }
  return px;
}
function blurPass(src, dst, w, h, radius, win, horizontal) {
  var outer = horizontal ? h : w, inner = horizontal ? w : h;
  for (var o = 0; o < outer; o++) {
    for (var c = 0; c < 4; c++) {
      var sum = 0, i, idx;
      for (i = -radius; i <= radius; i++) {
        idx = Math.min(inner - 1, Math.max(0, i));
        sum += src[((horizontal ? o * w + idx : idx * w + o) << 2) + c];
      }
      for (i = 0; i < inner; i++) {
        dst[((horizontal ? o * w + i : i * w + o) << 2) + c] = sum / win;
        var add = Math.min(inner - 1, i + radius + 1), sub = Math.max(0, i - radius);
        sum += src[((horizontal ? o * w + add : add * w + o) << 2) + c] - src[((horizontal ? o * w + sub : sub * w + o) << 2) + c];
      }
    }
  }
}

function makeCanvas(w, h) {
  if (typeof OffscreenCanvas !== 'undefined') return new OffscreenCanvas(w, h);
  var c = document.createElement('canvas');
  c.width = w; c.height = h;
  return c;
}

/* Tiny blurred copy of a photo, as plain pixels. Made once at import and
   handed to preview and export alike, so both show the same background. */
function makeBlurData(source, iw, ih) {
  var s = BLUR_EDGE / Math.max(iw, ih);
  var w = Math.max(2, Math.round(iw * s)), h = Math.max(2, Math.round(ih * s));
  var c = makeCanvas(w, h), ctx = c.getContext('2d', { willReadFrequently: true });
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.fillStyle = '#000000';
  ctx.fillRect(0, 0, w, h);
  ctx.drawImage(source, 0, 0, w, h);
  var img = ctx.getImageData(0, 0, w, h);
  boxBlur(img.data, w, h, 3, 3);
  c.width = c.height = 0;
  return img;
}

/* Turn stored blur pixels back into something drawImage accepts. */
function blurSource(imageData) {
  if (!imageData) return null;
  var c = makeCanvas(imageData.width, imageData.height);
  c.getContext('2d').putImageData(imageData, 0, 0);
  return c;
}

return {
  BLUR_EDGE: BLUR_EDGE,
  drawPhoto: drawPhoto, renderFrame: renderFrame,
  boxBlur: boxBlur, makeCanvas: makeCanvas, makeBlurData: makeBlurData, blurSource: blurSource
};
});
