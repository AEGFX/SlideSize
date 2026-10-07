// Frame level analysis of exported videos with ffmpeg. Used by the acceptance run.
const { execFileSync } = require('child_process');
const fs = require('fs');
const sharp = require('./lib.js').need('sharp');
function probe(file) {
  const out = execFileSync('ffprobe', ['-v', 'error', '-select_streams', 'v:0', '-count_frames', '-show_entries',
    'stream=codec_name,profile,width,height,r_frame_rate,avg_frame_rate,nb_read_frames,pix_fmt,color_space,color_primaries,color_transfer,color_range:format=duration,format_name,size',
    '-of', 'json', file]).toString();
  const j = JSON.parse(out), s = j.streams[0];
  return { codec: s.codec_name, profile: s.profile, w: s.width, h: s.height, rate: s.r_frame_rate, avg: s.avg_frame_rate, frames: +s.nb_read_frames, pix: s.pix_fmt,
           matrix: s.color_space, primaries: s.color_primaries, transfer: s.color_transfer, range: s.color_range, duration: +j.format.duration, format: j.format.format_name, size: +j.format.size };
}
// every frame as RGB at w x h
function rgbFrames(file, w, h) {
  const raw = execFileSync('ffmpeg', ['-v', 'error', '-i', file, '-vf', `scale=${w}:${h}:flags=area`, '-pix_fmt', 'rgb24', '-f', 'rawvideo', '-'], { maxBuffer: 1 << 30 });
  const size = w * h * 3, n = Math.floor(raw.length / size), frames = [];
  for (let i = 0; i < n; i++) frames.push(raw.subarray(i * size, (i + 1) * size));
  return frames;
}
function mad(a, b) { let s = 0; for (let i = 0; i < a.length; i++) s += Math.abs(a[i] - b[i]); return s / a.length; }
// d[k] = difference between frame k and frame k+1, with the last entry wrapping to frame 0
function stepDiffs(frames) { return frames.map((f, k) => mad(f, frames[(k + 1) % frames.length])); }
async function pngToRgb(buf, w, h) { return sharp(buf).removeAlpha().resize(w, h, { fit: 'fill', kernel: 'cubic' }).raw().toBuffer(); }
function median(a) { const s = a.slice().sort((x, y) => x - y); return s[Math.floor(s.length / 2)]; }
module.exports = { probe, rgbFrames, mad, stepDiffs, pngToRgb, median };
