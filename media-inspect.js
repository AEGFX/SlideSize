/* ============================================================
   Media Inspector engine  (mi*)
   ISO BMFF / H.264 SPS parser + playback risk rules.
   Folded into PPT Report so a dropped deck gets its videos
   inspected as part of the pre-flight.

   Source of truth: github.com/AEGFX/media-inspect
   Kept in modern syntax deliberately: this is a bitstream
   parser with a test matrix behind it, and a hand rewrite to
   var/ES5 house style would risk silent numeric bugs for no
   compatibility gain. Every feature used here is supported in
   every browser that can already run this page.
   ============================================================ */
window.MI = (function () {
'use strict';

/**
 * mp4-inspect.js
 * ISO BMFF (MP4 / MOV / M4V) container inspector with H.264 SPS and HEVC hvcC decoding.
 *
 * Zero dependencies. Runs in the browser and in Node.
 * Reads only box headers plus the moov atom, so a 20 GB file costs a handful of
 * range reads and never lands in memory.
 *
 * Usage (browser):
 *   import { inspect, BlobReader } from './mp4-inspect.js';
 *   const info = await inspect(new BlobReader(file), { name: file.name });
 *
 * Usage (bytes already in hand, e.g. pulled out of a PPTX by JSZip):
 *   import { inspect, BytesReader } from './mp4-inspect.js';
 *   const info = await inspect(new BytesReader(uint8), { name: 'media1.mp4' });
 *
 * AEGFX / SlideSize
 */

/* ------------------------------------------------------------------ *
 * Readers                                                             *
 * ------------------------------------------------------------------ */

class BlobReader {
  constructor(blob) { this.blob = blob; this.size = blob.size; }
  async read(offset, length) {
    if (offset >= this.size) return new Uint8Array(0);
    const end = Math.min(offset + length, this.size);
    return new Uint8Array(await this.blob.slice(offset, end).arrayBuffer());
  }
}

class BytesReader {
  constructor(bytes) { this.bytes = bytes; this.size = bytes.length; }
  async read(offset, length) {
    if (offset >= this.size) return new Uint8Array(0);
    return this.bytes.subarray(offset, Math.min(offset + length, this.size));
  }
}

/* ------------------------------------------------------------------ *
 * Byte helpers                                                        *
 * ------------------------------------------------------------------ */

const u8 = (b, o) => b[o];
const u16 = (b, o) => (b[o] << 8) | b[o + 1];
const i16 = (b, o) => { const v = u16(b, o); return v >= 0x8000 ? v - 0x10000 : v; };
const u24 = (b, o) => (b[o] << 16) | (b[o + 1] << 8) | b[o + 2];
const u32 = (b, o) => ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
const i32 = (b, o) => (b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3];
const u64 = (b, o) => u32(b, o) * 4294967296 + u32(b, o + 4);
const fx1616 = (b, o) => i32(b, o) / 65536;
const fx230 = (b, o) => i32(b, o) / 1073741824;
const str4 = (b, o) => String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]);

function cleanStr(b, o, len) {
  let s = '';
  for (let i = 0; i < len; i++) {
    const c = b[o + i];
    if (c === 0) break;
    s += String.fromCharCode(c);
  }
  return s.trim();
}

/* ------------------------------------------------------------------ *
 * Box walking                                                         *
 * ------------------------------------------------------------------ */

const CONTAINERS = new Set([
  'moov', 'trak', 'mdia', 'minf', 'stbl', 'edts', 'dinf', 'mvex',
  'udta', 'moof', 'traf', 'mfra', 'wave', 'tref',
]);

/** Read one box header at an absolute file offset. Async, used for the top level only. */
async function readHeader(reader, offset) {
  const head = await reader.read(offset, 16);
  if (head.length < 8) return null;
  let size = u32(head, 0);
  const type = str4(head, 4);
  let headerSize = 8;
  if (size === 1) {
    if (head.length < 16) return null;
    size = u64(head, 8);
    headerSize = 16;
  } else if (size === 0) {
    size = reader.size - offset;
  }
  if (size < headerSize) return null;
  return { type, offset, size, headerSize, dataOffset: offset + headerSize, dataSize: size - headerSize };
}

/**
 * Walk children of an in-memory box.
 * `bytes` is the whole buffer, [start,end) the payload range.
 */
function* children(bytes, start, end) {
  let off = start;
  while (off + 8 <= end) {
    let size = u32(bytes, off);
    const type = str4(bytes, off + 4);
    let hs = 8;
    if (size === 1) {
      if (off + 16 > end) return;
      size = u64(bytes, off + 8);
      hs = 16;
    } else if (size === 0) {
      size = end - off;
    }
    if (size < hs) return;
    const boxEnd = Math.min(off + size, end);   // tolerate truncation
    yield { type, start: off + hs, end: boxEnd, size };
    if (off + size <= off) return;              // paranoia against zero advance
    off += size;
  }
}

function findChild(bytes, box, type) {
  for (const c of children(bytes, box.start, box.end)) if (c.type === type) return c;
  return null;
}

function findPath(bytes, box, path) {
  let cur = box;
  for (const t of path) {
    cur = findChild(bytes, cur, t);
    if (!cur) return null;
  }
  return cur;
}

/* ------------------------------------------------------------------ *
 * Bit reader for SPS parsing                                          *
 * ------------------------------------------------------------------ */

class BitReader {
  constructor(bytes) { this.b = bytes; this.pos = 0; }
  get bitsLeft() { return this.b.length * 8 - this.pos; }
  bit() {
    if (this.pos >= this.b.length * 8) throw new Error('bitstream overrun');
    const v = (this.b[this.pos >> 3] >> (7 - (this.pos & 7))) & 1;
    this.pos++;
    return v;
  }
  u(n) { let v = 0; for (let i = 0; i < n; i++) v = v * 2 + this.bit(); return v; }
  ue() {
    let lz = 0;
    while (this.bit() === 0) { if (++lz > 32) throw new Error('exp-golomb overrun'); }
    return lz === 0 ? 0 : (Math.pow(2, lz) - 1) + this.u(lz);
  }
  se() { const k = this.ue(); return (k & 1) ? (k + 1) / 2 : -(k / 2); }
}

/** Strip emulation prevention bytes (00 00 03 -> 00 00). */
function rbsp(nal) {
  const out = new Uint8Array(nal.length);
  let o = 0, zeros = 0;
  for (let i = 0; i < nal.length; i++) {
    const b = nal[i];
    if (zeros >= 2 && b === 3) { zeros = 0; continue; }
    out[o++] = b;
    zeros = b === 0 ? zeros + 1 : 0;
  }
  return out.subarray(0, o);
}

/* ------------------------------------------------------------------ *
 * H.264 SPS                                                           *
 * ------------------------------------------------------------------ */

const HIGH_PROFILES = new Set([100, 110, 122, 244, 44, 83, 86, 118, 128, 138, 139, 134, 135]);

const H264_PROFILE_NAMES = {
  66: 'Baseline', 77: 'Main', 88: 'Extended', 100: 'High', 110: 'High 10',
  122: 'High 4:2:2', 244: 'High 4:4:4 Predictive', 44: 'CAVLC 4:4:4',
  83: 'Scalable Baseline', 86: 'Scalable High', 118: 'Stereo High',
  128: 'Multiview High', 138: 'Multiview Depth High',
};

const SAR_TABLE = {
  1: [1, 1], 2: [12, 11], 3: [10, 11], 4: [16, 11], 5: [40, 33], 6: [24, 11],
  7: [20, 11], 8: [32, 11], 9: [80, 33], 10: [18, 11], 11: [15, 11],
  12: [64, 33], 13: [160, 99], 14: [4, 3], 15: [3, 2], 16: [2, 1],
};

const CHROMA_NAMES = { 0: 'monochrome', 1: '4:2:0', 2: '4:2:2', 3: '4:4:4' };

function skipScalingList(r, size) {
  let last = 8, next = 8;
  for (let i = 0; i < size; i++) {
    if (next !== 0) {
      const delta = r.se();
      next = (last + delta + 256) % 256;
    }
    last = next === 0 ? last : next;
  }
}

function skipHrd(r) {
  const cpbCnt = r.ue() + 1;
  r.u(4); r.u(4);
  for (let i = 0; i < cpbCnt; i++) { r.ue(); r.ue(); r.u(1); }
  r.u(5); r.u(5); r.u(5); r.u(5);
}

function parseSPS(nalWithHeader) {
  const d = rbsp(nalWithHeader.subarray(1));
  const r = new BitReader(d);
  const sps = {};

  sps.profileIdc = r.u(8);
  const cflags = r.u(8);
  sps.constraintFlags = cflags;
  sps.levelIdc = r.u(8);
  sps.spsId = r.ue();

  sps.chromaFormatIdc = 1;
  sps.separateColourPlane = 0;
  sps.bitDepthLuma = 8;
  sps.bitDepthChroma = 8;

  if (HIGH_PROFILES.has(sps.profileIdc)) {
    sps.chromaFormatIdc = r.ue();
    if (sps.chromaFormatIdc === 3) sps.separateColourPlane = r.u(1);
    sps.bitDepthLuma = r.ue() + 8;
    sps.bitDepthChroma = r.ue() + 8;
    r.u(1); // qpprime_y_zero_transform_bypass_flag
    if (r.u(1)) { // seq_scaling_matrix_present_flag
      const n = sps.chromaFormatIdc !== 3 ? 8 : 12;
      for (let i = 0; i < n; i++) if (r.u(1)) skipScalingList(r, i < 6 ? 16 : 64);
    }
  }

  r.ue(); // log2_max_frame_num_minus4
  const pocType = r.ue();
  if (pocType === 0) {
    r.ue();
  } else if (pocType === 1) {
    r.u(1); r.se(); r.se();
    const n = r.ue();
    for (let i = 0; i < n; i++) r.se();
  }
  sps.maxNumRefFrames = r.ue();
  r.u(1); // gaps_in_frame_num_value_allowed_flag

  sps.picWidthInMbs = r.ue() + 1;
  sps.picHeightInMapUnits = r.ue() + 1;
  sps.frameMbsOnly = r.u(1);
  sps.mbAdaptiveFrameField = sps.frameMbsOnly ? 0 : r.u(1);
  r.u(1); // direct_8x8_inference_flag

  let crop = { l: 0, r: 0, t: 0, b: 0 };
  if (r.u(1)) crop = { l: r.ue(), r: r.ue(), t: r.ue(), b: r.ue() };
  sps.crop = crop;

  // Derived geometry
  const chromaArray = sps.separateColourPlane ? 0 : sps.chromaFormatIdc;
  const subW = chromaArray === 1 || chromaArray === 2 ? 2 : 1;
  const subH = chromaArray === 1 ? 2 : 1;
  const cropUnitX = chromaArray === 0 ? 1 : subW;
  const cropUnitY = (chromaArray === 0 ? 1 : subH) * (2 - sps.frameMbsOnly);

  sps.widthMb = sps.picWidthInMbs * 16;
  sps.heightMb = (2 - sps.frameMbsOnly) * sps.picHeightInMapUnits * 16;
  sps.width = sps.widthMb - cropUnitX * (crop.l + crop.r);
  sps.height = sps.heightMb - cropUnitY * (crop.t + crop.b);
  sps.macroblocks = sps.picWidthInMbs * ((2 - sps.frameMbsOnly) * sps.picHeightInMapUnits);

  // VUI
  sps.vui = null;
  try {
    if (r.u(1)) {
      const v = {};
      if (r.u(1)) { // aspect_ratio_info_present_flag
        const idc = r.u(8);
        if (idc === 255) v.sar = [r.u(16), r.u(16)];
        else if (SAR_TABLE[idc]) v.sar = SAR_TABLE[idc];
        v.aspectRatioIdc = idc;
      }
      if (r.u(1)) r.u(1); // overscan
      if (r.u(1)) { // video_signal_type_present_flag
        v.videoFormat = r.u(3);
        v.fullRange = r.u(1) === 1;
        if (r.u(1)) {
          v.colourPrimaries = r.u(8);
          v.transferCharacteristics = r.u(8);
          v.matrixCoefficients = r.u(8);
        }
      }
      if (r.u(1)) { v.chromaSampleLocTop = r.ue(); v.chromaSampleLocBottom = r.ue(); }
      if (r.u(1)) { // timing_info_present_flag
        v.numUnitsInTick = r.u(32);
        v.timeScale = r.u(32);
        v.fixedFrameRate = r.u(1) === 1;
        if (v.numUnitsInTick > 0) v.fps = v.timeScale / (2 * v.numUnitsInTick);
      }
      const nalHrd = r.u(1); if (nalHrd) skipHrd(r);
      const vclHrd = r.u(1); if (vclHrd) skipHrd(r);
      if (nalHrd || vclHrd) r.u(1);
      v.picStructPresent = r.u(1) === 1;
      sps.vui = v;
    }
  } catch (e) {
    sps.vuiError = String(e.message || e);
  }

  sps.profileName = H264_PROFILE_NAMES[sps.profileIdc] || `profile ${sps.profileIdc}`;
  sps.chromaName = CHROMA_NAMES[sps.chromaFormatIdc] ?? String(sps.chromaFormatIdc);
  return sps;
}

/* ------------------------------------------------------------------ *
 * Codec config boxes                                                  *
 * ------------------------------------------------------------------ */

function parseAvcC(bytes, box) {
  const s = box.start;
  const cfg = {
    configVersion: u8(bytes, s),
    profileIdc: u8(bytes, s + 1),
    profileCompat: u8(bytes, s + 2),
    levelIdc: u8(bytes, s + 3),
    lengthSize: (u8(bytes, s + 4) & 0x03) + 1,
    sps: [], pps: [],
  };
  let o = s + 5;
  const numSps = u8(bytes, o) & 0x1f; o += 1;
  for (let i = 0; i < numSps && o + 2 <= box.end; i++) {
    const len = u16(bytes, o); o += 2;
    cfg.sps.push(bytes.subarray(o, o + len)); o += len;
  }
  if (o < box.end) {
    const numPps = u8(bytes, o); o += 1;
    for (let i = 0; i < numPps && o + 2 <= box.end; i++) {
      const len = u16(bytes, o); o += 2;
      cfg.pps.push(bytes.subarray(o, o + len)); o += len;
    }
  }
  // Optional extension for High profiles
  if (o + 4 <= box.end && HIGH_PROFILES.has(cfg.profileIdc)) {
    cfg.extChromaFormat = u8(bytes, o) & 0x03;
    cfg.extBitDepthLuma = (u8(bytes, o + 1) & 0x07) + 8;
    cfg.extBitDepthChroma = (u8(bytes, o + 2) & 0x07) + 8;
  }
  return cfg;
}

function parseHvcC(bytes, box) {
  const s = box.start;
  const b1 = u8(bytes, s + 1);
  return {
    configVersion: u8(bytes, s),
    generalProfileSpace: (b1 >> 6) & 0x03,
    generalTierFlag: (b1 >> 5) & 0x01,
    generalProfileIdc: b1 & 0x1f,
    generalProfileCompat: u32(bytes, s + 2),
    generalLevelIdc: u8(bytes, s + 12),
    chromaFormatIdc: u8(bytes, s + 16) & 0x03,
    bitDepthLuma: (u8(bytes, s + 17) & 0x07) + 8,
    bitDepthChroma: (u8(bytes, s + 18) & 0x07) + 8,
    avgFrameRate: u16(bytes, s + 19) / 256,
    constantFrameRate: (u8(bytes, s + 21) >> 6) & 0x03,
    numTemporalLayers: (u8(bytes, s + 21) >> 3) & 0x07,
    lengthSize: (u8(bytes, s + 21) & 0x03) + 1,
  };
}

function parseEsds(bytes, box) {
  // Minimal descriptor walk: ES_Descriptor -> DecoderConfigDescriptor -> DecoderSpecificInfo
  let o = box.start + 4; // version + flags
  const out = {};
  const readLen = () => {
    let len = 0, b;
    do { b = u8(bytes, o++); len = (len << 7) | (b & 0x7f); } while (b & 0x80);
    return len;
  };
  try {
    if (u8(bytes, o) !== 0x03) return out;
    o++; readLen();
    o += 2;                                   // ES_ID
    const flags = u8(bytes, o++);
    if (flags & 0x80) o += 2;
    if (flags & 0x40) o += 1 + u8(bytes, o);
    if (flags & 0x20) o += 2;
    if (u8(bytes, o) !== 0x04) return out;
    o++; readLen();
    out.objectTypeIndication = u8(bytes, o); o += 1;
    out.streamType = u8(bytes, o) >> 2; o += 1;
    o += 3;                                   // bufferSizeDB
    out.maxBitrate = u32(bytes, o); o += 4;
    out.avgBitrate = u32(bytes, o); o += 4;
    if (u8(bytes, o) === 0x05) {
      o++; const len = readLen();
      const asc = bytes.subarray(o, o + len);
      if (asc.length >= 2) {
        const r = new BitReader(asc);
        let aot = r.u(5);
        if (aot === 31) aot = 32 + r.u(6);
        out.audioObjectType = aot;
        const sfi = r.u(4);
        out.samplingFrequencyIndex = sfi;
        if (sfi === 15) out.sampleRate = r.u(24);
        out.channelConfiguration = r.u(4);
      }
    }
  } catch { /* best effort */ }
  return out;
}

/* ------------------------------------------------------------------ *
 * Codec identity                                                      *
 * ------------------------------------------------------------------ */

const CODEC_NAMES = {
  avc1: 'H.264 / AVC', avc3: 'H.264 / AVC (in-band params)',
  hvc1: 'HEVC / H.265', hev1: 'HEVC / H.265 (in-band params)',
  vp09: 'VP9', vp08: 'VP8', av01: 'AV1', mp4v: 'MPEG-4 Part 2', s263: 'H.263',
  apco: 'Apple ProRes 422 Proxy', apcs: 'Apple ProRes 422 LT',
  apcn: 'Apple ProRes 422', apch: 'Apple ProRes 422 HQ',
  ap4h: 'Apple ProRes 4444', ap4x: 'Apple ProRes 4444 XQ', icpf: 'Apple ProRes',
  AVdn: 'Avid DNxHD / DNxHR', AVdh: 'Avid DNxHR',
  Hap1: 'HAP', Hap5: 'HAP Alpha', HapY: 'HAP Q', HapM: 'HAP Q Alpha', HapA: 'HAP Alpha Only',
  CFHD: 'GoPro CineForm', 'rle ': 'QuickTime Animation (RLE)',
  jpeg: 'Motion JPEG', mjpa: 'Motion JPEG A', mjpb: 'Motion JPEG B',
  png: 'PNG sequence', tiff: 'TIFF sequence',
  '2vuy': 'Uncompressed 8-bit 4:2:2', v210: 'Uncompressed 10-bit 4:2:2',
  v410: 'Uncompressed 10-bit 4:4:4', r210: 'Uncompressed 10-bit RGB',
  raw: 'Uncompressed RGB', dvc: 'DV', dvcp: 'DV PAL', dvh5: 'DVCPRO HD', dvh6: 'DVCPRO HD',
  mp4a: 'AAC / MPEG-4 Audio', 'ac-3': 'Dolby Digital (AC-3)', 'ec-3': 'Dolby Digital Plus',
  dtsc: 'DTS', alac: 'Apple Lossless', Opus: 'Opus', '.mp3': 'MP3',
  lpcm: 'Uncompressed PCM', sowt: 'PCM 16-bit LE', twos: 'PCM 16-bit BE',
  in24: 'PCM 24-bit', in32: 'PCM 32-bit', fl32: 'PCM 32-bit float', samr: 'AMR',
  tmcd: 'Timecode', c608: 'CEA-608 captions', c708: 'CEA-708 captions',
  tx3g: 'Timed text', wvtt: 'WebVTT',
};

const INTRA_ONLY = new Set([
  'apco', 'apcs', 'apcn', 'apch', 'ap4h', 'ap4x', 'icpf', 'AVdn', 'AVdh',
  'Hap1', 'Hap5', 'HapY', 'HapM', 'HapA', 'CFHD', 'rle ', 'jpeg', 'mjpa', 'mjpb',
  '2vuy', 'v210', 'v410', 'r210', 'raw', 'dvc', 'dvcp', 'dvh5', 'dvh6',
]);

const hex2 = (n) => n.toString(16).toUpperCase().padStart(2, '0');

function h264CodecString(cfg) {
  return `avc1.${hex2(cfg.profileIdc)}${hex2(cfg.profileCompat)}${hex2(cfg.levelIdc)}`;
}

function hevcCodecString(fourcc, c) {
  const space = ['', 'A', 'B', 'C'][c.generalProfileSpace] || '';
  // compatibility flags are signalled in reverse bit order
  let rev = 0;
  for (let i = 0; i < 32; i++) rev = (rev << 1) | ((c.generalProfileCompat >>> i) & 1);
  const tier = c.generalTierFlag ? 'H' : 'L';
  return `${fourcc}.${space}${c.generalProfileIdc}.${(rev >>> 0).toString(16).toUpperCase()}.${tier}${c.generalLevelIdc}`;
}

/* ------------------------------------------------------------------ *
 * Level tables                                                        *
 * ------------------------------------------------------------------ */

// H.264 Table A-1. MaxBR is the Baseline/Main/Extended figure in kbps.
const H264_LEVELS = [
  { level: 1.0, name: '1',   maxMBPS: 1485,     maxFS: 99,     maxDpbMbs: 396,    maxBR: 64 },
  { level: 1.05, name: '1b', maxMBPS: 1485,     maxFS: 99,     maxDpbMbs: 396,    maxBR: 128 },
  { level: 1.1, name: '1.1', maxMBPS: 3000,     maxFS: 396,    maxDpbMbs: 900,    maxBR: 192 },
  { level: 1.2, name: '1.2', maxMBPS: 6000,     maxFS: 396,    maxDpbMbs: 2376,   maxBR: 384 },
  { level: 1.3, name: '1.3', maxMBPS: 11880,    maxFS: 396,    maxDpbMbs: 2376,   maxBR: 768 },
  { level: 2.0, name: '2',   maxMBPS: 11880,    maxFS: 396,    maxDpbMbs: 2376,   maxBR: 2000 },
  { level: 2.1, name: '2.1', maxMBPS: 19800,    maxFS: 792,    maxDpbMbs: 4752,   maxBR: 4000 },
  { level: 2.2, name: '2.2', maxMBPS: 20250,    maxFS: 1620,   maxDpbMbs: 8100,   maxBR: 4000 },
  { level: 3.0, name: '3',   maxMBPS: 40500,    maxFS: 1620,   maxDpbMbs: 8100,   maxBR: 10000 },
  { level: 3.1, name: '3.1', maxMBPS: 108000,   maxFS: 3600,   maxDpbMbs: 18000,  maxBR: 14000 },
  { level: 3.2, name: '3.2', maxMBPS: 216000,   maxFS: 5120,   maxDpbMbs: 20480,  maxBR: 20000 },
  { level: 4.0, name: '4',   maxMBPS: 245760,   maxFS: 8192,   maxDpbMbs: 32768,  maxBR: 20000 },
  { level: 4.1, name: '4.1', maxMBPS: 245760,   maxFS: 8192,   maxDpbMbs: 32768,  maxBR: 50000 },
  { level: 4.2, name: '4.2', maxMBPS: 522240,   maxFS: 8704,   maxDpbMbs: 34816,  maxBR: 50000 },
  { level: 5.0, name: '5',   maxMBPS: 589824,   maxFS: 22080,  maxDpbMbs: 110400, maxBR: 135000 },
  { level: 5.1, name: '5.1', maxMBPS: 983040,   maxFS: 36864,  maxDpbMbs: 184320, maxBR: 240000 },
  { level: 5.2, name: '5.2', maxMBPS: 2073600,  maxFS: 36864,  maxDpbMbs: 184320, maxBR: 240000 },
  { level: 6.0, name: '6',   maxMBPS: 4177920,  maxFS: 139264, maxDpbMbs: 696320, maxBR: 240000 },
  { level: 6.1, name: '6.1', maxMBPS: 8355840,  maxFS: 139264, maxDpbMbs: 696320, maxBR: 480000 },
  { level: 6.2, name: '6.2', maxMBPS: 16711680, maxFS: 139264, maxDpbMbs: 696320, maxBR: 800000 },
];

// HEVC Table A.8 MaxLumaPs, and MaxLumaSr from Table A.9.
const HEVC_LEVELS = [
  { level: 1.0, name: '1',   maxLumaPs: 36864,    maxLumaSr: 552960 },
  { level: 2.0, name: '2',   maxLumaPs: 122880,   maxLumaSr: 3686400 },
  { level: 2.1, name: '2.1', maxLumaPs: 245760,   maxLumaSr: 7372800 },
  { level: 3.0, name: '3',   maxLumaPs: 552960,   maxLumaSr: 16588800 },
  { level: 3.1, name: '3.1', maxLumaPs: 983040,   maxLumaSr: 33177600 },
  { level: 4.0, name: '4',   maxLumaPs: 2228224,  maxLumaSr: 66846720 },
  { level: 4.1, name: '4.1', maxLumaPs: 2228224,  maxLumaSr: 133693440 },
  { level: 5.0, name: '5',   maxLumaPs: 8912896,  maxLumaSr: 267386880 },
  { level: 5.1, name: '5.1', maxLumaPs: 8912896,  maxLumaSr: 534773760 },
  { level: 5.2, name: '5.2', maxLumaPs: 8912896,  maxLumaSr: 1069547520 },
  { level: 6.0, name: '6',   maxLumaPs: 35651584, maxLumaSr: 1069547520 },
  { level: 6.1, name: '6.1', maxLumaPs: 35651584, maxLumaSr: 2139095040 },
  { level: 6.2, name: '6.2', maxLumaPs: 35651584, maxLumaSr: 4278190080 },
];

function h264LevelFromIdc(levelIdc, constraintFlags) {
  if (levelIdc === 11 && (constraintFlags & 0x10)) return H264_LEVELS[1]; // 1b
  const v = levelIdc / 10;
  return H264_LEVELS.find((l) => Math.abs(l.level - v) < 0.001) || null;
}

/**
 * Does the declared level actually cover this picture?
 * This is the check that catches Mac screen recordings that claim level 4.0
 * while carrying a frame that needs level 5.
 */
function checkH264Level(levelIdc, constraintFlags, macroblocks, fps, bitrateBps, profileIdc) {
  const declared = h264LevelFromIdc(levelIdc, constraintFlags);
  const mbps = macroblocks * (fps || 0);

  const byFrameSize = H264_LEVELS.find((l) => macroblocks <= l.maxFS) || null;
  const byThroughput = H264_LEVELS.find((l) => macroblocks <= l.maxFS && mbps <= l.maxMBPS) || null;

  // Profile-dependent bitrate multiplier (cpbBrVclFactor family)
  let brFactor = 1;
  if (profileIdc === 100) brFactor = 1.25;
  else if (profileIdc === 110) brFactor = 3;
  else if (profileIdc === 122 || profileIdc === 244) brFactor = 4;

  const declaredBrBps = declared ? declared.maxBR * 1000 * brFactor : null;

  return {
    declaredIdc: levelIdc,
    declared: declared ? declared.name : `unknown (${levelIdc})`,
    declaredLevel: declared ? declared.level : null,
    macroblocks,
    macroblocksPerSecond: Math.round(mbps),
    maxFS: declared ? declared.maxFS : null,
    maxMBPS: declared ? declared.maxMBPS : null,
    requiredByFrameSize: byFrameSize ? byFrameSize.name : null,
    requiredByThroughput: byThroughput ? byThroughput.name : null,
    frameSizeOk: declared ? macroblocks <= declared.maxFS : null,
    throughputOk: declared ? mbps <= declared.maxMBPS : null,
    bitrateOk: declaredBrBps && bitrateBps ? bitrateBps <= declaredBrBps : null,
    maxBitrateBps: declaredBrBps,
    conformant: declared ? macroblocks <= declared.maxFS && mbps <= declared.maxMBPS : null,
  };
}

function checkHevcLevel(levelIdc, lumaSamples, fps) {
  const v = levelIdc / 30;
  const declared = HEVC_LEVELS.find((l) => Math.abs(l.level - v) < 0.001) || null;
  const sr = lumaSamples * (fps || 0);
  const byFrameSize = HEVC_LEVELS.find((l) => lumaSamples <= l.maxLumaPs) || null;
  return {
    declared: declared ? declared.name : `unknown (${levelIdc})`,
    lumaSamples,
    maxLumaPs: declared ? declared.maxLumaPs : null,
    requiredByFrameSize: byFrameSize ? byFrameSize.name : null,
    frameSizeOk: declared ? lumaSamples <= declared.maxLumaPs : null,
    throughputOk: declared ? sr <= declared.maxLumaSr : null,
    conformant: declared ? lumaSamples <= declared.maxLumaPs && sr <= declared.maxLumaSr : null,
  };
}

/* ------------------------------------------------------------------ *
 * Colour signalling                                                   *
 * ------------------------------------------------------------------ */

const PRIMARIES = { 1: 'BT.709', 4: 'BT.470M', 5: 'BT.601 PAL', 6: 'BT.601 NTSC', 7: 'SMPTE 240M', 9: 'BT.2020', 11: 'DCI-P3', 12: 'Display P3' };
const TRANSFER = { 1: 'BT.709', 4: 'Gamma 2.2', 5: 'Gamma 2.8', 6: 'BT.601', 7: 'SMPTE 240M', 8: 'Linear', 13: 'sRGB', 14: 'BT.2020 10-bit', 15: 'BT.2020 12-bit', 16: 'PQ (HDR10)', 17: 'SMPTE 428', 18: 'HLG' };
const MATRIX = { 0: 'RGB', 1: 'BT.709', 4: 'FCC', 5: 'BT.470BG', 6: 'BT.601', 7: 'SMPTE 240M', 9: 'BT.2020 NCL', 10: 'BT.2020 CL' };

const nameOf = (table, v) => (v == null ? null : table[v] || `unknown (${v})`);

/* ------------------------------------------------------------------ *
 * Sample table decoding                                               *
 * ------------------------------------------------------------------ */

function parseStts(bytes, box) {
  const count = u32(bytes, box.start + 4);
  const entries = [];
  let samples = 0, duration = 0;
  for (let i = 0; i < count; i++) {
    const o = box.start + 8 + i * 8;
    if (o + 8 > box.end) break;
    const sampleCount = u32(bytes, o);
    const sampleDelta = u32(bytes, o + 4);
    entries.push({ sampleCount, sampleDelta });
    samples += sampleCount;
    duration += sampleCount * sampleDelta;
  }
  return { entries, samples, duration };
}

function parseStss(bytes, box) {
  const count = u32(bytes, box.start + 4);
  const out = [];
  for (let i = 0; i < count; i++) {
    const o = box.start + 8 + i * 4;
    if (o + 4 > box.end) break;
    out.push(u32(bytes, o));
  }
  return out;
}

function parseStsz(bytes, box) {
  const uniform = u32(bytes, box.start + 4);
  const count = u32(bytes, box.start + 8);
  if (uniform !== 0) return { count, total: uniform * count, max: uniform, uniform };
  let total = 0, max = 0;
  for (let i = 0; i < count; i++) {
    const o = box.start + 12 + i * 4;
    if (o + 4 > box.end) break;
    const s = u32(bytes, o);
    total += s;
    if (s > max) max = s;
  }
  return { count, total, max, uniform: 0 };
}

function parseStz2(bytes, box) {
  const fieldSize = u8(bytes, box.start + 7);
  const count = u32(bytes, box.start + 8);
  let total = 0, max = 0;
  for (let i = 0; i < count; i++) {
    let s = 0;
    if (fieldSize === 16) { const o = box.start + 12 + i * 2; if (o + 2 > box.end) break; s = u16(bytes, o); }
    else if (fieldSize === 8) { const o = box.start + 12 + i; if (o + 1 > box.end) break; s = u8(bytes, o); }
    else if (fieldSize === 4) { const o = box.start + 12 + (i >> 1); if (o + 1 > box.end) break; s = (i & 1) ? (u8(bytes, o) & 0x0f) : (u8(bytes, o) >> 4); }
    total += s;
    if (s > max) max = s;
  }
  return { count, total, max, uniform: 0 };
}

function parseElst(bytes, box) {
  const version = u8(bytes, box.start);
  const count = u32(bytes, box.start + 4);
  const out = [];
  let o = box.start + 8;
  for (let i = 0; i < count; i++) {
    if (version === 1) {
      if (o + 20 > box.end) break;
      const empty = u32(bytes, o + 8) === 0xffffffff && u32(bytes, o + 12) === 0xffffffff;
      out.push({ duration: u64(bytes, o), mediaTime: empty ? -1 : u64(bytes, o + 8), rate: fx1616(bytes, o + 16) });
      o += 20;
    } else {
      if (o + 12 > box.end) break;
      const mt = i32(bytes, o + 4);
      out.push({ duration: u32(bytes, o), mediaTime: mt, rate: fx1616(bytes, o + 8) });
      o += 12;
    }
  }
  return out;
}

/** Classify frame timing from the time-to-sample table. */
function classifyTiming(stts, timescale) {
  if (!stts || !stts.entries.length) return { mode: 'unknown' };
  const total = stts.samples;
  const byDelta = new Map();
  for (const e of stts.entries) byDelta.set(e.sampleDelta, (byDelta.get(e.sampleDelta) || 0) + e.sampleCount);

  const deltas = [...byDelta.entries()].sort((a, b) => b[1] - a[1]);
  const [modeDelta, modeCount] = deltas[0];
  const deltaVals = stts.entries.map((e) => e.sampleDelta).filter((d) => d > 0);
  const minDelta = Math.min(...deltaVals);
  const maxDelta = Math.max(...deltaVals);

  const avgFps = stts.duration > 0 ? (total * timescale) / stts.duration : null;
  const modeFps = modeDelta > 0 ? timescale / modeDelta : null;
  const share = total > 0 ? modeCount / total : 0;

  let mode;
  if (byDelta.size === 1) mode = 'CFR';
  else if (share >= 0.98) mode = 'near-CFR';
  else mode = 'VFR';

  return {
    mode,
    nominalFps: snapFrameRate(modeFps),
    distinctIntervals: byDelta.size,
    dominantShare: Number(share.toFixed(4)),
    fps: modeFps,
    avgFps,
    minFps: maxDelta > 0 ? timescale / maxDelta : null,
    maxFps: minDelta > 0 ? timescale / minDelta : null,
    sttsEntries: stts.entries.length,
    sampleCount: total,
  };
}

/** Snap a measured rate to a broadcast-standard rate if it is within tolerance. */
function snapFrameRate(fps) {
  if (!fps) return null;
  const std = [23.976, 24, 25, 29.97, 30, 47.952, 48, 50, 59.94, 60, 100, 119.88, 120];
  for (const s of std) if (Math.abs(fps - s) / s < 0.002) return s;
  return Number(fps.toFixed(3));
}

/* ------------------------------------------------------------------ *
 * Track parsing                                                       *
 * ------------------------------------------------------------------ */

function rotationFromMatrix(m) {
  // m = [a,b,u, c,d,v, x,y,w]
  const [a, b, , c, d] = m;
  if (a === 1 && b === 0 && c === 0 && d === 1) return 0;
  if (a === 0 && b === 1 && c === -1 && d === 0) return 90;
  if (a === -1 && b === 0 && c === 0 && d === -1) return 180;
  if (a === 0 && b === -1 && c === 1 && d === 0) return 270;
  const deg = Math.round((Math.atan2(b, a) * 180) / Math.PI);
  return ((deg % 360) + 360) % 360;
}

function parseVisualSampleEntry(bytes, entry) {
  const s = entry.start;
  const out = {
    fourcc: entry.type,
    codedWidth: u16(bytes, s + 24),
    codedHeight: u16(bytes, s + 26),
    horizResolution: fx1616(bytes, s + 28),
    vertResolution: fx1616(bytes, s + 32),
    frameCountPerSample: u16(bytes, s + 40),
    compressorName: cleanStr(bytes, s + 43, Math.min(31, u8(bytes, s + 42))),
    depth: u16(bytes, s + 74),
    boxes: {},
  };
  for (const c of children(bytes, s + 78, entry.end)) {
    out.boxes[c.type] = c;
  }
  return out;
}

function parseAudioSampleEntry(bytes, entry) {
  const s = entry.start;
  const version = u16(bytes, s + 8);
  const out = {
    fourcc: entry.type,
    version,
    channels: u16(bytes, s + 16),
    sampleSize: u16(bytes, s + 18),
    sampleRate: u16(bytes, s + 24),   // 16.16, integer part is what matters
    boxes: {},
  };
  let childStart = s + 28;
  if (version === 1) childStart = s + 28 + 16;
  else if (version === 2) childStart = s + 28 + 36;
  for (const c of children(bytes, childStart, entry.end)) out.boxes[c.type] = c;
  return out;
}

function parseTrack(bytes, trak, movieTimescale) {
  const tkhd = findChild(bytes, trak, 'tkhd');
  const mdia = findChild(bytes, trak, 'mdia');
  if (!mdia) return null;
  const mdhd = findChild(bytes, mdia, 'mdhd');
  const hdlr = findChild(bytes, mdia, 'hdlr');
  const stbl = findPath(bytes, mdia, ['minf', 'stbl']);

  const t = { warnings: [] };

  // tkhd
  if (tkhd) {
    const v = u8(bytes, tkhd.start);
    let o = tkhd.start + 4;
    t.enabled = (u24(bytes, tkhd.start + 1) & 0x01) === 1;
    if (v === 1) { o += 16; t.trackId = u32(bytes, o); o += 4 + 4 + 8; }
    else { o += 8; t.trackId = u32(bytes, o); o += 4 + 4 + 4; }
    o += 8;                                     // reserved
    o += 2 + 2 + 2 + 2;                         // layer, altgroup, volume, reserved
    const m = [];
    for (let i = 0; i < 9; i++) m.push(i === 2 || i === 5 || i === 8 ? fx230(bytes, o + i * 4) : fx1616(bytes, o + i * 4));
    o += 36;
    t.matrix = m;
    t.rotation = rotationFromMatrix(m);
    t.displayWidth = Math.round(fx1616(bytes, o));
    t.displayHeight = Math.round(fx1616(bytes, o + 4));
  }

  // mdhd
  let timescale = 1000, mediaDuration = 0;
  if (mdhd) {
    const v = u8(bytes, mdhd.start);
    if (v === 1) { timescale = u32(bytes, mdhd.start + 20); mediaDuration = u64(bytes, mdhd.start + 24); }
    else { timescale = u32(bytes, mdhd.start + 12); mediaDuration = u32(bytes, mdhd.start + 16); }
  }
  t.timescale = timescale;
  t.duration = timescale > 0 ? mediaDuration / timescale : 0;

  // hdlr
  if (hdlr) {
    t.handlerType = str4(bytes, hdlr.start + 8);
    t.handlerName = cleanStr(bytes, hdlr.start + 24, hdlr.end - hdlr.start - 24);
  }

  // edit list
  const elst = findPath(bytes, trak, ['edts', 'elst']);
  if (elst) {
    t.editList = parseElst(bytes, elst).map((e) => ({
      durationSeconds: movieTimescale ? e.duration / movieTimescale : null,
      mediaTimeSeconds: e.mediaTime === -1 ? null : e.mediaTime / timescale,
      empty: e.mediaTime === -1,
      rate: e.rate,
    }));
  }

  if (!stbl) { t.warnings.push('No sample table (fragmented or damaged file)'); return t; }

  // sample description
  const stsd = findChild(bytes, stbl, 'stsd');
  let entry = null;
  if (stsd) {
    const entryCount = u32(bytes, stsd.start + 4);
    t.sampleDescriptionCount = entryCount;
    if (entryCount > 1) t.warnings.push(`${entryCount} sample descriptions: codec or size changes mid-file`);
    for (const c of children(bytes, stsd.start + 8, stsd.end)) { entry = c; break; }
  }

  // sample tables
  const stts = findChild(bytes, stbl, 'stts');
  const stss = findChild(bytes, stbl, 'stss');
  const ctts = findChild(bytes, stbl, 'ctts');
  const stszBox = findChild(bytes, stbl, 'stsz');
  const stz2Box = findChild(bytes, stbl, 'stz2');

  const timing = stts ? classifyTiming(parseStts(bytes, stts), timescale) : { mode: 'unknown' };
  const sizes = stszBox ? parseStsz(bytes, stszBox) : stz2Box ? parseStz2(bytes, stz2Box) : null;

  t.frameRate = timing;
  t.sampleCount = timing.sampleCount ?? sizes?.count ?? 0;
  t.hasBFrames = !!ctts;
  if (sizes) {
    t.totalBytes = sizes.total;
    t.maxSampleBytes = sizes.max;
    t.bitrate = t.duration > 0 ? Math.round((sizes.total * 8) / t.duration) : null;
  }

  if (stss) {
    const kf = parseStss(bytes, stss);
    t.keyframes = {
      count: kf.length,
      allIntra: kf.length > 0 && kf.length === t.sampleCount,
      intervalFrames: kf.length > 1 ? Math.round(t.sampleCount / kf.length) : t.sampleCount,
      intervalSeconds: kf.length > 0 && t.duration > 0 ? Number((t.duration / kf.length).toFixed(2)) : null,
      first: kf.slice(0, 5),
    };
  } else if (t.sampleCount) {
    t.keyframes = { count: t.sampleCount, allIntra: true, intervalFrames: 1, intervalSeconds: 0 };
  }

  if (!entry) return t;

  const isVideo = t.handlerType === 'vide';
  const isAudio = t.handlerType === 'soun';

  if (isVideo) {
    const vse = parseVisualSampleEntry(bytes, entry);
    t.codec = { fourcc: vse.fourcc, name: CODEC_NAMES[vse.fourcc] || `unknown (${vse.fourcc})`, compressorName: vse.compressorName };
    t.coded = { width: vse.codedWidth, height: vse.codedHeight };
    t.depth = vse.depth;
    t.allIntraCodec = INTRA_ONLY.has(vse.fourcc);

    // pixel aspect
    if (vse.boxes.pasp) {
      const p = vse.boxes.pasp;
      t.pixelAspect = { h: u32(bytes, p.start), v: u32(bytes, p.start + 4), source: 'pasp' };
    }

    // clean aperture
    if (vse.boxes.clap) t.hasCleanAperture = true;

    // field ordering
    if (vse.boxes.fiel) {
      const f = vse.boxes.fiel;
      const fields = u8(bytes, f.start);
      const detail = f.end - f.start > 1 ? u8(bytes, f.start + 1) : 0;
      t.interlaced = {
        flag: fields === 2,
        source: 'fiel',
        fieldOrder: fields === 2 ? ({ 1: 'top first (TFF)', 6: 'bottom first (BFF)', 9: 'bottom first (BFF)', 14: 'top first (TFF)' }[detail] || `unknown (${detail})`) : 'progressive',
      };
    }

    // colour
    if (vse.boxes.colr) {
      const c = vse.boxes.colr;
      const type = str4(bytes, c.start);
      if (type === 'nclx' || type === 'nclc') {
        const pr = u16(bytes, c.start + 4), tr = u16(bytes, c.start + 6), mx = u16(bytes, c.start + 8);
        t.colour = {
          primaries: nameOf(PRIMARIES, pr), primariesIdc: pr,
          transfer: nameOf(TRANSFER, tr), transferIdc: tr,
          matrix: nameOf(MATRIX, mx), matrixIdc: mx,
          fullRange: type === 'nclx' && c.end - c.start > 10 ? (u8(bytes, c.start + 10) & 0x80) !== 0 : null,
          source: `colr/${type}`,
        };
      } else {
        t.colour = { iccProfile: true, source: `colr/${type}` };
      }
    }

    // codec configuration
    if (vse.boxes.avcC) {
      const cfg = parseAvcC(bytes, vse.boxes.avcC);
      t.codec.string = h264CodecString(cfg);
      t.codec.nalLengthSize = cfg.lengthSize;
      t.avcC = {
        profileIdc: cfg.profileIdc,
        profileName: H264_PROFILE_NAMES[cfg.profileIdc] || `profile ${cfg.profileIdc}`,
        levelIdc: cfg.levelIdc,
        level: (h264LevelFromIdc(cfg.levelIdc, 0) || {}).name || String(cfg.levelIdc / 10),
        spsCount: cfg.sps.length,
        ppsCount: cfg.pps.length,
      };
      if (cfg.sps.length) {
        try {
          const sps = parseSPS(cfg.sps[0]);
          t.sps = sps;
          t.profile = sps.profileName;
          t.level = (h264LevelFromIdc(sps.levelIdc, sps.constraintFlags) || {}).name || String(sps.levelIdc / 10);
          t.bitDepth = sps.bitDepthLuma;
          t.chroma = sps.chromaName;
          t.cropped = { width: sps.width, height: sps.height };
          t.macroblocks = sps.macroblocks;

          if (!t.interlaced) {
            t.interlaced = {
              flag: sps.frameMbsOnly === 0,
              source: 'SPS frame_mbs_only_flag',
              fieldOrder: sps.frameMbsOnly === 0 ? (sps.mbAdaptiveFrameField ? 'MBAFF' : 'PAFF or field coded') : 'progressive',
            };
          }
          if (sps.vui) {
            if (sps.vui.sar && !t.pixelAspect) t.pixelAspect = { h: sps.vui.sar[0], v: sps.vui.sar[1], source: 'SPS VUI' };
            if (sps.vui.fps) t.frameRate.vuiFps = snapFrameRate(sps.vui.fps);
            if (sps.vui.fixedFrameRate != null) t.frameRate.vuiFixedFrameRate = sps.vui.fixedFrameRate;
            if (!t.colour && sps.vui.colourPrimaries != null) {
              t.colour = {
                primaries: nameOf(PRIMARIES, sps.vui.colourPrimaries), primariesIdc: sps.vui.colourPrimaries,
                transfer: nameOf(TRANSFER, sps.vui.transferCharacteristics), transferIdc: sps.vui.transferCharacteristics,
                matrix: nameOf(MATRIX, sps.vui.matrixCoefficients), matrixIdc: sps.vui.matrixCoefficients,
                fullRange: sps.vui.fullRange, source: 'SPS VUI',
              };
            } else if (t.colour && t.colour.fullRange == null && sps.vui.fullRange != null) {
              t.colour.fullRange = sps.vui.fullRange;
            }
          }

          const fps = t.frameRate.nominalFps || t.frameRate.fps || t.frameRate.avgFps;
          t.levelCheck = checkH264Level(sps.levelIdc, sps.constraintFlags, sps.macroblocks, fps, t.bitrate, sps.profileIdc);
        } catch (e) {
          t.warnings.push(`SPS parse failed: ${e.message}`);
        }
      }
    } else if (vse.boxes.hvcC) {
      const c = parseHvcC(bytes, vse.boxes.hvcC);
      t.hvcC = c;
      t.codec.string = hevcCodecString(vse.fourcc, c);
      t.profile = `HEVC profile ${c.generalProfileIdc}${c.generalTierFlag ? ' (High tier)' : ' (Main tier)'}`;
      t.level = (c.generalLevelIdc / 30).toFixed(1).replace(/\.0$/, '');
      t.bitDepth = c.bitDepthLuma;
      t.chroma = CHROMA_NAMES[c.chromaFormatIdc] ?? String(c.chromaFormatIdc);
      const fps = t.frameRate.nominalFps || t.frameRate.fps || t.frameRate.avgFps;
      t.levelCheck = checkHevcLevel(c.generalLevelIdc, vse.codedWidth * vse.codedHeight, fps);
      if (vse.fourcc === 'hev1') t.warnings.push('hev1 keeps parameter sets in-band, which some players and hardware decoders reject');
    } else if (vse.boxes.av1C) {
      const s = vse.boxes.av1C.start;
      const b1 = u8(bytes, s + 1), b2 = u8(bytes, s + 2);
      t.codec.string = `av01.${(b1 >> 5) & 0x07}.${b1 & 0x1f}`;
      t.bitDepth = (b2 & 0x40) ? ((b2 & 0x20) ? 12 : 10) : 8;
      t.chroma = (b2 & 0x08) ? 'monochrome' : ((b2 & 0x04) ? ((b2 & 0x02) ? '4:2:0' : '4:2:2') : '4:4:4');
    } else if (vse.boxes.vpcC) {
      const s = vse.boxes.vpcC.start;
      t.codec.string = `vp09.${String(u8(bytes, s + 4)).padStart(2, '0')}.${String(u8(bytes, s + 5)).padStart(2, '0')}`;
      t.bitDepth = (u8(bytes, s + 6) >> 4) & 0x0f;
      const cs = (u8(bytes, s + 6) >> 1) & 0x07;
      t.chroma = { 0: '4:2:0', 1: '4:2:0', 2: '4:2:2', 3: '4:4:4' }[cs] ?? null;
    }

    // Infer bit depth for uncompressed and intra codecs where there is no config box
    // Codecs with no config box: infer from the fourcc first, then the depth field.
    if (t.bitDepth == null) {
      const guess = { v210: 10, r210: 10, v410: 10, ap4h: 12, ap4x: 12, apch: 10, apcn: 10, apcs: 10, apco: 10, icpf: 10, AVdn: 8, AVdh: 10, HapY: 8, CFHD: 10 };
      if (guess[vse.fourcc]) { t.bitDepth = guess[vse.fourcc]; t.bitDepthInferred = true; }
    }
    if (!t.chroma) {
      const cg = { apch: '4:2:2', apcn: '4:2:2', apcs: '4:2:2', apco: '4:2:2', ap4h: '4:4:4', ap4x: '4:4:4',
                   icpf: '4:2:2', v210: '4:2:2', '2vuy': '4:2:2', v410: '4:4:4', AVdn: '4:2:2', AVdh: '4:2:2',
                   r210: 'RGB', raw: 'RGB', 'rle ': 'RGB' };
      if (cg[vse.fourcc]) { t.chroma = cg[vse.fourcc]; t.chromaInferred = true; }
    }
    if (t.bitDepth == null && vse.depth) {
      if (vse.depth === 24 || vse.depth === 32 || vse.depth === 40) { t.bitDepth = 8; t.bitDepthInferred = true; }
      else if (vse.depth === 48 || vse.depth === 64) { t.bitDepth = 16; t.bitDepthInferred = true; }
      else if (vse.depth <= 8) { t.bitDepth = 8; t.bitDepthInferred = true; t.palettised = true; }
    }

    // Effective display geometry
    const cw = t.cropped?.width || t.coded.width;
    const ch = t.cropped?.height || t.coded.height;
    const pa = t.pixelAspect;
    const dw = pa && pa.h && pa.v ? Math.round((cw * pa.h) / pa.v) : cw;
    t.display = { width: dw, height: ch };
    t.storageAspect = simplifyRatio(cw, ch);
    t.displayAspect = simplifyRatio(dw, ch);
    t.squarePixels = !pa || pa.h === pa.v;
  }

  if (isAudio) {
    const ase = parseAudioSampleEntry(bytes, entry);
    t.codec = { fourcc: ase.fourcc, name: CODEC_NAMES[ase.fourcc] || `unknown (${ase.fourcc})` };
    t.channels = ase.channels;
    t.sampleSize = ase.sampleSize;
    t.sampleRate = ase.sampleRate;
    const esdsBox = ase.boxes.esds || (ase.boxes.wave ? findChild(bytes, ase.boxes.wave, 'esds') : null);
    if (esdsBox) {
      const e = parseEsds(bytes, esdsBox);
      t.esds = e;
      if (e.objectTypeIndication === 0x6b || e.objectTypeIndication === 0x69) {
        t.codec.name = 'MP3 (in MP4 container)';
        t.warnings.push('MP3 inside an MP4 container is poorly supported by some players');
      }
      if (e.channelConfiguration) t.channels = e.channelConfiguration;
      if (e.audioObjectType === 5 || e.audioObjectType === 29) t.codec.name = 'HE-AAC';
      if (e.avgBitrate) t.bitrate = e.avgBitrate;
    }
  }

  return t;
}

function gcd(a, b) { a = Math.abs(a); b = Math.abs(b); while (b) { [a, b] = [b, a % b]; } return a || 1; }
function simplifyRatio(w, h) {
  if (!w || !h) return null;
  const g = gcd(Math.round(w), Math.round(h)) || 1;
  return `${Math.round(w) / g}:${Math.round(h) / g}`;
}

/* ------------------------------------------------------------------ *
 * Main entry point                                                    *
 * ------------------------------------------------------------------ */

const BRAND_NAMES = {
  qt: 'QuickTime MOV', isom: 'MP4 (ISO base)', iso2: 'MP4 (ISO base v2)',
  mp41: 'MP4 v1', mp42: 'MP4 v2', avc1: 'MP4 (AVC)', M4V: 'iTunes M4V',
  M4A: 'iTunes M4A', dash: 'MP4 (DASH)', msdh: 'MP4 (DASH)', iso5: 'MP4 (ISO v5)',
  iso6: 'MP4 (ISO v6)', mmp4: 'Mobile MP4', '3gp4': '3GPP', 'hvc1': 'MP4 (HEVC)',
};

/**
 * Inspect an ISO BMFF file.
 * @param {{read:(o:number,l:number)=>Promise<Uint8Array>, size:number}} reader
 * @param {{name?:string, maxMoovBytes?:number}} opts
 */
async function inspect(reader, opts = {}) {
  const maxMoov = opts.maxMoovBytes ?? 64 * 1024 * 1024;
  const result = {
    ok: false,
    name: opts.name || null,
    fileSize: reader.size,
    container: null,
    brands: null,
    faststart: null,
    fragmented: false,
    duration: null,
    overallBitrate: null,
    video: [],
    audio: [],
    other: [],
    warnings: [],
    topLevelBoxes: [],
  };

  // Top-level scan. Only headers are read, so mdat costs nothing.
  let offset = 0, guard = 0, moovBox = null, mdatOffset = null, ftypBox = null;
  while (offset < reader.size && guard++ < 20000) {
    const h = await readHeader(reader, offset);
    if (!h) break;
    result.topLevelBoxes.push({ type: h.type, offset: h.offset, size: h.size });
    if (h.type === 'ftyp') ftypBox = h;
    if (h.type === 'moov') moovBox = h;
    if (h.type === 'mdat' && mdatOffset == null) mdatOffset = h.offset;
    if (h.type === 'moof') result.fragmented = true;
    offset = h.offset + h.size;
  }

  if (!ftypBox && !moovBox) {
    result.error = 'Not an ISO BMFF file (no ftyp or moov box found)';
    return result;
  }

  if (ftypBox) {
    const f = await reader.read(ftypBox.dataOffset, Math.min(ftypBox.dataSize, 256));
    const major = str4(f, 0).trim();
    const compatible = [];
    for (let o = 8; o + 4 <= f.length; o += 4) compatible.push(str4(f, o).trim());
    result.brands = { major, minor: u32(f, 4), compatible };
    result.container = BRAND_NAMES[major] || `MP4 (brand ${major})`;
  } else {
    result.container = 'QuickTime MOV (no ftyp)';
  }

  if (!moovBox) { result.error = 'No moov box: file is truncated or still being written'; return result; }
  if (moovBox.size > maxMoov) { result.error = `moov is ${moovBox.size} bytes, above the ${maxMoov} byte limit`; return result; }

  result.faststart = mdatOffset == null ? null : moovBox.offset < mdatOffset;
  if (result.faststart === false) result.warnings.push('moov sits after mdat (not faststart): slow to start over a network');
  if (result.fragmented) {
    result.incomplete = true;
    result.warnings.push('Fragmented MP4: sample tables live in moof boxes, so frame rate, frame count, keyframe spacing and bitrate cannot be read from the init segment');
  }

  const moovBytes = await reader.read(moovBox.dataOffset, moovBox.dataSize);
  const moov = { type: 'moov', start: 0, end: moovBytes.length };

  // mvhd
  let movieTimescale = 1000;
  const mvhd = findChild(moovBytes, moov, 'mvhd');
  if (mvhd) {
    const v = u8(moovBytes, mvhd.start);
    if (v === 1) { movieTimescale = u32(moovBytes, mvhd.start + 20); result.duration = u64(moovBytes, mvhd.start + 24) / movieTimescale; }
    else { movieTimescale = u32(moovBytes, mvhd.start + 12); result.duration = u32(moovBytes, mvhd.start + 16) / movieTimescale; }
  }
  result.movieTimescale = movieTimescale;

  if (findChild(moovBytes, moov, 'mvex')) {
    result.fragmented = true;
    result.incomplete = true;
    if (!result.warnings.some((w) => w.startsWith('Fragmented'))) {
      result.warnings.push('Fragmented MP4 signalled by mvex: frame rate, frame count, keyframe spacing and bitrate cannot be read from the init segment');
    }
  }

  for (const c of children(moovBytes, moov.start, moov.end)) {
    if (c.type !== 'trak') continue;
    const t = parseTrack(moovBytes, c, movieTimescale);
    if (!t) continue;
    if (t.handlerType === 'vide') result.video.push(t);
    else if (t.handlerType === 'soun') result.audio.push(t);
    else result.other.push(t);
  }

  if (result.duration && reader.size) {
    result.overallBitrate = Math.round((reader.size * 8) / result.duration);
  }

  // creation date from mvhd
  if (mvhd) {
    const v = u8(moovBytes, mvhd.start);
    const secs = v === 1 ? u64(moovBytes, mvhd.start + 4) : u32(moovBytes, mvhd.start + 4);
    if (secs > 0) {
      const d = new Date((secs - 2082844800) * 1000);   // 1904 epoch
      if (!Number.isNaN(d.getTime())) result.created = d.toISOString();
    }
  }

  // An edit list that skips into the media, or an empty edit, changes where playback starts.
  for (const t of [...result.video, ...result.audio]) {
    if (!t.editList || !t.editList.length) continue;
    const first = t.editList[0];
    if (first.empty) {
      t.startDelay = first.durationSeconds;
    } else {
      // Encoders write a small offset equal to the B-frame reorder delay on nearly
      // every file. Only flag an offset well beyond that.
      const fps = t.frameRate?.nominalFps || t.frameRate?.fps || 25;
      const floor = Math.max(0.5, 4 / fps);
      if (first.mediaTimeSeconds > floor) t.trimmedHeadSeconds = Number(first.mediaTimeSeconds.toFixed(3));
      else if (first.mediaTimeSeconds > 0) t.compositionOffsetSeconds = Number(first.mediaTimeSeconds.toFixed(3));
    }
    if (t.editList.some((e) => e.rate && Math.abs(e.rate - 1) > 0.001)) t.editListRateChange = true;
    const presented = t.editList.reduce((a, e) => a + (e.durationSeconds || 0), 0);
    if (t.duration && presented && Math.abs(presented - t.duration) > 0.05) {
      t.presentedDuration = Number(presented.toFixed(3));
    }
  }

  if (!result.video.length && !result.audio.length) result.warnings.push('No video or audio tracks found');
  result.ok = true;
  return result;
}

/** Convenience wrapper for a browser File or Blob. */
async function inspectFile(file, opts = {}) {
  return inspect(new BlobReader(file), { name: file.name, ...opts });
}

/** Convenience wrapper for raw bytes, e.g. a video pulled out of a PPTX. */
async function inspectBytes(bytes, name, opts = {}) {
  return inspect(new BytesReader(bytes), { name, ...opts });
}

/**
 * playback-risk.js
 * Turns the output of mp4-inspect.js into a ranked list of playback findings.
 *
 * The rules are deliberately separate from the parser so they can be tuned
 * without touching bitstream code. Everything here is opinion; the parser is fact.
 *
 * AEGFX / SlideSize
 */

const SEVERITY = { RED: 'red', AMBER: 'amber', INFO: 'info' };
const RANK = { red: 3, amber: 2, info: 1 };

/* ------------------------------------------------------------------ *
 * Target profiles                                                     *
 * ------------------------------------------------------------------ */

const TARGETS = {
  'powerpoint-win': {
    label: 'PowerPoint on Windows',
    note: 'The default show machine. Most conservative.',
    blockedCodecs: ['apco', 'apcs', 'apcn', 'apch', 'ap4h', 'ap4x', 'icpf', 'AVdn', 'AVdh',
      'Hap1', 'Hap5', 'HapY', 'HapM', 'HapA', 'CFHD', 'rle ', 'v210', 'v410', 'r210', '2vuy',
      'mjpa', 'mjpb', 'dvh5', 'dvh6', 'vp09', 'vp08', 'av01'],
    riskyCodecs: ['hvc1', 'hev1', 'mp4v', 'jpeg'],
    blockedAudio: ['dtsc', 'ec-3'],
    riskyAudio: ['ac-3', 'lpcm', 'in24', 'in32', 'fl32', 'twos', 'alac', 'Opus'],
    maxBitDepth: 8,
    allowedChroma: ['4:2:0'],
    maxWidth: 3840, maxHeight: 2160,
    maxBitrate: 60e6,
    maxGopSeconds: 5,
    allowVfr: false,
    requireLevelConformance: true,
  },
  'powerpoint-mac': {
    label: 'PowerPoint on macOS',
    note: 'VideoToolbox is far more forgiving. Use this only if the deck will never touch Windows.',
    blockedCodecs: ['Hap1', 'Hap5', 'HapY', 'HapM', 'HapA', 'AVdn', 'AVdh', 'CFHD', 'av01'],
    riskyCodecs: ['ap4x', 'v410', 'r210', 'vp09'],
    blockedAudio: ['dtsc'],
    riskyAudio: ['ec-3'],
    maxBitDepth: 10,
    allowedChroma: ['4:2:0', '4:2:2'],
    maxWidth: 4096, maxHeight: 4096,
    maxBitrate: 200e6,
    maxGopSeconds: 10,
    allowVfr: true,
    requireLevelConformance: false,
  },
  'media-server': {
    label: 'Media server (disguise / Watchout / Pixera / Resolume)',
    note: 'Wants all-intra codecs. Long-GOP H.264 is the problem here, not the fix.',
    blockedCodecs: [],
    riskyCodecs: ['avc1', 'avc3', 'hvc1', 'hev1', 'vp09', 'av01'],
    preferIntra: true,
    blockedAudio: [],
    riskyAudio: ['Opus'],
    maxBitDepth: 12,
    allowedChroma: ['4:2:0', '4:2:2', '4:4:4'],
    maxWidth: 16384, maxHeight: 16384,
    maxBitrate: 1200e6,
    maxGopSeconds: 1,
    allowVfr: false,
    requireLevelConformance: false,
  },
  'web': {
    label: 'Browser / web delivery',
    note: 'Baseline compatibility across Chrome, Safari, Firefox and Edge.',
    blockedCodecs: ['apco', 'apcs', 'apcn', 'apch', 'ap4h', 'ap4x', 'icpf', 'AVdn', 'AVdh',
      'Hap1', 'Hap5', 'HapY', 'HapM', 'HapA', 'CFHD', 'rle ', 'v210', 'r210', '2vuy', 'dvh5'],
    riskyCodecs: ['hvc1', 'hev1', 'av01', 'mp4v'],
    blockedAudio: ['dtsc', 'lpcm', 'twos', 'in24', 'in32'],
    riskyAudio: ['ac-3', 'ec-3', 'alac'],
    maxBitDepth: 8,
    allowedChroma: ['4:2:0'],
    maxWidth: 3840, maxHeight: 2160,
    maxBitrate: 40e6,
    maxGopSeconds: 5,
    allowVfr: true,
    requireLevelConformance: true,
    requireFaststart: true,
  },
};

/* ------------------------------------------------------------------ *
 * Assessment                                                          *
 * ------------------------------------------------------------------ */

/**
 * @param {object} info            result from mp4-inspect.inspect()
 * @param {object} opts
 * @param {string} opts.target     key into TARGETS
 * @param {number} [opts.projectFrameRate]  e.g. 25, 30, 50, 59.94
 * @param {number} [opts.slideWidth]  slide placeholder width in px, for upscale checks
 * @param {number} [opts.slideHeight]
 * @param {boolean} [opts.linked]  true if the PPTX references this as an external file
 */
function assess(info, opts = {}) {
  const target = TARGETS[opts.target] || TARGETS['powerpoint-win'];
  const f = [];
  const add = (severity, code, title, detail, fix) => f.push({ severity, code, title, detail, fix });

  if (!info.ok) {
    add(SEVERITY.RED, 'unreadable', 'File could not be parsed',
      info.error || 'The container structure is not readable.',
      'Re-export from the source application, or remux with: ffmpeg -i in.mp4 -c copy out.mp4');
    return finish(f, info, target);
  }

  if (opts.linked) {
    add(SEVERITY.RED, 'linked-media', 'Video is linked, not embedded',
      'PowerPoint holds only a path to this file. It will not travel with the deck and will fail on the show machine.',
      'In PowerPoint: Insert > Video > This Device, and pick Insert rather than Link to File.');
  }

  if (!info.video.length && !opts.linked) {
    add(SEVERITY.AMBER, 'no-video', 'No video track', 'The file contains no video track.', null);
  }

  for (const v of info.video) {
    const fourcc = v.codec?.fourcc;
    const label = v.codec?.name || fourcc;

    /* --- codec acceptance --- */
    if (target.blockedCodecs.includes(fourcc)) {
      add(SEVERITY.RED, 'codec-blocked', `${label} will not play`,
        `${target.label} cannot decode ${label}. This is the single most common cause of a black or missing video at showtime.`,
        ffmpegTranscode(v, opts));
    } else if (target.riskyCodecs.includes(fourcc)) {
      const why = fourcc === 'hvc1' || fourcc === 'hev1'
        ? 'HEVC needs the paid HEVC Video Extension on Windows, which is often missing on a hired show machine.'
        : target.preferIntra
          ? 'Long-GOP codecs decode poorly when a media server is scrubbing, looping or cueing mid-clip.'
          : 'Support is inconsistent across players and machines.';
      add(SEVERITY.AMBER, 'codec-risky', `${label} may not play`, why, ffmpegTranscode(v, opts));
    }

    if (target.preferIntra && v.allIntraCodec === false && !target.riskyCodecs.includes(fourcc)) {
      add(SEVERITY.INFO, 'not-intra', 'Not an all-intra codec',
        'Media servers prefer frame-independent codecs for reliable scrubbing.', null);
    }

    /* --- level conformance: the Mac screen recording bug --- */
    const lc = v.levelCheck;
    if (lc && lc.conformant === false && target.requireLevelConformance) {
      const parts = [];
      if (lc.frameSizeOk === false) {
        parts.push(`The frame is ${lc.macroblocks} macroblocks but level ${lc.declared} caps at ${lc.maxFS}. It needs level ${lc.requiredByFrameSize}.`);
      }
      if (lc.throughputOk === false) {
        parts.push(`Throughput is ${lc.macroblocksPerSecond} MB/s against a limit of ${lc.maxMBPS}.`);
      }
      add(SEVERITY.AMBER, 'level-mismatch', `Declared level ${lc.declared} is too low for this picture`,
        `${parts.join(' ')} Apple's decoder ignores this. Windows DXVA2 and D3D11VA size their decode buffers from the declared level, so hardware decode either refuses the stream or under-allocates and starts dropping frames.`,
        `Header-only fix, no re-encode: ffmpeg -i in.mp4 -c copy -bsf:v h264_metadata=level=${lc.requiredByThroughput || lc.requiredByFrameSize || '5'} out.mp4`);
    } else if (lc && lc.conformant === false) {
      add(SEVERITY.INFO, 'level-mismatch-info', `Declared level ${lc.declared} is below what the picture needs`,
        `Needs level ${lc.requiredByThroughput || lc.requiredByFrameSize}. Harmless on this target but will bite on Windows.`, null);
    }
    if (lc && lc.bitrateOk === false) {
      add(SEVERITY.INFO, 'level-bitrate', 'Bitrate exceeds the declared level',
        `Peak allowance for level ${lc.declared} is about ${fmtBitrate(lc.maxBitrateBps)}.`, null);
    }

    /* --- bit depth and chroma --- */
    if (v.bitDepth && v.bitDepth > target.maxBitDepth) {
      add(SEVERITY.RED, 'bit-depth', `${v.bitDepth}-bit video`,
        `${target.label} expects ${target.maxBitDepth}-bit. Deeper bit depths usually drop to software decode or fail outright.`,
        ffmpegTranscode(v, opts));
    }
    if (v.chroma && !target.allowedChroma.includes(v.chroma) && v.chroma !== 'monochrome') {
      add(SEVERITY.AMBER, 'chroma', `${v.chroma} chroma subsampling`,
        `${target.label} expects ${target.allowedChroma.join(' or ')}. ${v.chroma} disables hardware decode on most consumer GPUs.`,
        ffmpegTranscode(v, opts));
    }

    /* --- frame timing --- */
    const fr = v.frameRate || {};
    if (fr.mode === 'VFR' && !target.allowVfr) {
      add(SEVERITY.AMBER, 'vfr', 'Variable frame rate',
        `${fr.distinctIntervals} distinct frame intervals, spanning ${fmtFps(fr.minFps)} to ${fmtFps(fr.maxFps)}, with ${(fr.dominantShare * 100).toFixed(1)}% of frames at ${fmtFps(fr.nominalFps)}. Normal for screen recordings and phone footage, but it drifts out of sync in PowerPoint and breaks media server timelines.`,
        `ffmpeg -i in.mp4 -fps_mode cfr -r ${opts.projectFrameRate || fr.nominalFps || 25} -c:v libx264 -crf 18 -preset medium -pix_fmt yuv420p out.mp4`);
    } else if (fr.mode === 'near-CFR' && !target.allowVfr) {
      add(SEVERITY.INFO, 'near-cfr', 'Almost constant frame rate',
        `${fr.distinctIntervals} distinct frame intervals, but ${(fr.dominantShare * 100).toFixed(1)}% of frames share one. Usually harmless.`, null);
    }

    const fps = fr.nominalFps || fr.fps || fr.avgFps;
    if (opts.projectFrameRate && fps) {
      const ratio = opts.projectFrameRate / fps;
      const clean = Math.abs(ratio - Math.round(ratio)) < 0.001;
      if (!clean) {
        add(SEVERITY.AMBER, 'framerate-mismatch', `${fmtFps(fps)} does not divide into the ${opts.projectFrameRate} fps project rate`,
          `Every frame will be held for an uneven number of output frames, giving a regular judder that no amount of decoding headroom will fix.`,
          `ffmpeg -i in.mp4 -r ${opts.projectFrameRate} -c:v libx264 -crf 18 -preset medium -pix_fmt yuv420p out.mp4`);
      }
    }
    if (fps && fps > 60) {
      add(SEVERITY.INFO, 'high-framerate', `${fmtFps(fps)} frame rate`,
        'High frame rates raise decode load and rarely survive projection intact.', null);
    }

    /* --- geometry --- */
    const w = v.cropped?.width || v.coded?.width;
    const h = v.cropped?.height || v.coded?.height;
    if (w && h) {
      if (w > target.maxWidth || h > target.maxHeight) {
        add(SEVERITY.AMBER, 'oversize', `${w}x${h} exceeds the ${target.maxWidth}x${target.maxHeight} target`,
          'Above the resolution the target reliably decodes.', ffmpegTranscode(v, opts));
      }
      if (w % 2 || h % 2) {
        add(SEVERITY.AMBER, 'odd-dimensions', `${w}x${h} has an odd dimension`,
          'Odd dimensions are illegal for 4:2:0 chroma and break many encoders and decoders.',
          `ffmpeg -i in.mp4 -vf "pad=ceil(iw/2)*2:ceil(ih/2)*2" -c:v libx264 -crf 18 -pix_fmt yuv420p out.mp4`);
      } else if (h % 16 || w % 16) {
        add(SEVERITY.INFO, 'unaligned-dimensions', `${w}x${h} is not a multiple of 16`,
          'Legal, but some hardware decoders take a slower unaligned path. Common in screen recordings.', null);
      }
      const standard = ['1920x1080', '1280x720', '3840x2160', '2560x1440', '1080x1920', '720x1280', '1920x1200', '1024x768'];
      if (!standard.includes(`${w}x${h}`) && !(h % 16) && !(w % 16)) {
        add(SEVERITY.INFO, 'nonstandard-resolution', `${w}x${h} is a non-standard resolution`,
          'Not a fault, but worth knowing before it gets scaled onto a screen.', null);
      }
    }

    if (v.squarePixels === false && v.pixelAspect) {
      add(SEVERITY.AMBER, 'non-square-pixels', `Non-square pixels (${v.pixelAspect.h}:${v.pixelAspect.v})`,
        `Stored at ${w}x${h} but intended to display at ${v.display.width}x${v.display.height}. PowerPoint ignores pixel aspect, so this will show up stretched or squashed.`,
        `ffmpeg -i in.mp4 -vf "scale=${v.display?.width}:${v.display?.height},setsar=1" -c:v libx264 -crf 18 -pix_fmt yuv420p out.mp4`);
    }

    if (v.rotation) {
      add(SEVERITY.AMBER, 'rotation', `Rotation flag of ${v.rotation} degrees`,
        'The picture is stored rotated and relies on the player honouring the track matrix. PowerPoint does not always, so it can appear sideways.',
        `ffmpeg -i in.mp4 -vf "transpose=1" -metadata:s:v rotate=0 -c:v libx264 -crf 18 -pix_fmt yuv420p out.mp4`);
    }

    /* --- edit lists --- */
    if (v.trimmedHeadSeconds) {
      add(SEVERITY.AMBER, 'edit-list-trim', `Edit list skips the first ${v.trimmedHeadSeconds}s`,
        'The container tells the player to start part-way into the media. Players that honour the edit list and players that ignore it will show different first frames, and the clip will look like it starts late on one machine and early on another.',
        `ffmpeg -i in.mp4 -ignore_editlist 1 -c copy out.mp4    # or bake it in with a re-encode`);
    }
    if (v.compositionOffsetSeconds) {
      add(SEVERITY.INFO, 'composition-offset', `Edit list offset of ${v.compositionOffsetSeconds}s`,
        'Matches the B-frame reorder delay. Normal, written by most encoders.', null);
    }
    if (v.startDelay) {
      add(SEVERITY.AMBER, 'edit-list-delay', `Edit list holds a ${v.startDelay.toFixed(2)}s empty gap before the video starts`,
        'An empty edit at the head. Some players show black, some skip it, some drift the audio against it.', null);
    }
    if (v.editListRateChange) {
      add(SEVERITY.INFO, 'edit-list-rate', 'Edit list changes playback rate',
        'A non-unity rate in the edit list. Rarely honoured consistently.', null);
    }

    /* --- interlacing --- */
    if (v.interlaced?.flag) {
      add(SEVERITY.AMBER, 'interlaced', `Interlaced (${v.interlaced.fieldOrder})`,
        'Progressive displays and projectors will show combing on motion unless something deinterlaces it first.',
        `ffmpeg -i in.mp4 -vf "yadif=1" -c:v libx264 -crf 18 -pix_fmt yuv420p out.mp4`);
    }

    /* --- GOP structure --- */
    const kf = v.keyframes;
    if (kf && !kf.allIntra && kf.intervalSeconds != null && kf.intervalSeconds > target.maxGopSeconds) {
      add(SEVERITY.AMBER, 'long-gop', `Keyframe every ${kf.intervalSeconds}s`,
        `Only ${kf.count} keyframes across the clip. Straight playback is fine, but cueing into the middle, scrubbing or looping will hitch while the decoder walks back to the last keyframe.`,
        `ffmpeg -i in.mp4 -c:v libx264 -crf 18 -preset medium -g ${Math.round((opts.projectFrameRate || fps || 25) * 2)} -keyint_min ${Math.round(opts.projectFrameRate || fps || 25)} -pix_fmt yuv420p out.mp4`);
    }

    /* --- bitrate --- */
    if (v.bitrate && v.bitrate > target.maxBitrate) {
      add(SEVERITY.AMBER, 'high-bitrate', `${fmtBitrate(v.bitrate)} video bitrate`,
        `Above the ${fmtBitrate(target.maxBitrate)} this target handles comfortably. Playing off a laptop SSD alongside PowerPoint rendering is where this bites.`,
        `ffmpeg -i in.mp4 -c:v libx264 -crf 20 -preset medium -maxrate 25M -bufsize 50M -pix_fmt yuv420p out.mp4`);
    }
    if (v.maxSampleBytes && v.totalBytes && v.sampleCount) {
      const avgFrame = v.totalBytes / v.sampleCount;
      if (v.maxSampleBytes > avgFrame * 12) {
        add(SEVERITY.INFO, 'peaky-bitrate', 'Very uneven frame sizes',
          `Largest frame is ${(v.maxSampleBytes / 1024).toFixed(0)} KB against an average of ${(avgFrame / 1024).toFixed(1)} KB, a ${(v.maxSampleBytes / avgFrame).toFixed(0)}x peak. Usually just a keyframe, but big spikes can stall a decoder even when the average looks safe.`, null);
      }
    }

    /* --- colour --- */
    if (v.colour?.transferIdc === 16 || v.colour?.transferIdc === 18) {
      add(SEVERITY.RED, 'hdr', `HDR content (${v.colour.transfer})`,
        'On an SDR projector or LED wall this comes out washed out and grey. PowerPoint does no tone mapping.',
        `ffmpeg -i in.mp4 -vf "zscale=t=linear:npl=100,format=gbrpf32le,zscale=p=bt709,tonemap=hable,zscale=t=bt709:m=bt709:r=tv,format=yuv420p" -c:v libx264 -crf 18 out.mp4`);
    } else if (!v.colour) {
      add(SEVERITY.INFO, 'no-colour-tags', 'No colour signalling',
        'No primaries, transfer or matrix tagged. Players will assume BT.709, which is usually right for HD but can shift colour on SD-sized material.', null);
    } else if (v.colour.matrixIdc === 6 || v.colour.matrixIdc === 5) {
      if (h && h > 576) {
        add(SEVERITY.AMBER, 'colour-mismatch', 'BT.601 matrix on an HD picture',
          'HD content tagged as standard definition. Expect a visible shift in saturation between players.',
          `ffmpeg -i in.mp4 -c copy -bsf:v h264_metadata=matrix_coefficients=1:colour_primaries=1:transfer_characteristics=1 out.mp4`);
      }
    }
    if (v.colour?.fullRange === true) {
      add(SEVERITY.INFO, 'full-range', 'Full range (0-255) levels',
        'Common in screen recordings. Some players clamp to 16-235, which crushes blacks and clips whites.', null);
    }

    /* --- upscale on the slide --- */
    if (opts.slideWidth && opts.slideHeight && w && h) {
      const scale = Math.max(opts.slideWidth / w, opts.slideHeight / h);
      if (scale > 1.35) {
        add(SEVERITY.AMBER, 'upscaled', `Scaled ${scale.toFixed(1)}x on the slide`,
          `A ${w}x${h} source placed in a ${Math.round(opts.slideWidth)}x${Math.round(opts.slideHeight)} frame. It will look soft on a big screen.`, null);
      }
    }
  }

  /* --- audio --- */
  for (const a of info.audio) {
    const fourcc = a.codec?.fourcc;
    if (target.blockedAudio.includes(fourcc)) {
      add(SEVERITY.RED, 'audio-blocked', `${a.codec.name} audio will not play`,
        `${target.label} cannot decode this. Video may play silently, or the whole file may fail.`,
        `ffmpeg -i in.mp4 -c:v copy -c:a aac -b:a 192k out.mp4`);
    } else if (target.riskyAudio.includes(fourcc)) {
      add(SEVERITY.AMBER, 'audio-risky', `${a.codec.name} audio may not play`,
        'Support varies by machine and codec pack.', `ffmpeg -i in.mp4 -c:v copy -c:a aac -b:a 192k out.mp4`);
    }
    if (a.channels > 2) {
      add(SEVERITY.INFO, 'multichannel', `${a.channels} channel audio`,
        'Will be downmixed, or partly lost, on a stereo PA feed.', null);
    }
    if (a.sampleRate && ![44100, 48000, 96000].includes(a.sampleRate)) {
      add(SEVERITY.INFO, 'odd-samplerate', `${a.sampleRate} Hz sample rate`,
        'Non-standard. Resampling on the fly can glitch on some machines.', null);
    }
  }
  if (info.video.length && !info.audio.length) {
    add(SEVERITY.INFO, 'no-audio', 'No audio track', 'Silent clip. Worth confirming that is intentional.', null);
  }

  /* --- container --- */
  if (target.requireFaststart && info.faststart === false) {
    add(SEVERITY.AMBER, 'not-faststart', 'moov atom is at the end of the file',
      'Playback cannot start until the whole file has downloaded.',
      `ffmpeg -i in.mp4 -c copy -movflags +faststart out.mp4`);
  }
  if (info.fragmented) {
    add(SEVERITY.AMBER, 'fragmented', 'Fragmented MP4, inspection incomplete',
      'Sample tables live in the fragments rather than the header, so frame rate, frame count, keyframe spacing and true bitrate could not be read. Some desktop players and editing tools also handle these badly. Remux before trusting any of the timing figures above.',
      `ffmpeg -i in.mp4 -c copy -movflags +faststart out.mp4`);
  }
  if (info.brands?.major === 'qt' && opts.target === 'powerpoint-win') {
    add(SEVERITY.INFO, 'mov-container', 'QuickTime MOV container',
      'Modern PowerPoint reads MOV, but an MP4 wrapper removes one variable.',
      `ffmpeg -i in.mov -c copy -movflags +faststart out.mp4`);
  }

  for (const w of info.warnings) add(SEVERITY.INFO, 'parser-warning', w, '', null);
  for (const v of info.video) for (const w of (v.warnings || [])) add(SEVERITY.INFO, 'track-warning', w, '', null);

  return finish(f, info, target);
}

function finish(findings, info, target) {
  findings.sort((a, b) => RANK[b.severity] - RANK[a.severity]);
  const reds = findings.filter((x) => x.severity === SEVERITY.RED).length;
  const ambers = findings.filter((x) => x.severity === SEVERITY.AMBER).length;
  const verdict = reds ? SEVERITY.RED : ambers ? SEVERITY.AMBER : 'green';
  const summary = reds
    ? `Will not play reliably on ${target.label}.`
    : ambers
      ? `Should play, but ${ambers} thing${ambers > 1 ? 's' : ''} could bite.`
      : `No problems found for ${target.label}.`;
  return { verdict, summary, target: target.label, counts: { red: reds, amber: ambers, info: findings.length - reds - ambers }, findings };
}

/* ------------------------------------------------------------------ *
 * Fix suggestions                                                     *
 * ------------------------------------------------------------------ */

function ffmpegTranscode(v, opts) {
  const fps = opts.projectFrameRate ? ` -r ${opts.projectFrameRate}` : '';
  return `ffmpeg -i in.mp4 -c:v libx264 -profile:v high -crf 18 -preset medium${fps} -g ${Math.round((opts.projectFrameRate || 50))} -pix_fmt yuv420p -c:a aac -b:a 192k -movflags +faststart out.mp4`;
}

/**
 * Build a single conforming command for a file, folding every fix into one pass.
 * Preferred over stacking the per-finding suggestions, which would re-encode repeatedly.
 */
function buildFixCommand(info, assessment, opts = {}) {
  const v = info.video[0];
  if (!v) return null;
  const onlyLevel = assessment.findings.length > 0 &&
    assessment.findings.filter((x) => x.severity !== SEVERITY.INFO).every((x) => x.code === 'level-mismatch');
  if (onlyLevel) {
    const lc = v.levelCheck;
    const lvl = lc.requiredByThroughput || lc.requiredByFrameSize || '5';
    return { reencode: false, command: `ffmpeg -i "${info.name}" -c copy -bsf:v h264_metadata=level=${lvl} -movflags +faststart "fixed_${info.name}"` };
  }

  const filters = [];
  if (v.interlaced?.flag) filters.push('yadif=1');
  if (v.squarePixels === false) filters.push(`scale=${v.display.width}:${v.display.height}`);
  const maxW = opts.maxWidth || 1920, maxH = opts.maxHeight || 1080;
  const w = v.cropped?.width || v.coded?.width;
  const h = v.cropped?.height || v.coded?.height;
  if (w > maxW || h > maxH) filters.push(`scale='min(${maxW},iw)':'min(${maxH},ih)':force_original_aspect_ratio=decrease`);
  filters.push('scale=trunc(iw/2)*2:trunc(ih/2)*2', 'setsar=1');

  const fps = opts.projectFrameRate || 25;
  const parts = [
    `ffmpeg -i "${info.name}"`,
    `-vf "${filters.join(',')}"`,
    `-c:v libx264 -profile:v high -level 4.2 -crf 18 -preset medium`,
    `-r ${fps} -fps_mode cfr -g ${fps * 2} -keyint_min ${fps}`,
    `-pix_fmt yuv420p -colorspace bt709 -color_primaries bt709 -color_trc bt709`,
    info.audio.length ? `-c:a aac -b:a 192k -ar 48000 -ac 2` : `-an`,
    `-movflags +faststart "fixed_${(info.name || 'video.mp4').replace(/\.[^.]+$/, '')}.mp4"`,
  ];
  return { reencode: true, command: parts.join(' ') };
}

/* ------------------------------------------------------------------ *
 * Formatting helpers                                                  *
 * ------------------------------------------------------------------ */

function fmtBitrate(bps) {
  if (!bps) return 'unknown';
  if (bps >= 1e6) return `${(bps / 1e6).toFixed(2)} Mbps`;
  return `${Math.round(bps / 1000)} kbps`;
}

function fmtFps(fps) {
  if (!fps) return 'unknown';
  return `${Number(fps.toFixed(3))} fps`;
}

function fmtDuration(seconds) {
  if (seconds == null) return 'unknown';
  const s = Math.floor(seconds % 60), m = Math.floor((seconds / 60) % 60), h = Math.floor(seconds / 3600);
  const frac = (seconds % 1).toFixed(2).slice(1);
  return h ? `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}` : `${m}:${String(s).padStart(2, '0')}${frac}`;
}

function fmtBytes(n) {
  if (n == null) return 'unknown';
  const u = ['B', 'KB', 'MB', 'GB', 'TB'];
  let i = 0;
  while (n >= 1024 && i < u.length - 1) { n /= 1024; i++; }
  return `${n.toFixed(i ? 1 : 0)} ${u[i]}`;
}

/** Flat key/value summary, handy for a report table or a CSV export. */
function summarise(info) {
  const v = info.video[0];
  const a = info.audio[0];
  return {
    File: info.name,
    Size: fmtBytes(info.fileSize),
    Container: info.container,
    Brands: info.brands ? [info.brands.major, ...info.brands.compatible].filter(Boolean).join(' ') : null,
    Duration: fmtDuration(info.duration),
    'Overall bitrate': fmtBitrate(info.overallBitrate),
    Faststart: info.faststart == null ? 'n/a' : info.faststart ? 'yes' : 'no',
    'Video codec': v ? `${v.codec.name} (${v.codec.fourcc})` : 'none',
    'Codec string': v?.codec?.string || null,
    Profile: v?.profile || null,
    Level: v?.level || null,
    'Level conformant': v?.levelCheck ? (v.levelCheck.conformant ? 'yes' : `no, needs ${v.levelCheck.requiredByThroughput || v.levelCheck.requiredByFrameSize}`) : null,
    Resolution: v ? `${v.cropped?.width || v.coded.width}x${v.cropped?.height || v.coded.height}` : null,
    'Display size': v ? `${v.display.width}x${v.display.height}` : null,
    'Pixel aspect': v?.pixelAspect ? `${v.pixelAspect.h}:${v.pixelAspect.v} (${v.pixelAspect.source})` : '1:1 (square)',
    'Display aspect': v?.displayAspect || null,
    Rotation: v?.rotation ? `${v.rotation} degrees` : 'none',
    'Bit depth': v?.bitDepth ? `${v.bitDepth}-bit${v.bitDepthInferred ? ' (inferred)' : ''}` : null,
    Chroma: v?.chroma || null,
    'Frame rate': v ? `${fmtFps(v.frameRate.nominalFps || v.frameRate.fps)} ${v.frameRate.mode}${v.frameRate.mode !== 'CFR' ? ` (${fmtFps(v.frameRate.minFps)} to ${fmtFps(v.frameRate.maxFps)}, avg ${fmtFps(v.frameRate.avgFps)})` : ''}` : null,
    Frames: v?.sampleCount || null,
    Scan: v?.interlaced ? (v.interlaced.flag ? `interlaced, ${v.interlaced.fieldOrder}` : 'progressive') : 'progressive',
    'B-frames': v ? (v.hasBFrames ? 'yes' : 'no') : null,
    Keyframes: v?.keyframes ? (v.keyframes.allIntra ? 'all-intra' : `${v.keyframes.count}, every ${v.keyframes.intervalSeconds}s`) : null,
    'Video bitrate': v ? fmtBitrate(v.bitrate) : null,
    'Media duration': v && Math.abs((v.duration || 0) - (info.duration || 0)) > 0.05 ? `${fmtDuration(v.duration)} (container presents ${fmtDuration(info.duration)})` : null,
    'Edit list': v?.trimmedHeadSeconds ? `skips first ${v.trimmedHeadSeconds}s` : (v?.startDelay ? `${v.startDelay.toFixed(2)}s empty gap` : (v?.editList?.length ? 'present, no offset' : 'none')),
    Colour: v?.colour ? `${v.colour.primaries} / ${v.colour.transfer} / ${v.colour.matrix}${v.colour.fullRange ? ' full range' : ''}` : 'untagged',
    'Audio codec': a ? `${a.codec.name} (${a.codec.fourcc})` : 'none',
    Audio: a ? `${a.channels} ch, ${a.sampleRate} Hz` : null,
  };
}

/**
 * pptx-media.js
 * Pulls every video and audio reference out of a PPTX and reports how it is
 * attached: embedded, linked, or orphaned. Feeds each embedded file straight
 * into mp4-inspect.js.
 *
 * Zip-library agnostic. Supply an adapter with:
 *   { list(): string[], read(path): Promise<Uint8Array>, readText(path): Promise<string> }
 * Adapters for JSZip and zip.js are at the bottom.
 *
 * XML handling is a deliberate block-scan rather than a DOM parse. PPTX XML is
 * machine generated by PowerPoint, so the structure is predictable, and this
 * keeps the module dependency-free and identical in Node and the browser.
 *
 * AEGFX / SlideSize
 */


const EMU_PER_INCH = 914400;
const VIDEO_EXT = /\.(mp4|m4v|mov|avi|wmv|mkv|webm|mpg|mpeg|m2v|mts|m2ts|flv|3gp|ogv)$/i;
const AUDIO_EXT = /\.(mp3|m4a|wav|aiff?|wma|aac|flac|ogg|oga|mid|midi)$/i;
const INSPECTABLE = /\.(mp4|m4v|mov|m4a)$/i;

/* ------------------------------------------------------------------ *
 * Tiny XML helpers                                                    *
 * ------------------------------------------------------------------ */

function attrs(tagText) {
  const out = {};
  const re = /([\w:.-]+)\s*=\s*"([^"]*)"/g;
  let m;
  while ((m = re.exec(tagText))) out[m[1]] = decodeEntities(m[2]);
  return out;
}

function decodeEntities(s) {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, d) => String.fromCharCode(+d))
    .replace(/&amp;/g, '&');
}

/** All self-closing or opening tags of a given name, as attribute objects. */
function findTags(xml, name) {
  const re = new RegExp(`<${name}\\b([^>]*?)/?>`, 'g');
  const out = [];
  let m;
  while ((m = re.exec(xml))) out.push(attrs(m[1]));
  return out;
}

/** Text of each <name>...</name> block, nesting-aware for one level of the same tag. */
function findBlocks(xml, name) {
  const out = [];
  const open = new RegExp(`<${name}(\\s[^>]*)?>`, 'g');
  let m;
  while ((m = open.exec(xml))) {
    const start = m.index;
    let depth = 1;
    const scan = new RegExp(`<(/?)${name}(?:\\s[^>]*)?(/?)>`, 'g');
    scan.lastIndex = open.lastIndex;
    let s;
    while (depth > 0 && (s = scan.exec(xml))) {
      if (s[2] === '/') continue;
      depth += s[1] === '/' ? -1 : 1;
    }
    const end = depth === 0 && s ? s.index + s[0].length : xml.length;
    out.push(xml.slice(start, end));
    open.lastIndex = end;
  }
  return out;
}

/* ------------------------------------------------------------------ *
 * Relationship handling                                               *
 * ------------------------------------------------------------------ */

function relsPathFor(partPath) {
  const i = partPath.lastIndexOf('/');
  return `${partPath.slice(0, i)}/_rels/${partPath.slice(i + 1)}.rels`;
}

/** Resolve a relationship Target against the part that declared it. */
function resolveTarget(partPath, target) {
  if (/^[a-zA-Z]+:/.test(target) || target.startsWith('\\\\') || target.startsWith('/')) return target;
  const base = partPath.slice(0, partPath.lastIndexOf('/')).split('/');
  for (const seg of target.split('/')) {
    if (seg === '.') continue;
    else if (seg === '..') base.pop();
    else base.push(seg);
  }
  return base.join('/');
}

async function loadRels(zip, partPath) {
  const p = relsPathFor(partPath);
  const map = new Map();
  if (!zip.list().includes(p)) return map;
  const xml = await zip.readText(p);
  for (const r of findTags(xml, 'Relationship')) {
    map.set(r.Id, {
      id: r.Id,
      type: r.Type || '',
      target: r.Target || '',
      external: (r.TargetMode || '') === 'External',
      resolved: (r.TargetMode || '') === 'External' ? r.Target : resolveTarget(partPath, r.Target || ''),
    });
  }
  return map;
}

/* ------------------------------------------------------------------ *
 * Slide scanning                                                      *
 * ------------------------------------------------------------------ */

/** Pull the media references out of one slide's XML. */
function scanSlide(xml, rels, slideNumber, partPath) {
  const items = [];

  for (const pic of [...findBlocks(xml, 'p:pic'), ...findBlocks(xml, 'p:graphicFrame')]) {
    const nv = findTags(pic, 'p:cNvPr')[0] || {};
    const videoFile = findTags(pic, 'a:videoFile')[0];
    const audioFile = findTags(pic, 'a:audioFile')[0];
    const media = findTags(pic, 'p14:media')[0];
    const ref = videoFile || audioFile || media;
    if (!ref) continue;

    const rid = ref['r:link'] || ref['r:embed'] || ref['r:id'];
    const rel = rid ? rels.get(rid) : null;

    const ext = findTags(pic, 'a:ext').find((e) => e.cx && e.cy);
    const off = findTags(pic, 'a:off')[0];
    const xfrm = findTags(pic, 'a:xfrm')[0] || {};

    const trim = findTags(pic, 'p14:trim')[0];
    const nvPr = findTags(pic, 'p:nvPr')[0] || {};
    const hasPoster = /<a:blip\b/.test(pic);

    items.push({
      slide: slideNumber,
      part: partPath,
      name: nv.name || null,
      kind: audioFile ? 'audio' : 'video',
      relId: rid || null,
      linked: rel ? rel.external : null,
      target: rel ? rel.target : null,
      path: rel && !rel.external ? rel.resolved : null,
      displayEmu: ext ? { cx: Number(ext.cx), cy: Number(ext.cy) } : null,
      positionEmu: off ? { x: Number(off.x), y: Number(off.y) } : null,
      rotationDeg: xfrm.rot ? Number(xfrm.rot) / 60000 : 0,
      flipH: xfrm.flipH === '1',
      flipV: xfrm.flipV === '1',
      trimmed: trim ? { start: trim.st || null, end: trim.end || null } : null,
      posterFrame: hasPoster,
      loop: /\bloop="1"|<p:cMediaNode[^>]*\bloop="1"/.test(pic) || /playLst/.test(pic),
      hidden: nvPr.isPhoto === '1' ? false : undefined,
    });
  }

  return items;
}

/* ------------------------------------------------------------------ *
 * Main entry point                                                    *
 * ------------------------------------------------------------------ */

/**
 * @param {{list:()=>string[], read:(p:string)=>Promise<Uint8Array>, readText:(p:string)=>Promise<string>}} zip
 * @param {{inspect?:boolean, slidePixelWidth?:number, maxInspectBytes?:number}} opts
 */
async function scanPptx(zip, opts = {}) {
  const doInspect = opts.inspect !== false;
  const slidePx = opts.slidePixelWidth || 1920;
  const names = zip.list();

  const result = {
    slideCount: 0,
    slideSizeEmu: null,
    slideSizeInches: null,
    slideAspect: null,
    media: [],
    orphans: [],
    mediaPartCount: 0,
    mediaBytes: 0,
    totalBytes: names.reduce((a, n) => a + (zip.sizeOf ? zip.sizeOf(n) || 0 : 0), 0) || null,
    warnings: [],
  };

  // Slide dimensions, for the upscale check
  if (names.includes('ppt/presentation.xml')) {
    const pres = await zip.readText('ppt/presentation.xml');
    const sz = findTags(pres, 'p:sldSz')[0];
    if (sz) {
      const cx = Number(sz.cx), cy = Number(sz.cy);
      result.slideSizeEmu = { cx, cy };
      result.slideSizeInches = { w: +(cx / EMU_PER_INCH).toFixed(2), h: +(cy / EMU_PER_INCH).toFixed(2) };
      result.slideAspect = +(cx / cy).toFixed(4);
    }
  } else {
    result.warnings.push('No ppt/presentation.xml: this may not be a PPTX');
  }

  // Slides in presentation order where possible, numeric order otherwise
  const slideParts = names.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));
  result.slideCount = slideParts.length;

  const otherParts = names.filter((n) =>
    /^ppt\/(notesSlides\/notesSlide\d+|slideLayouts\/slideLayout\d+|slideMasters\/slideMaster\d+)\.xml$/.test(n));

  const referenced = new Set();

  for (const part of [...slideParts, ...otherParts]) {
    const num = /slides\/slide(\d+)/.test(part) ? Number(part.match(/slide(\d+)/)[1]) : null;
    const xml = await zip.readText(part);
    if (!/(a:videoFile|a:audioFile|p14:media)/.test(xml)) continue;
    const rels = await loadRels(zip, part);
    for (const item of scanSlide(xml, rels, num, part)) {
      if (item.path) referenced.add(item.path);
      result.media.push(item);
    }
  }

  // Anything in ppt/media that is a video or audio file
  const mediaParts = names.filter((n) => n.startsWith('ppt/media/'));
  for (const p of mediaParts) {
    if (!VIDEO_EXT.test(p) && !AUDIO_EXT.test(p)) continue;
    result.mediaPartCount++;
    if (zip.sizeOf) result.mediaBytes += zip.sizeOf(p) || 0;
    if (!referenced.has(p)) {
      result.orphans.push({ path: p, bytes: zip.sizeOf ? zip.sizeOf(p) : null });
    }
  }

  // Deduplicate: one media part can be placed on several slides
  const byPath = new Map();
  for (const item of result.media) {
    if (!item.path) continue;
    if (!byPath.has(item.path)) byPath.set(item.path, []);
    byPath.get(item.path).push(item);
  }

  // Inspect each distinct embedded file once, then attach the result to every placement
  if (doInspect) {
    for (const [path, items] of byPath) {
      const filename = path.split('/').pop();
      if (!INSPECTABLE.test(path)) {
        for (const it of items) {
          it.inspection = { ok: false, name: filename, error: `${filename.split('.').pop().toUpperCase()} is not an ISO BMFF container, so it cannot be inspected in the browser` };
        }
        continue;
      }
      let bytes;
      try {
        bytes = await zip.read(path);
      } catch (e) {
        for (const it of items) it.inspection = { ok: false, name: filename, error: `Could not read from the archive: ${e.message}` };
        continue;
      }
      let info;
      try {
        info = await inspectBytes(bytes, filename);
      } catch (e) {
        info = { ok: false, name: filename, error: `Parse failed: ${e.message}` };
      }
      for (const it of items) {
        it.inspection = info;
        it.bytes = bytes.length;
        // Native size against the size it is placed at on the slide
        const v = info.video && info.video[0];
        if (v && it.displayEmu && result.slideSizeEmu) {
          const pxPerEmu = slidePx / result.slideSizeEmu.cx;
          it.slidePx = {
            width: Math.round(it.displayEmu.cx * pxPerEmu),
            height: Math.round(it.displayEmu.cy * pxPerEmu),
          };
          const w = v.cropped?.width || v.coded?.width;
          const h = v.cropped?.height || v.coded?.height;
          if (w && h) {
            it.scaleFactor = +Math.max(it.slidePx.width / w, it.slidePx.height / h).toFixed(2);
            const srcAspect = w / h;
            const boxAspect = it.displayEmu.cx / it.displayEmu.cy;
            if (Math.abs(srcAspect - boxAspect) / srcAspect > 0.02) {
              it.aspectDistortion = +(boxAspect / srcAspect).toFixed(3);
            }
          }
        }
      }
    }
  }

  const linked = result.media.filter((m) => m.linked);
  if (linked.length) {
    result.warnings.push(`${linked.length} media reference${linked.length > 1 ? 's are' : ' is'} linked rather than embedded and will not travel with the deck`);
  }

  return result;
}

/* ------------------------------------------------------------------ *
 * Deck-level report                                                   *
 * ------------------------------------------------------------------ */

/**
 * Run the risk rules over every item in a scan and roll it up.
 * Pass the `assess` function from playback-risk.js so this module stays
 * independent of the rule set.
 */
function reportPptx(scan, assess, opts = {}) {
  const rows = [];
  for (const item of scan.media) {
    const risk = item.inspection
      ? assess(item.inspection, {
        ...opts,
        linked: item.linked === true,
        slideWidth: item.slidePx?.width,
        slideHeight: item.slidePx?.height,
      })
      : assess({ ok: true, name: item.target || item.name, video: [], audio: [], warnings: [], fileSize: null, faststart: null },
        { ...opts, linked: item.linked === true });

    if (item.aspectDistortion) {
      risk.findings.unshift({
        severity: 'amber',
        code: 'aspect-distortion',
        title: `Stretched ${item.aspectDistortion > 1 ? 'wider' : 'taller'} on the slide`,
        detail: `The placeholder is ${Math.abs((item.aspectDistortion - 1) * 100).toFixed(0)}% off the source aspect ratio, so the picture is distorted.`,
        fix: 'In PowerPoint, right-click the video, Size and Position, and reset the aspect ratio.',
      });
      if (risk.verdict === 'green') risk.verdict = 'amber';
    }

    rows.push({
      slide: item.slide,
      name: item.name,
      file: item.path ? item.path.split('/').pop() : item.target,
      linked: item.linked,
      bytes: item.bytes ?? null,
      scaleFactor: item.scaleFactor ?? null,
      posterFrame: item.posterFrame,
      trimmed: item.trimmed,
      verdict: risk.verdict,
      summary: risk.summary,
      findings: risk.findings,
    });
  }

  rows.sort((a, b) => (a.slide ?? 999) - (b.slide ?? 999));

  const reds = rows.filter((r) => r.verdict === 'red').length;
  const ambers = rows.filter((r) => r.verdict === 'amber').length;

  return {
    verdict: reds ? 'red' : ambers ? 'amber' : 'green',
    slideCount: scan.slideCount,
    mediaCount: rows.length,
    linkedCount: rows.filter((r) => r.linked).length,
    orphans: scan.orphans,
    orphanBytes: scan.orphans.reduce((a, o) => a + (o.bytes || 0), 0),
    counts: { red: reds, amber: ambers, green: rows.length - reds - ambers },
    summary: reds
      ? `${reds} of ${rows.length} clips will not play reliably.`
      : ambers
        ? `All ${rows.length} clips should play, but ${ambers} need a look.`
        : `All ${rows.length} clips look fine.`,
    rows,
    warnings: scan.warnings,
  };
}

/* ------------------------------------------------------------------ *
 * Zip adapters                                                        *
 * ------------------------------------------------------------------ */

/** JSZip. Simple, but loads the whole archive into memory. Fine under ~200 MB. */
function jszipAdapter(zip) {
  return {
    list: () => Object.keys(zip.files).filter((n) => !zip.files[n].dir),
    sizeOf: (p) => zip.files[p]?._data?.uncompressedSize ?? null,
    read: (p) => zip.file(p).async('uint8array'),
    readText: (p) => zip.file(p).async('string'),
  };
}

/**
 * zip.js. Reads from a Blob via the central directory, so a 2 GB deck full of
 * video does not have to sit in memory. Preferred for this tool.
 *
 *   const reader = new zip.ZipReader(new zip.BlobReader(file));
 *   const adapter = await zipjsAdapter(reader, zip);
 */
async function zipjsAdapter(zipReader, zipNS) {
  const entries = await zipReader.getEntries();
  const byName = new Map(entries.map((e) => [e.filename, e]));
  return {
    list: () => [...byName.keys()],
    sizeOf: (p) => byName.get(p)?.uncompressedSize ?? null,
    read: async (p) => new Uint8Array(await byName.get(p).getData(new zipNS.Uint8ArrayWriter())),
    readText: async (p) => byName.get(p).getData(new zipNS.TextWriter()),
  };
}

return {
  inspect: inspect, inspectFile: inspectFile, inspectBytes: inspectBytes,
  BlobReader: BlobReader, BytesReader: BytesReader,
  assess: assess, summarise: summarise, buildFixCommand: buildFixCommand,
  fmtBitrate: fmtBitrate, fmtFps: fmtFps, fmtDuration: fmtDuration, fmtBytes: fmtBytes,
  scanPptx: scanPptx, reportPptx: reportPptx,
  jszipAdapter: jszipAdapter, zipjsAdapter: zipjsAdapter,
  TARGETS: TARGETS, H264_LEVELS: H264_LEVELS, checkH264Level: checkH264Level
};
})();
