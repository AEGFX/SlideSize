/* key-inspect.js — Keynote Inspector engine
 *
 * Reads a Keynote document (single-file .key zip, or the package folder)
 * entirely in the browser and reports every media asset in Data/ with
 * its format, pixel size, bit depth, colour and the things that are
 * known to upset Keynote's renderer: 16-bit and float images, HDR
 * (gain map, PQ, HLG), HEIC, CMYK, oversized rasters, HDR video.
 *
 * Nothing is decoded. Every format is read from its header with random
 * access over the source file, so a multi-gigabyte deck is fine and the
 * only asset ever fully pulled into memory is one that the zip stored
 * deflated (Keynote stores media uncompressed, so that is rare).
 *
 * window.KI = {
 *   openSource(filesOrEntries)  -> Promise<source>   // .key File or [{path, file}]
 *   inspect(source, opts)       -> Promise<report>
 *   fmtBytes, fmtPixels
 * }
 *
 * Video analysis is delegated to media-inspect.js (window.MI) when it
 * is loaded; pass opts.MI.
 */
window.KI = (function () {
  'use strict';

  /* ------------------------------------------------------------------
     byte helpers
     ------------------------------------------------------------------ */
  const u8 = (b, o) => b[o];
  const u16be = (b, o) => (b[o] << 8) | b[o + 1];
  const u16le = (b, o) => b[o] | (b[o + 1] << 8);
  const u24le = (b, o) => b[o] | (b[o + 1] << 8) | (b[o + 2] << 16);
  const u32be = (b, o) => ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
  const u32le = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
  const i32le = (b, o) => b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24);
  const u64le = (b, o) => u32le(b, o) + u32le(b, o + 4) * 4294967296;
  const u64be = (b, o) => u32be(b, o) * 4294967296 + u32be(b, o + 4);
  const str = (b, o, n) => { let s = ''; for (let i = 0; i < n; i++) s += String.fromCharCode(b[o + i]); return s; };
  const utf8 = (b) => { try { return new TextDecoder().decode(b); } catch (e) { return str(b, 0, b.length); } };
  const hasBytes = (b, o, arr) => arr.every((v, i) => b[o + i] === v);

  function fmtBytes(n) {
    if (n == null) return '';
    if (n < 1024) return n + ' B';
    if (n < 1048576) return (n / 1024).toFixed(1) + ' KB';
    if (n < 1073741824) return (n / 1048576).toFixed(1) + ' MB';
    return (n / 1073741824).toFixed(2) + ' GB';
  }
  function fmtPixels(n) {
    if (!n) return '';
    if (n < 1e6) return (n / 1e3).toFixed(0) + ' K';
    return (n / 1e6).toFixed(1) + ' MP';
  }

  /* ------------------------------------------------------------------
     Random-access reader over a Blob slice
     ------------------------------------------------------------------ */
  function blobReader(blob, start, size) {
    start = start || 0;
    size = size == null ? blob.size - start : size;
    return {
      size,
      async read(off, len) {
        if (off >= size) return new Uint8Array(0);
        len = Math.min(len, size - off);
        const ab = await blob.slice(start + off, start + off + len).arrayBuffer();
        return new Uint8Array(ab);
      },
      slice(off, len) {
        len = len == null ? size - off : Math.min(len, size - off);
        return blob.slice(start + off, start + off + len);
      }
    };
  }
  function bytesReader(bytes) {
    return {
      size: bytes.length,
      async read(off, len) { return bytes.subarray(off, Math.min(bytes.length, off + len)); },
      slice(off, len) { len = len == null ? bytes.length - off : len; return new Blob([bytes.subarray(off, off + len)]); }
    };
  }

  /* ------------------------------------------------------------------
     Zip central directory reader (handles ZIP64)
     ------------------------------------------------------------------ */
  async function openZip(file) {
    const size = file.size;
    const tailLen = Math.min(size, 65557 + 20);
    const tail = new Uint8Array(await file.slice(size - tailLen, size).arrayBuffer());
    let eocd = -1;
    for (let i = tail.length - 22; i >= 0; i--) {
      if (tail[i] === 0x50 && tail[i + 1] === 0x4b && tail[i + 2] === 0x05 && tail[i + 3] === 0x06) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('Not a zip file (no end-of-central-directory record)');
    let count = u16le(tail, eocd + 10);
    let cdSize = u32le(tail, eocd + 12);
    let cdOffset = u32le(tail, eocd + 16);
    if (count === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff) {
      /* ZIP64 locator sits 20 bytes before the EOCD */
      const loc = eocd - 20;
      if (loc >= 0 && u32le(tail, loc) === 0x07064b50) {
        const z64Off = u64le(tail, loc + 8);
        const z = new Uint8Array(await file.slice(z64Off, z64Off + 56).arrayBuffer());
        if (u32le(z, 0) !== 0x06064b50) throw new Error('Bad ZIP64 record');
        count = u64le(z, 32);
        cdSize = u64le(z, 40);
        cdOffset = u64le(z, 48);
      }
    }
    const cd = new Uint8Array(await file.slice(cdOffset, cdOffset + cdSize).arrayBuffer());
    const entries = [];
    let p = 0;
    while (p + 46 <= cd.length && u32le(cd, p) === 0x02014b50) {
      const method = u16le(cd, p + 10);
      let csize = u32le(cd, p + 20);
      let usize = u32le(cd, p + 24);
      const nlen = u16le(cd, p + 28), elen = u16le(cd, p + 30), clen = u16le(cd, p + 32);
      let lho = u32le(cd, p + 42);
      const name = utf8(cd.subarray(p + 46, p + 46 + nlen));
      /* ZIP64 extra field */
      let ep = p + 46 + nlen; const eend = ep + elen;
      while (ep + 4 <= eend) {
        const id = u16le(cd, ep), sz = u16le(cd, ep + 2);
        if (id === 0x0001) {
          let q = ep + 4;
          if (usize === 0xffffffff) { usize = u64le(cd, q); q += 8; }
          if (csize === 0xffffffff) { csize = u64le(cd, q); q += 8; }
          if (lho === 0xffffffff) { lho = u64le(cd, q); q += 8; }
        }
        ep += 4 + sz;
      }
      if (!name.endsWith('/')) entries.push({ path: name, size: usize, csize, method, lho });
      p += 46 + nlen + elen + clen;
    }
    async function dataStart(e) {
      const lh = new Uint8Array(await file.slice(e.lho, e.lho + 30).arrayBuffer());
      if (u32le(lh, 0) !== 0x04034b50) throw new Error('Bad local header for ' + e.path);
      return e.lho + 30 + u16le(lh, 26) + u16le(lh, 28);
    }
    async function getReader(e) {
      const start = await dataStart(e);
      if (e.method === 0) return blobReader(file, start, e.csize);
      if (e.method === 8) {
        if (!window.DecompressionStream) throw new Error('Deflated entry and this browser cannot inflate it');
        if (e.size > 768 * 1048576) throw new Error('Deflated entry too large to inflate in memory');
        const ds = new DecompressionStream('deflate-raw');
        const resp = new Response(file.slice(start, start + e.csize).stream().pipeThrough(ds));
        return bytesReader(new Uint8Array(await resp.arrayBuffer()));
      }
      throw new Error('Unsupported zip compression method ' + e.method);
    }
    const media = entries.filter(e => /^(?:.*\/)?Data\//.test(e.path));
    return { kind: 'zip', name: file.name, size, entries, getReader, stored: media.every(e => e.method === 0) };
  }

  /* Folder source: [{ path, file }] with paths relative to the package root */
  function openFolder(list, name) {
    const entries = list.map(x => ({ path: x.path, size: x.file.size, csize: x.file.size, method: 0, file: x.file }));
    const total = entries.reduce((a, e) => a + e.size, 0);
    return {
      kind: 'folder', name: name || 'package', size: total, entries, stored: true,
      async getReader(e) { return blobReader(e.file, 0, e.file.size); }
    };
  }

  async function openSource(input) {
    if (input instanceof Blob) return openZip(input);
    if (Array.isArray(input)) return openFolder(input, input.rootName);
    throw new Error('Unsupported source');
  }

  /* ------------------------------------------------------------------
     Format sniffing
     ------------------------------------------------------------------ */
  function sniff(b, path) {
    const ext = (path.match(/\.([a-z0-9]+)$/i) || [, ''])[1].toLowerCase();
    if (b.length >= 8 && hasBytes(b, 0, [0x89, 0x50, 0x4e, 0x47])) return 'png';
    if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'jpeg';
    if (b.length >= 4 && (hasBytes(b, 0, [0x49, 0x49, 0x2a, 0x00]) || hasBytes(b, 0, [0x4d, 0x4d, 0x00, 0x2a]))) return 'tiff';
    if (b.length >= 4 && hasBytes(b, 0, [0x76, 0x2f, 0x31, 0x01])) return 'exr';
    if (b.length >= 4 && str(b, 0, 4) === '8BPS') return 'psd';
    if (b.length >= 6 && str(b, 0, 3) === 'GIF') return 'gif';
    if (b.length >= 12 && str(b, 0, 4) === 'RIFF' && str(b, 8, 4) === 'WEBP') return 'webp';
    if (b.length >= 2 && str(b, 0, 2) === 'BM') return 'bmp';
    if (b.length >= 5 && str(b, 0, 5) === '%PDF-') return 'pdf';
    if (b.length >= 12 && str(b, 4, 4) === 'ftyp') {
      const brand = str(b, 8, 4);
      if (/^(heic|heix|heim|heis|hevc|hevx|mif1|msf1|avif|avis)$/.test(brand)) return 'heif';
      return 'isobmff';
    }
    if (b.length >= 4 && (str(b, 0, 4) === 'FORM' || str(b, 0, 4) === 'RIFF' || str(b, 0, 4) === 'caff' || str(b, 0, 3) === 'ID3' || (b[0] === 0xff && (b[1] & 0xe0) === 0xe0))) return 'audio';
    const head = utf8(b.subarray(0, Math.min(b.length, 512))).trimStart();
    if (/^<\?xml|^<svg/i.test(head)) return 'svg';
    if (/^(mp4|m4v|mov|qt)$/.test(ext)) return 'isobmff';
    if (/^(mp3|m4a|aac|wav|aif|aiff|caf)$/.test(ext)) return 'audio';
    if (/^(ttf|otf|ttc)$/.test(ext)) return 'font';
    return 'other';
  }

  /* ------------------------------------------------------------------
     Image parsers. Each returns a partial asset record:
     { width, height, bits, channels, float, colour, hdr, alpha, extra:{...} }
     ------------------------------------------------------------------ */
  async function parsePNG(r) {
    const b = await r.read(0, Math.min(r.size, 1048576));
    const out = { format: 'PNG' };
    let p = 8;
    while (p + 8 <= b.length) {
      const len = u32be(b, p), type = str(b, p + 4, 4);
      if (type === 'IHDR') {
        out.width = u32be(b, p + 8); out.height = u32be(b, p + 12);
        out.bits = u8(b, p + 16);
        const ct = u8(b, p + 17);
        out.channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[ct] || 3;
        out.alpha = ct === 4 || ct === 6;
        out.colourType = ct;
        if (ct === 3) out.indexed = true;
        if (out.bits > 8) out.deep = true;
      } else if (type === 'cICP') {
        const tr = u8(b, p + 9);
        out.cicp = { primaries: u8(b, p + 8), transfer: tr };
        if (tr === 16 || tr === 18) out.hdr = tr === 16 ? 'PQ' : 'HLG';
      } else if (type === 'iCCP') {
        let n = ''; let q = p + 8; while (b[q] && q < p + 8 + 80) n += String.fromCharCode(b[q++]);
        out.icc = n || 'embedded';
      } else if (type === 'acTL') out.animated = true;
      else if (type === 'sRGB') out.srgb = true;
      else if (type === 'IDAT' || type === 'IEND') break;
      p += 12 + len;
    }
    return out;
  }

  async function parseJPEG(r) {
    const out = { format: 'JPEG', channels: 3, bits: 8 };
    let p = 2;
    let buf = await r.read(0, Math.min(r.size, 262144));
    let bufStart = 0;
    async function ensure(off, len) {
      if (off + len <= bufStart + buf.length && off >= bufStart) return;
      bufStart = off;
      buf = await r.read(off, Math.max(len, 65536));
    }
    const at = (o) => buf[o - bufStart];
    let guard = 0;
    while (p + 4 < r.size && guard++ < 200) {
      await ensure(p, 4);
      if (at(p) !== 0xff) { p++; continue; }
      const m = at(p + 1);
      if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) { p += 2; continue; }
      if (m === 0xff) { p++; continue; }
      const len = (at(p + 2) << 8) | at(p + 3);
      if (m === 0xda || m === 0xd9) break;
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) {
        await ensure(p, 10);
        out.bits = at(p + 4);
        out.height = (at(p + 5) << 8) | at(p + 6);
        out.width = (at(p + 7) << 8) | at(p + 8);
        out.channels = at(p + 9);
        out.progressive = m === 0xc2 || m === 0xc6 || m === 0xca || m === 0xce;
        if (m === 0xc3 || m === 0xc7 || m === 0xcb || m === 0xcf) out.lossless = true;
      } else if (m === 0xe1 || m === 0xe2 || m === 0xee) {
        const want = Math.min(len, 4096);
        await ensure(p, 4 + want);
        const seg = buf.subarray(p + 4 - bufStart, p + 4 - bufStart + want - 2);
        const tag = str(seg, 0, Math.min(seg.length, 29));
        if (m === 0xe1 && tag.startsWith('Exif')) out.exif = true;
        else if (m === 0xe1 && tag.startsWith('http://ns.adobe.com/xap/1.0/')) {
          const x = utf8(seg);
          out.xmp = true;
          if (/hdrgm:|HDRGainMap|hdrgainmap/i.test(x)) out.gainmap = 'xmp';
          if (/apple_desktop|iso21496/i.test(x)) out.gainmap = out.gainmap || 'xmp';
        } else if (m === 0xe2 && tag.startsWith('ICC_PROFILE')) {
          out.icc = out.icc || 'embedded';
          /* profile description is not trivial to pull, but the header has the colour space */
          const icc = seg.subarray(14);
          if (icc.length > 20) {
            const cs = str(icc, 16, 4);
            out.iccSpace = cs.trim();
            if (cs.trim() === 'CMYK') out.cmyk = true;
          }
        } else if (m === 0xe2 && tag.startsWith('MPF')) out.mpf = true;
        else if (m === 0xee && tag.startsWith('Adobe')) { out.adobe = true; out.adobeTransform = seg[11]; }
      }
      p += 2 + len;
    }
    if (out.channels === 4) out.cmyk = true;
    if (out.channels === 1) out.grey = true;
    if (out.mpf && out.gainmap) out.hdr = 'gain map';
    else if (out.gainmap) out.hdr = 'gain map (XMP)';
    else if (out.mpf) out.mpfOnly = true;
    if (out.bits && out.bits > 8) out.deep = true;
    return out;
  }

  async function parseTIFF(r) {
    const head = await r.read(0, 8);
    const le = head[0] === 0x49;
    const U16 = le ? u16le : u16be, U32 = le ? u32le : u32be;
    const out = { format: 'TIFF', channels: 3, bits: 8 };
    let ifdOff = U32(head, 4);
    let ifdCountBuf = await r.read(ifdOff, 2);
    const n = U16(ifdCountBuf, 0);
    const ifd = await r.read(ifdOff + 2, n * 12 + 4);
    const sizes = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 11: 4, 12: 8, 16: 8, 17: 8, 18: 8 };
    async function values(entry) {
      const type = U16(ifd, entry + 2), count = U32(ifd, entry + 4);
      const sz = (sizes[type] || 1) * count;
      let data;
      if (sz <= 4) data = ifd.subarray(entry + 8, entry + 12);
      else data = await r.read(U32(ifd, entry + 8), Math.min(sz, 4096));
      const vals = [];
      for (let i = 0; i < Math.min(count, 64); i++) {
        if (type === 3) vals.push(U16(data, i * 2));
        else if (type === 4) vals.push(U32(data, i * 4));
        else if (type === 1) vals.push(data[i]);
        else break;
      }
      return { type, count, vals };
    }
    for (let i = 0; i < n; i++) {
      const e = i * 12;
      const tag = U16(ifd, e);
      if (tag === 256) out.width = (await values(e)).vals[0];
      else if (tag === 257) out.height = (await values(e)).vals[0];
      else if (tag === 258) out.bits = (await values(e)).vals[0];
      else if (tag === 277) out.channels = (await values(e)).vals[0];
      else if (tag === 259) out.compression = (await values(e)).vals[0];
      else if (tag === 262) out.photometric = (await values(e)).vals[0];
      else if (tag === 339) out.sampleFormat = (await values(e)).vals[0];
      else if (tag === 338) out.alpha = true;
      else if (tag === 34377) out.photoshop = true;
      else if (tag === 37724) out.layers = true;   /* Photoshop ImageSourceData, layered TIFF */
      else if (tag === 34675) out.icc = 'embedded';
      else if (tag === 700) out.xmp = true;
    }
    if (out.sampleFormat === 3) out.float = true;
    if (out.photometric === 5) out.cmyk = true;
    if (out.photometric === 8) out.lab = true;
    if (out.photometric === 0 || out.photometric === 1) out.grey = true;
    if (out.bits > 8) out.deep = true;
    return out;
  }

  async function parseEXR(r) {
    const b = await r.read(0, Math.min(r.size, 65536));
    const out = { format: 'OpenEXR', float: true, hdr: 'scene-linear', bits: 16, channels: 0 };
    let p = 8;
    function cstr() { let s = ''; while (p < b.length && b[p]) s += String.fromCharCode(b[p++]); p++; return s; }
    while (p < b.length) {
      const name = cstr(); if (!name) break;
      const type = cstr();
      const size = u32le(b, p); p += 4;
      if (name === 'dataWindow') {
        const x0 = i32le(b, p), y0 = i32le(b, p + 4), x1 = i32le(b, p + 8), y1 = i32le(b, p + 12);
        out.width = x1 - x0 + 1; out.height = y1 - y0 + 1;
      } else if (name === 'channels') {
        let q = p; let maxBits = 16; let ch = 0;
        while (q < p + size) {
          let nm = ''; while (b[q]) nm += String.fromCharCode(b[q++]); q++;
          if (!nm) break;
          const pt = u32le(b, q); q += 16;
          ch++;
          if (pt === 2) maxBits = 32;
        }
        out.channels = ch; out.bits = maxBits;
      }
      p += size;
    }
    return out;
  }

  async function parsePSD(r) {
    const b = await r.read(0, 26);
    const ver = u16be(b, 4);
    const out = { format: ver === 2 ? 'PSB' : 'PSD' };
    out.channels = u16be(b, 12);
    out.height = u32be(b, 14); out.width = u32be(b, 18);
    out.bits = u16be(b, 22);
    const mode = u16be(b, 24);
    out.mode = { 0: 'Bitmap', 1: 'Greyscale', 2: 'Indexed', 3: 'RGB', 4: 'CMYK', 7: 'Multichannel', 8: 'Duotone', 9: 'Lab' }[mode] || ('mode ' + mode);
    if (mode === 4) out.cmyk = true;
    if (mode === 9) out.lab = true;
    if (out.bits > 8) out.deep = true;
    if (out.bits === 32) out.float = true;
    out.layered = true;
    return out;
  }

  async function parseGIF(r) {
    const b = await r.read(0, Math.min(r.size, 65536));
    const out = { format: 'GIF', bits: 8, channels: 3, indexed: true };
    out.width = u16le(b, 6); out.height = u16le(b, 8);
    const s = str(b, 0, b.length);
    if (s.indexOf('NETSCAPE2.0') >= 0) out.animated = true;
    return out;
  }

  async function parseWebP(r) {
    const b = await r.read(0, 64);
    const out = { format: 'WebP', bits: 8, channels: 3 };
    const chunk = str(b, 12, 4);
    if (chunk === 'VP8X') {
      const f = b[20];
      out.alpha = !!(f & 0x10); out.animated = !!(f & 0x02);
      out.width = 1 + u24le(b, 24); out.height = 1 + u24le(b, 27);
    } else if (chunk === 'VP8 ') {
      out.width = u16le(b, 26) & 0x3fff; out.height = u16le(b, 28) & 0x3fff;
    } else if (chunk === 'VP8L') {
      const bits = u32le(b, 21);
      out.width = (bits & 0x3fff) + 1; out.height = ((bits >> 14) & 0x3fff) + 1;
      out.alpha = !!((bits >> 28) & 1);
    }
    return out;
  }

  async function parseBMP(r) {
    const b = await r.read(0, 34);
    const out = { format: 'BMP', channels: 3 };
    out.width = Math.abs(i32le(b, 18)); out.height = Math.abs(i32le(b, 22));
    const bpp = u16le(b, 28);
    out.bits = bpp >= 24 ? 8 : bpp;
    out.channels = bpp === 32 ? 4 : 3;
    return out;
  }

  /* HEIF / HEIC via the meta box */
  function* boxes(b, start, end) {
    let p = start;
    while (p + 8 <= end) {
      let size = u32be(b, p); const type = str(b, p + 4, 4); let hdr = 8;
      if (size === 1) { size = u64be(b, p + 8); hdr = 16; }
      else if (size === 0) size = end - p;
      if (size < hdr) return;
      yield { type, start: p, body: p + hdr, end: Math.min(end, p + size) };
      p += size;
    }
  }
  async function parseHEIF(r) {
    const head = await r.read(0, Math.min(r.size, 32));
    const out = { format: 'HEIF', bits: 8, channels: 3 };
    const brand = str(head, 8, 4);
    out.brand = brand;
    if (/^avi/.test(brand)) out.format = 'AVIF';
    else out.format = 'HEIC';
    /* find meta box: walk top-level with small reads */
    let p = 0; let meta = null;
    while (p + 8 <= r.size) {
      const h = await r.read(p, 16);
      if (h.length < 8) break;
      let size = u32be(h, 0); const type = str(h, 4, 4); let hdr = 8;
      if (size === 1) { size = u64be(h, 8); hdr = 16; } else if (size === 0) size = r.size - p;
      if (type === 'meta') { meta = { start: p, size, hdr }; break; }
      if (type === 'mdat' && p > 64 * 1048576) break;
      p += size;
      if (size <= 0) break;
    }
    if (!meta) return out;
    const b = await r.read(meta.start, Math.min(meta.size, 8 * 1048576));
    const metaBody = meta.hdr + 4; /* fullbox */
    let primary = null; const items = {}; const props = []; const assoc = {};
    for (const bx of boxes(b, metaBody, b.length)) {
      if (bx.type === 'pitm') {
        const v = b[bx.body];
        primary = v === 0 ? u16be(b, bx.body + 4) : u32be(b, bx.body + 4);
      } else if (bx.type === 'iinf') {
        const v = b[bx.body];
        const cnt = v === 0 ? u16be(b, bx.body + 4) : u32be(b, bx.body + 4);
        const st = bx.body + (v === 0 ? 6 : 8);
        for (const inf of boxes(b, st, bx.end)) {
          if (inf.type !== 'infe') continue;
          const iv = b[inf.body];
          let id, type;
          if (iv >= 2) {
            id = iv === 2 ? u16be(b, inf.body + 4) : u32be(b, inf.body + 4);
            const o = inf.body + (iv === 2 ? 8 : 10);
            type = str(b, o, 4);
          }
          if (id != null) items[id] = { type };
        }
        void cnt;
      } else if (bx.type === 'iprp') {
        for (const ch of boxes(b, bx.body, bx.end)) {
          if (ch.type === 'ipco') {
            for (const pr of boxes(b, ch.body, ch.end)) {
              const o = { type: pr.type };
              if (pr.type === 'ispe') { o.width = u32be(b, pr.body + 4); o.height = u32be(b, pr.body + 8); }
              else if (pr.type === 'pixi') { const n = b[pr.body + 4]; o.bits = []; for (let i = 0; i < n; i++) o.bits.push(b[pr.body + 5 + i]); }
              else if (pr.type === 'colr') {
                const ct = str(b, pr.body, 4);
                if (ct === 'nclx') { o.primaries = u16be(b, pr.body + 4); o.transfer = u16be(b, pr.body + 6); o.matrix = u16be(b, pr.body + 8); o.fullRange = !!(b[pr.body + 10] & 0x80); }
                else o.icc = ct;
              }
              else if (pr.type === 'auxC') { let s = ''; let q = pr.body + 4; while (b[q] && q < pr.end) s += String.fromCharCode(b[q++]); o.urn = s; }
              else if (pr.type === 'hvcC') { o.bitDepthLuma = (b[pr.body + 21] & 0x07) + 8; }
              else if (pr.type === 'av1C') { const f = b[pr.body + 2]; o.highBitDepth = !!(f & 0x40); o.twelveBit = !!(f & 0x20); }
              props.push(o);
            }
          } else if (ch.type === 'ipma') {
            const v = b[ch.body], flags = u24le ? ((b[ch.body + 1] << 16) | (b[ch.body + 2] << 8) | b[ch.body + 3]) : 0;
            const cnt = u32be(b, ch.body + 4);
            let q = ch.body + 8;
            for (let i = 0; i < cnt && q < ch.end; i++) {
              const id = v < 1 ? u16be(b, q) : u32be(b, q); q += v < 1 ? 2 : 4;
              const ac = b[q++]; const list = [];
              for (let j = 0; j < ac; j++) {
                let idx;
                if (flags & 1) { idx = u16be(b, q) & 0x7fff; q += 2; } else { idx = b[q] & 0x7f; q += 1; }
                list.push(idx - 1);
              }
              assoc[id] = list;
            }
          }
        }
      }
    }
    const propsFor = (id) => (assoc[id] || []).map(i => props[i]).filter(Boolean);
    const prim = primary != null ? primary : Number(Object.keys(items)[0]);
    const pp = propsFor(prim);
    const ispe = pp.find(x => x.type === 'ispe');
    if (ispe) { out.width = ispe.width; out.height = ispe.height; }
    const pixi = pp.find(x => x.type === 'pixi');
    if (pixi && pixi.bits.length) { out.bits = Math.max.apply(null, pixi.bits); out.channels = pixi.bits.length; }
    const hvc = pp.find(x => x.type === 'hvcC');
    if (hvc && hvc.bitDepthLuma) out.bits = Math.max(out.bits, hvc.bitDepthLuma);
    const colr = pp.find(x => x.type === 'colr');
    if (colr) {
      if (colr.transfer === 16) out.hdr = 'PQ';
      else if (colr.transfer === 18) out.hdr = 'HLG';
      if (colr.primaries === 9) out.wideGamut = 'BT.2020';
      else if (colr.primaries === 12) out.wideGamut = 'Display P3';
      if (colr.icc) out.icc = colr.icc;
    }
    /* gain maps: ISO 21496 'tmap' item, or Apple's auxiliary HDR gain map */
    const types = Object.values(items).map(x => x.type);
    if (types.indexOf('tmap') >= 0) out.hdr = out.hdr || 'gain map (ISO)';
    for (const id in items) {
      const a = propsFor(Number(id)).find(x => x.type === 'auxC');
      if (a && /hdrgainmap/i.test(a.urn)) out.hdr = out.hdr || 'gain map (Apple)';
      if (a && /alpha/i.test(a.urn)) out.alpha = true;
    }
    if (items[prim] && items[prim].type === 'grid') out.tiled = true;
    if (out.bits > 8) out.deep = true;
    out.itemCount = Object.keys(items).length;
    return out;
  }

  async function parsePDF(r) {
    const out = { format: 'PDF', vector: true };
    const head = await r.read(0, 16);
    out.version = str(head, 5, 3);
    const scanLen = Math.min(r.size, 24 * 1048576);
    const b = await r.read(0, scanLen);
    const s = str(b, 0, b.length);
    const pages = (s.match(/\/Type\s*\/Page(?![s\w])/g) || []).length;
    out.pages = pages || null;
    out.scannedWhole = scanLen === r.size;
    if (/\/Subtype\s*\/Image/.test(s)) out.hasRaster = true;
    if (/\/BitsPerComponent\s*16/.test(s)) { out.deep = true; out.bits = 16; }
    if (/\/DeviceCMYK/.test(s)) out.cmyk = true;
    if (/\/SMask/.test(s)) out.alpha = true;
    if (/\/Encrypt/.test(s)) out.encrypted = true;
    return out;
  }

  /* ------------------------------------------------------------------
     Asset classification and rules
     ------------------------------------------------------------------ */
  const IMAGE_KINDS = ['png', 'jpeg', 'tiff', 'exr', 'psd', 'gif', 'webp', 'bmp', 'heif'];

  function estimateDecoded(a) {
    if (!a.width || !a.height) return 0;
    const ch = a.channels || 3;
    /* Keynote/CoreGraphics decode to RGBA. 16-bit and HDR paths decode to half float. */
    const bytes = (a.deep || a.float || a.hdr) ? 8 : 4;
    return a.width * a.height * bytes * (ch > 0 ? 1 : 1);
  }

  function assessImage(a) {
    const f = [];
    const add = (sev, code, title, detail) => f.push({ sev, code, title, detail });
    const mp = a.width && a.height ? a.width * a.height : 0;
    const side = Math.max(a.width || 0, a.height || 0);

    if (a.float) add('red', 'float', 'Floating-point image (' + a.format + (a.bits ? ', ' + a.bits + '-bit' : '') + ')',
      'Keynote renders this through its HDR path (RGBAf16). Convert to 8-bit sRGB PNG or JPEG.');
    else if (a.hdr) add('red', 'hdr', 'HDR image (' + a.hdr + ')',
      'Gain map or PQ/HLG content is decoded as 16-bit float and composited through Keynote\'s HDR headroom path. Export an SDR copy, 8-bit sRGB, no gain map.');
    else if (a.deep) add('red', 'deep', a.bits + '-bit image',
      'Anything over 8 bits per channel goes through the 16-bit float render path. Convert to 8-bit sRGB.');

    if (a.format === 'HEIC' && !a.hdr && !a.deep) add('amber', 'heic', 'HEIC image',
      'Stored in the deck as HEIC and decoded by the HEVC decoder at every render. Convert to JPEG or PNG to take the decoder out of the loop.');
    if (a.format === 'AVIF') add('amber', 'avif', 'AVIF image', 'Convert to JPEG or PNG for Keynote.');

    if (a.cmyk) add('amber', 'cmyk', 'CMYK colour', 'Colour-managed conversion on every draw and the colours will not match the designer\'s intent. Convert to RGB.');
    if (a.lab) add('amber', 'lab', 'Lab colour', 'Convert to RGB.');
    if (a.wideGamut && !a.hdr) add('info', 'gamut', a.wideGamut + ' primaries', 'Wide-gamut SDR. Fine on a Mac display, may shift on a projector.');
    if (a.icc && /^(CMYK)$/.test(a.iccSpace || '')) { /* already flagged */ }

    if (side > 8192) add('red', 'huge', 'Over 8192 px on a side (' + a.width + ' × ' + a.height + ')',
      'Past the safe texture size for most GPUs and heavier than any screen it will play on. Resample to the slide size at 2x at most.');
    else if (side > 4096 || mp > 16e6) add('amber', 'large', 'Large raster (' + a.width + ' × ' + a.height + ', ' + fmtPixels(mp) + ')',
      'Bigger than a 4K canvas. Resample to the slide size, or 2x the slide size for zoomed builds.');

    if (a.format === 'PSD' || a.format === 'PSB') add('amber', 'psd', 'Photoshop file',
      'Keynote keeps the whole layered file in the document and flattens it on the fly. Export a flattened PNG or JPEG.');
    if (a.layers) add('amber', 'layered-tiff', 'Layered TIFF', 'Contains Photoshop layer data. Flatten and export as PNG or JPEG.');
    if (a.format === 'TIFF' && !a.deep && !a.layers && !a.cmyk && !a.float) add('info', 'tiff', 'TIFF', 'Works, but PNG or JPEG is lighter and decodes faster.');
    if (a.animated) add('info', 'anim', 'Animated ' + a.format, 'Every frame is decoded and held. Consider a short MP4 instead.');
    if (a.lossless) add('amber', 'jpeg-lossless', 'Lossless JPEG', 'Not supported by every decoder. Re-save as a standard JPEG or PNG.');
    if (a.format === 'JPEG' && a.mpfOnly) add('info', 'mpf', 'Multi-picture JPEG', 'Carries a second embedded image (depth map or thumbnail). Harmless but adds weight. Re-save to strip it.');
    if (a.format === 'BMP') add('info', 'bmp', 'BMP', 'Uncompressed. Convert to PNG.');
    return f;
  }

  function assessPDF(a) {
    const f = [];
    const add = (sev, code, title, detail) => f.push({ sev, code, title, detail });
    if (a.encrypted) add('amber', 'pdf-enc', 'Encrypted PDF', 'Keynote may fail to render it. Re-export without security.');
    if (a.deep) add('red', 'pdf-16', 'PDF with 16-bit images', 'Embedded 16-bit images are rendered through the same float path. Flatten to an 8-bit PNG.');
    if (a.pages && a.pages > 1) add('amber', 'pdf-pages', a.pages + '-page PDF', 'Only the first page shows. The rest is dead weight the renderer still parses.');
    if (a.size > 10 * 1048576) add('amber', 'pdf-big', 'Large PDF (' + fmtBytes(a.size) + ')', 'Vector PDFs are re-rasterised at each zoom level and in every thumbnail. Flatten heavy artwork to PNG at 2x.');
    if (a.hasRaster && !a.deep) add('info', 'pdf-raster', 'PDF wraps a raster image', 'A JPEG or PNG inside a PDF. Export the image directly instead.');
    if (a.cmyk) add('amber', 'pdf-cmyk', 'CMYK PDF', 'Print-intent artwork. Re-export in RGB.');
    return f;
  }

  function assessVideo(a) {
    const f = [];
    const add = (sev, code, title, detail) => f.push({ sev, code, title, detail });
    const v = a.video;
    if (!v) return f;
    const codec = (v.codec && v.codec.string) || '';
    if (/^dv(h1|he|a1|av)/i.test(codec)) add('red', 'dolby-vision', 'Dolby Vision video', 'Composited through the HDR path. Re-encode as SDR H.264 or HEVC (BT.709).');
    else if (v.colour && (v.colour.transferIdc === 16 || v.colour.transferIdc === 18)) add('red', 'hdr-video', 'HDR video (' + v.colour.transfer + ')', 'Composited through the HDR headroom path. Tone-map to SDR BT.709 on export.');
    if (v.bitDepth && v.bitDepth > 8) add('amber', 'video-depth', v.bitDepth + '-bit video', '8-bit 4:2:0 is the safe target for Keynote playback.');
    const w = v.cropped ? v.cropped.width : (v.coded ? v.coded.width : 0);
    if (w > 3840) add('amber', 'video-size', 'Video wider than 4K (' + w + ' px)', 'Scale to the canvas size on export.');
    if (v.frameRate && v.frameRate.mode && v.frameRate.mode !== 'CFR') add('amber', 'vfr', 'Variable frame rate', 'Conform to a constant frame rate on export.');
    if (v.interlaced && v.interlaced.flag) add('amber', 'interlaced', 'Interlaced video', 'Deinterlace on export.');
    if (/^(ap4h|ap4x)$/i.test(codec)) add('info', 'prores-4444', 'ProRes 4444', 'Plays, but very high bitrate. ProRes 422 or H.264 unless the alpha is needed.');
    return f;
  }

  /* ------------------------------------------------------------------
     Main
     ------------------------------------------------------------------ */
  /* Header-only probe of a single file or blob. Used by Media Extractor for
     instant dimensions without decoding. Returns {kind, format, width, height, bits, ...} */
  async function probeReader(r, name) {
    const head = await r.read(0, Math.min(r.size, 64));
    const kind = sniff(head, name || '');
    let info = { kind, format: ((name || '').match(/\.([a-z0-9]+)$/i) || [, '?'])[1].toUpperCase() };
    let parsed = null;
    if (kind === 'png') parsed = await parsePNG(r);
    else if (kind === 'jpeg') parsed = await parseJPEG(r);
    else if (kind === 'tiff') parsed = await parseTIFF(r);
    else if (kind === 'exr') parsed = await parseEXR(r);
    else if (kind === 'psd') parsed = await parsePSD(r);
    else if (kind === 'gif') parsed = await parseGIF(r);
    else if (kind === 'webp') parsed = await parseWebP(r);
    else if (kind === 'bmp') parsed = await parseBMP(r);
    else if (kind === 'heif') parsed = await parseHEIF(r);
    else if (kind === 'pdf') parsed = await parsePDF(r);
    if (parsed) Object.assign(info, parsed);
    return info;
  }
  async function probeBlob(blob, name) {
    return probeReader(blobReader(blob, 0, blob.size), name || blob.name || '');
  }


  /* ------------------------------------------------------------------
     IWA (Index/*.iwa) reading: Snappy chunks of TSP.ArchiveInfo protobufs.
     Used to map each Data/ asset to the slide(s) that reference it.
     Type ids: 1 KN.DocumentArchive, 2 KN.ShowArchive, 4 KN.SlideNodeArchive,
     5 KN.SlideArchive, 6 KN.SlideArchive (master), 11006 TSP.PackageMetadata.
     ------------------------------------------------------------------ */
  function snappyUncompress(src) {
    let p = 0, len = 0, shift = 0, b;
    do { b = src[p++]; len |= (b & 0x7f) << shift; shift += 7; } while (b & 0x80);
    const out = new Uint8Array(len);
    let o = 0;
    while (p < src.length) {
      const tag = src[p++];
      const t = tag & 3;
      if (t === 0) {
        let n = tag >> 2;
        if (n >= 60) { const k = n - 59; n = 0; for (let i = 0; i < k; i++) n |= src[p++] << (8 * i); }
        n += 1;
        out.set(src.subarray(p, p + n), o); o += n; p += n;
      } else {
        let n, off;
        if (t === 1) { n = ((tag >> 2) & 7) + 4; off = ((tag >> 5) << 8) | src[p++]; }
        else if (t === 2) { n = (tag >> 2) + 1; off = src[p] | (src[p + 1] << 8); p += 2; }
        else { n = (tag >> 2) + 1; off = (src[p] | (src[p + 1] << 8) | (src[p + 2] << 16) | (src[p + 3] << 24)) >>> 0; p += 4; }
        if (off === 0 || off > o) throw new Error('bad snappy offset');
        for (let i = 0; i < n; i++) { out[o] = out[o - off]; o++; }
      }
    }
    return o === len ? out : out.subarray(0, o);
  }
  function iwaUncompress(bytes) {
    const parts = []; let p = 0, total = 0;
    while (p + 4 <= bytes.length) {
      if (bytes[p] !== 0) throw new Error('IWA chunk header');
      const len = bytes[p + 1] | (bytes[p + 2] << 8) | (bytes[p + 3] << 16);
      const chunk = bytes.subarray(p + 4, p + 4 + len);
      let out;
      try { out = snappyUncompress(chunk); } catch (e) { out = chunk; }
      parts.push(out); total += out.length;
      p += 4 + len;
    }
    const all = new Uint8Array(total); let o = 0;
    for (const x of parts) { all.set(x, o); o += x.length; }
    return all;
  }
  /* minimal protobuf wire reader */
  function pbVarint(b, p) {
    let v = 0, mul = 1, byte;
    do { byte = b[p++]; v += (byte & 0x7f) * mul; mul *= 128; } while (byte & 0x80);
    return [v, p];
  }
  function pbFields(b, start, end) {
    const out = []; let p = start == null ? 0 : start; end = end == null ? b.length : end;
    while (p < end) {
      let key; [key, p] = pbVarint(b, p);
      const field = Math.floor(key / 8), wire = key % 8;
      if (wire === 0) { let v; [v, p] = pbVarint(b, p); out.push({ field, wire, v }); }
      else if (wire === 1) { out.push({ field, wire, start: p, end: p + 8 }); p += 8; }
      else if (wire === 2) { let n; [n, p] = pbVarint(b, p); out.push({ field, wire, start: p, end: p + n }); p += n; }
      else if (wire === 5) { out.push({ field, wire, start: p, end: p + 4 }); p += 4; }
      else throw new Error('protobuf wire type ' + wire);
    }
    return out;
  }
  const pbStr = (b, f) => utf8(b.subarray(f.start, f.end));
  function pbPacked(b, f) {   /* repeated varint, packed or not */
    if (f.wire === 0) return [f.v];
    const out = []; let p = f.start;
    while (p < f.end) { let v; [v, p] = pbVarint(b, p); out.push(v); }
    return out;
  }
  function pbRefId(b, f) {    /* TSP.Reference / DataReference { identifier = 1 } */
    for (const x of pbFields(b, f.start, f.end)) if (x.field === 1 && x.wire === 0) return x.v;
    return null;
  }
  /* Walk one .iwa: yields { id, type, refs, dataRefs, body(bytes), bodyStart, bodyEnd } */
  function iwaArchives(bytes) {
    const b = iwaUncompress(bytes);
    const out = []; let p = 0;
    while (p < b.length) {
      let n; [n, p] = pbVarint(b, p);
      const infoEnd = p + n;
      let id = null; const msgs = [];
      for (const f of pbFields(b, p, infoEnd)) {
        if (f.field === 1 && f.wire === 0) id = f.v;
        else if (f.field === 2 && f.wire === 2) {
          const m = { type: 0, length: 0, refs: [], dataRefs: [] };
          for (const g of pbFields(b, f.start, f.end)) {
            if (g.field === 1) m.type = g.v;
            else if (g.field === 3) m.length = g.v;
            else if (g.field === 5) m.refs = m.refs.concat(pbPacked(b, g));
            else if (g.field === 6) m.dataRefs = m.dataRefs.concat(pbPacked(b, g));
          }
          msgs.push(m);
        }
      }
      p = infoEnd;
      msgs.forEach((m, i) => {
        const rec = { id, type: m.type, refs: m.refs, dataRefs: m.dataRefs, bodyStart: p, bodyEnd: p + m.length, b, primary: i === 0 };
        out.push(rec);
        p += m.length;
      });
    }
    return out;
  }

  /* Build { dataIdToFile, fileToSlides, slides[] } */
  async function buildSlideMap(source, entries) {
    const map = { ok: false, slides: [], fileSlides: {}, orphans: [] };
    const iwaEntries = entries.filter(e => /(^|\/)Index\/[^/]+\.iwa$/.test(e.path));
    if (!iwaEntries.length) return map;
    const objects = {};           /* id -> { type, comp, refs, dataRefs, rec } */
    const compDataRefs = {};      /* comp -> Set(dataId) */
    const dataFiles = {};         /* data id -> file name */
    let showTree = null;          /* ordered top-level slide node ids */
    const slideNodes = {};        /* node id -> { slideId, children[], skipped, thumbs[] } */
    const slideArchives = {};     /* slide id -> { name, comp, master:boolean } */

    for (const e of iwaEntries) {
      const comp = e.path.replace(/^.*\/Index\//, '').replace(/^Index\//, '').replace(/\.iwa$/, '');
      let bytes;
      try { const r = await source.getReader(e); bytes = await r.read(0, r.size); } catch (err) { continue; }
      let arcs;
      try { arcs = iwaArchives(bytes); } catch (err) { continue; }
      const set = compDataRefs[comp] || (compDataRefs[comp] = new Set());
      for (const a of arcs) {
        a.dataRefs.forEach(d => set.add(d));
        if (a.primary && a.id != null) objects[a.id] = { type: a.type, comp, refs: a.refs, dataRefs: a.dataRefs };
        const b = a.b;
        if (a.type === 11006) {
          for (const f of pbFields(b, a.bodyStart, a.bodyEnd)) {
            if (f.field === 4 && f.wire === 2) {
              let id = null, fn = null, pref = null;
              for (const g of pbFields(b, f.start, f.end)) {
                if (g.field === 1) id = g.v; else if (g.field === 3) pref = pbStr(b, g); else if (g.field === 4) fn = pbStr(b, g);
              }
              if (id != null) dataFiles[id] = fn || pref;
            }
          }
        } else if (a.type === 2) {
          for (const f of pbFields(b, a.bodyStart, a.bodyEnd)) {
            if (f.field === 3 && f.wire === 2) {
              showTree = [];
              for (const g of pbFields(b, f.start, f.end)) if (g.field === 2 && g.wire === 2) { const id = pbRefId(b, g); if (id != null) showTree.push(id); }
            }
          }
        } else if (a.type === 4 && a.id != null) {
          const node = { slideId: null, children: [], skipped: false, thumbs: [] };
          for (const f of pbFields(b, a.bodyStart, a.bodyEnd)) {
            if (f.field === 1 && f.wire === 2) { const id = pbRefId(b, f); if (id != null) node.children.push(id); }
            else if (f.field === 2 && f.wire === 2) node.slideId = pbRefId(b, f);
            else if (f.field === 4 && f.wire === 0) node.skipped = !!f.v;
            else if (f.field === 16 && f.wire === 2) { const id = pbRefId(b, f); if (id != null) node.thumbs.push(id); }
          }
          slideNodes[a.id] = node;
        } else if ((a.type === 5 || a.type === 6) && a.id != null) {
          let name = null;
          for (const f of pbFields(b, a.bodyStart, a.bodyEnd)) if (f.field === 10 && f.wire === 2) name = pbStr(b, f);
          slideArchives[a.id] = { name, comp, master: a.type === 6 || /^MasterSlide/.test(comp) };
        }
      }
    }
    if (!Object.keys(dataFiles).length) return map;
    map.ok = true;

    /* slide order: depth-first through the show tree, Keynote numbers every slide incl. skipped */
    const ordered = [];
    const seen = new Set();
    function walk(nodeId, depth) {
      if (seen.has(nodeId)) return; seen.add(nodeId);
      const node = slideNodes[nodeId]; if (!node) return;
      const sa = node.slideId != null ? slideArchives[node.slideId] : null;
      const comp = sa ? sa.comp : (objects[node.slideId] || {}).comp;
      ordered.push({ num: ordered.length + 1, nodeId, slideId: node.slideId, comp, name: sa && sa.name, skipped: node.skipped, depth, thumbs: node.thumbs });
      node.children.forEach(c => walk(c, depth + 1));
    }
    if (showTree && showTree.length) showTree.forEach(id => walk(id, 0));
    else {
      /* no tree found, fall back to Slide-N file order and say so */
      map.fileOrder = true;
      Object.keys(slideArchives).filter(id => !slideArchives[id].master)
        .sort((x, y) => parseInt((slideArchives[x].comp.match(/\d+/) || [0])[0]) - parseInt((slideArchives[y].comp.match(/\d+/) || [0])[0]))
        .forEach(id => ordered.push({ num: ordered.length + 1, slideId: Number(id), comp: slideArchives[id].comp, name: slideArchives[id].name, skipped: false, depth: 0, thumbs: [] }));
    }
    map.slides = ordered;

    /* component -> slide label(s); masters and the document itself get their own labels */
    const compLabel = {};
    ordered.forEach(sl => { if (sl.comp) (compLabel[sl.comp] = compLabel[sl.comp] || []).push(sl); });
    const masters = Object.keys(slideArchives).filter(id => slideArchives[id].master).map(id => slideArchives[id]);

    /* data id -> file -> slides */
    const fileSlides = map.fileSlides;
    const touch = (dataId, where) => {
      const fn = dataFiles[dataId]; if (!fn) return;
      const arr = fileSlides[fn] || (fileSlides[fn] = []);
      if (!arr.some(w => w.key === where.key)) arr.push(where);
    };
    for (const comp in compDataRefs) {
      const sls = compLabel[comp];
      if (sls) {
        for (const d of compDataRefs[comp]) sls.forEach(sl => touch(d, { key: 'slide:' + sl.num, kind: 'slide', num: sl.num, name: sl.name, skipped: sl.skipped }));
      } else if (/^MasterSlide/.test(comp)) {
        const m = masters.find(x => x.comp === comp);
        for (const d of compDataRefs[comp]) touch(d, { key: 'master:' + comp, kind: 'master', name: (m && m.name) || comp });
      } else {
        for (const d of compDataRefs[comp]) touch(d, { key: 'doc:' + comp, kind: 'document', name: comp });
      }
    }
    /* slide-node thumbnails live in Document.iwa but belong to a slide */
    ordered.forEach(sl => sl.thumbs.forEach(d => {
      const fn = dataFiles[d]; if (!fn) return;
      fileSlides[fn] = [{ key: 'thumb:' + sl.num, kind: 'thumbnail', num: sl.num, name: sl.name }];
    }));
    map.dataFiles = dataFiles;
    map.masterCount = masters.length;
    return map;
  }

  async function inspect(source, opts = {}) {
    const MI = opts.MI || null;
    const progress = opts.onProgress || function () {};
    const report = {
      name: source.name, kind: source.kind, size: source.size, stored: source.stored,
      slides: 0, masters: 0, versions: [], assets: [], findings: [], preview: null,
      totals: { images: 0, videos: 0, audio: 0, pdfs: 0, other: 0, mediaBytes: 0, decodedBytes: 0, deepCount: 0, hdrCount: 0 }
    };
    const entries = source.entries;
    const dataEntries = entries.filter(e => /^(?:.*\/)?Data\//.test(e.path));
    report.slides = entries.filter(e => /(^|\/)Index\/Slide-\d+\.iwa$/.test(e.path)).length;
    report.masters = entries.filter(e => /(^|\/)Index\/MasterSlide-\d+\.iwa$/.test(e.path)).length;
    report.iwaCount = entries.filter(e => /\.iwa$/.test(e.path)).length;
    report.isKeynote = entries.some(e => /(^|\/)Index\/Document\.iwa$/.test(e.path)) || entries.some(e => /(^|\/)index\.apxl$/.test(e.path));
    report.legacy = entries.some(e => /(^|\/)index\.apxl(\.gz)?$/.test(e.path));
    progress(0, 1, 'Reading slide index');
    let slideMap = { ok: false, slides: [], fileSlides: {} };
    try { slideMap = await buildSlideMap(source, entries); } catch (e) { report.slideMapError = e.message || String(e); }
    report.slideMap = slideMap.ok;
    report.slideOrderFromFiles = !!slideMap.fileOrder;
    if (slideMap.ok && slideMap.slides.length) { report.slides = slideMap.slides.length; report.slideList = slideMap.slides; }
    if (slideMap.ok && slideMap.masterCount) report.masters = slideMap.masterCount;

    /* Metadata/BuildVersionHistory.plist is XML, lists every Keynote build that touched the file */
    const bvh = entries.find(e => /(^|\/)Metadata\/BuildVersionHistory\.plist$/.test(e.path));
    if (bvh) {
      try {
        const r = await source.getReader(bvh);
        const txt = utf8(await r.read(0, Math.min(r.size, 65536)));
        report.versions = (txt.match(/<string>([^<]+)<\/string>/g) || []).map(s => s.replace(/<\/?string>/g, ''));
      } catch (e) { /* binary plist, skip */ }
    }
    const prev = entries.find(e => /(^|\/)preview\.jpg$/.test(e.path)) || entries.find(e => /(^|\/)preview-web\.jpg$/.test(e.path));
    if (prev) {
      try {
        const r = await source.getReader(prev);
        report.preview = r.slice(0);
        const pj = await parseJPEG(r);
        report.previewSize = { width: pj.width, height: pj.height };
      } catch (e) { /* ignore */ }
    }

    let done = 0;
    for (const e of dataEntries) {
      progress(done++, dataEntries.length, e.path);
      const name = e.path.replace(/^(?:.*\/)?Data\//, '');
      const a = { path: e.path, name, size: e.size, kind: 'other', format: (name.match(/\.([a-z0-9]+)$/i) || [, '?'])[1].toUpperCase(), findings: [] };
      try {
        const r = await source.getReader(e);
        const head = await r.read(0, Math.min(r.size, 64));
        const kind = sniff(head, name);
        a.kind = kind;
        let info = {};
        if (kind === 'png') info = await parsePNG(r);
        else if (kind === 'jpeg') info = await parseJPEG(r);
        else if (kind === 'tiff') info = await parseTIFF(r);
        else if (kind === 'exr') info = await parseEXR(r);
        else if (kind === 'psd') info = await parsePSD(r);
        else if (kind === 'gif') info = await parseGIF(r);
        else if (kind === 'webp') info = await parseWebP(r);
        else if (kind === 'bmp') info = await parseBMP(r);
        else if (kind === 'heif') info = await parseHEIF(r);
        else if (kind === 'pdf') info = await parsePDF(r);
        else if (kind === 'isobmff') {
          a.format = 'Video';
          if (MI) {
            const blob = r.slice(0);
            blob.name = name;
            try {
              const vi = await MI.inspectFile(new File([blob], name));
              if (vi.ok) {
                a.video = vi.video[0] || null;
                a.audioTrack = vi.audio[0] || null;
                a.duration = vi.duration;
                a.container = vi.container;
                if (a.video) {
                  a.width = a.video.cropped ? a.video.cropped.width : a.video.coded.width;
                  a.height = a.video.cropped ? a.video.cropped.height : a.video.coded.height;
                  a.bits = a.video.bitDepth || 8;
                  a.codec = a.video.codec.name + (a.video.profile ? ' ' + a.video.profile : '');
                  a.colour = a.video.colour ? a.video.colour.transfer : 'untagged';
                  a.fps = a.video.frameRate && (a.video.frameRate.nominalFps || a.video.frameRate.fps);
                }
              } else a.error = vi.error;
            } catch (err) { a.error = err.message || String(err); }
          }
        } else if (kind === 'audio') a.format = 'Audio';
        else if (kind === 'svg') { a.format = 'SVG'; a.vector = true; }
        else if (kind === 'font') a.format = 'Font';
        Object.assign(a, info);
      } catch (err) {
        a.error = err.message || String(err);
      }

      if (IMAGE_KINDS.indexOf(a.kind) >= 0) {
        report.totals.images++;
        a.findings = assessImage(a);
        a.decoded = estimateDecoded(a);
        report.totals.decodedBytes += a.decoded;
        if ((a.deep || a.float) && !a.hdr) report.totals.deepCount++;
        if (a.hdr) report.totals.hdrCount++;
      } else if (a.kind === 'pdf') { report.totals.pdfs++; a.findings = assessPDF(a); }
      else if (a.kind === 'isobmff') { report.totals.videos++; a.findings = assessVideo(a); if (a.hdr) report.totals.hdrCount++; }
      else if (a.kind === 'audio') report.totals.audio++;
      else report.totals.other++;
      report.totals.mediaBytes += a.size;
      a.worst = a.findings.reduce((w, f) => rank(f.sev) > rank(w) ? f.sev : w, 'green');
      if (slideMap.ok) {
        a.where = slideMap.fileSlides[a.name] || [];
        a.orphan = !a.where.length;
        if (a.where.some(w => w.kind === 'thumbnail')) a.thumbnail = true;
      }
      report.assets.push(a);
    }
    progress(dataEntries.length, dataEntries.length, '');

    /* deck-level findings */
    const t = report.totals;
    const add = (sev, code, title, detail) => report.findings.push({ sev, code, title, detail });
    if (!report.isKeynote) add('amber', 'not-keynote', 'This does not look like a Keynote document', 'No Index/Document.iwa found. The media scan still ran on anything under Data/.');
    if (report.legacy) add('amber', 'legacy', 'Keynote \'09 format', 'Old XML-based document. Open in Keynote and Save As to upgrade before anything else.');
    if (t.hdrCount) add('red', 'deck-hdr', t.hdrCount + ' HDR asset' + (t.hdrCount > 1 ? 's' : ''), 'Any HDR content puts the whole slide through Keynote\'s float compositing path. Replace every one.');
    if (t.deepCount) add('red', 'deck-deep', t.deepCount + ' image' + (t.deepCount > 1 ? 's' : '') + ' over 8 bits per channel', 'Convert each to 8-bit sRGB. Listed below.');
    if (t.decodedBytes > 3 * 1073741824) add('red', 'deck-memory', 'About ' + fmtBytes(t.decodedBytes) + ' of decoded image data', 'If Keynote decodes every image (light table, export, thumbnails) this is what it holds in RAM. Resample the big ones.');
    else if (t.decodedBytes > 1.5 * 1073741824) add('amber', 'deck-memory', 'About ' + fmtBytes(t.decodedBytes) + ' of decoded image data', 'Heavy. Resample images to the slide size.');
    if (t.images > 400) add('amber', 'deck-count', t.images + ' images in the deck', 'Keynote builds a layer for every one. Consider splitting the deck.');
    if (source.kind === 'zip' && !source.stored) add('info', 'deflated', 'Some media is deflate-compressed inside the file', 'Unusual for Keynote. The file may have been re-zipped by hand.');
    const cmykN = report.assets.filter(a => a.cmyk).length;
    if (cmykN) add('amber', 'deck-cmyk', cmykN + ' CMYK asset' + (cmykN > 1 ? 's' : ''), 'Print artwork pasted in. Convert to RGB.');

    /* which slides need attention */
    if (slideMap.ok) {
      const bySlide = {};
      report.assets.forEach(a => {
        if (!a.findings.some(f => f.sev === 'red' || f.sev === 'amber')) return;
        (a.where || []).forEach(w => {
          const k = w.key;
          const o = bySlide[k] || (bySlide[k] = { key: k, kind: w.kind, num: w.num, name: w.name, skipped: w.skipped, assets: [], worst: 'green' });
          o.assets.push(a.name);
          if (rank(a.worst) > rank(o.worst)) o.worst = a.worst;
        });
      });
      report.slidesToFix = Object.values(bySlide).sort((x, y) => rank(y.worst) - rank(x.worst) || (x.num || 9e9) - (y.num || 9e9));
      const orphanBytes = report.assets.filter(a => a.orphan && !a.thumbnail).reduce((s, a) => s + a.size, 0);
      const orphanN = report.assets.filter(a => a.orphan && !a.thumbnail).length;
      if (orphanN) add('info', 'orphans', orphanN + ' unreferenced file' + (orphanN > 1 ? 's' : '') + ' in Data/ (' + fmtBytes(orphanBytes) + ')', 'Not used by any slide or master. Keynote usually clears these on File > Reduce File Size, or on a Save As.');
      if (slideMap.fileOrder) add('info', 'slide-order', 'Slide numbers are from file order', 'The slide tree could not be read, so Slide-N file numbering is used. Usually right, not guaranteed.');
    } else if (report.isKeynote && !report.legacy) {
      add('info', 'no-slide-map', 'Could not map assets to slides', (report.slideMapError ? report.slideMapError + '. ' : '') + 'The findings below are per file only.');
    }

    const worstAsset = report.assets.reduce((w, a) => rank(a.worst) > rank(w) ? a.worst : w, 'green');
    const worstDeck = report.findings.reduce((w, f) => rank(f.sev) > rank(w) ? f.sev : w, 'green');
    report.verdict = rank(worstAsset) > rank(worstDeck) ? worstAsset : worstDeck;
    if (report.verdict === 'info') report.verdict = 'green';
    return report;
  }

  function rank(s) { return { green: 0, info: 1, amber: 2, red: 3 }[s] || 0; }

  /* ------------------------------------------------------------------
     Plain-text report for copy/paste to whoever is onsite
     ------------------------------------------------------------------ */
  function textReport(rep) {
    const L = [];
    L.push('KEYNOTE INSPECTOR  ·  ' + rep.name);
    L.push('Verdict ' + rep.verdict.toUpperCase() + '  ·  ' + (rep.slides ? rep.slides + ' slides  ·  ' : '') + rep.totals.images + ' images  ·  ' + rep.totals.videos + ' videos  ·  ' + fmtBytes(rep.size));
    if (rep.versions.length) L.push('Keynote builds ' + rep.versions.join(', '));
    L.push('');
    if (rep.findings.length) {
      L.push('DECK');
      rep.findings.forEach(f => L.push('  [' + f.sev.toUpperCase() + '] ' + f.title + '  ' + f.detail));
      L.push('');
    }
    if (rep.slidesToFix && rep.slidesToFix.length) {
      L.push('SLIDES TO CHECK');
      rep.slidesToFix.forEach(sl => {
        const label = sl.kind === 'slide' ? 'Slide ' + sl.num + (sl.name ? '  "' + sl.name + '"' : '') + (sl.skipped ? '  (skipped)' : '') : sl.kind === 'master' ? 'Master "' + sl.name + '"' : sl.name;
        L.push('  [' + sl.worst.toUpperCase() + '] ' + label + '  ' + sl.assets.join(', '));
      });
      L.push('');
    }
    const flagged = rep.assets.filter(a => a.findings.length).sort((a, b) => rank(b.worst) - rank(a.worst));
    if (flagged.length) {
      L.push('ASSETS TO FIX');
      flagged.forEach(a => {
        L.push('  ' + a.name + '  ' + describe(a) + (a.where ? '  [' + whereLabel(a) + ']' : ''));
        a.findings.forEach(f => L.push('     [' + f.sev.toUpperCase() + '] ' + f.title + '  ' + f.detail));
      });
      L.push('');
    }
    L.push('ALL ASSETS');
    rep.assets.forEach(a => L.push('  ' + a.name + '  ' + describe(a) + (a.where ? '  [' + whereLabel(a) + ']' : '')));
    return L.join('\n');
  }
  function whereLabel(a) {
    if (!a.where) return '';
    if (!a.where.length) return a.thumbnail ? 'navigator thumbnail' : 'not used on any slide';
    const slides = a.where.filter(w => w.kind === 'slide').map(w => w.num).sort((x, y) => x - y);
    const parts = [];
    if (slides.length) parts.push((slides.length === 1 ? 'Slide ' : 'Slides ') + slides.join(', '));
    a.where.filter(w => w.kind === 'master').forEach(w => parts.push('Master "' + w.name + '"'));
    a.where.filter(w => w.kind === 'document').forEach(w => parts.push(w.name === 'Document' ? 'Document' : w.name));
    if (a.where.some(w => w.kind === 'thumbnail')) parts.push('navigator thumbnail');
    return parts.join(' · ');
  }
  function describe(a) {
    const parts = [a.format];
    if (a.width && a.height) parts.push(a.width + 'x' + a.height);
    if (a.bits && a.kind !== 'pdf') parts.push(a.bits + '-bit' + (a.float ? ' float' : ''));
    if (a.hdr) parts.push('HDR ' + a.hdr);
    if (a.cmyk) parts.push('CMYK');
    if (a.codec) parts.push(a.codec);
    if (a.fps) parts.push(Number(a.fps).toFixed(2).replace(/\.?0+$/, '') + ' fps');
    parts.push(fmtBytes(a.size));
    return parts.join(' · ');
  }

  return { openSource, inspect, probeBlob, probeReader, textReport, describe, whereLabel, fmtBytes, fmtPixels, rank };
})();
