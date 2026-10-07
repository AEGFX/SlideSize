/* ============================================================
   PowerPoint to PDF: deck inspection  (PptPdfInspect)

   Reads what a presentation contains before PowerPoint is asked
   to convert it. Slide count, slide size, hidden slides, fonts,
   video and audio, links to outside files, animations and the
   things a static PDF cannot show.

   A .pptx is a zip. Only the central directory and the XML parts
   are read, by slicing the file, so a 2 GB deck costs a few
   megabytes of memory and the media inside it is never loaded.

   A legacy .ppt is an OLE compound file. Only the slide size,
   the slide count and whether it is encrypted are read. The rest
   is reported by PowerPoint when it opens the file.

   No DOM access. The same file runs in the page and in Node.

   AEGFX / SlideSize
   ============================================================ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PptPdfInspect = factory();
})(typeof self !== 'undefined' ? self : this, function () {
'use strict';

var EMU_PER_PT = 12700;
var MAX_XML_BYTES = 64 * 1024 * 1024;      /* one XML part larger than this is not parsed */
var MAX_PPT_STREAM = 96 * 1024 * 1024;     /* legacy document stream larger than this is not walked */

/* ------------------------------------------------------------------ *
 * Zip reader over a Blob                                              *
 * ------------------------------------------------------------------ */

function u16(dv, o) { return dv.getUint16(o, true); }
function u32(dv, o) { return dv.getUint32(o, true); }
function u64(dv, o) { return u32(dv, o) + u32(dv, o + 4) * 4294967296; }

async function sliceBytes(blob, start, end) {
  return new Uint8Array(await blob.slice(start, end).arrayBuffer());
}

function decodeText(bytes) {
  var s = new TextDecoder('utf-8').decode(bytes);
  return s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s;
}

async function inflateRaw(bytes) {
  var ds = new DecompressionStream('deflate-raw');
  var stream = new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/* Returns { names, has(name), entry(name), bytes(name), text(name) }.
   Throws an Error with .code = 'not-zip' or 'zip-damaged'. */
async function openZip(blob) {
  var size = blob.size;
  if (size < 22) throw coded('not-zip', 'The file is too small to be a presentation.');
  var tailLen = Math.min(size, 65557);
  var tail = await sliceBytes(blob, size - tailLen, size);
  var dv = new DataView(tail.buffer, tail.byteOffset, tail.byteLength);
  var e = -1, i;
  for (i = tail.length - 22; i >= 0; i--) {
    if (tail[i] === 0x50 && tail[i + 1] === 0x4b && tail[i + 2] === 0x05 && tail[i + 3] === 0x06) { e = i; break; }
  }
  if (e < 0) throw coded('not-zip', 'No zip directory was found. The file is not a .pptx or it is cut short.');
  var total = u16(dv, e + 10), cdSize = u32(dv, e + 12), cdOffset = u32(dv, e + 16);
  if (total === 0xFFFF || cdSize === 0xFFFFFFFF || cdOffset === 0xFFFFFFFF) {
    /* zip64: the locator sits 20 bytes before the end record */
    var l = e - 20;
    if (l < 0 || u32(dv, l) !== 0x07064b50) throw coded('zip-damaged', 'The zip64 directory locator is missing.');
    var z64 = u64(dv, l + 8);
    var zr = await sliceBytes(blob, z64, z64 + 56);
    var zv = new DataView(zr.buffer, zr.byteOffset, zr.byteLength);
    if (zr.length < 56 || u32(zv, 0) !== 0x06064b50) throw coded('zip-damaged', 'The zip64 directory is damaged.');
    total = u64(zv, 32); cdSize = u64(zv, 40); cdOffset = u64(zv, 48);
  }
  if (cdOffset + cdSize > size) throw coded('zip-damaged', 'The zip directory points past the end of the file. The file is cut short.');
  var cd = await sliceBytes(blob, cdOffset, cdOffset + cdSize);
  var cv = new DataView(cd.buffer, cd.byteOffset, cd.byteLength);
  var entries = Object.create(null), names = [], p = 0, td = new TextDecoder('utf-8');
  for (i = 0; i < total; i++) {
    if (p + 46 > cd.length || u32(cv, p) !== 0x02014b50) throw coded('zip-damaged', 'The zip directory is damaged.');
    var method = u16(cv, p + 10), crc = u32(cv, p + 16);
    var csize = u32(cv, p + 20), usize = u32(cv, p + 24);
    var nlen = u16(cv, p + 28), xlen = u16(cv, p + 30), clen = u16(cv, p + 32);
    var lho = u32(cv, p + 42), flags = u16(cv, p + 8);
    var name = td.decode(cd.subarray(p + 46, p + 46 + nlen));
    if (usize === 0xFFFFFFFF || csize === 0xFFFFFFFF || lho === 0xFFFFFFFF) {
      var x = p + 46 + nlen, xe = x + xlen;
      while (x + 4 <= xe) {
        var id = u16(cv, x), len = u16(cv, x + 2), q = x + 4;
        if (id === 0x0001) {
          if (usize === 0xFFFFFFFF) { usize = u64(cv, q); q += 8; }
          if (csize === 0xFFFFFFFF) { csize = u64(cv, q); q += 8; }
          if (lho === 0xFFFFFFFF) { lho = u64(cv, q); q += 8; }
        }
        x += 4 + len;
      }
    }
    entries[name] = { name: name, method: method, csize: csize, usize: usize, lho: lho, crc: crc, flags: flags };
    names.push(name);
    p += 46 + nlen + xlen + clen;
  }

  async function bytes(name, cap) {
    var en = entries[name];
    if (!en) throw coded('zip-missing', 'Part ' + name + ' is missing from the file.');
    if (cap && en.usize > cap) throw coded('zip-too-large', 'Part ' + name + ' is ' + en.usize + ' bytes, larger than this check reads.');
    if (en.flags & 1) throw coded('zip-encrypted', 'Part ' + name + ' is encrypted.');
    var head = await sliceBytes(blob, en.lho, en.lho + 30);
    var hv = new DataView(head.buffer, head.byteOffset, head.byteLength);
    if (head.length < 30 || u32(hv, 0) !== 0x04034b50) throw coded('zip-damaged', 'Part ' + name + ' has a damaged header.');
    var start = en.lho + 30 + u16(hv, 26) + u16(hv, 28);
    var raw = await sliceBytes(blob, start, start + en.csize);
    if (raw.length !== en.csize) throw coded('zip-damaged', 'Part ' + name + ' is cut short.');
    if (en.method === 0) return raw;
    if (en.method === 8) {
      try { return await inflateRaw(raw); }
      catch (err) { throw coded('zip-damaged', 'Part ' + name + ' could not be decompressed.'); }
    }
    throw coded('zip-damaged', 'Part ' + name + ' uses a compression method this check does not read.');
  }

  return {
    names: names,
    has: function (n) { return !!entries[n]; },
    entry: function (n) { return entries[n] || null; },
    bytes: bytes,
    text: async function (n) { return decodeText(await bytes(n, MAX_XML_BYTES)); }
  };
}

function coded(code, message) { var e = new Error(message); e.code = code; return e; }

/* ------------------------------------------------------------------ *
 * A small tag scanner. OOXML is machine written and regular, so this  *
 * walks start, end and empty tags without building a tree.            *
 * ------------------------------------------------------------------ */

var TAG_RE = /<(\/?)([A-Za-z_][\w:.\-]*)((?:\s+[\w:.\-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;

/* on(kind, name, attrs, stack) with kind 'open' | 'close'. An empty tag
   fires open then close. stack holds the names of the open ancestors. */
function scan(xml, on) {
  var m, stack = [];
  TAG_RE.lastIndex = 0;
  while ((m = TAG_RE.exec(xml))) {
    var name = m[2];
    if (m[1]) {
      for (var k = stack.length - 1; k >= 0; k--) {
        if (stack[k] === name) { stack.length = k; break; }
      }
      on('close', name, '', stack);
    } else {
      on('open', name, m[3], stack);
      if (m[4]) on('close', name, '', stack);
      else stack.push(name);
    }
  }
}

function attr(attrs, name) {
  if (!attrs) return null;
  var re = new RegExp('(?:^|\\s)' + name.replace(/[.:\-]/g, '\\$&') + '\\s*=\\s*(?:"([^"]*)"|\'([^\']*)\')');
  var m = re.exec(attrs);
  if (!m) return null;
  return unescapeXml(m[1] !== undefined ? m[1] : m[2]);
}

function unescapeXml(s) {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, function (all, e) {
    if (e === 'amp') return '&'; if (e === 'lt') return '<'; if (e === 'gt') return '>';
    if (e === 'quot') return '"'; if (e === 'apos') return "'";
    var n = e.charAt(1) === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
    try { return String.fromCodePoint(n); } catch (x) { return all; }
  });
}

/* ------------------------------------------------------------------ *
 * Package helpers                                                     *
 * ------------------------------------------------------------------ */

function dirOf(part) { var i = part.lastIndexOf('/'); return i < 0 ? '' : part.slice(0, i); }

function resolveTarget(basePart, target) {
  if (target.charAt(0) === '/') return target.slice(1);
  var segs = dirOf(basePart).split('/').filter(Boolean);
  target.split('/').forEach(function (s) {
    if (s === '..') segs.pop(); else if (s !== '.' && s !== '') segs.push(s);
  });
  return segs.join('/');
}

function relsPartOf(part) {
  var d = dirOf(part), f = part.slice(d ? d.length + 1 : 0);
  return (d ? d + '/' : '') + '_rels/' + f + '.rels';
}

/* { byId: { rId: {type, target, external, part} }, list: [...] } */
async function readRels(zip, part) {
  var rp = relsPartOf(part), out = { byId: Object.create(null), list: [] };
  if (!zip.has(rp)) return out;
  var xml = await zip.text(rp);
  scan(xml, function (kind, name, attrs) {
    if (kind !== 'open' || name.replace(/^.*:/, '') !== 'Relationship') return;
    var type = attr(attrs, 'Type') || '', target = attr(attrs, 'Target') || '';
    var external = attr(attrs, 'TargetMode') === 'External';
    var r = { id: attr(attrs, 'Id'), type: type, kind: type.slice(type.lastIndexOf('/') + 1),
              target: target, external: external, part: external ? null : resolveTarget(part, target) };
    out.byId[r.id] = r; out.list.push(r);
  });
  return out;
}

function isRemote(target) {
  return /^(https?|ftp|ftps|sftp|smb|webdav|dav):/i.test(target) || /^\\\\/.test(target) || /^\/\/[^/]/.test(target) ||
         /^file:\/\/[^/]/i.test(target);
}

/* ------------------------------------------------------------------ *
 * Fonts                                                               *
 * ------------------------------------------------------------------ */

var FONT_TAGS = { 'a:latin': 1, 'a:ea': 1, 'a:cs': 1, 'a:sym': 1, 'a:buFont': 1 };

function collectFonts(xml, into) {
  var re = /<a:(?:latin|ea|cs|sym|buFont)\b[^>]*?\btypeface\s*=\s*"([^"]*)"/g, m;
  while ((m = re.exec(xml))) { var t = unescapeXml(m[1]).trim(); if (t) into[t] = true; }
}

function hasText(xml) { return /<a:t>[^<]*[^\s<][^<]*<\/a:t>/.test(xml); }

async function readTheme(zip, part) {
  var out = { major: null, minor: null };
  if (!part || !zip.has(part)) return out;
  var xml = await zip.text(part), where = null;
  scan(xml, function (kind, name, attrs, stack) {
    if (kind !== 'open') return;
    if (name === 'a:majorFont') where = 'major';
    else if (name === 'a:minorFont') where = 'minor';
    else if (name === 'a:latin' && where && stack[stack.length - 1] === (where === 'major' ? 'a:majorFont' : 'a:minorFont')) {
      var t = attr(attrs, 'typeface'); if (t && !out[where]) out[where] = t;
    }
  });
  return out;
}

function resolveFontNames(set, theme) {
  var out = Object.create(null);
  Object.keys(set).forEach(function (n) {
    if (n.charAt(0) === '+') {
      var t = /^\+mj-/.test(n) ? theme.major : /^\+mn-/.test(n) ? theme.minor : null;
      if (t && /-lt$/.test(n)) out[t] = true;      /* east asian and complex script theme slots are usually empty */
    } else out[n] = true;
  });
  return out;
}

/* ------------------------------------------------------------------ *
 * One slide                                                           *
 * ------------------------------------------------------------------ */

var SHAPE_TAGS = { 'p:sp': 1, 'p:pic': 1, 'p:graphicFrame': 1, 'p:cxnSp': 1, 'p:grpSp': 1, 'p:contentPart': 1 };

function emuBox(b) {
  if (!b || b.x == null || b.cx == null) return null;
  return { left: b.x / EMU_PER_PT, top: b.y / EMU_PER_PT, width: b.cx / EMU_PER_PT, height: b.cy / EMU_PER_PT };
}

function overlapShare(a, b) {
  var x = Math.max(0, Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left));
  var y = Math.max(0, Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top));
  var small = Math.min(a.width * a.height, b.width * b.height);
  return small > 0 ? (x * y) / small : 0;
}

function parseSlide(xml, rels) {
  var s = {
    hidden: false, media: [], linkedImages: [], ole: 0, controls: 0, model3d: 0, ink: 0, hyperlinks: 0,
    slideNumberField: /<a:fld\b[^>]*\btype\s*=\s*"slidenum"/.test(xml),
    transition: null, anim: { entrance: 0, exit: 0, emphasis: 0, path: 0, total: 0, overlapping: 0, exitTargets: 0 },
    fonts: Object.create(null), hasText: hasText(xml)
  };
  collectFonts(xml, s.fonts);

  var shapes = [], shapeStack = [], byId = Object.create(null);
  var effects = [], animTargets = { entr: Object.create(null), exit: Object.create(null) };
  var inTransition = 0, groupDepth = 0;

  scan(xml, function (kind, name, attrs, stack) {
    if (kind === 'open') {
      if (name === 'p:sld') { if (attr(attrs, 'show') === '0') s.hidden = true; return; }

      if (SHAPE_TAGS[name]) {
        var sh = { tag: name, id: null, name: '', box: null, inGroup: groupDepth > 0, video: null, audio: null, blip: null, blipLink: null };
        shapes.push(sh); shapeStack.push(sh);
        if (name === 'p:grpSp') groupDepth++;
        return;
      }
      var cur = shapeStack[shapeStack.length - 1];

      if (name === 'p:cNvPr' && cur && cur.id === null) {
        cur.id = attr(attrs, 'id'); cur.name = attr(attrs, 'name') || '';
        if (cur.id !== null) byId[cur.id] = cur;
      } else if ((name === 'a:off' || name === 'a:ext') && cur) {
        var parent = stack[stack.length - 1];
        if (parent === 'a:xfrm' || parent === 'p:xfrm') {
          cur.box = cur.box || {};
          if (name === 'a:off' && cur.box.x == null) { cur.box.x = +attr(attrs, 'x'); cur.box.y = +attr(attrs, 'y'); }
          if (name === 'a:ext' && cur.box.cx == null) { cur.box.cx = +attr(attrs, 'cx'); cur.box.cy = +attr(attrs, 'cy'); }
        }
      } else if (name === 'a:videoFile' && cur) {
        cur.video = cur.video || {}; cur.video.link = attr(attrs, 'r:link');
      } else if ((name === 'a:audioFile' || name === 'a:wavAudioFile' || name === 'a:audioCd') && cur) {
        cur.audio = cur.audio || {}; cur.audio.link = attr(attrs, 'r:link') || attr(attrs, 'r:embed');
      } else if (name === 'p14:media' && cur) {
        var t = cur.video || cur.audio || (cur.video = {});
        t.embed = attr(attrs, 'r:embed'); t.mlink = attr(attrs, 'r:link');
      } else if (name === 'a:blip' && cur) {
        if (cur.blip === null) cur.blip = attr(attrs, 'r:embed');
        var bl = attr(attrs, 'r:link'); if (bl) cur.blipLink = bl;
      } else if (name === 'p:oleObj') s.ole++;
      else if (name === 'p:control') s.controls++;
      else if (name === 'am3d:model3d') s.model3d++;
      else if (name === 'p14:contentPart' || name === 'p:contentPart') s.ink++;
      else if (name === 'a:hlinkClick' || name === 'a:hlinkHover') s.hyperlinks++;
      else if (name === 'p:transition') { inTransition = stack.length + 1; if (!s.transition) s.transition = { type: 'cut', advanceAfterMs: attr(attrs, 'advTm') }; }
      else if (inTransition && stack.length === inTransition && s.transition && s.transition.type === 'cut') {
        s.transition.type = name.replace(/^.*:/, '');
      } else if (name === 'p:cTn') {
        var pc = attr(attrs, 'presetClass');
        effects.push(pc && attr(attrs, 'presetID') !== null ? { cls: pc, depth: stack.length } : null);
        if (pc && attr(attrs, 'presetID') !== null) {
          s.anim.total++;
          if (pc === 'entr') s.anim.entrance++; else if (pc === 'exit') s.anim.exit++;
          else if (pc === 'emph') s.anim.emphasis++; else if (pc === 'path') s.anim.path++;
        }
      } else if (name === 'p:spTgt') {
        var spid = attr(attrs, 'spid');
        for (var k = effects.length - 1; k >= 0; k--) {
          if (effects[k]) { if (animTargets[effects[k].cls]) animTargets[effects[k].cls][spid] = true; break; }
        }
      }
    } else {
      if (SHAPE_TAGS[name]) { shapeStack.pop(); if (name === 'p:grpSp') groupDepth--; }
      else if (name === 'p:cTn') effects.pop();
      else if (name === 'p:transition') inTransition = 0;
    }
  });

  /* media */
  shapes.forEach(function (sh) {
    var m = sh.video || sh.audio;
    if (!m) return;
    var rid = m.embed || m.mlink || m.link, rel = rid && rels.byId[rid];
    var linkedRel = (m.mlink && rels.byId[m.mlink]) || (m.link && rels.byId[m.link]);
    var linked = !m.embed && !!(linkedRel && linkedRel.external);
    var poster = sh.blip && rels.byId[sh.blip];
    s.media.push({
      type: sh.video ? 'video' : 'audio', shapeId: sh.id, name: sh.name,
      linked: linked, target: rel ? (rel.external ? rel.target : rel.part) : null,
      remote: !!(linked && linkedRel && isRemote(linkedRel.target)),
      box: sh.inGroup ? null : emuBox(sh.box), inGroup: sh.inGroup,
      posterPart: poster && !poster.external ? poster.part : null
    });
  });
  shapes.forEach(function (sh) {
    if (!sh.blipLink) return;
    var r = rels.byId[sh.blipLink];
    if (r && r.external) s.linkedImages.push({ target: r.target, remote: isRemote(r.target), hasEmbeddedCopy: !!sh.blip, shapeName: sh.name });
  });

  /* stacked animated shapes: things that appear over, or leave from over, something else */
  var animated = Object.keys(animTargets.entr).concat(Object.keys(animTargets.exit))
    .filter(function (id, i, a) { return a.indexOf(id) === i; })
    .map(function (id) { return byId[id]; })
    .filter(function (sh) { return sh && !sh.inGroup && emuBox(sh.box); });
  var pairs = 0;
  for (var i = 0; i < animated.length; i++) {
    for (var j = i + 1; j < animated.length; j++) {
      if (overlapShare(emuBox(animated[i].box), emuBox(animated[j].box)) >= 0.3) pairs++;
    }
  }
  s.anim.overlapping = pairs;
  s.anim.exitTargets = Object.keys(animTargets.exit).length;
  return s;
}

function notesWords(xml) {
  /* body placeholder text only. The slide number field and the slide image are not notes. */
  var words = 0, re = /<p:sp\b[\s\S]*?<\/p:sp>/g, m;
  while ((m = re.exec(xml))) {
    var sp = m[0];
    if (!/<p:ph\b[^>]*\btype\s*=\s*"body"/.test(sp)) continue;
    var t = sp.replace(/<a:fld\b[\s\S]*?<\/a:fld>/g, ''), tr = /<a:t>([^<]*)<\/a:t>/g, x;
    while ((x = tr.exec(t))) { var w = unescapeXml(x[1]).trim(); if (w) words += w.split(/\s+/).length; }
  }
  return words;
}

/* ------------------------------------------------------------------ *
 * A whole .pptx                                                       *
 * ------------------------------------------------------------------ */

async function inspectPptx(blob) {
  var zip = await openZip(blob);
  if (zip.has('EncryptedPackage')) throw coded('password', 'The file is password protected.');
  if (!zip.has('[Content_Types].xml')) throw coded('not-pptx', 'This zip is not an Office file.');

  var rootRels = await readRels(zip, '');
  var main = null, thumb = null;
  rootRels.list.forEach(function (r) {
    if (r.kind === 'officeDocument') main = r.part;
    if (r.kind === 'thumbnail') thumb = r.part;
  });
  if (!main || !zip.has(main)) throw coded('not-pptx', 'The file has no presentation part. It may be a Word or Excel file.');
  var types = await zip.text('[Content_Types].xml');
  var mainType = (new RegExp('PartName="/' + main.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"\\s+ContentType="([^"]*)"').exec(types) || [])[1] || '';
  if (!/presentationml/.test(mainType)) throw coded('not-pptx', 'The main part of this file is not a presentation.');

  var pxml = await zip.text(main), prels = await readRels(zip, main);
  var out = {
    kind: 'pptx', macro: /macroEnabled/.test(mainType), template: /template/.test(mainType), show: /slideshow/.test(mainType),
    slideWidthEmu: 0, slideHeightEmu: 0, widthPt: 0, heightPt: 0,
    slides: [], hiddenSlides: [], fonts: [], slideFonts: [], embeddedFonts: [], themeFonts: null,
    externalLinks: [], writeProtected: /<p:modifyVerifier\b/.test(pxml),
    thumbnailPart: thumb && zip.has(thumb) ? thumb : null,
    counts: { video: 0, audio: 0, notes: 0, transitions: 0, animated: 0, ole: 0, controls: 0, model3d: 0, linkedMedia: 0 },
    unread: []
  };

  var order = [];
  scan(pxml, function (kind, name, attrs, stack) {
    if (kind !== 'open') return;
    if (name === 'p:sldSz') { out.slideWidthEmu = +attr(attrs, 'cx'); out.slideHeightEmu = +attr(attrs, 'cy'); }
    else if (name === 'p:sldId' && stack[stack.length - 1] === 'p:sldIdLst') order.push(attr(attrs, 'r:id'));
    else if (name === 'p:font' && stack[stack.length - 1] === 'p:embeddedFont') { var t = attr(attrs, 'typeface'); if (t) out.embeddedFonts.push(t); }
  });
  if (!(out.slideWidthEmu > 0 && out.slideHeightEmu > 0)) throw coded('not-pptx', 'The presentation does not state a slide size.');
  out.widthPt = out.slideWidthEmu / EMU_PER_PT; out.heightPt = out.slideHeightEmu / EMU_PER_PT;

  var themeRel = prels.list.filter(function (r) { return r.kind === 'theme'; })[0];
  out.themeFonts = await readTheme(zip, themeRel && themeRel.part);

  var fontSet = Object.create(null), explicitSet = Object.create(null), layoutCache = Object.create(null);

  async function layoutFonts(layoutPart) {
    if (!layoutPart) return { fonts: {}, theme: out.themeFonts, slideNumberField: false };
    if (layoutCache[layoutPart]) return layoutCache[layoutPart];
    var res = { fonts: Object.create(null), theme: out.themeFonts, slideNumberField: false };
    layoutCache[layoutPart] = res;
    try {
      var lx = await zip.text(layoutPart); collectFonts(lx, res.fonts);
      res.slideNumberField = /<a:fld\b[^>]*\btype\s*=\s*"slidenum"/.test(lx);
      var lr = await readRels(zip, layoutPart);
      var mr = lr.list.filter(function (r) { return r.kind === 'slideMaster'; })[0];
      if (mr && zip.has(mr.part)) {
        var mx = await zip.text(mr.part); collectFonts(mx, res.fonts);
        if (/<a:fld\b[^>]*\btype\s*=\s*"slidenum"/.test(mx)) res.slideNumberField = true;
        var mrels = await readRels(zip, mr.part);
        var tr = mrels.list.filter(function (r) { return r.kind === 'theme'; })[0];
        if (tr) res.theme = await readTheme(zip, tr.part);
      }
    } catch (e) { out.unread.push(layoutPart); }
    return res;
  }

  for (var i = 0; i < order.length; i++) {
    var rel = prels.byId[order[i]];
    var n = i + 1;
    if (!rel || !zip.has(rel.part)) { out.slides.push({ n: n, unreadable: true }); out.unread.push(rel ? rel.part : 'slide ' + n); continue; }
    var sx, srels, s;
    try {
      sx = await zip.text(rel.part); srels = await readRels(zip, rel.part); s = parseSlide(sx, srels);
    } catch (e) { out.slides.push({ n: n, unreadable: true }); out.unread.push(rel.part); continue; }
    s.n = n; s.part = rel.part;

    var lrel = srels.list.filter(function (r) { return r.kind === 'slideLayout'; })[0];
    var lay = await layoutFonts(lrel && lrel.part);
    var used = resolveFontNames(s.fonts, lay.theme);
    Object.keys(used).forEach(function (k) { explicitSet[k] = true; });      /* named on the slide itself */
    if (s.hasText) {
      /* text with no font named on the run takes the layout, master and theme fonts */
      var inherited = resolveFontNames(lay.fonts, lay.theme);
      Object.keys(inherited).forEach(function (k) { used[k] = true; });
      if (lay.theme.minor) used[lay.theme.minor] = true;
      if (lay.theme.major) used[lay.theme.major] = true;
    }
    s.fonts = Object.keys(used).sort();
    s.fonts.forEach(function (f) { fontSet[f] = true; });

    var nrel = srels.list.filter(function (r) { return r.kind === 'notesSlide'; })[0];
    s.notesWords = 0;
    if (nrel && zip.has(nrel.part)) {
      try { s.notesWords = notesWords(await zip.text(nrel.part)); } catch (e) { out.unread.push(nrel.part); }
    }
    s.hasNotes = s.notesWords > 0;

    if (s.hidden) out.hiddenSlides.push(n);
    if (s.hasNotes) out.counts.notes++;
    if (s.transition) out.counts.transitions++;
    if (s.anim.total) out.counts.animated++;
    out.counts.ole += s.ole; out.counts.controls += s.controls; out.counts.model3d += s.model3d;
    s.media.forEach(function (m) {
      if (m.type === 'video') out.counts.video++; else out.counts.audio++;
      if (m.linked) { out.counts.linkedMedia++; out.externalLinks.push({ slide: n, kind: m.type, target: m.target, remote: m.remote }); }
    });
    s.linkedImages.forEach(function (l) { out.externalLinks.push({ slide: n, kind: 'image', target: l.target, remote: l.remote, hasEmbeddedCopy: l.hasEmbeddedCopy }); });
    out.slides.push(s);
  }

  out.fonts = Object.keys(fontSet).sort();
  out.slideFonts = Object.keys(explicitSet).sort();
  out.slideCount = out.slides.length;
  out.readBytes = async function (part, cap) { return zip.bytes(part, cap); };
  return out;
}

/* ------------------------------------------------------------------ *
 * OLE compound file (.ppt, and password protected .pptx)              *
 * ------------------------------------------------------------------ */

var CFB_MAGIC = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
var ENDOFCHAIN = 0xFFFFFFFE, FREESECT = 0xFFFFFFFF;

function isCfb(head) { for (var i = 0; i < 8; i++) if (head[i] !== CFB_MAGIC[i]) return false; return true; }
function isZip(head) { return head[0] === 0x50 && head[1] === 0x4b && (head[2] === 3 || head[2] === 5) ; }

async function openCfb(blob) {
  var head = await sliceBytes(blob, 0, 512);
  if (head.length < 512 || !isCfb(head)) throw coded('not-cfb', 'Not an OLE compound file.');
  var hv = new DataView(head.buffer, head.byteOffset, head.byteLength);
  var ss = 1 << u16(hv, 30), mss = 1 << u16(hv, 32);
  var nFat = u32(hv, 44), dirStart = u32(hv, 48), cutoff = u32(hv, 56);
  var miniFatStart = u32(hv, 60), difatStart = u32(hv, 68), nDifat = u32(hv, 72);
  if (ss !== 512 && ss !== 4096) throw coded('cfb-damaged', 'The file header is damaged.');
  var secOff = function (n) { return (n + 1) * ss; };
  var maxSectors = Math.ceil(blob.size / ss) + 1;

  var fatSecs = [], i;
  for (i = 0; i < 109 && fatSecs.length < nFat; i++) { var v = u32(hv, 76 + i * 4); if (v < ENDOFCHAIN) fatSecs.push(v); }
  var next = difatStart, guard = 0;
  while (nDifat > 0 && next < ENDOFCHAIN && guard++ < 4096) {
    var ds = await sliceBytes(blob, secOff(next), secOff(next) + ss), dvv = new DataView(ds.buffer, ds.byteOffset, ds.byteLength);
    for (i = 0; i < ss / 4 - 1 && fatSecs.length < nFat; i++) { var w = u32(dvv, i * 4); if (w < ENDOFCHAIN) fatSecs.push(w); }
    next = u32(dvv, ss - 4);
  }
  var per = ss / 4, fat = new Uint32Array(fatSecs.length * per);
  for (i = 0; i < fatSecs.length; i++) {
    var fs = await sliceBytes(blob, secOff(fatSecs[i]), secOff(fatSecs[i]) + ss);
    if (fs.length < ss) throw coded('cfb-damaged', 'The file is cut short.');
    var fv = new DataView(fs.buffer, fs.byteOffset, fs.byteLength);
    for (var j = 0; j < per; j++) fat[i * per + j] = u32(fv, j * 4);
  }
  function chain(start) {
    var out = [], n = start, g = 0;
    while (n < ENDOFCHAIN && n !== FREESECT && g++ < maxSectors) { out.push(n); n = n < fat.length ? fat[n] : ENDOFCHAIN; }
    return out;
  }
  async function readChain(start, size) {
    var secs = chain(start), want = size == null ? secs.length * ss : size;
    var buf = new Uint8Array(want), o = 0;
    for (var k = 0; k < secs.length && o < want; k++) {
      var part = await sliceBytes(blob, secOff(secs[k]), secOff(secs[k]) + Math.min(ss, want - o));
      buf.set(part, o); o += part.length;
      if (part.length === 0) break;
    }
    return buf;
  }

  var dir = await readChain(dirStart), dview = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  var entries = [], rootEntry = null;
  for (i = 0; i + 128 <= dir.length; i += 128) {
    var nameLen = u16(dview, i + 64), type = dir[i + 66];
    if (!type || nameLen < 2) continue;
    var nm = '';
    for (var c = 0; c < nameLen - 2 && c < 62; c += 2) nm += String.fromCharCode(u16(dview, i + c));
    var en = { name: nm, type: type, start: u32(dview, i + 116), size: u64(dview, i + 120) };
    if (type === 5) rootEntry = en; else entries.push(en);
  }
  var mini = null, miniFat = null;
  async function miniStream() {
    if (mini) return;
    mini = rootEntry ? await readChain(rootEntry.start, rootEntry.size) : new Uint8Array(0);
    var mf = miniFatStart < ENDOFCHAIN ? await readChain(miniFatStart) : new Uint8Array(0);
    miniFat = new Uint32Array(Math.floor(mf.length / 4));
    var mv = new DataView(mf.buffer, mf.byteOffset, mf.byteLength);
    for (var k = 0; k < miniFat.length; k++) miniFat[k] = u32(mv, k * 4);
  }
  return {
    names: entries.map(function (e) { return e.name; }),
    size: function (name) { var e = entries.filter(function (x) { return x.name === name; })[0]; return e ? e.size : -1; },
    read: async function (name) {
      var e = entries.filter(function (x) { return x.name === name && x.type === 2; })[0];
      if (!e) return null;
      if (e.size >= cutoff) return readChain(e.start, e.size);
      await miniStream();
      var out = new Uint8Array(e.size), o = 0, n = e.start, g = 0;
      while (n < ENDOFCHAIN && o < e.size && g++ < miniFat.length + 1) {
        var take = Math.min(mss, e.size - o);
        out.set(mini.subarray(n * mss, n * mss + take), o); o += take;
        n = n < miniFat.length ? miniFat[n] : ENDOFCHAIN;
      }
      return out;
    }
  };
}

/* PowerPoint 97-2003 binary. Sizes are in master units, 576 to the inch. */
async function inspectPpt(blob) {
  var cfb = await openCfb(blob);
  if (cfb.names.indexOf('EncryptedPackage') >= 0) throw coded('password', 'The file is password protected.');
  if (cfb.names.indexOf('PowerPoint Document') < 0) {
    throw coded('not-ppt', cfb.names.indexOf('WordDocument') >= 0 ? 'This is a Word document.' :
      cfb.names.indexOf('Workbook') >= 0 ? 'This is an Excel workbook.' : 'This file is not a PowerPoint presentation.');
  }
  var out = { kind: 'ppt', slideCount: null, widthPt: null, heightPt: null, hiddenSlides: null, slides: null,
              fonts: null, embeddedFonts: [], externalLinks: [], counts: null, writeProtected: null, thumbnailPart: null, unread: [] };

  var cu = await cfb.read('Current User');
  if (cu && cu.length >= 16) {
    var token = new DataView(cu.buffer, cu.byteOffset, cu.byteLength).getUint32(12, true);
    if (token === 0xF3D1C4DF) throw coded('password', 'The file is password protected.');
  }
  var size = cfb.size('PowerPoint Document');
  if (size > MAX_PPT_STREAM) { out.unread.push('PowerPoint Document'); return out; }
  var doc = await cfb.read('PowerPoint Document');
  var dv = new DataView(doc.buffer, doc.byteOffset, doc.byteLength);

  /* top level records. The last Document container is the current one. */
  var p = 0, docAt = -1, docLen = 0;
  while (p + 8 <= doc.length) {
    var type = u16(dv, p + 2), len = u32(dv, p + 4);
    if (type === 1000) { docAt = p + 8; docLen = len; }
    if (len > doc.length) break;
    p += 8 + len;
  }
  if (docAt < 0) { out.unread.push('PowerPoint Document'); return out; }
  var end = Math.min(doc.length, docAt + docLen), q = docAt, slides = null;
  while (q + 8 <= end) {
    var verInst = u16(dv, q), t = u16(dv, q + 2), l = u32(dv, q + 4);
    if (t === 1001 && l >= 8) {                         /* DocumentAtom */
      out.widthPt = dv.getInt32(q + 8, true) / 8; out.heightPt = dv.getInt32(q + 12, true) / 8;
    } else if (t === 4080 && (verInst >> 4) === 0) {    /* SlideListWithText, instance 0 is the slides */
      var c = q + 8, ce = Math.min(end, c + l), n = 0;
      while (c + 8 <= ce) { if (u16(dv, c + 2) === 1011) n++; c += 8 + u32(dv, c + 4); }
      slides = n;
    }
    q += 8 + l;
  }
  out.slideCount = slides;
  return out;
}

/* ------------------------------------------------------------------ *
 * Entry point                                                         *
 * ------------------------------------------------------------------ */

function extOf(name) { var m = /\.([A-Za-z0-9]+)$/.exec(name || ''); return m ? m[1].toLowerCase() : ''; }

/* Never throws. Returns { ok, kind, error: {code, message}, ...facts }. */
async function inspect(blob, name) {
  var ext = extOf(name);
  try {
    if (blob.size === 0) throw coded('empty', 'The file is empty.');
    var head = await sliceBytes(blob, 0, 8);
    if (isZip(head)) {
      var z = await inspectPptx(blob);
      z.ok = true; z.extMismatch = ext === 'ppt';
      return z;
    }
    if (isCfb(head)) {
      var c = await inspectPpt(blob);
      c.ok = true; c.extMismatch = ext === 'pptx';
      return c;
    }
    throw coded('unsupported', 'The file is not a PowerPoint presentation. It does not start like a .pptx or a .ppt.');
  } catch (e) {
    return { ok: false, kind: ext === 'ppt' ? 'ppt' : 'pptx',
             error: { code: e.code || 'unreadable', message: e.message || String(e) } };
  }
}

return {
  EMU_PER_PT: EMU_PER_PT,
  openZip: openZip, openCfb: openCfb, scan: scan, attr: attr, unescapeXml: unescapeXml,
  resolveTarget: resolveTarget, relsPartOf: relsPartOf, isRemote: isRemote,
  parseSlide: parseSlide, notesWords: notesWords,
  inspectPptx: inspectPptx, inspectPpt: inspectPpt, inspect: inspect, extOf: extOf
};
});
