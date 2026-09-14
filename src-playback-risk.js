/**
 * playback-risk.js
 * Turns the output of mp4-inspect.js into a ranked list of playback findings.
 *
 * The rules are deliberately separate from the parser so they can be tuned
 * without touching bitstream code. Everything here is opinion; the parser is fact.
 *
 * AEGFX / SlideSize
 */

export const SEVERITY = { RED: 'red', AMBER: 'amber', INFO: 'info' };
const RANK = { red: 3, amber: 2, info: 1 };

/* ------------------------------------------------------------------ *
 * Target profiles                                                     *
 * ------------------------------------------------------------------ */

export const TARGETS = {
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
export function assess(info, opts = {}) {
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
export function buildFixCommand(info, assessment, opts = {}) {
  const v = info.video[0];
  if (!v) return null;
  const real = assessment.findings.filter((x) => x.severity !== SEVERITY.INFO);
  if (!real.length) return null;                 // nothing to fix
  const codes = new Set(real.map((x) => x.code));
  const target = TARGETS[opts.target] || TARGETS['powerpoint-win'];
  const name = info.name || 'video.mp4';

  // A misdeclared level on its own needs a header rewrite, not a re-encode.
  if (real.every((x) => x.code === 'level-mismatch') && v.levelCheck) {
    const lvl = v.levelCheck.requiredByThroughput || v.levelCheck.requiredByFrameSize || '5';
    return {
      reencode: false,
      level: lvl,
      notes: [],
      command: `ffmpeg -i "${name}" -c copy -bsf:v h264_metadata=level=${lvl} -movflags +faststart "fixed_${name}"`,
    };
  }

  const srcW = v.cropped?.width || v.coded?.width;
  const srcH = v.cropped?.height || v.coded?.height;
  let outW = srcW, outH = srcH;
  const filters = [];

  if (v.interlaced?.flag) filters.push('yadif=1');

  // Non-square pixels: resample to the intended display geometry.
  if (codes.has('non-square-pixels') && v.display) {
    outW = Math.round(v.display.width / 2) * 2;
    outH = Math.round(v.display.height / 2) * 2;
    filters.push(`scale=${outW}:${outH}:flags=lanczos`);
  }

  // Only resize when the assessment actually said the picture is too big,
  // and then against the target's own limits rather than a hardcoded 1080.
  if (codes.has('oversize')) {
    const maxW = opts.maxWidth || target.maxWidth;
    const maxH = opts.maxHeight || target.maxHeight;
    const k = Math.min(maxW / outW, maxH / outH, 1);
    if (k < 1) {
      outW = Math.round((outW * k) / 2) * 2;
      outH = Math.round((outH * k) / 2) * 2;
      filters.push(`scale=${outW}:${outH}:flags=lanczos`);
    }
  }

  // 4:2:0 cannot carry odd dimensions.
  if (outW % 2 || outH % 2) {
    outW -= outW % 2; outH -= outH % 2;
    filters.push('crop=trunc(iw/2)*2:trunc(ih/2)*2');
  }
  if (v.squarePixels === false) filters.push('setsar=1');

  const fps = opts.projectFrameRate || v.frameRate?.nominalFps || v.frameRate?.fps || 25;
  const gop = Math.max(2, Math.round(fps * 2));

  // Pick the level the OUTPUT actually needs. Hardcoding one is how files end
  // up misdeclared in the first place, which is the bug this tool exists to find.
  const mbs = Math.ceil(outW / 16) * Math.ceil(outH / 16);
  const lvlEntry = H264_LEVELS.find((l) => l.level >= 3 && mbs <= l.maxFS && mbs * fps <= l.maxMBPS)
    || H264_LEVELS[H264_LEVELS.length - 1];

  const parts = [`ffmpeg -i "${name}"`];
  if (filters.length) parts.push(`-vf "${filters.join(',')}"`);
  parts.push(`-c:v libx264 -profile:v high -level ${lvlEntry.name} -crf 18 -preset medium`);
  parts.push(`-r ${fps} -fps_mode cfr -g ${gop} -keyint_min ${Math.round(fps)}`);
  parts.push('-pix_fmt yuv420p');

  // Only retag colour when it is missing or wrong; never claim a conversion
  // that these flags do not perform.
  if (codes.has('colour-mismatch') || !v.colour || !v.colour.primaries) {
    parts.push('-colorspace bt709 -color_primaries bt709 -color_trc bt709');
  }

  parts.push(info.audio.length ? '-c:a aac -b:a 192k -ar 48000 -ac 2' : '-an');
  parts.push(`-movflags +faststart "fixed_${name.replace(/\.[^.]+$/, '')}.mp4"`);

  // Consequences of this command that are not obvious from reading it.
  const notes = [];
  if (codes.has('edit-list-trim') && v.trimmedHeadSeconds) {
    notes.push(`ffmpeg honours the edit list, so this bakes the trim in and discards the first `
      + `${v.trimmedHeadSeconds}s permanently. That is usually what you want. To keep those frames `
      + 'instead, add `-ignore_editlist 1` before `-i`.');
  }
  if (v.bitrate && v.bitrate < 5e6) {
    notes.push(`CRF 18 is visually lossless and will likely make the file larger than the `
      + `${fmtBitrate(v.bitrate)} source. If size matters more than fidelity, raise it to \`-crf 22\`.`);
  }
  if (codes.has('vfr')) {
    notes.push('Converting variable to constant frame rate duplicates or drops frames to fill the gaps. '
      + 'Motion in the sparse sections will look slightly different.');
  }

  return { reencode: true, level: lvlEntry.name, outWidth: outW, outHeight: outH,
    notes: notes, command: parts.join(' ') };
}

/* ------------------------------------------------------------------ *
 * Formatting helpers                                                  *
 * ------------------------------------------------------------------ */

export function fmtBitrate(bps) {
  if (!bps) return 'unknown';
  if (bps >= 1e6) return `${(bps / 1e6).toFixed(2)} Mbps`;
  return `${Math.round(bps / 1000)} kbps`;
}

export function fmtFps(fps) {
  if (!fps) return 'unknown';
  return `${Number(fps.toFixed(3))} fps`;
}

export function fmtDuration(seconds) {
  if (seconds == null) return 'unknown';
  const s = Math.floor(seconds % 60), m = Math.floor((seconds / 60) % 60), h = Math.floor(seconds / 3600);
  const frac = (seconds % 1).toFixed(2).slice(1);
  return h ? `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}` : `${m}:${String(s).padStart(2, '0')}${frac}`;
}

export function fmtBytes(n) {
  if (n == null) return 'unknown';
  const u = ['B', 'KB', 'MB', 'GB', 'TB'];
  let i = 0;
  while (n >= 1024 && i < u.length - 1) { n /= 1024; i++; }
  return `${n.toFixed(i ? 1 : 0)} ${u[i]}`;
}

/** Flat key/value summary, handy for a report table or a CSV export. */
export function summarise(info) {
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
