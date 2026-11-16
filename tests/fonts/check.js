/* Engine checks for Font Extractor.

   Needs   python3 with python-pptx and fontTools, DejaVu Sans installed
           (to draw the decks), npm package jszip (or NODE_MODULES_DIR
           pointing at a folder that has it).
   Run     node tests/fonts/check.js
   Extra   REAL_DECKS=/folder/of/pptx also runs every deck in that folder
           through the engine and prints what it found. Use it with decks
           saved from PowerPoint, whose fonts are MTX compressed. */
'use strict';
const path = require('path'), fs = require('fs'), os = require('os'), vm = require('vm');
const { spawnSync } = require('child_process');
function need(name) {
  try { return require(name); } catch (e) {}
  if (process.env.NODE_MODULES_DIR) return require(path.join(process.env.NODE_MODULES_DIR, name));
  throw new Error('Cannot find ' + name + '. npm i ' + name + ' or set NODE_MODULES_DIR.');
}
const JSZip = need('jszip');
const root = path.resolve(__dirname, '..', '..');
const FX = require(path.join(root, 'font-extract.js'));
const box = {}; vm.createContext(box);
vm.runInContext(fs.readFileSync(path.join(root, 'mtx-decompressor.min.js'), 'utf8') + ';this.M = MTXDecompressor;', box);
const mtx = box.M;

const dir = process.env.WORK || path.join(os.tmpdir(), 'font-extractor-checks');
const r = spawnSync('python3', [path.join(__dirname, 'make_decks.py'), dir], { stdio: 'inherit' });
if (r.status !== 0) throw new Error('could not draw the decks');

let fails = 0;
function ok(cond, msg) { console.log((cond ? '  ok   ' : '  FAIL ') + msg); if (!cond) fails++; }
async function load(name) { return JSZip.loadAsync(fs.readFileSync(path.join(dir, name))); }
const slot = (d, s) => d.fonts.find(f => f.slot === s);

(async () => {
  console.log('full.pptx');
  let d = await FX.readDeck(await load('full.pptx'), { mtx });
  ok(d.fonts.length === 2, 'two fonts listed');
  ok(d.embedFlag, 'embed flag read');
  let reg = slot(d, 'regular'), bold = slot(d, 'bold');
  ok(reg.typeface === 'DejaVu Sans' && reg.container === 'eot' && !reg.error, 'regular unwrapped from EOT');
  ok(bold.eot.encrypted && FX.isSfnt(bold.sfnt), 'XOR-scrambled bold unwrapped');
  ok(reg.verdict.level === 'full' && bold.verdict.level === 'full', 'both judged whole');
  ok(reg.info.names.postscript === 'DejaVuSans' && reg.fileName === 'DejaVuSans.ttf', 'named from the font, ' + reg.fileName);
  ok(bold.fileName === 'DejaVuSans-Bold.ttf', 'bold named ' + bold.fileName);
  ok(reg.info.numGlyphs > 5000 && reg.info.chars.length > 5000, 'glyphs and characters counted, ' + reg.info.numGlyphs + ' / ' + reg.info.chars.length);
  ok(FX.permission(reg.info.fsType).key === 'installable', 'embedding permission read');

  console.log('subset.pptx');
  d = await FX.readDeck(await load('subset.pptx'), { mtx });
  reg = slot(d, 'regular'); bold = slot(d, 'bold');
  const ital = slot(d, 'italic'), bi = slot(d, 'boldItalic'), stray = d.fonts.find(f => f.orphan);
  ok(reg.verdict.level === 'subset' && reg.eot.subset, 'cut-down regular flagged, header flag seen');
  ok(/Missing basic/.test(reg.verdict.reasons.join()), 'missing characters reported');
  ok(reg.info.chars.map(c => String.fromCodePoint(c)).join('') === ' 26HWdelor', 'characters left are exactly " 26HWdelor"');
  ok(/-SUBSET\.ttf$/.test(reg.fileName), 'subset file name marked, ' + reg.fileName);
  ok(bold.verdict.level === 'subset' && !bold.eot.subset, 'hollowed bold flagged with no header flag');
  ok(/outlines removed/.test(bold.verdict.reasons.join()), 'emptied outlines counted, ' + bold.info.emptyGlyphs + ' of ' + bold.info.numGlyphs);
  ok(bold.info.chars.filter(c => c > 32 && c < 127).map(c => String.fromCodePoint(c)).join('') === '26HWdelor', 'mapped but emptied glyphs not counted as characters');
  ok(ital.verdict.level === 'check' && /lowercase/.test(ital.verdict.reasons.join()), 'font with no lowercase asks for a check, not a flag');
  ok(bi.sameAs === reg, 'bold italic spotted as the same file as regular');
  ok(stray && stray.container === 'plain' && stray.verdict.level === 'full', 'unlisted raw font part picked up');

  console.log('none.pptx');
  d = await FX.readDeck(await load('none.pptx'), { mtx });
  ok(!d.hasFonts, 'no fonts reported');

  console.log('strip');
  const zip = await load('subset.pptx');
  const s = await FX.stripDeck(zip);
  const buf = await s.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const back = await JSZip.loadAsync(buf);
  const names = Object.keys(back.files);
  const pres = await back.file('ppt/presentation.xml').async('string');
  const rels = await back.file('ppt/_rels/presentation.xml.rels').async('string');
  ok(s.removed.length === 4, 'four font parts removed, stray one included');
  ok(!names.some(n => /fonts\//.test(n)), 'no font parts left in the package');
  ok(!/embeddedFont|embedTrueTypeFonts/.test(pres), 'font list and embed flags gone from presentation.xml');
  ok(!/relationships\/font"/.test(rels) && /relationships\/slide"/.test(rels), 'font relationships gone, the rest kept');
  ok(names.filter(n => !/\/$/.test(n)).length === Object.keys((await load('subset.pptx')).files).filter(n => !/fonts\//.test(n) && !/\/$/.test(n)).length, 'every other part still there');
  d = await FX.readDeck(back, { mtx });
  ok(!d.hasFonts, 'stripped deck reads as having no fonts');
  fs.writeFileSync(path.join(dir, 'subset-stripped.pptx'), buf);
  const py = spawnSync('python3', ['-c', 'import sys;from pptx import Presentation;p=Presentation(sys.argv[1]);print(len(p.slides))', path.join(dir, 'subset-stripped.pptx')]);
  ok(py.status === 0 && String(py.stdout).trim() === '1', 'stripped deck opens in python-pptx');

  console.log('odd input');
  let u = FX.unwrap(new Uint8Array(200), { mtx });
  ok(u.container === 'unknown' && u.error, 'junk is refused, not crashed on');
  u = FX.unwrap(new Uint8Array(4), { mtx });
  ok(u.error, 'tiny file refused');

  if (process.env.REAL_DECKS) {
    console.log('real decks');
    for (const f of fs.readdirSync(process.env.REAL_DECKS).filter(n => /\.pp[ts][xm]$/i.test(n))) {
      const dk = await FX.readDeck(await JSZip.loadAsync(fs.readFileSync(path.join(process.env.REAL_DECKS, f))), { mtx });
      console.log('  ' + f + (dk.fonts.length ? '' : '  no embedded fonts'));
      dk.fonts.forEach(x => {
        console.log('    ' + [x.typeface, x.slot, x.container + (x.eot && x.eot.compressed ? '+mtx' : ''), x.verdict.level, x.fileName, x.error || ''].join('  '));
        ok(!x.error, f + ' ' + x.slot + ' unwrapped');
      });
    }
  }

  console.log(fails ? fails + ' FAILED' : 'all passed');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
