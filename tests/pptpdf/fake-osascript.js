#!/usr/bin/env node
/* A stand-in for macOS osascript, used only by the automated checks.

   It lets the Mac helper's shell logic run on Linux: the queue, the time
   limits, the two ways of leaving slides out and the result files. It
   pretends to be PowerPoint in the plainest way. A "presentation" is a
   small JSON file such as {"slides":6,"hidden":[3]} and a PDF is blank
   pages. The real AppleScript inside the helper is NOT run by this, so it
   says nothing about whether that AppleScript is right. */
'use strict';
const fs = require('fs'), path = require('path');
const args = process.argv.slice(2);
const stateFile = path.join(process.env.SLIDESIZE_FAKE_STATE || '/tmp', 'fake-powerpoint.json');
const load = () => { try { return JSON.parse(fs.readFileSync(stateFile, 'utf8')); } catch (e) { return { open: {} }; } };
const save = s => fs.writeFileSync(stateFile, JSON.stringify(s));
const fail = (msg, code) => { process.stderr.write('fake.applescript:1:2: execution error: ' + msg + ' (' + (code || -1728) + ')\n'); process.exit(1); };

function pdf(file, pages, w, h) {
  const objs = ['<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Count ' + pages + ' /Kids [ ' + Array.from({ length: pages }, (_, i) => (3 + i) + ' 0 R').join(' ') + ' ] >>'];
  for (let i = 0; i < pages; i++) objs.push('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ' + w + ' ' + h + '] /Resources << >> >>');
  let out = '%PDF-1.4\n'; const offs = [];
  objs.forEach((o, i) => { offs.push(out.length); out += (i + 1) + ' 0 obj\n' + o + '\nendobj\n'; });
  const xref = out.length;
  out += 'xref\n0 ' + (objs.length + 1) + '\n0000000000 65535 f \n' + offs.map(o => String(o).padStart(10, '0') + ' 00000 n \n').join('');
  out += 'trailer\n<< /Size ' + (objs.length + 1) + ' /Root 1 0 R /ID [<00112233445566778899aabbccddeeff> <00112233445566778899aabbccddeeff>] >>\nstartxref\n' + xref + '\n%%EOF\n';
  fs.writeFileSync(file, out, 'latin1');
}

if (args[0] === '-l' && args[1] === 'JavaScript' && args[2] === '-e') {
  /* the helper's own JavaScript runs for real, against stand-ins for the two macOS frameworks it uses */
  const list = a => ({ count: a.length, objectAtIndex: i => a[i] });
  const sandbox = {
    ObjC: { import() {}, unwrap: v => v },
    $: {
      NSFontManager: { sharedFontManager: { availableFontFamilies: list(['Arial', 'Carlito']), availableFonts: list(['ArialMT', 'Carlito-Bold']) } },
      NSURL: { fileURLWithPath: p => p },
      PDFDocument: { alloc: { initWithURL: p => {
        let n = null;
        try { n = (fs.readFileSync(p, 'latin1').match(/\/Type \/Page\b/g) || []).length; } catch (e) { n = null; }
        return { isNil: () => n === null || n === 0, pageCount: n };
      } } }
    }
  };
  const run = new Function('ObjC', '$', args[3] + '; return run;')(sandbox.ObjC, sandbox.$);
  process.stdout.write(String(run(args.slice(4))) + '\n');
  process.exit(0);
}
if (args[0] === '-e') { process.stdout.write((process.env.SLIDESIZE_FAKE_PICKED || '') + '\n'); process.exit(0); }

const verb = path.basename(args[0], '.applescript'), a = args.slice(1), st = load();
const deck = name => st.open[name] || fail('Can’t get presentation "' + name + '".');
switch (verb) {
  case 'version':
    if (process.env.SLIDESIZE_FAKE_START_FAILS === '1') fail('Not authorised to send Apple events to Microsoft PowerPoint.', -1743);
    console.log('16.0'); break;
  case 'others': console.log(Object.keys(st.open).filter(n => !n.startsWith(a[0])).length); break;
  case 'closestale': Object.keys(st.open).filter(n => n.startsWith(a[0])).forEach(n => delete st.open[n]); save(st); console.log('ok'); break;
  case 'open': {
    let d; try { d = JSON.parse(fs.readFileSync(a[0], 'utf8')); } catch (e) { fail('The file could not be opened.'); }
    if (d.stall) { setInterval(() => {}, 1000); return; }
    if (d.fail) fail(d.fail);
    const hidden = d.hidden || [];
    st.open[path.basename(a[0])] = { w: d.w || 960, h: d.h || 540, slides: Array.from({ length: d.slides }, (_, i) => ({ n: i + 1, hidden: hidden.includes(i + 1) })), shapes: [] };
    save(st); console.log('ok'); break;
  }
  case 'facts': { const d = deck(a[0]); console.log([d.slides.length, d.w, d.h].join('\t')); break; }
  case 'hidden': {
    if (process.env.SLIDESIZE_FAKE_NO_HIDDEN_TERM === '1') fail('Can’t get hidden of slide show transition.');
    const d = deck(a[0]); console.log('ok ' + d.slides.map((s, i) => s.hidden ? (i + 1) : '').filter(Boolean).join(' ') + ' '); break;
  }
  case 'sethidden': { const d = deck(a[0]); a.slice(2).forEach(i => { d.slides[i - 1].hidden = a[1] === '1'; }); save(st); console.log('ok'); break; }
  case 'delete': { const d = deck(a[0]); a.slice(1).forEach(i => d.slides.splice(i - 1, 1)); save(st); console.log('ok'); break; }
  case 'placeholder': { const d = deck(a[0]); if (+a[1] > d.slides.length) fail('Can’t get slide ' + a[1] + '.'); d.shapes.push(a.slice(1)); save(st); console.log('ok'); break; }
  case 'pdf': {
    const d = deck(a[0]);
    /* PowerPoint for Mac may or may not skip hidden slides. Both are played here. */
    const shown = process.env.SLIDESIZE_FAKE_EXPORTS_HIDDEN === '1' ? d.slides : d.slides.filter(s => !s.hidden);
    pdf(a[1], shown.length, d.w, d.h); console.log('ok'); break;
  }
  case 'png': {
    const d = deck(a[0]); fs.mkdirSync(a[1], { recursive: true });
    const shown = process.env.SLIDESIZE_FAKE_EXPORTS_HIDDEN === '1' ? d.slides : d.slides.filter(s => !s.hidden);
    shown.forEach((s, i) => fs.writeFileSync(path.join(a[1], 'Slide' + (i + 1) + '.png'), 'picture of original slide ' + s.n));
    console.log('ok'); break;
  }
  case 'close': delete st.open[a[0]]; save(st); console.log('ok'); break;
  case 'quit': st.open = {}; st.quit = true; save(st); console.log('ok'); break;
  default: fail('The stand-in does not know the script ' + verb + '.');
}
