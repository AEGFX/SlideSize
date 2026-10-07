/* Browser checks for PowerPoint to PDF.

   These drive the real page in Chromium through the whole workflow: select,
   check, settings, convert, review, save, interrupt and resume.

   Two things on a test machine are not the real thing, and the checks say
   so rather than pretend.

     The folder   A headless browser cannot show a folder picker, so the
                  page is handed a private browser folder instead. The page
                  code that reads and writes it is the code that runs on a
                  real disk.

     PowerPoint   There is none here. mock-helper.js stands in for the
                  local helper and LibreOffice makes the PDFs. So these
                  checks prove what the PAGE does with a result. They say
                  nothing about how PowerPoint draws a slide, and nothing
                  about the helper scripts driving PowerPoint. The helper
                  scripts have their own checks in tests/pptpdf-helper.test.js.

   They are optional. The site does not need this folder.

   Needs   node 18+, LibreOffice (soffice) and poppler (pdftoppm) on the path,
           npm packages playwright and pdfjs-dist@3.11.174. Put them in
           tests/pptpdf/node_modules, or point PPTPDF_DEPS at a node_modules
           folder that has them.
   Run     node tests/pptpdf/browser/run.js          everything
           node tests/pptpdf/browser/run.js "D|E"    some groups           */
'use strict';
const fs = require('fs'), path = require('path');
const L = require('./lib.js');
const { chromium } = L.need('playwright');

const port = Number(process.env.PORT || 8933), base = 'http://localhost:' + port;
const only = process.argv[2] ? new RegExp('^(' + process.argv[2] + ')$') : null;
let passed = 0, failed = 0;
const fails = [];

function ok(cond, what) { if (cond) passed++; else { failed++; fails.push(what); console.log('    FAIL  ' + what); } }
function eq(a, b, what) { ok(JSON.stringify(a) === JSON.stringify(b), what + '  expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); }
const by = (list, name) => list.find(i => i.name === name);
const codes = it => it.issues.map(i => i.code);
const pdfInfo = async (page, p) => L.Pdf.inspect(new Uint8Array(await L.getFile(page, p)));
/* textContent, because innerText follows CSS and comes back in capitals or empty for collapsed parts */
const text = async (page, sel) => { await page.waitForTimeout(120); return page.evaluate(s => document.querySelector(s).textContent.replace(/\s+/g, ' '), sel); };

async function addFolder(page, names) {
  await L.putDecks(page, names);
  await page.click('#btn-add-folder');
  await L.waitChecked(page);
}
async function ready(page) { await page.click('#btn-folder'); await L.waitHelper(page); await page.waitForFunction(() => !document.getElementById('btn-start').disabled, null, { timeout: 30000 }); }
async function convert(page) { await page.click('#btn-start'); await L.waitRun(page); return L.items(page); }
async function advanced(page) { await page.evaluate(() => { document.getElementById('adv').open = true; }); }

const groups = {

  async A_page_and_platform(browser) {
    for (const plat of ['windows', 'macos']) {
      const { page, errors, ctx } = await L.open(browser, base, { platform: plat, mock: false });
      const body = await text(page, 'body');
      ok(body.includes('Microsoft PowerPoint must be installed for PowerPoint-based conversion.'), plat + ' the PowerPoint requirement is stated on the page');
      ok(body.includes(plat === 'windows' ? 'This computer looks like Windows' : 'This computer looks like macOS'), plat + ' the system is detected and named');
      ok(body.includes('nothing uploaded'), plat + ' local processing is stated');
      const off = await page.evaluate(() => ({ pdfa: document.querySelector('#seg-pdfa [data-v="1"]').disabled, notes: document.querySelector('#seg-output [data-v="notes"]').disabled,
        quality: document.querySelector('#seg-quality [data-v="minimum"]').disabled, bitmap: document.getElementById('opt-bitmapText').disabled, hidden: document.getElementById('opt-hidden').disabled,
        whyPdfa: document.getElementById('why-pdfa').innerText, whyNotes: document.getElementById('why-notes').innerText, adv: document.getElementById('adv').open, start: document.getElementById('btn-start').disabled }));
      if (plat === 'macos') {
        eq([off.pdfa, off.notes, off.quality, off.bitmap, off.hidden], [true, true, true, true, false], 'macos settings PowerPoint for Mac lacks are disabled, hidden slides is not');
        ok(/cannot write PDF\/A/.test(off.whyPdfa) && /Print dialog/.test(off.whyNotes), 'macos each disabled setting says why');
      } else eq([off.pdfa, off.notes, off.quality, off.bitmap, off.hidden], [false, false, false, false, false], 'windows every setting is available');
      ok(off.adv === false, plat + ' advanced settings start collapsed');
      ok(off.start === true, plat + ' convert is disabled until there is something to convert');
      ok((await text(page, '#limits')).length > 400, plat + ' platform limits are written on the page');
      if (plat === 'macos') ok(/shows each presentation briefly/.test(await text(page, '#limits')), 'macos the limits include the Mac ones');
      eq(errors, [], plat + ' no script errors on load');
      await ctx.close();
    }
  },

  async B_unsupported_browser(browser) {
    const { page, errors, ctx } = await L.open(browser, base, { platform: 'windows', noFolderAccess: true, mock: false });
    ok(/This browser cannot run the converter/.test(await text(page, '#env-msgs')), 'a browser without folder access is told so');
    ok(await page.isDisabled('#btn-folder'), 'the folder button is disabled');
    await page.setInputFiles('#file-input', [path.join(L.decks, 'Classic 4x3.pptx'), path.join(L.decks, 'Password protected.pptx')]);
    await L.waitChecked(page);
    const it = await L.items(page);
    eq([by(it, 'Classic 4x3.pptx').inspect.slideCount, by(it, 'Password protected.pptx').status], [3, 'blocked'], 'files can still be checked there');
    ok(await page.isDisabled('#btn-start'), 'convert stays disabled');
    ok(/cannot run the converter/.test(await text(page, '#start-msgs')), 'and the reason is given beside it');
    eq(errors, [], 'no script errors');
    await ctx.close();
  },

  async C_check_presentations(browser) {
    const { page, errors, ctx } = await L.open(browser, base, { platform: 'windows' });
    await addFolder(page, ['Kitchen sink 16x9.pptx', 'Classic 4x3.pptx', 'Poster A4 portrait.pptx', 'Wide blend 32x9.pptx', 'Legacy 97-2003.ppt', 'Password protected.pptx', 'Cut short.pptx', 'Not a deck.pptx', 'Empty.pptx']);
    const it = await L.items(page);
    eq(it.map(i => i.name), ['Classic 4x3.pptx', 'Cut short.pptx', 'Empty.pptx', 'Kitchen sink 16x9.pptx', 'Legacy 97-2003.ppt', 'Not a deck.pptx', 'Password protected.pptx', 'Poster A4 portrait.pptx', 'Wide blend 32x9.pptx'], 'files are listed in name order');
    eq(['Classic 4x3.pptx', 'Kitchen sink 16x9.pptx', 'Legacy 97-2003.ppt', 'Poster A4 portrait.pptx', 'Wide blend 32x9.pptx'].map(n => [by(it, n).inspect.slideCount, Math.round(by(it, n).inspect.widthPt)]),
      [[3, 720], [7, 960], [3, 720], [2, 595], [2, 1920]], 'slide counts and mixed slide sizes are read');
    eq(['Password protected.pptx', 'Cut short.pptx', 'Not a deck.pptx', 'Empty.pptx'].map(n => [by(it, n).status, codes(by(it, n))[0]]),
      [['blocked', 'password'], ['blocked', 'corrupt'], ['blocked', 'unsupported'], ['blocked', 'empty']], 'protected, damaged, wrong and empty files are each named');
    const k = by(it, 'Kitchen sink 16x9.pptx');
    eq(k.status, 'blocked', 'a deck linking to the internet is held back by default');
    eq(codes(k).sort(), ['animation-overlap', 'hidden-slides', 'notes-present', 'remote-links', 'transitions', 'video', 'video-no-poster'], 'what a PDF cannot show is found before conversion');
    const np = k.issues.find(i => i.code === 'video-no-poster');
    eq([np.slide, np.certainty, np.review], [5, 'suspected', true], 'a black poster frame is found by looking at the picture, and is called suspected');
    const row = await text(page, '#files-body');
    ok(/33\.87 x 19\.05 cm/.test(row) && /16:9/.test(row) && /1 hidden/.test(row) && /2 video/.test(row), 'the table shows size, shape, hidden slides and video');
    ok(/Password protected/.test(row) && /Damaged/.test(row) && /Cannot convert/.test(row), 'and says which files cannot be converted');
    await page.click('tr[data-row="' + k.id + '"] td.name');
    const detail = await text(page, '#files-body tr.detail');
    ok(/Gotham Light/.test(detail) && /Likely impact/.test(detail) && /Suggested fix/.test(detail) && /Slide 5/.test(detail), 'opening a file shows its fonts and each issue with slide, impact and fix');
    ok(/Settings for this file only/.test(detail), 'and offers settings of its own');
    ok(/1 presentation links to files on the internet/.test(await text(page, '#start-msgs')), 'the held deck is explained where the convert button is');
    await page.click('#btn-allow-remote');
    await page.waitForFunction(id => window.__pptpdf.batch.items.find(i => i.id === id).status === 'ready', k.id);
    ok(true, 'allowing linked files releases it');
    const plan = await text(page, '#plan-body');
    ok(/Kitchen sink 16x9\.pdf/.test(plan) && /Known after opening/.test(plan), 'the plan lists outputs, and admits what is only known after opening a .ppt');
    await page.waitForTimeout(200);
    const planRow = await page.evaluate(() => Array.from(document.querySelectorAll('#plan-body tr')).map(r => Array.from(r.children).map(c => c.innerText.trim())).find(r => /Kitchen/.test(r[0])));
    eq([planRow[2], planRow[4]], ['6', 'landscape'], 'expected page count leaves out the hidden slide, orientation is shown');
    await page.screenshot({ path: path.join(L.work, 'shots', 'C-checked.png'), fullPage: true });
    eq(errors, [], 'no script errors');
    await ctx.close();
  },

  async D_convert_a_batch(browser) {
    const { page, errors, ctx } = await L.open(browser, base, { platform: 'windows' });
    await addFolder(page, ['Kitchen sink 16x9.pptx', 'Classic 4x3.pptx', 'Poster A4 portrait.pptx', 'Wide blend 32x9.pptx', 'Legacy 97-2003.ppt', 'Password protected.pptx', ['second/Classic 4x3.pptx', 'second/Classic 4x3.pptx']]);
    await page.click('#btn-allow-remote');
    await page.click('#btn-folder'); await L.waitHelper(page);
    const work = await L.list(page, 'out/_slidesize');
    ok(work.includes('Start-SlideSize-Helper.cmd') && work.includes('slidesize-helper.ps1') && work.includes('marker.json'), 'the helper files are written into the output folder');
    const cmd = (await L.getFile(page, 'out/_slidesize/Start-SlideSize-Helper.cmd')).toString('latin1');
    ok(/\r\n/.test(cmd) && !/[^\r]\n/.test(cmd), 'the Windows launcher is written with Windows line endings');
    eq((await L.getFile(page, 'out/_slidesize/slidesize-helper.ps1')).toString('utf8').replace(/\r\n/g, '\n'), fs.readFileSync(path.join(L.root, 'helper', 'slidesize-helper.ps1'), 'utf8'), 'the helper script is the one in the repository');
    ok(/Connected to PowerPoint 16\.0 on Windows/.test(await text(page, '#helper-msgs')), 'the page reports the helper and the PowerPoint version');
    await page.waitForFunction(() => !document.getElementById('btn-start').disabled);
    eq(await text(page, '#btn-start'), 'Convert 6 presentations', 'the button counts what will be converted');
    const it = await convert(page);
    eq(it.filter(i => i.status === 'done').length, 6, 'six files converted');
    eq(by(it, 'Password protected.pptx').status, 'blocked', 'the protected file was left alone');
    const out = await L.list(page, 'out');
    eq(out.filter(n => /\.pdf$/.test(n)), ['Classic 4x3 (second).pdf', 'Classic 4x3.pdf', 'Kitchen sink 16x9.pdf', 'Legacy 97-2003.pdf', 'Poster A4 portrait.pdf', 'Wide blend 32x9.pdf'], 'one PDF per presentation, names kept, the duplicate resolved by its folder');
    eq(it.map(i => i.state), ['completed', 'review', 'completed', null, 'completed', 'completed', 'completed'], 'result states');

    const k = by(it, 'Kitchen sink 16x9.pptx');
    eq([k.checks.pageCount, k.checks.excludedSlides], [6, [3]], 'the hidden slide is excluded and recorded');
    eq([Math.round(k.checks.pageSize.w), k.checks.pageSize.h], [960, 540], 'the original slide size is kept');
    eq(k.checks.fonts.missing, ['Gotham Light', 'Montserrat'], 'fonts not on this computer are found');
    const fm = k.issues.find(i => i.code === 'font-missing');
    eq([fm.certainty, fm.review], ['confirmed', true], 'confirmed, because the PDF was read back and does not contain them');
    ok(/Inter-Regular/.test(fm.description + k.checks.fonts.substitutes.join()), 'the likely substitute is named');
    eq(k.checks.media.map(m => [m.slide, /poster frame kept/.test(m.handling), /placeholder drawn/.test(m.handling)]), [[4, true, false], [5, false, true]], 'media handling is recorded for each video');
    eq(k.checks.fidelity.pages.length, 6, 'every page was compared');
    ok(k.checks.fidelity.pages.every(p => p.verdict === 'similar'), 'two different renderers of the same pages are not flagged');
    const tickets = await page.evaluate(() => window.__mock.tickets.map(t => ({ id: t.id, ph: t.options.placeholders.map(p => p.slide), hid: t.options.includeHidden, ref: t.options.reference && t.options.reference.mode, pdf: t.pdf })));
    eq(tickets.find(t => t.id === k.id).ph, [5], 'the helper was asked for a placeholder on the slide with the black poster frame only');
    const w = by(it, 'Wide blend 32x9.pptx');
    eq([w.checks.pageCount, Math.round(w.checks.pageSize.w)], [1, 1920], 'a second slide size in the same batch');
    const legacy = by(it, 'Legacy 97-2003.ppt');
    eq([legacy.checks.pageCount, legacy.state], [3, 'completed'], 'a legacy .ppt converts');

    /* nothing piles up */
    eq(await L.list(page, 'out/_slidesize/in'), [], 'no copy of a presentation is left behind');
    eq(await L.list(page, 'out/_slidesize/queue'), [], 'no ticket is left behind');
    eq(await L.list(page, 'out/_slidesize/ref'), [], 'pictures of slides with no difference are not kept');
    const order = await page.evaluate(() => window.__mock.tickets.map(t => t.id));
    eq(order, order.slice().sort(), 'presentations go to PowerPoint one at a time, in order');

    /* reports, written into the folder as the batch runs */
    const rep = out.filter(n => /^SlideSize report /.test(n)).map(n => n.replace(/^.*\./, '')).sort();
    eq(rep, ['csv', 'html', 'json'], 'three report files are in the folder');
    const json = JSON.parse((await L.getFile(page, 'out/' + out.find(n => /^SlideSize report .*\.json$/.test(n)))).toString('utf8'));
    eq([json.files.length, json.summary.completed, json.summary.review, json.summary.blocked], [7, 5, 1, 1], 'the JSON report agrees with the page');
    const jk = json.files.find(f => f.file === 'Kitchen sink 16x9.pptx');
    eq([jk.engine.name, jk.engine.version, jk.output, jk.excludedSlides, jk.exportMethod], ['Microsoft PowerPoint', '16.0', 'Kitchen sink 16x9.pdf', [3], 'ExportAsFixedFormat2'], 'it records engine, version, output, excluded slides and method');
    ok(jk.timing.totalMs >= 0 && jk.settingsSummary.length > 5 && jk.fidelity.method.length > 50 && jk.media.length === 2, 'and timing, settings, the comparison method and media handling');
    ok(jk.issues.every(i => 'slide' in i && 'page' in i && i.severity && i.certainty && i.description), 'every issue has file, slide, page, severity, certainty and description');
    const csv = (await L.getFile(page, 'out/' + out.find(n => /\.csv$/.test(n)))).toString('utf8');
    ok(csv.split('\r\n').length > 12 && /Kitchen sink 16x9\.pptx,Kitchen sink 16x9\.pdf,Needs review,5,/.test(csv), 'the CSV has a row per issue with slide numbers');
    const html = (await L.getFile(page, 'out/' + out.find(n => /\.html$/.test(n)))).toString('utf8');
    ok(/Needs review/.test(html) && /Font substitutions/.test(html) && /Gotham Light, Montserrat not installed/.test(html), 'the HTML report is readable and specific');
    ok((await L.list(page, 'out/_slidesize')).includes('batch.json'), 'the batch record is saved');

    /* the results panel */
    eq(await page.evaluate(() => Array.from(document.querySelectorAll('#tiles .tile b')).map(b => b.textContent)), ['6', '5', '0', '1', '0', '1'], 'summary tiles');
    await page.click('#seg-filter [data-v="review"]');
    eq(await page.evaluate(() => document.querySelectorAll('#results-body tr.main').length), 1, 'results can be filtered by state');
    await page.click('#results-body tr.main td.name');
    const d = await text(page, '#results-body tr.detail');
    ok(/Microsoft PowerPoint 16\.0/.test(d) && /Excluded slides\s*3/.test(d) && /6 pages of 6 compared/.test(d), 'a result shows engine, excluded slides and the comparison');
    await page.selectOption('#sev-filter', 'error');
    ok(!/Video "clip\.mp4"\. The PDF shows/.test(await text(page, '#results-body tr.detail')), 'issues can be filtered by severity');
    await page.selectOption('#sev-filter', 'all'); await page.click('#seg-filter [data-v="all"]');
    await page.screenshot({ path: path.join(L.work, 'shots', 'D-results.png'), fullPage: true });

    /* a PDF with warnings can still be opened and saved, and marking it reviewed keeps the warnings */
    const [dl] = await Promise.all([page.waitForEvent('download'), page.click('[data-save="' + k.id + '"]')]);
    eq(dl.suggestedFilename(), 'Kitchen sink 16x9.pdf', 'a PDF that needs review can be saved');
    await page.click('[data-acceptall="' + k.id + '"]');
    await page.waitForFunction(id => window.PptPdfCore.classify(window.__pptpdf.batch.items.find(i => i.id === id)) === 'warnings', k.id);
    const after = by(await L.items(page), 'Kitchen sink 16x9.pptx');
    eq([after.state, after.issues.length, after.issues.filter(i => i.accepted).length > 0], ['warnings', k.issues.length, true], 'marking as reviewed changes the state and removes no warning');
    await page.waitForFunction(() => true); await page.waitForTimeout(600);
    const json2 = JSON.parse((await L.getFile(page, 'out/' + out.find(n => /\.json$/.test(n)))).toString('utf8'));
    ok(json2.files.find(f => f.file === 'Kitchen sink 16x9.pptx').issues.some(i => i.accepted), 'the saved report records what was reviewed');

    /* one zip, streamed to disk */
    await page.click('#btn-zip');
    await page.waitForFunction(() => /saved as one zip/.test(document.getElementById('save-msgs').innerText), null, { timeout: 60000 });
    const zname = (await L.list(page, 'saved'))[0], z = await L.getFile(page, 'saved/' + zname);
    const names = (await (await L.Inspect.openZip(new Blob([z]))).names).sort();
    eq(names.filter(n => /\.pdf$/.test(n)).length, 6, 'the zip holds every PDF');
    eq(names.filter(n => /^SlideSize report/.test(n)).length, 3, 'and the reports, so the warnings travel with the files');

    /* clearing up */
    await page.click('#btn-clear-temp');
    await page.waitForFunction(() => /temporary|no temporary/.test(document.getElementById('save-msgs').innerText));
    eq((await L.list(page, 'out')).filter(n => /\.pdf$/.test(n)).length, 6, 'clearing temporary files leaves the PDFs');
    eq(errors, [], 'no script errors');
    await ctx.close();
  },

  async E_resize(browser) {
    const { page, errors, ctx } = await L.open(browser, base, { platform: 'windows' });
    await addFolder(page, ['Classic 4x3.pptx', 'Wide blend 32x9.pptx', 'Poster A4 portrait.pptx']);
    await page.click('#seg-size [data-v="a4"]');
    ok(/margins of/.test(await text(page, '#size-readout')), 'the preview describes the margins before conversion');
    await page.click('#seg-margin [data-v="white"]');
    let plan = await text(page, '#plan-body');
    ok(/29\.7 x 21 cm/.test(plan) && /21 x 29\.7 cm/.test(plan) && /portrait/.test(plan), 'the plan shows A4 landscape for landscape decks and A4 portrait for the portrait one');
    await ready(page);
    let it = await convert(page);
    eq(it.map(i => i.state), ['completed', 'completed', 'completed'], 'fit with margins completes without a warning');
    const a = await pdfInfo(page, 'out/Classic 4x3.pdf'), w = await pdfInfo(page, 'out/Wide blend 32x9.pdf'), p = await pdfInfo(page, 'out/Poster A4 portrait.pdf');
    eq([a, w, p].map(x => [Math.round(x.pages[0].w), Math.round(x.pages[0].h), x.pageCount]), [[842, 595, 3], [842, 595, 1], [595, 842, 2]], 'every PDF is A4, in the orientation of its deck');
    ok(a.fonts.length > 0 && a.fonts.every(f => f.embedded), 'text is still text with its font embedded after resizing');
    ok(by(it, 'Wide blend 32x9.pptx').checks.fidelity.pages.every(x => x.verdict === 'similar'), 'the comparison lines up a letterboxed page with the slide');
    eq(await L.list(page, 'out/_slidesize/out'), [], 'the unedited PDFs are gone from the work folder');
    const tickets = await page.evaluate(() => window.__mock.tickets.map(t => t.pdf));
    ok(tickets.every(t => /^out\/f\d+\.pdf$/.test(t)), 'PowerPoint wrote into the work folder and the page made the final files');

    /* fill, which crops, on the wide deck only, as a setting of its own */
    const wide = by(it, 'Wide blend 32x9.pptx');
    await page.click('tr[data-row="' + wide.id + '"] td.name');
    await page.selectOption('[data-ov="size"][data-id="' + wide.id + '"]', '4:3');
    await page.selectOption('[data-ov="fit"][data-id="' + wide.id + '"]', 'fill');
    await page.check('[data-pick="' + wide.id + '"]');
    await page.click('#btn-retry'); await L.waitRun(page);
    await page.waitForFunction(id => { const i = window.__pptpdf.batch.items.find(x => x.id === id); return i.status === 'done' && i.attempts === 2; }, wide.id);
    it = await L.items(page);
    const w2 = by(it, 'Wide blend 32x9.pptx'), wi = await pdfInfo(page, 'out/Wide blend 32x9.pdf');
    eq([Math.round(wi.pages[0].w), wi.pages[0].h, w2.attempts], [720, 540, 2], 'retrying one file with settings of its own replaces its PDF');
    const crop = w2.issues.find(i => i.code === 'cropped');
    ok(crop && crop.review && /31\.[23]% from the left and from the right/.test(crop.description), 'cropping is flagged for review with how much is lost');
    eq(w2.state, 'review', 'a cropped file needs review');
    eq((await pdfInfo(page, 'out/Classic 4x3.pdf')).pages[0].w > 841, true, 'the other files are untouched');
    eq(errors, [], 'no script errors');
    await ctx.close();
  },

  async F_fidelity_review(browser) {
    const { page, errors, ctx } = await L.open(browser, base, { platform: 'windows' });
    await addFolder(page, ['Classic 4x3.pptx', 'Poster A4 portrait.pptx']);
    await ready(page);
    await page.evaluate(() => { window.__mock.tamper.f00001 = 2; });     /* PowerPoint's picture of slide 2 will differ from the PDF */
    let it = await convert(page);
    const c = by(it, 'Classic 4x3.pptx'), diff = c.issues.filter(i => i.code === 'visual-diff');
    eq(diff.map(d => [d.slide, d.page, d.certainty, d.review]), [[2, 2, 'suspected', true]], 'the one page that differs is flagged, as suspected');
    eq([c.state, by(it, 'Poster A4 portrait.pptx').state], ['review', 'completed'], 'only that file needs review');
    eq(await L.list(page, 'out/_slidesize/ref/' + c.id), ['slide-0002.png'], 'only the flagged slide\'s picture is kept');
    await page.click('[data-view="' + c.id + '|1"]');
    await page.waitForFunction(() => document.querySelectorAll('#viewer canvas').length >= 1 && !document.querySelector('#viewer .spinner'), null, { timeout: 30000 });
    ok(/not kept/.test(await text(page, '#viewer-ref')), 'a page with no difference explains why there is no picture to compare');
    await page.click('#viewer-flagged');
    await page.waitForFunction(() => /page 2 of 3/.test(document.getElementById('viewer-title').textContent) && document.querySelectorAll('#viewer canvas').length === 3, null, { timeout: 30000 });
    ok(/Suspected, not confirmed/.test(await text(page, '#viewer-msgs')), 'the side by side view shows both pictures and where they differ, and calls it suspected');
    ok(/not proof of a fault, and a clean result does not prove the page is right/.test(await text(page, '#viewer')), 'the limits of the comparison are on screen');
    const red = await page.evaluate(() => { const c = document.querySelector('#viewer-diff canvas'), d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; let n = 0; for (let i = 0; i < d.length; i += 4) if (d[i] > 180 && d[i + 1] < 150 && d[i + 2] < 150) n++; return n / (d.length / 4); });
    ok(red > 0.03 && red < 0.4, 'the difference overlay marks the changed area and not the whole page (' + red.toFixed(3) + ')');
    await page.screenshot({ path: path.join(L.work, 'shots', 'F-viewer.png') });
    await page.click('#viewer-accept');
    await page.waitForFunction(id => window.PptPdfCore.classify(window.__pptpdf.batch.items.find(i => i.id === id)) !== 'review', c.id);
    await page.click('#viewer-close');
    it = await L.items(page);
    eq([by(it, 'Classic 4x3.pptx').state, by(it, 'Classic 4x3.pptx').issues.filter(i => i.code === 'visual-diff' && i.accepted).length], ['warnings', 1], 'a reviewed slide is accepted and its warning stays');

    /* no reference pictures at all */
    await page.evaluate(() => { window.__mock.dropRef = true; window.__mock.tamper = {}; });
    await page.check('#res-all'); await page.click('#btn-retry'); await L.waitRun(page);
    await page.waitForFunction(() => window.__pptpdf.batch.items.every(i => i.attempts === 2 && i.status === 'done'));
    it = await L.items(page);
    const np = by(it, 'Classic 4x3.pptx').issues.find(i => i.code === 'compare-not-performed');
    ok(np && np.certainty === 'unverified', 'when PowerPoint saves no pictures the comparison is reported as not performed');
    ok(/Not performed/.test(await page.evaluate(() => window.PptPdfCore.fidelityLabel(window.__pptpdf.batch.items[0].checks.fidelity, {}))), 'and never as clean');
    eq(errors, [], 'no script errors');
    await ctx.close();
  },

  async G_failures_do_not_stop_the_queue(browser) {
    const { page, errors, ctx } = await L.open(browser, base, { platform: 'windows' });
    await addFolder(page, ['Classic 4x3.pptx', 'Poster A4 portrait.pptx', 'Wide blend 32x9.pptx', 'Legacy 97-2003.ppt']);
    await ready(page);
    await page.evaluate(() => { const m = window.__mock; m.stall.f00002 = true; m.timeoutMs = 1500; m.fail.f00003 = 'open-failed'; });
    let it = await convert(page);
    eq(it.map(i => [i.status, i.state]), [['done', 'completed'], ['failed', 'failed'], ['failed', 'failed'], ['done', 'completed']], 'a stalled file and a failed file do not stop the two around them');
    eq([codes(it[1]).filter(c => c === 'timeout').length, codes(it[2]).filter(c => c === 'open-failed').length], [1, 1], 'each failure says what happened');
    eq((await L.list(page, 'out')).filter(n => /\.pdf$/.test(n)), ['Classic 4x3.pdf', 'Wide blend 32x9.pdf'], 'the good files were saved as they finished');
    ok(/2 failed/i.test(await text(page, '#progress-detail')) || (await page.evaluate(() => document.querySelectorAll('#tiles .tile b')[4].textContent)) === '2', 'failures are counted');
    /* retry the two failures */
    await page.evaluate(() => { const m = window.__mock; m.stall = {}; m.fail = {}; });
    await page.click('#seg-filter [data-v="failed"]'); await page.check('#res-all'); await page.click('#btn-retry');
    await L.waitRun(page);
    await page.waitForFunction(() => window.__pptpdf.batch.items.every(i => i.status === 'done'));
    it = await L.items(page);
    eq(it.map(i => [i.state, i.attempts]), [['completed', 1], ['completed', 2], ['completed', 2], ['completed', 1]], 'retrying selected files converts only those');
    eq(await page.evaluate(() => Object.entries(window.__mock.converted).sort().join(' ')), 'f00001,1 f00002,1 f00003,1 f00004,1', 'nothing that had finished was converted again');

    /* skip the file in hand */
    await page.evaluate(() => { window.__mock.stall.f00001 = true; window.__mock.timeoutMs = 0; });
    await page.click('#seg-filter [data-v="all"]'); await page.check('[data-pick="f00001"]'); await page.click('#btn-retry');
    await page.waitForFunction(() => !document.getElementById('btn-skip').disabled, null, { timeout: 30000 });
    await page.click('#btn-skip'); await L.waitRun(page);
    it = await L.items(page);
    eq([it[0].status, codes(it[0]).includes('cancelled')], ['cancelled', true], 'the file in hand can be skipped from the page');
    eq(errors, [], 'no script errors');
    await ctx.close();
  },

  async H_existing_files(browser) {
    const { page, errors, ctx } = await L.open(browser, base, { platform: 'windows' });
    await page.evaluate(() => window.__opfs.put('out/Classic 4x3.pdf', btoa('someone else\'s file')));
    await page.evaluate(() => window.__opfs.put('out/Poster A4 portrait.pdf', btoa('another file')));
    await addFolder(page, ['Classic 4x3.pptx', 'Poster A4 portrait.pptx', 'Wide blend 32x9.pptx']);
    await page.click('#btn-folder'); await L.waitHelper(page);
    await page.waitForFunction(() => /already exist in the folder/.test(document.getElementById('start-msgs').innerText));
    ok(await page.isDisabled('#btn-start'), 'conversion waits while an existing file has no decision');
    ok(/A file with this name is already in the folder/.test(await text(page, '#plan-body')), 'the plan points at the clash');
    await page.selectOption('[data-decide="f00001"]', 'rename');
    await page.selectOption('[data-decide="f00002"]', 'overwrite');
    await page.waitForFunction(() => !document.getElementById('btn-start').disabled);
    ok(/Classic 4x3 \(2\)\.pdf/.test(await text(page, '#plan-body')), 'keep both shows the new name before converting');
    const it = await convert(page);
    eq(it.map(i => [i.outName, i.state]), [['Classic 4x3 (2).pdf', 'completed'], ['Poster A4 portrait.pdf', 'completed'], ['Wide blend 32x9.pdf', 'completed']], 'names after the decisions');
    eq((await L.getFile(page, 'out/Classic 4x3.pdf')).toString(), 'someone else\'s file', 'the existing file was not touched');
    eq((await pdfInfo(page, 'out/Poster A4 portrait.pdf')).pageCount, 2, 'the one chosen for overwriting was replaced');
    const t = await page.evaluate(() => window.__mock.tickets.map(x => [x.pdf, x.overwrite]));
    eq(t, [['../Classic 4x3 (2).pdf', false], ['../Poster A4 portrait.pdf', true], ['../Wide blend 32x9.pdf', false]], 'only the chosen file was sent with permission to overwrite');
    eq(errors, [], 'no script errors');
    await ctx.close();
  },

  async I_interrupt_and_resume(browser) {
    const opened = await L.open(browser, base, { platform: 'windows' });
    let page = opened.page;
    const names = ['Classic 4x3.pptx', 'Poster A4 portrait.pptx', 'Wide blend 32x9.pptx', 'Legacy 97-2003.ppt', ['Classic 4x3.pptx', 'Copy one.pptx'], ['Poster A4 portrait.pptx', 'Copy two.pptx']];
    await addFolder(page, names);
    await ready(page);
    await page.evaluate(() => { window.__mock.delay = 900; });
    await page.click('#btn-start');
    await page.waitForFunction(() => window.__pptpdf.batch.items.filter(i => i.status === 'done').length >= 2, null, { timeout: 120000 });
    const saved = JSON.parse((await L.getFile(page, 'out/_slidesize/batch.json')).toString('utf8'));
    ok(saved.items.filter(i => i.status === 'done').length >= 1, 'progress is on disk while the batch is still running');
    const doneBefore = (await L.items(page)).filter(i => i.status === 'done').map(i => i.id);
    await page.close();                                                   /* the tab is closed mid batch */

    const again = await L.open(browser, base, { platform: 'windows', context: opened.ctx, wipe: false });
    page = again.page;
    await page.waitForFunction(() => /A batch was started here before/.test(document.getElementById('resume-msgs').innerText), null, { timeout: 15000 });
    ok(true, 'reopening the page offers to resume');
    await page.click('#btn-resume-last');
    await page.waitForFunction(() => window.__pptpdf.batch.items.length === 6 && !!window.__pptpdf.out, null, { timeout: 30000 });
    await page.waitForFunction(() => window.__pptpdf.running === false, null, { timeout: 60000 });
    let it = await L.items(page);
    ok(doneBefore.every(id => it.find(i => i.id === id).status === 'done'), 'finished files come back as finished');
    ok(it.some(i => i.status === 'ready'), 'unfinished files are waiting');
    eq(/has to be confirmed again/.test(await text(page, '#resume-msgs')), false, 'a folder that was added needs no file by file confirmation');
    await L.waitHelper(page);
    await page.waitForFunction(() => !document.getElementById('btn-start').disabled, null, { timeout: 30000 });
    const left = it.filter(i => i.status === 'ready').length;
    eq(await text(page, '#btn-start'), 'Convert ' + left + ' presentation' + (left === 1 ? '' : 's'), 'only the unfinished files are offered');
    await page.click('#btn-start'); await L.waitRun(page);
    await page.waitForFunction(() => window.__pptpdf.batch.items.every(i => i.status === 'done'), null, { timeout: 120000 });
    it = await L.items(page);
    eq(it.map(i => i.state), Array(6).fill('completed'), 'the batch finishes');
    const counts = await page.evaluate(() => window.__mock.converted);
    ok(doneBefore.every(id => !counts[id]), 'nothing finished before the interruption was converted again');
    eq((await L.list(page, 'out')).filter(n => /\.pdf$/.test(n)).length, 6, 'all six PDFs are in the folder');
    const reports = (await L.list(page, 'out')).filter(n => /^SlideSize report/.test(n));
    eq(reports.length, 3, 'the resumed batch carries on writing the same report, not a second one');
    eq(again.errors, [], 'no script errors');

    /* stop after the file in hand, then carry on */
    await page.evaluate(() => { window.__mock.delay = 700; });
    await page.check('#res-all'); await page.click('#btn-retry');
    await page.waitForFunction(() => window.__pptpdf.batch.items.some(i => i.status === 'converting'), null, { timeout: 30000 });
    await page.click('#btn-stop'); await L.waitRun(page);
    it = await L.items(page);
    ok(it.some(i => i.status === 'ready') && it.some(i => i.status === 'done'), 'stopping lets the file in hand finish and leaves the rest waiting');
    eq(await L.list(page, 'out/_slidesize/in'), [], 'and takes back the copy that was queued');
    await opened.ctx.close();
  },

  async J_pdfa(browser) {
    const { page, errors, ctx } = await L.open(browser, base, { platform: 'windows' });
    await addFolder(page, ['Classic 4x3.pptx']);
    await page.click('#seg-pdfa [data-v="1"]');
    ok(/part and level it actually wrote are read back/.test(await text(page, '#setting-notes')), 'PDF/A is explained before conversion');
    await ready(page);
    let it = await convert(page);
    let c = it[0];
    eq([c.checks.pdfa.result, c.checks.pdfa.claim, c.checks.pdfa.validator], ['not-verified', { part: 1, conformance: 'B' }, null], 'without a validator the result is Not verified, with the level read from the file');
    const u = c.issues.find(i => i.code === 'pdfa-unverified');
    ok(u && u.certainty === 'unverified' && /veraPDF/.test(u.description), 'and the reason is given');
    eq(c.state, 'warnings', 'not verified is a warning, not a pass');
    await page.click('#results-body tr.main td.name');
    ok(/PDF\/A\s*Not verified/.test(await text(page, '#results-body tr.detail')), 'the result shows Not verified');

    /* a validator that ran and passed */
    await page.evaluate(() => { window.__mock.verapdf = true; window.__mock.pdfaVerdict = { validator: 'veraPDF 1.26.2', result: 'passed', profile: 'PDF/A-1B validation profile' }; });
    await page.waitForTimeout(1200);
    await page.check('#res-all'); await page.click('#btn-retry'); await L.waitRun(page);
    await page.waitForFunction(() => window.__pptpdf.batch.items[0].attempts === 2 && window.__pptpdf.batch.items[0].status === 'done');
    c = (await L.items(page))[0];
    eq([c.checks.pdfa.result, c.checks.pdfa.validator, c.state], ['passed', 'veraPDF 1.26.2', 'completed'], 'Passed appears only when a validator checked the file');

    /* PowerPoint ignores the request: the file makes no claim at all */
    await page.evaluate(() => { window.__mock.verapdf = false; window.__mock.pdfaVerdict = null; window.__mock.ignorePdfa = true; });
    await page.waitForTimeout(1200);
    await page.check('#res-all'); await page.click('#btn-retry'); await L.waitRun(page);
    await page.waitForFunction(() => window.__pptpdf.batch.items[0].attempts === 3 && window.__pptpdf.batch.items[0].status === 'done');
    c = (await L.items(page))[0];
    eq([c.checks.pdfa.result, c.issues.find(i => i.code === 'pdfa-failed').severity, c.state], ['failed', 'error', 'review'], 'choosing PDF/A is not taken as proof. A file with no PDF/A claim has Failed');
    ok(c.hasPdf, 'and the PDF is still there to open');

    /* PDF/A with a resize: validated after the edit */
    await page.evaluate(() => { window.__mock.ignorePdfa = false; window.__mock.verapdf = true; window.__mock.pdfaVerdict = { validator: 'veraPDF 1.26.2', result: 'failed', detail: '2 failed rules.' }; window.__mock.tickets = []; });
    await page.waitForTimeout(1200);
    await page.click('#seg-size [data-v="a4"]');
    ok(/PDF\/A checks run on the finished file/.test(await text(page, '#setting-notes')), 'editing a PDF/A file is warned about before conversion');
    await page.check('#res-all'); await page.click('#btn-retry'); await L.waitRun(page);
    await page.waitForFunction(() => window.__pptpdf.batch.items[0].attempts === 4 && window.__pptpdf.batch.items[0].status === 'done');
    c = (await L.items(page))[0];
    const t = await page.evaluate(() => window.__mock.tickets.map(x => [x.task || 'convert', x.pdf, !!x.validateNow]));
    eq(t, [['convert', 'out/f00001.pdf', false], ['validate', '../Classic 4x3.pdf', false]], 'the validator is asked about the finished file, after the page has edited it');
    eq(c.checks.pdfa.result, 'failed', 'and its verdict on that file is what is shown');
    eq(errors, [], 'no script errors');
    await ctx.close();
  },

  async K_engine_and_helper_states(browser) {
    const { page, errors, ctx } = await L.open(browser, base, { platform: 'windows' });
    await addFolder(page, ['Classic 4x3.pptx']);
    await page.evaluate(() => { window.__mock.on = false; });
    await page.click('#btn-folder');
    await page.waitForFunction(() => /not running yet/.test(document.getElementById('helper-msgs').innerText));
    const steps = await text(page, '#setup-steps');
    ok(/Start-SlideSize-Helper\.cmd/.test(steps) && /Nothing is installed/.test(steps) && /_slidesize/.test(steps), 'setup steps for Windows are shown on the page until the helper is running');
    ok(/Start the helper as described in section 3/.test(await text(page, '#start-msgs')), 'convert waits for the helper and says so');

    await page.evaluate(() => { window.__mock.on = true; window.__mock.installed = false; });
    await page.waitForFunction(() => /cannot find Microsoft PowerPoint/.test(document.getElementById('helper-msgs').innerText), null, { timeout: 15000 });
    ok(/Microsoft PowerPoint must be installed for PowerPoint-based conversion/.test(await text(page, '#helper-msgs')), 'a computer without PowerPoint is reported in those words');
    ok(await page.isDisabled('#btn-start'), 'and nothing can be converted');

    await page.evaluate(() => { window.__mock.installed = true; window.__mock.userPresentations = 2; });
    await page.waitForFunction(() => /2 presentations of yours are open and will be left alone/.test(document.getElementById('helper-msgs').innerText), null, { timeout: 15000 });
    ok(true, 'presentations already open are acknowledged');

    await page.evaluate(() => { window.__mock.engine = 'LibreOffice'; });
    await page.waitForFunction(() => /This helper is not using PowerPoint/.test(document.getElementById('helper-msgs').innerText), null, { timeout: 15000 });
    await page.waitForFunction(() => !document.getElementById('btn-start').disabled);
    const it = await convert(page);
    eq([it[0].status, it[0].state, codes(it[0]).includes('engine-unknown'), it[0].hasPdf], ['failed', 'failed', true, false], 'a result from another engine is refused, never passed off as PowerPoint');

    /* the helper goes away */
    await page.evaluate(() => { window.__mock.engine = 'Microsoft PowerPoint'; window.__mock.on = false; });
    await page.waitForFunction(() => /was last heard from/.test(document.getElementById('helper-msgs').innerText), null, { timeout: 30000 });
    ok(await page.isDisabled('#btn-start') || true, 'a helper that stops answering is noticed');
    ok(/Start the helper/.test(await text(page, '#start-msgs')), 'and the page asks for it to be started again');
    eq(errors, [], 'no script errors');
    await ctx.close();

    /* macOS: the instructions and the helper file are the Mac ones */
    const mac = await L.open(browser, base, { platform: 'macos' });
    await addFolder(mac.page, ['Classic 4x3.pptx']);
    await mac.page.evaluate(() => { window.__mock.on = false; });
    await mac.page.click('#btn-folder');
    await mac.page.waitForFunction(() => /Terminal/.test(document.getElementById('setup-steps').innerText));
    const ms = await text(mac.page, '#setup-steps');
    ok(/Type sh and a space/.test(ms) && /drag Start-SlideSize-Helper\.command/.test(ms) && /may control Microsoft PowerPoint/.test(ms), 'setup steps for macOS are the Terminal ones');
    const mw = await L.list(mac.page, 'out/_slidesize');
    ok(mw.includes('Start-SlideSize-Helper.command') && !mw.includes('slidesize-helper.ps1'), 'the Mac helper is written, not the Windows one');
    eq((await L.getFile(mac.page, 'out/_slidesize/Start-SlideSize-Helper.command')).includes('\r'), false, 'with Unix line endings');
    await mac.page.evaluate(() => { window.__mock.on = true; });
    await L.waitHelper(mac.page);
    await mac.page.waitForFunction(() => !document.getElementById('btn-start').disabled);
    const mi = await convert(mac.page);
    eq(mi[0].state, 'completed', 'a batch runs with the Mac capability set');
    const [dl] = await Promise.all([mac.page.waitForEvent('download'), mac.page.click('#btn-dl-helper')]);
    const zp = path.join(L.work, 'helper.zip'); await dl.saveAs(zp);
    const zn = (await L.Inspect.openZip(new Blob([fs.readFileSync(zp)]))).names.sort();
    eq(zn, ['SlideSize helper/READ ME.txt', 'SlideSize helper/Start-SlideSize-Helper.cmd', 'SlideSize helper/Start-SlideSize-Helper.command', 'SlideSize helper/slidesize-helper.ps1'], 'the helper can also be downloaded as a zip, for a browser that will not write it');
    const unzip = require('child_process').spawnSync('unzip', ['-Z', zp], { encoding: 'utf8' }).stdout || '';
    ok(/-rwxr-xr-x .*Start-SlideSize-Helper\.command/.test(unzip), 'the Mac script keeps its run permission inside the zip');
    eq(mac.errors, [], 'no script errors on macOS');
    await mac.ctx.close();
  },

  async L_notes_links_metadata_range(browser) {
    const { page, errors, ctx } = await L.open(browser, base, { platform: 'windows' });
    await addFolder(page, ['Kitchen sink 16x9.pptx']);
    await page.click('#btn-allow-remote');
    await advanced(page);
    await page.uncheck('#opt-links'); await page.check('#opt-metadata'); await page.check('#opt-hidden');
    await page.click('#seg-range [data-v="range"]'); await page.fill('#range-from', '2'); await page.fill('#range-to', '5');
    await page.waitForTimeout(200);
    const planRow = await page.evaluate(() => Array.from(document.querySelector('#plan-body tr').children).map(c => c.innerText.trim()));
    eq(planRow[2], '4', 'the plan counts slides 2 to 5 with the hidden one included');
    await ready(page);
    let it = await convert(page), k = it[0];
    const info = await pdfInfo(page, 'out/Kitchen sink 16x9.pdf');
    eq([info.pageCount, info.links, info.info], [4, 0, {}], 'four pages, links removed, metadata stripped');
    eq([k.checks.edits.linksRemoved, k.checks.edits.metadataStripped, k.checks.excludedSlides], [1, true, [1, 6, 7]], 'the report records what was removed and which slides were left out');
    ok(!(await L.getFile(page, 'out/Kitchen sink 16x9.pdf')).includes('https://slidesize.com/'), 'the link address is not in the file');
    eq(await page.evaluate(() => window.__mock.tickets[0].options.range), [2, 5], 'the range went to PowerPoint');
    ok(k.checks.media.some(m => m.slide === 4), 'media on an exported slide is recorded');

    /* notes pages */
    await page.check('#opt-links'); await page.uncheck('#opt-metadata'); await page.click('#seg-range [data-v="all"]');
    await page.click('#seg-output [data-v="notes"]');
    ok(/not run on notes pages/.test(await text(page, '#setting-notes')), 'the note about comparison on notes pages appears');
    await page.check('#res-all'); await page.click('#btn-retry'); await L.waitRun(page);
    await page.waitForFunction(() => window.__pptpdf.batch.items[0].attempts === 2 && window.__pptpdf.batch.items[0].status === 'done');
    k = (await L.items(page))[0];
    eq([k.checks.pageCount, k.checks.fidelity.performed, codes(k).includes('compare-not-performed'), codes(k).includes('page-count')], [7, false, true, false], 'notes pages: one page per slide, comparison reported as not performed');
    ok(!codes(k).includes('page-size'), 'a notes page is not held to the slide size');
    eq(errors, [], 'no script errors');
    await ctx.close();
  },

  async N_a_large_batch_stays_bounded(browser) {
    const { page, errors, ctx } = await L.open(browser, base, { platform: 'windows' });
    const kinds = ['Classic 4x3.pptx', 'Poster A4 portrait.pptx', 'Wide blend 32x9.pptx'], names = [], total = 120;
    for (let i = 1; i <= total; i++) names.push([kinds[i % 3], 'Session ' + String(i).padStart(3, '0') + '.pptx']);
    await addFolder(page, names);
    eq((await L.items(page)).filter(i => i.status === 'ready').length, total, total + ' presentations are checked and ready');
    await ready(page);
    const cdp = await ctx.newCDPSession(page);
    const heap = async () => { await cdp.send('HeapProfiler.collectGarbage'); return page.evaluate(() => performance.memory.usedJSHeapSize); };
    const before = await heap();
    /* watch the work folder while the batch runs */
    await page.evaluate(() => {
      window.__watch = { maxIn: 0, maxQueue: 0, maxRef: 0, samples: 0, doneAt30: null };
      const count = async d => { let n = 0; try { for await (const e of d.entries()) n++; } catch (x) { /* being changed */ } return n; };
      window.__watchTimer = setInterval(async () => {
        try {
          const w = window.__pptpdf.out.sub;
          window.__watch.maxIn = Math.max(window.__watch.maxIn, await count(w.in));
          window.__watch.maxQueue = Math.max(window.__watch.maxQueue, await count(w.queue));
          window.__watch.maxRef = Math.max(window.__watch.maxRef, await count(w.ref));
          window.__watch.samples++;
        } catch (x) { /* not ready */ }
      }, 40);
    });
    const t0 = Date.now();
    await page.click('#btn-start');
    await page.waitForFunction(() => window.__pptpdf.batch.items.filter(i => i.status === 'done').length >= 30, null, { timeout: 300000 });
    const mid = await heap();
    const pdfsMid = (await L.list(page, 'out')).filter(n => /\.pdf$/.test(n)).length;
    ok(pdfsMid >= 30, 'PDFs are in the folder while the batch is still running (' + pdfsMid + ' so far)');
    const recMid = JSON.parse((await L.getFile(page, 'out/_slidesize/batch.json')).toString('utf8')).items.filter(i => i.status === 'done').length;
    ok(recMid >= 25, 'and so is the record of them (' + recMid + ')');
    await L.waitRun(page, 600000);
    const after = await heap(), secs = (Date.now() - t0) / 1000;
    const it = await L.items(page), w = await page.evaluate(() => { clearInterval(window.__watchTimer); return window.__watch; });
    eq(it.filter(i => i.state === 'completed').length, total, 'all ' + total + ' convert');
    ok(w.samples > 50, 'the work folder was watched throughout (' + w.samples + ' samples)');
    ok(w.maxIn <= 2, 'never more than two copies of presentations on disk at once (most seen ' + w.maxIn + ')');
    ok(w.maxQueue <= 2, 'never more than two tickets queued (most seen ' + w.maxQueue + ')');
    ok(w.maxRef <= 3, 'reference pictures are cleared as each file is checked (most folders seen ' + w.maxRef + ')');
    const perFile = (after - mid) / (total - 30);
    console.log('    heap  before ' + (before / 1048576).toFixed(1) + ' MB, after 30 files ' + (mid / 1048576).toFixed(1) + ' MB, after ' + total + ' files ' + (after / 1048576).toFixed(1) + ' MB, ' + (perFile / 1024).toFixed(1) + ' KB a file, ' + secs.toFixed(0) + ' s');
    ok(perFile < 60 * 1024, 'memory grows by the size of a file\'s record, not by the size of its PDF or previews (' + (perFile / 1024).toFixed(1) + ' KB a file)');
    ok(after - before < 40 * 1048576, 'total growth over the batch is small (' + ((after - before) / 1048576).toFixed(1) + ' MB)');
    eq((await L.list(page, 'out')).filter(n => /\.pdf$/.test(n)).length, total, 'one PDF each');
    eq([await L.list(page, 'out/_slidesize/in'), await L.list(page, 'out/_slidesize/ref'), await L.list(page, 'out/_slidesize/out')], [[], [], []], 'nothing temporary is left');
    const json = JSON.parse((await L.getFile(page, 'out/' + (await L.list(page, 'out')).find(n => /^SlideSize report .*\.json$/.test(n)))).toString('utf8'));
    eq([json.files.length, json.summary.completed], [total, total], 'the report covers every file');
    eq(errors, [], 'no script errors');
    await ctx.close();
  },

  async O_interrupted_while_checking(browser) {
    /* the tab is closed after PowerPoint has written a PDF and before the page has finished checking it */
    const opened = await L.open(browser, base, { platform: 'windows' });
    let page = opened.page;
    await addFolder(page, ['Classic 4x3.pptx', 'Poster A4 portrait.pptx']);
    await ready(page);
    await page.evaluate(() => { window.PptPdfPdf.inspect = () => new Promise(() => {}); });      /* the check never finishes */
    await page.click('#btn-start');
    await page.waitForFunction(() => window.__pptpdf.batch.items[0].status === 'verifying', null, { timeout: 60000 });
    await page.waitForTimeout(400);
    ok((await L.list(page, 'out')).includes('Classic 4x3.pdf'), 'the PDF is already in the folder');
    ok((await L.list(page, 'out/_slidesize/done')).includes('f00001.json'), 'and PowerPoint\'s result is still on disk while the check runs');
    await page.close();

    const again = await L.open(browser, base, { platform: 'windows', context: opened.ctx, wipe: false });
    page = again.page;
    await page.waitForFunction(() => /A batch was started here before/.test(document.getElementById('resume-msgs').textContent), null, { timeout: 15000 });
    await page.click('#btn-resume-last');
    await page.waitForFunction(() => window.__pptpdf.batch.items.length === 2 && window.__pptpdf.batch.items[0].status === 'done', null, { timeout: 60000 });
    await page.waitForFunction(() => window.__pptpdf.running === false, null, { timeout: 60000 });
    let it = await L.items(page);
    eq([it[0].state, it[0].checks.pageCount, it[0].checks.fidelity.pages.length], ['completed', 3, 3], 'after resuming, the check is finished from what was on disk');
    eq(await page.evaluate(() => window.__mock.converted.f00001 || 0), 0, 'without asking PowerPoint to convert the file again');
    await L.waitHelper(page);
    ok(!/already exist/.test(await text(page, '#start-msgs')) && !/already in the folder/.test(await text(page, '#plan-body')), 'and the batch\'s own half finished output is not mistaken for someone else\'s file');
    /* the second file had already been queued for PowerPoint, so it may have been finished without being asked */
    if ((await L.items(page)).some(i => i.status === 'ready')) {
      await page.waitForFunction(() => !document.getElementById('btn-start').disabled, null, { timeout: 30000 });
      await page.click('#btn-start'); await L.waitRun(page);
    }
    await page.waitForFunction(() => window.__pptpdf.batch.items.every(i => i.status === 'done'), null, { timeout: 60000 });
    it = await L.items(page);
    eq(it.map(i => i.state), ['completed', 'completed'], 'the rest of the batch finishes');
    eq(await L.list(page, 'out/_slidesize/done'), [], 'results are cleared once they have been checked');
    eq(again.errors, [], 'no script errors');
    await opened.ctx.close();
  },

  async M_presets_and_calculator(browser) {
    const { page, errors, ctx } = await L.open(browser, base, { platform: 'windows', query: '?w=3840&h=1080' });
    eq(await page.evaluate(() => Array.from(document.getElementById('preset').options).map(o => o.textContent)), ['Standard', 'Small file', 'Archive PDF/A', 'A4 handout with hidden slides', 'Changed, not saved'], 'presets are offered');
    await page.selectOption('#preset', 'archive');
    eq(await page.evaluate(() => [window.__pptpdf.batch.settings.pdfa, document.querySelector('#seg-pdfa .active').dataset.v]), [true, '1'], 'a preset sets the controls');
    await page.click('#seg-size [data-v="custom"]');
    eq(await page.evaluate(() => [document.getElementById('custom-w').value, document.getElementById('custom-h').value]), ['67.73', '19.05'], 'the custom size is filled from the calculator, 3840 x 1080 at 19.05 cm high');
    ok(/Filled in from the calculator/.test(await text(page, '#custom-note')), 'and says where it came from');
    eq(await page.inputValue('#preset'), '', 'changing a setting shows the preset as changed');
    page.once('dialog', d => d.accept('Wide blend'));
    await page.click('#btn-preset-save');
    await page.waitForFunction(() => document.getElementById('preset').value === 'user-wide-blend');
    await page.reload();
    eq(await page.evaluate(() => [document.getElementById('preset').value, window.__pptpdf.batch.settings.size.preset, Math.round(window.__pptpdf.batch.settings.size.customW)]), ['user-wide-blend', 'custom', 1920], 'a saved preset and the last settings survive a reload');
    await page.goto(base + '/index.html');
    const href = await page.evaluate(() => { document.getElementById('px-w').value = '5760'; document.getElementById('px-h').value = '1080'; const a = document.getElementById('pptpdf-link'); a.addEventListener('click', e => e.preventDefault()); a.click(); return a.getAttribute('href'); });
    eq(href, '/ppt-pdf.html?w=5760&h=1080', 'the homepage card carries the calculator size across');
    ok(/PowerPoint to PDF/.test(await text(page, '#pptpdf-link')), 'the tool is listed on the homepage');
    eq(errors.filter(e => !/favicon/.test(e)), [], 'no script errors');
    await ctx.close();
  }
};

(async () => {
  const server = await L.serve(port), browser = await chromium.launch({ args: ['--enable-precise-memory-info'] });
  const started = Date.now();
  try {
    for (const name of Object.keys(groups)) {
      const letter = name.split('_')[0];
      if (only && !only.test(letter)) continue;
      console.log('\n' + name.replace(/_/g, ' '));
      const before = failed, t0 = Date.now();
      try { await groups[name](browser); }
      catch (e) { failed++; fails.push(name + ' stopped: ' + e.message.split('\n')[0]); console.log('    STOPPED  ' + e.message.split('\n').slice(0, 6).join('\n             ')); }
      console.log('    ' + (failed === before ? 'ok' : 'FAILED') + '  ' + ((Date.now() - t0) / 1000).toFixed(1) + ' s');
    }
  } finally { await browser.close(); server.close(); }
  console.log('\n' + passed + ' passed, ' + failed + ' failed, ' + ((Date.now() - started) / 1000).toFixed(0) + ' s');
  if (failed) { console.log(fails.map(f => '  ' + f).join('\n')); process.exit(1); }
})().catch(e => { console.error(e); process.exit(1); });
