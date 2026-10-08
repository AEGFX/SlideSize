/* Browser checks for Photo Montage.

   These drive the real page in Chromium, export real videos and measure
   them with ffmpeg: frame counts, frame rate, colour tags, the frames
   either side of the loop point, and preview against export.

   They are optional. The site does not need this folder.

   Needs   node 18+, ffmpeg and ffprobe on the path,
           npm packages playwright and sharp,
           python3 with Pillow and numpy (to draw the test photos).
   Run     node tests/browser/run.js            everything
           node tests/browser/run.js export A   one group of export checks
           node tests/browser/run.js ui "F|G"   some groups of interface checks

   Chromium without Google's codecs has no H.264 encoder, so there the
   exports use VP9 and AV1 and the page's handling of a missing H.264
   encoder is what gets checked. Run with a full Chrome to cover H.264. */
'use strict';
const { spawn, spawnSync } = require('child_process');
const path = require('path'), fs = require('fs'), os = require('os'), http = require('http');

const root = path.resolve(__dirname, '..', '..');
const work = process.env.WORK || path.join(os.tmpdir(), 'slidesize-browser-checks');
const port = Number(process.env.PORT || 8912);
const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json' };

function serve() {
  return new Promise(resolve => {
    const server = http.createServer((req, res) => {
      let p = decodeURIComponent(req.url.split('?')[0]);
      if (p === '/') p = '/index.html';
      const file = path.join(root, p);
      if (!file.startsWith(root) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); res.end('not found'); return; }
      res.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-store' });
      fs.createReadStream(file).pipe(res);
    });
    server.listen(port, '127.0.0.1', () => resolve(server));
  });
}

(async () => {
  for (const d of ['out', 'shots', 'photos']) fs.mkdirSync(path.join(work, d), { recursive: true });
  if (!fs.existsSync(path.join(work, 'photos', 'set40'))) {
    for (const args of [['set12', '12', 'extras'], ['set40', '40']]) {
      const r = spawnSync('python3', [path.join(__dirname, 'make_photos.py'), path.join(work, 'photos', args[0])].concat(args.slice(1)), { stdio: 'inherit' });
      if (r.status !== 0) throw new Error('Could not draw the test photos. Python with Pillow and numpy is needed.');
    }
  }
  const server = await serve();
  const which = process.argv[2], filter = process.argv[3];
  const jobs = [['export', 'export-checks.js'], ['ui', 'ui-checks.js']].filter(j => !which || which === j[0]);
  let failed = false;
  for (const job of jobs) {
    const code = await new Promise(resolve => {
      const child = spawn(process.execPath, [path.join(__dirname, job[1])].concat(filter ? [filter] : []),
        { stdio: 'inherit', env: Object.assign({}, process.env, { WORK: work, BASE: 'http://localhost:' + port }) });
      child.on('exit', resolve);
    });
    if (code !== 0) failed = true;
  }
  server.close();
  console.log('\nFiles and screenshots are in ' + work);
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error(e.message || e); process.exit(1); });
