/* Engine checks for Deck Compare.

   Needs   python3 with python-pptx and Pillow (to draw the decks),
           npm packages jszip and @xmldom/xmldom (or NODE_MODULES_DIR
           pointing at a folder that has them).
   Run     node tests/compare/check.js */
'use strict';
const path = require('path'), fs = require('fs'), os = require('os');
const { spawnSync } = require('child_process');
function need(name) {
  try { return require(name); } catch (e) {}
  if (process.env.NODE_MODULES_DIR) return require(path.join(process.env.NODE_MODULES_DIR, name));
  throw new Error('Cannot find ' + name + '. npm i ' + name + ' or set NODE_MODULES_DIR.');
}
const JSZip = need('jszip');
const { DOMParser } = need('@xmldom/xmldom');
const DC = require('../../deck-compare.js');

const dir = process.env.WORK || path.join(os.tmpdir(), 'deck-compare-checks');
const r = spawnSync('python3', [path.join(__dirname, 'make_decks.py'), dir], { stdio: 'inherit' });
if (r.status !== 0) throw new Error('could not draw the decks');

async function load(name) {
  const buf = fs.readFileSync(path.join(dir, name));
  const zip = await JSZip.loadAsync(buf);
  return DC.parse(zip, { name, size: buf.length, parseXml: s => new DOMParser({ onError: () => {} }).parseFromString(s, 'application/xml') });
}

let fails = 0;
function ok(cond, msg) { console.log((cond ? '  ok   ' : '  FAIL ') + msg); if (!cond) fails++; }
function row(rep, title) { return rep.rows.find(r => DC.titleOf(r) === title); }
function has(r, re) { return !!r && r.changes.some(c => re.test(c.text)); }

(async () => {
  const v1 = await load('v1.pptx'), v2 = await load('v2.pptx');
  const rep = DC.compare(v1, v2);
  console.log(DC.toText(rep));

  console.log('v1 → v2');
  ok(rep.idsTrusted, 'shared history spotted, IDs trusted');
  ok(rep.counts.added === 1 && row(rep, 'Housekeeping').status === 'added', 'Housekeeping added');
  ok(rep.counts.deleted === 1 && row(rep, 'Guest speaker').status === 'deleted', 'Guest speaker deleted');
  ok(rep.counts.moved === 1, 'exactly one slide flagged as moved (' + rep.counts.moved + ')');
  ok(has(row(rep, 'Annual Kick-off 2026'), /"14" → "15"/), 'date edit on the title slide');
  const ag = row(rep, 'Agenda');
  ok(has(ag, /now hidden/), 'agenda hidden');
  ok(has(ag, /"yaer" → "year"/), 'agenda typo fix');
  ok(ag.changes.length === 2, 'agenda has no noise (' + ag.changes.length + ' changes)');
  const pic = row(rep, 'Our new office');
  ok(has(pic, /picture replaced/), 'picture replaced');
  ok(has(pic, /moved 2 cm right/), 'picture moved 2 cm right');
  ok(has(row(rep, 'Sizzle reel'), /video file replaced/), 'video replaced');
  ok(has(row(rep, 'Numbers'), /"8\.1m" → "8\.3m"|"1m" → "3m"|"8" ?.*→/), 'table figure');
  const ty = row(rep, 'Thank you');
  ok(has(ty, /formatting changed.*size from layout → 36 pt/), 'font size bump');
  ok(has(ty, /Speaker notes added/), 'notes added');
  ok(rep.rows.every(r => r.status !== 'unchanged'), 'nothing reported unchanged that changed');

  console.log('\nv1 → v1');
  const same = DC.compare(v1, await load('v1.pptx'));
  ok(same.rows.every(r => r.status === 'unchanged' && !r.changes.length), 'identical decks show no changes');
  ok(!same.deck.length, 'identical decks show no deck changes');

  console.log('\nv1 → v2 with a pasted slide');
  const pasted = DC.compare(v1, await load('v2-pasted.pptx'));
  const pr = row(pasted, 'Our new office');
  ok(pr && pr.match === 'content' && pr.status === 'changed', 'pasted slide still paired, as a probable match');
  ok(pasted.counts.added === 1 && pasted.counts.deleted === 1, 'no extra add/delete from the pasted slide');

  console.log('\nv1 → unrelated deck');
  const other = DC.compare(v1, await load('other.pptx'));
  ok(!other.idsTrusted, 'colliding IDs not trusted');
  ok(other.rows.filter(r => r.a && r.b).length <= 1, 'unrelated slides not paired (' + other.rows.filter(r => r.a && r.b).length + ')');

  console.log('\nword diff');
  ok(JSON.stringify(DC.wordDiff('a b c', 'a x c')) === JSON.stringify([['=', 'a '], ['-', 'b'], ['+', 'x'], ['=', ' c']]), 'single word');
  ok(DC.diffSummary(DC.wordDiff('one two three four', 'one five six four')) === '"two three" → "five six"', 'adjacent words merge into one edit');

  console.log(fails ? '\n' + fails + ' failed' : '\nall passed');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
