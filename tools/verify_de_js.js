/* Recompute the bundled workbook in JavaScript and diff against Excel's values.
 *   node tools/verify_de_js.js <bundle.json> <bundle.expect.json>            */
const fs = require('fs');
const XL = require('../src/xlengine.js');

const bundle = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const expect = JSON.parse(fs.readFileSync(process.argv[3], 'utf8'));

const g = new XL.Graph(bundle.cells);
const t0 = Date.now();
let total = 0, bad = 0;
const problems = [];
for (const ref of Object.keys(expect)) {
  const exp = expect[ref];
  const got = g.get(ref);
  total++;
  if (typeof got !== 'number' || Math.abs(exp - got) > 0.005) {
    bad++;
    if (problems.length < 10) problems.push([ref, exp, got, (bundle.cells[ref] || {}).f]);
  }
}
const ms = Date.now() - t0;
console.log(`cells checked : ${total}`);
console.log(`mismatched    : ${bad}  ${bad === 0 ? 'ALL MATCH' : '<-- LOOK'}`);
for (const [ref, e, got, f] of problems) {
  console.log(`   ${ref.padEnd(16)} excel=${String(e).padEnd(14)} js=${String(got).padEnd(14)} ${String(f).slice(0, 52)}`);
}
const errs = Object.keys(g.errors);
if (errs.length) {
  console.log(`\nformulas the engine could not evaluate: ${errs.length}`);
  errs.slice(0, 8).forEach(k => console.log(`   ${k}  ${g.errors[k]}`));
}
console.log(`\nfull recompute took ${ms} ms`);
process.exit(bad ? 1 : 0);
