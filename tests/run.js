'use strict';

/* Hand-worked checks of estimate.js, to the cent. No DOM, no network.
   Run: /usr/local/opt/node/bin/node tests/run.js   (from the app folder; any cwd works)
   Each case in cases.json carries its working in the "working" array; the figures asserted here are copied from that working.
   Prints PASS/FAIL per case and exits 1 if anything failed. */
const path = require('path');
const fs = require('fs');
const E = require(path.join(__dirname, '..', 'estimate.js'));
const rates = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'rates.json'), 'utf8'));
const spec = JSON.parse(fs.readFileSync(path.join(__dirname, 'cases.json'), 'utf8'));

const near = (got, want) => (typeof want === 'number' ? typeof got === 'number' && Math.abs(got - want) < 0.005 : got === want);
const show = (v) => (v === undefined ? 'undefined' : JSON.stringify(v));
const workbook = (ref) => (ref == null ? null : typeof ref === 'string' ? spec[ref] : ref);

let failed = 0, passed = 0;
function check(problems, label, got, want) {
  if (!near(got, want)) problems.push(`${label}: expected ${show(want)}, got ${show(got)}`);
}
function runCase(c) {
  const problems = [];
  const d = workbook(c.d);
  const est = E.estimatePosition(d, c.a, rates, spec.fy);
  for (const [k, v] of Object.entries(c.expect || {})) check(problems, k, est[k], v);
  for (const k of c.expectAbsent || []) if (est[k] !== undefined) problems.push(`${k} should be absent, got ${show(est[k])}`);
  const codes = (est.warnings || []).map((w) => w.code);
  for (const w of c.expectWarnings || []) if (!codes.includes(w)) problems.push(`warning ${w} missing (have: ${codes.join(', ') || 'none'})`);
  for (const w of c.expectNoWarnings || []) if (codes.includes(w)) problems.push(`warning ${w} should not fire`);
  if (c.expectCgt) for (const [k, v] of Object.entries(c.expectCgt)) check(problems, `cgt.${k}`, est.cgt && est.cgt[k], v);
  if (c.expectRows) c.expectRows.forEach((want, i) => {
    const row = est.cgt && est.cgt.rows[i];
    if (!row) { problems.push(`cgt row ${i} missing`); return; }
    for (const [k, v] of Object.entries(want)) check(problems, `cgt.rows[${i}].${k}`, row[k], v);
  });
  for (const hc of c.heldChecks || []) check(problems, `heldOver12Months(${hc.bought}, ${hc.sold})`, E.heldOver12Months(hc.bought, hc.sold, rates), hc.expect);
  if (c.defaultCheck !== undefined) check(problems, 'standardDeductionDefault', E.standardDeductionDefault(d, c.a, rates), c.defaultCheck);
  if (c.noSalaryDefault) check(problems, 'standardDeductionDefault (no salary)', E.standardDeductionDefault(c.noSalaryDefault.d, c.noSalaryDefault.a, rates), c.noSalaryDefault.expect);
  if (c.also) {
    const est2 = E.estimatePosition(d, c.also.a, rates, spec.fy);
    for (const [k, v] of Object.entries(c.also.expect || {})) check(problems, `also.${k}`, est2[k], v);
    const codes2 = (est2.warnings || []).map((w) => w.code);
    for (const w of c.also.expectNoWarnings || []) if (codes2.includes(w)) problems.push(`also: warning ${w} should not fire`);
    for (const w of c.also.expectWarnings || []) if (!codes2.includes(w)) problems.push(`also: warning ${w} missing`);
  }
  // Every chain must add up: the listed rows reproduce taxable income, the total and the result.
  if (est.chain) {
    const val = (k) => est.chain.find((r) => r.key === k).value;
    const incomeSide = est.chain.filter((r) => r.kind === 'income' || r.kind === 'deduction').reduce((a, r) => a + r.value, 0);
    check(problems, 'chain adds to taxable income', Math.max(0, Math.round(incomeSide * 100) / 100), val('taxableIncome'));
    const taxSide = est.chain.filter((r) => ['tax', 'offset'].includes(r.kind)).reduce((a, r) => a + r.value, 0);
    check(problems, 'chain adds to the total', Math.round(taxSide * 100) / 100, val('totalLiability'));
    check(problems, 'chain result = credits - total', Math.round((-val('credits') - val('totalLiability')) * 100) / 100, val('result'));
    const dash = est.chain.concat(est.warnings).map((r) => r.label || r.text).join(' ');
    if (/[–—]/.test(dash)) problems.push('an em or en dash crept into UI text');
    if (/\bAI\b/.test(dash)) problems.push('the word AI crept into UI text');
  }
  if (problems.length) { failed++; console.log(`FAIL ${c.id}: ${c.name}`); for (const p of problems) console.log(`     ${p}`); }
  else { passed++; console.log(`PASS ${c.id}: ${c.name}`); }
}

for (const c of spec.cases) runCase(c);

// rates.json itself: the checkpoints the ATO publishes must agree with the brackets (guards against a typo in the file).
{
  const problems = [];
  for (const [at, want] of Object.entries(rates.incomeTax.checkpoints)) {
    const est = E.estimatePosition(null, { salary: Number(at), useStandardDeduction: false, savedAt: 'x' }, rates, spec.fy);
    check(problems, `tax at ${at}`, est.incomeTax, want);
  }
  // Medicare phase-in meets the full rate at the upper threshold (10% x 7,002 = 700.20 against 2% x 35,013 = 700.26)
  const up = E.estimatePosition(null, { salary: rates.medicareLevy.lowIncome.single.upper, useStandardDeduction: false, savedAt: 'x' }, rates, spec.fy);
  check(problems, 'medicare at the upper threshold', up.medicareLevy, 700.20);
  if (problems.length) { failed++; console.log('FAIL rates: checkpoints'); for (const p of problems) console.log(`     ${p}`); }
  else { passed++; console.log('PASS rates: ATO checkpoints 45,000 / 135,000 / 190,000 and the Medicare phase-in meet'); }
}

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
