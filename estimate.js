'use strict';

/* Refund-or-bill estimate: the pure maths behind the Position tab.
   No DOM and no network, so index.html loads it before app.js and tests/run.js loads it in Node.
   Every figure comes from rates.json (passed in as `rates`); nothing here is hard-coded. */
(function (root) {
  const DAY = 86400000;
  const r2 = (x) => Math.round((x + Number.EPSILON) * 100) / 100;
  // A blank field is null; anything else becomes a number (0 when it cannot be read).
  const num = (v) => {
    if (v == null || v === '') return null;
    const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/[$,\s]/g, ''));
    return Number.isFinite(n) ? n : null;
  };
  const amt = (v) => num(v) || 0;
  const parseDay = (s) => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null; };
  const addMonths = (d, m) => { const x = new Date(d.getTime()); x.setUTCMonth(x.getUTCMonth() + m); return x; };

  // The workbook's work-related PTR rows: these are what the $1,000 standard deduction replaces.
  // Union and professional association fees are claimed on top of it (ATO, standard deduction page), so they stay with the other PTR rows.
  const WORK_RELATED = [/other work-related/i, /work-related clothing/i, /work-related travel/i, /self-education/i, /home office/i];
  const isWorkRelated = (label) => WORK_RELATED.some((re) => re.test(label));
  // PTR rows that are offsets or rebates (LITO, the private health rebate, the super co-contribution): not deductions, listed separately.
  const isOffsetRow = (label) => /offset|rebate|co-contribution/i.test(label);
  const sumRows = (rows) => rows.reduce((a, r) => a + (r.total || 0), 0);

  /* ----- the workbook side ----- */
  function workbookParts(d) {
    const items = (d && d.items) || [];
    const income = items.filter((i) => i.section === 'income' && !/capital gain/i.test(i.label));
    const ptrAll = items.filter((i) => i.section === 'ptr');
    const ptr = ptrAll.filter((i) => !isOffsetRow(i.label));
    return {
      incomeRows: income, incomeTotal: sumRows(income), rentalIncome: sumRows(income.filter((i) => /rental/i.test(i.label))),
      workRelated: ptr.filter((i) => isWorkRelated(i.label)), ptrOther: ptr.filter((i) => !isWorkRelated(i.label)),
      notDeductions: ptrAll.filter((i) => isOffsetRow(i.label)),
      // A personal super row in the PTR block is a reportable super contribution (added back for the surcharge and HELP tests);
      // an investment property row there joins the IP section in the net rental loss test.
      ptrSuper: ptr.filter((i) => /personal superannuation/i.test(i.label)), ptrIp: ptr.filter((i) => /investment property/i.test(i.label)),
      ip: items.filter((i) => i.section === 'ip'), ctr: items.filter((i) => i.section === 'ctr'),
    };
  }
  // Default for the "Use the $1,000 standard deduction instead" toggle: on when the work-related rows are under the cap and there is salary.
  function standardDeductionDefault(d, a, rates) {
    const wr = sumRows(workbookParts(d).workRelated);
    return wr < rates.standardDeduction.max && amt(a && a.salary) > 0;
  }

  /* ----- capital gains ----- */
  // held > 12 months: the day of purchase and the day of sale do not count, so the sale must fall after the first anniversary.
  function heldOver12Months(bought, sold, rates) {
    const b = parseDay(bought), s = parseDay(sold);
    if (!b || !s) return null;
    return s.getTime() > addMonths(b, rates.cgt.holdMonths).getTime();
  }
  function cgtRow(ev, rates) {
    const costBase = num(ev.costBase), proceeds = num(ev.proceeds);
    const held = heldOver12Months(ev.bought, ev.sold, rates);
    const b = parseDay(ev.bought), s = parseDay(ev.sold);
    const days = b && s ? Math.max(0, Math.round((s - b) / DAY) - 1) : null;
    const gain = costBase != null && proceeds != null ? r2(proceeds - costBase) : null;
    return {
      id: ev.id || '', asset: ev.asset || '', bought: ev.bought || '', sold: ev.sold || '', note: ev.note || '',
      costBase, proceeds, gain, daysHeld: days, heldOver12Months: held,
      missingCostBase: costBase == null, discountApplies: gain != null && gain > 0 && held === true,
    };
  }
  // Rules per rates.cgt: losses (carried-forward first) come off before the discount, and off the non-discount gains first.
  function netCapitalGain(cgt, rates, fy) {
    const events = (cgt && cgt.events) || [];
    const inYear = (row) => { const s = parseDay(row.sold); return !s || (s.getTime() >= Date.UTC(fy - 1, 6, 1) && s.getTime() <= Date.UTC(fy, 5, 30)); };
    const rows = events.map((ev) => cgtRow(ev, rates)).map((row) => ({ ...row, inYear: inYear(row) }));
    const counted = rows.filter((row) => row.gain != null && row.inYear);
    const discountGains = counted.filter((row) => row.gain > 0 && row.heldOver12Months === true).reduce((a, row) => a + row.gain, 0);
    const otherGains = counted.filter((row) => row.gain > 0 && row.heldOver12Months !== true).reduce((a, row) => a + row.gain, 0);
    const yearLosses = counted.filter((row) => row.gain < 0).reduce((a, row) => a - row.gain, 0);
    const carried = amt(cgt && cgt.carriedLoss);
    let losses = yearLosses + carried;
    const offOther = Math.min(otherGains, losses); losses -= offOther;
    const offDiscount = Math.min(discountGains, losses); losses -= offDiscount;
    const discountable = discountGains - offDiscount;
    const discount = r2(discountable * rates.cgt.discount);
    const net = r2(Math.max(0, (otherGains - offOther) + discountable - discount));
    return {
      rows, totalGains: r2(discountGains + otherGains), discountGains: r2(discountGains), otherGains: r2(otherGains),
      yearLosses: r2(yearLosses), carriedLoss: r2(carried), lossesApplied: r2(offOther + offDiscount), lossToCarry: r2(losses),
      discount, net, missing: rows.filter((row) => row.missingCostBase),
    };
  }

  /* ----- the tax pieces, each straight from rates.json ----- */
  function bracketTax(taxable, brackets) {
    let tax = 0, prev = 0, marginal = 0;
    for (const b of brackets) {
      const top = b.upTo == null ? Infinity : b.upTo;
      if (taxable > prev) tax += (Math.min(taxable, top) - prev) * b.rate;
      if (taxable <= top) { marginal = b.rate; break; }
      prev = top;
    }
    return { tax, marginal };
  }
  function lito(taxable, L) {
    const clamp = (x, lo, hi) => Math.min(hi, Math.max(lo, x));
    const off = L.max - L.taper1.rate * clamp(taxable - L.taper1.from, 0, L.taper1.to - L.taper1.from)
      - L.taper2.rate * clamp(taxable - L.taper2.from, 0, L.taper2.to - L.taper2.from);
    return Math.max(0, off);
  }
  // Below the lower threshold nothing; in the phase-in band 10% of the excess (always the lesser of the two); above it 2% of taxable income.
  function medicare(taxable, M) {
    const li = M.lowIncome.single;
    if (taxable <= li.lower) return { levy: 0, marginal: 0, phaseIn: false };
    const full = taxable * M.rate, shaded = (taxable - li.lower) * li.phaseInRate;
    return shaded < full ? { levy: shaded, marginal: li.phaseInRate, phaseIn: true } : { levy: full, marginal: M.rate, phaseIn: false };
  }
  function mlsTier(income, tiers) {
    for (let i = 0; i < tiers.length; i++) if (tiers[i].upTo == null || income <= tiers[i].upTo) return { tier: i, rate: tiers[i].rate, upTo: tiers[i].upTo, from: i ? tiers[i - 1].upTo : 0 };
    return { tier: tiers.length - 1, rate: tiers[tiers.length - 1].rate, upTo: null, from: 0 };
  }
  function helpRepayment(income, H) {
    for (const b of H.bands) {
      if (b.upTo != null && income > b.upTo) continue;
      if (b.flatRateOnTotalIncome != null) return { amount: income * b.flatRateOnTotalIncome, marginal: b.flatRateOnTotalIncome, band: b };
      return { amount: (b.base || 0) + (b.rate || 0) * Math.max(0, income - (b.over || 0)), marginal: b.rate || 0, band: b };
    }
    return { amount: 0, marginal: 0, band: null };
  }

  // Assumptions "exist" once the sheet has been saved, or when any figure in it is filled in.
  function hasAssumptions(a) {
    if (!a || typeof a !== 'object') return false;
    if (a.savedAt) return true;
    const fields = ['salary', 'paygWithheld', 'paygPerPay', 'dividends', 'franking', 'interest', 'trust', 'other', 'super', 'incomeProtection'];
    if (fields.some((k) => amt(a[k]) > 0)) return true;
    if (a.help && a.help.has) return true;
    return !!(a.cgt && (a.cgt.events || []).some((e) => num(e.proceeds) != null || num(e.costBase) != null));
  }
  const projectedPayg = (a) => r2(amt(a.paygWithheld) + amt(a.paygPerPay) * amt(a.paysLeft));

  /* ----- the chain ----- */
  function estimatePosition(d, a, rates, fy) {
    if (!hasAssumptions(a) || !rates) return { hasAssumptions: false, warnings: [] };
    const warnings = [];
    const warn = (code, text) => warnings.push({ code, text });
    const wb = workbookParts(d);
    const standardDeductionOn = a.useStandardDeduction == null ? standardDeductionDefault(d, a, rates) : !!a.useStandardDeduction;
    const workRelatedTotal = r2(sumRows(wb.workRelated));
    const ctrTotal = r2(sumRows(wb.ctr));
    const includeCtr = !!a.includeCtr;
    const salary = r2(amt(a.salary)), dividends = r2(amt(a.dividends)), franking = r2(amt(a.franking)), interest = r2(amt(a.interest));
    const trust = r2(amt(a.trust)), other = r2(amt(a.other)), superDed = r2(amt(a.super)), incomeProtection = r2(amt(a.incomeProtection));
    // The standard deduction is the lesser of the cap and labour income, and the ATO reduces it dollar for dollar by actual claims,
    // so the work-related part used is never below the itemised rows.
    const standardDeductionAmount = r2(Math.min(rates.standardDeduction.max, salary));
    const workRelatedUsed = r2(standardDeductionOn ? Math.max(workRelatedTotal, standardDeductionAmount) : workRelatedTotal);
    const standardReplaced = standardDeductionOn && standardDeductionAmount > workRelatedTotal;
    const ptrUsed = r2(sumRows(wb.ptrOther) + workRelatedUsed);
    const notDeductions = r2(sumRows(wb.notDeductions));
    const ipTotal = r2(sumRows(wb.ip));
    // A negatively geared property: the loss comes off taxable income but is added back for the surcharge and HELP tests.
    // Investment property rows recorded in the PTR block count with the IP section here.
    const netRentalLoss = r2(Math.max(0, ipTotal + sumRows(wb.ptrIp) - wb.rentalIncome));
    const workbookIncome = r2(wb.incomeTotal);
    const workbookDeductions = r2(ptrUsed + ipTotal + (includeCtr ? ctrTotal : 0));
    const otherIncome = r2(salary + dividends + franking + interest + trust + other);

    const cgt = netCapitalGain(a.cgt, rates, fy);
    const taxableIncome = r2(Math.max(0, workbookIncome + otherIncome + cgt.net - workbookDeductions - superDed - incomeProtection));

    // Medicare levy surcharge: only with hospital cover for the whole year is there none. Unanswered counts as no cover until it is answered.
    // Tier by the single or family table (family when so marked or when there are dependent children; the threshold grows per extra child).
    const hospitalCover = a.hospitalCover === true;
    const hospitalUnanswered = a.hospitalCover == null;
    const family = a.family || {};
    const children = amt(family.children) || 0;
    const isFamily = family.status === 'family' || children > 0;
    const extraKids = Math.max(0, children - 1) * rates.mls.familyPerExtraChild;
    const tiers = (isFamily ? rates.mls.family : rates.mls.single).map((t) => ({ ...t, upTo: t.upTo == null ? null : t.upTo + (isFamily ? extraKids : 0) }));
    // Income for both tests adds back reportable super (deductible personal super from the sheet and the workbook's personal super row)
    // and the net rental loss, per rates.mls/help.incomeBasis.
    const ptrSuper = r2(sumRows(wb.ptrSuper));
    const addBacks = r2(superDed + ptrSuper + netRentalLoss);
    const help = a.help || {};
    const helpOn = !!help.has;
    const helpBalance = num(help.balance);

    // The pieces that depend on taxable income, so the same chain can be re-run at one dollar more for the marginal rate.
    function piecesAt(ti) {
      const bt = bracketTax(ti, rates.incomeTax.brackets);
      const litoRaw = Math.min(lito(ti, rates.lito), bt.tax);
      const med = medicare(ti, rates.medicareLevy);
      const mlsIncome = ti + addBacks + (isFamily ? amt(family.spouseIncome) : 0);
      const tier = mlsTier(mlsIncome, tiers);
      const mls = hospitalCover ? 0 : ti * tier.rate;
      // HELP: repayment income = taxable income + the add-backs (reportable fringe benefits and exempt foreign income are not modelled).
      const helpIncome = ti + addBacks;
      const hr = helpOn ? helpRepayment(helpIncome, rates.help) : { amount: 0, marginal: 0, band: null };
      const helpRepay = helpOn && helpBalance != null ? Math.min(hr.amount, Math.max(0, helpBalance)) : hr.amount;
      return { bt, litoRaw, med, mlsIncome, tier, mls, helpIncome, hr, helpRepay, total: bt.tax - litoRaw + med.levy + mls + helpRepay };
    }
    const p = piecesAt(taxableIncome);
    const incomeTax = r2(p.bt.tax);
    const litoOffset = r2(p.litoRaw);
    const med = p.med;
    const medicareLevy = r2(med.levy);
    const mlsIncome = r2(p.mlsIncome);
    const tier = p.tier;
    const mlsRate = hospitalCover ? 0 : tier.rate;
    const mls = r2(p.mls);
    const helpIncome = r2(p.helpIncome);
    const helpUncapped = r2(p.hr.amount);
    const helpRepay = r2(p.helpRepay);
    const helpCapped = helpOn && helpBalance != null && helpUncapped > helpBalance;

    const totalLiability = r2(incomeTax - litoOffset + medicareLevy + mls + helpRepay - franking);
    const credits = projectedPayg(a);
    const result = r2(credits - totalLiability);
    // The rate on the next dollar, worked out numerically: the whole chain at one dollar more, less the chain now (to four decimals).
    const marginal = Math.round((piecesAt(taxableIncome + 1).total - p.total) * 10000) / 10000;

    /* warnings */
    const addBackText = [superDed || ptrSuper ? `personal super ${fmt(superDed + ptrSuper)}` : '', netRentalLoss ? `net rental loss ${fmt(netRentalLoss)}` : ''].filter(Boolean).join(' and ');
    if (standardDeductionOn) warn('standard-deduction', standardReplaced
      ? `Standard deduction of ${fmt(standardDeductionAmount)} used instead of the workbook's ${fmt(workRelatedTotal)} work-related rows`
      : `Standard deduction on, but the workbook's ${fmt(workRelatedTotal)} work-related rows are more than the ${fmt(standardDeductionAmount)} standard amount, so ${fmt(workRelatedUsed)} is used`);
    if (!includeCtr && ctrTotal) warn('ctr-excluded', `CTR Deductions of ${fmt(ctrTotal)} left out`);
    if (med.phaseIn || (taxableIncome > 0 && taxableIncome <= rates.medicareLevy.lowIncome.single.upper)) warn('medicare-provisional', `Medicare low-income thresholds are ${rates.medicareLevy.lowIncome.year} figures, ${rates.label} not yet published`);
    if (hospitalUnanswered) warn('hospital-unanswered', 'Hospital cover not answered: the surcharge is included until you say you have cover');
    if (!hospitalCover) {
      const near = [tier.from, tier.upTo].filter((x) => x).some((edge) => Math.abs(mlsIncome - edge) <= 5000);
      if (near) warn('mls-boundary', `Within $5,000 of a Medicare levy surcharge tier boundary (${isFamily ? 'family' : 'single'} test on ${fmt(mlsIncome)})`);
      if (addBacks) warn('mls-basis', `Surcharge test income ${fmt(mlsIncome)} adds back ${addBackText}; reportable fringe benefits are not modelled`);
    }
    if (helpOn) {
      const cliff = rates.help.bands.find((b) => b.flatRateOnTotalIncome != null);
      const top = rates.help.bands[rates.help.bands.indexOf(cliff) - 1].upTo;
      if (helpIncome > top - 5000 && helpIncome <= top) warn('help-cliff', `Within $5,000 of the HELP top band: over ${fmt(top)} the repayment becomes ${Math.round(cliff.flatRateOnTotalIncome * 100)}% of your whole income (about ${fmt((top + 1) * cliff.flatRateOnTotalIncome)})`);
      else if (helpIncome > top && helpIncome <= top + 5000) warn('help-cliff', `Just over the HELP top band (${fmt(top)}): the repayment is ${Math.round(cliff.flatRateOnTotalIncome * 100)}% of your whole income`);
      if (helpBalance == null) warn('help-balance', 'HELP balance not entered, so the repayment is not capped');
      if (helpCapped) warn('help-capped', `HELP repayment capped at the ${fmt(helpBalance)} balance`);
      warn('help-basis', `HELP repayment income ${fmt(helpIncome)} is taxable income${addBacks ? ` plus ${addBackText}` : ''}; reportable fringe benefits are not modelled`);
    }
    for (const row of cgt.missing) warn('cgt-costbase', `${row.asset || 'A CGT event'}: no cost base yet${row.proceeds != null ? ', so its gain is not in the estimate' : ''}`);
    // A gain with a date missing cannot be tested for the 12-month discount, so it is counted in full until both dates are in.
    for (const row of cgt.rows) if (row.inYear && row.gain != null && row.gain > 0 && row.heldOver12Months == null) {
      warn('cgt-dates', `${row.asset || 'A CGT event'}: enter the bought and sold dates to test the 12-month discount; counted in full for now${row.sold ? '' : ` and taken as sold in FY${fy % 100}`}`);
    }
    for (const row of cgt.rows) if (!row.inYear && row.gain != null) warn('cgt-outside-year', `${row.asset}: sold outside FY${fy % 100}, not counted`);
    if (cgt.lossToCarry) warn('cgt-loss-carry', `${fmt(cgt.lossToCarry)} of capital losses left to carry forward`);

    /* the chain as the UI lists it */
    const chain = [
      { key: 'workbookIncome', label: 'Income in the workbook', value: workbookIncome, kind: 'income' },
      { key: 'salary', label: 'Salary and wages', value: salary, kind: 'income', optional: true },
      { key: 'dividends', label: 'Dividends', value: dividends, kind: 'income', optional: true },
      { key: 'franking', label: 'Franking credits (grossed up)', value: franking, kind: 'income', optional: true },
      { key: 'interest', label: 'Interest', value: interest, kind: 'income', optional: true },
      { key: 'trust', label: 'Trust distributions', value: trust, kind: 'income', optional: true },
      { key: 'other', label: 'Other income', value: other, kind: 'income', optional: true },
      { key: 'netCapitalGain', label: 'Net capital gain', value: cgt.net, kind: 'income', optional: true },
      { key: 'workbookDeductions', label: standardReplaced ? 'Workbook deductions (work-related rows replaced by the standard deduction)' : 'Workbook deductions', value: -workbookDeductions, kind: 'deduction' },
      { key: 'super', label: 'Personal super contributions', value: -superDed, kind: 'deduction', optional: true },
      { key: 'incomeProtection', label: 'Income protection premiums', value: -incomeProtection, kind: 'deduction', optional: true },
      { key: 'taxableIncome', label: 'Taxable income', value: taxableIncome, kind: 'subtotal' },
      { key: 'incomeTax', label: 'Income tax', value: incomeTax, kind: 'tax' },
      { key: 'lito', label: 'Low income tax offset', value: -litoOffset, kind: 'offset', optional: true },
      { key: 'medicareLevy', label: 'Medicare levy', value: medicareLevy, kind: 'tax' },
      { key: 'mls', label: 'Medicare levy surcharge', value: mls, kind: 'tax', optional: !mls && hospitalCover },
      { key: 'help', label: 'HELP repayment', value: helpRepay, kind: 'tax', optional: !helpOn },
      { key: 'frankingOffset', label: 'Franking credit offset', value: -franking, kind: 'offset', optional: true },
      { key: 'totalLiability', label: 'Tax, levies and repayments', value: totalLiability, kind: 'subtotal' },
      { key: 'credits', label: 'PAYG withheld (projected to 30 June)', value: -credits, kind: 'credit' },
      { key: 'result', label: result >= 0 ? 'Estimated refund' : 'Estimated amount owing', value: result, kind: 'result' },
    ];
    const assumptionsUsed = [
      salary ? `salary ${fmt(salary)}` : '', credits ? `PAYG withheld ${fmt(credits)}` : '',
      dividends ? `dividends ${fmt(dividends)}` : '', franking ? `franking credits ${fmt(franking)}` : '', interest ? `interest ${fmt(interest)}` : '',
      trust ? `trust distributions ${fmt(trust)}` : '', other ? `other income ${fmt(other)}` : '',
      superDed ? `personal super ${fmt(superDed)}` : '', incomeProtection ? `income protection ${fmt(incomeProtection)}` : '',
      standardReplaced ? `standard deduction ${fmt(standardDeductionAmount)}` : '', includeCtr ? 'CTR rows included' : '',
      helpOn ? `HELP debt${helpBalance != null ? ` ${fmt(helpBalance)}` : ''}` : 'no HELP debt',
      hospitalCover ? 'hospital cover all year' : `${hospitalUnanswered ? 'hospital cover not answered, surcharge included' : 'no hospital cover'} (${isFamily ? 'family' : 'single'})`,
      cgt.rows.length ? `${cgt.rows.length} capital gains event${cgt.rows.length === 1 ? '' : 's'}` : '',
    ].filter(Boolean);

    return {
      hasAssumptions: true, fy,
      workbookIncome, workbookDeductions, ptrUsed, notDeductions, ipTotal, netRentalLoss, ctrTotal, includeCtr,
      workRelatedTotal, standardDeductionOn, standardDeductionAmount, workRelatedUsed, standardReplaced,
      salary, dividends, franking, interest, trust, other, otherIncome, super: superDed, incomeProtection, ptrSuper, addBacks,
      cgt, netCapitalGain: cgt.net, taxableIncome,
      incomeTax, lito: litoOffset, medicareLevy, hospitalCover, hospitalUnanswered, isFamily, mlsIncome, mlsTier: hospitalCover ? null : tier.tier, mlsRate, mls,
      helpOn, helpIncome, helpUncapped, help: helpRepay, helpCapped, helpBalance, frankingOffset: franking,
      totalLiability, credits, result, marginal, warnings, chain, assumptionsUsed,
      provisional: [
        { label: 'Medicare levy low-income thresholds', year: rates.medicareLevy.lowIncome.year, provisional: !!rates.medicareLevy.lowIncome.provisional },
        { label: 'Working from home fixed rate', year: rates.wfhFixedRatePerHour.year, provisional: !!rates.wfhFixedRatePerHour.provisional },
      ].filter((p) => p.provisional),
    };
  }
  const fmt = (n) => (n < 0 ? '-' : '') + '$' + Math.abs(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');

  const api = { estimatePosition, hasAssumptions, standardDeductionDefault, cgtRow, netCapitalGain, heldOver12Months, projectedPayg, workbookParts, num, fmtMoney: fmt };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Estimate = api;
})(typeof window !== 'undefined' ? window : globalThis);
