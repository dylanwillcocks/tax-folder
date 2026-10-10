'use strict';

/* Phone rendering review helper (round 1): the same mocked page as render.js, with element screenshots of the
   refund-or-owing hero, the "How it is worked out" card, the Capital gains card and the Assumptions dialog in
   light and dark, plus DOM measurements: overflow past the viewport, clipped text, overlapping kv labels,
   tap targets under 44 px, text contrast and the dialog's internal scrolling.
   Run: /usr/local/opt/node/bin/node tests/measure.js */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SHOTS = path.join(__dirname, 'shots');
const FIXTURE_XLSX = path.join(ROOT, '..', 'ptr27.xlsx');
const PW_CANDIDATES = [
  process.env.PW_PATH,
  '/Users/dylanwillcocks/oakwood-segment-recovery/cdn/pw/node_modules/playwright',
  '/Users/dylanwillcocks/oakwood-segment-recovery/cdn/pw/node_modules/playwright-core',
  '/Users/dylanwillcocks/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright-core',
  'playwright', 'playwright-core',
].filter(Boolean);
function loadPlaywright() {
  for (const p of PW_CANDIDATES) { try { return { pw: require(p), from: p }; } catch (e) { /* next */ } }
  throw new Error('playwright or playwright-core not found. Set PW_PATH to its folder.');
}

function syntheticWorkbook() {
  const XLSX = require(path.join(ROOT, 'xlsx.min.js'));
  const streams = ['TRUSTS 1&2:  1 Sample Road, Sampleton QLD 4000', '12 Sample Street, Sampleton QLD 4000', 'Gym (casual)', 'Fitness (self employed)', 'Company QLD', 'Other'];
  const row = (label, amounts, note) => [label, ...streams.map((_, i) => amounts[i] == null ? null : amounts[i]), null, note || null];
  const blank = () => [null];
  const serial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  const aoa = [
    ['', ...streams, 'Total', 'Notes'],
    ['Income'],
    row('Rental Income', [null, 26000]),
    row('Business Income (if self-employed or from a small business)', [null, null, null, 4800]),
    row('Allowances, Earnings, Tips, Honoraria', [null, null, 2150]),
    blank(),
    ['PTR Deductions'],
    row('Cost of Managing Tax Affairs (e.g., tax agent fees)', [null, null, null, null, null, 330]),
    row('Other Work-Related Expenses (e.g., tools, equipment, software)', [null, null, 212.5]),
    row('Home Office Expenses (e.g., electricity, internet, phone)', [null, null, null, 187.59]),
    blank(),
    ['CTR Deductions'],
    row('Software Subscriptions', [null, null, null, null, 648]),
    blank(),
    ['Investment Property Deductions'],
    row('Council Rates', [null, 877.09]),
    row('Insurance: Landlord Insurance', [null, 1430.4]),
    row('Interest on Loans (e.g., mortgage interest)', [null, 9870.21]),
    row('Repairs and Maintenance: General Repairs', [null, 1265]),
    blank(),
    ['Other'],
    blank(),
    ['Key Dates'],
    ['ABC shares sold', serial(2026, 7, 6)],
    ['Sample Street contract signed', serial(2026, 11, 14)],
    blank(),
    ['Surplus'],
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'OVERVIEW');
  return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
}
const workbookBytes = fs.existsSync(FIXTURE_XLSX) ? fs.readFileSync(FIXTURE_XLSX) : syntheticWorkbook();
const assumptionsFixture = fs.readFileSync(path.join(__dirname, 'fixtures', 'assumptions.json'), 'utf8');
/* a refund variant: the same sheet with enough PAYG to land on the green side */
const refundFixture = JSON.stringify({ ...JSON.parse(assumptionsFixture), paygWithheld: 40000, paygPerPay: 1150, paysLeft: 18 });

const stressFixture = JSON.stringify({ ...JSON.parse(assumptionsFixture), salary: 3000000, paygWithheld: 0, paygPerPay: 0, paysLeft: 0, help: { has: true, balance: 500000 }, hospitalCover: false,
  cgt: { carriedLoss: 2500, events: [
    { id: 'wbc-shares', asset: 'ABC shares', bought: '2025-03-01', costBase: 10000, sold: '2026-07-06', proceeds: 13775, note: '' },
    { id: 'short', asset: 'Some very long managed fund holding name that goes on', bought: '2026-05-01', costBase: 5000, sold: '2026-08-01', proceeds: 4000, note: '' },
    { id: 'noproc', asset: 'Unit 12/345 Example Parade, Somewhere Else QLD 4000', bought: '2020-01-15', costBase: 650000, sold: '2026-09-30', proceeds: null, note: '' } ] } });
const ITEMS = {
  root: { id: 'root1', name: 'Tax27', folder: { childCount: 3 }, webUrl: 'https://onedrive.live.com/mock/Tax27' },
  workbook: { id: 'wb1', name: 'PTR Calculations 27.xlsx', file: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }, size: workbookBytes.length,
    lastModifiedDateTime: '2026-10-09T22:14:00Z', webUrl: 'https://onedrive.live.com/mock/PTR', '@microsoft.graph.downloadUrl': 'https://graph.microsoft.com/mock/download/wb1' },
  inbox: { id: 'inbox1', name: 'Inbox', folder: { childCount: 1 }, webUrl: 'https://onedrive.live.com/mock/Inbox' },
  receipts: { id: 'rec1', name: 'Receipts', folder: { childCount: 12 }, webUrl: 'https://onedrive.live.com/mock/Receipts' },
};
const json = (route, status, body) => route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(body) });
const notFound = (route) => json(route, 404, { error: { code: 'itemNotFound', message: 'The resource could not be found.' } });
const mockGraph = (fixture) => (route) => {
  const p = decodeURIComponent(new URL(route.request().url()).pathname);
  const as = { id: 'as1', name: 'assumptions.json', file: { mimeType: 'application/json' }, size: fixture.length, eTag: '"etag-as1"', '@microsoft.graph.downloadUrl': 'https://graph.microsoft.com/mock/download/as1' };
  if (p.startsWith('/mock/download/wb1')) return route.fulfill({ status: 200, contentType: 'application/octet-stream', body: workbookBytes });
  if (p.startsWith('/mock/download/as1')) return route.fulfill({ status: 200, contentType: 'application/json', body: fixture });
  if (/^\/v1\.0\/me\/drive\/root:\/Personal Documents\/Tax\/Tax27$/.test(p)) return json(route, 200, ITEMS.root);
  if (p === '/v1.0/me/drive/items/root1/children') return json(route, 200, { value: [ITEMS.workbook, ITEMS.inbox, ITEMS.receipts] });
  if (p === '/v1.0/me/drive/items/inbox1:/assumptions.json') return fixture ? json(route, 200, as) : notFound(route);
  return notFound(route);
};

/* ----- in-page measurement (serialised into the browser) ----- */
function pageProbe(scope) {
  const VW = document.documentElement.clientWidth, VH = window.innerHeight;
  const root = scope ? document.querySelector(scope) : document;
  const vis = (el) => { const r = el.getBoundingClientRect(); const cs = getComputedStyle(el); return r.width > 0 && r.height > 0 && cs.visibility !== 'hidden' && cs.display !== 'none'; };
  const desc = (el) => {
    const t = (el.innerText || el.value || el.placeholder || el.getAttribute('aria-label') || '').replace(/\s+/g, ' ').trim().slice(0, 50);
    return `${el.tagName.toLowerCase()}${el.id ? '#' + el.id : ''}${el.className && typeof el.className === 'string' ? '.' + el.className.trim().split(/\s+/).join('.') : ''} "${t}"`;
  };
  const rr = (el) => { const r = el.getBoundingClientRect(); return { l: +r.left.toFixed(1), t: +r.top.toFixed(1), r: +r.right.toFixed(1), b: +r.bottom.toFixed(1), w: +r.width.toFixed(1), h: +r.height.toFixed(1) }; };
  const out = { VW, VH, overflow: [], clipped: [], overlaps: [], small: [], contrast: [] };
  const all = [...root.querySelectorAll('*')].filter(vis);
  /* anything poking past the viewport's right edge (the page has no horizontal scroll but a fixed/absolute child could still be cut) */
  for (const el of all) { const r = el.getBoundingClientRect(); if (r.right > VW + 0.5 || r.left < -0.5) out.overflow.push({ el: desc(el), rect: rr(el) }); }
  /* text that cannot fit its box: nowrap values wider than the box, or scrollWidth past clientWidth with overflow hidden */
  for (const el of all) {
    const cs = getComputedStyle(el);
    if (el.scrollWidth > el.clientWidth + 1 && (cs.overflowX === 'hidden' || cs.overflowX === 'clip' || cs.textOverflow === 'ellipsis')) out.clipped.push({ el: desc(el), scrollWidth: el.scrollWidth, clientWidth: el.clientWidth });
    if (cs.whiteSpace === 'nowrap' && el.parentElement) { const r = el.getBoundingClientRect(), p = el.parentElement.getBoundingClientRect(); if (r.right > p.right + 0.5 || r.left < p.left - 0.5) out.clipped.push({ el: desc(el), rect: rr(el), parent: rr(el.parentElement) }); }
  }
  /* kv rows and stream heads: the label must end before the value starts */
  for (const row of root.querySelectorAll('.kv, .stream-head, .progress-label, .cgt-foot')) {
    const kids = [...row.children].filter(vis);
    for (let i = 1; i < kids.length; i++) { const a = kids[i - 1].getBoundingClientRect(), b = kids[i].getBoundingClientRect(); if (a.right > b.left + 0.5 && a.bottom > b.top && b.bottom > a.top) out.overlaps.push({ row: desc(row), a: rr(kids[i - 1]), b: rr(kids[i]) }); }
  }
  /* tap targets */
  for (const el of root.querySelectorAll('button, a[href], input, select, label.toggle, [role="button"], [role="checkbox"]')) {
    if (!vis(el)) continue;
    if (el.closest('label.toggle') && el.tagName === 'INPUT') continue;   // the checkbox sits inside a 44px label
    const r = el.getBoundingClientRect();
    if (r.height < 44 || r.width < 44) out.small.push({ el: desc(el), w: +r.width.toFixed(1), h: +r.height.toFixed(1) });
  }
  /* contrast: text colour against the first opaque ancestor background (gradients reported by their CSS) */
  const parse = (c) => { const m = c.match(/rgba?\(([^)]+)\)/); if (!m) return null; const v = m[1].split(',').map(Number); return { r: v[0], g: v[1], b: v[2], a: v.length > 3 ? v[3] : 1 }; };
  const lum = (c) => { const f = (x) => { x /= 255; return x <= 0.03928 ? x / 12.92 : Math.pow((x + 0.055) / 1.055, 2.4); }; return 0.2126 * f(c.r) + 0.7152 * f(c.g) + 0.0722 * f(c.b); };
  const ratio = (a, b) => { const l1 = lum(a), l2 = lum(b); return +((Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05)).toFixed(2); };
  const hex = (s) => { const out = []; for (const m of s.matchAll(/rgba?\(([^)]+)\)/g)) { const v = m[1].split(',').map(Number); out.push({ r: v[0], g: v[1], b: v[2], a: 1 }); } for (const m of s.matchAll(/#([0-9a-f]{6})\b/gi)) out.push({ r: parseInt(m[1].slice(0, 2), 16), g: parseInt(m[1].slice(2, 4), 16), b: parseInt(m[1].slice(4, 6), 16), a: 1 }); return out; };
  const bgOf = (el) => {
    let e = el;
    while (e && e !== document.documentElement) {
      const cs = getComputedStyle(e);
      if (cs.backgroundImage && cs.backgroundImage !== 'none') { const hs = hex(cs.backgroundImage); if (hs.length) return { kind: 'gradient', colours: hs, css: cs.backgroundImage.slice(0, 60) }; }
      const c = parse(cs.backgroundColor); if (c && c.a >= 0.99) return { kind: 'solid', colours: [c] };
      e = e.parentElement;
    }
    const c = parse(getComputedStyle(document.body).backgroundColor); return { kind: 'body', colours: [c] };
  };
  const seen = new Set();
  const targets = root.querySelectorAll('.hero .eyebrow, .hero .big, .hero .sub, .hero .progress-label span, .pill, .kv > span, .kv b, .note, .sect, .muted, .eyebrow2, label.field, .toggle small, .btn, .tile small, .tile b, .stream small, .stream-name, .task b, .task small, input, select, .cgt-calc b, h2, .banner');
  for (const el of targets) {
    if (!vis(el)) continue;
    const cs = getComputedStyle(el);
    const fg = parse(cs.color); if (!fg) continue;
    const bg = bgOf(el);
    const ratios = bg.colours.map((c) => ratio(fg, c));
    const key = `${desc(el).replace(/ ".*"$/, '')}|${cs.color}|${bg.colours.map((c) => `${c.r},${c.g},${c.b}`).join('/')}`;
    if (seen.has(key)) continue; seen.add(key);
    const size = parseFloat(cs.fontSize), bold = parseInt(cs.fontWeight, 10) >= 700;
    const large = size >= 24 || (size >= 18.66 && bold);
    const min = Math.min(...ratios);
    out.contrast.push({ el: desc(el), fg: cs.color, bg: bg.kind === 'gradient' ? bg.css : `rgb(${bg.colours[0].r},${bg.colours[0].g},${bg.colours[0].b})`, font: `${size}px/${cs.fontWeight}`, ratio: min, aa: min >= (large ? 3 : 4.5), aaLarge: min >= 3 });
  }
  out.contrast.sort((a, b) => a.ratio - b.ratio);
  return out;
}

(async () => {
  const { pw, from } = loadPlaywright();
  console.log(`playwright from ${from}; workbook: ${fs.existsSync(FIXTURE_XLSX) ? 'ptr27.xlsx fixture' : 'in-memory workbook (ptr27.xlsx fixture not present)'}`);
  fs.mkdirSync(SHOTS, { recursive: true });
  const browser = await pw.chromium.launch({ channel: 'chrome', headless: true, args: ['--allow-file-access-from-files'] });
  const indexUrl = 'file://' + path.join(ROOT, 'index.html');
  const report = {};

  async function open(colorScheme, fixture) {
    const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, colorScheme, isMobile: true, hasTouch: true, serviceWorkers: 'block' });
    await ctx.addInitScript((seed) => {
      localStorage.setItem('tf.clientId', JSON.stringify('11111111-2222-3333-4444-555555555555'));
      localStorage.setItem('tf.tokens', JSON.stringify({ access: 'mock-access', refresh: 'mock-refresh', exp: Date.now() + 365 * 86400000 }));
      if (seed) localStorage.setItem('tf.assumptions', seed);
      localStorage.setItem('tf.accountantEmail', JSON.stringify('accountant@example.com'));
    }, fixture || '');
    await ctx.route('https://login.microsoftonline.com/**', (route) => json(route, 200, { access_token: 'mock-access', expires_in: 3600 }));
    await ctx.route('https://graph.microsoft.com/**', mockGraph(fixture));
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto(indexUrl);
    if (fixture) { await page.waitForSelector('.hero .big.refund, .hero .big.owing', { timeout: 20000 }); await page.waitForSelector('text=How it is worked out', { timeout: 20000 }); }
    else { await page.waitForSelector('text=Add your salary and PAYG to see your refund or bill', { timeout: 20000 }); }
    await page.waitForTimeout(300);
    return { ctx, page, errors };
  }
  /* screenshot of a .sect heading plus the card that follows it */
  async function shotSection(page, title, file) {
    const box = await page.evaluate((t) => {
      const s = [...document.querySelectorAll('.sect')].find((e) => e.textContent.trim() === t);
      if (!s) return null;
      const a = s.getBoundingClientRect(), b = s.nextElementSibling.getBoundingClientRect();
      return { x: 0, y: a.top + window.scrollY - 4, width: 390, height: b.bottom - a.top + 12 };
    }, title);
    if (box) await page.screenshot({ path: path.join(SHOTS, file), fullPage: true, clip: box });
    return box;
  }

  for (const scheme of ['light', 'dark']) {
    const { ctx, page, errors } = await open(scheme, assumptionsFixture);
    const r = report[`position-${scheme}`] = { errors };
    r.probe = await page.evaluate(pageProbe, null);
    r.sections = {};
    r.sections.chain = await shotSection(page, 'How it is worked out', `m-chain-${scheme}.png`);
    r.sections.cgt = await shotSection(page, 'Capital gains', `m-cgt-${scheme}.png`);
    r.sections.todo = await shotSection(page, 'To do', `m-todo-${scheme}.png`);
    r.sections.hero = await page.$eval('.hero', (el) => { const b = el.getBoundingClientRect(); return { x: 0, y: b.top + window.scrollY - 4, width: 390, height: b.height + 8 }; });
    await page.screenshot({ path: path.join(SHOTS, `m-hero-${scheme}.png`), fullPage: true, clip: r.sections.hero });
    /* the bottom of the page: the Assumptions button in the actions row */
    r.btnAssume = await page.$eval('#btn-assume', (el) => { const b = el.getBoundingClientRect(); return { w: b.width, h: b.height, text: el.textContent, disabled: el.disabled }; });
    await page.$eval('#btn-assume', (el) => el.scrollIntoView({ block: 'center' }));
    await page.screenshot({ path: path.join(SHOTS, `m-actions-${scheme}.png`), fullPage: false });
    /* the kv rows in the chain card, with label and value widths */
    r.chainRows = await page.evaluate(() => {
      const s = [...document.querySelectorAll('.sect')].find((e) => e.textContent.trim() === 'How it is worked out');
      return [...s.nextElementSibling.querySelectorAll('.kv')].map((kv) => { const a = kv.children[0].getBoundingClientRect(), b = kv.children[1].getBoundingClientRect(); return `${kv.children[0].textContent} [${a.width.toFixed(0)}w ${a.height.toFixed(0)}h] | ${kv.children[1].textContent} [${b.width.toFixed(0)}w] gap ${(b.left - a.right).toFixed(0)}`; });
    });

    /* the Assumptions dialog */
    await page.click('#btn-assume');
    await page.waitForSelector('dialog#assume[open]');
    await page.waitForTimeout(200);
    const dlg = r.dialog = {};
    dlg.style = await page.$eval('dialog#assume', (el) => { const cs = getComputedStyle(el); const b = el.getBoundingClientRect(); return { overflowY: cs.overflowY, maxHeight: cs.maxHeight, rect: { l: b.left, t: b.top, r: b.right, b: b.bottom, w: b.width, h: b.height }, scrollHeight: el.scrollHeight, clientHeight: el.clientHeight, scrollsInternally: el.scrollHeight > el.clientHeight && /auto|scroll/.test(cs.overflowY) }; });
    dlg.focusOnOpen = await page.evaluate(() => { const a = document.activeElement; return a ? `${a.tagName.toLowerCase()}${a.id ? '#' + a.id : ''}` : null; });
    dlg.bodyLocked = await page.evaluate(() => ({ htmlOverflow: getComputedStyle(document.documentElement).overflow, bodyOverflow: getComputedStyle(document.body).overflow, scrollY: window.scrollY }));
    dlg.probeTop = await page.evaluate(pageProbe, 'dialog#assume');
    await page.screenshot({ path: path.join(SHOTS, `m-dialog-top-${scheme}.png`), fullPage: false });
    /* the middle: deductions + toggles + HELP */
    await page.$eval('#a-std', (el) => el.scrollIntoView({ block: 'start' }));
    await page.waitForTimeout(100);
    await page.screenshot({ path: path.join(SHOTS, `m-dialog-mid-${scheme}.png`), fullPage: false });
    /* HELP on: the balance field appears; family: spouse income appears */
    await page.selectOption('#a-family', 'family');
    await page.waitForTimeout(100);
    await page.$eval('#a-help', (el) => el.scrollIntoView({ block: 'start' }));
    await page.screenshot({ path: path.join(SHOTS, `m-dialog-help-${scheme}.png`), fullPage: false });
    await page.selectOption('#a-family', 'single');
    /* the bottom: CGT rows + Save/Cancel */
    await page.$eval('dialog#assume', (el) => { el.scrollTop = el.scrollHeight; });
    await page.waitForTimeout(100);
    await page.screenshot({ path: path.join(SHOTS, `m-dialog-bottom-${scheme}.png`), fullPage: false });
    dlg.saveCancel = await page.evaluate(() => ['a-cancel', 'a-save'].map((id) => { const el = document.getElementById(id); const b = el.getBoundingClientRect(); return { id, l: +b.left.toFixed(1), t: +b.top.toFixed(1), r: +b.right.toFixed(1), b: +b.bottom.toFixed(1), w: +b.width.toFixed(1), h: +b.height.toFixed(1), inViewport: b.top >= 0 && b.bottom <= window.innerHeight }; }));
    dlg.probeBottom = await page.evaluate(pageProbe, 'dialog#assume');
    dlg.dateInputs = await page.$$eval('dialog#assume input[type="date"]', (els) => els.map((el) => { const b = el.getBoundingClientRect(); const cs = getComputedStyle(el); return { value: el.value, w: +b.width.toFixed(1), h: +b.height.toFixed(1), padding: cs.padding, fontSize: cs.fontSize, scrollWidth: el.scrollWidth, clientWidth: el.clientWidth }; }));
    /* the date input zoom: the first row's Sold date (the indicator sits at its right edge) */
    const sold = await page.$('dialog#assume .cgt-row input[type="date"]');
    if (sold) { await sold.scrollIntoViewIfNeeded(); const b = await sold.boundingBox(); await page.screenshot({ path: path.join(SHOTS, `m-dateinput-${scheme}.png`), clip: { x: b.x - 10, y: b.y - 30, width: 390 - b.x, height: b.height + 60 } }); }
    /* Escape closes, and the page behind did not scroll sideways */
    await page.keyboard.press('Escape');
    await page.waitForTimeout(100);
    dlg.escapeCloses = await page.$eval('dialog#assume', (el) => !el.open);
    dlg.pageAfter = await page.evaluate(() => ({ scrollWidth: document.documentElement.scrollWidth, clientWidth: document.documentElement.clientWidth }));
    await ctx.close();
  }

  /* the refund side of the hero, both schemes */
  for (const scheme of ['light', 'dark']) {
    const { ctx, page } = await open(scheme, refundFixture);
    const r = report[`refund-${scheme}`] = {};
    r.hero = await page.$eval('.hero', (el) => el.innerText.replace(/\s+/g, ' ').trim());
    const box = await page.$eval('.hero', (el) => { const b = el.getBoundingClientRect(); return { x: 0, y: b.top + window.scrollY - 4, width: 390, height: b.height + 8 }; });
    await page.screenshot({ path: path.join(SHOTS, `m-hero-refund-${scheme}.png`), fullPage: true, clip: box });
    r.contrast = (await page.evaluate(pageProbe, '.hero')).contrast;
    r.resultRow = await page.evaluate(() => { const kv = document.querySelector('.kv.result'); if (!kv) return null; const cs = getComputedStyle(kv.querySelector('b')); return { text: kv.innerText.replace(/\s+/g, ' '), color: cs.color }; });
    await shotSection(page, 'How it is worked out', `m-chain-refund-${scheme}.png`);
    await ctx.close();
  }

  /* stress: long values and labels */
  for (const scheme of ['light', 'dark']) {
    const { ctx, page, errors } = await open(scheme, stressFixture);
    const r = report[`stress-${scheme}`] = { errors };
    r.probe = await page.evaluate(pageProbe, null);
    r.hero = await page.$eval('.hero', (el) => el.innerText.replace(/\s+/g, ' ').trim());
    const box = await page.$eval('.hero', (el) => { const b = el.getBoundingClientRect(); return { x: 0, y: b.top + window.scrollY - 4, width: 390, height: b.height + 8 }; });
    await page.screenshot({ path: path.join(SHOTS, `m-hero-stress-${scheme}.png`), fullPage: true, clip: box });
    await shotSection(page, 'How it is worked out', `m-chain-stress-${scheme}.png`);
    await shotSection(page, 'Capital gains', `m-cgt-stress-${scheme}.png`);
    await page.click('#btn-assume');
    await page.waitForSelector('dialog#assume[open]');
    await page.$eval('dialog#assume', (el) => { el.scrollTop = el.scrollHeight; });
    await page.waitForTimeout(100);
    await page.screenshot({ path: path.join(SHOTS, `m-dialog-stress-bottom-${scheme}.png`), fullPage: false });
    r.dialogProbe = await page.evaluate(pageProbe, 'dialog#assume');
    await ctx.close();
  }
  /* the sheet opened from the empty state: the seeded rows */
  {
    const { ctx, page } = await open('light', '');
    const r = report['empty-sheet-light'] = {};
    await page.click('text=Add your salary and PAYG to see your refund or bill >> xpath=.. >> button');
    await page.waitForSelector('dialog#assume[open]');
    await page.waitForTimeout(150);
    r.focusOnOpen = await page.evaluate(() => { const a = document.activeElement; return a ? `${a.tagName.toLowerCase()}${a.id ? '#' + a.id : ''}` : null; });
    r.sheet = await page.$eval('dialog#assume', (el) => ({ rows: el.querySelectorAll('.cgt-row').length, std: el.querySelector('#a-std').checked, calc: [...el.querySelectorAll('.cgt-calc')].map((c) => c.textContent.trim()), assets: [...el.querySelectorAll('.cgt-row input[type="text"]:first-of-type')].map((i) => i.value) }));
    r.probe = await page.evaluate(pageProbe, 'dialog#assume');
    await page.screenshot({ path: path.join(SHOTS, 'm-empty-sheet-top.png'), fullPage: false });
    await page.$eval('dialog#assume', (el) => { el.scrollTop = el.scrollHeight; });
    await page.waitForTimeout(100);
    await page.screenshot({ path: path.join(SHOTS, 'm-empty-sheet-bottom.png'), fullPage: false });
    await ctx.close();
  }

  /* the no-assumptions view */
  for (const scheme of ['light', 'dark']) {
    const { ctx, page } = await open(scheme, '');
    const r = report[`empty-${scheme}`] = {};
    r.probe = await page.evaluate(pageProbe, null);
    const box = await page.evaluate(() => { const b = [...document.querySelectorAll('.card b')].find((e) => /Add your salary/.test(e.textContent)); const c = b.closest('.card').getBoundingClientRect(); return { x: 0, y: c.top + window.scrollY - 4, width: 390, height: c.height + 8 }; });
    await page.screenshot({ path: path.join(SHOTS, `m-empty-prompt-${scheme}.png`), fullPage: true, clip: box });
    await shotSection(page, 'Indicative tax', `m-empty-indicative-${scheme}.png`);
    await ctx.close();
  }

  await browser.close();
  fs.writeFileSync(path.join(SHOTS, 'measure.json'), JSON.stringify(report, null, 2));
  /* a readable digest */
  for (const [name, r] of Object.entries(report)) {
    console.log(`\n== ${name}`);
    if (r.errors && r.errors.length) console.log('  page errors:', r.errors.join(' | '));
    const digest = (p, label) => {
      if (!p) return;
      console.log(`  [${label}] viewport ${p.VW}x${p.VH}`);
      console.log(`   overflow (${p.overflow.length}):`, p.overflow.slice(0, 12).map((o) => `${o.el} ${JSON.stringify(o.rect)}`).join('\n     ') || 'none');
      console.log(`   clipped (${p.clipped.length}):`, p.clipped.slice(0, 12).map((o) => `${o.el} ${JSON.stringify(o.rect || o)}`).join('\n     ') || 'none');
      console.log(`   overlaps (${p.overlaps.length}):`, p.overlaps.slice(0, 12).map((o) => `${o.row} a=${JSON.stringify(o.a)} b=${JSON.stringify(o.b)}`).join('\n     ') || 'none');
      console.log(`   small targets (${p.small.length}):`, p.small.slice(0, 30).map((o) => `${o.el} ${o.w}x${o.h}`).join('\n     ') || 'none');
      console.log(`   contrast under 4.5 (${p.contrast.filter((c) => c.ratio < 4.5).length}):`, p.contrast.filter((c) => c.ratio < 4.5).slice(0, 30).map((c) => `${c.ratio} ${c.el} fg ${c.fg} on ${c.bg} ${c.font}${c.aa ? ' (large text, passes AA)' : ''}`).join('\n     ') || 'none');
    };
    digest(r.probe, 'page');
    if (r.chainRows) console.log('  chain rows:\n     ' + r.chainRows.join('\n     '));
    if (r.btnAssume) console.log('  #btn-assume:', JSON.stringify(r.btnAssume));
    if (r.dialog) {
      console.log('  dialog:', JSON.stringify(r.dialog.style), JSON.stringify(r.dialog.bodyLocked));
      console.log('  save/cancel at bottom:', JSON.stringify(r.dialog.saveCancel));
      console.log('  date inputs:', JSON.stringify(r.dialog.dateInputs));
      console.log('  escape closes:', r.dialog.escapeCloses, 'page after:', JSON.stringify(r.dialog.pageAfter));
      digest(r.dialog.probeTop, 'dialog top');
      digest(r.dialog.probeBottom, 'dialog bottom');
    }
    if (r.dialog && r.dialog.focusOnOpen) console.log('  focus on open:', r.dialog.focusOnOpen);
    if (r.focusOnOpen) console.log('  focus on open:', r.focusOnOpen, 'sheet:', JSON.stringify(r.sheet));
    if (r.dialogProbe) digest(r.dialogProbe, 'dialog bottom');
    if (r.hero && !r.contrast) console.log('  hero:', r.hero);
    if (r.hero && r.contrast) console.log('  hero:', r.hero, '\n  result row:', JSON.stringify(r.resultRow), '\n  hero contrast:', r.contrast.map((c) => `${c.ratio} ${c.el} ${c.fg}`).join(' | '));
  }
})().catch((e) => { console.error('MEASURE ERROR', e); process.exit(1); });
