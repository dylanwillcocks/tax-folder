'use strict';

/* Renders the Position tab at phone size with everything mocked, and saves screenshots.
   Run: /usr/local/opt/node/bin/node tests/render.js
   - opens index.html from a file:// URL in headless Chrome (playwright-core driving the installed Google Chrome, no download)
   - mocks login.microsoftonline.com and graph.microsoft.com: the Tax27 folder, the workbook (bytes from the ptr27.xlsx fixture
     when it exists, otherwise a workbook built in memory with SheetJS) and an Inbox folder holding assumptions.json
   - seeds localStorage (client id, a token that will not expire, the assumptions cache)
   - screenshots tests/shots/position-light.png, position-dark.png and assumptions-dialog.png at 390x844
   - reports document.documentElement.scrollWidth against clientWidth for each (horizontal scroll = FAIL) */
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

/* ----- workbook bytes ----- */
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
const workbookSource = fs.existsSync(FIXTURE_XLSX) ? 'ptr27.xlsx fixture' : 'workbook built in memory (ptr27.xlsx fixture not present)';
const assumptionsFixture = fs.readFileSync(path.join(__dirname, 'fixtures', 'assumptions.json'), 'utf8');

/* ----- Graph mocks ----- */
const ITEMS = {
  root: { id: 'root1', name: 'Tax27', folder: { childCount: 3 }, webUrl: 'https://onedrive.live.com/mock/Tax27' },
  workbook: { id: 'wb1', name: 'PTR Calculations 27.xlsx', file: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }, size: workbookBytes.length,
    lastModifiedDateTime: '2026-10-09T22:14:00Z', webUrl: 'https://onedrive.live.com/mock/PTR', '@microsoft.graph.downloadUrl': 'https://graph.microsoft.com/mock/download/wb1' },
  inbox: { id: 'inbox1', name: 'Inbox', folder: { childCount: 1 }, webUrl: 'https://onedrive.live.com/mock/Inbox' },
  receipts: { id: 'rec1', name: 'Receipts', folder: { childCount: 12 }, webUrl: 'https://onedrive.live.com/mock/Receipts' },
  assumptions: { id: 'as1', name: 'assumptions.json', file: { mimeType: 'application/json' }, size: assumptionsFixture.length, eTag: '"etag-as1"',
    '@microsoft.graph.downloadUrl': 'https://graph.microsoft.com/mock/download/as1' },
};
const json = (route, status, body) => route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(body) });
const notFound = (route) => json(route, 404, { error: { code: 'itemNotFound', message: 'The resource could not be found.' } });
const graphLog = [];
function mockGraph(route) {
  const url = new URL(route.request().url());
  const p = decodeURIComponent(url.pathname);
  graphLog.push(`${route.request().method()} ${p}`);
  if (p.startsWith('/mock/download/wb1')) return route.fulfill({ status: 200, contentType: 'application/octet-stream', body: workbookBytes });
  if (p.startsWith('/mock/download/as1')) return route.fulfill({ status: 200, contentType: 'application/json', body: assumptionsFixture });
  if (/^\/v1\.0\/me\/drive\/root:\/Personal Documents\/Tax\/Tax27$/.test(p)) return json(route, 200, ITEMS.root);
  if (p === '/v1.0/me/drive/items/root1/children') return json(route, 200, { value: [ITEMS.workbook, ITEMS.inbox, ITEMS.receipts] });
  if (p === '/v1.0/me/drive/items/inbox1:/assumptions.json') return json(route, 200, ITEMS.assumptions);
  if (p === '/v1.0/me/drive/items/inbox1:/assumptions.json:/content') return json(route, 200, { ...ITEMS.assumptions, eTag: '"etag-as2"' });
  if (p.startsWith('/v1.0/me/drive/items/inbox1:/')) return notFound(route);     // inbox.json, profile.md, briefing.json, questions.json
  return notFound(route);
}

/* ----- the run ----- */
(async () => {
  const { pw, from } = loadPlaywright();
  console.log(`playwright from ${from}; workbook: ${workbookSource}`);
  fs.mkdirSync(SHOTS, { recursive: true });
  const browser = await pw.chromium.launch({ channel: 'chrome', headless: true, args: ['--allow-file-access-from-files'] });
  const results = [];
  let failed = false;
  const indexUrl = 'file://' + path.join(ROOT, 'index.html');

  async function openPosition(colorScheme) {
    const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, colorScheme, isMobile: true, hasTouch: true, serviceWorkers: 'block' });
    await ctx.addInitScript((seed) => {
      localStorage.setItem('tf.clientId', JSON.stringify('11111111-2222-3333-4444-555555555555'));
      localStorage.setItem('tf.tokens', JSON.stringify({ access: 'mock-access', refresh: 'mock-refresh', exp: Date.now() + 365 * 86400000 }));
      localStorage.setItem('tf.assumptions', seed);
      localStorage.setItem('tf.accountantEmail', JSON.stringify('accountant@example.com'));
    }, assumptionsFixture);
    await ctx.route('https://login.microsoftonline.com/**', (route) => json(route, 200, { access_token: 'mock-access', expires_in: 3600 }));
    await ctx.route('https://graph.microsoft.com/**', mockGraph);
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    page.on('console', (m) => { if (m.type() === 'error' && !/Failed to load resource/.test(m.text())) errors.push(m.text()); });   // the mocked 404s are expected
    await page.goto(indexUrl);
    await page.waitForSelector('.hero .big.refund, .hero .big.owing', { timeout: 20000 });
    await page.waitForSelector('text=How it is worked out', { timeout: 20000 });
    return { ctx, page, errors };
  }
  async function measure(page, name, extraSelector) {
    const m = await page.evaluate((sel) => {
      const el = sel ? document.querySelector(sel) : null;
      return { scrollWidth: document.documentElement.scrollWidth, clientWidth: document.documentElement.clientWidth,
        extra: el ? { scrollWidth: el.scrollWidth, clientWidth: el.clientWidth } : null };
    }, extraSelector || null);
    const ok = m.scrollWidth <= m.clientWidth && (!m.extra || m.extra.scrollWidth <= m.extra.clientWidth);
    if (!ok) failed = true;
    results.push(`${ok ? 'PASS' : 'FAIL'} ${name}: scrollWidth ${m.scrollWidth} vs clientWidth ${m.clientWidth}${m.extra ? ` (dialog ${m.extra.scrollWidth} vs ${m.extra.clientWidth})` : ''}`);
  }

  for (const scheme of ['light', 'dark']) {
    const { ctx, page, errors } = await openPosition(scheme);
    const hero = await page.$eval('.hero', (el) => el.innerText.replace(/\s+/g, ' ').trim());
    const file = path.join(SHOTS, `position-${scheme}.png`);
    await page.screenshot({ path: file, fullPage: true });
    await page.screenshot({ path: path.join(SHOTS, `position-${scheme}-hero.png`), clip: { x: 0, y: 0, width: 390, height: 560 } });   // the first screen only
    await measure(page, `position-${scheme}`);
    results.push(`     hero: ${hero}`);
    if (errors.length) { failed = true; results.push(`     page errors: ${errors.join(' | ')}`); }
    if (scheme === 'light') {
      const chain = await page.$$eval('.kv', (els) => els.map((e) => e.innerText.replace(/\s+/g, ' ').trim()));
      results.push(`     chain rows: ${chain.filter((t) => /Taxable income|Income tax|Medicare levy|HELP|Estimated/.test(t)).join(' | ')}`);
      const pills = await page.$$eval('.pills .pill', (els) => els.map((e) => e.textContent.trim()));
      results.push(`     pills: ${pills.join(' | ')}`);
      const tasks = await page.$$eval('.task b', (els) => els.map((e) => e.textContent.trim()));
      results.push(`     tasks: ${tasks.join(' | ')}`);
      const mail = await page.$eval('a[href^="mailto:"]', (a) => decodeURIComponent(a.getAttribute('href')));
      results.push(`     email has chain: ${/Estimated (refund|amount owing)/.test(mail) && /Assumptions used/.test(mail)}`);
      // the Assumptions sheet
      await page.click('#btn-assume');
      await page.waitForSelector('dialog#assume[open]');
      await page.screenshot({ path: path.join(SHOTS, 'assumptions-dialog.png'), fullPage: false });
      await measure(page, 'assumptions-dialog', 'dialog#assume');
      const sheet = await page.$eval('dialog#assume', (el) => ({ rows: el.querySelectorAll('.cgt-row').length, payg: el.querySelector('#a-payg-total').textContent, std: el.querySelector('#a-std').checked, calc: [...el.querySelectorAll('.cgt-calc')].map((c) => c.textContent.trim()) }));
      results.push(`     sheet: ${JSON.stringify(sheet)}`);
      // the dialog scrolled to the bottom (the CGT rows) for a second look
      await page.$eval('dialog#assume', (el) => { el.scrollTop = el.scrollHeight; });
      await page.screenshot({ path: path.join(SHOTS, 'assumptions-dialog-bottom.png'), fullPage: false });
      await measure(page, 'assumptions-dialog-bottom', 'dialog#assume');
    }
    await ctx.close();
  }

  // the no-assumptions view: the empty-state card and the round-1 indicative card must still be there
  {
    const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, colorScheme: 'light', serviceWorkers: 'block' });
    await ctx.addInitScript(() => {
      localStorage.setItem('tf.clientId', JSON.stringify('11111111-2222-3333-4444-555555555555'));
      localStorage.setItem('tf.tokens', JSON.stringify({ access: 'mock-access', refresh: 'mock-refresh', exp: Date.now() + 365 * 86400000 }));
    });
    await ctx.route('https://login.microsoftonline.com/**', (route) => json(route, 200, {}));
    await ctx.route('https://graph.microsoft.com/**', (route) => {
      const p = decodeURIComponent(new URL(route.request().url()).pathname);
      if (p === '/v1.0/me/drive/items/inbox1:/assumptions.json') return notFound(route);
      return mockGraph(route);
    });
    const page = await ctx.newPage();
    await page.goto(indexUrl);
    await page.waitForSelector('text=Indicative tax', { timeout: 20000 });
    await page.waitForSelector('text=Add your salary and PAYG to see your refund or bill', { timeout: 20000 });
    const hasChain = await page.$('text=How it is worked out');
    results.push(`${hasChain ? 'FAIL' : 'PASS'} no-assumptions view shows the round-1 card and the prompt only`);
    if (hasChain) failed = true;
    await page.screenshot({ path: path.join(SHOTS, 'position-empty.png'), fullPage: true });
    await measure(page, 'position-empty');
    await ctx.close();
  }

  await browser.close();
  console.log(results.join('\n'));
  console.log(`graph calls: ${[...new Set(graphLog)].join(', ')}`);
  console.log(failed ? '\nRENDER FAIL' : '\nRENDER OK');
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error('RENDER ERROR', e); process.exit(1); });
