'use strict';

/* Tax Folder: browse a OneDrive folder and file receipts into it.
   Talks straight to Microsoft Graph from the browser. Sign-in is OAuth authorization
   code + PKCE, so there is no server and no secret. */

const GRAPH = 'https://graph.microsoft.com/v1.0';
const AUTH = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
const SCOPES = 'Files.ReadWrite offline_access';
const DEFAULTS = { rootPath: 'Personal Documents/Tax/Tax27', receiptsFolder: 'Receipts', workbook: 'PTR Calculations 27.xlsx' };
const LOCK_AFTER_MS = 2 * 60 * 1000;
const MAX_UPLOAD = 100 * 1024 * 1024;

const $ = (id) => document.getElementById(id);

/* ---------- storage (wrapped: localStorage can throw in private windows) ---------- */
const ls = {
  get(k, d = null) {
    try { const v = localStorage.getItem('tf.' + k); return v == null ? d : JSON.parse(v); } catch { return d; }
  },
  set(k, v) { try { localStorage.setItem('tf.' + k, JSON.stringify(v)); } catch { /* ignore */ } },
  del(k) { try { localStorage.removeItem('tf.' + k); } catch { /* ignore */ } },
};
const cfg = () => ({
  rootPath: (ls.get('rootPath') || DEFAULTS.rootPath).replace(/^\/+|\/+$/g, ''),
  receiptsFolder: ls.get('receiptsFolder') || DEFAULTS.receiptsFolder,
  workbook: ls.get('workbook') || DEFAULTS.workbook,
  accountantName: ls.get('accountantName') || 'Brayden',
  accountantEmail: ls.get('accountantEmail') || '',
  clientId: ls.get('clientId') || '',
});

/* ---------- small DOM helper (text goes in as text nodes, never HTML) ---------- */
function h(tag, props = {}, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else if (v === true) el.setAttribute(k, '');
    else el.setAttribute(k, v);
  }
  for (const kid of kids.flat()) if (kid != null && kid !== false) el.append(kid);
  return el;
}
const svg = (inner) => {
  const wrap = document.createElement('span');
  wrap.innerHTML = `<svg viewBox="0 0 24 24" aria-hidden="true">${inner}</svg>`; // static markup only
  return wrap.firstChild;
};
const FOLDER_SVG = '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>';

/* ---------- auth: PKCE ---------- */
class AuthError extends Error {}
const redirectUri = () => new URL('./', location.href).href;
const b64url = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const rand = (n) => b64url(crypto.getRandomValues(new Uint8Array(n)));

async function signIn() {
  const verifier = rand(48);
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = rand(16);
  ls.set('pkce', { verifier, state, redirect: redirectUri() });
  location.assign(`${AUTH}/authorize?` + new URLSearchParams({
    client_id: cfg().clientId, response_type: 'code', redirect_uri: redirectUri(),
    response_mode: 'query', scope: SCOPES, state,
    code_challenge: challenge, code_challenge_method: 'S256',
  }));
}

async function handleRedirect() {
  const q = new URLSearchParams(location.search);
  if (!q.has('code') && !q.has('error')) return;
  const pk = ls.get('pkce'); ls.del('pkce');
  history.replaceState(null, '', redirectUri());
  if (q.has('error')) throw new Error((q.get('error_description') || q.get('error')).split(/\r?\n/)[0]);
  if (!pk || pk.state !== q.get('state')) throw new Error('Sign-in was interrupted. Please try again.');
  await tokenRequest({ grant_type: 'authorization_code', code: q.get('code'), redirect_uri: pk.redirect, code_verifier: pk.verifier });
}

async function tokenRequest(params) {
  const res = await fetch(`${AUTH}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: cfg().clientId, scope: SCOPES, ...params }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const e = new Error((data.error_description || data.error || 'Sign-in failed').split(/\r?\n/)[0]);
    e.code = data.error;
    throw e;
  }
  const old = ls.get('tokens') || {};
  ls.set('tokens', {
    access: data.access_token,
    refresh: data.refresh_token || old.refresh,
    exp: Date.now() + data.expires_in * 1000,
  });
  return data.access_token;
}

let refreshing = null;
async function getToken() {
  const t = ls.get('tokens');
  if (!t) throw new AuthError('Signed out');
  if (t.exp - Date.now() > 60_000) return t.access;
  if (!t.refresh) throw new AuthError('Session expired');
  if (!refreshing) {
    refreshing = tokenRequest({ grant_type: 'refresh_token', refresh_token: t.refresh })
      .catch((e) => {
        if (e.code === 'invalid_grant' || e.code === 'interaction_required') { ls.del('tokens'); throw new AuthError('Session expired'); }
        throw e;
      })
      .finally(() => { refreshing = null; });
  }
  return refreshing;
}

/* ---------- Graph ---------- */
async function graph(path, { method = 'GET', body, headers = {} } = {}, retry = true) {
  const token = await getToken();
  const res = await fetch(path.startsWith('http') ? path : GRAPH + path, {
    method, body, headers: { Authorization: `Bearer ${token}`, ...headers },
  });
  if (res.status === 401 && retry) {
    const t = ls.get('tokens');
    if (t) ls.set('tokens', { ...t, exp: 0 });
    return graph(path, { method, body, headers }, false);
  }
  if (!res.ok) {
    let msg = res.statusText || `HTTP ${res.status}`;
    try { msg = (await res.json()).error.message || msg; } catch { /* keep default */ }
    const e = new Error(msg); e.status = res.status; throw e;
  }
  return res.status === 204 ? null : res.json();
}
const encodePath = (p) => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');

async function listChildren(id) {
  const base = `/me/drive/items/${id}/children?$top=200`;
  const items = [];
  let next = base + '&$expand=thumbnails';
  let first = true;
  while (next) {
    let page;
    try { page = await graph(next); }
    catch (e) {
      if (first && !(e instanceof AuthError) && e.status !== 401 && e.status !== 404) { next = base; first = false; continue; }
      throw e;
    }
    first = false;
    items.push(...page.value);
    next = page['@odata.nextLink'];
  }
  return items;
}

/* ---------- state ---------- */
const state = {
  tab: 'home',
  rootItems: [],      // children of the root folder (workbook, Receipts, ...)
  rootError: '',
  position: null,     // parsed tax workbook snapshot
  positionError: '',
  positionBusy: false,
  stack: [],          // [{id, name, webUrl}], stack[0] is the configured root
  rootFolders: [],    // child folders of the root, for the "Save into" menu
  sort: ls.get('sort', 'name'),
  seq: 0,             // guards against out-of-order responses
  dirty: false,       // a receipt was saved, so the listing is stale
  pending: null,      // {blob, ext, name, previewUrl}
  saving: false,
};
const current = () => state.stack[state.stack.length - 1];

/* ---------- views ---------- */
function show(name) {
  state.tab = name;
  for (const v of ['setup', 'signin', 'home', 'files', 'add']) $('view-' + v).hidden = v !== name;
  $('nav').hidden = !['home', 'files', 'add'].includes(name);
  for (const t of ['home', 'files', 'add']) {
    const b = $('tab-' + t);
    if (t === name) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current');
  }
  const showUp = name === 'files' && state.stack.length > 1;
  $('btn-up').hidden = !showUp;
  $('logo').hidden = showUp;
  $('title').textContent = name === 'files' && current() ? current().name
    : name === 'add' ? 'Add receipt'
    : name === 'home' ? `FY${fyEnd() % 100} tax position` : 'Tax Folder';
  window.scrollTo(0, 0);
}
function banner(msg) { const b = $('banner'); b.textContent = msg || ''; b.hidden = !msg; }
let toastTimer;
function toast(msg) {
  const t = $('toast'); t.textContent = msg; t.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, 4000);
}
function friendly(e) {
  if (e instanceof TypeError) return 'No connection. Check your signal and try again.';
  return e.message || 'Something went wrong.';
}
function needSignIn() { ls.del('tokens'); banner(''); show('signin'); }

/* ---------- files ---------- */
const fmtDate = (iso) => new Date(iso).toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: 'numeric' });
const fmtSize = (n) => n < 1024 ? `${n} B` : n < 1048576 ? `${Math.round(n / 1024)} KB` : `${(n / 1048576).toFixed(1)} MB`;
function fileKind(name) {
  const ext = (name.split('.').pop() || '').toLowerCase();
  if (ext === 'pdf') return ['PDF', 't-pdf'];
  if (['jpg', 'jpeg', 'png', 'heic', 'gif', 'webp'].includes(ext)) return [ext.toUpperCase().slice(0, 4), 't-img'];
  if (['xls', 'xlsx', 'xlsm', 'csv'].includes(ext)) return [ext.toUpperCase().slice(0, 4), 't-xls'];
  if (['doc', 'docx', 'txt', 'rtf'].includes(ext)) return [ext.toUpperCase().slice(0, 4), 't-doc'];
  return [ext && ext.length <= 4 && ext !== name.toLowerCase() ? ext.toUpperCase() : 'FILE', ''];
}
function relPath(item) {
  let p = item.parentReference && item.parentReference.path;
  if (!p) return '';
  try { p = decodeURIComponent(p); } catch { /* keep raw */ }
  p = p.replace(/^.*?root:\/?/, '');
  const root = cfg().rootPath;
  if (p === root) return state.stack[0] ? state.stack[0].name : '';
  return p.startsWith(root + '/') ? p.slice(root.length + 1) : p;
}

function sortItems(items, searching) {
  const byName = (a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
  const byNew = (a, b) => (b.lastModifiedDateTime || '').localeCompare(a.lastModifiedDateTime || '');
  const cmp = state.sort === 'new' ? byNew : byName;
  return [...items].sort((a, b) => {
    if (!searching && !!a.folder !== !!b.folder) return a.folder ? -1 : 1;
    return cmp(a, b);
  });
}

function renderList(items, { searching = false } = {}) {
  const list = $('list'), status = $('files-state');
  list.replaceChildren();
  if (!items.length) {
    status.replaceChildren(searching ? 'No matches in this folder.' : 'This folder is empty.');
    status.hidden = false; list.hidden = true; return;
  }
  status.hidden = true; list.hidden = false;
  for (const it of sortItems(items, searching)) {
    const sub = searching
      ? [relPath(it), it.folder ? '' : fmtSize(it.size || 0)].filter(Boolean).join(' · ')
      : it.folder
        ? `${it.folder.childCount} item${it.folder.childCount === 1 ? '' : 's'}`
        : [fmtSize(it.size || 0), it.lastModifiedDateTime && fmtDate(it.lastModifiedDateTime)].filter(Boolean).join(' · ');
    const thumb = it.thumbnails && it.thumbnails[0] && it.thumbnails[0].small && it.thumbnails[0].small.url;
    let badge;
    if (it.folder) badge = h('div', { class: 'badge folder' }, svg(FOLDER_SVG));
    else {
      const [label, cls] = fileKind(it.name);
      badge = h('div', { class: 'badge ' + cls }, thumb ? h('img', { src: thumb, alt: '', loading: 'lazy' }) : label);
    }
    const body = [badge, h('div', { class: 'meta' }, h('b', {}, it.name), h('small', {}, sub))];
    let row;
    if (it.folder) {
      row = h('button', { class: 'item', type: 'button', onclick: () => openFolder({ id: it.id, name: it.name, webUrl: it.webUrl }) }, ...body, h('span', { class: 'chev', 'aria-hidden': 'true' }, '›'));
    } else {
      row = h('a', { class: 'item', href: it.webUrl, target: '_blank', rel: 'noopener' }, ...body);
    }
    list.append(h('li', {}, row));
  }
}

function renderCrumbs() {
  const box = $('crumbs'); box.replaceChildren();
  state.stack.forEach((c, i) => {
    if (i) box.append(h('span', { 'aria-hidden': 'true' }, '›'));
    box.append(h('button', { type: 'button', onclick: () => goTo(i) }, c.name));
  });
  box.scrollLeft = box.scrollWidth;
  const cur = current();
  const link = $('btn-onedrive');
  link.hidden = !(cur && cur.webUrl);
  if (cur && cur.webUrl) link.href = cur.webUrl;
  $('btn-sort').textContent = 'Sort: ' + (state.sort === 'new' ? 'Newest' : 'Name');
  $('btn-up').hidden = state.stack.length <= 1 || state.tab !== 'files';
  $('logo').hidden = !$('btn-up').hidden;
  if (state.tab === 'files') $('title').textContent = cur ? cur.name : 'Tax Folder';
}

let lastItems = [];
async function loadFolder() {
  const mine = ++state.seq;
  const cur = current();
  renderCrumbs();
  $('search').value = '';
  $('list').hidden = true;
  const status = $('files-state');
  status.replaceChildren(h('div', { class: 'spinner' }), 'Loading…'); status.hidden = false;
  banner('');
  try {
    const items = await listChildren(cur.id);
    if (mine !== state.seq) return;
    lastItems = items;
    state.dirty = false;
    if (state.stack.length === 1) { state.rootItems = items; state.rootFolders = items.filter((i) => i.folder); }
    renderList(items);
  } catch (e) {
    if (mine !== state.seq) return;
    if (e instanceof AuthError) return needSignIn();
    status.replaceChildren(friendly(e));
  }
}

function openFolder(entry) { state.stack.push(entry); loadFolder(); }
function goTo(i) { if (i >= state.stack.length - 1) return; state.stack.length = i + 1; loadFolder(); }

async function openRoot() {
  state.rootError = '';
  state.position = ls.get('snapshot');
  show('home');
  renderHome();
  try {
    const root = await graph('/me/drive/root:/' + encodePath(cfg().rootPath));
    if (!root.folder) throw new Error(`"${cfg().rootPath}" is not a folder.`);
    state.stack = [{ id: root.id, name: root.name, webUrl: root.webUrl }];
    await loadFolder();
    renderHome();
    refreshPosition();
  } catch (e) {
    if (e instanceof AuthError) return needSignIn();
    state.stack = [];
    renderCrumbs();
    state.rootError = e.status === 404
      ? `Couldn't find "${cfg().rootPath}" in your OneDrive. Check the folder path in Settings (the ⋯ button).`
      : friendly(e);
    $('files-state').replaceChildren(state.rootError); $('files-state').hidden = false;
    renderHome();
  }
}

async function refreshRootItems() {
  if (!state.stack.length) return;
  try {
    state.rootItems = await listChildren(state.stack[0].id);
    state.rootFolders = state.rootItems.filter((i) => i.folder);
  } catch (e) { if (e instanceof AuthError) needSignIn(); }
}

let searchTimer;
async function runSearch() {
  const q = $('search').value.trim();
  if (q.length < 2) { if (!q) renderList(lastItems); return; }
  const mine = ++state.seq;
  const status = $('files-state');
  status.replaceChildren(h('div', { class: 'spinner' }), 'Searching…'); status.hidden = false; $('list').hidden = true;
  try {
    const safe = encodeURIComponent(q.replace(/'/g, "''"));
    const res = await graph(`/me/drive/items/${current().id}/search(q='${safe}')?$top=100`);
    if (mine !== state.seq) return;
    renderList(res.value, { searching: true });
  } catch (e) {
    if (mine !== state.seq) return;
    if (e instanceof AuthError) return needSignIn();
    status.replaceChildren(friendly(e));
  }
}

/* ---------- add receipt ---------- */
const today = () => new Date().toLocaleDateString('en-CA'); // YYYY-MM-DD in local time
const cleanName = (s) => (s || '').replace(/[\\/:*?"<>|#%~&{}]/g, ' ').replace(/\s+/g, ' ').trim().replace(/\.+$/, '').slice(0, 80);
function parseAmount(s) {
  const n = parseFloat(String(s || '').replace(/[^0-9.]/g, ''));
  return Number.isFinite(n) && n > 0 ? n.toFixed(2) : '';
}
function buildName(date, vendor, amount, ext) {
  const parts = [date || today(), cleanName(vendor) || 'Receipt'];
  const a = parseAmount(amount);
  if (a) parts.push('$' + a);
  return parts.join(' ') + ext;
}
function updateNamePreview() {
  const p = state.pending;
  const ext = p ? p.ext : '.jpg';
  $('f-preview-name').textContent = buildName($('f-date').value, $('f-vendor').value, $('f-amount').value, ext);
  $('btn-save').disabled = !p || state.saving;
}

function fillFolderMenu() {
  const sel = $('f-folder');
  const keep = sel.value;
  sel.replaceChildren();
  const root = state.stack[0];
  if (!root) return;
  sel.append(h('option', { value: root.id }, root.name + ' (main folder)'));
  for (const f of [...state.rootFolders].sort((a, b) => a.name.localeCompare(b.name))) sel.append(h('option', { value: f.id }, f.name));
  const want = state.rootFolders.find((f) => f.name.toLowerCase() === cfg().receiptsFolder.toLowerCase());
  sel.value = [...sel.options].some((o) => o.value === keep) ? keep : want ? want.id : root.id;
}

async function compressImage(file, maxSide = 2200, quality = 0.85) {
  let src, w, h2;
  try {
    src = await createImageBitmap(file, { imageOrientation: 'from-image' });
    w = src.width; h2 = src.height;
  } catch {
    src = await new Promise((resolve, reject) => {
      const img = new Image(); const url = URL.createObjectURL(file);
      img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("Couldn't read that image.")); };
      img.src = url;
    });
    w = src.naturalWidth; h2 = src.naturalHeight;
  }
  const scale = Math.min(1, maxSide / Math.max(w, h2));
  const c = document.createElement('canvas');
  c.width = Math.round(w * scale); c.height = Math.round(h2 * scale);
  c.getContext('2d').drawImage(src, 0, 0, c.width, c.height);
  return new Promise((resolve, reject) => c.toBlob((b) => (b ? resolve(b) : reject(new Error("Couldn't process that image."))), 'image/jpeg', quality));
}

function clearPending() {
  if (state.pending && state.pending.previewUrl) URL.revokeObjectURL(state.pending.previewUrl);
  state.pending = null;
  $('preview').hidden = true; $('preview').removeAttribute('src');
  $('capture-empty').hidden = false;
  $('btn-clear').hidden = true;
  $('file-name').hidden = true;
  $('in-camera').value = ''; $('in-file').value = '';
  updateNamePreview();
}

async function onPicked(file) {
  if (!file) return;
  if (file.size > MAX_UPLOAD) { toast('That file is over 100 MB.'); return; }
  banner('');
  try {
    let blob = file, ext;
    const isImage = file.type.startsWith('image/') && !/svg|gif/.test(file.type);
    if (isImage) { blob = await compressImage(file); ext = '.jpg'; }
    else {
      const m = /\.[a-z0-9]{1,5}$/i.exec(file.name);
      ext = m ? m[0].toLowerCase() : file.type === 'application/pdf' ? '.pdf' : '';
    }
    setPending(blob, ext, file.name, {});
    updateNamePreview();
    if (!$('f-vendor').value) $('f-vendor').focus({ preventScroll: true });
  } catch (e) { banner(friendly(e)); }
}

// Shows a file ready to be saved. For a scan the blob is a PDF and previewBlob is its first page.
function setPending(blob, ext, name, { previewBlob = null, info = '' } = {}) {
  if (state.pending && state.pending.previewUrl) URL.revokeObjectURL(state.pending.previewUrl);
  const shown = previewBlob || (blob.type.startsWith('image/') ? blob : null);
  const previewUrl = shown ? URL.createObjectURL(shown) : '';
  state.pending = { blob, ext, name, previewUrl };
  $('capture-empty').hidden = true; $('btn-clear').hidden = false;
  if (previewUrl) { $('preview').src = previewUrl; $('preview').hidden = false; } else { $('preview').hidden = true; }
  const text = info || (previewUrl ? '' : `${name} · ${fmtSize(blob.size)}`);
  $('file-name').textContent = text; $('file-name').hidden = !text;
}
function onScanned({ blob, ext, preview, count }) {
  setPending(blob, ext, 'Scan', { previewBlob: preview, info: `${count} page${count === 1 ? '' : 's'} · ${fmtSize(blob.size)}${ext === '.pdf' ? ' · PDF' : ''}` });
  updateNamePreview();
  if (!$('f-vendor').value) $('f-vendor').focus({ preventScroll: true });
}

async function saveReceipt() {
  const p = state.pending;
  if (!p || state.saving) return;
  state.saving = true;
  const btn = $('btn-save');
  btn.textContent = 'Saving…'; btn.disabled = true; banner('');
  const name = buildName($('f-date').value, $('f-vendor').value, $('f-amount').value, p.ext);
  const folderId = $('f-folder').value;
  const folderName = $('f-folder').selectedOptions[0].textContent.replace(' (main folder)', '');
  try {
    const item = await graph(`/me/drive/items/${folderId}:/${encodeURIComponent(name)}:/content?@microsoft.graph.conflictBehavior=rename`, {
      method: 'PUT', body: p.blob, headers: { 'Content-Type': p.blob.type || 'application/octet-stream' },
    });
    toast(`Saved to ${folderName}: ${item.name}`);
    clearPending();
    $('f-vendor').value = ''; $('f-amount').value = ''; $('f-date').value = today();
    state.dirty = true;
  } catch (e) {
    if (e instanceof AuthError) { state.saving = false; btn.textContent = 'Save to OneDrive'; return needSignIn(); }
    banner(friendly(e));
  } finally {
    state.saving = false;
    btn.textContent = 'Save to OneDrive';
    updateNamePreview();
  }
}

/* ---------- tax position: read from the PTR workbook's OVERVIEW sheet ---------- */
const SECTION_KEYS = {
  'income': 'income', 'ptr deductions': 'ptr', 'ctr deductions': 'ctr',
  'investment property deductions': 'ip', 'other': 'other', 'key dates': 'dates',
};
const SECTION_TITLES = { ptr: 'PTR deductions', ctr: 'CTR deductions', ip: 'Investment property deductions' };
const DEDUCTION_SECTIONS = ['ptr', 'ctr', 'ip'];

const num = (v) => {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') { const n = parseFloat(v.replace(/[$,\s]/g, '')); return Number.isFinite(n) ? n : 0; }
  return 0;
};
const serialToDate = (n) => new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);

// rows = sheet_to_json(header:1). Row 1 holds the income-stream names, then Total and Notes.
function parseOverview(rows) {
  const head = rows[0] || [];
  const streams = [];
  let totalIdx = -1, notesIdx = -1;
  head.forEach((v, i) => {
    const n = String(v == null ? '' : v).trim();
    if (i === 0 || !n) return;
    if (/^total$/i.test(n)) totalIdx = i;
    else if (/^notes?$/i.test(n)) notesIdx = i;
    else streams.push({ idx: i, name: n });
  });
  let section = null;
  const items = [], dates = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const label = String(row[0] == null ? '' : row[0]).trim();
    if (!label) continue;
    const key = SECTION_KEYS[label.replace(/:\s*$/, '').toLowerCase()];
    if (key) { section = key; continue; }
    if (/^surplus$/i.test(label)) { section = null; continue; }
    if (section === 'dates') {
      if (typeof row[1] === 'number' && row[1] > 20000) dates.push({ label, date: serialToDate(row[1]) });
      continue;
    }
    if (!section) continue;
    const amounts = streams.map((s) => num(row[s.idx]));
    let total = totalIdx > 0 ? num(row[totalIdx]) : 0;
    if (!total) total = amounts.reduce((a, b) => a + b, 0);
    if (!total && !amounts.some(Boolean)) continue;
    items.push({ section, label, amounts, total, note: notesIdx > 0 && row[notesIdx] ? String(row[notesIdx]) : '' });
  }
  const sum = (secs, i) => items.filter((x) => secs.includes(x.section)).reduce((a, x) => a + (i == null ? x.total : x.amounts[i]), 0);
  const income = sum(['income']), deductions = sum(DEDUCTION_SECTIONS);
  return {
    streams: streams.map((s, i) => ({ name: s.name, income: sum(['income'], i), deductions: sum(DEDUCTION_SECTIONS, i) })),
    items, dates, income, deductions, other: sum(['other']), net: income - deductions,
  };
}

let xlsxPromise;
const loadXLSX = () => (xlsxPromise = xlsxPromise || new Promise((resolve, reject) => {
  const el = document.createElement('script');
  el.src = 'xlsx.min.js';
  el.onload = () => resolve(window.XLSX);
  el.onerror = () => { xlsxPromise = null; reject(new Error("Couldn't load the spreadsheet reader.")); };
  document.head.append(el);
}));

async function parseWorkbook(buf) {
  const XLSX = await loadXLSX();
  let wb = XLSX.read(buf, { type: 'array', sheets: ['OVERVIEW'] });
  let name = wb.SheetNames.find((n) => /^overview$/i.test(n));
  if (!name) { wb = XLSX.read(buf, { type: 'array' }); name = wb.SheetNames.find((n) => /^overview$/i.test(n)); }
  if (!name) throw new Error('The workbook has no OVERVIEW sheet.');
  return parseOverview(XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null }));
}

async function downloadItem(item) {
  const url = item['@microsoft.graph.downloadUrl'];
  if (url) {
    try { const r = await fetch(url); if (r.ok) return r.arrayBuffer(); } catch { /* fall through to the authenticated route */ }
  }
  const token = await getToken();
  const r = await fetch(`${GRAPH}/me/drive/items/${item.id}/content`, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) throw new Error(`Couldn't download the workbook (${r.status}).`);
  return r.arrayBuffer();
}

function findWorkbook() {
  const want = cfg().workbook.toLowerCase();
  const files = state.rootItems.filter((i) => i.file);
  return files.find((i) => i.name.toLowerCase() === want) || files.find((i) => /^ptr calculations.*\.xlsx$/i.test(i.name));
}

async function refreshPosition(force = false) {
  const item = findWorkbook();
  if (!item) { state.position = null; state.positionError = `Couldn't find "${cfg().workbook}" in ${state.stack[0] ? state.stack[0].name : 'the folder'}.`; renderHome(); return; }
  const snap = state.position;
  if (!force && snap && snap.id === item.id && snap.modified === item.lastModifiedDateTime && snap.data) { state.positionError = ''; renderHome(); return; }
  state.positionBusy = true; state.positionError = ''; renderHome();
  try {
    const data = await parseWorkbook(await downloadItem(item));
    state.position = { id: item.id, name: item.name, modified: item.lastModifiedDateTime, webUrl: item.webUrl, at: Date.now(), data };
    ls.set('snapshot', state.position);
  } catch (e) {
    if (e instanceof AuthError) { state.positionBusy = false; return needSignIn(); }
    state.positionError = friendly(e);
  }
  state.positionBusy = false;
  renderHome();
}

/* Indicative resident-rate tax. The first bracket was 16% to 2025-26, then 15% for 2026-27 and 14% from 2027-28. */
function estimateTax(taxable, fy) {
  const low = fy >= 2028 ? 0.14 : fy === 2027 ? 0.15 : 0.16;
  const t = Math.max(0, taxable);
  const tiers = [[18200, 0], [45000, low], [135000, 0.30], [190000, 0.37], [Infinity, 0.45]];
  let tax = 0, prev = 0, marginal = 0, next = null;
  for (let i = 0; i < tiers.length; i++) {
    const [top, rate] = tiers[i];
    if (t > prev) tax += (Math.min(t, top) - prev) * rate;
    if (t <= top) { marginal = rate; if (top !== Infinity) next = { room: top - t, rate: tiers[i + 1][1] }; break; }
    prev = top;
  }
  const medicare = t * 0.02;
  return { tax, medicare, total: tax + medicare, effective: t ? (tax + medicare) / t : 0, marginal: marginal + 0.02, next };
}


/* ---------- accountant: a ready-to-send email draft (opens in the phone's mail app; you press Send) ---------- */
function summaryEmailHref() {
  const c = cfg(), p = state.position, d = p && p.data;
  if (!d || !c.accountantEmail) return '';
  const fy = fyEnd();
  const fmtD = (ms) => { const t = new Date(ms); return `${t.getUTCDate()} ${MONTHS[t.getUTCMonth()]} ${t.getUTCFullYear()}`; };
  const lines = [
    `Hi ${c.accountantName},`, '',
    `Here is where my FY${fy % 100} tax position stands, from my tax workbook (read ${new Date(p.at).toLocaleDateString('en-AU', { day: 'numeric', month: 'long', year: 'numeric' })}).`, '',
    `Income: ${money(d.income, true)}`, `Deductions: ${money(d.deductions, true)}`, `Net income: ${money(d.net, true)}`, '',
  ];
  const active = d.streams.filter((s) => s.income || s.deductions);
  if (active.length) {
    lines.push('By income stream:');
    for (const s of active) lines.push(`- ${streamLabel(s.name).main}: ${[s.income ? `income ${money(s.income, true)}` : '', s.deductions ? `deductions ${money(s.deductions, true)}` : ''].filter(Boolean).join(', ')}`);
    lines.push('');
  }
  const ded = d.items.filter((i) => DEDUCTION_SECTIONS.includes(i.section));
  if (ded.length) {
    lines.push('Deductions:');
    for (const it of ded) {
      const who = it.amounts.map((a, i) => (a ? streamLabel(d.streams[i].name).main : '')).filter(Boolean).join(', ');
      lines.push(`- ${shortLabel(it.label)}: ${money(it.total, true)}${who ? ` (${who})` : ''}`);
    }
    lines.push('');
  }
  const hasGain = d.items.some((i) => /capital gain/i.test(i.label) && i.total);
  if (d.dates.length) {
    lines.push('Key dates:');
    for (const x of d.dates) {
      const ms = new Date(x.date).getTime();
      lines.push(`- ${fmtD(ms)}: ${x.label}${!hasGain && /\b(sold|sale|disposed)\b/i.test(x.label) ? ' (no capital gain recorded yet)' : ''}`);
    }
    lines.push('');
  }
  lines.push('Let me know if you need anything else.', '', 'Thanks,', 'Dylan');
  return `mailto:${c.accountantEmail}?subject=${encodeURIComponent(`FY${fy % 100} tax position: Dylan Willcocks`)}&body=${encodeURIComponent(lines.join('\n'))}`;
}


/* ---------- tax tasks (checklist) + calendar reminders (.ics) ---------- */
const MONTHS_LONG = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const TASKS_AHEAD_DAYS = 45, TASKS_STALE_DAYS = 40;
const slug = (t) => t.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40);

// ms === null means "do it now"; otherwise the day it falls due (UTC midnight of that calendar day).
function buildTasks() {
  const fy = fyEnd(), acc = cfg().accountantName;
  const d = state.position && state.position.data;
  const out = [];
  const add = (id, title, ms, note) => out.push({ id, title, ms, note });
  for (let m = 1; m <= 12; m++) {               // on the 1st: file the month that just finished
    const prev = new Date(Date.UTC(fy - 1, 5 + m, 1));
    add(`receipts-${prev.getUTCFullYear()}-${prev.getUTCMonth() + 1}`, `File ${MONTHS_LONG[prev.getUTCMonth()]}'s receipts`,
      Date.UTC(fy - 1, 6 + m, 1), 'Add the month\'s receipts to your Receipts folder.');
  }
  add(`accountant-${fy}-sep`, `Send ${acc} an update (September quarter)`, Date.UTC(fy - 1, 9, 7), 'Use the Email summary button on the Position tab.');
  add(`accountant-${fy}-dec`, `Send ${acc} an update (December quarter)`, Date.UTC(fy, 0, 7), 'Use the Email summary button on the Position tab.');
  add(`accountant-${fy}-mar`, `Send ${acc} an update (March quarter)`, Date.UTC(fy, 3, 7), 'Use the Email summary button on the Position tab.');
  add(`sweep-${fy}`, 'Final receipts sweep before 30 June', Date.UTC(fy, 5, 15), 'Check Receipts and the workbook are complete for the year.');
  add(`accountant-${fy}-final`, `Send ${acc} your full-year summary`, Date.UTC(fy, 6, 1), 'Use the Email summary button on the Position tab.');
  add(`lodge-${fy}`, `Lodgement: confirm the plan with ${acc}`, Date.UTC(fy, 9, 1), 'The self-lodgement deadline is 31 October.');
  if (d) {
    const hasGain = d.items.some((i) => /capital gain/i.test(i.label) && i.total);
    for (const x of d.dates) {
      if (!hasGain && /\b(sold|sale|disposed)\b/i.test(x.label)) {
        add(`cgt-${slug(x.label)}`, `Work out the capital gain: ${x.label}`, null, 'Add the gain to the workbook so your tax position includes it.');
      }
    }
  }
  return out;
}
const tasksDone = () => ls.get('tasksDone', {});
function dueText(ms) {
  if (ms === null) return 'Do now';
  const n = Math.round((ms - todayUTC()) / DAY);
  const dt = new Date(ms);
  const when = `${dt.getUTCDate()} ${MONTHS[dt.getUTCMonth()]}`;
  return n < 0 ? `Overdue: ${when}` : n === 0 ? 'Due today' : n === 1 ? 'Due tomorrow' : `Due ${when} (in ${n} days)`;
}
function visibleTasks() {
  const t0 = todayUTC(), done = tasksDone();
  const all = buildTasks();
  const shown = all.filter((t) => t.ms === null || (t.ms - t0 <= TASKS_AHEAD_DAYS * DAY && (t0 - t.ms <= TASKS_STALE_DAYS * DAY || done[t.id])));
  const key = (t) => (t.ms === null ? t0 : t.ms);
  return {
    open: shown.filter((t) => !done[t.id]).sort((a, b) => key(a) - key(b)),
    closed: shown.filter((t) => done[t.id] && (t.ms === null || t0 - t.ms <= TASKS_STALE_DAYS * DAY)),
    later: all.filter((t) => t.ms !== null && t.ms - t0 > TASKS_AHEAD_DAYS * DAY && !done[t.id]).length,
  };
}
function toggleTask(id) {
  const done = tasksDone();
  if (done[id]) delete done[id]; else done[id] = Date.now();
  ls.set('tasksDone', done);
  renderHome();
}
function taskRow(t, isDone) {
  const late = t.ms !== null && t.ms < todayUTC() && !isDone;
  return h('button', { class: 'task', type: 'button', role: 'checkbox', 'aria-checked': String(isDone), onclick: () => toggleTask(t.id) },
    h('span', { class: 'check', 'aria-hidden': 'true' }, isDone ? svg('<path d="M5 12.5l4.5 4.5L19 7.5" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>') : null),
    h('span', { class: 'meta' }, h('b', {}, t.title), h('small', { class: late ? 'late' : '' }, isDone ? 'Done' : dueText(t.ms)), isDone ? null : h('small', {}, t.note)));
}
function renderTasks(box) {
  const { open, closed, later } = visibleTasks();
  box.append(sect('To do'));
  const wrap = card();
  if (!open.length) wrap.append(h('p', { class: 'muted' }, 'Nothing due in the next 45 days.'));
  for (const t of open) wrap.append(taskRow(t, false));
  for (const t of closed) wrap.append(taskRow(t, true));
  if (later) wrap.append(h('p', { class: 'note', style: 'margin-top:8px' }, `${later} more later this year.`));
  wrap.append(h('button', { class: 'btn block', type: 'button', style: 'margin-top:12px', onclick: exportReminders }, 'Add reminders to my phone calendar'));
  wrap.append(h('p', { class: 'note', style: 'margin-top:8px' }, 'Adds an alert at 9 am on each due date to your calendar app. Re-adding updates the same reminders rather than doubling them.'));
  box.append(wrap);
}

const icsText = (v) => String(v).replace(/\\/g, '\\\\').replace(/\r?\n/g, '\\n').replace(/,/g, '\\,').replace(/;/g, '\;');
const icsFold = (line) => { const out = []; let l = line; while (l.length > 74) { out.push(l.slice(0, 74)); l = ' ' + l.slice(74); } out.push(l); return out.join('\r\n'); };
const p2 = (n) => String(n).padStart(2, '0');
function buildIcs() {
  const t0 = todayUTC(), done = tasksDone();
  const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+Z$/, 'Z');
  const todo = buildTasks().filter((t) => !done[t.id] && (t.ms === null || t.ms >= t0));
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Tax Folder//EN', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH'];
  for (const t of todo) {
    const dt = new Date(t.ms === null ? t0 + DAY : t.ms);
    const day = `${dt.getUTCFullYear()}${p2(dt.getUTCMonth() + 1)}${p2(dt.getUTCDate())}`;
    lines.push('BEGIN:VEVENT', `UID:taxfolder-${t.id}@tax-folder`, `DTSTAMP:${stamp}`,
      `DTSTART:${day}T090000`, `DTEND:${day}T091500`,
      `SUMMARY:${icsText('Tax: ' + t.title)}`, `DESCRIPTION:${icsText(t.note + ' Open the Tax Folder app: ' + redirectUri())}`,
      'BEGIN:VALARM', 'ACTION:DISPLAY', `DESCRIPTION:${icsText('Tax: ' + t.title)}`, 'TRIGGER:PT0M', 'END:VALARM', 'END:VEVENT');
  }
  lines.push('END:VCALENDAR');
  return { count: todo.length, text: lines.map(icsFold).join('\r\n') + '\r\n' };
}
async function exportReminders() {
  const ics = buildIcs();
  if (!ics.count) { toast('Nothing to add right now.'); return; }
  const file = new File([ics.text], 'tax-reminders.ics', { type: 'text/calendar' });
  try {
    if (navigator.canShare && navigator.canShare({ files: [file] })) { await navigator.share({ files: [file], title: 'Tax reminders' }); return; }
  } catch (e) { if (e && e.name === 'AbortError') return; }
  const url = URL.createObjectURL(file);
  const a = h('a', { href: url, download: 'tax-reminders.ics' });
  document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
  toast(`${ics.count} reminders ready. Open the file to add them to your calendar.`);
}

/* ---------- home screen ---------- */
const aud0 = new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD', maximumFractionDigits: 0 });
const aud2 = new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD' });
const money = (n, cents = false) => (n < 0 ? '−' : '') + (cents ? aud2 : aud0).format(Math.abs(n));
const pct = (n) => `${(n * 100).toFixed(n * 100 % 1 ? 1 : 0)}%`;
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const DAY = 86400000;
const todayUTC = () => { const d = new Date(); return Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()); };
function fyEnd() {
  const m = /(\d{2})(?!\d)/.exec(cfg().workbook.replace(/\.[^.]+$/, ''));
  if (m) return 2000 + Number(m[1]);
  const d = new Date(); return d.getMonth() >= 6 ? d.getFullYear() + 1 : d.getFullYear();
}
function relDays(ms) {
  const n = Math.round((ms - todayUTC()) / DAY);
  return n === 0 ? 'today' : n === 1 ? 'tomorrow' : n === -1 ? 'yesterday' : n > 0 ? `in ${n} days` : `${-n} days ago`;
}
function ago(ts) {
  const m = Math.round((Date.now() - ts) / 60000);
  return m < 1 ? 'just now' : m < 60 ? `${m} min ago` : m < 1440 ? `${Math.round(m / 60)} h ago` : `${Math.round(m / 1440)} d ago`;
}
function streamLabel(name) {
  const [main, ...rest] = name.split(':');
  const m = main.trim();
  const nice = m === m.toUpperCase() ? m.toLowerCase().replace(/(^|\s)\S/g, (c) => c.toUpperCase()) : m;
  return { main: nice, sub: rest.join(':').trim() };
}
const shortLabel = (l) => l.replace(/\s*\((?:e\.g\.|for|from|related|not|if|where|including|fees)[^)]*\)/gi, '').trim();

function kv(label, value, cls = '') { return h('div', { class: 'kv ' + cls }, h('span', {}, label), h('b', {}, value)); }
function card(...kids) { return h('div', { class: 'card' }, ...kids); }
const sect = (t) => h('div', { class: 'sect' }, t);

function renderHome() {
  const box = $('home');
  box.replaceChildren();
  const fy = fyEnd();
  const p = state.position;
  const d = p && p.data;

  if (state.rootError && !d) { box.append(card(h('b', {}, 'Folder not found'), h('p', { class: 'muted' }, state.rootError))); return; }
  if (!d) {
    box.append(h('div', { class: 'state' },
      state.positionError ? state.positionError : [h('div', { class: 'spinner' }), 'Reading your tax workbook…']));
    if (state.positionError) box.append(h('button', { class: 'btn block', type: 'button', onclick: () => refreshPosition(true) }, 'Try again'));
    return;
  }

  /* hero: net income + how far through the year */
  const start = Date.UTC(fy - 1, 6, 1), end = Date.UTC(fy, 5, 30);
  const total = Math.round((end - start) / DAY) + 1;
  const dayNo = Math.min(total, Math.max(0, Math.floor((todayUTC() - start) / DAY) + 1));
  box.append(h('div', { class: 'hero' },
    h('div', { class: 'eyebrow' }, `FY${fy % 100} net income recorded`),
    h('div', { class: 'big' }, money(d.net)),
    h('div', { class: 'sub' }, `${money(d.income, true)} income less ${money(d.deductions, true)} deductions`),
    h('div', { class: 'progress', role: 'img', 'aria-label': `Day ${dayNo} of ${total}` }, h('i', { style: `width:${(dayNo / total * 100).toFixed(1)}%` })),
    h('div', { class: 'progress-label' }, h('span', {}, dayNo ? `Day ${dayNo} of ${total}` : `Starts ${relDays(start)}`), h('span', {}, dayNo >= total ? 'Year complete' : `${total - dayNo} days left`))));

  box.append(h('div', { class: 'tiles' },
    h('div', { class: 'tile in' }, h('small', {}, 'Income'), h('b', {}, money(d.income))),
    h('div', { class: 'tile' }, h('small', {}, 'Deductions'), h('b', {}, money(d.deductions)))));

  renderTasks(box);

  /* indicative tax */
  const est = estimateTax(d.net, fy);
  const taxRows = [
    kv('Income tax', money(est.tax, true)),
    kv('Medicare levy (2%)', money(est.medicare, true)),
    kv('Estimated total', money(est.total, true), 'total'),
    kv('Effective rate', pct(est.effective)),
    kv('Rate on your next dollar', pct(est.marginal)),
  ];
  if (est.next) taxRows.push(kv(`Room before the ${pct(est.next.rate + 0.02)} rate`, money(est.next.room)));
  box.append(sect('Indicative tax'),
    card(...taxRows,
      h('p', { class: 'note', style: 'margin-top:10px' }, h('span', { class: 'pill warn' }, 'Estimate'), ' ',
        `Assumes resident rates for ${fy - 1}–${String(fy).slice(2)}, all net income taxed to you personally, and no offsets, PAYG credits or capital gains. Trust distributions and the share of income taxed to you can change this a lot, so use it as a guide and check with your accountant.`)));

  /* by income stream */
  const active = d.streams.filter((s) => s.income || s.deductions);
  const idle = d.streams.filter((s) => !s.income && !s.deductions).map((s) => streamLabel(s.name).main);
  if (active.length) {
    box.append(sect('Where it comes from'));
    const wrap = card();
    for (const s of active) {
      const { main, sub } = streamLabel(s.name);
      const net = s.income - s.deductions, whole = s.income + s.deductions;
      wrap.append(h('div', { class: 'stream' },
        h('div', { class: 'stream-head' }, h('span', { class: 'stream-name' }, main), h('b', {}, money(net, true))),
        sub ? h('small', {}, sub) : null,
        h('small', {}, [s.income ? `Income ${money(s.income)}` : '', s.deductions ? `Deductions ${money(s.deductions, true)}` : ''].filter(Boolean).join(' · ')),
        h('div', { class: 'bar', 'aria-hidden': 'true' },
          s.income ? h('i', { class: 'in', style: `width:${(s.income / whole * 100).toFixed(1)}%` }) : null,
          s.deductions ? h('i', { class: 'out', style: `width:${(s.deductions / whole * 100).toFixed(1)}%` }) : null)));
    }
    if (idle.length) wrap.append(h('p', { class: 'note', style: 'margin-top:8px' }, `Nothing recorded yet: ${idle.join(', ')}.`));
    box.append(wrap);
  }

  /* deductions by section */
  const dedGroups = DEDUCTION_SECTIONS.map((sec) => ({ sec, items: d.items.filter((i) => i.section === sec) })).filter((g) => g.items.length);
  if (dedGroups.length) {
    box.append(sect('Deductions'));
    const wrap = card();
    dedGroups.forEach((g, gi) => {
      const sub = g.items.reduce((a, i) => a + i.total, 0);
      wrap.append(h('div', { class: 'kv total', style: gi ? 'margin-top:6px' : '' }, h('span', {}, SECTION_TITLES[g.sec]), h('b', {}, money(sub, true))));
      for (const it of g.items) {
        const who = it.amounts.map((a, i) => (a ? streamLabel(d.streams[i].name).main : '')).filter(Boolean).join(', ');
        wrap.append(h('div', { class: 'kv' }, h('span', {}, shortLabel(it.label), who ? h('small', { style: 'display:block' }, who) : null), h('b', {}, money(it.total, true))));
      }
    });
    box.append(wrap);
  }
  const others = d.items.filter((i) => i.section === 'other');
  if (others.length) {
    box.append(sect('Other tax items'), card(...others.map((it) => kv(shortLabel(it.label), money(it.total, true)))));
  }

  /* key dates: the workbook's own, plus the fixed ones for this year */
  const hasGain = d.items.some((i) => /capital gain/i.test(i.label) && i.total);
  const dates = [
    ...d.dates.map((x) => ({ label: x.label, ms: new Date(x.date).getTime(), cgt: !hasGain && /\b(sold|sale|disposed)\b/i.test(x.label) })),
    { label: `Financial year ends`, ms: end, fixed: true },
    { label: `Self-lodgement due`, ms: Date.UTC(fy, 9, 31), fixed: true },
  ].sort((a, b) => a.ms - b.ms);
  box.append(sect('Key dates'));
  const dwrap = card();
  for (const x of dates) {
    const dt = new Date(x.ms);
    dwrap.append(h('div', { class: 'date-row' },
      h('div', { class: 'date-box' }, h('b', {}, String(dt.getUTCDate())), h('small', {}, MONTHS[dt.getUTCMonth()])),
      h('div', { class: 'meta' }, h('b', {}, x.label), h('small', {}, `${dt.getUTCFullYear()} · ${relDays(x.ms)}`),
        x.cgt ? h('div', { style: 'margin-top:4px' }, h('span', { class: 'pill warn' }, 'CGT event: gain not in these figures yet')) : null)));
  }
  box.append(dwrap);

  const acc = cfg();
  box.append(sect('Your accountant'));
  box.append(card(
    h('div', { class: 'kv total' }, h('span', {}, acc.accountantName), h('b', { style: 'font-weight:500;font-size:14px;color:var(--muted)' }, acc.accountantEmail || 'No email saved yet')),
    h('p', { class: 'note', style: 'margin:6px 0 12px' }, 'Opens an email in your mail app with these figures already written. Nothing is sent until you press Send.'),
    acc.accountantEmail
      ? h('a', { class: 'btn primary block', href: summaryEmailHref() }, `Email ${acc.accountantName} this summary`)
      : h('button', { class: 'btn block', type: 'button', onclick: openSettings }, `Add ${acc.accountantName}'s email`)));

  /* documents + actions */
  const rec = state.rootItems.find((i) => i.folder && i.name.toLowerCase() === cfg().receiptsFolder.toLowerCase());
  box.append(sect('Documents'));
  box.append(card(
    kv('Receipts filed', rec ? String(rec.folder.childCount) : 'No Receipts folder'),
    kv('Items in the tax folder', String(state.rootItems.length)),
    kv('Workbook last saved', p.modified ? fmtDate(p.modified) : 'Unknown')));
  box.append(h('div', { class: 'actions', style: 'margin-top:12px' },
    h('button', { class: 'btn primary', type: 'button', onclick: () => $('tab-add').click() }, 'Add receipt'),
    p.webUrl ? h('a', { class: 'btn', href: p.webUrl, target: '_blank', rel: 'noopener' }, 'Open workbook') : null));
  box.append(h('p', { class: 'note', style: 'text-align:center;margin-top:14px' },
    state.positionBusy ? 'Updating from OneDrive…' : state.positionError ? `Couldn't update (${state.positionError}). Showing figures from ${ago(p.at)}.` : `Figures from ${p.name}, read ${ago(p.at)}.`,
    ' ', h('button', { class: 'btn quiet', type: 'button', style: 'min-height:36px;padding:0 8px', onclick: () => refreshPosition(true) }, 'Refresh')));
}

/* ---------- app lock: Face ID / Touch ID / fingerprint, through the phone's passkey prompt ---------- */
const BIO = /iPhone|iPad/.test(navigator.userAgent) ? 'Face ID' : 'Biometric';
const bytes = (n) => crypto.getRandomValues(new Uint8Array(n));
const fromB64url = (s) => Uint8Array.from(atob(s.replace(/-/g, '+').replace(/_/g, '/')), (c) => c.charCodeAt(0));
const lockOn = () => !!ls.get('lock');

async function lockSupported() {
  try { return !!(window.PublicKeyCredential && await PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable()); } catch { return false; }
}
async function enableLock() {
  if (!(await lockSupported())) throw new Error(`This browser can't use ${BIO} for web apps. On iPhone, open the app from the home-screen icon.`);
  const cred = await navigator.credentials.create({ publicKey: {
    rp: { name: 'Tax Folder', id: location.hostname },
    user: { id: bytes(16), name: 'tax-folder', displayName: 'Tax Folder' },
    challenge: bytes(32),
    pubKeyCredParams: [{ type: 'public-key', alg: -7 }, { type: 'public-key', alg: -257 }],
    authenticatorSelection: { authenticatorAttachment: 'platform', userVerification: 'required', residentKey: 'preferred' },
    attestation: 'none', timeout: 60000,
  } });
  ls.set('lock', { id: b64url(cred.rawId) });
}
async function verifyLock() {
  const l = ls.get('lock');
  await navigator.credentials.get({ publicKey: {
    challenge: bytes(32), rpId: location.hostname, userVerification: 'required', timeout: 60000,
    allowCredentials: [{ type: 'public-key', id: fromB64url(l.id), transports: ['internal'] }],
  } });
}
function lockError(e) {
  if (e && e.name === 'NotAllowedError') return `${BIO} was cancelled or didn't match. Try again.`;
  if (e && e.name === 'SecurityError') return `${BIO} needs the secure https address.`;
  return friendly(e);
}
function showLock(msg) { $('btn-unlock').textContent = `Unlock with ${BIO}`; $('lock-msg').textContent = msg || 'Locked'; $('lock').hidden = false; }
let unlocking = false;
async function unlock(auto = false) {
  if (unlocking) return;
  unlocking = true;
  try { await verifyLock(); $('lock').hidden = true; }
  catch (e) { showLock(auto && e && e.name === 'NotAllowedError' ? `Tap to unlock with ${BIO}` : lockError(e)); }
  finally { unlocking = false; }
}
let hiddenAt = 0;
document.addEventListener('visibilitychange', () => {
  if (!lockOn()) return;
  if (document.hidden) { hiddenAt = Date.now(); showLock('Locked'); }              // also hides the app-switcher preview
  else if (Date.now() - hiddenAt > LOCK_AFTER_MS) { showLock('Locked'); unlock(true); }
  else $('lock').hidden = true;
});
function renderLockRow() {
  const on = lockOn();
  $('s-lock-title').textContent = `${BIO} lock`;
  $('s-lock-state').textContent = on ? 'On. Asked when you open the app or come back after 2 minutes.' : 'Off';
  $('s-lock-toggle').textContent = on ? 'Turn off' : 'Turn on';
}

/* ---------- settings ---------- */
function openSettings() {
  const c = cfg();
  $('s-root').value = c.rootPath; $('s-receipts').value = c.receiptsFolder; $('s-client').value = c.clientId;
  $('s-redirect').textContent = redirectUri();
  $('s-workbook').value = c.workbook;
  $('s-acc-name').value = c.accountantName; $('s-acc-email').value = c.accountantEmail;
  renderLockRow();
  $('settings').showModal();
}
function signOut() { for (const k of ['tokens', 'pkce', 'lock', 'snapshot']) ls.del(k); location.replace(redirectUri()); }

/* ---------- wire up ---------- */
function wire() {
  $('redirect-uri').value = redirectUri();
  $('save-client').addEventListener('click', () => {
    const v = $('client-id').value.trim();
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(v)) { banner('That doesn\'t look like an Application (client) ID. It is 32 characters in 5 groups separated by dashes.'); return; }
    banner(''); ls.set('clientId', v); show('signin');
  });
  $('btn-signin').addEventListener('click', () => signIn().catch((e) => banner(friendly(e))));
  $('btn-settings').addEventListener('click', openSettings);
  $('btn-up').addEventListener('click', () => goTo(state.stack.length - 2));
  $('btn-refresh').addEventListener('click', loadFolder);
  $('btn-sort').addEventListener('click', () => {
    state.sort = state.sort === 'name' ? 'new' : 'name'; ls.set('sort', state.sort);
    $('btn-sort').textContent = 'Sort: ' + (state.sort === 'new' ? 'Newest' : 'Name');
    if ($('search').value.trim().length >= 2) runSearch(); else renderList(lastItems);
  });
  $('search').addEventListener('input', () => { clearTimeout(searchTimer); searchTimer = setTimeout(runSearch, 400); });
  $('tab-home').addEventListener('click', async () => {
    show('home'); renderHome();
    if (state.dirty && state.stack.length) { await refreshRootItems(); renderHome(); }
  });
  $('tab-files').addEventListener('click', () => { show('files'); if (state.dirty && current()) loadFolder(); });
  $('tab-add').addEventListener('click', () => { show('add'); fillFolderMenu(); updateNamePreview(); });

  $('in-camera').addEventListener('change', (e) => onPicked(e.target.files[0]));
  $('in-file').addEventListener('change', (e) => onPicked(e.target.files[0]));
  $('btn-clear').addEventListener('click', clearPending);
  $('btn-scan').addEventListener('click', () => { Scanner.wire(); Scanner.open(onScanned); });
  for (const id of ['f-date', 'f-vendor', 'f-amount']) $(id).addEventListener('input', updateNamePreview);
  $('f-amount').addEventListener('blur', () => { const a = parseAmount($('f-amount').value); if (a) $('f-amount').value = a; updateNamePreview(); });
  $('btn-save').addEventListener('click', saveReceipt);
  $('f-date').value = today();

  $('settings').addEventListener('close', () => {
    if ($('settings').returnValue !== 'save') return;
    const c = cfg();
    const next = { rootPath: $('s-root').value.trim().replace(/^\/+|\/+$/g, '') || DEFAULTS.rootPath, receiptsFolder: $('s-receipts').value.trim() || DEFAULTS.receiptsFolder, workbook: $('s-workbook').value.trim() || DEFAULTS.workbook, clientId: $('s-client').value.trim() };
    ls.set('rootPath', next.rootPath); ls.set('receiptsFolder', next.receiptsFolder); ls.set('workbook', next.workbook);
    if (next.workbook !== c.workbook) ls.del('snapshot');
    ls.set('accountantName', $('s-acc-name').value.trim() || 'Brayden'); ls.set('accountantEmail', $('s-acc-email').value.trim());
    if (next.clientId && next.clientId !== c.clientId) { ls.set('clientId', next.clientId); ls.del('tokens'); }
    location.replace(redirectUri());
  });
  $('settings').addEventListener('click', (e) => { if (e.target === $('settings')) $('settings').close('cancel'); });
  $('s-cancel').addEventListener('click', () => $('settings').close('cancel'));
  $('s-lock-toggle').addEventListener('click', async () => {
    const msg = $('s-lock-state');
    try {
      if (lockOn()) { await verifyLock(); ls.del('lock'); } else await enableLock();
    } catch (e) { msg.textContent = lockError(e); return; }
    renderLockRow();
  });
  $('btn-unlock').addEventListener('click', () => unlock(false));
  $('btn-lock-reset').addEventListener('click', signOut);
  $('s-signout').addEventListener('click', signOut);
}

async function boot() {
  wire();
  if (lockOn()) { showLock('Locked'); unlock(true); }
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  if (!cfg().clientId) return show('setup');
  try { await handleRedirect(); } catch (e) { banner(friendly(e)); }
  if (!ls.get('tokens')) return show('signin');
  await openRoot();
}
boot();
