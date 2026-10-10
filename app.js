'use strict';

/* Tax Folder: browse a OneDrive folder and file receipts into it.
   Talks straight to Microsoft Graph from the browser. Sign-in is OAuth authorization
   code + PKCE, so there is no server and no secret. */

const GRAPH = 'https://graph.microsoft.com/v1.0';
const AUTH = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
const SCOPES = 'Files.ReadWrite offline_access';
const DEFAULTS = { rootPath: 'Personal Documents/Tax/Tax27', receiptsFolder: 'Receipts', workbook: 'PTR Calculations 27.xlsx' };
const LOCK_AFTER_MS = 2 * 60 * 1000;
const APP_VERSION = '16';
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
  inboxFolder: ls.get('inboxFolder') || 'Inbox',
  apiKey: ls.get('apiKey') || '',
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

// fetch that gives up instead of hanging forever (a stalled request used to leave "Saving…" spinning)
async function fetchT(url, opts = {}, ms = 30000) {
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), ms);
  try { return await fetch(url, { ...opts, signal: ctl.signal }); }
  catch (e) {
    if (e && e.name === 'AbortError') { const err = new Error('The connection timed out. Check your signal and try again.'); err.timeout = true; throw err; }
    throw e;
  } finally { clearTimeout(timer); }
}

async function tokenRequest(params) {
  const res = await fetchT(`${AUTH}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: cfg().clientId, scope: SCOPES, ...params }),
  }, 20000);
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
async function graph(path, { method = 'GET', body, headers = {}, ms } = {}, retry = true) {
  const token = await getToken();
  const res = await fetchT(path.startsWith('http') ? path : GRAPH + path, {
    method, body, headers: { Authorization: `Bearer ${token}`, ...headers },
  }, ms || (method === 'GET' ? 30000 : 60000));
  if (res.status === 401 && retry) {
    const t = ls.get('tokens');
    if (t) ls.set('tokens', { ...t, exp: 0 });
    return graph(path, { method, body, headers, ms }, false);
  }
  if (!res.ok) {
    let msg = res.statusText || `HTTP ${res.status}`;
    try { msg = (await res.json()).error.message || msg; } catch { /* keep default */ }
    const e = new Error(msg); e.status = res.status; throw e;
  }
  return res.status === 204 ? null : res.json();
}

// File uploads go through XMLHttpRequest so we can show real progress and notice a stalled connection.
async function uploadContent(path, blob, onProgress, retry = true) {
  const token = await getToken();
  try {
    return await new Promise((resolve, reject) => {
      const x = new XMLHttpRequest();
      let last = Date.now(), done = false;
      const finish = (fn, v) => { if (done) return; done = true; clearInterval(stall); fn(v); };
      const stall = setInterval(() => {
        if (done || Date.now() - last < 45000) return;
        const e = new Error('The upload stalled. Check your signal and try again.'); e.timeout = true;
        finish(reject, e); x.abort();
      }, 5000);
      x.open('PUT', GRAPH + path);
      x.setRequestHeader('Authorization', 'Bearer ' + token);
      x.setRequestHeader('Content-Type', blob.type || 'application/octet-stream');
      x.upload.onprogress = (e) => { last = Date.now(); if (e.lengthComputable && onProgress) onProgress(e.loaded / e.total); };
      x.onload = () => {
        if (x.status >= 200 && x.status < 300) { let j = {}; try { j = JSON.parse(x.responseText); } catch { /* empty body */ } return finish(resolve, j); }
        let msg = x.statusText || `HTTP ${x.status}`;
        try { msg = JSON.parse(x.responseText).error.message || msg; } catch { /* keep default */ }
        const e = new Error(msg); e.status = x.status; finish(reject, e);
      };
      x.onerror = () => finish(reject, new TypeError('Network error'));
      x.onabort = () => { const e = new Error('The upload was cancelled.'); e.timeout = true; finish(reject, e); };
      x.send(blob);
    });
  } catch (e) {
    if (e.status === 401 && retry) { const t = ls.get('tokens'); if (t) ls.set('tokens', { ...t, exp: 0 }); return uploadContent(path, blob, onProgress, false); }
    throw e;
  }
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
  rates: null,        // rates.json, loaded once per open
  ratesError: '',
  assume: { data: ls.get('assumptions'), draft: null, etag: '', loaded: false, busy: false, saving: false, error: '' },
  vehicle: undefined, // mileage summary from Inbox/vehicle.json; undefined = not fetched yet
};
const current = () => state.stack[state.stack.length - 1];

/* ---------- views ---------- */
function show(name) {
  state.tab = name;
  for (const v of ['setup', 'signin', 'home', 'inbox', 'ask', 'files', 'add']) $('view-' + v).hidden = v !== name;
  $('nav').hidden = !['home', 'inbox', 'ask', 'files', 'add'].includes(name);
  for (const t of ['home', 'inbox', 'ask', 'files', 'add']) {
    const b = $('tab-' + t);
    if (t === name) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current');
  }
  const showUp = name === 'files' && state.stack.length > 1;
  $('btn-up').hidden = !showUp;
  $('logo').hidden = showUp;
  $('title').textContent = name === 'files' && current() ? current().name
    : name === 'add' ? 'Add receipt'
    : name === 'inbox' ? 'Inbox'
    : name === 'ask' ? 'Ask'
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
  loadRates().then(() => { if (state.tab === 'home') renderHome(); });
  try {
    const root = await graph('/me/drive/root:/' + encodePath(cfg().rootPath));
    if (!root.folder) throw new Error(`"${cfg().rootPath}" is not a folder.`);
    state.stack = [{ id: root.id, name: root.name, webUrl: root.webUrl }];
    await loadFolder();
    renderHome();
    refreshPosition();
    loadInbox();
    loadAssumptions();
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
  sel.append(h('option', { value: 'auto' }, 'Auto: file it by what it is'));
  sel.append(h('option', { value: root.id }, root.name + ' (main folder)'));
  for (const f of [...state.rootFolders].sort((a, b) => a.name.localeCompare(b.name))) sel.append(h('option', { value: f.id }, f.name));
  sel.value = [...sel.options].some((o) => o.value === keep) ? keep : 'auto';
}
function fillAddOptions() {
  const t = $('f-treatment');
  if (!t.options.length) t.append(h('option', { value: '' }, 'Not sure yet'), ...Object.entries(TREATMENTS).map(([k, v]) => h('option', { value: k }, v)));
  const s = $('f-stream'), keep = s.value;
  s.replaceChildren(h('option', { value: '' }, 'Not sure which stream'), ...streamOptions().map((n) => h('option', { value: n }, n)));
  s.value = keep;
  $('f-cats').replaceChildren(...categoryOptions().map((c) => h('option', { value: c })));
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
  state.aiSeq = (state.aiSeq || 0) + 1; state.addAI = null;
  for (const id of ['f-treatment', 'f-stream', 'f-category', 'f-note']) $(id).value = '';
  delete $('f-date').dataset.touched;
  refreshAddCard();
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

/* ---- paste a receipt ------------------------------------------------------
   Two routes, because browsers differ: the paste event (⌘V) carries the file
   directly, while the button has to ask for the clipboard and needs a gesture. */
const PASTEABLE = (t) => /^image\/(png|jpe?g|webp|heic|heif|tiff|bmp)$/i.test(t) || t === 'application/pdf';

function pastedName(type) {
  const d = new Date(), p = (n) => String(n).padStart(2, '0');
  const ext = type === 'application/pdf' ? 'pdf' : (type.split('/')[1] || 'png').replace('jpeg', 'jpg');
  return `Pasted ${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}${p(d.getMinutes())}.${ext}`;
}

async function acceptPaste(file) {
  if (!file) { toast('Nothing on the clipboard I can read. Copy an image or a PDF first.'); return false; }
  if (!PASTEABLE(file.type)) { toast(`Can't use ${file.type || 'that'} — copy an image or a PDF.`); return false; }
  if (state.tab !== 'add') { $('tab-add').click(); }
  document.body.classList.add('pasting');
  try { await onPicked(file); } finally { document.body.classList.remove('pasting'); }
  return true;
}

/* ⌘V anywhere in the app, as long as focus isn't in a text field */
async function onPasteEvent(e) {
  const t = e.target;
  if (t && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName))) return;
  const dt = e.clipboardData;
  if (!dt) return;
  let file = [...(dt.files || [])].find((f) => PASTEABLE(f.type));
  if (!file) {
    const item = [...(dt.items || [])].find((i) => i.kind === 'file' && PASTEABLE(i.type));
    if (item) file = item.getAsFile();
  }
  if (!file) return;                      // let normal paste happen
  e.preventDefault();
  if (!file.name || file.name === 'image.png') file = new File([file], pastedName(file.type), { type: file.type });
  await acceptPaste(file);
}

/* the button: ask the clipboard directly */
async function pasteFromClipboard() {
  if (!navigator.clipboard || !navigator.clipboard.read) {
    toast('This browser won\'t let me open the clipboard — press ⌘V instead.'); return;
  }
  try {
    for (const item of await navigator.clipboard.read()) {
      const type = item.types.find(PASTEABLE);
      if (!type) continue;
      const blob = await item.getType(type);
      return void acceptPaste(new File([blob], pastedName(type), { type }));
    }
    toast('Nothing on the clipboard I can read. Copy an image or a PDF first.');
  } catch (e) {
    toast(/denied|permission/i.test(String(e))
      ? 'Clipboard access was blocked — press ⌘V instead.'
      : 'Couldn\'t read the clipboard — press ⌘V instead.');
  }
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
  state.addAI = null; fillAddOptions(); refreshAddCard(); analyzePending();
}
function onScanned({ blob, ext, preview, count }) {
  setPending(blob, ext, 'Scan', { previewBlob: preview, info: `${count} page${count === 1 ? '' : 's'} · ${fmtSize(blob.size)}${ext === '.pdf' ? ' · PDF' : ''}` });
  updateNamePreview();
  if (!$('f-vendor').value) $('f-vendor').focus({ preventScroll: true });
}

// Subfolders under Receipts, chosen by what the document turned out to be.
const TREAT_FOLDER = { claim: '', xero: 'Company costs', cgt: 'Capital gains', reimbursed: 'Not claimable', personal: 'Not claimable', skip: 'Not claimable' };
async function ensureFolder(parentId, name) {
  try { const f = await graph(`/me/drive/items/${parentId}:/${encodeURIComponent(name)}`); if (f.folder) return f.id; }
  catch (e) { if (e.status !== 404) throw e; }
  const made = await graph(`/me/drive/items/${parentId}/children`, {
    method: 'POST', body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' }), headers: { 'Content-Type': 'application/json' },
  });
  return made.id;
}
async function destinationFor(treatment) {
  const sel = $('f-folder');
  if (sel.value && sel.value !== 'auto') return { id: sel.value, label: sel.selectedOptions[0].textContent.replace(' (main folder)', '') };
  const rec = state.rootFolders.find((f) => f.name.toLowerCase() === cfg().receiptsFolder.toLowerCase());
  const base = rec ? rec.id : state.stack[0].id, baseName = rec ? rec.name : state.stack[0].name;
  const sub = TREAT_FOLDER[treatment] || '';
  return sub ? { id: await ensureFolder(base, sub), label: `${baseName} / ${sub}` } : { id: base, label: baseName };
}
// Records an uploaded document as an already-decided item, so it reaches the workbook log with the rest.
async function addUploadedItem(file, dest, treatment) {
  const ai = (state.addAI && state.addAI.result) || {};
  const amount = parseAmount($('f-amount').value);
  const entry = {
    id: 'up-' + Date.now().toString(36), source: 'upload',
    date: $('f-date').value || today(), vendor: $('f-vendor').value.trim() || ai.vendor || '', docType: ai.docType || 'receipt',
    description: ai.description || '', amount: amount ? Number(amount) : null, gst: ai.gst == null ? null : ai.gst,
    payment: ai.payment || null, ref: ai.ref || null, attachments: [file.name],
    file: { id: file.id || '', name: file.name, webUrl: file.webUrl || '', folder: dest.label }, viewUrl: file.webUrl || '',
    relevance: 'likely', status: 'decided',
    suggestion: ai.treatment ? { treatment: ai.treatment, stream: ai.stream || '', category: ai.category || '', confidence: ai.confidence || '', reason: ai.reason || '', question: ai.question || null } : null,
    decision: { treatment, stream: $('f-stream').value, category: $('f-category').value.trim(), note: $('f-note').value.trim(), at: new Date().toISOString() },
  };
  await mutateInbox((data) => { data.items = data.items || []; data.items.push(entry); });
}

async function saveReceipt() {
  const p = state.pending;
  if (!p || state.saving) return;
  state.saving = true;
  const btn = $('btn-save');
  btn.disabled = true; btn.textContent = 'Saving…'; banner('');
  const treatment = $('f-treatment').value;
  const name = buildName($('f-date').value, $('f-vendor').value, $('f-amount').value, p.ext);
  try {
    const dest = await destinationFor(treatment);
    const file = await uploadContent(`/me/drive/items/${dest.id}:/${encodeURIComponent(name)}:/content?@microsoft.graph.conflictBehavior=rename`, p.blob,
      (f) => { btn.textContent = `Saving… ${Math.round(f * 100)}%`; });
    let recorded = true;
    if (treatment) { try { await addUploadedItem(file, dest, treatment); } catch (e) { recorded = false; } }
    toast(`Saved to ${dest.label}: ${file.name || name}${treatment && !recorded ? ". I couldn't record it in the inbox." : ''}`);
    clearPending();
    $('f-vendor').value = ''; $('f-amount').value = ''; $('f-date').value = today();
    state.dirty = true;
  } catch (e) {
    if (e instanceof AuthError) { state.saving = false; btn.textContent = 'Save to OneDrive'; return needSignIn(); }
    banner(`Couldn't save: ${friendly(e)}${e.status ? ` (code ${e.status})` : ''}`);
  } finally {
    state.saving = false;
    btn.textContent = 'Save to OneDrive';
    updateNamePreview();
  }
}

/* ---------- reading a document with Claude ---------- */
const EXTRACT_SCHEMA = {
  type: 'object', additionalProperties: false,
  required: ['vendor', 'date', 'total', 'gst', 'payment', 'docType', 'description', 'ref', 'address', 'treatment', 'stream', 'category', 'confidence', 'reason', 'question'],
  properties: {
    vendor: { type: 'string' }, date: { anyOf: [{ type: 'string' }, { type: 'null' }] },
    total: { anyOf: [{ type: 'number' }, { type: 'null' }] }, gst: { anyOf: [{ type: 'number' }, { type: 'null' }] },
    payment: { anyOf: [{ type: 'string' }, { type: 'null' }] },
    docType: { type: 'string', enum: ['receipt', 'tax invoice', 'bill', 'statement', 'order confirmation', 'letter', 'other'] },
    description: { type: 'string' }, ref: { anyOf: [{ type: 'string' }, { type: 'null' }] }, address: { anyOf: [{ type: 'string' }, { type: 'null' }] },
    treatment: { type: 'string', enum: ['claim', 'xero', 'reimbursed', 'personal', 'cgt', 'skip', 'unsure'] },
    stream: { type: 'string' }, category: { type: 'string' }, confidence: { type: 'string', enum: ['low', 'medium', 'high'] },
    reason: { type: 'string' }, question: { anyOf: [{ type: 'string' }, { type: 'null' }] },
  },
};
const blobToBase64 = (blob) => new Promise((resolve, reject) => {
  const r = new FileReader();
  r.onload = () => resolve(String(r.result).split(',')[1] || '');
  r.onerror = () => reject(new Error("Couldn't read the file."));
  r.readAsDataURL(blob);
});
function buildExtractPrompt() {
  const d = state.position && state.position.data, l = d && d.labels, fy = fyEnd();
  const streams = streamOptions();
  const list = (a) => (a && a.length ? a.join(' | ') : '(not available)');
  return [
    `You read one document (a receipt, tax invoice, bill, statement or letter) for Dylan's personal tax records and classify it. The financial year is 1 July ${fy - 1} to 30 June ${fy}. Today is ${new Date().toISOString().slice(0, 10)}.`,
    'Return ONLY the JSON object that matches the schema. Use null (or an empty string for stream and category) when something is not visible or you are not sure. Never invent amounts, dates or numbers.',
    'vendor = the business. date = the document date as YYYY-MM-DD (the invoice or transaction date, not a print date). total = what was paid or is payable in AUD including GST. gst = the GST amount if shown. payment = card brand and last four digits, PayPal, bank transfer and so on, if shown. description = one short line. ref = invoice, order, policy or assessment number. address = a property address if the document is about a property.',
    'treatment: claim (a deduction in his personal return), xero (a company cost: anything paid on the business card or clearly Oakwood business; it is not part of his personal return and he does not use Xero himself, so never mention Xero), reimbursed (already claimed back from Oakwood), personal (not claimable), cgt (a cost of buying, improving or selling a property that belongs in the capital gain), skip (not a receipt), unsure.',
    `stream must be exactly one of: ${list(streams)}, or empty. category must be copied exactly from the lists below, or empty.`,
    'reason = one or two plain sentences. question = one specific question whose answer would change the treatment, or null if you are confident. Use his profile for his cards and standing rules: spend on the business card is xero, personal-card spend is personal, anything Oakwood reimbursed is not claimable, selling costs are cgt, utilities and rates are only deductible for a property that is rented or genuinely available for rent.',
    '', 'ABOUT DYLAN', state.ask.profile || '(no profile available)',
    '', 'CATEGORY LISTS (copy exactly)',
    `Personal return deductions: ${list(l && l.ptr)}`,
    `Investment property deductions: ${list(l && l.ip)}`,
    `Company deductions (company costs, not his return): ${list(l && l.ctr)}`,
  ].join('\n');
}
async function extractDocument(blob) {
  const isPdf = blob.type === 'application/pdf', isImg = /^image\/(jpeg|png|webp|gif)$/.test(blob.type);
  if (!isPdf && !isImg) throw new Error('I can read photos and PDFs. Choose what this one is below.');
  const b64 = await blobToBase64(blob);
  const content = [
    isPdf ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: b64 } }
      : { type: 'image', source: { type: 'base64', media_type: blob.type, data: b64 } },
    { type: 'text', text: 'Read this document and return the JSON.' },
  ];
  const base = { model: ASK_MODEL, max_tokens: 4000, system: buildExtractPrompt(), messages: [{ role: 'user', content }] };
  const withSchema = { ...base, output_config: { format: { type: 'json_schema', schema: EXTRACT_SCHEMA } } };
  let data, lastErr;
  for (const go of [() => postClaude(withSchema, true), () => postClaude(withSchema, false), () => postClaude(base, false)]) {
    try { data = await go(); break; } catch (e) { lastErr = e; if (e.status !== 400) throw e; }
  }
  if (!data) throw lastErr;
  if (data.stop_reason === 'refusal') throw new Error("Claude wouldn't read that one. Choose what it is below.");
  const text = (data.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('').trim();
  let r;
  try { r = JSON.parse(text); } catch { const m = /\{[\s\S]*\}/.exec(text); if (!m) throw new Error('Claude sent back something I could not read. Try again.'); r = JSON.parse(m[0]); }
  return r;
}
async function analyzePending() {
  const p = state.pending;
  if (!p || !cfg().apiKey) return;
  if (p.blob.size > 20 * 1024 * 1024) { state.addAI = { busy: false, error: 'That file is too large for me to read here. Choose what it is below.' }; refreshAddCard(); return; }
  const mine = state.aiSeq = (state.aiSeq || 0) + 1;
  state.addAI = { busy: true }; refreshAddCard();
  try {
    if (!state.ask.loaded) await loadAskContext();
    const result = await extractDocument(p.blob);
    if (mine !== state.aiSeq || state.pending !== p) return;
    state.addAI = { busy: false, result };
    applyAIResult(result);
  } catch (e) {
    if (mine !== state.aiSeq) return;
    if (e instanceof AuthError) return needSignIn();
    state.addAI = { busy: false, error: askErrorText(e) };
  }
  refreshAddCard();
}
function applyAIResult(r) {
  if (r.date && /^\d{4}-\d{2}-\d{2}$/.test(r.date) && !$('f-date').dataset.touched) $('f-date').value = r.date;
  if (r.vendor && !$('f-vendor').value) $('f-vendor').value = r.vendor;
  if (r.total != null && !$('f-amount').value) $('f-amount').value = Number(r.total).toFixed(2);
  $('f-treatment').value = TREATMENTS[r.treatment] ? r.treatment : '';
  $('f-stream').value = [...$('f-stream').options].some((o) => o.value === r.stream) ? r.stream : '';
  $('f-category').value = r.category || '';
  updateNamePreview();
}
function refreshAddCard() {
  const p = state.pending, card = $('ai-card');
  card.hidden = !p;
  if (!p) return;
  const a = state.addAI || {}, st = $('ai-status'), gb = $('ai-guess');
  gb.hidden = true;
  if (!cfg().apiKey) { st.textContent = 'Choose what it is below. Add a Claude API key in Settings and I will read and classify documents for you.'; return; }
  if (a.busy) { st.replaceChildren(h('span', { class: 'spinner', style: 'display:inline-block;vertical-align:middle;margin:0 8px 0 0;width:16px;height:16px' }), 'Claude is reading it…'); return; }
  if (a.error) { st.replaceChildren(a.error + ' ', h('button', { class: 'chip', type: 'button', onclick: analyzePending }, 'Try again')); return; }
  if (!a.result) { st.replaceChildren(h('button', { class: 'chip', type: 'button', onclick: analyzePending }, 'Read it with Claude')); return; }
  const r = a.result;
  st.textContent = '';
  gb.hidden = false;
  gb.replaceChildren(
    h('div', { class: 'eyebrow2' }, `My guess · ${r.confidence} confidence`),
    h('b', {}, guessLine({ treatment: r.treatment, stream: r.stream, category: r.category })),
    r.reason ? h('p', { class: 'note', style: 'margin-top:4px' }, r.reason) : null,
    r.question ? h('p', { class: 'ask' }, h('b', {}, 'Question: '), r.question) : null);
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
// `live[i]` is a Map(label -> amount) added up straight from stream i's own sheet (or null to use the saved total).
function parseOverview(rows, live = []) {
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
  const entries = [], dates = [];
  const labels = { income: [], ptr: [], ctr: [], ip: [], other: [] };
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
    if (labels[section] && !/:\s*$/.test(label)) labels[section].push(label);
    entries.push({ section, label, row });
  }
  const items = [];
  for (const { section, label, row } of entries) {
    const amounts = streams.map((s, i) => (live[i] ? (live[i].get(label.toLowerCase()) || 0) : num(row[s.idx])));
    const total = amounts.reduce((a, b) => a + b, 0);
    if (!total && !amounts.some(Boolean)) continue;
    items.push({ section, label, amounts, total, note: notesIdx > 0 && row[notesIdx] ? String(row[notesIdx]) : '' });
  }
  const sum = (secs, i) => items.filter((x) => secs.includes(x.section)).reduce((a, x) => a + (i == null ? x.total : x.amounts[i]), 0);
  const income = sum(['income']), deductions = sum(DEDUCTION_SECTIONS);
  return {
    v: 3, streams: streams.map((s, i) => ({ name: s.name, income: sum(['income'], i), deductions: sum(DEDUCTION_SECTIONS, i), sheet: live[i] ? live[i].sheet : null })),
    streamCols: streams.map((s) => s.idx), items, dates, labels, income, deductions, other: sum(['other']), net: income - deductions,
    live: live.some(Boolean),
  };
}

// Each stream's own sheet is the source of truth (the OVERVIEW sheet just adds them up with SUMIF), so adding the sheets
// ourselves keeps the figures right even when the OVERVIEW's saved totals are out of date.
function liveTotals(XLSX, wb, ws, rows) {
  const probe = parseOverview(rows);
  const labelSet = new Set(Object.values(probe.labels).flat().map((l) => l.toLowerCase()));
  const ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
  return probe.streamCols.map((c) => {
    if (!ref) return null;
    let sheet = null;
    for (let r = 1; r <= ref.e.r && !sheet; r++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      const m = cell && typeof cell.f === 'string' && /SUMIF\(\s*(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!/i.exec(cell.f);
      if (m) sheet = m[1] ? m[1].replace(/''/g, "'") : m[2];
    }
    const sh = sheet && wb.Sheets[sheet];
    if (!sh || !sh['!ref']) return null;
    const e = XLSX.utils.decode_range(sh['!ref']).e;
    const data = XLSX.utils.sheet_to_json(sh, { header: 1, raw: true, defval: null, range: { s: { r: 0, c: 0 }, e } });
    const map = new Map();
    for (let i = 2; i < data.length; i++) {          // SUMIF reads rows 3 onward: category in column B, amount in column F
      const cat = String(data[i] && data[i][1] != null ? data[i][1] : '').trim().toLowerCase();
      if (!cat || !labelSet.has(cat)) continue;
      const amt = num(data[i][5]);
      if (amt) map.set(cat, (map.get(cat) || 0) + amt);
    }
    map.sheet = sheet;
    return map;
  });
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
  const wb = XLSX.read(buf, { type: 'array' });
  const name = wb.SheetNames.find((n) => /^overview$/i.test(n));
  if (!name) throw new Error('The workbook has no OVERVIEW sheet.');
  const ws = wb.Sheets[name];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  let live = [];
  try { live = liveTotals(XLSX, wb, ws, rows); } catch (e) { live = []; }
  return parseOverview(rows, live);
}

async function downloadItem(item) {
  const url = item['@microsoft.graph.downloadUrl'];
  if (url) {
    try { const r = await fetchT(url, {}, 45000); if (r.ok) return r.arrayBuffer(); } catch { /* fall through to the authenticated route */ }
  }
  const token = await getToken();
  const r = await fetchT(`${GRAPH}/me/drive/items/${item.id}/content`, { headers: { Authorization: `Bearer ${token}` } }, 45000);
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
  if (!force && snap && snap.id === item.id && snap.modified === item.lastModifiedDateTime && snap.data && snap.data.v === 3) { state.positionError = ''; renderHome(); return; }
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

/* ---------- refund or bill estimate: rates.json + your assumptions; the maths lives in estimate.js ---------- */
let ratesPromise;
// rates.json holds every 2026-27 figure (verified, with sources). If it cannot be read the round-1 card is shown with a note.
function loadRates() {
  return (ratesPromise = ratesPromise || fetchT('rates.json?v=' + APP_VERSION, { cache: 'no-cache' }, 15000)
    .then((r) => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.json(); })
    .then((j) => { state.rates = j; state.ratesError = ''; return j; })
    .catch((e) => { ratesPromise = null; state.ratesError = friendly(e); return null; }));
}
const currentEstimate = (d, fy) => (state.rates ? Estimate.estimatePosition(d, state.assume.data, state.rates, fy) : null);

// First-time contents of the sheet: everything blank, plus one capital gains row for each sale the workbook's Key Dates
// mention ("Sold ... worth $x" on a date), with the proceeds and sale date filled and the cost base left for the user.
function seededAssumptions(d) {
  const events = ((d && d.dates) || []).filter((x) => /\b(sold|sale|disposed)\b/i.test(x.label)).map((x, i) => {
    const amt = /\$\s?([\d,]+(?:\.\d+)?)/.exec(x.label);
    const asset = x.label.replace(/^\s*sold\s+/i, '').replace(/\s+(worth|for)\s+\$[\d,.]+.*$/i, '').replace(/\s*[,(]?\s*\b(sold|sale|disposed)\b.*$/i, '').trim() || x.label;
    const dt = new Date(x.date);
    return { id: 'keydate-' + i, asset, bought: '', costBase: null, sold: Number.isNaN(dt.getTime()) ? '' : dt.toISOString().slice(0, 10),
      proceeds: amt ? Number(amt[1].replace(/,/g, '')) : null,
      note: `From the workbook key date "${x.label}". Enter the cost base: what it cost plus buying, improving and selling costs. The contract date is the CGT date.` };
  });
  return {
    version: 1, savedAt: null,
    salary: null, paygWithheld: null, paygPerPay: null, paysLeft: null,
    dividends: null, franking: null, interest: null, trust: null, other: null,
    super: null, incomeProtection: null, useStandardDeduction: null, includeCtr: false,
    help: { has: false, balance: null }, hospitalCover: null, family: { status: 'single', children: 0, spouseIncome: null },
    cgt: { carriedLoss: null, events },
  };
}

// Inbox/assumptions.json in OneDrive is the record; the `assumptions` cache lets Position render before it arrives.
// A save that reached the phone but not OneDrive has a newer savedAt than the file (or there is no file yet): that copy wins and is pushed again.
async function loadAssumptions() {
  const as = state.assume;
  as.busy = true;
  try {
    const r = await readInboxFile('assumptions.json');
    if (r) {
      as.etag = r.etag;
      const remote = r.text ? JSON.parse(r.text) : null;
      const local = as.data;
      const phoneNewer = !!(local && local.savedAt && (!remote || !remote.savedAt || local.savedAt > remote.savedAt));
      if (phoneNewer) {
        try { await writeAssumptions(local); toast('Assumptions saved to OneDrive.'); }
        catch (e) { if (e instanceof AuthError) throw e; toast(`Assumptions are on this phone only. Couldn't save to OneDrive: ${friendly(e)}`); }
      } else if (remote) { as.data = remote; ls.set('assumptions', remote); }
    }
    as.loaded = true; as.error = '';
  } catch (e) {
    if (e instanceof AuthError) { as.busy = false; return needSignIn(); }
    as.error = friendly(e);
  }
  as.busy = false;
  if (state.tab === 'home') renderHome();
}
// PUT with If-Match like questions.json; a 404 on the read means the file does not exist yet, so the PUT creates it.
async function writeAssumptions(data) {
  for (let attempt = 0; attempt < 2; attempt++) {
    const r = await readInboxFile('assumptions.json');
    if (!r) throw new Error('The Inbox folder is missing in OneDrive.');
    try {
      const saved = await graph(`/me/drive/items/${r.folderId}:/assumptions.json:/content`, {
        method: 'PUT', body: JSON.stringify(data, null, 2),
        headers: { 'Content-Type': 'application/json', ...(r.etag ? { 'If-Match': r.etag } : {}) },
      });
      state.assume.etag = (saved && saved.eTag) || '';
      return;
    } catch (e) { if (e.status !== 412 || attempt) throw e; }
  }
}
async function saveAssumptions(data) {
  const as = state.assume;
  data.savedAt = isoNow();
  as.data = data; ls.set('assumptions', data);
  $('assume').close('save');
  renderHome();
  as.saving = true;
  try { await writeAssumptions(data); toast('Assumptions saved to OneDrive.'); }
  catch (e) {
    if (e instanceof AuthError) { as.saving = false; return needSignIn(); }
    toast(`Saved on this phone. Couldn't save to OneDrive: ${friendly(e)}`);
  }
  as.saving = false;
}

/* ----- the Assumptions sheet ----- */
const ASSUME_OFFLINE_NOTE = 'Assumptions are from this phone; OneDrive could not be read.';
const numField = (id) => Estimate.num($(id).value);
const setNum = (id, v) => { $(id).value = v == null ? '' : String(v); };
function openAssumptions() {
  const as = state.assume, d = state.position && state.position.data, rates = state.rates;
  const a = as.draft = JSON.parse(JSON.stringify(as.data || seededAssumptions(state.position && state.position.data)));
  for (const [id, k] of [['a-salary', 'salary'], ['a-payg', 'paygWithheld'], ['a-perpay', 'paygPerPay'], ['a-paysleft', 'paysLeft'], ['a-dividends', 'dividends'], ['a-franking', 'franking'],
    ['a-interest', 'interest'], ['a-trust', 'trust'], ['a-other', 'other'], ['a-super', 'super'], ['a-ip', 'incomeProtection']]) setNum(id, a[k]);
  $('a-ctr').checked = !!a.includeCtr;
  $('a-help').checked = !!(a.help && a.help.has); setNum('a-help-balance', a.help && a.help.balance);
  $('a-hospital').checked = a.hospitalCover === true; $('a-hospital').dataset.touched = '';
  $('a-family').value = a.family && a.family.status === 'family' ? 'family' : 'single';
  setNum('a-children', a.family && a.family.children); setNum('a-spouse', a.family && a.family.spouseIncome);
  setNum('a-carried', a.cgt && a.cgt.carriedLoss);
  delete $('a-std').dataset.touched;
  /* the workbook's work-related rows, and what the standard deduction would replace */
  const parts = Estimate.workbookParts(d);
  const wr = $('a-wr'); wr.replaceChildren();
  if (parts.workRelated.length) {
    wr.append(h('p', { class: 'note', style: 'margin-bottom:4px' }, 'Work-related rows in the workbook'));
    for (const it of parts.workRelated) wr.append(kv(shortLabel(it.label), money(it.total, true)));
    wr.append(kv('Work-related total', money(parts.workRelated.reduce((s, i) => s + i.total, 0), true), 'total'));
  } else wr.append(h('p', { class: 'note' }, 'No work-related rows in the workbook yet.'));
  if (rates) {
    $('a-std-label').textContent = `Use the ${money(rates.standardDeduction.max)} standard deduction instead`;
    $('a-std-note').textContent = `For work-related expenses with no receipts, when you have salary or director fees. The default is on when the workbook's work-related rows are under ${money(rates.standardDeduction.max)}.`;
  }
  const ctrTotal = parts.ctr.reduce((s, i) => s + i.total, 0);
  $('a-ctr-total').textContent = ctrTotal ? `Currently ${money(ctrTotal, true)} is left out.` : 'Nothing recorded there yet.';
  $('a-status').textContent = as.error ? ASSUME_OFFLINE_NOTE : '';
  $('a-status').hidden = !as.error;
  refreshAssumptionsSheet();
  renderCgtRows();
  $('assume').showModal();
  $('assume').focus();   // the dialog itself, not the first field, so the keyboard does not pop on open
  $('assume').scrollTop = 0;
}
// Live bits of the sheet: the projected PAYG total, the standard-deduction default, which fields apply.
function refreshAssumptionsSheet() {
  const d = state.position && state.position.data, rates = state.rates;
  $('a-payg-total').textContent = money(Estimate.projectedPayg({ paygWithheld: numField('a-payg'), paygPerPay: numField('a-perpay'), paysLeft: numField('a-paysleft') }), true);
  const a = state.assume.draft || {};
  if (rates && !$('a-std').dataset.touched) $('a-std').checked = a.useStandardDeduction == null ? Estimate.standardDeductionDefault(d, { salary: numField('a-salary') }, rates) : !!a.useStandardDeduction;
  $('a-help-balance-wrap').hidden = !$('a-help').checked;
  $('a-spouse-wrap').hidden = $('a-family').value !== 'family';
}
function blankEvent() { return { id: 'ev-' + Date.now().toString(36), asset: '', bought: '', costBase: null, sold: '', proceeds: null, note: '' }; }
function renderCgtRows() {
  const box = $('a-cgt-rows'), a = state.assume.draft, rates = state.rates;
  box.replaceChildren();
  a.cgt = a.cgt || { carriedLoss: null, events: [] };
  if (!a.cgt.events.length) box.append(h('p', { class: 'note' }, 'No capital gains events. Add one for each asset sold this year.'));
  a.cgt.events.forEach((ev, i) => {
    const calc = h('div', { class: 'cgt-calc' });
    const update = () => {
      calc.replaceChildren();
      if (!rates) return;
      const row = Estimate.cgtRow(ev, rates);
      if (row.gain == null) { calc.append(h('span', { class: 'pill warn' }, row.missingCostBase ? 'Cost base missing' : 'Proceeds missing')); return; }
      calc.append(h('b', {}, `${row.gain >= 0 ? 'Gain' : 'Loss'} ${money(Math.abs(row.gain), true)}`));
      if (row.heldOver12Months != null) calc.append(h('span', { class: 'pill' + (row.heldOver12Months ? '' : ' warn') }, row.heldOver12Months ? 'Held over 12 months' : 'Held under 12 months'));
      calc.append(h('span', { class: 'pill' + (row.discountApplies ? '' : ' warn') }, row.discountApplies ? `${pct(rates.cgt.discount)} discount applies` : row.gain > 0 && row.heldOver12Months == null ? 'Enter both dates for the discount test' : 'No discount'));
    };
    const field = (label, key, type, extra = {}) => h('label', { class: 'field' }, label,
      h('input', { type, value: ev[key] == null ? '' : String(ev[key]), autocomplete: 'off', ...extra,
        oninput: (e) => { ev[key] = type === 'date' || key === 'asset' ? e.target.value : Estimate.num(e.target.value); if (key !== 'asset') update(); } }));
    update();
    box.append(h('div', { class: 'cgt-row' },
      field('Asset', 'asset', 'text', { placeholder: 'e.g. ABC shares' }),
      h('div', { class: 'row2' }, field('Bought', 'bought', 'date'), field('Cost base ($)', 'costBase', 'text', { inputmode: 'decimal', placeholder: '0.00' })),
      h('div', { class: 'row2' }, field('Sold (contract date)', 'sold', 'date'), field('Proceeds ($)', 'proceeds', 'text', { inputmode: 'decimal', placeholder: '0.00' })),
      ev.note ? h('p', { class: 'note' }, ev.note) : null,
      h('div', { class: 'cgt-foot' }, calc, h('button', { class: 'btn', type: 'button', onclick: () => { a.cgt.events.splice(i, 1); renderCgtRows(); } }, 'Remove'))));
  });
}
function collectAssumptions() {
  const a = state.assume.draft;
  const touched = !!$('a-std').dataset.touched;
  return {
    ...a, version: 1,
    salary: numField('a-salary'), paygWithheld: numField('a-payg'), paygPerPay: numField('a-perpay'), paysLeft: numField('a-paysleft'),
    dividends: numField('a-dividends'), franking: numField('a-franking'), interest: numField('a-interest'), trust: numField('a-trust'), other: numField('a-other'),
    super: numField('a-super'), incomeProtection: numField('a-ip'),
    useStandardDeduction: touched ? $('a-std').checked : a.useStandardDeduction == null ? null : $('a-std').checked,
    includeCtr: $('a-ctr').checked,
    help: { has: $('a-help').checked, balance: numField('a-help-balance') },
    hospitalCover: $('a-hospital').dataset.touched ? $('a-hospital').checked : a.hospitalCover,   // null stays null until the box is touched
    family: { status: $('a-family').value, children: numField('a-children') || 0, spouseIncome: numField('a-spouse') },
    cgt: { carriedLoss: numField('a-carried'), events: (a.cgt && a.cgt.events) || [] },
  };
}
function wireAssumptions() {
  for (const id of ['a-payg', 'a-perpay', 'a-paysleft', 'a-salary']) $(id).addEventListener('input', refreshAssumptionsSheet);
  $('a-std').addEventListener('change', () => { $('a-std').dataset.touched = '1'; });
  $('a-hospital').addEventListener('change', () => { $('a-hospital').dataset.touched = '1'; });
  $('a-help').addEventListener('change', refreshAssumptionsSheet);
  $('a-family').addEventListener('change', refreshAssumptionsSheet);
  $('a-cgt-add').addEventListener('click', () => { state.assume.draft.cgt.events.push(blankEvent()); renderCgtRows(); });
  $('a-cancel').addEventListener('click', () => $('assume').close('cancel'));
  $('a-save').addEventListener('click', () => saveAssumptions(collectAssumptions()));
  // Close only on a true backdrop tap: a tap on the sheet's own rim (target is the dialog, but inside its box) does nothing.
  $('assume').addEventListener('click', (e) => {
    const dlg = $('assume');
    if (e.target !== dlg) return;
    const r = dlg.getBoundingClientRect();
    const outside = e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom;
    if (outside) dlg.close('cancel');
  });
}

/* ----- Position cards ----- */
function estimateHero(est, fy) {
  const refund = est.result >= 0;
  const today = new Date().toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: 'numeric' });
  return h('div', { class: 'hero' },
    h('div', { class: 'eyebrow' + (refund ? ' refund' : '') }, refund ? 'Estimated refund' : 'Estimated amount owing'),
    h('div', { class: 'big ' + (refund ? 'refund' : 'owing') }, money(Math.abs(est.result))),
    h('div', { class: 'sub' }, `FY${fy % 100}, as at today (${today})`),
    h('div', { class: 'sub', style: 'margin-top:4px' }, `Taxable income ${money(est.taxableIncome, true)} (estimate) · ${Math.round(est.marginal * 1000) / 10}% on your next dollar`));
}
function renderChainCard(box, est, rates, fy) {
  box.append(sect('How it is worked out'));
  const rows = est.chain.filter((r) => !(r.optional && !r.value)).map((r) => {
    const cls = r.kind === 'subtotal' ? 'total' : r.kind === 'result' ? 'total result' + (r.value < 0 ? ' owing' : '') : '';
    return kv(r.label, money(r.kind === 'result' ? Math.abs(r.value) : r.value, true), cls);
  });
  if (est.notDeductions) rows.push(kv('Not deductions (offsets in the workbook)', money(est.notDeductions, true)));
  const pills = est.warnings.length ? h('div', { class: 'pills' }, ...est.warnings.map((w) => h('span', { class: 'pill warn' }, w.text))) : null;
  const provisional = est.provisional.map((p) => `${p.label}: ${p.year} figure, ${rates.label} not yet published`);
  box.append(card(...rows, pills,
    h('p', { class: 'note', style: 'margin-top:10px' }, h('span', { class: 'pill warn' }, 'Estimate'), ' ',
      `Resident rates for ${rates.label}; the workbook figures are as recorded so far and the salary, PAYG and other figures come from your assumptions. Trust distributions and reportable fringe benefits can move this, so check it with your accountant.`),
    h('p', { class: 'note', style: 'margin-top:8px' }, `Rates verified ${fmtYmd(rates.verifiedOn)}.${provisional.length ? ' ' + provisional.join('. ') + '.' : ''}`),
    h('button', { class: 'btn block', type: 'button', style: 'margin-top:12px', onclick: openAssumptions }, 'Edit assumptions')));
}
function renderCgtCard(box, est, rates) {
  const c = est.cgt;
  if (!c.rows.length) return;
  box.append(sect('Capital gains'));
  const wrap = card();
  for (const row of c.rows) {
    const when = [row.bought && `bought ${fmtYmd(row.bought)}`, row.sold && `sold ${fmtYmd(row.sold)}`].filter(Boolean).join(', ');
    wrap.append(h('div', { class: 'stream' },
      h('div', { class: 'stream-head' }, h('span', { class: 'stream-name' }, row.asset || 'Asset'),
        h('b', {}, row.gain == null ? '' : `${row.gain < 0 ? 'Loss ' : 'Gain '}${money(Math.abs(row.gain), true)}`)),
      when ? h('small', {}, when) : null,
      h('div', { class: 'pills', style: 'margin-top:6px' },
        row.missingCostBase ? h('span', { class: 'pill warn' }, 'Cost base missing') : null,
        !row.missingCostBase && row.proceeds == null ? h('span', { class: 'pill warn' }, 'Proceeds missing') : null,
        row.heldOver12Months != null ? h('span', { class: 'pill' + (row.heldOver12Months ? '' : ' warn') }, row.heldOver12Months ? 'Held over 12 months' : 'Held under 12 months') : null,
        row.gain != null && row.gain > 0 ? h('span', { class: 'pill' + (row.discountApplies ? '' : ' warn') }, row.discountApplies ? `${pct(rates.cgt.discount)} discount applied` : row.heldOver12Months == null ? 'Enter both dates for the discount test' : 'No discount') : null,
        !row.inYear ? h('span', { class: 'pill warn' }, 'Sold outside this year') : null)));
  }
  const totals = [
    kv('Gains', money(c.totalGains, true)),
    c.yearLosses ? kv('Losses this year', money(-c.yearLosses, true)) : null,
    c.carriedLoss ? kv('Loss carried forward', money(-c.carriedLoss, true)) : null,
    c.discount ? kv(`${pct(rates.cgt.discount)} discount (held over ${rates.cgt.holdMonths} months)`, money(-c.discount, true)) : null,
    kv('Net capital gain', money(c.net, true), 'total'),
  ].filter(Boolean);
  wrap.append(h('div', { style: 'margin-top:10px;border-top:1px solid var(--line);padding-top:6px' }, ...totals));
  wrap.append(h('p', { class: 'note', style: 'margin-top:8px' }, 'Losses come off first, then the discount. Selling costs go in the cost base, not deductions.'));
  box.append(wrap);
}
function renderAssumptionsPrompt(box) {
  box.append(card(
    h('b', {}, 'Add your salary and PAYG to see your refund or bill'),
    h('p', { class: 'muted', style: 'margin:6px 0 10px' }, state.ratesError
      ? `The rates file could not be loaded (${state.ratesError}), so only the indicative figure below is available. Try again later.`
      : 'The workbook has your deductions and side income. Add what it cannot know and the Position tab shows your estimated refund or amount owing, with every step listed.'),
    h('button', { class: 'btn primary block', type: 'button', disabled: !state.rates, onclick: openAssumptions }, 'Assumptions')));
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
  const est = currentEstimate(d, fy);
  if (est && est.hasAssumptions) {
    lines.push(`Estimated ${est.result >= 0 ? 'refund' : 'amount owing'}: ${money(Math.abs(est.result), true)} (an estimate as at ${new Date().toLocaleDateString('en-AU', { day: 'numeric', month: 'long', year: 'numeric' })})`,
      `Taxable income (estimate): ${money(est.taxableIncome, true)}`,
      `Worked out as: ${est.chain.filter((r) => ['tax', 'offset', 'credit'].includes(r.kind) && r.value).map((r) => `${r.label} ${money(r.value, true)}`).join(', ')}.`,
      `Assumptions used: ${est.assumptionsUsed.join('; ')}.`);
    if (est.warnings.length) lines.push(`Notes: ${est.warnings.map((w) => w.text).join('; ')}.`);
    lines.push('');
  }
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
  add(`resolutions-${fy}`, 'Sign trust distribution resolutions (due 30 June)', Date.UTC(fy, 5, 15), 'Income resolutions must be made by 30 June.');
  add(`gain-streaming-${fy}`, 'Record in writing who receives any capital gain (due 31 August)', Date.UTC(fy, 7, 15), 'Needed if a trust sells property this year.');
  add(`lodge-${fy}`, `Lodgement: confirm the plan with ${acc}`, Date.UTC(fy, 9, 1), 'The self-lodgement deadline is 31 October.');
  const a = state.assume.data;
  if (!(a && (Estimate.num(a.salary) > 0 || Estimate.projectedPayg(a) > 0))) {
    add(`assumptions-${fy}`, 'Enter your salary and PAYG in Assumptions', Date.UTC(fy - 1, 9, 31), 'Open Assumptions on the Position tab so your refund or bill can be estimated.');
  }
  for (const ev of (a && a.cgt && a.cgt.events) || []) {
    if (Estimate.num(ev.costBase) == null) add(`costbase-${slug(ev.asset || ev.id || 'asset')}`, `Find the cost base: ${ev.asset || 'asset'}`, null, 'Price paid plus buying costs, improvements and selling costs. Enter it in Assumptions.');
  }
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


/* ---------- inbox: receipts Claude found in Gmail, waiting for your say-so ---------- */
const TREATMENTS = {
  claim: 'Claim it',
  xero: 'Company cost (not my return)',
  reimbursed: 'Claimed back from Oakwood',
  personal: 'Personal, not claimable',
  cgt: 'Part of the capital gain',
  skip: 'Not a receipt',
};
const FALLBACK_CATEGORIES = ['Cost of Managing Tax Affairs', 'Council Rates', 'Insurance: Landlord Insurance', 'Interest on Loans', 'Repairs and Maintenance: General Repairs', 'Water Charges', 'Other Work-Related Expenses', 'Home Office Expenses', 'Gifts or Donations'];
state.inbox = { data: null, error: '', busy: false, saving: '', thinking: '', drafts: {}, filter: 'review', open: null };

const inboxFolder = () => state.rootItems.find((i) => i.folder && i.name.toLowerCase() === cfg().inboxFolder.toLowerCase());
const fmtYmd = (s) => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? `${Number(m[3])} ${MONTHS[Number(m[2]) - 1]} ${m[1]}` : ''; };
const inboxPending = (d) => ((d && d.items) || []).filter((i) => i.status === 'pending' && i.relevance !== 'no');

async function fetchInbox() {
  const f = inboxFolder();
  if (!f) return { folderId: '', data: null, etag: '' };
  let meta;
  try { meta = await graph(`/me/drive/items/${f.id}:/inbox.json`); }
  catch (e) { if (e.status === 404) return { folderId: f.id, data: null, etag: '' }; throw e; }
  const buf = await downloadItem(meta);
  return { folderId: f.id, data: JSON.parse(new TextDecoder().decode(buf)), etag: meta.eTag || '' };
}
async function loadInbox() {
  const ib = state.inbox;
  ib.busy = true; ib.error = '';
  if (state.tab === 'inbox') renderInbox();
  try { const r = await fetchInbox(); ib.data = r.data; }
  catch (e) { if (e instanceof AuthError) { ib.busy = false; return needSignIn(); } ib.error = friendly(e); }
  ib.busy = false;
  updateInboxBadge();
  if (state.tab === 'inbox') renderInbox();
  if (state.tab === 'home') renderHome();
  refreshPositionIfLogged();
}
// When the Mac has added receipts to the workbook since this phone last read it, fetch the workbook again so Position shows the new figures.
async function refreshPositionIfLogged() {
  const items = (state.inbox.data && state.inbox.data.items) || [];
  const last = Math.max(0, ...items.filter((i) => i.logged && i.logged.sheet && !i.logged.duplicate).map((i) => Date.parse(i.loggedAt) || 0));
  if (!last || !state.stack[0] || (state.position && state.position.at > last + 120000)) return;
  try {
    const items2 = await listChildren(state.stack[0].id);
    state.rootItems = items2; state.rootFolders = items2.filter((i) => i.folder);
    await refreshPosition();
  } catch (e) { /* the next open tries again */ }
}
// Read the latest file, let `fn` change it, write it back (retrying once if the file changed underneath us).
async function mutateInbox(fn) {
  for (let attempt = 0; attempt < 2; attempt++) {
    const r = await fetchInbox();
    if (!r.folderId) throw new Error('The Inbox folder is missing in OneDrive.');
    const data = r.data || { version: 1, source: 'gmail', items: [], skippedIds: [] };
    fn(data);
    data.updated = new Date().toISOString();
    try {
      await graph(`/me/drive/items/${r.folderId}:/inbox.json:/content`, {
        method: 'PUT', body: JSON.stringify(data, null, 2),
        headers: { 'Content-Type': 'application/json', ...(r.etag ? { 'If-Match': r.etag } : {}) },
      });
      state.inbox.data = data; updateInboxBadge();
      return data;
    } catch (e) { if (e.status !== 412 || attempt) throw e; }
  }
}
async function updateItem(id, change) {
  await mutateInbox((data) => {
    const it = (data.items || []).find((x) => x.id === id);
    if (!it) throw new Error('That receipt is no longer in the inbox.');
    change(it);
  });
}
async function decide(id, decision) {
  const ib = state.inbox; ib.saving = id; renderInbox();
  try {
    await updateItem(id, (it) => {
      if (decision) { it.decision = { ...decision, at: new Date().toISOString() }; it.status = 'decided'; }
      else { delete it.decision; it.status = 'pending'; }
    });
    ib.open = null; updateInboxBadge();
    toast(decision ? `Saved: ${TREATMENTS[decision.treatment]}` : 'Moved back to review');
  } catch (e) { if (e instanceof AuthError) { ib.saving = ''; return needSignIn(); } banner(friendly(e)); }
  ib.saving = ''; renderInbox();
}
async function bringBack(id) {
  state.inbox.saving = id; renderInbox();
  try { await updateItem(id, (it) => { it.relevance = 'maybe'; }); updateInboxBadge(); }
  catch (e) { if (e instanceof AuthError) { state.inbox.saving = ''; return needSignIn(); } banner(friendly(e)); }
  state.inbox.saving = ''; renderInbox();
}
function updateInboxBadge() {
  const n = inboxPending(state.inbox.data).length, b = $('inbox-badge');
  b.textContent = String(n); b.hidden = !n;
}

function categoryOptions() {
  const l = state.position && state.position.data && state.position.data.labels;
  const fromSheet = l ? [...l.ptr, ...l.ip].map(shortLabel) : [];
  return [...new Set([...fromSheet, ...FALLBACK_CATEGORIES])];
}
function streamOptions() {
  const d = state.position && state.position.data;
  return d ? d.streams.map((x) => streamLabel(x.name).main) : [];
}
function guessLine(g) {
  return [TREATMENTS[g.treatment] || 'Not sure', g.stream, g.category].filter(Boolean).join(' · ');
}

function editorFor(it) {
  const dec = it.decision || it.suggestion || {};
  const tSel = h('select', { id: 'ed-treatment', 'aria-label': 'What to do with it' }, ...Object.entries(TREATMENTS).map(([k, v]) => h('option', { value: k }, v)));
  tSel.value = dec.treatment || 'claim';
  const sSel = h('select', { id: 'ed-stream', 'aria-label': 'Which income stream' }, h('option', { value: '' }, 'Not sure which stream'), ...streamOptions().map((n) => h('option', { value: n }, n)));
  sSel.value = dec.stream || '';
  const dl = h('datalist', { id: 'ed-cats' }, ...categoryOptions().map((c) => h('option', { value: c })));
  const cat = h('input', { type: 'text', id: 'ed-category', list: 'ed-cats', placeholder: 'Category, e.g. Council Rates', autocomplete: 'off', value: dec.category || '' });
  return h('div', { class: 'stack', style: 'margin-top:10px' },
    h('label', { class: 'field' }, 'What should happen to it?', tSel),
    h('label', { class: 'field' }, 'Income stream', sSel),
    h('label', { class: 'field' }, 'Category', cat, dl),
    h('button', { class: 'btn primary block', type: 'button', onclick: () => decide(it.id, {
      treatment: tSel.value, stream: sSel.value, category: cat.value.trim(), note: ($(`note-${it.id}`) || {}).value || '',
    }) }, 'Save my choice'));
}

// Where a decided receipt has got to: filed under Receipts, and added to the workbook or the capital-gains list.
// The filing and logging are done on the Mac by Inbox/process_decisions.py, which stamps filedAt / loggedAt on the item.
function processStatus(it) {
  const d = it.decision || {}, lg = it.logged || {}, f = it.file || {};
  const where = [(f.folder || '').replace(/\s*\/\s*/g, '/'), f.name].filter(Boolean).join('/');
  const lines = [];
  const filed = !!(it.filedAt || f.webUrl);
  if (filed && f.name) lines.push(`Filed: ${where}`);
  if (lg.sheet) lines.push(lg.duplicate ? `Already in the workbook (${lg.sheet}, row ${lg.row}), so I did not add it again.` : `Added to the workbook: ${lg.sheet}, row ${lg.row}.`);
  else if (lg.csv) lines.push('Added to your capital gains costs list (cgt-costs.csv), not to the deductions.');
  if (it.warnNote) lines.push(`Heads up: ${it.warnNote}`);
  if (it.processNote && !it.loggedAt) lines.push(`Not in the workbook yet: ${it.processNote}.`);
  else if (!filed) lines.push(d.treatment === 'claim' ? 'Waiting to be filed and added to the workbook.' : 'Waiting to be filed.');
  else if (!it.loggedAt && (d.treatment === 'claim' || d.treatment === 'cgt')) lines.push(d.treatment === 'claim' ? 'Waiting to be added to the workbook.' : 'Waiting to be added to your capital gains list.');
  return { lines, canUndo: !filed && !f.name };
}

function receiptCard(it, mode) {
  const g = it.suggestion || {}, saving = state.inbox.saving === it.id;
  const top = [
    h('div', { class: 'stream-head' }, h('span', { class: 'stream-name' }, it.vendor || it.from || 'Unknown sender'), h('b', {}, it.amount != null ? money(it.amount, true) : '—')),
    h('small', { style: 'display:block;color:var(--muted);font-size:13px' }, [fmtYmd(it.date), it.docType, it.description].filter(Boolean).join(' · ')),
  ];
  const docUrl = it.file && it.file.webUrl;
  const links = (it.viewUrl || docUrl) ? h('a', { class: 'chip', style: 'margin-top:8px', href: docUrl || it.viewUrl, target: '_blank', rel: 'noopener' }, docUrl ? 'Open the document ↗' : 'Open the email ↗') : null;

  if (mode === 'done') {
    const d = it.decision || {}, st = processStatus(it);
    return h('div', { class: 'rcpt' }, ...top,
      h('div', { style: 'margin-top:6px' }, h('span', { class: 'pill' + (d.treatment === 'claim' ? '' : ' warn') }, TREATMENTS[d.treatment] || 'Decided'), ' ',
        h('small', {}, [d.stream, d.category].filter(Boolean).join(' · '))),
      d.note ? h('p', { class: 'note', style: 'margin-top:4px' }, `Your note: ${d.note}`) : null,
      ...st.lines.map((l) => h('p', { class: 'note', style: 'margin-top:4px' }, l)),
      h('div', { class: 'rcpt-actions' }, links, st.canUndo ? h('button', { class: 'chip', type: 'button', disabled: saving, onclick: () => decide(it.id, null) }, 'Undo') : null));
  }
  if (mode === 'notax') {
    return h('div', { class: 'rcpt' }, ...top,
      g.reason ? h('p', { class: 'note', style: 'margin-top:4px' }, g.reason) : null,
      h('div', { class: 'rcpt-actions' }, links,
        h('button', { class: 'chip', type: 'button', disabled: saving, onclick: () => bringBack(it.id) }, 'Actually, review this'),
        h('button', { class: 'chip', type: 'button', disabled: saving, onclick: () => decide(it.id, { treatment: 'skip', stream: '', category: '', note: '' }) }, 'Confirm: not tax')));
  }
  const open = state.inbox.open === it.id;
  const thinking = state.inbox.thinking === it.id, busy = saving || thinking;
  const note = h('input', {
    type: 'text', id: `note-${it.id}`, value: state.inbox.drafts[it.id] || '', autocomplete: 'off', enterkeyhint: 'send',
    placeholder: g.question ? 'Type your answer here' : 'Add a note',
    oninput: (e) => { state.inbox.drafts[it.id] = e.target.value; },
    onkeydown: (e) => { if (e.key === 'Enter') { e.preventDefault(); rethink(it); } },
  });
  const ask = g.question ? h('p', { class: 'ask' }, h('b', {}, 'Question: '), g.question) : null;
  return h('div', { class: 'rcpt' }, ...top,
    h('div', { class: 'guess' },
      h('div', { class: 'eyebrow2' }, `My guess${g.confidence ? ` · ${g.confidence} confidence` : ''}`),
      h('b', {}, guessLine(g)),
      g.reason ? h('p', { class: 'note', style: 'margin-top:4px' }, g.reason) : null,
      ask,
      it.answer ? h('p', { class: 'note', style: 'margin-top:6px' }, h('b', {}, 'You said: '), it.answer) : null),
    h('div', { class: 'noterow' }, note,
      h('button', { class: 'btn orange', type: 'button', disabled: busy, onclick: () => rethink(it) }, thinking ? 'Thinking…' : 'Send')),
    h('div', { class: 'rcpt-actions' },
      g.treatment ? h('button', { class: 'btn primary', type: 'button', disabled: busy, onclick: () => decide(it.id, { treatment: g.treatment, stream: g.stream || '', category: g.category || '', note: (note.value || it.answer || '') }) }, saving ? 'Saving…' : "Yes, that's right") : null,
      h('button', { class: 'btn', type: 'button', disabled: busy, onclick: () => { state.inbox.open = open ? null : it.id; renderInbox(); } }, open ? 'Close' : 'Something else')),
    open ? editorFor(it) : null,
    links ? h('div', {}, links) : null);
}

// ----- Send: save Dylan's answer and let Claude re-decide the item with it -----
const RETHINK_SCHEMA = {
  type: 'object', additionalProperties: false,
  required: ['treatment', 'stream', 'category', 'confidence', 'reason', 'question'],
  properties: {
    treatment: { type: 'string', enum: ['claim', 'xero', 'reimbursed', 'personal', 'cgt', 'skip', 'unsure'] },
    stream: { type: 'string' }, category: { type: 'string' }, confidence: { type: 'string', enum: ['low', 'medium', 'high'] },
    reason: { type: 'string' }, question: { anyOf: [{ type: 'string' }, { type: 'null' }] },
  },
};
async function claudeJson(base, schema) {
  const withSchema = { ...base, output_config: { format: { type: 'json_schema', schema } } };
  let data, lastErr;
  for (const go of [() => postClaude(withSchema, true), () => postClaude(withSchema, false), () => postClaude(base, false)]) {
    try { data = await go(); break; } catch (e) { lastErr = e; if (e.status !== 400) throw e; }
  }
  if (!data) throw lastErr;
  if (data.stop_reason === 'refusal') throw new Error("Claude wouldn't answer that one.");
  const text = (data.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('').trim();
  try { return JSON.parse(text); } catch { const m = /\{[\s\S]*\}/.exec(text); if (!m) throw new Error('Claude sent back something I could not read. Try again.'); return JSON.parse(m[0]); }
}
async function rethinkItem(it, answer) {
  const system = buildExtractPrompt() + '\n\nYOUR JOB NOW: you are not reading a document. You are re-deciding ONE receipt after Dylan answered your question. Treat his answer as a fact. If it settles the treatment, set confidence to high or medium and question to null. If it opens a new uncertainty, ask one specific follow-up question instead.';
  const item = { vendor: it.vendor, date: it.date, amount: it.amount, gst: it.gst, payment: it.payment, description: it.description, docType: it.docType, ref: it.ref, yourPreviousGuess: it.suggestion || null, dylansAnswer: answer };
  const r = await claudeJson({ model: ASK_MODEL, max_tokens: 3000, system, messages: [{ role: 'user', content: 'Re-decide this receipt and return the JSON.\n' + JSON.stringify(item, null, 2) }] }, RETHINK_SCHEMA);
  return { treatment: r.treatment === 'unsure' ? null : r.treatment, stream: r.stream || '', category: r.category || '', confidence: r.confidence, reason: r.reason, question: r.question || null };
}
async function rethink(it) {
  const ib = state.inbox, input = $(`note-${it.id}`);
  const text = ((input && input.value) || ib.drafts[it.id] || '').trim();
  if (!text) { toast('Type your answer first, then tap Send.'); return; }
  if (ib.saving === it.id || ib.thinking === it.id) return;
  ib.drafts[it.id] = text; ib.thinking = it.id; renderInbox();
  try {
    let suggestion = null;
    if (cfg().apiKey) { if (!state.ask.loaded) await loadAskContext(); suggestion = await rethinkItem(it, text); }
    await updateItem(it.id, (x) => {
      x.answer = text; x.answeredAt = new Date().toISOString();
      if (suggestion) { x.suggestion = suggestion; x.answerHandled = true; } else x.answerHandled = false;
    });
    delete ib.drafts[it.id];
    toast(suggestion ? 'Updated my guess from your answer.' : 'Saved your answer. Add a Claude API key in Settings so I can re-think it.');
  } catch (e) {
    if (e instanceof AuthError) { ib.thinking = ''; return needSignIn(); }
    banner(`Couldn't send that: ${askErrorText(e)}`);
  }
  ib.thinking = ''; renderInbox();
}

function renderInbox() {
  const box = $('inbox'), ib = state.inbox;
  const keepFocus = document.activeElement && document.activeElement.id;
  box.replaceChildren();
  if (ib.busy && !ib.data) { box.append(h('div', { class: 'state' }, h('div', { class: 'spinner' }), 'Opening your inbox…')); return; }
  if (ib.error && !ib.data) { box.append(card(h('b', {}, "Couldn't open the inbox"), h('p', { class: 'muted' }, ib.error), h('button', { class: 'btn block', type: 'button', style: 'margin-top:10px', onclick: loadInbox }, 'Try again'))); return; }
  if (!ib.data) {
    box.append(card(h('b', {}, 'Nothing here yet'),
      h('p', { class: 'muted', style: 'margin-top:6px' }, 'When Claude scans your Gmail for receipts they land here, each with a guess at what it is for and a question where it is not obvious. Ask Claude to "sync my Gmail receipts".')));
    return;
  }
  const items = [...(ib.data.items || [])].sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  const review = items.filter((i) => i.status === 'pending' && i.relevance !== 'no');
  const done = items.filter((i) => i.status === 'decided');
  const notax = items.filter((i) => i.status === 'pending' && i.relevance === 'no');
  const f = ib.filter;
  box.append(h('div', { class: 'hero' },
    h('div', { class: 'eyebrow' }, 'Receipts from Gmail'),
    h('div', { class: 'big' }, `${review.length} to review`),
    h('div', { class: 'sub' }, ib.data.updated ? `Last scanned ${fmtYmd(ib.data.syncedThrough || ib.data.updated)} · ${items.length} found this year` : `${items.length} found this year`)));
  box.append(h('div', { class: 'chips', style: 'margin:0' },
    ...[['review', `To review (${review.length})`], ['done', `Done (${done.length})`], ['notax', `Probably not tax (${notax.length})`]]
      .map(([k, label]) => h('button', { class: 'chip', type: 'button', 'aria-pressed': String(f === k), onclick: () => { ib.filter = k; ib.open = null; renderInbox(); } }, label)),
    h('button', { class: 'chip', type: 'button', onclick: loadInbox }, ib.busy ? 'Refreshing…' : 'Refresh')));
  const list = f === 'done' ? done : f === 'notax' ? notax : review;
  if (!list.length) {
    box.append(h('div', { class: 'state' }, f === 'review' ? 'All caught up. Nothing waiting for you.' : 'Nothing here.'));
  } else {
    const wrap = card();
    for (const it of list) wrap.append(receiptCard(it, f));
    box.append(wrap);
  }
  if (keepFocus) { const el = document.getElementById(keepFocus); if (el) el.focus(); }
}

function renderInboxPrompt(box) {
  const n = inboxPending(state.inbox.data).length;
  if (!n) return;
  box.append(h('button', { class: 'btn orange block', type: 'button', onclick: () => $('tab-inbox').click() }, `${n} receipt${n === 1 ? '' : 's'} from Gmail to review`));
}


/* ---------- ask: questions answered by Claude, who knows your setup ---------- */
const ASK_MODEL = 'claude-opus-5-5';
const ASK_API = 'https://api.anthropic.com/v1/messages';
state.ask = { thread: ls.get('askThread', []), busy: false, view: 'chat', briefing: null, profile: '', loaded: false };

async function readInboxFile(name) {
  const f = inboxFolder();
  if (!f) return null;
  let meta;
  try { meta = await graph(`/me/drive/items/${f.id}:/${encodeURIComponent(name)}`); }
  catch (e) { if (e.status === 404) return { folderId: f.id, text: null, etag: '' }; throw e; }
  return { folderId: f.id, text: new TextDecoder().decode(await downloadItem(meta)), etag: meta.eTag || '' };
}
async function loadAskContext() {
  const a = state.ask;
  try {
    const [p, b] = await Promise.all([readInboxFile('profile.md'), readInboxFile('briefing.json')]);
    a.profile = (p && p.text) || '';
    a.briefing = b && b.text ? JSON.parse(b.text) : null;
    a.loaded = true;
  } catch (e) { if (e instanceof AuthError) return needSignIn(); }
  syncQueue();
  if (state.tab === 'ask') renderAsk();
}

const persistAsk = () => ls.set('askThread', state.ask.thread.slice(-30));
const isoNow = () => new Date().toISOString();

// ----- the system prompt: generic rules in code, everything personal comes from your OneDrive at question time -----
function buildSystemPrompt() {
  const a = state.ask, p = state.position, d = p && p.data, fy = fyEnd();
  const today = new Date().toLocaleDateString('en-AU', { day: 'numeric', month: 'long', year: 'numeric' });
  const lines = [
    `You are the personal tax assistant inside Dylan's tax app. Today is ${today}. The current Australian financial year is 1 July ${fy - 1} to 30 June ${fy}. Address Dylan as "you". Be direct, warm and plain-English, with Australian spelling.`,
    `You are not a registered tax agent. You give general information and a reasoned view on his own situation. Say so in a few words when it matters, and point to ${cfg().accountantName} (his accountant) for anything material or uncertain. Never pretend to certainty.`,
    '',
    'HOW TO ANSWER',
    '1. Start with a verdict in bold, one of: "Likely claimable", "Probably not claimable", "It depends", "Not yours to claim (it belongs to the business or a trust)", "Part of the capital gain, not a deduction", or "Not tax related".',
    '2. Then two to five short bullets or sentences: the rule that decides it, applied to HIS situation, naming the income stream, card, property or entity involved.',
    '3. Then "To be sure:" with one to three specific questions whose answers would change the verdict.',
    '4. Then "Keep:" the record to keep, and "Log it:" where it goes in his workbook (income stream and category), or say it is a company cost when it is a business-card or company expense (he does not use Xero, so never mention it).',
    'Keep answers under about 250 words unless he asks for more. Never invent figures: use only numbers in the data below or that he gives you. If a rate, threshold or rule for the current year is not in the data below and you are not sure of it, say so and say what to check. If a question needs a fact you do not have, ask for it instead of assuming.',
    '',
    'ABOUT DYLAN (his profile, written by him and Claude; trust it over guesses)',
    a.profile || '(No profile file found. Ask him a short question about his structure before answering anything that depends on it.)',
  ];
  if (d) {
    lines.push('', `HIS NUMBERS FROM THE WORKBOOK ${p.name} (read ${new Date(p.at).toLocaleDateString('en-AU')})`,
      `Income recorded ${money(d.income, true)}, deductions recorded ${money(d.deductions, true)}, net ${money(d.net, true)}.`);
    for (const s of d.streams.filter((x) => x.income || x.deductions)) lines.push(`- ${streamLabel(s.name).main}: income ${money(s.income, true)}, deductions ${money(s.deductions, true)}`);
    for (const it of d.items.filter((x) => DEDUCTION_SECTIONS.includes(x.section))) lines.push(`- Deduction: ${shortLabel(it.label)} ${money(it.total, true)}`);
    for (const x of d.dates) lines.push(`- Key date: ${new Date(x.date).toISOString().slice(0, 10)} ${x.label}`);
    const est = estimateTax(d.net, fy);
    lines.push(`Indicative tax on that net income if all taxed to him personally: ${money(est.total, true)} including Medicare (marginal rate ${pct(est.marginal)}). Trust distributions and offsets are not included.`);
  }
  const inbox = state.inbox.data;
  if (inbox && inbox.items) {
    const decided = inbox.items.filter((i) => i.status === 'decided');
    const open = inboxPending(inbox).length;
    lines.push('', `RECEIPT DECISIONS SO FAR (${decided.length} decided, ${open} waiting for his answer)`);
    for (const i of decided.slice(0, 40)) {
      const dec = i.decision || {};
      lines.push(`- ${i.date} ${i.vendor} ${i.amount != null ? money(i.amount, true) : ''}: ${TREATMENTS[dec.treatment] || dec.treatment}${dec.category ? ', ' + dec.category : ''}${dec.stream ? ' (' + dec.stream + ')' : ''}${dec.note ? '. His note: ' + dec.note : ''}`);
    }
  }
  const b = a.briefing;
  if (b && b.items && b.items.length) {
    lines.push('', `CURRENT TAX NEWS, CHECKED ${b.generated || ''} (use these for current-year rules; each item has its status)`);
    for (const it of b.items) lines.push(`- [${it.status}] ${it.topic}: ${it.headline}${it.effective ? ' (' + it.effective + ')' : ''}. ${it.what || ''}`);
  }
  return lines.join('\n');
}

// the API wants strictly alternating turns that start with the user
function apiMessages() {
  const out = [];
  for (const m of state.ask.thread.slice(-14)) {
    if (m.error || m.pending || !m.text) continue;
    const last = out[out.length - 1];
    if (last && last.role === m.role) last.content += '\n\n' + m.text; else out.push({ role: m.role, content: m.text });
  }
  while (out.length && out[0].role !== 'user') out.shift();
  return out;
}

async function postClaude(body, withFallback) {
  const headers = { 'content-type': 'application/json', 'x-api-key': cfg().apiKey, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' };
  let payload = body;
  if (withFallback) { headers['anthropic-beta'] = 'server-side-fallback-2026-07-01'; payload = { ...body, fallbacks: 'default' }; }
  const res = await fetchT(ASK_API, { method: 'POST', headers, body: JSON.stringify(payload) }, 120000);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const e = new Error((data.error && data.error.message) || `Claude returned ${res.status}`);
    e.status = res.status; throw e;
  }
  return data;
}
async function callClaude() {
  const body = { model: ASK_MODEL, max_tokens: 8000, system: buildSystemPrompt(), messages: apiMessages() };
  let data;
  try { data = await postClaude(body, true); }
  catch (e) { if (e.status === 400) data = await postClaude(body, false); else throw e; }   // if the fallback option is not accepted, ask plainly
  if (data.stop_reason === 'refusal') throw new Error("Claude wouldn't answer that one. Try rewording the question.");
  const text = (data.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('\n\n').trim();
  if (!text) throw new Error('Claude sent back an empty answer. Try again.');
  return text;
}
function askErrorText(e) {
  if (e instanceof TypeError) return 'No connection. Check your signal and try again.';
  if (e.status === 401) return 'Claude rejected the API key. Check it in Settings (the ⋯ button).';
  if (e.status === 403) return 'That API key is not allowed to use this model. Check the key in the Claude Console.';
  if (e.status === 429) return 'Claude is busy or the key has hit a limit. Try again in a minute.';
  if (e.status === 402 || /credit|balance|billing/i.test(e.message || '')) return 'The Claude account is out of credit. Add credit in the Claude Console.';
  return e.message || 'Something went wrong.';
}

// ----- no API key: leave the question in your OneDrive for Claude to answer next time it runs -----
async function updateQuestions(change) {
  for (let attempt = 0; attempt < 2; attempt++) {
    const r = await readInboxFile('questions.json');
    if (!r) throw new Error('The Inbox folder is missing in OneDrive.');
    const data = r.text ? JSON.parse(r.text) : { version: 1, items: [] };
    change(data);
    try {
      await graph(`/me/drive/items/${r.folderId}:/questions.json:/content`, { method: 'PUT', body: JSON.stringify(data, null, 2), headers: { 'Content-Type': 'application/json', ...(r.etag ? { 'If-Match': r.etag } : {}) } });
      return data;
    } catch (e) { if (e.status !== 412 || attempt) throw e; }
  }
}
async function queueQuestion(text) {
  const qid = 'q' + Date.now().toString(36);
  await updateQuestions((d) => d.items.push({ id: qid, at: isoNow(), q: text, status: 'open', answer: null }));
  return qid;
}
async function syncQueue() {
  const a = state.ask;
  if (!a.thread.some((m) => m.pending)) return;
  try {
    const r = await readInboxFile('questions.json');
    if (!r || !r.text) return;
    const items = JSON.parse(r.text).items || [];
    let changed = false;
    for (const m of a.thread) {
      if (!m.pending) continue;
      const it = items.find((x) => x.id === m.qid);
      if (it && it.answer) { m.pending = false; m.text = it.answer; m.at = Date.now(); changed = true; }
    }
    if (changed) { persistAsk(); if (state.tab === 'ask') renderAsk(); }
  } catch (e) { /* try again next time */ }
}

async function sendQuestion(text) {
  const a = state.ask;
  text = (text || '').trim();
  if (!text || a.busy) return;
  a.thread.push({ role: 'user', text, at: Date.now() });
  a.view = 'chat'; a.busy = true; persistAsk(); renderAsk();
  try {
    if (cfg().apiKey) a.thread.push({ role: 'assistant', text: await callClaude(), at: Date.now() });
    else a.thread.push({ role: 'assistant', pending: true, qid: await queueQuestion(text), text: '', at: Date.now() });
  } catch (e) {
    if (e instanceof AuthError) { a.busy = false; return needSignIn(); }
    a.thread.push({ role: 'assistant', error: true, text: askErrorText(e), at: Date.now() });
  }
  a.busy = false; persistAsk(); renderAsk();
  const box = $('ask'); if (box && box.lastElementChild) box.lastElementChild.scrollIntoView({ block: 'nearest' });
}
function askAbout(q) {
  state.ask.view = 'chat'; renderAsk();
  const input = $('ask-input'); input.value = q; input.focus();
}

// ----- display -----
function inlineFmt(s) {
  return s.split(/(\*\*[^*]+\*\*)/g).filter(Boolean).map((p) => (/^\*\*[^*]+\*\*$/.test(p) ? h('b', {}, p.slice(2, -2)) : p));
}
function formatAnswer(text) {
  const nodes = []; let list = null;
  for (const raw of text.split('\n')) {
    const line = raw.trim();
    if (!line) { list = null; continue; }
    const m = /^[-•*]\s+(.*)$/.exec(line);
    if (m) { if (!list) { list = h('ul', { class: 'ans-list' }); nodes.push(list); } list.append(h('li', {}, ...inlineFmt(m[1]))); }
    else { list = null; nodes.push(h('p', { class: 'ans-p' }, ...inlineFmt(line.replace(/^#+\s*/, '')))); }
  }
  return nodes;
}
const GENERIC_SUGGESTIONS = ['Can I claim my phone and internet?', 'Can I claim a meal with a client?', 'Can I claim a new fridge for a rental?'];

function renderAsk() {
  const box = $('ask'), a = state.ask;
  box.replaceChildren();
  const nBrief = a.briefing && a.briefing.items ? a.briefing.items.length : 0;
  box.append(h('div', { class: 'chips', style: 'margin:0' },
    h('button', { class: 'chip', type: 'button', 'aria-pressed': String(a.view === 'chat'), onclick: () => { a.view = 'chat'; renderAsk(); } }, 'Ask Claude'),
    h('button', { class: 'chip', type: 'button', 'aria-pressed': String(a.view === 'briefing'), onclick: () => { a.view = 'briefing'; renderAsk(); } }, `For you${nBrief ? ` (${nBrief})` : ''}`)));
  $('ask-form').hidden = a.view !== 'chat';
  if (a.view === 'briefing') { renderBriefing(box); return; }

  if (!a.thread.length) {
    const sugg = [...new Set([...(a.briefing && a.briefing.items ? a.briefing.items.flatMap((i) => (i.questions || []).slice(0, 1)) : []).slice(0, 3), ...GENERIC_SUGGESTIONS])].slice(0, 4);
    box.append(card(
      h('b', {}, 'Ask whether you can claim something'),
      h('p', { class: 'muted', style: 'margin:6px 0 10px' }, cfg().apiKey
        ? 'Claude knows your setup from your profile, your workbook and your receipt decisions. Answers are general information, not tax advice. Anything material goes to your accountant.'
        : 'No API key yet, so questions are saved for Claude to answer the next time you open it on your Mac. Add an API key in Settings for instant answers. Your questions and tax figures are sent to Claude to answer.'),
      h('div', { class: 'qchips' }, ...sugg.map((q) => h('button', { class: 'qchip', type: 'button', onclick: () => askAbout(q) }, q)))));
  }
  for (const m of a.thread) {
    if (m.role === 'user') box.append(h('div', { class: 'bubble me' }, m.text));
    else if (m.pending) box.append(h('div', { class: 'bubble ai wait' }, 'Saved for Claude. The answer appears here once Claude has looked at it.', h('button', { class: 'chip', type: 'button', style: 'margin-left:8px', onclick: syncQueue }, 'Check now')));
    else if (m.error) box.append(h('div', { class: 'bubble ai err' }, m.text));
    else box.append(h('div', { class: 'bubble ai' }, h('div', { class: 'who' }, 'Claude'), ...formatAnswer(m.text)));
  }
  if (a.busy) box.append(h('div', { class: 'bubble ai wait' }, h('div', { class: 'spinner', style: 'margin:0 8px 0 0;display:inline-block;vertical-align:middle' }), 'Thinking…'));
  $('ask-send').disabled = a.busy;
}

function renderBriefing(box) {
  const b = state.ask.briefing;
  if (!b || !b.items || !b.items.length) {
    box.append(card(h('b', {}, 'No briefing yet'), h('p', { class: 'muted', style: 'margin-top:6px' }, 'Claude writes this from the latest Budget, ATO and Queensland announcements, matched to your setup. Ask Claude to "update my tax briefing".')));
    return;
  }
  box.append(h('p', { class: 'note', style: 'padding:0 4px' }, `Checked ${fmtYmd(b.generated)}. ${b.note || ''} Tap a question to ask Claude about it.`));
  for (const it of b.items) {
    const firm = /in effect|legislated/i.test(it.status || '');
    box.append(card(
      h('div', { class: 'stream-head' }, h('span', { class: 'eyebrow2' }, it.topic), h('span', { class: 'pill' + (firm ? '' : ' warn') }, it.status || '')),
      h('b', { style: 'display:block;margin:4px 0 8px' }, it.headline),
      it.you ? h('div', { class: 'guess', style: 'margin-top:0' }, h('div', { class: 'eyebrow2' }, 'How it could affect you'), h('p', { style: 'margin:0;font-size:15px' }, it.you)) : null,
      it.questions && it.questions.length ? h('div', { class: 'qchips' }, ...it.questions.map((q) => h('button', { class: 'qchip', type: 'button', onclick: () => askAbout(q) }, q))) : null,
      h('details', { class: 'det' }, h('summary', {}, 'What changed and the source'),
        it.effective ? h('p', { class: 'note', style: 'margin:8px 0 0' }, it.effective) : null,
        h('p', { class: 'muted', style: 'font-size:14px;margin:6px 0' }, it.what || ''),
        it.source && it.source.url ? h('a', { class: 'chip', href: it.source.url, target: '_blank', rel: 'noopener' }, `${it.source.title || 'Source'} ↗`) : null)));
  }
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

/* ---- driving (mileage) ------------------------------------------------ */
/* Hues validated against both surfaces for CVD separation and contrast; keep the order fixed. */
const DRIVE_HUES = {
  light: { work: '#1b5fd1', rental: '#e8590c', gym: '#8b5cf6', personal: '#0f9b8e' },
  dark:  { work: '#4d8df0', rental: '#d16b16', gym: '#9470e8', personal: '#1fa894' },
};
const driveHue = (key) =>
  (matchMedia('(prefers-color-scheme: dark)').matches ? DRIVE_HUES.dark : DRIVE_HUES.light)[key] || 'var(--muted)';
const km = (n) => `${n.toLocaleString('en-AU', { minimumFractionDigits: 0, maximumFractionDigits: 0 })} km`;

async function loadVehicle() {
  if (state.vehicle !== undefined) return;
  state.vehicle = null;                        // claim the slot so we only fetch once
  try {
    const f = await readInboxFile('vehicle.json');
    state.vehicle = f && f.text ? JSON.parse(f.text) : null;
  } catch (e) { if (e instanceof AuthError) return needSignIn(); state.vehicle = null; }
  if (state.tab === 'home') renderHome();
}

/* SVG donut. r chosen so the circumference is a round 100 => dasharray is literally a percentage. */
function donut(slices, centre, sub) {
  const R = 15.915, C = 2 * Math.PI * R;
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 42 42');
  svg.setAttribute('class', 'donut');
  svg.setAttribute('role', 'img');
  svg.setAttribute('aria-label', slices.map((s) => `${s.label} ${s.pct.toFixed(0)}%`).join(', '));
  const el = (tag, attrs) => {
    const n = document.createElementNS('http://www.w3.org/2000/svg', tag);
    for (const [k, v] of Object.entries(attrs)) n.setAttribute(k, v);
    return n;
  };
  svg.append(el('circle', { class: 'track', cx: 21, cy: 21, r: R }));
  let at = 0;
  for (const s of slices) {
    if (s.pct <= 0) continue;
    const len = s.pct / 100 * C;
    svg.append(el('circle', {
      class: 'seg', cx: 21, cy: 21, r: R, stroke: s.colour,
      /* 1.2 of surface between segments so neighbours never touch */
      'stroke-dasharray': `${Math.max(0, len - 0.9)} ${C - Math.max(0, len - 0.9)}`,
      'stroke-dashoffset': -at, 'stroke-linecap': 'butt',
    }));
    at += len;
  }
  const t = el('text', { x: 21, y: 20.2, 'text-anchor': 'middle', class: 'donut-mid' });
  t.textContent = centre;
  const s2 = el('text', { x: 21, y: 24.6, 'text-anchor': 'middle', class: 'donut-sub' });
  s2.textContent = sub;
  svg.append(t, s2);
  return svg;
}

function renderDriving(box) {
  const v = state.vehicle;
  if (!v || !v.periods || !v.periods.length) return;
  const keys = v.buckets.map((b) => b.key);
  const ytd = {};
  for (const k of keys) ytd[k] = v.periods.reduce((a, p) => a + (p.km[k] || 0), 0);
  const ytdTotal = keys.reduce((a, k) => a + ytd[k], 0);
  if (!ytdTotal) return;
  const trips = v.periods.reduce((a, p) => a + (p.trips || 0), 0);

  const meta = Object.fromEntries(v.buckets.map((b) => [b.key, b]));
  const claimKeys = keys.filter((k) => meta[k].claimable);
  const claimKm = claimKeys.reduce((a, k) => a + ytd[k], 0);
  const claim = claimKm * v.ratePerKm;

  box.append(sect('Driving'));
  const c = card();

  /* headline: what the year is worth so far */
  c.append(h('div', { class: 'kv total' },
    h('span', {}, `FY${v.fy % 100} claimable travel`), h('b', {}, money(claim, true))));
  c.append(h('div', { class: 'kv' },
    h('span', {}, `${km(claimKm)} at ${Math.round(v.ratePerKm * 100)}c`),
    h('b', {}, `${km(ytdTotal)} driven · ${trips} trips`)));

  /* YTD composition */
  const slices = v.buckets.map((b) => ({ ...b, value: ytd[b.key], pct: ytd[b.key] / ytdTotal * 100, colour: driveHue(b.key) }))
    .filter((s) => s.value > 0);
  c.append(h('div', { class: 'drive-wrap', style: 'margin-top:14px' },
    donut(slices, km(claimKm).replace(' km', ''), 'CLAIMED'),
    h('div', { class: 'drive-legend' }, ...slices.map((s) => h('div', { class: 'row' },
      h('i', { class: 'dot', style: `background:${s.colour}` }),
      h('span', { class: 'nm' }, s.label),
      h('span', { class: 'vl' }, km(s.value)),
      h('span', { class: 'pc' }, `${s.pct.toFixed(0)}%`))))));

  /* month on month: one stacked bar per month, shared scale so heights compare */
  const peak = Math.max(...v.periods.map((p) => keys.reduce((a, k) => a + (p.km[k] || 0), 0)));
  const months = h('div', { class: 'mom' });
  for (const p of v.periods) {
    const tot = keys.reduce((a, k) => a + (p.km[k] || 0), 0);
    const col = h('div', { class: 'mom-col' });
    const stack = h('div', { class: 'mom-stack', style: `height:${Math.max(4, tot / peak * 92).toFixed(1)}px` });
    for (const k of keys) {
      const val = p.km[k] || 0;
      if (!val) continue;
      stack.append(h('i', { style: `flex:${val};background:${driveHue(k)}`, title: `${meta[k].label} ${km(val)}` }));
    }
    col.append(h('b', {}, km(p.km[claimKeys[0]] || 0).replace(' km', '')), stack,
      h('small', {}, p.label.slice(0, 3) + (p.partial ? '*' : '')));
    months.append(col);
  }
  c.append(h('div', { class: 'sect', style: 'margin:16px 0 6px' }, 'Month on month'), months);
  c.append(h('p', { class: 'note', style: 'margin-top:10px' }, 'Figure above each bar is the claimable km for that month. Bar height is total distance driven.'));

  /* the cap */
  const capPct = Math.min(100, claimKm / v.capKm * 100);
  c.append(h('div', { class: 'kv', style: 'margin-top:12px' },
    h('span', {}, `Against the ${v.capKm.toLocaleString('en-AU')} km cents-per-km cap`),
    h('b', {}, `${capPct.toFixed(0)}%`)));
  c.append(h('div', { class: 'progress', role: 'img', 'aria-label': `${capPct.toFixed(0)} per cent of the cap` },
    h('i', { style: `width:${capPct.toFixed(1)}%` })));

  const notClaim = slices.filter((s) => !s.claimable && s.why);
  if (notClaim.length) {
    const why = h('div', { class: 'drive-why' });
    for (const s of notClaim) why.append(h('p', { style: 'margin:0 0 6px' }, h('b', {}, s.label + ': '), s.why));
    c.append(why);
  }
  const part = v.periods.filter((p) => p.partial);
  if (part.length) c.append(h('p', { class: 'note', style: 'margin-top:8px' },
    h('span', { class: 'pill warn' }, 'Part month'), ' ', part.map((p) => p.partialNote).filter(Boolean).join(' ')));
  box.append(c);
}
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
  /* with assumptions saved, the refund-or-bill estimate leads and the net-income hero drops to second */
  const est = currentEstimate(d, fy);
  const hasEst = !!(est && est.hasAssumptions);
  if (hasEst) box.append(estimateHero(est, fy));
  box.append(h('div', { class: 'hero' },
    h('div', { class: 'eyebrow' }, `FY${fy % 100} net income recorded`),
    h('div', { class: 'big' }, money(d.net)),
    h('div', { class: 'sub' }, `${money(d.income, true)} income less ${money(d.deductions, true)} deductions`),
    h('div', { class: 'progress', role: 'img', 'aria-label': `Day ${dayNo} of ${total}` }, h('i', { style: `width:${(dayNo / total * 100).toFixed(1)}%` })),
    h('div', { class: 'progress-label' }, h('span', {}, dayNo ? `Day ${dayNo} of ${total}` : `Starts ${relDays(start)}`), h('span', {}, dayNo >= total ? 'Year complete' : `${total - dayNo} days left`))));
  /* OneDrive could not be read: the phone's copy of the assumptions is in use */
  if (state.assume.error) box.append(h('p', { class: 'note', style: 'text-align:center;margin:-4px 0 10px' }, ASSUME_OFFLINE_NOTE));

  box.append(h('div', { class: 'tiles' },
    h('div', { class: 'tile in' }, h('small', {}, 'Income'), h('b', {}, money(d.income))),
    h('div', { class: 'tile' }, h('small', {}, 'Deductions'), h('b', {}, money(d.deductions)))));

  renderInboxPrompt(box);
  renderTasks(box);
  renderDriving(box);
  loadVehicle();

  if (hasEst) {
    renderChainCard(box, est, state.rates, fy);
    renderCgtCard(box, est, state.rates);
  } else {
    renderAssumptionsPrompt(box);
    /* indicative tax (round 1): shown only until assumptions exist */
    const est1 = estimateTax(d.net, fy);
    const taxRows = [
      kv('Income tax', money(est1.tax, true)),
      kv('Medicare levy (2%)', money(est1.medicare, true)),
      kv('Estimated total', money(est1.total, true), 'total'),
      kv('Effective rate', pct(est1.effective)),
      kv('Rate on your next dollar', pct(est1.marginal)),
    ];
    if (est1.next) taxRows.push(kv(`Room before the ${pct(est1.next.rate + 0.02)} rate`, money(est1.next.room)));
    box.append(sect('Indicative tax'),
      card(...taxRows,
        h('p', { class: 'note', style: 'margin-top:10px' }, h('span', { class: 'pill warn' }, 'Estimate'), ' ',
          `Assumes resident rates for ${fy - 1}-${String(fy).slice(2)}, all net income taxed to you personally, and no offsets, PAYG credits or capital gains. Trust distributions and the share of income taxed to you can change this a lot, so use it as a guide and check with your accountant.`)));
  }

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
    p.webUrl ? h('a', { class: 'btn', href: p.webUrl, target: '_blank', rel: 'noopener' }, 'Open workbook') : null,
    h('button', { class: 'btn', type: 'button', id: 'btn-assume', disabled: !state.rates, onclick: openAssumptions }, 'Assumptions')));
  box.append(h('p', { class: 'note', style: 'text-align:center;margin-top:14px' },
    state.positionBusy ? 'Updating from OneDrive…' : state.positionError ? `Couldn't update (${state.positionError}). Showing figures from ${ago(p.at)}.` : `Figures from ${p.name}, read ${ago(p.at)}.`,
    ' ', h('button', { class: 'btn quiet', type: 'button', style: 'min-height:44px;padding:0 8px', onclick: () => refreshPosition(true) }, 'Refresh')));
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
  $('s-apikey').value = c.apiKey; $('s-version').textContent = APP_VERSION;
  $('s-acc-name').value = c.accountantName; $('s-acc-email').value = c.accountantEmail;
  renderLockRow();
  $('settings').showModal();
}
function signOut() { for (const k of ['tokens', 'pkce', 'lock', 'snapshot', 'apiKey', 'askThread', 'assumptions']) ls.del(k); location.replace(redirectUri()); }

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
  $('tab-inbox').addEventListener('click', () => { show('inbox'); renderInbox(); loadInbox(); });
  $('tab-ask').addEventListener('click', () => { show('ask'); renderAsk(); loadAskContext(); });
  $('ask-form').addEventListener('submit', (e) => { e.preventDefault(); const i = $('ask-input'); const t = i.value; i.value = ''; i.style.height = ''; sendQuestion(t); });
  $('ask-input').addEventListener('input', (e) => { e.target.style.height = 'auto'; e.target.style.height = Math.min(140, e.target.scrollHeight) + 'px'; });
  $('tab-files').addEventListener('click', () => { show('files'); if (state.dirty && current()) loadFolder(); });
  $('tab-add').addEventListener('click', () => { show('add'); fillFolderMenu(); fillAddOptions(); refreshAddCard(); updateNamePreview(); });

  $('in-camera').addEventListener('change', (e) => onPicked(e.target.files[0]));
  $('in-file').addEventListener('change', (e) => onPicked(e.target.files[0]));
  $('btn-clear').addEventListener('click', clearPending);
  $('btn-paste').addEventListener('click', pasteFromClipboard);
  document.addEventListener('paste', onPasteEvent);
  $('btn-scan').addEventListener('click', async () => {
    try {
      if (typeof Scanner === 'undefined') {
        await new Promise((resolve, reject) => {
          const el = document.createElement('script'); el.src = 'scan.js?v=' + APP_VERSION;
          el.onload = resolve; el.onerror = () => reject(new Error('scan.js did not load'));
          document.head.append(el);
        });
      }
      Scanner.wire(); Scanner.open(onScanned);
    } catch (e) { banner(`The scanner couldn't start (${e.message}). Close the app completely and open it again.`); }
  });
  for (const id of ['f-date', 'f-vendor', 'f-amount']) $(id).addEventListener('input', updateNamePreview);
  $('f-amount').addEventListener('blur', () => { const a = parseAmount($('f-amount').value); if (a) $('f-amount').value = a; updateNamePreview(); });
  $('btn-save').addEventListener('click', saveReceipt);
  $('f-date').value = today();
  $('f-date').addEventListener('input', () => { $('f-date').dataset.touched = '1'; });

  $('settings').addEventListener('close', () => {
    if ($('settings').returnValue !== 'save') return;
    const c = cfg();
    const next = { rootPath: $('s-root').value.trim().replace(/^\/+|\/+$/g, '') || DEFAULTS.rootPath, receiptsFolder: $('s-receipts').value.trim() || DEFAULTS.receiptsFolder, workbook: $('s-workbook').value.trim() || DEFAULTS.workbook, clientId: $('s-client').value.trim() };
    ls.set('rootPath', next.rootPath); ls.set('receiptsFolder', next.receiptsFolder); ls.set('workbook', next.workbook);
    if (next.workbook !== c.workbook) ls.del('snapshot');
    ls.set('apiKey', $('s-apikey').value.trim());
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
  wireAssumptions();
}

async function boot() {
  wire();
  if (lockOn()) { showLock('Locked'); unlock(true); }
  if ('serviceWorker' in navigator) {
    const hadController = !!navigator.serviceWorker.controller;
    let reloaded = false;
    navigator.serviceWorker.register('sw.js').catch(() => {});
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (!hadController || reloaded) return;
      if (state.pending || !$('scan').hidden) { toast('Updated. Reopen the app soon to use the latest version.'); return; }
      reloaded = true; location.reload();
    });
  }
  if (!cfg().clientId) return show('setup');
  try { await handleRedirect(); } catch (e) { banner(friendly(e)); }
  if (!ls.get('tokens')) return show('signin');
  await openRoot();
}
boot();
