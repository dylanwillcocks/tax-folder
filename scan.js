'use strict';

/* Document scanner: photo -> find the page -> adjust corners -> flatten perspective -> clean up -> pages -> PDF.
   Plain canvas maths and a tiny PDF writer: no libraries, nothing leaves the phone until you save. */

const Scanner = (() => {
  const WORK_MAX = 2000;   // longest side of the working copy of the photo
  const PAGE_MAX = 1700;   // longest side of a finished page
  const $ = (id) => document.getElementById(id);

  let work = null;         // canvas holding the current photo
  let corners = [];        // [{x, y}] top-left, top-right, bottom-right, bottom-left, in work coordinates
  let base = null;         // flattened page before the filter, as a canvas
  let filter = 'clean';
  let fmt = 'pdf';
  let pages = [];          // [{blob, url, w, h}]
  let step = 'capture';
  let done = null;
  let view = { scale: 1, w: 0, h: 0 };

  /* ---------- geometry ---------- */
  const dist = (a, b) => Math.hypot(a.x - b.x, a.y - b.y);

  // Solve for the 3x3 homography that maps output pixel (u, v) to photo pixel (x, y).
  function homography(dst, src) {
    const A = [], b = [];
    for (let i = 0; i < 4; i++) {
      const { x: u, y: v } = dst[i], { x, y } = src[i];
      A.push([u, v, 1, 0, 0, 0, -u * x, -v * x]); b.push(x);
      A.push([0, 0, 0, u, v, 1, -u * y, -v * y]); b.push(y);
    }
    const n = 8;
    for (let c = 0; c < n; c++) {                    // Gaussian elimination with partial pivoting
      let p = c;
      for (let r = c + 1; r < n; r++) if (Math.abs(A[r][c]) > Math.abs(A[p][c])) p = r;
      [A[c], A[p]] = [A[p], A[c]]; [b[c], b[p]] = [b[p], b[c]];
      const d = A[c][c] || 1e-12;
      for (let r = c + 1; r < n; r++) {
        const f = A[r][c] / d;
        for (let k = c; k < n; k++) A[r][k] -= f * A[c][k];
        b[r] -= f * b[c];
      }
    }
    const h = new Array(n);
    for (let r = n - 1; r >= 0; r--) {
      let s = b[r];
      for (let k = r + 1; k < n; k++) s -= A[r][k] * h[k];
      h[r] = s / (A[r][r] || 1e-12);
    }
    return h;
  }

  function warp(srcCanvas, quad) {
    const wTop = dist(quad[0], quad[1]), wBot = dist(quad[3], quad[2]);
    const hL = dist(quad[0], quad[3]), hR = dist(quad[1], quad[2]);
    let W = Math.max(wTop, wBot), H = Math.max(hL, hR);
    const k = Math.min(1, PAGE_MAX / Math.max(W, H));
    W = Math.max(8, Math.round(W * k)); H = Math.max(8, Math.round(H * k));
    const h = homography([{ x: 0, y: 0 }, { x: W, y: 0 }, { x: W, y: H }, { x: 0, y: H }], quad);
    const sw = srcCanvas.width, sh = srcCanvas.height;
    const src = srcCanvas.getContext('2d').getImageData(0, 0, sw, sh).data;
    const out = document.createElement('canvas'); out.width = W; out.height = H;
    const octx = out.getContext('2d');
    const img = octx.createImageData(W, H), dst = img.data;
    for (let v = 0; v < H; v++) {
      for (let u = 0; u < W; u++) {
        const q = h[6] * u + h[7] * v + 1;
        let x = (h[0] * u + h[1] * v + h[2]) / q, y = (h[3] * u + h[4] * v + h[5]) / q;
        x = x < 0 ? 0 : x > sw - 1.001 ? sw - 1.001 : x; y = y < 0 ? 0 : y > sh - 1.001 ? sh - 1.001 : y;
        const x0 = x | 0, y0 = y | 0, fx = x - x0, fy = y - y0;
        const i00 = (y0 * sw + x0) * 4, i10 = i00 + 4, i01 = i00 + sw * 4, i11 = i01 + 4;
        const o = (v * W + u) * 4;
        for (let c = 0; c < 3; c++) {
          dst[o + c] = (src[i00 + c] * (1 - fx) + src[i10 + c] * fx) * (1 - fy) + (src[i01 + c] * (1 - fx) + src[i11 + c] * fx) * fy;
        }
        dst[o + 3] = 255;
      }
    }
    octx.putImageData(img, 0, 0);
    return out;
  }

  /* ---------- finding the page ---------- */
  // Paper is usually the largest bright region. Threshold the photo (Otsu), keep the biggest bright blob,
  // and take its four extreme points. If that looks wrong the user simply drags the corners.
  function insetCorners(w, h, m = 0.06) {
    return [{ x: w * m, y: h * m }, { x: w * (1 - m), y: h * m }, { x: w * (1 - m), y: h * (1 - m) }, { x: w * m, y: h * (1 - m) }];
  }
  function detectCorners(src) {
    const S = 260, sc = Math.min(1, S / Math.max(src.width, src.height));
    const w = Math.max(8, Math.round(src.width * sc)), h = Math.max(8, Math.round(src.height * sc));
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    const g = c.getContext('2d'); g.drawImage(src, 0, 0, w, h);
    const d = g.getImageData(0, 0, w, h).data;
    let gray = new Float32Array(w * h);
    for (let i = 0; i < w * h; i++) gray[i] = 0.299 * d[i * 4] + 0.587 * d[i * 4 + 1] + 0.114 * d[i * 4 + 2];
    const blurred = new Float32Array(w * h);          // 3x3 box blur, which calms text and noise
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      let s = 0, n = 0;
      for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
        const xx = x + dx, yy = y + dy;
        if (xx >= 0 && yy >= 0 && xx < w && yy < h) { s += gray[yy * w + xx]; n++; }
      }
      blurred[y * w + x] = s / n;
    }
    gray = blurred;
    const hist = new Array(256).fill(0);
    for (let i = 0; i < w * h; i++) hist[Math.min(255, gray[i] | 0)]++;
    let sumAll = 0; for (let i = 0; i < 256; i++) sumAll += i * hist[i];
    let wB = 0, sumB = 0, best = 0, thr = 128;
    for (let t = 0; t < 256; t++) {
      wB += hist[t]; if (!wB) continue;
      const wF = w * h - wB; if (!wF) break;
      sumB += t * hist[t];
      const mB = sumB / wB, mF = (sumAll - sumB) / wF, between = wB * wF * (mB - mF) * (mB - mF);
      if (between > best) { best = between; thr = t; }
    }
    const mask = new Uint8Array(w * h);
    for (let i = 0; i < w * h; i++) mask[i] = gray[i] > thr ? 1 : 0;
    const seen = new Uint8Array(w * h);
    let bestComp = null;
    const stack = [];
    for (let i0 = 0; i0 < w * h; i0++) {
      if (!mask[i0] || seen[i0]) continue;
      const pts = [];
      stack.push(i0); seen[i0] = 1;
      while (stack.length) {
        const i = stack.pop(); pts.push(i);
        const x = i % w, y = (i / w) | 0;
        if (x > 0 && mask[i - 1] && !seen[i - 1]) { seen[i - 1] = 1; stack.push(i - 1); }
        if (x < w - 1 && mask[i + 1] && !seen[i + 1]) { seen[i + 1] = 1; stack.push(i + 1); }
        if (y > 0 && mask[i - w] && !seen[i - w]) { seen[i - w] = 1; stack.push(i - w); }
        if (y < h - 1 && mask[i + w] && !seen[i + w]) { seen[i + w] = 1; stack.push(i + w); }
      }
      if (!bestComp || pts.length > bestComp.length) bestComp = pts;
    }
    const fallback = insetCorners(src.width, src.height);
    if (!bestComp) return { corners: fallback, found: false };
    const frac = bestComp.length / (w * h);
    if (frac < 0.12 || frac > 0.97) return { corners: fallback, found: false };
    let tl = null, tr = null, br = null, bl = null;
    let mTL = Infinity, mTR = -Infinity, mBR = -Infinity, mBL = Infinity;
    for (const i of bestComp) {
      const x = i % w, y = (i / w) | 0, s = x + y, df = x - y;
      if (s < mTL) { mTL = s; tl = { x, y }; }
      if (s > mBR) { mBR = s; br = { x, y }; }
      if (df > mTR) { mTR = df; tr = { x, y }; }
      if (df < mBL) { mBL = df; bl = { x, y }; }
    }
    const quad = [tl, tr, br, bl];
    const area = Math.abs(quad.reduce((a, p, i) => { const q = quad[(i + 1) % 4]; return a + (p.x * q.y - q.x * p.y); }, 0)) / 2;
    if (area < 0.5 * bestComp.length) return { corners: fallback, found: false };
    return { corners: quad.map((p) => ({ x: (p.x + 0.5) / sc, y: (p.y + 0.5) / sc })), found: true };
  }

  /* ---------- clean-up filters ---------- */
  // Lighting estimate: for each 16 px block take the average of its brighter half (that is the paper, not the ink),
  // then blend between block centres so the result is smooth.
  function paperLevel(data, w, h) {
    const B = 16, bw = Math.ceil(w / B), bh = Math.ceil(h / B);
    const lum = new Float32Array(w * h);
    for (let i = 0; i < w * h; i++) lum[i] = 0.299 * data[i * 4] + 0.587 * data[i * 4 + 1] + 0.114 * data[i * 4 + 2];
    const level = new Float32Array(bw * bh);
    for (let by = 0; by < bh; by++) for (let bx = 0; bx < bw; bx++) {
      let s = 0, n = 0;
      for (let y = by * B; y < Math.min(h, (by + 1) * B); y++) for (let x = bx * B; x < Math.min(w, (bx + 1) * B); x++) { s += lum[y * w + x]; n++; }
      const mean = s / n; let s2 = 0, n2 = 0;
      for (let y = by * B; y < Math.min(h, (by + 1) * B); y++) for (let x = bx * B; x < Math.min(w, (bx + 1) * B); x++) { const v = lum[y * w + x]; if (v >= mean) { s2 += v; n2++; } }
      level[by * bw + bx] = n2 ? s2 / n2 : mean;
    }
    return { lum, at(x, y) {
      const fx = Math.min(bw - 1, Math.max(0, x / B - 0.5)), fy = Math.min(bh - 1, Math.max(0, y / B - 0.5));
      const x0 = fx | 0, y0 = fy | 0, x1 = Math.min(bw - 1, x0 + 1), y1 = Math.min(bh - 1, y0 + 1), ax = fx - x0, ay = fy - y0;
      return (level[y0 * bw + x0] * (1 - ax) + level[y0 * bw + x1] * ax) * (1 - ay) + (level[y1 * bw + x0] * (1 - ax) + level[y1 * bw + x1] * ax) * ay;
    } };
  }
  function applyFilter(srcCanvas, kind) {
    const w = srcCanvas.width, h = srcCanvas.height;
    const out = document.createElement('canvas'); out.width = w; out.height = h;
    const octx = out.getContext('2d');
    if (kind === 'original') { octx.drawImage(srcCanvas, 0, 0); return out; }
    const img = srcCanvas.getContext('2d').getImageData(0, 0, w, h), d = img.data;
    const paper = paperLevel(d, w, h);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      const i = (y * w + x) * 4, k = 255 / Math.max(60, paper.at(x, y));
      let r = Math.min(255, d[i] * k), g = Math.min(255, d[i + 1] * k), b = Math.min(255, d[i + 2] * k);
      if (kind === 'bw') {
        const L = 0.299 * r + 0.587 * g + 0.114 * b;
        const t = Math.min(1, Math.max(0, (L - 120) / 70));          // ink below ~120 goes black, paper above ~190 goes white
        r = g = b = 255 * t * t * (3 - 2 * t);
      } else {
        r = 255 * Math.pow(r / 255, 1.35); g = 255 * Math.pow(g / 255, 1.35); b = 255 * Math.pow(b / 255, 1.35);   // darker ink, white paper
      }
      d[i] = r; d[i + 1] = g; d[i + 2] = b;
    }
    octx.putImageData(img, 0, 0);
    return out;
  }

  /* ---------- PDF writer: one JPEG per page ---------- */
  async function buildPdf(list) {
    const enc = new TextEncoder();
    const chunks = [], offsets = [];
    let pos = 0;
    const push = (x) => { const b = typeof x === 'string' ? enc.encode(x) : x; chunks.push(b); pos += b.length; };
    const total = 2 + list.length * 3;
    push('%PDF-1.4\n');
    offsets[1] = pos; push('1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n');
    offsets[2] = pos;
    push(`2 0 obj\n<< /Type /Pages /Kids [${list.map((_, i) => `${3 + i * 3} 0 R`).join(' ')}] /Count ${list.length} >>\nendobj\n`);
    for (let i = 0; i < list.length; i++) {
      const pg = list[i], pw = 595.28, ph = +(pw * pg.h / pg.w).toFixed(2);
      const pageObj = 3 + i * 3, contObj = 4 + i * 3, imgObj = 5 + i * 3;
      const bytes = new Uint8Array(await pg.blob.arrayBuffer());
      const content = `q ${pw} 0 0 ${ph} 0 0 cm /Im0 Do Q`;
      offsets[pageObj] = pos;
      push(`${pageObj} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pw} ${ph}] /Resources << /XObject << /Im0 ${imgObj} 0 R >> >> /Contents ${contObj} 0 R >>\nendobj\n`);
      offsets[contObj] = pos;
      push(`${contObj} 0 obj\n<< /Length ${content.length} >>\nstream\n${content}\nendstream\nendobj\n`);
      offsets[imgObj] = pos;
      push(`${imgObj} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${pg.w} /Height ${pg.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${bytes.length} >>\nstream\n`);
      push(bytes); push('\nendstream\nendobj\n');
    }
    const xref = pos;
    let table = `xref\n0 ${total + 1}\n0000000000 65535 f \n`;
    for (let n = 1; n <= total; n++) table += `${String(offsets[n]).padStart(10, '0')} 00000 n \n`;
    push(table);
    push(`trailer\n<< /Size ${total + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
    return new Blob(chunks, { type: 'application/pdf' });
  }

  /* ---------- loading a photo ---------- */
  async function loadPhoto(file) {
    let src, w, h;
    try { src = await createImageBitmap(file, { imageOrientation: 'from-image' }); w = src.width; h = src.height; }
    catch {
      src = await new Promise((resolve, reject) => {
        const img = new Image(), url = URL.createObjectURL(file);
        img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
        img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("Couldn't read that photo.")); };
        img.src = url;
      });
      w = src.naturalWidth; h = src.naturalHeight;
    }
    const k = Math.min(1, WORK_MAX / Math.max(w, h));
    const c = document.createElement('canvas'); c.width = Math.round(w * k); c.height = Math.round(h * k);
    c.getContext('2d').drawImage(src, 0, 0, c.width, c.height);
    return c;
  }
  function rotated(c) {
    const o = document.createElement('canvas'); o.width = c.height; o.height = c.width;
    const g = o.getContext('2d'); g.translate(o.width, 0); g.rotate(Math.PI / 2); g.drawImage(c, 0, 0);
    return o;
  }

  /* ---------- screens ---------- */
  const PANES = ['capture', 'crop', 'enhance', 'review'];
  function show(name) {
    step = name;
    for (const p of PANES) $('scan-' + p).hidden = p !== name;
    $('scan-title').textContent = { capture: 'Scan', crop: 'Adjust the corners', enhance: 'Choose a look', review: `${pages.length} page${pages.length === 1 ? '' : 's'}` }[name];
    const primary = $('scan-primary');
    primary.hidden = name === 'capture' && !pages.length;
    primary.textContent = { capture: 'Done', crop: 'Next', enhance: 'Add page', review: 'Save' }[name];
    $('scan-back').hidden = !(name === 'crop' || name === 'enhance');
    $('scan-cancel').hidden = name === 'crop' || name === 'enhance';
    $('scan-msg').hidden = true;
  }
  function msg(text, kind = '') { const m = $('scan-msg'); m.textContent = text || ''; m.className = 'banner' + (kind ? ' ' + kind : ''); m.hidden = !text; }
  function busy(on, text) { const b = $('scan-busy'); b.hidden = !on; b.querySelector('span').textContent = text || 'Working…'; }

  function layoutCrop() {
    const stage = $('crop-stage'), avail = $('scan-crop').clientWidth || 340;
    const maxH = Math.max(240, (window.innerHeight || 700) - 260);
    const scale = Math.min(avail / work.width, maxH / work.height);
    view = { scale, w: Math.round(work.width * scale), h: Math.round(work.height * scale) };
    stage.style.width = view.w + 'px'; stage.style.height = view.h + 'px';
    for (const id of ['crop-canvas', 'crop-overlay']) { const c = $(id); c.width = view.w; c.height = view.h; c.style.width = view.w + 'px'; c.style.height = view.h + 'px'; }
    $('crop-canvas').getContext('2d').drawImage(work, 0, 0, view.w, view.h);
    drawOverlay();
  }
  function drawOverlay() {
    const c = $('crop-overlay'), g = c.getContext('2d');
    g.clearRect(0, 0, c.width, c.height);
    const pts = corners.map((p) => ({ x: p.x * view.scale, y: p.y * view.scale }));
    g.fillStyle = 'rgba(5, 20, 60, 0.45)';
    g.beginPath(); g.rect(0, 0, c.width, c.height);
    g.moveTo(pts[0].x, pts[0].y); for (let i = 3; i >= 0; i--) g.lineTo(pts[i].x, pts[i].y); g.closePath();
    g.fill('evenodd');
    g.strokeStyle = '#ff8a1f'; g.lineWidth = 2.5; g.beginPath();
    pts.forEach((p, i) => (i ? g.lineTo(p.x, p.y) : g.moveTo(p.x, p.y))); g.closePath(); g.stroke();
    document.querySelectorAll('.crop-handle').forEach((hd, i) => { hd.style.left = pts[i].x + 'px'; hd.style.top = pts[i].y + 'px'; });
  }
  function showLoupe(i) {
    const l = $('crop-loupe'), g = l.getContext('2d'), p = corners[i], r = 36;
    l.hidden = false;
    l.style.left = (pts0(i).x > view.w / 2 ? 8 : view.w - l.width - 8) + 'px'; l.style.top = '8px';
    g.clearRect(0, 0, l.width, l.height);
    g.drawImage(work, p.x - r, p.y - r, r * 2, r * 2, 0, 0, l.width, l.height);
    g.strokeStyle = '#ff8a1f'; g.lineWidth = 1.5; g.beginPath();
    g.moveTo(l.width / 2, 0); g.lineTo(l.width / 2, l.height); g.moveTo(0, l.height / 2); g.lineTo(l.width, l.height / 2); g.stroke();
  }
  const pts0 = (i) => ({ x: corners[i].x * view.scale, y: corners[i].y * view.scale });

  function wireHandles() {
    document.querySelectorAll('.crop-handle').forEach((hd) => {
      const i = Number(hd.dataset.i);
      hd.addEventListener('pointerdown', (e) => { hd.setPointerCapture(e.pointerId); hd.dataset.drag = '1'; showLoupe(i); e.preventDefault(); });
      hd.addEventListener('pointermove', (e) => {
        if (!hd.dataset.drag) return;
        const r = $('crop-stage').getBoundingClientRect();
        corners[i] = { x: Math.min(work.width, Math.max(0, (e.clientX - r.left) / view.scale)), y: Math.min(work.height, Math.max(0, (e.clientY - r.top) / view.scale)) };
        drawOverlay(); showLoupe(i);
      });
      const end = () => { delete hd.dataset.drag; $('crop-loupe').hidden = true; };
      hd.addEventListener('pointerup', end); hd.addEventListener('pointercancel', end);
    });
  }

  async function takePhoto(file) {
    if (!file) return;
    busy(true, 'Opening photo…');
    try {
      work = await loadPhoto(file);
      const { corners: c, found } = detectCorners(work);
      corners = c;
      show('crop'); layoutCrop();
      msg(found ? 'Page found. Drag the orange corners if any are off.' : "Couldn't see the page edges. Drag the corners onto the page.", found ? 'info' : 'warn');
    } catch (e) { msg(e.message || "Couldn't open that photo."); }
    busy(false);
    $('scan-cam').value = ''; $('scan-pick').value = '';
  }

  async function toEnhance() {
    busy(true, 'Flattening the page…');
    await new Promise((r) => setTimeout(r, 30));            // let the spinner paint
    try {
      base = warp(work, corners);
      show('enhance'); await renderEnhance();
    } catch (e) { msg(e.message || "Couldn't process that page."); }
    busy(false);
  }
  async function renderEnhance() {
    busy(true, 'Cleaning up…');
    await new Promise((r) => setTimeout(r, 30));
    const out = applyFilter(base, filter);
    const view2 = $('enhance-canvas');
    const maxW = $('scan-enhance').clientWidth || 340, maxH = Math.max(240, (window.innerHeight || 700) - 300);
    const s = Math.min(maxW / out.width, maxH / out.height);
    view2.width = Math.round(out.width * s); view2.height = Math.round(out.height * s);
    view2.getContext('2d').drawImage(out, 0, 0, view2.width, view2.height);
    view2._full = out;
    document.querySelectorAll('#scan-filters .chip').forEach((c) => c.setAttribute('aria-pressed', String(c.dataset.f === filter)));
    busy(false);
  }
  async function addPage() {
    const full = $('enhance-canvas')._full;
    const blob = await new Promise((res) => full.toBlob(res, 'image/jpeg', 0.85));
    pages.push({ blob, url: URL.createObjectURL(blob), w: full.width, h: full.height });
    renderReview(); show('review');
  }
  function renderReview() {
    const grid = $('scan-grid'); grid.replaceChildren();
    pages.forEach((p, i) => {
      const cell = document.createElement('div'); cell.className = 'scan-thumb';
      const img = document.createElement('img'); img.src = p.url; img.alt = `Page ${i + 1}`;
      const num = document.createElement('span'); num.textContent = String(i + 1);
      const del = document.createElement('button'); del.type = 'button'; del.className = 'scan-del'; del.setAttribute('aria-label', `Delete page ${i + 1}`); del.textContent = '×';
      del.addEventListener('click', () => { URL.revokeObjectURL(p.url); pages.splice(i, 1); if (!pages.length) { show('capture'); } else { renderReview(); show('review'); } });
      cell.append(img, num, del); grid.append(cell);
    });
    if (pages.length > 1) fmt = 'pdf';
    document.querySelectorAll('#scan-format .chip').forEach((c) => { c.setAttribute('aria-pressed', String(c.dataset.f === fmt)); c.disabled = c.dataset.f === 'jpg' && pages.length > 1; });
  }
  async function finish() {
    busy(true, 'Making your file…');
    try {
      let blob, ext;
      if (fmt === 'jpg' && pages.length === 1) { blob = pages[0].blob; ext = '.jpg'; }
      else { blob = await buildPdf(pages); ext = '.pdf'; }
      const preview = pages[0].blob;
      const count = pages.length;
      close();
      if (done) done({ blob, ext, preview, count });
    } catch (e) { msg(e.message || "Couldn't make the file."); busy(false); }
  }

  /* ---------- open / close ---------- */
  function close() {
    pages.forEach((p) => URL.revokeObjectURL(p.url)); pages = []; work = null; base = null;
    busy(false); $('scan').hidden = true; document.body.style.overflow = '';
  }
  function open(cb) {
    done = cb; pages = []; filter = 'clean'; fmt = 'pdf';
    $('scan').hidden = false; document.body.style.overflow = 'hidden';
    show('capture');
  }

  let wired = false;
  function wire() {
    if (wired) return; wired = true;
    $('scan-cam').addEventListener('change', (e) => takePhoto(e.target.files[0]));
    $('scan-pick').addEventListener('change', (e) => takePhoto(e.target.files[0]));
    $('scan-cancel').addEventListener('click', close);
    $('scan-back').addEventListener('click', () => show(step === 'enhance' ? 'crop' : pages.length ? 'review' : 'capture'));
    $('scan-primary').addEventListener('click', () => {
      if (step === 'capture' || step === 'review') return pages.length ? finish() : null;
      if (step === 'crop') return toEnhance();
      if (step === 'enhance') return addPage();
    });
    $('scan-auto').addEventListener('click', () => { const r = detectCorners(work); corners = r.corners; drawOverlay(); msg(r.found ? 'Page found.' : "Couldn't see the page edges. Drag the corners onto the page.", r.found ? 'info' : 'warn'); });
    $('scan-full').addEventListener('click', () => { corners = insetCorners(work.width, work.height, 0.01); drawOverlay(); });
    $('scan-rotate').addEventListener('click', () => { work = rotated(work); corners = detectCorners(work).corners; layoutCrop(); });
    $('scan-add').addEventListener('click', () => show('capture'));
    document.querySelectorAll('#scan-filters .chip').forEach((c) => c.addEventListener('click', () => { filter = c.dataset.f; renderEnhance(); }));
    document.querySelectorAll('#scan-format .chip').forEach((c) => c.addEventListener('click', () => { if (!c.disabled) { fmt = c.dataset.f; renderReview(); } }));
    wireHandles();
    window.addEventListener('resize', () => { if (step === 'crop' && work) layoutCrop(); });
  }

  return { open, wire, detectCorners, warp, applyFilter, buildPdf, homography, insetCorners, _state: () => ({ pages, step, corners }) };
})();
