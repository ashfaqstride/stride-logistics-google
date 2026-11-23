'use strict';
/* =====================================================================
   Stride Logistics: offline-first phone app that writes into an Excel
   Google Sheet in Google Drive (Sheets and Drive APIs).
   ===================================================================== */
const CFG = window.CONFIG;
const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad = n => String(n).padStart(2, '0');
const ld = d => d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
const today = () => ld(new Date());
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const fd = d => { if (!d) return ''; const p = String(d).split('-'); return p.length === 3 ? p[2] + '-' + MON[+p[1] - 1] + '-' + p[0] : d; };
const ago = (days, months) => { const d = new Date(); d.setDate(d.getDate() - days); d.setMonth(d.getMonth() - months); return ld(d); };
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const enc = p => p.split('/').map(encodeURIComponent).join('/');
const serial = iso => iso ? Math.round((Date.parse(iso + 'T00:00:00Z') - Date.UTC(1899, 11, 30)) / 86400000) : '';
const iso = n => typeof n === 'number' ? new Date(Date.UTC(1899, 11, 30) + n * 86400000).toISOString().slice(0, 10) : (n || '');
// our own formulas are wrapped as {formula}; all other text is written as text
const cell = v => v == null ? '' : v;   // text is protected from number/date guessing when written (enterVal)

/* ---------- lists (editable in Settings) ---------- */
const LISTS = {
  handlers: { label: 'Sales handlers', col: 'A', max: 10 },
  brands:   { label: 'Brands', col: 'C', max: 30 },
  types:    { label: 'Item types', col: 'G', max: 30 },
  storage:  { label: 'Storage locations', col: 'B', max: 6 },
  units:    { label: 'Units', col: 'D', max: 10 },
  modes:    { label: 'Dispatch modes', col: 'F', max: 10 },
  people:   { label: 'Delivery people', col: 'H', max: 10 }
};
const DEFAULTS = {
  handlers: ['Vishnu', 'Vishnupriyan', 'Cijith', 'Brijith', 'Sreerag'],
  brands: ['Endress+Hauser', 'HORIBA'],
  types: ['Pressure transmitter', 'Temperature transmitter', 'Level transmitter', 'Flow meter', 'Thermowell', 'Thermocouple', 'RTD', 'Pressure gauge', 'Analyser', 'Gas detector', 'Junction box', 'Cable gland', 'Valve'],
  storage: ['Storely', 'Warehouse'],
  units: ['pcs', 'sets', 'nos', 'm', 'kg', 'box'],
  modes: ['Company vehicle', 'Courier', 'Air cargo', 'Sea freight', 'Customer pickup'],
  people: []
};
/* where each list value is used, for rename propagation and "in use" warnings */
const FIELDS = {
  handlers: { pos: ['handler'], dsp: ['handler'] }, brands: { items: ['brand'] }, types: { items: ['type'] },
  storage: { items: ['store'] }, units: { items: ['unit'] }, modes: { dsp: ['mode'] }, people: { dsp: ['by'], items: ['by'] }
};
const REMOTE_COLS = {
  handlers: [['PO Register', 'D', 5, 204], ['Dispatch', 'M', 5, 504]], brands: [['Receiving', 'F', 5, 304]],
  types: [['Receiving', 'E', 5, 304]], storage: [['Receiving', 'I', 5, 304]], units: [['Receiving', 'H', 5, 304]],
  modes: [['Dispatch', 'P', 5, 504]], people: [['Dispatch', 'Q', 5, 504], ['Receiving', 'L', 5, 304]]
};
const OPKEY = { po: 'pos', receive: 'items', dispatch: 'dsp' };

/* ---------- storage: state in IndexedDB, photos as blobs ---------- */
const DB = (() => {
  let db;
  const open = () => new Promise((res, rej) => {
    if (db) return res(db);
    const r = indexedDB.open('stride-logistics', 1);
    r.onupgradeneeded = () => { r.result.createObjectStore('kv'); r.result.createObjectStore('photos'); };
    r.onsuccess = () => { db = r.result; res(db); };
    r.onerror = () => rej(r.error);
  });
  const run = async (store, mode, fn) => {
    const d = await open();
    return new Promise((res, rej) => {
      const t = d.transaction(store, mode); const rq = fn(t.objectStore(store));
      t.oncomplete = () => res(rq.result); t.onerror = t.onabort = () => rej(t.error);
    });
  };
  return {
    get: (k, s = 'kv') => run(s, 'readonly', o => o.get(k)),
    set: (k, v, s = 'kv') => run(s, 'readwrite', o => o.put(v, k)),
    del: (k, s = 'kv') => run(s, 'readwrite', o => o.delete(k))
  };
})();

let S;
const blank = () => ({
  v: 1, user: null, queue: [], idmap: {}, listsDirty: false, lastError: null,
  cache: { items: [], dsp: [], pos: [], at: 0 },
  lists: JSON.parse(JSON.stringify(DEFAULTS)),
  cfg: { maxPhotos: 5, folder: CFG.defaultFolder, file: CFG.defaultFile }
});
async function loadState() {
  let s = null;
  try { s = await DB.get('state'); } catch (e) { /* private mode: start blank */ }
  S = Object.assign(blank(), s || {});
  S.cfg = Object.assign(blank().cfg, S.cfg || {});
}
let saving = Promise.resolve();
const save = () => { saving = saving.then(() => DB.set('state', JSON.parse(JSON.stringify(S)))).catch(() => { }); return saving; };

/* ---------- merged model: workbook cache + entries still waiting to upload ---------- */
function M() {
  const items = S.cache.items.map(i => ({ ...i, key: i.id }));
  const dsp = S.cache.dsp.map(d => ({ ...d, key: d.id }));
  const pos = S.cache.pos.map(p => ({ ...p }));
  for (const op of S.queue) {
    const d = op.data || {};
    if (op.type === 'po') pos.push({ ...d, pending: true });
    else if (op.type === 'receive') items.push({ ...d, key: 'L:' + op.uid, id: null, pending: true, photoCount: op.photos.length });
    else if (op.type === 'dispatch') dsp.push({ ...d, key: 'L:' + op.uid, id: null, pending: true, photoCount: op.photos.length });
  }
  for (const op of S.queue) if (op.type === 'deliver') {
    const x = dsp.find(z => z.key === op.ref || z.id === op.ref);
    if (x) { x.status = 'Delivered'; x.delDate = op.date; x.pendingDeliver = true; }
  }
  return { items, dsp, pos };
}
const sentOf = (m, key) => m.dsp.filter(d => d.item === key).reduce((a, d) => a + (+d.qty || 0), 0);
const balOf = (m, it) => it.qty - sentOf(m, it.key);
const pending = () => S.queue.length + (S.listsDirty ? 1 : 0);

/* ---------- Google sign-in (Google Identity Services) ---------- */
let tokClient = null, account = null;
const TEST_TOKEN = () => { try { return localStorage.getItem('stride_test_token'); } catch (e) { return null; } };
const lsGet = k => { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch (e) { return null; } };
const lsSet = (k, v) => { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, JSON.stringify(v)); } catch (e) { } };
const gsiReady = (ms = 5000) => new Promise(res => {
  const t0 = Date.now();
  (function chk() { if (window.google && google.accounts && google.accounts.oauth2) return res(true); if (Date.now() - t0 > ms) return res(false); setTimeout(chk, 150); })();
});
async function initAuth() {
  if (TEST_TOKEN()) { account = { name: 'Test User', username: 'test@example.com' }; return; }
  account = lsGet('stride_acct');
  if (await gsiReady(account ? 1500 : 5000)) {
    tokClient = google.accounts.oauth2.initTokenClient({ client_id: CFG.googleClientId, scope: CFG.scopes.join(' '), callback: () => { } });
  }
}
function requestToken(prompt, ms) {
  return new Promise((res, rej) => {
    if (!tokClient) return rej(Object.assign(new Error('Google sign-in is not loaded. Check your connection.'), { auth: true }));
    const timer = ms ? setTimeout(() => rej(Object.assign(new Error('Sign in again'), { auth: true })), ms) : null;
    const done = f => v => { clearTimeout(timer); f(v); };
    tokClient.callback = done(r => r.error ? rej(Object.assign(new Error(r.error_description || r.error), { auth: true })) : res(r));
    tokClient.error_callback = done(e => rej(Object.assign(new Error((e && (e.message || e.type)) || 'Sign-in cancelled'), { auth: true })));
    const opt = { prompt };
    if (account && account.username) opt.hint = account.username;
    tokClient.requestAccessToken(opt);
  });
}
function keepToken(r) { lsSet('stride_tok', { t: r.access_token, exp: Date.now() + (+r.expires_in || 3600) * 1000 }); return r.access_token; }
async function getToken() {
  const t = TEST_TOKEN(); if (t) return t;
  const c = lsGet('stride_tok');
  if (c && c.exp - 60000 > Date.now()) return c.t;
  if (!account) { const e = new Error('Sign in again'); e.auth = true; throw e; }
  return keepToken(await requestToken('none', 6000));   // silent refresh; fails if Google needs a tap
}
async function doLogin() {
  try {
    if (!tokClient && await gsiReady(3000)) tokClient = google.accounts.oauth2.initTokenClient({ client_id: CFG.googleClientId, scope: CFG.scopes.join(' '), callback: () => { } });
    const b = $('lg-btn'); if (b) b.disabled = true;
    const tok = keepToken(await requestToken(account ? '' : 'select_account'));
    const u = await (await fetch('https://www.googleapis.com/oauth2/v3/userinfo', { headers: { Authorization: 'Bearer ' + tok } })).json();
    account = { name: u.name || u.email, username: u.email };
    lsSet('stride_acct', account);
    S.user = account.name || 'User'; S.lastError = null; await save();
    if (current === 'login') show('home'); else paint();
    trySync(true);
  } catch (e) {
    const m = $('lg-err'); if (m) { m.style.display = 'block'; m.textContent = 'Sign-in problem: ' + (e.message || e); }
  } finally { const b = $('lg-btn'); if (b) b.disabled = false; }
}
async function doLogout() {
  if (!confirm('Sign out? Entries waiting to upload stay on this phone.')) return;
  try { const c = lsGet('stride_tok'); if (c && window.google) google.accounts.oauth2.revoke(c.t, () => { }); } catch (e) { }
  lsSet('stride_tok', null); lsSet('stride_acct', null); account = null;
  S.user = null; await save(); show('login');
}

/* ---------- Google Sheets + Drive helpers ---------- */
const isNet = e => e instanceof TypeError || !navigator.onLine;
const DRV = 'https://www.googleapis.com/drive/v3', UPL = 'https://www.googleapis.com/upload/drive/v3', SHT = 'https://sheets.googleapis.com/v4/spreadsheets';
async function gapi(url, o = {}) {
  for (let a = 0; ; a++) {
    const token = await getToken();
    const isBlob = o.body instanceof Blob;
    const res = await fetch(url, {
      method: o.method || 'GET',
      headers: Object.assign({ Authorization: 'Bearer ' + token }, o.body && !isBlob ? { 'Content-Type': 'application/json' } : {}, o.headers || {}),
      body: o.body ? ((isBlob || typeof o.body === 'string') ? o.body : JSON.stringify(o.body)) : undefined
    });
    if ((res.status === 429 || res.status === 500 || res.status === 503) && a < 3) { await new Promise(r => setTimeout(r, 1500 * (a + 1))); continue; }
    if (res.status === 401) { lsSet('stride_tok', null); if (a < 1 && !TEST_TOKEN()) continue; const e = new Error('Sign in again'); e.auth = true; throw e; }
    if (!res.ok) {
      const t = await res.text().catch(() => '');
      const m = /"message"\s*:\s*"([^"]*)"/.exec(t);
      const e = new Error('Google ' + res.status + ': ' + (m ? m[1] : t.slice(0, 150)));
      e.status = res.status; throw e;
    }
    return res.status === 204 ? null : await res.json();
  }
}
// multipart upload (metadata + file) to Drive
async function driveUpload(meta, blob, mime) {
  const b = 'stride' + uid();
  const body = new Blob([`--${b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n--${b}\r\nContent-Type: ${mime}\r\n\r\n`, blob, `\r\n--${b}--`]);
  return gapi(`${UPL}/files?uploadType=multipart&supportsAllDrives=true&fields=id,webViewLink`, { method: 'POST', body, headers: { 'Content-Type': 'multipart/related; boundary=' + b } });
}
const qs = s => String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
async function findOne(q) { const r = await gapi(`${DRV}/files?q=${encodeURIComponent(q)}&fields=files(id,webViewLink)&pageSize=1&spaces=drive`); return r.files && r.files[0]; }
async function folderPath(path) {
  S.fids = S.fids || {};
  let parent = 'root', acc = '';
  for (const seg of path.split('/').filter(Boolean)) {
    acc += '/' + seg;
    if (S.fids[acc]) { parent = S.fids[acc]; continue; }
    let f = await findOne(`name='${qs(seg)}' and mimeType='application/vnd.google-apps.folder' and '${parent}' in parents and trashed=false`);
    if (!f) f = await gapi(`${DRV}/files?fields=id,webViewLink`, { method: 'POST', body: { name: seg, mimeType: 'application/vnd.google-apps.folder', parents: [parent] } });
    S.fids[acc] = f.id; parent = f.id;
  }
  return parent;
}
// the sheet is created from the bundled template the first time the app syncs
async function ensureSheet() {
  const MIME = 'application/vnd.google-apps.spreadsheet';
  if (S.cfg.fileId) {
    try { const f = await gapi(`${DRV}/files/${S.cfg.fileId}?fields=id,trashed`); if (!f.trashed) return f.id; }
    catch (e) { if (e.status !== 404) throw e; }
    S.cfg.fileId = null;
  }
  const folder = await folderPath(S.cfg.folder);
  let f = await findOne(`name='${qs(S.cfg.file)}' and mimeType='${MIME}' and '${folder}' in parents and trashed=false`);
  if (!f) {
    const r = await fetch('template.xlsx'); if (!r.ok) throw new Error('Template workbook is missing from the app.');
    f = await driveUpload({ name: S.cfg.file, mimeType: MIME, parents: [folder] }, await r.blob(), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  }
  S.cfg.fileId = f.id; await save();
  return f.id;
}
const rangeUrl = (sh, addr) => ({ sh, addr });
const colN = s => [...s].reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0);
function dims(addr) {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(addr);
  const c1 = colN(m[1]), r1 = +m[2], c2 = m[3] ? colN(m[3]) : c1, r2 = m[4] ? +m[4] : r1;
  return { rows: r2 - r1 + 1, cols: c2 - c1 + 1 };
}
const A1 = (sh, addr) => encodeURIComponent(`'${sh}'!${addr}`);
function grid(vals, d) {   // the API drops empty trailing cells: pad to a full rectangle
  const out = [];
  for (let i = 0; i < d.rows; i++) { const r = (vals && vals[i]) || []; out.push(Array.from({ length: d.cols }, (_, j) => r[j] === undefined ? '' : r[j])); }
  return out;
}
const enterVal = v => {
  if (v && typeof v === 'object' && v.formula) return v.formula;          // our own formulas
  if (typeof v === 'string') return v === '' ? '' : "'" + v;               // text stays text (no date/number guessing)
  return v == null ? '' : v;
};
// compatibility shim: wb(ctx, rangeUrl(...), {method:'PATCH', body:{values|formulas}})
async function wb(ctx, p, o = {}) {
  const d = dims(p.addr);
  if (o.method === 'PATCH') {
    const rows = (o.body.formulas || o.body.values).map(r => r.map(enterVal));
    return gapi(`${SHT}/${ctx.fileId}/values/${A1(p.sh, p.addr)}?valueInputOption=USER_ENTERED`, { method: 'PUT', body: { range: `'${p.sh}'!${p.addr}`, majorDimension: 'ROWS', values: rows } });
  }
  const base = `${SHT}/${ctx.fileId}/values/${A1(p.sh, p.addr)}?majorDimension=ROWS`;
  const [v, f] = await Promise.all([gapi(base + '&valueRenderOption=UNFORMATTED_VALUE'), gapi(base + '&valueRenderOption=FORMULA')]);
  return { values: grid(v.values, d), formulas: grid(f.values, d) };
}
const putRow = (ctx, sh, addr, row) => wb(ctx, rangeUrl(sh, addr), { method: 'PATCH', body: { formulas: [row] } });
async function nextRow(ctx, sh, idCol, keyCol, r1, r2) {
  const d = await wb(ctx, rangeUrl(sh, `${idCol}${r1}:${keyCol}${r2}`));
  const last = d.values[0].length - 1;
  const i = d.values.findIndex(r => r[last] === '' || r[last] == null);
  if (i < 0) throw new Error(sh + ' sheet is full. Start a new workbook.');
  return { row: r1 + i, id: d.values[i][0] };
}
async function uploadPhotos(ctx, keys, date, id, prefix) {
  if (!keys || !keys.length) return null;
  const folder = await folderPath(`${S.cfg.folder}/Photos/${(date || today()).slice(0, 7)}/${id}`);
  let n = 0;
  for (const k of keys) {
    const blob = await DB.get(k, 'photos'); if (!blob) continue;
    n++;
    await driveUpload({ name: prefix + '_' + n + '.jpg', parents: [folder] }, blob, 'image/jpeg');
  }
  if (!n) return null;
  let count = n;
  try { const c = await gapi(`${DRV}/files?q=${encodeURIComponent(`'${folder}' in parents and trashed=false`)}&fields=files(id)&pageSize=100`); count = c.files.length; } catch (e) { /* keep n */ }
  return { url: 'https://drive.google.com/drive/folders/' + folder, count };
}
const linkF = p => ({ formula: `=HYPERLINK("${p.url}","${p.count} photo${p.count > 1 ? 's' : ''}")` });
const dropPhotos = async keys => { for (const k of keys || []) { try { await DB.del(k, 'photos'); } catch (e) { } } };
const resolve = k => (k && String(k).startsWith('L:')) ? (S.idmap[k] || k) : k;
function relink(uidKey, id) {
  for (const op of S.queue) {
    if (op.data && op.data.item === uidKey) op.data.item = id;
    if (op.ref === uidKey) op.ref = id;
  }
}

/* ---------- sync: upload queued entries, then refresh from the workbook ---------- */
async function pushLists(ctx) {
  for (const m of Object.entries(LISTS)) {
    const [k, meta] = m;
    const vals = S.lists[k].slice(0, meta.max); while (vals.length < meta.max) vals.push('');
    await wb(ctx, rangeUrl('Settings', `${meta.col}7:${meta.col}${6 + meta.max}`), { method: 'PATCH', body: { values: vals.map(v => [cell(v)]) } });
  }
  S.listsDirty = false;
}
async function renameCells(ctx, sh, col, r1, r2, fn) {
  const d = await wb(ctx, rangeUrl(sh, `${col}${r1}:${col}${r2}`));
  let changed = false;
  const out = d.values.map(r => { const v = r[0]; if (typeof v !== 'string') return [v]; const n = fn(v); if (n !== v) changed = true; return [n]; });
  if (changed) await wb(ctx, rangeUrl(sh, `${col}${r1}:${col}${r2}`), { method: 'PATCH', body: { values: out } });
}
async function runOp(op, ctx) {
  const d = op.data || {};
  switch (op.type) {
    case 'po': {
      const { row } = await nextRow(ctx, 'PO Register', 'A', 'A', 5, 204);
      await putRow(ctx, 'PO Register', `A${row}:E${row}`, [d.po, serial(d.date), d.supplier || '', d.handler || '', d.remarks || '']);
      S.cache.pos.push({ ...d }); break;
    }
    case 'receive': {
      const { row, id } = await nextRow(ctx, 'Receiving', 'A', 'B', 5, 304);
      const ph = await uploadPhotos(ctx, op.photos, d.date, id, 'received');
      await putRow(ctx, 'Receiving', `B${row}:N${row}`, [d.po, serial(d.date), d.code || '', d.type, d.brand || '', d.qty, d.unit || '', d.store || '', d.dest || '', serial(d.plan), d.by || '', ph ? linkF(ph) : '', d.remarks || '']);
      S.cache.items.push({ ...d, id, photoUrl: ph ? ph.url : '', photoCount: ph ? ph.count : 0 });
      S.idmap['L:' + op.uid] = id; relink('L:' + op.uid, id); await dropPhotos(op.photos); break;
    }
    case 'dispatch': {
      const itemId = resolve(d.item);
      const it = S.cache.items.find(i => i.id === itemId);
      if (!it) throw new Error('Item ' + d.item + ' is not in the workbook yet. Sync the receive entry first.');
      const { row, id } = await nextRow(ctx, 'Dispatch', 'A', 'C', 5, 504);
      const ph = await uploadPhotos(ctx, op.photos, d.date, id, 'dispatch');
      const label = [it.id, it.po, it.code || '', it.type].join(' | ');
      // details first, item label last: a failed attempt then leaves the row free to reuse
      await putRow(ctx, 'Dispatch', `K${row}:U${row}`, [d.sq, d.client || '', d.handler || '', d.qty, d.dest || '', d.mode || '', d.by || '', d.status || 'In Transit', serial(d.delDate), ph ? linkF(ph) : '', d.remarks || '']);
      await putRow(ctx, 'Dispatch', `B${row}:C${row}`, [serial(d.date), label]);
      S.cache.dsp.push({ ...d, item: it.id, id, photoUrl: ph ? ph.url : '', photoCount: ph ? ph.count : 0 });
      S.idmap['L:' + op.uid] = id; relink('L:' + op.uid, id); await dropPhotos(op.photos); break;
    }
    case 'deliver': {
      const id = resolve(op.ref);
      if (String(id).startsWith('L:')) throw new Error('The dispatch to mark as delivered is not uploaded yet.');
      const col = await wb(ctx, rangeUrl('Dispatch', 'A5:C504'));
      const idx = col.values.findIndex(r => r[0] === id && r[2] !== '');
      if (idx < 0) throw new Error('Dispatch ' + id + ' was not found in the workbook.');
      const row = 5 + idx;
      const ph = await uploadPhotos(ctx, op.photos, op.date, id, 'delivered');
      await putRow(ctx, 'Dispatch', `R${row}:S${row}`, ['Delivered', serial(op.date)]);
      if (ph) await putRow(ctx, 'Dispatch', `T${row}`, [linkF(ph)]);
      const c = S.cache.dsp.find(x => x.id === id);
      if (c) { c.status = 'Delivered'; c.delDate = op.date; if (ph) { c.photoUrl = ph.url; c.photoCount = ph.count; } }
      await dropPhotos(op.photos); break;
    }
    case 'rename': {
      for (const [sh, col, r1, r2] of REMOTE_COLS[op.list]) await renameCells(ctx, sh, col, r1, r2, v => v === op.from ? op.to : v);
      if (op.list === 'types') await renameCells(ctx, 'Dispatch', 'C', 5, 504, v => v.endsWith(' | ' + op.from) ? v.slice(0, -op.from.length) + op.to : v);
      break;
    }
    default: throw new Error('Unknown entry type ' + op.type);
  }
}
async function pull(ctx) {
  const g = (sh, a) => wb(ctx, rangeUrl(sh, a));
  const rec = await g('Receiving', 'A5:O304'), dsp = await g('Dispatch', 'A5:U504'), po = await g('PO Register', 'A5:E204'), set = await g('Settings', 'A7:H36');
  const link = f => { const m = /HYPERLINK\("([^"]+)"/.exec(f || ''); return m ? m[1] : ''; };
  const cnt = f => { const m = /HYPERLINK\("[^"]+"[,;]"(\d+)/.exec(f || ''); return m ? +m[1] : 0; };
  const str = v => v == null ? '' : String(v);
  const items = [], dsps = [], pos = [];
  rec.values.forEach((r, i) => {
    if (r[1] === '' || r[1] == null) return; const f = rec.formulas[i];
    items.push({ id: str(r[0]), po: str(r[1]), date: iso(r[2]), code: str(r[3]), type: str(r[4]), brand: str(r[5]), qty: +r[6] || 0, unit: str(r[7]), store: str(r[8]), dest: str(r[9]), plan: iso(r[10]), by: str(r[11]), photoUrl: link(f[12]), photoCount: cnt(f[12]), remarks: str(r[13]) });
  });
  dsp.values.forEach((r, i) => {
    if (r[2] === '' || r[2] == null) return; const f = dsp.formulas[i];
    dsps.push({ id: str(r[0]), date: iso(r[1]), item: str(r[2]).split(' | ')[0], sq: str(r[10]), client: str(r[11]), handler: str(r[12]), qty: +r[13] || 0, dest: str(r[14]), mode: str(r[15]), by: str(r[16]), status: str(r[17]), delDate: iso(r[18]), photoUrl: link(f[19]), photoCount: cnt(f[19]), remarks: str(r[20]) });
  });
  po.values.forEach(r => { if (r[0] === '' || r[0] == null) return; pos.push({ po: str(r[0]), date: iso(r[1]), supplier: str(r[2]), handler: str(r[3]), remarks: str(r[4]) }); });
  S.cache = { items, dsp: dsps, pos, at: Date.now() };
  for (const [k, meta] of Object.entries(LISTS)) {
    const idx = meta.col.charCodeAt(0) - 65;
    const remote = set.values.slice(0, meta.max).map(r => str(r[idx]).trim()).filter(Boolean);
    if (S.listsDirty) continue;
    if (remote.length) S.lists[k] = remote;
    else if (S.lists[k].length) S.listsDirty = true;   // workbook list empty: push ours on the next sync
  }
}
let syncing = false;
async function sync(manual) {
  if (syncing) return;
  if (!navigator.onLine) { if (manual) toast('No signal. Entries stay on the phone.'); return; }
  if (!S.user) return;
  syncing = true; paint();
  const ctx = {};
  try {
    ctx.fileId = await ensureSheet();
    if (S.listsDirty) { await pushLists(ctx); await save(); }
    while (S.queue.length) {
      const op = S.queue[0];
      try { await runOp(op, ctx); }
      catch (e) { if (isNet(e) || e.auth) throw e; op.error = e.message; await save(); throw e; }
      S.queue.shift(); await save(); paint();
    }
    await pull(ctx);
    S.lastError = null; await save();
    if (manual) toast('Synced with Google Sheets');
  } catch (e) {
    if (e.auth) S.lastError = 'Please sign in again to sync.';
    else if (!isNet(e)) S.lastError = e.message;
    await save();
  } finally {
    syncing = false; paint(); refreshOpen();
  }
}
const trySync = manual => { sync(manual).catch(() => { }); };

/* =====================================================================
   UI
   ===================================================================== */
function toast(m) { const t = $('toast'); t.textContent = m; t.classList.add('on'); clearTimeout(t._h); t._h = setTimeout(() => t.classList.remove('on'), 2800); }
let current = 'login';
function show(id) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('on'));
  $('s-' + id).classList.add('on'); current = id; paint(); refreshOpen();
  if (id === 'po' && !$('po-d').value) $('po-d').value = today();
  if (id === 'po') fillSelect('po-h', S.lists.handlers, false, true, 'Not assigned');
  if (id === 'stk') { fillSeg('s-loc', ['All', ...S.lists.storage], true); renderStock(); }
  if (id === 'fnd') { $('po-dl').innerHTML = M().pos.map(p => `<option value="${esc(p.po)}">`).join(''); renderFind(); }
  if (id === 'syn') renderSync();
  if (id === 'set') renderSettings();
}
function refreshOpen() { if (current === 'stk') renderStock(); else if (current === 'fnd') renderFind(); else if (current === 'syn') renderSync(); }
function pick(g, el) { [...$(g).children].forEach(c => c.classList.remove('sel')); el.classList.add('sel'); }
const val = g => { const e = $(g).querySelector('.sel'); return e ? e.textContent.trim() : ''; };
function fillSeg(id, arr, selFirst) {
  const keep = val(id);
  $(id).innerHTML = arr.map(v => `<div class="chip" onclick="pick('${id}',this)">${esc(v)}</div>`).join('');
  const hit = [...$(id).children].find(c => c.textContent === keep) || (selFirst ? $(id).children[0] : $(id).children[0]);
  if (hit) hit.classList.add('sel');
  if (id === 's-loc') [...$(id).children].forEach(c => c.onclick = () => { pick(id, c); renderStock(); });
}
function fillSelect(id, arr, other, keep, blankLabel) {
  const el = $(id), cur = keep ? el.value : '';
  el.innerHTML = (blankLabel ? `<option value="">${esc(blankLabel)}</option>` : '') + arr.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('') + (other ? '<option value="__o">Other…</option>' : '');
  if (cur && [...el.options].some(o => o.value === cur)) el.value = cur;
}
const selVal = (id, oid) => { const v = $(id).value; return v === '__o' ? $(oid).value.trim() : v; };
const onOther = (id, oid) => { $(oid).style.display = $(id).value === '__o' ? 'block' : 'none'; };
function addToList(key, v) {
  const l = S.lists[key];
  if (v && !l.some(x => x.toLowerCase() === v.toLowerCase()) && l.length < LISTS[key].max) { l.push(v); S.listsDirty = true; }
}

/* ---------- status / home ---------- */
const online = () => navigator.onLine;
function paint() {
  document.querySelectorAll('.dot').forEach(d => d.classList.toggle('off', !online()));
  document.querySelectorAll('.mp').forEach(e => e.textContent = S.cfg.maxPhotos);
  const p = pending(), b = $('pb');
  b.textContent = syncing ? 'Syncing…' : (p ? p + ' waiting to sync' : 'All synced'); b.className = 'badge' + (p && !syncing ? ' pend' : '');
  const n = $('sy-n'); n.style.display = p ? 'block' : 'none'; n.textContent = p;
  const at = S.cache.at ? new Date(S.cache.at).toLocaleString() : 'never';
  $('cx').textContent = (online() ? 'Online. ' : 'Offline. Entries are saved on this phone. ') + 'Last synced: ' + at + '.';
  const e = $('h-err'); e.style.display = S.lastError ? 'block' : 'none'; e.textContent = S.lastError || '';
  if (S.user) $('hi').textContent = 'Hello, ' + String(S.user).split(' ')[0];
}
window.addEventListener('online', () => { paint(); trySync(); });
window.addEventListener('offline', paint);

/* ---------- photos: up to the Settings limit, shrunk before saving ---------- */
const P = { r: [], d: [], v: [] }; let pctx = 'r';
function addPhoto(c) { if (P[c].length >= S.cfg.maxPhotos) return toast('Maximum ' + S.cfg.maxPhotos + ' photos (change it in Settings)'); pctx = c; $('ph-in').click(); }
$('ph-in').onchange = async function () {
  const f = this.files[0]; this.value = ''; if (!f) return;
  try {
    const bmp = await createImageBitmap(f);
    const sc = Math.min(1, 1280 / Math.max(bmp.width, bmp.height));
    const cv = document.createElement('canvas'); cv.width = Math.round(bmp.width * sc); cv.height = Math.round(bmp.height * sc);
    cv.getContext('2d').drawImage(bmp, 0, 0, cv.width, cv.height);
    const blob = await new Promise(r => cv.toBlob(r, 'image/jpeg', 0.72));
    const key = 'p_' + uid(); await DB.set(key, blob, 'photos');
    P[pctx].push({ key, url: URL.createObjectURL(blob) }); renderPh(pctx);
  } catch (e) { toast('Could not read that photo'); }
};
function renderPh(c) {
  $(c + '-ph').innerHTML = P[c].map((p, i) => `<div><img src="${p.url}"><button onclick="rmPhoto('${c}',${i})">&times;</button></div>`).join('');
  $(c + '-phn').textContent = P[c].length;
}
async function rmPhoto(c, i) { const p = P[c].splice(i, 1)[0]; if (p) await dropPhotos([p.key]); renderPh(c); }
const takePhotos = c => { const k = P[c].map(p => p.key); P[c] = []; renderPh(c); return k; };

async function addOp(op) {
  op.uid = uid(); S.queue.push(op); await save(); paint();
  if (online()) trySync();
}

/* ---------- new PO ---------- */
async function savePO() {
  const po = $('po-no').value.trim(); if (!po) return toast('Enter the PO number');
  if (M().pos.some(p => p.po.toLowerCase() === po.toLowerCase())) return toast('This PO already exists');
  await addOp({ type: 'po', data: { po, date: $('po-d').value || today(), supplier: $('po-sup').value.trim(), handler: $('po-h').value, remarks: $('po-rm').value.trim() } });
  ['po-no', 'po-sup', 'po-rm'].forEach(i => $(i).value = ''); toast('PO saved');
  openReceive(po);
}

/* ---------- receive ---------- */
let curType = null;
function poOptions(sel, keep) { fillSelect(sel, M().pos.map(p => p.po)); if (keep) $(sel).value = keep; }
function openReceive(po) {
  poOptions('r-po', po); $('r-date').value = today(); curType = null;
  $('r-types').innerHTML = [...S.lists.types, 'Other'].map(t => `<div class="chip" onclick="setType(this)">${esc(t)}</div>`).join('');
  $('r-type-o').style.display = 'none';
  fillSelect('r-brand', S.lists.brands, true); $('r-brand-o').style.display = 'none';
  fillSelect('r-unit', S.lists.units, false);
  fillSeg('r-store', S.lists.storage, true);
  fillSelect('r-by', S.lists.people, true, false, 'Not set'); $('r-by-o').style.display = 'none';
  show('rec');
}
function setType(el) {
  document.querySelectorAll('#r-types .chip').forEach(c => c.classList.remove('sel')); el.classList.add('sel');
  curType = el.textContent; $('r-type-o').style.display = curType === 'Other' ? 'block' : 'none';
}
async function saveItem() {
  if (!$('r-po').value) return toast('Create a PO first');
  let t = curType; if (!t) return toast('Pick an item type');
  if (t === 'Other') { t = $('r-type-o').value.trim(); if (!t) return toast('Type the item name'); }
  const q = parseFloat($('r-qty').value); if (!(q > 0)) return toast('Enter the quantity');
  const brand = selVal('r-brand', 'r-brand-o'), by = selVal('r-by', 'r-by-o');
  addToList('types', t); addToList('brands', brand); addToList('people', by);
  const po = $('r-po').value;
  await addOp({ type: 'receive', photos: takePhotos('r'), data: { po, date: $('r-date').value || today(), code: $('r-code').value.trim(), type: t, brand, qty: q, unit: $('r-unit').value, store: val('r-store'), dest: $('r-dest').value.trim(), plan: $('r-plan').value, by, remarks: $('r-rm').value.trim() } });
  ['r-code', 'r-qty', 'r-dest', 'r-plan', 'r-rm'].forEach(i => $(i).value = ''); toast('Item saved');
  openReceive(po);
}

/* ---------- dispatch ---------- */
function openDispatch() {
  poOptions('d-po'); $('d-date').value = today(); fillItems();
  fillSelect('d-h', S.lists.handlers, false); fillSelect('d-by', S.lists.people, true, false, 'Not set'); $('d-by-o').style.display = 'none';
  fillSelect('d-mode', S.lists.modes, false); show('dsp');
}
function fillItems() {
  const m = M(), po = $('d-po').value, list = m.items.filter(i => i.po === po);
  $('d-item').innerHTML = list.map(i => `<option value="${esc(i.key)}">${esc(i.type)} - ${esc(i.code || 'no code')} (left: ${balOf(m, i)})</option>`).join('') || '<option value="">No items received for this PO</option>';
}
async function saveDsp() {
  const m = M(), key = $('d-item').value, it = m.items.find(i => i.key === key); if (!it) return toast('Pick an item');
  const q = parseFloat($('d-qty').value); if (!(q > 0)) return toast('Enter quantity sent');
  if (q > balOf(m, it)) return toast('Only ' + balOf(m, it) + ' left in stock. Cannot send ' + q);
  const sq = $('d-sq').value.trim(); if (!sq) return toast('Enter the SQ number');
  const by = selVal('d-by', 'd-by-o'); addToList('people', by);
  const status = val('d-st'), date = $('d-date').value || today();
  await addOp({ type: 'dispatch', photos: takePhotos('d'), data: { item: key, date, sq, client: $('d-cl').value.trim(), handler: $('d-h').value, qty: q, dest: $('d-dest').value.trim(), mode: $('d-mode').value, by, status, delDate: status === 'Delivered' ? date : '', remarks: $('d-rm').value.trim() } });
  ['d-sq', 'd-qty', 'd-cl', 'd-dest', 'd-rm'].forEach(i => $(i).value = ''); toast('Dispatch saved');
  fillItems(); fillSelect('d-by', S.lists.people, true, true, 'Not set');
}

/* ---------- mark delivered ---------- */
let dlvKey = null;
function openDeliver(key) {
  const m = M(), d = m.dsp.find(x => x.key === key); if (!d) return;
  const it = m.items.find(i => i.key === d.item) || {};
  dlvKey = key; $('v-date').value = today();
  $('v-info').innerHTML = `<b>${esc(it.type || d.item)}</b><div class="muted">${esc(it.po || '')} &rarr; ${esc(d.sq)} &middot; ${d.qty} sent ${fd(d.date)}</div>`;
  show('dlv');
}
async function saveDeliver() {
  await addOp({ type: 'deliver', ref: dlvKey, date: $('v-date').value || today(), photos: takePhotos('v') });
  toast('Marked as delivered'); show('fnd');
}

/* ---------- stock ---------- */
let per = '1 month';
function setPer(el) { pick('s-per', el); per = el.textContent.trim(); $('s-dates').style.display = per === 'Custom' ? 'flex' : 'none'; renderStock(); }
function range() {
  const to = today();
  if (per === '48 hrs') return [ago(2, 0), to]; if (per === '1 month') return [ago(0, 1), to]; if (per === '2 months') return [ago(0, 2), to]; if (per === '3 months') return [ago(0, 3), to];
  if (per === 'Custom') return [$('s-from').value || '', $('s-to').value || to]; return ['', ''];
}
const inR = (d, r) => !!d && (!r[0] || d >= r[0]) && (!r[1] || d <= r[1]);
function renderStock() {
  const m = M(), r = range(), loc = val('s-loc') || 'All', q = $('s-q').value.trim().toLowerCase(), view = val('s-view');
  $('s-range').textContent = per === 'All' ? 'Showing all dates' : 'Showing ' + (r[0] ? fd(r[0]) : 'start') + ' to ' + fd(r[1]);
  const match = i => (loc === 'All' || i.store === loc) && (!q || (i.po + ' ' + i.type + ' ' + i.code + ' ' + i.brand).toLowerCase().includes(q));
  let h = '';
  if (view === 'Items') {
    const list = m.items.filter(i => match(i) && (per === 'All' || inR(i.date, r) || m.dsp.some(d => d.item === i.key && inR(d.date, r))));
    h = list.map(i => {
      const ds = m.dsp.filter(d => d.item === i.key), sent = sentOf(m, i.key);
      const inP = inR(i.date, r) ? i.qty : 0, outP = ds.filter(d => inR(d.date, r)).reduce((a, d) => a + d.qty, 0);
      return `<div class="card"><div class="row"><b>${esc(i.type)}</b><span class="pill ${i.store === S.lists.storage[0] ? 's' : 'w'}">${esc(i.store)}</span></div>
        <div class="muted">${esc(i.po)} &middot; ${esc(i.brand)}${i.code ? ' &middot; ' + esc(i.code) : ''} &middot; received ${fd(i.date)}${i.pending ? ' <span class="pill p">on phone</span>' : ''}</div>
        <div class="n"><div><b>${i.qty}</b>Received</div><div><b>${sent}</b>Sent</div><div><b>${i.qty - sent}</b>In stock</div></div>
        ${per === 'All' ? '' : `<div class="muted" style="margin-top:6px">In this period: +${inP} in, &minus;${outP} out</div>`}
        ${ds.map(d => `<span class="pill">${esc(d.sq)} &times; ${d.qty} (${fd(d.date)})</span>`).join('')}</div>`;
    }).join('');
  } else {
    const ev = [];
    m.items.filter(match).forEach(i => {
      if (per === 'All' || inR(i.date, r)) ev.push({ d: i.date, t: 'IN', i, txt: i.qty + ' ' + i.unit + ' received' });
      m.dsp.filter(x => x.item === i.key && (per === 'All' || inR(x.date, r))).forEach(x => ev.push({ d: x.date, t: 'OUT', i, txt: x.qty + ' sent to ' + x.sq }));
    });
    ev.sort((a, b) => String(b.d).localeCompare(String(a.d)));
    h = ev.map(e => `<div class="card"><div class="row"><span class="${e.t === 'IN' ? 'in' : 'out'}">${e.t}</span><span class="muted">${fd(e.d)}</span></div><b>${esc(e.i.type)}</b><div class="muted">${esc(e.i.po)} &middot; ${esc(e.txt)}</div></div>`).join('');
  }
  $('s-list').innerHTML = h || '<div class="muted">Nothing in this period.</div>';
}

/* ---------- find PO: full history ---------- */
const photoBit = x => x.photoUrl ? `<div><a href="${esc(x.photoUrl)}" target="_blank" rel="noopener">&#128247; ${x.photoCount} photo${x.photoCount > 1 ? 's' : ''}</a></div>` : (x.photoCount ? `<div class="muted">&#128247; ${x.photoCount} on phone, uploading</div>` : '');
function renderFind() {
  const q = $('f-q').value.trim().toLowerCase();
  if (!q) { $('f-list').innerHTML = '<div class="muted">Type a PO number to see its items and everything that happened to each one.</div>'; return; }
  const m = M(), list = m.pos.filter(p => p.po.toLowerCase().includes(q));
  $('f-list').innerHTML = list.map(p => {
    const its = m.items.filter(i => i.po === p.po), rec = its.reduce((a, i) => a + i.qty, 0), sent = its.reduce((a, i) => a + sentOf(m, i.key), 0);
    return `<div class="card"><div class="row"><b>${esc(p.po)}</b><span class="muted">${fd(p.date)}</span></div>
      <div class="muted">${esc(p.supplier || 'No supplier')}${p.handler ? ' &middot; ' + esc(p.handler) : ''}${p.pending ? ' <span class="pill p">on phone</span>' : ''}</div>
      <div class="n"><div><b>${rec}</b>Received</div><div><b>${sent}</b>Sent</div><div><b>${rec - sent}</b>In stock</div></div></div>` +
      (its.length ? its.map(i => {
        const ds = m.dsp.filter(d => d.item === i.key).sort((a, b) => String(a.date).localeCompare(String(b.date))), bal = balOf(m, i);
        return `<div class="card"><div class="row"><b>${esc(i.type)}</b><span class="pill ${i.store === S.lists.storage[0] ? 's' : 'w'}">${esc(i.store)}</span></div>
          <div class="muted">${esc(i.brand)}${i.code ? ' &middot; ' + esc(i.code) : ''} &middot; in stock now: <b>${bal}</b> of ${i.qty} ${esc(i.unit)}</div>
          <div class="tl"><div><span class="in">IN</span> ${fd(i.date)}: ${i.qty} ${esc(i.unit)} received, stored at ${esc(i.store)}${i.pending ? ' <span class="pill p">on phone</span>' : ''}${photoBit(i)}</div>
          ${ds.map(d => `<div><span class="out">OUT</span> ${fd(d.date)}: ${d.qty} &rarr; ${esc(d.sq)}${d.client ? ' (' + esc(d.client) + ')' : ''}${d.by ? ', by ' + esc(d.by) : ''}${d.mode ? ', ' + esc(d.mode) : ''}
            <span class="pill ${d.status === 'Delivered' ? 'ok' : 'p'}">${esc(d.status || 'Pending')}${d.status === 'Delivered' && d.delDate ? ' ' + fd(d.delDate) : ''}</span>${d.pending ? ' <span class="pill p">on phone</span>' : ''}${photoBit(d)}
            ${d.status !== 'Delivered' ? `<div><button class="chip" onclick="openDeliver('${esc(d.key)}')">Mark delivered</button></div>` : ''}</div>`).join('')}
          ${bal > 0 ? `<div class="muted">Still here: ${bal} ${esc(i.unit)} in ${esc(i.store)}${i.dest ? ', planned for ' + esc(i.dest) : ''}</div>` : '<div class="muted">All dispatched.</div>'}</div></div>`;
      }).join('') : '<div class="muted" style="margin:0 0 10px">No items received for this PO yet.</div>');
  }).join('') || `<div class="muted">No PO matches "${esc($('f-q').value)}".</div>`;
}

/* ---------- sync screen ---------- */
function describe(op) {
  const d = op.data || {};
  if (op.type === 'po') return 'New PO ' + d.po;
  if (op.type === 'receive') return `Received ${d.qty} ${d.type} (${d.po})` + (op.photos.length ? `, ${op.photos.length} photo(s)` : '');
  if (op.type === 'dispatch') return `Dispatch ${d.qty} to ${d.sq}` + (op.photos.length ? `, ${op.photos.length} photo(s)` : '');
  if (op.type === 'deliver') return 'Mark delivered ' + op.ref;
  if (op.type === 'rename') return `Rename ${LISTS[op.list].label.toLowerCase()}: ${op.from} → ${op.to}`;
  return op.type;
}
function renderSync() {
  const p = pending();
  $('sy-t').textContent = syncing ? 'Syncing…' : (p ? p + ' waiting to upload' : 'All synced');
  $('sy-s').textContent = (online() ? 'Online' : 'Offline') + '. Last synced: ' + (S.cache.at ? new Date(S.cache.at).toLocaleString() : 'never') + '.';
  $('sy-b').disabled = syncing;
  const e = $('sy-err'); e.style.display = S.lastError ? 'block' : 'none'; e.textContent = S.lastError || '';
  if (S.lastError && /sign in/i.test(S.lastError)) e.innerHTML = esc(S.lastError) + ' <button class="btn sm" onclick="doLogin()">Sign in</button>';
  $('sy-q').innerHTML = S.queue.map((op, i) => `<div class="card"><div>${esc(describe(op))}</div>${op.error ? `<div class="muted" style="color:var(--r)">${esc(op.error)}</div>` : ''}<button class="btn red sm" onclick="discardOp(${i})">Discard</button></div>`).join('')
    + (S.listsDirty ? '<div class="card"><div>Settings lists changed</div></div>' : '') || '<div class="muted">Nothing waiting.</div>';
}
async function discardOp(i) {
  const op = S.queue[i]; if (!op || !confirm('Discard this entry? It will not be uploaded.')) return;
  S.queue.splice(i, 1); await dropPhotos(op.photos); await save(); paint(); renderSync();
}

/* ---------- settings ---------- */
function renderSettings() {
  $('c-mp').value = S.cfg.maxPhotos; $('c-folder').value = S.cfg.folder; $('c-file').value = S.cfg.file;
  { const o = $('c-open'); if (o) { o.style.display = S.cfg.fileId ? 'inline-block' : 'none'; o.href = 'https://docs.google.com/spreadsheets/d/' + S.cfg.fileId; } }
  $('c-user').textContent = 'Signed in as ' + (S.user || '') + (account && account.username ? ' (' + account.username + ')' : '');
  $('set-lists').innerHTML = Object.entries(LISTS).map(([k, meta]) => `<div class="card"><b>${meta.label}</b> <span class="muted">(max ${meta.max})</span>
    ${S.lists[k].map((v, i) => `<div class="li"><span>${esc(v)}</span><button onclick="renameItem('${k}',${i})">&#9998;</button><button class="x" onclick="deleteItem('${k}',${i})">&times;</button></div>`).join('') || '<div class="muted">Empty</div>'}
    <div class="add"><input id="add-${k}" placeholder="Add new"><button onclick="addItem('${k}')">Add</button></div></div>`).join('');
}
async function saveCfg() {
  const mp = Math.max(1, Math.min(20, parseInt($('c-mp').value) || 5));
  S.cfg = { fileId: (S.cfg.fileId && $('c-folder').value.trim() === S.cfg.folder && $('c-file').value.trim() === S.cfg.file) ? S.cfg.fileId : null, maxPhotos: mp, folder: $('c-folder').value.trim().replace(/^\/+|\/+$/g, '') || CFG.defaultFolder, file: $('c-file').value.trim() || CFG.defaultFile };
  await save(); paint(); toast('Saved'); trySync();
}
async function addItem(k) {
  const v = $('add-' + k).value.trim(); if (!v) return;
  if (S.lists[k].some(x => x.toLowerCase() === v.toLowerCase())) return toast('Already in the list');
  if (S.lists[k].length >= LISTS[k].max) return toast('This list is full (max ' + LISTS[k].max + ')');
  S.lists[k].push(v); S.listsDirty = true; await save(); renderSettings(); paint(); if (online()) trySync();
}
function applyRenameLocal(k, from, to) {
  const f = FIELDS[k];
  for (const [coll, fields] of Object.entries(f)) {
    for (const o of S.cache[coll]) for (const fl of fields) if (o[fl] === from) o[fl] = to;
    for (const op of S.queue) if (op.data && OPKEY[op.type] === coll) for (const fl of fields) if (op.data[fl] === from) op.data[fl] = to;
  }
}
async function renameItem(k, i) {
  const from = S.lists[k][i], to = (prompt('Rename "' + from + '" to:', from) || '').trim();
  if (!to || to === from) return;
  if (S.lists[k].some(x => x.toLowerCase() === to.toLowerCase())) return toast('That name is already in the list');
  S.lists[k][i] = to; S.listsDirty = true; applyRenameLocal(k, from, to);
  S.queue.push({ type: 'rename', uid: uid(), list: k, from, to, photos: [] });
  await save(); renderSettings(); paint(); toast('Renamed. Old entries update on the next sync.'); if (online()) trySync();
}
async function deleteItem(k, i) {
  const v = S.lists[k][i], m = M();
  const used = Object.entries(FIELDS[k]).reduce((a, [coll, fl]) => a + m[coll].filter(o => fl.some(f => o[f] === v)).length, 0);
  const msg = used ? `"${v}" is used by ${used} existing entr${used > 1 ? 'ies' : 'y'}. Those entries keep the name, but it will no longer appear in the dropdown. Remove it?` : `Remove "${v}"?`;
  if (!confirm(msg)) return;
  S.lists[k].splice(i, 1); S.listsDirty = true; await save(); renderSettings(); paint(); if (online()) trySync();
}

/* ---------- start ---------- */
async function boot() {
  await loadState();
  try { await initAuth(); } catch (e) { S.lastError = 'Sign-in problem: ' + (e.message || e); }
  if (account && !S.user) S.user = account.name || account.username || 'User';
  await save();
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => { });
  if (S.user) { show('home'); trySync(); }
  else { show('login'); if (S.lastError) { $('lg-err').style.display = 'block'; $('lg-err').textContent = S.lastError; } }
  setInterval(() => { if (S.user && online() && (S.queue.length || S.listsDirty)) trySync(); }, 60000);
}
boot();
