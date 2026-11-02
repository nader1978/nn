/* =====================================================================
   CONFIG + MICROSOFT SIGN-IN (MSAL) — same pattern as SC Mngt system
   ===================================================================== */
/* Connection settings are SHARED with "NFPC SC Mngt Support System" (same keys, same web address = same settings):
   Tenant ID / Client ID / Site URL / token / light-dark / colour theme are read from and written to SC Mngt's storage. */
const SC_LS = 'sp_dashboard_secure_v12';
const SC_KEY = { ms: SC_LS + '_ms_cfg', global: SC_LS + '_global', theme: SC_LS + '_theme', accent: SC_LS + '_accent', mini: SC_LS + '_navmini' };
const scGet = k => { try { return JSON.parse(localStorage.getItem(k) || 'null') || {}; } catch (e) { return {}; } };
const scSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} };
function cfg() {
  const c = lsGet('cfg', {}), ms = scGet(SC_KEY.ms), gl = scGet(SC_KEY.global);
  const tenantId = ms.tenantId || c.tenantId || MS_AUTH_DEFAULT.tenantId, clientId = ms.clientId || c.clientId || MS_AUTH_DEFAULT.clientId;
  const scSite = (gl.siteUrl || ms.siteUrl || MS_AUTH_DEFAULT.siteUrl || '').replace(/\/+$/, ''), rabSite = (c.rabSiteUrl || '').replace(/\/+$/, '');
  const siteUrl = rabSite || scSite || (c.siteUrl || '').replace(/\/+$/, '');
  return { mode: c.mode || (siteUrl && ((tenantId && clientId) || gl.accessToken) ? 'sp' : 'demo'), tenantId, clientId, siteUrl, scSite, rabSite, listNames: c.listNames || {} };
}
function saveCfg(p) {
  const { tenantId, clientId, siteUrl, ...rest } = p;
  if (tenantId !== undefined || clientId !== undefined) { const ms = scGet(SC_KEY.ms); scSet(SC_KEY.ms, { ...ms, ...(tenantId !== undefined ? { tenantId } : {}), ...(clientId !== undefined ? { clientId } : {}) }); }
  if (siteUrl !== undefined) { const gl = scGet(SC_KEY.global); scSet(SC_KEY.global, { ...gl, siteUrl, updatedAt: Date.now(), updatedBy: (window.Auth && Auth.user && Auth.user.name) || 'RAB' }); }
  lsSet('cfg', { ...lsGet('cfg', {}), ...rest });
}
const listName = k => cfg().listNames[k] || LISTS[k].name;

const MS = {
  app: null, account: null, token: null, exp: 0,
  loadLib() {
    if (window.msal) return Promise.resolve();
    const urls = ['https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js', 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js'];
    return urls.reduce((p, u) => p.catch(() => new Promise((res, rej) => { const s = document.createElement('script'); s.src = u; s.onload = res; s.onerror = () => rej(new Error('Could not load the Microsoft sign-in library.')); document.head.appendChild(s); })), Promise.reject());
  },
  async init() {
    if (this.app) return this.app;
    const c = cfg(); if (!c.tenantId || !c.clientId) throw new Error('Enter the Tenant ID and Client ID in Settings first.');
    if (location.protocol === 'file:') throw new Error('Microsoft sign-in does not work from a local file. Open this page from https:// (e.g. a SharePoint document library) or http://localhost.');
    await this.loadLib();
    this.app = new msal.PublicClientApplication({ auth: { clientId: c.clientId, authority: 'https://login.microsoftonline.com/' + c.tenantId, redirectUri: location.origin + location.pathname }, cache: { cacheLocation: 'localStorage' } });
    if (this.app.initialize) await this.app.initialize();
    return this.app;
  },
  scopes: ['https://graph.microsoft.com/Sites.ReadWrite.All', 'https://graph.microsoft.com/User.Read'],
  manageScopes: ['https://graph.microsoft.com/Sites.Manage.All'], manage: false, mtoken: null, mexp: 0,
  async getManageToken(interactive = true) {   // only needed to create lists/columns (Settings page)
    if (this.mtoken && Date.now() < this.mexp - 120000) return this.mtoken;
    const app = await this.init(); const acct = this.account || app.getAllAccounts()[0]; let res;
    try { res = await app.acquireTokenSilent({ scopes: this.manageScopes, account: acct }); } catch (e) { if (!interactive) throw e; res = await app.acquireTokenPopup({ scopes: this.manageScopes, account: acct }); }
    this.mtoken = res.accessToken; this.mexp = +new Date(res.expiresOn || Date.now() + 3000000); return this.mtoken;
  },
  /* shared Bearer token (Settings → Global Access Token, or the sp_token.json file) — used when Microsoft sign-in is not set up */
  shared() { const t = scGet(SC_KEY.global).accessToken || ''; if (!t) return ''; try { const j = JSON.parse(atob(t.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'))); if (j.exp && j.exp * 1000 < Date.now() + 60000) return ''; } catch (e) {} return t; },
  async getToken(interactive = false) {
    { const c = cfg(); if (!c.tenantId || !c.clientId) { const t = this.shared(); if (t) return t; } }
    if (this.manage) return this.getManageToken(interactive);
    if (this.token && Date.now() < this.exp - 120000) return this.token;
    const app = await this.init(); let acct = this.account || app.getAllAccounts()[0]; let res;
    if (!acct) { if (!interactive) throw new Error('Not signed in to Microsoft.'); res = await app.loginPopup({ scopes: this.scopes }); acct = res.account; }
    try { res = await app.acquireTokenSilent({ scopes: this.scopes, account: acct }); }
    catch (e) { if (!interactive) throw e; res = await app.acquireTokenPopup({ scopes: this.scopes, account: acct }); }
    this.account = acct; this.token = res.accessToken; this.exp = +new Date(res.expiresOn || Date.now() + 3000000);
    try { const gl = scGet(SC_KEY.global); scSet(SC_KEY.global, { ...gl, siteUrl: gl.siteUrl || cfg().scSite || cfg().siteUrl, accessToken: res.accessToken, updatedAt: Date.now(), updatedBy: (acct && acct.username) || 'RAB' }); } catch (e) {}   // SC Mngt sees the same fresh token
    return this.token;
  },
  async signOut() { try { const app = await this.init(); const a = app.getAllAccounts()[0]; this.token = null; if (a) await app.logoutPopup({ account: a }); } catch (e) {} },
  user() { const a = this.account || (this.app && this.app.getAllAccounts()[0]); return a ? (a.name || a.username) : ''; }
};

/* Graph fetch with retry on throttling (429/503) */
async function gfetch(url, opt = {}, tries = 5) {
  const token = await MS.getToken(false);
  for (let i = 0; i < tries; i++) {
    const r = await fetch(url.startsWith('http') ? url : 'https://graph.microsoft.com/v1.0' + url, { ...opt, headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json', Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly', ...(opt.headers || {}) } });
    if (r.status === 429 || r.status === 503 || r.status === 504) { const ra = +(r.headers.get('Retry-After') || 0); await new Promise(s => setTimeout(s, (ra || 2 ** i) * 1000)); continue; }
    if (!r.ok) { let m = r.status + ' ' + r.statusText; try { const j = await r.json(); m = (j.error && j.error.message) || m; } catch (e) {} const er = new Error(m); er.status = r.status; throw er; }
    return r.status === 204 ? null : r.json();
  }
  throw new Error('SharePoint is busy (throttled). Try again in a minute.');
}
const SP = {
  siteIds: {}, siteLists: {}, listIds: {}, mapped: {},
  /* existing lists live on their own site (from the Access link); new lists go to the RAB/NFPC lists site */
  siteUrlFor(k) { const L = LISTS[k], c = cfg(); if (L && L.site) { try { return new URL(c.scSite || c.siteUrl).origin + '/' + L.site; } catch (e) { return ''; } } return c.siteUrl; },
  async siteIdFor(url) {
    if (!url) throw new Error('No SharePoint site URL — set it in Settings.');
    if (this.siteIds[url]) return this.siteIds[url];
    const u = new URL(url); const path = u.pathname.replace(/\/+$/, '');
    const j = await gfetch(`/sites/${u.hostname}${path && path !== '/' ? ':' + path : ''}`);
    return (this.siteIds[url] = j.id);
  },
  async site(k) { return this.siteIdFor(this.siteUrlFor(k)); },
  async findList(k) {
    if (this.listIds[k]) return this.listIds[k];
    const sid = await this.site(k), nm = listName(k);
    if (!this.siteLists[sid]) { const out = []; let url = `/sites/${sid}/lists?$select=id,displayName,name&$top=999`; while (url) { const j = await gfetch(url); out.push(...(j.value || [])); url = j['@odata.nextLink'] || null; } this.siteLists[sid] = out; }
    const hit = this.siteLists[sid].find(l => l.displayName === nm || l.name === nm || l.displayName.toLowerCase() === nm.toLowerCase());
    if (hit) this.listIds[k] = hit.id; return hit ? hit.id : null;
  },
  async listId(k) { const id = await this.findList(k); if (!id) throw new Error(`SharePoint list "${listName(k)}" does not exist yet — open Settings → SharePoint lists → Create.`); return id; },
  /* read the real columns of the list and map Access field names to SharePoint internal names */
  async mapColumns(k) {
    if (this.mapped[k]) return; const L = LISTS[k], sid = await this.site(k), lid = await this.listId(k);
    const j = await gfetch(`/sites/${sid}/lists/${lid}/columns?$top=999`); const cols = (j.value || []).filter(c => !c.hidden);
    const used = new Set();
    for (const f of L.fields) {
      const c = cols.find(c => norm(c.displayName) === norm(f.n)) || cols.find(c => norm(c.name) === norm(f.n)) || cols.find(c => norm(c.name) === norm(f.sp));
      if (!c) { f.missing = true; continue; }
      used.add(c.name); f.sp = c.name; f.missing = false; spType(f, c);
    }
    for (const c of cols) { if (used.has(c.name) || READONLY_SYS.has(c.name)) continue; const f = { n: c.displayName, sp: c.name, l: c.displayName, extra: true }; spType(f, c); L.fields.push(f); }
    L.byName = Object.fromEntries(L.fields.map(f => [f.n, f])); this.mapped[k] = true;
  },
  fromSP(k, it) {
    const L = LISTS[k], f = it.fields || {}, rec = { _id: String(it.id), _etag: f['@odata.etag'] };
    for (const d of L.fields) {
      if (d.missing) continue;
      let v = f[d.sp]; if ((v === undefined || v === null) && f[d.sp + 'LookupId'] !== undefined) v = f[d.sp + 'LookupId'];
      if (Array.isArray(v)) v = v.map(x => x && typeof x === 'object' ? (x.LookupValue || x.Email || JSON.stringify(x)) : x).join('; ');
      else if (v && typeof v === 'object') v = v.LookupValue || v.Email || v.Url || JSON.stringify(v);
      if (v === undefined || v === null) { rec[d.n] = d.t === 'b' ? false : null; continue; }
      if (d.t === 'd') { const x = new Date(v); v = isNaN(x) ? String(v).slice(0, 10) : `${x.getFullYear()}-${pad(x.getMonth() + 1)}-${pad(x.getDate())}`; }
      else if (d.t === 'n' || d.t === 'c' || d.t === 'p') v = num(v);
      else if (d.t === 'b') v = v === true || v === 'true' || v === 1 || v === 'Yes';
      rec[d.n] = v;
    }
    return rec;
  },
  toSP(k, rec) {
    const L = LISTS[k], o = {};
    if (!L.byName.Title || L.byName.Title.missing) o.Title = String(L.title(rec) || '').slice(0, 255);
    for (const d of L.fields) {
      if (d.missing || d.ro || !(d.n in rec)) continue; let v = rec[d.n];
      if (v === '' || v === undefined) v = null;
      if (v !== null) {
        if (d.t === 'd') v = String(v).slice(0, 10) + 'T12:00:00Z';
        else if (d.t === 'dt') { const x = new Date(v); v = isNaN(x) ? null : x.toISOString(); }
        else if (d.t === 'n' || d.t === 'c' || d.t === 'p') v = num(v);
        else if (d.t === 'b') v = !!v;
        else v = String(v);
      } else if (d.t === 'b') v = false;
      o[d.sp] = v;
    }
    return o;
  },
  async all(k) {
    if (LISTS[k].existing) await this.mapColumns(k);
    const sid = await this.site(k), lid = await this.listId(k), out = [];
    let url = `/sites/${sid}/lists/${lid}/items?$expand=fields&$top=999`;
    while (url) { const j = await gfetch(url); for (const it of j.value || []) out.push(this.fromSP(k, it)); url = j['@odata.nextLink'] || null; }
    return out;
  },
  async create(k, rec) { const sid = await this.site(k), lid = await this.listId(k); const j = await gfetch(`/sites/${sid}/lists/${lid}/items`, { method: 'POST', body: JSON.stringify({ fields: this.toSP(k, rec) }) }); return this.fromSP(k, j); },
  async update(k, rec) { const sid = await this.site(k), lid = await this.listId(k); await gfetch(`/sites/${sid}/lists/${lid}/items/${rec._id}/fields`, { method: 'PATCH', body: JSON.stringify(this.toSP(k, rec)) }); return rec; },
  async remove(k, rec) { const sid = await this.site(k), lid = await this.listId(k); await gfetch(`/sites/${sid}/lists/${lid}/items/${rec._id}`, { method: 'DELETE' }); },
  colDef(d) {
    const base = { name: d.sp, displayName: d.l.slice(0, 80), description: 'Access field: ' + d.n };
    switch (d.t) {
      case 'm': return { ...base, text: { allowMultipleLines: true, linesForEditing: 6, textType: 'plain' } };
      case 'n': case 'p': return { ...base, number: { decimalPlaces: 'automatic' } };
      case 'c': return { ...base, number: { decimalPlaces: 'two' } };
      case 'd': return { ...base, dateTime: { format: 'dateOnly' } };
      case 'dt': return { ...base, dateTime: { format: 'dateTime' } };
      case 'b': return { ...base, boolean: {} };
      case 'ch': return { ...base, choice: { allowTextEntry: true, choices: d.opts || [], displayAs: 'dropDownMenu' } };
      default: return { ...base, text: { maxLength: 255 } };
    }
  },
  async inspect(k) {
    const lid = await this.findList(k); if (!lid) return { exists: false, missing: LISTS[k].fields.filter(f => !f.extra) };
    if (LISTS[k].existing) { await this.mapColumns(k); return { exists: true, missing: [], existing: true, unmatched: LISTS[k].fields.filter(f => f.missing).length }; }
    const sid = await this.site(k); const j = await gfetch(`/sites/${sid}/lists/${lid}/columns?$select=name&$top=999`);
    const have = new Set((j.value || []).map(c => c.name.toLowerCase()));
    return { exists: true, missing: LISTS[k].fields.filter(d => !d.extra && !have.has(d.sp.toLowerCase())) };
  },
  async provision(k, onStep) {
    if (LISTS[k].existing) { const lid = await this.findList(k); if (!lid) throw new Error('Existing list not found on ' + this.siteUrlFor(k)); await this.mapColumns(k); return 0; }
    const sid = await this.site(k); let lid = await this.findList(k);
    if (!lid) { const j = await gfetch(`/sites/${sid}/lists`, { method: 'POST', body: JSON.stringify({ displayName: listName(k), description: LISTS[k].label + ' — created by ' + APP.name, list: { template: 'genericList' } }) }); lid = this.listIds[k] = j.id; (this.siteLists[sid] = this.siteLists[sid] || []).push({ id: j.id, displayName: listName(k), name: listName(k) }); onStep && onStep('List created'); }
    const { missing } = await this.inspect(k); let n = 0;
    for (const d of missing) { try { await gfetch(`/sites/${sid}/lists/${lid}/columns`, { method: 'POST', body: JSON.stringify(this.colDef(d)) }); } catch (e) { if (!/already exists|duplicate/i.test(e.message)) throw new Error(d.n + ': ' + e.message); } n++; onStep && onStep(`${n}/${missing.length} columns`); }
    try { await gfetch(`/sites/${sid}/lists/${lid}/columns/Title`, { method: 'PATCH', body: JSON.stringify({ required: false }) }); } catch (e) {}
    return missing.length;
  }
};
const READONLY_SYS = new Set('ID,ContentType,Modified,Created,Author,Editor,_UIVersionString,Attachments,Edit,LinkTitleNoMenu,LinkTitle,DocIcon,ItemChildCount,FolderChildCount,_ComplianceFlags,_ComplianceTag,_ComplianceTagWrittenTime,_ComplianceTagUserId,_IsRecord,AppAuthor,AppEditor,_ColorTag,ComplianceAssetId,_DisplayName'.split(','));
function spType(f, c) {
  if (c.readOnly || c.lookup || c.personOrGroup || c.calculated || c.hyperlinkOrPicture || c.thumbnail || c.geolocation || c.term) f.ro = !!(c.readOnly || c.calculated || c.lookup || c.personOrGroup || c.term);
  if (c.text) f.t = c.text.allowMultipleLines ? 'm' : (f.t === 'm' ? 't' : (['t', 'ch'].includes(f.t) ? f.t : 't'));
  else if (c.number) f.t = ['c', 'n', 'p'].includes(f.t) ? f.t : 'n';
  else if (c.currency) f.t = 'c';
  else if (c.dateTime) f.t = c.dateTime.format === 'dateOnly' ? 'd' : 'dt';
  else if (c.boolean) f.t = 'b';
  else if (c.choice) { f.t = 'ch'; f.opts = c.choice.choices || []; }
  else f.t = f.t || 't';
}
/* Demo backend: this browser only (localStorage) — lets you try every page before SharePoint is ready */
const DEMO = {
  key: k => 'demo_' + k,
  all(k) { return lsGet(this.key(k), null); },
  put(k, a) { lsSet(this.key(k), a); },
  seqNext() { const s = lsGet('demo_seq', 1000) + 1; lsSet('demo_seq', s); return String(s); },
};

/* =====================================================================
   DB — one API for both backends, with an in-memory cache per list
   ===================================================================== */
const DB = {
  cache: {}, loading: {},
  isSP() { return cfg().mode === 'sp'; },
  async all(k, force = false) {
    if (this.cache[k] && !force) return this.cache[k];
    if (this.loading[k] && !force) return this.loading[k];
    const p = (async () => {
      let rows;
      if (this.isSP()) rows = await SP.all(k);
      else { rows = DEMO.all(k); if (!rows) { rows = (SEED[k] ? SEED[k]() : []).map(r => ({ _id: DEMO.seqNext(), ...r })); DEMO.put(k, rows); } }
      this.cache[k] = rows; return rows;
    })();
    this.loading[k] = p; try { return await p; } finally { delete this.loading[k]; }
  },
  peek(k) { return this.cache[k] || []; },
  nextKey(k) { const pk = LISTS[k].pk; if (!pk) return null; return Math.max(0, ...this.peek(k).map(r => num(r[pk]))) + 1; },
  async save(k, rec) {
    const L = LISTS[k]; await this.all(k);
    if (L.pk && !num(rec[L.pk])) { if (this.isSP()) await this.all(k, true); rec[L.pk] = this.nextKey(k); }
    let out;
    if (this.isSP()) out = rec._id ? await SP.update(k, rec) : await SP.create(k, rec);
    else { const a = DEMO.all(k) || []; if (rec._id) { const i = a.findIndex(x => x._id === rec._id); a[i >= 0 ? i : a.length] = { ...rec }; } else { rec._id = DEMO.seqNext(); a.push({ ...rec }); } DEMO.put(k, a); out = { ...rec }; }
    const c = this.cache[k] || (this.cache[k] = []); const i = c.findIndex(x => x._id === out._id);
    const merged = { ...rec, ...out }; if (i >= 0) c[i] = merged; else c.push(merged);
    Bus.emit('data', k); return merged;
  },
  async remove(k, rec) {
    if (!rec || !rec._id) return;
    if (this.isSP()) await SP.remove(k, rec); else DEMO.put(k, (DEMO.all(k) || []).filter(x => x._id !== rec._id));
    if (this.cache[k]) this.cache[k] = this.cache[k].filter(x => x._id !== rec._id);
    Bus.emit('data', k);
  },
  /* run many saves/deletes with limited parallelism + progress */
  async bulk(jobs, onProg, par = 4) { let done = 0, i = 0, errs = []; const run = async () => { while (i < jobs.length) { const j = jobs[i++]; try { await j(); } catch (e) { errs.push(e.message); } done++; onProg && onProg(done, jobs.length); } }; await Promise.all(Array.from({ length: Math.min(par, jobs.length || 1) }, run)); return errs; },
  resetCache() { this.cache = {}; SP.siteIds = {}; SP.siteLists = {}; SP.listIds = {}; SP.mapped = {}; },
};
const Bus = { h: {}, on(e, f) { (this.h[e] = this.h[e] || []).push(f); }, emit(e, ...a) { (this.h[e] || []).forEach(f => { try { f(...a); } catch (x) { console.error(x); } }); } };

/* =====================================================================
   USERS & ROLES — the same user store, roles and password format as NFPC SC Mngt
   (localStorage "sp_dashboard_secure_v12_users"): on the same web address both systems share
   their users. Optional SharePoint login control (LoginUserList, settings shared with SC Mngt).
   ===================================================================== */
const LS_BASE = SC_LS;
const USERS_LIST = 'LoginUserList';
const ROLES = { viewer: { label: 'Viewer' }, operator: { label: 'Operator' }, admin: { label: 'Admin' } };
const EMAIL_RX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const uidGen = () => Math.random().toString(36).slice(2, 10);
async function hashPassword(password, salt) { let h = salt + ':' + password; for (let i = 0; i < 10000; i++) h = await sha256(h); return h; }
function randomSalt() { const a = new Uint8Array(16); crypto.getRandomValues(a); return [...a].map(b => b.toString(16).padStart(2, '0')).join(''); }
/* Built-in administrator of SC Mngt (salted hash only — the password itself is not in the file) */
const BUILTIN_USER = { username: 'nader', displayName: 'Nader Al Najjar', role: 'admin', salt: '38235b55aa4aa5abc05404b8afcdf9c4', hash: '78e0211c590b958a30c0d5d1bc1063a064a37e9ae76b64961e92b47c3c869c1a' };
const Users = {
  list: [],
  save() { try { localStorage.setItem(LS_BASE + '_users', JSON.stringify(this.list)); } catch (e) {} },
  find(n) { const u = String(n || '').trim().toLowerCase(); return this.list.find(x => String(x.username).toLowerCase() === u || (x.email && String(x.email).toLowerCase() === u)); },
  async load() {
    try { this.list = JSON.parse(localStorage.getItem(LS_BASE + '_users') || '[]') || []; } catch (e) { this.list = []; }
    const b = this.list.find(x => String(x.username).toLowerCase() === BUILTIN_USER.username);
    if (b) { b.builtin = true; b.role = 'admin'; if (!b.salt || !b.hash) { b.salt = BUILTIN_USER.salt; b.hash = BUILTIN_USER.hash; } }
    else this.list.unshift({ id: uidGen(), username: BUILTIN_USER.username, displayName: BUILTIN_USER.displayName, role: 'admin', salt: BUILTIN_USER.salt, hash: BUILTIN_USER.hash, builtin: true, mustChangePassword: false, allowedLists: [], editableLists: [], createdAt: Date.now(), lastLogin: null });
    if (!this.list.some(x => !x.builtin)) {   // first run on this web address: admin / admin (+ a demo operator in demo mode)
      const s = randomSalt(); this.list.push({ id: uidGen(), username: 'admin', displayName: 'Administrator', role: 'admin', salt: s, hash: await hashPassword('admin', s), mustChangePassword: !(cfg().mode === 'demo'), allowedLists: [], editableLists: [], createdAt: Date.now(), lastLogin: null });
      if (cfg().mode === 'demo') { const s2 = randomSalt(); this.list.push({ id: uidGen(), username: 'store', displayName: 'Imran Khan', role: 'operator', jobTitle: 'Storekeeper', email: 'store@example.com', salt: s2, hash: await hashPassword('1234', s2), allowedLists: [], editableLists: [], createdAt: Date.now(), lastLogin: null }); }
    }
    this.save();
  },
};
/* ---- login control: users checked against a SharePoint list (same settings key as SC Mngt) ---- */
const LUL_DEFAULT = { on: false, list: 'LoginUserList', allow: 'Allow_More_VMs', user: '', adm: true, pwOn: true, pw: 'Password', role: 'Role' };
const LUL_KEY = LS_BASE + '_lulcfg';
function lulCfg() { let c = {}; try { c = JSON.parse(localStorage.getItem(LUL_KEY) || '{}') || {}; } catch (e) {} const D = LUL_DEFAULT; return { on: c.on === undefined ? D.on : !!c.on, list: (c.list || D.list).trim(), allow: (c.allow || D.allow).trim(), user: (c.user || '').trim(), adm: c.adm === undefined ? D.adm : !!c.adm, pwOn: c.pwOn === undefined ? D.pwOn : !!c.pwOn, pw: (c.pw || D.pw).trim(), role: (c.role || D.role).trim() }; }
const lulNorm = x => String(x || '').toLowerCase().replace(/[^a-z0-9]/g, '');
const lulTrue = v => v === true || /^(true|yes|y|1|-1)$/i.test(String(v == null ? '' : (typeof v === 'object' ? v.Value : v)).trim());
function lulVals(v) { if (v == null) return []; if (Array.isArray(v)) return v.flatMap(lulVals); if (typeof v === 'object') return [v.Email, v.LookupValue, v.Value, v.Title, v.DisplayName, v.EMail].filter(Boolean).map(String); return [String(v)]; }
function lulKey(f, want) { if (want in f) return want; const n = lulNorm(want); return Object.keys(f).find(k => lulNorm(k) === n) || null; }
function lulUserKeys(f, col) { if (col) { const k = lulKey(f, col); return k ? [k] : []; } return Object.keys(f).filter(k => !/^(@|_|id$|attachments|content|edit|modified|created|author|editor|linktitle|appauthor|appeditor|itemchild|folderchild|compliance)/i.test(k)); }
/* raw SharePoint items of any list by name (existing NFPC lists keep their own site) */
async function spRawItems(name) {
  const k = Object.keys(LISTS).find(x => norm(x) === norm(name) || norm(LISTS[x].name) === norm(name));
  const sid = k ? await SP.site(k) : await SP.siteIdFor(cfg().scSite || cfg().siteUrl);
  if (!SP.siteLists[sid]) { const out = []; let url = `/sites/${sid}/lists?$select=id,displayName,name&$top=999`; while (url) { const j = await gfetch(url); out.push(...(j.value || [])); url = j['@odata.nextLink'] || null; } SP.siteLists[sid] = out; }
  const hit = SP.siteLists[sid].find(l => norm(l.displayName) === norm(name) || norm(l.name) === norm(name)); if (!hit) { const e = new Error(`List “${name}” not found`); e.status = 404; throw e; }
  const items = []; let url = `/sites/${sid}/lists/${hit.id}/items?expand=fields&$top=999`; while (url) { const j = await gfetch(url); items.push(...(j.value || [])); url = j['@odata.nextLink'] || null; }
  return { sid, lid: hit.id, items };
}
async function lulFetch() {
  if (!DB.isSP()) return { state: 'notoken', msg: 'Demo data mode' };
  try { await MS.getToken(false); } catch (e) { return { state: 'notoken', msg: e.message }; }
  try { return { state: 'ok', ...(await spRawItems(lulCfg().list)) }; }
  catch (e) { return { state: e.status === 404 ? 'nolist' : /\b(401|403)\b/.test(String(e.status)) ? 'notoken' : 'error', msg: e.message }; }
}
function lulIdentities(u) { const ids = new Set(); const add = v => { v = String(v || '').trim().toLowerCase(); if (v) ids.add(v); }; if (u) { add(u.email); add(u.username); add(u.displayName); } try { const a = MS.app && MS.app.getAllAccounts()[0]; if (a) { add(a.username); add(a.name); } } catch (e) {} return [...ids]; }
function lulJudge(items, u) {
  const c = lulCfg(), ids = lulIdentities(u), M = [];
  items.forEach(it => { const f = it.fields || {}, keys = lulUserKeys(f, c.user); let tier = -1; for (let i = 0; i < ids.length && tier < 0; i++) if (keys.some(k => lulVals(f[k]).some(v => String(v).trim().toLowerCase() === ids[i]))) tier = i;
    if (tier >= 0) { const ak = lulKey(f, c.allow); M.push({ f, tier, allow: ak ? lulTrue(f[ak]) : false, hasCol: !!ak, raw: ak ? lulVals(f[ak]).join('/') || '(empty)' : '(no column)' }); } });
  if (!M.length) return { found: false, allowed: false, ids };
  const t = Math.min(...M.map(h => h.tier)), hits = M.filter(h => h.tier === t), detail = `Matched on “${ids[t]}” — ${hits.length} row(s), ${c.allow} = ${hits.map(h => h.raw).join(' , ')}`;
  if (!hits.some(h => h.hasCol)) return { found: true, allowed: false, noCol: true, ids, detail };
  return { found: true, allowed: hits.some(h => h.allow), ids, detail, f: (hits.find(h => h.allow) || hits[0]).f };
}
async function lulLogin(username, password) {
  const c = lulCfg(), un = String(username || '').trim().toLowerCase(), bad = { state: 'bad', msg: 'Invalid username or password.' };
  const r = await lulFetch(); if (r.state !== 'ok') return { state: 'unreach', msg: r.msg || r.state };
  const rows = r.items.filter(it => { const f = it.fields || {}; return lulUserKeys(f, c.user).some(k => lulVals(f[k]).some(v => String(v).trim().toLowerCase() === un)); });
  if (!rows.length) return bad;
  let ok = false, off = false, role = '', email = '', name = '', title = '', loc = '', id = null;
  for (const it of rows) { const f = it.fields || {}, pk = lulKey(f, c.pw); if (!pk) continue; const stored = String(lulVals(f[pk])[0] || '').trim(); if (!stored) continue;
    const s2 = stored.replace(/^sha256:/i, ''), pass = /^[a-f0-9]{64}$/i.test(s2) ? (await sha256(password)).toLowerCase() === s2.toLowerCase() : stored === password; if (!pass) continue;
    const ak = lulKey(f, c.allow); if (!ak || !lulTrue(f[ak])) { off = true; continue; }
    ok = true; id = String(it.id); const rk = lulKey(f, c.role) || lulKey(f, 'Access_Level'); role = rk ? String(lulVals(f[rk])[0] || '').trim().toLowerCase() : '';
    const ek = Object.keys(f).find(k => /^e-?mail|email|mail$|email_id/i.test(k)); email = ek ? String(lulVals(f[ek])[0] || '') : '';
    const nk = Object.keys(f).find(k => /display ?name|fullname|^name$|^title$/i.test(k)); name = nk ? String(lulVals(f[nk])[0] || '') : '';
    const tk = Object.keys(f).find(k => /job.?title|designation/i.test(k)); title = tk ? String(lulVals(f[tk])[0] || '') : '';
    const lk = Object.keys(f).find(k => lulNorm(k) === 'locationid'); loc = lk ? String(lulVals(f[lk])[0] || '') : ''; break; }
  if (!ok) return off ? { state: 'off', msg: 'Your account is deactivated. Please contact the administrator.' } : bad;
  role = /admin|manager|super/.test(role) ? 'admin' : /view|read/.test(role) ? 'viewer' : 'operator';
  let u = Users.find(username); if (u && u.builtin) return { state: 'ok', user: u };
  if (!u) { u = { id: uidGen(), username: String(username).trim(), displayName: name || String(username).trim(), role, salt: '', hash: '', email, managerEmail: '', jobTitle: title, allowedLists: [], editableLists: [], createdAt: Date.now(), lastLogin: null, fromSP: true }; Users.list.push(u); }
  else { u.role = role; if (email && !u.email) u.email = email; if (name && !u.displayName) u.displayName = name; if (title) u.jobTitle = title; u.fromSP = true; }
  if (loc && u.locUser !== 1) u.locationId = loc; u.spItem = id; return { state: 'ok', user: u };
}
/* ---- per-user options (Users & Roles → 🎛 Options; absent = allowed) ---- */
const FEAT_LIST = [
  ['excel', '📊 Excel export', 'Download .xlsx files from any list or report'],
  ['email', '✉ Email', 'Send documents and lists by email'],
  ['print', '🖨 Print', 'Print records, lists and reports'],
  ['import', '⬆ Excel import', 'Import rows from Excel / CSV into a list'],
  ['delete', '🗑 Delete records', 'Delete records and lines'],
  ['post', '📌 Post / approve', 'Post stock movements, approve LPOs and repairs, post GL batches'],
  ['xlstyle', '🎨 Excel style settings', 'Change the Excel export style and the company name'],
];
const Auth = {
  user: null, rec: null,
  isAdmin() { const u = this.rec; return !!u && u.role === 'admin'; },
  feat(k) { const u = this.rec; return !u || u.role === 'admin' || !(u.nfpcFeatures && u.nfpcFeatures[k] === false); },
  block(k) { if (this.feat(k)) return false; const f = FEAT_LIST.find(x => x[0] === k); toast('🔒 ' + (f ? f[1].replace(/^\S+\s/, '') : k) + ' is switched off for your account. Contact an administrator.', 'warn', 5000); return true; },
  listVisible(k) { const u = this.rec; if (!u || u.role === 'admin') return true; const a = u.nfpcLists && u.nfpcLists.visible; return !a || !a.length || a.includes(k); },
  listEditable(k) { const u = this.rec; if (!u || u.role === 'admin') return true; if (u.role === 'viewer') return false; const L = u.nfpcLists || {}; const e = L.editable && L.editable.length ? L.editable : L.visible; return !e || !e.length || e.includes(k); },
  /* 'full' | 'view' | 'none' for a page */
  level(page) {
    const u = this.rec; if (!u) return 'none'; if (u.role === 'admin') return 'full';
    if (['home', 'help', 'settings', 'dash', 'reports'].includes(page)) return u.role === 'viewer' && page !== 'settings' ? 'view' : 'full';
    const P = PAGES.find(p => p.id === page); if (!P) return 'none';
    if (P.sec === 'Admin' || P.adminOnly) return 'none';
    if (P.gen && !this.listVisible(P.gen.table)) return 'none';
    const ov = u.nfpcSections || {};
    if (page in ov) return ov[page] === false ? 'none' : ov[page] === 'view' ? 'view' : (u.role === 'viewer' ? 'view' : 'full');
    if (typeof tpAllows === 'function' && !tpAllows(u, P)) return 'none';
    if (u.role === 'viewer') return 'view';
    if (P.gen && !this.listEditable(P.gen.table)) return 'view';
    return 'full';
  },
  can(page) { return this.level(page) !== 'none'; },
  ro(page) { return this.level(page || (Nav.cur && Nav.cur.id)) === 'view'; },
  session(u) { this.rec = u; this.user = { id: u.id, name: u.username, full: u.displayName || u.username, level: (u.builtin ? 'Built-in admin' : (ROLES[u.role] || ROLES.operator).label), role: u.role, super: !!u.builtin, email: u.email, jobTitle: u.jobTitle, locationId: u.locationId, _id: u.spItem || null }; sessionStorage.setItem(APP.ls + '_user', u.id); return this.user; },
  restore() { const id = sessionStorage.getItem(APP.ls + '_user'); const u = id && Users.list.find(x => x.id === id); return u ? this.session(u) : null; },
  async login(name, pw) {
    const c = lulCfg(), spOn = c.on && c.pwOn && DB.isSP();
    let u = Users.find(name);
    if (u && u.hash && (await hashPassword(pw, u.salt)) === u.hash) { /* local password */ }
    else if (spOn && !(u && u.builtin)) { const r = await lulLogin(name, pw); if (r.state !== 'ok') throw new Error(r.state === 'unreach' ? 'Could not reach the SharePoint login list: ' + r.msg : r.msg); u = r.user; }
    else throw new Error('Invalid username or password.');
    u.lastLogin = Date.now(); Users.save();
    if (u.spItem && DB.isSP()) { try { const rows = await DB.all(USERS_LIST); const e = rows.find(x => x._id === u.spItem); if (e) { e.Last_Login_TimeStamp = nowIso(); e.Current_Login_Status = 'Online'; DB._sys = 1; DB.save(USERS_LIST, e).catch(() => {}).finally(() => { DB._sys = 0; }); } } catch (e) {} }
    return this.session(u);
  },
  async logout() { const u = this.rec; if (u && u.spItem && DB.isSP()) { const e = DB.peek(USERS_LIST).find(x => x._id === u.spItem); if (e) { e.Current_Login_Status = 'Offline'; e.Last_LogOut_TimeStamp = nowIso(); DB._sys = 1; DB.save(USERS_LIST, e).catch(() => {}); } } this.user = this.rec = null; sessionStorage.removeItem(APP.ls + '_user'); },
};
const me = () => (Auth.user && (Auth.user.full || Auth.user.name)) || 'System';
