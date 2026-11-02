/* ---------------------------------------------------------------- SETTINGS & SHAREPOINT */
const ACCENTS = [['blue', 'Blue', 'linear-gradient(120deg,#12306b,#3b82f6)'], ['green', 'Green', 'linear-gradient(120deg,#14532d,#34b863)'], ['navy', 'Dark blue', 'linear-gradient(120deg,#071433,#1e40af)'], ['gray', 'Gray', 'linear-gradient(120deg,#1f2937,#8795a8)'], ['red', 'Dark red', 'linear-gradient(120deg,#5b0f1b,#d6455a)']];
const ACCENT_KEY = SC_KEY.accent;   // shared with NFPC SC Mngt — one colour choice for both systems
const accentGet = () => { try { return localStorage.getItem(ACCENT_KEY) || 'blue'; } catch (e) { return 'blue'; } };
function accentApply(a) { if (!a || a === 'blue') document.documentElement.removeAttribute('data-accent'); else document.documentElement.setAttribute('data-accent', a); $$('#setAccent button').forEach(b => b.classList.toggle('on', b.dataset.a === a)); }
const BRAND_ORIG = { name: 'NFPC Supply Chain System', tag: 'Warehouse · Distribution · Fleet · Finance · HR', logoLogin: '🔐', logoSide: '🏭', img: LOGO_FILES.icon, full: LOGO_FILES.full };
const brandGet = () => ({ ...BRAND_DEFAULT, ...Object.fromEntries(Object.entries(lsGet('brand', {})).filter(([, v]) => v)) });
function brandApply(b = brandGet()) {
  const name = b.name || BRAND_ORIG.name, img = b.img || (b.logo ? '' : BRAND_ORIG.img);
  $$('[data-brand="name"]').forEach(e => e.textContent = name);
  $$('[data-brand="tag"]').forEach(e => e.textContent = b.tag || BRAND_ORIG.tag);
  $$('[data-brand="logo"]').forEach(e => { if (img) { e.innerHTML = `<img alt="" src="${esc(img)}">`; e.classList.add('has-img'); } else { e.textContent = b.logo || (e.closest('.login-brand') ? BRAND_ORIG.logoLogin : BRAND_ORIG.logoSide); e.classList.remove('has-img'); } });
  $$('[data-brand="full"]').forEach(e => { const show = !b.img && !b.logo && BRAND_ORIG.full; e.style.display = show ? '' : 'none'; if (show) e.src = BRAND_ORIG.full; const lc = e.closest('.login-card'); if (lc) lc.classList.toggle('has-full', !!show); });
  document.title = name + ' — Secure';
}

page('settings', 'Configuration', '⚙', 'Settings', 'Global + personal settings', async v => {
  const c = cfg(), b = brandGet(), adm = Auth.isAdmin();
  v.innerHTML = `
  ${adm ? `<div class="card"><div class="card-header"><div class="card-title">🏷 System name &amp; logo <span class="badge role-admin">ADMIN ONLY</span></div></div>
    <div class="card-desc">Change the name and logo shown on the sign-in page, the side panel and the browser tab. Logo = a short text / emoji, or an uploaded picture. Saved on this browser; use <b>Copy default</b> and paste the result over <code>BRAND_DEFAULT</code> in the config (<code>js/00-config.js</code> or the top of the HTML file) to give it to every user. Documents still print the company from <i>Company Info</i>.</div>
    <div class="grid three"><div class="form-group"><label>System name</label><input type="text" id="brName" maxlength="60" value="${esc(b.name)}" placeholder="${esc(BRAND_ORIG.name)}"></div><div class="form-group"><label>Tagline (sign-in page)</label><input type="text" id="brTag" maxlength="60" value="${esc(b.tag)}" placeholder="${esc(BRAND_ORIG.tag)}"></div><div class="form-group"><label>Logo text / emoji</label><input type="text" id="brLogo" maxlength="4" value="${esc(b.logo)}" placeholder="🏭"></div></div>
    <div class="btn-row" style="margin-top:12px"><button class="secondary small" id="brFile">🖼 Upload logo picture…</button><button class="ghost small" id="brNoImg">✖ Remove picture</button><input type="file" id="brPick" accept="image/*" hidden><span class="spacer"></span><button id="brSave">💾 Save</button><button class="secondary" id="brCopy">📋 Copy default</button><button class="ghost" id="brReset">↺ Reset to original</button></div><div id="brMsg"></div></div>

  <div class="card"><div class="card-header"><div class="card-title">🌐 Global Connection Settings <span class="badge role-admin">ADMIN ONLY</span></div></div>
    <div class="card-desc">These values are <b>shared with NFPC SC Mngt Support System</b>: when both pages are opened from the same web address (e.g. <code>${esc(location.origin)}</code>), the Tenant ID, Client ID, Site URL, Microsoft sign-in and light/dark mode set in one are used by the other — configure once.</div>
    <div class="grid two"><div class="form-group"><label>Global SharePoint Site URL <span class="hint">— same as SC Mngt</span></label><input type="text" id="sSite" value="${esc(c.scSite)}" placeholder="https://tenant.sharepoint.com/sites/YourSite"></div>
      <div class="form-group"><label>New lists site <span class="hint">— where the lists created from the Access back-ends live (blank = global site). Existing lists keep their own site, shown in the table below.</span></label><input type="text" id="sRabSite" value="${esc(c.rabSite)}" placeholder="(use the global site)"></div>
      <div class="form-group"><label>Data mode</label><select id="sMode2"><option value="sp" ${c.mode === 'sp' ? 'selected' : ''}>SharePoint (Microsoft 365)</option><option value="demo" ${c.mode === 'demo' ? 'selected' : ''}>Demo data (this browser only)</option></select></div></div>
    <hr style="border:none;border-top:1px solid var(--border);margin:14px 0">
    <div class="card-title" style="font-size:13px;margin-bottom:6px">🔐 Microsoft sign-in <span class="hint" style="font-weight:400">— no pasted token; each user signs in with their own account</span></div>
    <div class="card-desc">Register the app once in Microsoft Entra → App registrations → <em>Single-page application</em> redirect URI = the address this page is opened from: <code>${esc(location.origin + location.pathname)}</code> (add it next to the SC Mngt address in the <b>same</b> app registration). API permissions: Microsoft Graph <code>Sites.ReadWrite.All</code> + <code>User.Read</code> (+ <code>Sites.Manage.All</code> to create the lists), then grant admin consent.</div>
    <div class="grid two"><div class="form-group"><label>Tenant ID</label><input type="text" id="sTen" value="${esc(c.tenantId)}" placeholder="00000000-0000-0000-0000-000000000000"></div><div class="form-group"><label>Client (Application) ID</label><input type="text" id="sCli" value="${esc(c.clientId)}" placeholder="00000000-0000-0000-0000-000000000000"></div></div>
    <div class="btn-row" style="margin-top:12px"><button class="secondary" id="sSave">💾 Save IDs</button><button id="sSign">🔐 Sign in with Microsoft</button><button class="ghost" id="sOut">Sign out</button><span class="muted" id="sWho"></span></div><div id="sMsg"></div>
    <div class="muted" style="margin-top:8px">${(() => { const g = scGet(SC_KEY.global); return g.updatedAt ? `Last changed ${fdt(new Date(g.updatedAt).toISOString())} by ${esc(g.updatedBy || '')}` : ''; })()}</div></div>

  <div class="card"><div class="card-header"><div class="card-title">🗂 SharePoint lists <span class="hint" style="font-weight:400">— one per Access table</span> <span class="badge role-admin">ADMIN ONLY</span></div><div class="btn-row"><button class="secondary small" id="sChk">🔍 Check lists</button><button class="success small" id="sProv">🛠 Create / repair all lists</button></div></div>
    <div class="card-desc"><b>${Object.values(LISTS).filter(L => L.existing).length} lists already exist</b> on <code>${esc(Object.values(LISTS).find(L => L.existing)?.site || '')}</code> (the ones the SC Mngt system uses) — they are read and written as they are, columns are matched by name and never changed. The other ${Object.values(LISTS).filter(L => !L.existing).length} lists come from the Access back-end tables and are created on the <i>New lists site</i>. Safe to run again — existing data is never touched.</div>
    <div class="tbl-wrap" style="max-height:none"><table class="dt"><thead><tr><th>Access table</th><th>SharePoint list name</th><th>Site</th><th>Columns</th><th>Status</th></tr></thead><tbody>${Object.entries(LISTS).sort((a, b) => (b[1].existing - a[1].existing) || a[0].localeCompare(b[0])).map(([k, L]) => `<tr><td>${esc(k)} ${L.existing ? '<span class="badge ok">existing list</span>' : '<span class="badge info">new</span>'}<div class="muted">${esc(NFPC.tables[k].group || '')}</div></td><td><input type="text" data-ln="${k}" value="${esc(listName(k))}" style="max-width:280px"></td><td class="muted">${esc(L.existing ? L.site : 'new lists site')}</td><td class="num">${L.fields.length}</td><td data-ls="${k}"><span class="muted">—</span></td></tr>`).join('')}</tbody></table></div>
    <div class="progress" id="sProg" style="display:none"><div></div></div></div>

  <div class="card"><div class="card-header"><div class="card-title">📥 Move your Access data <span class="badge role-admin">ADMIN ONLY</span></div></div><div class="card-desc">In Access: right-click each table → Export → Excel. Then import the files here (the file/sheet name picks the list; columns match by name). Import header tables (IngdMvmntHdr, LPOHeader, PL Header, GLBatchHeader…) before their line tables — the link numbers (headerID, PINVID, Relation…) are kept.</div>
    <div class="btn-row"><button id="sImp">⬆ Import Excel / CSV files</button><button class="secondary" id="sBak">⬇ Backup everything to Excel</button>${c.mode === 'demo' ? '<button class="ghost" id="sReset">↺ Reset demo data</button>' : ''}</div></div>` : `<div class="status-msg info" style="margin:0 0 18px">Connection settings are managed by an administrator.</div>`}

  <div class="card"><div class="card-header"><div class="card-title">⚙ Personal Settings</div></div>
    <div class="grid"><div class="form-group spanall"><label>Colour theme <span class="hint">— headers, buttons and highlights across the whole system — shared with SC Mngt</span></label><div id="setAccent">${ACCENTS.map(t => `<button type="button" data-a="${t[0]}"><i style="background:${t[2]}"></i>${t[1]}</button>`).join('')}</div></div>
      <div class="form-group"><label>Light / dark</label><div class="muted" style="padding-top:6px">Follows the 🌙 button in the top bar.</div></div></div>
    <div class="btn-row" style="margin-top:14px">${Auth.user && Auth.user._id ? '<button class="secondary" id="chgPw">🔑 Change My Password</button>' : ''}<button class="danger" id="clrMine">🗑 Clear My Saved Data</button></div></div>`;
  accentApply(accentGet());
  $('#setAccent', v).onclick = e => { const x = e.target.closest('[data-a]'); if (!x) return; try { localStorage.setItem(ACCENT_KEY, x.dataset.a); } catch (er) {} accentApply(x.dataset.a); };
  $('#chgPw', v) && ($('#chgPw', v).onclick = changeMyPassword);
  $('#clrMine', v).onclick = async () => { if (!(await confirmBox('Clear your saved filters, layout and activity log on this browser? Data in SharePoint is not touched.', 'Clear', true))) return; Object.keys(localStorage).filter(k => k.startsWith(APP.ls + '_') && !k.startsWith(APP.ls + '_demo_') && !k.endsWith('_cfg') && !k.endsWith('_brand')).forEach(k => localStorage.removeItem(k)); toast('Cleared', 'success'); };
  if (!adm) return;
  /* brand */
  let img = b.img || '';
  const bmsg = (t, cls = 'success') => $('#brMsg', v).innerHTML = `<div class="status-msg ${cls}">${t}</div>`;
  const draft = () => ({ name: $('#brName', v).value.trim(), tag: $('#brTag', v).value.trim(), logo: $('#brLogo', v).value.trim(), img });
  ['#brName', '#brTag', '#brLogo'].forEach(s => $(s, v).oninput = () => brandApply({ ...BRAND_DEFAULT, ...draft() }));
  $('#brFile', v).onclick = () => $('#brPick', v).click();
  $('#brPick', v).onchange = e => { const f = e.target.files[0]; if (!f) return; const rd = new FileReader(); rd.onload = () => { const im = new Image(); im.onload = () => { const cv = document.createElement('canvas'), s = Math.min(1, 160 / Math.max(im.width, im.height)); cv.width = Math.round(im.width * s); cv.height = Math.round(im.height * s); cv.getContext('2d').drawImage(im, 0, 0, cv.width, cv.height); img = cv.toDataURL('image/png'); brandApply({ ...BRAND_DEFAULT, ...draft() }); }; im.src = rd.result; }; rd.readAsDataURL(f); };
  $('#brNoImg', v).onclick = () => { img = ''; brandApply({ ...BRAND_DEFAULT, ...draft() }); };
  $('#brSave', v).onclick = () => { lsSet('brand', draft()); brandApply(); bmsg('✔ Saved on this browser.'); };
  $('#brCopy', v).onclick = async () => { const t = 'const BRAND_DEFAULT = ' + JSON.stringify(draft()) + ';'; try { await navigator.clipboard.writeText(t); bmsg('✔ Copied — paste it over the BRAND_DEFAULT line in the config (js/00-config.js, or the top of the HTML file).'); } catch (e) { prompt('Copy this line:', t); } };
  $('#brReset', v).onclick = () => { lsSet('brand', {}); img = ''; ['#brName', '#brTag', '#brLogo'].forEach(s => $(s, v).value = ''); brandApply(); bmsg('↺ Back to the original name and logo.'); };
  /* connection */
  const msg = (t, cls = 'info') => $('#sMsg', v).innerHTML = `<div class="status-msg ${cls}">${t}</div>`;
  const store = () => { const ln = {}; $$('[data-ln]', v).forEach(i => { if (i.value.trim() && i.value.trim() !== LISTS[i.dataset.ln].name) ln[i.dataset.ln] = i.value.trim(); }); saveCfg({ mode: $('#sMode2', v).value, siteUrl: $('#sSite', v).value.trim().replace(/\/+$/, ''), rabSiteUrl: $('#sRabSite', v).value.trim(), tenantId: $('#sTen', v).value.trim(), clientId: $('#sCli', v).value.trim(), listNames: ln }); MS.app = null; MS.token = null; DB.resetCache(); updateConn(); };
  $('#sSave', v).onclick = () => { store(); msg('✅ Saved on this browser (also used by SC Mngt here). To give the IDs to every user, also put them into MS_AUTH_DEFAULT in the config (js/00-config.js or the top of the HTML file).', 'success'); };
  $('#sSign', v).onclick = async () => { store(); if (cfg().mode !== 'sp') return msg('Switch Data mode to SharePoint first.', 'warn'); try { msg('Signing in…'); await MS.getToken(true); await SP.site(); $('#sWho', v).textContent = '✔ ' + MS.user(); msg(`✅ Signed in as ${esc(MS.user())} and connected to the site. The token renews automatically.`, 'success'); updateConn(); } catch (e) { msg('✘ ' + esc(e.message), 'error'); } };
  $('#sOut', v).onclick = async () => { await MS.signOut(); updateConn(); msg('Signed out.', 'warn'); };
  const status = (k, html) => { const td = $(`[data-ls="${k}"]`, v); if (td) td.innerHTML = html; };
  $('#sChk', v).onclick = async () => { store(); if (cfg().mode !== 'sp') return msg('Switch Data mode to SharePoint first.', 'warn'); try { await MS.getToken(true); for (const k of Object.keys(LISTS)) { status(k, '…'); try { const r = await SP.inspect(k); status(k, !r.exists ? '<span class="badge err">missing</span>' : r.missing.length ? `<span class="badge warn">${r.missing.length} column(s) missing</span>` : '<span class="badge ok">ready</span>'); } catch (e) { status(k, `<span class="badge err">${esc(e.message)}</span>`); } } } catch (e) { msg('✘ ' + esc(e.message), 'error'); } };
  $('#sProv', v).onclick = async () => { store(); if (cfg().mode !== 'sp') return msg('Switch Data mode to SharePoint first.', 'warn');
    try { await MS.getToken(true); await MS.getManageToken(true); } catch (e) { return msg('✘ ' + esc(e.message) + ' — the app needs the Sites.Manage.All permission.', 'error'); }
    MS.manage = true; const keys = Object.keys(LISTS); $('#sProg', v).style.display = '';
    try { for (let i = 0; i < keys.length; i++) { const k = keys[i]; status(k, 'working…'); try { const n = await SP.provision(k, s => status(k, esc(s))); status(k, `<span class="badge ok">ready${n ? ` (+${n})` : ''}</span>`); } catch (e) { status(k, `<span class="badge err">${esc(e.message)}</span>`); } $('#sProg div', v).style.width = ((i + 1) / keys.length * 100) + '%'; } msg('✔ Lists are ready. Import your Access data below.', 'success'); Activity.add('🛠', 'SharePoint lists created / repaired'); }
    finally { MS.manage = false; DB.resetCache(); } };
  $('#sImp', v).onclick = () => importDialog(null, () => { DB.resetCache(); Activity.add('⬆', 'Data imported from Excel'); });
  $('#sBak', v).onclick = async () => { try { const wb = XLSX.utils.book_new(); for (const k of Object.keys(LISTS)) { const rows = await DB.all(k); const ws = XLSX.utils.json_to_sheet(rows.map(r => Object.fromEntries(LISTS[k].fields.map(f => [f.n, r[f.n] ?? ''])))); XLSX.utils.book_append_sheet(wb, ws, LISTS[k].name.slice(0, 31)); } XLSX.writeFile(wb, `NFPC_backup_${today()}.xlsx`); } catch (e) { toast(e.message, 'error'); } };
  $('#sReset', v) && ($('#sReset', v).onclick = async () => { if (!(await confirmBox('Reset all demo data in this browser?', 'Reset', true))) return; Object.keys(LISTS).forEach(k => { try { localStorage.removeItem(APP.ls + '_demo_' + k); } catch (e) {} }); DB.resetCache(); toast('Demo data reset', 'success'); go('home'); });
}, 'Settings window, ConnectSPWindow, Admin_Control Window');
function changeMyPassword() {
  const body = Modal.open('Change my password', `<div class="grid two"><div class="form-group"><label>Current password</label><input type="password" id="pw0" autocomplete="current-password"></div><div></div><div class="form-group"><label>New password <span class="hint">min 6 characters</span></label><input type="password" id="pw1" autocomplete="new-password"></div><div class="form-group"><label>Repeat new password</label><input type="password" id="pw2" autocomplete="new-password"></div></div>`, [{ label: 'Cancel', cls: 'ghost' }, { label: '🔑 Change', onClick: async () => {
    const e = DB.peek(USERS_LIST).find(x => x._id === Auth.user._id); if (!e) return toast('User record not found.', 'error');
    const old = $('#pw0', body).value, n1 = $('#pw1', body).value, h0 = await sha256(old);
    const pw = String(e.Password || ''); if (!(pw === 'sha256:' + h0 || pw.toLowerCase() === h0 || pw === old)) return toast('Current password is wrong.', 'warn');
    if (n1.length < 6) return toast('New password: at least 6 characters.', 'warn'); if (n1 !== $('#pw2', body).value) return toast('The two new passwords differ.', 'warn');
    try { await DB.save(USERS_LIST, { ...e, Password: /^(sha256:)?[0-9a-f]{64}$/i.test(pw) ? 'sha256:' + await sha256(n1) : n1 }); /* same format the list already uses (SC Mngt reads it too) */ Modal.close(); toast('Password changed ✔', 'success'); Activity.add('🔑', 'Password changed'); } catch (x) { toast(x.message, 'error'); }
  } }], 'narrow');
}

/* ---------------------------------------------------------------- HELP / FORM MAP */
page('help', 'Configuration', '❔', 'Help & Form Map', 'Which Access form became which page, and which SharePoint list it uses', async v => {
  const rows = PAGES.filter(p => p.form && !p.hidden);
  v.innerHTML = `<div class="card"><div class="toolbar"><input type="search" placeholder="Search form, page or list…" data-q><span class="spacer" style="flex:1"></span><span class="muted">${rows.length} pages · ${Object.keys(LISTS).length} lists · ${NFPC.reports.length} reports</span></div><div data-g></div></div>
    <div class="card"><div class="card-title" style="margin-bottom:10px">Smart features</div><ul style="font-size:13px;line-height:1.9;margin-left:18px">
      <li>Every Access form is a page: the same fields on the <b>Main</b> tab, every other column of the list on <b>All fields</b>, and each subform as an editable line grid under the record.</li>
      <li><b>FEFO stock control</b> for RM &amp; PM: posting a receiving document creates LOTs (with expiry dates); posting a releasing document takes the quantity from the LOTs that expire first and updates the item stock. Line grids show the open LOT balance and the next FEFO expiry while you type.</li>
      <li>Line amounts = quantity × unit price (− discount %), header totals = sum of the lines; document numbers (RCV-/REL-/LPO-/PL-/CPO-YYMM-0001) and created-by / time-stamp fields are filled automatically.</li>
      <li>Approve / send / cancel LPOs, approve and post vehicle repairs, post inspections and GL batches (debit = credit check) — buttons appear in the record window when the action is allowed.</li>
      <li>Alerts for low stock, expiring and expired LOTs, LPOs waiting for approval, vehicle insurance / registration / service dates and staff document expiry.</li>
      <li>The 8 existing SharePoint lists used by SC Mngt (LoginUserList, Inventory_Item_List, Main_Inventory_Movements_tbl…) are read and written as they are; the other lists come from the Access back-ends.</li>
      <li>Type-ahead suggestions from your own data on every text field (like the Access “SELECT DISTINCT” combo boxes); every list exports to Excel and imports from Excel / CSV; every record and report prints as a branded A4 document.</li>
      <li>Shortcuts: <span class="kbd">Ctrl K</span> search everything · <span class="kbd">Esc</span> close a window.</li></ul></div>`;
  const cols = [{ l: 'Page', html: p => `<a href="#${p.id}">${p.icon} ${esc(p.label)}</a>` }, { f: 'sec', l: 'Group' }, { f: 'form', l: 'Access form / report' }, { l: 'SharePoint list', get: p => p.gen ? LISTS[p.gen.table].name + (LISTS[p.gen.table].existing ? ' (existing)' : '') : '—' }, { l: 'Subforms', get: p => p.gen ? p.gen.subs.map(s => s.table).join(', ') : '' }];
  const g = makeGrid($('[data-g]', v), { cols, rows, pageSize: 400 });
  $('[data-q]', v).oninput = debounce(e => g.redraw(textFilter(rows, $('[data-q]', v).value, ['label', 'form', 'sec', p => p.gen ? p.gen.table : ''])), 150);
}, 'About');

/* =====================================================================
   GLOBAL SEARCH (Ctrl+K) — pages + every loaded list
   ===================================================================== */
function openSearch() {
  const body = Modal.open('Search everything', `<input type="search" id="gq" placeholder="Page, item, LPO nr, vehicle plate, employee, customer…" autofocus><div class="muted" style="margin-top:6px">Searches the pages and the lists already opened in this session (${Object.keys(DB.cache).length} loaded).</div><div class="search-results" id="gr" style="margin-top:10px"></div>`, [], 'wide');
  const q = $('#gq', body); let hits = [], hl = 0;
  const run = () => {
    const s = q.value.trim(); if (s.length < 2) { $('#gr', body).innerHTML = '<div class="empty">Type at least 2 characters.</div>'; hits = []; return; }
    hits = PAGES.filter(p => !p.hidden && Auth.can(p.id) && (p.label + ' ' + (p.form || '')).toLowerCase().includes(s.toLowerCase())).slice(0, 8).map(p => ({ i: p.icon, t: p.label, s: p.sec + ' · page', go: () => go(p.id) }));
    for (const [k, rows] of Object.entries(DB.cache)) { if (!LISTS[k] || !rows.length) continue; const P = pageFor(k); if (!P || !Auth.can(P.id)) continue;
      hits.push(...textFilter(rows, s, LISTS[k].fields.filter(f => f.t === 't').map(f => f.n)).slice(0, 6).map(r => ({ i: P.icon, t: LISTS[k].title(r), s: `${LISTS[k].name} · ${P.label}`, go: () => go(P.id, 'rec:' + r._id) }))); if (hits.length > 80) break; }
    hl = 0; $('#gr', body).innerHTML = hits.length ? hits.map((h, i) => `<div class="sr-item ${i === hl ? 'hl' : ''}" data-i="${i}"><span style="font-size:18px">${h.i}</span><div><div><b>${esc(h.t)}</b></div><div class="muted">${esc(h.s)}</div></div></div>`).join('') : '<div class="empty">Nothing found.</div>';
    $$('.sr-item', body).forEach(e => e.onclick = () => { Modal.close(); hits[+e.dataset.i].go(); });
  };
  q.oninput = debounce(run, 150); run();
  q.onkeydown = e => { const items = $$('.sr-item', body); if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); hl = Math.max(0, Math.min(items.length - 1, hl + (e.key === 'ArrowDown' ? 1 : -1))); items.forEach((x, i) => x.classList.toggle('hl', i === hl)); items[hl]?.scrollIntoView({ block: 'nearest' }); } if (e.key === 'Enter' && hits[hl]) { Modal.close(); hits[hl].go(); } };
  setTimeout(() => q.focus(), 50);
}

/* ---------------------------------------------------------------- SWITCHBOARD (main page, as in SC Mngt) */
const VISITS_KEY = () => 'visits_' + ((Auth.user && Auth.user.name) || 'anon');
function trackVisit(id) { if (!id || id === 'home') return; const h = lsGet(VISITS_KEY(), {}), d = today(); (h[d] = h[d] || {})[id] = (h[d][id] || 0) + 1; Object.keys(h).sort().slice(0, -14).forEach(x => delete h[x]); lsSet(VISITS_KEY(), h); }
page('home', 'Main', '🏠', 'Switchboard', 'Main page — pick a group to open its pages', async v => {
  const days = [...Array(7)].map((_, i) => addDays(today(), i - 6)), h = lsGet(VISITS_KEY(), {});
  const secs = []; PAGES.forEach(p => { if (!p.hidden && Auth.can(p.id) && !secs.includes(p.sec)) secs.push(p.sec); });
  const groups = secs.map(s => { const items = PAGES.filter(p => p.sec === s && !p.hidden && Auth.can(p.id)); const per = days.map(d => items.reduce((a, p) => a + ((h[d] || {})[p.id] || 0), 0)); return { name: s, items, per, total: per.reduce((a, b) => a + b, 0) }; });
  const use = {}; days.forEach(d => Object.entries(h[d] || {}).forEach(([k, n]) => use[k] = (use[k] || 0) + n));
  const top = Object.entries(use).sort((a, b) => b[1] - a[1])[0], topP = top && PAGES.find(p => p.id === top[0]);
  const hr = new Date().getHours(), hello = (hr < 12 ? 'Good morning' : hr < 18 ? 'Good afternoon' : 'Good evening') + ', ' + esc((Auth.user.full || Auth.user.name || '').split(' ')[0]);
  const COL = ['#2563eb', '#0e7490', '#7c3aed', '#b45309', '#047857', '#be185d', '#4338ca', '#0f766e'];
  const bars = (g, c) => { const mx = Math.max(1, ...g.per); return `<div class="hm-bars" style="height:44px" title="${g.total ? 'Your visits, last 7 days' : 'No visits recorded yet'}">${g.per.map(x => `<i style="height:${Math.max(3, Math.round(x / mx * 44))}px;background:${c};opacity:${x ? .9 : .25}"></i>`).join('')}</div>`; };
  const kpis = [['Pages you can open', groups.reduce((a, g) => a + g.items.length, 0), ''], ['Your visits today', ((h[today()] || {}) && Object.values(h[today()] || {}).reduce((a, b) => a + b, 0)), ''], ['Most used (7 days)', topP ? topP.label : '—', topP ? topP.id : '']];
  v.innerHTML = `<div class="hm-ban"><div class="hm-ban-l">${BRAND_ORIG.full && !brandGet().img && !brandGet().logo ? `<div class="hm-logo"><img alt="" src="${esc(BRAND_ORIG.full)}"></div>` : ''}<div><div class="hm-eye">${esc(company().CompanyFullName || 'National Food Products Company')}</div><h1>${hello}</h1><p>Pick a group to open its pages. The bars show your own activity over the last 7 days.</p></div></div>
    <label class="hm-search"><span>🔍</span><input id="hmQ" placeholder="Search any page…" aria-label="Search pages"></label></div>
    <div class="hm-wrap"><div class="hm-kpis" id="hmK">${kpis.map(k => `<div class="hm-kpi" ${k[2] ? `data-go="${k[2]}"` : ''}><span>${k[0]}</span><b>${esc(k[1])}</b></div>`).join('')}</div>
    <div class="hm-grid">${groups.map((g, i) => `<section class="hm-card" data-name="${esc((g.name + ' ' + g.items.map(p => p.label).join(' ')).toLowerCase())}"><header data-go="${g.items[0].id}"><span class="hm-ico" style="background:${COL[i % 8]}1f;color:${COL[i % 8]}">${g.items[0].icon}</span><b>${esc(g.name)}</b><em>${g.items.length}</em></header>${bars(g, COL[i % 8])}<div class="hm-chips">${g.items.map(p => `<a class="hm-chip" data-go="${p.id}" data-pn="${esc(p.label.toLowerCase())}"><span>${p.icon}</span>${esc(p.label)}</a>`).join('')}</div></section>`).join('')}</div></div>
    ${groups.some(g => g.total) ? '' : '<p class="hm-note">Charts fill in as you open pages — they show your own activity on this browser.</p>'}`;
  v.onclick = e => { const g = e.target.closest('[data-go]'); if (g) { e.preventDefault(); go(g.dataset.go); } };
  $('#hmQ', v).oninput = e => { const q = e.target.value.trim().toLowerCase(); $$('.hm-card', v).forEach(c => { c.style.display = !q || c.dataset.name.includes(q) ? '' : 'none'; $$('.hm-chip', c).forEach(x => x.style.outline = q && x.dataset.pn.includes(q) ? '2px solid var(--acc-2)' : ''); }); };
  $('#hmQ', v).onkeydown = e => { if (e.key === 'Enter') { const x = $$('.hm-chip', v).find(c => c.dataset.pn.includes(e.target.value.trim().toLowerCase()) && e.target.value.trim()); if (x) go(x.dataset.go); } };
  /* live business figures under the KPIs (load in the background) */
  loadAll(['ItemTable', 'LOTTable', 'LPOHeader', 'Vehicles_tbl']).then(() => { if (!v.isConnected) return; const lots = DB.peek('LOTTable').filter(l => num(l.BalanceQty) > 0 && !l.CloseLot && l.ExpiryDate && daysBetween(today(), l.ExpiryDate) <= 30);
    const add = [['Items below minimum', stockRows().filter(r => r._status !== 'OK').length, 'stock'], ['LOTs expiring ≤ 30 d', lots.length, 'fefo'], ['LPOs to approve', DB.peek('LPOHeader').filter(h => !h.ApproveLPO && !h.CancelPO).length, 'dash'], ['Fleet alerts ≤ 30 d', fleetAlerts(30).length, 'fleetc']].filter(x => Auth.can(x[2]));
    $('#hmK', v)?.insertAdjacentHTML('afterbegin', add.map(k => `<div class="hm-kpi" data-go="${k[2]}"><span>${k[0]}</span><b>${esc(k[1])}</b></div>`).join('')); }).catch(() => {});
});
PAGES.unshift(PAGES.pop()); PAGES[0].hidden = true;
for (const p of PAGES) if (['settings', 'help'].includes(p.id)) p.sec = 'Configuration';
for (const id of ['settings', 'help']) { const i = PAGES.findIndex(p => p.id === id); PAGES.push(PAGES.splice(i, 1)[0]); }

/* ---------------------------------------------------------------- RECENT ACTIVITY (drawer, as in SC Mngt) */
const Activity = {
  list() { return lsGet('activity', []); },
  add(icon, text) { const a = this.list(); a.unshift({ t: nowIso(), i: icon, x: text, u: me() }); lsSet('activity', a.slice(0, 200)); lsSet('activityUnseen', lsGet('activityUnseen', 0) + 1); this.badge(); if ($('#activityDrawer').classList.contains('open')) this.render(); },
  badge() { const n = lsGet('activityUnseen', 0), b = $('#activityToggle'); b.dataset.count = n > 99 ? '99+' : n; b.classList.toggle('has-badge', n > 0); },
  render() { const a = this.list(); $('#activityBody').innerHTML = a.length ? a.map(x => `<div class="activity-item"><span class="ai">${x.i}</span><div>${esc(x.x)}<small>${fdt(x.t)} · ${esc(x.u || '')}</small></div></div>`).join('') : '<div class="activity-empty">No activity yet.</div>'; },
  open(on) { $('#activityDrawer').classList.toggle('open', on); $('#activityBackdrop').classList.toggle('open', on); if (on) { lsSet('activityUnseen', 0); this.badge(); this.render(); } },
};
const LOGGED = new Proxy({}, { get: (t, k) => LISTS[k] && k !== USERS_LIST ? [iconFor(k), r => `${k.replace(/_/g, ' ')} — ${LISTS[k].title(r)}`] : undefined });
{ const _save = DB.save.bind(DB), _rm = DB.remove.bind(DB);
  DB.save = async (k, rec) => { const isNew = !rec._id, out = await _save(k, rec); if (LOGGED[k] && !Activity.quiet) Activity.add(LOGGED[k][0], (isNew ? 'Added: ' : 'Saved: ') + LOGGED[k][1](out)); return out; };
  DB.remove = async (k, rec) => { await _rm(k, rec); if (LOGGED[k]) Activity.add('🗑', 'Deleted: ' + LOGGED[k][1](rec)); }; }

/* =====================================================================
   BOOT
   ===================================================================== */
function updateConn() {
  const c = cfg(), p = $('#connPill'), t = $('#connText');
  if (c.mode !== 'sp') { p.className = 'token-pill throttled'; t.textContent = 'Demo data'; $('#sbMode').textContent = 'Demo data (this browser)'; return; }
  if (!c.siteUrl) { p.className = 'token-pill throttled'; t.textContent = MS.token ? 'Token ready — Site URL missing' : 'No token — Admin required'; return; }
  p.className = 'token-pill' + (MS.token ? ' ok' : ''); t.textContent = MS.token ? 'Token ready' : 'No token — sign in'; $('#sbMode').textContent = 'SharePoint · ' + (c.siteUrl.split('/').pop() || '');
}
function applyBranding() { brandApply(); const c = company(); if (c.CompanyFullName) $('[data-brand="tag2"]').textContent = c.CompanyFullName; }
function buildNav() {
  const secs = []; for (const p of PAGES) if (!p.hidden && Auth.can(p.id) && !secs.includes(p.sec)) secs.push(p.sec);
  const col = lsGet('navCollapsed', {});
  $('#nav').innerHTML = secs.map(s => { const ps = PAGES.filter(p => p.sec === s && !p.hidden && Auth.can(p.id)); return `<div class="nav-group ${col[s] ? 'collapsed' : ''}" data-g="${esc(s)}"><div class="nav-group-head"><span class="caret">▾</span><span class="gname">${esc(s)}</span><span class="gcount">${ps.length}</span></div><div class="nav-group-body">` + ps.map(p => `<div class="nav-item" data-id="${p.id}" title="${esc(p.label)}"><span class="icon">${p.icon}</span><span>${esc(p.label)}</span><span class="badge alert" data-badge="${p.id}" style="display:none"></span></div>`).join('') + '</div></div>'; }).join('');
  $$('.nav-item').forEach(n => n.onclick = () => { document.body.classList.remove('m-nav-open'); go(n.dataset.id); });
  const save = () => { const c = {}; $$('.nav-group').forEach(g => { if (g.classList.contains('collapsed')) c[g.dataset.g] = 1; }); lsSet('navCollapsed', c); };
  $$('.nav-group-head').forEach(h => h.onclick = () => { h.parentElement.classList.toggle('collapsed'); save(); });
  $('#navCollapseAll').onclick = () => { $$('.nav-group').forEach(g => g.classList.add('collapsed')); save(); };
  $('#navExpandAll').onclick = () => { $$('.nav-group').forEach(g => g.classList.remove('collapsed')); save(); };
}
function navMini(on) { document.body.classList.toggle('sb-mini', on); $('#sbToggle').textContent = on ? '»' : '«'; $('#sbToggle').title = on ? 'Expand the side panel' : 'Collapse the side panel'; lsSet('navMini', on ? 1 : 0); }
function updateBadges() { const b = (id, n) => { const e = $(`[data-badge="${id}"]`); if (e) { e.textContent = n; e.style.display = n ? '' : 'none'; } };
  b('stock', stockRows().filter(r => r._status !== 'OK' && DB.peek('LOTTable').length).length); b('fefo', DB.peek('LOTTable').filter(l => num(l.BalanceQty) > 0 && !l.CloseLot && l.ExpiryDate && daysBetween(today(), l.ExpiryDate) <= 30).length);
  b('lpo-pending-for-approval', DB.peek('LPOHeader').filter(h => !h.ApproveLPO && !h.CancelPO).length); b('fleetc', fleetAlerts(30).length); b('docexp', staffAlerts(30).length); }
Bus.on('data', debounce(updateBadges, 300));
Bus.on('loadError', debounce((k, e) => toast(`Could not load ${listName(k)}: ${e.message}`, 'error', 6000), 300));
/* page history for the ← Back button */
const Hist = { stack: [], last: null };
window.addEventListener('hashchange', () => { if (Hist.last !== null && Hist.last !== location.hash && !Hist.goingBack) Hist.stack.push(Hist.last); Hist.goingBack = false; Hist.last = location.hash; $('#backBtn').disabled = !Hist.stack.length; });
function applyTheme(t) { document.body.classList.toggle('dark', t === 'dark'); $('#themeBtn').textContent = t === 'dark' ? '☀' : '🌙'; try { localStorage.setItem(SC_KEY.theme, t); } catch (e) {} }
async function enterApp() {
  $('#loginScreen').style.display = 'none'; $('#appShell').classList.add('visible');
  const u = Auth.user; $('#uName').textContent = u.full || u.name; $('#uRole').textContent = (u.super ? 'Super admin · ' : '') + u.level; $('#uAvatar').textContent = String(u.full || u.name).trim().charAt(0).toUpperCase();
  sessionStorage.setItem(APP.ls + '_user', JSON.stringify(u));
  buildNav(); updateConn(); Activity.badge(); Hist.last = location.hash;
  loadCore().then(() => { applyBranding(); loadAll(['ItemTable', 'LOTTable', 'LPOHeader', 'Vehicles_tbl', 'Employees']).then(updateBadges); });
  if (u.bootstrap) setTimeout(() => toast('No system users exist yet — you are signed in as administrator. Add users to the LoginUserList list (Admin → Login Users).', 'warn', 9000), 800);
  route();
}
const _route = route; route = async function () { await _route(); const id = Nav.cur && Nav.cur.id; document.body.classList.toggle('on-home', id === 'home'); trackVisit(id); };
async function initLogin() {
  const c = cfg(), sp = c.mode === 'sp';
  $('#msBox').style.display = sp ? '' : 'none';
  $('#loginHint').style.display = sp ? 'none' : '';
  $('#loginHint').innerHTML = '<b>Demo mode</b> — data stays in this browser. Sign in with <code>admin</code> / <code>admin</code> (or <code>store</code> / <code>1234</code> as a normal user). Connect SharePoint in Settings.';
  if (sp && location.protocol === 'file:') {
    $('#msLoginState').innerHTML = '⚠ This page was opened as a <b>local file</b>. Microsoft (SharePoint) sign-in only works from an <b>https://</b> address or <b>http://localhost</b>.';
    $('#msLoginBtn').textContent = 'Use demo data on this computer'; $('#msLoginBtn').onclick = () => { saveCfg({ mode: 'demo' }); location.reload(); };
    setTimeout(() => $('#loginUser').focus(), 50); return;
  }
  if (sp) { try { await MS.getToken(false); $('#msLoginState').innerHTML = '✔ Connected to Microsoft 365 as <b>' + esc(MS.user()) + '</b>.'; $('#msLoginBtn').textContent = 'Renew Microsoft session'; } catch (e) { if (!/Not signed/.test(e.message)) $('#msLoginState').textContent = '⚠ ' + e.message; } }
  const saved = JSON.parse(sessionStorage.getItem(APP.ls + '_user') || 'null');
  if (saved && (!sp || MS.token)) { Auth.user = saved; return enterApp(); }
  setTimeout(() => $('#loginUser').focus(), 50);
}
async function doLogin(e) {
  if (e) e.preventDefault();
  const err = $('#loginError'); err.classList.remove('visible'); const b = $('#loginBtn');
  if (!$('#loginUser').value.trim() || !$('#loginPw').value) { err.textContent = 'Enter both username and password.'; err.classList.add('visible'); return; }
  b.disabled = true; b.textContent = DB.isSP() ? 'Checking SharePoint…' : 'Signing in…';
  try { if (DB.isSP() && location.protocol === 'file:') throw new Error('Opened as a local file — SharePoint sign-in is not possible here. Click “Use demo data on this computer”, or open the page from https:// or http://localhost.'); if (DB.isSP()) await MS.getToken(true); await Auth.login($('#loginUser').value, $('#loginPw').value); Activity.quiet = true; enterApp(); setTimeout(() => { Activity.quiet = false; Activity.add('🔓', 'Signed in'); }, 1500); }
  catch (x) { err.textContent = x.message; err.classList.add('visible'); } finally { b.disabled = false; b.textContent = 'Sign in'; }
}
document.addEventListener('DOMContentLoaded', () => {
  let th = null; try { th = localStorage.getItem(SC_KEY.theme); } catch (e) {}
  applyTheme(th || (lsGet('theme', '') === 'dark' ? 'dark' : 'light')); accentApply(accentGet()); brandApply(); try { preloadDocImages(); } catch (e) {} if (lsGet('navMini', 0)) navMini(true);
  $('#loginForm').addEventListener('submit', doLogin);
  $('#msLoginBtn').onclick = async () => { try { await MS.getToken(true); $('#msLoginState').innerHTML = '✔ Connected to Microsoft 365 as <b>' + esc(MS.user()) + '</b>.'; } catch (e) { $('#msLoginState').textContent = '✘ ' + e.message; } };
  $('#logoutBtn').onclick = async () => { Activity.add('🔒', 'Signed out'); await Auth.logout(); location.hash = ''; location.reload(); };
  $('#myInfoBtn').onclick = () => { const u = Auth.user || {}, e = DB.peek(USERS_LIST).find(x => x._id === u._id) || {}; Modal.open('My account', `<div class="grid two"><div class="form-group"><label>User name</label><input readonly value="${esc(u.name || '')}"></div><div class="form-group"><label>Access level</label><input readonly value="${esc((u.super ? 'Super admin · ' : '') + (u.level || ''))}"></div><div class="form-group"><label>E-mail</label><input readonly value="${esc(e.Email_ID || u.email || '')}"></div><div class="form-group"><label>Microsoft account</label><input readonly value="${esc(MS.user() || '—')}"></div></div>`, [...(u._id ? [{ label: '🔑 Change my password', cls: 'secondary', onClick: () => { Modal.close(); changeMyPassword(); } }] : []), { label: 'Close' }], 'narrow'); };
  $('#themeBtn').onclick = () => { applyTheme(document.body.classList.contains('dark') ? 'light' : 'dark'); if (Nav.cur && Nav.cur.id === 'dash') route(); };
  $('#menuBtn').onclick = () => document.body.classList.toggle('m-nav-open'); $('#mBackdrop').onclick = () => document.body.classList.remove('m-nav-open');
  $('#sbToggle').onclick = () => navMini(!document.body.classList.contains('sb-mini'));
  $('#sbBrand').onclick = () => go('home');
  $('#backBtn').onclick = () => { if (!Hist.stack.length) return; Hist.goingBack = true; location.hash = Hist.stack.pop(); $('#backBtn').disabled = !Hist.stack.length; };
  $('#refreshBtn').onclick = async () => { DB.resetCache(); try { await loadCore(); toast('Data reloaded', 'success'); } catch (e) { toast(e.message, 'error'); } route(); };
  $('#msReloginBtn').onclick = async () => { if (!DB.isSP()) return toast('Demo data — no Microsoft sign-in needed. Connect SharePoint in Settings.', ''); try { MS.token = null; await MS.getToken(true); updateConn(); toast('Microsoft session renewed ✔', 'success'); } catch (e) { toast(e.message, 'error', 6000); } };
  $('#connPill').onclick = () => Auth.isAdmin() ? go('settings') : null;
  $('#searchPill').onclick = openSearch;
  $('#activityToggle').onclick = () => Activity.open(true); $('#activityCloseBtn').onclick = $('#activityBackdrop').onclick = () => Activity.open(false);
  $('#activityClearBtn').onclick = () => { lsSet('activity', []); Activity.render(); };
  $('#activityExportBtn').onclick = () => exportXlsx('Activity', Activity.list(), [{ f: 't', l: 'Time', get: r => fdt(r.t) }, { f: 'u', l: 'User' }, { f: 'x', l: 'Activity' }]);
  $('#modalX').onclick = Modal.close; $('#modalBg').addEventListener('mousedown', e => { if (e.target.id === 'modalBg') Modal.close(); });
  document.addEventListener('keydown', e => { if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k' && Auth.user) { e.preventDefault(); openSearch(); } if (e.key === 'Escape') { if ($('#modalBg').classList.contains('open')) Modal.close(); else Activity.open(false); } if (e.altKey && e.key === 'ArrowLeft' && !(Nav.cur && ['trans', 'inq'].includes(Nav.cur.id))) $('#backBtn').click(); });
  window.addEventListener('hashchange', route);
  setInterval(() => { if (DB.isSP() && Auth.user) MS.getToken(false).then(updateConn).catch(() => updateConn()); }, 40 * 60 * 1000);
  initLogin();
});
