/* =====================================================================
   CONFIG, THEMES & ADMIN ENGINE — same settings as NFPC SC Mngt v128
   (brand & logos, colour themes, side-panel layout, switchboard, login control,
   App_Config_tbl central settings, Excel styles, email sending, user options)
   ===================================================================== */

/* ---------------- central settings in SharePoint: App_Config_tbl (Title, Config_Value) ---------------- */
const CFG_LIST = 'App_Config_tbl';
const AppCfg = {
  async rows() { const r = await spRawItems(CFG_LIST); return r; },
  async get(title) { const r = await this.rows(); const it = r.items.find(x => String((x.fields || {}).Title || '').trim() === title); return it ? String((it.fields || {}).Config_Value || '') : ''; },
  async getMany(titles) { const r = await this.rows(), o = {}; for (const t of titles) { const it = r.items.find(x => String((x.fields || {}).Title || '').trim() === t); o[t] = it ? String((it.fields || {}).Config_Value || '') : ''; } return o; },
  async put(map) {
    const r = await this.rows();
    for (const [title, val] of Object.entries(map)) {
      const ex = r.items.find(x => String((x.fields || {}).Title || '').trim() === title);
      if (ex) await gfetch(`/sites/${r.sid}/lists/${r.lid}/items/${ex.id}/fields`, { method: 'PATCH', body: JSON.stringify({ Config_Value: val }) });
      else await gfetch(`/sites/${r.sid}/lists/${r.lid}/items`, { method: 'POST', body: JSON.stringify({ fields: { Title: title, Config_Value: val } }) });
    }
  },
  ready() { return DB.isSP(); },
};
const jget = (k, d) => { try { const v = JSON.parse(localStorage.getItem(k) || 'null'); return v == null ? d : v; } catch (e) { return d; } };
const jset = (k, v) => { try { if (v == null) localStorage.removeItem(k); else localStorage.setItem(k, JSON.stringify(v)); } catch (e) { toast('Could not save on this browser: ' + e.message, 'error'); } };

/* ---------------- colour themes (key shared with SC Mngt) ---------------- */
const ACCENTS = [['nfpc', 'NFPC purple', 'linear-gradient(120deg,#4a1a63,#8d45ab)'], ['rab', 'RAB green & gold', 'linear-gradient(120deg,#1f2321 0%,#0b4d27 45%,#0e7a3a 80%,#c9a227 100%)'], ['blue', 'Blue', 'linear-gradient(120deg,#12306b,#3b82f6)'], ['green', 'Green', 'linear-gradient(120deg,#14532d,#34b863)'], ['navy', 'Dark blue', 'linear-gradient(120deg,#071433,#1e40af)'], ['gray', 'Gray', 'linear-gradient(120deg,#1f2937,#8795a8)'], ['red', 'Dark red', 'linear-gradient(120deg,#5b0f1b,#d6455a)']];
const ACCENT_KEY = SC_KEY.accent, BR_KEY = APP.ls + '_brand', BR_CEN = APP.ls + '_brand_cen', LOGIN_BG_KEY = APP.ls + '_loginbg';
const brandCen = () => jget(BR_CEN, {}) || {};
const accentGet = () => { const z = brandCen(); let own = null; try { own = localStorage.getItem(ACCENT_KEY); } catch (e) {} return (z.force && z.accent) ? z.accent : (own || z.accent || 'nfpc'); };
const loginBgGet = () => { const z = brandCen(); return (z.login || jget(LOGIN_BG_KEY, '') || 'theme') === 'white' ? 'white' : 'theme'; };
function accentApply(a) {
  a = a || accentGet(); const h = document.documentElement;
  if (a === 'blue') h.removeAttribute('data-accent'); else h.setAttribute('data-accent', a);
  if (loginBgGet() === 'white') h.setAttribute('data-login', 'white'); else h.removeAttribute('data-login');
  $$('#setAccent button').forEach(b => b.classList.toggle('on', b.dataset.a === a));
}
function accentSet(a) { const z = brandCen(); if (z.force && z.accent) return toast('The administrator has set the colour theme for everyone.', 'warn'); try { localStorage.setItem(ACCENT_KEY, a); } catch (e) {} accentApply(a); }

/* ---------------- system name, logos, company, side-panel fonts, corners ---------------- */
const BRAND_ORIG = { name: 'NFPC Supply Chain System', tag: 'Warehouse · Distribution · Fleet · Finance · HR', logoLogin: '🔐', logoSide: '🏭', img: LOGO_FILES.icon || NFPC_LOGO, imgDark: NFPC_LOGO_DARK, full: LOGO_FILES.full || NFPC_LOGO, wide: true };
const CO_FIELDS = [['name', 'Trade name', '🏢'], ['addr', 'Address', '📍'], ['phone', 'Tel', '📞'], ['fax', 'Fax', '📠'], ['mail', 'Email', '✉'], ['web', 'Website', '🌐'], ['reg', 'Registration / Tax no.', '🧾']];
/* SC Mngt settings on the same web address are used as a fallback for logos, company details, fonts and corners */
const scBrand = () => { const a = jget(LS_BASE + '_brand_cen', {}) || {}, b = jget(LS_BASE + '_brand', {}) || {}; return { img: b.img || a.img || '', imgDark: b.imgDark || a.imgDark || '', wide: b.wide ?? a.wide, co: { ...(a.co || {}), ...(b.co || {}) }, sb: { ...(a.sb || {}), ...(b.sb || {}) }, rad: b.rad ?? a.rad }; };
function brandGet() {
  const c = jget(BR_KEY, {}) || {}, z = brandCen(), f = !!z.force, D = BRAND_DEFAULT || {}, S = scBrand();
  const P = k => f && z[k] != null && z[k] !== '' ? z[k] : (c[k] != null && c[k] !== '' ? c[k] : (z[k] != null && z[k] !== '' ? z[k] : (D[k] || '')));
  const M = k => f ? { ...(S[k] || {}), ...(D[k] || {}), ...(c[k] || {}), ...(z[k] || {}) } : { ...(S[k] || {}), ...(D[k] || {}), ...(z[k] || {}), ...(c[k] || {}) };
  const pic = (k, orig) => { const v = P(k); return v === 'none' ? '' : (v || S[k] || orig || ''); };
  const logo = P('logo'), img = logo && !P('img') ? '' : pic('img', BRAND_ORIG.img), imgDark = logo && !P('imgDark') ? '' : pic('imgDark', BRAND_ORIG.imgDark);
  const wv = P('wide'); const wide = wv === '' ? (S.wide ?? BRAND_ORIG.wide) : !!wv;
  const rad = [f ? z.rad : null, c.rad, z.rad, D.rad, S.rad].find(x => x != null && x !== '');
  return { name: P('name'), tag: P('tag'), logo, img, imgDark, wide: !!wide && !!(img || imgDark), co: M('co'), sb: M('sb'), rad: rad == null ? '' : rad };
}
function brandApply(b = brandGet()) {
  const name = b.name || BRAND_ORIG.name;
  $$('[data-brand="name"]').forEach(e => e.textContent = name);
  $$('[data-brand="tag"]').forEach(e => e.textContent = b.tag || BRAND_ORIG.tag);
  document.body.classList.toggle('brand-wide', !!b.wide);
  $$('[data-brand="logo"]').forEach(e => { const lg = !!e.closest('.login-brand'), pic = lg ? b.img : (b.imgDark || b.img); if (pic) { e.innerHTML = `<img alt="" src="${esc(pic)}" style="width:100%;height:100%;object-fit:contain;border-radius:inherit">`; e.classList.add('has-img'); e.style.background = 'transparent'; e.style.boxShadow = 'none'; e.style.padding = '0'; } else { e.textContent = b.logo || (lg ? BRAND_ORIG.logoLogin : BRAND_ORIG.logoSide); e.classList.remove('has-img'); e.style.background = e.style.boxShadow = e.style.padding = ''; } });
  $$('[data-brand="full"]').forEach(e => { e.style.display = 'none'; const lc = e.closest('.login-card'); if (lc) lc.classList.remove('has-full'); });
  document.title = name + ' — Secure';
  favApply(b); sbApply(b.sb); radApply(b.rad);
}
function favApply(b) { try { let l = document.querySelector('link[rel~="icon"]'); if (!l) { l = document.createElement('link'); l.rel = 'icon'; document.head.appendChild(l); } if (!l.dataset.orig) l.dataset.orig = l.getAttribute('href') || ''; if (b.img) { l.removeAttribute('type'); l.href = b.img; return; } if (b.logo) { const cv = document.createElement('canvas'); cv.width = cv.height = 64; const g = cv.getContext('2d'); g.fillStyle = CSSV('--acc-2') || '#1F4E79'; g.fillRect(0, 0, 64, 64); g.fillStyle = '#fff'; g.font = 'bold 34px Segoe UI,Arial,sans-serif'; g.textAlign = 'center'; g.textBaseline = 'middle'; g.fillText(String(b.logo).slice(0, 2), 32, 35); l.href = cv.toDataURL('image/png'); return; } l.href = l.dataset.orig; } catch (e) {} }
function radApply(v) {
  let st = $('#radCss'); if (!st) { st = document.createElement('style'); st.id = 'radCss'; document.head.appendChild(st); }
  const n = v === '' || v == null ? NaN : parseInt(v, 10); if (!isFinite(n) || n < 0) { st.textContent = ''; return; } const r = Math.min(40, n), i = Math.round(r * .75), P = 'html body ';
  st.textContent = `${P}.card,${P}.modal,${P}.login-card{border-radius:${r}px!important}${P}.kpi,${P}.hm-kpi,${P}.hm-card,${P}.hm-pan,${P}.hm-tile,${P}.hm-zone{border-radius:${i}px!important}`;
}
function sbApply(o) {
  o = o || {}; let st = $('#sbFontCss'); if (!st) { st = document.createElement('style'); st.id = 'sbFontCss'; document.head.appendChild(st); }
  const fam = v => v ? "font-family:'" + String(v).replace(/['\\<>]/g, '') + "',Segoe UI,Arial,sans-serif!important;" : '', n = parseInt(o.size, 10), set = n >= 9 && n <= 24, z = set ? n : 13, hn = parseInt(o.hsize, 10), h = hn >= 10 && hn <= 28 ? hn : Math.max(14, z + 1), f = fam(o.font), hf = fam(o.hfont);
  st.textContent = (f ? '.sidebar,.sidebar *{' + f + '}' : '') + (hf ? '.sidebar .nav-group-head,.sidebar .nav-group-head *{' + hf + '}' : '')
    + (set ? '.sidebar .nav-item,.sidebar .nav-item>span:not(.icon):not(.badge){font-size:' + z + 'px!important}.sidebar .nav-item .icon{font-size:' + (z + 3) + 'px!important}' : '')
    + (hn ? '.sidebar .nav-group-head,.sidebar .nav-group-head .gname{font-size:' + h + 'px!important}' : '');
  const L = 'rgba(255,255,255,.4)', AC = 'var(--acc-3,#4d7cf0)', G = String(o.gstyle || '').toUpperCase(), P = 'body:not(.sb-mini) .sidebar .nav-group-body', I = P + ' .nav-item';
  st.textContent += ({
    A: P + '{margin-left:22px;border-left:2px solid ' + L + '}' + I + '{position:relative;margin-left:16px!important}' + I + '::before{content:"";position:absolute;left:-16px;top:50%;width:14px;border-top:2px solid ' + L + '}',
    B: P + '{margin-left:22px;border-left:4px solid ' + AC + ';border-radius:3px}' + I + '{margin-left:10px!important}',
    C: P + '{margin:2px 6px 8px 22px;padding:4px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.28);border-radius:10px}' + I + '{margin-left:0!important}',
    D: P + '{margin-left:22px;border-left:2px dotted ' + L + '}' + I + '{position:relative;margin-left:18px!important}' + I + '::before{content:"";position:absolute;left:-23px;top:50%;margin-top:-5px;width:10px;height:10px;border-radius:50%;box-sizing:border-box;background:var(--acc-1);border:2px solid ' + AC + '}' + I + '.active::before{background:' + AC + '}',
    E: P + '{position:relative;margin-left:22px}' + P + '::before{content:"";position:absolute;left:0;top:0;bottom:20px;border-left:2px solid ' + L + '}' + I + '{position:relative;margin-left:16px!important}' + I + '::before{content:"";position:absolute;left:-16px;top:50%;width:14px;height:0;border-top:2px solid ' + L + '}'
  })[G] || '';
}
/* company details for documents: Company Info list (Active row) + the brand's company card (brand wins when filled) */
function docCompany() { const c = { ...company() }, k = brandGet().co || {}; const m = { name: 'CompanyFullName', addr: 'Address', phone: 'TelephoneNumber', fax: 'FaxNumber', mail: 'Email', web: 'Website', reg: 'TRN' }; for (const [a, b] of Object.entries(m)) if (k[a]) c[b] = k[a]; if (k.addr) { c.City = ''; c.Country = ''; } return c; }
/* central brand (App_Config_tbl rows NFPC_Branding / NFPC_BrandingLogo / NFPC_BrandingLogoDark) */
const BR_ROWS = { j: 'NFPC_Branding', a: 'NFPC_BrandingLogo', b: 'NFPC_BrandingLogoDark' };
async function brandPull(silent) {
  if (!AppCfg.ready()) return false;
  try { const v = await AppCfg.getMany(Object.values(BR_ROWS)); if (!v[BR_ROWS.j]) { if (localStorage.getItem(BR_CEN)) { jset(BR_CEN, null); brandApply(); accentApply(); } return false; }
    const z = JSON.parse(v[BR_ROWS.j]); z.img = z.hasImg ? v[BR_ROWS.a] : ''; z.imgDark = z.hasImgDark ? v[BR_ROWS.b] : '';
    if (JSON.stringify(z) !== JSON.stringify(brandCen())) { jset(BR_CEN, z); brandApply(); accentApply(); if (Nav.cur && Nav.cur.id === 'home') route(); } return true; }
  catch (e) { if (!silent) toast('Could not read ' + CFG_LIST + ': ' + e.message, 'warn', 6000); return false; }
}

/* ---------------- Excel export styles (key shared with SC Mngt) ---------------- */
const XL_THEMES = {
  A: { n: 'A · Corporate Blue — blue title band, banded rows', tf: '1F4E79', tc: 'FFFFFF', ac: '2E75B6', h: '1F4E79', hc: 'FFFFFF', hb: 'B8C7DB', b: 'B8C7DB', bo: 'box', a: 'EAF1FB', s: 'D6E4F5' },
  B: { n: 'B · Fresh Teal — clean white title, teal header', tf: null, tc: '0F766E', ac: '14B8A6', h: '0F766E', hc: 'FFFFFF', hb: '14B8A6', b: 'CBD5D1', bo: 'row', a: 'F0F9F8', s: 'CCEBE8' },
  C: { n: 'C · Classic Grey — print friendly, grey header, no colour', tf: null, tc: '222222', ac: '7F7F7F', h: 'D9D9D9', hc: '000000', hb: '999999', b: '999999', bo: 'box', a: null, s: 'EDEDED' },
  D: { n: 'D · Executive Dark — charcoal title, gold header text', tf: '262B33', tc: 'FFFFFF', ac: 'F5C542', h: '262B33', hc: 'F5C542', hb: 'F5C542', b: 'D8D2C2', bo: 'row', a: 'F4F1EA', s: 'E8DFC8' },
  N: { n: 'N · NFPC Purple — purple title band, leaf-green line', tf: '4A1A63', tc: 'FFFFFF', ac: '74B236', h: '702C8C', hc: 'FFFFFF', hb: 'DCC5E8', b: 'DCC5E8', bo: 'box', a: 'F6EEFA', s: 'EAD9F2' },
};
const XL_PREF = () => jget(LS_BASE + '_xlstyle', {}) || {};
const XL_PAL = () => XL_THEMES[XL_PREF().style] || XL_THEMES.A;
let EJ_FAIL = 0;
function loadExcelJs() { if (window.ExcelJS) return Promise.resolve(window.ExcelJS); if (EJ_FAIL > 1) return Promise.reject(new Error('Excel styling library not reachable')); return Promise.race([loadExcelJs0(), new Promise((_, no) => setTimeout(() => { EJ_FAIL++; no(new Error('Excel styling library did not load in time (no internet?)')); }, 6000))]); }
function loadExcelJs0() { return window.ExcelJS ? Promise.resolve(window.ExcelJS) : new Promise((ok, no) => { const sc = document.createElement('script'); sc.src = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js'; sc.onload = () => window.ExcelJS ? ok(window.ExcelJS) : no(new Error('ExcelJS did not load')); sc.onerror = () => no(new Error('Could not load the Excel styling library (no internet?)')); document.head.appendChild(sc); }); }
async function xlBuild(wb, fileName) {
  const EJ = await loadExcelJs(), pref = XL_PREF(), P = XL_PAL(), bi = brandGet(), bk = bi.co || {}, co = (pref.company || '').trim() || bk.name || (company().CompanyFullName) || 'NFPC · Supply Chain';
  const coLine = [bk.addr, bk.phone && 'Tel ' + bk.phone, bk.fax && 'Fax ' + bk.fax, bk.mail && 'Email ' + bk.mail, bk.web && 'Web ' + bk.web, bk.reg && 'Reg/Tax ' + bk.reg].filter(Boolean).join('  ·  '), who = me(), now = new Date().toLocaleString(), out = new EJ.Workbook();
  out.creator = who || 'NFPC'; out.created = new Date();
  let logoId = null; { const m = /^data:image\/(png|jpe?g);base64,/i.exec(bi.img || ''); if (m) { try { logoId = out.addImage({ base64: bi.img, extension: /png/i.test(m[1]) ? 'png' : 'jpeg' }); } catch (e) { logoId = null; } } }
  const thin = c => ({ style: 'thin', color: { argb: c } }), box = c => ({ top: thin(c), left: thin(c), bottom: thin(c), right: thin(c) }), fill = c => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: c } });
  const base = String(fileName || 'Export').replace(/\.xlsx$/i, '').replace(/[_ -]*\d{4}-\d\d-\d\d.*$/, '').replace(/_+/g, ' ').trim() || 'Export', idHdr = /(^|[\s_])(nr|no|id|code|sku|batch|phone|mobile|barcode|plate|trip|doc|document)([\s_.:]|$)|number/i;
  wb.SheetNames.forEach(sn => {
    const ws0 = wb.Sheets[sn], name = String(sn).slice(0, 31); if (!ws0 || !ws0['!ref']) { out.addWorksheet(name); return; }
    const M = XLSX.utils.sheet_to_json(ws0, { header: 1, defval: '', raw: true }); let C = 0; M.forEach(r => { C = Math.max(C, r.length); }); if (!M.length || !C) { out.addWorksheet(name); return; }
    const R = M.length, ws = out.addWorksheet(name), big = R * C > 150000, tc = Math.max(C, 4), generic = /^(sheet\d*|list|export|data)$/i.test(sn);
    let hi = M.findIndex(r => r.filter(v => v !== '' && v != null).length >= Math.max(2, Math.ceil(C / 3))); if (hi < 0) hi = 0;
    ws.mergeCells(1, 1, 1, tc); const t = ws.getCell(1, 1); t.value = `${co}  |  ${(generic ? base : sn).toUpperCase()}`; if (P.tf) t.fill = fill('FF' + P.tf); t.font = { bold: true, size: 14, color: { argb: 'FF' + P.tc } }; t.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; ws.getRow(1).height = 28;
    ws.mergeCells(2, 1, 2, tc); const inf = ws.getCell(2, 1); inf.value = `${base}  ·  ${now}${who ? '  ·  Prepared by ' + who : ''}  ·  ${Math.max(0, R - 1 - hi)} row(s)`; inf.font = { italic: true, size: 10, color: { argb: 'FF555555' } }; inf.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; for (let c = 1; c <= tc; c++) ws.getCell(2, c).border = { bottom: { style: 'medium', color: { argb: 'FF' + P.ac } } }; ws.getRow(2).height = 18;
    if (logoId !== null) { ws.getRow(1).height = 34; try { ws.addImage(logoId, { tl: { col: tc - 0.95, row: 0.06 }, ext: { width: 58, height: 30 }, editAs: 'oneCell' }); } catch (e) {} }
    let top = 3; if (coLine) { ws.mergeCells(3, 1, 3, tc); const cl = ws.getCell(3, 1); cl.value = coLine; cl.font = { size: 9, color: { argb: 'FF555555' } }; top = 4; }
    for (let r = 0; r < hi; r++) { const row = ws.getRow(top + r); M[r].forEach((v, c) => { if (v !== '' && v != null) { const cell = row.getCell(c + 1); cell.value = v; cell.font = { bold: true, size: 10 }; } }); }
    const hr = top + hi, H = M[hi] || [];
    for (let c = 0; c < C; c++) { const cell = ws.getCell(hr, c + 1); cell.value = H[c] === '' || H[c] == null ? null : H[c]; cell.fill = fill('FF' + P.h); cell.font = { bold: true, size: 10, color: { argb: 'FF' + P.hc } }; cell.border = box('FF' + P.hb); cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; }
    ws.getRow(hr).height = 24;
    const isTot = r => /^\s*total/i.test(String(M[r][0] == null ? '' : M[r][0])), colFmt = [];
    for (let c = 0; c < C; c++) colFmt[c] = idHdr.test(String(H[c] == null ? '' : H[c])) ? '0' : null;
    let last = hr; const widths = Array(C).fill(0); H.forEach((h, c) => { widths[c] = String(h == null ? '' : h).length + 4; });
    for (let r = hi + 1; r < R; r++) {
      const row = ws.getRow(top + r), tot = isTot(r), band = (r - hi) % 2 === 0; if (!tot) last = top + r;
      for (let c = 0; c < C; c++) { const v = M[r][c], cell = row.getCell(c + 1); if (v !== '' && v != null) cell.value = v; if (r - hi < 400) { const L = v instanceof Date ? 10 : String(v == null ? '' : v).length; if (L + 2 > widths[c]) widths[c] = L + 2; }
        if (big) continue; const bc = 'FF' + P.b, ac = 'FF' + P.ac; cell.border = P.bo === 'row' ? { bottom: thin(bc) } : box(bc); cell.font = tot ? { bold: true, size: 10 } : { size: 10 };
        if (tot) { cell.fill = fill('FF' + P.s); cell.border = { top: { style: 'medium', color: { argb: ac } }, bottom: { style: 'medium', color: { argb: ac } } }; } else if (band && P.a) cell.fill = fill('FF' + P.a);
        if (typeof v === 'number') { cell.alignment = { horizontal: colFmt[c] === '0' ? 'left' : 'right', vertical: 'middle' }; cell.numFmt = colFmt[c] === '0' ? '0' : (Number.isInteger(v) ? '#,##0' : '#,##0.00#'); } else if (v instanceof Date) { cell.numFmt = 'yyyy-mm-dd'; cell.alignment = { horizontal: 'center', vertical: 'middle' }; } else cell.alignment = { vertical: 'middle', wrapText: String(v).length > 60 }; }
    }
    widths.forEach((w, c) => { ws.getColumn(c + 1).width = Math.max(10, Math.min(46, w)); });
    ws.views = [{ state: 'frozen', ySplit: hr, xSplit: 0 }]; if (last > hr) ws.autoFilter = { from: { row: hr, column: 1 }, to: { row: last, column: C } };
    ws.pageSetup = { orientation: C > 7 ? 'landscape' : 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9, printTitlesRow: `${hr}:${hr}`, margins: { left: .4, right: .4, top: .5, bottom: .6, header: .3, footer: .3 } };
    ws.headerFooter = { oddFooter: `&L${co.replace(/&/g, '&&')}&C${base.replace(/&/g, '&&')}&RPage &P of &N` };
  });
  return out;
}
async function xlBuf(wb, name) { return (await xlBuild(wb, name)).xlsx.writeBuffer(); }
async function xlB64(wb, name) { if (XL_PREF().style !== 'plain') { try { const u = new Uint8Array(await xlBuf(wb, name)); let bin = ''; for (let i = 0; i < u.length; i += 0x8000) bin += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(bin); } catch (e) {} } return XLSX.write(wb, { type: 'base64', bookType: 'xlsx' }); }
(function () {   // every Excel download gets the chosen style (or plain), and obeys the user's Excel option
  if (typeof XLSX === 'undefined' || XLSX.__styled) return; XLSX.__styled = true; const plain = XLSX.writeFile.bind(XLSX);
  XLSX.writeFile = function (wb, name, opts) {
    if (Auth.block('excel')) return;
    if (XL_PREF().style === 'plain' || !/\.xlsx$/i.test(String(name || '')) || (opts && opts.bookType && opts.bookType !== 'xlsx')) return plain(wb, name, opts);
    xlBuf(wb, name).then(buf => { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })); a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000); })
      .catch(err => { try { plain(wb, name, opts); toast('Saved without colours — ' + err.message, 'warn'); } catch (e) {} });
  };
})();

/* ---------------- email: Microsoft 365 (Graph sendMail) · Power Automate flow · Outlook draft (keys shared with SC Mngt) ---------------- */
const MAIL_SEND_DEFAULT = { method: 'graph', flowUrl: '' };
const MAIL_SEND_KEY = LS_BASE + '_mailsend';
function mailCfg() { const c = jget(MAIL_SEND_KEY, {}) || {}; return { method: c.method || MAIL_SEND_DEFAULT.method, flowUrl: (c.flowUrl !== undefined ? c.flowUrl : MAIL_SEND_DEFAULT.flowUrl) || '' }; }
const mailList = t => String(t || '').split(/[;,\s]+/).map(x => x.trim()).filter(Boolean);
async function mailGraphToken() { const app = await MS.init(); let acct = MS.account || app.getAllAccounts()[0]; const req = { scopes: ['https://graph.microsoft.com/Mail.Send'] }; if (!acct) { const r = await app.loginPopup({ scopes: [...MS.scopes, 'https://graph.microsoft.com/Mail.Send'] }); acct = r.account; } try { return (await app.acquireTokenSilent({ ...req, account: acct })).accessToken; } catch (e) { return (await app.acquireTokenPopup({ ...req, account: acct })).accessToken; } }
async function mailSend(m) {   // m: { to, cc, subject, html, attachment:{fileName,dataB64,contentType}|null }
  const c = mailCfg(), to = mailList(m.to), cc = mailList(m.cc), bad = to.concat(cc).filter(x => !EMAIL_RX.test(x));
  if (!to.length) throw new Error('Enter at least one recipient in "To".'); if (bad.length) throw new Error('Not a valid e-mail address: ' + bad.join(', '));
  const att = m.attachment, ctype = (att && att.contentType) || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  if (c.method === 'graph') {
    if (!cfg().clientId) throw new Error('Microsoft sign-in is not set up (Settings → Microsoft sign-in). Use “Outlook draft” instead, or ask the admin.');
    if (att && att.dataB64.length > 3300000) throw new Error('The attachment is too large for Microsoft Graph (about 3 MB).');
    const token = await mailGraphToken(), rc = a => a.map(x => ({ emailAddress: { address: x } }));
    const r = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', { method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify({ message: { subject: m.subject, body: { contentType: 'HTML', content: m.html }, toRecipients: rc(to), ccRecipients: rc(cc), attachments: att ? [{ '@odata.type': '#microsoft.graph.fileAttachment', name: att.fileName, contentType: ctype, contentBytes: att.dataB64 }] : [] }, saveToSentItems: true }) });
    if (!r.ok) { let t = ''; try { t = (await r.json()).error.message; } catch (e) {} throw new Error(`Microsoft Graph said HTTP ${r.status}${t ? ': ' + t : ''}${r.status === 403 ? ' — Mail.Send permission is missing or not consented.' : ''}`); }
    return 'sent through your Microsoft 365 mailbox (a copy is in Sent Items)';
  }
  if (c.method === 'flow') {
    if (!/^https:\/\//i.test(c.flowUrl)) throw new Error('The Power Automate URL is not configured (Settings → Send email directly).');
    const payload = JSON.stringify({ to: to.join(';'), cc: cc.join(';'), subject: m.subject, bodyHtml: m.html, sender: (Auth.user && (Auth.user.email || Auth.user.name)) || '', attachments: att ? [{ name: att.fileName, contentType: ctype, contentBase64: att.dataB64 }] : [] });
    try { const r = await fetch(c.flowUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: payload }); if (!r.ok) throw new Error(`The flow answered HTTP ${r.status}. Check the flow's run history.`); return 'handed to the Power Automate flow (it sends the mail)'; }
    catch (e) { if (!(e instanceof TypeError)) throw e; await fetch(c.flowUrl, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain' }, body: payload }); return 'posted to the Power Automate flow (check the flow run history to confirm)'; }
  }
  throw new Error('Direct sending is not configured — use “Outlook draft”.');
}
function mailDraft(m) {   // .eml file that Outlook opens as a new draft
  const b = 'NFPC_' + Date.now(), att = m.attachment, enc = s => '=?UTF-8?B?' + btoa(unescape(encodeURIComponent(s))) + '?=';
  let eml = `To: ${mailList(m.to).join(', ')}\r\n${m.cc ? `Cc: ${mailList(m.cc).join(', ')}\r\n` : ''}Subject: ${enc(m.subject)}\r\nX-Unsent: 1\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary="${b}"\r\n\r\n--${b}\r\nContent-Type: text/html; charset=UTF-8\r\nContent-Transfer-Encoding: base64\r\n\r\n${btoa(unescape(encodeURIComponent(m.html))).replace(/.{76}/g, '$&\r\n')}\r\n`;
  if (att) eml += `--${b}\r\nContent-Type: ${att.contentType || 'application/octet-stream'}; name="${att.fileName}"\r\nContent-Disposition: attachment; filename="${att.fileName}"\r\nContent-Transfer-Encoding: base64\r\n\r\n${att.dataB64.replace(/.{76}/g, '$&\r\n')}\r\n`;
  eml += `--${b}--\r\n`; const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([eml], { type: 'message/rfc822' })); a.download = (m.subject || 'email').replace(/[^\w\- ]+/g, '').slice(0, 60) + '.eml'; document.body.appendChild(a); a.click(); a.remove();
}
/* the e-mail window (To = line manager, CC = me by default, like SC Mngt) */
function mailDialog(m) {
  if (Auth.block('email')) return; const u = Auth.rec || {}, c = mailCfg();
  const body = Modal.open('✉ Share by email', `<div class="grid two"><div class="form-group span2"><label>To</label><input type="text" id="mTo" value="${esc(m.to || u.managerEmail || '')}" placeholder="name@company.com; …"></div><div class="form-group span2"><label>CC</label><input type="text" id="mCc" value="${esc(m.cc || u.email || '')}"></div><div class="form-group span2"><label>Subject</label><input type="text" id="mSub" value="${esc(m.subject || '')}"></div><div class="form-group span2"><label>Message</label><textarea id="mMsg" rows="3">Dear all,\n\nPlease find ${esc(m.what || 'the document')} below${m.attachment ? ' / attached' : ''}.\n\nBest regards,\n${esc(me())}</textarea></div></div>${m.attachment ? `<div class="muted">📎 ${esc(m.attachment.fileName)}</div>` : ''}<div id="mSt" class="muted" style="margin-top:8px">Sending method: <b>${c.method === 'flow' ? 'Power Automate' : c.method === 'graph' ? 'Microsoft 365' : 'Outlook draft only'}</b> (Settings → Send email directly).</div>`, [
    { label: 'Cancel', cls: 'ghost' }, { label: '📝 Outlook draft (.eml)', cls: 'secondary', onClick: () => { mailDraft(read()); Modal.close(); } },
    ...(c.method !== 'draft' ? [{ label: '🚀 Send now', cls: 'success', onClick: async () => { const st = $('#mSt', body); st.textContent = 'Sending…'; try { st.textContent = '✔ ' + await mailSend(read()); toast('E-mail ' + st.textContent.slice(2), 'success', 6000); Activity.add('✉', `E-mail sent: ${read().subject}`); setTimeout(() => Modal.close(), 900); } catch (e) { st.innerHTML = `<span class="neg">✘ ${esc(e.message)}</span>`; } } }] : [])], 'wide');
  const read = () => ({ to: $('#mTo', body).value, cc: $('#mCc', body).value, subject: $('#mSub', body).value, html: `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:13px">${esc($('#mMsg', body).value).replace(/\n/g, '<br>')}</div><hr>${m.html || ''}`, attachment: m.attachment || null });
}
window.NFPCmailDoc = (title, html) => mailDialog({ subject: title, what: title, html });
async function mailList_xlsx(name, rows, cols) { if (typeof XLSX === 'undefined') return null; const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.map(r => Object.fromEntries(cols.map(c => [c.l, typeof c.get === 'function' ? c.get(r) : r[c.f]])))), String(name).slice(0, 31)); return xlB64(wb, name + '.xlsx'); }

/* ---------------- EmailJS notifications (Notifications page) ---------------- */
const NOTIF_KEY = APP.ls + '_notifications';
const notifCfg = () => jget(NOTIF_KEY, {}) || {};
function loadEmailJs() { return window.emailjs ? Promise.resolve(window.emailjs) : new Promise((ok, no) => { const s = document.createElement('script'); s.src = 'https://cdn.jsdelivr.net/npm/@emailjs/browser@4/dist/email.min.js'; s.onload = () => ok(window.emailjs); s.onerror = () => no(new Error('Could not load EmailJS (no internet?)')); document.head.appendChild(s); }); }
async function notifySend(kind, subject, message) { const n = notifCfg(); if (!n.serviceId || !n.templateId || !n.publicKey) throw new Error('EmailJS is not configured (Admin → Email Notifications).'); const ej = await loadEmailJs(); ej.init({ publicKey: n.publicKey }); await ej.send(n.serviceId, n.templateId, { kind, subject, message, profile: 'NFPC Supply Chain', user: (Auth.user && Auth.user.name) || 'system', timestamp: new Date().toLocaleString(), to_email: n.recipient }); }
function alertsDigest() {
  const low = stockRows().filter(r => r._status !== 'OK'), lots = DB.peek('LOTTable').filter(l => num(l.BalanceQty) > 0 && !l.CloseLot && l.ExpiryDate && daysBetween(today(), l.ExpiryDate) <= 30), pend = DB.peek('LPOHeader').filter(h => !h.ApproveLPO && !h.CancelPO), fa = fleetAlerts(30), sa = staffAlerts(30);
  return [`NFPC Supply Chain — alerts on ${fdate(today())}`, '', `Items below minimum / out of stock: ${low.length}`, ...low.slice(0, 15).map(r => `  • ${r.ItemDesc}: ${intf(r._bal)} (min ${intf(r._min)})`), `LOTs expiring ≤ 30 days (or expired): ${lots.length}`, ...lots.slice(0, 15).map(l => `  • ${l.ItemDesc} — ${fdate(l.ExpiryDate)} — ${intf(l.BalanceQty)}`), `LPOs pending approval: ${pend.length}`, `Fleet items due ≤ 30 days: ${fa.length}`, ...fa.slice(0, 10).map(a => `  • ${a.plate} ${a.what} ${fdate(a.date)}`), `Staff documents ≤ 30 days: ${sa.length}`, ...sa.slice(0, 10).map(a => `  • ${a.name} ${a.what} ${fdate(a.date)}`)].join('\n');
}

/* ---------------- shared token file (sp_token.json, same as SC Mngt) ---------------- */
const TF_KEY = LS_BASE + '_tokenfile';
const tfCfg = () => ({ url: '', auto: true, ...(jget(TF_KEY, {}) || {}) });
const tfExp = tok => { try { const j = JSON.parse(atob(tok.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'))); return j.exp ? j.exp * 1000 : 0; } catch (e) { return 0; } };
function tfApply(o) { if (!o || !o.accessToken) throw new Error('The file has no accessToken.'); const exp = o.expiresAt || tfExp(o.accessToken); if (exp && exp < Date.now()) throw new Error('The token in the file expired on ' + new Date(exp).toLocaleString() + ' — ask the admin to export a fresh file.'); const g = scGet(SC_KEY.global); if (g.accessToken === o.accessToken) return { changed: false, exp }; scSet(SC_KEY.global, { ...g, siteUrl: o.siteUrl || g.siteUrl, accessToken: o.accessToken, updatedAt: Date.now(), updatedBy: 'token file' }); DB.resetCache(); return { changed: true, exp }; }
async function tfLoadUrl(quiet) { const c = tfCfg(); if (!c.url) { if (!quiet) throw new Error('Enter the token file location first.'); return false; } const r = await fetch(c.url, { cache: 'no-store', credentials: 'include' }); if (!r.ok) throw new Error('HTTP ' + r.status); return tfApply(await r.json()); }

/* ---------------- login control: lock screen + check every 5 minutes ---------------- */
const lul = { ok: 0, busy: false, timer: null, warned: false };
const lulExempt = () => !Auth.rec || Auth.rec.builtin || (lulCfg().adm && Auth.rec.role === 'admin') || !DB.isSP();
function lulShow(ico, title, text, btns) {
  let el = $('#lulLock'); if (!el) { el = document.createElement('div'); el.id = 'lulLock'; el.style.cssText = 'display:none;position:fixed;inset:0;z-index:100050;background:var(--acc-grad,#12306b);align-items:center;justify-content:center;padding:20px'; el.innerHTML = '<div style="background:var(--card,#fff);color:var(--gray-800,#222);border-radius:18px;max-width:520px;width:100%;padding:28px 30px;box-shadow:0 30px 80px rgba(0,0,0,.45);text-align:center"><div id="lulIco" style="font-size:44px"></div><h2 id="lulTitle" style="margin:8px 0 6px;font-size:20px"></h2><div id="lulText" style="font-size:13.5px;color:var(--gray-600,#555);line-height:1.55"></div><div id="lulBtns" class="btn-row" style="justify-content:center;margin-top:18px"></div></div>'; document.body.appendChild(el); }
  el.style.display = 'flex'; $('#lulIco').textContent = ico; $('#lulTitle').textContent = title; $('#lulText').innerHTML = text; const b = $('#lulBtns'); b.innerHTML = '';
  (btns || []).forEach(([t, fn, cls]) => { const x = document.createElement('button'); x.type = 'button'; x.className = cls || ''; x.textContent = t; x.onclick = fn; b.appendChild(x); });
}
const lulHide = () => { const el = $('#lulLock'); if (el) el.style.display = 'none'; };
async function lulSignOut() { lulHide(); await Auth.logout(); location.hash = ''; location.reload(); }
async function lulRun() {
  if (lul.busy || !Auth.rec) return; const c = lulCfg(); if (!c.on || lulExempt()) { lulHide(); return; } lul.busy = true;
  try {
    if (!lul.ok) lulShow('🔐', 'Checking your access…', 'Verifying your account with SharePoint (' + esc(c.list) + ')…', []);
    const r = await lulFetch();
    if (r.state === 'ok') { const j = lulJudge(r.items, Auth.rec);
      if (j.allowed) { lul.ok = Date.now(); lulHide(); }
      else { Activity.add('⛔', `Access denied by ${c.list}: ${Auth.rec.username}`);
        lulShow('⛔', j.found ? 'Your access is deactivated' : 'You are not on the access list', j.found ? (j.noCol ? `Your row in <b>${esc(c.list)}</b> has no “${esc(c.allow)}” column.` : `Your account is switched off (<b>${esc(c.allow)}</b> is not Yes). Please contact the administrator.<div style="margin-top:10px;font-size:12px;color:#888">${esc(j.detail || '')}</div>`) : `No row for <b>${esc(j.ids.join(' / '))}</b> was found in <b>${esc(c.list)}</b>. Please contact the administrator.`, [['Sign out', lulSignOut], ['Check again', () => { lul.ok = 0; lul.busy = false; lulRun(); }, 'secondary']]); } }
    else if (r.state === 'nolist') { lulHide(); if (!lul.warned) { lul.warned = true; toast('Login control is on but the list “' + c.list + '” was not found — access is not being checked.', 'warn', 7000); } }
    else if (r.state === 'notoken') { if (!lul.ok) lulShow('🔑', 'Sign in with Microsoft', 'To confirm your access the system needs your Microsoft sign-in.' + (r.msg ? `<div style="margin-top:10px;font-size:12px;color:#c0392b">${esc(r.msg)}</div>` : ''), [['Sign in with Microsoft', async () => { lul.busy = false; try { await MS.getToken(true); updateConn(); lulRun(); } catch (e) { $('#lulText').insertAdjacentHTML('beforeend', `<div style="margin-top:10px;font-size:12px;color:#c0392b">✘ ${esc(e.message)}</div>`); } }], ['Retry', () => { lul.busy = false; lulRun(); }, 'secondary'], ['Sign out', lulSignOut, 'ghost']]); }
    else if (!lul.ok) lulShow('⚠', 'Could not check your access', esc(r.msg || 'SharePoint did not answer.'), [['Retry', () => { lul.busy = false; lulRun(); }], ['Sign out', lulSignOut, 'ghost']]);
  } finally { lul.busy = false; }
}
function lulStart() { lul.ok = 0; lul.warned = false; clearInterval(lul.timer); if (!lulCfg().on || lulExempt()) { lulHide(); return; } lulRun(); lul.timer = setInterval(() => { if (!Auth.rec) return clearInterval(lul.timer); lulRun(); }, 5 * 60 * 1000); }
document.addEventListener('visibilitychange', () => { if (!document.hidden && Auth.rec && lulCfg().on && !lulExempt() && Date.now() - lul.ok > 60000) lulRun(); });

/* ---------------- job-title profiles: which side-panel groups open for a user (App_Config_tbl row NFPC_TitleProfiles) ---------------- */
const TP_KEY = APP.ls + '_tprof', TP_ITEM = 'NFPC_TitleProfiles';
const TP_DEFAULT = { v: 1, profiles: [
  { id: 'store', name: 'Storekeeper', titles: 'storekeeper, store, warehouse, whs', groups: ['Main', 'RM & PM', 'Warehouse', 'Reports'] },
  { id: 'asst', name: 'Assistant Storekeeper', titles: 'assistant, asst, helper, forklift', groups: ['Main', 'Warehouse', 'RM & PM'] },
  { id: 'log', name: 'Logistics', titles: 'logistic, dispatch, distribution', groups: ['Main', 'Distributions', 'Fleet', 'Reports'] },
  { id: 'fleet', name: 'Fleet', titles: 'fleet, tech, mechanic, driver', groups: ['Main', 'Fleet'] },
  { id: 'plan', name: 'Planning', titles: 'planner, planning, demand', groups: ['Main', 'Demand & Planning', 'Production', 'RM & PM', 'Reports'] },
  { id: 'fin', name: 'Finance', titles: 'account, finance, cashier', groups: ['Main', 'Finance', 'Reports'] },
  { id: 'hr', name: 'HR', titles: 'hr, human, payroll, pro', groups: ['Main', 'Human Capital', 'Reports'] }] };
const tpCfg = () => { const j = jget(TP_KEY, null); return j && Array.isArray(j.profiles) && j.profiles.length ? j : JSON.parse(JSON.stringify(TP_DEFAULT)); };
function tpMatchTitle(t, c) { t = String(t || '').toLowerCase(); if (!t) return null; return (c || tpCfg()).profiles.find(p => String(p.titles || '').split(',').map(x => x.trim().toLowerCase()).filter(Boolean).some(k => t.includes(k))) || null; }
function tpProfileFor(u) { if (!u || u.role === 'admin' || u.builtin) return null; const c = tpCfg(); if (u.nfpcProfile === 'none') return null; if (u.nfpcProfile) return c.profiles.find(p => p.id === u.nfpcProfile) || null; return tpMatchTitle(u.jobTitle, c); }
function tpAllows(u, P) { const p = tpProfileFor(u); if (!p) return true; const sec = navSecOf(P.id) || P.sec; return ['home', 'help', 'settings'].includes(P.id) || (p.groups || []).includes(sec) || (p.groups || []).includes(P.sec); }
const tpLabel = u => { const p = tpProfileFor(u); return p ? p.name : ''; };
async function tpPull(silent) { if (!AppCfg.ready()) return false; try { const v = await AppCfg.get(TP_ITEM); if (!v) return false; const j = JSON.parse(v); if (j && Array.isArray(j.profiles)) { jset(TP_KEY, j); return true; } } catch (e) { if (!silent) toast('Could not read ' + CFG_LIST + ': ' + e.message, 'warn'); } return false; }

/* ---------------- side-panel layout: groups, subgroups, order, start collapsed (per user + published for all) ---------------- */
const NAV_ITEM = 'NFPC_SidePanelLayout', NAV_CEN = APP.ls + '_nav_cen', NAV_COL = () => APP.ls + '_navcol_' + ((Auth.rec && Auth.rec.id) || 'anon'), NAV_USER = () => APP.ls + '_nav_' + ((Auth.rec && Auth.rec.id) || 'anon');
function navDefault() { const groups = []; for (const p of PAGES) { if (p.hidden) continue; let g = groups.find(x => x.id === p.sec); if (!g) groups.push(g = { id: p.sec, name: p.sec, items: [], subs: [], collapsed: !['Main', 'Configuration'].includes(p.sec) }); g.items.push(p.id); } return { v: 1, groups }; }
function navReconcile(lay) {
  const d = JSON.parse(JSON.stringify(lay)), known = new Set(PAGES.filter(p => !p.hidden).map(p => p.id)), seen = new Set();
  d.groups = (d.groups || []).map(g => ({ id: g.id || g.name, name: g.name || 'Group', collapsed: !!g.collapsed, items: (g.items || []).filter(v => known.has(v) && !seen.has(v) && seen.add(v)), subs: (g.subs || []).map(s => ({ name: s.name || 'Subgroup', collapsed: !!s.collapsed, items: (s.items || []).filter(v => known.has(v) && !seen.has(v) && seen.add(v)) })) }));
  for (const p of PAGES) { if (p.hidden || seen.has(p.id)) continue; let g = d.groups.find(x => x.id === p.sec) || d.groups.find(x => x.name === p.sec); if (!g) d.groups.push(g = { id: p.sec, name: p.sec, items: [], subs: [], collapsed: true }); g.items.push(p.id); seen.add(p.id); }
  return d;
}
const NAVC = { k: null, v: null };
function navLayout() { let rk = ''; try { rk = (localStorage.getItem(NAV_CEN) || '') + '|' + (localStorage.getItem(NAV_USER()) || '') + '|' + PAGES.length; } catch (e) {} if (NAVC.k === rk && NAVC.v) return JSON.parse(JSON.stringify(NAVC.v)); const r = navLayout0(); NAVC.k = rk; NAVC.v = r; return JSON.parse(JSON.stringify(r)); }
function navLayout0() { const z = jget(NAV_CEN, null), own = jget(NAV_USER(), null); const base = (z && z.force && z.groups) ? z : (own && own.groups) ? own : (z && z.groups) ? z : navDefault(); return navReconcile(base); }
function navSecOf(id) { navLayout(); const L = NAVC.v; const g = L.groups.find(x => x.items.includes(id) || x.subs.some(s => s.items.includes(id))); return g ? g.name : null; }
async function navPull(silent) { if (!AppCfg.ready()) return false; try { const v = await AppCfg.get(NAV_ITEM); if (!v) return false; const j = JSON.parse(v); if (!j || !Array.isArray(j.groups)) return false; const old = JSON.stringify(jget(NAV_CEN, null)); jset(NAV_CEN, j); if (JSON.stringify(j) !== old && Auth.user) buildNav(); if (!silent) toast('✔ Side-panel layout loaded from SharePoint.', 'success'); return true; } catch (e) { if (!silent) toast('Could not read ' + CFG_LIST + ': ' + e.message, 'warn'); return false; } }
function navSuggestSubs(g) {
  const R = [['📋 Lists & search', /list|search|datasheet|overview|data sheet|view$/i], ['📊 Reports & analysis', /report|analysis|tracking|turnover|history|compliance|expiry|balance/i], ['⚙ Setup & masters', /uom|type|group|config|brand|model|window_al|parameter|permeter|currancy|categor|dept|location update|settings/i]];
  const all = [...g.items, ...g.subs.flatMap(s => s.items)], left = [], subs = R.map(r => ({ name: r[0], collapsed: true, items: [] }));
  for (const v of all) { const P = PAGES.find(p => p.id === v), t = P ? P.label + ' ' + (P.form || '') : v, i = R.findIndex(r => r[1].test(t)); if (i >= 0 && !['stock', 'fefo', 'fleetc', 'docexp', 'dash'].includes(v)) subs[i].items.push(v); else left.push(v); }
  const keep = subs.filter(s => s.items.length >= 2); subs.filter(s => s.items.length < 2).forEach(s => left.push(...s.items));
  g.items = left; g.subs = keep; return keep.length;
}

/* ---------------- options switched off for the user → hide their buttons ---------------- */
function featApply() { for (const k of ['excel', 'print', 'import', 'delete', 'email', 'post']) document.body.classList.toggle('no-' + k, !Auth.feat(k)); }
{ const _x = exportXlsx; exportXlsx = (...a) => { if (Auth.block('excel')) return; return _x(...a); }; }
{ const _p = printDoc; printDoc = (...a) => { if (Auth.block('print')) return; return _p(...a); }; }
{ const _i = importDialog; importDialog = (...a) => { if (Auth.block('import') || Auth.ro()) return Auth.ro() && toast('View-only access — import is switched off for you.', 'warn'); return _i(...a); }; }
