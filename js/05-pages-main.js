/* =====================================================================
   PAGES — one per Access form (see Help → Form map)
   ===================================================================== */
const PAGES = [];
const page = (id, sec, icon, label, sub, render, form) => PAGES.push({ id, sec, icon, label, sub, render, form });
const Nav = { cur: null, guard: null };
function go(id, param) { const h = '#' + id + (param !== undefined && param !== null && param !== '' ? '/' + encodeURIComponent(param) : ''); if (location.hash === h) route(); else location.hash = h; }
async function route() {
  if (!Auth.user) return;
  let [id, param] = location.hash.replace(/^#/, '').split('/'); param = param ? decodeURIComponent(param) : '';
  let P = PAGES.find(p => p.id === id); if (!P || !Auth.can(P.id)) { P = PAGES.find(p => Auth.can(p.id)); param = ''; }
  if (Nav.guard && Nav.cur && !(Nav.cur.id === P.id && Nav.cur.param === param)) { const g = Nav.guard; if (!(await g())) { history.replaceState(null, '', '#' + Nav.cur.id + (Nav.cur.param ? '/' + encodeURIComponent(Nav.cur.param) : '')); return; } }
  Nav.guard = null; Nav.cur = { id: P.id, param };
  $$('.nav-item').forEach(n => n.classList.toggle('active', n.dataset.id === P.id));
  $('#pgTitle').textContent = P.label; $('#pgSub').textContent = P.sub || ''; $('#sidebar').classList.remove('open');
  const c = $('#content'); c.innerHTML = ''; const v = document.createElement('div'); v.className = 'view active'; v.id = 'view-' + P.id; c.appendChild(v); c.scrollTop = 0;
  try { v.innerHTML = '<div class="empty">Loading…</div>'; await P.render(v, param); } catch (e) { console.error(e); v.innerHTML = `<div class="card"><div class="status-msg error">⚠ ${esc(e.message)}</div><div class="muted" style="margin-top:8px">If this is a SharePoint error, check Settings → SharePoint lists.</div></div>`; }
}
function setDirtyGuard(isDirty) { Nav.guard = async () => !isDirty() || await confirmBox('You have unsaved changes on this page. Leave without saving?', 'Leave', true); }
const STATUS_CLS = { ok: /^(ok|closed|done|delivered|completed|approved|active|posted|yes|paid|received|online)$/i, err: /^(cancel+ed|rejected|expired|out of stock|returned|blocked|deactive|inactive|no|offline)$/i, warn: /^(pending|hold|due|below minimum|in progress|partly|partial|open)$/i, info: /^(new|processing|sent|draft)$/i };
const badgeFor = s => { const c = Object.keys(STATUS_CLS).find(k => STATUS_CLS[k].test(String(s || '').trim())) || ''; return `<span class="badge ${c}">${esc(s || '—')}</span>`; };
const dayBadge = d => d === null || d === undefined ? '' : `<span class="badge ${d < 0 ? 'err' : d <= 30 ? 'warn' : d <= 90 ? 'info' : 'ok'}">${d < 0 ? 'expired ' + -d + ' d' : d + ' d'}</span>`;
function monthly(rows, df, fields, n = 12) {
  const labels = [], keys = []; const d = new Date(); d.setDate(1);
  for (let i = n - 1; i >= 0; i--) { const x = new Date(d.getFullYear(), d.getMonth() - i, 1); keys.push(`${x.getFullYear()}-${pad(x.getMonth() + 1)}`); labels.push(MON[x.getMonth()] + ' ' + String(x.getFullYear()).slice(2)); }
  const values = fields.map(f => keys.map(k => sum(rows.filter(r => ym(typeof df === 'function' ? df(r) : r[df]) === k), f)));
  return { labels, values };
}
const loadAll = ks => Promise.all(ks.filter(k => LISTS[k]).map(k => DB.all(k).catch(e => { DB.cache[k] = DB.cache[k] || []; Bus.emit('loadError', k, e); })));
const kpi = (l, val, sub, col, link) => `<div class="kpi" style="--kc:${col};${link ? 'cursor:pointer' : ''}" ${link ? `onclick="go('${link}')"` : ''}><div class="k-label">${l}</div><div class="k-val">${val}</div><div class="k-sub">${sub || ''}</div></div>`;
const pageFor = (k, prefer) => (prefer && PAGES.find(p => p.id === prefer)) || PAGES.find(p => p.gen && p.gen.table === k && p.gen.subs.length) || PAGES.find(p => p.gen && p.gen.table === k);

/* =====================================================================
   DASHBOARD (Access main menu / INICIO switchboard)
   ===================================================================== */
const DASH_LISTS = ['ItemTable', 'LOTTable', 'IngdMvmntData', 'IngdMvmntHdr', 'LPOHeader', 'Vehicles_tbl', 'Vehicles_Defect_Reg_tbl', 'Employees', 'SC_Deliveries_Orders_tbl'];
page('dash', 'Main', '📊', 'Dashboard', 'Warehouse, RM & PM, distribution, fleet and HR at a glance', async v => {
  await loadAll(DASH_LISTS);
  const C = [CSSV('--c1'), CSSV('--c2'), CSSV('--c3'), CSSV('--c4')];
  const st = stockRows(), low = st.filter(r => r._status !== 'OK'), lots = DB.peek('LOTTable').filter(l => num(l.BalanceQty) > 0 && !l.CloseLot && l.ExpiryDate);
  const exp30 = lots.filter(l => daysBetween(today(), l.ExpiryDate) <= 30 && daysBetween(today(), l.ExpiryDate) >= 0), expired = lots.filter(l => daysBetween(today(), l.ExpiryDate) < 0);
  const lpo = DB.peek('LPOHeader'), pend = lpo.filter(h => !h.ApproveLPO && !h.CancelPO), lpoM = lpo.filter(h => ym(h['LPO Date']) === ym(today()) && !h.CancelPO);
  const fa = fleetAlerts(30), rep = DB.peek('Vehicles_Defect_Reg_tbl'), openRep = rep.filter(r => !r.Post), sa = staffAlerts(30);
  const del = DB.peek('SC_Deliveries_Orders_tbl'), delM = del.filter(r => ym(r.OrderDate || r.TransactionDate) === ym(today()));
  const unposted = DB.peek('IngdMvmntHdr').filter(h => !h.Post);
  const mv = DB.peek('IngdMvmntData'), mvD = r => r.DocDate || (DB.peek('IngdMvmntHdr').find(h => num(h.headerID) === num(r.Relation)) || {}).DocumentDate;
  const mM = monthly(mv, mvD, [r => r.ReceivedQtyBase ?? r.ReceivedQty, r => r.ReleasedQtyBase ?? r.ReleasedQty]);
  const mD = monthly(del, r => r.OrderDate || r.TransactionDate, ['OrderedQuantity', 'DeliveredQuantity']);
  const mL = monthly(lpo.filter(h => !h.CancelPO), 'LPO Date', ['LPOAmount']);
  const mR = monthly(rep, 'WS Doc Date', ['Repairing Cost']);
  const pg = id => (PAGES.find(p => p.id === id) || {}).id || '';
  v.innerHTML = `
  <div class="btn-row" style="margin-bottom:14px"><span class="muted">As of ${fdate(today())}</span><span class="spacer"></span>
    ${Auth.can(pg('ingd-receiving-window')) ? `<button onclick="go('ingd-receiving-window','new')">＋ RM/PM receiving</button><button class="secondary" onclick="go('ingd-releasing-window','new')">＋ RM/PM releasing</button>` : ''}
    ${Auth.can(pg('lpo-window')) ? `<button class="secondary" onclick="go('lpo-window','new')">＋ LPO</button>` : ''}${Auth.can(pg('vehicle-repairs-window')) ? `<button class="secondary" onclick="go('vehicle-repairs-window','new')">＋ Vehicle repair</button>` : ''}</div>
  <div class="kpis">
    ${kpi('Items in stock', intf(st.filter(r => r._bal > 0).length), `of ${st.length} items`, C[0], 'stock')}${kpi('Below minimum / out', `<span class="${low.length ? 'neg' : ''}">${low.length}</span>`, 'items to re-order', C[2], 'stock')}
    ${kpi('LOTs expiring ≤ 30 d', `<span class="${exp30.length ? 'neg' : ''}">${exp30.length}</span>`, intf(sum(exp30, 'BalanceQty')) + ' units', C[3], 'fefo')}${kpi('Expired LOT stock', `<span class="${expired.length ? 'neg' : ''}">${expired.length}</span>`, intf(sum(expired, 'BalanceQty')) + ' units on hand', C[2], 'fefo')}
    ${kpi('LPOs pending approval', pend.length, money(sum(pend, 'LPOAmount')) + ' AED', C[0], pg('lpo-pending-for-approval') || 'lpo-list')}${kpi('LPO value this month', money(sum(lpoM, 'LPOAmount')), lpoM.length + ' LPOs', C[1], 'lpo-list')}
    ${kpi('Deliveries this month', intf(sum(delM, 'DeliveredQuantity')), `${delM.length} order lines · ${del.length ? (sum(del, 'DeliveredQuantity') / Math.max(1, sum(del, 'OrderedQuantity')) * 100).toFixed(1) : '0'}% fill rate`, C[1], 'sc-deliveries-orders-ds')}
    ${kpi('Fleet alerts ≤ 30 d', `<span class="${fa.length ? 'neg' : ''}">${fa.length}</span>`, 'insurance, registration, service', C[3], 'fleetc')}${kpi('Open vehicle repairs', openRep.length, money(sum(openRep, 'Repairing Cost')) + ' AED', C[2], 'vehicle-repairs-list')}
    ${kpi('Staff documents ≤ 30 d', `<span class="${sa.length ? 'neg' : ''}">${sa.length}</span>`, 'visa, passport, EID, licence', C[0], 'docexp')}
  </div>
  <div class="charts">
    <div class="card"><div class="card-title" style="margin-bottom:8px">RM &amp; PM — received vs released (qty)</div>${barChart(mM.labels, [{ name: 'Received', values: mM.values[0], color: C[0] }, { name: 'Released', values: mM.values[1], color: C[2] }])}</div>
    <div class="card"><div class="card-title" style="margin-bottom:8px">Distribution — ordered vs delivered (qty)</div>${barChart(mD.labels, [{ name: 'Ordered', values: mD.values[0], color: C[3] }, { name: 'Delivered', values: mD.values[1], color: C[1], line: true }])}</div>
    <div class="card"><div class="card-title" style="margin-bottom:8px">LPO value per month (AED)</div>${barChart(mL.labels, [{ name: 'LPO amount', values: mL.values[0], color: C[0] }])}</div>
    <div class="card"><div class="card-title" style="margin-bottom:8px">Vehicle repair cost per month (AED)</div>${barChart(mR.labels, [{ name: 'Repair cost', values: mR.values[0], color: C[2] }])}</div>
  </div>
  <div class="grid two" style="margin-top:4px">
    <div class="card"><div class="card-title" style="margin-bottom:12px">🔔 Alerts</div><div class="alert-list">
      ${[['📉 Items below minimum stock / out of stock', low.length, "go('stock','low')"], ['⏳ LOTs expiring in 30 days', exp30.length, "go('fefo','30')"], ['⛔ Expired LOTs still on hand', expired.length, "go('fefo','expired')"], ['📌 RM/PM documents not posted', unposted.length, `go('${pg('items-lot-list') || 'ingd-receiving-window'}')`], ['🧾 LPOs pending approval', pend.length, `go('${pg('lpo-pending-for-approval') || 'lpo-list'}')`], ['🚛 Vehicle documents / service due ≤ 30 days', fa.length, "go('fleetc')"], ['🔧 Vehicle repairs not posted', openRep.length, "go('vehicle-repairs-list')"], ['🪪 Staff documents expiring ≤ 30 days', sa.length, "go('docexp')"]]
      .map(([t, n, a]) => `<div class="alert-item" onclick="${a}">${t}<span class="n ${n ? 'neg' : 'pos'}">${n}</span></div>`).join('')}</div></div>
    <div class="card"><div class="card-title" style="margin-bottom:12px">⏳ Next LOTs to release (FEFO)</div>
      ${lots.length ? `<div class="tbl-wrap"><table class="dt"><thead><tr><th>Item</th><th>Expiry</th><th class="num">Balance</th><th>Left</th></tr></thead><tbody>${lots.sort((a, b) => String(a.ExpiryDate).localeCompare(String(b.ExpiryDate))).slice(0, 8).map(l => `<tr onclick="go('fefo')" style="cursor:pointer"><td>${esc(l.ItemDesc)}</td><td>${fdate(l.ExpiryDate)}</td><td class="num">${intf(l.BalanceQty)}</td><td>${dayBadge(daysBetween(today(), l.ExpiryDate))}</td></tr>`).join('')}</tbody></table></div>` : '<div class="empty">No open LOTs.</div>'}</div>
  </div>`;
}, 'INICIO (main menu), Main Switchboard, Managment Stocks Report');

/* =====================================================================
   WAREHOUSE — Stock balance (ItemTable + LOTTable + IngdMvmntData)
   ===================================================================== */
page('stock', 'Warehouse', '📦', 'Stock Balance', 'Live stock per item from posted receipts / releases and open LOTs — with FEFO drill-down', async (v, param) => {
  await loadAll(['ItemTable', 'LOTTable', 'IngdMvmntData', 'IngdMvmntHdr']);
  const rows = stockRows();
  v.innerHTML = `<div class="card"><div class="toolbar"><input type="search" placeholder="Search item…" data-q><select data-c><option value="">Category: all</option>${uniq(rows.map(r => r.ItemCategory || r.ItemType)).map(x => `<option>${esc(x)}</option>`).join('')}</select><select data-s><option value="">Status: all</option>${['OK', 'Below minimum', 'Out of stock'].map(x => `<option ${param === 'low' && x === 'Below minimum' ? '' : ''}>${x}</option>`).join('')}<option value="low" ${param === 'low' ? 'selected' : ''}>Below minimum + out</option></select><span class="spacer" style="flex:1"></span><span class="muted" data-n></span><button class="ghost small" data-x>⬇ Excel</button><button class="ghost small" data-p>🖨 Print</button></div><div data-g></div></div>`;
  const cols = [{ f: 'ItemCode', l: 'Code' }, { f: 'ItemDesc', l: 'Item' }, { l: 'Category', get: r => r.ItemCategory || r.ItemType }, { f: 'UOMID', l: 'UOM' }, { f: '_rec', l: 'Received', t: 'n', sum: 1 }, { f: '_rel', l: 'Released', t: 'n', sum: 1 }, { f: '_bal', l: 'On hand', t: 'n', sum: 1 }, { f: '_min', l: 'Minimum', t: 'n' }, { f: '_lots', l: 'Open LOTs', t: 'n' }, { f: '_next', l: 'Next expiry', t: 'd' }, { l: 'Left', html: r => dayBadge(r._days) }, { f: '_status', l: 'Status', html: r => badgeFor(r._status) }, { l: 'Stock value', t: 'c', sum: 1, get: r => r._bal * num(r.CostPrice) }];
  const now = () => { let r = rows; const c = $('[data-c]', v).value, s = $('[data-s]', v).value; if (c) r = r.filter(x => (x.ItemCategory || x.ItemType) === c); if (s === 'low') r = r.filter(x => x._status !== 'OK'); else if (s) r = r.filter(x => x._status === s); return textFilter(r, $('[data-q]', v).value, ['ItemDesc', 'ItemCode', 'ItemCategory', 'ItemLocation']); };
  const g = makeGrid($('[data-g]', v), { cols, rows: now(), sortBy: 'ItemDesc', desc: false, rowClass: r => r._status !== 'OK' ? 'warn-row' : '', onRow: r => lotDrill(r.ItemDesc) });
  const re = () => { const r = now(); $('[data-n]', v).textContent = r.length + ' items'; g.redraw(r); }; re();
  $('[data-q]', v).oninput = debounce(re, 200); $('[data-c]', v).onchange = $('[data-s]', v).onchange = re;
  $('[data-x]', v).onclick = () => exportXlsx('Stock_Balance', now(), cols.filter(c => c.f || c.get).map(c => ({ ...c, f: c.f || c.l })));
  $('[data-p]', v).onclick = () => printTable('Stock Balance', now(), cols.filter(c => !c.html));
}, 'StockList, ItemsStockList, ItemsListStockLevelTrackingWindow, ItemsInventoryStocksByLocation');
function lotDrill(item) {
  const L = DB.peek('LOTTable').filter(l => norm(l.ItemDesc) === norm(item)).sort((a, b) => String(a.ExpiryDate || '9999').localeCompare(String(b.ExpiryDate || '9999')));
  const open = L.filter(l => num(l.BalanceQty) > 0 && !l.CloseLot);
  Modal.open('LOTs — ' + item, `<div class="status-msg info" style="margin:0 0 10px">FEFO order: the first open LOT below is the one the next release takes from. On hand: <b>${intf(sum(open, 'BalanceQty'))}</b> in ${open.length} open LOT(s).</div><div class="tbl-wrap"><table class="dt"><thead><tr><th>#</th><th>LOT</th><th>Expiry</th><th>Left</th><th>Received</th><th class="num">Rcvd</th><th class="num">Released</th><th class="num">Balance</th><th>Supplier / ref.</th></tr></thead><tbody>${L.map((l, i) => `<tr class="${num(l.BalanceQty) <= 0 || l.CloseLot ? 'muted' : ''}"><td>${num(l.BalanceQty) > 0 && !l.CloseLot ? open.indexOf(l) + 1 : '—'}</td><td>${esc(l.LOTID ?? '')}</td><td>${fdate(l.ExpiryDate)}</td><td>${l.ExpiryDate ? dayBadge(daysBetween(today(), l.ExpiryDate)) : ''}</td><td>${fdate(l.RcvShpDate)}</td><td class="num">${intf(l.ReceivedQty)}</td><td class="num">${intf(l.ReleasedQty)}</td><td class="num"><b>${intf(l.BalanceQty)}</b></td><td>${esc([l.SupplierName, l.ShipmentRef].filter(Boolean).join(' · '))}</td></tr>`).join('') || '<tr><td colspan="9" class="empty">No LOTs for this item.</td></tr>'}</tbody></table></div>`, [{ label: 'Close' }], 'wide');
}

/* =====================================================================
   RM & PM — LOT expiry & FEFO (LOTTable, Items Lot Numbers Recalculate Tool)
   ===================================================================== */
page('fefo', 'RM & PM', '⏳', 'LOT Expiry & FEFO', 'Open LOTs by expiry, expired stock, and a FEFO release planner (first-expiry-first-out)', async (v, param) => {
  await loadAll(['LOTTable', 'ItemTable']);
  const open = DB.peek('LOTTable').filter(l => num(l.BalanceQty) > 0 && !l.CloseLot).map(l => ({ ...l, _d: l.ExpiryDate ? daysBetween(today(), l.ExpiryDate) : 99999 }));
  const B = [['expired', 'Expired', l => l._d < 0], ['30', '≤ 30 days', l => l._d >= 0 && l._d <= 30], ['90', '31–90 days', l => l._d > 30 && l._d <= 90], ['180', '91–180 days', l => l._d > 90 && l._d <= 180], ['later', '> 180 days / no date', l => l._d > 180]];
  let cur = B.find(b => b[0] === param) ? param : '';
  const C = [CSSV('--c3'), CSSV('--c4'), CSSV('--c1'), CSSV('--c2'), CSSV('--c2')];
  v.innerHTML = `<div class="kpis">${B.map((b, i) => kpi(b[1], open.filter(b[2]).length, intf(sum(open.filter(b[2]), 'BalanceQty')) + ' units', C[i])).join('')}</div>
    <div class="card"><div class="card-header"><div class="card-title">🧮 FEFO release planner</div></div><div class="card-desc">Pick an item and the quantity to release: the LOTs are taken in order of expiry date (earliest first), exactly as the posting of a releasing document does. Expired LOTs are skipped unless you tick “use expired”.</div>
      <div class="grid three"><div class="form-group"><label>Item</label><input list="fefoItems" id="fItem"><datalist id="fefoItems">${uniq(open.map(l => l.ItemDesc)).map(x => `<option value="${esc(x)}">`).join('')}</datalist></div><div class="form-group"><label>Quantity to release</label><input type="number" id="fQty" min="0"></div><label class="chk" style="align-self:end"><input type="checkbox" id="fExp"> use expired LOTs</label></div><div id="fOut" style="margin-top:10px"></div></div>
    <div class="card"><div class="toolbar"><input type="search" placeholder="Search item, supplier, shipment…" data-q><select data-b><option value="">Expiry: all</option>${B.map(b => `<option value="${b[0]}" ${cur === b[0] ? 'selected' : ''}>${b[1]}</option>`).join('')}</select><span class="spacer" style="flex:1"></span><span class="muted" data-n></span><button class="ghost small" data-x>⬇ Excel</button><button class="ghost small" data-p>🖨 Print</button></div><div data-g></div></div>`;
  const cols = [{ f: 'LOTID', l: 'LOT', t: 'n' }, { f: 'ItemDesc', l: 'Item' }, { f: 'ExpiryDate', l: 'Expiry', t: 'd' }, { f: '_d', l: 'Days left', html: l => l._d === 99999 ? '' : dayBadge(l._d) }, { f: 'RcvShpDate', l: 'Received on', t: 'd' }, { f: 'ReceivedQty', l: 'Received', t: 'n', sum: 1 }, { f: 'ReleasedQty', l: 'Released', t: 'n', sum: 1 }, { f: 'BalanceQty', l: 'Balance', t: 'n', sum: 1 }, { f: 'SupplierName', l: 'Supplier' }, { f: 'ShipmentRef', l: 'Shipment' }];
  const now = () => { let r = open; const b = B.find(x => x[0] === $('[data-b]', v).value); if (b) r = r.filter(b[2]); return textFilter(r, $('[data-q]', v).value, ['ItemDesc', 'SupplierName', 'ShipmentRef', 'CustomerName']); };
  const g = makeGrid($('[data-g]', v), { cols, rows: now(), sortBy: 'ExpiryDate', desc: false, rowClass: l => l._d < 0 ? 'err-row' : l._d <= 30 ? 'warn-row' : '', onRow: l => lotDrill(l.ItemDesc) });
  const re = () => { const r = now(); $('[data-n]', v).textContent = r.length + ' open LOTs'; g.redraw(r); }; re();
  $('[data-q]', v).oninput = debounce(re, 200); $('[data-b]', v).onchange = re;
  $('[data-x]', v).onclick = () => exportXlsx('LOT_Expiry', now(), cols.map(c => ({ ...c, html: null })));
  $('[data-p]', v).onclick = () => printTable('LOT Expiry Report (FEFO)', now(), cols.filter(c => !c.html));
  const plan = () => { const it = $('#fItem', v).value.trim(), q = num($('#fQty', v).value), useExp = $('#fExp', v).checked; if (!it || q <= 0) { $('#fOut', v).innerHTML = ''; return; }
    let left = q; const out = []; for (const l of lotsFEFO(it)) { if (left <= 0) break; const d = l.ExpiryDate ? daysBetween(today(), l.ExpiryDate) : 99999; if (d < 0 && !useExp) continue; const t = Math.min(left, num(l.BalanceQty)); out.push([l, t, d]); left = r2(left - t); }
    $('#fOut', v).innerHTML = `<div class="tbl-wrap"><table class="dt"><thead><tr><th>#</th><th>LOT</th><th>Expiry</th><th>Left</th><th class="num">LOT balance</th><th class="num">Take</th><th class="num">Remaining in LOT</th></tr></thead><tbody>${out.map(([l, t, d], i) => `<tr><td>${i + 1}</td><td>${esc(l.LOTID ?? '')}</td><td>${fdate(l.ExpiryDate)}</td><td>${d === 99999 ? '' : dayBadge(d)}</td><td class="num">${intf(l.BalanceQty)}</td><td class="num"><b>${intf(t)}</b></td><td class="num">${intf(num(l.BalanceQty) - t)}</td></tr>`).join('') || '<tr><td colspan="7" class="empty">No usable LOTs.</td></tr>'}</tbody></table></div>${left > 0 ? `<div class="status-msg warn" style="margin-top:8px">⚠ Short by <b>${intf(left)}</b> — not enough ${useExp ? '' : 'non-expired '}stock.</div>` : '<div class="status-msg success" style="margin-top:8px">✔ Fully covered by FEFO LOTs.</div>'}`; };
  ['#fItem', '#fQty', '#fExp'].forEach(s => { $(s, v).oninput = $(s, v).onchange = debounce(plan, 150); });
}, 'Items Lot Numbers Recalculate Tool, ItemsLotList, ItemsLotTrackingRelease, ItemTrackingWindow');

/* =====================================================================
   FLEET — compliance (Vehicles_tbl expiry & service dates, repairs)
   ===================================================================== */
page('fleetc', 'Fleet', '🛡', 'Fleet Compliance', 'Insurance, registration and service due dates per vehicle, and repair cost', async v => {
  await loadAll(['Vehicles_tbl', 'Vehicles_Defect_Reg_tbl']);
  const al = fleetAlerts(90), veh = DB.peek('Vehicles_tbl'), rep = DB.peek('Vehicles_Defect_Reg_tbl');
  const cost = p => sum(rep.filter(r => norm(r['Vehicle Plate Number']) === norm(p)), 'Repairing Cost');
  v.innerHTML = `<div class="kpis">${kpi('Vehicles', veh.length, uniq(veh.map(x => x.Asset_Regoin)).join(' · '), CSSV('--c1'))}${kpi('Overdue', al.filter(a => a.days < 0).length, 'already expired / past due', CSSV('--c3'))}${kpi('Due ≤ 30 days', al.filter(a => a.days >= 0 && a.days <= 30).length, '', CSSV('--c4'))}${kpi('Repair cost (all)', money(sum(rep, 'Repairing Cost')), rep.length + ' repairs', CSSV('--c2'))}</div>
    <div class="card"><div class="card-title" style="margin-bottom:10px">Due in the next 90 days</div><div data-a></div></div>
    <div class="card"><div class="card-title" style="margin-bottom:10px">Vehicles</div><div data-v></div></div>`;
  makeGrid($('[data-a]', v), { cols: [{ f: 'plate', l: 'Plate' }, { f: 'what', l: 'Item' }, { f: 'date', l: 'Due date', t: 'd' }, { l: 'Left', html: a => dayBadge(a.days) }, { l: 'Type', get: a => a.v.VehicleType }, { l: 'Region', get: a => a.v.Asset_Regoin }, { l: 'In charge', get: a => a.v['Staff in charge'] }], rows: al, sortBy: 'date', desc: false, onRow: a => go(pageFor('Vehicles_tbl', 'vehicle-window').id, 'rec:' + a.v._id), empty: 'Nothing due in the next 90 days ✔' });
  makeGrid($('[data-v]', v), { cols: [{ f: 'Vehicle_Plate_No', l: 'Plate' }, { f: 'VehicleType', l: 'Type' }, { f: 'VehicleModel', l: 'Model' }, { f: 'Asset_Regoin', l: 'Region' }, ...VEH_EXP.slice(0, 4).map(([f, l]) => ({ f, l, t: 'd' })), { l: 'Repair cost', t: 'c', sum: 1, get: x => cost(x.Vehicle_Plate_No) }], rows: veh, sortBy: 'Vehicle_Plate_No', desc: false, onRow: x => go(pageFor('Vehicles_tbl', 'vehicle-window').id, 'rec:' + x._id) });
}, 'Vehicle info List, VehicleSearchList, Vehicle Window (expiry tab)');

/* =====================================================================
   HUMAN CAPITAL — staff document expiry
   ===================================================================== */
page('docexp', 'Human Capital', '🪪', 'Staff Document Expiry', 'Visa, passport, Emirates ID, labour card, health card and driving licence due dates', async v => {
  await loadAll(['Employees']);
  const al = staffAlerts(90);
  v.innerHTML = `<div class="card"><div class="toolbar"><span class="muted">${al.length} document(s) expire within 90 days (or already expired)</span><span class="spacer" style="flex:1"></span><button class="ghost small" data-p>🖨 Print</button></div><div data-g></div></div>`;
  const cols = [{ f: 'name', l: 'Staff' }, { f: 'what', l: 'Document' }, { f: 'date', l: 'Expiry', t: 'd' }, { l: 'Left', html: a => dayBadge(a.days) }];
  makeGrid($('[data-g]', v), { cols, rows: al, sortBy: 'date', desc: false, onRow: () => go(pageFor('Employees', 'employees').id), empty: 'No documents expiring in the next 90 days ✔' });
  $('[data-p]', v).onclick = () => printTable('Staff Document Expiry', al, [...cols.slice(0, 3), { f: 'days', l: 'Days left', t: 'n' }]);
}, 'Staff Profile (documents), EmployeesList');
