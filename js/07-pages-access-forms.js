/* =====================================================================
   GENERIC ACCESS-FORM PAGES — one smart page per Access form (NFPC.pages)
   record source → list; subforms → editable line grids; branded print
   ===================================================================== */
const LABELS = { 'Ingd Receiving Window': 'RM/PM Receiving (GRN)', 'Ingd Releasing Window': 'RM/PM Releasing (Issue)', 'LPO Window': 'LPO — Purchase Order', 'LPO Window NAV': 'LPO (NAV items)', 'Item Window': 'Items Master', 'Vehicle Window': 'Vehicles', 'Vehicle Repairs Window': 'Vehicle Repairs', 'GL Transection Posting Window': 'GL Journal Posting', 'GL Transection Posting Window Multi': 'GL Journal Posting (multi)', 'CompanyInfo': 'Company Info', 'LoginUserList_Window_App': 'Login Users (SharePoint)', 'PL Window': 'Packing List', 'PL List': 'Packing Lists', 'Clients Purchase Orders': 'Customer Purchase Orders', 'Vehicle Inspection Window_New': 'Vehicle Inspection (new)', 'ItemsLotList': 'RM/PM Documents List', 'MOM Window': 'Minutes of Meeting' };
const ICONS = [[/lpo|\bpo\b|purchase|vendor/i, '🧾'], [/repair|workshop|defect/i, '🔧'], [/inspection/i, '🔍'], [/petrol|fuel/i, '⛽'], [/milg|meter/i, '🧭'], [/vehicle|fleet|truck|brand|model/i, '🚛'], [/attend|shift/i, '⏱'], [/login|user|rights|credential/i, '🔑'], [/employee|staff|profile|helper|driver/i, '👷'], [/budget|gl|currancy|cash|payment|pay tran|invoice/i, '💰'], [/customer|client|nav ?card/i, '👥'], [/trip|deliver|pick|\bpl\b|pl data|incentive/i, '🚚'], [/receiving/i, '📥'], [/releasing|release/i, '📤'], [/lot|expiry/i, '⏳'], [/uom|group/i, '📐'], [/item|stock|inventory/i, '📦'], [/machine/i, '⚙'], [/import/i, '🔗'], [/fg|prod|bom|rm |rp items|rpm/i, '🏭'], [/plan|deploy|p&c|order|mrp/i, '📈'], [/company/i, '🏢'], [/mom|task|facilit/i, '🗒'], [/list|search/i, '📋']];
const iconFor = s => (ICONS.find(([rx]) => rx.test(s)) || [0, '📄'])[1];
const humanize = s => LABELS[s] || s.replace(/_/g, ' ').replace(/\b(Window|Wndw|tbl|frm|DS)\b/gi, ' ').replace(/([a-z])([A-Z][a-z])/g, '$1 $2').replace(/\s+/g, ' ').trim() || s;
/* page-specific filters & defaults (what the Access form's own filter / default values did) */
const PAGE_RULE = {
  'Ingd Receiving Window': { where: r => !!r.Receive, defaults: () => ({ Receive: true, DocumentDate: today(), ReceivingDate: today() }), noun: 'receiving document' },
  'Ingd Releasing Window': { where: r => !r.Receive, defaults: () => ({ Receive: false, DocumentDate: today() }), noun: 'releasing document' },
  'LPO Pending For Approval': { where: r => !r.ApproveLPO && !r.CancelPO },
  'LPOListRmndr': { where: r => r.ApproveLPO && !r.LPOSent && !r.CancelPO },
  'LPO Window': { defaults: () => ({ 'LPO Date': today(), Currancy: 'AED', 'Prepared By': me() }), noun: 'LPO' },
  'VehicleRepairsApproval': { where: r => !r.Approved },
  'Vehicle Repairs Window': { defaults: () => ({ 'WS Doc Date': today(), Ticket_Case_Status: 'Open' }), noun: 'repair' },
  'Vehicle History Posted records': { where: r => !!r.Posted }, 'Vehicle Events Window Posted': { where: r => !!r.Posted },
};
const AUDIT_RX = /^(createdby|createby|recordgntrdby|timestamp|creationtimestamp|creatingtimestamp|creation_timestamp|creatingdate|lstmodifedby|modifiedby|lastmodftimestamp|modificationtimestamp|postedby|postedstamp|upsize.*)$/i;
const isEditableSub = (P, S) => /relation|link|hdr|header|batch|pinvid/i.test(S.link.c) || (LISTS[P.table] && S.link.p === LISTS[P.table].pk);
const linkVal = v => v === null || v === undefined || v === '' ? null : String(v);

function genericRender(P) {
  return async (v, param) => {
    const k = P.table, L = LISTS[k], R = PAGE_RULE[P.form] || {};
    const subs = P.subs.filter(S => LISTS[S.table]);
    await loadAll([k, ...subs.map(S => S.table), ...(k === 'IngdMvmntHdr' ? ['LOTTable', 'ItemTable'] : [])]);
    const pf = [...new Set(P.fields)].map(n => L.byName[n]).filter(f => f && !f.missing);
    const main = (pf.length ? pf : L.fields.filter(f => !f.missing)).map(f => f.n);
    const rest = L.fields.filter(f => !f.missing && !main.includes(f.n)).map(f => f.n);
    const visible = main.map(n => L.byName[n]).filter(f => f.t !== 'm' && !AUDIT_RX.test(f.n));
    const colF = [...(L.pk && L.byName[L.pk] ? [L.byName[L.pk]] : []), ...visible.filter(f => f.n !== L.pk && !/id$|link$/i.test(f.n)), ...visible.filter(f => f.n !== L.pk && /id$|link$/i.test(f.n))].slice(0, 9);
    const cols = colF.map(f => ({ f: f.n, l: f.l, t: f.t, sum: f.t === 'c' ? 1 : 0, html: /status$|^status/i.test(f.n) ? r => r[f.n] ? badgeFor(r[f.n]) : '' : null }));
    const rows0 = DB.peek(k);
    const filters = L.fields.filter(f => !f.missing && ['t', 'ch'].includes(f.t) && /status|type|region|regoin|category|department|dept|location|whs|level|group/i.test(f.n) && main.includes(f.n)).map(f => ({ f: f.n, l: f.l, opts: uniq(rows0.map(r => r[f.n])).slice(0, 60) })).filter(x => x.opts.length > 1 && x.opts.length <= 60).slice(0, 3);
    const totalF = subs.length ? L.fields.find(f => !f.missing && f.t === 'c' && HDR_TOTAL_RX.test(f.n)) : null;
    const dateF = L.fields.find(f => f.t === 'd' && main.includes(f.n));
    let ED = null;   // line editors of the open record
    const lp = listPage(v, {
      k, noun: R.noun || 'record', cols, filters, where: R.where, sortBy: L.pk || (dateF && dateF.n), desc: true,
      search: L.fields.filter(f => !f.missing && ['t', 'ch'].includes(f.t)).map(f => f.n),
      tabs: [{ title: 'Main', sections: [{ fields: main }] }, ...(rest.length ? [{ title: `All fields (${rest.length})`, sections: [{ fields: rest }] }] : [])],
      ro: L.fields.filter(f => f.ro || AUDIT_RX.test(f.n)).map(f => f.n), calc: totalF ? [totalF.n] : [],
      defaults: R.defaults,
      extra: subs.length ? (host, rec) => { let first = true; ED = subs.map(S => { const hasAmt = !!findF(S.table, AMT_RX) && isEditableSub(P, S); const e = lineEditor(host, P, S, rec, hasAmt && first ? totalF : null); if (hasAmt) first = false; return e; }); } : null,
      before: async (out, isNew) => { stampAudit(k, out, isNew); if (totalF && ED) { const e = ED.find(x => x.amounts && x.drivesTotal); if (e) out[totalF.n] = r2(sum(e.read(), l => lineAmount(e.S.table, l))); } },
      after: async (saved, isNew) => { if (ED) for (const e of ED) await e.save(saved); Activity.quiet || null; },
      onDelete: async rec => { for (const S of subs.filter(S => isEditableSub(P, S))) { const pv = linkVal(rec[S.link.p]); if (pv === null) continue; const ls = DB.peek(S.table).filter(l => linkVal(l[S.link.c]) === pv); if (ls.length && await confirmBox(`Also delete the ${ls.length} line(s) of this record (${esc(S.table)})?`, 'Delete lines', true)) await DB.bulk(ls.map(l => () => DB.remove(S.table, l))); } },
      buttons: rec => [
        ...(rec ? [{ label: '🖨 Print', cls: 'ghost', onClick: () => printRecord(P, DB.peek(k).find(x => x._id === rec._id) || rec) }] : []),
        ...(rec ? (ACTIONS[k] || []).filter(a => a.when(rec) && (!a.admin || Auth.isAdmin())).map(a => ({ label: a.label, cls: a.cls, onClick: async () => {
          try { const fresh = { ...(DB.peek(k).find(x => x._id === rec._id) || rec) }; const ctx = { lines: t => { const S = subs.find(s => s.table === t); if (!S) return []; const pv = linkVal(fresh[S.link.p]); return pv === null ? [] : DB.peek(t).filter(l => linkVal(l[S.link.c]) === pv).map(l => ({ ...l })); } };
            const m = await a.run(fresh, ctx); if (m === null) return; toast(m, 'success', 6000); Activity.add('📌', `${humanize(P.form)}: ${m}`); Modal.close(); lp.refresh(); } catch (e) { toast(e.message, 'error', 7000); } } })) : [])],
    });
    const tb = $('.toolbar', v);
    tb.insertAdjacentHTML('beforeend', `<button class="ghost small" data-pl title="Print this list">🖨 Print</button>`);
    $('[data-pl]', v).onclick = () => printTable(humanize(P.form), lp.rows(), cols.map(c => ({ ...c, html: null })));
    if (R.where) tb.insertAdjacentHTML('afterbegin', `<span class="badge info" title="Same filter as the Access form">filtered view</span>`);
    lp.refresh();
    if (param === 'new') lp.edit(null);
    else if (param && param.startsWith('rec:')) { const r = DB.peek(k).find(x => String(x._id) === param.slice(4)); if (r) lp.edit(r); }
  };
}

/* ---------- line grid (Access subform) ---------- */
function lineEditor(host, P, S, rec, totalF) {
  const SL = LISTS[S.table], editable = isEditableSub(P, S), pv = linkVal(rec[S.link.p]);
  let lines = pv === null ? [] : DB.peek(S.table).filter(l => linkVal(l[S.link.c]) === pv).map(l => ({ ...l }));
  lines.sort((a, b) => num(a.LineNr ?? a[SL.pk] ?? a._id) - num(b.LineNr ?? b[SL.pk] ?? b._id));
  const removed = [];
  let shown = [...new Set(S.fields)].map(n => SL.byName[n]).filter(f => f && !f.missing && f.n !== S.link.c && f.n !== SL.pk && (f.t !== 'm' || /desc|detail|name/i.test(f.n)));
  if (!shown.length) shown = SL.fields.filter(f => !f.missing && f.n !== S.link.c && f.n !== SL.pk && (f.t !== 'm' || /desc|detail|name/i.test(f.n)));
  const pri = f => /^item.?desc|^item description$|^itemdetails$/i.test(f.n) ? 0 : /^item.?(id|code)$/i.test(f.n) ? 1 : 2;
  shown = shown.map((f, i) => [f, i]).sort((a, b) => pri(a[0]) - pri(b[0]) || a[1] - b[1]).map(x => x[0]).slice(0, 11);
  const amtF = findF(S.table, AMT_RX), qF = findF(S.table, QTY_RX), pF = findF(S.table, PRICE_RX), autoAmt = amtF && qF && pF;
  const fefo = S.table === 'IngdMvmntData';
  const box = document.createElement('div'); box.className = 'lines'; host.appendChild(box);
  const body = host.closest('.modal-body') || document;
  const totals = () => { if (!amtF) return; const t = r2(sum(read(), l => lineAmount(S.table, l))); const el = $('[data-lt]', box); if (el) el.textContent = money(t); if (totalF) setVal(body, totalF.n, t); };
  const hint = l => { const it = itemKeyOf(l); if (!it) return ''; const ls = usableLots(it), bal = sum(ls, 'BalanceQty'), n = ls[0]; return `<span class="badge ${bal > 0 ? 'ok' : 'err'}" title="Usable (non-expired) LOT balance">${intf(bal)}</span>${n && n.ExpiryDate ? ` <span class="muted">FEFO ${fdate(n.ExpiryDate)}</span>` : ''}`; };
  function draw() {
    box.innerHTML = `<div class="sec-title" style="display:flex;align-items:center;gap:10px">${esc(humanize(S.form))} <span class="muted" style="font-weight:400">${lines.length} line(s) · list ${esc(S.table)}</span>${editable ? '' : ' <span class="badge">read-only history</span>'}<span class="spacer" style="flex:1"></span>${amtF ? `<span class="muted">Total <b data-lt>0.00</b></span>` : ''}</div>
      <div class="tbl-wrap" style="max-height:340px"><table class="dt lines-t"><thead><tr><th style="width:26px">#</th>${shown.map(f => `<th class="${['c', 'n'].includes(f.t) ? 'num' : ''}">${esc(f.l)}</th>`).join('')}${fefo ? '<th>Stock / FEFO</th>' : ''}${editable ? '<th style="width:30px"></th>' : ''}</tr></thead><tbody>${lines.map((l, i) => `<tr data-li="${i}"><td class="muted">${i + 1}</td>${shown.map(f => `<td>${editable ? inputFor(S.table, f.t === 'm' ? { ...f, t: 't' } : f, l[f.n], { ro: autoAmt && f === amtF, calc: autoAmt && f === amtF, noSuggest: !['t', 'ch'].includes(f.t) }).replace(/data-f=/g, 'data-lf=') : fmtCell({ f: f.n, t: f.t }, l)}</td>`).join('')}${fefo ? `<td data-h>${hint(l)}</td>` : ''}${editable ? `<td><button type="button" class="x-btn" data-rm title="Remove line">✖</button></td>` : ''}</tr>`).join('') || `<tr><td colspan="${shown.length + 3}" class="empty">No lines yet.</td></tr>`}</tbody></table></div>
      ${editable ? `<div class="btn-row" style="margin-top:8px"><button type="button" class="secondary small" data-add>＋ Add line</button>${lines.length ? '<button type="button" class="ghost small" data-dup>⧉ Copy last line</button>' : ''}<span class="muted">${autoAmt ? `${esc(amtF.l)} = ${esc(qF.l)} × ${esc(pF.l)} — calculated` : ''}</span></div>` : ''}`;
    $$('tr[data-li]', box).forEach(tr => {
      const i = +tr.dataset.li;
      tr.oninput = tr.onchange = () => { readRow(tr, lines[i]); if (autoAmt) { calcLine(S.table, lines[i]); const a = $(`[data-lf="${CSS.escape(amtF.n)}"]`, tr); if (a) a.value = r2(lines[i][amtF.n]); } if (fefo) $('[data-h]', tr).innerHTML = hint(lines[i]); totals(); };
      const rm = $('[data-rm]', tr); if (rm) rm.onclick = () => { readAll(); const [x] = lines.splice(i, 1); if (x && x._id) removed.push(x); draw(); };
    });
    const add = $('[data-add]', box); if (add) add.onclick = () => { readAll(); lines.push(newLine()); draw(); const tr = $$('tr[data-li]', box).pop(); tr && $('input,select', tr)?.focus(); };
    const dup = $('[data-dup]', box); if (dup) dup.onclick = () => { readAll(); const c = { ...lines[lines.length - 1] }; delete c._id; if (SL.pk) delete c[SL.pk]; lines.push(c); draw(); };
    totals();
  }
  const newLine = () => { const l = { LineNr: lines.length + 1 }; if (!SL.byName.LineNr) delete l.LineNr; if (fefo) { l.DocDate = getVal(body, 'DocumentDate') || today(); l.DocNumber = getVal(body, 'DocumentNo'); } return l; };
  function readRow(tr, l) { $$('[data-lf]', tr).forEach(el => { const n = el.dataset.lf; l[n] = el.type === 'checkbox' ? el.checked : el.type === 'number' ? (el.value === '' ? null : num(el.value)) : el.type === 'datetime-local' ? (el.value ? new Date(el.value).toISOString() : null) : (el.value.trim() || null); }); return l; }
  function readAll() { if (editable) $$('tr[data-li]', box).forEach(tr => readRow(tr, lines[+tr.dataset.li])); return lines; }
  const read = () => readAll().filter(l => shown.some(f => l[f.n] !== null && l[f.n] !== undefined && l[f.n] !== '' && l[f.n] !== false));
  draw();
  return {
    S, amounts: !!amtF, drivesTotal: !!totalF, read,
    async save(saved) {
      if (!editable) return;
      const p = linkVal(saved[S.link.p]); const keep = read();
      if (p === null) { if (keep.length) toast(`Lines not saved: “${S.link.p}” is empty on the record.`, 'warn', 6000); return; }
      const lf = SL.byName[S.link.c], linkV = lf && ['n', 'c'].includes(lf.t) ? num(p) : p;
      Activity.quiet = true;
      try {
        const errs = await DB.bulk([...removed.map(x => () => DB.remove(S.table, x)), ...keep.map((l, i) => () => { l[S.link.c] = linkV; if (SL.byName.LineNr && !l.LineNr) l.LineNr = i + 1; if (fefo) { l.DocNumber = l.DocNumber || saved.DocumentNo; l.DocDate = l.DocDate || saved.DocumentDate; } calcLine(S.table, l); return DB.save(S.table, l); })], null, 3);
        if (errs.length) toast(`${errs.length} line(s) not saved: ${errs[0]}`, 'error', 7000);
      } finally { Activity.quiet = false; }
    },
  };
}

/* ---------- branded record print (any form) ---------- */
function printRecord(P, rec) {
  const L = LISTS[P.table], fs = [...new Set(P.fields)].map(n => L.byName[n]).filter(f => f && rec[f.n] !== null && rec[f.n] !== undefined && rec[f.n] !== '' && f.t !== 'b' || (f && f.t === 'b' && rec[f.n]));
  let html = `<div class="box">${fs.map(f => `<div><b>${esc(f.l)}</b> ${f.t === 'm' ? esc(rec[f.n]).replace(/\n/g, '<br>') : fmtCell({ f: f.n, t: f.t }, rec)}</div>`).join('')}</div>`;
  for (const S of P.subs.filter(S => LISTS[S.table])) {
    const SL = LISTS[S.table], pv = linkVal(rec[S.link.p]); if (pv === null) continue;
    const ls = DB.peek(S.table).filter(l => linkVal(l[S.link.c]) === pv); if (!ls.length) continue;
    const cs = [...new Set(S.fields)].map(n => SL.byName[n]).filter(f => f && f.n !== S.link.c && (f.t !== 'm' || /desc|detail|name/i.test(f.n)) && ls.some(l => l[f.n] !== null && l[f.n] !== undefined && l[f.n] !== '')).slice(0, 9);
    const amt = findF(S.table, AMT_RX);
    html += `<h3 style="margin:16px 0 4px;font-size:13px">${esc(humanize(S.form))}</h3><table><thead><tr><th>#</th>${cs.map(f => `<th class="${['c', 'n'].includes(f.t) ? 'n' : ''}">${esc(f.l)}</th>`).join('')}</tr></thead><tbody>${ls.map((l, i) => `<tr><td>${i + 1}</td>${cs.map(f => `<td class="${['c', 'n'].includes(f.t) ? 'n' : ''}">${fmtCell({ f: f.n, t: f.t }, l)}</td>`).join('')}</tr>`).join('')}</tbody>${amt && cs.includes(amt) ? `<tfoot><tr><td></td>${cs.map(f => `<td class="n">${f === amt ? money(sum(ls, f.n)) : ''}</td>`).join('')}</tr></tfoot>` : ''}</table>`;
    if (amt) { const t = sum(ls, amt.n); html += `<div class="tot"><div class="g"><span>Total</span><span>${money(t)}</span></div></div><div class="words">${esc(amountInWords(t))}</div>`; }
  }
  html += `<div class="sig"><div>Prepared by</div><div>Checked by</div><div>Approved by</div></div>`;
  const no = rec[DOCNO[P.table]?.[0]] || rec[L.pk] || '';
  printDoc(humanize(P.form), html, { no, meta: [['Document', no], ['Date', fdate(rec[(L.fields.find(f => f.t === 'd' && rec[f.n]) || {}).n])], ['Printed by', me()]] });
}

/* ---------- register one page per Access form, grouped like the Access menu ---------- */
(() => {
  const order = NFPC.groups.map(g => g[0]), used = new Set(PAGES.map(p => p.id));
  const gen = NFPC.pages.filter(p => LISTS[p.table]).map(p => ({ ...p, label: humanize(p.form) }));
  gen.sort((a, b) => (order.indexOf(a.group) - order.indexOf(b.group)) || ((b.subs.length > 0) - (a.subs.length > 0)) || a.label.localeCompare(b.label));
  const custom = PAGES.splice(0);   // dashboard + custom pages first, then each group's own pages
  PAGES.push(...custom.filter(p => p.sec === 'Main'));
  for (const g of order) {
    PAGES.push(...custom.filter(p => p.sec === g));
    for (const p of gen.filter(x => x.group === g)) { let id = p.id; while (used.has(id)) id += '-2'; used.add(id); PAGES.push({ id, sec: g, icon: iconFor(p.form), label: p.label, sub: `Access form “${p.form}” · list ${LISTS[p.table].name}${LISTS[p.table].existing ? ' (existing)' : ''}`, render: genericRender(p), form: p.form, gen: p }); }
  }
  PAGES.push(...custom.filter(p => p.sec !== 'Main' && !order.includes(p.sec)));
})();
