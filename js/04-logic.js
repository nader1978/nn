/* =====================================================================
   NFPC BUSINESS RULES — smart behaviour added on top of the generic Access-form pages
   ===================================================================== */
const has = (k, n) => !!(LISTS[k] && LISTS[k].byName[n] && !LISTS[k].byName[n].missing);
const findF = (k, rx) => (LISTS[k] ? LISTS[k].fields.find(f => !f.missing && rx.test(f.n)) : null);
const QTY_RX = /^(qty|quantity|orderqty|order_?qty|ordered_?qty|ordered_?quantity|receivedqty|releasedqty|rcvdqty|trnsction_qty|qty_ordered)$/i;
const PRICE_RX = /^(unit ?price|unit_?price|price|unit_?cost|unitcostprice|unitpurprice|unitsalesprice)$/i;
const AMT_RX = /^(amount|line_?amount|lineamount|net_?amount|netamount|total_?amount|purlineamount|saleslineamount)$/i;
const DISC_RX = /^discount%?$/i;
const HDR_TOTAL_RX = /^(lpoamount|rcrdtotalamount|cpoamount|total_?amount|totalamount|invoice_?amount|amount)$/i;
const AUDIT_NEW = { CreatedBy: () => me(), RecordGntrdBy: () => me(), Createdby: () => me(), CreateBy: () => me(), 'Prepared By': () => me(), TimeStamp: () => nowIso(), CreationTimeStamp: () => nowIso(), CreatingTimeStamp: () => nowIso(), Creation_Timestamp: () => nowIso(), CreatingDate: () => today() };
const AUDIT_EDIT = { LstModifedBy: () => me(), Modifiedby: () => me(), LastModfTimeStamp: () => nowIso(), ModificationTimeStamp: () => nowIso() };
const DOCNO = { 'IngdMvmntHdr': ['DocumentNo', r => (r.Receive ? 'RCV' : 'REL')], 'LPOHeader': ['LPO Number', () => 'LPO'], 'PL Header': ['PL Number', () => 'PL'], 'CustomerPOHeader_tbl': ['CPONumber', () => 'CPO'], 'GLBatchHeader': ['GLBatchName', () => 'GLB'], 'RPM_Batch_Header_Table': ['RPM_Batch_Nr', () => 'RPM'], 'MOM_HDR_tbl': ['MOM_Nr', () => 'MOM'] };

/* line maths: Amount = Qty × Unit price (− discount %) */
function calcLine(k, l) {
  const q = findF(k, QTY_RX), p = findF(k, PRICE_RX), a = findF(k, AMT_RX), d = findF(k, DISC_RX);
  if (q && p && a) { const disc = d ? num(l[d.n]) : 0; l[a.n] = r2(num(l[q.n]) * num(l[p.n]) * (1 - disc / 100)); }
  return l;
}
function lineAmount(k, l) { const a = findF(k, AMT_RX); return a ? num(l[a.n]) : 0; }
function stampAudit(k, rec, isNew) {
  for (const [n, fn] of Object.entries(isNew ? AUDIT_NEW : {})) if (has(k, n) && !rec[n]) rec[n] = fn();
  for (const [n, fn] of Object.entries(AUDIT_EDIT)) if (has(k, n)) rec[n] = fn();
  const dn = DOCNO[k]; if (dn && has(k, dn[0]) && !rec[dn[0]]) { const yy = today().slice(2, 7).replace('-', ''); const used = DB.peek(k).map(x => String(x[dn[0]] || '')).filter(x => x.startsWith(dn[1](rec) + '-' + yy)).map(x => num(x.split('-').pop())); rec[dn[0]] = `${dn[1](rec)}-${yy}-${String(Math.max(0, ...used) + 1).padStart(4, '0')}`; }
  return rec;
}

/* ---------- Warehouse: stock, LOT/FEFO ---------- */
const itemKeyOf = r => String(r.ItemDesc || r['Item Description'] || r.ItemID || '').trim();
function lotsFEFO(item) { return DB.peek('LOTTable').filter(l => norm(l.ItemDesc) === norm(item) && num(l.BalanceQty) > 0 && !l.CloseLot).sort((a, b) => String(a.ExpiryDate || '9999').localeCompare(String(b.ExpiryDate || '9999'))); }
const usableLots = item => lotsFEFO(item).filter(l => !l.ExpiryDate || String(l.ExpiryDate).slice(0, 10) >= today());   // FEFO never issues expired LOTs
function stockOf(item) { return sum(lotsFEFO(item), 'BalanceQty'); }
function stockRows() {
  const items = DB.peek('ItemTable'), mv = DB.peek('IngdMvmntData');
  return items.map(it => {
    const key = it.ItemDesc, m = mv.filter(x => norm(x.ItemDesc) === norm(key) || (it.ItemID && num(x.ItemID) === num(it.ItemID)));
    const rec = sum(m, x => x.ReceivedQtyBase ?? x.ReceivedQty), rel = sum(m, x => x.ReleasedQtyBase ?? x.ReleasedQty);
    const lots = lotsFEFO(key), lotBal = sum(lots, 'BalanceQty'), next = lots[0];
    const bal = lots.length ? lotBal : (num(it.AvalibleStock) || rec - rel), min = num(it['MinimumQtyLevel-Pkt'] ?? it.MinimumQtyLevelPkt);
    return { ...it, _rec: rec, _rel: rel, _bal: bal, _min: min, _lots: lots.length, _next: next ? next.ExpiryDate : null, _days: next && next.ExpiryDate ? daysBetween(today(), next.ExpiryDate) : null, _status: bal <= 0 ? 'Out of stock' : min && bal < min ? 'Below minimum' : 'OK' };
  });
}
/* Post a receiving / releasing document: receive → new LOT, release → FEFO from the oldest-expiring LOTs */
async function postMovement(h, lines) {
  await Promise.all(['LOTTable', 'ItemTable'].map(k => DB.all(k).catch(() => [])));
  const jobs = [], notes = [];
  for (const l of lines) {
    const item = itemKeyOf(l), rq = num(l.ReceivedQtyBase ?? l.ReceivedQty) || num(l.ReceivedQty), xq = num(l.ReleasedQtyBase ?? l.ReleasedQty) || num(l.ReleasedQty);
    if (!item) continue;
    if (rq > 0) jobs.push(() => DB.save('LOTTable', { ItemDesc: item, ExpiryDate: l.ExpiryDate || null, ReceivedQty: rq, ReleasedQty: 0, BalanceQty: rq, ShipmentRef: h.ContainerFIRS || h.DocumentNo, SupplierName: h.VendorName, RcvShpHdrNr: String(h.headerID ?? ''), RcvShpLineNr: String(l.MvmntID ?? ''), RcvShpDate: h.ReceivingDate || h.DocumentDate || today(), Rcvby: me(), RcvTimeStamp: nowIso(), headerID: h.headerID, PostLot: true, DocumentDate: h.DocumentDate }));
    if (xq > 0) {
      let left = xq;
      for (const lot of usableLots(item)) { if (left <= 0) break; const take = Math.min(left, num(lot.BalanceQty)); left = r2(left - take); const L = { ...lot, ReleasedQty: r2(num(lot.ReleasedQty) + take), BalanceQty: r2(num(lot.BalanceQty) - take), CustomerName: h.CustomerName || lot.CustomerName, RlsDNHdrNr: String(h.headerID ?? ''), RlsDNDate: h.DocumentDate || today(), Rlsdby: me(), RlsTimeStamp: nowIso() }; if (L.BalanceQty <= 0) L.CloseLot = true; jobs.push(() => DB.save('LOTTable', L)); if (!l.ExpiryDate) l.ExpiryDate = lot.ExpiryDate; }
      if (left > 0) notes.push(`${item}: ${left} not covered by non-expired LOTs`);
    }
    const it = DB.peek('ItemTable').find(x => norm(x.ItemDesc) === norm(item)); if (it && has('ItemTable', 'AvalibleStock')) jobs.push(() => DB.save('ItemTable', { ...it, AvalibleStock: r2(num(it.AvalibleStock) + rq - xq) }));
  }
  const errs = await DB.bulk(jobs, null, 3); if (errs.length) throw new Error(errs[0]);
  return notes;
}

/* ---------- Fleet ---------- */
const VEH_EXP = [['InsuranceExpiryDate', 'Insurance'], ['Reg_License_Expry_Date', 'Registration'], ['InsuLicenseExpryDate', 'Insurance licence'], ['DateofNextOilChngApprox', 'Oil change'], ['DateofNextTiresChngApprox', 'Tyres'], ['DateofNextBattaryChngApprox', 'Battery']];
function fleetAlerts(days = 30) { const out = []; for (const v of DB.peek('Vehicles_tbl')) for (const [f, l] of VEH_EXP) if (v[f]) { const d = daysBetween(today(), v[f]); if (d <= days) out.push({ plate: v.Vehicle_Plate_No, what: l, date: v[f], days: d, v }); } return out.sort((a, b) => a.days - b.days); }
const EMP_EXP = [['VisaExpiryDate', 'Visa'], ['PassPortExpiryDate', 'Passport'], ['EIDExpiryDate', 'Emirates ID'], ['DLExpiryDate', 'Driving licence'], ['LaborCardExpiryDate', 'Labour card'], ['HealthCardExpiryDate', 'Health card']];
function staffAlerts(days = 30) { const out = []; for (const e of DB.peek('Employees')) for (const [f, l] of EMP_EXP) if (e[f]) { const d = daysBetween(today(), e[f]); if (d <= days) out.push({ name: [e.FirstName, e.LastName].filter(Boolean).join(' '), what: l, date: e[f], days: d }); } return out.sort((a, b) => a.days - b.days); }

/* ---------- action buttons per table (shown in the record window) ---------- */
const ACTIONS = {
  IngdMvmntHdr: [{ label: '📌 Post & update stock', cls: 'success', when: r => r._id && !r.Post, run: async (r, ctx) => { const notes = await postMovement(r, ctx.lines('IngdMvmntData')); Object.assign(r, { Post: true, PostedBy: me(), PostedStamp: nowIso() }); await DB.save('IngdMvmntHdr', r); return 'Posted — stock and LOT balances updated (FEFO).' + (notes.length ? ' ⚠ ' + notes.join('; ') : ''); } }],
  LPOHeader: [
    { label: '✔ Approve LPO', cls: 'success', when: r => r._id && !r.ApproveLPO && !r.CancelPO, admin: true, run: async r => { Object.assign(r, { ApproveLPO: true, POApprovedByName: me(), POApprovedTimeStamp: nowIso() }); await DB.save('LPOHeader', r); return 'LPO approved.'; } },
    { label: '📧 Mark as sent', cls: 'secondary', when: r => r._id && r.ApproveLPO && !r.LPOSent, run: async r => { Object.assign(r, { LPOSent: true, DateOfSendingEmail: today() }); await DB.save('LPOHeader', r); return 'Marked as sent to the supplier.'; } },
    { label: '✖ Cancel LPO', cls: 'danger', when: r => r._id && !r.CancelPO, run: async r => { if (!(await confirmBox('Cancel this LPO?', 'Cancel LPO', true))) return null; Object.assign(r, { CancelPO: true, CancelledBy: me(), CancellationStamp: nowIso() }); await DB.save('LPOHeader', r); return 'LPO cancelled.'; } }],
  Vehicles_Defect_Reg_tbl: [{ label: '✔ Approve repair', cls: 'success', when: r => r._id && !r.Approved, admin: true, run: async r => { r.Approved = true; await DB.save('Vehicles_Defect_Reg_tbl', r); return 'Repair approved.'; } }, { label: '📌 Post (completed)', cls: 'secondary', when: r => r._id && r.Approved && !r.Post, run: async r => { r.Post = true; r.Ticket_Case_Status = r.Ticket_Case_Status || 'Closed'; await DB.save('Vehicles_Defect_Reg_tbl', r); return 'Repair posted.'; } }],
  'Vehicle History records': [{ label: '📌 Post inspection', cls: 'success', when: r => r._id && !r.Posted, run: async r => { r.Posted = true; await DB.save('Vehicle History records', r); return 'Inspection posted.'; } }],
  GLBatchHeader: [{ label: '📌 Post batch', cls: 'success', when: r => r._id, admin: true, run: async (r, ctx) => { const L = ctx.lines('GLTransectionsTable'); const dr = sum(L, 'DebitAmount'), cr = sum(L, 'CreditAmount'); if (Math.abs(dr - cr) > 0.005 && !(await confirmBox(`Debit ${money(dr)} ≠ credit ${money(cr)}. Post anyway?`, 'Post'))) return null; await DB.bulk(L.filter(l => l._id).map(l => () => DB.save('GLTransectionsTable', { ...l, PostTransection: 'Yes', PostedBy: me(), PostingTimeStamp: nowIso() }))); return `Batch posted — ${L.length} line(s).`; } }],
};

/* =====================================================================
   DEMO DATA (Demo mode only) — realistic rows for the main processes, sample rows elsewhere
   ===================================================================== */
const SEED = (() => {
  let built = null;
  function sampleVal(f, i) {
    const n = f.n, l = n.toLowerCase();
    switch (f.t) {
      case 'd': return addDays(today(), (i * 37 % 120) - 60);
      case 'dt': return new Date(Date.now() - i * 86400000 * 3).toISOString();
      case 'b': return i % 3 === 0;
      case 'n': return /id$|link|relation/i.test(n) ? i + 1 : (i * 7 % 50) + 1;
      case 'c': return r2(((i * 137) % 900) + 50);
      case 'm': return '';
      default:
        if (/region|location|whs?_?desc/.test(l)) return ['Dubai', 'Abu Dhabi', 'Sharjah', 'Al Ain', 'Muscat'][i % 5];
        if (/status/.test(l)) return ['Open', 'In progress', 'Closed'][i % 3];
        if (/name|desc|details|title/.test(l)) return `${n.replace(/_/g, ' ')} ${i + 1}`;
        if (/code|nr|no$|number/.test(l)) return `${n.slice(0, 3).toUpperCase()}-${1000 + i}`;
        return `${n.replace(/_/g, ' ').slice(0, 18)} ${i + 1}`;
    }
  }
  function generic(k, n = 5) { return Array.from({ length: n }, (_, i) => Object.fromEntries(LISTS[k].fields.filter(f => !f.extra).map(f => [f.n, sampleVal(f, i)]))); }
  function build() {
    if (built) return built;
    const B = {};
    for (const k of Object.keys(LISTS)) B[k] = generic(k, 4);
    const items = [['Milk Powder 25kg', 'RM', 'BAG', 20], ['Sugar 50kg', 'RM', 'BAG', 30], ['PET Bottle 500ml', 'PM', 'PCS', 2000], ['Carton Box 24x500', 'PM', 'PCS', 1500], ['Strawberry Flavour', 'RM', 'KG', 10], ['Label 500ml Laban', 'PM', 'ROLL', 15], ['Cap 38mm White', 'PM', 'PCS', 1500], ['Stabiliser E471', 'RM', 'KG', 8]];
    B.ItemTable = items.map(([d, c, u, m], i) => ({ ItemID: i + 1, ItemDesc: d, ItemCode: `IT-${1001 + i}`, ItemCategory: c, ItemType: c, UOMID: u, 'MinimumQtyLevel-Pkt': m, AvalibleStock: 0, CostPrice: r2(5 + i * 3.5), SalesPrice: r2(8 + i * 4), ItemLocation: ['Dubai WH', 'Al Ain WH'][i % 2], LotNumberTrack: 'Yes', CreateBy: 'admin', CreatingDate: addDays(today(), -200) }));
    B.LOTTable = []; B.IngdMvmntHdr = []; B.IngdMvmntData = []; let lot = 1, hid = 100, mid = 1000;
    for (let d = 0; d < 10; d++) {
      const rcv = d % 3 !== 2, h = { headerID: ++hid, DocumentNo: `${rcv ? 'RCV' : 'REL'}-${today().slice(2, 7).replace('-', '')}-${String(d + 1).padStart(4, '0')}`, DocumentDate: addDays(today(), -60 + d * 6), ReceivingDate: addDays(today(), -60 + d * 6), Receive: rcv, Post: d < 8, VendorName: rcv ? ['Almarai Supplies', 'Gulf Packaging LLC', 'Emirates Sugar'][d % 3] : null, CustomerName: rcv ? null : 'Production Line ' + (d % 2 + 1), ContainerFIRS: rcv ? `CONT-${7700 + d}` : null, CreatedBy: 'admin', TimeStamp: new Date(Date.now() - (60 - d * 6) * 864e5).toISOString(), PostedBy: d < 8 ? 'admin' : null };
      B.IngdMvmntHdr.push(h);
      for (let j = 0; j < 3; j++) { const it = B.ItemTable[(d * 2 + j) % B.ItemTable.length]; let q = 10 + ((d * 7 + j * 5) % 40) * (it.UOMID === 'PCS' ? 100 : 1); const exp = addDays(today(), 20 + ((d * 31 + j * 17) % 300) - 40);
        if (!rcv) { const have = B.LOTTable.filter(x => x.ItemDesc === it.ItemDesc).reduce((s, x) => s + x.BalanceQty, 0); q = Math.min(Math.round(q / 3), have); if (q <= 0) continue; let left = h.Post ? q : 0; for (const lt of B.LOTTable.filter(x => x.ItemDesc === it.ItemDesc && x.BalanceQty > 0).sort((a, b) => a.ExpiryDate.localeCompare(b.ExpiryDate))) { if (left <= 0) break; const t = Math.min(left, lt.BalanceQty); lt.BalanceQty -= t; lt.ReleasedQty += t; left -= t; } }
        const line = { MvmntID: ++mid, Relation: h.headerID, ItemID: it.ItemID, ItemDesc: it.ItemDesc, ItemUOMRcd: it.UOMID, DocDate: h.DocumentDate, DocNumber: h.DocumentNo, ReceivedQty: rcv ? q : 0, ReleasedQty: rcv ? 0 : q, ReceivedQtyBase: rcv ? q : 0, ReleasedQtyBase: rcv ? 0 : q, ExpiryDate: exp, UnitPurPrice: it.CostPrice, PurLineAmount: r2(q * it.CostPrice), ItemLocation: it.ItemLocation };
        B.IngdMvmntData.push(line);
        if (rcv && h.Post) B.LOTTable.push({ LOTID: lot++, ItemDesc: it.ItemDesc, ExpiryDate: exp, ReceivedQty: q, ReleasedQty: 0, BalanceQty: q, ShipmentRef: h.ContainerFIRS, SupplierName: h.VendorName, RcvShpHdrNr: String(h.headerID), RcvShpDate: h.ReceivingDate, Rcvby: 'admin', PostLot: true, headerID: h.headerID }); }
    }
    for (const lt of B.LOTTable) if (lt.BalanceQty <= 0) lt.CloseLot = true;
    { const o = B.LOTTable.filter(x => x.BalanceQty > 0); if (o.length > 3 && !o.some(x => x.ExpiryDate < today())) o[o.length - 1].ExpiryDate = addDays(today(), -6); }
    for (const it of B.ItemTable) it.AvalibleStock = B.LOTTable.filter(x => x.ItemDesc === it.ItemDesc).reduce((s, x) => s + x.BalanceQty, 0);
    const plates = ['D 12345', 'D 22871', 'AD 5521', 'SHJ 7781', 'D 99012', 'AA 4410'];
    B.Vehicles_tbl = plates.map((p, i) => ({ Vehicle_ID: i + 1, Vehicle_Plate_No: p, VehicleType: ['Reefer truck', 'Van', 'Pickup', 'Trailer'][i % 4], VehicleModel: ['Isuzu NPR', 'Mitsubishi Canter', 'Toyota Hilux', 'Volvo FH'][i % 4], Make: ['Isuzu', 'Mitsubishi', 'Toyota', 'Volvo'][i % 4], ModelYear: String(2019 + i % 5), Asset_Regoin: ['Dubai', 'Abu Dhabi', 'Al Ain'][i % 3], 'Staff in charge': ['Ravi', 'Imran', 'Joseph', 'Ali', 'Sunil', 'Omar'][i], InsuranceExpiryDate: addDays(today(), [12, 80, -3, 200, 25, 300][i]), Reg_License_Expry_Date: addDays(today(), [150, 20, 90, 5, 400, 60][i]), DateofNextOilChngApprox: addDays(today(), [3, 40, 9, 70, -2, 25][i]), Vehicle_Activity_Status: 'Active', Purchase_Price: 120000 + i * 15000, LastOliMtrRding: 80000 + i * 12000 }));
    B.Vehicles_Defect_Reg_tbl = plates.slice(0, 5).map((p, i) => ({ 'Vehicle Plate Number': p, Region: ['Dubai', 'Abu Dhabi', 'Al Ain'][i % 3], 'Defect Type': ['A/C', 'Brakes', 'Tyres', 'Engine', 'Body'][i], 'Defect Description': ['A/C not cooling', 'Brake pads worn', 'Rear tyre puncture', 'Engine warning light', 'Rear door hinge'][i], Workshop: ['Al Futtaim Motors', 'Quick Fix Garage'][i % 2], LaborCharges: 150 + i * 80, SparePartsAmount: 300 + i * 120, 'Repairing Cost': 450 + i * 200, Approved: i < 3, Post: i < 2, Ticket_Case_Status: i < 2 ? 'Closed' : 'Open', 'WS Doc Date': addDays(today(), -i * 9) }));
    const vendors = ['Gulf Packaging LLC', 'Almarai Supplies', 'Emirates Sugar', 'Tetra Pak Gulf'];
    B.LPOHeader = []; B.LPOData = [];
    for (let i = 0; i < 6; i++) { const h = { PINVID: i + 1, 'LPO Number': `LPO-${today().slice(2, 7).replace('-', '')}-${String(i + 1).padStart(4, '0')}`, 'LPO Date': addDays(today(), -i * 8), 'Vendor Details': vendors[i % 4], Currancy: 'AED', 'Delivery Terms': 'DAP Dubai', 'Prepared By': 'admin', ApproveLPO: i > 1, LPOSent: i > 2, ExpectedDateforDelivery: addDays(today(), 20 - i * 3), RecordGntrdBy: 'admin', TimeStamp: new Date(Date.now() - i * 8 * 864e5).toISOString() }; let tot = 0;
      for (let j = 0; j < 3; j++) { const it = B.ItemTable[(i + j) % B.ItemTable.length], q = 10 + j * 15, l = { Relation: h.PINVID, ItemID: it.ItemID, ItemCode: it.ItemCode, 'Item Description': it.ItemDesc, UOM: it.UOMID, Qty: q, 'Unit Price': it.CostPrice, Amount: r2(q * it.CostPrice), 'LPO Number': h['LPO Number'], 'LPO Date': h['LPO Date'] }; tot += l.Amount; B.LPOData.push(l); }
      h.LPOAmount = r2(tot); B.LPOHeader.push(h); }
    B.Employees = [['Nader', 'Alnajjar', 'Supply Chain Manager', 'Management'], ['Ravi', 'Kumar', 'Driver', 'Logistics'], ['Imran', 'Khan', 'Storekeeper', 'Warehouse'], ['Joseph', 'Mathew', 'Planner', 'Planning'], ['Sara', 'Ali', 'Accountant', 'Finance']].map(([a, b, c, d], i) => ({ Employee1ID: i + 1, FirstName: a, LastName: b, Designation: c, Department: d, JobRole: c, Region: 'Dubai', EmploymentStatus: 'Active', VisaExpiryDate: addDays(today(), [400, 18, 90, -5, 220][i]), PassPortExpiryDate: addDays(today(), [900, 300, 25, 700, 60][i]), EIDExpiryDate: addDays(today(), [380, 15, 100, 500, 40][i]), DLExpiryDate: addDays(today(), [600, 8, 200, 300, 900][i]) }));
    B.LoginUserList = [{ User_Name: 'admin', Title: 'Nader Alnajjar', Password: 'admin', Access_Level: 'Admin', Email_ID: 'admin@example.com', Region: 'Dubai' }, { User_Name: 'store', Title: 'Imran Khan', Password: '1234', Access_Level: 'Operator', Email_ID: 'store@example.com', Region: 'Dubai' }];
    B.SC_Deliveries_Orders_tbl = Array.from({ length: 60 }, (_, i) => { const d = addDays(today(), -(i * 5) % 180); const q = 20 + (i * 13) % 80; return { OrderNumber: `SO-${50000 + i}`, OrderDate: d, TransactionDate: d, CustomerCode: `C${100 + i % 12}`, ItemCode: `FG-${200 + i % 9}`, ItemName: ['Laban 1L', 'Milk 2L', 'Juice 250ml', 'Yoghurt 500g', 'Water 1.5L'][i % 5], OrderedQuantity: q, DeliveredQuantity: q - (i % 7 === 0 ? 5 : 0), Price: 4.5, Amount: r2(q * 4.5), NetAmount: r2(q * 4.5), Region: ['Dubai', 'Abu Dhabi', 'Al Ain', 'Sharjah'][i % 4], OrderStatus: i % 9 === 0 ? 'Returned' : 'Delivered', DeliveredByCode: ['DRV-01', 'DRV-02', 'DRV-03'][i % 3] }; });
    B.GLBatchHeader = [{ GLBatchID: 1, GLBatchName: 'GLB-2611-0001', BatchRemarks: 'October accruals' }];
    B.GLTransectionsTable = [{ GLTrnsID: 1, GLBatchID: 1, TransDate: today(), TransNr: 'JV-1', GLDebitDesc: 'Fuel expense', GLCreditDesc: 'Accrued expenses', DebitAmount: 12500, CreditAmount: 0, Dept: 'Logistics' }, { GLTrnsID: 2, GLBatchID: 1, TransDate: today(), TransNr: 'JV-1', GLDebitDesc: 'Fuel expense', GLCreditDesc: 'Accrued expenses', DebitAmount: 0, CreditAmount: 12500, Dept: 'Logistics' }];
    B.CompanyInfoNN = [{ CompanyFullName: 'National Food Products Company', ArabicName: 'الشركة الوطنية للمنتجات الغذائية', Address: 'Abu Dhabi Industrial City', City: 'Abu Dhabi', Country: 'UAE', TelephoneNumber: '+971 2 000 0000', Email: 'info@example.com', Website: 'nfpc.ae', Active: true, PLAutoNrLetter: 'PL', POAutoNrLetter: 'LPO' }];
    return (built = B);
  }
  return new Proxy({}, { get: (t, k) => () => (build()[k] || []).map(x => ({ ...x })), has: (t, k) => k in build() });
})();
const company = () => DB.peek('CompanyInfoNN').find(c => c.Active) || DB.peek('CompanyInfoNN')[0] || {};
async function loadCore() { await Promise.all(['LoginUserList', 'CompanyInfoNN'].map(k => DB.all(k).catch(e => { DB.cache[k] = DB.cache[k] || []; }))); }
