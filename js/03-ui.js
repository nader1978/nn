/* =====================================================================
   SMART FORM ENGINE — renders Access-like forms from the schema
   ===================================================================== */
/* Type-ahead like the Access "SELECT DISTINCT" combo boxes: values of the same field in this list
   and in every other loaded list that has a field with the same name. */
const SUGGEST_FROM = {};
function suggestions(k, f) {
  if (!['t', 'ch'].includes(f.t)) return [];
  const vals = [];
  for (const [lk, rows] of Object.entries(DB.cache)) { const L = LISTS[lk]; if (!L || !rows.length) continue; const ff = L.byName[f.n]; if (!ff || !(lk === k || ff.t === 't')) continue; for (const r of rows) vals.push(r[f.n]); if (vals.length > 4000) break; }
  for (const [lk, fn] of (SUGGEST_FROM[f.n] || [])) for (const r of DB.peek(lk)) vals.push(r[fn]);
  return uniq(vals).slice(0, 400);
}
let dlSeq = 0;
function inputFor(k, f, v, o = {}) {
  const ro = o.ro ? ' readonly tabindex="-1"' : '', cls = o.calc ? ' class="calc"' : '', id = o.id ? ` id="${o.id}"` : '';
  const a = `data-f="${f.n}"${id}${ro}${cls}`;
  v = v ?? '';
  switch (f.t) {
    case 'm': return `<textarea ${a} rows="${o.rows || 2}">${esc(v)}</textarea>`;
    case 'n': return `<input type="number" step="any" ${a} value="${esc(v)}">`;
    case 'c': return `<input type="number" step="0.01" ${a} value="${v === '' || v === null ? '' : r2(v)}">`;
    case 'p': return `<input type="number" step="0.01" ${a} value="${esc(v)}" placeholder="%">`;
    case 'd': return `<input type="date" ${a} value="${esc(String(v).slice(0, 10))}">`;
    case 'dt': { let s = ''; if (v) { const d = new Date(v); if (!isNaN(d)) s = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`; } return `<input type="datetime-local" ${a} value="${s}">`; }
    case 'b': return `<input type="checkbox" ${a} ${v ? 'checked' : ''} ${o.ro ? 'disabled' : ''}>`;
    case 'ch': { const opts = uniq([...(f.opts || []), ...(v ? [v] : [])]); if (o.ro) return `<input type="text" ${a} value="${esc(v)}">`; return `<select ${a}><option value=""></option>${(f.opts || []).concat(opts.filter(x => !(f.opts || []).includes(x))).map(x => `<option ${x === v ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select>`; }
    default: {
      if (f.n === 'Pword') return `<input type="password" ${a} value="" placeholder="${v ? '•••••• (leave empty to keep)' : 'Set a password'}" autocomplete="new-password">`;
      const sug = o.noSuggest || o.ro ? [] : suggestions(k, f); let dl = '';
      if (sug.length) { const dlid = 'dl' + (++dlSeq); dl = `<datalist id="${dlid}">${sug.map(x => `<option value="${esc(x)}">`).join('')}</datalist>`; return `<input type="text" ${a} list="${dlid}" value="${esc(v)}" autocomplete="off">${dl}`; }
      return `<input type="text" ${a} value="${esc(v)}">`;
    }
  }
}
function fieldBox(k, name, rec, o = {}) {
  const f = typeof name === 'object' ? name : fieldOf(k, name), v = rec ? rec[f.n] : '';
  const ro = (o.ro || []).includes(f.n), calc = (o.calc || []).includes(f.n), req = (o.req || []).includes(f.n);
  const span = (o.wide || []).includes(f.n) || f.t === 'm' ? ' span2' : '';
  if (f.t === 'b') return `<label class="chk${span}">${inputFor(k, f, v, { ro })} ${esc(o.labels?.[f.n] || f.l)}</label>`;
  return `<div class="form-group${span}"><label>${esc(o.labels?.[f.n] || f.l)}${req ? ' <span class="req">*</span>' : ''}${calc ? ' <span class="hint">auto</span>' : ''}</label>${inputFor(k, f, v, { ro: ro || calc, calc, rows: o.rows })}</div>`;
}
/* sections: [{title, fields:[...]}] */
function renderSections(k, rec, sections, o = {}) {
  return `<div class="grid${o.cols ? ' ' + o.cols : ''}">` + sections.map(s => (s.title ? `<div class="sec-title">${esc(s.title)}</div>` : '') + s.fields.map(n => fieldBox(k, n, rec, o)).join('')).join('') + '</div>';
}
function readForm(root, k, base = {}) {
  const out = { ...base };
  $$('[data-f]', root).forEach(el => {
    const f = fieldOf(k, el.dataset.f); let v;
    if (el.type === 'checkbox') v = el.checked;
    else if (el.type === 'number') v = el.value === '' ? null : num(el.value);
    else if (el.type === 'datetime-local') v = el.value ? new Date(el.value).toISOString() : null;
    else v = el.value.trim() === '' ? null : el.value.trim();
    if (f.n === 'Pword') { if (!v) return; }
    out[f.n] = v;
  });
  return out;
}
const setVal = (root, n, v) => { const el = $(`[data-f="${n}"]`, root); if (!el) return; if (el.type === 'checkbox') el.checked = !!v; else if (el.type === 'number' && v !== null && v !== '' && v !== undefined) el.value = Number.isInteger(v) ? v : r2(v); else el.value = v ?? ''; };
const getVal = (root, n) => { const el = $(`[data-f="${n}"]`, root); if (!el) return null; if (el.type === 'checkbox') return el.checked; if (el.type === 'number') return el.value === '' ? null : num(el.value); return el.value; };
function validate(root, req) { let ok = true; for (const n of req) { const el = $(`[data-f="${n}"]`, root); if (!el) continue; const bad = el.type === 'checkbox' ? false : !String(el.value).trim(); el.classList.toggle('invalid', bad); if (bad) ok = false; } if (!ok) toast('Please fill the required fields (marked *).', 'warn'); return ok; }

/* =====================================================================
   DATA GRID — search, sort, paging, totals, Excel export
   ===================================================================== */
function fmtCell(c, r) {
  const v = typeof c.get === 'function' ? c.get(r) : r[c.f];
  if (c.html) return c.html(r, v);
  const t = c.t || (c.k ? fieldOf(c.k, c.f).t : 't');
  if (v === null || v === undefined || v === '') return '';
  if (t === 'c') return `<span class="${num(v) < 0 ? 'neg' : ''}">${money(v)}</span>`;
  if (t === 'n') return intf(v);
  if (t === 'p') return num(v) + '%';
  if (t === 'd') return fdate(v);
  if (t === 'dt') return fdt(v);
  if (t === 'b') return v ? '✔' : '';
  return esc(v);
}
function makeGrid(host, o) {
  /* o: {cols:[{f,l,t,get,html,w,sum}], rows, onRow, pageSize, sortBy, desc, selectable, rowClass, empty} */
  const st = { sort: o.sortBy || null, desc: o.desc ?? true, page: 0, sel: new Set() };
  const ps = o.pageSize || 150;
  function sortedRows() {
    const rows = o.rows.slice(); if (!st.sort) return rows;
    const c = o.cols.find(x => (x.f || x.l) === st.sort) || {}; const g = r => typeof c.get === 'function' ? c.get(r) : r[c.f];
    const numeric = ['c', 'n', 'p'].includes(c.t);
    rows.sort((a, b) => { const x = g(a), y = g(b); const r = numeric ? num(x) - num(y) : String(x ?? '').localeCompare(String(y ?? ''), undefined, { numeric: true }); return st.desc ? -r : r; });
    return rows;
  }
  function draw() {
    const rows = sortedRows(), pages = Math.max(1, Math.ceil(rows.length / ps)); if (st.page >= pages) st.page = pages - 1;
    const view = rows.slice(st.page * ps, st.page * ps + ps);
    const hasSum = o.cols.some(c => c.sum);
    host.innerHTML = `<div class="tbl-wrap" style="${o.maxH ? 'max-height:' + o.maxH : ''}"><table class="dt"><thead><tr>${o.selectable ? '<th style="width:28px"><input type="checkbox" data-all></th>' : ''}${o.cols.map(c => `<th data-s="${esc(c.f || c.l)}" class="${['c', 'n', 'p'].includes(c.t) ? 'num ' : ''}${st.sort === (c.f || c.l) ? (st.desc ? 'sorted-desc' : 'sorted-asc') : ''}" style="${c.w ? 'width:' + c.w : ''}">${esc(c.l)}</th>`).join('')}</tr></thead>
      <tbody>${view.length ? view.map((r, i) => `<tr data-i="${st.page * ps + i}" class="${o.rowClass ? o.rowClass(r) : ''}${st.sel.has(r) ? ' sel' : ''}">${o.selectable ? `<td><input type="checkbox" data-sel ${st.sel.has(r) ? 'checked' : ''}></td>` : ''}${o.cols.map(c => `<td class="${['c', 'n', 'p'].includes(c.t) ? 'num' : ''}">${fmtCell(c, r)}</td>`).join('')}</tr>`).join('') : `<tr><td colspan="${o.cols.length + (o.selectable ? 1 : 0)}" class="empty">${o.empty || 'No records found.'}</td></tr>`}</tbody>
      ${hasSum && rows.length ? `<tfoot><tr>${o.selectable ? '<td></td>' : ''}${o.cols.map((c, i) => `<td class="${c.sum ? 'num' : ''}">${c.sum ? (c.t === 'c' ? money(sum(rows, r => typeof c.get === 'function' ? c.get(r) : r[c.f])) : intf(sum(rows, r => typeof c.get === 'function' ? c.get(r) : r[c.f]))) : i === 0 ? rows.length + ' rec.' : ''}</td>`).join('')}</tr></tfoot>` : ''}</table></div>
      ${pages > 1 ? `<div class="pager">Page ${st.page + 1} of ${pages} · ${rows.length} records <button class="ghost small" data-p="-1" ${st.page ? '' : 'disabled'}>‹ Prev</button><button class="ghost small" data-p="1" ${st.page < pages - 1 ? '' : 'disabled'}>Next ›</button></div>` : ''}`;
    $$('th[data-s]', host).forEach(th => th.onclick = () => { const s = th.dataset.s; if (st.sort === s) st.desc = !st.desc; else { st.sort = s; st.desc = false; } draw(); });
    $$('[data-p]', host).forEach(b => b.onclick = () => { st.page += +b.dataset.p; draw(); });
    $$('tbody tr[data-i]', host).forEach(tr => tr.onclick = e => {
      const r = rows[+tr.dataset.i];
      if (o.selectable && (e.target.matches('[data-sel]') || !o.onRow)) { if (st.sel.has(r)) st.sel.delete(r); else st.sel.add(r); tr.classList.toggle('sel'); const cb = $('[data-sel]', tr); if (cb) cb.checked = st.sel.has(r); o.onSelect && o.onSelect([...st.sel]); return; }
      o.onRow && o.onRow(r, e);
    });
    const all = $('[data-all]', host); if (all) all.onclick = e => { e.stopPropagation(); if (all.checked) rows.forEach(r => st.sel.add(r)); else st.sel.clear(); draw(); o.onSelect && o.onSelect([...st.sel]); };
  }
  draw();
  return { redraw(rows) { if (rows) o.rows = rows; draw(); }, selected: () => [...st.sel], clearSel() { st.sel.clear(); draw(); }, rows: () => sortedRows() };
}
function exportXlsx(name, rows, cols) {
  if (!window.XLSX) { toast('Excel library not loaded (offline?)', 'error'); return; }
  const data = rows.map(r => Object.fromEntries(cols.map(c => { let v = typeof c.get === 'function' ? c.get(r) : r[c.f]; if (c.t === 'd' && v) v = String(v).slice(0, 10); return [c.l, v ?? '']; })));
  const ws = XLSX.utils.json_to_sheet(data), wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 30).replace(/[\\/?*[\]:]/g, '-'));
  XLSX.writeFile(wb, `${name.replace(/[^\w-]+/g, '_')}_${today()}.xlsx`);
}
function textFilter(rows, q, fields) { q = String(q || '').trim().toLowerCase(); if (!q) return rows; const terms = q.split(/\s+/); return rows.filter(r => { const hay = fields.map(f => typeof f === 'function' ? f(r) : r[f]).join(' ').toLowerCase(); return terms.every(t => hay.includes(t)); }); }

/* =====================================================================
   GENERIC LIST PAGE (search + grid + editor) — used by simple forms
   ===================================================================== */
function listPage(view, o) {
  /* o: {k, cols, sections, search, filters:[{f,l,opts}], defaults(), title, req, ro, calc, before(rec), after(rec), canDelete, extra(rec) } */
  const k = o.k, L = LISTS[k];
  view.innerHTML = `<div class="card"><div class="toolbar"><input type="search" placeholder="Search…" data-q>${(o.filters || []).map((f, i) => `<select data-flt="${i}"><option value="">${esc(f.l)}: all</option>${f.opts.map(x => `<option>${esc(x)}</option>`).join('')}</select>`).join('')}<span class="spacer" style="flex:1"></span>
    <span class="muted" data-count></span><button class="ghost small" data-imp>⬆ Import</button><button class="ghost small" data-exp>⬇ Excel</button><button class="small" data-new>＋ New ${esc(o.noun || 'record')}</button></div><div data-grid></div></div>`;
  let grid;
  const rowsNow = () => { let rows = DB.peek(k); (o.filters || []).forEach((f, i) => { const v = $(`[data-flt="${i}"]`, view).value; if (v) rows = rows.filter(r => (f.get ? f.get(r) : r[f.f]) === v); }); if (o.where) rows = rows.filter(o.where); return textFilter(rows, $('[data-q]', view).value, o.search || L.fields.filter(f => f.t === 't').map(f => f.n)); };
  const refresh = () => { const rows = rowsNow(); $('[data-count]', view).textContent = rows.length + ' records'; if (!grid) grid = makeGrid($('[data-grid]', view), { cols: o.cols, rows, onRow: r => edit(r), sortBy: o.sortBy || L.pk, desc: o.desc ?? true }); else grid.redraw(rows); };
  function edit(rec) {
    const isNew = !rec, r = rec ? { ...rec } : { ...(o.defaults ? o.defaults() : {}) };
    const fo = { ro: [L.pk, ...(o.ro || [])].filter(Boolean), calc: o.calc, req: o.req, wide: o.wide };
    const inner = o.tabs ? `<div class="tabs">${o.tabs.map((t, i) => `<button class="tab ${i ? '' : 'active'}" data-mt="${i}">${esc(t.title)}</button>`).join('')}</div>${o.tabs.map((t, i) => `<div class="tab-pane ${i ? '' : 'active'}" data-mp="${i}">${renderSections(k, r, t.sections, fo)}</div>`).join('')}` : renderSections(k, r, o.sections, fo);
    const body = Modal.open((isNew ? 'New ' : 'Edit ') + (o.noun || 'record') + (isNew ? '' : ' — ' + L.title(r)), inner + (o.extra ? `<div data-extra style="margin-top:14px"></div>` : ''), [
      ...(!isNew && o.canDelete !== false ? [{ label: '🗑 Delete', cls: 'danger', onClick: async () => { if (!(await confirmBox('Delete this record permanently?', 'Delete', true))) return edit(rec); try { if (o.onDelete) await o.onDelete(rec); await DB.remove(k, rec); toast('Deleted', 'success'); refresh(); } catch (e) { toast(e.message, 'error'); } } }] : []),
      ...(!isNew ? [{ label: '⧉ Duplicate', cls: 'ghost', onClick: () => { const c = { ...rec }; delete c._id; if (L.pk) delete c[L.pk]; Modal.close(); editNew(c); } }] : []),
      ...((o.buttons ? o.buttons(rec, () => save(), body => body) : []) || []),
      { label: 'Cancel', cls: 'ghost' },
      { label: '💾 Save', cls: 'success', onClick: save }], 'wide');
    $$('[data-mt]', body).forEach(t => t.onclick = () => { $$('[data-mt]', body).forEach(x => x.classList.toggle('active', x === t)); $$('[data-mp]', body).forEach(p => p.classList.toggle('active', p.dataset.mp === t.dataset.mt)); });
    if (o.extra) o.extra($('[data-extra]', body), r);
    if (o.wire) o.wire(body, r);
    async function save() {
      if (o.req && !validate(body, o.req)) return;
      let out = readForm(body, k, r);
      try { if (o.before) { const res = await o.before(out, isNew); if (res === false) return; } const saved = await DB.save(k, out); if (o.after) await o.after(saved, isNew); toast('Saved ✔', 'success'); Modal.close(); refresh(); } catch (e) { toast('Save failed: ' + e.message, 'error', 6000); }
    }
  }
  function editNew(pre) { edit(null); const body = $('#modalBody'); for (const [n, v] of Object.entries(pre)) setVal(body, n, v); }
  $('[data-q]', view).oninput = debounce(refresh, 200);
  $$('[data-flt]', view).forEach(s => s.onchange = refresh);
  $('[data-new]', view).onclick = () => edit(null);
  $('[data-exp]', view).onclick = () => exportXlsx(L.name, rowsNow(), L.fields.map(f => ({ f: f.n, l: f.n, t: f.t })));
  $('[data-imp]', view).onclick = () => importDialog(k, refresh);
  return { refresh, edit, editNew, rows: rowsNow };
}

/* =====================================================================
   IMPORT (Excel / CSV exported from Access) — columns matched by name
   ===================================================================== */
function matchField(k, header) { const h = norm(header); return LISTS[k].fields.find(f => norm(f.n) === h || norm(f.sp) === h || (f.alias && norm(f.alias) === h) || norm(f.l) === h) || (h === 'id' ? LISTS[k].byName.AccessID || (LISTS[k].pk && LISTS[k].byName[LISTS[k].pk]) : null); }
function coerce(f, v) {
  if (v === null || v === undefined || v === '') return f.t === 'b' ? false : null;
  if (f.t === 'b') return v === true || /^(true|yes|-1|1|y|✔)$/i.test(String(v).trim());
  if (['n', 'c', 'p'].includes(f.t)) return num(v);
  if (f.t === 'd' || f.t === 'dt') {
    let d = v instanceof Date ? v : null;
    if (!d && typeof v === 'number') d = new Date(Math.round((v - 25569) * 86400000));
    if (!d) { const s = String(v).trim(); const m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/); d = m ? new Date(+(m[3].length === 2 ? '20' + m[3] : m[3]), +m[2] - 1, +m[1]) : new Date(s); }
    if (isNaN(d)) return null;
    return f.t === 'd' ? `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` : d.toISOString();
  }
  return String(v);
}
async function readWorkbookRows(file) { const buf = await file.arrayBuffer(); const wb = XLSX.read(buf, { type: 'array', cellDates: true }); return wb.SheetNames.map(n => ({ sheet: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: null, raw: true }) })); }
function guessList(name) { const n = norm(name); return Object.keys(LISTS).find(k => norm(LISTS[k].name) === n || norm(listName(k)) === n) || Object.keys(LISTS).find(k => n.includes(norm(LISTS[k].name))); }
function importDialog(fixedK, done) {
  const body = Modal.open('Import data from Excel / CSV', `<div class="card-desc">Export your Access tables to Excel (External Data → Excel), then pick the file(s) here. Each sheet or file is matched to a list by its name; columns are matched by field name. ${fixedK ? `Target list: <b>${esc(listName(fixedK))}</b>.` : ''}</div>
    <input type="file" accept=".xlsx,.xls,.csv" multiple data-file><div data-plan style="margin-top:14px"></div><div class="progress" style="display:none"><div></div></div><div data-st></div>`, [{ label: 'Close', cls: 'ghost' }, { label: '⬆ Import', cls: 'success', onClick: () => go() }], 'wide');
  let plan = [];
  $('[data-file]', body).onchange = async e => {
    plan = [];
    for (const file of e.target.files) for (const s of await readWorkbookRows(file)) {
      const k = fixedK || guessList(s.sheet) || guessList(file.name.replace(/\.\w+$/, ''));
      plan.push({ label: file.name + (s.sheet && !/^sheet\d*$/i.test(s.sheet) ? ' › ' + s.sheet : ''), rows: s.rows, k });
    }
    $('[data-plan]', body).innerHTML = `<table class="dt"><thead><tr><th>Source</th><th>Rows</th><th>Target list</th><th>Matched columns</th></tr></thead><tbody>${plan.map((p, i) => { const hs = Object.keys(p.rows[0] || {}); const m = p.k ? hs.filter(h => matchField(p.k, h)) : []; return `<tr><td>${esc(p.label)}</td><td>${p.rows.length}</td><td><select data-pk="${i}"><option value="">— skip —</option>${Object.keys(LISTS).map(k => `<option value="${k}" ${k === p.k ? 'selected' : ''}>${esc(listName(k))}</option>`).join('')}</select></td><td class="muted" style="white-space:normal">${p.k ? `${m.length}/${hs.length}: ${esc(m.slice(0, 12).join(', '))}${m.length > 12 ? '…' : ''}` : ''}</td></tr>`; }).join('')}</tbody></table>
      <label class="chk"><input type="checkbox" data-replace> Replace: delete existing records in the target list first</label>`;
    $$('[data-pk]', body).forEach(s => s.onchange = () => { plan[+s.dataset.pk].k = s.value || null; });
  };
  async function go() {
    const todo = plan.filter(p => p.k && p.rows.length); if (!todo.length) { toast('Nothing to import.', 'warn'); return; }
    const prog = $('.progress', body), bar = $('.progress div', body), st = $('[data-st]', body); prog.style.display = '';
    const replace = $('[data-replace]', body)?.checked; let total = 0, errs = [];
    for (const p of todo) {
      const L = LISTS[p.k]; await DB.all(p.k);
      if (replace) { st.innerHTML = `<div class="status-msg info">Deleting old records in ${esc(listName(p.k))}…</div>`; errs.push(...await DB.bulk(DB.peek(p.k).slice().map(r => () => DB.remove(p.k, r)), (d, n) => bar.style.width = (d / n * 100) + '%')); }
      const recs = p.rows.map(row => { const rec = {}; for (const [h, v] of Object.entries(row)) { const f = matchField(p.k, h); if (f) rec[f.n] = coerce(f, v); } if (p.k === 'emp' && rec.Pword && !String(rec.Pword).startsWith('sha256:')) { /* keeps Access password; upgraded to hash at first login */ } return rec; });
      st.innerHTML = `<div class="status-msg info">Importing ${recs.length} rows into ${esc(listName(p.k))}…</div>`;
      errs.push(...await DB.bulk(recs.map(r => () => DB.save(p.k, r)), (d, n) => { bar.style.width = (d / n * 100) + '%'; st.innerHTML = `<div class="status-msg info">${esc(listName(p.k))}: ${d} / ${n}</div>`; }));
      total += recs.length;
    }
    st.innerHTML = `<div class="status-msg ${errs.length ? 'warn' : 'success'}">Imported ${total - errs.length} of ${total} rows.${errs.length ? ' Errors: ' + esc(uniq(errs).slice(0, 5).join(' | ')) : ''}</div>`;
    done && done();
  }
}

/* =====================================================================
   SVG CHARTS (no library)
   ===================================================================== */
function barChart(labels, series, o = {}) {
  /* series: [{name, values, color, line}] */
  const W = 560, H = o.h || 230, P = { l: 52, r: 10, t: 12, b: 34 }, iw = W - P.l - P.r, ih = H - P.t - P.b;
  const bars = series.filter(s => !s.line), lines = series.filter(s => s.line);
  const all = series.flatMap(s => s.values.map(num)); let mx = Math.max(0, ...all), mn = Math.min(0, ...all); if (mx === mn) mx = mn + 1;
  const raw = (mx - mn) / 4, mag = 10 ** Math.floor(Math.log10(raw || 1)), step = [1, 2, 2.5, 5, 10].map(m => m * mag).find(s => s >= raw) || raw; mx = Math.ceil(mx / step) * step; mn = Math.floor(mn / step) * step;
  const y = v => P.t + ih - (num(v) - mn) / (mx - mn) * ih, gw = iw / Math.max(1, labels.length), bw = Math.max(3, Math.min(26, (gw * 0.75) / Math.max(1, bars.length)));
  const fmtA = v => Math.abs(v) >= 1e6 ? (v / 1e6).toFixed(1) + 'M' : Math.abs(v) >= 1e3 ? (v / 1e3).toFixed(0) + 'K' : String(Math.round(v));
  let s = `<svg viewBox="0 0 ${W} ${H}" role="img">`;
  for (let v = mn, i = 0; v <= mx + step / 2 && i < 12; v += step, i++) { const yy = y(v); s += `<line class="ax" x1="${P.l}" x2="${W - P.r}" y1="${yy}" y2="${yy}" stroke-dasharray="${Math.abs(v) > step / 2 ? '3 3' : ''}"/><text x="${P.l - 6}" y="${yy + 3}" text-anchor="end">${fmtA(v)}</text>`; }
  labels.forEach((lb, i) => {
    const x0 = P.l + gw * i + (gw - bw * bars.length) / 2;
    bars.forEach((b, j) => { const v = num(b.values[i]); const yy = y(Math.max(0, v)), h = Math.abs(y(v) - y(0)); s += `<rect x="${x0 + j * bw}" y="${yy}" width="${bw - 1}" height="${Math.max(0.5, h)}" fill="${b.color}" rx="2"><title>${esc(b.name)} · ${esc(lb)}: ${money(v)}</title></rect>`; });
    if (labels.length <= 14 || i % Math.ceil(labels.length / 12) === 0) s += `<text x="${P.l + gw * i + gw / 2}" y="${H - P.b + 14}" text-anchor="middle">${esc(lb)}</text>`;
  });
  lines.forEach(l => { const pts = l.values.map((v, i) => `${P.l + gw * i + gw / 2},${y(v)}`).join(' '); s += `<polyline points="${pts}" fill="none" stroke="${l.color}" stroke-width="2.2"/>` + l.values.map((v, i) => `<circle cx="${P.l + gw * i + gw / 2}" cy="${y(v)}" r="3" fill="${l.color}"><title>${esc(l.name)} · ${esc(labels[i])}: ${money(v)}</title></circle>`).join(''); });
  s += '</svg>';
  return `<div class="chart">${s}<div class="legend">${series.map(x => `<span><i style="background:${x.color}"></i>${esc(x.name)}</span>`).join('')}</div></div>`;
}
const CSSV = n => getComputedStyle(document.body).getPropertyValue(n).trim();
