'use strict';
/* =====================================================================
   NFPC Supply Chain System — HTML/SharePoint edition of
   Sales_Expenses_App.accdb. Each Access table = one SharePoint list,
   each Access form = one page in this file.
   ===================================================================== */
const APP = { name: 'NFPC Supply Chain System', ver: '1.0', ls: 'nfpcsc' };

/* ---------------- small helpers ---------------- */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const num = v => { const n = typeof v === 'number' ? v : parseFloat(String(v ?? '').replace(/,/g, '')); return isFinite(n) ? n : 0; };
const r2 = v => Math.round(num(v) * 100) / 100;
const money = v => num(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const intf = v => num(v).toLocaleString('en-US', { maximumFractionDigits: 0 });
const pad = n => String(n).padStart(2, '0');
const today = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const nowIso = () => new Date().toISOString();
const addDays = (iso, n) => { const d = new Date(iso + 'T00:00:00'); d.setDate(d.getDate() + n); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const daysBetween = (a, b) => Math.round((new Date(b + 'T00:00:00') - new Date(a + 'T00:00:00')) / 86400000);
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const fdate = iso => { if (!iso) return ''; const s = String(iso).slice(0, 10); const [y, m, d] = s.split('-'); return y && m ? `${d}-${MON[+m - 1]}-${y}` : s; };
const fdt = iso => { if (!iso) return ''; const d = new Date(iso); return isNaN(d) ? String(iso) : `${fdate(d.toISOString().slice(0, 10))} ${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const ym = iso => String(iso || '').slice(0, 7);
const norm = s => String(s ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
const debounce = (fn, ms = 250) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
const uniq = a => [...new Set(a.filter(v => v !== null && v !== undefined && String(v).trim() !== ''))].sort((x, y) => String(x).localeCompare(String(y)));
const sum = (a, f) => a.reduce((s, x) => s + num(typeof f === 'function' ? f(x) : x[f]), 0);
const groupBy = (a, f) => { const m = new Map(); for (const x of a) { const k = typeof f === 'function' ? f(x) : x[f]; if (!m.has(k)) m.set(k, []); m.get(k).push(x); } return m; };
const lsGet = (k, d) => { try { const v = localStorage.getItem(APP.ls + '_' + k); return v === null ? d : JSON.parse(v); } catch (e) { return d; } };
const lsSet = (k, v) => { try { localStorage.setItem(APP.ls + '_' + k, JSON.stringify(v)); } catch (e) {} };
async function sha256(t) { const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(String(t))); return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, '0')).join(''); }

function toast(msg, type = '', ms = 3200) { const t = document.createElement('div'); t.className = 'toast ' + type; t.textContent = msg; $('#toasts').appendChild(t); setTimeout(() => t.remove(), ms); }
const Modal = {
  open(title, bodyHtml, buttons = [], cls = '') {
    $('#modal').className = 'modal ' + cls; $('#modalTitle').textContent = title; $('#modalBody').innerHTML = bodyHtml;
    const f = $('#modalFoot'); f.innerHTML = '';
    for (const b of buttons) { const el = document.createElement('button'); el.textContent = b.label; el.className = b.cls || ''; el.onclick = () => b.onClick ? b.onClick() : Modal.close(); f.appendChild(el); }
    f.style.display = buttons.length ? '' : 'none'; $('#modalBg').classList.add('open'); return $('#modalBody');
  },
  close() { $('#modalBg').classList.remove('open'); $('#modalBody').innerHTML = ''; }
};
function confirmBox(msg, okLabel = 'Yes', danger = false) {
  return new Promise(res => Modal.open('Please confirm', `<div style="font-size:14px;line-height:1.6">${msg}</div>`, [
    { label: 'Cancel', cls: 'ghost', onClick: () => { Modal.close(); res(false); } },
    { label: okLabel, cls: danger ? 'danger' : '', onClick: () => { Modal.close(); res(true); } }], 'narrow'));
}

/* =====================================================================
   SCHEMA — recovered from the Access front-end (forms, queries, reports).
   Field types: t text · m multi-line · n number · c money · p percent
                d date · dt date+time · b yes/no · ch choice (free text allowed)
   ===================================================================== */
/* =====================================================================
   SCHEMA — generated from "NFPC Supply Chain Prog_2022.accdb":
   every table used by a form/report → one SharePoint list (existing lists are reused).
   Field types: t text · m multi-line · n number · c money · d date · dt date+time · b yes/no · ch choice
   ===================================================================== */
const F = (n, t = 't', l, x = {}) => ({ n, t, l: l || n.replace(/_/g, ' ').replace(/([a-z])([A-Z])/g, '$1 $2').replace(/\s+/g, ' ').trim(), ...x });
const LISTS = {};
const SP_RESERVED = new Set('title,id,created,modified,author,editor,attachments,contenttype,order,guid,fileref,fsobjtype,path,linktitle,edit,docicon,_uiversionstring'.split(','));
const spSafe = n => { let s = n.replace(/[^A-Za-z0-9_]/g, '_').replace(/_+/g, '_').replace(/^_+|_+$/g, ''); if (/^\d/.test(s)) s = 'F' + s; if (SP_RESERVED.has(s.toLowerCase())) s += '_f'; return s.slice(0, 32) || 'Field'; };
function buildLists() {
  for (const [k, T] of Object.entries(NFPC.tables)) {
    const fields = T.fields.map(f => F(f.n, f.t));
    const seen = new Set();
    for (const f of fields) { let s = spSafe(f.n); while (seen.has(s.toLowerCase())) s = s.slice(0, 29) + '_' + Math.floor(Math.random() * 90 + 10); seen.add(s.toLowerCase()); f.sp = s; }
    const tf = fields.find(f => f.t === 't' && !/id$|code$|link/i.test(f.n)) || fields.find(f => f.t === 't');
    LISTS[k] = { key: k, name: T.list, label: k.replace(/_/g, ' '), pk: T.pk, site: T.site, existing: T.existing, fields,
      title: r => String((T.pk && r[T.pk] ? r[T.pk] + ' · ' : '') + (tf ? r[tf.n] ?? '' : '')).trim() || ('Record ' + (r._id || '')) };
    LISTS[k].byName = Object.fromEntries(fields.map(f => [f.n, f]));
  }
}
const fieldOf = (k, n) => (LISTS[k] && LISTS[k].byName[n]) || F(n);
buildLists();
