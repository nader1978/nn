/* =====================================================================
   PRINT ENGINE — branded A4 documents (logo, company header, watermark, footer)
   ===================================================================== */
function amountInWords(n) {
  const a = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'], b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  const w = x => x < 20 ? a[x] : x < 100 ? b[Math.floor(x / 10)] + (x % 10 ? '-' + a[x % 10] : '') : x < 1000 ? a[Math.floor(x / 100)] + ' Hundred' + (x % 100 ? ' and ' + w(x % 100) : '') : x < 1e6 ? w(Math.floor(x / 1000)) + ' Thousand' + (x % 1000 ? ' ' + w(x % 1000) : '') : w(Math.floor(x / 1e6)) + ' Million' + (x % 1e6 ? ' ' + w(x % 1e6) : '');
  const v = Math.abs(r2(n)), i = Math.floor(v), f = Math.round((v - i) * 100);
  return (i ? w(i) : 'Zero') + ' Dirhams' + (f ? ' and ' + w(f) + ' Fils' : '') + ' Only';
}
/* RAB document guidelines — every printed document uses the logo and the logo colours:
   charcoal #212123 (text, footer) · RAB green #0e7a3a (headings, table headers) · gold #c9a227 (rules, highlights) */
/* document colours follow the selected colour theme (Settings → Colour theme, shared with SC Mngt) */
const docColors = () => { const g = n => CSSV(n) || null; return { green: g('--acc-2') || '#1d4ed8', greenD: g('--acc-1') || '#12306b', soft: g('--acc-soft') || '#eaf0fb', line: g('--acc-line') || '#c9d7f0', gold: g('--acc-3') || '#3b82f6', goldL: g('--acc-soft') || '#eaf0fb', char: '#1f2937' }; };
const absUrl = u => !u ? '' : /^(data:|https?:)/.test(u) ? u : new URL(u, location.href).href;
/* logo pictures are embedded into each document (print windows cannot load picture files reliably) */
const DOC_IMG = {};
async function docImg(u) { if (!u || u.startsWith('data:')) return u || ''; if (DOC_IMG[u]) return DOC_IMG[u]; try { const r = await fetch(absUrl(u)); const bl = await r.blob(); DOC_IMG[u] = await new Promise(res => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.readAsDataURL(bl); }); } catch (e) { DOC_IMG[u] = absUrl(u); } return DOC_IMG[u]; }
function preloadDocImages() { [brandGet().img].forEach(u => u && docImg(u)); }
const docSrc = u => !u ? '' : u.startsWith('data:') ? u : (DOC_IMG[u] || absUrl(u));
function printDoc(title, inner, o = {}) {
  const c = docCompany(), b = brandGet(), w = window.open('', '_blank'); if (!w) { toast('Allow pop-ups to print.', 'warn'); return; }
  const logo = docSrc(b.img || ''), mark = logo, wmark = logo;
  const meta = (o.meta || []).filter(m => m[1] !== undefined && m[1] !== null && m[1] !== '');
  const contact = [c.TelephoneNumber && '☎ ' + c.TelephoneNumber, c.Email && '✉ ' + c.Email, c.Website && '🌐 ' + c.Website].filter(Boolean).map(esc).join('&nbsp;&nbsp;·&nbsp;&nbsp;');
  const D = docColors();
  w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(title)}${o.no ? ' ' + esc(o.no) : ''}</title><style>
    @page{size:A4 ${o.landscape ? 'landscape' : 'portrait'};margin:12mm 12mm 22mm}
    *{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}
    body{font-family:'Segoe UI',Arial,sans-serif;color:${D.char};margin:0;padding:26px 32px 90px;font-size:12.5px;background:#fff}
    .wm{position:fixed;left:50%;top:54%;transform:translate(-50%,-50%) rotate(-8deg);height:${o.landscape ? '62vh' : '52vh'};max-height:560px;width:auto;opacity:.07;pointer-events:none;z-index:0}
    .doc{position:relative;z-index:1}
    .wmt{height:auto;width:auto;font-size:190px;font-weight:900;letter-spacing:14px;color:${D.greenD};opacity:.05}
    .hd{display:flex;justify-content:space-between;align-items:center;gap:20px;padding-bottom:14px}
    .brand{display:flex;align-items:center;gap:16px}.brand img{height:82px;width:auto}
    .co{font-size:19px;font-weight:800;color:${D.char};letter-spacing:.2px}.co2{font-size:15px;color:${D.greenD};direction:rtl;font-weight:600;margin-top:1px}
    .meta{color:#555;font-size:11px;line-height:1.55;margin-top:3px}
    .ttl{text-align:right}.ttl h1{font-size:22px;margin:0;color:${D.green};text-transform:uppercase;letter-spacing:2px;font-weight:800}
    .ttl .sub{color:#666;margin-top:3px}
    .mbox{display:inline-grid;grid-template-columns:auto auto;gap:3px 14px;margin-top:8px;padding:8px 12px;border:1px solid ${D.line};border-top:3px solid ${D.gold};border-radius:6px;background:${D.soft};text-align:left;font-size:11.5px}
    .mbox span{color:#666}.mbox b{color:${D.char}}
    .rule{height:5px;background:linear-gradient(90deg,${D.char} 0 22%,${D.greenD} 22% 60%,${D.green} 60% 88%,${D.gold} 88%);border-radius:3px;margin-bottom:18px}
    table{width:100%;border-collapse:collapse;margin-top:10px}
    thead{display:table-header-group} tfoot{display:table-row-group} tr{page-break-inside:avoid}
    table.page{margin:0}table.page>thead>tr>td,table.page>tbody>tr>td,table.page>tfoot>tr>td{padding:0;border:0;background:none!important}table.page>tfoot{display:table-footer-group}.fspace{height:0}
    th{background:${D.green};color:#fff;padding:8px 9px;text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid ${D.gold}}
    td{padding:7px 9px;border-bottom:1px solid ${D.line};vertical-align:top} tbody tr:nth-child(even) td{background:#f7fbf8}
    td.n,th.n{text-align:right;font-variant-numeric:tabular-nums}
    tfoot td{font-weight:700;background:${D.soft};border-top:2px solid ${D.green};color:${D.char}}
    .box{display:grid;grid-template-columns:1fr 1fr;gap:6px 30px;margin:4px 0 12px;padding:10px 14px;border:1px solid ${D.line};border-left:4px solid ${D.green};border-radius:6px}
    .box div{border-bottom:1px dotted #c9d6cd;padding:3px 0}.box b{display:inline-block;min-width:130px;color:${D.greenD};font-weight:600}
    .tot{margin-left:auto;width:330px;margin-top:14px;border:1px solid ${D.line};border-radius:6px;overflow:hidden}
    .tot div{display:flex;justify-content:space-between;padding:6px 10px;border-bottom:1px solid ${D.line}}
    .tot .g{background:${D.char};color:#fff;font-weight:700;font-size:14.5px;border-bottom:0;border-top:3px solid ${D.gold}}.tot .g span:last-child{color:${D.gold}}
    .words{margin-top:12px;padding:9px 12px;background:${D.goldL};border-left:4px solid ${D.gold};font-style:italic;border-radius:0 6px 6px 0}
    .sig{display:flex;justify-content:space-between;margin-top:64px;gap:30px}.sig div{flex:1;border-top:1.5px solid ${D.char};padding-top:6px;text-align:center;color:#555;font-size:11.5px}
    .neg{color:#a80000}.grp td{background:${D.soft}!important;font-weight:700;color:${D.greenD};border-top:1px solid ${D.green}}
    .stamp{position:absolute;right:30px;top:150px;transform:rotate(-12deg);border:3px solid currentColor;border-radius:8px;padding:4px 16px;font-size:22px;font-weight:900;letter-spacing:3px;opacity:.55}
    .stamp.paid{color:${D.green}}.stamp.unpaid{color:#b42318}.stamp.due{color:#b7791f}.stamp.quote{color:${D.gold}}
    .ft{position:fixed;left:0;right:0;bottom:0;background:${D.char};color:#e9e9e9;border-top:3px solid ${D.gold};padding:8px 32px;font-size:10.5px;display:flex;justify-content:space-between;gap:16px;align-items:center}
    .ft img{height:24px;background:#fff;border-radius:4px;padding:2px}.ft .l{display:flex;align-items:center;gap:10px}.ft b{color:${D.gold}}
    .bar{position:fixed;bottom:56px;right:14px;z-index:5}.bar button{padding:9px 16px;background:${D.green};color:#fff;border:none;border-bottom:3px solid ${D.gold};border-radius:6px;cursor:pointer;font-weight:700}
    p{margin:8px 0}
    @media print{.bar{display:none}body{padding:0}.ft{padding:5px 12mm;font-size:9.5px}.fspace{height:16mm}}</style></head><body>
    ${wmark ? `<img class="wm" src="${wmark}" alt="">` : `<div class="wm wmt">${esc((brandGet().name || BRAND_ORIG.name).split(' ')[0])}</div>`}
    <div class="bar"><button onclick="try{opener.NFPCmailDoc(document.title,'<style>'+document.querySelector('style').textContent.split('position:fixed').join('position:static')+'</style>'+document.querySelector('.doc').outerHTML)}catch(e){alert('Open the e-mail window from the system page.')}" style="margin-right:6px">✉ Email</button><button onclick="print()">🖨 Print / Save PDF</button></div>
    <table class="page"><tbody><tr><td><div class="doc">
    ${o.stamp ? `<div class="stamp ${o.stamp.cls}">${esc(o.stamp.t)}</div>` : ''}
    <div class="hd"><div class="brand">${logo ? `<img src="${logo}" alt="logo">` : ''}<div><div class="co">${esc(c.CompanyFullName || 'Company name')}</div>${c.ArabicName ? `<div class="co2">${esc(c.ArabicName)}</div>` : ''}<div class="meta">${esc([c.Address, c.City, c.Country].filter(Boolean).join(', '))}${c.POBox ? ' · P.O. Box ' + esc(c.POBox) : ''}${c.BusinessType ? '<br>' + esc(c.BusinessType) : ''}</div></div></div>
    <div class="ttl"><h1>${esc(title)}</h1>${o.sub ? `<div class="sub">${o.sub}</div>` : ''}${meta.length ? `<div class="mbox">${meta.map(m => `<span>${esc(m[0])}</span><b>${esc(m[1])}</b>`).join('')}</div>` : ''}</div></div>
    <div class="rule"></div>${inner}</div></td></tr></tbody><tfoot><tr><td><div class="fspace"></div></td></tr></tfoot></table>
    <div class="ft"><div class="l">${mark ? `<img src="${mark}" alt="">` : ''}<span><b>${esc(c.CompanyFullName || '')}</b>${contact ? '&nbsp;&nbsp;·&nbsp;&nbsp;' + contact : ''}</span></div><span>${o.foot || 'Computer generated document · E.O.E.'} · Printed ${fdt(nowIso())} by ${esc(me())}</span></div></body></html>`);
  w.document.close();
  if (o.auto !== false) { const t0 = Date.now(); const go = () => { const imgs = [...w.document.images]; if (imgs.every(i => i.complete) || Date.now() - t0 > 3000) { try { w.focus(); w.print(); } catch (e) {} } else setTimeout(go, 150); }; setTimeout(go, 300); }
}
function printTable(title, rows, cols) {
  printDoc(title, `<table><thead><tr>${cols.map(c => `<th class="${['c', 'n'].includes(c.t) ? 'n' : ''}">${esc(c.l)}</th>`).join('')}</tr></thead><tbody>${rows.map(r => `<tr>${cols.map(c => `<td class="${['c', 'n'].includes(c.t) ? 'n' : ''}">${fmtCell(c, r)}</td>`).join('')}</tr>`).join('')}</tbody><tfoot><tr>${cols.map((c, i) => `<td class="n">${c.sum ? (c.t === 'c' ? money(sum(rows, r => typeof c.get === 'function' ? c.get(r) : r[c.f])) : intf(sum(rows, c.f))) : i === 0 ? rows.length + ' rows' : ''}</td>`).join('')}</tr></tfoot></table>`, { landscape: cols.length > 6, meta: [['Date', fdate(today())], ['Rows', rows.length]] });
}
