// خادم Node لـ Railway — البيانات على قرص دائم (Volume) في DATA_DIR
import JSZipLib from 'jszip';
import { webcrypto } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import { fileURLToPath } from 'node:url';
if (!globalThis.crypto) globalThis.crypto = webcrypto;
globalThis.JSZip = JSZipLib;
const indexedDB = undefined, sessionStorage = undefined;
const DATA_DIR = process.env.DATA_DIR || process.env.RAILWAY_VOLUME_MOUNT_PATH || path.join(path.dirname(fileURLToPath(import.meta.url)), 'data');
fs.mkdirSync(DATA_DIR, { recursive: true });
function getStore() {   // نفس واجهة Netlify Blobs لكن على القرص
  const f = (k) => path.join(DATA_DIR, encodeURIComponent(k));
  const write = (k, buf) => { const tmp = f(k) + '.tmp' + process.pid; fs.writeFileSync(tmp, buf); fs.renameSync(tmp, f(k)); };
  return {
    async get(k, o = {}) {
      if (!fs.existsSync(f(k))) return null;
      const b = fs.readFileSync(f(k));
      if (o.type === 'json') return JSON.parse(b.toString());
      if (o.type === 'text') return b.toString();
      return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
    },
    async set(k, v) { write(k, typeof v === 'string' ? v : Buffer.from(v)); },
    async setJSON(k, v) { write(k, JSON.stringify(v)); },
    async delete(k) { if (fs.existsSync(f(k))) fs.unlinkSync(f(k)); },
  };
}

/* قراءة وكتابة ملفات Excel داخل المتصفح (منقول من excel_io.py) */
'use strict';
const XL = (() => {
  const AR = '٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', EN = '01234567890123456789';
  const normDigits = (s) => (s == null ? '' : String(s)).replace(/[٠-٩۰-۹]/g, (c) => EN[AR.indexOf(c)]);
  const normHeader = (s) => normDigits(s || '').trim().toLowerCase()
    .replace(/[ً-ْـ]/g, '').replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي')
    .replace(/[\s_\-:./()]+/g, ' ').trim();
  const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (m, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-f]+);/gi, (m, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, '&');
  const xesc = (s) => String(s).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const attr = (tag, name) => { const m = tag.match(new RegExp('(?:^|\\s)' + name.replace(':', '\\:') + '="([^"]*)"')); return m ? m[1] : null; };

  const FIELD_LABELS = { mrn: 'رقم الملف الطبي', name: 'الاسم', phone: 'رقم الجوال', phone2: 'رقم جوال إضافي', clinic: 'العيادة', city: 'المدينة',
    address: 'العنوان الوطني المختصر', district: 'الحي', pharmacy: 'الصيدلية', periodic: 'يرغب بالشحن الدوري (نعم/لا)',
    review_date: 'التاريخ (موعد أو آخر صرف)', notes: 'ملاحظات' };
  const SYN = {
    mrn: ['رقم الملف', 'رقم الملف الطبي', 'الملف الطبي', 'الملف', 'رقم المريض', 'mrn', 'file no', 'file number', 'medical record', 'medical record number', 'رقم السجل', 'السجل الطبي'],
    name: ['الاسم', 'اسم المريض', 'الاسم الرباعي', 'اسم المستلم', 'الاسم الكامل', 'name', 'patient name', 'full name', 'المريض'],
    phone2: ['جوال اضافي', 'رقم جوال اضافي', 'رقم اضافي', 'جوال 2', 'جوال2', 'رقم اخر', 'جوال اخر', 'phone2', 'phone 2', 'mobile 2', 'رقم التواصل 2', 'هاتف اضافي'],
    phone: ['الجوال', 'رقم الجوال', 'جوال', 'الهاتف', 'رقم الهاتف', 'رقم التواصل', 'الموبايل', 'mobile', 'phone', 'رقم المستلم', 'جوال المستلم', 'mobile number', 'phone number'],
    clinic: ['العياده', 'عياده', 'clinic', 'القسم', 'department'],
    city: ['المدينه', 'مدينه', 'city', 'المدينه المستلم'],
    address: ['العنوان', 'العنوان الوطني', 'العنوان المختصر', 'العنوان الوطني المختصر', 'address', 'short address', 'national address', 'عنوان المستلم'],
    review_date: ['تاريخ المراجعه', 'تاريخ الاستحقاق', 'الموعد', 'موعد الصرف', 'تاريخ الصرف القادم', 'تاريخ الصرف', 'الاستحقاق', 'review date', 'due date', 'date', 'التاريخ', 'موعد المراجعه', 'تاريخ الاستحقاق القادم'],
    notes: ['ملاحظات', 'ملاحظه', 'notes', 'note', 'remarks'],
    district: ['الحي', 'حي', 'district', 'neighborhood', 'الحي السكني'],
    pharmacy: ['الصيدليه', 'صيدليه', 'pharmacy', 'الصيدليه المستلمه'],
    periodic: ['نعم لا', 'نعم او لا', 'الشحن الدوري', 'شحن دوري', 'دوري', 'يرغب بالشحن'],
  };
  const norm = (o) => Object.fromEntries(Object.entries(o).map(([k, v]) => [k, v.map(normHeader)]));
  const SYN_N = norm(SYN);
  function bestMatch(header, table, dflt) {
    const h = normHeader(header);
    if (!h) return [dflt, 0];
    let best = dflt, score = 0;
    for (const [f, syns] of Object.entries(table)) for (const s of syns) {
      let sc = 0;
      if (h === s) sc = 100 + s.length; else if (s.length >= 3 && h.includes(s)) sc = 50 + s.length; else continue;
      if (sc > score) { best = f; score = sc; }
    }
    return [best, score];
  }
  const matchField = (h) => bestMatch(h, SYN_N, null);
  function autoMapping(headers) {
    const c = [];
    headers.forEach((h, i) => { const [f, sc] = matchField(h); if (f) c.push([sc, i, f]); });
    c.sort((a, b) => b[0] - a[0] || b[1] - a[1]);
    const mapping = {}, used = new Set();
    for (const [, i, f] of c) { if (f in mapping || used.has(i)) continue; mapping[f] = i; used.add(i); }
    return mapping;
  }

  // ---------------- ألوان
  const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
  const IDX = '000000,FFFFFF,FF0000,00FF00,0000FF,FFFF00,FF00FF,00FFFF,000000,FFFFFF,FF0000,00FF00,0000FF,FFFF00,FF00FF,00FFFF,800000,008000,000080,808000,800080,008080,C0C0C0,808080,9999FF,993366,FFFFCC,CCFFFF,660066,FF8080,0066CC,CCCCFF,000080,FF00FF,FFFF00,00FFFF,800080,800000,008080,0000FF,00CCFF,CCFFFF,CCFFCC,FFFF99,99CCFF,FF99CC,CC99FF,FFCC99,3366FF,33CCCC,99CC00,FFCC00,FF9900,FF6600,666699,969696,003366,339966,003300,333300,993300,993366,333399,333333'.split(',');
  const COLOR_LABELS = { red: 'أحمر', orange: 'برتقالي', yellow: 'أصفر', green: 'أخضر', blue: 'أزرق', purple: 'بنفسجي', none: 'بدون لون' };
  const COLOR_CLASSES = ['red', 'orange', 'yellow', 'green', 'blue', 'purple'];
  function colorHex(a, theme) {
    if (!a) return null;
    const rgb = attr(a, 'rgb'), th = attr(a, 'theme'), ix = attr(a, 'indexed');
    if (rgb) { if (/^00000000$/i.test(rgb)) return null; return rgb.slice(-6).toUpperCase(); }
    if (th != null) return theme[+th] || null;
    if (ix != null) return IDX[+ix] || null;
    return null;
  }
  function classify(hx) {
    if (!hx) return 'none';
    const r = parseInt(hx.slice(0, 2), 16) / 255, g = parseInt(hx.slice(2, 4), 16) / 255, b = parseInt(hx.slice(4, 6), 16) / 255;
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2;
    let h = 0, s = 0;
    if (mx !== mn) {
      const d = mx - mn;
      s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
      h = mx === r ? ((g - b) / d + (g < b ? 6 : 0)) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
      h /= 6;
    }
    if (s < 0.22 || l > 0.96 || l < 0.08) return 'none';
    const d = h * 360;
    if (d < 15 || d >= 325) return 'red'; if (d < 40) return 'orange'; if (d < 75) return 'yellow';
    if (d < 165) return 'green'; if (d < 260) return 'blue'; return 'purple';
  }

  // ---------------- قيم
  function cellText(v) {
    if (v == null) return '';
    if (typeof v === 'number') return Number.isInteger(v) ? String(v) : String(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return normDigits(String(v)).trim();
  }
  const pad = (n) => String(n).padStart(2, '0');
  function serialToISO(n) { const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * 864e5); return d.toISOString().slice(0, 10); }
  function validYMD(y, m, d) { const dt = new Date(Date.UTC(y, m - 1, d)); return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d; }
  function parseDate(v) {
    if (v == null || v === '') return [null, null];
    if (typeof v === 'number') return v > 20000 && v < 80000 ? [serialToISO(v), null] : [null, 'تاريخ غير مفهوم'];
    let s = normDigits(v).trim().replace(/\s*(م|هـ|ه)$/, '');
    if (/^\d{5}$/.test(s)) return parseDate(+s);
    let y, mo, d, m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T].*)?$/);
    if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; } else {
      m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(?:[ T].*)?$/);
      if (!m) return [null, 'تاريخ غير مفهوم'];
      d = +m[1]; mo = +m[2]; y = +m[3];
    }
    if (y >= 1300 && y <= 1500) return [null, 'تاريخ هجري - يُرجى إدخاله ميلاديًا'];
    if (!validYMD(y, mo, d)) return [null, 'تاريخ غير صحيح'];
    return [`${y}-${pad(mo)}-${pad(d)}`, null];
  }
  function normPhone(s) {
    s = normDigits(s || '').replace(/\D/g, '');
    if (!s) return ['', null];
    if (s.startsWith('00966')) s = s.slice(5); else if (s.startsWith('966')) s = s.slice(3);
    if (s.length === 9 && s.startsWith('5')) s = '0' + s;
    return /^05\d{8}$/.test(s) ? [s, null] : [s, 'رقم جوال غير صحيح'];
  }
  const normMrn = (s) => normDigits(s || '').replace(/\s+/g, '').trim();

  // ---------------- قارئ xlsx خفيف (يقرأ الألوان)
  async function openXlsx(buf) {
    const zip = await JSZip.loadAsync(buf);
    const txt = async (p) => { const f = zip.file(p); return f ? f.async('string') : ''; };
    const wbx = await txt('xl/workbook.xml'), rels = await txt('xl/_rels/workbook.xml.rels');
    const sheets = [...wbx.matchAll(/<sheet\b[^>]*>/g)].map((m) => ({ name: unesc(attr(m[0], 'name') || ''), rid: attr(m[0], 'r:id') }));
    const relMap = {};
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const t = attr(m[0], 'Target') || '';
      relMap[attr(m[0], 'Id')] = t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, '');
    }
    return { zip, txt, sheets, pathOf: (name) => relMap[(sheets.find((s) => s.name === name) || sheets[0]).rid] || 'xl/worksheets/sheet1.xml' };
  }
  async function readStyles(x) {
    const st = await x.txt('xl/styles.xml'), th = await x.txt('xl/theme/theme1.xml');
    let theme = DEFAULT_THEME;
    const cs = th.match(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/);
    if (cs) {
      const get = (tag) => { const m = cs[0].match(new RegExp(`<a:${tag}>([\\s\\S]*?)</a:${tag}>`)); const c = m && m[1].match(/(?:val|lastClr)="([0-9A-Fa-f]{6})"/); return c ? c[1].toUpperCase() : null; };
      const v = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map(get);
      theme = [v[1], v[0], v[3], v[2], ...v.slice(4)].map((c, i) => c || DEFAULT_THEME[i]);
    }
    const numFmts = {};
    for (const m of st.matchAll(/<numFmt\b[^>]*>/g)) numFmts[attr(m[0], 'numFmtId')] = unesc(attr(m[0], 'formatCode') || '');
    const fills = [], fonts = [];
    const fs = st.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/);
    if (fs) for (const m of fs[1].matchAll(/<fill>([\s\S]*?)<\/fill>|<fill\/>/g)) {
      const body = m[1] || '';
      const pt = body.match(/<patternFill\b([^>]*)/), fg = body.match(/<fgColor\b([^>]*?)\/?>/), bg = body.match(/<bgColor\b([^>]*?)\/?>/);
      const type = pt ? attr(pt[1], 'patternType') : null;
      let hx = null;
      if (type && type !== 'none') { hx = colorHex(fg && fg[1], theme); if (!hx && type !== 'solid') hx = colorHex(bg && bg[1], theme); }
      fills.push(hx);
    }
    const fo = st.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/);
    if (fo) for (const m of fo[1].matchAll(/<font\b[^>]*?(?:\/>|>([\s\S]*?)<\/font>)/g)) {
      const c = (m[1] || '').match(/<color\b([^>]*?)\/?>/);
      fonts.push(c ? colorHex(c[1], theme) : null);
    }
    const xfs = [];
    const cx = st.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
    const isDateFmt = (id) => {
      id = +id;
      if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47)) return true;
      const code = numFmts[id];
      return !!code && /[dmy]/i.test(code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''));
    };
    if (cx) for (const m of cx[1].matchAll(/<xf\b([^>]*?)(?:\/>|>)/g)) {
      xfs.push({ fill: fills[+(attr(m[1], 'fillId') || 0)] || null, font: fonts[+(attr(m[1], 'fontId') || 0)] || null, date: isDateFmt(attr(m[1], 'numFmtId') || 0) });
    }
    return { xfs };
  }
  async function readShared(x) {
    const ss = await x.txt('xl/sharedStrings.xml');
    const out = [];
    for (const m of ss.matchAll(/<si>([\s\S]*?)<\/si>|<si\/>/g)) {
      const body = (m[1] || '').replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
      out.push([...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>|<t\/>/g)].map((t) => unesc(t[1] || '')).join(''));
    }
    return out;
  }
  const colIndex = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };
  const colLetter = (n) => { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };

  // يعيد صفوف الورقة: [{row, cells: [{v, fill, font, type}]}]
  async function readSheetRows(buf, sheetName, maxRows = 20000, maxCols = 60) {
    const x = await openXlsx(buf);
    const name = sheetName && x.sheets.some((s) => s.name === sheetName) ? sheetName : (x.sheets[0] || {}).name;
    const path = x.pathOf(name);
    const [{ xfs }, shared, xml] = await Promise.all([readStyles(x), readShared(x), x.txt(path)]);
    const rows = [];
    for (const rm of xml.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const rn = +(attr(rm[1], 'r') || rows.length + 1);
      if (rows.length >= maxRows) break;
      const cells = [];
      let auto = 0;
      for (const cm of (rm[2] || '').matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const a = cm[1], body = cm[2] || '';
        const ref = attr(a, 'r');
        const ci = ref ? colIndex(ref) : auto;
        auto = ci + 1;
        if (ci >= maxCols) continue;
        const t = attr(a, 't'), s = +(attr(a, 's') || 0), xf = xfs[s] || {};
        const vm = body.match(/<v>([\s\S]*?)<\/v>/);
        let v = null, type = 'empty';
        if (t === 's') { v = vm ? shared[+vm[1]] ?? '' : ''; type = 'str'; }
        else if (t === 'inlineStr') { v = [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((q) => unesc(q[1])).join(''); type = 'str'; }
        else if (t === 'str') { v = vm ? unesc(vm[1]) : ''; type = 'str'; }
        else if (t === 'b') { v = vm ? vm[1] === '1' : null; type = 'bool'; }
        else if (t === 'e') { v = null; }
        else if (vm) { v = parseFloat(vm[1]); type = 'num'; if (xf.date && v > 0 && v < 2958466) { v = serialToISO(v); type = 'date'; } }
        cells[ci] = { v, fill: xf.fill || null, font: xf.font || null, type };
      }
      rows.push({ row: rn, cells });
    }
    return { sheets: x.sheets.map((s) => s.name), sheet: name, path, rows };
  }

  async function readWorkbook(buf, sheet, useFont = true) {
    const { sheets, sheet: sname, rows: raw } = await readSheetRows(buf, sheet);
    let hdr = 0, best = 0;
    raw.slice(0, 15).forEach((r, i) => {
      const n = r.cells.filter((c) => c && matchField(cellText(c.v))[0]).length;
      if (n > best) { best = n; hdr = i; }
    });
    let headers = raw.length ? Array.from({ length: raw[hdr].cells.length }, (_, i) => cellText(raw[hdr].cells[i]?.v)) : [];
    while (headers.length && !headers[headers.length - 1]) headers.pop();
    const ncol = Math.max(headers.length, 1);
    headers = headers.map((h, i) => h || `عمود ${i + 1}`);
    const rows = [];
    for (const r of raw.slice(hdr + 1)) {
      const vals = Array.from({ length: ncol }, (_, i) => r.cells[i] ? r.cells[i].v : null);
      if (vals.every((v) => cellText(v) === '')) continue;
      let classes = [];
      for (let i = 0; i < ncol; i++) { const c = r.cells[i]; if (c && c.fill) { const k = classify(c.fill); if (k !== 'none') classes.push([k, c.fill]); } }
      let source = 'fill';
      if (!classes.length && useFont) {
        for (let i = 0; i < ncol; i++) { const c = r.cells[i]; if (c && c.font && cellText(c.v) !== '') { const k = classify(c.font); if (k !== 'none') classes.push([k, c.font]); } }
        source = 'font';
      }
      let color = 'none', hex = null;
      if (classes.length) {
        const cnt = {}; classes.forEach(([k]) => cnt[k] = (cnt[k] || 0) + 1);
        color = Object.entries(cnt).sort((a, b) => b[1] - a[1])[0][0];
        hex = classes.find(([k]) => k === color)[1];
      } else source = null;
      rows.push({ row: r.row, values: vals, color, hex, color_source: source });
    }
    return { sheets, sheet: sname, header_row: raw[hdr] ? raw[hdr].row : 1, headers, rows, mapping: autoMapping(headers) };
  }

  function extractRow(r, mapping) {
    const vals = r.values;
    const get = (f) => { const i = mapping[f]; return i == null || i === '' || +i >= vals.length ? null : vals[+i]; };
    const out = {}, errors = [], warns = [];
    out.mrn = normMrn(cellText(get('mrn')));
    out.name = cellText(get('name'));
    let [p, e] = normPhone(cellText(get('phone'))); out.phone = p; if (e) warns.push('الجوال: ' + e);
    let [p2, e2] = normPhone(cellText(get('phone2'))); out.phone2 = p2; if (e2) warns.push('الجوال الإضافي: ' + e2);
    out.clinic = cellText(get('clinic')); out.city = cellText(get('city'));
    out.address = cellText(get('address')).replace(/[^A-Za-z0-9]/g, '').toUpperCase();
    out.district = cellText(get('district')); out.pharmacy = cellText(get('pharmacy'));
    const pv = cellText(get('periodic')).trim();
    out.periodic = ['نعم', 'ن', 'yes', 'Yes', 'Y', '1'].includes(pv) ? 1 : ['لا', 'no', 'No', 'N', '0'].includes(pv) ? 0 : null;
    if (pv && out.periodic == null) warns.push('عمود الشحن الدوري: قيمة غير مفهومة');
    out.notes = cellText(get('notes'));
    let rd = get('review_date');
    if (typeof rd === 'string' && /^\d{4}-\d{2}-\d{2}/.test(rd)) rd = rd.slice(0, 10);
    const [d, de] = parseDate(rd);
    out.review_date = d; out.date_raw = de ? cellText(rd) : '';
    if (de) warns.push('التاريخ: ' + de);
    if (out.address && !/^[A-Z]{4}\d{4}$/.test(out.address)) warns.push('العنوان الوطني بصيغة غير معتادة');
    if (!out.mrn) errors.push('رقم الملف مفقود');
    if (!out.name) warns.push('الاسم مفقود');
    return [out, errors, warns];
  }

  // ---------------- التصدير
  const EXPORT_FIELDS = { seq: 'مسلسل', mrn: 'رقم الملف الطبي', name: 'الاسم', phone: 'رقم الجوال', phone_intl: 'رقم الجوال (966)', phone2: 'رقم جوال إضافي',
    clinic: 'العيادة', city: 'المدينة', address: 'العنوان الوطني المختصر', district: 'الحي', pharmacy: 'الصيدلية', name_phone: 'الاسم + الجوال (9 أرقام)',
    phone9: 'الجوال بدون صفر (9 أرقام)', ref_seq: 'رقم مرجعي متسلسل', out_of_city: 'خارج المدينة (نعم/لا)', periodic: 'يرغب بالشحن الدوري (نعم/لا)',
    review_date: 'تاريخ الصرف القادم', last_note: 'ملاحظة الصيدلي', notes: 'ملاحظات الملف', assigned_name: 'الموظف المسؤول', export_date: 'تاريخ التصدير', empty: '(فارغ)' };
  const EXP_SYN = Object.fromEntries(Object.entries(SYN).filter(([k]) => k !== 'review_date'));
  EXP_SYN.seq = ['م', 'مسلسل', 'الرقم التسلسلي', '#', 'no', 'رقم'];
  EXP_SYN.export_date = ['تاريخ التصدير', 'تاريخ الشحن', 'تاريخ الطلب', 'التاريخ', 'date'];
  const EXP_N = norm(EXP_SYN);
  const exportMatch = (h) => bestMatch(h, EXP_N, 'empty')[0];

  function inferFromSamples(hf, vals) {
    vals = vals.filter((v) => v != null && v !== '');
    if (!vals.length) return [hf, null];
    const txt = vals.map(cellText);
    if (txt.every((t) => /^\D+\s+\d{9}$/.test(t.trim()))) return ['name_phone', null];
    if (!['empty', 'seq'].includes(hf)) return [hf, null];
    if (txt.length >= 2 && txt.every((t) => /^\d+$/.test(t))) {
      const n = txt.map(Number);
      if (n.every((x, i) => i === 0 || x === n[i - 1] + 1)) return ['ref_seq', n[0]];
    }
    if (hf === 'empty' && new Set(txt).size === 1) return ['fixed:' + txt[0], null];
    return [hf, null];
  }

  async function readTemplateHeaders(buf) {
    const { sheet, path, rows } = await readSheetRows(buf, null, 40, 60);
    const byRow = new Map(rows.map((r) => [r.row, r]));
    const cellsOf = (rn) => (byRow.get(rn) || { cells: [] }).cells;
    let hdr = 1, best = -1;
    for (let rn = 1; rn <= 15; rn++) {
      const vals = cellsOf(rn).map((c) => cellText(c && c.v));
      const n = vals.filter(Boolean).length, known = vals.filter((v) => v && exportMatch(v) !== 'empty').length;
      const sc = known * 10 + n;
      if (n && sc > best) { best = sc; hdr = rn; }
    }
    const samples = [1, 2, 3, 4, 5].map((k) => hdr + k);
    const cols = []; let refStart = null;
    const hc = cellsOf(hdr);
    for (let ci = 0; ci < hc.length; ci++) {
      const h = cellText(hc[ci] && hc[ci].v);
      if (!h) continue;
      const sv = samples.map((rn) => (cellsOf(rn)[ci] || {}).v);
      const [field, start] = inferFromSamples(exportMatch(h), sv);
      if (start != null) refStart = start;
      const typ = samples.some((rn) => (cellsOf(rn)[ci] || {}).type === 'num') ? 'num' : 'str';
      cols.push({ col: ci + 1, header: h, field, type: typ });
    }
    return { sheet, header_row: hdr, columns: cols, ref_start: refStart || 1, sheet_path: path };
  }

  const TEXT_FIELDS = new Set(['mrn', 'phone', 'phone2', 'phone_intl', 'phone9', 'address']);
  function fieldValue(p, field, i, exportDate, ctx = {}) {
    if (field.startsWith('fixed:')) return field.slice(6);
    if (field === 'seq') return i + 1;
    if (field === 'ref_seq') return (+ctx.ref_start || 1) + i;
    const ph = p.phone || '', ph9 = ph.startsWith('05') ? ph.slice(1) : ph;
    if (field === 'phone_intl') return ph.startsWith('05') ? '966' + ph.slice(1) : ph;
    if (field === 'phone9') return ph9;
    if (field === 'name_phone') return ((p.name || '').trim() + ' ' + ph9).trim();
    if (field === 'out_of_city') return p.out_of_city ? 'نعم' : 'لا';
    if (field === 'periodic') return p.periodic === 1 ? 'نعم' : p.periodic === 0 ? 'لا' : '';
    if (field === 'export_date') return exportDate;
    if (field === 'city') return p.city || ctx.home_city || '';
    if (field === 'empty') return null;
    return p[field] ?? '';
  }

  // ---- كاتب ZIP (بدون ضغط) يطابق رؤوس ملف القالب
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (u8) => { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function parseZipEntries(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= 0; i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    const n = dv.getUint16(eocd + 10, true); let pos = dv.getUint32(eocd + 16, true);
    const out = [];
    for (let k = 0; k < n; k++) {
      const nl = dv.getUint16(pos + 28, true), el = dv.getUint16(pos + 30, true), cl = dv.getUint16(pos + 32, true);
      out.push({ vmade: dv.getUint16(pos + 4, true), vneed: dv.getUint16(pos + 6, true), flags: dv.getUint16(pos + 8, true), time: dv.getUint16(pos + 12, true),
        date: dv.getUint16(pos + 14, true), iattr: dv.getUint16(pos + 36, true), eattr: dv.getUint32(pos + 38, true),
        name: new TextDecoder().decode(u8.subarray(pos + 46, pos + 46 + nl)) });
      pos += 46 + nl + el + cl;
    }
    return out;
  }
  function writeStoredZip(entries) { // entries: [{name, data(Uint8Array), hdr}]
    const enc = new TextEncoder(), parts = [], central = [];
    let off = 0;
    for (const e of entries) {
      const nm = enc.encode(e.name), h = e.hdr || { vmade: 20, vneed: 20, flags: 0x800, time: 0, date: 33, iattr: 0, eattr: 0 };
      const crc = crc32(e.data), sz = e.data.length;
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, h.vneed, true); lh.setUint16(6, h.flags, true); lh.setUint16(8, 0, true);
      lh.setUint16(10, h.time, true); lh.setUint16(12, h.date, true); lh.setUint32(14, crc, true); lh.setUint32(18, sz, true); lh.setUint32(22, sz, true);
      lh.setUint16(26, nm.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), nm, e.data);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, h.vmade, true); ch.setUint16(6, h.vneed, true); ch.setUint16(8, h.flags, true); ch.setUint16(10, 0, true);
      ch.setUint16(12, h.time, true); ch.setUint16(14, h.date, true); ch.setUint32(16, crc, true); ch.setUint32(20, sz, true); ch.setUint32(24, sz, true);
      ch.setUint16(28, nm.length, true); ch.setUint16(30, 0, true); ch.setUint16(32, 0, true); ch.setUint16(34, 0, true);
      ch.setUint16(36, h.iattr, true); ch.setUint32(38, h.eattr, true); ch.setUint32(42, off, true);
      central.push(new Uint8Array(ch.buffer), nm);
      off += 30 + nm.length + sz;
    }
    const cdSize = central.reduce((n, a) => n + a.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, off, true);
    const all = [...parts, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((n, a) => n + a.length, 0));
    let p = 0; for (const a of all) { out.set(a, p); p += a.length; }
    return out;
  }

  async function rawSheet(tplBuf, cfg, plist, exportDate, ctx) {
    const u8 = new Uint8Array(tplBuf);
    const zip = await JSZip.loadAsync(u8);
    const hdrs = parseZipEntries(u8);
    const path = cfg.sheet_path || 'xl/worksheets/sheet1.xml';
    const xml = await zip.file(path).async('string');
    const hr = +cfg.header_row || 1;
    const m = xml.match(/<sheetData\s*\/>|<sheetData>([\s\S]*?)<\/sheetData>/);
    const inner = m ? (m[1] || '') : '';
    const kept = [];
    for (const rm of inner.matchAll(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)) {
      const rn = rm[0].match(/\br="(\d+)"/);
      if (rn && +rn[1] <= hr) kept.push(rm[0]);
    }
    const cols = [...(cfg.columns || [])].sort((a, b) => a.col - b.col);
    const lastCol = Math.max(1, ...cols.map((c) => +c.col));
    const rowsXml = plist.map((p, i) => {
      const r = hr + 1 + i;
      return `<row r="${r}">` + cols.map((c) => {
        const f = c.field || 'empty', v = fieldValue(p, f, i, exportDate, ctx), ref = colLetter(+c.col) + r;
        if (v == null) return '';
        if (c.type === 'num' && !TEXT_FIELDS.has(f) && /^-?\d+(\.\d+)?$/.test(String(v))) return `<c r="${ref}"><v>${v}</v></c>`;
        return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xesc(v)}</t></is></c>`;
      }).join('') + '</row>';
    });
    let nx = m ? xml.slice(0, m.index) + '<sheetData>' + kept.join('') + rowsXml.join('') + '</sheetData>' + xml.slice(m.index + m[0].length) : xml;
    nx = nx.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:${colLetter(lastCol)}${Math.max(hr + plist.length, hr)}"/>`);
    const entries = [];
    for (const h of hdrs) {
      const f = zip.file(h.name);
      if (!f) continue;
      const data = h.name === path ? new TextEncoder().encode(nx) : await f.async('uint8array');
      entries.push({ name: h.name, data, hdr: h });
    }
    return writeStoredZip(entries);
  }

  function simpleXlsx(headers, rows, sheetName = 'Sheet1') { // ملف بسيط بدون قالب
    const enc = new TextEncoder();
    const cell = (ref, v) => (typeof v === 'number') ? `<c r="${ref}"><v>${v}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xesc(v ?? '')}</t></is></c>`;
    const all = [headers, ...rows];
    const sd = all.map((r, i) => `<row r="${i + 1}">` + r.map((v, j) => cell(colLetter(j + 1) + (i + 1), v)).join('') + '</row>').join('');
    const files = {
      '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
      '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
      'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${xesc(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
      'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews><sheetData>${sd}</sheetData></worksheet>`,
    };
    return writeStoredZip(Object.entries(files).map(([name, s]) => ({ name, data: enc.encode(s) })));
  }

  async function buildExport(patients, cfg, tplBuf, exportDate, split, ctx) {
    let groups = [['المستحقون', patients]];
    if (split) groups = [['داخل_المدينة', patients.filter((p) => !p.out_of_city)], ['خارج_المدينة', patients.filter((p) => p.out_of_city)]].filter((g) => g[1].length);
    const one = async (plist) => {
      if (tplBuf) return rawSheet(tplBuf, cfg, plist, exportDate, ctx);
      const cols = cfg.columns || [];
      return simpleXlsx(cols.map((c) => c.header || EXPORT_FIELDS[c.field] || ''), plist.map((p, i) => cols.map((c) => {
        const f = c.field || 'empty'; let v = fieldValue(p, f, i, exportDate, ctx);
        if (v == null) return '';
        if (!TEXT_FIELDS.has(f) && c.type === 'num' && /^-?\d+$/.test(String(v))) return +v;
        return String(v);
      })));
    };
    if (groups.length === 1) return [await one(groups[0][1]), 'xlsx'];
    const z = new JSZip();
    for (const [n, pl] of groups) z.file(n + '.xlsx', await one(pl));
    return [await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), 'zip'];
  }

  return { normDigits, normHeader, normPhone, normMrn, cellText, parseDate, classify, readWorkbook, extractRow, readTemplateHeaders, buildExport,
    simpleXlsx, FIELD_LABELS, COLOR_LABELS, COLOR_CLASSES, EXPORT_FIELDS };
})();

/* الخادم المحلي داخل المتصفح — منقول من app.py، والبيانات تُحفظ في IndexedDB على هذا الجهاز */
'use strict';
const LOCAL = (() => {
  // ------------------------------------------------ تخزين
  const IDB = {
    db: null,
    open() {
      if (this.db) return Promise.resolve(this.db);
      return new Promise((res, rej) => {
        const r = indexedDB.open('pharma-system', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('kv');
        r.onsuccess = () => { this.db = r.result; res(this.db); };
        r.onerror = () => rej(r.error);
      });
    },
    async get(k) { const db = await this.open(); return new Promise((res, rej) => { const t = db.transaction('kv').objectStore('kv').get(k); t.onsuccess = () => res(t.result); t.onerror = () => rej(t.error); }); },
    async set(k, v) { const db = await this.open(); return new Promise((res, rej) => { const tx = db.transaction('kv', 'readwrite'); tx.objectStore('kv').put(v, k); tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error); }); },
    async setMany(obj) { const db = await this.open(); return new Promise((res, rej) => { const tx = db.transaction('kv', 'readwrite'); const st = tx.objectStore('kv'); for (const [k, v] of Object.entries(obj)) st.put(v, k); tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error); }); },
    async del(k) { const db = await this.open(); return new Promise((res, rej) => { const tx = db.transaction('kv', 'readwrite'); tx.objectStore('kv').delete(k); tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error); }); },
  };

  // مخزن البيانات والجلسة قابلان للاستبدال: المتصفح (IndexedDB) أو الخادم (Netlify Blobs + كوكيز)
  const ENV = { server: false, seedAdminHash: null };
  let KV = IDB;
  const SESS = {
    get: (k) => { try { return sessionStorage.getItem(k); } catch (e) { return null; } },
    set: (k, v) => { try { sessionStorage.setItem(k, v); } catch (e) {} },
    del: (k) => { try { sessionStorage.removeItem(k); } catch (e) {} },
  };
  const TABLES = ['users', 'patients', 'reviews', 'px_cases', 'px_contacts', 'requests', 'assignments_log', 'field_history', 'export_batches', 'export_items', 'audit'];
  let D = null; // قاعدة البيانات في الذاكرة
  let byId = {};
  const DEFAULT_SETTINGS = {
    org_name: 'التجمع الصحي', org_sub: 'صيدلية الرعاية المنزلية — طلبات العلاج والتوصيل', logo: '', brand_color: '#0B6E73',
    lead_days: 4, home_city: '', pharmacist_scope: 'assigned', refill_options: [30, 60, 90], refill_default: 30,
    color_actions: { red: 'not_eligible', yellow: 'out_of_city', green: 'ignore', orange: 'ignore', blue: 'ignore', purple: 'ignore' },
    export_template: { has_file: true, builtin: true, sheet: 'Sheet1', header_row: 1, sheet_path: 'xl/worksheets/sheet1.xml', file_name: 'نموذج البريد السعودي SPL (مدمج)', ref_start: 1001, columns: [
      { col: 1, header: 'اسم العميل', field: 'name_phone', type: 'str' }, { col: 2, header: 'رقم الجوال', field: 'phone', type: 'str' },
      { col: 3, header: 'عدد البوليصات', field: 'fixed:1', type: 'num' }, { col: 4, header: 'وزن الشحنه', field: 'fixed:1', type: 'num' },
      { col: 5, header: 'العنوان مختصر', field: 'address', type: 'str' }, { col: 6, header: 'المدينة', field: 'city', type: 'str' },
      { col: 7, header: 'الحي', field: 'district', type: 'str' }, { col: 8, header: 'الرقم المرجعي', field: 'ref_seq', type: 'str' },
      { col: 9, header: 'وصف الشحنه', field: 'fixed:ادوية', type: 'str' }] },
    wa_templates: {
      not_eligible: 'السلام عليكم {الاسم}،\nنفيدكم من {الجهة} بأن موعد استحقاق صرف علاجكم القادم بتاريخ {تاريخ_الاستحقاق}.\n{ملاحظة_للمريض}\nرقم الملف: {رقم_الملف}',
      no_rx: 'السلام عليكم {الاسم}،\nنفيدكم من {الجهة} بعدم وجود وصفة فعّالة في ملفكم حاليًا.\n{ملاحظة_للمريض}\nنأمل مراجعة العيادة لتجديد الوصفة. رقم الملف: {رقم_الملف}',
      wrong_data: 'السلام عليكم {الاسم}،\nنأمل تزويدنا بالبيانات الصحيحة التالية: {المطلوب_تصحيحه}\n{ملاحظة_للمريض}\nرقم الملف: {رقم_الملف}',
      needs_info: 'السلام عليكم {الاسم}،\nبخصوص طلبكم رقم {رقم_الطلب}، نأمل تزويدنا بما يلي: {المطلوب}',
    },
  };
  const clone = (o) => JSON.parse(JSON.stringify(o));
  const BUILTIN_TPL_B64 = 'UEsDBBQAAAgAAAAAIQBuYbgNLQIAAC0CAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbDw/eG1sIHZlcnNpb249IjEuMCIgZW5jb2Rpbmc9IlVURi04IiBzdGFuZGFsb25lPSJ5ZXMiPz48VHlwZXMgeG1sbnM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9wYWNrYWdlLzIwMDYvY29udGVudC10eXBlcyI+PERlZmF1bHQgRXh0ZW5zaW9uPSJyZWxzIiBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLXBhY2thZ2UucmVsYXRpb25zaGlwcyt4bWwiLz48RGVmYXVsdCBFeHRlbnNpb249InhtbCIgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3htbCIvPjxPdmVycmlkZSBQYXJ0TmFtZT0iL3hsL3dvcmtib29rLnhtbCIgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1vZmZpY2Vkb2N1bWVudC5zcHJlYWRzaGVldG1sLnNoZWV0Lm1haW4reG1sIi8+PE92ZXJyaWRlIFBhcnROYW1lPSIveGwvd29ya3NoZWV0cy9zaGVldDEueG1sIiBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50LnNwcmVhZHNoZWV0bWwud29ya3NoZWV0K3htbCIvPjwvVHlwZXM+UEsDBBQAAAgAAAAAIQCY2uuLJwEAACcBAAALAAAAX3JlbHMvLnJlbHM8P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJVVEYtOCIgc3RhbmRhbG9uZT0ieWVzIj8+PFJlbGF0aW9uc2hpcHMgeG1sbnM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9wYWNrYWdlLzIwMDYvcmVsYXRpb25zaGlwcyI+PFJlbGF0aW9uc2hpcCBJZD0icklkMSIgVHlwZT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcy9vZmZpY2VEb2N1bWVudCIgVGFyZ2V0PSJ4bC93b3JrYm9vay54bWwiLz48L1JlbGF0aW9uc2hpcHM+UEsDBBQAAAgAAAAAIQCdbEO9GwEAABsBAAAPAAAAeGwvd29ya2Jvb2sueG1sPD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pjx3b3JrYm9vayB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3NwcmVhZHNoZWV0bWwvMjAwNi9tYWluIiB4bWxuczpyPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9yZWxhdGlvbnNoaXBzIj48c2hlZXRzPjxzaGVldCBuYW1lPSJTaGVldDEiIHNoZWV0SWQ9IjEiIHI6aWQ9InJJZDEiLz48L3NoZWV0cz48L3dvcmtib29rPlBLAwQUAAAIAAAAACEAWv2CaygBAAAoAQAAGgAAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzPD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/PjxSZWxhdGlvbnNoaXBzIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvcGFja2FnZS8yMDA2L3JlbGF0aW9uc2hpcHMiPjxSZWxhdGlvbnNoaXAgSWQ9InJJZDEiIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvd29ya3NoZWV0IiBUYXJnZXQ9IndvcmtzaGVldHMvc2hlZXQxLnhtbCIvPjwvUmVsYXRpb25zaGlwcz5QSwMEFAAACAAAAAAhAKoj5x/3AwAA9wMAABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWw8P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJVVEYtOCIgc3RhbmRhbG9uZT0ieWVzIj8+PHdvcmtzaGVldCB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3NwcmVhZHNoZWV0bWwvMjAwNi9tYWluIj48c2hlZXRWaWV3cz48c2hlZXRWaWV3IHJpZ2h0VG9MZWZ0PSIxIiB3b3JrYm9va1ZpZXdJZD0iMCIvPjwvc2hlZXRWaWV3cz48c2hlZXREYXRhPjxyb3cgcj0iMSI+PGMgcj0iQTEiIHQ9ImlubGluZVN0ciI+PGlzPjx0IHhtbDpzcGFjZT0icHJlc2VydmUiPtin2LPZhSDYp9mE2LnZhdmK2YQgPC90PjwvaXM+PC9jPjxjIHI9IkIxIiB0PSJpbmxpbmVTdHIiPjxpcz48dCB4bWw6c3BhY2U9InByZXNlcnZlIj7YsdmC2YUg2KfZhNis2YjYp9mEPC90PjwvaXM+PC9jPjxjIHI9IkMxIiB0PSJpbmxpbmVTdHIiPjxpcz48dCB4bWw6c3BhY2U9InByZXNlcnZlIj7Yudiv2K8g2KfZhNio2YjZhNmK2LXYp9iqPC90PjwvaXM+PC9jPjxjIHI9IkQxIiB0PSJpbmxpbmVTdHIiPjxpcz48dCB4bWw6c3BhY2U9InByZXNlcnZlIj7ZiNiy2YYg2KfZhNi02K3ZhtmHPC90PjwvaXM+PC9jPjxjIHI9IkUxIiB0PSJpbmxpbmVTdHIiPjxpcz48dCB4bWw6c3BhY2U9InByZXNlcnZlIj7Yp9mE2LnZhtmI2KfZhiDZhdiu2KrYtdixPC90PjwvaXM+PC9jPjxjIHI9IkYxIiB0PSJpbmxpbmVTdHIiPjxpcz48dCB4bWw6c3BhY2U9InByZXNlcnZlIj7Yp9mE2YXYr9mK2YbYqTwvdD48L2lzPjwvYz48YyByPSJHMSIgdD0iaW5saW5lU3RyIj48aXM+PHQgeG1sOnNwYWNlPSJwcmVzZXJ2ZSI+2KfZhNit2Yo8L3Q+PC9pcz48L2M+PGMgcj0iSDEiIHQ9ImlubGluZVN0ciI+PGlzPjx0IHhtbDpzcGFjZT0icHJlc2VydmUiPtin2YTYsdmC2YUg2KfZhNmF2LHYrNi52Yo8L3Q+PC9pcz48L2M+PGMgcj0iSTEiIHQ9ImlubGluZVN0ciI+PGlzPjx0IHhtbDpzcGFjZT0icHJlc2VydmUiPtmI2LXZgSDYp9mE2LTYrdmG2Yc8L3Q+PC9pcz48L2M+PC9yb3c+PC9zaGVldERhdGE+PC93b3Jrc2hlZXQ+UEsBAhQAFAAACAAAAAAhAG5huA0tAgAALQIAABMAAAAAAAAAAAAAAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECFAAUAAAIAAAAACEAmNrriycBAAAnAQAACwAAAAAAAAAAAAAAAABeAgAAX3JlbHMvLnJlbHNQSwECFAAUAAAIAAAAACEAnWxDvRsBAAAbAQAADwAAAAAAAAAAAAAAAACuAwAAeGwvd29ya2Jvb2sueG1sUEsBAhQAFAAACAAAAAAhAFr9gmsoAQAAKAEAABoAAAAAAAAAAAAAAAAA9gQAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzUEsBAhQAFAAACAAAAAAhAKoj5x/3AwAA9wMAABgAAAAAAAAAAAAAAAAAVgYAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbFBLBQYAAAAABQAFAEUBAACDCgAAAAA=';
  const builtinTpl = () => Uint8Array.from(atob(BUILTIN_TPL_B64), (c) => c.charCodeAt(0)).buffer;
  async function tplBytes(tpl) { return (!tpl.builtin && (await KV.get('export_template'))) || builtinTpl(); }
  const reindex = () => { byId = {}; for (const t of TABLES) { byId[t] = new Map(); for (const r of D[t]) byId[t].set(r.id, r); } };
  // تزامن بين النوافذ: رقم إصدار يتغير مع كل حفظ، وأي نافذة ترى إصدارًا أحدث تعيد التحميل قبل التنفيذ
  let saveChain = Promise.resolve(), myVer = null;
  const save = () => (saveChain = saveChain.then(() => { myVer = Date.now() + ':' + Math.random(); return KV.setMany({ db: D, ver: myVer }); }));
  async function load() {
    if (D) {
      const v = await KV.get('ver');
      if (v === myVer) return;
      const fresh = await KV.get('db');
      if (fresh) { D = fresh; for (const t of TABLES) D[t] = D[t] || []; reindex(); }
      myVer = v;
      return;
    }
    myVer = await KV.get('ver');
    D = await KV.get('db');
    if (!D) {
      D = { seq: {}, settings: {}, first_password: null };
      for (const t of TABLES) D[t] = [];
      const pw = ENV.seedAdminHash ? null : genPassword();
      D.users.push({ id: nextId('users'), name: 'مدير النظام', username: 'admin', pass_hash: ENV.seedAdminHash || await hashPw(pw), role: 'admin', can_review: 1, active: 1, must_change_pw: 1, last_login: null, created_at: nowIso() });
      D.first_password = pw;
      myVer = 'init:' + Date.now();
      await KV.setMany({ db: D, ver: myVer });
    }
    for (const t of TABLES) D[t] = D[t] || [];
    reindex();
  }
  function nextId(t) { D.seq[t] = (D.seq[t] || 0) + 1; return D.seq[t]; }
  function insert(t, row) { row.id = nextId(t); D[t].push(row); byId[t].set(row.id, row); return row; }

  // ------------------------------------------------ أدوات
  class ApiError extends Error { constructor(msg, code = 400, extra = {}) { super(msg); this.code = code; this.extra = extra; } }
  const nowIso = () => new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const todayS = () => new Date(Date.now() + 3 * 3600e3).toISOString().slice(0, 10);
  const addDays = (s, n) => { const d = new Date(s + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const validDate = (s) => typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(new Date(s + 'T00:00:00Z'));
  const daysBetween = (a, b) => Math.round((new Date(b + 'T00:00:00Z') - new Date(a + 'T00:00:00Z')) / 864e5);
  function genPassword(n = 10) {
    const a = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
    for (;;) {
      const r = crypto.getRandomValues(new Uint32Array(n));
      const p = [...r].map((x) => a[x % a.length]).join('');
      if (/\d/.test(p) && /[A-Za-z]/.test(p)) return p;
    }
  }
  const hex = (u8) => [...new Uint8Array(u8)].map((b) => b.toString(16).padStart(2, '0')).join('');
  async function hashPw(pw, salt) {
    salt = salt || hex(crypto.getRandomValues(new Uint8Array(16)));
    const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveBits']);
    const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', salt: new TextEncoder().encode(salt), iterations: 120000, hash: 'SHA-256' }, key, 256);
    return 'p2$' + salt + '$' + hex(bits);
  }
  async function checkPw(hash, pw) {
    if (!hash) return false;
    if (hash.startsWith('p2$')) return (await hashPw(pw, hash.split('$')[1])) === hash;
    const [salt] = hash.split('$');   // صيغة قديمة (نسخة سابقة)
    const old = salt + '$' + hex(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(salt + ':' + pw)));
    return old === hash;
  }
  const S = (k) => {
    const v = D.settings[k];
    if (v === undefined || (k === 'export_template' && !v.has_file)) return clone(DEFAULT_SETTINGS[k]);
    if (DEFAULT_SETTINGS[k] && typeof DEFAULT_SETTINGS[k] === 'object' && !Array.isArray(DEFAULT_SETTINGS[k])) return { ...clone(DEFAULT_SETTINGS[k]), ...v };
    return v;
  };
  const setS = (k, v) => { D.settings[k] = v; };
  let USER = null;
  function audit(action, entity = null, entity_id = null, details = null, user = USER) {
    insert('audit', { user_id: user ? user.id : null, user_name: user ? user.name : 'بوابة المريض', action, entity, entity_id,
      details: details != null ? JSON.stringify(details) : null, at: nowIso() });
  }
  const userName = (id) => (id && byId.users.get(id) ? byId.users.get(id).name : null);
  const ROLE_LABELS = { admin: 'المدير', employee: 'موظف', px: 'تجربة المريض' };
  const STATUS_LABELS = { pending: 'بانتظار المراجعة', eligible: 'مستحق - بانتظار التصدير', exported: 'تم التصدير', not_eligible: 'غير مستحق حاليًا', no_rx: 'لا توجد وصفة فعّالة', wrong_data: 'بيانات خاطئة' };
  const PX_STATUS = { new: 'جديد', contacting: 'جارٍ التواصل', awaiting: 'بانتظار رد المريض', done: 'تمت المعالجة', unreachable: 'تعذر التواصل' };
  const PX_REASON = { not_eligible: 'غير مستحق حاليًا', no_rx: 'لا توجد وصفة فعّالة', wrong_data: 'بيانات خاطئة' };
  const FIELD_AR = { mrn: 'رقم الملف', name: 'الاسم', phone: 'الجوال', phone2: 'الجوال الإضافي', clinic: 'العيادة', city: 'المدينة', address: 'العنوان الوطني',
    district: 'الحي', pharmacy: 'الصيدلية', periodic: 'الشحن الدوري', out_of_city: 'خارج المدينة', notes: 'ملاحظات', review_date: 'تاريخ المراجعة' };
  const REQ_TYPES = { delivery: 'طلب توصيل علاج', shortage: 'علاج ناقص / لم يصل', complaint: 'تواصل / شكوى', treatment: 'طلب علاج', inquiry: 'استفسار' };
  const PORTAL_TYPES = ['delivery', 'shortage', 'complaint'];
  const REQ_STATUS = { new: 'جديد', in_progress: 'قيد المعالجة', needs_info: 'يحتاج استكمال بيانات', done: 'مكتمل', rejected: 'مغلق/مرفوض' };
  const DELIVERY_STATUS = { new: 'جديد', coordinating: 'قيد التنسيق', ready: 'جاهز للتوصيل', out: 'خرج للتوصيل', delivered: 'تم التسليم', failed: 'متعذر' };
  const ACTION_LABELS = { not_eligible: 'غير مستحق', eligible: 'مستحق', out_of_city: 'خارج المدينة', no_rx: 'لا توجد وصفة فعّالة', ignore: 'تجاهل اللون (بانتظار المراجعة)', skip: 'استبعاد الصف' };
  const like = (v, q) => String(v || '').toLowerCase().includes(String(q).toLowerCase());
  const isOpenReq = (r) => !['done', 'rejected'].includes(r.status);

  // ------------------------------------------------ الصلاحيات
  function currentUser() {
    const id = +SESS.get('sid');
    const u = id && byId.users.get(id);
    return u && u.active ? u : null;
  }
  function auth(o = {}) {
    const u = currentUser();
    if (!u) throw new ApiError('يجب تسجيل الدخول', 401);
    if (u.must_change_pw && !o.pw) throw new ApiError('يجب تغيير كلمة المرور أولًا', 403, { must_change_pw: true });
    if (o.roles && !o.roles.includes(u.role)) throw new ApiError('ليست لديك صلاحية لهذا الإجراء', 403);
    if (o.review && !(u.role === 'admin' || u.can_review)) throw new ApiError('هذا الإجراء خاص بالصيدلي', 403);
    USER = u;
    return u;
  }
  const isAdmin = () => USER.role === 'admin';
  const openCase = (pid) => { let c = null; for (const x of D.px_cases) if (x.patient_id === pid && !x.closed_at && (!c || x.id > c.id)) c = x; return c; };
  function visible(p) {
    const u = USER;
    if (u.role === 'admin') return true;
    if (u.role === 'px') return D.px_cases.some((c) => c.patient_id === p.id);
    if (u.can_review && S('pharmacist_scope') === 'all') return true;
    return p.assigned_to === u.id;
  }
  function getPatientChecked(id) { const p = byId.patients.get(+id); if (!p || !visible(p)) throw new ApiError('المريض غير موجود أو غير مسند إليك', 404); return p; }
  function decorate(p, ctx) {
    const lead = +S('lead_days'), t = todayS();
    const c = ctx ? ctx.cases.get(p.id) : openCase(p.id);
    const o = { ...p, assigned_name: userName(p.assigned_to), px_case_id: c ? c.id : null, px_status: c ? c.status : null, px_reason: c ? c.reason : null,
      open_requests: ctx ? (ctx.reqs.get(p.id) || 0) : D.requests.filter((r) => r.patient_id === p.id && isOpenReq(r)).length };
    o.status_label = STATUS_LABELS[o.status] || o.status;
    o.px_status_label = o.px_status ? PX_STATUS[o.px_status] : null;
    o.px_reason_label = o.px_reason ? PX_REASON[o.px_reason] : null;
    o.current_handler = o.px_case_id ? 'تجربة المريض' : (o.assigned_name || 'غير مسند');
    const rd = o.review_date;
    o.in_queue = !!(rd && ['pending', 'not_eligible', 'exported'].includes(o.status) && rd <= addDays(t, lead));
    o.overdue = !!(o.in_queue && rd < t);
    o.days_left = rd && validDate(rd) ? daysBetween(t, rd) : null;
    return o;
  }
  function decoCtx() {
    const cases = new Map(), reqs = new Map();
    for (const c of D.px_cases) if (!c.closed_at) { const x = cases.get(c.patient_id); if (!x || c.id > x.id) cases.set(c.patient_id, c); }
    for (const r of D.requests) if (r.patient_id && isOpenReq(r)) reqs.set(r.patient_id, (reqs.get(r.patient_id) || 0) + 1);
    return { cases, reqs };
  }
  const inQueue = (p, lim) => ['pending', 'not_eligible', 'exported'].includes(p.status) && p.review_date && p.review_date <= lim;

  // ------------------------------------------------ عمليات مشتركة
  function openLoad(uid) {
    const lim = addDays(todayS(), +S('lead_days'));
    let n = 0;
    for (const p of D.patients) if (p.assigned_to === uid && (['pending', 'eligible', 'no_rx', 'wrong_data'].includes(p.status) || (['not_eligible', 'exported'].includes(p.status) && p.review_date && p.review_date <= lim))) n++;
    for (const r of D.requests) if (r.assigned_to === uid && isOpenReq(r)) n++;
    return n;
  }
  function assignPatient(pid, uid, reason) {
    const p = byId.patients.get(pid);
    if (!p || p.assigned_to === uid) return false;
    const from = p.assigned_to;
    p.assigned_to = uid; p.updated_at = nowIso();
    for (const r of D.requests) if (r.patient_id === pid && isOpenReq(r)) r.assigned_to = uid;
    for (const c of D.px_cases) if (c.patient_id === pid && !c.closed_at) c.original_owner = uid;
    insert('assignments_log', { patient_id: pid, from_user: from, to_user: uid, by_user: USER.id, reason, at: nowIso() });
    return true;
  }
  function distribute(pids, uids, mode, reason) {
    if (!uids.length) throw new ApiError('اختر موظفًا واحدًا على الأقل');
    const valid = uids.filter((id) => { const u = byId.users.get(id); return u && u.role === 'employee' && u.active; });
    if (!valid.length) throw new ApiError('لا يوجد موظف نشط ضمن الاختيار');
    const loads = Object.fromEntries(valid.map((u) => [u, openLoad(u)]));
    const plan = new Map();
    if (mode === 'load') {
      for (const pid of pids) { const uid = valid.reduce((a, b) => (loads[b] < loads[a] ? b : a)); plan.set(pid, uid); loads[uid]++; }
    } else {
      const order = [...valid].sort((a, b) => loads[a] - loads[b]);
      pids.forEach((pid, i) => plan.set(pid, order[i % order.length]));
    }
    const counts = {};
    for (const [pid, uid] of plan) { assignPatient(pid, uid, reason); counts[userName(uid)] = (counts[userName(uid)] || 0) + 1; }
    return counts;
  }
  function applyChanges(pid, changes, source) {
    const p = byId.patients.get(pid), real = {};
    for (const [k, v] of Object.entries(changes)) {
      const old = p[k];
      if (k === 'out_of_city') { if (+(old || 0) === +(v || 0)) continue; }
      else if (k === 'periodic') { if ((old ?? null) === (v ?? null)) continue; }
      else if (String(old ?? '') === String(v ?? '')) continue;
      real[k] = v;
      insert('field_history', { patient_id: pid, field: k, old_value: old == null ? null : String(old), new_value: v == null ? null : String(v), by_user: USER.id, by_name: USER.name, source, at: nowIso() });
    }
    if (Object.keys(real).length) { Object.assign(p, real); p.updated_at = nowIso(); }
    return real;
  }
  function cleanFields(d, allowed) {
    const out = {};
    for (const k of allowed) {
      if (!(k in d)) continue;
      let v = d[k];
      if (k === 'phone' || k === 'phone2') { const [ph, e] = XL.normPhone(v); if (e && ph) throw new ApiError(`${FIELD_AR[k]}: ${e}`); v = ph; }
      else if (k === 'mrn') { v = XL.normMrn(v); if (!v) throw new ApiError('رقم الملف مطلوب'); }
      else if (k === 'address') v = XL.normDigits(v || '').replace(/[^A-Za-z0-9]/g, '').toUpperCase();
      else if (k === 'out_of_city') v = [1, true, '1', 'true'].includes(v) ? 1 : 0;
      else if (k === 'periodic') v = [1, true, '1', 'true', 'yes'].includes(v) ? 1 : [0, false, '0', 'false', 'no'].includes(v) ? 0 : null;
      else if (k === 'review_date') { v = v || null; if (v && !validDate(v)) throw new ApiError('تاريخ غير صحيح'); }
      else v = typeof v === 'string' ? v.trim() : v;
      out[k] = v;
    }
    return out;
  }
  function homeCityFlag(city) {
    const hc = XL.normHeader(S('home_city') || '');
    if (!hc || !city) return null;
    return XL.normHeader(city) === hc ? 0 : 1;
  }
  function openPxCase(pid, reason, note, pnote, wf, corr, nd, owner) {
    for (const c of D.px_cases) if (c.patient_id === pid && !c.closed_at) Object.assign(c, { closed_at: nowIso(), closed_by: USER.id, close_note: 'أُغلقت تلقائيًا بإحالة جديدة', status: 'done' });
    return insert('px_cases', { patient_id: pid, reason, pharmacist_note: note, patient_note: pnote, wrong_field: wf, correction: corr, next_date: nd, status: 'new',
      last_message: null, original_owner: owner, created_by: USER.id, created_at: nowIso(), updated_at: nowIso(), closed_at: null, closed_by: null, close_note: null }).id;
  }
  function fillTemplate(tpl, ctx) {
    let out = tpl || '';
    for (const [k, v] of Object.entries(ctx)) out = out.split('{' + k + '}').join(v || '');
    return out.replace(/\n\s*\n+/g, '\n').trim();
  }
  function caseMessage(c, p) {
    return fillTemplate(S('wa_templates')[c.reason], { 'الاسم': p.name, 'رقم_الملف': p.mrn, 'تاريخ_الاستحقاق': c.next_date || p.review_date || '', 'العيادة': p.clinic,
      'الجهة': S('org_name'), 'ملاحظة_للمريض': c.patient_note, 'المطلوب_تصحيحه': c.correction });
  }
  function waLink(phone, text) {
    let ph = String(phone || '').replace(/\D/g, '');
    if (ph.startsWith('05')) ph = '966' + ph.slice(1);
    return ph ? `https://wa.me/${ph}?text=${encodeURIComponent(text || '')}` : null;
  }
  const publicUser = (u) => ({ id: u.id, name: u.name, username: u.username, role: u.role, role_label: ROLE_LABELS[u.role], can_review: !!u.can_review,
    active: !!u.active, must_change_pw: !!u.must_change_pw, last_login: u.last_login, created_at: u.created_at });
  const fileBlob = (data, name, type, inline) => ({ __file: true, data, name, inline: !!inline, type: type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

  // ------------------------------------------------ المسارات
  const R = [];
  const route = (method, pattern, fn) => R.push({ method, re: new RegExp('^' + pattern.replace(/:(\w+)/g, '(?<$1>[^/]+)') + '$'), fn });

  // الدخول
  route('POST', '/api/login', async ({ body }) => {
    const u = D.users.find((x) => x.username.toLowerCase() === String(body.username || '').trim().toLowerCase());
    if (u && u.locked_until && u.locked_until > Date.now()) throw new ApiError('تم إيقاف الدخول مؤقتًا بسبب محاولات خاطئة متكررة، حاول بعد 10 دقائق', 429);
    if (!u || !(await checkPw(u.pass_hash, body.password || ''))) {
      if (u) { u.failed = (u.failed || 0) + 1; if (u.failed >= 6) { u.locked_until = Date.now() + 10 * 60e3; u.failed = 0; } await save(); }
      throw new ApiError('اسم المستخدم أو كلمة المرور غير صحيحة', 401);
    }
    u.failed = 0; u.locked_until = null;
    if (!u.active) throw new ApiError('الحساب معطّل، تواصل مع المدير', 403);
    SESS.set('sid', String(u.id));
    u.last_login = nowIso();
    audit('تسجيل دخول', 'user', u.id, null, u);
    return { ok: true, must_change_pw: !!u.must_change_pw };
  });
  route('POST', '/api/logout', () => { SESS.del('sid'); return { ok: true }; });
  route('GET', '/api/me', () => { const u = auth({ pw: true }); return { user: publicUser(u), settings: { org_name: S('org_name'), org_sub: S('org_sub'), logo: S('logo'), brand_color: S('brand_color'), lead_days: S('lead_days'), refill_options: S('refill_options'), refill_default: S('refill_default'), today: todayS() } }; });
  route('GET', '/api/public-info', async () => { await load(); return { org_name: S('org_name'), org_sub: S('org_sub'), logo: S('logo'), brand_color: S('brand_color'), first_password: ENV.server ? null : D.first_password }; });
  route('POST', '/api/change-password', async ({ body }) => {
    const u = auth({ pw: true });
    if (!(await checkPw(u.pass_hash, body.old_password || ''))) throw new ApiError('كلمة المرور الحالية غير صحيحة');
    const n = body.new_password || '';
    if (n.length < 8 || !/\d/.test(n) || !/\D/.test(n)) throw new ApiError('كلمة المرور يجب أن تكون 8 أحرف على الأقل وتحتوي حروفًا وأرقامًا');
    u.pass_hash = await hashPw(n); u.must_change_pw = 0;
    if (u.username === 'admin') D.first_password = null;
    audit('تغيير كلمة المرور', 'user', u.id);
    return { ok: true };
  });

  // المستخدمون
  route('GET', '/api/users', () => {
    auth();
    const rows = [...D.users].sort((a, b) => b.active - a.active || a.role.localeCompare(b.role) || a.name.localeCompare(b.name));
    if (!isAdmin()) return rows.map((u) => ({ id: u.id, name: u.name, role: u.role, active: !!u.active }));
    return rows.map((u) => ({ ...publicUser(u), assigned_count: D.patients.filter((p) => p.assigned_to === u.id).length }));
  });
  route('POST', '/api/users', async ({ body: d }) => {
    auth({ roles: ['admin'] });
    const name = (d.name || '').trim(), username = (d.username || '').trim();
    if (!name || !/^[A-Za-z0-9_.\-]{3,30}$/.test(username)) throw new ApiError('أدخل الاسم واسم مستخدم (حروف إنجليزية/أرقام، 3 أحرف على الأقل)');
    if (!ROLE_LABELS[d.role]) throw new ApiError('الدور غير صحيح');
    if (D.users.some((u) => u.username.toLowerCase() === username.toLowerCase())) throw new ApiError('اسم المستخدم مستخدم مسبقًا');
    const pw = genPassword();
    const u = insert('users', { name, username, pass_hash: await hashPw(pw), role: d.role, can_review: d.can_review || d.role === 'admin' ? 1 : 0, active: 1, must_change_pw: 1, last_login: null, created_at: nowIso() });
    audit('إنشاء حساب', 'user', u.id, { name, username, role: ROLE_LABELS[d.role], 'صيدلي': !!d.can_review });
    return { ok: true, id: u.id, temp_password: pw };
  });
  route('PATCH', '/api/users/:id', ({ p, body: d }) => {
    auth({ roles: ['admin'] });
    const u = byId.users.get(+p.id); if (!u) throw new ApiError('المستخدم غير موجود', 404);
    const ch = {};
    if (d.name && d.name.trim()) ch.name = d.name.trim();
    if (d.role && ROLE_LABELS[d.role] && d.role !== u.role) {
      if (u.role === 'admin' && D.users.filter((x) => x.role === 'admin' && x.active).length <= 1) throw new ApiError('لا يمكن تغيير دور آخر مدير في النظام');
      if (u.role === 'employee' && D.patients.some((x) => x.assigned_to === u.id)) throw new ApiError('أعد إسناد مرضى هذا الموظف قبل تغيير دوره');
      ch.role = d.role;
    }
    if ('can_review' in d) ch.can_review = d.can_review ? 1 : 0;
    Object.assign(u, ch);
    if (Object.keys(ch).length) audit('تعديل حساب', 'user', u.id, ch);
    return { ok: true };
  });
  route('POST', '/api/users/:id/reset-password', async ({ p }) => {
    auth({ roles: ['admin'] });
    const u = byId.users.get(+p.id); if (!u) throw new ApiError('المستخدم غير موجود', 404);
    const pw = genPassword(); u.pass_hash = await hashPw(pw); u.must_change_pw = 1;
    audit('إعادة تعيين كلمة المرور', 'user', u.id);
    return { ok: true, temp_password: pw };
  });
  route('POST', '/api/users/:id/deactivate', ({ p, body: d }) => {
    auth({ roles: ['admin'] });
    const u = byId.users.get(+p.id); if (!u) throw new ApiError('المستخدم غير موجود', 404);
    if (u.id === USER.id) throw new ApiError('لا يمكنك تعطيل حسابك');
    if (u.role === 'admin' && D.users.filter((x) => x.role === 'admin' && x.active).length <= 1) throw new ApiError('لا يمكن تعطيل آخر مدير');
    const pids = D.patients.filter((x) => x.assigned_to === u.id).map((x) => x.id);
    let moved = {};
    if (pids.length) {
      if (!d.reassign) throw new ApiError('لدى الموظف مرضى مسندون، اختر طريقة إعادة الإسناد', 409, { assigned_count: pids.length });
      const reason = 'تعطيل حساب ' + u.name;
      moved = d.reassign === 'auto' ? distribute(pids, D.users.filter((x) => x.role === 'employee' && x.active && x.id !== u.id).map((x) => x.id), 'load', reason)
        : distribute(pids, [+d.reassign], 'equal', reason);
    }
    u.active = 0;
    audit('تعطيل حساب', 'user', u.id, Object.keys(moved).length ? { 'أعيد إسناد': moved } : null);
    return { ok: true, moved };
  });
  route('POST', '/api/users/:id/activate', ({ p }) => { auth({ roles: ['admin'] }); const u = byId.users.get(+p.id); if (u) u.active = 1; audit('تفعيل حساب', 'user', +p.id); return { ok: true }; });

  // المرضى
  function filterPatients(a) {
    const t = todayS(), lim = addDays(t, +S('lead_days'));
    const q = (a.q || '').trim(), qd = XL.normDigits(q);
    const sts = a.status ? a.status.split(',').filter((s) => STATUS_LABELS[s]) : [];
    const ctx = decoCtx();
    return D.patients.filter((p) => {
      if (!visible(p)) return false;
      if (q && !(like(p.mrn, qd) || like(p.name, q) || like(p.phone, qd) || like(p.phone2, qd))) return false;
      if (a.assigned === 'none' && p.assigned_to) return false;
      if (a.assigned && /^\d+$/.test(a.assigned) && p.assigned_to !== +a.assigned) return false;
      if (a.clinic && p.clinic !== a.clinic) return false;
      if (sts.length && !sts.includes(p.status)) return false;
      if (['0', '1'].includes(a.periodic) && p.periodic !== +a.periodic) return false;
      if (['0', '1'].includes(a.out_of_city) && +(p.out_of_city || 0) !== +a.out_of_city) return false;
      if (a.from && validDate(a.from) && !(p.review_date && p.review_date >= a.from)) return false;
      if (a.to && validDate(a.to) && !(p.review_date && p.review_date <= a.to)) return false;
      if (['queue', 'overdue', 'upcoming'].includes(a.due)) {
        if (!inQueue(p, lim)) return false;
        if (a.due === 'overdue' && !(p.review_date < t)) return false;
        if (a.due === 'upcoming' && !(p.review_date > t)) return false;
      }
      if (a.due === 'nodate' && !(!p.review_date && ['pending', 'exported', 'not_eligible'].includes(p.status))) return false;
      if (a.px === '1' && !ctx.cases.has(p.id)) return false;
      return true;
    });
  }
  const sortByDate = (arr) => arr.sort((x, y) => (!x.review_date) - (!y.review_date) || String(x.review_date).localeCompare(String(y.review_date)) || x.id - y.id);
  route('GET', '/api/patients', ({ q: a }) => {
    auth({ roles: ['admin', 'employee', 'px'] });
    const all = sortByDate(filterPatients(a));
    const page = Math.max(+a.page || 1, 1), size = Math.min(+a.size || 50, 500);
    const ctx = decoCtx();
    const clinics = [...new Set(D.patients.map((p) => p.clinic).filter(Boolean))].sort();
    return { total: all.length, page, size, items: all.slice((page - 1) * size, page * size).map((p) => decorate(p, ctx)), clinics };
  });
  route('GET', '/api/patients/ids', ({ q: a }) => { auth({ roles: ['admin'] }); return filterPatients(a).map((p) => p.id); });
  route('GET', '/api/patients/:id', ({ p: prm }) => {
    auth({ roles: ['admin', 'employee', 'px'] });
    const p = decorate(getPatientChecked(prm.id)); const pid = p.id;
    p.reviews = D.reviews.filter((r) => r.patient_id === pid).sort((a, b) => b.id - a.id);
    p.cases = D.px_cases.filter((c) => c.patient_id === pid).sort((a, b) => b.id - a.id).map((c) => ({ ...c, reason_label: PX_REASON[c.reason], status_label: PX_STATUS[c.status] }));
    p.requests = D.requests.filter((r) => r.patient_id === pid).sort((a, b) => b.id - a.id).map((r) => ({ id: r.id, ref: r.ref, type: r.type, status: r.status, created_at: r.created_at, type_label: REQ_TYPES[r.type], status_label: REQ_STATUS[r.status] }));
    p.history = D.field_history.filter((h) => h.patient_id === pid).sort((a, b) => b.id - a.id).map((h) => ({ ...h, field_label: FIELD_AR[h.field] || h.field }));
    p.transfers = D.assignments_log.filter((x) => x.patient_id === pid).sort((a, b) => b.id - a.id).map((x) => ({ ...x, from_name: userName(x.from_user), to_name: userName(x.to_user), by_name: userName(x.by_user) }));
    p.log = D.audit.filter((x) => x.entity === 'patient' && x.entity_id === pid).sort((a, b) => b.id - a.id).slice(0, 100);
    p.exports = D.export_items.filter((x) => x.patient_id === pid).map((x) => byId.export_batches.get(x.batch_id)).filter(Boolean).sort((a, b) => b.id - a.id).map((b) => ({ id: b.id, at: b.at, by_name: b.by_name }));
    return p;
  });
  route('POST', '/api/patients', ({ body: d }) => {
    auth({ roles: ['admin', 'employee'] });
    const f = cleanFields(d, ['mrn', 'name', 'phone', 'phone2', 'clinic', 'city', 'district', 'pharmacy', 'address', 'out_of_city', 'periodic', 'notes', 'review_date']);
    if (!f.mrn || !f.name) throw new ApiError('رقم الملف والاسم مطلوبان');
    if (D.patients.some((p) => p.mrn === f.mrn)) throw new ApiError('يوجد مريض بنفس رقم الملف');
    if (!('out_of_city' in d)) { const fl = homeCityFlag(f.city); if (fl != null) f.out_of_city = fl; }
    const assigned = USER.role === 'employee' ? USER.id : (d.assigned_to ? +d.assigned_to : null);
    const p = insert('patients', { phone: '', phone2: '', clinic: '', city: '', district: '', pharmacy: '', address: '', notes: '', out_of_city: 0, periodic: null,
      ...f, review_date: f.review_date || todayS(), status: 'pending', last_note: null, last_decision_at: null, assigned_to: assigned, exported_batch_id: null, created_at: nowIso(), updated_at: nowIso() });
    if (assigned) insert('assignments_log', { patient_id: p.id, from_user: null, to_user: assigned, by_user: USER.id, reason: 'إضافة يدوية', at: nowIso() });
    audit('إضافة مريض يدويًا', 'patient', p.id, { mrn: f.mrn });
    return { ok: true, id: p.id };
  });
  route('PATCH', '/api/patients/:id', ({ p: prm, body: d }) => {
    auth({ roles: ['admin', 'employee', 'px'] });
    const p = getPatientChecked(prm.id);
    const role = USER.role;
    const allowed = role === 'admin' ? ['mrn', 'name', 'phone', 'phone2', 'clinic', 'city', 'district', 'pharmacy', 'address', 'out_of_city', 'periodic', 'notes', 'review_date']
      : role === 'employee' ? ['name', 'phone', 'phone2', 'clinic', 'city', 'district', 'pharmacy', 'address', 'out_of_city', 'periodic', 'notes'] : ['phone', 'phone2', 'city', 'district', 'address'];
    const f = cleanFields(d, allowed);
    if ('mrn' in f && D.patients.some((x) => x.mrn === f.mrn && x.id !== p.id)) throw new ApiError('رقم الملف مستخدم لمريض آخر');
    if ('city' in f && !('out_of_city' in d) && role !== 'px') { const fl = homeCityFlag(f.city); if (fl != null) f.out_of_city = fl; }
    const real = applyChanges(p.id, f, role !== 'px' ? 'تعديل' : 'تصحيح - تجربة المريض');
    if (Object.keys(real).length) audit('تعديل بيانات المريض', 'patient', p.id, Object.fromEntries(Object.entries(real).map(([k, v]) => [FIELD_AR[k] || k, v])));
    return { ok: true, changed: Object.keys(real) };
  });

  // الإسناد
  route('POST', '/api/assign', ({ body: d }) => {
    auth({ roles: ['admin'] });
    const ids = (d.patient_ids || []).map(Number), reason = (d.reason || '').trim() || 'إسناد من المدير';
    if (!ids.length) throw new ApiError('لم يتم اختيار مرضى');
    if (d.to_user == null || d.to_user === '' || d.to_user === 'none') {
      let n = 0;
      for (const id of ids) { const p = byId.patients.get(id); if (p && p.assigned_to) { insert('assignments_log', { patient_id: id, from_user: p.assigned_to, to_user: null, by_user: USER.id, reason, at: nowIso() }); p.assigned_to = null; n++; } }
      audit('إلغاء إسناد', 'patient', null, { 'العدد': n });
      return { ok: true, moved: n };
    }
    const tu = byId.users.get(+d.to_user);
    if (!tu || tu.role !== 'employee' || !tu.active) throw new ApiError('الموظف غير صحيح أو معطّل');
    const n = ids.filter((id) => assignPatient(id, tu.id, reason)).length;
    audit('إسناد مرضى', 'patient', null, { 'إلى': tu.name, 'العدد': n, 'السبب': reason });
    return { ok: true, moved: n };
  });
  route('POST', '/api/assign/auto', ({ body: d }) => {
    auth({ roles: ['admin'] });
    const mode = d.mode === 'load' ? 'load' : 'equal';
    const users = (d.user_ids || []).map(Number).length ? d.user_ids.map(Number) : D.users.filter((u) => u.role === 'employee' && u.active).map((u) => u.id);
    const ids = d.scope === 'selected' ? (d.patient_ids || []).map(Number) : sortByDate(D.patients.filter((p) => !p.assigned_to)).map((p) => p.id);
    if (!ids.length) throw new ApiError('لا يوجد مرضى للتوزيع');
    const res = distribute(ids, users, mode, 'توزيع تلقائي (' + (mode === 'load' ? 'حسب الحالات المفتوحة' : 'بالتساوي') + ')');
    audit('توزيع تلقائي', 'patient', null, { 'الطريقة': mode, 'النتيجة': res });
    return { ok: true, result: res };
  });

  // مراجعة الصيدلي
  route('GET', '/api/review/queue', ({ q: a }) => {
    auth({ roles: ['admin', 'employee'], review: true });
    let lim = addDays(todayS(), +S('lead_days'));
    if (a.until && validDate(a.until) && a.until > lim) lim = a.until;
    const ownOnly = USER.role === 'employee' && S('pharmacist_scope') !== 'all';
    const ctx = decoCtx();
    const rows = D.patients.filter((p) => (ownOnly ? p.assigned_to === USER.id : visible(p)) && inQueue(p, lim)
      && (!a.assigned || !/^\d+$/.test(a.assigned) || p.assigned_to === +a.assigned) && (!a.clinic || p.clinic === a.clinic)
      && (!a.q || like(p.mrn, XL.normDigits(a.q)) || like(p.name, a.q)));
    return sortByDate(rows).slice(0, 3000).map((p) => decorate(p, ctx));
  });
  route('POST', '/api/patients/:id/review', ({ p: prm, body: d }) => {
    auth({ roles: ['admin', 'employee'], review: true });
    const p = byId.patients.get(+prm.id);
    if (!p || (USER.role === 'employee' && S('pharmacist_scope') !== 'all' && p.assigned_to !== USER.id)) throw new ApiError('المريض غير موجود أو غير مسند إليك', 404);
    const dec = d.decision, note = (d.note || '').trim(), pnote = (d.patient_note || '').trim(); let nd = (d.next_date || '').trim() || null; const wf = (d.wrong_field || '').trim(), corr = (d.correction || '').trim();
    if (!['eligible', 'not_eligible', 'no_rx', 'wrong_data'].includes(dec)) throw new ApiError('اختر قرار المراجعة');
    if (nd && !validDate(nd)) throw new ApiError('تاريخ الاستحقاق القادم غير صحيح');
    if (dec === 'not_eligible') { if (!nd) throw new ApiError('تاريخ الاستحقاق القادم إلزامي لغير المستحق'); if (nd <= todayS()) throw new ApiError('تاريخ الاستحقاق القادم يجب أن يكون بعد اليوم'); }
    if (dec === 'no_rx' && !note && !pnote) throw new ApiError('اكتب ملاحظة توضح الإجراء المطلوب من المريض');
    if (dec === 'eligible' && !nd) nd = addDays(todayS(), +S('refill_default') || 30);
    if (dec === 'eligible' && nd <= todayS()) throw new ApiError('تاريخ الصرف القادم يجب أن يكون بعد اليوم');
    if (dec === 'wrong_data' && (!wf || !corr)) throw new ApiError('حدد الحقل الخاطئ ووصف التصحيح المطلوب');
    insert('reviews', { patient_id: p.id, decision: dec, note, patient_note: pnote, next_date: nd, wrong_field: wf, correction: corr, by_user: USER.id, by_name: USER.name, at: nowIso() });
    let caseId = null;
    if (dec === 'eligible') Object.assign(p, { status: 'eligible', review_date: nd, last_note: note, last_decision_at: nowIso(), exported_batch_id: null, updated_at: nowIso() });
    else {
      Object.assign(p, { status: dec, review_date: dec === 'not_eligible' ? nd : null, last_note: note, last_decision_at: nowIso(), updated_at: nowIso() });
      caseId = openPxCase(p.id, dec, note, pnote, wf, corr, nd, p.assigned_to);
    }
    audit('قرار الصيدلي: ' + STATUS_LABELS[dec], 'patient', p.id, { 'التاريخ القادم': nd, 'ملاحظة': note, 'إحالة لتجربة المريض': !!caseId });
    return { ok: true, px_case_id: caseId };
  });
  route('POST', '/api/patients/:id/refer', ({ p: prm, body: d }) => {
    auth({ roles: ['admin', 'employee'] });
    const p = getPatientChecked(prm.id);
    const wf = (d.wrong_field || '').trim(), corr = (d.correction || '').trim();
    if (!wf || !corr) throw new ApiError('حدد الحقل الخاطئ ووصف التصحيح المطلوب');
    const prev = p.status; p.status = 'wrong_data'; p.updated_at = nowIso();
    const cid = openPxCase(p.id, 'wrong_data', (d.note || '').trim(), (d.patient_note || '').trim(), wf, corr, null, p.assigned_to);
    audit('إحالة لتجربة المريض: بيانات خاطئة', 'patient', p.id, { 'الحقل': wf, 'الحالة السابقة': prev });
    return { ok: true, px_case_id: cid };
  });

  // تجربة المريض
  function getCase(id) {
    const c = byId.px_cases.get(+id); if (!c) throw new ApiError('الحالة غير موجودة', 404);
    if (USER.role === 'employee' && byId.patients.get(c.patient_id).assigned_to !== USER.id) throw new ApiError('الحالة غير موجودة', 404);
    return c;
  }
  route('GET', '/api/px/cases', ({ q: a }) => {
    auth({ roles: ['admin', 'px', 'employee'] });
    const st = a.status || 'open';
    const rows = D.px_cases.filter((c) => {
      const p = byId.patients.get(c.patient_id); if (!p) return false;
      if (USER.role === 'employee' && p.assigned_to !== USER.id) return false;
      if (st === 'open' && c.closed_at) return false;
      if (PX_STATUS[st] && c.status !== st) return false;
      if (PX_REASON[a.reason] && c.reason !== a.reason) return false;
      if (a.q && !(like(p.mrn, XL.normDigits(a.q)) || like(p.name, a.q) || like(p.phone, XL.normDigits(a.q)))) return false;
      return true;
    }).sort((x, y) => (!!x.closed_at) - (!!y.closed_at) || y.id - x.id).slice(0, 1000);
    return rows.map((c) => {
      const p = byId.patients.get(c.patient_id);
      const contacts = D.px_contacts.filter((x) => x.case_id === c.id).map((x) => x.at).sort();
      return { ...c, mrn: p.mrn, name: p.name, phone: p.phone, phone2: p.phone2, clinic: p.clinic, review_date: p.review_date, owner_name: userName(p.assigned_to),
        created_by_name: userName(c.created_by), last_contact: contacts[contacts.length - 1] || null, reason_label: PX_REASON[c.reason], status_label: PX_STATUS[c.status] };
    });
  });
  route('GET', '/api/px/cases/:id', ({ p: prm }) => {
    auth({ roles: ['admin', 'px', 'employee'] });
    const c = getCase(prm.id), p = byId.patients.get(c.patient_id);
    const msg = caseMessage(c, p);
    return { case: { ...c, reason_label: PX_REASON[c.reason], status_label: PX_STATUS[c.status], created_by_name: userName(c.created_by), closed_by_name: userName(c.closed_by) },
      patient: { ...p, assigned_name: userName(p.assigned_to) }, contacts: D.px_contacts.filter((x) => x.case_id === c.id).sort((a, b) => b.id - a.id),
      message: c.last_message || msg, default_message: msg,
      history: D.field_history.filter((h) => h.patient_id === p.id).sort((a, b) => b.id - a.id).slice(0, 30).map((h) => ({ ...h, field_label: FIELD_AR[h.field] || h.field })) };
  });
  route('POST', '/api/px/cases/:id/contact', ({ p: prm, body: d }) => {
    auth({ roles: ['admin', 'px'] });
    const c = getCase(prm.id); if (c.closed_at) throw new ApiError('الحالة مغلقة');
    const labels = { sent: 'تم إرسال الرسالة', replied: 'رد المريض', no_answer: 'لم يرد', wrong_number: 'رقم غير صحيح', called: 'اتصال هاتفي', other: 'أخرى' };
    const result = d.result || 'sent'; if (!labels[result]) throw new ApiError('نتيجة التواصل غير صحيحة');
    const msg = (d.message || '').trim();
    insert('px_contacts', { case_id: c.id, channel: d.channel || 'whatsapp', message: msg, result: labels[result], note: (d.note || '').trim(), by_user: USER.id, by_name: USER.name, at: nowIso() });
    const ns = PX_STATUS[d.status] && d.status !== 'done' ? d.status : (result === 'sent' ? 'awaiting' : 'contacting');
    Object.assign(c, { status: ns, last_message: msg || c.last_message, updated_at: nowIso() });
    audit('تواصل مع المريض: ' + labels[result], 'patient', c.patient_id, { 'الحالة': PX_STATUS[ns] });
    return { ok: true };
  });
  route('POST', '/api/px/cases/:id/status', ({ p: prm, body: d }) => {
    auth({ roles: ['admin', 'px'] });
    const c = getCase(prm.id);
    if (!PX_STATUS[d.status] || d.status === 'done') throw new ApiError('الحالة غير صحيحة');
    const closed = d.status === 'unreachable';
    Object.assign(c, { status: d.status, updated_at: nowIso(), closed_at: closed ? nowIso() : null, closed_by: closed ? USER.id : null, close_note: (d.note || '').trim() || null });
    if (closed) {
      const p = byId.patients.get(c.patient_id);
      if (p && ['no_rx', 'wrong_data'].includes(p.status)) {
        Object.assign(p, { status: 'pending', review_date: addDays(todayS(), 7), last_note: 'تعذر التواصل — أُعيد للمراجعة بعد 7 أيام', updated_at: nowIso() });
      }
    }
    audit('تحديث حالة تجربة المريض: ' + PX_STATUS[d.status], 'patient', c.patient_id, { 'ملاحظة': d.note });
    return { ok: true };
  });
  route('POST', '/api/px/cases/:id/close', ({ p: prm, body: d }) => {
    auth({ roles: ['admin', 'px'] });
    const c = getCase(prm.id); if (c.closed_at) throw new ApiError('الحالة مغلقة مسبقًا');
    const note = (d.note || '').trim();
    let back = !!d.return_to_review || ['no_rx', 'wrong_data'].includes(c.reason);
    if (back && c.reason === 'not_eligible' && !d.confirm_early) throw new ApiError('المريض غير مستحق حتى تاريخه القادم، وسيعود تلقائيًا للصيدلي قبل موعده. تأكيد الإعادة المبكرة مطلوب', 409, { need_confirm: true });
    Object.assign(c, { status: 'done', closed_at: nowIso(), closed_by: USER.id, close_note: note, updated_at: nowIso() });
    if (back) Object.assign(byId.patients.get(c.patient_id), { status: 'pending', review_date: todayS(), updated_at: nowIso() });
    audit('تمت المعالجة - تجربة المريض' + (back ? ' وإعادة للمراجعة' : ''), 'patient', c.patient_id, { 'ملاحظة': note });
    return { ok: true, returned: back };
  });

  // طلبات المرضى
  const reqVisible = (r) => isAdmin() || (USER.role === 'px' && r.type === 'complaint') || (USER.role === 'employee' && r.type !== 'complaint' && r.assigned_to === USER.id);
  function getReq(id) { const r = byId.requests.get(+id); if (!r || !reqVisible(r)) throw new ApiError('الطلب غير موجود', 404); return r; }
  route('GET', '/api/requests', ({ q: a }) => {
    auth({ roles: ['admin', 'employee', 'px'] });
    const st = a.status || 'open';
    return D.requests.filter((r) => {
      if (!reqVisible(r)) return false;
      if (REQ_TYPES[a.type] && r.type !== a.type) return false;
      if (st === 'open' && !isOpenReq(r)) return false;
      if (REQ_STATUS[st] && r.status !== st) return false;
      if (a.assigned === 'none' && r.assigned_to) return false;
      if (/^\d+$/.test(a.assigned || '') && r.assigned_to !== +a.assigned) return false;
      if (a.q) { const p = byId.patients.get(r.patient_id); const s = XL.normDigits(a.q); if (!(like(r.ref, s) || like(r.mrn, s) || like(r.full_name, a.q) || (p && like(p.name, a.q)))) return false; }
      return true;
    }).sort((x, y) => y.id - x.id).slice(0, 1000).map((r) => {
      const p = byId.patients.get(r.patient_id);
      return { id: r.id, ref: r.ref, type: r.type, status: r.status, delivery_status: r.delivery_status, mrn: r.mrn, created_at: r.created_at, possible_duplicate: r.possible_duplicate,
        patient_id: r.patient_id, has_attachment: !!r.attachment, details: (r.details || '').slice(0, 120), patient_name: p ? p.name : r.full_name, assigned_name: userName(r.assigned_to),
        type_label: REQ_TYPES[r.type], status_label: REQ_STATUS[r.status], delivery_label: DELIVERY_STATUS[r.delivery_status], subtype: r.subtype, phone: r.phone, full_name: r.full_name };
    });
  });
  route('GET', '/api/requests/:id', async ({ p: prm }) => {
    auth({ roles: ['admin', 'employee', 'px'] });
    const r = { ...getReq(prm.id) };
    const p = r.patient_id ? byId.patients.get(r.patient_id) : null;
    Object.assign(r, { type_label: REQ_TYPES[r.type], status_label: REQ_STATUS[r.status], delivery_label: DELIVERY_STATUS[r.delivery_status], assigned_name: userName(r.assigned_to) });
    r.patient = p ? { ...p, assigned_name: userName(p.assigned_to) } : null;
    const rel = r.related_ref ? D.requests.find((x) => x.ref === r.related_ref) : null;
    r.related = rel ? { id: rel.id, ref: rel.ref, type: rel.type, details: rel.details, status: rel.status, created_at: rel.created_at } : null;
    r.log = D.audit.filter((x) => x.entity === 'request' && x.entity_id === r.id).sort((a, b) => b.id - a.id);
    if (r.type === 'complaint' && r.phone) r.wa_reply = waLink(r.phone, `السلام عليكم ${r.full_name || ''}،\nبخصوص ${r.subtype || 'طلبكم'} رقم ${r.ref}، `);
    if (r.status === 'needs_info' && (p || r.phone)) r.wa_link = waLink(p ? p.phone : r.phone, fillTemplate(S('wa_templates').needs_info, { 'الاسم': p ? p.name : r.full_name, 'رقم_الطلب': r.ref, 'المطلوب': r.needs_info_note || '' }));
    if (r.attachment) {
      if (ENV.server) r.attachment_url = `/api/requests/${r.id}/attachment`;
      else { const f = await KV.get('att:' + r.attachment); if (f) r.attachment_url = URL.createObjectURL(f); }
    }
    return r;
  });
  route('GET', '/api/requests/:id/attachment', async ({ p: prm }) => {
    auth({ roles: ['admin', 'employee', 'px'] });
    const r = getReq(prm.id);
    if (!r.attachment) throw new ApiError('لا يوجد مرفق', 404);
    const f = await KV.get('att:' + r.attachment);
    if (!f) throw new ApiError('المرفق غير موجود', 404);
    const nm = r.attachment_name || 'attachment';
    const ext = (nm.split('.').pop() || '').toLowerCase();
    const mime = { pdf: 'application/pdf', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', heic: 'image/heic', heif: 'image/heif' }[ext] || 'application/octet-stream';
    return fileBlob(f, nm, mime, true);
  });
  route('PATCH', '/api/requests/:id', ({ p: prm, body: d }) => {
    auth({ roles: ['admin', 'employee', 'px'] });
    const r = getReq(prm.id), ch = {};
    if ('status' in d) { if (!REQ_STATUS[d.status]) throw new ApiError('حالة غير صحيحة'); if (d.status === 'needs_info' && !(d.needs_info_note || r.needs_info_note)) throw new ApiError('حدد المطلوب من المريض لاستكمال البيانات'); ch.status = d.status; }
    if ('delivery_status' in d) { if (!DELIVERY_STATUS[d.delivery_status]) throw new ApiError('حالة توصيل غير صحيحة'); ch.delivery_status = d.delivery_status; }
    for (const k of ['needs_info_note', 'staff_note', 'shortage_requested', 'shortage_dispensed', 'shortage_received', 'shortage_action']) if (k in d) ch[k] = (d[k] || '').trim();
    if (r.type === 'shortage' && ch.status === 'done' && !(ch.shortage_action || r.shortage_action)) throw new ApiError('سجّل الإجراء المتخذ لطلب العلاج الناقص قبل الإغلاق');
    if (!Object.keys(ch).length) return { ok: true };
    Object.assign(r, ch, { updated_at: nowIso() });
    const lab = {};
    if (ch.status) lab['الحالة'] = REQ_STATUS[ch.status];
    if (ch.delivery_status) lab['التوصيل'] = DELIVERY_STATUS[ch.delivery_status];
    if (ch.needs_info_note) lab['المطلوب من المريض'] = ch.needs_info_note;
    if (ch.shortage_action) lab['الإجراء'] = ch.shortage_action;
    if (ch.staff_note) lab['ملاحظة'] = ch.staff_note;
    audit('تحديث طلب ' + r.ref, 'request', r.id, lab);
    return { ok: true };
  });
  route('POST', '/api/requests/:id/assign', ({ p: prm, body: d }) => { auth({ roles: ['admin'] }); const r = getReq(prm.id); r.assigned_to = d.to_user ? +d.to_user : null; audit('إسناد طلب', 'request', r.id, { 'إلى': userName(r.assigned_to) }); return { ok: true }; });
  route('POST', '/api/requests/:id/link', ({ p: prm, body: d }) => {
    auth({ roles: ['admin'] }); const r = getReq(prm.id);
    const p = D.patients.find((x) => x.mrn === XL.normMrn(d.mrn)); if (!p) throw new ApiError('لا يوجد مريض بهذا الرقم');
    Object.assign(r, { patient_id: p.id, mrn: p.mrn, assigned_to: p.assigned_to || r.assigned_to });
    audit('ربط الطلب بملف مريض', 'request', r.id, { mrn: p.mrn });
    return { ok: true };
  });
  route('POST', '/api/requests/:id/create-patient', ({ p: prm }) => {
    auth({ roles: ['admin'] }); const r = getReq(prm.id);
    if (r.patient_id) throw new ApiError('الطلب مرتبط بملف مسبقًا');
    if (D.patients.some((x) => x.mrn === r.mrn)) throw new ApiError('يوجد ملف بنفس الرقم، استخدم الربط');
    if (!r.full_name) throw new ApiError('الطلب لا يحتوي اسمًا لإنشاء الملف');
    const p = insert('patients', { mrn: r.mrn, name: r.full_name, phone: r.phone, phone2: r.phone2, clinic: r.clinic, city: '', district: '', pharmacy: '', address: r.short_address, notes: '',
      out_of_city: 0, periodic: null, status: 'pending', review_date: todayS(), last_note: null, last_decision_at: null, assigned_to: null, exported_batch_id: null, created_at: nowIso(), updated_at: nowIso() });
    r.patient_id = p.id;
    audit('إنشاء ملف مريض من طلب', 'patient', p.id, { 'الطلب': r.ref });
    return { ok: true, patient_id: p.id };
  });
  route('POST', '/api/requests/:id/apply', ({ p: prm }) => {
    auth({ roles: ['admin', 'employee'] }); const r = getReq(prm.id);
    if (!r.patient_id) throw new ApiError('الطلب غير مرتبط بملف مريض');
    const ch = {};
    for (const [s, t] of [['full_name', 'name'], ['phone', 'phone'], ['phone2', 'phone2'], ['clinic', 'clinic'], ['short_address', 'address']]) if (r[s]) ch[t] = r[s];
    const real = applyChanges(r.patient_id, ch, 'من طلب التوصيل ' + r.ref);
    audit('اعتماد بيانات التوصيل في الملف', 'patient', r.patient_id, Object.keys(real).length ? Object.fromEntries(Object.entries(real).map(([k, v]) => [FIELD_AR[k] || k, v])) : { 'نتيجة': 'لا تغييرات' });
    return { ok: true, changed: Object.keys(real) };
  });

  // بوابة المريض — الدخول برقم الملف + رقم الجوال، ولا يشترط وجود المريض في النظام
  const portalSess = () => { try { const x = JSON.parse(SESS.get('pt') || 'null'); if (x && x.mrn && x.phone) return x; } catch (e) {} throw new ApiError('انتهت الجلسة، أدخل رقم الملف ورقم الجوال مرة أخرى', 401); };
  route('POST', '/api/portal/login', ({ body: d }) => {
    const mrn = XL.normMrn(d.mrn);
    if (!/^[A-Za-z0-9\-]{3,20}$/.test(mrn || '')) throw new ApiError('أدخل رقم ملف طبي صحيح');
    const [ph, e] = XL.normPhone(d.phone);
    if (!ph || e) throw new ApiError('أدخل رقم جوال صحيح (مثال: 05XXXXXXXX)');
    SESS.set('pt', JSON.stringify({ mrn, phone: ph }));
    return { ok: true, mrn };
  });
  route('POST', '/api/portal/logout', () => { SESS.del('pt'); return { ok: true }; });
  route('GET', '/api/portal/me', () => {
    const { mrn, phone } = portalSess();
    const p = D.patients.find((x) => x.mrn === mrn);
    const known = !!(p && (p.phone === phone || p.phone2 === phone));
    return { mrn, phone, org_name: S('org_name'),
      prefill: known ? { full_name: p.name || '', phone2: (p.phone === phone ? p.phone2 : p.phone) || '', short_address: p.address || '', clinic: p.clinic || '' } : { full_name: '', phone2: '', short_address: '', clinic: '' },
      requests: D.requests.filter((r) => r.mrn === mrn && r.login_phone === phone).sort((a, b) => b.id - a.id).slice(0, 50).map((r) => ({
        ref: r.ref, type: r.type, status: r.status, delivery_status: r.delivery_status, created_at: r.created_at, type_label: REQ_TYPES[r.type], status_label: REQ_STATUS[r.status],
        delivery_label: DELIVERY_STATUS[r.delivery_status], subtype: r.subtype, needs_info_note: r.status === 'needs_info' ? r.needs_info_note : null })) };
  });
  route('POST', '/api/portal/requests', async ({ body: f }) => {
    const sess = portalSess(), mrn = sess.mrn;
    const g = (k) => String((f.get ? f.get(k) : f[k]) ?? '').trim();
    const t = g('type');
    if (!PORTAL_TYPES.includes(t)) throw new ApiError('نوع الطلب غير صحيح');
    const name = g('full_name').replace(/\s+/g, ' ');
    const [ph, e] = XL.normPhone(g('phone') || sess.phone); if (!ph || e) throw new ApiError('رقم الجوال غير صحيح (مثال: 05XXXXXXXX)');
    const [ph2, e2] = XL.normPhone(g('phone2')); if (ph2 && e2) throw new ApiError('رقم الجوال الإضافي غير صحيح');
    const sa = XL.normDigits(g('short_address')).replace(/[^A-Za-z0-9]/g, '').toUpperCase();
    const details = g('details').slice(0, 3000), med = g('med_name').slice(0, 500);
    const up = f.get ? f.get('attachment') : null;
    const hasFile = !!(up && up.name && up.size);
    if (!name) throw new ApiError('اكتب اسم المريض');
    if (t !== 'complaint') {
      if (name.split(' ').length < (t === 'delivery' ? 4 : 2)) throw new ApiError(t === 'delivery' ? 'اكتب اسم المريض الرباعي' : 'اكتب اسم المريض كاملًا');
      if (!/^[A-Z]{4}\d{4}$/.test(sa)) throw new ApiError('العنوان الوطني المختصر غير صحيح (4 حروف إنجليزية ثم 4 أرقام، مثال: RRRD2929)');
    }
    if (t === 'shortage' && !med && !hasFile) throw new ApiError('اكتب اسم العلاج الناقص أو أرفق صورة له');
    if (t === 'complaint' && !details) throw new ApiError('اكتب ملاحظاتك');
    let att = null, attName = null;
    if (hasFile) {
      if (!/\.(jpe?g|png|gif|webp|heic|heif|pdf)$/i.test(up.name) && !/^image\//.test(up.type || '')) throw new ApiError('المرفق يجب أن يكون صورة أو PDF');
      if (up.size > (ENV.server ? 4.5 : 10) * 1024 * 1024) throw new ApiError('حجم المرفق كبير، أرسل صورة أصغر');
      att = Math.random().toString(36).slice(2) + Date.now().toString(36);
      await KV.set('att:' + att, ENV.server ? await up.arrayBuffer() : up);
      attName = up.name.slice(-120);
    }
    const p = D.patients.find((x) => x.mrn === mrn);
    const since = new Date(Date.now() - 3 * 864e5).toISOString();
    const dup = D.requests.some((r) => r.mrn === mrn && r.type === t && r.created_at >= since && isOpenReq(r)) ? 1 : 0;
    const r = insert('requests', { ref: null, patient_id: p ? p.id : null, mrn, login_phone: sess.phone, type: t, subtype: t === 'complaint' ? (g('subtype') === 'شكوى' ? 'شكوى' : 'تواصل') : null,
      details: t === 'shortage' && med ? ('العلاج: ' + med + (details ? '\n' + details : '')) : details, med_name: med || null, attachment: att, attachment_name: attName,
      full_name: name, phone: ph, phone2: ph2 || null, clinic: g('clinic') || null, short_address: sa || null, related_ref: null, status: 'new',
      delivery_status: t === 'delivery' ? 'new' : null, needs_info_note: null, staff_note: null, shortage_requested: med || null, shortage_dispensed: null, shortage_received: null,
      shortage_action: null, possible_duplicate: dup, assigned_to: t === 'complaint' ? null : (p ? p.assigned_to : null), created_at: nowIso(), updated_at: nowIso() });
    const k = new Date(Date.now() + 3 * 3600e3);
    r.ref = 'R' + String(k.getUTCFullYear()).slice(2) + String(k.getUTCMonth() + 1).padStart(2, '0') + '-' + String(r.id).padStart(5, '0');
    audit('طلب جديد من بوابة المريض: ' + REQ_TYPES[t], 'request', r.id, { 'المرجع': r.ref, mrn, 'مريض جديد': !p }, null);
    return { ok: true, ref: r.ref };
  });

  // بحث تجربة المريض: حالة المريض وطلباته ولو لم يكن مسجلًا
  route('GET', '/api/px/search', ({ q: a }) => {
    auth({ roles: ['admin', 'px', 'employee'] });
    const q = (a.q || '').trim(); if (q.length < 2) return [];
    const qd = XL.normDigits(q);
    const lim = addDays(todayS(), +S('lead_days')), t = todayS();
    const ctx = decoCtx();
    const scope = (p) => USER.role !== 'employee' || p.assigned_to === USER.id;
    const pats = D.patients.filter((p) => scope(p) && (like(p.mrn, qd) || like(p.name, q) || like(p.phone, qd) || like(p.phone2, qd))).slice(0, 30);
    const seen = new Set(pats.map((p) => p.mrn));
    const out = pats.map((p) => {
      const d = decorate(p, ctx);
      let stage;
      if (d.px_case_id) stage = { tone: 'amber', text: 'لدى تجربة المريض — ' + (d.px_reason_label || '') + ' (' + (d.px_status_label || '') + ')' };
      else if (p.status === 'eligible') stage = { tone: 'green', text: 'مستحق — بانتظار التصدير للبريد' };
      else if (inQueue(p, lim)) stage = p.review_date < t ? { tone: 'red', text: `بانتظار مراجعة الصيدلي — متأخر ${daysBetween(p.review_date, t)} يوم` } : { tone: 'blue', text: 'بانتظار مراجعة الصيدلي — موعده ' + p.review_date };
      else if (p.status === 'exported') stage = { tone: 'green', text: 'تم التصدير للبريد — الصرف القادم ' + (p.review_date || '—') };
      else if (p.status === 'not_eligible') stage = { tone: 'gray', text: 'غير مستحق حاليًا — الاستحقاق القادم ' + (p.review_date || '—') };
      else if (p.review_date) stage = { tone: 'gray', text: 'المراجعة القادمة ' + p.review_date };
      else stage = { tone: 'red', text: 'بلا موعد متابعة — يحتاج إجراء' };
      const lastRev = D.reviews.filter((r) => r.patient_id === p.id).sort((x, y) => y.id - x.id)[0];
      return { kind: 'patient', id: p.id, mrn: p.mrn, name: p.name, phone: p.phone, clinic: p.clinic, status_label: d.status_label, review_date: p.review_date,
        assigned_name: d.assigned_name, current_handler: d.current_handler, px_case_id: d.px_case_id, stage,
        last_review: lastRev ? { decision: STATUS_LABELS[lastRev.decision] || lastRev.decision, by: lastRev.by_name, at: lastRev.at } : null,
        requests: D.requests.filter((r) => r.mrn === p.mrn).sort((x, y) => y.id - x.id).slice(0, 10).map((r) => ({ id: r.id, ref: r.ref, type_label: REQ_TYPES[r.type], status: r.status, status_label: REQ_STATUS[r.status], delivery_label: DELIVERY_STATUS[r.delivery_status], created_at: r.created_at })) };
    });
    if (USER.role !== 'employee') {
      const extra = new Map();
      for (const r of D.requests) {
        if (seen.has(r.mrn)) continue;
        if (like(r.mrn, qd) || like(r.full_name, q) || like(r.phone, qd) || like(r.ref, qd)) { if (!extra.has(r.mrn)) extra.set(r.mrn, []); extra.get(r.mrn).push(r); }
      }
      for (const [mrn, rs] of [...extra].slice(0, 20)) {
        rs.sort((x, y) => y.id - x.id);
        out.push({ kind: 'unregistered', mrn, name: rs[0].full_name, phone: rs[0].phone, stage: { tone: 'amber', text: 'مريض غير مسجل في النظام — لديه طلبات من البوابة' },
          requests: rs.slice(0, 10).map((r) => ({ id: r.id, ref: r.ref, type_label: REQ_TYPES[r.type], status: r.status, status_label: REQ_STATUS[r.status], delivery_label: DELIVERY_STATUS[r.delivery_status], created_at: r.created_at })) });
      }
    }
    return out;
  });

  // ضمان عدم التأخير: كل مريض يجب أن يكون له مسار متابعة دائمًا
  function monitor(scopeUid) {
    const t = todayS(), lim = addDays(t, +S('lead_days')), now = Date.now();
    const ago = (iso, h) => iso && (now - new Date(iso).getTime()) > h * 3600e3;
    const ctx = decoCtx();
    const inScope = (p) => !scopeUid || p.assigned_to === scopeUid;
    const pats = D.patients.filter(inScope);
    const activeEmp = new Set(D.users.filter((u) => u.role === 'employee' && u.active).map((u) => u.id));
    const hasPath = (p) => p.status === 'eligible' || ctx.cases.has(p.id) || (p.review_date && ['pending', 'not_eligible', 'exported'].includes(p.status));
    const row = (p) => ({ id: p.id, mrn: p.mrn, name: p.name, review_date: p.review_date, status_label: STATUS_LABELS[p.status], assigned_name: userName(p.assigned_to), days: p.review_date ? daysBetween(p.review_date, t) : null });
    const cats = [];
    const add = (key, level, title, hint, list, action) => cats.push({ key, level, title, hint, count: list.length, items: list.slice(0, 200), action });
    const orphans = pats.filter((p) => !hasPath(p));
    add('orphans', 'critical', 'مرضى بلا مسار متابعة', 'لا موعد مراجعة، ولا إحالة مفتوحة، وليسوا في قائمة التصدير — قد يُنسون تمامًا', orphans.map(row), 'fix_orphans');
    const overdue = pats.filter((p) => inQueue(p, lim) && p.review_date < t).sort((a, b) => a.review_date.localeCompare(b.review_date));
    add('overdue7', 'critical', 'متأخرون أكثر من 7 أيام عن موعد المراجعة', 'يجب مراجعتهم اليوم', overdue.filter((p) => daysBetween(p.review_date, t) > 7).map(row));
    add('overdue', 'high', 'متأخرون 1–7 أيام', 'ضمن قائمة الصيدلي، تجاوزوا الموعد', overdue.filter((p) => daysBetween(p.review_date, t) <= 7).map(row));
    add('today', 'info', 'موعدهم اليوم', 'يجب أن تُسجَّل قراراتهم قبل نهاية الدوام', pats.filter((p) => inQueue(p, lim) && p.review_date === t).map(row));
    if (!scopeUid) {
      add('unassigned', 'high', 'مرضى غير مسندين لموظف', 'لن يظهروا لأي موظف حتى يُسندوا', pats.filter((p) => !p.assigned_to).map(row), 'distribute');
      add('inactive', 'critical', 'مسندون لحساب معطّل', 'أعد إسنادهم لموظف نشط', pats.filter((p) => p.assigned_to && !activeEmp.has(p.assigned_to)).map(row));
    }
    add('elig_stale', 'high', 'مستحقون لم يُصدَّروا منذ أكثر من يومين', 'اعتمدهم الصيدلي ولم يُرسلوا للبريد', pats.filter((p) => p.status === 'eligible' && ago(p.last_decision_at, 48)).map(row));
    const pxOld = D.px_cases.filter((c) => !c.closed_at && ago(c.created_at, 72) && inScope(byId.patients.get(c.patient_id) || {}));
    add('px_stale', 'high', 'حالات لدى تجربة المريض منذ أكثر من 3 أيام', 'لم تُغلق بعد — قد تؤخر الصرف', pxOld.map((c) => ({ ...row(byId.patients.get(c.patient_id)), case_id: c.id, reason: PX_REASON[c.reason], px_status: PX_STATUS[c.status] })));
    const reqs = D.requests.filter((r) => isOpenReq(r) && (!scopeUid || r.assigned_to === scopeUid));
    const rrow = (r) => ({ req_id: r.id, mrn: r.mrn, name: (byId.patients.get(r.patient_id) || {}).name || r.full_name, ref: r.ref, type: REQ_TYPES[r.type], status_label: REQ_STATUS[r.status], created_at: r.created_at, assigned_name: userName(r.assigned_to) });
    add('req_new', 'high', 'طلبات بوابة بلا إجراء منذ أكثر من 24 ساعة', 'لم يفتحها أحد أو لم تتغير حالتها', reqs.filter((r) => r.status === 'new' && ago(r.created_at, 24)).map(rrow));
    if (!scopeUid) add('req_orphan', 'high', 'طلبات غير مسندة أو لمريض غير مسجل', 'اربطها بملف أو أنشئ ملفًا ثم أسندها', reqs.filter((r) => r.type !== 'complaint' && (!r.patient_id || !r.assigned_to)).map(rrow));
    const level = cats.some((c) => c.count && c.level === 'critical') ? 'critical' : cats.some((c) => c.count && c.level === 'high') ? 'high' : 'ok';
    const issues = cats.filter((c) => c.count && c.level !== 'info').reduce((n, c) => n + c.count, 0);
    return { level, issues, categories: cats, checked_at: nowIso(), total_patients: pats.length,
      guarantee: { with_path: pats.length - orphans.length, total: pats.length } };
  }
  route('GET', '/api/monitor', () => { auth({ roles: ['admin', 'employee'] }); return monitor(isAdmin() ? null : USER.id); });
  route('POST', '/api/monitor/fix-orphans', () => {
    auth({ roles: ['admin'] });
    const ctx = decoCtx(); let n = 0;
    for (const p of D.patients) {
      const hasPath = p.status === 'eligible' || ctx.cases.has(p.id) || (p.review_date && ['pending', 'not_eligible', 'exported'].includes(p.status));
      if (!hasPath) { Object.assign(p, { status: 'pending', review_date: todayS(), updated_at: nowIso() }); n++; }
    }
    audit('إعادة المرضى بلا مسار إلى مراجعة الصيدلي', 'patient', null, { 'العدد': n });
    return { ok: true, fixed: n };
  });

  // لوحة التحكم
  route('GET', '/api/dashboard', () => {
    auth({ roles: ['admin', 'employee', 'px'] });
    const t = todayS(), lead = +S('lead_days'), lim = addDays(t, lead);
    const week = new Date(Date.now() - 7 * 864e5).toISOString();
    const dayStart = new Date(Date.UTC(+t.slice(0, 4), +t.slice(5, 7) - 1, +t.slice(8, 10)) - 3 * 3600e3).toISOString();
    const openCases = new Set(D.px_cases.filter((c) => !c.closed_at).map((c) => c.patient_id));
    function stats(pred) {
      const ps = D.patients.filter(pred), ids = new Set(ps.map((p) => p.id));
      return { patients: ps.length, due: ps.filter((p) => inQueue(p, lim)).length, overdue: ps.filter((p) => inQueue(p, lim) && p.review_date < t).length,
        eligible: ps.filter((p) => p.status === 'eligible').length, at_px: ps.filter((p) => openCases.has(p.id)).length,
        reviewed_today: D.reviews.filter((r) => r.at >= dayStart && ids.has(r.patient_id)).length, reviewed_week: D.reviews.filter((r) => r.at >= week && ids.has(r.patient_id)).length,
        nodate: ps.filter((p) => !p.review_date && ['pending', 'exported', 'not_eligible'].includes(p.status)).length };
    }
    const res = { today: t, lead_days: lead };
    let rq = null;
    if (USER.role === 'admin') {
      res.totals = { ...stats(() => true), unassigned: D.patients.filter((p) => !p.assigned_to).length };
      res.employees = D.users.filter((u) => u.role === 'employee').sort((a, b) => b.active - a.active || a.name.localeCompare(b.name)).map((u) => ({
        id: u.id, name: u.name, active: !!u.active, can_review: !!u.can_review, ...stats((p) => p.assigned_to === u.id),
        open_requests: D.requests.filter((r) => r.assigned_to === u.id && isOpenReq(r)).length,
        requests_done_week: D.requests.filter((r) => r.assigned_to === u.id && !isOpenReq(r) && r.updated_at >= week).length }));
      rq = () => true;
    } else if (USER.role === 'employee') { res.totals = stats((p) => p.assigned_to === USER.id); rq = (r) => r.assigned_to === USER.id; }
    if (rq) {
      const open = D.requests.filter((r) => rq(r) && isOpenReq(r));
      res.requests = Object.fromEntries(Object.keys(REQ_TYPES).map((k) => [k, open.filter((r) => r.type === k).length]));
      res.requests.new = D.requests.filter((r) => rq(r) && r.status === 'new').length;
      res.requests.needs_info = D.requests.filter((r) => rq(r) && r.status === 'needs_info').length;
      if (USER.role === 'admin') res.requests.unlinked = open.filter((r) => !r.patient_id || !r.assigned_to).length;
    }
    const pxs = D.px_cases.filter((c) => USER.role !== 'employee' || (byId.patients.get(c.patient_id) || {}).assigned_to === USER.id);
    res.px = Object.fromEntries(['new', 'contacting', 'awaiting'].map((k) => [k, pxs.filter((c) => c.status === k && !c.closed_at).length]));
    res.px.by_reason = Object.fromEntries(Object.keys(PX_REASON).map((k) => [k, pxs.filter((c) => c.reason === k && !c.closed_at).length]));
    res.px.done_week = pxs.filter((c) => c.closed_at && c.closed_at >= week).length;
    if (USER.role !== 'px') { const m = monitor(isAdmin() ? null : USER.id); res.safety = { level: m.level, issues: m.issues, top: m.categories.filter((c) => c.count && c.level !== 'info').slice(0, 4).map((c) => ({ key: c.key, title: c.title, count: c.count, level: c.level })) }; }
    if (USER.role === 'px') res.complaints = D.requests.filter((r) => r.type === 'complaint' && isOpenReq(r)).length;
    return res;
  });

  // الاستيراد
  const IMPORTS = {};
  route('POST', '/api/import/parse', async ({ body: f }) => {
    auth({ roles: ['admin'] });
    const file = f.get('file');
    if (!file || !/\.(xlsx|xlsm)$/i.test(file.name)) throw new ApiError('ارفع ملف Excel بصيغة xlsx (إن كان xls قديمًا احفظه بصيغة xlsx أولًا)');
    const buf = await file.arrayBuffer();
    let parsed;
    try { parsed = await XL.readWorkbook(buf, f.get('sheet') || null, f.get('use_font') !== '0'); } catch (e) { throw new ApiError('تعذر قراءة الملف: ' + String(e.message || e).slice(0, 200)); }
    const tok = Math.random().toString(36).slice(2);
    IMPORTS[tok] = { data: parsed, raw: buf, name: file.name, useFont: f.get('use_font') !== '0' };
    if (ENV.server) await KV.set('imp:' + tok, JSON.stringify({ name: file.name, useFont: f.get('use_font') !== '0', sheet: parsed.sheet })), await KV.set('impraw:' + tok, buf);
    return { token: tok, file: file.name, sheets: parsed.sheets, sheet: parsed.sheet, header_row: parsed.header_row, headers: parsed.headers, mapping: parsed.mapping,
      fields: XL.FIELD_LABELS, color_actions: S('color_actions'), color_labels: XL.COLOR_LABELS, analysis: analyze(parsed, parsed.mapping, S('color_actions'), {}) };
  });
  async function imp(tok) {
    if (IMPORTS[tok]) return IMPORTS[tok];
    if (ENV.server && tok && /^[a-z0-9]+$/.test(tok)) {
      const meta = await KV.get('imp:' + tok), raw = await KV.get('impraw:' + tok);
      if (meta && raw) { const m = JSON.parse(typeof meta === 'string' ? meta : new TextDecoder().decode(meta)); IMPORTS[tok] = { data: await XL.readWorkbook(raw, m.sheet, m.useFont), raw, name: m.name, useFont: m.useFont }; return IMPORTS[tok]; }
    }
    throw new ApiError('انتهت صلاحية ملف الاستيراد، ارفع الملف مرة أخرى');
  }
  route('POST', '/api/import/sheet', async ({ body: d }) => {
    auth({ roles: ['admin'] }); const it = await imp(d.token);
    const parsed = await XL.readWorkbook(it.raw, d.sheet, d.use_font !== false); it.data = parsed;
    if (ENV.server) await KV.set('imp:' + d.token, JSON.stringify({ name: it.name, useFont: d.use_font !== false, sheet: parsed.sheet }));
    return { sheet: parsed.sheet, headers: parsed.headers, mapping: parsed.mapping, header_row: parsed.header_row, analysis: analyze(parsed, parsed.mapping, S('color_actions'), {}) };
  });
  function analyze(parsed, mapping, colorActions, opts, build = false) {
    mapping = Object.fromEntries(Object.entries(mapping || {}).filter(([, v]) => v !== null && v !== '' && v !== -1 && v !== '-1').map(([k, v]) => [k, +v]));
    const dateMode = opts.date_mode || 'due';
    const interval = Math.max(1, Math.min(365, +opts.interval || 30));
    const t = todayS(), seen = {}, out = [], errors = [], warnings = [], counts = {}, colors = {}, unknown = {};
    const existing = new Map(D.patients.map((p) => [p.mrn, p.id]));
    let existingN = 0, pastN = 0;
    const inc = (o, k) => o[k] = (o[k] || 0) + 1;
    for (const r of parsed.rows) {
      const [f, errs, ws] = XL.extractRow(r, mapping);
      if (f.review_date && f.review_date < t) pastN++;
      if (f.review_date && dateMode === 'last') f.review_date = addDays(f.review_date, interval);
      if (f.date_raw) f.notes = (f.notes ? f.notes + ' | ' : '') + 'التاريخ في الملف: ' + f.date_raw;
      inc(colors, r.color);
      const action = r.color !== 'none' ? (colorActions[r.color] || 'ignore') : 'none';
      if (r.color !== 'none' && !(r.color in colorActions)) unknown[r.color] = r.hex;
      if (action === 'skip') { inc(counts, 'skipped'); continue; }
      if (errs.length) { errors.push({ row: r.row, msg: errs.join('، ') }); continue; }
      if (seen[f.mrn]) { errors.push({ row: r.row, msg: `رقم الملف ${f.mrn} مكرر في الملف (الصف ${seen[f.mrn]})` }); continue; }
      seen[f.mrn] = r.row;
      let status = 'pending', oc = null;
      if (action === 'not_eligible') { status = 'not_eligible'; if (!f.review_date) { ws.push('أحمر بدون تاريخ: سيظهر للصيدلي اليوم لتحديد الموعد'); f.review_date = t; } }
      else if (action === 'eligible') status = 'eligible';
      else if (action === 'no_rx') status = 'no_rx';
      else if (action === 'out_of_city') oc = 1;
      if (oc == null) { const fl = homeCityFlag(f.city); if (fl != null) oc = fl; }
      if (status === 'pending' && !f.review_date) f.review_date = t;
      const ex = existing.has(f.mrn); if (ex) existingN++;
      inc(counts, status); if (oc) inc(counts, 'out_of_city'); if (f.periodic === 0) inc(counts, 'not_periodic');
      if (ws.length) warnings.push({ row: r.row, msg: ws.join('، ') });
      out.push({ row: r.row, color: r.color, hex: r.hex, action, status, status_label: STATUS_LABELS[status], out_of_city: oc, existing: ex, warn: ws.join('، '), ...f });
    }
    return { total: parsed.rows.length, valid: out.length, errors, warnings, existing: existingN, new: out.length - existingN, counts, colors, unknown_colors: unknown,
      past_dates: pastN, date_mode: dateMode, interval, preview: build ? out : out.slice(0, 200) };
  }
  route('POST', '/api/import/preview', async ({ body: d }) => { auth({ roles: ['admin'] }); return analyze((await imp(d.token)).data, d.mapping, d.color_actions || S('color_actions'), d); });
  route('POST', '/api/import/commit', async ({ body: d }) => {
    auth({ roles: ['admin'] });
    const it = await imp(d.token), ca = d.color_actions || S('color_actions');
    if (d.save_colors) setS('color_actions', ca);
    const res = analyze(it.data, d.mapping, ca, d, true);
    const upd = d.existing === 'update', refer = !!d.refer_not_eligible;
    let updated = 0, skipped = 0, referred = 0; const newIds = [];
    const byMrn = new Map(D.patients.map((p) => [p.mrn, p]));
    for (const r of res.preview) {
      const fields = Object.fromEntries(['name', 'phone', 'phone2', 'clinic', 'city', 'district', 'pharmacy', 'address', 'notes'].map((k) => [k, r[k]]));
      let pid;
      if (r.existing) {
        if (!upd) { skipped++; continue; }
        const p = byMrn.get(r.mrn);
        const ch = Object.fromEntries(Object.entries(fields).filter(([, v]) => v));
        if (r.out_of_city != null) ch.out_of_city = r.out_of_city;
        if (r.periodic != null) ch.periodic = r.periodic;
        applyChanges(p.id, ch, 'استيراد ' + it.name);
        if (r.color !== 'none' && ['not_eligible', 'eligible', 'no_rx'].includes(r.action)) Object.assign(p, { status: r.status, review_date: r.review_date, updated_at: nowIso() });
        else if (d.update_dates && r.review_date && ['pending', 'not_eligible', 'exported'].includes(p.status)) p.review_date = r.review_date;
        updated++; pid = p.id;
      } else {
        const p = insert('patients', { mrn: r.mrn, ...fields, out_of_city: r.out_of_city || 0, periodic: r.periodic, status: r.status,
          review_date: r.status === 'eligible' ? addDays(todayS(), +S('refill_default') || 30) : r.status === 'no_rx' ? null : r.review_date,
          last_note: null, last_decision_at: null, assigned_to: null, exported_batch_id: null, created_at: nowIso(), updated_at: nowIso() });
        pid = p.id; newIds.push(pid); byMrn.set(p.mrn, p);
      }
      if ((refer && r.status === 'not_eligible') || r.status === 'no_rx') {
        openPxCase(pid, r.status, 'مستورد من ملف Excel (لون ' + XL.COLOR_LABELS[r.color] + ')', '', '', '', r.status === 'not_eligible' ? r.review_date : null, byId.patients.get(pid).assigned_to);
        referred++;
      }
    }
    let dist = null;
    const am = d.assign_mode;
    if (newIds.length && (am === 'equal' || am === 'load')) {
      const users = (d.assign_users || []).length ? d.assign_users.map(Number) : D.users.filter((u) => u.role === 'employee' && u.active).map((u) => u.id);
      if (users.length) dist = distribute(newIds, users, am, 'استيراد ' + it.name);
    } else if (newIds.length && /^\d+$/.test(String(am || ''))) dist = distribute(newIds, [+am], 'equal', 'استيراد ' + it.name);
    audit('استيراد Excel', 'import', null, { 'الملف': it.name, 'جديد': newIds.length, 'محدّث': updated, 'متجاهل': skipped, 'أخطاء': res.errors.length, 'إحالات': referred, 'التوزيع': dist });
    delete IMPORTS[d.token];
    if (ENV.server) { await KV.del('imp:' + d.token); await KV.del('impraw:' + d.token); }
    return { ok: true, created: newIds.length, updated, skipped, errors: res.errors.length, referred, distribution: dist };
  });
  route('GET', '/api/import/sample', () => {
    auth({ roles: ['admin'] });
    const t = todayS();
    const data = XL.simpleXlsx(['رقم الملف', 'الاسم', 'الجوال', 'جوال إضافي', 'العيادة', 'المدينة', 'الحي', 'العنوان الوطني المختصر', 'تاريخ المراجعة', 'نعم / لا'],
      [['100001', 'مثال بدون لون', '0500000001', '', 'الباطنية', '', 'حي تجريبي', 'RRRD2929', t, 'نعم'], ['100002', 'مثال ثانٍ', '0500000002', '', 'القلب', '', 'حي تجريبي', 'ABCD1234', addDays(t, 20), 'لا']]);
    return fileBlob(data, 'نموذج_استيراد.xlsx');
  });

  // التصدير
  function candidates(a) {
    const ctx = decoCtx();
    const ids = a.patient_ids ? new Set(a.patient_ids.map(Number)) : null;
    const loc = a.location || 'all';
    return D.patients.filter((p) => p.status === 'eligible' && (USER.role !== 'employee' || p.assigned_to === USER.id)
      && (USER.role === 'employee' || !/^\d+$/.test(String(a.assigned || '')) || p.assigned_to === +a.assigned)
      && (loc !== 'in' || !p.out_of_city) && (loc !== 'out' || p.out_of_city) && (!a.clinic || p.clinic === a.clinic)
      && (!['0', '1'].includes(String(a.periodic ?? '')) || (p.periodic ?? 1) === +a.periodic) && (!ids || ids.has(p.id)))
      .sort((x, y) => (x.out_of_city || 0) - (y.out_of_city || 0) || String(x.clinic || '').localeCompare(String(y.clinic || '')) || String(x.name).localeCompare(String(y.name)))
      .map((p) => decorate(p, ctx));
  }
  route('GET', '/api/export/candidates', ({ q: a }) => { auth({ roles: ['admin', 'employee'] }); return candidates(a); });
  const exportCtx = (tpl) => ({ home_city: S('home_city') || '', ref_start: tpl.ref_start || 1 });
  const stamp = () => { const k = new Date(Date.now() + 3 * 3600e3).toISOString(); return k.slice(0, 10).replace(/-/g, '') + '_' + k.slice(11, 16).replace(':', ''); };
  route('POST', '/api/export', async ({ body: d }) => {
    auth({ roles: ['admin', 'employee'] });
    const pats = candidates(d);
    if (!pats.length) throw new ApiError('لا يوجد مرضى مستحقون ضمن الاختيار');
    const tpl = S('export_template');
    const tb = await tplBytes(tpl);
    const [data, ext] = await XL.buildExport(pats, tpl, tb, todayS(), !!d.split && (d.location || 'all') === 'all', exportCtx(tpl));
    const loc = { in: 'داخل_المدينة', out: 'خارج_المدينة' }[d.location] || 'الكل';
    const fname = `المستحقون_${loc}_${stamp()}.${ext}`;
    if (d.mark !== false) {
      const b = insert('export_batches', { created_by: USER.id, by_name: USER.name, at: nowIso(), count: pats.length, filter: JSON.stringify({ location: d.location, clinic: d.clinic }), file: fname });
      await KV.set('export:' + b.id, data);
      for (const p of pats) {
        insert('export_items', { batch_id: b.id, patient_id: p.id });
        const pp = byId.patients.get(p.id);
        Object.assign(pp, { status: 'exported', exported_batch_id: b.id, updated_at: nowIso() });
        if (!pp.review_date || pp.review_date <= todayS()) pp.review_date = addDays(todayS(), +S('refill_default') || 30);
      }
      audit('تصدير المستحقين', 'export', b.id, { 'العدد': pats.length, 'الملف': fname });
    }
    return fileBlob(data, fname, ext === 'zip' ? 'application/zip' : null);
  });
  route('GET', '/api/export/batches', () => { auth({ roles: ['admin', 'employee'] }); return D.export_batches.filter((b) => isAdmin() || b.created_by === USER.id).sort((a, b) => b.id - a.id).slice(0, 200); });
  route('GET', '/api/export/batches/:id/file', async ({ p: prm }) => {
    auth({ roles: ['admin', 'employee'] });
    const b = byId.export_batches.get(+prm.id); if (!b || (!isAdmin() && b.created_by !== USER.id)) throw new ApiError('الدفعة غير موجودة', 404);
    return fileBlob(await KV.get('export:' + b.id), b.file, b.file.endsWith('.zip') ? 'application/zip' : null);
  });
  route('POST', '/api/export/batches/:id/undo', ({ p: prm }) => {
    auth({ roles: ['admin'] }); let n = 0;
    for (const it of D.export_items.filter((x) => x.batch_id === +prm.id)) { const p = byId.patients.get(it.patient_id); if (p && p.status === 'exported' && p.exported_batch_id === +prm.id) { p.status = 'eligible'; p.exported_batch_id = null; n++; } }
    audit('إلغاء دفعة تصدير', 'export', +prm.id, { 'أعيد للمستحقين': n });
    return { ok: true, restored: n };
  });
  route('POST', '/api/export/test', async () => {
    auth({ roles: ['admin'] });
    let pats = candidates({}).slice(0, 5);
    if (!pats.length) pats = [{ mrn: '100001', name: 'مريض تجريبي للتجربة', phone: '0500000000', phone2: '', clinic: 'الباطنية', city: '', district: 'حي تجريبي', address: 'RRRD2929', out_of_city: 0, review_date: todayS() }];
    const tpl = S('export_template');
    const tb = await tplBytes(tpl);
    const [data, ext] = await XL.buildExport(pats, tpl, tb, todayS(), false, exportCtx(tpl));
    return fileBlob(data, 'تجربة_التصدير.' + ext);
  });

  // الإعدادات
  route('GET', '/api/settings', () => {
    auth();
    return { values: Object.fromEntries(['org_name', 'org_sub', 'logo', 'brand_color', 'lead_days', 'home_city', 'pharmacist_scope', 'refill_options', 'refill_default', 'color_actions', 'export_template', 'wa_templates'].map((k) => [k, S(k)])),
      export_fields: XL.EXPORT_FIELDS, color_labels: XL.COLOR_LABELS, action_labels: ACTION_LABELS };
  });
  route('PUT', '/api/settings', ({ body: d }) => {
    auth({ roles: ['admin'] }); const ch = [];
    if ('org_name' in d) { setS('org_name', String(d.org_name || '').trim().slice(0, 100)); ch.push('اسم الجهة'); }
    if ('org_sub' in d) { setS('org_sub', String(d.org_sub || '').trim().slice(0, 120)); ch.push('الوصف'); }
    if ('logo' in d) { const l = String(d.logo || ''); if (l && !/^data:image\//.test(l)) throw new ApiError('الشعار يجب أن يكون صورة'); if (l.length > 700000) throw new ApiError('حجم الشعار كبير، استخدم صورة أصغر من 500 كيلوبايت'); setS('logo', l); ch.push('الشعار'); }
    if ('brand_color' in d) { if (!/^#[0-9a-fA-F]{6}$/.test(d.brand_color || '')) throw new ApiError('لون غير صحيح'); setS('brand_color', d.brand_color); ch.push('لون الهوية'); }
    if ('refill_default' in d) { const n = +d.refill_default; if (!(n >= 1 && n <= 365)) throw new ApiError('مدة الصرف بين 1 و365 يومًا'); setS('refill_default', n); ch.push('مدة الصرف الافتراضية'); }
    if ('lead_days' in d) { const n = +d.lead_days; if (!(n >= 0 && n <= 60)) throw new ApiError('عدد الأيام بين 0 و60'); setS('lead_days', n); ch.push('أيام الظهور المسبق'); }
    if ('home_city' in d) { setS('home_city', String(d.home_city || '').trim()); ch.push('المدينة'); }
    if (['assigned', 'all'].includes(d.pharmacist_scope)) { setS('pharmacist_scope', d.pharmacist_scope); ch.push('نطاق الصيدلي'); }
    if ('refill_options' in d) { setS('refill_options', [...new Set(d.refill_options.map(Number).filter((x) => x >= 1 && x <= 365))].sort((a, b) => a - b)); ch.push('مدد الصرف'); }
    if (d.color_actions) { setS('color_actions', Object.fromEntries(Object.entries(d.color_actions).filter(([k, v]) => XL.COLOR_CLASSES.includes(k) && ACTION_LABELS[v]))); ch.push('ألوان الاستيراد'); }
    if (d.wa_templates) { setS('wa_templates', { ...S('wa_templates'), ...Object.fromEntries(Object.entries(d.wa_templates).filter(([k]) => k in DEFAULT_SETTINGS.wa_templates)) }); ch.push('رسائل واتساب'); }
    if (d.export_columns) {
      const tpl = S('export_template');
      tpl.columns = d.export_columns.map((c, i) => {
        let f = c.field || 'empty'; if (!XL.EXPORT_FIELDS[f] && !f.startsWith('fixed:')) f = 'empty';
        return { col: +c.col || i + 1, header: String(c.header || '').trim(), field: f, type: ['num', 'str'].includes(c.type) ? c.type : 'str' };
      });
      if ('ref_start' in d) tpl.ref_start = Math.max(0, +d.ref_start || 1);
      setS('export_template', tpl); ch.push('أعمدة التصدير');
    }
    audit('تعديل الإعدادات', 'settings', null, ch);
    return { ok: true };
  });
  route('POST', '/api/settings/export-template', async ({ body: f }) => {
    auth({ roles: ['admin'] });
    const file = f.get('file');
    if (!file || !/\.(xlsx|xlsm)$/i.test(file.name)) throw new ApiError('ارفع قالب Excel بصيغة xlsx');
    const buf = await file.arrayBuffer();
    let info;
    try { info = await XL.readTemplateHeaders(buf); } catch (e) { throw new ApiError('تعذر قراءة القالب: ' + String(e.message || e).slice(0, 200)); }
    if (!info.columns.length) throw new ApiError('لم أجد صف عناوين في القالب');
    await KV.set('export_template', buf);
    const tpl = { has_file: true, builtin: false, sheet: info.sheet, header_row: info.header_row, file_name: file.name, columns: info.columns, ref_start: info.ref_start, sheet_path: info.sheet_path };
    setS('export_template', tpl);
    audit('رفع قالب التصدير', 'settings', null, { 'الملف': file.name, 'الأعمدة': info.columns.length });
    return { ok: true, template: tpl };
  });
  route('DELETE', '/api/settings/export-template', async () => {
    auth({ roles: ['admin'] });
    setS('export_template', clone(DEFAULT_SETTINGS.export_template));
    await KV.del('export_template');
    audit('حذف قالب التصدير', 'settings');
    return { ok: true };
  });
  route('GET', '/api/settings/export-template/file', async () => {
    auth({ roles: ['admin'] });
    const tpl = S('export_template');
    return fileBlob(await tplBytes(tpl), tpl.builtin ? 'SPL_template.xlsx' : (tpl.file_name || 'template.xlsx'));
  });

  // السجل والنسخ الاحتياطي
  route('GET', '/api/audit', ({ q: a }) => {
    auth({ roles: ['admin'] });
    const rows = D.audit.filter((x) => (!/^\d+$/.test(a.user || '') || x.user_id === +a.user) && (!a.from || x.at >= a.from) && (!a.to || x.at < addDays(a.to, 1))
      && (!a.q || like(x.action, a.q) || like(x.details, a.q))).sort((x, y) => y.id - x.id);
    const page = Math.max(+a.page || 1, 1);
    return { total: rows.length, items: rows.slice((page - 1) * 100, page * 100).map((x) => ({ ...x, mrn: x.entity === 'patient' && byId.patients.get(x.entity_id) ? byId.patients.get(x.entity_id).mrn : null })) };
  });
  route('GET', '/api/backup', async () => {
    auth({ roles: ['admin'] });
    audit('تنزيل نسخة احتياطية', 'backup');
    const tpl = await KV.get('export_template');
    const blob = new Blob([JSON.stringify({ version: 1, db: D, template: tpl ? [...new Uint8Array(tpl)] : null })], { type: 'application/json' });
    return fileBlob(blob, `نسخة_احتياطية_${stamp()}.json`, 'application/json');
  });
  route('POST', '/api/restore', async ({ body: f }) => {
    auth({ roles: ['admin'] });
    const file = f.get('file');
    let j; try { j = JSON.parse(await file.text()); } catch (e) { throw new ApiError('ملف النسخة الاحتياطية غير صالح'); }
    if (!j.db || !j.db.users) throw new ApiError('ملف النسخة الاحتياطية غير صالح');
    D = j.db; for (const t of TABLES) D[t] = D[t] || []; reindex(); await save();
    if (j.template) await KV.set('export_template', new Uint8Array(j.template).buffer);
    SESS.del('sid');
    return { ok: true };
  });

  // ------------------------------------------------ المنفذ
  const READONLY = new Set(['GET']);
  let queue = Promise.resolve();
  function handle(method, url, body, session) {   // تنفيذ الطلبات بالتتابع داخل النسخة الواحدة
    const run = queue.then(() => { if (session) Object.assign(SESS, session); return handleNow(method, url, body); });
    queue = run.catch(() => {});
    return run;
  }
  async function handleNow(method, url, body) {
    const u = new URL(url, 'http://x');
    const q = Object.fromEntries(u.searchParams.entries());
    const r = R.find((x) => x.method === method && x.re.test(u.pathname));
    if (!r) return { ok: false, status: 404, data: { error: 'غير موجود' } };
    const m = u.pathname.match(r.re);
    const writes = !READONLY.has(method) || u.pathname === '/api/backup';
    for (let attempt = 0; attempt < 4; attempt++) {
      await load();
      const startVer = myVer;
      USER = null;
      try {
        const data = await r.fn({ p: m.groups || {}, q, body: body ?? {} });
        if (writes) {
          if (ENV.server && attempt < 3) {   // لو عدّلت نسخة أخرى البيانات أثناء التنفيذ: أعد التحميل وكرر
            const cur = await KV.get('ver');
            if (cur !== startVer) { D = null; myVer = null; continue; }
          }
          await save();
        }
        return { ok: true, status: 200, data };
      } catch (e) {
        if (e instanceof ApiError) return { ok: false, status: e.code, data: { error: e.message, ...e.extra } };
        console.error(e);
        return { ok: false, status: 500, data: { error: 'خطأ غير متوقع: ' + (e.message || e) } };
      }
    }
    return { ok: false, status: 503, data: { error: 'النظام مشغول، حاول مرة أخرى' } };
  }
  function configure(o) { if (o.kv) KV = o.kv; if (o.session) Object.assign(SESS, o.session); if ('server' in o) ENV.server = o.server; if (o.seedAdminHash) ENV.seedAdminHash = o.seedAdminHash; }
  async function firstRunInfo() { await load(); return D.first_password; }
  async function resetAll() { await KV.del('ver'); await KV.del('db'); await KV.del('export_template'); D = null; try { sessionStorage.clear(); } catch (e) {} }
  return { handle, firstRunInfo, resetAll, configure, hashPw };
})();


// ------------------------------------------------------------------ ربط الخادم
const SEED_ADMIN_HASH = 'p2$e3f9eb61e642d19b1fa407273e4e567c$5b44960c6c0107686b040b0cbbb1e3931328755e1997167d551b27ba1b8c6122';
let store = null;
const st = () => (store ||= getStore({ name: 'pharma-system', consistency: 'strong' }));
const textKey = (k) => k === 'ver' || k.startsWith('imp:') || k === 'secret';
const KVS = {
  async get(k) {
    if (k === 'db') return st().get(k, { type: 'json' });
    if (textKey(k)) return st().get(k, { type: 'text' });
    return st().get(k, { type: 'arrayBuffer' });
  },
  async set(k, v) {
    if (k === 'db') return st().setJSON(k, v);
    if (v instanceof Uint8Array) v = v.buffer.slice(v.byteOffset, v.byteOffset + v.byteLength);
    else if (v && typeof v.arrayBuffer === 'function' && !(v instanceof ArrayBuffer)) v = await v.arrayBuffer();
    return st().set(k, v);
  },
  async setMany(o) { for (const [k, v] of Object.entries(o)) if (k !== 'ver') await this.set(k, v); if ('ver' in o) await this.set('ver', o.ver); },
  async del(k) { return st().delete(k); },
};

// ---- كوكيز موقعة
let SECRET = null;
async function secret() {
  if (SECRET) return SECRET;
  let s = process.env.SESSION_SECRET || await st().get('secret', { type: 'text' });
  if (!s) { s = [...crypto.getRandomValues(new Uint8Array(32))].map((b) => b.toString(16).padStart(2, '0')).join(''); await st().set('secret', s); }
  SECRET = await crypto.subtle.importKey('raw', new TextEncoder().encode(s), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign', 'verify']);
  return SECRET;
}
const b64u = (buf) => Buffer.from(buf).toString('base64url');
async function sign(v) { return b64u(await crypto.subtle.sign('HMAC', await secret(), new TextEncoder().encode(v))); }
const MAXAGE = { sid: 12 * 3600, pt: 2 * 3600 };
async function makeCookie(name, value) {
  const exp = Math.floor(Date.now() / 1000) + MAXAGE[name];
  const payload = b64u(new TextEncoder().encode(JSON.stringify({ v: value, exp })));
  return `${name}=${payload}.${await sign(name + '.' + payload)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${MAXAGE[name]}`;
}
async function readCookie(cookies, name) {
  const raw = cookies[name]; if (!raw) return null;
  const [payload, sig] = raw.split('.');
  if (!payload || !sig || sig !== await sign(name + '.' + payload)) return null;
  try { const o = JSON.parse(Buffer.from(payload, 'base64url').toString()); return o.exp > Date.now() / 1000 ? o.v : null; } catch (e) { return null; }
}

LOCAL.configure({ kv: KVS, server: true, seedAdminHash: SEED_ADMIN_HASH });

const DEFAULT_HANDLER = async (req) => {
  const url = new URL(req.url);
  const cookies = Object.fromEntries((req.headers.get('cookie') || '').split(/;\s*/).filter(Boolean).map((c) => { const i = c.indexOf('='); return [c.slice(0, i), decodeURIComponent(c.slice(i + 1))]; }));
  const sess = { sid: await readCookie(cookies, 'sid'), pt: await readCookie(cookies, 'pt') };
  const setCookies = [];
  const session = {
    get: (k) => sess[k] ?? null,
    set: (k, v) => { sess[k] = v; setCookies.push(['set', k, v]); },
    del: (k) => { sess[k] = null; setCookies.push(['del', k]); },
  };
  let body;
  const ct = req.headers.get('content-type') || '';
  if (req.method !== 'GET' && req.method !== 'HEAD') {
    if (ct.includes('multipart/form-data')) body = await req.formData();
    else { const t = await req.text(); try { body = t ? JSON.parse(t) : {}; } catch (e) { body = {}; } }
  }
  if (req.method !== 'GET') {   // حماية من الطلبات القادمة من مواقع أخرى
    const origin = req.headers.get('origin');
    if (origin && new URL(origin).host !== url.host) return Response.json({ error: 'طلب غير مسموح' }, { status: 403 });
  }
  const r = await LOCAL.handle(req.method, url.pathname + url.search, body, session);
  const headers = new Headers({ 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
  for (const c of setCookies) headers.append('Set-Cookie', c[0] === 'set' ? await makeCookie(c[1], c[2]) : `${c[1]}=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0`);
  if (r.ok && r.data && r.data.__file) {
    const f = r.data;
    let data = f.data;
    if (data && typeof data.arrayBuffer === 'function' && !(data instanceof ArrayBuffer)) data = await data.arrayBuffer();
    headers.set('Content-Type', f.type || 'application/octet-stream');
    headers.set('Content-Disposition', `${f.inline ? 'inline' : 'attachment'}; filename*=UTF-8''${encodeURIComponent(f.name)}`);
    return new Response(data, { status: 200, headers });
  }
  headers.set('Content-Type', 'application/json; charset=utf-8');
  return new Response(JSON.stringify(r.data), { status: r.status, headers });
};



// ------------------------------------------------------------------ خادم HTTP
const PUB = path.join(path.dirname(fileURLToPath(import.meta.url)), 'public');
const handler = (await Promise.resolve(DEFAULT_HANDLER));
function backupDaily() {   // نسخة احتياطية يومية على القرص (آخر 14)
  try {
    const dir = path.join(DATA_DIR, 'backups'); fs.mkdirSync(dir, { recursive: true });
    const name = path.join(dir, 'db_' + new Date(Date.now() + 3 * 3600e3).toISOString().slice(0, 10) + '.json');
    const src = path.join(DATA_DIR, 'db');
    if (fs.existsSync(src) && !fs.existsSync(name)) fs.copyFileSync(src, name);
    fs.readdirSync(dir).filter((x) => x.startsWith('db_')).sort().slice(0, -14).forEach((x) => fs.unlinkSync(path.join(dir, x)));
  } catch (e) { console.error('backup', e); }
}
backupDaily(); setInterval(backupDaily, 3600e3);
http.createServer(async (req, res) => {
  try {
    const proto = req.headers['x-forwarded-proto'] || 'http';
    const url = new URL(req.url, `${proto}://${req.headers.host}`);
    if (url.pathname.startsWith('/api/')) {
      const chunks = []; let size = 0;
      for await (const c of req) { size += c.length; if (size > 25e6) { res.writeHead(413); return res.end('{"error":"حجم الطلب كبير"}'); } chunks.push(c); }
      const body = ['GET', 'HEAD'].includes(req.method) ? undefined : Buffer.concat(chunks);
      const r = await handler(new Request(url, { method: req.method, headers: req.headers, body, duplex: 'half' }));
      const h = {}; r.headers.forEach((v, k) => { if (k !== 'set-cookie') h[k] = v; });
      const sc = r.headers.getSetCookie(); if (sc.length) h['set-cookie'] = sc;
      res.writeHead(r.status, h); return res.end(Buffer.from(await r.arrayBuffer()));
    }
    if (url.pathname === '/health') { res.writeHead(200); return res.end('ok'); }
    const file = url.pathname === '/portal' || url.pathname === '/portal/' || url.pathname === '/portal.html' ? 'portal.html' : 'index.html';
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-cache', 'X-Frame-Options': 'DENY', 'X-Content-Type-Options': 'nosniff', 'Referrer-Policy': 'no-referrer' });
    res.end(fs.readFileSync(path.join(PUB, file)));
  } catch (e) { console.error(e); res.writeHead(500); res.end('{"error":"خطأ في الخادم"}'); }
}).listen(+process.env.PORT || 8080, '0.0.0.0', () => console.log('pharmacy-system listening on', process.env.PORT || 8080, 'data:', DATA_DIR));
