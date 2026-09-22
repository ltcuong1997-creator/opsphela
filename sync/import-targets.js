/**
 * Nạp target doanh thu tháng × cửa hàng từ Excel/CSV lên targets/{YYYY-MM}_{storeId}
 * (cùng khuôn với trang "Target" trên web: { year, month, storeId, storeName, target }).
 *
 * Tự nhận 2 kiểu file:
 *  - Dọc: mỗi dòng 1 cửa hàng × 1 tháng, có cột Tháng, Năm (hoặc "Tháng" dạng 09/2026),
 *    Chi nhánh / Mã CH, Target.
 *  - Ngang: mỗi dòng 1 cửa hàng, mỗi cột 1 tháng ("Tháng 9", "T9", "09/2026", "2026-09"...).
 *    Cột tháng không ghi năm thì lấy năm ở --year (mặc định năm nay).
 * Tên chi nhánh được khớp với mã Fabi qua danh mục stores (bỏ dấu, bỏ tiền tố "Phê La").
 * Chỉ ghi target mới hoặc đổi số.
 *
 *   node import-targets.js --dry-run target.xlsx      xem đọc được gì, không ghi
 *   node import-targets.js target.xlsx [--year 2026]
 */

const XLSX = require('xlsx');
const path = require('path');
const { loadEnv, initFirebase } = require('./lib');
const admin = require('firebase-admin');

const arg = (n) => { const i = process.argv.indexOf(n); return i > 0 ? process.argv[i + 1] : null; };
const norm = (s) => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/đ/gi, 'd').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
// "Phê La – 145 Trích Sài" / "PL 145 Trích Sài" / "145 trich sai" -> "145 trich sai"
const storeKey = (s) => norm(s).replace(/^(phe la|pl)\s+/, '');
const toNumber = (v) => (typeof v === 'number' ? v : Number(String(v || '').replace(/[^0-9-]/g, '')) || 0);

const COLS = {
  store: ['chi nhanh', 'ten chi nhanh', 'cua hang', 'ten cua hang', 'store'],
  storeId: ['ma ch', 'ma cua hang', 'store id', 'ma chi nhanh'],
  month: ['thang', 'month', 'ky'],
  year: ['nam', 'year'],
  target: ['target', 'muc tieu', 'doanh thu muc tieu', 'chi tieu', 'ke hoach'],
};

// Tiêu đề cột trông như 1 tháng -> { year?, month }
function monthOfHeader(h, defaultYear) {
  const t = norm(h);
  let m = t.match(/^(?:thang|t|th|month)\s*(\d{1,2})(?:\s+(\d{4}))?$/);
  if (m && +m[1] >= 1 && +m[1] <= 12) return { month: +m[1], year: m[2] ? +m[2] : defaultYear };
  m = t.match(/^(\d{1,2})\s+(\d{4})$/); // 09/2026 -> "09 2026"
  if (m && +m[1] <= 12) return { month: +m[1], year: +m[2] };
  m = t.match(/^(\d{4})\s+(\d{1,2})$/); // 2026-09
  if (m && +m[2] <= 12) return { month: +m[2], year: +m[1] };
  return null;
}

function parseSheet(matrix, defaultYear) {
  // Dòng tiêu đề: trong 20 dòng đầu, dòng có cột cửa hàng + (cột target hoặc nhiều cột tháng)
  for (let r = 0; r < Math.min(matrix.length, 20); r++) {
    const heads = matrix[r] || [];
    const find = (names) => heads.findIndex((h) => names.includes(norm(h)));
    const c = { store: find(COLS.store), storeId: find(COLS.storeId), month: find(COLS.month), year: find(COLS.year), target: find(COLS.target) };
    if (c.store < 0 && c.storeId < 0) continue;
    const monthCols = heads.map((h, i) => [i, monthOfHeader(h, defaultYear)]).filter(([, m]) => m);
    const out = [];
    const body = matrix.slice(r + 1).filter((l) => l && l.some((v) => v != null && v !== ''));
    if (c.target >= 0 && c.month >= 0) {
      for (const l of body) {
        let month = l[c.month], year = c.year >= 0 ? +l[c.year] : defaultYear;
        const mm = typeof month === 'string' && monthOfHeader(month.includes('/') || month.includes('-') ? month : 'thang ' + month, year);
        if (mm) ({ month, year } = { month: mm.month, year: mm.year || year });
        out.push({ store: l[c.store], storeId: c.storeId >= 0 ? l[c.storeId] : null, year: +year, month: +month, target: toNumber(l[c.target]) });
      }
      return { kind: 'dọc', rows: out };
    }
    if (monthCols.length) {
      for (const l of body) {
        for (const [i, m] of monthCols) {
          if (l[i] == null || l[i] === '') continue;
          out.push({ store: l[c.store], storeId: c.storeId >= 0 ? l[c.storeId] : null, year: m.year, month: m.month, target: toNumber(l[i]) });
        }
      }
      return { kind: 'ngang', rows: out };
    }
  }
  return null;
}

async function main() {
  loadEnv();
  const dryRun = process.argv.includes('--dry-run');
  const defaultYear = +(arg('--year') || new Date().getFullYear());
  const files = process.argv.slice(2).filter((a, i, all) => !a.startsWith('--') && all[i - 1] !== '--year');
  if (!files.length) throw new Error('Cách dùng: node import-targets.js [--dry-run] [--year 2026] <file.xlsx> ...');

  const db = initFirebase();
  const stores = {};
  (await db.collection('stores').get()).forEach((d) => { stores[storeKey(d.data().storeName)] = { id: d.id, name: d.data().storeName }; });
  const existing = {};
  (await db.collection('targets').get()).forEach((d) => { existing[d.id] = d.data().target; });

  for (const file of files) {
    const wb = XLSX.readFile(file);
    const parsed = wb.SheetNames.map((n) => parseSheet(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null }), defaultYear)).filter(Boolean);
    if (!parsed.length) throw new Error(`${path.basename(file)}: không tìm thấy cột Chi nhánh + Target/Tháng`);

    const ok = {}, unmatched = new Set(), bad = [];
    for (const p of parsed) {
      for (const r of p.rows) {
        const id = r.storeId && /^\d+$/.test(String(r.storeId).trim()) ? String(r.storeId).trim() : (stores[storeKey(r.store)] || {}).id;
        if (!id) { if (r.store) unmatched.add(String(r.store)); continue; }
        if (!(r.month >= 1 && r.month <= 12) || !(r.year > 2000) || !r.target) { bad.push(r); continue; }
        const key = `${r.year}-${String(r.month).padStart(2, '0')}_${id}`;
        const name = (Object.values(stores).find((s) => s.id === id) || {}).name || String(r.store || '');
        ok[key] = { year: r.year, month: r.month, storeId: id, storeName: name, target: (ok[key]?.target || 0) + r.target };
      }
    }
    const list = Object.entries(ok);
    const changed = list.filter(([k, t]) => existing[k] !== t.target);
    const byMonth = {};
    for (const [, t] of list) byMonth[`${t.month}/${t.year}`] = (byMonth[`${t.month}/${t.year}`] || 0) + t.target;
    console.log(`${path.basename(file)} (kiểu ${parsed.map((p) => p.kind).join('+')}): ${list.length} target, ${new Set(list.map(([, t]) => t.storeId)).size} cửa hàng`);
    for (const [m, v] of Object.entries(byMonth)) console.log(`  tháng ${m}: tổng target ${v.toLocaleString('vi-VN')} ₫`);
    if (unmatched.size) console.log(`  ⚠ ${unmatched.size} chi nhánh không khớp được mã Fabi: ${[...unmatched].join(' | ')}`);
    if (bad.length) console.log(`  ⚠ bỏ ${bad.length} dòng thiếu tháng/năm/target`);
    console.log(`  ${changed.length} target mới hoặc đổi số, ${list.length - changed.length} giữ nguyên`);
    if (dryRun || !changed.length) continue;

    for (let i = 0; i < changed.length; i += 400) {
      const batch = db.batch();
      for (const [k, t] of changed.slice(i, i + 400)) {
        batch.set(db.collection('targets').doc(k), { ...t, updatedAt: admin.firestore.FieldValue.serverTimestamp() }, { merge: true });
      }
      await batch.commit();
    }
    console.log(`  ✅ đã ghi ${changed.length} target`);
  }
}

main().catch((e) => { console.error('❌ Lỗi:', e.message); process.exit(1); });
