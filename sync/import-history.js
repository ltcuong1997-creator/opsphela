/**
 * Nạp doanh thu lịch sử từ file báo cáo nội bộ (vd d05.xlsx, sheet "Data": mỗi dòng
 * 1 cửa hàng × 1 ngày) lên histDays/{YYYY-MM-DD}, và điền AM/DM/Khu vực/City còn trống
 * trong danh mục stores.
 *
 * Doanh thu trong file KHÁC cách tính của Fabi D05:
 *   Doanh thu file = (Doanh thu D05 + Hoa hồng app) / 1,08
 * (trước VAT 8%, chưa trừ hoa hồng Shopee/Be/Xanh Ngon) - đã đối chiếu khớp 75/75 cửa
 * hàng ngày 01/09/2026. Vì vậy không ghi vào d05Daily/d05Days mà để collection riêng:
 *
 *   histDays/{date} = { date, revenue, tc, storeCount, source: 'excel',
 *     stores: { [mã Fabi]: { name, revenue, tc, sources: { [key]: { name, revenue, tc } },
 *                            hours: { '07': doanh thu, ... } } } }
 *
 * Dòng có doanh thu > 300tr mà không có TC là tổng CẢ THÁNG ghi vào 1 ngày (file có ở
 * 01/03, 01/04/2025...) -> tách sang histMonths/{YYYY-MM} = { month, revenue,
 * stores: { [mã Fabi]: { name, revenue } } } để biểu đồ ngày không bị vọt.
 *
 * Ngày lấy theo 3 cột Năm/Tháng/Ngày (cột "Ngày tháng năm" lệch 1 ngày do múi giờ).
 * Bỏ dòng có ngày ngoài khoảng 2024-01-01 → hôm nay (gõ nhầm năm 2056...).
 * Tên cửa hàng khớp mã Fabi như import-targets.js (+ store-aliases.json).
 * Chỉ ghi ngày mới hoặc đổi số. Danh mục stores: chỉ điền ô trống, không đè thứ đã sửa tay.
 *
 *   node import-history.js --dry-run d05.xlsx [--sheet Data]
 *   node import-history.js d05.xlsx [--sheet Data] [--no-stores]
 */

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const admin = require('firebase-admin');
const { loadEnv, initFirebase } = require('./lib');

const arg = (n) => { const i = process.argv.indexOf(n); return i > 0 ? process.argv[i + 1] : null; };
const norm = (s) => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/đ/gi, 'd').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const storeKey = (s) => norm(s).replace(/^(phe la|pl)\s+/, '');
const keyOf = (s) => norm(s).replace(/ /g, '-') || 'khong-ro';
const num = (v) => (typeof v === 'number' ? v : Number(String(v == null ? '' : v).replace(/[^0-9.-]/g, '')) || 0);
const pad = (n) => String(n).padStart(2, '0');

// Cột kênh bán: "MANG VỀ" / "TC MANG VỀ", riêng giao hàng gộp là "DT DELI" / "TC DELI"
const SOURCES = ['MANG VỀ', 'TẠI CHỖ', 'ShopeeFood', 'GRABFOOD', 'Ahamove', 'BEFOOD', 'GOJEK (GOVIET)', 'Xanh Ngon', 'APP_PHELA'];
const MONTH_LUMP = 3e8; // dòng không TC mà lớn hơn mức này = tổng tháng
// Khu vực / City trong file -> tên dùng trên web (seed-stores.js)
const REGION = { 'hai phong': 'Miền Bắc' };
const CITY = { 'ho chi minh': 'TP.HCM', 'tp ho chi minh': 'TP.HCM', 'tphcm': 'TP.HCM', 'hcm': 'TP.HCM' };

async function main() {
  loadEnv();
  const dryRun = process.argv.includes('--dry-run');
  const sheet = arg('--sheet') || 'Data';
  const file = process.argv.slice(2).find((a, i, all) => !a.startsWith('--') && all[i - 1] !== '--sheet');
  if (!file) throw new Error('Cách dùng: node import-history.js [--dry-run] [--sheet Data] <file.xlsx>');

  const db = initFirebase();
  const stores = {}, meta = {};
  (await db.collection('stores').get()).forEach((d) => {
    stores[storeKey(d.data().storeName)] = d.id;
    meta[d.id] = d.data();
  });
  const aliasFile = path.join(__dirname, 'store-aliases.json');
  for (const [name, id] of Object.entries(fs.existsSync(aliasFile) ? require(aliasFile) : {})) {
    if (!name.startsWith('_')) stores[storeKey(name)] = id;
  }

  const wb = XLSX.readFile(file, { sheets: [sheet] });
  if (!wb.Sheets[sheet]) throw new Error(`${path.basename(file)}: không có sheet "${sheet}"`);
  const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: null });
  const head = (matrix[0] || []).map((h) => norm(h));
  const col = (name) => head.indexOf(norm(name));
  const c = {
    year: col('Năm'), month: col('Tháng'), day: col('Ngày'), store: col('Cửa hàng'),
    revenue: col('Doanh Thu'), tc: col('TC'), am: col('AM'), dm: col('DM'), region: col('Khu vực'), city: col('Cities'),
  };
  for (const [k, i] of Object.entries(c)) if (i < 0) throw new Error(`Sheet "${sheet}" thiếu cột ${k}`);
  const hourCols = head.map((h, i) => [h.match(/^(\d{2})$/), i]).filter(([m]) => m).map(([m, i]) => [m[1], i]);

  const today = new Date().toISOString().slice(0, 10);
  const days = {}, months = {}, unmatched = {}, badDates = [], dup = [];
  const latestMeta = {}; // mã Fabi -> { date, am, dm, region, city } của dòng mới nhất
  let empty = 0;
  for (const l of matrix.slice(1)) {
    if (!l || !l[c.year]) continue;
    const date = `${l[c.year]}-${pad(l[c.month])}-${pad(l[c.day])}`;
    const revenue = num(l[c.revenue]), tc = num(l[c.tc]);
    if (!revenue && !tc) { empty++; continue; }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date < '2024-01-01' || date > today) { badDates.push(`${date} ${l[c.store]}`); continue; }
    const id = stores[storeKey(l[c.store])];
    if (!id) { unmatched[l[c.store]] = (unmatched[l[c.store]] || 0) + 1; continue; }

    if (!tc && revenue > MONTH_LUMP) {
      const m = (months[date.slice(0, 7)] ||= { month: date.slice(0, 7), revenue: 0, stores: {} });
      if (!m.stores[id]) { m.stores[id] = { name: (meta[id] || {}).storeName || String(l[c.store]), revenue: Math.round(revenue) }; m.revenue += Math.round(revenue); }
      continue;
    }
    const d = (days[date] ||= { date, revenue: 0, tc: 0, storeCount: 0, source: 'excel', stores: {} });
    if (d.stores[id]) { dup.push(`${date} ${l[c.store]}`); continue; } // file có dòng lặp y hệt
    const sources = {};
    for (const name of SOURCES) {
      const r = num(l[col(name)]), t = num(l[col('TC ' + name)]);
      if (r || t) sources[keyOf(name)] = { name, revenue: Math.round(r), tc: t };
    }
    const hours = {};
    for (const [h, i] of hourCols) if (num(l[i])) hours[h] = Math.round(num(l[i]));
    d.stores[id] = { name: (meta[id] || {}).storeName || String(l[c.store]), revenue: Math.round(revenue), tc, sources, hours };
    d.revenue += Math.round(revenue); d.tc += tc; d.storeCount++;

    if (!latestMeta[id] || latestMeta[id].date < date) {
      latestMeta[id] = { date, am: l[c.am], dm: l[c.dm], region: l[c.region], city: l[c.city] };
    }
  }

  const list = Object.values(days).sort((a, b) => a.date.localeCompare(b.date));
  const byMonth = {};
  for (const d of list) {
    const m = (byMonth[d.date.slice(0, 7)] ||= { days: 0, revenue: 0, stores: new Set() });
    m.days++; m.revenue += d.revenue; Object.keys(d.stores).forEach((s) => m.stores.add(s));
  }
  console.log(`${path.basename(file)} / ${sheet}: ${list.length} ngày ${list[0]?.date} → ${list[list.length - 1]?.date}, bỏ ${empty} dòng trống`);
  for (const [m, v] of Object.entries(byMonth)) {
    console.log(`  ${m}: ${v.days} ngày, ${v.stores.size} cửa hàng, doanh thu (trước VAT) ${Math.round(v.revenue).toLocaleString('vi-VN')} ₫`);
  }
  for (const m of Object.values(months)) console.log(`  histMonths ${m.month}: tổng tháng của ${Object.keys(m.stores).length} cửa hàng, ${m.revenue.toLocaleString('vi-VN')} ₫ (dòng không có TC)`);
  if (badDates.length) console.log(`  ⚠ bỏ ${badDates.length} dòng ngày sai: ${badDates.slice(0, 5).join(' | ')}${badDates.length > 5 ? ' …' : ''}`);
  if (dup.length) console.log(`  ⚠ bỏ ${dup.length} dòng trùng cửa hàng-ngày: ${dup.slice(0, 5).join(' | ')}`);
  const um = Object.entries(unmatched);
  if (um.length) console.log(`  ⚠ ${um.length} tên không khớp mã Fabi: ${um.map(([n, k]) => `${n} (${k} dòng)`).join(' | ')}`);

  // Danh mục stores: điền ô trống từ dòng mới nhất của mỗi cửa hàng
  const fill = [];
  if (!process.argv.includes('--no-stores')) {
    for (const [id, m] of Object.entries(latestMeta)) {
      const cur = meta[id] || {};
      const city = m.city ? CITY[norm(m.city)] || String(m.city).trim() : null;
      const add = {};
      if (!cur.am && m.am) add.am = String(m.am).trim();
      if (!cur.dm && m.dm) add.dm = String(m.dm).trim();
      if (!cur.region && m.region) add.region = REGION[norm(m.region)] || String(m.region).trim();
      if (!cur.city && city) add.city = city;
      if (Object.keys(add).length) fill.push([id, cur.storeName || id, add]);
    }
    const vals = (k) => [...new Set(fill.map(([, , a]) => a[k]).filter(Boolean))].join(', ');
    console.log(`Danh mục stores: điền thêm cho ${fill.length} cửa hàng (chỉ ô trống)`);
    console.log(`  Khu vực: ${vals('region')}\n  City: ${vals('city')}`);
    for (const [id, name, add] of fill) console.log(`  ${id} ${name}: ${Object.entries(add).map(([k, v]) => `${k}=${v}`).join(', ')}`);
  }

  // So với histDays đã có -> chỉ ghi ngày đổi số
  const existing = {};
  for (let i = 0; i < list.length; i += 300) {
    const refs = list.slice(i, i + 300).map((d) => db.collection('histDays').doc(d.date));
    (await db.getAll(...refs)).forEach((s) => { if (s.exists) existing[s.id] = JSON.stringify(s.data().stores); });
  }
  const changed = list.filter((d) => existing[d.date] !== JSON.stringify(d.stores));
  console.log(`histDays: ${changed.length} ngày mới hoặc đổi số, ${list.length - changed.length} giữ nguyên`);
  if (dryRun) { console.log('(chạy thử, chưa ghi gì)'); return; }

  for (let i = 0; i < changed.length; i += 20) { // doc ~60KB, 20 doc/batch cho dưới giới hạn 10MB
    const batch = db.batch();
    for (const d of changed.slice(i, i + 20)) batch.set(db.collection('histDays').doc(d.date), { ...d, updatedAt: admin.firestore.FieldValue.serverTimestamp() });
    await batch.commit();
  }
  console.log(`  ✅ đã ghi ${changed.length} ngày vào histDays`);
  const mlist = Object.values(months);
  if (mlist.length) {
    const batch = db.batch();
    for (const m of mlist) batch.set(db.collection('histMonths').doc(m.month), { ...m, source: 'excel', updatedAt: admin.firestore.FieldValue.serverTimestamp() });
    await batch.commit();
    console.log(`  ✅ đã ghi ${mlist.length} tháng vào histMonths`);
  }
  if (fill.length) {
    const batch = db.batch();
    for (const [id, , add] of fill) batch.set(db.collection('stores').doc(id), add, { merge: true });
    await batch.commit();
    console.log(`  ✅ đã điền danh mục cho ${fill.length} cửa hàng`);
  }
}

main().then(() => process.exit(0)).catch((e) => { console.error('❌ Lỗi:', e.message); process.exit(1); });
