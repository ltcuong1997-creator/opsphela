/**
 * Phần dùng chung cho sync-report.js (tự động, hôm nay) và import-file.js
 * (nạp tay file Excel D05 cũ): đọc file, gộp ngày × cửa hàng × món, và chỉ
 * ghi lên Firestore những document thật sự thay đổi.
 *
 * Cấu trúc Firestore:
 *   d05Daily/{date}_{storeId}    1 doc = 1 cửa hàng trong 1 ngày (storeId = mã Fabi)
 *     { date, storeId, storeName, revenue, quantity, bills, hash, updatedAt,
 *       items: { [itemKey]: { code, name, group, quantity, revenue } } }
 *   syncState/d05_{date}         { hashes: { [docId]: hash } }
 *   d05Days/{date}               tóm tắt cả ngày cho trang web: tổng ngày, tổng từng
 *     cửa hàng (stores{}), tổng từng món cả chuỗi (items{}) - web đọc 1 doc/ngày
 *     thay vì ~90 doc nặng, xem 1 tháng chỉ tốn ~30 lượt đọc.
 *
 * Mốc EXCEL_UNTIL: các ngày tới hết mốc lấy số từ báo cáo nội bộ Excel (hist-to-days.js
 * chép histDays -> d05Days); Fabi chỉ tự nạp các ngày SAU mốc. Để cùng cách tính với
 * Excel và target, d05Days của ngày sau mốc ghi doanh thu TRƯỚC VAT:
 *   revenue = (doanh thu Fabi + hoa hồng app) / 1,08     (số gốc giữ ở revenueGross)
 * Riêng doanh thu từng món (items) vẫn là số gốc Fabi.
 *
 * Mỗi lần ghi, đọc 1 doc syncState của ngày đó (1 lượt đọc), so hash từng
 * cửa hàng, chỉ ghi các cửa hàng có số liệu khác đi + cập nhật lại syncState.
 * Không có gì đổi thì không ghi gì cả.
 */

const XLSX = require('xlsx');
const admin = require('firebase-admin');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const COLLECTION = 'd05Daily';
const STATE_COLLECTION = 'syncState';
const DAYS_COLLECTION = 'd05Days';
const STORES_COLLECTION = 'stores';

// Tới hết ngày này: số trên d05Days lấy từ Excel, không để Fabi ghi đè (xem đầu file).
const EXCEL_UNTIL = '2026-09-20';
const VAT = 1.08;
// Nguồn đơn mà Fabi ghi doanh thu ĐÃ trừ hoa hồng app (Grab, Ahamove thì không) -
// đối chiếu với báo cáo nội bộ ngày 01/09/2026.
const COMMISSION_SOURCES = ['shopeefood', 'befood', 'xanh-ngon'];

// Tên cột trong file Excel D05 (khác bảng trên web: "Ngày" thay cho "Thời
// gian", "Hoá đơn" thay cho "Mã hoá đơn", và có thêm cột "Cửa hàng").
const COL = {
  time: 'Ngày',
  bill: 'Hoá đơn',
  itemCode: 'Mã hàng',
  itemName: 'Tên hàng',
  group: 'Nhóm món',
  type: 'Loại món',
  quantity: 'Số lượng',
  revenue: 'Tổng tiền',
  store: 'Cửa hàng',
  source: 'Nguồn',
  payment: 'PTTT',
  area: 'Khu vực',
  discount: 'Giảm giá',
  serviceFee: 'Phí dịch vụ',
  tax: 'Thuế',
  shipFee: 'Phí ship',
  commission: 'Hoa hồng',
};

function loadEnv() {
  const envFile = path.join(__dirname, '.env');
  if (fs.existsSync(envFile)) process.loadEnvFile(envFile);
}

function initFirebase() {
  let json = process.env.FIREBASE_SERVICE_ACCOUNT;
  if (!json && process.env.FIREBASE_SERVICE_ACCOUNT_FILE) {
    json = fs.readFileSync(process.env.FIREBASE_SERVICE_ACCOUNT_FILE, 'utf8');
  }
  if (!json) throw new Error('Thiếu FIREBASE_SERVICE_ACCOUNT hoặc FIREBASE_SERVICE_ACCOUNT_FILE');
  admin.initializeApp({ credential: admin.credential.cert(JSON.parse(json)) });
  return admin.firestore();
}

// Tên cột có thể gặp cho từng trường (so khớp NGUYÊN tên sau khi bỏ dấu, nên "Giảm giá VAT"
// không bị nhầm thành "Giảm giá", "Tổng tiền (không bao gồm VAT)" không nhầm "Tổng tiền").
// Thêm tên mới vào đây nếu file Excel dùng tên cột khác.
const ALIASES = {
  time: ['Ngày', 'Thời gian', 'Ngày bán', 'Ngày tạo', 'Ngày giờ'],
  bill: ['Hoá đơn', 'Hóa đơn', 'Mã hoá đơn', 'Mã hóa đơn', 'Mã HĐ'],
  itemCode: ['Mã hàng', 'Mã món'],
  itemName: ['Tên hàng', 'Tên món'],
  group: ['Nhóm món', 'Nhóm hàng'],
  type: ['Loại món', 'Loại hàng'],
  quantity: ['Số lượng', 'SL'],
  revenue: ['Tổng tiền'],
  store: ['Cửa hàng', 'Chi nhánh', 'Tên cửa hàng'],
  source: ['Nguồn', 'Nguồn đơn', 'Nguồn đơn hàng'],
  payment: ['PTTT', 'Phương thức thanh toán', 'Hình thức thanh toán'],
  area: ['Khu vực'],
  discount: ['Giảm giá'],
  serviceFee: ['Phí dịch vụ'],
  tax: ['Thuế'],
  shipFee: ['Phí ship', 'Phí giao hàng'],
  commission: ['Hoa hồng'],
};
const REQUIRED = ['time', 'bill', 'quantity', 'revenue'];

// Tìm dòng tiêu đề (trong 30 dòng đầu, dòng khớp nhiều tên cột nhất) và vị trí từng trường.
function detectColumns(matrix) {
  let best = { row: -1, map: {}, score: 0 };
  for (let r = 0; r < Math.min(matrix.length, 30); r++) {
    const heads = (matrix[r] || []).map((h) => (h == null ? '' : keyOf(h)));
    const map = {};
    for (const [field, names] of Object.entries(ALIASES)) {
      const i = heads.findIndex((h) => names.some((n) => keyOf(n) === h));
      if (i >= 0) map[field] = i;
    }
    const score = Object.keys(map).length;
    if (score > best.score) best = { row: r, map, score };
  }
  return best;
}

// Đọc mọi file Excel bán hàng kiểu D05 thành các dòng có key = tên cột chuẩn (COL), dù file
// đặt tên cột khác hay có vài dòng tiêu đề phụ ở trên. rows.info cho biết đã nhận ra
// những cột nào, thiếu cột nào, đọc những sheet nào.
//
// File D05 "tất cả cửa hàng" của Fabi: mỗi cửa hàng có bán là 1 sheet, tên sheet dạng
// "108654-Phê La – 145 Trích..." (mã cửa hàng trên Fabi + tên bị cắt). Dòng cuối mỗi sheet
// là dòng tổng cộng (không có Ngày) - aggregate() tự bỏ qua. Ngoài ra có 1 sheet "Tất cả
// cửa hàng" chứa LẠI toàn bộ dòng của các sheet trên - phải bỏ, không thì doanh thu bị nhân
// đôi. File không có sheet theo mã cửa hàng thì đọc mọi sheet có đủ cột.
function readRows(filePath) {
  const workbook = XLSX.readFile(filePath, { cellDates: true });
  const storeSheets = workbook.SheetNames.filter((n) => /^\d+-/.test(n));
  const sheets = storeSheets.length ? storeSheets : workbook.SheetNames;
  const rows = [];
  const info = { sheets: 0, skippedSheets: [], found: new Set(), missing: [] };
  for (const name of sheets) {
    const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: null });
    const { row, map } = detectColumns(matrix);
    if (REQUIRED.some((f) => map[f] == null)) {
      info.skippedSheets.push(name);
      continue;
    }
    info.sheets++;
    Object.keys(map).forEach((f) => info.found.add(f));
    const storeId = name.match(/^(\d+)-/)?.[1] || null;
    for (let r = row + 1; r < matrix.length; r++) {
      const line = matrix[r];
      if (!line) continue;
      const obj = { __storeId: storeId };
      for (const [field, i] of Object.entries(map)) obj[COL[field]] = line[i];
      rows.push(obj);
    }
  }
  info.found = [...info.found];
  info.missing = Object.keys(ALIASES).filter((f) => !info.found.includes(f));
  rows.info = info;
  return rows;
}

function toNumber(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  // VND không có phần thập phân; dấu "," / "." chỉ là phân cách hàng nghìn.
  return Number(String(v).replace(/[^0-9-]/g, '')) || 0;
}

function normalizeDate(v) {
  // "22/09/2026 06:05" (dd/mm/yyyy) - new Date() sẽ hiểu nhầm thành mm/dd.
  if (!v) return null;
  if (v instanceof Date) {
    if (isNaN(v)) return null;
    const p = (n) => String(n).padStart(2, '0');
    return `${v.getFullYear()}-${p(v.getMonth() + 1)}-${p(v.getDate())}`;
  }
  const m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  const iso = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return iso ? iso[0] : null;
}

function keyOf(s) {
  // Firestore doc id / field name an toàn: bỏ dấu, chỉ giữ chữ số và "-".
  return String(s)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/gi, 'd')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '') || 'khong-ro';
}

// rows (từng dòng món trong hoá đơn) -> { [docId]: doc } gộp theo ngày × cửa hàng.
function aggregate(rows, storeIdByName = {}) {
  const docs = {};
  const bills = {};
  const sourceBills = {}; // id -> { sourceKey: Set(billId) } - để tính AOV riêng theo nguồn đơn
  for (const row of rows) {
    const date = normalizeDate(row[COL.time]);
    if (!date) continue;
    const storeName = String(row[COL.store] || 'Không rõ').trim();
    // File không có sheet theo mã cửa hàng: tra mã Fabi theo tên (danh mục stores), không thì
    // cùng 1 cửa hàng sẽ ra 2 doc khác nhau (mã số và tên-không-dấu).
    const storeId = row.__storeId || storeIdByName[keyOf(storeName)] || keyOf(storeName);
    const id = `${date}_${storeId}`;
    const doc = (docs[id] ||= {
      date, storeId, storeName, revenue: 0, quantity: 0, bills: 0,
      discount: 0, serviceFee: 0, tax: 0, shipFee: 0, commission: 0,
      items: {}, sources: {}, payments: {}, areas: {},
    });
    const qty = toNumber(row[COL.quantity]);
    const rev = toNumber(row[COL.revenue]);
    doc.revenue += rev;
    doc.quantity += qty;
    doc.discount += toNumber(row[COL.discount]);
    doc.serviceFee += toNumber(row[COL.serviceFee]);
    doc.tax += toNumber(row[COL.tax]);
    doc.shipFee += toNumber(row[COL.shipFee]);
    doc.commission += toNumber(row[COL.commission]);
    (bills[id] ||= new Set()).add(row[COL.bill]);
    const srcKey = keyOf(row[COL.source] || 'Không rõ');
    (((sourceBills[id] ||= {})[srcKey] ||= new Set())).add(row[COL.bill]);

    const code = row[COL.itemCode] || row[COL.itemName];
    const item = (doc.items[keyOf(code)] ||= {
      code: row[COL.itemCode] || null,
      name: row[COL.itemName] || null,
      group: row[COL.group] || null,
      type: row[COL.type] || null,
      quantity: 0,
      revenue: 0,
    });
    item.quantity += qty;
    item.revenue += rev;

    addBreakdown(doc.sources, row[COL.source], qty, rev);
    addBreakdown(doc.payments, row[COL.payment], qty, rev);
    addBreakdown(doc.areas, row[COL.area], qty, rev);
  }
  for (const id in docs) {
    docs[id].bills = bills[id].size;
    for (const [k, set] of Object.entries(sourceBills[id] || {})) {
      if (docs[id].sources[k]) docs[id].sources[k].bills = set.size;
    }
  }
  return docs;
}

// Gộp 1 chiều theo tên (nguồn đơn / PTTT / khu vực) -> { [key]: { name, quantity, revenue } }.
// sources{} còn có thêm .bills (số hoá đơn riêng theo nguồn, gán ở aggregate()) để tính AOV theo nguồn.
function addBreakdown(map, name, qty, rev) {
  name = String(name || 'Không rõ').trim() || 'Không rõ';
  const entry = (map[keyOf(name)] ||= { name, quantity: 0, revenue: 0 });
  entry.quantity += qty;
  entry.revenue += rev;
}

function sortedEntries(map) {
  return Object.keys(map).sort().map((k) => [k, map[k]]);
}

function hashOf(doc) {
  const stable = JSON.stringify([
    doc.date, doc.storeId, doc.storeName, doc.revenue, doc.quantity, doc.bills,
    doc.discount, doc.serviceFee, doc.tax, doc.shipFee, doc.commission,
    sortedEntries(doc.items), sortedEntries(doc.sources), sortedEntries(doc.payments), sortedEntries(doc.areas),
  ]);
  return crypto.createHash('sha1').update(stable).digest('hex').slice(0, 16);
}

// Chỉ ghi các doc có hash khác với lần ghi trước. Trả về số doc đã ghi.
async function writeChanged(db, docs) {
  const byDate = {};
  for (const [id, doc] of Object.entries(docs)) (byDate[doc.date] ||= {})[id] = doc;

  let written = 0;
  let skipped = 0;
  const changedDocs = [];
  const locked = Object.keys(byDate).filter((d) => d <= EXCEL_UNTIL);
  if (locked.length && !process.argv.includes('--force')) {
    throw new Error(`Ngày ${locked.join(', ')} nằm trong khoảng lấy số từ Excel (tới ${EXCEL_UNTIL}) - không nạp Fabi đè lên. `
      + 'Thêm --force nếu chỉ muốn lưu món/PTTT gốc vào d05Daily (d05Days vẫn giữ số Excel).');
  }
  for (const [date, dayDocs] of Object.entries(byDate)) {
    const stateRef = db.collection(STATE_COLLECTION).doc(`d05_${date}`);
    const oldHashes = (await stateRef.get()).data()?.hashes || {};
    const newHashes = { ...oldHashes };
    const changed = [];
    for (const [id, doc] of Object.entries(dayDocs)) {
      const h = hashOf(doc);
      newHashes[id] = h;
      if (oldHashes[id] === h) skipped++;
      else changed.push([id, { ...doc, hash: h }]);
    }
    if (!changed.length) continue;

    // Tóm tắt ngày cần đủ mọi cửa hàng của ngày đó. File xuất "tất cả cửa hàng" thì
    // dayDocs đã đủ; lỡ chỉ nạp 1 phần (file vài cửa hàng) thì đọc bù phần còn lại.
    let allDay = dayDocs;
    const missing = Object.keys(oldHashes).filter((id) => !dayDocs[id]);
    if (missing.length) {
      const snap = await db.collection(COLLECTION).where('date', '==', date).get();
      allDay = {};
      snap.forEach((d) => (allDay[d.id] = d.data()));
      Object.assign(allDay, dayDocs);
    }
    const summary = daySummary(date, allDay);

    // Chia batch 100 doc (~1 MB) - batch lớn hay bị DEADLINE_EXCEEDED khi mạng chậm.
    for (let i = 0; i < changed.length; i += 100) {
      const batch = db.batch();
      for (const [id, doc] of changed.slice(i, i + 100)) {
        batch.set(db.collection(COLLECTION).doc(id), {
          ...doc,
          updatedAt: admin.firestore.FieldValue.serverTimestamp(),
        });
      }
      if (i + 100 >= changed.length) {
        batch.set(stateRef, { hashes: newHashes });
        if (date > EXCEL_UNTIL) batch.set(db.collection(DAYS_COLLECTION).doc(date), {
          ...summary,
          updatedAt: admin.firestore.FieldValue.serverTimestamp(),
        });
      }
      await batch.commit();
    }
    written += changed.length;
    for (const [, doc] of changed) changedDocs.push(doc);
  }

  // Danh mục cửa hàng dùng cho trang "Chi nhánh" (gán Region/City/AM). Merge để không
  // đụng field admin đã gán trên web - script chỉ giữ storeId/storeName luôn đủ và đúng.
  // Chỉ các cửa hàng vừa có doc thay đổi: cửa hàng mới / đổi tên luôn làm hash đổi nên
  // vẫn được cập nhật, còn lần chạy không có gì mới thì không tốn thêm lượt ghi nào.
  const stores = new Map();
  for (const doc of changedDocs) stores.set(doc.storeId, doc.storeName);
  if (stores.size) {
    const catalogBatch = db.batch();
    for (const [storeId, storeName] of stores) {
      catalogBatch.set(db.collection(STORES_COLLECTION).doc(storeId), {
        storeId, storeName, updatedAt: admin.firestore.FieldValue.serverTimestamp(),
      }, { merge: true });
    }
    await catalogBatch.commit();
  }

  return { written, skipped };
}

// "ET" = tại chỗ + mang về, "D" = các nguồn còn lại (giao hàng/app). Tách sẵn theo từng
// cửa hàng để trang Summary tính được AOV-ET / AOV-D khi lọc theo vùng, mà không phải
// đọc d05Daily (nặng gấp ~90 lần).
const ET_KEYS = ['tai-cho', 'mang-ve'];
function etSplit(sources) {
  const et = { revenue: 0, bills: 0 };
  const dlv = { revenue: 0, bills: 0 };
  for (const [k, v] of Object.entries(sources || {})) {
    const bucket = ET_KEYS.includes(k) ? et : dlv;
    bucket.revenue += v.revenue || 0;
    bucket.bills += v.bills || 0;
  }
  return { et, dlv };
}

// Doc cửa hàng (số gốc Fabi) -> số trước VAT như báo cáo nội bộ. Hoa hồng chỉ có tổng
// theo cửa hàng nên chia cho các nguồn ShopeeFood/BeFood/Xanh Ngon theo tỉ lệ doanh thu;
// PTTT và khu vực thì nhân cùng tỉ lệ net/gốc của cửa hàng. Ngày nạp bằng bản script cũ
// chưa có hoa hồng -> chỉ chia 1,08 (thấp hơn thực tế một chút), đánh dấu noCommission.
function toNet(d) {
  const gross = d.revenue;
  const com = d.commission || 0;
  const net = (gross + com) / VAT;
  const f = gross ? net / gross : 1 / VAT;
  const src = d.sources || {};
  const platform = COMMISSION_SOURCES.reduce((a, k) => a + ((src[k] && src[k].revenue) || 0), 0);
  const sources = {};
  for (const [k, v] of Object.entries(src)) {
    const share = COMMISSION_SOURCES.includes(k) && platform ? com * v.revenue / platform : 0;
    sources[k] = { ...v, revenue: Math.round((v.revenue + share) / VAT) };
  }
  const scale = (map) => Object.fromEntries(Object.entries(map || {}).map(([k, v]) => [k, { ...v, revenue: Math.round(v.revenue * f) }]));
  return { ...d, revenue: Math.round(net), revenueGross: gross, sources, payments: scale(d.payments), areas: scale(d.areas) };
}

// Gộp các doc cửa hàng của 1 ngày thành doc tóm tắt d05Days/{date}.
function daySummary(date, dayDocs) {
  const s = {
    date, basis: 'net', revenueGross: 0, noCommission: 0, revenue: 0, bills: 0, quantity: 0, storeCount: 0,
    discount: 0, serviceFee: 0, tax: 0, shipFee: 0, commission: 0,
    stores: {}, items: {}, sources: {}, payments: {}, areas: {},
  };
  for (const raw of Object.values(dayDocs)) {
    const d = toNet(raw);
    if (raw.commission == null) s.noCommission++;
    s.revenueGross += d.revenueGross;
    s.revenue += d.revenue;
    s.bills += d.bills;
    s.quantity += d.quantity;
    s.discount += d.discount || 0;
    s.serviceFee += d.serviceFee || 0;
    s.tax += d.tax || 0;
    s.shipFee += d.shipFee || 0;
    s.commission += d.commission || 0;
    s.storeCount++;
    s.stores[d.storeId] = {
      name: d.storeName, revenue: d.revenue, bills: d.bills, quantity: d.quantity, commission: d.commission || 0,
      ...etSplit(d.sources),
    };
    for (const [k, it] of Object.entries(d.items || {})) {
      // Doc nạp bằng bản script cũ không có nhóm/loại món -> null (Firestore không nhận undefined)
      const x = (s.items[k] ||= { name: it.name, group: it.group || null, type: it.type || null, quantity: 0, revenue: 0 });
      x.quantity += it.quantity;
      x.revenue += it.revenue;
    }
    for (const dim of ['sources', 'payments', 'areas']) {
      for (const [k, v] of Object.entries(d[dim] || {})) {
        const x = (s[dim][k] ||= { name: v.name, quantity: 0, revenue: 0, bills: 0 });
        x.quantity += v.quantity;
        x.revenue += v.revenue;
        x.bills += v.bills || 0;
      }
    }
  }
  return s;
}

// Dựng lại d05Days cho các ngày đã có trong d05Daily (dùng 1 lần khi mới thêm d05Days,
// hoặc khi nghi ngờ tóm tắt lệch). Mỗi ngày: đọc ~90 doc, ghi 1 doc.
async function rebuildDaySummaries(db, dates) {
  for (const date of dates) {
    if (date <= EXCEL_UNTIL) { console.log(`${date}: bỏ qua - ngày lấy số từ Excel`); continue; }
    const snap = await db.collection(COLLECTION).where('date', '==', date).get();
    if (snap.empty) continue;
    const dayDocs = {};
    snap.forEach((d) => (dayDocs[d.id] = d.data()));
    await db.collection(DAYS_COLLECTION).doc(date).set({
      ...daySummary(date, dayDocs),
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    console.log(`${date}: ${snap.size} cửa hàng`);
  }
}

function summarize(docs) {
  const list = Object.values(docs);
  const dates = [...new Set(list.map((d) => d.date))].sort();
  const revenue = list.reduce((s, d) => s + d.revenue, 0);
  return `${list.length} cửa hàng-ngày, ngày ${dates[0]} → ${dates[dates.length - 1]}, tổng tiền ${revenue.toLocaleString('vi-VN')} ₫`;
}

// { [tên cửa hàng không dấu]: mã Fabi } từ danh mục stores - dùng khi nạp file không có mã.
async function loadStoreIds(db) {
  const snap = await db.collection(STORES_COLLECTION).get();
  const map = {};
  snap.forEach((d) => { if (d.data().storeName) map[keyOf(d.data().storeName)] = d.id; });
  return map;
}

module.exports = { EXCEL_UNTIL, ET_KEYS, toNet, daySummary, loadStoreIds, ALIASES, loadEnv, initFirebase, readRows, aggregate, writeChanged, summarize, rebuildDaySummaries, COL };
