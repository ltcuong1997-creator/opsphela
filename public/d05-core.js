/**
 * Lõi đọc & tổng hợp file D05 của Fabi - DÙNG CHUNG cho script Node (sync/lib.js) và nút
 * "Nạp file D05" trên web (tab Data), để hai đường nạp ra đúng cùng một con số.
 * Chỉ chứa hàm thuần: không Firestore, không đọc file. Node: require('../public/d05-core.js');
 * trình duyệt: <script src="/d05-core.js"> -> window.D05.
 */
(function (root) {
'use strict';

// Tới hết ngày này: số trên d05Days lấy từ báo cáo nội bộ Excel (histDays), Fabi không
// ghi đè. Từ 01/2026 trở đi dữ liệu nạp 100% từ Fabi nên mốc dừng ở cuối 2025.
const EXCEL_UNTIL = '2025-12-31';
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
function rowsFromWorkbook(workbook, XLSX) {
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

// Giờ bán "07" từ cột Ngày ("02/09/2026 07:31:40" - Fabi xuất dạng chữ); null nếu không có giờ.
function hourOf(v) {
  if (v instanceof Date) return isNaN(v) ? null : String(v.getHours()).padStart(2, '0');
  const m = String(v || '').match(/\s(\d{1,2}):\d{2}/);
  return m ? m[1].padStart(2, '0') : null;
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
  const hourBills = {}; // id -> { '07': Set(billId) } - TC theo giờ
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
      items: {}, sources: {}, payments: {}, areas: {}, hours: {},
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
    // Doanh thu theo giờ: số gốc Fabi + hoa hồng riêng từng dòng, toNet() quy về trước VAT
    const hr = hourOf(row[COL.time]);
    if (hr) {
      const x = (doc.hours[hr] ||= { revenue: 0, commission: 0, bills: 0 });
      x.revenue += rev;
      x.commission += toNumber(row[COL.commission]);
      (((hourBills[id] ||= {})[hr] ||= new Set())).add(row[COL.bill]);
    }

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
    for (const [hr, set] of Object.entries(hourBills[id] || {})) docs[id].hours[hr].bills = set.size;
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

function hashInput(doc) {
  return JSON.stringify([
    doc.date, doc.storeId, doc.storeName, doc.revenue, doc.quantity, doc.bills,
    doc.discount, doc.serviceFee, doc.tax, doc.shipFee, doc.commission,
    sortedEntries(doc.items), sortedEntries(doc.sources), sortedEntries(doc.payments), sortedEntries(doc.areas),
    sortedEntries(doc.hours || {}),
  ]);
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
  const hours = {};
  for (const [k, v] of Object.entries(d.hours || {})) hours[k] = { bills: v.bills || 0, revenue: Math.round(((v.revenue || 0) + (v.commission || 0)) / VAT) };
  return { ...d, revenue: Math.round(net), revenueGross: gross, sources, payments: scale(d.payments), areas: scale(d.areas), hours };
}

// Số bán từng món theo từng cửa hàng trong 1 ngày -> d05ItemStores/{date}.
// { date, items, stores, payload } - payload là CHUỖI JSON:
//   { [mã món]: { name, group, stores: { [mã CH]: [số lượng, doanh thu] } } }
// Để chuỗi thay vì map lồng vì Firestore đánh chỉ mục từng khoá của map: ~170 món × ~80
// cửa hàng vượt giới hạn chỉ mục của 1 document (lỗi "too many index entries").
// Doanh thu món là số gốc Fabi (gồm VAT), giống items trong d05Days.
function itemStoresDoc(date, dayDocs) {
  const items = {};
  const stores = new Set();
  for (const d of Object.values(dayDocs)) {
    for (const [k, it] of Object.entries(d.items || {})) {
      const x = (items[k] ||= { name: it.name, group: it.group || null, stores: {} });
      const st = (x.stores[d.storeId] ||= [0, 0]);
      st[0] += it.quantity || 0;
      st[1] += it.revenue || 0;
      stores.add(d.storeId);
    }
  }
  return { date, items: Object.keys(items).length, stores: stores.size, payload: JSON.stringify(items) };
}

// Gộp các doc cửa hàng của 1 ngày thành doc tóm tắt d05Days/{date}.
function daySummary(date, dayDocs) {
  const s = {
    date, basis: 'net', revenueGross: 0, noCommission: 0, revenue: 0, bills: 0, quantity: 0, storeCount: 0,
    discount: 0, serviceFee: 0, tax: 0, shipFee: 0, commission: 0,
    stores: {}, items: {}, sources: {}, payments: {}, areas: {}, hours: {}, hourBills: {},
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
    // Theo giờ: hours = doanh thu (trước VAT), hourBills = TC. Ngày nạp trước khi có giờ thì bỏ trống.
    const hrs = Object.entries(d.hours || {});
    if (hrs.length) {
      const st = s.stores[d.storeId];
      st.hours = {}; st.hourBills = {};
      for (const [k, v] of hrs) {
        st.hours[k] = v.revenue; st.hourBills[k] = v.bills;
        s.hours[k] = (s.hours[k] || 0) + v.revenue;
        s.hourBills[k] = (s.hourBills[k] || 0) + v.bills;
      }
    }
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


const api = {
  EXCEL_UNTIL, VAT, COMMISSION_SOURCES, COL, ALIASES, REQUIRED, ET_KEYS,
  detectColumns, rowsFromWorkbook, toNumber, normalizeDate, hourOf, keyOf, aggregate, addBreakdown,
  sortedEntries, hashInput, etSplit, toNet, itemStoresDoc, daySummary,
};
if (typeof module !== 'undefined' && module.exports) module.exports = api;
else root.D05 = api;
})(typeof self !== 'undefined' ? self : this);
