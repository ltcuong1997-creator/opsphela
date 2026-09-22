/**
 * Dựng d05Days - bộ số mà mọi trang web đang đọc - cho các ngày tới hết mốc EXCEL_UNTIL
 * (lib.js) từ báo cáo nội bộ Excel (histDays, nạp bằng import-history.js). Sau mốc thì
 * Fabi tự nạp.
 *
 * Ghép theo từng cửa hàng:
 *  - Cửa hàng có trong Excel: doanh thu (trước VAT), TC (-> bills), kênh bán lấy từ Excel.
 *  - Cửa hàng Excel không có mà Fabi có (trước 03/2026 Excel thiếu nhóm Hà Nội; kênh TMDT,
 *    PĐT không bao giờ có trong Excel): lấy số Fabi quy về trước VAT (toNet trong lib.js).
 *  - Món, PTTT, khu vực, số món, giảm giá/phí/hoa hồng: chỉ Fabi có -> lấy của Fabi.
 * Phần Fabi đọc từ d05Daily (số gốc, không bao giờ bị ghi đè) nên chạy lại cho cùng kết quả.
 *
 *   node hist-to-days.js --dry-run     so sánh, không ghi
 *   node hist-to-days.js
 */

const admin = require('firebase-admin');
const { loadEnv, initFirebase, EXCEL_UNTIL, ET_KEYS, toNet, daySummary } = require('./lib');

const FEES = ['discount', 'serviceFee', 'tax', 'shipFee', 'commission'];

// h: doc histDays; fabiDocs: { [docId]: doc d05Daily } của ngày đó (có thể rỗng)
function buildDay(h, fabiDocs) {
  const fabiList = Object.values(fabiDocs);
  const fabi = fabiList.length ? daySummary(h.date, fabiDocs) : null;
  const day = {
    date: h.date, basis: 'net', source: fabi ? 'excel+fabi' : 'excel',
    revenue: 0, bills: 0, quantity: fabi ? fabi.quantity : 0, storeCount: 0,
    discount: 0, serviceFee: 0, tax: 0, shipFee: 0, commission: 0,
    stores: {}, sources: {}, hours: {},
    items: fabi ? fabi.items : {}, payments: fabi ? fabi.payments : {}, areas: fabi ? fabi.areas : {},
    fromFabi: [],
  };
  if (fabi) for (const k of FEES) day[k] = fabi[k];
  const addSource = (k, v, bills) => {
    const x = (day.sources[k] ||= { name: v.name, quantity: 0, revenue: 0, bills: 0 });
    x.revenue += v.revenue; x.bills += bills; x.quantity += v.quantity || 0;
  };

  for (const [id, st] of Object.entries(h.stores)) {
    const fst = (fabi && fabi.stores[id]) || {};
    const et = { revenue: 0, bills: 0 }, dlv = { revenue: 0, bills: 0 };
    for (const [k, v] of Object.entries(st.sources || {})) {
      const b = ET_KEYS.includes(k) ? et : dlv;
      b.revenue += v.revenue; b.bills += v.tc;
      addSource(k, v, v.tc);
    }
    for (const [hr, v] of Object.entries(st.hours || {})) day.hours[hr] = (day.hours[hr] || 0) + v;
    day.stores[id] = { name: st.name, revenue: st.revenue, bills: st.tc, quantity: fst.quantity || 0, commission: fst.commission || 0, et, dlv };
  }
  for (const raw of fabiList) {
    if (day.stores[raw.storeId]) continue;
    const d = toNet(raw);
    day.stores[raw.storeId] = fabi.stores[raw.storeId];
    for (const [k, v] of Object.entries(d.sources || {})) addSource(k, v, v.bills || 0);
    day.fromFabi.push(raw.storeName);
  }
  for (const st of Object.values(day.stores)) { day.revenue += st.revenue; day.bills += st.bills; day.storeCount++; }
  return day;
}

async function main() {
  loadEnv();
  const dryRun = process.argv.includes('--dry-run');
  const db = initFirebase();
  const hist = (await db.collection('histDays').where(admin.firestore.FieldPath.documentId(), '<=', EXCEL_UNTIL).get()).docs.map((d) => d.data());
  const old = {};
  (await db.collection('d05Days').where('date', '<=', EXCEL_UNTIL).select('revenue', 'storeCount', 'source').get()).forEach((d) => { old[d.id] = d.data(); });
  // Ngày nào Fabi từng nạp: syncState/d05_{date} (1 doc/ngày) - rồi mới đọc d05Daily của những ngày đó
  const fabiDates = (await db.collection('syncState').get()).docs.map((d) => d.id.replace(/^d05_/, '')).filter((d) => d <= EXCEL_UNTIL);
  const fabiByDate = {};
  for (const date of fabiDates) {
    const snap = await db.collection('d05Daily').where('date', '==', date).get();
    fabiByDate[date] = {};
    snap.forEach((d) => { fabiByDate[date][d.id] = d.data(); });
  }

  const out = hist.map((h) => buildDay(h, fabiByDate[h.date] || {}));
  const merged = out.filter((d) => d.source === 'excel+fabi');
  console.log(`Tới ${EXCEL_UNTIL}: ${out.length} ngày (${out[0]?.date} → ${out[out.length - 1]?.date}), ${merged.length} ngày ghép thêm Fabi`);
  for (const d of merged) {
    const o = old[d.date];
    console.log(`  ${d.date}: ${d.storeCount} CH (Excel ${d.storeCount - d.fromFabi.length} + Fabi ${d.fromFabi.length}), DT ${d.revenue.toLocaleString('vi-VN')}`
      + ` | đang có ${o ? Math.round(o.revenue).toLocaleString('vi-VN') + ' (' + o.storeCount + ' CH)' : '—'}`);
  }
  const fabiOnly = [...new Set(merged.flatMap((d) => d.fromFabi))];
  console.log(`  Cửa hàng lấy từ Fabi (${fabiOnly.length}): ${fabiOnly.join(' | ')}`);
  const extra = Object.keys(old).filter((k) => !hist.some((h) => h.date === k));
  if (extra.length) console.log(`  ⚠ ${extra.length} ngày có trên d05Days mà Excel không có (giữ nguyên): ${extra.join(', ')}`);
  if (dryRun) { console.log('(chạy thử, chưa ghi gì)'); return; }

  // Doc ngày có kèm danh sách món khá nặng: 5 doc/batch, lỗi mạng (DEADLINE_EXCEEDED) thì thử lại
  for (let i = 0; i < out.length; i += 5) {
    for (let attempt = 1; ; attempt++) {
      try {
        const batch = db.batch();
        for (const d of out.slice(i, i + 5)) batch.set(db.collection('d05Days').doc(d.date), { ...d, updatedAt: admin.firestore.FieldValue.serverTimestamp() });
        await batch.commit();
        break;
      } catch (e) {
        if (attempt >= 4) throw e;
        console.log(`  thử lại ${out[i].date}… (${e.message.slice(0, 40)})`);
      }
    }
  }
  console.log(`  ✅ đã ghi ${out.length} ngày vào d05Days`);
}

main().then(() => process.exit(0)).catch((e) => { console.error('❌ Lỗi:', e.message); process.exit(1); });
