/**
 * Nạp tay file Excel bán hàng (D05 tải từ Fabi, hoặc file Excel tương tự) lên Firestore.
 * Tự tìm dòng tiêu đề và tự nhận cột theo tên (xem ALIASES trong lib.js), rồi gộp và
 * "chỉ ghi phần thay đổi" giống sync-report.js - nạp lại cùng 1 file không tốn thêm lượt ghi.
 *
 *   node import-file.js file1.xlsx file2.xlsx ...
 *   node import-file.js thu-muc-chua-file/          (nạp mọi .xlsx/.xls/.csv trong thư mục)
 *   node import-file.js --dry-run file.xlsx         (chỉ đọc + in thống kê, không ghi)
 *
 * File rất lớn (cả tháng, hàng triệu dòng) có thể cần thêm RAM:
 *   node --max-old-space-size=8192 import-file.js ...
 */

const fs = require('fs');
const path = require('path');
const { loadEnv, initFirebase, readRows, aggregate, writeChanged, summarize, loadStoreIds, COL } = require('./lib');

async function main() {
  loadEnv();
  const dryRun = process.argv.includes('--dry-run');
  const files = process.argv
    .slice(2)
    .filter((a) => !a.startsWith('--'))
    .flatMap((p) =>
      fs.statSync(p).isDirectory()
        ? fs.readdirSync(p).filter((f) => /\.(xlsx|xls|csv)$/i.test(f) && !f.startsWith('~$')).sort().map((f) => path.join(p, f))
        : [p]
    );
  if (!files.length) throw new Error('Cách dùng: node import-file.js <file.xlsx | thư mục> ...');

  // Có Firestore thì tra được mã cửa hàng theo tên; --dry-run không có thì vẫn đọc được.
  let db = null;
  let storeIds = {};
  try {
    db = initFirebase();
    storeIds = await loadStoreIds(db);
  } catch (e) {
    if (!dryRun) throw e;
  }

  const failed = [];
  for (const file of files) {
    const name = path.basename(file);
    try {
      const rows = readRows(file);
      const { info } = rows;
      if (!info.sheets) {
        throw new Error(`không thấy sheet nào đủ cột Ngày / Hoá đơn / Số lượng / Tổng tiền (${info.skippedSheets.join(', ')})`);
      }
      const docs = aggregate(rows, storeIds);
      console.log(`${name}: ${summarize(docs)}`);
      console.log(`  đọc ${info.sheets} sheet · nhận cột: ${info.found.map((f) => COL[f]).join(', ')}`);
      if (info.missing.length) console.log(`  thiếu cột (để 0 / "Không rõ"): ${info.missing.map((f) => COL[f]).join(', ')}`);
      const unknown = new Set(Object.values(docs).filter((d) => !/^\d+$/.test(d.storeId)).map((d) => d.storeName));
      if (unknown.size) console.log(`  ⚠ ${unknown.size} cửa hàng chưa có mã Fabi (tra theo tên không ra): ${[...unknown].slice(0, 5).join(', ')}${unknown.size > 5 ? '…' : ''}`);
      if (dryRun) continue;
      const { written, skipped } = await writeChanged(db, docs);
      console.log(`  ✅ ghi ${written} doc, bỏ qua ${skipped} doc không đổi`);
    } catch (e) {
      console.error(`❌ ${name}: ${e.message}`);
      failed.push(name);
    }
  }
  if (failed.length) throw new Error(`${failed.length} file lỗi: ${failed.join(', ')}`);
}

main().catch((err) => {
  console.error('❌ Lỗi:', err.message);
  process.exit(1);
});
