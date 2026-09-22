/**
 * Nạp tay các file Excel D05 đã tải từ Fabi (dữ liệu ngày cũ) lên Firestore.
 * Dùng chung cách gộp và cơ chế "chỉ ghi phần thay đổi" với sync-report.js,
 * nên nạp lại cùng 1 file nhiều lần cũng không tốn thêm lượt ghi.
 *
 *   node import-file.js file1.xlsx file2.xlsx ...
 *   node import-file.js thu-muc-chua-file/          (nạp mọi .xlsx trong thư mục)
 *   node import-file.js --dry-run file.xlsx         (chỉ in thống kê, không ghi)
 */

const fs = require('fs');
const path = require('path');
const { loadEnv, initFirebase, readRows, aggregate, writeChanged, summarize } = require('./lib');

async function main() {
  loadEnv();
  const dryRun = process.argv.includes('--dry-run');
  const files = process.argv
    .slice(2)
    .filter((a) => !a.startsWith('--'))
    .flatMap((p) =>
      fs.statSync(p).isDirectory()
        ? fs.readdirSync(p).filter((f) => /\.xlsx?$/i.test(f)).map((f) => path.join(p, f))
        : [p]
    );
  if (!files.length) throw new Error('Cách dùng: node import-file.js <file.xlsx | thư mục> ...');

  const db = dryRun ? null : initFirebase();
  for (const file of files) {
    const docs = aggregate(readRows(file));
    console.log(`${path.basename(file)}: ${summarize(docs)}`);
    if (dryRun) continue;
    const { written, skipped } = await writeChanged(db, docs);
    console.log(`  ✅ ghi ${written} doc, bỏ qua ${skipped} doc không đổi`);
  }
}

main().catch((err) => {
  console.error('❌ Lỗi:', err.message);
  process.exit(1);
});
