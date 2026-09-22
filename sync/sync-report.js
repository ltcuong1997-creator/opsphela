/**
 * Đồng bộ tự động D05 - Báo cáo bán hàng của NGÀY HÔM NAY từ Fabi POS CMS
 * lên Firestore (GitHub Actions chạy mỗi giờ). Dữ liệu các ngày cũ nạp tay
 * bằng import-file.js.
 *
 * Luồng: đăng nhập Fabi -> mở D05 (mặc định đang lọc ngày hôm nay) -> Xuất
 * báo cáo tất cả cửa hàng -> chờ hộp "Tiến trình xuất báo cáo" xong -> tải file ->
 * gộp ngày × cửa hàng × món -> chỉ ghi những cửa hàng có số liệu thay đổi.
 *
 *   node sync-report.js            chạy thật
 *   node sync-report.js --dry-run  tải file và in thống kê, KHÔNG ghi Firestore
 *
 * BIẾN MÔI TRƯỜNG (local: file sync/.env; GitHub: Secrets) - KHÔNG ghi thẳng
 * vào code:
 *  - FABI_EMAIL, FABI_PASSWORD
 *  - FIREBASE_SERVICE_ACCOUNT (chuỗi JSON) hoặc FIREBASE_SERVICE_ACCOUNT_FILE (đường dẫn)
 */

const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const { loadEnv, initFirebase, readRows, aggregate, writeChanged, summarize } = require('./lib');

const FABI_BASE_URL = 'https://fabi.ipos.vn';
const REPORT_URL = `${FABI_BASE_URL}/report/accounting/sale-audit`; // D05
const DOWNLOAD_DIR = path.join(__dirname, 'downloads');
const MAX_WAIT_MS = 15 * 60 * 1000;
const POLL_INTERVAL_MS = 5000;
const DRY_RUN = process.argv.includes('--dry-run');

async function main() {
  loadEnv();
  assertEnv(['FABI_EMAIL', 'FABI_PASSWORD']);
  fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });
  const db = DRY_RUN ? null : initFirebase();

  const browser = await chromium.launch({ headless: !process.argv.includes('--headed') });
  // Bộ lọc ngày mặc định của Fabi tính theo giờ trình duyệt; máy GitHub chạy giờ
  // UTC nên phải ép về giờ VN, không thì 0h-7h sáng sẽ xuất nhầm ngày hôm qua.
  const context = await browser.newContext({ acceptDownloads: true, timezoneId: 'Asia/Ho_Chi_Minh' });
  const page = await context.newPage();

  try {
    await login(page);
    await goToReport(page);
    const filePath = await exportAndDownload(page);
    console.log(`Đã tải: ${path.basename(filePath)}`);

    const docs = aggregate(readRows(filePath));
    console.log(summarize(docs));
    if (DRY_RUN) return;

    const { written, skipped } = await writeChanged(db, docs);
    console.log(`✅ Ghi ${written} doc thay đổi, bỏ qua ${skipped} doc không đổi.`);
    fs.unlinkSync(filePath);
  } finally {
    await browser.close();
  }
}

function assertEnv(names) {
  const missing = names.filter((n) => !process.env[n]);
  if (missing.length) throw new Error(`Thiếu biến môi trường: ${missing.join(', ')}`);
}

async function login(page) {
  await page.goto(`${FABI_BASE_URL}/login`, { waitUntil: 'networkidle' });
  await page.fill('input[name="email_input"]', process.env.FABI_EMAIL);
  await page.fill('input[type="password"]', process.env.FABI_PASSWORD);
  await page.click('button[type="submit"]');
  // Fabi là SPA: sau đăng nhập chỉ đổi route, trang đích không bao giờ "load" xong
  // hẳn, nên chỉ chờ URL rời khỏi /login (waitUntil: 'commit').
  await page
    .waitForURL((url) => !url.pathname.startsWith('/login'), { timeout: 45000, waitUntil: 'commit' })
    .catch(async () => {
      const msg = await page.locator('form .invalid-feedback, form .alert, .toast').allInnerTexts().catch(() => []);
      throw new Error(`Đăng nhập Fabi thất bại (đang ở ${page.url()}). ${msg.join(' ').trim()}`);
    });
}

async function goToReport(page) {
  // Trang D05 gọi mạng liên tục nên không bao giờ "networkidle" - chờ nút xuất.
  await page.goto(REPORT_URL, { waitUntil: 'domcontentloaded' });
  await page.getByRole('button', { name: /Xuất báo cáo/ }).waitFor({ timeout: 60000 });
  await page.waitForTimeout(2000);
  const closeNotice = page.getByRole('button', { name: 'Đóng' });
  if (await closeNotice.isVisible().catch(() => false)) {
    await closeNotice.click();
  }
}

// Fabi dựng sẵn mọi hộp thoại (.kit-modal__container), cái nào cũng mang class
// "show" - chỉ ẩn/hiện bằng display. Vì vậy luôn lọc theo tiêu đề + :visible.
function dialog(page, title) {
  return page.locator('.kit-modal__container:visible').filter({ hasText: title });
}

// Bấm "Xuất báo cáo tất cả cửa hàng" -> hộp "Tiến trình xuất báo cáo" hiện %
// (92 cửa hàng mất ~4 phút) -> xong thì chính hộp đó có nút "Tải file"
// (.ep__download). Hộp này chỉ theo dõi lần xuất vừa bấm, nên không lẫn với file
// người khác trong "Lịch sử xuất".
async function exportAndDownload(page) {
  await page.getByRole('button', { name: /Xuất báo cáo/ }).click();
  await page.getByText('Xuất báo cáo tất cả cửa hàng', { exact: true }).click();

  const progress = dialog(page, 'Tiến trình xuất báo cáo');
  const already = dialog(page, 'Dữ liệu này đã được xuất trước đó');
  await progress.or(already).first().waitFor({ timeout: 60000 });
  if (await already.isVisible()) {
    // Chưa gặp với dữ liệu hôm nay; nếu gặp thì chọn xuất mới, không lấy file cũ.
    const again = already.getByRole('button', { name: /Xuất lại|Xuất mới|Tiếp tục xuất/ });
    if (!(await again.count())) {
      throw new Error(`Gặp hộp "đã xuất trước đó" chưa xử lý: ${(await already.innerText()).replace(/\s+/g, ' ')}`);
    }
    await again.first().click();
    await progress.waitFor({ timeout: 60000 });
  }

  const downloadBtn = progress.locator('button.ep__download');
  const start = Date.now();
  while (!(await downloadBtn.isVisible())) {
    const text = (await progress.innerText().catch(() => '')).replace(/\s+/g, ' ');
    if (/Đã huỷ|Đã hủy|Lỗi|Thất bại/i.test(text)) throw new Error(`Job xuất bị lỗi/huỷ: ${text.slice(0, 200)}`);
    if (Date.now() - start > MAX_WAIT_MS) {
      throw new Error('Quá thời gian chờ xuất báo cáo - kiểm tra "Lịch sử xuất" trên Fabi.');
    }
    await page.waitForTimeout(POLL_INTERVAL_MS);
  }

  const [download] = await Promise.all([page.waitForEvent('download'), downloadBtn.click()]);
  const downloadPath = path.join(DOWNLOAD_DIR, download.suggestedFilename());
  await download.saveAs(downloadPath);
  return downloadPath;
}

main().catch((err) => {
  console.error('❌ Lỗi:', err);
  process.exit(1);
});
