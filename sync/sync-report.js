/**
 * Đồng bộ tự động D05 - Báo cáo bán hàng của NGÀY HÔM NAY từ Fabi POS CMS
 * lên Firestore (GitHub Actions chạy mỗi giờ). Dữ liệu các ngày cũ nạp tay
 * bằng import-file.js.
 *
 * Luồng: đăng nhập Fabi -> mở D05 (mặc định đang lọc ngày hôm nay) -> Xuất
 * báo cáo tất cả cửa hàng -> chờ hộp "Tiến trình xuất báo cáo" xong -> tải file ->
 * gộp ngày × cửa hàng × món -> chỉ ghi những cửa hàng có số liệu thay đổi.
 *
 *   node sync-report.js            chạy thật (ngày hôm nay)
 *   node sync-report.js --dry-run  tải file và in thống kê, KHÔNG ghi Firestore
 *   node sync-report.js --from 2026-09-01 --to 2026-09-21
 *                                  nạp bù ngày cũ, mỗi ngày 1 lần xuất (~4-5 phút/ngày)
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
const arg = (name) => { const i = process.argv.indexOf(name); return i > 0 ? process.argv[i + 1] : null; };
// Repo công khai: log trên GitHub Actions ai cũng đọc được -> không in doanh thu.
const PUBLIC_LOG = !!process.env.GITHUB_ACTIONS;

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

  // Không truyền --from/--to: ngày hôm nay (giữ bộ lọc mặc định của Fabi). Riêng lần
  // chạy trước 7h sáng (lịch 6h15) thì chốt lại TRỌN ngày hôm qua - lần chạy cuối
  // 23h15 chưa có doanh thu bán sau giờ đó. Chỉ ghi phần thay đổi nên không tốn thêm.
  let days = arg('--from') ? dayList(arg('--from'), arg('--to') || arg('--from')) : [vnHour() < 7 ? vnYesterday() : null];
  // --skip-done: bỏ các ngày đã có syncState (đã nạp xong) - chạy lại sau khi bị ngắt giữa chừng
  if (db && days[0] && process.argv.includes('--skip-done')) {
    const done = await Promise.all(days.map((d) => db.collection('syncState').doc(`d05_${d}`).get()));
    days = days.filter((d, i) => !done[i].exists);
    console.log(`Cần nạp ${days.length} ngày: ${days.join(', ') || '(không còn ngày nào)'}`);
    if (!days.length) return;
  }
  // --skip-complete: bỏ ngày đã có ĐỦ các trường mới (nguồn đơn kèm số hoá đơn, hoa hồng...).
  // Khác --skip-done ở chỗ ngày nạp bằng bản script cũ vẫn được nạp lại. Dùng để chạy tiếp
  // một lượt nạp bù dài bị ngắt giữa chừng mà không phải tải lại từ đầu (mỗi ngày ~10 phút).
  if (db && days[0] && process.argv.includes('--skip-complete')) {
    const probes = await Promise.all(days.map((d) => db.collection('d05Daily').where('date', '==', d).limit(1).get()));
    const before = days.length;
    days = days.filter((d, i) => {
      const doc = probes[i].docs[0];
      if (!doc) return true;
      const x = doc.data();
      return !(x.commission != null && Object.values(x.sources || {}).some((s) => s.bills != null));
    });
    console.log(`Bỏ qua ${before - days.length} ngày đã đủ trường; còn ${days.length} ngày cần nạp.`);
    if (!days.length) return;
  }

  try {
    await login(page);
    if (days.length === 1) {
      await goToReport(page);
      await syncDay(page, days[0], db);
      return;
    }

    // Nạp bù nhiều ngày: file do máy chủ Fabi dựng (~5 phút/ngày), nên mở nhiều tab
    // (chung phiên đăng nhập) cho mỗi tab xuất 1 ngày cùng lúc. Ngày lỗi thì ghi lại,
    // tab đó mở lại trang báo cáo và lấy ngày kế tiếp.
    const parallel = Math.max(1, Math.min(Number(arg('--parallel')) || 4, days.length));
    const queue = [...days];
    const failed = [];
    const pages = [page];
    while (pages.length < parallel) pages.push(await context.newPage());
    await Promise.all(
      pages.map(async (p) => {
        await goToReport(p);
        for (let day; (day = queue.shift()); ) {
          try {
            await syncDay(p, day, db);
          } catch (err) {
            // Mạng chậm: Firestore đã ghi xong nhưng máy không kịp nhận phản hồi -> báo
            // DEADLINE_EXCEEDED dù dữ liệu đủ. syncState ghi ở batch cuối cùng của ngày,
            // có nó nghĩa là cả ngày đã vào hết.
            if (!DRY_RUN && /DEADLINE_EXCEEDED/.test(err.message)
              && (await db.collection('syncState').doc(`d05_${day}`).get().then((d) => d.exists, () => false))) {
              console.log(`[${day}] ✅ Đã ghi (báo quá giờ nhưng kiểm tra lại thấy đủ).`);
              continue;
            }
            console.error(`❌ ${day}: ${err.message.split('\n')[0]}`);
            failed.push(day);
            await goToReport(p).catch(() => {});
          }
        }
      })
    );
    if (failed.length) throw new Error(`Các ngày lỗi, chạy lại riêng: ${failed.sort().join(', ')}`);
  } finally {
    await browser.close();
  }
}

// day = null: giữ bộ lọc mặc định (hôm nay).
// File xlsx là zip: phải có "End of central directory" (PK\x05\x06) trong 64 KB cuối.
// Mạng đứt giữa chừng thì file vẫn nằm đó nhưng cụt đuôi, xlsx báo "Unsupported ZIP
// Compression method NaN" - bắt sớm ở đây để tải lại thay vì bỏ cả ngày.
function zipLooksComplete(filePath) {
  const size = fs.statSync(filePath).size;
  if (size < 1000) return false;
  const len = Math.min(size, 65557);
  const buf = Buffer.alloc(len);
  const fd = fs.openSync(filePath, 'r');
  try { fs.readSync(fd, buf, 0, len, size - len); } finally { fs.closeSync(fd); }
  return buf.includes(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
}

async function syncDay(page, day, db) {
  const tag = day || 'hôm nay';
  let filePath = null;
  for (let attempt = 1; ; attempt++) {
    if (day) await pickDay(page, day);
    filePath = await exportAndDownload(page, day);
    if (zipLooksComplete(filePath)) break;
    try { fs.unlinkSync(filePath); } catch (e) { /* kệ */ }
    if (attempt >= 3) throw new Error('File tải về bị cụt 3 lần liên tiếp (mạng đứt giữa chừng)');
    console.log(`[${tag}] file tải về bị cụt, tải lại (lần ${attempt})…`);
    await page.waitForTimeout(3000 * attempt);
  }
  const docs = aggregate(readRows(filePath));
  console.log(`[${tag}] ${path.basename(filePath)}: ${PUBLIC_LOG ? Object.keys(docs).length + ' cửa hàng-ngày' : summarize(docs)}`);
  if (DRY_RUN) return;
  const { written, skipped } = await writeChanged(db, docs);
  console.log(`[${tag}] ✅ Ghi ${written} doc thay đổi, bỏ qua ${skipped} doc không đổi.`);
  // Dọn file tạm: đã ghi xong rồi, file biến mất vì lý do gì cũng không phải lỗi của ngày này.
  try { fs.unlinkSync(filePath); } catch (e) { /* đã bị xoá trước đó */ }
}

// Chữ lấy từ hộp thoại Fabi có email người xuất - che đi trước khi đưa vào log công khai.
const noEmail = (t) => String(t).replace(/\s+/g, ' ').replace(/[\w.+-]+@[\w-]+(\.[\w-]+)+/g, '***@***');

function vnNow() {
  // Giờ VN = UTC+7, không có giờ mùa hè
  return new Date(Date.now() + 7 * 3600 * 1000);
}
const vnHour = () => vnNow().getUTCHours();
const vnYesterday = () => new Date(vnNow().getTime() - 864e5).toISOString().slice(0, 10);

function assertEnv(names) {
  const missing = names.filter((n) => !process.env[n]);
  if (missing.length) throw new Error(`Thiếu biến môi trường: ${missing.join(', ')}`);
}

// Fabi thỉnh thoảng tải chậm hoặc chối đăng nhập một lần; hỏng bước này là mất cả mẻ
// nạp bù nên thử lại vài lần trước khi bỏ cuộc.
async function login(page, times = 4) {
  for (let i = 1; ; i++) {
    try { return await loginOnce(page); } catch (e) {
      if (i >= times) throw e;
      console.log(`Đăng nhập Fabi lỗi (lần ${i}): ${String(e.message).split(String.fromCharCode(10))[0].slice(0, 70)} - thử lại…`);
      await page.waitForTimeout(5000 * i);
    }
  }
}

async function loginOnce(page) {
  // KHÔNG chờ 'networkidle': trang Fabi có analytics chạy nền liên tục, mạng chậm là
  // không bao giờ đứng yên -> chờ ô nhập email hiện ra là đủ.
  await page.goto(`${FABI_BASE_URL}/login`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForSelector('input[name="email_input"]', { timeout: 60000 });
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

function dayList(from, to) {
  const out = [];
  for (let d = new Date(from + 'T00:00:00Z'); d <= new Date(to + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + 1)) {
    out.push(d.toISOString().slice(0, 10));
  }
  if (!out.length) throw new Error(`Khoảng ngày không hợp lệ: ${from} → ${to}`);
  return out;
}

const vnDate = (day) => day.split('-').reverse().join('/'); // 2026-09-01 -> 01/09/2026

// Ô chọn ngày của D05 là vue2-daterange-picker, tự áp dụng: bấm ngày bắt đầu rồi
// ngày kết thúc (cùng 1 ngày thì bấm 2 lần) là xong. Có 2 lịch cạnh nhau; bấm mũi
// tên "‹" của lịch trái cho tới khi 1 trong 2 lịch hiện đúng tháng cần chọn.
async function pickDay(page, day) {
  const [y, m, d] = day.split('-').map(Number);
  await page.locator('.reportrange-text').click();
  const picker = page.locator('.daterangepicker:visible');
  await picker.waitFor({ timeout: 10000 });

  const monthOf = (c) => c.evaluate((el) => {
    // Đọc chữ đang hiện ("Tháng 9") thay vì value của ô chọn tháng - value bản
    // Fabi đánh số khác thư viện gốc, đọc value là lệch 1 tháng.
    const sel = el.querySelector('select.monthselect');
    const text = sel ? sel.selectedOptions[0]?.textContent : el.querySelector('th.month')?.textContent;
    return { month: Number((text || '').match(/\d+/)?.[0]), year: Number(el.querySelector('.yearselect')?.value) };
  });

  // Lật ĐÚNG CHIỀU: nạp bù chạy tăng dần nên sau khi xong tháng 1, lịch đang ở tháng 1,
  // muốn tới tháng 2 phải bấm "›". Chỉ bấm "‹" thì không bao giờ tới được.
  let cal = null;
  const target = y * 12 + m;
  for (let i = 0; i < 40 && !cal; i++) {
    let left = null;
    for (const side of ['left', 'right']) {
      const c = picker.locator(`.drp-calendar.${side}`);
      const shown = await monthOf(c);
      if (side === 'left') left = shown;
      if (shown.year === y && shown.month === m) cal = c;
    }
    if (cal) break;
    const at = left && left.year ? left.year * 12 + left.month : target;
    const arrow = target < at ? '.drp-calendar.left .prev' : '.drp-calendar.right .next';
    await picker.locator(arrow).click();
  }
  if (!cal) throw new Error(`Không lật được lịch Fabi tới tháng ${m}/${y}`);

  // Bấm theo chữ số đang hiện. KHÔNG dùng thuộc tính data-date của ô: bản Fabi bị lệch
  // múi giờ, ô hiện "14" lại mang data-date="...-13".
  const cell = cal.locator('td:not(.off):not(.week)').filter({ hasText: new RegExp(`^\\s*${d}\\s*$`) });
  await cell.click();
  await cell.click();

  const want = `${vnDate(day)} - ${vnDate(day)}`;
  await page.locator('.reportrange-text').filter({ hasText: want }).waitFor({ timeout: 10000 });
  await page.waitForTimeout(2500); // bảng tải lại theo ngày mới
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
async function exportAndDownload(page, day) {
  await page.getByRole('button', { name: /Xuất báo cáo/ }).click();
  await page.getByText('Xuất báo cáo tất cả cửa hàng', { exact: true }).click();

  const progress = dialog(page, 'Tiến trình xuất báo cáo');
  const already = dialog(page, 'Dữ liệu này đã được xuất trước đó');
  await progress.or(already).first().waitFor({ timeout: 60000 });
  if (await already.isVisible()) {
    // Chưa gặp với dữ liệu hôm nay; nếu gặp thì chọn xuất mới, không lấy file cũ.
    const again = already.getByRole('button', { name: /Xuất lại|Xuất mới|Tiếp tục xuất/ });
    if (!(await again.count())) {
      throw new Error(`Gặp hộp "đã xuất trước đó" chưa xử lý: ${noEmail(await already.innerText())}`);
    }
    await again.first().click();
    await progress.waitFor({ timeout: 60000 });
  }

  const downloadBtn = progress.locator('button.ep__download');
  const start = Date.now();
  while (!(await downloadBtn.isVisible())) {
    const text = (await progress.innerText().catch(() => '')).replace(/\s+/g, ' ');
    if (/Đã huỷ|Đã hủy|Lỗi|Thất bại/i.test(text)) throw new Error(`Job xuất bị lỗi/huỷ: ${noEmail(text).slice(0, 200)}`);
    if (Date.now() - start > MAX_WAIT_MS) {
      throw new Error('Quá thời gian chờ xuất báo cáo - kiểm tra "Lịch sử xuất" trên Fabi.');
    }
    await page.waitForTimeout(POLL_INTERVAL_MS);
  }

  const [download] = await Promise.all([page.waitForEvent('download', { timeout: 120000 }), downloadBtn.click()]);
  // Tên file dạng d05-01-09-2026-to-01-09-2026.xlsx - chặn trường hợp chọn ngày trượt
  const tag = day ? day.split('-').reverse().join('-') : null;
  if (tag && !download.suggestedFilename().includes(`${tag}-to-${tag}`)) {
    throw new Error(`File tải về (${download.suggestedFilename()}) không khớp ngày ${day}`);
  }
  const downloadPath = path.join(DOWNLOAD_DIR, download.suggestedFilename());
  await download.saveAs(downloadPath);
  // Hộp tiến trình không tự đóng - để nguyên thì nó chặn cú bấm chọn ngày kế tiếp.
  await progress.locator('button.close').click().catch(() => {});
  await progress.waitFor({ state: 'hidden', timeout: 10000 }).catch(() => {});
  return downloadPath;
}

main().catch((err) => {
  console.error('❌ Lỗi:', err);
  process.exit(1);
});
