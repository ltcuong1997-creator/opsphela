# Phê La Dashboard

Trang báo cáo dạng "Looker Studio thu nhỏ": đăng nhập mới xem được, sidebar liệt kê các loại báo cáo, chọn khoảng thời gian, xem KPI + biểu đồ + bảng chi tiết. Dữ liệu được đổ vào từ Fabi POS CMS qua script RPA (thư mục `sync/`), tất cả về chung 1 cấu trúc dữ liệu trong Firestore để trang web (`public/`) đọc và tự vẽ, không cần sửa code mỗi khi thêm báo cáo mới.

```
phela-dashboard/
  firebase.json / .firebaserc / firestore.rules   -> cấu hình Firebase Hosting + Firestore
  public/                                          -> trang web (Hosting) - phela.web.app
    index.html / styles.css / app.js
  sync/                                            -> script RPA đồng bộ dữ liệu từ Fabi
    sync-report.js / seed-report-types.js
  .github/workflows/sync-report.yml                -> chạy tự động theo lịch (GitHub Actions)
```

## Các bước làm theo thứ tự

### 1. Tạo Firebase project

1. Vào https://console.firebase.google.com → "Add project" → đặt tên (vd `phela-dashboard`).
2. Vào mục **Build → Authentication** → tab "Sign-in method" → bật **Email/Password**.
3. Vào **Build → Authentication → Users** → "Add user" → tạo tài khoản đăng nhập cho chính bạn (và bất kỳ ai cần xem dashboard).
4. Vào **Build → Firestore Database** → "Create database" → chọn chế độ **Production** → chọn khu vực (vd `asia-southeast1`).
5. Vào **Build → Hosting** → "Get started" (chỉ cần bấm qua, sẽ deploy bằng CLI ở bước sau).

### 2. Gắn tên miền phela.web.app

Domain `<gì-đó>.web.app` mặc định trùng project ID. Có 2 cách để ra đúng `phela.web.app`:

- **Cách 1 (dễ nhất):** khi tạo project ở bước 1, đặt Project ID là `phela` (Firebase cho sửa Project ID ngay lúc tạo, trước khi bấm Create) — nếu ID này chưa ai dùng thì domain mặc định sẽ là `phela.web.app` và `phela.firebaseapp.com`.
- **Cách 2 (nếu ID `phela` đã bị người khác dùng):** tạo project với ID khác tuỳ ý, sau đó chạy `firebase hosting:sites:create phela` (lệnh CLI, xem bước 3) để tạo thêm 1 Hosting site tên `phela` trong project đó → site này sẽ có domain `phela.web.app` dù project ID là gì.

### 3. Cài Firebase CLI và deploy trang web

Trên máy bạn (cần Node.js đã cài sẵn):

```bash
npm install -g firebase-tools
firebase login

cd phela-dashboard
firebase use --add        # chọn project vừa tạo, đặt alias "default"
# Nếu dùng Cách 2 ở trên:
# firebase hosting:sites:create phela
# rồi sửa firebase.json thêm "site": "phela" trong mục "hosting"

firebase deploy --only firestore:rules,hosting
```

Sau lệnh deploy, trang sẽ chạy tại `https://phela.web.app` (hoặc domain Firebase báo ra) — nhưng lúc này vẫn chưa xem được gì vì `app.js` còn thiếu cấu hình.

### 4. Lấy cấu hình Firebase Web App và điền vào app.js

1. Firebase Console → Project settings (biểu tượng bánh răng) → cuộn xuống "Your apps" → bấm biểu tượng `</>` để tạo 1 Web app (nếu chưa có) → đặt tên bất kỳ.
2. Copy đoạn `firebaseConfig = {...}` được hiển ra.
3. Mở `public/app.js`, thay toàn bộ object `firebaseConfig` (đang để `"REPLACE_ME"`) bằng đoạn vừa copy.
4. Chạy lại `firebase deploy --only hosting` để cập nhật.
5. Mở `https://phela.web.app`, đăng nhập bằng tài khoản đã tạo ở bước 1.3 — sẽ thấy màn hình dashboard nhưng sidebar báo "Chưa có báo cáo nào" vì Firestore chưa có dữ liệu.

### 5. Lấy Firebase service account key (để script RPA ghi được dữ liệu)

Firebase Console → Project settings → tab "Service accounts" → "Generate new private key" → tải về file JSON. Giữ file này riêng tư, không commit lên Git, không gửi qua chat.

### 6. Tạo các "loại báo cáo" (reportTypes)

```bash
cd sync
npm install
FIREBASE_SERVICE_ACCOUNT="$(cat /đường/dẫn/service-account.json)" node seed-report-types.js
```

Lệnh này tạo sẵn 2 loại báo cáo mẫu trong `seed-report-types.js` (doanh thu tổng quan, bán hàng theo cửa hàng). Mở lại trang `phela.web.app`, sidebar sẽ hiện 2 mục này (dù bên trong chưa có số liệu).

### 7. Hoàn thiện script RPA và chạy thử

`sync/sync-report.js` đã viết sẵn luồng: đăng nhập Fabi → xuất báo cáo D05 → đọc file → map vào cấu trúc chung → ghi Firestore. Hai chỗ cần bạn (hoặc Claude Code) kiểm tra lại trước khi chạy thật, vì mình chưa tự ý nhập thử tài khoản Fabi của bạn:

- Selector ô email/mật khẩu/nút đăng nhập ở `https://fabi.ipos.vn/login` (hàm `login()`).
- Nút "Tải file" xuất hiện sau khi job xuất báo cáo chạy xong (hàm `exportAndDownload()`).
- Tên cột thật trong file Excel D05 có khớp với `mapRow()` không (vd `'Thời gian'`, `'Tổng tiền'`, `'Số lượng'`) — mở thử 1 file đã xuất để đối chiếu.

Chạy thử ở máy local:

```bash
cd sync
export FABI_EMAIL="..."
export FABI_PASSWORD="..."
export FIREBASE_SERVICE_ACCOUNT="$(cat /đường/dẫn/service-account.json)"
node sync-report.js
```

Nếu chạy xong không lỗi, mở `phela.web.app`, chọn "Bán hàng theo cửa hàng" — sẽ thấy biểu đồ và bảng có dữ liệu thật.

### 8. Đặt lịch chạy tự động (GitHub Actions)

1. Đẩy toàn bộ thư mục `phela-dashboard` lên 1 GitHub repo **riêng tư**.
2. Repo → Settings → Secrets and variables → Actions → tạo 3 secret: `FABI_EMAIL`, `FABI_PASSWORD`, `FIREBASE_SERVICE_ACCOUNT`.
3. Vào tab Actions → workflow "Sync Fabi reports to Firebase" → "Run workflow" để test chạy tay.
4. Nếu ổn, để lịch mặc định (mỗi giờ) hoặc sửa dòng `cron` trong `.github/workflows/sync-report.yml` theo tần suất bạn muốn (giờ UTC = giờ VN - 7).

## Thêm một báo cáo mới sau này

Không cần sửa `public/app.js`. Chỉ cần:

1. Thêm 1 object mô tả vào `REPORT_TYPES` trong `seed-report-types.js` (label, chartType `line`/`bar`/`table`, dimensionKey, metricKeys, metricLabels, units) → chạy lại `npm run seed`.
2. Thêm 1 object vào `REPORTS` trong `sync-report.js` với `reportUrl` trỏ tới trang báo cáo Fabi tương ứng và hàm `mapRow()` map đúng cột Excel của báo cáo đó vào `{ date, dimensions, metrics }` khớp với `dimensionKey`/`metricKeys` đã khai ở bước 1.
3. Chạy lại `node sync-report.js` (hoặc đợi lịch tự động) — báo cáo mới sẽ tự xuất hiện trên sidebar.

## Về sau (chưa làm ở bản này, có thể mở rộng khi cần)

- Bộ lọc theo cửa hàng cụ thể (hiện đang gộp tất cả).
- Chọn khoảng ngày tuỳ ý thay vì chỉ 4 preset.
- Giao diện sidebar dạng drawer cho di động (hiện bị ẩn ở màn hình nhỏ).
- Phân quyền nhiều cấp (vd quản lý vùng chỉ thấy cửa hàng của mình) qua Firestore Security Rules + custom claims.
