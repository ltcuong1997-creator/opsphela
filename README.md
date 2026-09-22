# Phê La — Báo cáo bán hàng

Dashboard doanh thu Phê La, dữ liệu lấy tự động từ báo cáo **D05 – Báo cáo bán hàng** trên Fabi POS CMS.

```
public/index.html          trang web (1 file, React + Babel qua CDN, không build) -> opsphela.web.app
sync/sync-report.js        đăng nhập Fabi, xuất D05, gộp số, ghi Firestore (chỉ phần thay đổi)
sync/lib.js                đọc file D05, gộp ngày × cửa hàng × món, ghi theo hash
sync/import-file.js        nạp tay file D05 đã tải sẵn
sync/preview-server.js     xem trước trên máy (localhost:5180), đọc Firestore bằng service account
.github/workflows/         chạy sync-report.js mỗi giờ 6h15–23h15 giờ VN
firestore.rules            chỉ người đăng nhập mới đọc được, không ai ghi từ trình duyệt
```

## Dữ liệu trên Firestore

- `d05Daily/{ngày}_{mã cửa hàng Fabi}` — 1 cửa hàng trong 1 ngày: `revenue` (cột "Tổng tiền"), `bills`, `quantity`, `items{}`.
- `syncState/d05_{ngày}` — hash từng cửa hàng; lần đồng bộ sau chỉ ghi cửa hàng có số liệu khác đi.

## Bí mật — không bao giờ nằm trong repo

| Ở đâu | Gồm |
|---|---|
| `sync/.env` (máy local, đã `.gitignore`) | `FABI_EMAIL`, `FABI_PASSWORD`, `FIREBASE_SERVICE_ACCOUNT_FILE` (đường dẫn file JSON để ngoài repo) |
| GitHub → Settings → Secrets → Actions | `FABI_EMAIL`, `FABI_PASSWORD`, `FIREBASE_SERVICE_ACCOUNT` (nội dung file JSON) |

Repo công khai nên log GitHub Actions cũng công khai: script không in doanh thu và che email khi chạy trên GitHub.

## Lệnh hay dùng (trong thư mục `sync/`)

```bash
npm install && npx playwright install chromium   # lần đầu

npm run preview                                   # xem trước: http://localhost:5180
npm start                                         # đồng bộ ngày hôm nay
node sync-report.js --from 2026-09-01 --to 2026-09-21 --skip-done --parallel 2
                                                  # nạp bù ngày cũ, bỏ ngày đã có, 2 tab song song
npm run import -- <file.xlsx | thư-mục>            # nạp file D05 đã tải tay
```

Triển khai web: `firebase deploy --only hosting` (rules: `--only firestore:rules`).

## Khi có sự cố

- Đổi mật khẩu Fabi → sửa `sync/.env` và secret `FABI_PASSWORD`.
- Fabi đổi giao diện → selector nằm trong `login()`, `pickDay()`, `exportAndDownload()` của `sync-report.js`.
- Tài khoản xem dashboard: Firebase Console → Authentication → Users.
