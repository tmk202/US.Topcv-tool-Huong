# TopCV CV crawler to Lark

Tool này mở TopCV bằng Playwright, lấy dữ liệu ứng viên/CV rồi gửi từng bản ghi vào Lark webhook theo schema:

```json
{
  "job_id": "123456",
  "job_title": "Nhan vien kinh doanh",
  "apply_at": "2026-06-02 10:04:53",
  "candidate_name": "Nguyen Van A",
  "candidate_email": "candidate.test@example.com",
  "candidate_phone": "0901234567",
  "download_url": "https://tuyendung-api.topcv.vn/api/v1/cv-management/onetime-download?token=..."
}
```

## Cai dat

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
cp .env.example .env
```

Sau đó sửa `.env`:

```env
LARK_WEBHOOK_URL=<webhook Lark cua ban>
TOPCV_START_URL=<URL trang danh sach CV/ung vien tren TopCV>
CHROME_CDP_URL=
HEADLESS=false
MAX_PAGES=0
SEND_EXISTING=true
```

`MAX_PAGES=0` nghĩa là không giới hạn số trang. `SEND_EXISTING=false` chỉ xem dữ liệu và lưu state, không gửi các CV đã có sẵn trong lần quét đầu.

## Chay

```bash
source .venv/bin/activate
python scripts/topcv_to_lark.py
```

Lần đầu chạy, trình duyệt sẽ mở ra. Đăng nhập TopCV và vào đúng trang danh sách CV như ảnh bạn gửi. Script sẽ giữ session trong `.topcv-session/`, các lần sau không cần đăng nhập lại nếu cookie còn hạn.

## Chay bang file Excel export

Nếu nút "Xuất danh sách CV" trả về file Excel, nên ưu tiên cách này vì không cần attach Chrome.

Cách 1: tải file `.xlsx` từ TopCV về máy rồi chạy:

```bash
source .venv/bin/activate
python scripts/topcv_excel_to_lark.py --xlsx /path/to/topcv-export.xlsx
```

Cách 2: copy URL `export-excel` vào `.env`:

```env
TOPCV_EXPORT_URL=<URL export-excel cua TopCV>
```

Rồi chạy:

```bash
source .venv/bin/activate
python scripts/topcv_excel_to_lark.py
```

Chạy thử không gửi Lark:

```bash
python scripts/topcv_excel_to_lark.py --xlsx /path/to/topcv-export.xlsx --dry-run
```

File `excel-state.json` lưu các dòng đã gửi để tránh gửi trùng.

## Preview export bang Cookie

Nếu muốn tool tự gọi API export bằng cookie rồi chỉ in kết quả ra terminal:

1. Mở TopCV, bấm `F12`
2. Vào tab `Network`
3. Bấm một request TopCV bất kỳ hoặc request `export-excel`
4. Trong `Request Headers`, copy toàn bộ dòng `Cookie`
5. Dán vào `.env`:

```env
TOPCV_COOKIE=<cookie copy tu DevTools>
```

Chạy:

```bash
source .venv/bin/activate
python scripts/topcv_cookie_export_preview.py
```

Nếu endpoint mặc định cần token, copy nguyên link `export-excel?...___token=...` vào `.env`:

```env
TOPCV_EXPORT_URL=<full export-excel URL>
```

In JSON lines thay vì bảng:

```bash
python scripts/topcv_cookie_export_preview.py --json --limit 0
```

## TUI doc cookie JSON

Nếu bạn export cookie thành file JSON, chỉ cần đặt file đó trong thư mục `crawl-topcv`, ví dụ:

```text
crawl-topcv/topcv-cookies.json
```

Sau đó chạy UI trong terminal:

```bash
source .venv/bin/activate
python scripts/topcv_export_tui.py
```

Menu sẽ tự quét các file `.json` trong thư mục project, lọc file nào có cookie hợp lệ, rồi cho người dùng chọn.

Nếu muốn chỉ định một file cụ thể, đặt đường dẫn trong `.env`:

```env
TOPCV_COOKIE_JSON=/path/to/topcv-cookies.json
```

Tool hỗ trợ các format phổ biến:

- Playwright storage state: `{ "cookies": [{ "name": "...", "value": "...", "domain": "..." }] }`
- Cookie editor export: `[{ "name": "...", "value": "...", "domain": "..." }]`
- Mapping đơn giản: `{ "cookie_name": "cookie_value" }`

Menu sẽ cho chọn nguồn cookie, endpoint export, số dòng hiển thị, định dạng bảng/JSON, và tùy chọn lưu raw file export.

Sau lần chọn đầu tiên, tool lưu lựa chọn vào `.topcv-export-tui.json`. Các lần sau chỉ cần chạy:

```bash
python scripts/topcv_export_tui.py
```

Với cấu hình hiện tại, tool sẽ tải file export, parse toàn bộ cột TopCV, rồi gửi từng dòng vào Lark webhook trong `.env`. Payload gồm:

- Các trường chuẩn: `job_id`, `job_title`, `apply_at`, `candidate_name`, `candidate_email`, `candidate_phone`, `download_url`
- Toàn bộ cột TopCV dạng key phẳng: `topcv_ho_ten`, `topcv_so_dien_thoai`, `topcv_ma_cdtd`, ...
- Toàn bộ cột gốc trong object `topcv_fields`

Chạy preview không gửi Lark:

```bash
python scripts/topcv_export_tui.py --preview-only
```

Để đổi lựa chọn đã lưu:

```bash
python scripts/topcv_export_tui.py --interactive
```

Nếu direct API bị Cloudflare `403`, chọn fetch mode mặc định:

```text
Open export URL in browser, then parse latest download
```

Tool sẽ mở URL export bằng Chrome/browser thật, chờ file `.xlsx` hoặc `.csv` mới xuất hiện trong `~/Downloads`, rồi parse và in ra terminal.

## Dung Chrome dang mo

Playwright chỉ attach được vào Chrome đang chạy nếu Chrome được mở với remote debugging port. Chrome mới yêu cầu remote debugging dùng profile riêng, nên mở bằng script:

```bash
./scripts/open_debug_chrome.sh
```

Sau đó sửa `.env`:

```env
CHROME_CDP_URL=http://localhost:9222
```

Mở tab TopCV danh sách CV trong Chrome đó, rồi chạy script. Khi dùng `CHROME_CDP_URL`, script sẽ bám vào tab TopCV đang có thay vì mở browser riêng.

## Ghi chu

- Không đưa webhook vào code. Đặt trong `.env`.
- File `state.json` lưu các CV đã gửi, tránh gửi trùng khi chạy lại.
- Nếu TopCV đổi giao diện/API, chạy với `HEADLESS=false` để xem trình duyệt đang đứng ở đâu.
