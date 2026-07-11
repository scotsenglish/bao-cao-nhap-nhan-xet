# Dashboard Theo dõi mốc nhập nhận xét — Scots English

Repo này build tự động file `index.html` (dashboard) mỗi khi bạn cập nhật dữ liệu, dùng GitHub Actions.

## Cấu trúc thư mục

```
.
├── data/
│   ├── class_info.xlsx      <- file "All_Branch_Class_Info" export từ LMS
│   ├── student_issue.xlsx   <- file "LMS_Student_Issue" export từ LMS
│   └── grade_report.xlsx    <- file "BC điểm CK" export từ LMS
├── template.html             <- khung dashboard (KHÔNG cần sửa, trừ khi đổi giao diện)
├── build.py                  <- script build (đọc 3 file trên -> tạo index.html)
├── index.html                 <- dashboard hoàn chỉnh (Actions tự tạo/ghi đè, đây là file GitHub Pages sẽ serve)
└── .github/workflows/build.yml   <- workflow tự động chạy build.py
```

> ⚠️ Thư mục `.github/` bắt đầu bằng dấu chấm nên **ẩn mặc định** trên File Explorer/Finder.
> Khi upload qua giao diện web GitHub thì không vấn đề gì (cứ tạo đúng path `.github/workflows/build.yml`
> qua nút "Create new file" và gõ nguyên đường dẫn vào ô tên file, GitHub sẽ tự tạo thư mục).

## Setup lần đầu

1. Tạo repo mới trên GitHub (hoặc dùng repo đã có sẵn từ trước cho dashboard này).
2. Upload đủ các file/thư mục ở trên lên repo (giữ nguyên cấu trúc thư mục).
3. Vào **Settings → Actions → General → Workflow permissions** → chọn **"Read and write permissions"** → Save.
   (Nếu bỏ qua bước này, Action sẽ chạy build nhưng **không commit lại được** index.html mới.)
4. Vào **Settings → Pages** → chọn source là branch `main`, thư mục `/ (root)` → Save.
   → GitHub sẽ cấp cho bạn 1 URL dạng `https://<username>.github.io/<repo-name>/`.
5. Vào tab **Actions**, nếu workflow "Build dashboard" chưa tự chạy, bấm vào workflow đó →
   **"Run workflow"** để chạy lần đầu thủ công.

Sau bước 5, `index.html` sẽ được tạo/cập nhật, và link GitHub Pages ở bước 4 sẽ hiển thị dashboard.

## Cập nhật dữ liệu hàng tháng (hoặc bất cứ khi nào có dữ liệu mới)

1. Xuất 3 file mới từ LMS như thường lệ (All_Branch_Class_Info, LMS_Student_Issue, BC điểm CK).
2. Vào thư mục `data/` trên GitHub (giao diện web), bấm vào từng file cũ → **"Upload files"**
   (hoặc kéo-thả) → chọn file mới → **ĐẶT ĐÚNG TÊN CŨ** (`class_info.xlsx`, `student_issue.xlsx`,
   `grade_report.xlsx`) → Commit.
3. Xong! GitHub Actions sẽ tự động:
   - Phát hiện file trong `data/` vừa thay đổi
   - Chạy lại `build.py`
   - Commit `index.html` mới
   - Link GitHub Pages cũ **tự động** hiển thị dữ liệu mới, không cần làm gì thêm.

Bạn có thể theo dõi tiến trình build ở tab **Actions** — thường mất khoảng 30-60 giây.

### Chỉ có 1-2 trong 3 file mới?

Không sao — `build.py` được viết để chạy được ngay cả khi thiếu file (sẽ in cảnh báo trong log Actions,
nhưng vẫn build ra dashboard với phần dữ liệu đang có). Ví dụ nếu tháng này chỉ có `class_info.xlsx`
mới còn 2 file kia chưa export, cứ upload đè mỗi `class_info.xlsx` là được, 2 file kia giữ nguyên bản cũ.

## Khi nào cần sửa `build.py`?

- **Có chi nhánh mới / đổi AS phụ trách** → sửa biến `REGION_MAPPING_RAW` ở đầu file.
- **Có lịch nghỉ lễ năm mới được công bố** → thêm vào hàm `build_vn_holidays()`.
- Các phần còn lại (logic phân loại mốc, tính tiến độ...) không cần đụng vào trừ khi
  đổi cấu trúc cột của file Excel nguồn.

## Khi nào cần sửa `template.html`?

Chỉ khi muốn đổi giao diện/thêm tính năng mới cho dashboard (bố cục, cột, biểu đồ, KPI...).
File này không chứa dữ liệu thật — dữ liệu được `build.py` "bơm" vào 2 chỗ placeholder:
`__DATA_JSON__` và `__COMMENTS_JSON__`.

## Đồng bộ tick tay (Google Sheet)

URL Google Apps Script cho tính năng tick tay đồng bộ đã được nhúng sẵn trong `template.html`,
không cần cấu hình lại. Nếu sau này đổi sang Google Sheet khác, tìm biến `SHEET_API_URL`
trong `template.html` và thay URL mới vào.
