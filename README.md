# Báo điểm 12A10 — Web gửi PH (public)

🌐 Web public: **https://kimkenpkl.github.io/bao-diem-12A10-web/**

Tra điểm KTĐK **Lần 1 → Lần 8** lớp 12A10 (THPT Chi Lăng), bấm **Copy tin nhắn** để gửi Zalo cho phụ huynh.

## Ba chế độ gửi (mặc định: đợt mới Lần 6–8)

| Chế độ | Nội dung tin nhắn | Dùng khi |
|---|---|---|
| **Gửi: Lần 6–8 (đợt mới)** ← mặc định | 3 hàng Lần 6, 7, 8 | PH đã nhận Lần 1–5, chỉ gửi 3 lần mới |
| Gửi: đủ 8 Lần/tin | 8 hàng Lần 1 → 8 | PH chưa từng nhận tin nào |
| Gửi: từng Lần riêng | 1 hàng của Lần đang chọn | Gửi bù / gửi lại đúng 1 lần |

Mỗi thẻ học sinh còn có hàng **Gửi nhanh: L6 · L7 · L8** — bấm là copy ngay tin của riêng lần đó, không cần đổi chế độ.

## Form tin nhắn (đúng yêu cầu — KHÔNG thứ hạng, KHÔNG nhận xét)

```
Kính gửi phụ huynh em [Họ tên] (lớp 12A10),
Kết quả KTĐK các lần gần nhất của em:
- Lần 6: Toán 5.5, Hoá 7.6, Sinh 8.25, Văn — | Tổng B: 21.35
- Lần 7: ...
- Lần 8: ...
Kính mong phụ huynh nhắc nhở, động viên em ôn tập.
Trân trọng!
[GVCN]
```

- Văn chỉ thi Lần 1 và 3 → các Lần 2/4/5/6/7/8 hiện `—`.
- Tổng khối B = Toán + Hoá + Sinh (tự tính trong Excel, web giữ nguyên).
- Nguồn: file KTĐK hằng tuần (`...Bảng điểm định kì hằng tuần (8).xlsx`, sheet Lần 1–8) + file thống kê (`...THỐNG KÊ ĐIỂM...xlsx`).
- Sĩ số 33: đã **xoá em Phạm Hữu Nam Huy** (nghỉ học) khỏi cả 7 Lần, SBD chuẩn theo Lần 7; **thêm em Nguyễn Trịnh Bảo Ngọc** (mới, có điểm từ Lần 5; Lần 1–4 hiện "chưa có điểm").
- **21/10/2026**: cập nhật lại Lần 6 theo file (8) — Phan Khánh Hằng Toán 6.5 → **7.0** (Tổng B 19.35 → 19.85) và xếp hạng Lần 6 dịch theo; thêm chế độ gửi đợt mới Lần 5–7.
- **08/11/2026**: cập nhật **Lần 8** (33 HS) từ file Bảng điểm định kì hằng tuần mới nhất; mặc định đợt mới chuyển thành **Lần 6–8**.

## Cách dùng hằng tuần / hằng tháng

1. Mở web trên điện thoại/máy tính → chọn **Lần 8** (mới nhất) → gõ tên tìm em → **Copy tin nhắn** → dán qua Zalo PH.
2. Muốn ký tên: nhập tên GVCN vào ô trên cùng (lưu trên máy, lần sau tự nhớ).
3. Nút **Copy toàn bộ danh sách** để dán ra Excel đối chiếu.

## Cập nhật Lần 8, 9… (2 phút)

1. Bổ sung điểm vào file Excel (sheet Lần mới), đối chiếu Tổng B = Toán + Hoá + Sinh.
2. Chạy script build `data.js` từ Excel (hoặc nhắn tôi làm).
3. `git add data.js && git commit -m "cap nhat Lan 8" && git push` — web tự cập nhật sau ~1 phút (GitHub Pages).

## Điểm giữa kì 1 (11 môn) — trang riêng

Trang **`gk1.html`**: https://kimkenpkl.github.io/bao-diem-12A10-web/gk1.html

- Đủ **11 môn**: Toán, Lý, Anh, Văn, Sử, Địa, Hoá, Sinh, Tin, KTPL, GDQP.
- **B00 = Toán + Hoá + Sinh** và **Tổng = tổng 11 môn** (web tự tính lại bằng Python, không tin công thức Excel).
- Web có hiện **hạng B00 / hạng tổng** để thầy theo dõi; **tin nhắn gửi PH không có xếp hạng**.
- Link "📘 Giữa kì 1" ở trang chủ **chỉ hiện khi `data.js` đã có điểm GK1 thật** (tránh đưa trang rỗng lên mạng).
- Khi chưa có điểm, `gk1.html` hiện thông báo "Chưa có điểm giữa kì 1".

### Quy trình cập nhật GK1 (3 bước)

1. **Nhập điểm** vào file mẫu `12A10-GK1-MAU.xlsx` (33 HS, cột D–N là 11 môn; O/P/Q/R tự tính B00, Tổng, hạng).
   - Hoặc dùng trang nhập nhanh local `nhap-diem.html` (mở bằng trình duyệt, gõ điểm, Enter xuống dòng cùng môn, Tab sang môn; dữ liệu tự lưu trong máy) rồi copy TSV dán vào ô D4 của Excel.
2. **Sinh `data.js`**:
   ```powershell
   $env:PYTHONIOENCODING='utf-8'
   python build_data.py --gk1 "12A10-GK1-MAU.xlsx"
   ```
   - File GK1 rỗng → script **không** thêm key `"Giữa kì 1"` (an toàn).
   - Có thể xem trước bằng `--out data_test.js`.
3. **Push**:
   ```powershell
   git add data.js && git commit -m "cap nhat diem giua ki 1" && git push
   ```

### Công cụ build

`build_data.py` đọc Excel → tính lại B00/Tổng/hạng → ghi `data.js`.

```powershell
# cập nhật điểm tuần (sheet Lần N)
python build_data.py --weekly "Bang diem tuan.xlsx"

# cập nhật GK1, giữ nguyên các Lần đang có
python build_data.py --gk1 "GK1.xlsx"

# chỉ ghi đè khi tên học sinh lệch so với data.js (mặc định: bỏ qua sheet lệch)
python build_data.py --weekly "..." --force
```

Chốt an toàn: sheet có tên học sinh **lệch** so với `data.js` hiện có sẽ **bị bỏ qua** (trừ khi `--force`) — tránh ghi đè nhầm khi file Excel đổi danh sách.

Sau khi ghi `data.js`, script **tự tăng cache-buster** `data.js?v=<ngày_giờ>` trong `index.html` và `gk1.html` (bỏ qua bằng `--no-bump`) — nhờ đó trình duyệt/PH không còn thấy dữ liệu cũ sau khi push. Mặc định trang chủ cũng **tự lấy Lần mới nhất**, không hard-code.

## File trong repo

- `index.html` — web tĩnh điểm tuần (không backend, mở offline vẫn chạy).
- `gk1.html` — trang điểm giữa kì 1 (11 môn + B00 + Tổng).
- `data.js` — điểm Lần 1–8 (33 HS × Toán/Hoá/Sinh/Văn/B00/XH) + `"Giữa kì 1"` khi có điểm.
- `build_data.py` — script sinh `data.js` từ Excel.
- `.nojekyll` — cấu hình GitHub Pages.

> `nhap-diem.html` (trang nhập điểm local) và `danhba-private.js` (SĐT phụ huynh) nằm trong `.gitignore` — **không** đưa lên mạng.
