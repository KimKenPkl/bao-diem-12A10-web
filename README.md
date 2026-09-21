# Báo điểm 12A10 — Web gửi PH (public)

🌐 Web public: **https://kimkenpkl.github.io/bao-diem-12A10-web/**

Tra điểm KTĐK **Lần 1 → Lần 4** lớp 12A10 (THPT Chi Lăng), bấm **Copy tin nhắn** để gửi Zalo cho phụ huynh.

## Form tin nhắn (đúng yêu cầu — KHÔNG thứ hạng, KHÔNG nhận xét)

```
Kính gửi phụ huynh em [Họ tên] (lớp 12A10),
Kết quả kiểm tra định kì LẦN 4 của em như sau:
Toán: 7.25
Hoá: 8.5
Sinh: 4.6
Văn: —
Tổng điểm khối B (Toán + Hoá + Sinh): 20.35
Kính mong phụ huynh nhắc nhở, động viên em ôn tập.
Trân trọng!
[GVCN]
```

- Văn chỉ thi Lần 1 → các Lần 2/3/4 hiện `—`.
- Tổng khối B = Toán + Hoá + Sinh (tự tính trong Excel, web giữ nguyên).
- Nguồn: `2026-2027- 12A10 ĐK LẦN 4.xlsx` (33 em, 4 sheet Lần 1-4).

## Cách dùng hằng tuần / hằng tháng

1. Mở web trên điện thoại/máy tính → chọn **Lần 4** (mới nhất) → gõ tên tìm em → **Copy tin nhắn** → dán qua Zalo PH.
2. Muốn ký tên: nhập tên GVCN vào ô trên cùng (lưu trên máy, lần sau tự nhớ).
3. Nút **Copy toàn bộ danh sách** để dán ra Excel đối chiếu.

## Cập nhật Lần 5, 6… (2 phút)

1. Mở file Excel điểm mới, copy sheet vào file `2026-2027- 12A10 ĐK LẦN 4.xlsx` hoặc gửi file mới cho tôi.
2. Chạy: `python dump_data.py` (trong thư mục dữ liệu) để sinh lại `data.js`, hoặc nhắn tôi làm.
3. `git add data.js && git commit -m "cap nhat Lan 5" && git push` — web tự cập nhật sau ~1 phút (GitHub Pages).

## File trong repo

- `index.html` — web tĩnh (không backend, mở offline vẫn chạy).
- `data.js` — điểm 4 lần (33 HS × Toán/Hoá/Sinh/Văn/B00).
- `.nojekyll` — cấu hình GitHub Pages.
