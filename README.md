# Báo điểm 12A10 — Web gửi PH (public)

🌐 Web public: **https://kimkenpkl.github.io/bao-diem-12A10-web/**

Tra điểm KTĐK **Lần 1 → Lần 6** lớp 12A10 (THPT Chi Lăng), bấm **Copy tin nhắn** để gửi Zalo cho phụ huynh.

## Form tin nhắn (đúng yêu cầu — KHÔNG thứ hạng, KHÔNG nhận xét)

```
Kính gửi phụ huynh em [Họ tên] (lớp 12A10),
Kết quả KTĐK LẦN 6 của em:
Toán 7.75, Hoá 7.75, Sinh 6, Văn — | Tổng B: 21.5
Kính mong phụ huynh nhắc nhở, động viên em ôn tập.
Trân trọng!
[GVCN]
```

- Văn chỉ thi Lần 1 và 3 → các Lần 2/4/5/6 hiện `—`.
- Tổng khối B = Toán + Hoá + Sinh (tự tính trong Excel, web giữ nguyên).
- Nguồn: file KTĐK hằng tuần (`...Bảng điểm định kì hằng tuần (5).xlsx`, sheet Lần 1–6) + file thống kê (`...THỐNG KÊ ĐIỂM...xlsx`).
- Sĩ số 33: đã **xoá em Phạm Hữu Nam Huy** (nghỉ học) khỏi cả 6 Lần, SBD chuẩn theo Lần 6; **thêm em Nguyễn Trịnh Bảo Ngọc** (mới, có điểm từ Lần 5; Lần 1–4 hiện "chưa có điểm").

## Cách dùng hằng tuần / hằng tháng

1. Mở web trên điện thoại/máy tính → chọn **Lần 6** (mới nhất) → gõ tên tìm em → **Copy tin nhắn** → dán qua Zalo PH.
2. Muốn ký tên: nhập tên GVCN vào ô trên cùng (lưu trên máy, lần sau tự nhớ).
3. Nút **Copy toàn bộ danh sách** để dán ra Excel đối chiếu.

## Cập nhật Lần 7, 8… (2 phút)

1. Bổ sung điểm vào file Excel (sheet Lần mới), đối chiếu Tổng B = Toán + Hoá + Sinh.
2. Chạy script build `data.js` từ Excel (hoặc nhắn tôi làm).
3. `git add data.js && git commit -m "cap nhat Lan 7" && git push` — web tự cập nhật sau ~1 phút (GitHub Pages).

## File trong repo

- `index.html` — web tĩnh (không backend, mở offline vẫn chạy).
- `data.js` — điểm 6 lần (33 HS × Toán/Hoá/Sinh/Văn/B00/XH).
- `.nojekyll` — cấu hình GitHub Pages.
