# Báo điểm 12A10 — Web gửi PH (public)

🌐 Web public: **https://kimkenpkl.github.io/bao-diem-12A10-web/**

Tra điểm KTĐK **Lần 1 → Lần 7** lớp 12A10 (THPT Chi Lăng), bấm **Copy tin nhắn** để gửi Zalo cho phụ huynh.

## Ba chế độ gửi (mặc định: đợt mới Lần 5–7)

| Chế độ | Nội dung tin nhắn | Dùng khi |
|---|---|---|
| **Gửi: Lần 5–7 (đợt mới)** ← mặc định | 3 hàng Lần 5, 6, 7 | PH đã nhận Lần 1–4, chỉ gửi 3 lần mới |
| Gửi: đủ 7 Lần/tin | 7 hàng Lần 1 → 7 | PH chưa từng nhận tin nào |
| Gửi: từng Lần riêng | 1 hàng của Lần đang chọn | Gửi bù / gửi lại đúng 1 lần |

Mỗi thẻ học sinh còn có hàng **Gửi nhanh: L5 · L6 · L7** — bấm là copy ngay tin của riêng lần đó, không cần đổi chế độ.

## Form tin nhắn (đúng yêu cầu — KHÔNG thứ hạng, KHÔNG nhận xét)

```
Kính gửi phụ huynh em [Họ tên] (lớp 12A10),
Kết quả KTĐK các lần gần nhất của em:
- Lần 5: Toán 5.5, Hoá 7.6, Sinh 8.25, Văn — | Tổng B: 21.35
- Lần 6: ...
- Lần 7: ...
Kính mong phụ huynh nhắc nhở, động viên em ôn tập.
Trân trọng!
[GVCN]
```

- Văn chỉ thi Lần 1 và 3 → các Lần 2/4/5/6/7 hiện `—`.
- Tổng khối B = Toán + Hoá + Sinh (tự tính trong Excel, web giữ nguyên).
- Nguồn: file KTĐK hằng tuần (`...Bảng điểm định kì hằng tuần (8).xlsx`, sheet Lần 1–7) + file thống kê (`...THỐNG KÊ ĐIỂM...xlsx`).
- Sĩ số 33: đã **xoá em Phạm Hữu Nam Huy** (nghỉ học) khỏi cả 7 Lần, SBD chuẩn theo Lần 7; **thêm em Nguyễn Trịnh Bảo Ngọc** (mới, có điểm từ Lần 5; Lần 1–4 hiện "chưa có điểm").
- **21/10/2026**: cập nhật lại Lần 6 theo file (8) — Phan Khánh Hằng Toán 6.5 → **7.0** (Tổng B 19.35 → 19.85) và xếp hạng Lần 6 dịch theo; thêm chế độ gửi đợt mới Lần 5–7.

## Cách dùng hằng tuần / hằng tháng

1. Mở web trên điện thoại/máy tính → chọn **Lần 7** (mới nhất) → gõ tên tìm em → **Copy tin nhắn** → dán qua Zalo PH.
2. Muốn ký tên: nhập tên GVCN vào ô trên cùng (lưu trên máy, lần sau tự nhớ).
3. Nút **Copy toàn bộ danh sách** để dán ra Excel đối chiếu.

## Cập nhật Lần 7, 8… (2 phút)

1. Bổ sung điểm vào file Excel (sheet Lần mới), đối chiếu Tổng B = Toán + Hoá + Sinh.
2. Chạy script build `data.js` từ Excel (hoặc nhắn tôi làm).
3. `git add data.js && git commit -m "cap nhat Lan 7" && git push` — web tự cập nhật sau ~1 phút (GitHub Pages).

## File trong repo

- `index.html` — web tĩnh (không backend, mở offline vẫn chạy).
- `data.js` — điểm 7 lần (33 HS × Toán/Hoá/Sinh/Văn/B00/XH).
- `.nojekyll` — cấu hình GitHub Pages.
