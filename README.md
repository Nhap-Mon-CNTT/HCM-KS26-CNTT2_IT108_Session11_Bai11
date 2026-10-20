

# Phần 1 – Thiết kế kiến trúc

## 1. Sơ đồ luồng dữ liệu

| Thành phần | Loại | Giải thích |
|---|---|---|
| Excel | **Nguồn** | Chỗ duy nhất nhập và sửa số liệu. Sửa ở đây thì Word và Slide cập nhật theo. |
| Word | **Hiển thị** (chi tiết) | Trình bày số liệu dạng bảng; đồng thời là nơi chốt nội dung văn bản (lịch trình, giải thưởng, quy chế) mà Slide trích lại. |
| Slide | **Hiển thị** (rút gọn) | Mỗi slide một thông điệp rút ra từ Excel và Word; không chứa dữ liệu gốc. |

## 2. Cấu trúc Heading của tài liệu Word (19 dòng tiêu đề, 3 cấp)

```
H1  1. Giới thiệu hội nghị
H2    1.1. Mục tiêu
H2    1.2. Thời gian và địa điểm
H2    1.3. Quy mô và đối tượng tham dự
H1  2. Cơ cấu tổ chức
H2    2.1. Ban Chỉ đạo
H2    2.2. Ban Kỹ thuật
H3      2.2.1. Nhiệm vụ của Trưởng Ban Kỹ thuật
H3      2.2.2. Bộ hồ sơ kỹ thuật
H1  3. Lịch trình hội nghị
H2    3.1. Buổi sáng
H2    3.2. Buổi chiều
H1  4. Cơ cấu giải thưởng
H2    4.1. Các giải thưởng
H2    4.2. Tiêu chí chấm điểm
H1  5. Quy chế hội nghị
H2    5.1. Quy định đối với báo cáo viên
H2    5.2. Quy định đối với Ban giám khảo
H2    5.3. Quy định đối với người tham dự
```

---

# Phần 2 – Sản phẩm: các điểm đã đáp ứng ràng buộc

## 3/ **Word**
File: [KeHoach_HoiNghi_PTIT2027.docx](KeHoach_HoiNghi_PTIT2027.docx)

## 4/ **Excel**
File: [DanhSachBaoCao_PTIT2027.xlsx](DanhSachBaoCao_PTIT2027.xlsx)

## 5/ **Slide**
File:  [SlideKhaiMac_PTIT2027.pptx](SlideKhaiMac_PTIT2027.pttx)



## 6/ Báo cáo xử lý 3 Edge Case

### Edge Case 1 (Word): đổi màu Heading 1 từ đỏ sang navy cho tài liệu 40 trang

| Cách làm | Số thao tác |
|---|---|
| **Dùng Heading Styles** | 1 thao tác (khoảng 4 cú nhấp: chuột phải Heading 1, Modify, Font Color, chọn Navy, OK) | 
| **Định dạng thủ công** | Khoảng 60 thao tác nếu 40 trang có 20 mục H1 (mỗi mục 3 thao tác: bôi đen, đổi màu sang Navy,và chỉnh kích thước chữ) | 



### Edge Case 2 (Excel): 5 báo cáo viên mời ngoài chưa có trong `DanhMucPhong`

- **VLOOKUP hiển thị:** lỗi `#N/A`: không tìm thấy khóa
- **Cách xử lý không sửa `DanhMucPhong`:** bọc VLOOKUP bằng `IFERROR`:
  - Phòng hiển thị "Chưa xếp phòng".
  - Thù lao hiển thị 0 để tổng vẫn tính đúng.

### Edge Case 3 (Slide): ban giám khảo yêu cầu bảng số liệu đầy đủ

- Giữ slide chính ngắn gọn: bảng số liệu không đưa vào slide trình chiếu.
- Thêm slide phụ lục ẩn khi bấm vào khi bấm vào mới hiển thị. Khi trình chiếu, slide này bị bỏ qua nên Rule 10/20/30 của những silde chính không bị phá.

