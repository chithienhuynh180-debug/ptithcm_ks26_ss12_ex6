# Bài Smart-Home

## Phần 1. Phân tích

bảng xấu:
- có ô merge, có hàng trống hoặc cột trống
- 1 ô chứa nhiều giá trị (vd "Q1, 2 phòng ngủ, 3 tỷ")
- không có dòng tiêu đề rõ, trộn nhiều bảng trong 1 sheet
- ghi số kèm chữ (vd "3 tỷ") nên không tính được

bảng đúng:
- dòng 1 là tiêu đề, mỗi dòng là 1 record, mỗi cột là 1 field
- 1 ô 1 giá trị, không merge, không hàng trống
- số để dạng số nên tính được

## Phần 2. Triển khai

### Word. SRS Smart-Home

# 1. Giới thiệu
## 1.1 Mục đích
## 1.2 Phạm vi
# 2. Yêu cầu chức năng
## 2.1 Quản lý tin nhà đất
### 2.1.1 Đăng tin
### 2.1.2 Duyệt tin
## 2.2 Đấu giá trực tuyến
### 2.2.1 Đặt giá
### 2.2.2 Chốt người thắng
# 3. Yêu cầu phi chức năng
## 3.1 Bảo mật
## 3.2 Hiệu năng

### Excel. Danh sách nhà đất

| Mã nhà | Địa chỉ | Giá niêm yết | Chiết khấu | Giá sau chiết khấu | Mã môi giới | Hoa hồng | Giá thuê |
|---|---|---|---|---|---|---|---|
| N001 | Quận 7 | 3000000000 | 5% | `=C2*(1-D2)` | MG01 | `=VLOOKUP(F2,$K$2:$L$4,2,FALSE)` | `=C2*0.004` |

bảng tra hoa hồng (cột K, L):

| Mã môi giới | % hoa hồng |
|---|---|
| MG01 | 2% |
| MG02 | 1.5% |
| MG03 | 1% |

### PowerPoint. Pitch deck (8 slide, mỗi slide 1 ý, để nhiều khoảng trắng)

1. Smart-Home. đấu giá nhà đất online
2. Mua bán nhà đất còn chậm và thiếu minh bạch
3. Giải pháp: đấu giá trực tuyến công khai
4. Thị trường nhà đất rất lớn
5. Mô hình thu phí: hoa hồng mỗi giao dịch
6. Đội ngũ sáng lập
7. Cần gọi vốn để phát triển
8. Cảm ơn

## Phần 3. Kiểm thử

kéo công thức xuống thì vùng tra `K2:L4` bị trôi thành `K3:L5`, `K4:L6`, nên tra ra #N/A. sửa: khóa vùng bằng dấu $ thành `$K$2:$L$4` (bấm F4). dấu $ giữ cố định cột và hàng, còn lookup_value `F2` vẫn để tự đổi theo dòng.