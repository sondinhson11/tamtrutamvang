import random
from openpyxl import Workbook
import os

# Chỉnh Sửa Ở Đây
so_luong_nguoi = 16  # Số Người
ma_tinh = '026'  # Mã tỉnh
nam_sinh_nho_nhat = 1970  # Năm Sinh Nhỏ Nhất Trong Danh Sách
nam_sinh_lon_nhat = 1995  # Năm Sinh Lớn Nhất Trong Danh Sách
noi_lam_viec = "Nhân Đạo"
dan_toc = "Kinh"
quoc_tich = "Việt Nam"
diachi_quocgia = "Cộng hòa xã hội chủ nghĩa Việt Nam"
thanh_pho = "Vĩnh Phúc"
quan_huyen = "Sông Lô"
phuong_xa = "Nhân Đạo"
dia_chi = "Nhân Đạo"
loai_cu_tru = "Tạm trú"
ngay_den = "08/07/2023"
ngay_di = "09/07/2023"
ly_do = "Du Lịch"

#
#

# Kiểm tra nếu tệp Excel đã tồn tại, thì xóa nó
if os.path.exists('danh_sach_nguoi_yeu_cau.xlsx'):
    os.remove('danh_sach_nguoi_yeu_cau.xlsx')


# Tạo một danh sách các họ, tên đệm và tên
ho_list = [
    'Nguyễn', 'Trần', 'Lê', 'Phạm', 'Hoàng', 'Huỳnh', 'Phan', 'Vũ', 'Đặng', 'Bùi',
    'Đỗ', 'Hồ', 'Ngô', 'Dương', 'Lý', 'Trương', 'Võ', 'Đoàn', 'Trịnh',
    'Lương', 'Đào', 'Mai', 'Hà', 'Đinh', 'Thái', 'Vương', 'Trịnh'
]
ten_dem_nam_list = [
    'Văn', 'Đức', 'Hữu', 'Thành', 'Minh', 'Quốc', 'Công', 'Đình', 'Phước', 'Gia',
    'Nhật', 'Sơn', 'Tuấn', 'Huy', 'Hải', 'Khánh', 'Khoa', 'Kiên', 'Long', 'Phúc'
]
ten_dem_nu_list = ['Thị', 'Thu', 'Ngọc', 'Phương', 'Thảo', 'Linh', 'Hạnh']
# Tạo một danh sách các tên nam và nữ
ten_nam_list = [
    'An', 'Bình', 'Cường', 'Dũng', 'Dương', 'Hiếu', 'Linh', 'Minh', 'Quân', 'Sơn',
    'Thành', 'Thủy', 'Tuấn', 'Hưng', 'Hải', 'Khánh', 'Khoa', 'Kiên', 'Long', 'Phúc'
]

ten_nu_list = [
    'Anh', 'Bích', 'Cẩm', 'Diễm', 'Hà', 'Hạnh', 'Lan', 'Linh', 'Mai', 'My',
    'Ngọc', 'Nhi', 'Oanh', 'Phương', 'Quỳnh', 'Thảo', 'Trâm', 'Trang', 'Xuân', 'Yến'
]
# Tạo một danh sách Excel mới
wb = Workbook()
ws = wb.active

# Tạo 80 người yêu cầu
cccd_suffix_list = []  # Danh sách lưu trữ các số cuối của căn cước đã được tạo ra

for i in range(so_luong_nguoi):
    ho = random.choice(ho_list)
    gioi_tinh = random.choice(['Giới tính Nam', 'Giới tính Nữ'])
    ngay_sinh = random.randint(1, 28)
    thang_sinh = random.randint(1, 12)
    nam_sinh = random.randint(nam_sinh_nho_nhat, nam_sinh_lon_nhat)

    # Chỉ chọn người yêu cầu đăng ký khai sinh tại Hà Nội
    # Lấy mã giới tính
    if nam_sinh >= 1900 and nam_sinh <= 1999:
        ma_gioi_tinh = '0' if gioi_tinh == 'Giới tính Nam' else '1'
    elif nam_sinh >= 2000 and nam_sinh <= 2099:
        ma_gioi_tinh = '2' if gioi_tinh == 'Giới tính Nam' else '3'
    elif nam_sinh >= 2100 and nam_sinh <= 2199:
        ma_gioi_tinh = '4' if gioi_tinh == 'Giới tính Nam' else '5'
    elif nam_sinh >= 2200 and nam_sinh <= 2299:
        ma_gioi_tinh = '6' if gioi_tinh == 'Giới tính Nam' else '7'
    elif nam_sinh >= 2300 and nam_sinh <= 2399:
        ma_gioi_tinh = '8' if gioi_tinh == 'Giới tính Nam' else '9'

    # Lấy mã năm sinh
    ma_nam_sinh = str(nam_sinh)[-2:]

    # Tạo số căn cước
    while True:
        suffix = random.randint(1000, 9999)  # Số cuối của căn cước
        if suffix not in cccd_suffix_list:
            cccd_suffix_list.append(suffix)
            break

    random_numbers = ''.join(str(random.randint(0, 9)) for _ in range(6))

    so_can_cuoc = f'{ma_tinh}{ma_gioi_tinh}{ma_nam_sinh}{random_numbers}'

    if gioi_tinh == 'Giới tính Nam':
        ten_dem = random.choice(ten_dem_nam_list)
    else:
        ten_dem = random.choice(ten_dem_nu_list)

    if gioi_tinh == 'Giới tính Nam':
        ten = random.choice(ten_nam_list)
    else:
        ten = random.choice(ten_nu_list)

    nghe_nghiep = 'Tự do'

    ho_ten = f'{ho} {ten_dem} {ten}'

    ws.append([ho_ten, f'{ngay_sinh:02d}/{thang_sinh:02d}/{nam_sinh}', gioi_tinh, so_can_cuoc, '', '', '', nghe_nghiep, noi_lam_viec,
              dan_toc, quoc_tich, diachi_quocgia, thanh_pho, quan_huyen, phuong_xa, dia_chi, loai_cu_tru, ngay_den, ngay_di, ly_do])

# Lưu danh sách vào tệp Excel
wb.save('danh_sach_nguoi_yeu_cau.xlsx')
print("Đã tạo thành công danh sách người yêu cầu!")


# 01: Thành phố Hà Nội
# 02: Tỉnh Hà Giang
# 04: Tỉnh Cao Bằng
# 06: Tỉnh Bắc Kạn
# 08: Tỉnh Tuyên Quang
# 10: Tỉnh Lào Cai
# 11: Tỉnh Điện Biên
# 12: Tỉnh Lai Châu
# 14: Tỉnh Sơn La
# 15: Tỉnh Yên Bái
# 17: Tỉnh Hòa Bình
# 19: Tỉnh Thái Nguyên
# 20: Tỉnh Lạng Sơn
# 22: Tỉnh Quảng Ninh
# 24: Tỉnh Bắc Giang
# 25: Tỉnh Phú Thọ
# 26: Tỉnh Vĩnh Phúc
# 27: Tỉnh Bắc Ninh
# 30: Tỉnh Hải Dương
# 31: Thành phố Hải Phòng
# 33: Tỉnh Hưng Yên
# 34: Tỉnh Thái Bình
# 35: Tỉnh Hà Nam
# 36: Tỉnh Nam Định
# 37: Tỉnh Ninh Bình
# 38: Tỉnh Thanh Hóa
# 40: Tỉnh Nghệ An
# 42: Tỉnh Hà Tĩnh
# 44: Tỉnh Quảng Bình
# 45: Tỉnh Quảng Trị
# 46: Tỉnh Thừa Thiên Huế
# 48: Thành phố Đà Nẵng
# 49: Tỉnh Quảng Nam
# 51: Tỉnh Quảng Ngãi
# 52: Tỉnh Bình Định
# 54: Tỉnh Phú Yên
# 56: Tỉnh Khánh Hòa
# 58: Tỉnh Ninh Thuận
# 60: Tỉnh Bình Thuận
# 62: Tỉnh Kon Tum
# 64: Tỉnh Gia Lai
# 66: Tỉnh Đắk Lắk
# 67: Tỉnh Đắk Nông
# 68: Tỉnh Lâm Đồng
# 70: Tỉnh Bình Phước
# 72: Tỉnh Tây Ninh
# 74: Tỉnh Bình Dương
# 75: Tỉnh Đồng Nai
# 77: Tỉnh Bà Rịa - Vũng Tàu
# 79: Thành phố Hồ Chí Minh
# 80: Tỉnh Long An
# 82: Tỉnh Tiền Giang
# 83: Tỉnh Bến Tre
# 84: Tỉnh Trà Vinh
# 86: Tỉnh Vĩnh Long
# 87: Tỉnh Đồng Tháp
# 89: Tỉnh An Giang
# 91: Tỉnh Kiên Giang
# 92: Thành phố Cần Thơ
# 93: Tỉnh Hậu Giang
# 94: Tỉnh Sóc Trăng
# 95: Tỉnh Bạc Liêu
# 96: Tỉnh Cà Mau
