import random
from openpyxl import Workbook

# Tạo một danh sách các họ, tên đệm và tên
ho_list = [
    'Nguyễn', 'Trần', 'Lê', 'Phạm', 'Hoàng', 'Huỳnh', 'Phan', 'Vũ', 'Đặng', 'Bùi',
    'Đỗ', 'Hồ', 'Ngô', 'Dương', 'Lý', 'Quách', 'Trương', 'Võ', 'Đoàn', 'Trịnh',
    'Lương', 'Đào', 'Mai', 'Hà', 'Đinh', 'Thái', 'Vương', 'Trịnh', 'Bành', 'Lục'
]
ten_dem_list = [
    'Văn', 'Thị', 'Đức', 'Hữu', 'Đức', 'Thịnh', 'Hải', 'Thu', 'Đình', 'Phước',
    'Thanh', 'Như', 'Xuân', 'Công', 'Ngọc', 'Thế', 'Tuấn', 'Nguyệt', 'Tâm', 'Gia'
]

# Tạo một danh sách các tên
ten_list = [
    'An', 'Bình', 'Cường', 'Dung', 'Dũng', 'Dương', 'Hiếu', 'Linh', 'Minh', 'Trang',
    'Quân', 'Hoài', 'Sơn', 'Thu', 'Thủy', 'Yến', 'Hoa', 'Hương', 'Long', 'Mạnh'
]

# Tạo một danh sách Excel mới
wb = Workbook()
ws = wb.active

# Thêm tiêu đề cho các cột
ws.append(['Họ và tên', 'Ngày sinh', 'Giới tính', 'Số căn cước'])

# Tạo 80 người yêu cầu
for i in range(80):
    ho = random.choice(ho_list)
    ten_dem = random.choice(ten_dem_list)
    ten = random.choice(ten_list)
    gioi_tinh = random.choice(['Giới tính nam', 'Giới tính nữ'])
    ngay_sinh = random.randint(1, 28)
    thang_sinh = random.randint(1, 12)
    nam_sinh = random.randint(1960, 1980)
    
    # Chỉ chọn người yêu cầu đăng ký khai sinh tại Hà Nội
    ma_tinh = '001'  # Mã tỉnh Hà Nội
    
    # Lấy mã giới tính
    if nam_sinh >= 1900 and nam_sinh <= 1999:
        ma_gioi_tinh = '0' if gioi_tinh == 'Giới tính nam' else '1'
    elif nam_sinh >= 2000 and nam_sinh <= 2099:
        ma_gioi_tinh = '2' if gioi_tinh == 'Giới tính nam' else '3'
    elif nam_sinh >= 2100 and nam_sinh <= 2199:
        ma_gioi_tinh = '4' if gioi_tinh == 'Giới tính nam' else '5'
    elif nam_sinh >= 2200 and nam_sinh <= 2299:
        ma_gioi_tinh = '6' if gioi_tinh == 'Giới tính nam' else '7'
    elif nam_sinh >= 2300 and nam_sinh <= 2399:
        ma_gioi_tinh = '8' if gioi_tinh == 'Giới tính nam' else '9'
    
    # Lấy mã năm sinh
    ma_nam_sinh = str(nam_sinh)[-2:]
    
    # Tạo số căn cước
    so_can_cuoc = f'{ma_tinh}{ma_gioi_tinh}{ma_nam_sinh}{random.randint(100000, 999999):06}'
    
    ho_ten = f'{ho} {ten_dem} {ten}'
    
    ws.append([ho_ten, f'{ngay_sinh:02d}/{thang_sinh:02d}/{nam_sinh}', gioi_tinh, so_can_cuoc])

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
