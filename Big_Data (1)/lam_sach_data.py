import pandas as pd
import numpy as np

# Đọc dữ liệu từ file Excel
df = pd.read_excel('theodoitainha.xlsx', sheet_name='Sheet1')

# Kiểm tra dữ liệu ban đầu
print("Các cột trong DataFrame:", df.columns)
print("Kiểu dữ liệu các cột:\n", df.dtypes)
print("Số lượng hàng ban đầu:", len(df))

# Chuyển các cột có thể là Categorical sang object
for col in df.columns:
    if df[col].dtype.name == 'category':
        df[col] = df[col].astype('object')

# 1. Giới tính
df['Giới tính'] = df['Giới tính'].astype(str).str.strip().str.lower().replace({
    'nam': 'Nam', 'nữ': 'Nữ', 'nu': 'Nữ', 'female': 'Nữ', 'male': 'Nam'
})
df['Giới tính'] = df['Giới tính'].replace('nan', 'Không xác định').fillna('Không xác định')

# 2. Năm sinh
df['Năm sinh'] = pd.to_numeric(df['Năm sinh'], errors='coerce')
df.loc[(df['Năm sinh'] < 1900) | (df['Năm sinh'] > 2023), 'Năm sinh'] = np.nan
df['Năm sinh'] = df['Năm sinh'].fillna(df['Năm sinh'].mean())

# 3. Quốc tịch
df['Quốc tịch'] = df['Quốc tịch'].astype(str).str.strip().str.title().replace({
    'Vn': 'Việt Nam', 'Usa': 'Hoa Kỳ', 'Uk': 'Anh'
})
df['Quốc tịch'] = df['Quốc tịch'].replace('Nan', 'Không xác định').fillna('Không xác định')

# 4. Tỉnh/TP, Huyện, Phường/Xã
df['Tỉnh/TP'] = df['Tỉnh/TP'].astype(str).str.strip().str.title().replace({
    'Hn': 'Hà Nội', 'Hcm': 'Hồ Chí Minh'
})
df['Huyện'] = df['Huyện'].astype(str).str.strip().str.title()
df['Phường xã'] = df['Phường xã'].astype(str).str.strip().str.title()

for col in ['Tỉnh/TP', 'Huyện', 'Phường xã']:
    df[col] = df[col].replace('Nan', 'Không xác định').fillna('Không xác định')

# 6. Ngày chuẩn đoán xác định
df['Ngày chẩn đoán xác định'] = pd.to_datetime(df['Ngày chẩn đoán xác định'], dayfirst=True, errors='coerce')
df['Ngày chẩn đoán xác định'] = df['Ngày chẩn đoán xác định'].fillna(pd.Timestamp.now())

# 7. Đơn vị xác nhận
df['Đơn vị xác nhận'] = df['Đơn vị xác nhận'].astype(str).str.strip().str.title().replace({
    'Bv Chợ Rẫy': 'Bệnh viện Chợ Rẫy'
})
df['Đơn vị xác nhận'] = df['Đơn vị xác nhận'].replace('Nan', 'Không xác định').fillna('Không xác định')

# 8. Ngày bệnh thứ
df['Ngày bệnh thứ'] = pd.to_numeric(df['Ngày bệnh thứ'], errors='coerce')
df.loc[df['Ngày bệnh thứ'] < 0, 'Ngày bệnh thứ'] = 0
df['Ngày bệnh thứ'] = df['Ngày bệnh thứ'].fillna(0)

# 9. Tình trạng tiêm vắc xin
df['Tình trạng tiêm vắc xin'] = df['Tình trạng tiêm vắc xin'].astype(str).str.strip().str.lower().replace({
    'đã tiêm': 'Đã tiêm', 'chưa tiêm': 'Chưa tiêm', 'tiêm 1 mũi': 'Đã tiêm'
})
df['Tình trạng tiêm vắc xin'] = df['Tình trạng tiêm vắc xin'].replace('nan', 'Không rõ').fillna('Không rõ')

# 10. Thuộc đối tượng có thai
df['Thuộc đối tượng có thai'] = df['Thuộc đối tượng có thai'].astype(str).str.strip().str.lower().replace({
    'có': 'Có', 'không': 'Không'
})
df.loc[df['Giới tính'] == 'Nam', 'Thuộc đối tượng có thai'] = 'Không'
df['Thuộc đối tượng có thai'] = df['Thuộc đối tượng có thai'].replace('nan', 'Không').fillna('Không')

# 11. Nhóm tuổi
df['Tuổi'] = 2023 - df['Năm sinh']
df['Nhóm tuổi'] = pd.cut(df['Tuổi'], bins=[0, 18, 35, 60, 100], labels=['0-18', '19-35', '36-60', 'trên 60'])
df['Nhóm tuổi'] = df['Nhóm tuổi'].astype(str).replace('nan', 'Không xác định').fillna('Không xác định')

# 12. Bệnh nền
df['Bệnh nền'] = df['Bệnh nền'].astype(str).str.strip().str.title().replace({
    'Đái tháo đường': 'Tiểu đường', 'Cao HA': 'Cao huyết áp'
})
df['Bệnh nền'] = df['Bệnh nền'].replace('Nan', 'Không có').fillna('Không có')

# 13. Đơn vị theo dõi
df['Đơn vị theo dõi'] = df['Đơn vị theo dõi'].astype(str).str.strip().str.title().replace({
    'Ttyt Q.1': 'Trung tâm Y tế Quận 1'
})
df['Đơn vị theo dõi'] = df['Đơn vị theo dõi'].replace('Nan', 'Không xác định').fillna('Không xác định')

# 14. Trạng thái theo dõi
df['Trạng thái theo dõi'] = df['Trạng thái theo dõi'].astype(str).str.strip().str.title().replace({
    'Theo dõi': 'Đang theo dõi', 'Kết thúc': 'Đã kết thúc'
})
df['Trạng thái theo dõi'] = df['Trạng thái theo dõi'].replace('Nan', 'Không rõ').fillna('Không rõ')

# 15. Thời gian khai báo sức khỏe gần nhất
df['Thời gian khai báo sức khỏe gần nhất'] = pd.to_datetime(df['Thời gian khai báo sức khỏe gần nhất'], dayfirst=True, errors='coerce')
df['Thời gian khai báo sức khỏe gần nhất'] = df['Thời gian khai báo sức khỏe gần nhất'].fillna(pd.Timestamp.now())

# Kiểm tra dữ liệu trước khi lưu
print("\nDữ liệu mẫu của tất cả các cột:")
print(df.head(10))
print("Số lượng hàng sau khi làm sạch:", len(df))
try:
    df.to_excel('du_lieu_sach.xlsx', index=False, engine='openpyxl')
    print("Đã lưu file 'du_lieu_sach.xlsx' thành công bằng openpyxl!")
except Exception as e:
    print(f"Lỗi khi lưu file với openpyxl: {e}")