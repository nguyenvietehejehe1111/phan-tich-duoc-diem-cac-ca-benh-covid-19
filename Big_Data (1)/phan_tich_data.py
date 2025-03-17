import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Đọc dữ liệu từ file Excel đã làm sạch
df = pd.read_excel('du_lieu_sach.xlsx', sheet_name='Sheet1')

# Kiểm tra dữ liệu
print("Các cột trong DataFrame:", df.columns)
print(df.head())

# --- Phân tích theo giới tính ---
# Đếm số lượng theo giới tính
gender_counts = df['Giới tính'].value_counts()
print("\nPhân bố theo giới tính:")
print(gender_counts)

# --- Phân tích theo độ tuổi ---
# Nếu đã có cột 'Nhóm tuổi', sử dụng trực tiếp
if 'Nhóm tuổi' in df.columns:
    age_group_counts = df['Nhóm tuổi'].value_counts()
    print("\nPhân bố theo nhóm tuổi:")
    print(age_group_counts)
else:
    # Nếu chỉ có 'Năm sinh', tính tuổi và phân nhóm
    current_year = 2025  # Cập nhật năm hiện tại (ngày 16/03/2025)
    df['Tuổi'] = current_year - df['Năm sinh']
    df['Nhóm tuổi'] = pd.cut(df['Tuổi'], bins=[0, 18, 35, 60, 100], labels=['0-18', '19-35', '36-60', 'trên 60'])
    age_group_counts = df['Nhóm tuổi'].value_counts()
    print("\nPhân bố theo nhóm tuổi (tính từ Năm sinh):")
    print(age_group_counts)

# --- Phân tích theo Huyện ---
# Đếm số lượng ca theo Huyện
huyen_counts = df['Huyện'].value_counts()
print("\nPhân bố theo Huyện:")
print(huyen_counts)

# --- Phân tích theo Phường xã ---
# Đếm số lượng ca theo Phường xã
phuong_xa_counts = df['Phường xã'].value_counts()
print("\nPhân bố theo Phường xã:")
print(phuong_xa_counts)

# --- Trực quan hóa dữ liệu ---
# Thiết lập giao diện biểu đồ
sns.set_style("whitegrid")  # Sử dụng kiểu của seaborn
plt.rcParams['font.family'] = 'Arial'  # Đặt font chữ

# 1. Biểu đồ cột cho Giới tính
plt.figure(figsize=(8, 6))
sns.barplot(x=gender_counts.index, y=gender_counts.values, palette='viridis')
plt.title('Phân bố theo Giới tính', fontsize=14)
plt.xlabel('Giới tính', fontsize=12)
plt.ylabel('Số lượng', fontsize=12)
plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig('phan_bo_gioi_tinh.png')  # Lưu biểu đồ
plt.show()

# 2. Biểu đồ cột cho Nhóm tuổi
plt.figure(figsize=(10, 6))
sns.barplot(x=age_group_counts.index, y=age_group_counts.values, palette='magma')
plt.title('Phân bố theo Nhóm tuổi', fontsize=14)
plt.xlabel('Nhóm tuổi', fontsize=12)
plt.ylabel('Số lượng', fontsize=12)
plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig('phan_bo_nhom_tuoi.png')  # Lưu biểu đồ
plt.show()

# 3. Biểu đồ kết hợp Giới tính và Nhóm tuổi (pivot table)
pivot_table = df.pivot_table(index='Nhóm tuổi', columns='Giới tính', aggfunc='size', fill_value=0)
plt.figure(figsize=(12, 8))
sns.heatmap(pivot_table, annot=True, fmt='d', cmap='YlGnBu')
plt.title('Phân bố theo Giới tính và Nhóm tuổi', fontsize=14)
plt.xlabel('Giới tính', fontsize=12)
plt.ylabel('Nhóm tuổi', fontsize=12)
plt.tight_layout()
plt.savefig('phan_bo_gioi_tinh_nhom_tuoi.png')  # Lưu biểu đồ
plt.show()

# 4. Biểu đồ cột cho Huyện (chỉ hiển thị top 10 để tránh quá tải)
top_huyen = huyen_counts.head(10)
plt.figure(figsize=(12, 6))
sns.barplot(x=top_huyen.index, y=top_huyen.values, palette='viridis')
plt.title('Phân bố ca bệnh theo Huyện (Top 10)', fontsize=14)
plt.xlabel('Huyện', fontsize=12)
plt.ylabel('Số lượng ca', fontsize=12)
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig('phan_bo_huyen.png')
plt.show()

# 5. Biểu đồ cột cho Phường xã (chỉ hiển thị top 10 để tránh quá tải)
top_phuong_xa = phuong_xa_counts.head(10)
plt.figure(figsize=(12, 6))
sns.barplot(x=top_phuong_xa.index, y=top_phuong_xa.values, palette='magma')
plt.title('Phân bố ca bệnh theo Phường xã (Top 10)', fontsize=14)
plt.xlabel('Phường xã', fontsize=12)
plt.ylabel('Số lượng ca', fontsize=12)
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig('phan_bo_phuong_xa.png')
plt.show()

# 6. Biểu đồ nhiệt kết hợp Huyện và Phường xã (chỉ lấy top 5 Huyện và top 10 Phường xã)
top_5_huyen = huyen_counts.head(5).index
top_10_phuong_xa = phuong_xa_counts.head(10).index
filtered_df = df[df['Huyện'].isin(top_5_huyen) & df['Phường xã'].isin(top_10_phuong_xa)]
pivot_table_huyen_phuong = filtered_df.pivot_table(index='Huyện', columns='Phường xã', aggfunc='size', fill_value=0)
plt.figure(figsize=(14, 8))
sns.heatmap(pivot_table_huyen_phuong, annot=True, fmt='d', cmap='YlGnBu')
plt.title('Phân bố ca bệnh theo Huyện và Phường xã (Top 5 Huyện, Top 10 Phường xã)', fontsize=14)
plt.xlabel('Phường xã', fontsize=12)
plt.ylabel('Huyện', fontsize=12)
plt.tight_layout()
plt.savefig('phan_bo_huyen_phuong_xa.png')
plt.show()