# Import các thư viện cần thiết
import pandas as pd
import plotly.express as px

# Đọc dữ liệu từ file Excel
df = pd.read_excel('dataset-416.xlsx')

# In danh sách cột để kiểm tra
print("Các cột trong DataFrame:", df.columns.tolist())

# Xem trước dữ liệu
print(df.head())

# Kiểm tra và đổi tên cột nếu cần
column_mapping = {}
if 'Language' in df.columns:
    column_mapping['Language'] = 'Ngôn ngữ'
elif 'ngôn ngữ' in df.columns:
    column_mapping['ngôn ngữ'] = 'Ngôn ngữ'
if 'Loại môn học ' in df.columns:
    column_mapping['Loại môn học '] = 'Loại môn học'
if 'Học kỳ' in df.columns:
    column_mapping['Học kỳ'] = 'Học Kỳ'
if 'Tên học phần ' in df.columns:
    column_mapping['Tên học phần '] = 'Tên học phần'

# Đổi tên cột
df.rename(columns=column_mapping, inplace=True)

# Kiểm tra lại cột sau khi đổi tên
print("Các cột sau khi đổi tên:", df.columns.tolist())

# Kiểm tra và xử lý dữ liệu trống
df['Học Kỳ'] = df['Học Kỳ'].fillna(0)
df['Loại môn học'] = df['Loại môn học'].fillna('Không xác định')
df['Tên học phần'] = df['Tên học phần'].fillna('Không xác định')

# Chuẩn bị dữ liệu cho biểu đồ Sunburst
try:
    df_grouped = df.groupby(['Học Kỳ', 'Loại môn học', 'Tên học phần']).size().reset_index(name='Số lượng')
except KeyError as e:
    print(f"Lỗi KeyError: Cột '{e}' không tồn tại. Vui lòng kiểm tra tên cột trong file Excel.")
    exit()

# Tạo biểu đồ Sunburst với Plotly
fig = px.sunburst(
    df_grouped,
    path=['Học Kỳ', 'Loại môn học', 'Tên học phần'],  # Vòng trong cùng: Học Kỳ -> Loại môn học -> Tên học phần
    values='Số lượng',
    title='Biểu đồ Nested Pie Chart của Học phần theo Học Kỳ, Loại và Tên Môn Học',
    color='Loại môn học',  # Gán màu theo Loại môn học (Bắt buộc/Tự chọn)
    color_discrete_map={  # Tùy chỉnh màu sắc
        'Bắt buộc': '#1f77b4',  # Màu xanh dương
        'Tự chọn': '#ff7f0e',   # Màu cam
        'Không xác định': '#d62728'  # Màu đỏ (cho dữ liệu trống)
    }
)

# Cấu hình layout cho biểu đồ
fig.update_layout(
    margin=dict(t=50, l=25, r=25, b=25),
    title_font_size=20,
    width=1000,  # Tăng kích thước để hiển thị chi tiết
    height=1000
)

# Xuất biểu đồ thành file HTML, nhúng Plotly để chạy offline
fig.write_html('416_10k.html', include_plotlyjs=True)

print("Đã tạo file HTML: 416_10k.html")