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
# Đảm bảo các cột cần thiết tồn tại và có tên chính xác
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

# Kiểm tra và xử lý dữ liệu trống trong cột 'Học Kỳ'
df['Học Kỳ'] = df['Học Kỳ'].fillna(0)

# Chuẩn bị dữ liệu cho biểu đồ Sunburst
try:
    df_grouped = df.groupby(['Học Kỳ', 'Loại môn học', 'Tên học phần']).size().reset_index(name='Số lượng')
except KeyError as e:
    print(f"Lỗi KeyError: Cột '{e}' không tồn tại. Vui lòng kiểm tra tên cột trong file Excel.")
    exit()

# Tạo biểu đồ Sunburst với Plotly
fig = px.sunburst(
    df_grouped,
    path=['Học Kỳ', 'Loại môn học', 'Tên học phần'],  # Thay đổi thứ tự: Học Kỳ -> Loại môn học -> Tên học phần
    values='Số lượng',
    title='Biểu đồ Nested Pie Chart của Học phần theo Học Kỳ, Loại và Tên Môn Học',
    color='Số lượng',
    color_continuous_scale='Blues'
)

# Cấu hình layout cho biểu đồ
fig.update_layout(
    margin=dict(t=50, l=25, r=25, b=25),
    title_font_size=20,
    width=800,
    height=800
)

# Xuất biểu đồ thành file HTML
fig.write_html('416_10k.html')

print("Đã tạo file HTML: 416_10k.html")