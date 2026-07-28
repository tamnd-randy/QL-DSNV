import sys
import pandas as pd

# Đảm bảo hiển thị Tiếng Việt trong Console Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

input_file = r'd:\DSNV\VTM-DSNV-07.2026.xlsx'
output_file = r'd:\DSNV\data.xlsx'

def loc_danh_sach_ten_t(file_in, file_out):
    # Đọc file nháp không có header để tự động tìm dòng tiêu đề
    df_raw = pd.read_excel(file_in, header=None)
    
    header_row = 0
    col_name = None
    possible_names = ['Họ và tên', 'Họ tên', 'Tên', 'HO VA TEN', 'Full Name']
    
    # Tìm dòng tiêu đề chứa các từ khóa tên
    for idx, row in df_raw.iterrows():
        row_values = [str(val).strip() for val in row.values if pd.notna(val)]
        for name_key in possible_names:
            if any(name_key.lower() == val.lower() for val in row_values):
                header_row = idx
                break
        if header_row != 0:
            break
            
    print(f"-> Đã phát hiện dòng tiêu đề cột tại dòng số: {header_row + 1}")
    
    # Đọc lại Excel từ dòng tiêu đề đã tìm thấy
    df = pd.read_excel(file_in, header=header_row)
    
    # Xác định chính xác cột Tên
    for col in df.columns:
        if str(col).strip().lower() in [name.lower() for name in possible_names]:
            col_name = col
            break
            
    if not col_name:
        print("Không tìm thấy cột 'Họ và tên'!")
        return

    print(f"-> Đã xác định cột Họ và tên: '{col_name}'")

    # Hàm lọc TÊN (từ cuối cùng của Họ và Tên) bắt đầu bằng chữ T (Tuấn, Trang, Thành, Tùng, Tiên...)
    def check_ten_t(full_name):
        if pd.isna(full_name):
            return False
        parts = str(full_name).strip().split()
        if not parts:
            return False
        first_name = parts[-1] # Lấy tên chính (từ cuối)
        return first_name.upper().startswith('T')

    # Thực hiện lọc
    df_filtered = df[df[col_name].apply(check_ten_t)]

    # Lưu kết quả xuất ra file Excel
    df_filtered.to_excel(file_out, index=False)
    print(f"-> ĐÃ LỌC THÀNH CÔNG: {len(df_filtered)} người có tên bắt đầu bằng chữ 'T'.")
    print(f"-> File đầu ra đã được tạo: {file_out}")

if __name__ == '__main__':
    loc_danh_sach_ten_t(input_file, output_file)
