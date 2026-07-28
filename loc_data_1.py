import sys
import unicodedata
import pandas as pd

# Thiết lập UTF-8 cho Windows Console
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

input_file = r'd:\DSNV\VTM-DSNV-07.2026.xlsx'
output_file = r'd:\DSNV\data 1.xlsx'

# Danh sách 9 người cần lọc
target_names = [
    "Vũ Huy Thành",
    "Tô Văn Đạt",
    "Nguyễn Thị Thu Dung",
    "Nguyễn Thị Mai Ngọc",
    "Đào Ngọc Thủy",
    "Đậu Thị Hằng",
    "Phạm Văn Cường",
    "Dương Thị Khương",
    "Bùi Thị Miền"
]

def normalize_text(text):
    if pd.isna(text):
        return ""
    text_str = str(text).strip()
    text_str = unicodedata.normalize('NFC', text_str)
    return " ".join(text_str.upper().split())

def loc_danh_sach_theo_ten(file_in, file_out, list_names):
    df_raw = pd.read_excel(file_in, header=None)
    
    header_row = 0
    possible_names = ['Họ và tên', 'Họ tên', 'Tên', 'HO VA TEN', 'Full Name']
    
    for idx, row in df_raw.iterrows():
        row_values = [str(val).strip() for val in row.values if pd.notna(val)]
        for name_key in possible_names:
            if any(name_key.lower() == val.lower() for val in row_values):
                header_row = idx
                break
        if header_row != 0:
            break
            
    df = pd.read_excel(file_in, header=header_row)
    
    col_name = None
    for col in df.columns:
        if str(col).strip().lower() in [name.lower() for name in possible_names]:
            col_name = col
            break

    if not col_name:
        print("Không tìm thấy cột 'Họ và tên'!")
        return

    # Chuẩn hóa danh sách tên cần tìm
    normalized_target_set = {normalize_text(name) for name in list_names}

    # Hàm kiểm tra trùng khớp tên
    def matches_target(full_name):
        norm_name = normalize_text(full_name)
        return norm_name in normalized_target_set

    df_filtered = df[df[col_name].apply(matches_target)]

    # Lưu kết quả
    try:
        df_filtered.to_excel(file_out, index=False)
        print(f"-> Số người cần lọc: {len(target_names)}")
        print(f"-> Tìm thấy thành công: {len(df_filtered)} kết quả.")
    except PermissionError:
        print(f"[CẢNH BÁO] File '{file_out}' đang được mở bởi ứng dụng khác (Excel).")
        alt_output = r'd:\DSNV\data 1_new.xlsx'
        df_filtered.to_excel(alt_output, index=False)
        print(f"-> Đã lưu kết quả thay thế vào file: '{alt_output}'")
        return

    found_names = df_filtered[col_name].unique()
    print("Danh sách tìm thấy:")
    for name in found_names:
        print(f"  + {name}")

if __name__ == '__main__':
    loc_danh_sach_theo_ten(input_file, output_file, target_names)
