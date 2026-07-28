import sys
import locale
import pandas as pd

# Cấu hình UTF-8 cho Windows Console
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')
    try:
        locale.setlocale(locale.LC_ALL, 'Vietnamese_Vietnam.1258')
    except Exception:
        try:
            locale.setlocale(locale.LC_ALL, 'vi_VN')
        except Exception:
            pass

input_file = r'd:\DSNV\data 1_new.xlsx'
output_file = r'd:\DSNV\data 1.xlsx'
alt_output_file = r'd:\DSNV\data 1_new.xlsx'

def sap_xep_va_loc_cot(file_in):
    df = pd.read_excel(file_in)
    
    # 1. Xác định các cột cần giữ lại
    col_ma_nv = 'Mã nhân viên'
    col_ho_ten = 'Họ và tên'
    col_chuc_danh = 'Tên chức danh nghề nghiệp'  # Cột chứa thông tin chức danh đầy đủ
    col_sdt = 'Điện thoại'
    
    # Đảm bảo tồn tại đủ các cột
    selected_cols = [col_ma_nv, col_ho_ten, col_chuc_danh, col_sdt]
    df_result = df[selected_cols].copy()
    
    # Đổi tên cột cho chuẩn theo yêu cầu của user
    df_result.rename(columns={
        col_ho_ten: 'Họ và tên',
        col_chuc_danh: 'Chức danh',
        col_sdt: 'Số điện thoại'
    }, inplace=True)
    
    # 2. Hàm tạo khóa sắp xếp chuẩn ABC Tiếng Việt (Tên -> Họ -> Tên đệm)
    def vietnamese_sort_key(full_name):
        if pd.isna(full_name):
            return ''
        parts = str(full_name).strip().split()
        if not parts:
            return ''
        # Thứ tự ưu tiên sắp xếp: Tên chính (từ cuối) -> Họ (từ đầu) -> Các từ đệm
        ordered = [parts[-1]] + [parts[0]] + parts[1:-1]
        return [locale.strxfrm(p.lower()) for p in ordered]

    # 3. Sắp xếp danh sách theo bảng chữ cái ABC Tiếng Việt
    df_result['sort_key'] = df_result['Họ và tên'].apply(vietnamese_sort_key)
    df_result.sort_values(by='sort_key', inplace=True)
    df_result.drop(columns=['sort_key'], inplace=True)

    # 4. Lưu ra file Excel
    saved_path = output_file
    try:
        df_result.to_excel(output_file, index=False)
        print(f"-> Đã sắp xếp và lưu thành công vào file: '{output_file}'")
    except PermissionError:
        print(f"[CẢNH BÁO] File '{output_file}' đang mở. Lưu thay thế vào '{alt_output_file}'")
        df_result.to_excel(alt_output_file, index=False)
        saved_path = alt_output_file

    print("\nDanh sách sau khi sắp xếp ABC:")
    print(df_result.to_string(index=False))

if __name__ == '__main__':
    sap_xep_va_loc_cot(input_file)
