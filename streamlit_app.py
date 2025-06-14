import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle, Alignment
from datetime import datetime
import io
import os
import re # Import regex module

# --- Cấu hình trang Streamlit ---
st.set_page_config(layout="wide", page_title="Đồng bộ dữ liệu SSE") # Changed to wide layout for more space

# Đường dẫn đến các file cần thiết (giả định cùng thư mục với script)
LOGO_PATH = "Logo.png"
DATA_FILE_PATH = "Data.xlsx" # Tên chính xác của file dữ liệu

# Định nghĩa tiêu đề cho file UpSSE.xlsx (Di chuyển lên đây để luôn có sẵn)
headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng",
           "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế",
           "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm",
           "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế",
           "Nhóm Hàng", "Ghi chú", "Tiền thuế"]

# --- Hàm trợ giúp chuyển đổi giá trị sang float an toàn ---
def to_float(value):
    """Chuyển đổi giá trị sang float, trả về 0.0 nếu không thể chuyển đổi."""
    try:
        # Xóa tất cả các khoảng trắng và dấu phẩy (nếu là định dạng số có dấu phẩy)
        if isinstance(value, str):
            value = value.replace(",", "").strip()
        return float(value)
    except (ValueError, TypeError):
        return 0.0

# --- Hàm làm sạch chuỗi (loại bỏ mọi loại khoảng trắng và chuẩn hóa) ---
def clean_string(s):
    if s is None:
        return ""
    # Thay thế tất cả các loại khoảng trắng (bao gồm non-breaking space, tabs, newlines) bằng một khoảng trắng đơn
    s = re.sub(r'\s+', ' ', str(s)).strip()
    return s

# --- Hàm đọc dữ liệu tĩnh và bảng tra cứu từ Data.xlsx ---
@st.cache_data
def get_static_data_from_excel(file_path):
    """
    Đọc dữ liệu và xây dựng các bảng tra cứu từ Data.xlsx.
    Sử dụng openpyxl để đọc dữ liệu. Kết quả được cache.
    """
    try:
        wb = load_workbook(file_path, data_only=True)
        ws = wb.active

        listbox_data = [] # Dữ liệu cho combobox chọn CHXD (cột K)
        chxd_detail_map = {} # Map để lưu thông tin chi tiết CHXD (G5, H5, F5, B5)
        
        # Đọc dữ liệu từ bảng CHXD để xây dựng listbox_data và chxd_detail_map
        # Giả định:
        # - Tên CHXD ở cột K (index 10)
        # - Mã khách (G5) ở cột P (index 15)
        # - Mã kho (F5) ở cột Q (index 16)
        # - Khu vực (H5) ở cột S (index 17)
        # Bắt đầu đọc từ hàng 4 (index 3)
        for row_idx in range(4, ws.max_row + 1):
            row_data_values = [cell.value for cell in ws[row_idx]]

            if len(row_data_values) >= 18: # Đảm bảo đủ cột để truy cập index 17 (cột S)
                raw_chxd_name = row_data_values[10] # Cột K (index 10)

                if raw_chxd_name is not None and clean_string(raw_chxd_name) != '':
                    chxd_name_str = clean_string(raw_chxd_name)
                    
                    if chxd_name_str and chxd_name_str not in listbox_data: # Tránh trùng lặp trong listbox
                        listbox_data.append(chxd_name_str)

                    # Lấy các giá trị cho G5, H5, F5_full, B5 từ các cột tương ứng
                    g5_val = row_data_values[15] if len(row_data_values) > 15 and pd.notna(row_data_values[15]) else None # Cột P (index 15)
                    f5_val_full = clean_string(row_data_values[16]) if len(row_data_values) > 16 and pd.notna(row_data_values[16]) else '' # Cột Q (index 16)
                    h5_val = clean_string(row_data_values[17]).lower() if len(row_data_values) > 17 and pd.notna(row_data_values[17]) else '' # Cột S (index 17)
                    b5_val = chxd_name_str # B5 chính là tên CHXD

                    if f5_val_full: # Chỉ thêm vào map nếu Mã kho có giá trị
                        chxd_detail_map[chxd_name_str] = {
                            'g5_val': g5_val,
                            'h5_val': h5_val,
                            'f5_val_full': f5_val_full,
                            'b5_val': b5_val
                        }
        
        # Đọc các bảng tra cứu khác theo phạm vi cụ thể
        # I4:J7 (Mã hàng <-> Tên mặt hàng)
        lookup_table = {}
        for row in ws.iter_rows(min_row=4, max_row=7, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                lookup_table[clean_string(row[0]).lower()] = row[1]
        
        # I10:J13 (TMT Lookup table - Tên mặt hàng <-> Mức phí BVMT)
        tmt_lookup_table = {}
        for row in ws.iter_rows(min_row=10, max_row=13, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                tmt_lookup_table[clean_string(row[0]).lower()] = to_float(row[1]) # Chuyển sang float
        
        # I29:J31 (S Lookup table - Khu vực <-> Tk nợ)
        s_lookup_table = {}
        for row in ws.iter_rows(min_row=29, max_row=31, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                s_lookup_table[clean_string(row[0]).lower()] = row[1]
        
        # I33:J35 (T Lookup table - Khu vực <-> Tk doanh thu)
        t_lookup_table = {}
        for row in ws.iter_rows(min_row=33, max_row=35, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                t_lookup_table[clean_string(row[0]).lower()] = row[1]
        
        # I53:J55 (V Lookup table - Khu vực <-> Tk thuế có)
        v_lookup_table = {}
        for row in ws.iter_rows(min_row=53, max_row=55, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                v_lookup_table[clean_string(row[0]).lower()] = row[1]
        
        # I17:J20 (X Lookup table - Tên mặt hàng <-> Mã vụ việc)
        x_lookup_table = {}
        for row in ws.iter_rows(min_row=17, max_row=20, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                x_lookup_table[clean_string(row[0]).lower()] = row[1]

        # Đọc giá trị J36 (u_value)
        u_value = ws['J36'].value

        wb.close()
        
        return {
            "listbox_data": listbox_data,
            "lookup_table": lookup_table,
            "tmt_lookup_table": tmt_lookup_table,
            "s_lookup_table": s_lookup_table,
            "t_lookup_table": t_lookup_table,
            "v_lookup_table": v_lookup_table,
            "x_lookup_table": x_lookup_table,
            "u_value": u_value,
            "chxd_detail_map": chxd_detail_map
        }
    except FileNotFoundError:
        st.error(f"Lỗi: Không tìm thấy file {file_path}. Vui lòng đảm bảo file tồn tại.")
        st.stop()
    except Exception as e:
        st.error(f"Lỗi không mong muốn khi đọc file Data.xlsx: {e}")
        st.exception(e)
        st.stop()

# --- Hàm thêm dòng tổng hợp cho "Người mua không lấy hóa đơn" ---
def add_summary_row_for_no_invoice(data_for_summary_product, bkhd_source_ws, product_name, headers_list,
                                   g5_val, b5_val, s_lookup, t_lookup, v_lookup, x_lookup, u_val, h5_val, common_lookup_table):
    """
    Tạo một dòng tổng hợp cho "Người mua không lấy hóa đơn" cho một mặt hàng cụ thể.
    data_for_summary_product: Danh sách các dòng thô đã được xử lý khớp với tiêu chí "Người mua không lấy hóa đơn" cho mặt hàng này.
    """
    new_row = [''] * len(headers_list)
    new_row[0] = g5_val
    new_row[1] = f"Khách hàng mua {product_name} không lấy hóa đơn"

    # Lấy giá trị Ngày (C) và Ký hiệu (E) từ dòng đầu tiên có liên quan trong nhóm này
    c_val_from_first_row = None
    e_val_from_first_row = None
    if data_for_summary_product:
        # Giả định data_for_summary_product chứa các dòng đã xử lý (như new_row_for_upsse)
        # trong đó cột C là chỉ mục 2 và cột E là chỉ mục 4
        c_val_from_first_row = data_for_summary_product[0][2] 
        e_val_from_first_row = data_for_summary_product[0][4] 

    c_val_from_first_row = c_val_from_first_row if c_val_from_first_row is not None else ""
    e_val_from_first_row = e_val_from_first_row if e_val_from_first_row is not None else ""

    new_row[2] = c_val_from_first_row # Cột C (Ngày)
    new_row[4] = e_val_from_first_row # Cột E (Ký hiệu)

    value_C = clean_string(new_row[2])
    value_E = clean_string(new_row[4])

    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name, "")

    if b5_val == "Nguyễn Huệ":
        value_D = f"HNBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh":
        value_D = f"MMBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    else:
        value_D = f"{value_E[-2:]}BK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    new_row[3] = value_D # Cột D (Số hóa đơn)

    new_row[5] = f"Xuất bán lẻ theo hóa đơn số " + new_row[3] # Cột F (Diễn giải)
    new_row[7] = product_name # Cột H (Tên mặt hàng)
    new_row[6] = common_lookup_table.get(clean_string(product_name).lower(), '') # Cột G (Mã hàng)
    new_row[8] = "Lít" # Cột I (Đvt)
    new_row[9] = g5_val # Cột J (Mã kho)
    new_row[10] = '' # Cột K (Mã vị trí)
    new_row[11] = '' # Cột L (Mã lô)

    # Tính tổng số lượng (cột M) cho các mục không có hóa đơn của sản phẩm này
    total_M = sum(to_float(r[12]) for r in data_for_summary_product)
    new_row[12] = total_M # Cột M (Số lượng)
    
    # Tính Giá bán lớn nhất (cột N) cho các mục không có hóa đơn của sản phẩm này
    max_value_N = 0.0
    if data_for_summary_product:
        max_value_N = max(to_float(r[13]) for r in data_for_summary_product) # r[13] là 'Giá bán' (cột N)
    new_row[13] = max_value_N # Cột N (Giá bán)

    # Tính tổng số tiền (Tiền hàng) dựa trên BKHD gốc cho các dòng không có hóa đơn
    # Tổng này cần lấy từ BKHD gốc như logic của UpSSE.2025.py
    tien_hang_hd_from_bkhd_original = sum(to_float(r[11]) for r in bkhd_source_ws.iter_rows(min_row=2, max_row=bkhd_source_ws.max_row, values_only=True)
                                         if clean_string(r[5]) == "Người mua không lấy hóa đơn" and clean_string(r[8]) == product_name)
    price_per_liter_map = {"Xăng E5 RON 92-II": 1900, "Xăng RON 95-III": 2000, "Dầu DO 0,05S-II": 1000, "Dầu DO 0,001S-V": 1000}
    current_price_per_liter = price_per_liter_map.get(product_name, 0)
    
    # Tính Tiền hàng (Cột O) cho dòng tổng hợp
    new_row[14] = tien_hang_hd_from_bkhd_original - round(total_M * current_price_per_liter, 0) # Cột O (Tiền hàng)

    new_row[15] = '' # Mã nt
    new_row[16] = '' # Tỷ giá
    new_row[17] = 10 # Mã thuế (Cột R)

    new_row[18] = s_lookup.get(h5_val, '') # Tk nợ (Cột S)
    new_row[19] = t_lookup.get(h5_val, '') # Tk doanh thu (Cột T)
    new_row[20] = u_val # Tk giá vốn (Cột U)
    new_row[21] = v_lookup.get(h5_val, '') # Tk thuế có (Cột V)
    new_row[22] = '' # Cục thuế (Cột W)
    new_row[23] = x_lookup.get(clean_string(product_name).lower(), '') # Vụ việc (Cột X)

    new_row[24] = '' # Bộ phận
    new_row[25] = '' # Lsx
    new_row[26] = '' # Sản phẩm
    new_row[27] = '' # Hợp đồng
    new_row[28] = '' # Phí
    new_row[29] = '' # Khế ước
    new_row[30] = '' # Nhân viên bán

    new_row[31] = f"Khách mua {product_name} không lấy hóa đơn" # Tên KH(thuế) (Cột AF)
    new_row[32] = "" # Địa chỉ (thuế) (Cột AG) - Không có trong file gốc
    new_row[33] = "" # Mã số Thuế (Cột AH) - Không có trong file gốc
    new_row[34] = '' # Nhóm Hàng
    new_row[35] = '' # Ghi chú

    # Tính Tiền thuế (Cột AK) cho dòng tổng hợp
    tien_thue_hd_from_bkhd_original = sum(to_float(r[12]) for r in bkhd_source_ws.iter_rows(min_row=2, max_row=bkhd_source_ws.max_row, values_only=True)
                                         if clean_string(r[5]) == "Người mua không lấy hóa đơn" and clean_string(r[8]) == product_name)
    new_row[36] = tien_thue_hd_from_bkhd_original - round(total_M * current_price_per_liter * 0.1, 0) # Cột AK (Tiền thuế)

    return new_row

# --- Hàm tạo dòng TMT theo hóa đơn ---
def create_per_invoice_tmt_row(original_row_data, tmt_value, headers_list, g5_val, s_lookup, t_lookup, v_lookup, u_val, h5_val):
    """
    Tạo một dòng TMT dựa trên một dòng hóa đơn gốc.
    """
    tmt_row = list(original_row_data) # Bắt đầu bằng một bản sao của dòng gốc

    # Áp dụng các phép biến đổi TMT cho các cột cụ thể dựa trên logic của UpSSE.2025.py
    # Mã hàng (Cột G)
    tmt_row[6] = "TMT" 
    # Tên mặt hàng (Cột H)
    tmt_row[7] = "Thuế bảo vệ môi trường" 
    # Đvt (Cột I)
    tmt_row[8] = "VNĐ" 
    # Mã kho (Cột J)
    tmt_row[9] = g5_val 
    # Giá bán (Cột N)
    tmt_row[13] = tmt_value 
    # Tiền hàng (Cột O) = tmt_value * Số lượng (Cột M)
    tmt_row[14] = round(tmt_value * to_float(original_row_data[12]), 0) 
    # Mã thuế (Cột R) - xóa nó
    tmt_row[17] = '' 
    # Tk nợ (Cột S)
    tmt_row[18] = s_lookup.get(h5_val, '') 
    # Tk doanh thu (Cột T)
    tmt_row[19] = t_lookup.get(h5_val, '') 
    # Tk giá vốn (Cột U)
    tmt_row[20] = u_val 
    # Tk thuế có (Cột V)
    tmt_row[21] = v_lookup.get(h5_val, '') 
    # Vụ việc (Cột X) - xóa nó
    tmt_row[23] = '' 
    # Tên KH(thuế) (Cột AF) - Đặt rõ ràng là trống cho các dòng TMT
    tmt_row[31] = "" 
    # Tiền thuế (Cột AK) = tmt_value * Số lượng (Cột M) * 0.1
    tmt_row[36] = round(tmt_value * to_float(original_row_data[12]) * 0.1, 0) 

    # Xóa các trường không liên quan khác cho dòng TMT (điều chỉnh chỉ mục nếu cần dựa trên lý trí)
    for idx in [5, 10, 11, 15, 16, 22, 24, 25, 26, 27, 28, 29, 30, 32, 33, 34, 35]:
        if idx < len(tmt_row): # Đảm bảo chỉ mục nằm trong giới hạn
            tmt_row[idx] = ''
    
    return tmt_row

# --- Hàm thêm dòng tổng hợp thuế bảo vệ môi trường (TMT) (cho các bản tóm tắt không có hóa đơn) ---
def add_tmt_summary_row(product_name_full, total_bvmt_amount, headers_list, g5_val, s_lookup, t_lookup, v_lookup, u_val, h5_val, 
                        representative_date, representative_symbol, total_quantity_for_tmt, tmt_unit_value_for_summary, b5_val, customer_name_for_summary_row):
    """
    Tạo một dòng tổng hợp cho Thuế bảo vệ môi trường (đặc biệt cho các bản tóm tắt không có hóa đơn).
    Bây giờ nhận ngày đại diện, ký hiệu, tổng số lượng và giá trị đơn vị TMT để tính toán.
    """
    new_tmt_row = [''] * len(headers_list)
    new_tmt_row[0] = g5_val # Mã khách
    
    # "Tên khách hàng" (Cột B) - Lấy từ customer_name_for_summary_row
    new_tmt_row[1] = customer_name_for_summary_row
    
    # Điền Ngày (C), Số hóa đơn (D), Ký hiệu (E)
    new_tmt_row[2] = representative_date # Cột C (Ngày)
    
    value_C = clean_string(representative_date)
    value_E = clean_string(representative_symbol)
    
    # Tạo Số hóa đơn (D) - Tương tự dòng tổng hợp, nhưng cho tổng hợp TMT
    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name_full, "") # Sử dụng product_name_full làm khóa

    if b5_val == "Nguyễn Huệ":
        value_D_tmt_summary = f"HNBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh":
        value_D_tmt_summary = f"MMBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    else:
        value_D_tmt_summary = f"{value_E[-2:]}BK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    new_tmt_row[3] = value_D_tmt_summary # Cột D (Số hóa đơn)

    new_tmt_row[4] = representative_symbol # Cột E (Ký hiệu)

    new_tmt_row[6] = "TMT" # Mã hàng
    new_tmt_row[7] = "Thuế bảo vệ môi trường" # Tên mặt hàng
    new_tmt_row[8] = "VNĐ" # Đvt
    new_tmt_row[9] = g5_val # Mã kho (giống Mã khách)
    
    # Điền Số lượng (M), Giá bán (N), Tiền hàng (O)
    new_tmt_row[12] = total_quantity_for_tmt # Cột M (Số lượng)
    new_tmt_row[13] = tmt_unit_value_for_summary # Cột N (Giá bán)
    new_tmt_row[14] = round(to_float(total_quantity_for_tmt) * to_float(tmt_unit_value_for_summary), 0) # Cột O (Tiền hàng)

    # Tài khoản
    new_tmt_row[18] = s_lookup.get(h5_val, '') # Tk nợ
    new_tmt_row[19] = t_lookup.get(h5_val, '') # Tk doanh thu
    new_tmt_row[20] = u_val # Tk giá vốn
    new_tmt_row[21] = v_lookup.get(h5_val, '') # Tk thuế có

    # Tiền thuế (Cột AK)
    new_tmt_row[36] = total_bvmt_amount 
    
    # Tên KH(thuế) (Cột AF) - Đặt rõ ràng là trống cho các dòng TMT
    new_tmt_row[31] = ""

    # Xóa các trường không liên quan khác cho dòng tổng hợp TMT
    # Mã thuế (R - chỉ mục 17) phải trống cho dòng TMT như mẫu gốc
    for idx in [5,10,11,15,16,17,22,23,24,25,26,27,28,29,30,32,33,34,35]:
        if idx < len(new_tmt_row):
            new_tmt_row[idx] = ''
    
    return new_tmt_row


# Tải dữ liệu tĩnh và bản đồ tra cứu từ Data.xlsx
static_data = get_static_data_from_excel(DATA_FILE_PATH)
listbox_data = static_data["listbox_data"]
lookup_table = static_data["lookup_table"]
tmt_lookup_table = static_data["tmt_lookup_table"]
s_lookup_table = static_data["s_lookup_table"]
t_lookup_table = static_data["t_lookup_table"]
v_lookup_table = static_data["v_lookup_table"]
x_lookup_table = static_data["x_lookup_table"]
u_value = static_data["u_value"]
chxd_detail_map = static_data["chxd_detail_map"] 

# --- Giao diện người dùng Streamlit ---
# Dùng st.container để bọc nội dung và căn giữa
with st.container():
    # Căn giữa nội dung chính
    st.markdown(
        """
        <style>
        .container {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 2rem;
            max-width: 800px; /* Giới hạn chiều rộng để cân đối */
            margin: auto; /* Căn giữa theo chiều ngang */
        }
        .stButton>button {
            background-color: #FF69B4; /* Pink */
            color: white;
            border-radius: 0.5rem;
            padding: 0.75rem 1.5rem;
            font-weight: bold;
            border: none;
            cursor: pointer;
            transition: background-color 0.3s ease;
        }
        .stButton>button:hover {
            background-color: #FF1493; /* Deeper Pink */
        }
        .stDownloadButton>button {
            background-color: #FF69B4; /* Pink */
            color: white;
            border-radius: 0.5rem;
            padding: 0.75rem 1.5rem;
            font-weight: bold;
            border: none;
            cursor: pointer;
            transition: background-color 0.3s ease;
        }
        .stDownloadButton>button:hover {
            background-color: #FF1493; /* Deeper Pink */
        }
        .stFileUploader > div > div > button {
            background-color: #FF69B4;
            color: white;
            border-radius: 0.5rem;
        }
        .stFileUploader > div > div > button:hover {
            background-color: #FF1493;
        }
        /* Define the blinking animation */
        @keyframes blink-important-note {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.2; } /* Makes it fade more noticeably */
        }
        /* Style for the important note box */
        .important-note-box {
            font-size: 1.5em; /* Larger font size: 1.5 times the default */
            font-weight: bold; /* Make it bold */
            color: red; /* Text color red */
            background-color: yellow; /* Background color yellow */
            text-align: center; /* Center the text */
            animation: blink-important-note 1s step-start 0s infinite; /* Apply blinking animation: 1s duration, instant steps, infinite loop */
            padding: 15px; /* Add some padding for better visual appearance */
            border-radius: 8px; /* Slightly rounded corners for the box */
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2); /* Add a subtle shadow */
            margin-bottom: 1rem; /* Add some space below the box */
        }
        /* New CSS for the header container */
        .header-main-flex-container {
            display: flex;
            align-items: center; /* Vertically align items in the center */
            justify-content: center; /* Center content horizontally */
            width: 100%; /* Take full width */
            margin-bottom: 1rem; /* Space below the header */
        }
        .header-text-column {
            display: flex;
            flex-direction: column;
            justify-content: center; /* Center text vertically */
            text-align: left; /* Align text to the left within its container */
            padding-left: 20px; /* Space between logo and text */
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # Logo và tiêu đề
    # Use an outer flex container to center the columns block
    st.markdown('<div class="header-main-flex-container">', unsafe_allow_html=True)
    # Use st.columns to place logo and text side-by-side within the centered container
    # Adjusted ratio for better balance and logo visibility
    logo_col, text_col = st.columns([1.5, 5]) 

    with logo_col:
        # Using st.image for reliable logo display
        st.image(LOGO_PATH, width=160) # Adjusted logo width for balance

    with text_col:
        st.markdown(
            """
            <div class="header-text-column">
                <h1 style="color: red; font-size: 26px; margin: 0px;">CÔNG TY CỔ PHẦN XĂNG DẦU</h1>
                <h2 style="color: red; font-size: 26px; margin: 0px;">DẦU KHÍ NAM ĐỊNH</h2>
            </div>
            """,
            unsafe_allow_html=True
        )
    st.markdown('</div>', unsafe_allow_html=True) # Close the outer flex container

    st.title("Đồng bộ dữ liệu SSE")

    # Bổ sung nội dung lưu ý quan trọng tại đây bằng HTML tùy chỉnh
    st.markdown(
        """
        <div class="important-note-box">
            Lưu ý quan trọng: Bạn cần mở file bảng kê hóa đơn, lưu lại (Ấn phím Ctrl+S hoặc vào File/save) trước khi chạy ứng dụng.
        </div>
        """,
        unsafe_allow_html=True
    )

    selected_value = st.selectbox(
        "Chọn CHXD:",
        options=[""] + listbox_data, # Thêm lựa chọn trống để khuyến khích người dùng chọn
        key='selected_chxd'
    )

    uploaded_file = st.file_uploader("Tải lên file bảng kê hóa đơn (.xlsx)", type=["xlsx"])

    # --- Xử lý chính khi nút "Xử lý" được bấm ---
    if st.button("Xử lý", key='process_button'):
        if not selected_value:
            st.warning("Vui lòng chọn một giá trị từ danh sách CHXD.")
        elif uploaded_file is None:
            st.warning("Vui lòng tải lên file bảng kê hóa đơn.")
        else:
            try:
                selected_value_normalized = clean_string(selected_value)

                if selected_value_normalized not in chxd_detail_map:
                    st.error("Không tìm thấy thông tin chi tiết cho CHXD đã chọn trong Data.xlsx. Vui lòng kiểm tra lại tên CHXD.")
                    st.error(f"Debug Info: Giá trị CHXD đã chọn: '{selected_value_normalized}'")
                    st.error(f"Debug Info: Các CHXD có trong map: {list(chxd_detail_map.keys())}")
                    st.stop()
                
                # Lấy các giá trị động từ chxd_detail_map
                chxd_details = chxd_detail_map[selected_value_normalized]
                g5_value = chxd_details['g5_val']
                h5_value = chxd_details['h5_val']
                f5_value_full = chxd_details['f5_val_full'] 
                b5_value = chxd_details['b5_val'] 

                # Đọc file bảng kê hóa đơn từ dữ liệu đã tải lên
                bkhd_wb = load_workbook(uploaded_file)
                bkhd_ws = bkhd_wb.active # bkhd_ws sẽ là worksheet gốc

                # Kiểm tra độ dài địa chỉ (cột H)
                long_cells = []
                for r_idx, cell in enumerate(bkhd_ws['H']):
                    if cell.value is not None and len(str(cell.value)) > 128:
                        long_cells.append(f"H{r_idx+1}")
                if long_cells:
                    st.error("Địa chỉ trên ô " + ', '.join(long_cells) + " quá dài, hãy điều chỉnh và thử lại.")
                    st.stop()

                # --- Chuẩn bị dữ liệu bảng kê: xóa 3 hàng đầu và sắp xếp lại cột ---
                temp_bkhd_data_for_processing = []
                for row in bkhd_ws.iter_rows(min_row=1, values_only=True):
                    temp_bkhd_data_for_processing.append(list(row)) 

                # Người dùng chỉ định dữ liệu bắt đầu từ dòng 5, nghĩa là các dòng 1-4 là tiêu đề/siêu dữ liệu
                if len(temp_bkhd_data_for_processing) >= 4: # Thay đổi từ 3 thành 4, để đảm bảo chúng ta bỏ qua 4 dòng nếu có
                    temp_bkhd_data_for_processing = temp_bkhd_data_for_processing[4:] # Thay đổi từ 3 thành 4 (danh sách dựa trên chỉ mục 0 nghĩa là bỏ qua các chỉ mục 0, 1, 2, 3)
                else:
                    temp_bkhd_data_for_processing = []

                temp_bkhd_ws_processed = Workbook().active
                for row_data in temp_bkhd_data_for_processing:
                    temp_bkhd_ws_processed.append(row_data)

                vi_tri_cu_idx = [0, 1, 2, 3, 4, 5, 7, 6, 8, 10, 11, 13, 14, 16]

                intermediate_data_rows = [] 
                for row_idx, row_values_original in enumerate(temp_bkhd_ws_processed.iter_rows(min_row=1, values_only=True)):
                    new_row_for_temp = [''] * (len(vi_tri_cu_idx) + 1) 
                    if len(row_values_original) <= max(vi_tri_cu_idx): 
                        continue 

                    for idx_new_col, col_old_idx in enumerate(vi_tri_cu_idx):
                        cell_value = row_values_original[col_old_idx] 

                        if idx_new_col == 3 and cell_value: 
                            cell_value_str = str(cell_value)[:10] 
                            try:
                                date_obj = datetime.strptime(cell_value_str, '%d-%m-%Y')
                                cell_value = date_obj.strftime('%Y-%m-%d')
                            except ValueError:
                                pass 
                        new_row_for_temp[idx_new_col] = cell_value 
                    
                    ma_kh_value = new_row_for_temp[4] 
                    if ma_kh_value is None or len(clean_string(ma_kh_value)) > 9:
                        new_row_for_temp.append("No") 
                    else:
                        new_row_for_temp.append("Yes") 
                    
                    intermediate_data_rows.append(new_row_for_temp)

                temp_bkhd_ws_with_cong_no = Workbook().active
                for row_data in intermediate_data_rows:
                    temp_bkhd_ws_with_cong_no.append(row_data)

                b2_bkhd_value = ""
                # Kiểm tra xem có ít nhất một dòng dữ liệu để lấy giá trị B2 (sau khi xóa tiêu đề ban đầu)
                if temp_bkhd_ws_with_cong_no.max_row >= 1: 
                    b2_bkhd_value = clean_string(temp_bkhd_ws_with_cong_no['B1'].value) # Bây giờ B1 là cột B của dòng dữ liệu đầu tiên
                
                normalized_f5_value_full = clean_string(f5_value_full)
                if normalized_f5_value_full.startswith('1'):
                    normalized_f5_value_full = normalized_f5_value_full[1:]

                if normalized_f5_value_full != b2_bkhd_value:
                    st.error("Bảng kê hóa đơn không phải của cửa hàng bạn chọn.")
                    st.stop()

                # --- Vòng lặp xử lý chính: Xây dựng final_upsse_output_rows và thu thập tất cả các dòng TMT ---
                final_upsse_output_rows = []
                all_tmt_rows = [] # Danh sách mới để thu thập tất cả các dòng TMT

                # Thêm 4 dòng trống và dòng tiêu đề
                for _ in range(4):
                    final_upsse_output_rows.append([''] * len(headers))
                final_upsse_output_rows.append(headers)

                # Danh sách để lưu trữ các dòng "Người mua không lấy hóa đơn" để tổng hợp sau
                no_invoice_e5_rows = []
                no_invoice_95_rows = []
                no_invoice_do_rows = []
                no_invoice_d1_rows = []

                # Lặp qua các dòng từ temp_bkhd_ws_with_cong_no (dữ liệu thô từ bảng kê hóa đơn, chỉ mục 0)
                for row_idx_from_bkhd, row_values_from_bkhd in enumerate(temp_bkhd_ws_with_cong_no.iter_rows(min_row=1, values_only=True)):
                    new_row_for_upsse = [''] * len(headers)

                    cong_no_status = row_values_from_bkhd[-1] 
                    
                    # Điền new_row_for_upsse dựa trên logic gốc
                    if cong_no_status == 'No':
                        new_row_for_upsse[0] = g5_value 
                    elif cong_no_status == 'Yes':
                        if row_values_from_bkhd[4] is None or clean_string(row_values_from_bkhd[4]) == '': 
                            new_row_for_upsse[0] = g5_value 
                        else:
                            new_row_for_upsse[0] = clean_string(row_values_from_bkhd[4]) 

                    new_row_for_upsse[1] = clean_string(row_values_from_bkhd[5]) 
                    new_row_for_upsse[2] = row_values_from_bkhd[3] 

                    value_C_for_D = new_row_for_upsse[2] 
                    value_B_for_D_original = row_values_from_bkhd[1] 
                    value_C_for_D_original = row_values_from_bkhd[2] 

                    if b5_value == "Nguyễn Huệ":
                        new_row_for_upsse[3] = "HN" + clean_string(value_C_for_D_original)[-6:] 
                    elif b5_value == "Mai Linh":
                        new_row_for_upsse[3] = "MM" + clean_string(value_C_for_D_original)[-6:] 
                    else:
                        new_row_for_upsse[3] = clean_string(value_B_for_D_original)[-2:] + clean_string(value_C_for_D_original)[-6:] 

                    new_row_for_upsse[4] = "1" + clean_string(value_B_for_D_original) if value_B_for_D_original else '' 
                    new_row_for_upsse[5] = "Xuất bán lẻ theo hóa đơn số " + new_row_for_upsse[3] 

                    new_row_for_upsse[7] = clean_string(row_values_from_bkhd[8]) 
                    new_row_for_upsse[6] = lookup_table.get(clean_string(new_row_for_upsse[7]).lower(), '') 
                    new_row_for_upsse[8] = "Lít" 
                    new_row_for_upsse[9] = g5_value 
                    new_row_for_upsse[10] = '' 
                    new_row_for_upsse[11] = '' 

                    new_row_for_upsse[12] = to_float(row_values_from_bkhd[9]) 
                    tmt_value = to_float(tmt_lookup_table.get(clean_string(new_row_for_upsse[7]).lower(), 0))

                    new_row_for_upsse[13] = round(to_float(row_values_from_bkhd[10]) / 1.1 - tmt_value, 2) if row_values_from_bkhd[10] is not None else 0.0 

                    tmt_calculation_for_row = round(tmt_value * to_float(new_row_for_upsse[12])) if new_row_for_upsse[12] is not None else 0
                    new_row_for_upsse[14] = to_float(row_values_from_bkhd[11]) - tmt_calculation_for_row if row_values_from_bkhd[11] is not None else 0.0 

                    new_row_for_upsse[15] = '' 
                    new_row_for_upsse[16] = '' 
                    new_row_for_upsse[17] = 10 

                    new_row_for_upsse[18] = s_lookup_table.get(h5_value, '') 
                    new_row_for_upsse[19] = t_lookup_table.get(h5_value, '') 
                    new_row_for_upsse[20] = u_value 
                    new_row_for_upsse[21] = v_lookup_table.get(h5_value, '') 
                    new_row_for_upsse[22] = '' 

                    h_value_for_x_lookup = clean_string(new_row_for_upsse[7]).lower()
                    x_value_for_x = x_lookup_table.get(h_value_for_x_lookup, '')
                    new_row_for_upsse[23] = x_value_for_x 

                    new_row_for_upsse[24] = '' 
                    new_row_for_upsse[25] = '' 
                    new_row_for_upsse[26] = '' 
                    new_row_for_upsse[27] = '' 
                    new_row_for_upsse[28] = '' 
                    new_row_for_upsse[29] = '' 
                    new_row_for_upsse[30] = '' 

                    new_row_for_upsse[31] = new_row_for_upsse[1] 
                    new_row_for_upsse[32] = row_values_from_bkhd[6] 
                    new_row_for_upsse[33] = row_values_from_bkhd[7] 
                    new_row_for_upsse[34] = '' 
                    new_row_for_upsse[35] = '' 

                    thue_cua_tmt_for_row_bvmt = 0.0 
                    if new_row_for_upsse[12] is not None and tmt_value is not None:
                        thue_cua_tmt_for_row_bvmt = round(to_float(new_row_for_upsse[12]) * to_float(tmt_value) * 0.1, 0)
                        new_row_for_upsse[36] = to_float(row_values_from_bkhd[12]) - thue_cua_tmt_for_row_bvmt 
                    else:
                        new_row_for_upsse[36] = to_float(row_values_from_bkhd[12]) 
                    
                    # Lọc các dòng chi tiết "Người mua không lấy hóa đơn".
                    # Thu thập chúng để tổng hợp thành các dòng tóm tắt sau.
                    # Chỉ thêm các dòng khác trực tiếp vào final_upsse_output_rows.
                    if clean_string(new_row_for_upsse[1]) == "Người mua không lấy hóa đơn":
                        if clean_string(new_row_for_upsse[7]) == "Xăng E5 RON 92-II":
                            no_invoice_e5_rows.append(new_row_for_upsse)
                        elif clean_string(new_row_for_upsse[7]) == "Xăng RON 95-III":
                            no_invoice_95_rows.append(new_row_for_upsse)
                        elif clean_string(new_row_for_upsse[7]) == "Dầu DO 0,05S-II":
                            no_invoice_do_rows.append(new_row_for_upsse)
                        elif clean_string(new_row_for_upsse[7]) == "Dầu DO 0,001S-V":
                            no_invoice_d1_rows.append(new_row_for_upsse)
                    else:
                        # Thêm dòng gốc vào đầu ra cuối cùng
                        final_upsse_output_rows.append(new_row_for_upsse)

                        # Thêm dòng TMT tương ứng vào all_tmt_rows (nếu áp dụng)
                        if tmt_value > 0 and to_float(new_row_for_upsse[12]) > 0:
                            tmt_per_invoice_row = create_per_invoice_tmt_row(
                                new_row_for_upsse, tmt_value, headers, g5_value, s_lookup_table, t_lookup_table, v_lookup_table, u_value, h5_value
                            )
                            all_tmt_rows.append(tmt_per_invoice_row)
                
                # --- Sau khi xử lý tất cả các dòng gốc, thêm các dòng tổng hợp "Người mua không lấy hóa đơn" ---
                # Và các dòng TMT tương ứng của chúng, cũng được thu thập vào all_tmt_rows.
                
                # Xử lý tóm tắt "Xăng E5 RON 92-II" không có hóa đơn
                if no_invoice_e5_rows:
                    summary_e5_row = add_summary_row_for_no_invoice(no_invoice_e5_rows, bkhd_ws, "Xăng E5 RON 92-II", headers,
                                     g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table)
                    final_upsse_output_rows.append(summary_e5_row)
                    
                    # total_bvmt_e5_summary được tính từ các dòng thô (no_invoice_e5_rows), không phải từ chính dòng tổng hợp
                    total_bvmt_e5_summary = sum(round(to_float(r[12]) * to_float(tmt_lookup_table.get(clean_string(r[7]).lower(), 0)) * 0.1, 0) for r in no_invoice_e5_rows)
                    if total_bvmt_e5_summary > 0:
                        tmt_unit_value = tmt_lookup_table.get(clean_string("Xăng E5 RON 92-II").lower(), 0)
                        total_quantity = sum(to_float(r[12]) for r in no_invoice_e5_rows)
                        customer_name_for_summary_row = summary_e5_row[1] 
                        all_tmt_rows.append(add_tmt_summary_row("Xăng E5 RON 92-II", total_bvmt_e5_summary, headers, g5_value, s_lookup_table, t_lookup_table, v_lookup_table, u_value, h5_value,
                                                                 summary_e5_row[2], summary_e5_row[4], total_quantity, tmt_unit_value, b5_value, customer_name_for_summary_row)) # Truyền các giá trị cần thiết

                # Xử lý tóm tắt "Xăng RON 95-III" không có hóa đơn
                if no_invoice_95_rows:
                    summary_95_row = add_summary_row_for_no_invoice(no_invoice_95_rows, bkhd_ws, "Xăng RON 95-III", headers,
                                     g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table)
                    final_upsse_output_rows.append(summary_95_row)

                    total_bvmt_95_summary = sum(round(to_float(r[12]) * to_float(tmt_lookup_table.get(clean_string(r[7]).lower(), 0)) * 0.1, 0) for r in no_invoice_95_rows)
                    if total_bvmt_95_summary > 0:
                        tmt_unit_value = tmt_lookup_table.get(clean_string("Xăng RON 95-III").lower(), 0)
                        total_quantity = sum(to_float(r[12]) for r in no_invoice_95_rows)
                        customer_name_for_summary_row = summary_95_row[1]
                        all_tmt_rows.append(add_tmt_summary_row("Xăng RON 95-III", total_bvmt_95_summary, headers, g5_value, s_lookup_table, t_lookup_table, v_lookup_table, u_value, h5_value,
                                                                 summary_95_row[2], summary_95_row[4], total_quantity, tmt_unit_value, b5_value, customer_name_for_summary_row)) # Truyền các giá trị cần thiết

                # Xử lý tóm tắt "Dầu DO 0,05S-II" không có hóa đơn
                if no_invoice_do_rows:
                    summary_do_row = add_summary_row_for_no_invoice(no_invoice_do_rows, bkhd_ws, "Dầu DO 0,05S-II", headers,
                                     g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table)
                    final_upsse_output_rows.append(summary_do_row)

                    total_bvmt_do_summary = sum(round(to_float(r[12]) * to_float(tmt_lookup_table.get(clean_string(r[7]).lower(), 0)) * 0.1, 0) for r in no_invoice_do_rows)
                    if total_bvmt_do_summary > 0:
                        tmt_unit_value = tmt_lookup_table.get(clean_string("Dầu DO 0,05S-II").lower(), 0)
                        total_quantity = sum(to_float(r[12]) for r in no_invoice_do_rows)
                        customer_name_for_summary_row = summary_do_row[1]
                        all_tmt_rows.append(add_tmt_summary_row("Dầu DO 0,05S-II", total_bvmt_do_summary, headers, g5_value, s_lookup_table, t_lookup_table, v_lookup_table, u_value, h5_value,
                                                                 summary_do_row[2], summary_do_row[4], total_quantity, tmt_unit_value, b5_value, customer_name_for_summary_row)) # Truyền các giá trị cần thiết

                # Xử lý tóm tắt "Dầu DO 0,001S-V" không có hóa đơn
                if no_invoice_d1_rows:
                    summary_d1_row = add_summary_row_for_no_invoice(no_invoice_d1_rows, bkhd_ws, "Dầu DO 0,001S-V", headers,
                                     g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table)
                    final_upsse_output_rows.append(summary_d1_row)

                    total_bvmt_d1_summary = sum(round(to_float(r[12]) * to_float(tmt_lookup_table.get(clean_string(r[7]).lower(), 0)) * 0.1, 0) for r in no_invoice_d1_rows)
                    if total_bvmt_d1_summary > 0:
                        tmt_unit_value = tmt_lookup_table.get(clean_string("Dầu DO 0,001S-V").lower(), 0)
                        total_quantity = sum(to_float(r[12]) for r in no_invoice_d1_rows)
                        customer_name_for_summary_row = summary_d1_row[1]
                        all_tmt_rows.append(add_tmt_summary_row("Dầu DO 0,001S-V", total_bvmt_d1_summary, headers, g5_value, s_lookup_table, t_lookup_table, v_lookup_table, u_value, h5_value,
                                                                 summary_d1_row[2], summary_d1_row[4], total_quantity, tmt_unit_value, b5_value, customer_name_for_summary_row)) # Truyền các giá trị cần thiết


                # --- Nối tất cả các dòng TMT đã thu thập vào cuối cùng ---
                final_upsse_output_rows.extend(all_tmt_rows)


                # --- Ghi dữ liệu cuối cùng vào worksheet mới và áp dụng định dạng ---
                up_sse_wb_final = Workbook()
                up_sse_ws_final = up_sse_wb_final.active
                for row_data in final_upsse_output_rows:
                    up_sse_ws_final.append(row_data)

                # up_sse_ws và up_sse_wb giờ đã được thay thế bằng up_sse_ws_final và up_sse_wb_final ở phía dưới.
                # Do đó, hai dòng gán lại biến này đã bị xóa để tránh nhầm lẫn đối tượng.
                # up_sse_ws = up_sse_ws_final 
                # up_sse_wb = up_sse_wb_final

                # Định nghĩa các NamedStyle cho định dạng
                text_style = NamedStyle(name="text_style")
                text_style.number_format = '@'

                date_style = NamedStyle(name="date_style")
                date_style.number_format = 'DD/MM/YYYY'

                # Các cột không cần chỉnh sửa định dạng sang văn bản (sử dụng chỉ mục dựa trên 1)
                exclude_columns_idx = {3, 13, 14, 15, 18, 19, 20, 21, 22, 37} 

                for r_idx in range(1, up_sse_ws_final.max_row + 1): # Sử dụng up_sse_ws_final
                    for c_idx in range(1, up_sse_ws_final.max_column + 1):
                        cell = up_sse_ws_final.cell(row=r_idx, column=c_idx)
                        if cell.value is not None and clean_string(cell.value) != "None": 
                            # Định dạng cột C (Ngày)
                            if c_idx == 3: 
                                if isinstance(cell.value, str):
                                    try:
                                        cell.value = datetime.strptime(clean_string(cell.value), '%Y-%m-%d').date() 
                                    except ValueError:
                                        pass 
                                if isinstance(cell.value, datetime):
                                    cell.number_format = 'DD/MM/YYYY' 
                                    cell.style = date_style
                            # Chuyển các cột khác sang văn bản trừ các cột loại trừ
                            elif c_idx not in exclude_columns_idx:
                                cell.value = clean_string(cell.value) 
                                cell.style = text_style

                for r_idx in range(1, up_sse_ws_final.max_row + 1): # Sử dụng up_sse_ws_final
                    for c_idx in range(18, 23): # Cột R (18) đến V (22)
                        cell = up_sse_ws_final.cell(row=r_idx, column=c_idx)
                        cell.number_format = '@' 

                up_sse_ws_final.column_dimensions['C'].width = 12 # Sử dụng up_sse_ws_final
                up_sse_ws_final.column_dimensions['D'].width = 12 # Sử dụng up_sse_ws_final
                up_sse_ws_final.column_dimensions['B'].width = 35 # Sử dụng up_sse_ws_final

                # Ghi file kết quả vào bộ nhớ đệm
                output = io.BytesIO()
                up_sse_wb_final.save(output) # Sử dụng up_sse_wb_final
                processed_data = output.getvalue()

                st.success("Đã tạo file UpSSE.xlsx thành công!")
                st.download_button(
                    label="Tải xuống file UpSSE.xlsx",
                    data=processed_data,
                    file_name="UpSSE.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Lỗi trong quá trình xử lý file: {e}")
                st.exception(e)
