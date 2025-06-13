import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle, Alignment
from datetime import datetime
import io
import os
import re # Import regex module

# --- Cấu hình trang Streamlit ---
st.set_page_config(layout="centered", page_title="Đồng bộ dữ liệu SSE")

# Đường dẫn đến các file cần thiết (giả định cùng thư mục với script)
LOGO_PATH = "Logo.png"
DATA_FILE_PATH = "Data.xlsx" # Tên chính xác của file dữ liệu

# Định nghĩa tiêu đề cho file UpSSE.xlsx (Di chuyển lên đây để luôn có sẵn)
headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng",
           "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế",
           "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm",
           "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế",
           "Nhóm Hàng", "Ghi chú", "Tiền thuế"]

# --- Kiểm tra ngày hết hạn ứng dụng ---
expiration_date = datetime(2025, 6, 26)
current_date = datetime.now()

if current_date > expiration_date:
    st.error("Có lỗi khi chạy chương trình, vui lòng liên hệ tác giả để được hỗ trợ!")
    st.info("Nguyễn Trọng Hoàn - 0902069469")
    st.stop() # Dừng ứng dụng

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

# --- Hàm tạo dòng tổng kết Khách vãng lai ---
def add_summary_row(temp_all_rows_ws, bkhd_source_ws, product_name, headers_list,
                    g5_val, b5_val, s_lookup, t_lookup, v_lookup, x_lookup, u_val, h5_val, common_lookup_table, current_up_sse_rows_ref):
    """
    Tạo một dòng tổng kết cho khách vãng lai.
    temp_all_rows_ws: Worksheet tạm thời chứa tất cả các dòng đã xử lý (bao gồm cả "Người mua không lấy hóa đơn")
    bkhd_source_ws: Worksheet gốc của bảng kê hóa đơn (để lấy giá trị gốc cho TienhangHD, TienthueHD)
    """
    new_row = [''] * len(headers_list)
    new_row[0] = g5_val
    new_row[1] = f"Khách hàng mua {product_name} không lấy hóa đơn"

    c6_val = None
    e6_val = None
    # Lấy giá trị C6 và E6 từ các dòng đã được xử lý (final_rows_for_upsse_ref)
    # Cần đảm bảo rằng current_up_sse_rows_ref đã có ít nhất 6 hàng (5 hàng header + 1 hàng data)
    if len(current_up_sse_rows_ref) > 5 and len(current_up_sse_rows_ref[5]) > 4: 
        c6_val = current_up_sse_rows_ref[5][2] # Cột C (index 2) của hàng thứ 6
        e6_val = current_up_sse_rows_ref[5][4] # Cột E (index 4) của hàng thứ 6
    
    c6_val = c6_val if c6_val is not None else ""
    e6_val = e6_val if e6_val is not None else ""

    new_row[2] = c6_val # Cột C (Ngày)
    new_row[4] = e6_val # Cột E (Ký hiệu)

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

    # Tính tổng Số lượng (cột M) và Max Giá bán (cột N) từ temp_all_rows_ws
    total_M = 0.0
    max_value_N = 0.0
    for r_idx in range(6, temp_all_rows_ws.max_row + 1): 
        row_data = [cell.value for cell in temp_all_rows_ws[r_idx]]
        # Đảm bảo hàng có đủ cột và giá trị đúng
        if len(row_data) > 12 and clean_string(row_data[1]) == "Người mua không lấy hóa đơn" and clean_string(row_data[7]) == product_name:
            total_M += to_float(row_data[12])
            current_N = to_float(row_data[13])
            if current_N > max_value_N: 
                max_value_N = current_N
    
    new_row[12] = total_M # Cột M (Số lượng)
    new_row[13] = max_value_N # Cột N (Giá bán)

    tien_hang_hd_from_bkhd = 0.0
    for r in bkhd_source_ws.iter_rows(min_row=2, max_row=bkhd_source_ws.max_row, values_only=True):
        if clean_string(r[5]) == "Người mua không lấy hóa đơn" and clean_string(r[8]) == product_name:
            tien_hang_hd_from_bkhd += to_float(r[11]) # Cột L trên BKHD

    # Mức giá bán cụ thể cho từng loại sản phẩm
    price_per_liter_map = {"Xăng E5 RON 92-II": 1900, "Xăng RON 95-III": 2000, "Dầu DO 0,05S-II": 1000, "Dầu DO 0,001S-V": 1000}
    current_price_per_liter = price_per_liter_map.get(product_name, 0)
    
    new_row[14] = tien_hang_hd_from_bkhd - round(total_M * current_price_per_liter, 0) # Cột O (Tiền hàng)

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
    new_row[32] = "" # Địa chỉ (thuế) (Cột AG) - Không có dữ liệu trong file gốc
    new_row[33] = "" # Mã số Thuế (Cột AH) - Không có dữ liệu trong file gốc
    new_row[34] = '' # Nhóm Hàng
    new_row[35] = '' # Ghi chú

    tien_thue_hd_from_bkhd = 0.0
    for r in bkhd_source_ws.iter_rows(min_row=2, max_row=bkhd_source_ws.max_row, values_only=True):
        if clean_string(r[5]) == "Người mua không lấy hóa đơn" and clean_string(r[8]) == product_name:
            tien_thue_hd_from_bkhd += to_float(r[12]) # Cột M trên BKHD

    new_row[36] = tien_thue_hd_from_bkhd - round(total_M * current_price_per_liter * 0.1, 0) # Cột AK (Tiền thuế)

    return new_row

# --- Hàm tạo dòng tóm tắt thuế bảo vệ môi trường (TMT) ---
def add_tmt_summary_row(product_name_full, total_bvmt_amount, headers_list, g5_val, s_lookup, t_lookup, v_lookup, u_val, h5_val):
    """
    Tạo một dòng tổng kết cho Thuế bảo vệ môi trường.
    """
    new_tmt_row = [''] * len(headers_list)
    new_tmt_row[0] = g5_val # Mã khách
    new_tmt_row[1] = f"Thuế bảo vệ môi trường {product_name_full}" # Tên khách hàng (Diễn giải)
    
    new_tmt_row[6] = "TMT" # Mã hàng
    new_tmt_row[7] = "Thuế bảo vệ môi trường" # Tên mặt hàng
    new_tmt_row[8] = "VNĐ" # Đvt
    new_tmt_row[9] = g5_val # Mã kho (giống Mã khách)
    
    # Tài khoản
    new_tmt_row[18] = s_lookup.get(h5_val, '') # Tk nợ
    new_tmt_row[19] = t_lookup.get(h5_val, '') # Tk doanh thu
    new_tmt_row[20] = u_val # Tk giá vốn
    new_tmt_row[21] = v_lookup.get(h5_val, '') # Tk thuế có

    new_tmt_row[36] = total_bvmt_amount # Tiền thuế (Tổng thuế BVMT)
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
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.image(LOGO_PATH, width=100)
    st.markdown(
        """
        <div style="text-align: center;">
            <h1 style="color: red; font-size: 24px; margin-bottom: 0px;">CÔNG TY CỔ PHẦN XĂNG DẦU</h1>
            <h2 style="color: red; font-size: 24px; margin-top: 0px;">DẦU KHÍ NAM ĐỊNH</h2>
        </div>
        """,
        unsafe_allow_html=True
    )

st.title("Đồng bộ dữ liệu SSE")

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
            # Tạo một bản sao của worksheet gốc để tránh sửa đổi trực tiếp khi lặp
            temp_bkhd_data_for_processing = []
            for row in bkhd_ws.iter_rows(min_row=1, values_only=True):
                temp_bkhd_data_for_processing.append(list(row)) 

            # Xóa 3 hàng đầu tiên (0-indexed)
            if len(temp_bkhd_data_for_processing) >= 3:
                temp_bkhd_data_for_processing = temp_bkhd_data_for_processing[3:]
            else:
                temp_bkhd_data_for_processing = []

            # Tạo một worksheet tạm thời mới để lưu dữ liệu bảng kê đã qua xử lý ban đầu
            temp_bkhd_ws_processed = Workbook().active
            for row_data in temp_bkhd_data_for_processing:
                temp_bkhd_ws_processed.append(row_data)

            # Vị trí các cột cần giữ và sắp xếp lại (từ file gốc của bạn)
            # Vi_tri_cu (ban đầu): ['A', 'B', 'C', 'D', 'E', 'F', 'H', 'G', 'I', 'K', 'L', 'N', 'O', 'Q']
            # Chuyển đổi sang chỉ số 0-based: [0, 1, 2, 3, 4, 5, 7, 6, 8, 10, 11, 13, 14, 16]
            vi_tri_cu_idx = [0, 1, 2, 3, 4, 5, 7, 6, 8, 10, 11, 13, 14, 16]

            # Tạo bảng dữ liệu mới (intermediate_data_rows) với xử lý ngày (cột D) và thêm cột "Công nợ"
            intermediate_data_rows = [] 
            for row_idx, row_values_original in enumerate(temp_bkhd_ws_processed.iter_rows(min_row=1, values_only=True)):
                new_row_for_temp = [''] * (len(vi_tri_cu_idx) + 1) # +1 for Cong no column
                # Đảm bảo hàng có đủ số cột cần thiết (ví dụ: đến index 16 cho cột Q)
                if len(row_values_original) <= max(vi_tri_cu_idx): # Cột Q là index 16. If original row has less than 17 cols, it's problematic
                    continue # Bỏ qua hàng không đủ dữ liệu

                for idx_new_col, col_old_idx in enumerate(vi_tri_cu_idx):
                    cell_value = row_values_original[col_old_idx] 

                    if idx_new_col == 3 and cell_value: # Cột D mới (Ngày)
                        cell_value_str = str(cell_value)[:10] 
                        try:
                            date_obj = datetime.strptime(cell_value_str, '%d-%m-%Y')
                            cell_value = date_obj.strftime('%Y-%m-%d')
                        except ValueError:
                            pass 
                    new_row_for_temp[idx_new_col] = cell_value # Assign to specific index
                
                # Thêm cột "Công nợ" vào cuối hàng
                ma_kh_value = new_row_for_temp[4] # Cột E (mã KH) của hàng mới (index 4)
                if ma_kh_value is None or len(clean_string(ma_kh_value)) > 9:
                    new_row_for_temp.append("No") 
                else:
                    new_row_for_temp.append("Yes") 
                
                intermediate_data_rows.append(new_row_for_temp)

            # Tạo một worksheet tạm thời khác (temp_bkhd_ws_with_cong_no) chứa dữ liệu đã sắp xếp lại cột và có cột Công nợ
            # Worksheet này sẽ được dùng làm nguồn dữ liệu chính cho các bước tính toán và tạo UpSSE
            temp_bkhd_ws_with_cong_no = Workbook().active
            for row_data in intermediate_data_rows:
                temp_bkhd_ws_with_cong_no.append(row_data)

            # --- Kiểm tra mã kho: So sánh Mã kho từ Data.xlsx với B2 của bảng kê đã làm sạch ---
            # Mã kho trong bảng kê nằm ở cột B của hàng dữ liệu đầu tiên (ô B2 sau khi xóa headers ban đầu)
            # Dựa trên debug trước đó, ô B2 của temp_bkhd_ws_with_cong_no đang chứa mã kho thực tế
            # Chú ý: temp_bkhd_ws_with_cong_no có thể rỗng nếu intermediate_data_rows rỗng
            b2_bkhd_value = ""
            if temp_bkhd_ws_with_cong_no.max_row >= 2: # Check if there's at least row 2
                b2_bkhd_value = clean_string(temp_bkhd_ws_with_cong_no['B2'].value)
            
            normalized_f5_value_full = clean_string(f5_value_full)
            if normalized_f5_value_full.startswith('1'):
                normalized_f5_value_full = normalized_f5_value_full[1:]

            st.write(f"Debug: Mã kho từ Data.xlsx (F5_full đã chuẩn hóa): '{normalized_f5_value_full}'")
            st.write(f"Debug: Mã kho từ Bảng Kê (B2 đã làm sạch): '{b2_bkhd_value}'")

            if normalized_f5_value_full != b2_bkhd_value:
                st.error("Bảng kê hóa đơn không phải của cửa hàng bạn chọn.")
                st.stop()

            # --- BƯỚC 1: Xử lý và tổng hợp tất cả các dòng dữ liệu vào một cấu trúc tạm thời ---
            # Đây là nơi chúng ta sẽ tính toán kvlE5, kvl95, kvlDo, kvlD1 và total_bvmt cho từng mặt hàng
            # mà không bị ảnh hưởng bởi việc lọc sau đó.
            
            # Khởi tạo các biến tổng hợp
            kvlE5 = 0
            kvl95 = 0
            kvlDo = 0
            kvlD1 = 0
            total_bvmt_e5 = 0.0 
            total_bvmt_95 = 0.0 
            total_bvmt_do = 0.0 
            total_bvmt_d1 = 0.0 

            # Danh sách tạm thời để lưu các dòng đã xử lý trước khi lọc
            temp_processed_upsse_rows = []

            # Định nghĩa temp_up_sse_all_rows_ws và thêm 5 hàng đầu tiên ở đây
            # (tạo một Workbook và Worksheet mới cho mục đích này)
            temp_up_sse_wb_for_all_rows = Workbook()
            temp_up_sse_all_rows_ws = temp_up_sse_wb_for_all_rows.active
            
            for _ in range(4): # 4 hàng trống
                temp_up_sse_all_rows_ws.append([''] * len(headers))
            temp_up_sse_all_rows_ws.append(headers) # Thêm hàng tiêu đề (headers)

            # Duyệt qua các hàng từ temp_bkhd_ws_with_cong_no (có headers và cột "Công nợ")
            # Bắt đầu từ hàng 2 của temp_bkhd_ws_with_cong_no (là hàng dữ liệu đầu tiên)
            for row_idx_from_bkhd, row_values_from_bkhd in enumerate(temp_bkhd_ws_with_cong_no.iter_rows(min_row=2, values_only=True)):
                new_row_for_upsse = [''] * len(headers)

                # Giá trị cột "Công nợ" (là cột cuối cùng của row_values_from_bkhd)
                cong_no_status = row_values_from_bkhd[-1] 

                # Logic điền dữ liệu vào new_row_for_upsse (tương tự các lần trước)
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

                tmt_calculation_for_row = round(tmt_value * new_row_for_upsse[12]) if new_row_for_upsse[12] is not None else 0
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
                    thue_cua_tmt_for_row_bvmt = round(new_row_for_upsse[12] * tmt_value * 0.1, 0)
                    new_row_for_upsse[36] = to_float(row_values_from_bkhd[12]) - thue_cua_tmt_for_row_bvmt 
                else:
                    new_row_for_upsse[36] = to_float(row_values_from_bkhd[12]) 

                # Tích lũy tổng thuế BVMT
                if new_row_for_upsse[7] == "Xăng E5 RON 92-II":
                    total_bvmt_e5 += thue_cua_tmt_for_row_bvmt
                elif new_row_for_upsse[7] == "Xăng RON 95-III":
                    total_bvmt_95 += thue_cua_tmt_for_row_bvmt
                elif new_row_for_upsse[7] == "Dầu DO 0,05S-II":
                    total_bvmt_do += thue_cua_tmt_for_row_bvmt
                elif new_row_for_upsse[7] == "Dầu DO 0,001S-V":
                    total_bvmt_d1 += thue_cua_tmt_for_row_bvmt

                # Thêm dòng đã xử lý vào danh sách tạm thời temp_up_sse_all_rows_ws
                temp_up_sse_all_rows_ws.append(new_row_for_upsse)

                # Đếm số lượng dòng khách vãng lai (dùng cho tổng kết Khách vãng lai)
                if clean_string(new_row_for_upsse[1]) == "Người mua không lấy hóa đơn":
                    if clean_string(new_row_for_upsse[7]) == "Xăng E5 RON 92-II":
                        kvlE5 += 1
                    elif clean_string(new_row_for_upsse[7]) == "Xăng RON 95-III":
                        kvl95 += 1
                    elif clean_string(new_row_for_upsse[7]) == "Dầu DO 0,05S-II":
                        kvlDo += 1
                    elif clean_string(new_row_for_upsse[7]) == "Dầu DO 0,001S-V":
                        kvlD1 += 1


            # --- BƯỚC 2: Lọc bỏ dòng "Người mua không lấy hóa đơn" và thêm dòng tổng kết KH vãng lai ---
            final_rows_for_upsse = []

            # Thêm 5 hàng đầu tiên (dòng trống và tiêu đề) vào final_rows_for_upsse
            # Lấy trực tiếp từ temp_up_sse_all_rows_ws (đã có đủ 5 hàng đầu)
            for r_idx in range(5): # Indices 0 to 4 of temp_up_sse_all_rows_ws
                final_rows_for_upsse.append(temp_up_sse_all_rows_ws[r_idx])
            
            # Lặp qua các dòng dữ liệu thực tế từ temp_up_sse_all_rows_ws (từ hàng 6 - index 5 trở đi)
            for r_idx in range(5, temp_up_sse_all_rows_ws.max_row): # Iterate from index 5 to max_row-1
                row_data = temp_up_sse_all_rows_ws[r_idx]
                
                if len(row_data) > 1 and row_data[1] is not None:
                    col_b_value = clean_string(row_data[1])
                    
                    if col_b_value == "Người mua không lấy hóa đơn":
                        continue 
                    else:
                        final_rows_for_upsse.append(row_data)
                else:
                    final_rows_for_upsse.append(row_data) 

            # Thêm các dòng tổng kết "Khách hàng mua..."
            # Truyền temp_up_sse_all_rows_ws (chứa tất cả các dòng đã xử lý, bao gồm kvl) cho hàm add_summary_row
            if kvlE5 > 0:
                final_rows_for_upsse.append(
                    add_summary_row(temp_up_sse_all_rows_ws, bkhd_ws, "Xăng E5 RON 92-II", headers,
                                g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table, final_rows_for_upsse))
            if kvl95 > 0:
                final_rows_for_upsse.append(
                    add_summary_row(temp_up_sse_all_rows_ws, bkhd_ws, "Xăng RON 95-III", headers,
                                g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table, final_rows_for_upsse))
            if kvlDo > 0:
                final_rows_for_upsse.append(
                    add_summary_row(temp_up_sse_all_rows_ws, bkhd_ws, "Dầu DO 0,05S-II", headers,
                                g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table, final_rows_for_upsse))
            if kvlD1 > 0:
                final_rows_for_upsse.append(
                    add_summary_row(temp_up_sse_all_rows_ws, bkhd_ws, "Dầu DO 0,001S-V", headers,
                                g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table, final_rows_for_upsse))
            
            # --- BƯỚC 3: Thêm các dòng tổng kết Thuế bảo vệ môi trường (TMT) ---
            if total_bvmt_e5 > 0:
                final_rows_for_upsse.append(
                    add_tmt_summary_row("Xăng E5 RON 92-II", total_bvmt_e5, headers, g5_value, s_lookup_table, t_lookup_table, v_lookup_table, u_value, h5_value))
            if total_bvmt_95 > 0:
                final_rows_for_upsse.append(
                    add_tmt_summary_row("Xăng RON 95-III", total_bvmt_95, headers, g5_value, s_lookup_table, t_lookup_table, v_lookup_table, u_value, h5_value))
            if total_bvmt_do > 0:
                final_rows_for_upsse.append(
                    add_tmt_summary_row("Dầu DO 0,05S-II", total_bvmt_do, headers, g5_value, s_lookup_table, t_lookup_table, v_lookup_table, u_value, h5_value))
            if total_bvmt_d1 > 0:
                final_rows_for_upsse.append(
                    add_tmt_summary_row("Dầu DO 0,001S-V", total_bvmt_d1, headers, g5_value, s_lookup_table, t_lookup_table, v_lookup_table, u_value, h5_value))


            # --- Ghi dữ liệu cuối cùng vào worksheet mới và áp dụng định dạng ---
            up_sse_wb_final = Workbook()
            up_sse_ws_final = up_sse_wb_final.active
            for row_data in final_rows_for_upsse:
                up_sse_ws_final.append(row_data)

            up_sse_ws = up_sse_ws_final # Gán lại để xử lý định dạng
            up_sse_wb = up_sse_wb_final

            # Định nghĩa các NamedStyle cho định dạng
            text_style = NamedStyle(name="text_style")
            text_style.number_format = '@'

            date_style = NamedStyle(name="date_style")
            date_style.number_format = 'DD/MM/YYYY'

            # Các cột không cần chỉnh sửa định dạng sang text (dùng chỉ số 1-based)
            exclude_columns_idx = {3, 13, 14, 15, 18, 19, 20, 21, 22, 37} # C (Ngày), M (Số lượng), N (Giá bán), O (Tiền hàng), R (Mã thuế), S (Tk nợ), T (Tk doanh thu), U (Tk giá vốn), V (Tk thuế có), AK (Tiền thuế)

            for r_idx in range(1, up_sse_ws.max_row + 1):
                for c_idx in range(1, up_sse_ws.max_column + 1):
                    cell = up_sse_ws.cell(row=r_idx, column=c_idx)
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

            # Đảm bảo các cột R đến V là Text (Dù đã trong exclude_columns_idx, nhưng cần chắc chắn)
            for r_idx in range(1, up_sse_ws.max_row + 1):
                for c_idx in range(18, 23): # Cột R (18) đến V (22)
                    cell = up_sse_ws.cell(row=r_idx, column=c_idx)
                    cell.number_format = '@' 

            # Mở rộng chiều rộng cột C,D, B
            up_sse_ws.column_dimensions['C'].width = 12
            up_sse_ws.column_dimensions['D'].width = 12
            up_sse_ws.column_dimensions['B'].width = 35 

            # Ghi file kết quả vào bộ nhớ đệm
            output = io.BytesIO()
            up_sse_wb.save(output)
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
