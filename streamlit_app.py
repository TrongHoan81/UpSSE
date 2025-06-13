import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle, Alignment
from datetime import datetime
import io
import os

# --- Cấu hình trang Streamlit ---
st.set_page_config(layout="centered", page_title="Đồng bộ dữ liệu SSE")

# Đường dẫn đến các file cần thiết (giả định cùng thư mục với script)
LOGO_PATH = "Logo.png"
DATA_FILE_PATH = "Data.xlsx" # Tên chính xác của file dữ liệu

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
        return float(value)
    except (ValueError, TypeError):
        return 0.0

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
        # - Khu vực (H5) ở cột S (index 18) -- Đã chỉnh sửa thành index 17 dựa trên lỗi cũ
        # Bắt đầu đọc từ hàng 4 (index 3)
        for row_idx in range(4, ws.max_row + 1):
            row_data_values = [cell.value for cell in ws[row_idx]]

            if len(row_data_values) >= 18: # Đảm bảo đủ cột để truy cập index 17 (cột S)
                raw_chxd_name = row_data_values[10] # Cột K (index 10)

                if raw_chxd_name is not None and str(raw_chxd_name).strip() != '':
                    chxd_name_str = str(raw_chxd_name).strip()
                    
                    if chxd_name_str and chxd_name_str not in listbox_data: # Tránh trùng lặp trong listbox
                        listbox_data.append(chxd_name_str)

                    # Lấy các giá trị cho G5, H5, F5_full, B5 từ các cột tương ứng
                    g5_val = row_data_values[15] if len(row_data_values) > 15 and pd.notna(row_data_values[15]) else None # Cột P (index 15)
                    f5_val_full = str(row_data_values[16]).strip() if len(row_data_values) > 16 and pd.notna(row_data_values[16]) else '' # Cột Q (index 16)
                    h5_val = str(row_data_values[17]).strip().lower() if len(row_data_values) > 17 and pd.notna(row_data_values[17]) else '' # Cột S (index 17)
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
                lookup_table[str(row[0]).strip().lower()] = row[1]
        
        # I10:J13 (TMT Lookup table - Tên mặt hàng <-> Mức phí BVMT)
        tmt_lookup_table = {}
        for row in ws.iter_rows(min_row=10, max_row=13, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                tmt_lookup_table[str(row[0]).strip().lower()] = to_float(row[1]) # Chuyển sang float
        
        # I29:J31 (S Lookup table - Khu vực <-> Tk nợ)
        s_lookup_table = {}
        for row in ws.iter_rows(min_row=29, max_row=31, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                s_lookup_table[str(row[0]).strip().lower()] = row[1]
        
        # I33:J35 (T Lookup table - Khu vực <-> Tk doanh thu)
        t_lookup_table = {}
        for row in ws.iter_rows(min_row=33, max_row=35, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                t_lookup_table[str(row[0]).strip().lower()] = row[1]
        
        # I53:J55 (V Lookup table - Khu vực <-> Tk thuế có)
        v_lookup_table = {}
        for row in ws.iter_rows(min_row=53, max_row=55, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                v_lookup_table[str(row[0]).strip().lower()] = row[1]
        
        # I17:J20 (X Lookup table - Tên mặt hàng <-> Mã vụ việc)
        x_lookup_table = {}
        for row in ws.iter_rows(min_row=17, max_row=20, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                x_lookup_table[str(row[0]).strip().lower()] = row[1]

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

    # Lấy giá trị C6 và E6 từ các dòng đã được xử lý (final_rows_for_upsse_ref)
    c6_val = None
    e6_val = None
    if len(current_up_sse_rows_ref) > 5 and len(current_up_sse_rows_ref[5]) > 4: 
        c6_val = current_up_sse_rows_ref[5][2]
        e6_val = current_up_sse_rows_ref[5][4]
    
    c6_val = c6_val if c6_val is not None else ""
    e6_val = e6_val if e6_val is not None else ""

    new_row[2] = c6_val # Cột C (Ngày)
    new_row[4] = e6_val # Cột E (Ký hiệu)

    value_C = new_row[2] if new_row[2] else ""
    value_E = new_row[4] if new_row[4] else ""

    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name, "")

    if b5_val == "Nguyễn Huệ":
        value_D = f"HNBK{str(value_C)[-2:]}.{str(value_C)[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh":
        value_D = f"MMBK{str(value_C)[-2:]}.{str(value_C)[5:7]}.{suffix_d}"
    else:
        value_D = f"{str(value_E)[-2:]}BK{str(value_C)[-2:]}.{str(value_C)[5:7]}.{suffix_d}"
    new_row[3] = value_D # Cột D (Số hóa đơn)

    new_row[5] = f"Xuất bán lẻ theo hóa đơn số " + new_row[3] # Cột F (Diễn giải)
    new_row[7] = product_name # Cột H (Tên mặt hàng)
    new_row[6] = common_lookup_table.get(product_name.strip().lower(), '') # Cột G (Mã hàng)
    new_row[8] = "Lít" # Cột I (Đvt)
    new_row[9] = g5_val # Cột J (Mã kho)
    new_row[10] = '' # Cột K (Mã vị trí)
    new_row[11] = '' # Cột L (Mã lô)

    # Tính tổng Số lượng (cột M) và Max Giá bán (cột N) từ temp_all_rows_ws
    total_M = 0.0
    max_value_N = 0.0
    for r_idx in range(6, temp_all_rows_ws.max_row + 1): 
        row_data = [cell.value for cell in temp_all_rows_ws[r_idx]]
        if len(row_data) > 12 and str(row_data[1]).strip() == "Người mua không lấy hóa đơn" and str(row_data[7]).strip() == product_name:
            total_M += to_float(row_data[12])
            current_N = to_float(row_data[13])
            if current_N > max_value_N: # Tìm giá trị N lớn nhất
                max_value_N = current_N
    
    new_row[12] = total_M # Cột M (Số lượng)
    new_row[13] = max_value_N # Cột N (Giá bán)

    # Tính Tiền hàng (cột O)
    tien_hang_hd_from_bkhd = 0.0
    for r in bkhd_source_ws.iter_rows(min_row=2, max_row=bkhd_source_ws.max_row, values_only=True):
        if str(r[5]).strip() == "Người mua không lấy hóa đơn" and str(r[8]).strip() == product_name:
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
    new_row[23] = x_lookup.get(product_name.strip().lower(), '') # Vụ việc (Cột X)

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
        if str(r[5]).strip() == "Người mua không lấy hóa đơn" and str(r[8]).strip() == product_name:
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
            selected_value_normalized = selected_value.strip()

            if selected_value_normalized not in chxd_detail_map:
                st.error("Không tìm thấy thông tin chi tiết cho CHXD đã chọn trong Data.xlsx. Vui lòng kiểm tra lại tên CHXD.")
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

            # Xóa 3 hàng đầu tiên của bảng kê (trên bản copy trong bộ nhớ)
            # Tạo một bản sao của worksheet gốc để tránh sửa đổi trực tiếp khi lặp
            temp_bkhd_data_for_processing = []
            for row in bkhd_ws.iter_rows(min_row=1, values_only=True):
                temp_bkhd_data_for_processing.append(list(row)) # Convert tuple to list for mutability

            # Apply delete_rows logic on temp_bkhd_data_for_processing
            # Remove first 3 rows (0-indexed)
            if len(temp_bkhd_data_for_processing) >= 3:
                temp_bkhd_data_for_processing = temp_bkhd_data_for_processing[3:]
            else:
                temp_bkhd_data_for_processing = [] # No data after removing headers

            # Tạo một worksheet tạm thời mới để lưu dữ liệu bảng kê đã qua xử lý ban đầu
            # Các phép tính kvl và bvmt sẽ sử dụng worksheet này
            temp_bkhd_ws_processed = Workbook().active
            for row_data in temp_bkhd_data_for_processing:
                temp_bkhd_ws_processed.append(row_data)

            # Vị trí các cột cần giữ và sắp xếp lại
            vi_tri_cu = ['A', 'B', 'C', 'D', 'E', 'F', 'H', 'G', 'I', 'K', 'L', 'N', 'O', 'Q']
            # Chuyển đổi sang chỉ số 0-based để dễ dàng truy cập trong list
            vi_tri_cu_idx = [ord(c) - ord('A') for c in vi_tri_cu]

            # Tạo bảng dữ liệu mới (data) với xử lý ngày (cột D) và "Công nợ"
            intermediate_data_rows = [] # Chứa các hàng đã được chọn cột và xử lý ngày
            for row_idx, row_values_original in enumerate(temp_bkhd_ws_processed.iter_rows(min_row=1, values_only=True)):
                new_row_for_temp = []
                for idx_new_col, col_old_idx in enumerate(vi_tri_cu_idx):
                    cell_value = row_values_original[col_old_idx] if col_old_idx < len(row_values_original) else None

                    if idx_new_col == 3 and cell_value: # Cột D mới (Ngày)
                        cell_value_str = str(cell_value)[:10] 
                        try:
                            date_obj = datetime.strptime(cell_value_str, '%d-%m-%Y')
                            cell_value = date_obj.strftime('%Y-%m-%d')
                        except ValueError:
                            pass 
                    new_row_for_temp.append(cell_value)
                intermediate_data_rows.append(new_row_for_temp)

            # Thêm cột "Công nợ" vào intermediate_data_rows
            # Giả định cột E (mã KH) là index 4 trong new_row_for_temp
            for r_data in intermediate_data_rows:
                ma_kh_value = r_data[4] # Cột E (mã KH)
                if ma_kh_value is None or len(str(ma_kh_value)) > 9:
                    r_data.append("No") # Thêm 'No' vào cuối hàng
                else:
                    r_data.append("Yes") # Thêm 'Yes' vào cuối hàng

            # Tạo một worksheet tạm thời khác (temp_bkhd_ws_with_cong_no) chứa dữ liệu đã sắp xếp lại cột và có cột Công nợ
            temp_bkhd_ws_with_cong_no = Workbook().active
            for row_data in intermediate_data_rows:
                temp_bkhd_ws_with_cong_no.append(row_data)

            # Kiểm tra Mã kho (F5_full) với ô B2 trên bkhd_ws gốc (sau khi đã xử lý 3 hàng đầu)
            # Lấy giá trị B2 từ temp_bkhd_ws_processed (là hàng đầu tiên của bảng kê sau khi xóa 3 dòng đầu)
            b2_bkhd_value = str(temp_bkhd_ws_processed['B1'].value).strip() if temp_bkhd_ws_processed['B1'].value else '' # B1 của temp_bkhd_ws_processed tương ứng với B4 của bkhd_ws gốc

            normalized_f5_value_full = f5_value_full
            if normalized_f5_value_full.startswith('1'):
                normalized_f5_value_full = normalized_f5_value_full[1:]

            if normalized_f5_value_full != b2_bkhd_value:
                st.error("Bảng kê hóa đơn không phải của cửa hàng bạn chọn.")
                st.stop()

            # --- Bước 1: Tạo UpSSE tạm thời với TẤT CẢ các dòng dữ liệu từ bảng kê ---
            temp_up_sse_all_rows_ws = Workbook().active

            # Thêm các dòng trống và tiêu đề
            for _ in range(4):
                temp_up_sse_all_rows_ws.append([''] * len(headers)) # Thêm hàng trống với đủ số cột
            temp_up_sse_all_rows_ws.append(headers)

            kvlE5 = 0
            kvl95 = 0
            kvlDo = 0
            kvlD1 = 0
            total_bvmt_e5 = 0.0
            total_bvmt_95 = 0.0
            total_bvmt_do = 0.0
            total_bvmt_d1 = 0.0

            # Duyệt qua các hàng từ temp_bkhd_ws_with_cong_no (đã có cột Công nợ cuối cùng)
            for row_idx, row_values_from_bkhd in enumerate(temp_bkhd_ws_with_cong_no.iter_rows(min_row=1, values_only=True)):
                new_row_for_upsse = [''] * len(headers)

                # Cột O (index 14) của row_values_from_bkhd là cột "Công nợ" đã được thêm vào cuối
                cong_no_status = row_values_from_bkhd[14] # Index 14 là cột 'Công nợ' mới thêm

                # Điều kiện cho cột A (Mã khách)
                if cong_no_status == 'No':
                    new_row_for_upsse[0] = g5_value
                elif cong_no_status == 'Yes':
                    if row_values_from_bkhd[4] is None or str(row_values_from_bkhd[4]).strip() == '': # Cột E (mã KH) của bkhd
                        new_row_for_upsse[0] = g5_value
                    else:
                        new_row_for_upsse[0] = str(row_values_from_bkhd[4]) # Giá trị của cột E của bkhd

                new_row_for_upsse[1] = row_values_from_bkhd[5] # Cột B (Tên khách hàng) từ cột F của bkhd
                new_row_for_upsse[2] = row_values_from_bkhd[3] # Cột C (Ngày) từ cột D của bkhd

                # Cột D (Số hóa đơn)
                value_C_for_D = new_row_for_upsse[2] # Lấy từ cột C mới
                value_B_for_D_original = row_values_from_bkhd[1] # Lấy từ cột B gốc của bkhd (index 1)
                value_C_for_D_original = row_values_from_bkhd[2] # Lấy từ cột C gốc của bkhd (index 2)

                if b5_value == "Nguyễn Huệ":
                    new_row_for_upsse[3] = "HN" + str(value_C_for_D_original)[-6:] 
                elif b5_value == "Mai Linh":
                    new_row_for_upsse[3] = "MM" + str(value_C_for_D_original)[-6:] 
                else:
                    new_row_for_upsse[3] = str(value_B_for_D_original)[-2:] + str(value_C_for_D_original)[-6:]  

                new_row_for_upsse[4] = "1" + str(value_B_for_D_original) if value_B_for_D_original else '' # Cột E (Ký hiệu)
                new_row_for_upsse[5] = "Xuất bán lẻ theo hóa đơn số " + new_row_for_upsse[3] # Cột F (Diễn giải)

                new_row_for_upsse[7] = row_values_from_bkhd[8] # Cột H (Tên mặt hàng) từ cột I của bkhd
                new_row_for_upsse[6] = lookup_table.get(str(new_row_for_upsse[7]).strip().lower(), '') # Cột G (Mã hàng)
                new_row_for_upsse[8] = "Lít" # Cột I (Đvt)
                new_row_for_upsse[9] = g5_value # Cột J (Mã kho)
                new_row_for_upsse[10] = '' # Cột K (Mã vị trí)
                new_row_for_upsse[11] = '' # Cột L (Mã lô)

                new_row_for_upsse[12] = to_float(row_values_from_bkhd[9]) # Cột M (Số lượng) từ cột J của bkhd
                tmt_value = to_float(tmt_lookup_table.get(str(new_row_for_upsse[7]).strip().lower(), 0))

                new_row_for_upsse[13] = round(to_float(row_values_from_bkhd[10]) / 1.1 - tmt_value, 2) if row_values_from_bkhd[10] is not None else 0.0 # Cột N (Giá bán) từ cột K của bkhd

                tmt_calculation_for_row = round(tmt_value * new_row_for_upsse[12]) if new_row_for_upsse[12] is not None else 0
                new_row_for_upsse[14] = to_float(row_values_from_bkhd[11]) - tmt_calculation_for_row if row_values_from_bkhd[11] is not None else 0.0 # Cột O (Tiền hàng) từ cột L của bkhd

                new_row_for_upsse[15] = '' # Mã nt
                new_row_for_upsse[16] = '' # Tỷ giá
                new_row_for_upsse[17] = 10 # Mã thuế (Cột R)

                new_row_for_upsse[18] = s_lookup_table.get(h5_value, '') # Tk nợ (Cột S)
                new_row_for_upsse[19] = t_lookup_table.get(h5_value, '') # Tk doanh thu (Cột T)
                new_row_for_upsse[20] = u_value # Tk giá vốn (Cột U)
                new_row_for_upsse[21] = v_lookup_table.get(h5_value, '') # Tk thuế có (Cột V)
                new_row_for_upsse[22] = '' # Cục thuế (Cột W)

                h_value_for_x_lookup = str(new_row_for_upsse[7]).strip().lower()
                x_value_for_x = x_lookup_table.get(h_value_for_x_lookup, '')
                new_row_for_upsse[23] = x_value_for_x # Vụ việc (Cột X)

                new_row_for_upsse[24] = '' # Bộ phận
                new_row_for_upsse[25] = '' # Lsx
                new_row_for_upsse[26] = '' # Sản phẩm
                new_row_for_upsse[27] = '' # Hợp đồng
                new_row_for_upsse[28] = '' # Phí
                new_row_for_upsse[29] = '' # Khế ước
                new_row_for_upsse[30] = '' # Nhân viên bán

                new_row_for_upsse[31] = new_row_for_upsse[1] # Tên KH(thuế) (Cột AF)
                new_row_for_upsse[32] = row_values_from_bkhd[6] # Địa chỉ (thuế) từ cột G của bkhd
                new_row_for_upsse[33] = row_values_from_bkhd[7] # Mã số Thuế từ cột H của bkhd
                new_row_for_upsse[34] = '' # Nhóm Hàng
                new_row_for_upsse[35] = '' # Ghi chú

                # Cột AK (Tiền thuế) - tính toán từ gốc
                thue_cua_tmt_for_row_bvmt = 0.0 # Tích lũy cho total_bvmt
                if new_row_for_upsse[12] is not None and tmt_value is not None:
                    thue_cua_tmt_for_row_bvmt = round(new_row_for_upsse[12] * tmt_value * 0.1, 0)
                    new_row_for_upsse[36] = to_float(row_values_from_bkhd[12]) - thue_cua_tmt_for_row_bvmt # Cột M gốc của bkhd trừ thuế BVMT
                else:
                    new_row_for_upsse[36] = to_float(row_values_from_bkhd[12]) # Giữ nguyên giá trị gốc nếu không có TMT

                # Tích lũy tổng thuế BVMT
                if new_row_for_upsse[7] == "Xăng E5 RON 92-II":
                    total_bvmt_e5 += thue_cua_tmt_for_row_bvmt
                elif new_row_for_upsse[7] == "Xăng RON 95-III":
                    total_bvmt_95 += thue_cua_tmt_for_row_bvmt
                elif new_row_for_upsse[7] == "Dầu DO 0,05S-II":
                    total_bvmt_do += thue_cua_tmt_for_row_bvmt
                elif new_row_for_upsse[7] == "Dầu DO 0,001S-V":
                    total_bvmt_d1 += thue_cua_tmt_for_row_bvmt

                # Thêm dòng mới vào worksheet tạm thời chứa tất cả các dòng đã xử lý
                temp_up_sse_all_rows_ws.append(new_row_for_upsse)

                # Đếm số lượng dòng khách vãng lai
                if str(new_row_for_upsse[1]).strip() == "Người mua không lấy hóa đơn":
                    if str(new_row_for_upsse[7]).strip() == "Xăng E5 RON 92-II":
                        kvlE5 += 1
                    elif str(new_row_for_upsse[7]).strip() == "Xăng RON 95-III":
                        kvl95 += 1
                    elif str(new_row_for_upsse[7]).strip() == "Dầu DO 0,05S-II":
                        kvlDo += 1
                    elif str(new_row_for_upsse[7]).strip() == "Dầu DO 0,001S-V":
                        kvlD1 += 1


            # --- Bước 2 & 3: Lọc bỏ dòng "Người mua không lấy hóa đơn" và thêm dòng tổng kết KH vãng lai ---
            final_rows_for_upsse = []

            # Giữ lại 5 hàng đầu tiên (tiêu đề và các dòng trống)
            for r_idx in range(1, 6):
                row_values = [cell.value for cell in temp_up_sse_all_rows_ws[r_idx]]
                final_rows_for_upsse.append(row_values)
            
            # Lặp qua các dòng dữ liệu thực tế (từ hàng 6) từ temp_up_sse_all_rows_ws
            for r_idx in range(6, temp_up_sse_all_rows_ws.max_row + 1):
                row_data = [cell.value for cell in temp_up_sse_all_rows_ws[r_idx]]
                
                if len(row_data) > 1 and row_data[1] is not None:
                    col_b_value = str(row_data[1]).strip()
                    
                    if col_b_value == "Người mua không lấy hóa đơn":
                        continue # Bỏ qua dòng này
                    else:
                        final_rows_for_upsse.append(row_data)
                else:
                    final_rows_for_upsse.append(row_data) # Giữ lại nếu không đủ cột hoặc cột B rỗng (dòng trắng)

            # Thêm các dòng tổng kết "Khách hàng mua..."
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
            
            # --- Bước 3: Thêm các dòng tổng kết Thuế bảo vệ môi trường (TMT) ---
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
                    if cell.value is not None and str(cell.value) != "None":
                        # Định dạng cột C (Ngày)
                        if c_idx == 3: 
                            if isinstance(cell.value, str):
                                try:
                                    cell.value = datetime.strptime(cell.value, '%Y-%m-%d').date()
                                except ValueError:
                                    pass 
                            if isinstance(cell.value, datetime):
                                cell.number_format = 'DD/MM/YYYY' 
                                cell.style = date_style
                        # Chuyển các cột khác sang văn bản trừ các cột loại trừ
                        elif c_idx not in exclude_columns_idx:
                            cell.value = str(cell.value)
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
