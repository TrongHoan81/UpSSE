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
        # ******** START CHANGE / BẮT ĐẦU THAY ĐỔI ********
        # Map để lưu mã "Vụ việc" cho từng cửa hàng, tái tạo logic VLOOKUP
        store_specific_x_lookup = {}
        # ******** END CHANGE / KẾT THÚC THAY ĐỔI ********

        # Đọc dữ liệu từ bảng CHXD để xây dựng listbox_data và các map tra cứu
        for row_idx in range(4, ws.max_row + 1):
            # Lấy tất cả giá trị của dòng để truy cập bằng chỉ mục
            row_data_values = [cell.value for cell in ws[row_idx]]

            # Đảm bảo có đủ cột để tránh lỗi IndexError
            if len(row_data_values) < 18:
                continue

            raw_chxd_name = row_data_values[10] # Cột K (index 10)

            if raw_chxd_name is not None and clean_string(raw_chxd_name) != '':
                chxd_name_str = clean_string(raw_chxd_name)
                
                if chxd_name_str and chxd_name_str not in listbox_data:
                    listbox_data.append(chxd_name_str)

                # --- Xây dựng chxd_detail_map (thông tin chung của cửa hàng) ---
                g5_val = row_data_values[15] if pd.notna(row_data_values[15]) else None # Cột P
                f5_val_full = clean_string(row_data_values[16]) if pd.notna(row_data_values[16]) else '' # Cột Q
                h5_val = clean_string(row_data_values[17]).lower() if pd.notna(row_data_values[17]) else '' # Cột S
                b5_val = chxd_name_str

                if f5_val_full:
                    chxd_detail_map[chxd_name_str] = {
                        'g5_val': g5_val, 'h5_val': h5_val,
                        'f5_val_full': f5_val_full, 'b5_val': b5_val
                    }

                # ******** START CHANGE / BẮT ĐẦU THAY ĐỔI ********
                # --- Xây dựng store_specific_x_lookup (Mã Vụ việc theo cửa hàng) ---
                # Tái tạo lại logic VLOOKUP và tham chiếu ô
                # Công thức tại J18 = B5; B5 = VLOOKUP(A5,$K:$P,2,0) -> Cột L
                vu_viec_95 = row_data_values[11]
                # Công thức tại J19 = C5; C5 = VLOOKUP(A5,$K:$P,3,0) -> Cột M
                vu_viec_do = row_data_values[12]
                # Công thức tại J17 = D5; D5 = VLOOKUP(A5,$K:$P,4,0) -> Cột N
                vu_viec_e5 = row_data_values[13]
                # Công thức tại J20 = E5; E5 = VLOOKUP(A5,$K:$P,5,0) -> Cột O
                vu_viec_d1 = row_data_values[14]

                store_specific_x_lookup[chxd_name_str] = {
                    "xăng e5 ron 92-ii": vu_viec_e5,
                    "xăng ron 95-iii": vu_viec_95,
                    "dầu do 0,05s-ii": vu_viec_do,
                    "dầu do 0,001s-v": vu_viec_d1
                }
                # ******** END CHANGE / KẾT THÚC THAY ĐỔI ********
        
        # Đọc các bảng tra cứu khác theo phạm vi cụ thể
        lookup_table = {} # I4:J7 (Mã hàng <-> Tên mặt hàng)
        for row in ws.iter_rows(min_row=4, max_row=7, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: lookup_table[clean_string(row[0]).lower()] = row[1]
        
        tmt_lookup_table = {} # I10:J13 (TMT)
        for row in ws.iter_rows(min_row=10, max_row=13, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: tmt_lookup_table[clean_string(row[0]).lower()] = to_float(row[1])
        
        s_lookup_table = {} # I29:J31 (Tk nợ)
        for row in ws.iter_rows(min_row=29, max_row=31, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: s_lookup_table[clean_string(row[0]).lower()] = row[1]
        
        t_lookup_table = {} # I33:J35 (Tk doanh thu)
        for row in ws.iter_rows(min_row=33, max_row=35, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: t_lookup_table[clean_string(row[0]).lower()] = row[1]
        
        v_lookup_table = {} # I53:J55 (Tk thuế có)
        for row in ws.iter_rows(min_row=53, max_row=55, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: v_lookup_table[clean_string(row[0]).lower()] = row[1]
        
        u_value = ws['J36'].value
        wb.close()
        
        return {
            "listbox_data": listbox_data, "lookup_table": lookup_table,
            "tmt_lookup_table": tmt_lookup_table, "s_lookup_table": s_lookup_table,
            "t_lookup_table": t_lookup_table, "v_lookup_table": v_lookup_table,
            "u_value": u_value, "chxd_detail_map": chxd_detail_map,
            # ******** START CHANGE / BẮT ĐẦU THAY ĐỔI ********
            "store_specific_x_lookup": store_specific_x_lookup
            # ******** END CHANGE / KẾT THÚC THAY ĐỔI ********
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
                    g5_val, b5_val, s_lookup, t_lookup, v_lookup, x_lookup_for_store, u_val, h5_val, common_lookup_table):
    """
    Tạo một dòng tổng hợp cho "Người mua không lấy hóa đơn" cho một mặt hàng cụ thể.
    x_lookup_for_store: Bảng tra cứu "Vụ việc" cụ thể cho cửa hàng đã chọn.
    """
    new_row = [''] * len(headers_list)
    new_row[0] = g5_val
    new_row[1] = f"Khách hàng mua {product_name} không lấy hóa đơn"

    c_val_from_first_row = data_for_summary_product[0][2] if data_for_summary_product else ""
    e_val_from_first_row = data_for_summary_product[0][4] if data_for_summary_product else ""
    new_row[2] = c_val_from_first_row
    new_row[4] = e_val_from_first_row

    value_C = clean_string(new_row[2])
    value_E = clean_string(new_row[4])

    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name, "")
    if b5_val == "Nguyễn Huệ": value_D = f"HNBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh": value_D = f"MMBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    else: value_D = f"{value_E[-2:]}BK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    new_row[3] = value_D

    new_row[5] = f"Xuất bán lẻ theo hóa đơn số {new_row[3]}"
    new_row[7] = product_name
    new_row[6] = common_lookup_table.get(clean_string(product_name).lower(), '')
    new_row[8] = "Lít"
    new_row[9] = g5_val
    new_row[10] = ''
    new_row[11] = ''

    total_M = sum(to_float(r[12]) for r in data_for_summary_product)
    new_row[12] = total_M
    
    max_value_N = max((to_float(r[13]) for r in data_for_summary_product), default=0.0)
    new_row[13] = max_value_N

    tien_hang_hd_from_bkhd_original = sum(to_float(r[11]) for r in bkhd_source_ws.iter_rows(min_row=2, values_only=True)
                                 if clean_string(r[5]) == "Người mua không lấy hóa đơn" and clean_string(r[8]) == product_name)
    price_per_liter_map = {"Xăng E5 RON 92-II": 1900, "Xăng RON 95-III": 2000, "Dầu DO 0,05S-II": 1000, "Dầu DO 0,001S-V": 1000}
    current_price_per_liter = price_per_liter_map.get(product_name, 0)
    new_row[14] = tien_hang_hd_from_bkhd_original - round(total_M * current_price_per_liter, 0)

    new_row[15], new_row[16] = '', ''
    new_row[17] = 10

    new_row[18] = s_lookup.get(h5_val, '')
    new_row[19] = t_lookup.get(h5_val, '')
    new_row[20] = u_val
    new_row[21] = v_lookup.get(h5_val, '')
    new_row[22] = ''
    # ******** START CHANGE / BẮT ĐẦU THAY ĐỔI ********
    new_row[23] = x_lookup_for_store.get(clean_string(product_name).lower(), '') # Vụ việc
    # ******** END CHANGE / KẾT THÚC THAY ĐỔI ********

    for i in range(24, 31): new_row[i] = ''
    new_row[31] = f"Khách mua {product_name} không lấy hóa đơn"
    new_row[32], new_row[33], new_row[34], new_row[35] = "", "", '', ''

    tien_thue_hd_from_bkhd_original = sum(to_float(r[12]) for r in bkhd_source_ws.iter_rows(min_row=2, values_only=True)
                                       if clean_string(r[5]) == "Người mua không lấy hóa đơn" and clean_string(r[8]) == product_name)
    new_row[36] = tien_thue_hd_from_bkhd_original - round(total_M * current_price_per_liter * 0.1, 0)

    return new_row

# --- Hàm tạo dòng TMT theo hóa đơn ---
def create_per_invoice_tmt_row(original_row_data, tmt_value, g5_val, s_lookup, t_lookup, v_lookup, u_val, h5_val):
    tmt_row = list(original_row_data)
    tmt_row[6] = "TMT"
    tmt_row[7] = "Thuế bảo vệ môi trường"
    tmt_row[8] = "VNĐ"
    tmt_row[9] = g5_val
    tmt_row[13] = tmt_value
    tmt_row[14] = round(tmt_value * to_float(original_row_data[12]), 0)
    tmt_row[17] = ''
    tmt_row[18] = s_lookup.get(h5_val, '')
    tmt_row[19] = t_lookup.get(h5_val, '')
    tmt_row[20] = u_val
    tmt_row[21] = v_lookup.get(h5_val, '')
    tmt_row[23] = ''
    tmt_row[31] = ""
    tmt_row[36] = round(tmt_value * to_float(original_row_data[12]) * 0.1, 0)
    for idx in [5, 10, 11, 15, 16, 22, 24, 25, 26, 27, 28, 29, 30, 32, 33, 34, 35]:
        if idx < len(tmt_row): tmt_row[idx] = ''
    return tmt_row

# --- Hàm thêm dòng tổng hợp thuế bảo vệ môi trường (TMT) ---
def add_tmt_summary_row(product_name_full, total_bvmt_amount, g5_val, s_lookup, t_lookup, v_lookup, u_val, h5_val, 
                        representative_date, representative_symbol, total_quantity_for_tmt, tmt_unit_value_for_summary, b5_val, customer_name_for_summary_row):
    new_tmt_row = [''] * len(headers)
    new_tmt_row[0] = g5_val
    new_tmt_row[1] = customer_name_for_summary_row
    new_tmt_row[2] = representative_date
    value_C = clean_string(representative_date)
    value_E = clean_string(representative_symbol)
    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name_full, "")
    if b5_val == "Nguyễn Huệ": value_D_tmt_summary = f"HNBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh": value_D_tmt_summary = f"MMBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    else: value_D_tmt_summary = f"{value_E[-2:]}BK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    new_tmt_row[3] = value_D_tmt_summary
    new_tmt_row[4] = representative_symbol
    new_tmt_row[6], new_tmt_row[7], new_tmt_row[8] = "TMT", "Thuế bảo vệ môi trường", "VNĐ"
    new_tmt_row[9] = g5_val
    new_tmt_row[12] = total_quantity_for_tmt
    new_tmt_row[13] = tmt_unit_value_for_summary
    new_tmt_row[14] = round(to_float(total_quantity_for_tmt) * to_float(tmt_unit_value_for_summary), 0)
    new_tmt_row[18], new_tmt_row[19] = s_lookup.get(h5_val, ''), t_lookup.get(h5_val, '')
    new_tmt_row[20], new_tmt_row[21] = u_val, v_lookup.get(h5_val, '')
    new_tmt_row[36] = total_bvmt_amount
    new_tmt_row[31] = ""
    for idx in [5,10,11,15,16,17,22,23,24,25,26,27,28,29,30,32,33,34,35]:
        if idx < len(new_tmt_row): new_tmt_row[idx] = ''
    return new_tmt_row

# --- Tải dữ liệu tĩnh ---
static_data = get_static_data_from_excel(DATA_FILE_PATH)
listbox_data = static_data["listbox_data"]
lookup_table = static_data["lookup_table"]
tmt_lookup_table = static_data["tmt_lookup_table"]
s_lookup_table = static_data["s_lookup_table"]
t_lookup_table = static_data["t_lookup_table"]
v_lookup_table = static_data["v_lookup_table"]
u_value = static_data["u_value"]
chxd_detail_map = static_data["chxd_detail_map"]
# ******** START CHANGE / BẮT ĐẦU THAY ĐỔI ********
store_specific_x_lookup = static_data["store_specific_x_lookup"]
# ******** END CHANGE / KẾT THÚC THAY ĐỔI ********

# --- Giao diện người dùng Streamlit ---
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=100)
    st.markdown("""<div style="text-align: center;"><h1 style="color: red; font-size: 24px; margin-bottom: 0px;">CÔNG TY CỔ PHẦN XĂNG DẦU</h1><h2 style="color: red; font-size: 24px; margin-top: 0px;">DẦU KHÍ NAM ĐỊNH</h2></div>""", unsafe_allow_html=True)
st.title("Đồng bộ dữ liệu SSE")
selected_value = st.selectbox("Chọn CHXD:", options=[""] + listbox_data, key='selected_chxd')
uploaded_file = st.file_uploader("Tải lên file bảng kê hóa đơn (.xlsx)", type=["xlsx"])

# --- Xử lý chính ---
if st.button("Xử lý", key='process_button'):
    if not selected_value: st.warning("Vui lòng chọn một giá trị từ danh sách CHXD.")
    elif uploaded_file is None: st.warning("Vui lòng tải lên file bảng kê hóa đơn.")
    else:
        try:
            selected_value_normalized = clean_string(selected_value)
            if selected_value_normalized not in chxd_detail_map:
                st.error(f"Không tìm thấy thông tin chi tiết cho CHXD: '{selected_value_normalized}'")
                st.stop()
            
            chxd_details = chxd_detail_map[selected_value_normalized]
            g5_value, h5_value, f5_value_full, b5_value = chxd_details['g5_val'], chxd_details['h5_val'], chxd_details['f5_val_full'], chxd_details['b5_val']
            # ******** START CHANGE / BẮT ĐẦU THAY ĐỔI ********
            # Lấy bảng tra cứu "Vụ việc" cho cửa hàng đã chọn
            x_lookup_for_store = store_specific_x_lookup.get(selected_value_normalized, {})
            if not x_lookup_for_store:
                st.warning(f"Không tìm thấy mã Vụ việc cho cửa hàng '{selected_value_normalized}' trong Data.xlsx.")
            # ******** END CHANGE / KẾT THÚC THAY ĐỔI ********

            bkhd_wb = load_workbook(uploaded_file)
            bkhd_ws = bkhd_wb.active

            long_cells = [f"H{r_idx+1}" for r_idx, cell in enumerate(bkhd_ws['H']) if cell.value and len(str(cell.value)) > 128]
            if long_cells:
                st.error("Địa chỉ trên ô " + ', '.join(long_cells) + " quá dài, hãy điều chỉnh và thử lại.")
                st.stop()

            # --- Chuẩn bị dữ liệu bảng kê ---
            all_rows_from_bkhd = list(bkhd_ws.iter_rows(values_only=True))
            temp_bkhd_data_for_processing = all_rows_from_bkhd[4:] if len(all_rows_from_bkhd) >= 4 else []
            
            vi_tri_cu_idx = [0, 1, 2, 3, 4, 5, 7, 6, 8, 10, 11, 13, 14, 16]
            intermediate_data_rows = []
            for row_values_original in temp_bkhd_data_for_processing:
                if len(row_values_original) <= max(vi_tri_cu_idx): continue
                new_row_for_temp = [row_values_original[col_old_idx] for col_old_idx in vi_tri_cu_idx]
                # Chuyển đổi ngày
                if new_row_for_temp[3]:
                    try: new_row_for_temp[3] = datetime.strptime(str(new_row_for_temp[3])[:10], '%d-%m-%Y').strftime('%Y-%m-%d')
                    except ValueError: pass
                # Thêm cột "Công nợ"
                ma_kh_value = new_row_for_temp[4]
                new_row_for_temp.append("No" if ma_kh_value is None or len(clean_string(ma_kh_value)) > 9 else "Yes")
                intermediate_data_rows.append(new_row_for_temp)

            # --- Kiểm tra mã kho ---
            if not intermediate_data_rows:
                st.error("Không có dữ liệu hợp lệ trong file bảng kê sau khi xử lý.")
                st.stop()

            b2_bkhd_value = clean_string(intermediate_data_rows[0][1])
            normalized_f5_value_full = clean_string(f5_value_full)
            if normalized_f5_value_full.startswith('1'): normalized_f5_value_full = normalized_f5_value_full[1:]
            if normalized_f5_value_full != b2_bkhd_value:
                st.error("Bảng kê hóa đơn không phải của cửa hàng bạn chọn.")
                st.stop()

            # --- Vòng lặp xử lý chính ---
            final_upsse_output_rows = [[''] * len(headers) for _ in range(4)] + [headers]
            all_tmt_rows = []
            no_invoice_rows = {"Xăng E5 RON 92-II": [], "Xăng RON 95-III": [], "Dầu DO 0,05S-II": [], "Dầu DO 0,001S-V": []}

            for row_values_from_bkhd in intermediate_data_rows:
                new_row_for_upsse = [''] * len(headers)
                cong_no_status, ma_kh = row_values_from_bkhd[-1], row_values_from_bkhd[4]
                new_row_for_upsse[0] = clean_string(ma_kh) if cong_no_status == 'Yes' and ma_kh and clean_string(ma_kh) != '' else g5_value
                new_row_for_upsse[1], new_row_for_upsse[2] = clean_string(row_values_from_bkhd[5]), row_values_from_bkhd[3]
                
                b_orig, c_orig = clean_string(row_values_from_bkhd[1]), clean_string(row_values_from_bkhd[2])
                if b5_value == "Nguyễn Huệ": new_row_for_upsse[3] = f"HN{c_orig[-6:]}"
                elif b5_value == "Mai Linh": new_row_for_upsse[3] = f"MM{c_orig[-6:]}"
                else: new_row_for_upsse[3] = f"{b_orig[-2:]}{c_orig[-6:]}"
                new_row_for_upsse[4] = f"1{b_orig}" if b_orig else ''
                new_row_for_upsse[5] = f"Xuất bán lẻ theo hóa đơn số {new_row_for_upsse[3]}"

                product_name = clean_string(row_values_from_bkhd[8])
                new_row_for_upsse[7] = product_name
                new_row_for_upsse[6] = lookup_table.get(product_name.lower(), '')
                new_row_for_upsse[8], new_row_for_upsse[9] = "Lít", g5_value
                new_row_for_upsse[10], new_row_for_upsse[11] = '', ''

                new_row_for_upsse[12] = to_float(row_values_from_bkhd[9])
                tmt_value = tmt_lookup_table.get(product_name.lower(), 0.0)
                new_row_for_upsse[13] = round(to_float(row_values_from_bkhd[10]) / 1.1 - tmt_value, 2)
                tmt_calc = round(tmt_value * new_row_for_upsse[12])
                new_row_for_upsse[14] = to_float(row_values_from_bkhd[11]) - tmt_calc

                new_row_for_upsse[15], new_row_for_upsse[16] = '', ''
                new_row_for_upsse[17] = 10
                new_row_for_upsse[18], new_row_for_upsse[19] = s_lookup_table.get(h5_value, ''), t_lookup_table.get(h5_value, '')
                new_row_for_upsse[20], new_row_for_upsse[21] = u_value, v_lookup_table.get(h5_value, '')
                new_row_for_upsse[22] = ''
                # ******** START CHANGE / BẮT ĐẦU THAY ĐỔI ********
                new_row_for_upsse[23] = x_lookup_for_store.get(product_name.lower(), '')
                # ******** END CHANGE / KẾT THÚC THAY ĐỔI ********
                for i in range(24, 31): new_row_for_upsse[i] = ''
                
                new_row_for_upsse[31] = new_row_for_upsse[1]
                new_row_for_upsse[32], new_row_for_upsse[33] = row_values_from_bkhd[6], row_values_from_bkhd[7]
                new_row_for_upsse[34], new_row_for_upsse[35] = '', ''
                
                thue_tmt = round(new_row_for_upsse[12] * tmt_value * 0.1)
                new_row_for_upsse[36] = to_float(row_values_from_bkhd[12]) - thue_tmt

                if new_row_for_upsse[1] == "Người mua không lấy hóa đơn" and product_name in no_invoice_rows:
                    no_invoice_rows[product_name].append(new_row_for_upsse)
                else:
                    final_upsse_output_rows.append(new_row_for_upsse)
                    if tmt_value > 0 and new_row_for_upsse[12] > 0:
                        all_tmt_rows.append(create_per_invoice_tmt_row(new_row_for_upsse, tmt_value, g5_value, s_lookup_table, t_lookup_table, v_lookup_table, u_value, h5_value))

            # --- Thêm các dòng tổng hợp "Người mua không lấy hóa đơn" ---
            for product_name, rows in no_invoice_rows.items():
                if rows:
                    summary_row = add_summary_row_for_no_invoice(rows, bkhd_ws, product_name, headers, g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_for_store, u_value, h5_value, lookup_table)
                    final_upsse_output_rows.append(summary_row)
                    total_bvmt = sum(round(to_float(r[12]) * tmt_lookup_table.get(clean_string(r[7]).lower(), 0) * 0.1, 0) for r in rows)
                    if total_bvmt > 0:
                        tmt_unit = tmt_lookup_table.get(product_name.lower(), 0)
                        total_qty = sum(to_float(r[12]) for r in rows)
                        all_tmt_rows.append(add_tmt_summary_row(product_name, total_bvmt, g5_value, s_lookup_table, t_lookup_table, v_lookup_table, u_value, h5_value, summary_row[2], summary_row[4], total_qty, tmt_unit, b5_value, summary_row[1]))

            final_upsse_output_rows.extend(all_tmt_rows)

            # --- Ghi file kết quả ---
            up_sse_wb_final = Workbook()
            up_sse_ws_final = up_sse_wb_final.active
            for row_data in final_upsse_output_rows: up_sse_ws_final.append(row_data)

            text_style, date_style = NamedStyle(name="text_style", number_format='@'), NamedStyle(name="date_style", number_format='DD/MM/YYYY')
            exclude_cols = {3, 13, 14, 15, 18, 19, 20, 21, 22, 37}
            for r in range(1, up_sse_ws_final.max_row + 1):
                for c in range(1, up_sse_ws_final.max_column + 1):
                    cell = up_sse_ws_final.cell(row=r, column=c)
                    if not cell.value or clean_string(cell.value) == "None": continue
                    if c == 3:
                        try:
                            cell.value = datetime.strptime(clean_string(cell.value), '%Y-%m-%d').date()
                            cell.style = date_style
                        except (ValueError, TypeError): pass
                    elif c not in exclude_cols: cell.style = text_style
            for r in range(1, up_sse_ws_final.max_row + 1):
                for c in range(18, 23): up_sse_ws_final.cell(row=r, column=c).number_format = '@'

            up_sse_ws_final.column_dimensions['B'].width = 35
            up_sse_ws_final.column_dimensions['C'].width = 12
            up_sse_ws_final.column_dimensions['D'].width = 12

            output = io.BytesIO()
            up_sse_wb_final.save(output)
            st.success("Đã tạo file UpSSE.xlsx thành công!")
            st.download_button("Tải xuống file UpSSE.xlsx", output.getvalue(), "UpSSE.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"Lỗi trong quá trình xử lý file: {e}")
            st.exception(e)
