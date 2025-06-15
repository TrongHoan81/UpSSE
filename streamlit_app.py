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
        if isinstance(value, str):
            # Xử lý trường hợp chuỗi rỗng hoặc "None"
            if not value.strip() or value.strip().lower() == "none":
                return 0.0
            value = value.replace(",", "").strip()
        return float(value)
    except (ValueError, TypeError):
        return 0.0

# --- Hàm làm sạch chuỗi (loại bỏ mọi loại khoảng trắng và chuẩn hóa) ---
def clean_string(s):
    if s is None:
        return ""
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

        listbox_data = []
        chxd_detail_map = {}
        store_specific_x_lookup = {}
        
        for row_idx in range(4, ws.max_row + 1):
            row_data_values = [cell.value for cell in ws[row_idx]]

            if len(row_data_values) < 18: continue

            raw_chxd_name = row_data_values[10]
            if raw_chxd_name and clean_string(raw_chxd_name):
                chxd_name_str = clean_string(raw_chxd_name)
                
                if chxd_name_str and chxd_name_str not in listbox_data:
                    listbox_data.append(chxd_name_str)

                g5_val = row_data_values[15] if pd.notna(row_data_values[15]) else None
                f5_val_full = clean_string(row_data_values[16]) if pd.notna(row_data_values[16]) else ''
                h5_val = clean_string(row_data_values[17]).lower() if pd.notna(row_data_values[17]) else ''
                
                if f5_val_full:
                    chxd_detail_map[chxd_name_str] = {
                        'g5_val': g5_val, 'h5_val': h5_val,
                        'f5_val_full': f5_val_full, 'b5_val': chxd_name_str # Đây là B5.value từ Data.xlsx
                    }
                
                # Cập nhật x_lookup_for_store với các giá trị từ cột L, M, N, O (index 11,12,13,14) trong Data.xlsx
                store_specific_x_lookup[chxd_name_str] = {
                    "xăng e5 ron 92-ii": row_data_values[11], # Cột L
                    "xăng ron 95-iii":   row_data_values[12], # Cột M
                    "dầu do 0,05s-ii":   row_data_values[13], # Cột N
                    "dầu do 0,001s-v":   row_data_values[14]  # Cột O
                }
        
        lookup_table = {} # Bảng tra cứu Mã hàng (cột I, J)
        for row in ws.iter_rows(min_row=4, max_row=7, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: lookup_table[clean_string(row[0]).lower()] = row[1]
        
        tmt_lookup_table = {} # Bảng tra cứu Thuế BVMT (cột I, J)
        for row in ws.iter_rows(min_row=10, max_row=13, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: tmt_lookup_table[clean_string(row[0]).lower()] = to_float(row[1])
        
        s_lookup_table = {} # Bảng tra cứu Tk nợ (cột I, J)
        for row in ws.iter_rows(min_row=29, max_row=31, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: s_lookup_table[clean_string(row[0]).lower()] = row[1]
        
        t_lookup_regular = {} # Bảng tra cứu Tk doanh thu (cột I, J) cho hóa đơn thông thường
        for row in ws.iter_rows(min_row=33, max_row=35, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: t_lookup_regular[clean_string(row[0]).lower()] = row[1]
        
        t_lookup_tmt = {} # Bảng tra cứu Tk doanh thu (cột I, J) cho thuế BVMT
        for row in ws.iter_rows(min_row=48, max_row=50, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: t_lookup_tmt[clean_string(row[0]).lower()] = row[1]

        v_lookup_table = {} # Bảng tra cứu Tk thuế có (cột I, J)
        for row in ws.iter_rows(min_row=53, max_row=55, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: v_lookup_table[clean_string(row[0]).lower()] = row[1]
        
        u_value = ws['J36'].value # Giá trị ô J36 (Tk giá vốn)
        wb.close()
        
        return {
            "listbox_data": listbox_data, "lookup_table": lookup_table,
            "tmt_lookup_table": tmt_lookup_table, "s_lookup_table": s_lookup_table,
            "t_lookup_regular": t_lookup_regular, "t_lookup_tmt": t_lookup_tmt,
            "v_lookup_table": v_lookup_table, "u_value": u_value,
            "chxd_detail_map": chxd_detail_map, "store_specific_x_lookup": store_specific_x_lookup
        }
    except FileNotFoundError:
        st.error(f"Lỗi: Không tìm thấy file {file_path}. Vui lòng đảm bảo file tồn tại.")
        st.stop()
    except Exception as e:
        st.error(f"Lỗi không mong muốn khi đọc file Data.xlsx: {e}")
        st.exception(e)
        st.stop()

def add_summary_row_for_no_invoice(data_for_summary_product, bkhd_source_ws, product_name, headers_list,
                    g5_val, b5_val, s_lookup, t_lookup, v_lookup, x_lookup_for_store, u_val, h5_val, common_lookup_table, tmt_lookup_table):
    new_row = [''] * len(headers_list)
    new_row[0], new_row[1] = g5_val, f"Khách hàng mua {product_name} không lấy hóa đơn"
    new_row[2] = data_for_summary_product[0][2] if data_for_summary_product else ""
    new_row[4] = data_for_summary_product[0][4] if data_for_summary_product else ""
    value_C, value_E = clean_string(new_row[2]), clean_string(new_row[4])
    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name, "")
    
    # Logic tính số hóa đơn tổng hợp
    if b5_val == "Nguyễn Huệ": new_row[3] = f"HNBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh": new_row[3] = f"MMBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    else: new_row[3] = f"{value_E[-2:]}BK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    
    new_row[5] = f"Xuất bán lẻ theo hóa đơn số {new_row[3]}"
    new_row[6], new_row[7], new_row[8], new_row[9] = common_lookup_table.get(clean_string(product_name).lower(), ''), product_name, "Lít", g5_val
    new_row[10], new_row[11] = '', ''
    
    # Tính tổng số lượng từ các dòng chi tiết của khách vãng lai cho mặt hàng này
    total_M = sum(to_float(r[12]) for r in data_for_summary_product)
    new_row[12] = total_M

    # Lấy giá bán cao nhất từ các dòng chi tiết
    new_row[13] = max((to_float(r[13]) for r in data_for_summary_product), default=0.0)

    # Tính Tiền hàng tổng hợp (Cột O trong UpSSE)
    # Lấy tổng tiền hàng từ BKHD gốc cho các hóa đơn khách vãng lai của mặt hàng này
    tien_hang_hd = sum(to_float(r[11]) for r in bkhd_source_ws.iter_rows(min_row=2, values_only=True) if clean_string(r[5]) == "Người mua không lấy hóa đơn" and clean_string(r[8]) == product_name)
    
    # Lấy giá đơn vị (price_per_liter) để tính toán dựa trên UpSSE.2025.py
    # Đây là giá trị đơn vị (ví dụ 1900 cho E5, 2000 cho 95, 1000 cho DO)
    price_per_liter = {"Xăng E5 RON 92-II": 1900, "Xăng RON 95-III": 2000, "Dầu DO 0,05S-II": 1000, "Dầu DO 0,001S-V": 1000}.get(product_name, 0)
    new_row[14] = tien_hang_hd - round(total_M * price_per_liter, 0) # Tiền hàng tổng = Tổng Tiền hàng BKHD - (Tổng số lượng * Giá đơn vị)

    new_row[15], new_row[16], new_row[17] = '', '', 10
    new_row[18], new_row[19] = s_lookup.get(h5_val, ''), t_lookup.get(h5_val, '')
    new_row[20], new_row[21] = u_val, v_lookup.get(h5_val, '')
    new_row[22] = ''
    new_row[23] = x_lookup_for_store.get(clean_string(product_name).lower(), '')
    for i in range(24, 31): new_row[i] = ''
    new_row[31] = f"Khách mua {product_name} không lấy hóa đơn"
    new_row[32], new_row[33], new_row[34], new_row[35] = "", "", '', ''
    
    # --- ĐIỀU CHỈNH CÔNG THỨC CỘT AK (Tiền thuế) CHO DÒNG TỔNG HỢP TIỀN HÀNG ---
    # Theo yêu cầu: "Ở dòng thể hiện tiền hàng thì nó bằng tổng tiền thuế (ô AK)
    # của tất cả các ô tiền thuế (AK) của các hóa đơn mà khách hàng là 'Người mua không lấy hóa đơn'"
    
    # Tính tổng Tiền thuế (AK) từ các dòng chi tiết đã được xử lý (no_invoice_rows)
    total_ak_from_individual_rows = sum(to_float(r[36]) for r in data_for_summary_product)
    new_row[36] = total_ak_from_individual_rows # Gán trực tiếp tổng AK của các dòng con
    # --- KẾT THÚC ĐIỀU CHỈNH ---

    return new_row

def create_per_invoice_tmt_row(original_row_data, tmt_value, g5_val, s_lookup, t_lookup_tmt, v_lookup, u_val, h5_val):
    tmt_row = list(original_row_data)
    tmt_row[6], tmt_row[7], tmt_row[8] = "TMT", "Thuế bảo vệ môi trường", "VNĐ"
    tmt_row[9] = g5_val
    tmt_row[13] = tmt_value
    tmt_row[14] = round(tmt_value * to_float(original_row_data[12]), 0) # Giá trị Tiền hàng cho dòng TMT
    tmt_row[17] = 10
    tmt_row[18] = s_lookup.get(h5_val, '')
    tmt_row[19] = t_lookup_tmt.get(h5_val, '')
    tmt_row[20], tmt_row[21] = u_val, v_lookup.get(h5_val, '')
    tmt_row[31] = ""
    tmt_row[36] = round(tmt_value * to_float(original_row_data[12]) * 0.1, 0) # Tiền thuế cho dòng TMT
    for idx in [5, 10, 11, 15, 16, 22, 24, 25, 26, 27, 28, 29, 30, 32, 33, 34, 35]:
        if idx < len(tmt_row): tmt_row[idx] = ''
    return tmt_row

def add_tmt_summary_row(product_name_full, total_bvmt_amount, g5_val, s_lookup, t_lookup_tmt, v_lookup, u_val, h5_val, 
                        representative_date, representative_symbol, total_quantity_for_tmt, tmt_unit_value_for_summary, b5_val, customer_name_for_summary_row, x_lookup_for_store):
    new_tmt_row = [''] * len(headers)
    new_tmt_row[0], new_tmt_row[1], new_tmt_row[2] = g5_val, customer_name_for_summary_row, representative_date
    value_C, value_E = clean_string(representative_date), clean_string(representative_symbol)
    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name_full, "")
    
    # Logic tính số hóa đơn tổng hợp
    if b5_val == "Nguyễn Huệ": new_tmt_row[3] = f"HNBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh": new_tmt_row[3] = f"MMBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    else: new_tmt_row[3] = f"{value_E[-2:]}BK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    
    new_tmt_row[4] = representative_symbol
    new_tmt_row[6], new_tmt_row[7], new_tmt_row[8] = "TMT", "Thuế bảo vệ môi trường", "VNĐ"
    new_tmt_row[9], new_tmt_row[12] = g5_val, total_quantity_for_tmt
    new_tmt_row[13] = tmt_unit_value_for_summary
    new_tmt_row[14] = round(to_float(total_quantity_for_tmt) * to_float(tmt_unit_value_for_summary), 0) # Tiền hàng cho dòng TMT tổng hợp
    new_tmt_row[17] = 10
    new_tmt_row[18] = s_lookup.get(h5_val, '')
    new_tmt_row[19] = t_lookup_tmt.get(h5_val, '')
    new_tmt_row[20], new_tmt_row[21] = u_val, v_lookup.get(h5_val, '')
    new_tmt_row[23] = x_lookup_for_store.get(product_name_full.lower(), '')
    new_tmt_row[36], new_tmt_row[31] = total_bvmt_amount, "" # Tiền thuế (AK) cho dòng TMT tổng hợp
    for idx in [5,10,11,15,16,22,24,25,26,27,28,29,30,32,33,34,35]:
        if idx != 23 and idx < len(new_tmt_row): new_tmt_row[idx] = ''
    return new_tmt_row

# --- Tải dữ liệu tĩnh ---
static_data = get_static_data_from_excel(DATA_FILE_PATH)
listbox_data = static_data["listbox_data"]
lookup_table = static_data["lookup_table"]
tmt_lookup_table = static_data["tmt_lookup_table"]
s_lookup_table = static_data["s_lookup_table"]
t_lookup_regular = static_data["t_lookup_regular"]
t_lookup_tmt = static_data["t_lookup_tmt"]
v_lookup_table = static_data["v_lookup_table"]
u_value = static_data["u_value"]
chxd_detail_map = static_data["chxd_detail_map"]
store_specific_x_lookup = static_data["store_specific_x_lookup"]

# --- Giao diện người dùng Streamlit ---
col1, col2 = st.columns([1, 4]) 
with col1:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=140)
with col2:
    st.markdown("""
    <div style="display: flex; flex-direction: column; justify-content: center; align-items: center; height: 100%; padding-top: 10px;">
        <h2 style="color: red; font-weight: bold; margin-bottom: 0px; font-size: 24px;">CÔNG TY CỔ PHẦN XĂNG DẦU</h2>
        <h2 style="color: red; font-weight: bold; margin-top: 0px; font-size: 24px;">DẦU KHÍ NAM ĐỊNH</h2>
    </div>
    """, unsafe_allow_html=True)

st.title("Đồng bộ dữ liệu SSE")

st.markdown("""
<style>
@keyframes blinker { 50% { opacity: 0.7; } }
.blinking-warning { padding: 12px; background-color: #FFFACD; border: 1px solid #FFD700; border-radius: 8px; text-align: center; animation: blinker 1.5s linear infinite; }
.blinking-warning p { color: #DC143C; font-weight: bold; margin: 0; font-size: 16px; }
</style>
<div class="blinking-warning">
  <p>Lưu ý quan trọng: Để tránh lỗi, sau khi tải file bảng kê từ POS về, bạn hãy mở lên và lưu lại (ấn Ctrl+S hoặc chọn File/Save) trước khi đưa vào ứng dụng để xử lý.</p>
</div>
<br>
""", unsafe_allow_html=True)

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
            x_lookup_for_store = store_specific_x_lookup.get(selected_value_normalized, {})
            if not x_lookup_for_store:
                st.warning(f"Không tìm thấy mã Vụ việc cho cửa hàng '{selected_value_normalized}' trong Data.xlsx.")

            bkhd_wb = load_workbook(uploaded_file)
            bkhd_ws = bkhd_wb.active

            long_cells = [f"H{r_idx+1}" for r_idx, cell in enumerate(bkhd_ws['H']) if cell.value and len(str(cell.value)) > 128]
            if long_cells:
                st.error("Địa chỉ trên ô " + ', '.join(long_cells) + " quá dài, hãy điều chỉnh và thử lại.")
                st.stop()

            # Bỏ qua 4 dòng đầu tiên của BKHD (tiêu đề không cần thiết)
            all_rows_from_bkhd = list(bkhd_ws.iter_rows(values_only=True))
            temp_bkhd_data = all_rows_from_bkhd[4:] if len(all_rows_from_bkhd) >= 4 else []
            
            # Định nghĩa lại vị trí các cột cần lấy từ BKHD gốc để tạo intermediate_data
            # Dựa trên phân tích UpSSE.2025.py và cấu trúc dữ liệu, các cột được ánh xạ như sau:
            # Original Col Index -> intermediate_data Index
            # A (0) -> (không sử dụng trực tiếp trong vòng lặp chính tạo upsse_row)
            # B (1) -> intermediate_data[1]
            # C (2) -> intermediate_data[2]
            # D (3) -> intermediate_data[3] (Ngày)
            # E (4) -> intermediate_data[4] (Mã khách)
            # F (5) -> intermediate_data[5] (Tên khách hàng)
            # G (6) -> intermediate_data[7] (Tên KH(thuế))
            # H (7) -> intermediate_data[6] (Địa chỉ (thuế))
            # I (8) -> intermediate_data[8] (Tên mặt hàng)
            # J (9) -> intermediate_data[9] (Số lượng)
            # K (10) -> intermediate_data[10] (Giá bán - original)
            # L (11) -> intermediate_data[11] (Tiền hàng - original)
            # M (12) -> (không sử dụng trực tiếp cho intermediate_data - Tiền thuế BKHD gốc, sẽ dùng cho tien_thue_hd)
            # N (13) -> intermediate_data[12] (Giá bán - nguyên gốc từ BKHD, sau này dùng cho Tiền hàng UpSSE N)
            # O (14) -> intermediate_data[13] (Tiền hàng - nguyên gốc từ BKHD, sau này dùng cho Tiền hàng UpSSE O)
            # P (15) -> (không sử dụng trực tiếp cho intermediate_data)
            # Q (16) -> intermediate_data[12] (Tiền thuế - original, sau này dùng cho Tiền thuế UpSSE AK) -> Lỗi: Vị trí 12 bị trùng với N (13), nên thực tế row[12] sẽ là Q.
            # Cần chỉnh lại vi_tri_cu_idx để khớp với UpSSE.2025.py.
            # Trong UpSSE.2025.py, cột M của BKHD là Tiền thuế (r[12]).
            # Cột N của BKHD là Giá bán (r[13]). Cột O của BKHD là Tiền hàng (r[14])
            # Trong streamlit_app, `intermediate_data` được tạo từ `vi_tri_cu_idx`.
            # `vi_tri_cu_idx = [0, 1, 2, 3, 4, 5, 7, 6, 8, 10, 11, 13, 14, 16]`
            # Điều này có nghĩa là:
            # new_row[9] (Số lượng) = row[9] (Original J)
            # new_row[10] (Giá bán) = row[10] (Original K)
            # new_row[11] (Tiền hàng) = row[11] (Original L)
            # new_row[12] (Tiền thuế/Số lượng) = row[13] (Original N - Giá bán) HOẶC row[14] (Original O - Tiền hàng) HOẶC row[16] (Original Q - Tiền thuế)
            # Dựa trên việc sử dụng `row[12]` (cho `tien_thue_hd`) và `row[9]` (cho `upsse_row[12]`), có vẻ như:
            # row[9] (tức intermediate_data[9]) là Số lượng
            # row[12] (tức intermediate_data[12]) là Tiền thuế
            # Và trong BKHD gốc thì Tiền thuế là cột M (index 12), Số lượng là cột J (index 9)

            # Khớp lại index cho intermediate_data với UpSSE.2025.py logic:
            # BKHD gốc: D (Ngày - index 3), E (Mã khách - index 4), F (Tên khách hàng - index 5),
            # G (Tên KH(thuế) - index 6), H (Địa chỉ (thuế) - index 7), I (Tên mặt hàng - index 8),
            # J (Số lượng - index 9), K (Giá bán - index 10), L (Tiền hàng - index 11), M (Tiền thuế - index 12)
            #
            # Streamlit `intermediate_data` index (mục đích sử dụng sau này):
            # idx 0: (dường như không dùng trực tiếp)
            # idx 1: B (tên KH - original) -> upsse_row[1]
            # idx 2: C (Số HĐ - original) -> upsse_row[3]
            # idx 3: D (Ngày - original) -> upsse_row[2]
            # idx 4: E (Mã khách - original) -> upsse_row[0]
            # idx 5: F (Tên KH - original) -> upsse_row[1]
            # idx 6: H (Địa chỉ thuế - original) -> upsse_row[32]
            # idx 7: G (Tên KH thuế - original) -> upsse_row[31]
            # idx 8: I (Tên mặt hàng - original) -> upsse_row[7]
            # idx 9: J (Số lượng - original) -> upsse_row[12]
            # idx 10: K (Giá bán - original) -> upsse_row[13]
            # idx 11: L (Tiền hàng - original) -> upsse_row[14]
            # idx 12: M (Tiền thuế - original) -> upsse_row[36]

            # Dựa trên phân tích trên, `vi_tri_cu_idx` của streamlit_app.py
            # `vi_tri_cu_idx = [0, 1, 2, 3, 4, 5, 7, 6, 8, 10, 11, 13, 14, 16]`
            # Có vẻ như `row[12]` (intermediate_data[12]) đang lấy từ `original_row[16]` (cột Q).
            # Trong UpSSE.2025.py, `row[12]` là cột M (Tiền thuế).
            # Để khớp, `intermediate_data` cần lấy cột M (index 12) từ BKHD gốc cho `Tiền thuế`.
            # Và cột J (index 9) cho `Số lượng`.

            # Giả định cấu trúc của `temp_bkhd_data` (các dòng từ BKHD gốc sau khi bỏ 4 dòng đầu):
            # Các cột cần là: D, E, F, G, H, I, J, K, L, M
            # indices:         3, 4, 5, 6, 7, 8, 9, 10,11, 12
            
            # `vi_tri_cu_idx` nên là các index của cột trong `all_rows_from_bkhd` (original BKHD)
            # mà chúng ta muốn giữ. Các cột này sẽ trở thành các cột trong `new_row` (của `intermediate_data`).
            
            # Các cột cần để tạo upsse_row:
            # 0: Mã khách (từ E)
            # 1: Tên khách hàng (từ F)
            # 2: Ngày (từ D)
            # 3: Số hóa đơn (tự tạo)
            # 4: Ký hiệu (tự tạo)
            # 5: Diễn giải (tự tạo)
            # 6: Mã hàng (từ I và lookup)
            # 7: Tên mặt hàng (từ I)
            # 8: Đvt (cố định "Lít")
            # 9: Mã kho (từ G5)
            # 10: Mã vị trí (trống)
            # 11: Mã lô (trống)
            # 12: Số lượng (từ J)
            # 13: Giá bán (từ K và tính toán)
            # 14: Tiền hàng (từ L và tính toán)
            # 15: Mã nt (trống)
            # 16: Tỷ giá (trống)
            # 17: Mã thuế (cố định 10)
            # 18: Tk nợ (từ H5 và lookup)
            # 19: Tk doanh thu (từ H5 và lookup)
            # 20: Tk giá vốn (từ J36)
            # 21: Tk thuế có (từ H5 và lookup)
            # 22: Cục thuế (trống)
            # 23: Vụ việc (từ I và lookup)
            # 24-30: (trống)
            # 31: Tên KH(thuế) (từ F)
            # 32: Địa chỉ (thuế) (từ G)
            # 33: Mã số Thuế (từ H)
            # 34-35: (trống)
            # 36: Tiền thuế (từ M và tính toán)

            # Để đơn giản và khớp với cách truy cập của UpSSE.2025.py,
            # chúng ta sẽ lấy các cột cần thiết từ BKHD gốc (bkhd_ws)
            # và sử dụng index trực tiếp khi tạo `upsse_row`.
            # Bỏ qua việc tạo `intermediate_data` phức tạp nếu không cần thiết.
            
            # Tuy nhiên, `streamlit_app.py` đã định nghĩa `vi_tri_cu_idx` và sử dụng nó.
            # Nếu `row[12]` trong `intermediate_data` thực sự là `original_row[16]` (cột Q),
            # và `UpSSE.2025.py` sử dụng `original_row[12]` (cột M), đây sẽ là nguồn gốc của sai lệch.
            # Dựa trên file UpSSE.2025.py:
            # col_L (Tiền hàng) = row[11]
            # col_M (Tiền thuế) = row[12]
            # Vậy `streamlit_app.py` đang sử dụng sai index cho Tiền thuế của BKHD gốc (`row[12]` trong `intermediate_data` nên là `original_row[12]` chứ không phải `original_row[16]`).
            #
            # Điều chỉnh `vi_tri_cu_idx` để `intermediate_data[12]` là Tiền thuế (original BKHD column M, index 12)
            # Điều chỉnh `vi_tri_cu_idx[11]` để `intermediate_data[11]` là Tiền hàng (original BKHD column L, index 11)
            # Điều chỉnh `vi_tri_cu_idx[10]` để `intermediate_data[10]` là Giá bán (original BKHD column K, index 10)
            # Điều chỉnh `vi_tri_cu_idx[9]` để `intermediate_data[9]` là Số lượng (original BKHD column J, index 9)
            #
            # Danh sách các cột cần lấy từ BKHD gốc (theo thứ tự ban đầu của BKHD):
            # Index: 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, ...
            # Cột:   A, B, C, D, E, F, G, H, I, J, K,  L,  M, ...
            #
            # `vi_tri_cu_idx` phải ánh xạ đúng các cột của BKHD gốc sang `intermediate_data`:
            # D (Ngày): original_index=3 -> intermediate_index=3
            # E (Mã khách): original_index=4 -> intermediate_index=4
            # F (Tên khách hàng): original_index=5 -> intermediate_index=5
            # G (Tên KH(thuế)): original_index=6 -> intermediate_index=6
            # H (Địa chỉ (thuế)): original_index=7 -> intermediate_index=7
            # I (Tên mặt hàng): original_index=8 -> intermediate_index=8
            # J (Số lượng): original_index=9 -> intermediate_index=9
            # K (Giá bán): original_index=10 -> intermediate_index=10
            # L (Tiền hàng): original_index=11 -> intermediate_index=11
            # M (Tiền thuế): original_index=12 -> intermediate_index=12 (Đây là điểm mấu chốt!)

            # `vi_tri_cu_idx` hiện tại: `[0, 1, 2, 3, 4, 5, 7, 6, 8, 10, 11, 13, 14, 16]`
            # Các index cần sửa để lấy đúng cột:
            # Old:                                                    New:
            # vi_tri_cu_idx[9] (original K) = 10 (Giá bán)            -> original_index=9 (J - Số lượng)
            # vi_tri_cu_idx[10] (original L) = 11 (Tiền hàng)         -> original_index=10 (K - Giá bán)
            # vi_tri_cu_idx[11] (original N) = 13 (Giá bán - not K)  -> original_index=11 (L - Tiền hàng)
            # vi_tri_cu_idx[12] (original O) = 14 (Tiền hàng - not L) -> original_index=12 (M - Tiền thuế)

            # Cột cần trong `intermediate_data` và index BKHD gốc tương ứng:
            # intermediate_idx  -> Original BKHD Col (index)
            # 0                 -> A (0) (nếu dùng) - thực tế không dùng
            # 1                 -> B (1) (dùng cho Tên KH trong Số HĐ)
            # 2                 -> C (2) (dùng cho Số HĐ)
            # 3                 -> D (3) (Ngày)
            # 4                 -> E (4) (Mã khách)
            # 5                 -> F (5) (Tên khách hàng)
            # 6                 -> H (7) (Địa chỉ (thuế))
            # 7                 -> G (6) (Tên KH(thuế))
            # 8                 -> I (8) (Tên mặt hàng)
            # 9                 -> J (9) (Số lượng)
            # 10                -> K (10) (Giá bán)
            # 11                -> L (11) (Tiền hàng)
            # 12                -> M (12) (Tiền thuế)

            # `vi_tri_cu_idx` ban đầu của streamlit_app.py có vẻ là một lỗi copy/paste hoặc hiểu sai cấu trúc BKHD.
            # Nó đang lấy `original_row[13]` (N), `original_row[14]` (O), `original_row[16]` (Q) thay vì J, K, L, M.
            
            # Điều chỉnh `vi_tri_cu_idx` để nó lấy đúng các cột J, K, L, M (index 9, 10, 11, 12)
            # và các cột khác theo đúng thứ tự mà `intermediate_data` đang mong đợi.
            
            # Đây là các index CỦA CÁC CỘT TRONG BKHD GỐC:
            # Để đảm bảo consistent, tôi sẽ liệt kê các cột được sử dụng từ BKHD gốc và index của chúng:
            # Col B (1), Col C (2), Col D (3), Col E (4), Col F (5), Col G (6), Col H (7), Col I (8), Col J (9), Col K (10), Col L (11), Col M (12)
            
            # `temp_bkhd_data` (rows from original BKHD)
            # column mapping from original BKHD to `row` in `intermediate_data` loop
            # row[1]: Original B
            # row[2]: Original C
            # row[3]: Original D
            # row[4]: Original E
            # row[5]: Original F
            # row[6]: Original G
            # row[7]: Original H
            # row[8]: Original I
            # row[9]: Original J
            # row[10]: Original K
            # row[11]: Original L
            # row[12]: Original M

            # Streamlit `intermediate_data` (cấu trúc mới, đơn giản hơn, ánh xạ trực tiếp)
            # Sử dụng list comprehension thay vì `vi_tri_cu_idx` phức tạp.
            
            intermediate_data = []
            for row_original_bkhd in temp_bkhd_data:
                if len(row_original_bkhd) < 13: # Đảm bảo có đủ các cột cần thiết (ít nhất đến M/index 12)
                    continue

                new_row_intermediate = [
                    row_original_bkhd[1],   # Original Col B (for internal logic, e.g., for new_row[3] based on B)
                    row_original_bkhd[2],   # Original Col C (for internal logic, e.g., for new_row[3] based on C)
                    row_original_bkhd[3],   # Original Col D (Ngày -> upsse_row[2])
                    row_original_bkhd[4],   # Original Col E (Mã khách -> upsse_row[0])
                    row_original_bkhd[5],   # Original Col F (Tên khách hàng -> upsse_row[1])
                    row_original_bkhd[6],   # Original Col G (Tên KH(thuế) -> upsse_row[32])
                    row_original_bkhd[7],   # Original Col H (Địa chỉ (thuế) -> upsse_row[33])
                    row_original_bkhd[8],   # Original Col I (Tên mặt hàng -> upsse_row[7])
                    row_original_bkhd[9],   # Original Col J (Số lượng -> upsse_row[12])
                    row_original_bkhd[10],  # Original Col K (Giá bán -> upsse_row[13])
                    row_original_bkhd[11],  # Original Col L (Tiền hàng -> upsse_row[14])
                    row_original_bkhd[12]   # Original Col M (Tiền thuế -> upsse_row[36])
                ]
                
                # Chuyển đổi định dạng ngày cho cột D (index 2 trong new_row_intermediate)
                if new_row_intermediate[2]:
                    try:
                        # Cố gắng chuyển đổi từ nhiều định dạng ngày có thể có trong Excel
                        if isinstance(new_row_intermediate[2], datetime):
                            new_row_intermediate[2] = new_row_intermediate[2].strftime('%Y-%m-%d')
                        elif isinstance(new_row_intermediate[2], str):
                            # Thử các định dạng phổ biến
                            for fmt in ('%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d'):
                                try:
                                    date_obj = datetime.strptime(new_row_intermediate[2][:10], fmt)
                                    new_row_intermediate[2] = date_obj.strftime('%Y-%m-%d')
                                    break
                                except ValueError:
                                    continue
                    except ValueError:
                        pass # Giữ nguyên nếu không chuyển đổi được
                
                # Thêm cột "Công nợ" (Yes/No)
                ma_kh = new_row_intermediate[3] # Original E (Mã khách)
                new_row_intermediate.append("No" if ma_kh is None or len(clean_string(ma_kh)) > 9 else "Yes")
                
                intermediate_data.append(new_row_intermediate)


            if not intermediate_data:
                st.error("Không có dữ liệu hợp lệ trong file bảng kê sau khi xử lý.")
                st.stop()

            # Kiểm tra CHXD
            # Dòng đầu tiên của BKHD gốc chứa tên cửa hàng ở cột B (index 1)
            b2_bkhd = clean_string(all_rows_from_bkhd[1][1]) # All_rows_from_bkhd[1] là dòng thứ 2 (index 1) của file gốc
            f5_norm = clean_string(f5_value_full) # f5_value_full là F5.value từ Data.xlsx
            if f5_norm.startswith('1'): f5_norm = f5_norm[1:] # Xóa ký tự '1' nếu có
            if f5_norm != b2_bkhd:
                st.error("Bảng kê hóa đơn không phải của cửa hàng bạn chọn.")
                st.stop()

            final_rows = [[''] * len(headers) for _ in range(4)] + [headers] # Dòng tiêu đề và 4 dòng trống đầu
            all_tmt_rows = []
            no_invoice_rows = {p: [] for p in ["Xăng E5 RON 92-II", "Xăng RON 95-III", "Dầu DO 0,05S-II", "Dầu DO 0,001S-V"]}

            for row_intermediate in intermediate_data:
                upsse_row = [''] * len(headers)
                
                # Ánh xạ từ intermediate_data sang upsse_row (cấu trúc output cuối cùng)
                # Lưu ý: Index của `row_intermediate` khác với `headers` của `upsse_row`
                
                # Cột A (Mã khách): Từ original E (intermediate_data[3]), hoặc G5
                upsse_row[0] = clean_string(row_intermediate[3]) if row_intermediate[-1] == 'Yes' and row_intermediate[3] and clean_string(row_intermediate[3]) else g5_value
                
                # Cột B (Tên khách hàng): Từ original F (intermediate_data[4])
                upsse_row[1] = clean_string(row_intermediate[4])
                
                # Cột C (Ngày): Từ original D (intermediate_data[2])
                upsse_row[2] = row_intermediate[2]

                # Cột D (Số hóa đơn): Logic phức tạp
                b_orig_for_d = clean_string(row_intermediate[0]) # Original B
                c_orig_for_d = clean_string(row_intermediate[1]) # Original C
                if b5_value == "Nguyễn Huệ": upsse_row[3] = f"HN{c_orig_for_d[-6:]}"
                elif b5_value == "Mai Linh": upsse_row[3] = f"MM{c_orig_for_d[-6:]}"
                else: upsse_row[3] = f"{b_orig_for_d[-2:]}{c_orig_for_d[-6:]}"
                
                # Cột E (Ký hiệu): Từ original B (intermediate_data[0])
                upsse_row[4] = f"1{b_orig_for_d}" if b_orig_for_d else ''
                
                # Cột F (Diễn giải): Tự tạo
                upsse_row[5] = f"Xuất bán lẻ theo hóa đơn số {upsse_row[3]}"
                
                # Cột H (Tên mặt hàng): Từ original I (intermediate_data[7])
                product_name = clean_string(row_intermediate[7])
                upsse_row[7] = product_name
                
                # Cột G (Mã hàng): Dò tìm từ Tên mặt hàng (upsse_row[7])
                upsse_row[6] = lookup_table.get(product_name.lower(), '')
                
                # Cột I (Đvt): Cố định "Lít"
                upsse_row[8] = "Lít"
                
                # Cột J (Mã kho): Từ G5 của Data.xlsx
                upsse_row[9] = g5_value
                
                # Cột K và L (Mã vị trí, Mã lô): Để trống
                upsse_row[10], upsse_row[11] = '', ''
                
                # Cột M (Số lượng): Từ original J (intermediate_data[8])
                upsse_row[12] = to_float(row_intermediate[8])
                
                # Lấy giá trị TMT theo Tên mặt hàng (upsse_row[7])
                tmt_value = tmt_lookup_table.get(product_name.lower(), 0.0)
                
                # Cột N (Giá bán): Từ original K (intermediate_data[9]), tính toán
                upsse_row[13] = round(to_float(row_intermediate[9]) / 1.1 - tmt_value, 2)
                
                # Cột O (Tiền hàng): Từ original L (intermediate_data[10]), tính toán
                upsse_row[14] = to_float(row_intermediate[10]) - round(tmt_value * upsse_row[12])
                
                # Cột P, Q (Mã nt, Tỷ giá): Để trống
                upsse_row[15], upsse_row[16] = '', ''
                
                # Cột R (Mã thuế): Cố định 10
                upsse_row[17] = 10
                
                # Cột S (Tk nợ): Từ H5 (Data.xlsx) và lookup
                upsse_row[18] = s_lookup_table.get(h5_value, '')
                
                # Cột T (Tk doanh thu): Từ H5 (Data.xlsx) và lookup (regular)
                upsse_row[19] = t_lookup_regular.get(h5_value, '')
                
                # Cột U (Tk giá vốn): Từ J36 (Data.xlsx)
                upsse_row[20] = u_value
                
                # Cột V (Tk thuế có): Từ H5 (Data.xlsx) và lookup
                upsse_row[21] = v_lookup_table.get(h5_value, '')
                
                # Cột W (Cục thuế): Để trống
                upsse_row[22] = ''
                
                # Cột X (Vụ việc): Từ Tên mặt hàng (upsse_row[7]) và lookup riêng của cửa hàng
                upsse_row[23] = x_lookup_for_store.get(product_name.lower(), '')
                
                # Các cột Y-AE (Lsx, Sản phẩm, Hợp đồng, Phí, Khế ước, Nhân viên bán): Để trống
                for i in range(24, 31): upsse_row[i] = ''
                
                # Cột AF (Tên KH(thuế)): Từ Tên khách hàng (upsse_row[1])
                upsse_row[31] = upsse_row[1]
                
                # Cột AG (Địa chỉ (thuế)): Từ original G (intermediate_data[5])
                upsse_row[32] = row_intermediate[5]
                
                # Cột AH (Mã số Thuế): Từ original H (intermediate_data[6])
                upsse_row[33] = row_intermediate[6]
                
                # Cột AI, AJ (Nhóm Hàng, Ghi chú): Để trống
                upsse_row[34], upsse_row[35] = '', ''
                
                # Cột AK (Tiền thuế): Từ original M (intermediate_data[11]), tính toán
                # Sử dụng intermediate_data[11] (original L - Tiền hàng) cho `row[12]`
                # Sử dụng intermediate_data[12] (original M - Tiền thuế) cho `row[12]` (cho Tien_Thue_Tren_BKHD)
                # `original_row_data[12]` (intermediate_data[12]) là original M (Tiền thuế)
                # `original_row_data[8]` (intermediate_data[8]) là original J (Số lượng)
                tien_thue_tren_bkhd_goc = to_float(row_intermediate[11]) # Tiền thuế từ cột L của BKHD (đã được điều chỉnh)
                so_luong_goc = to_float(row_intermediate[8]) # Số lượng từ cột J của BKHD
                
                # Đây là công thức Tiền thuế (AK) cho MỘT DÒNG HÓA ĐƠN RIÊNG LẺ
                # Nó bằng Tiền thuế gốc từ BKHD trừ đi phần thuế BVMT đã tính toán dựa trên số lượng và TMT value
                upsse_row[36] = tien_thue_tren_bkhd_goc - round(so_luong_goc * tmt_value * 0.1)

                if upsse_row[1] == "Người mua không lấy hóa đơn" and product_name in no_invoice_rows:
                    no_invoice_rows[product_name].append(upsse_row)
                else:
                    final_rows.append(upsse_row)
                    if tmt_value > 0 and upsse_row[12] > 0: # Chỉ tạo dòng TMT nếu có thuế và số lượng > 0
                        all_tmt_rows.append(create_per_invoice_tmt_row(upsse_row, tmt_value, g5_value, s_lookup_table, t_lookup_tmt, v_lookup_table, u_value, h5_value))

            # Xử lý các dòng tổng hợp cho "Người mua không lấy hóa đơn"
            for product_name, rows in no_invoice_rows.items():
                if rows:
                    # Tạo dòng tổng hợp Tiền hàng
                    summary_row = add_summary_row_for_no_invoice(rows, bkhd_ws, product_name, headers, g5_value, b5_value, s_lookup_table, t_lookup_regular, v_lookup_table, x_lookup_for_store, u_value, h5_value, lookup_table, tmt_lookup_table)
                    final_rows.append(summary_row)
                    
                    # Tính tổng số tiền thuế BVMT cho dòng tổng hợp TMT
                    total_bvmt = sum(round(to_float(r[12]) * tmt_lookup_table.get(clean_string(r[7]).lower(), 0) * 0.1, 0) for r in rows)
                    if total_bvmt > 0:
                        tmt_unit = tmt_lookup_table.get(product_name.lower(), 0)
                        total_qty = sum(to_float(r[12]) for r in rows)
                        # Tạo dòng tổng hợp Tiền thuế BVMT
                        # `summary_row[2]` là Ngày, `summary_row[4]` là Ký hiệu
                        # `summary_row[1]` là Tên khách hàng
                        all_tmt_rows.append(add_tmt_summary_row(product_name, total_bvmt, g5_value, s_lookup_table, t_lookup_tmt, v_lookup_table, u_value, h5_value, summary_row[2], summary_row[4], total_qty, tmt_unit, b5_value, summary_row[1], x_lookup_for_store))

            final_rows.extend(all_tmt_rows)

            up_sse_wb_final = Workbook()
            up_sse_ws_final = up_sse_wb_final.active
            for row_data in final_rows: up_sse_ws_final.append(row_data)

            # Định dạng cột và kiểu dữ liệu
            text_style, date_style = NamedStyle(name="text_style", number_format='@'), NamedStyle(name="date_style", number_format='DD/MM/YYYY')
            # Cột không cần định dạng text (ví dụ: cột chứa số, ngày tháng tự động định dạng)
            # Cột C (Ngày), N (Giá bán), O (Tiền hàng), R (Mã thuế), AK (Tiền thuế)
            # Cột S, T, U, V cần định dạng Text (là mã tài khoản)
            exclude_cols_from_text_format = {3, 14, 15, 18, 37} # C (index 2), N (index 13), O (index 14), R (index 17), AK (index 36)
            # Dưới đây là các cột sẽ được định dạng là TEXT (mã tài khoản)
            text_only_cols = {19, 20, 21, 22} # S, T, U, V

            for r in range(1, up_sse_ws_final.max_row + 1):
                for c in range(1, up_sse_ws_final.max_column + 1):
                    cell = up_sse_ws_final.cell(row=r, column=c)
                    if not cell.value or clean_string(cell.value) == "None": continue
                    
                    # Định dạng cột Ngày (C)
                    if c == 3: # Cột C (index 2 + 1)
                        try:
                            # Đảm bảo giá trị là chuỗi YYYY-MM-DD trước khi chuyển đổi
                            if isinstance(cell.value, datetime):
                                cell.value = cell.value.date() # Lấy phần ngày nếu là datetime object
                            elif isinstance(cell.value, str):
                                cell.value = datetime.strptime(clean_string(cell.value), '%Y-%m-%d').date()
                            cell.style = date_style
                        except (ValueError, TypeError): 
                            pass # Giữ nguyên định dạng nếu không phải ngày hợp lệ
                    # Định dạng các cột mã tài khoản (S, T, U, V) thành Text
                    elif c in text_only_cols:
                        cell.number_format = '@'
                    # Các cột khác (trừ những cột loại trừ) định dạng Text
                    elif c not in exclude_cols_from_text_format:
                        cell.style = text_style # Áp dụng kiểu text đã định nghĩa

            # Điều chỉnh chiều rộng cột
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

