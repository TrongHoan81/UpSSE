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
        
        # Đọc trực tiếp giá trị từ ô B5 và F5 trong Data.xlsx
        # Đây là thay đổi quan trọng để khớp với logic của UpSSE.2025.py
        b5_val_from_data = clean_string(ws['B5'].value) if pd.notna(ws['B5'].value) else ''
        f5_val_full_raw = clean_string(ws['F5'].value) if pd.notna(ws['F5'].value) else ''
        # Lấy 6 ký tự cuối cùng của F5, tương tự UpSSE.2025.py
        f5_val_for_comparison = f5_val_full_raw[-6:]

        for row_idx in range(4, ws.max_row + 1):
            row_data_values = [cell.value for cell in ws[row_idx]]

            if len(row_data_values) < 18: continue

            raw_chxd_name = row_data_values[10] # Cột K: Tên CHXD để đưa vào danh sách lựa chọn
            if raw_chxd_name and clean_string(raw_chxd_name):
                chxd_name_str = clean_string(raw_chxd_name)
                
                if chxd_name_str and chxd_name_str not in listbox_data:
                    listbox_data.append(chxd_name_str)

                g5_val = row_data_values[15] if pd.notna(row_data_values[15]) else None # Cột P: Mã kho (tương ứng G5)
                h5_val = clean_string(row_data_values[17]).lower() if pd.notna(row_data_values[17]) else '' # Cột R: Mã chi nhánh (tương ứng H5)
                
                chxd_detail_map[chxd_name_str] = {
                    'g5_val': g5_val,
                    'h5_val': h5_val,
                    'f5_val_full': f5_val_for_comparison, # Sử dụng giá trị đã xử lý 6 ký tự cuối
                    'b5_val': b5_val_from_data # Giá trị B5.value từ Data.xlsx
                }
                
                # Cập nhật x_lookup_for_store với các giá trị từ cột L, M, N, O (index 11,12,13,14) trong Data.xlsx
                store_specific_x_lookup[chxd_name_str] = {
                    "xăng e5 ron 92-ii": row_data_values[11], # Cột L
                    "xăng ron 95-iii":   row_data_values[12], # Cột M
                    "dầu do 0,05s-ii":   row_data_values[13], # Cột N
                    "dầu do 0,001s-v":   row_data_values[14]  # Cột O
                }
        
        lookup_table = {} # Bảng tra cứu Mã hàng (cột I, J) from Data.xlsx (rows 4-7)
        for row in ws.iter_rows(min_row=4, max_row=7, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: lookup_table[clean_string(row[0]).lower()] = row[1]
        
        tmt_lookup_table = {} # Bảng tra cứu Thuế BVMT (cột I, J) from Data.xlsx (rows 10-13)
        for row in ws.iter_rows(min_row=10, max_row=13, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: tmt_lookup_table[clean_string(row[0]).lower()] = to_float(row[1])
        
        s_lookup_table = {} # Bảng tra cứu Tk nợ (cột I, J) from Data.xlsx (rows 29-31)
        for row in ws.iter_rows(min_row=29, max_row=31, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: s_lookup_table[clean_string(row[0]).lower()] = row[1]
        
        t_lookup_regular = {} # Bảng tra cứu Tk doanh thu (cột I, J) cho hóa đơn thông thường from Data.xlsx (rows 33-35)
        for row in ws.iter_rows(min_row=33, max_row=35, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: t_lookup_regular[clean_string(row[0]).lower()] = row[1]
        
        t_lookup_tmt = {} # Bảng tra cứu Tk doanh thu (cột I, J) cho thuế BVMT from Data.xlsx (rows 48-50)
        for row in ws.iter_rows(min_row=48, max_row=50, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]: t_lookup_tmt[clean_string(row[0]).lower()] = row[1]

        v_lookup_table = {} # Bảng tra cứu Tk thuế có (cột I, J) from Data.xlsx (rows 53-55)
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
            g5_value, h5_value, f5_value_for_comparison, b5_value = chxd_details['g5_val'], chxd_details['h5_val'], chxd_details['f5_val_full'], chxd_details['b5_val']
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
            
            intermediate_data = []
            for row_original_bkhd in temp_bkhd_data:
                if len(row_original_bkhd) < 13: # Đảm bảo có đủ các cột cần thiết (ít nhất đến M/index 12)
                    continue

                new_row_intermediate = [
                    row_original_bkhd[1],   # Original Col B (dùng cho logic số hóa đơn, ký hiệu)
                    row_original_bkhd[2],   # Original Col C (dùng cho logic số hóa đơn)
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
            b2_bkhd = clean_string(all_rows_from_bkhd[1][1]) # all_rows_from_bkhd[1] là dòng thứ 2 (index 1) của file gốc

            # So sánh f5_value_for_comparison (đã lấy 6 ký tự cuối từ F5 Data.xlsx) với b2_bkhd
            # Logic này giờ khớp với UpSSE.2025.py
            if f5_value_for_comparison != b2_bkhd:
                st.error("Bảng kê hóa đơn không phải của cửa hàng bạn chọn.")
                st.stop()

            final_rows = [[''] * len(headers) for _ in range(4)] + [headers] # Dòng tiêu đề và 4 dòng trống đầu
            all_tmt_rows = []
            no_invoice_rows = {p: [] for p in ["Xăng E5 RON 92-II", "Xăng RON 95-III", "Dầu DO 0,05S-II", "Dầu DO 0,001S-V"]}

            for row_intermediate in intermediate_data:
                upsse_row = [''] * len(headers)
                
                # Ánh xạ từ intermediate_data sang upsse_row (cấu trúc output cuối cùng)
                
                # Cột A (Mã khách): Từ original E (intermediate_data[3]), hoặc G5
                upsse_row[0] = clean_string(row_intermediate[3]) if row_intermediate[-1] == 'Yes' and row_intermediate[3] and clean_string(row_intermediate[3]) else g5_value
                
                # Cột B (Tên khách hàng): Từ original F (intermediate_data[4])
                upsse_row[1] = clean_string(row_intermediate[4])
                
                # Cột C (Ngày): Từ original D (intermediate_data[2])
                upsse_row[2] = row_intermediate[2]

                # Cột D (Số hóa đơn): Logic phức tạp dựa trên B5.value từ Data.xlsx, và các cột B, C của BKHD
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
                
                # Cột AK (Tiền thuế): Từ original M (intermediate_data[12]), tính toán
                # tien_thue_tren_bkhd_goc = to_float(row_intermediate[11]) # Dòng này bị sai index ở bản cũ
                tien_thue_tren_bkhd_goc = to_float(row_intermediate[12]) # Lấy giá trị Tiền thuế từ cột M của BKHD gốc
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

