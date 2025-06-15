import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle, Alignment
from datetime import datetime
import io
import os
import re

# --- Cấu hình trang Streamlit ---
st.set_page_config(layout="centered", page_title="Đồng bộ dữ liệu SSE")

# Đường dẫn đến các file cần thiết (giả định cùng thư mục với script)
LOGO_PATH = "Logo.png"
DATA_FILE_PATH = "Data.xlsx" # Tên chính xác của file dữ liệu

# Định nghĩa tiêu đề cho file UpSSE.xlsx
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
            value = value.replace(",", "").strip()
        return float(value)
    except (ValueError, TypeError):
        return 0.0

# --- Hàm làm sạch chuỗi (loại bỏ mọi loại khoảng trắng và chuẩn hóa) ---
def clean_string(s):
    if s is None:
        return ""
    # Thay thế nhiều khoảng trắng bằng một khoảng trắng duy nhất, sau đó loại bỏ khoảng trắng đầu/cuối
    s = re.sub(r'\s+', ' ', str(s)).strip()
    return s

# --- Hàm đọc dữ liệu tĩnh và bảng tra cứu từ Data.xlsx ---
@st.cache_data
def get_static_data_from_excel(file_path):
    """
    Đọc dữ liệu và xây dựng các bảng tra cứu từ Data.xlsx.
    Sử dụng openpyxl để đọc dữ liệu. Kết quả được cache.
    Cung cấp các lookup map cho CHXD details, X-lookup, TMT lookup, v.v.
    """
    try:
        wb = load_workbook(file_path, data_only=True)
        ws = wb.active

        listbox_data = []
        chxd_detail_map = {} # Map cho các giá trị G5, F5, H5, B5 dựa trên tên CHXD
        store_specific_x_lookup = {} # Map cho 'Vụ việc' dựa trên tên CHXD và loại sản phẩm
        
        # Đọc dữ liệu từ bảng chính trong Data.xlsx để tạo bản đồ tìm kiếm cho CHXD
        # Giả định cột K (index 10) chứa tên CHXD, các cột P (15), Q (16), R (17) chứa
        # các giá trị tương ứng cho G5, F5_full, H5_val
        for row_idx in range(4, ws.max_row + 1):
            row_data_values = [cell.value for cell in ws[row_idx]]

            # Đảm bảo hàng có đủ cột để tham chiếu
            if len(row_data_values) < 18: 
                continue

            raw_chxd_name = row_data_values[10] # Cột K (index 10) - Tên CHXD
            if raw_chxd_name is not None and clean_string(raw_chxd_name):
                # Chuẩn hóa tên CHXD cho khóa bản đồ và danh sách dropdown
                chxd_name_key = clean_string(raw_chxd_name).lower() # Chuyển sang chữ thường để nhất quán
                chxd_name_display = clean_string(raw_chxd_name) # Giữ nguyên để hiển thị trong dropdown

                if chxd_name_display not in listbox_data:
                    listbox_data.append(chxd_name_display)

                g5_val = row_data_values[15] if pd.notna(row_data_values[15]) else None # Cột P (index 15)
                f5_val_full = clean_string(row_data_values[16]) if pd.notna(row_data_values[16]) else '' # Cột Q (index 16)
                h5_val = clean_string(row_data_values[17]).lower() if pd.notna(row_data_values[17]) else '' # Cột R (index 17)
                
                # Chỉ thêm vào bản đồ nếu có đủ thông tin chi tiết
                if g5_val is not None and f5_val_full and h5_val:
                    chxd_detail_map[chxd_name_key] = { # Sử dụng chxd_name_key (chữ thường) làm khóa
                        'g5_val': g5_val,
                        'h5_val': h5_val,
                        'f5_val_full': f5_val_full,
                        'b5_val': chxd_name_display # b5_val vẫn là tên hiển thị
                    }
                
                # Bản đồ tìm kiếm cho 'Vụ việc' (cột X) dựa trên tên CHXD và loại sản phẩm
                # Lấy giá trị từ các cột L (11), M (12), N (13), O (14)
                store_specific_x_lookup[chxd_name_key] = { # Sử dụng chxd_name_key (chữ thường) làm khóa
                    "xăng e5 ron 92-ii": row_data_values[11] if len(row_data_values) > 11 and pd.notna(row_data_values[11]) else '',
                    "xăng ron 95-iii":   row_data_values[12] if len(row_data_values) > 12 and pd.notna(row_data_values[12]) else '',
                    "dầu do 0,05s-ii":   row_data_values[13] if len(row_data_values) > 13 and pd.notna(row_data_values[13]) else '',
                    "dầu do 0,001s-v":   row_data_values[14] if len(row_data_values) > 14 and pd.notna(row_data_values[14]) else ''
                }
        
        # Các bảng tìm kiếm tĩnh khác (không phụ thuộc vào A5, đọc trực tiếp)
        # Các dải ô này đã được xác định từ file UpSSE.2025.py
        lookup_table = {} # I4:J7 - Mã hàng / Tên mặt hàng
        for row in ws.iter_rows(min_row=4, max_row=7, min_col=9, max_col=10, values_only=True): 
            if row[0] and row[1]: lookup_table[clean_string(row[0]).lower()] = row[1]
        
        tmt_lookup_table = {} # I10:J13 - TMT lookup (Đây là mức phí BVMT trên mỗi lít)
        for row in ws.iter_rows(min_row=10, max_row=13, min_col=9, max_col=10, values_only=True): 
            if row[0] and row[1]: tmt_lookup_table[clean_string(row[0]).lower()] = to_float(row[1])
        
        s_lookup_table = {} # I29:J31 - Tk nợ lookup
        for row in ws.iter_rows(min_row=29, max_row=31, min_col=9, max_col=10, values_only=True): 
            if row[0] and row[1]: s_lookup_table[clean_string(row[0]).lower()] = row[1]
        
        t_lookup_regular = {} # I33:J35 - Tk doanh thu lookup (regular)
        for row in ws.iter_rows(min_row=33, max_row=35, min_col=9, max_col=10, values_only=True): 
            if row[0] and row[1]: t_lookup_regular[clean_string(row[0]).lower()] = row[1]
        
        t_lookup_tmt = {} # I48:J50 - Tk doanh thu lookup (TMT)
        for row in ws.iter_rows(min_row=48, max_row=50, min_col=9, max_col=10, values_only=True): 
            if row[0] and row[1]: t_lookup_tmt[clean_string(row[0]).lower()] = row[1]

        v_lookup_table = {} # I53:J55 - Tk thuế có lookup
        for row in ws.iter_rows(min_row=53, max_row=55, min_col=9, max_col=10, values_only=True): 
            if row[0] and row[1]: v_lookup_table[clean_string(row[0]).lower()] = row[1]
        
        u_value = ws['J36'].value # Giá trị ô J36
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

def add_summary_row_for_no_invoice(data_for_summary_product, raw_bkhd_all_rows, product_name, headers_list,
                    g5_val, b5_val, s_lookup, t_lookup, v_lookup, x_lookup_for_store, u_val, h5_val, common_lookup_table, tmt_lookup_table_for_bvmt):
    new_row = [''] * len(headers_list)
    new_row[0], new_row[1] = g5_val, f"Khách hàng mua {product_name} không lấy hóa đơn"
    new_row[2] = data_for_summary_product[0][2] if data_for_summary_product else "" # Ngày
    new_row[4] = data_for_summary_product[0][4] if data_for_summary_product else "" # Ký hiệu
    value_C, value_E = clean_string(new_row[2]), clean_string(new_row[4])
    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name, "")
    if b5_val == "Nguyễn Huệ": new_row[3] = f"HNBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh": new_row[3] = f"MMBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    else: new_row[3] = f"{value_E[-2:]}BK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    new_row[5] = f"Xuất bán lẻ theo hóa đơn số {new_row[3]}"
    new_row[6], new_row[7], new_row[8], new_row[9] = common_lookup_table.get(clean_string(product_name).lower(), ''), product_name, "Lít", g5_val
    new_row[10], new_row[11] = '', ''
    total_M = sum(to_float(r[12]) for r in data_for_summary_product) # Tổng số lượng cột M (index 12)
    new_row[12] = total_M
    new_row[13] = max((to_float(r[13]) for r in data_for_summary_product), default=0.0) # Giá bán (cột N - index 13)
    
    # Tính toán Tiền hàng (cột O)
    tien_hang_hd_from_bkhd_original = sum(to_float(row[11]) for row in raw_bkhd_all_rows[4:] # Bắt đầu từ dòng dữ liệu
                                          if len(row) > 8 and clean_string(row[5]) == "Người mua không lấy hóa đơn" and clean_string(row[8]) == product_name)
    
    # Mức giá tham khảo, có thể thay đổi tùy loại mặt hàng.
    # Sử dụng các giá trị hardcode như trong UpSSE.2025.py cho tính toán Tiền hàng của dòng tổng hợp
    price_per_liter_for_tienhang_summary = {
        "Xăng E5 RON 92-II": 1900,
        "Xăng RON 95-III": 2000,
        "Dầu DO 0,05S-II": 1000,
        "Dầu DO 0,001S-V": 1000
    }.get(product_name, 0)
    new_row[14] = tien_hang_hd_from_bkhd_original - round(total_M * price_per_liter_for_tienhang_summary, 0)

    new_row[15], new_row[16], new_row[17] = '', '', 10
    new_row[18], new_row[19] = s_lookup.get(h5_val, ''), t_lookup.get(h5_val, '')
    new_row[20], new_row[21] = u_val, v_lookup.get(h5_val, '')
    new_row[22] = ''
    new_row[23] = x_lookup_for_store.get(clean_string(product_name).lower(), '')
    for i in range(24, 31): new_row[i] = ''
    new_row[31] = f"Khách mua {product_name} không lấy hóa đơn"
    new_row[32], new_row[33], new_row[34], new_row[35] = "", "", '', ''
    
    # === CHỈNH SỬA THUẬT TOÁN TÍNH TIỀN THUẾ (cột AK) CHO DÒNG TỔNG HỢP ===
    # Lấy tổng Tiền thuế từ các dòng hóa đơn gốc (row[12] là cột M gốc trong BKHD)
    # Đây chính là TienthueHD trong bản gốc
    total_original_tax_from_bkhd_original = sum(to_float(row[12]) for row in raw_bkhd_all_rows[4:] 
                                                 if len(row) > 8 and clean_string(row[5]) == "Người mua không lấy hóa đơn" and clean_string(row[8]) == product_name)
    
    # Mức giá sử dụng để trừ đi trong công thức tính Tiền thuế dòng tổng hợp
    # Các giá trị này được hardcode trong UpSSE.2025.py (1900, 2000, 1000)
    # Chúng KHÔNG PHẢI là mức phí BVMT từ tmt_lookup_table.
    price_for_tax_deduction_summary_row = {
        "Xăng E5 RON 92-II": 1900,
        "Xăng RON 95-III": 2000,
        "Dầu DO 0,05S-II": 1000,
        "Dầu DO 0,001S-V": 1000
    }.get(product_name, 0) # Lấy giá trị tương ứng với sản phẩm
    
    # Tính phần thuế cần trừ đi (làm tròn đến hàng đơn vị)
    # Mô phỏng chính xác logic từ UpSSE.2025.py: Tienthue1 = TienthueHD - round(total_M*1900*0.1)
    tax_to_deduct_summary_row = round(total_M * price_for_tax_deduction_summary_row * 0.1) 
    
    new_row[36] = total_original_tax_from_bkhd_original - tax_to_deduct_summary_row
    return new_row

def create_per_invoice_tmt_row(original_row_data, tmt_value, g5_val, s_lookup, t_lookup_tmt, v_lookup, u_val, h5_val):
    tmt_row = list(original_row_data)
    tmt_row[6], tmt_row[7], tmt_row[8] = "TMT", "Thuế bảo vệ môi trường", "VNĐ"
    tmt_row[9] = g5_val
    tmt_row[13] = tmt_value
    tmt_row[14] = round(tmt_value * to_float(original_row_data[12]), 0)
    tmt_row[17] = 10
    tmt_row[18] = s_lookup.get(h5_val, '')
    tmt_row[19] = t_lookup_tmt.get(h5_val, '')
    tmt_row[20], tmt_row[21] = u_val, v_lookup.get(h5_val, '')
    tmt_row[31] = ""
    tmt_row[36] = round(tmt_value * to_float(original_row_data[12]) * 0.1, 0)
    for idx in [5, 10, 11, 15, 16, 22, 24, 25, 26, 27, 28, 29, 30, 32, 33, 34, 35]:
        if idx < len(tmt_row): tmt_row[idx] = ''
    return tmt_row

def add_tmt_summary_row(product_name_full, total_bvmt_amount, g5_val, s_lookup, t_lookup_tmt, v_lookup, u_val, h5_val, 
                        representative_date, representative_symbol, total_quantity_for_tmt, tmt_unit_value_for_summary, b5_val, customer_name_for_summary_row, x_lookup_for_store):
    new_tmt_row = [''] * len(headers) # Sử dụng headers toàn cục
    new_tmt_row[0], new_tmt_row[1], new_tmt_row[2] = g5_val, customer_name_for_summary_row, representative_date
    value_C, value_E = clean_string(representative_date), clean_string(representative_symbol)
    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name_full, "")
    if b5_val == "Nguyễn Huệ": new_tmt_row[3] = f"HNBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh": new_tmt_row[3] = f"MMBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    else: new_tmt_row[3] = f"{value_E[-2:]}BK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    new_tmt_row[4] = representative_symbol
    new_tmt_row[6], new_tmt_row[7], new_tmt_row[8] = "TMT", "Thuế bảo vệ môi trường", "VNĐ"
    new_tmt_row[9], new_tmt_row[12] = g5_val, total_quantity_for_tmt
    new_tmt_row[13] = tmt_unit_value_for_summary
    new_tmt_row[14] = round(to_float(total_quantity_for_tmt) * to_float(tmt_unit_value_for_summary), 0)
    new_tmt_row[17] = 10
    new_tmt_row[18] = s_lookup.get(h5_val, '')
    new_tmt_row[19] = t_lookup_tmt.get(h5_val, '')
    new_tmt_row[20], new_tmt_row[21] = u_val, v_lookup.get(h5_val, '')
    new_tmt_row[23] = x_lookup_for_store.get(product_name_full.lower(), '')
    new_tmt_row[36], new_tmt_row[31] = total_bvmt_amount, ""
    for idx in [5,10,11,15,16,22,24,25,26,27,28,29,30,32,33,34,35]:
        if idx != 23 and idx < len(new_tmt_row): new_tmt_row[idx] = ''
    return new_tmt_row

# --- Tải dữ liệu tĩnh ---
static_data = get_static_data_from_excel(DATA_FILE_PATH)
listbox_data = static_data["listbox_data"]
lookup_table = static_data["lookup_table"]
tmt_lookup_table = static_data["tmt_lookup_table"] # Đây là bảng tra cứu mức phí BVMT
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
            # Chuẩn hóa giá trị đã chọn để tìm kiếm trong bản đồ
            selected_value_normalized = clean_string(selected_value).lower()
            
            if selected_value_normalized not in chxd_detail_map:
                st.error(f"Không tìm thấy thông tin chi tiết cho CHXD: '{selected_value}'. Vui lòng kiểm tra lại file Data.xlsx. Đảm bảo tên CHXD trong cột K của Data.xlsx khớp chính xác với lựa chọn của bạn (không phân biệt chữ hoa/thường, không có khoảng trắng thừa).")
                st.stop()
            
            # Lấy thông tin chi tiết từ bản đồ
            chxd_details = chxd_detail_map[selected_value_normalized]
            g5_value, h5_value, f5_value_full, b5_value = chxd_details['g5_val'], chxd_details['h5_val'], chxd_details['f5_val_full'], chxd_details['b5_val']
            
            # Lấy lookup map 'Vụ việc' cụ thể cho CHXD đã chọn
            x_lookup_for_store = store_specific_x_lookup.get(selected_value_normalized, {})
            if not x_lookup_for_store:
                st.warning(f"Không tìm thấy mã Vụ việc cụ thể cho cửa hàng '{selected_value}'. Có thể dẫn đến thiếu dữ liệu ở cột 'Vụ việc'.")

            bkhd_wb = load_workbook(uploaded_file)
            bkhd_ws = bkhd_wb.active

            long_cells = [f"H{r_idx+1}" for r_idx, cell in enumerate(bkhd_ws['H']) if cell.value and len(str(cell.value)) > 128]
            if long_cells:
                st.error("Địa chỉ trên ô " + ', '.join(long_cells) + " quá dài, hãy điều chỉnh và thử lại.")
                st.stop()

            # Lấy tất cả các dòng từ file BKHD gốc trước khi xử lý
            raw_bkhd_all_rows = list(bkhd_ws.iter_rows(values_only=True))

            temp_bkhd_data = raw_bkhd_all_rows[4:] if len(raw_bkhd_all_rows) >= 4 else []
            
            intermediate_data = []
            for row in temp_bkhd_data:
                # Cần đảm bảo row có đủ độ dài trước khi truy cập chỉ mục
                if len(row) < 17: # Index 16 là cột Q. Đảm bảo có đủ dữ liệu cột gốc
                    continue
                # Trích xuất dữ liệu theo các cột tương ứng trong file gốc và sắp xếp lại
                # Đây là quá trình "làm sạch" và "sắp xếp lại" cột tương tự UpSSE.2025.py
                # new_row (intermediate_data) sẽ có các chỉ mục sau:
                # 0: A gốc (Mã khách từ BKHD)
                # 1: B gốc (Ký hiệu HĐ)
                # 2: C gốc (Số HĐ)
                # 3: D gốc (Ngày)
                # 4: E gốc (Mã KH)
                # 5: F gốc (Tên khách hàng)
                # 6: H gốc (Địa chỉ)
                # 7: G gốc (Mã số thuế)
                # 8: I gốc (Tên mặt hàng)
                # 9: J gốc (Số lượng)
                # 10: K gốc (Giá bán / Tổng tiền hàng khi nhân số lượng)
                # 11: L gốc (Tiền hàng thuần)
                # 12: M gốc (Tiền thuế)
                # 13: Q gốc (Một cột khác, được giữ lại trong vi_tri_cu_idx của bản gốc)
                # 14: Cột mới "Công nợ" (Yes/No)

                new_row = [
                    row[0], # A
                    row[1], # B
                    row[2], # C
                    row[3], # D
                    row[4], # E
                    row[5], # F
                    row[7], # H (gốc)
                    row[6], # G (gốc)
                    row[8], # I (gốc)
                    row[9], # J (gốc)
                    row[10], # K (gốc)
                    row[11], # L (gốc)
                    row[12], # M (gốc)
                    row[16]  # Q (gốc)
                ]

                # Xử lý ngày (cột D - new_row[3])
                if new_row[3]:
                    try: 
                        date_obj = datetime.strptime(str(new_row[3])[:10], '%d-%m-%Y')
                        new_row[3] = date_obj.strftime('%Y-%m-%d')
                    except ValueError: 
                        pass # Giữ nguyên nếu không phải định dạng ngày

                # Thêm cột "Công nợ" (dựa trên cột E gốc - new_row[4])
                ma_kh = new_row[4] # Mã khách (original E column)
                new_row.append("No" if ma_kh is None or len(clean_string(ma_kh)) > 9 else "Yes")
                intermediate_data.append(new_row)

            if not intermediate_data:
                st.error("Không có dữ liệu hợp lệ trong file bảng kê sau khi xử lý.")
                st.stop()

            # Kiểm tra B2 của BKHD gốc với F5 của Data
            # raw_bkhd_all_rows[1][1] là giá trị ô B2 trong BKHD gốc (dòng 2, cột B)
            # F5_value_full có thể bắt đầu bằng '1' trong Data.xlsx, cần loại bỏ để so sánh
            b2_bkhd_original = clean_string(raw_bkhd_all_rows[1][1])
            f5_norm_compare = f5_value_full
            if f5_norm_compare.startswith('1'): 
                f5_norm_compare = f5_norm_compare[1:]

            if f5_norm_compare != b2_bkhd_original:
                st.error(f"Bảng kê hóa đơn không phải của cửa hàng bạn chọn. Giá trị F5 của Data.xlsx ({f5_value_full}) không khớp với B2 của Bảng kê hóa đơn ({b2_bkhd_original}).")
                st.stop()

            final_rows, all_tmt_rows = [[''] * len(headers) for _ in range(4)] + [headers], []
            no_invoice_rows = {p: [] for p in ["Xăng E5 RON 92-II", "Xăng RON 95-III", "Dầu DO 0,05S-II", "Dầu DO 0,001S-V"]}

            for row in intermediate_data:
                upsse_row = [''] * len(headers)
                # Sử dụng các chỉ mục của `row` (là `intermediate_data` đã được sắp xếp/chuyển đổi)
                # row[9] là Số lượng (J gốc BKHD)
                # row[10] là Giá bán (K gốc BKHD - giá nhân số lượng)
                # row[11] là Tiền hàng (L gốc BKHD - tiền hàng thuần)
                # row[12] là Tiền thuế (M gốc BKHD)

                upsse_row[0] = clean_string(row[4]) if row[-1] == 'Yes' and row[4] and clean_string(row[4]) else g5_value
                upsse_row[1], upsse_row[2] = clean_string(row[5]), row[3] # Tên KH, Ngày
                b_orig_bkhd = clean_string(row[1]) # Ký hiệu HĐ gốc (Cột B BKHD)
                c_orig_bkhd = clean_string(row[2]) # Số HĐ gốc (Cột C BKHD)
                
                if b5_value == "Nguyễn Huệ": upsse_row[3] = f"HN{c_orig_bkhd[-6:]}"
                elif b5_value == "Mai Linh": upsse_row[3] = f"MM{c_orig_bkhd[-6:]}"
                else: upsse_row[3] = f"{b_orig_bkhd[-2:]}{c_orig_bkhd[-6:]}"
                
                upsse_row[4] = f"1{b_orig_bkhd}" if b_orig_bkhd else '' # Ký hiệu
                upsse_row[5] = f"Xuất bán lẻ theo hóa đơn số {upsse_row[3]}" # Diễn giải
                
                product_name = clean_string(row[8]) # Tên mặt hàng (I gốc BKHD)
                upsse_row[6], upsse_row[7] = lookup_table.get(product_name.lower(), ''), product_name # Mã hàng, Tên mặt hàng
                upsse_row[8], upsse_row[9] = "Lít", g5_value # Đvt, Mã kho
                upsse_row[10], upsse_row[11] = '', '' # Mã vị trí, Mã lô
                upsse_row[12] = to_float(row[9]) # Số lượng (J gốc BKHD)
                
                tmt_value = tmt_lookup_table.get(product_name.lower(), 0.0)
                
                # Giá bán (cột N - upsse_row[13]): Từ K gốc (intermediate_data[10]) chia 1.1 trừ TMT
                upsse_row[13] = round(to_float(row[10]) / 1.1 - tmt_value, 2)
                
                # Tiền hàng (cột O - upsse_row[14]): Từ L gốc (intermediate_data[11]) trừ TMT * Số lượng
                upsse_row[14] = to_float(row[11]) - round(tmt_value * upsse_row[12])
                
                upsse_row[15], upsse_row[16], upsse_row[17] = '', '', 10 # Mã nt, Tỷ giá, Mã thuế
                upsse_row[18] = s_lookup_table.get(h5_value, '') # Tk nợ
                upsse_row[19] = t_lookup_regular.get(h5_value, '') # Tk doanh thu
                upsse_row[20], upsse_row[21] = u_value, v_lookup_table.get(h5_value, '') # Tk giá vốn, Tk thuế có
                upsse_row[22] = '' # Cục thuế
                upsse_row[23] = x_lookup_for_store.get(product_name.lower(), '') # Vụ việc
                for i in range(24, 31): upsse_row[i] = '' # Bộ phận đến Nhân viên bán
                upsse_row[31] = upsse_row[1] # Tên KH (thuế)
                upsse_row[32], upsse_row[33] = row[6], row[7] # Địa chỉ (thuế), Mã số thuế (Gốc H, G)
                upsse_row[34], upsse_row[35] = '', '' # Nhóm hàng, Ghi chú
                
                # Tiền thuế (cột AK - upsse_row[36]): Từ M gốc (intermediate_data[12]) trừ TMT * Số lượng * 0.1
                upsse_row[36] = to_float(row[12]) - round(upsse_row[12] * tmt_value * 0.1)

                if upsse_row[1] == "Người mua không lấy hóa đơn" and product_name in no_invoice_rows:
                    no_invoice_rows[product_name].append(upsse_row)
                else:
                    final_rows.append(upsse_row)
                    if tmt_value > 0 and upsse_row[12] > 0:
                        all_tmt_rows.append(create_per_invoice_tmt_row(upsse_row, tmt_value, g5_value, s_lookup_table, t_lookup_tmt, v_lookup_table, u_value, h5_value))

            for product_name, rows in no_invoice_rows.items():
                if rows:
                    # Truyền raw_bkhd_all_rows và tmt_lookup_table để hàm summary có thể lấy dữ liệu gốc và mức phí BVMT
                    summary_row = add_summary_row_for_no_invoice(rows, raw_bkhd_all_rows, product_name, headers, g5_value, b5_value, s_lookup_table, t_lookup_regular, v_lookup_table, x_lookup_for_store, u_value, h5_value, lookup_table, tmt_lookup_table)
                    final_rows.append(summary_row)
                    
                    # Tính tổng BVMT cho dòng tóm tắt TMT của khách hàng không lấy hóa đơn
                    # Mức phí BVMT cho dòng TMT summary cũng phải lấy từ tmt_lookup_table
                    bvmt_fee_for_summary_tmt = tmt_lookup_table.get(clean_string(product_name).lower(), 0.0)
                    total_bvmt_tax = sum(round(to_float(r[12]) * bvmt_fee_for_summary_tmt * 0.1, 0) for r in rows)
                    
                    if total_bvmt_tax > 0:
                        tmt_unit = bvmt_fee_for_summary_tmt
                        total_qty = sum(to_float(r[12]) for r in rows)
                        
                        all_tmt_rows.append(add_tmt_summary_row(product_name, total_bvmt_tax, g5_value, s_lookup_table, t_lookup_tmt, v_lookup_table, u_value, h5_value, summary_row[2], summary_row[4], total_qty, tmt_unit, b5_value, summary_row[1], x_lookup_for_store))

            final_rows.extend(all_tmt_rows)

            up_sse_wb_final = Workbook()
            up_sse_ws_final = up_sse_wb_final.active
            for row_data in final_rows: up_sse_ws_final.append(row_data)

            text_style, date_style = NamedStyle(name="text_style", number_format='@'), NamedStyle(name="date_style", number_format='DD/MM/YYYY')
            exclude_cols = {3, 13, 14, 15, 18, 19, 20, 21, 22, 37} # Cột 37 là cột AK, loại khỏi việc chuyển đổi thành text để không bị mất giá trị số
            for r in range(1, up_sse_ws_final.max_row + 1):
                for c in range(1, up_sse_ws_final.max_column + 1):
                    cell = up_sse_ws_final.cell(row=r, column=c)
                    if not cell.value or clean_string(cell.value) == "None": continue
                    if c == 3: # Cột C (index 3) là cột ngày
                        try:
                            # Chuyển đổi giá trị sang dạng ngày nếu là chuỗi YYYY-MM-DD
                            if isinstance(cell.value, str) and len(cell.value) == 10 and cell.value[4] == '-' and cell.value[7] == '-':
                                cell.value = datetime.strptime(clean_string(cell.value), '%Y-%m-%d').date()
                                cell.style = date_style
                            # Nếu giá trị đã là datetime object (ví dụ từ openpyxl), vẫn áp dụng style
                            elif isinstance(cell.value, datetime):
                                cell.value = cell.value.date()
                                cell.style = date_style
                        except (ValueError, TypeError): 
                            pass # Giữ nguyên nếu không phải định dạng ngày
                    elif c not in exclude_cols: 
                        cell.style = text_style
            
            # Đảm bảo các cột R, S, T, U, V (index 18-22) được định dạng là text
            for r in range(1, up_sse_ws_final.max_row + 1):
                for c in range(18, 23): # Cột R (index 18) đến V (index 22)
                    up_sse_ws_final.cell(row=r, column=c).number_format = '@'

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
