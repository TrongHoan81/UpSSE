import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle, Alignment
from datetime import datetime
import io
import os

# --- Thiết lập giao diện Streamlit ---
st.set_page_config(layout="centered", page_title="Đồng bộ dữ liệu SSE")

# Đường dẫn đến file Logo.png và Data.xlsx
LOGO_PATH = "Logo.png"
DATA_FILE_PATH = "Data.xlsx"

# Kiểm tra ngày hết hạn (giữ nguyên logic từ code gốc)
expiration_date = datetime(2025, 6, 26)
current_date = datetime.now()

if current_date > expiration_date:
    st.error("Có lỗi khi chạy chương trình, vui lòng liên hệ tác giả để được hỗ trợ!")
    st.info("Nguyễn Trọng Hoàn - 0902069469")
    st.stop() # Dừng ứng dụng Streamlit

# Helper function to safely convert to float
def to_float(value):
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0

# --- Hàm hỗ trợ đọc dữ liệu từ Data.xlsx ---
@st.cache_data
def get_static_data_from_excel(file_path):
    """
    Hàm đọc dữ liệu tĩnh và bảng tra cứu từ Data.xlsx bằng openpyxl.
    Sử dụng @st.cache_data để cache kết quả, tránh đọc lại mỗi lần chạy lại script.
    """
    try:
        wb = load_workbook(file_path, data_only=True)
        ws = wb.active

        data_listbox = []
        chxd_detail_map = {} # Map để lưu thông tin chi tiết CHXD

        for row_idx in range(4, ws.max_row + 1): # Bắt đầu từ hàng 4 Excel (index 3 Python)
            # Lấy tất cả các giá trị của hàng
            row_data_values = [cell.value for cell in ws[row_idx]]

            # Đảm bảo hàng có đủ cột để truy cập các chỉ số cần thiết (đến S - index 17)
            if len(row_data_values) >= 18: # Cần ít nhất 18 cột để truy cập index 17
                raw_chxd_name = row_data_values[10] # Column K (index 10)
                
                if raw_chxd_name is not None and str(raw_chxd_name).strip() != '':
                    chxd_name_str = str(raw_chxd_name).strip()
                    
                    if chxd_name_str: # Ensure it's not an empty string after stripping
                        if chxd_name_str not in data_listbox: # Tránh thêm trùng lặp
                            data_listbox.append(chxd_name_str)

                        g5_val = row_data_values[15] if len(row_data_values) > 15 and pd.notna(row_data_values[15]) else None # Column P (index 15)
                        f5_val_full = str(row_data_values[16]).strip() if len(row_data_values) > 16 and pd.notna(row_data_values[16]) else '' # Column Q (index 16)
                        h5_val = str(row_data_values[17]).strip().lower() if len(row_data_values) > 17 and pd.notna(row_data_values[17]) else '' # Column S (index 17)
                        b5_val = chxd_name_str # B5 is the CHXD name

                        # Only add to map if we have minimal required info (CHXD name and f5_val_full for comparison)
                        if f5_val_full: # Assuming f5_val_full (Mã kho) is essential for the check
                            chxd_detail_map[chxd_name_str] = {
                                'g5_val': g5_val,
                                'h5_val': h5_val,
                                'f5_val_full': f5_val_full,
                                'b5_val': b5_val
                            }
        # Re-read for specific static values that might not be in the main CHXD details table
        # For instance, u_value from J36
        u_value = ws['J36'].value # Read from J36 (openpyxl's default indexing)

        # Build other lookup tables using openpyxl (as their ranges are specific)
        # I4:J7 (Lookup table)
        lookup_table = {}
        for row in ws.iter_rows(min_row=4, max_row=7, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                lookup_table[str(row[0]).strip().lower()] = row[1]
        # I10:J13 (TMT Lookup table)
        tmt_lookup_table = {}
        for row in ws.iter_rows(min_row=10, max_row=13, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                tmt_lookup_table[str(row[0]).strip().lower()] = to_float(row[1]) # Convert to float here
        # I29:J31 (S Lookup table)
        s_lookup_table = {}
        for row in ws.iter_rows(min_row=29, max_row=31, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                s_lookup_table[str(row[0]).strip().lower()] = row[1]
        # I33:J35 (T Lookup table)
        t_lookup_table = {}
        for row in ws.iter_rows(min_row=33, min_col=9, max_col=10, max_row=35, values_only=True):
            if row[0] and row[1]:
                t_lookup_table[str(row[0]).strip().lower()] = row[1]
        # I53:J55 (V Lookup table)
        v_lookup_table = {}
        for row in ws.iter_rows(min_row=53, max_row=55, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                v_lookup_table[str(row[0]).strip().lower()] = row[1]
        # I17:J20 (X Lookup table)
        x_lookup_table = {}
        for row in ws.iter_rows(min_row=17, max_row=20, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                x_lookup_table[str(row[0]).strip().lower()] = row[1]

        wb.close()
        
        return {
            "listbox_data": data_listbox,
            "lookup_table": lookup_table,
            "tmt_lookup_table": tmt_lookup_table,
            "s_lookup_table": s_lookup_table,
            "t_lookup_table": t_lookup_table,
            "v_lookup_table": v_lookup_table,
            "x_lookup_table": x_lookup_table,
            "u_value": u_value,
            "chxd_detail_map": chxd_detail_map # Trả về bản đồ chi tiết CHXD
        }
    except FileNotFoundError:
        st.error(f"Lỗi: Không tìm thấy file {file_path}. Vui lòng đảm bảo file tồn tại.")
        st.stop()
    except Exception as e:
        st.error(f"Lỗi không mong muốn khi đọc file Data.xlsx: {e}")
        st.exception(e) # Hiển thị stack trace để debug
        st.stop()

# --- Hàm tạo dòng khách vãng lai ---
def add_summary_row(ws_target, ws_source, product_name, sum_m_col_count, price_per_liter, suffix_d, headers_list,
                    g5_val, b5_val, s_lookup, t_lookup, v_lookup, x_lookup, u_val, h5_val, lookup_table, final_rows_for_upsse_ref):
    new_row = [''] * len(headers_list)
    new_row[0] = g5_val
    new_row[1] = f"Khách hàng mua {product_name} không lấy hóa đơn"

    c6_val = None
    e6_val = None
    if len(final_rows_for_upsse_ref) > 5 and len(final_rows_for_upsse_ref[5]) > 4: 
        c6_val = final_rows_for_upsse_ref[5][2]
        e6_val = final_rows_for_upsse_ref[5][4]
    
    c6_val = c6_val if c6_val is not None else ""
    e6_val = e6_val if e6_val is not None else ""


    new_row[2] = c6_val # Cột C
    new_row[4] = e6_val # Cột E

    value_C = new_row[2] if new_row[2] else ""
    value_E = new_row[4] if new_row[4] else ""

    if b5_val == "Nguyễn Huệ":
        value_D = f"HNBK{str(value_C)[-2:]}.{str(value_C)[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh":
        value_D = f"MMBK{str(value_C)[-2:]}.{str(value_C)[5:7]}.{suffix_d}"
    else:
        value_D = f"{str(value_E)[-2:]}BK{str(value_C)[-2:]}.{str(value_C)[5:7]}.{suffix_d}"
    new_row[3] = value_D
    new_row[5] = f"Xuất bán lẻ theo hóa đơn số " + new_row[3]
    new_row[7] = product_name
    new_row[6] = lookup_table.get(product_name.strip().lower(), '')
    new_row[8] = "Lít"
    new_row[9] = g5_val
    new_row[10] = ''
    new_row[11] = ''

    total_M = 0
    max_value_N = None
    for r_idx in range(6, ws_target.max_row + 1): 
        row_data = [cell.value for cell in ws_target[r_idx]]
        if len(row_data) > 12 and row_data[1] == "Người mua không lấy hóa đơn" and row_data[7] == product_name:
            total_M += to_float(row_data[12]) # Convert to float
            if max_value_N is None or (to_float(row_data[13]) is not None and to_float(row_data[13]) > max_value_N): # Convert to float
                max_value_N = to_float(row_data[13]) # Convert to float
    
    new_row[12] = total_M
    new_row[13] = max_value_N

    tien_hang_hd = 0
    for r in ws_source.iter_rows(min_row=2, max_row=ws_source.max_row, values_only=True):
        if r[5] == "Người mua không lấy hóa đơn" and r[8] == product_name:
            tien_hang_hd += to_float(r[11]) # Convert to float
    new_row[14] = tien_hang_hd - round(total_M * price_per_liter, 0)

    new_row[17] = 10
    new_row[18] = s_lookup.get(h5_val, '')
    new_row[19] = t_lookup.get(h5_val, '')
    new_row[20] = u_val
    new_row[21] = v_lookup.get(h5_val, '')
    new_row[23] = x_lookup.get(product_name.strip().lower(), '')
    new_row[31] = f"Khách mua {product_name} không lấy hóa đơn"

    tien_thue_hd = 0
    for r in ws_source.iter_rows(min_row=2, max_row=ws_source.max_row, values_only=True):
        if r[5] == "Người mua không lấy hóa đơn" and r[8] == product_name:
            tien_thue_hd += to_float(r[12]) # Convert to float
    new_row[36] = tien_thue_hd - round(total_M * price_per_liter * 0.1, 0) # Convert to float

    return new_row

# --- Hàm tạo dòng tóm tắt thuế bảo vệ môi trường (TMT) ---
def add_tmt_summary_row(product_name_full, total_bvmt_amount, headers_list, g5_val, s_lookup, t_lookup, v_lookup, u_val, h5_val):
    new_tmt_row = [''] * len(headers_list)
    new_tmt_row[0] = g5_val # Mã khách (or generic code)
    new_tmt_row[1] = f"Thuế bảo vệ môi trường {product_name_full}" # Tên khách hàng
    
    # Các cột khác có thể để trống hoặc điền giá trị mặc định nếu cần
    new_tmt_row[6] = "TMT" # Mã hàng
    new_tmt_row[7] = "Thuế bảo vệ môi trường" # Tên mặt hàng
    new_tmt_row[8] = "VNĐ" # Đvt (assuming monetary unit)
    new_tmt_row[9] = g5_val # Mã kho (can be same as G5)
    
    # Tài khoản
    new_tmt_row[18] = s_lookup.get(h5_val, '') # Tk nợ (same as for sales)
    new_tmt_row[19] = t_lookup.get(h5_val, '') # Tk doanh thu (probably not relevant for TMT, maybe leave empty or specific)
    new_tmt_row[20] = u_val # Tk giá vốn (likely not relevant)
    new_tmt_row[21] = v_lookup.get(h5_val, '') # Tk thuế có (same as for sales)

    new_tmt_row[36] = total_bvmt_amount # Tiền thuế (total environmental tax)
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
chxd_detail_map = static_data["chxd_detail_map"] # Lấy bản đồ chi tiết CHXD

# --- Header ứng dụng ---
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

# --- Lựa chọn CHXD ---
selected_value = st.selectbox(
    "Chọn CHXD:",
    options=[""] + listbox_data, # Thêm lựa chọn trống
    key='selected_chxd'
)

uploaded_file = st.file_uploader("Tải lên file bảng kê hóa đơn (.xlsx)", type=["xlsx"])

# --- Xử lý file khi có đủ thông tin ---
if st.button("Xử lý", key='process_button'):
    if not selected_value:
        st.warning("Vui lòng chọn một giá trị từ danh sách CHXD.")
    elif uploaded_file is None:
        st.warning("Vui lòng tải lên file bảng kê hóa đơn.")
    else:
        try:
            # Lấy các giá trị G5, H5, F5_full, B5_value từ bản đồ tra cứu
            # Đảm bảo selected_value được chuẩn hóa (strip) trước khi tra cứu
            selected_value_normalized = selected_value.strip()

            if selected_value_normalized not in chxd_detail_map:
                st.error("Không tìm thấy thông tin chi tiết cho CHXD đã chọn trong Data.xlsx. Vui lòng kiểm tra lại tên CHXD.")
                st.stop()
            
            chxd_details = chxd_detail_map[selected_value_normalized]
            g5_value = chxd_details['g5_val']
            h5_value = chxd_details['h5_val']
            f5_value_full = chxd_details['f5_val_full'] # Lấy giá trị F5 đầy đủ
            b5_value = chxd_details['b5_val'] # Lấy giá trị B5 (tên CHXD)

            # Đọc file bảng kê hóa đơn từ dữ liệu đã tải lên
            temp_wb = load_workbook(uploaded_file)
            temp_ws = temp_wb.active

            # Kiểm tra các ô trong cột H có trên 128 ký tự
            long_cells = []
            for r_idx, cell in enumerate(temp_ws['H']):
                if cell.value is not None and len(str(cell.value)) > 128:
                    long_cells.append(f"H{r_idx+1}") # Lấy tọa độ ô
            if long_cells:
                st.error("Địa chỉ trên ô " + ', '.join(long_cells) + " quá dài, hãy điều chỉnh và thử lại.")
                st.stop()

            # Xóa 3 hàng đầu tiên
            temp_ws.delete_rows(1, 3)

            # Vị trí các cột cần giữ và sắp xếp lại
            vi_tri_cu = ['A', 'B', 'C', 'D', 'E', 'F', 'H', 'G', 'I', 'K', 'L', 'N', 'O', 'Q']
            # Chuyển đổi sang chỉ số 0-based để dễ dàng truy cập trong list
            vi_tri_cu_idx = [ord(c) - ord('A') for c in vi_tri_cu]

            # Tạo bảng dữ liệu mới với xử lý ngày (cột D)
            data = []
            for row_idx, row_values in enumerate(temp_ws.iter_rows(min_row=1, values_only=True)):
                new_row = []
                for idx_new_col, col_old_idx in enumerate(vi_tri_cu_idx):
                    cell_value = row_values[col_old_idx] if col_old_idx < len(row_values) else None

                    # Nếu là cột D (original index 3), chỉ lấy 10 ký tự đầu và chuyển đổi ngày tháng
                    if idx_new_col == 3 and cell_value: # idx_new_col 3 tương ứng với cột D mới
                        cell_value_str = str(cell_value)[:10] 
                        try:
                            date_obj = datetime.strptime(cell_value_str, '%d-%m-%Y')
                            cell_value = date_obj.strftime('%Y-%m-%d')
                        except ValueError:
                            pass # Giữ nguyên giá trị nếu không phải ngày hợp lệ
                    new_row.append(cell_value)
                data.append(new_row)

            # Tạo một Workbook mới để ghi dữ liệu đã xử lý
            bkhd_wb_processed = Workbook()
            bkhd_ws_processed = bkhd_wb_processed.active

            for row_data in data:
                bkhd_ws_processed.append(row_data)

            # Thêm cột "Công nợ" (O = cột 15)
            bkhd_ws_processed.cell(row=1, column=15).value = "Công nợ"
            for row in range(2, bkhd_ws_processed.max_row + 1):
                value = bkhd_ws_processed.cell(row=row, column=5).value  # Cột E = mã KH
                if value is None or len(str(value)) > 9:
                    bkhd_ws_processed.cell(row=row, column=15).value = "No"
                else:
                    bkhd_ws_processed.cell(row=row, column=15).value = "Yes"

            bkhd_ws = bkhd_ws_processed # Gán lại để sử dụng tên biến cũ

            # Lấy giá trị B2 từ bảng kê và loại bỏ khoảng trắng
            b2_bkhd_value = str(bkhd_ws['B2'].value).strip() if bkhd_ws['B2'].value else ''

            # Chuẩn hóa giá trị f5_value_full (Mã kho từ Data.xlsx)
            normalized_f5_value_full = f5_value_full
            if normalized_f5_value_full.startswith('1'):
                normalized_f5_value_full = normalized_f5_value_full[1:]


            if normalized_f5_value_full != b2_bkhd_value:
                st.error("Bảng kê hóa đơn không phải của cửa hàng bạn chọn.")
                st.stop()

            # Tiếp tục thực hiện các bước nếu trùng
            # Tạo file Excel mới - file UpSSE để chứa dữ liệu tạm thời
            temp_up_sse_wb = Workbook()
            temp_up_sse_ws = temp_up_sse_wb.active

            # Thêm các dòng trống trước tiêu đề
            for _ in range(4):
                temp_up_sse_ws.append([])

            # Điền tiêu đề vào dòng thứ 5 của file UpSSE
            headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng",
                       "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế",
                       "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm",
                       "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế",
                       "Nhóm Hàng", "Ghi chú", "Tiền thuế"]
            temp_up_sse_ws.append(headers)

            # Khởi tạo các biến phụ cho khách vãng lai và thuế bảo vệ môi trường
            kvlE5 = 0
            kvl95 = 0
            kvlDo = 0
            kvlD1 = 0
            total_bvmt_e5 = 0.0 # Initialize as float
            total_bvmt_95 = 0.0 # Initialize as float
            total_bvmt_do = 0.0 # Initialize as float
            total_bvmt_d1 = 0.0 # Initialize as float

            # Duyệt qua từng dòng của BKHD để tính toán và điền dữ liệu vào temp_up_sse_ws
            # Trong vòng lặp này, TẤT CẢ các dòng (bao gồm "Người mua không lấy hóa đơn") sẽ được thêm vào temp_up_sse_ws.
            for row_idx, row in enumerate(bkhd_ws.iter_rows(min_row=2, values_only=True)):
                new_row = [''] * len(headers)

                # Điều kiện cho cột A (Mã khách)
                if row[14] == 'No':  # Cột O (index 14) của BKHD
                    new_row[0] = g5_value 
                elif row[14] == 'Yes':
                    if row[4] is None or row[4] == '':  # Cột E (index 4) of BKHD
                        new_row[0] = g5_value 
                    else:
                        new_row[0] = str((row[4])) 

                # Cột B (Tên khách hàng): Điền bằng giá trị của cột F trên BKHD
                new_row[1] = row[5]  # Cột F (index 5) of BKHD

                # Cột C (Ngày): Điền bằng giá trị của cột D trên BKHD
                new_row[2] = row[3]  # Cột D (index 3) of BKHD

                # Cột D (Số hóa đơn): Điền là chuỗi ký tự bao gồm 2 ký tự cuối của cột B trên BKHD + 6 ký tự cuối của cột C trên BKHD
                if b5_value == "Nguyễn Huệ": 
                    new_row[3] = "HN" + str(row[2])[-6:] 
                elif b5_value == "Mai Linh": 
                    new_row[3] = "MM" + str(row[2])[-6:] 
                else:
                    new_row[3] = str(row[1])[-2:] + str(row[2])[-6:]  

                # Cột E (Ký hiệu): Điền bao gồm ký tự "1" và sau đó là giá trị cột B trên BKHD
                if row[1]:  
                    new_row[4] = "1" + str(row[1])
                else:
                    new_row[4] = ''

                # Cột F (Diễn giải): Điền một dãy ký tự bao gồm: “Xuất bán lẻ theo hóa đơn số” + giá trị Cột D (file UpSSE)
                new_row[5] = "Xuất bán lẻ theo hóa đơn số " + new_row[3]

                # Cột H (Tên mặt hàng): Điền bằng giá trị cột I trên BKHD
                new_row[7] = row[8]  # Cột I (index 8) of BKHD

                # Cột G (Mã hàng): Dò tìm giá trị của ô cùng dòng trên cột H trong ô I4:J7 của file Data.xlsx
                new_row[6] = lookup_table.get(str(new_row[7]).strip().lower(), '')

                # Cột I (Đvt): Điền dãy ký tự "Lít"
                new_row[8] = "Lít"

                # Cột J (Mã kho): Điền giá trị của ô G5 trên file Data.xlsx (sử dụng g5_value động)
                new_row[9] = g5_value 

                # Cột K (Mã vị trí) và L (Mã lô): Để trống
                new_row[10] = ''
                new_row[11] = ''

                # Cột M (Số lượng): Điền bằng giá trị của cột J trên BKHD
                new_row[12] = to_float(row[9])  # Convert to float

                # Tính toán TMT dựa trên giá trị cột H (Tên mặt hàng) của UpSSE
                tmt_value = to_float(tmt_lookup_table.get(str(new_row[7]).strip().lower(), 0)) # Ensure tmt_value is float

                # Cột N (Giá bán): Giá trị cột K trên BKHD chia cho 1.1 rồi trừ TMT, làm tròn tới 2 chữ số thập phân
                if row[10] is not None:  
                    new_row[13] = round(to_float(row[10]) / 1.1 - tmt_value, 2) # Convert to float
                else:
                    new_row[13] = 0.0

                # Cột O (Tiền hàng): Bằng giá trị cột L trên file BKHD trừ đi (TMT nhân với giá trị cột M trên file UpSSE)
                if row[11] is not None and new_row[12] is not None:  
                    tmt_calculation = round(tmt_value * new_row[12])
                    new_row[14] = to_float(row[11]) - tmt_calculation # Convert to float
                else:
                    new_row[14] = 0.0

                # Cột P (Mã nt) và Q (Tỷ giá): Để trống
                new_row[15] = ''
                new_row[16] = ''

                # Cột R (Mã thuế): Điền giá trị 10
                new_row[17] = 10

                # Cột S (Tk nợ): Dò tìm giá trị của ô H5 trong vùng I29:J31 của file Data.xlsx (sử dụng h5_value động)
                s_value_from_lookup = s_lookup_table.get(h5_value, '')
                new_row[18] = s_value_from_lookup

                # Cột T (Tk doanh thu): Dò tìm giá trị của ô H5 trong vùng I33:J35 của file Data.xlsx (sử dụng h5_value động)
                t_value_from_lookup = t_lookup_table.get(h5_value, '')
                new_row[19] = t_value_from_lookup

                # Cột U (Tk giá vốn): Điền giá trị tại ô J36 của file Data.xlsx (u_value cố định)
                new_row[20] = u_value

                # Cột V (Tk thuế có): Dò tìm giá trị của ô H5 trong vùng I53:J55 của file Data.xlsx (sử dụng h5_value động)
                v_value_from_lookup = v_lookup_table.get(h5_value, '')
                new_row[21] = v_value_from_lookup

                # Cột X (Vụ việc): Dò tìm giá trị của ô cùng dòng trên cột H trong vùng I17:J20 của file Data.xlsx
                h_value_for_x = str(new_row[7]).strip().lower()
                x_value_from_lookup = x_lookup_table.get(h_value_for_x, '')
                new_row[23] = x_value_for_x

                # Các cột Y, Z, AA, AB, AC, AD, AE: Để trống
                new_row[24] = '' 
                new_row[25] = '' 
                new_row[26] = '' 
                new_row[27] = '' 
                new_row[28] = '' 
                new_row[29] = '' 
                new_row[30] = '' 

                # Cột AF (Tên KH(thuế)): Điền bằng giá trị của cột B (Tên khách hàng) của file UpSSE.xlsx
                new_row[31] = new_row[1]

                # Cột AG (Địa chỉ (thuế)): Điền bằng giá trị của cột G trên file BKHD
                new_row[32] = row[6]  # Cột G (index 6) of BKHD

                # Cột AH (Mã số Thuế): Điền giá trị của cột H trên file BKHD
                new_row[33] = row[7]  # Cột H (index 7) of BKHD

                # Cột AI (Nhóm Hàng) và AJ (Ghi chú): Để trống
                new_row[34] = ''
                new_row[35] = ''

                # Cột AK (Tiền thuế): Tạo biến phụ Thue_Cua_TMT, làm tròn và tính toán
                thue_cua_tmt_for_row = 0.0 # Initialize as float
                if new_row[12] is not None and tmt_value is not None:  
                    thue_cua_tmt_for_row = round(new_row[12] * tmt_value * 0.1, 0) # Use new_row[12] (already float)
                    new_row[36] = new_row[14] - thue_cua_tmt_for_row  # Use new_row[14] (Tiền hàng) for calculation
                else:
                    new_row[36] = new_row[14] # If no TMT or M is none, keep Tiền hàng as is

                # Tích lũy tổng thuế BVMT cho từng loại sản phẩm
                if new_row[7] == "Xăng E5 RON 92-II":
                    total_bvmt_e5 += thue_cua_tmt_for_row
                elif new_row[7] == "Xăng RON 95-III":
                    total_bvmt_95 += thue_cua_tmt_for_row
                elif new_row[7] == "Dầu DO 0,05S-II":
                    total_bvmt_do += thue_cua_tmt_for_row
                elif new_row[7] == "Dầu DO 0,001S-V":
                    total_bvmt_d1 += thue_cua_tmt_for_row

                # Thêm dòng mới vào UpSSE tạm thời
                temp_up_sse_ws.append(new_row)

                # Đếm số lượng dòng thỏa mãn điều kiện cho kvl (dùng cho tổng kết Khách vãng lai)
                if new_row[1] == "Người mua không lấy hóa đơn" and new_row[7] == "Xăng E5 RON 92-II":
                    kvlE5 += 1
                elif new_row[1] == "Người mua không lấy hóa đơn" and new_row[7] == "Xăng RON 95-III":
                    kvl95 += 1
                elif new_row[1] == "Người mua không lấy hóa đơn" and new_row[7] == "Dầu DO 0,05S-II":
                    kvlDo += 1
                elif new_row[1] == "Người mua không lấy hóa đơn" and new_row[7] == "Dầu DO 0,001S-V":
                    kvlD1 += 1


            # --- Bắt đầu quá trình lọc và xây dựng lại worksheet ---
            final_rows_for_upsse = []

            # Giữ lại 5 hàng đầu tiên (tiêu đề và các dòng trống)
            for r_idx in range(1, 6): # Rows 1 to 5 (index 0 to 4 in 0-based list)
                row_values = [cell.value for cell in temp_up_sse_ws[r_idx]]
                final_rows_for_upsse.append(row_values)
            
            # Lặp qua các dòng dữ liệu thực tế từ temp_up_sse_ws (từ hàng 6 trở đi)
            for r_idx in range(6, temp_up_sse_ws.max_row + 1):
                row_data = [cell.value for cell in temp_up_sse_ws[r_idx]]
                
                # Kiểm tra nếu hàng có đủ cột và cột B có giá trị
                if len(row_data) > 1 and row_data[1] is not None:
                    col_b_value = str(row_data[1]).strip() # Loại bỏ khoảng trắng và chuyển về chuỗi
                    
                    # Nếu cột B là "Người mua không lấy hóa đơn", bỏ qua hàng này
                    if col_b_value == "Người mua không lấy hóa đơn":
                        continue 
                    
                    # Đối với các dòng dữ liệu khác, giữ lại
                    else:
                        final_rows_for_upsse.append(row_data)
                else:
                    # Nếu hàng không có đủ cột hoặc cột B rỗng, giữ lại để tránh mất dữ liệu không mong muốn
                    final_rows_for_upsse.append(row_data) 

            # Thêm các dòng tổng kết "Khách hàng mua..." vào cuối final_rows_for_upsse
            if kvlE5 > 0:
                final_rows_for_upsse.append(
                    add_summary_row(temp_up_sse_ws, bkhd_ws, "Xăng E5 RON 92-II", kvlE5, 1900, "1", headers,
                                g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table, final_rows_for_upsse))
            if kvl95 > 0:
                final_rows_for_upsse.append(
                    add_summary_row(temp_up_sse_ws, bkhd_ws, "Xăng RON 95-III", kvl95, 2000, "2", headers,
                                g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table, final_rows_for_upsse))
            if kvlDo > 0:
                final_rows_for_upsse.append(
                    add_summary_row(temp_up_sse_ws, bkhd_ws, "Dầu DO 0,05S-II", kvlDo, 1000, "3", headers,
                                g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table, final_rows_for_upsse))
            if kvlD1 > 0:
                final_rows_for_upsse.append(
                    add_summary_row(temp_up_sse_ws, bkhd_ws, "Dầu DO 0,001S-V", kvlD1, 1000, "4", headers,
                                g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table, final_rows_for_upsse))
            
            # Thêm các dòng tổng kết Thuế bảo vệ môi trường (TMT) vào cuối final_rows_for_upsse
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


            # Finalize the worksheet with the filtered and summarized data
            up_sse_wb_final = Workbook()
            up_sse_ws_final = up_sse_wb_final.active
            for row_data in final_rows_for_upsse:
                up_sse_ws_final.append(row_data)

            # Assign the final worksheet to up_sse_ws for subsequent processing
            up_sse_ws = up_sse_ws_final
            up_sse_wb = up_sse_wb_final # Also update workbook reference


            # --- Duyệt qua các hàng để thêm thuế TMT và format lại ---
            # Tạo một kiểu định dạng văn bản
            text_style = NamedStyle(name="text_style")
            text_style.number_format = '@'

            # Tạo một kiểu định dạng ngày
            date_style = NamedStyle(name="date_style")
            date_style.number_format = 'DD/MM/YYYY'

            # Các cột không cần chỉnh sửa định dạng
            # Dùng chỉ số 1-based vì up_sse_ws.cell(row, column) sử dụng 1-based
            exclude_columns_idx = {3, 13, 14, 15, 18, 19, 20, 21, 22, 37} # C, M, N, O, R, S, T, U, V, AK

            for row_idx in range(6, up_sse_ws.max_row + 1): # Bắt đầu từ dòng có dữ liệu (dòng 6)
                # Lấy giá trị cần thiết cho logic thuế TMT
                # Sử dụng 1-based indexing cho column trong openpyxl cell()
                current_h_value = up_sse_ws.cell(row=row_idx, column=8).value # Cột H
                current_m_value = up_sse_ws.cell(row=row_idx, column=13).value # Cột M
                current_n_value = up_sse_ws.cell(row=row_idx, column=14).value # Cột N
                current_af_value = up_sse_ws.cell(row=row_idx, column=32).value # Cột AF

                # Logic cập nhật cho các dòng thuế TMT
                # Điều kiện này được giữ lại nếu nó có tác dụng cập nhật các ô cụ thể trên các dòng hiện có,
                # nhưng nó KHÔNG TẠO RA DÒNG MỚI. Dòng TMT mới sẽ được thêm bằng add_tmt_summary_row.
                if (current_n_value is None or current_n_value == "") and current_h_value is not None:
                    lookup_key = str(current_h_value).strip().lower()
                    tmt_value = tmt_lookup_table.get(lookup_key, 0) # Already converted to float in tmt_lookup_table
                    
                    # Cột N (Giá bán)
                    up_sse_ws.cell(row=row_idx, column=14).value = tmt_value
                    # Cột O (Tiền hàng)
                    up_sse_ws.cell(row=row_idx, column=15).value = round(tmt_value * to_float(current_m_value), 0) if current_m_value is not None else 0.0 # Ensure current_m_value is float
                    # Cột S (Tk nợ)
                    up_sse_ws.cell(row=row_idx, column=19).value = s_lookup_table.get(h5_value, '') 
                    # Cột T (Tk doanh thu)
                    up_sse_ws.cell(row=row_idx, column=20).value = t_lookup_table.get(h5_value, '') 
                    # Cột U (Tk giá vốn)
                    up_sse_ws.cell(row=row_idx, column=21).value = u_value 
                    # Cột V (Tk thuế có)
                    up_sse_ws.cell(row=row_idx, column=22).value = v_lookup_table.get(h5_value, '') 
                    # Cột AK (Tiền thuế)
                    up_sse_ws.cell(row=row_idx, column=37).value = round(tmt_value * to_float(current_m_value) * 0.1, 0) if current_m_value is not None else 0.0 # Ensure current_m_value is float

                # Logic cập nhật cho các dòng "TMT" và "Thuế bảo vệ môi trường"
                # This part is still present, but if it's meant to *create* new rows, it won't work here.
                # If it's meant to modify existing rows that meet the criteria, it will.
                if (current_af_value is None or current_af_value == "") and current_h_value is not None:
                    # Cột G (Mã hàng)
                    up_sse_ws.cell(row=row_idx, column=7).value = "TMT"
                    # Cột H (Tên mặt hàng)
                    up_sse_ws.cell(row=row_idx, column=8).value = "Thuế bảo vệ môi trường"


            # Duyệt qua các ô để định dạng
            for r_idx in range(1, up_sse_ws.max_row + 1):
                for c_idx in range(1, up_sse_ws.max_column + 1):
                    cell = up_sse_ws.cell(row=r_idx, column=c_idx)
                    if cell.value is not None and str(cell.value) != "None":
                        # Chỉnh định dạng cột C (Ngày)
                        if c_idx == 3: # Cột C (1-based index)
                            if isinstance(cell.value, str):
                                try:
                                    cell.value = datetime.strptime(cell.value, '%Y-%m-%d').date()
                                except ValueError:
                                    pass # Giữ nguyên nếu không thể chuyển đổi
                            if isinstance(cell.value, datetime):
                                cell.number_format = 'DD/MM/YYYY' # Áp dụng định dạng ngày
                                cell.style = date_style
                        # Chuyển các cột khác sang văn bản trừ các cột loại trừ
                        elif c_idx not in exclude_columns_idx:
                            cell.value = str(cell.value)
                            cell.style = text_style

            # Chuyển ngược các cột R đến V thành text (đảm bảo chúng là văn bản dù không trong exclude_columns)
            # R (18) đến V (22)
            for r_idx in range(1, up_sse_ws.max_row + 1):
                for c_idx in range(18, 23):
                    cell = up_sse_ws.cell(row=r_idx, column=c_idx)
                    cell.number_format = '@' # Đặt định dạng văn bản

            # Mở rộng chiều rộng cột C,D cho khớp
            up_sse_ws.column_dimensions['C'].width = 12
            up_sse_ws.column_dimensions['D'].width = 12
            up_sse_ws.column_dimensions['B'].width = 35 # Tên khách hàng có thể dài


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
            st.exception(e) # Hiển thị stack trace để debug
