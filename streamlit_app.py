import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle, Alignment
from datetime import datetime
import io
import os
import re # Import regex module

# --- Streamlit page configuration ---
st.set_page_config(layout="centered", page_title="Đồng bộ dữ liệu SSE")

# Path to necessary files (assuming they are in the same directory as the script)
LOGO_PATH = "Logo.png"
DATA_FILE_PATH = "Data.xlsx" # Exact name of the data file

# Define headers for UpSSE.xlsx (Moved here to be always available)
headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng",
           "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế",
           "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm",
           "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế",
           "Nhóm Hàng", "Ghi chú", "Tiền thuế"]

# --- Check application expiration date ---
expiration_date = datetime(2025, 6, 26)
current_date = datetime.now()

if current_date > expiration_date:
    st.error("Có lỗi khi chạy chương trình, vui lòng liên hệ tác giả để được hỗ trợ!")
    st.info("Nguyễn Trọng Hoàn - 0902069469")
    st.stop() # Stop the application

# --- Helper function to safely convert values to float ---
def to_float(value):
    """Converts a value to float, returns 0.0 if conversion fails."""
    try:
        # Remove all spaces and commas (if it's a number format with commas)
        if isinstance(value, str):
            value = value.replace(",", "").strip()
        return float(value)
    except (ValueError, TypeError):
        return 0.0

# --- Helper function to clean string (remove all types of whitespace and standardize) ---
def clean_string(s):
    if s is None:
        return ""
    # Replace all types of whitespace (including non-breaking space, tabs, newlines) with a single space
    s = re.sub(r'\s+', ' ', str(s)).strip()
    return s

# --- Function to read static data and lookup tables from Data.xlsx ---
@st.cache_data
def get_static_data_from_excel(file_path):
    """
    Reads data and builds lookup tables from Data.xlsx.
    Uses openpyxl to read data. The result is cached.
    """
    try:
        wb = load_workbook(file_path, data_only=True)
        ws = wb.active

        listbox_data = [] # Data for the combobox to select gas station
        chxd_detail_map = {} # Map to store detailed gas station information (G5, H5, F5, B5)
        
        # Read data from the gas station table to build listbox_data and chxd_detail_map
        # Assumptions:
        # - Gas station name in column K (index 10)
        # - Customer code (G5) in column P (index 15)
        # - Warehouse code (F5) in column Q (index 16)
        # - Area (H5) in column S (index 17)
        # Start reading from row 4 (index 3)
        for row_idx in range(4, ws.max_row + 1):
            row_data_values = [cell.value for cell in ws[row_idx]]

            if len(row_data_values) >= 18: # Ensure enough columns to access index 17 (column S)
                raw_chxd_name = row_data_values[10] # Column K (index 10)

                if raw_chxd_name is not None and clean_string(raw_chxd_name) != '':
                    chxd_name_str = clean_string(raw_chxd_name)
                    
                    if chxd_name_str and chxd_name_str not in listbox_data: # Avoid duplication in listbox
                        listbox_data.append(chxd_name_str)

                    # Get values for G5, H5, F5_full, B5 from corresponding columns
                    g5_val = row_data_values[15] if len(row_data_values) > 15 and pd.notna(row_data_values[15]) else None # Column P (index 15)
                    f5_val_full = clean_string(row_data_values[16]) if len(row_data_values) > 16 and pd.notna(row_data_values[16]) else '' # Column Q (index 16)
                    h5_val = clean_string(row_data_values[17]).lower() if len(row_data_values) > 17 and pd.notna(row_data_values[17]) else '' # Column S (index 17)
                    b5_val = chxd_name_str # B5 is the gas station name

                    if f5_val_full: # Only add to map if Warehouse code has a value
                        chxd_detail_map[chxd_name_str] = {
                            'g5_val': g5_val,
                            'h5_val': h5_val,
                            'f5_val_full': f5_val_full,
                            'b5_val': b5_val
                        }
        
        # Read other lookup tables according to specific ranges
        # I4:J7 (Mã hàng <-> Tên mặt hàng)
        lookup_table = {}
        for row in ws.iter_rows(min_row=4, max_row=7, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                lookup_table[clean_string(row[0]).lower()] = row[1]
        
        # I10:J13 (TMT Lookup table - Tên mặt hàng <-> Mức phí BVMT)
        tmt_lookup_table = {}
        for row in ws.iter_rows(min_row=10, max_row=13, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                tmt_lookup_table[clean_string(row[0]).lower()] = to_float(row[1]) # Convert to float
        
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

        # Read J36 value (u_value)
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

# --- Function to add summary row for "Khách vãng lai" ---
def add_summary_row(all_temp_upsse_rows_list, bkhd_source_ws, product_name, headers_list,
                    g5_val, b5_val, s_lookup, t_lookup, v_lookup, x_lookup, u_val, h5_val, common_lookup_table, current_up_sse_rows_ref):
    """
    Creates a summary row for transient customers.
    all_temp_upsse_rows_list: Python list containing all processed UpSSE rows (including "Người mua không lấy hóa đơn")
    bkhd_source_ws: Original worksheet of the invoice statement (to get original values for TienhangHD, TienthueHD)
    """
    new_row = [''] * len(headers_list)
    new_row[0] = g5_val
    new_row[1] = f"Khách hàng mua {product_name} không lấy hóa đơn"

    c6_val = None
    e6_val = None
    # Get C6 and E6 values from the processed rows (current_up_sse_rows_ref)
    if len(current_up_sse_rows_ref) > 5 and len(current_up_sse_rows_ref[5]) > 4: 
        c6_val = current_up_sse_rows_ref[5][2] # Column C (index 2) of the 6th row
        e6_val = current_up_sse_rows_ref[5][4] # Column E (index 4) of the 6th row
    
    c6_val = c6_val if c6_val is not None else ""
    e6_val = e6_val if e6_val is not None else ""

    new_row[2] = c6_val # Column C (Date)
    new_row[4] = e6_val # Column E (Symbol)

    value_C = clean_string(new_row[2])
    value_E = clean_string(new_row[4])

    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name, "")

    if b5_val == "Nguyễn Huệ":
        value_D = f"HNBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh":
        value_D = f"MMBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    else:
        value_D = f"{value_E[-2:]}BK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    new_row[3] = value_D # Column D (Invoice Number)

    new_row[5] = f"Xuất bán lẻ theo hóa đơn số " + new_row[3] # Column F (Explanation)
    new_row[7] = product_name # Column H (Item Name)
    new_row[6] = common_lookup_table.get(clean_string(product_name).lower(), '') # Column G (Item Code)
    new_row[8] = "Lít" # Column I (Unit)
    new_row[9] = g5_val # Column J (Warehouse Code)
    new_row[10] = '' # Column K (Location Code)
    new_row[11] = '' # Column L (Batch Code)

    # Calculate total quantity (column M) and Max Selling Price (column N) from all_temp_upsse_rows_list
    total_M = 0.0
    max_value_N = 0.0
    # Loop starts from index 5 to skip initial 5 rows (headers, etc.)
    for r_data in all_temp_upsse_rows_list[5:]: # Iterate over actual data rows
        # Ensure the row has enough columns and correct values
        if len(r_data) > 12 and clean_string(r_data[1]) == "Người mua không lấy hóa đơn" and clean_string(r_data[7]) == product_name:
            total_M += to_float(r_data[12])
            current_N = to_float(r_data[13])
            if current_N > max_value_N: 
                max_value_N = current_N
    
    new_row[12] = total_M # Column M (Quantity)
    new_row[13] = max_value_N # Column N (Selling Price)

    tien_hang_hd_from_bkhd = 0.0
    for r in bkhd_source_ws.iter_rows(min_row=2, max_row=bkhd_source_ws.max_row, values_only=True):
        if clean_string(r[5]) == "Người mua không lấy hóa đơn" and clean_string(r[8]) == product_name:
            tien_hang_hd_from_bkhd += to_float(r[11]) # Column L on BKHD

    # Specific selling price for each product type
    price_per_liter_map = {"Xăng E5 RON 92-II": 1900, "Xăng RON 95-III": 2000, "Dầu DO 0,05S-II": 1000, "Dầu DO 0,001S-V": 1000}
    current_price_per_liter = price_per_liter_map.get(product_name, 0)
    
    new_row[14] = tien_hang_hd_from_bkhd - round(total_M * current_price_per_liter, 0) # Column O (Amount)

    new_row[15] = '' # Currency code
    new_row[16] = '' # Exchange rate
    new_row[17] = 10 # Tax code (Column R)

    new_row[18] = s_lookup.get(h5_val, '') # Debit account (Column S)
    new_row[19] = t_lookup.get(h5_val, '') # Revenue account (Column T)
    new_row[20] = u_val # Cost of goods sold account (Column U)
    new_row[21] = v_lookup.get(h5_val, '') # Tax credit account (Column V)
    new_row[22] = '' # Tax department (Column W)
    new_row[23] = x_lookup.get(clean_string(product_name).lower(), '') # Case (Column X)

    new_row[24] = '' # Department
    new_row[25] = '' # Production batch
    new_row[26] = '' # Product
    new_row[27] = '' # Contract
    new_row[28] = '' # Fee
    new_row[29] = '' # Pledge
    new_row[30] = '' # Sales employee

    new_row[31] = f"Khách mua {product_name} không lấy hóa đơn" # Customer name (tax) (Column AF)
    new_row[32] = "" # Address (tax) (Column AG) - No data in original file
    new_row[33] = "" # Tax code (Column AH) - No data in original file
    new_row[34] = '' # Item group
    new_row[35] = '' # Note

    tien_thue_hd_from_bkhd = 0.0
    for r in bkhd_source_ws.iter_rows(min_row=2, max_row=bkhd_source_ws.max_row, values_only=True):
        if clean_string(r[5]) == "Người mua không lấy hóa đơn" and clean_string(r[8]) == product_name:
            tien_thue_hd_from_bkhd += to_float(r[12]) # Column M on BKHD

    new_row[36] = tien_thue_hd_from_bkhd - round(total_M * current_price_per_liter * 0.1, 0) # Column AK (Tax amount)

    return new_row

# --- Function to add environmental protection tax (TMT) summary row ---
def add_tmt_summary_row(product_name_full, total_bvmt_amount, headers_list, g5_val, s_lookup, t_lookup, v_lookup, u_val, h5_val):
    """
    Creates a summary row for Environmental Protection Tax.
    """
    new_tmt_row = [''] * len(headers_list)
    new_tmt_row[0] = g5_val # Customer code
    new_tmt_row[1] = f"Thuế bảo vệ môi trường {product_name_full}" # Customer name (Explanation)
    
    new_tmt_row[6] = "TMT" # Item code
    new_tmt_row[7] = "Thuế bảo vệ môi trường" # Item name
    new_tmt_row[8] = "VNĐ" # Unit
    new_tmt_row[9] = g5_val # Warehouse code (same as Customer code)
    
    # Accounts
    new_tmt_row[18] = s_lookup.get(h5_val, '') # Debit account
    new_tmt_row[19] = t_lookup.get(h5_val, '') # Revenue account
    new_tmt_row[20] = u_val # Cost of goods sold account
    new_tmt_row[21] = v_lookup.get(h5_val, '') # Tax credit account

    new_tmt_row[36] = total_bvmt_amount # Tax amount (Total BVMT tax)
    return new_tmt_row


# Load static data and lookup maps from Data.xlsx
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

# --- Streamlit user interface ---
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
    options=[""] + listbox_data, # Add empty option to encourage user selection
    key='selected_chxd'
)

uploaded_file = st.file_uploader("Tải lên file bảng kê hóa đơn (.xlsx)", type=["xlsx"])

# --- Main processing when "Process" button is clicked ---
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
            
            # Get dynamic values from chxd_detail_map
            chxd_details = chxd_detail_map[selected_value_normalized]
            g5_value = chxd_details['g5_val']
            h5_value = chxd_details['h5_val']
            f5_value_full = chxd_details['f5_val_full'] 
            b5_value = chxd_details['b5_val'] 

            # Read invoice statement file from uploaded data
            bkhd_wb = load_workbook(uploaded_file)
            bkhd_ws = bkhd_wb.active # bkhd_ws will be the original worksheet

            # Check address length (column H)
            long_cells = []
            for r_idx, cell in enumerate(bkhd_ws['H']):
                if cell.value is not None and len(str(cell.value)) > 128:
                    long_cells.append(f"H{r_idx+1}")
            if long_cells:
                st.error("Địa chỉ trên ô " + ', '.join(long_cells) + " quá dài, hãy điều chỉnh và thử lại.")
                st.stop()

            # --- Prepare invoice statement data: delete first 3 rows and reorder columns ---
            # Create a copy of the original worksheet to avoid direct modification when iterating
            temp_bkhd_data_for_processing = []
            for row in bkhd_ws.iter_rows(min_row=1, values_only=True):
                temp_bkhd_data_for_processing.append(list(row)) 

            # Delete the first 3 rows (0-indexed)
            if len(temp_bkhd_data_for_processing) >= 3:
                temp_bkhd_data_for_processing = temp_bkhd_data_for_processing[3:]
            else:
                temp_bkhd_data_for_processing = []

            # Create a new temporary worksheet to save the initially processed invoice statement data
            temp_bkhd_ws_processed = Workbook().active
            for row_data in temp_bkhd_data_for_processing:
                temp_bkhd_ws_processed.append(row_data)

            # Column positions to keep and reorder (from your original file)
            # Old positions (original): ['A', 'B', 'C', 'D', 'E', 'F', 'H', 'G', 'I', 'K', 'L', 'N', 'O', 'Q']
            # Convert to 0-based indices: [0, 1, 2, 3, 4, 5, 7, 6, 8, 10, 11, 13, 14, 16]
            vi_tri_cu_idx = [0, 1, 2, 3, 4, 5, 7, 6, 8, 10, 11, 13, 14, 16]

            # Create new data table (intermediate_data_rows) with date processing (column D) and add "Công nợ" column
            intermediate_data_rows = [] 
            for row_idx, row_values_original in enumerate(temp_bkhd_ws_processed.iter_rows(min_row=1, values_only=True)):
                new_row_for_temp = [''] * (len(vi_tri_cu_idx) + 1) # +1 for Cong no column
                # Ensure the row has enough necessary columns (e.g., up to index 16 for column Q)
                if len(row_values_original) <= max(vi_tri_cu_idx): # Column Q is index 16. If original row has less than 17 cols, it's problematic
                    continue # Skip rows with insufficient data

                for idx_new_col, col_old_idx in enumerate(vi_tri_cu_idx):
                    cell_value = row_values_original[col_old_idx] 

                    if idx_new_col == 3 and cell_value: # New column D (Date)
                        cell_value_str = str(cell_value)[:10] 
                        try:
                            date_obj = datetime.strptime(cell_value_str, '%d-%m-%Y')
                            cell_value = date_obj.strftime('%Y-%m-%d')
                        except ValueError:
                            pass 
                    new_row_for_temp[idx_new_col] = cell_value # Assign to specific index
                
                # Add "Công nợ" column to the end of the row
                ma_kh_value = new_row_for_temp[4] # Column E (customer code) of the new row (index 4)
                if ma_kh_value is None or len(clean_string(ma_kh_value)) > 9:
                    new_row_for_temp.append("No") 
                else:
                    new_row_for_temp.append("Yes") 
                
                intermediate_data_rows.append(new_row_for_temp)

            # Create another temporary worksheet (temp_bkhd_ws_with_cong_no) containing the reordered columns and "Công nợ" column
            # This worksheet will be used as the main data source for calculations and UpSSE creation
            temp_bkhd_ws_with_cong_no = Workbook().active
            for row_data in intermediate_data_rows:
                temp_bkhd_ws_with_cong_no.append(row_data)

            # --- Check warehouse code: Compare warehouse code from Data.xlsx with B2 of the cleaned invoice statement ---
            # The warehouse code in the invoice statement is in column B of the first data row (cell B2 after deleting initial headers)
            # Based on previous debugging, cell B2 of temp_bkhd_ws_with_cong_no currently contains the actual warehouse code
            # Note: temp_bkhd_ws_with_cong_no might be empty if intermediate_data_rows is empty
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

            # --- STEP 1: Process and aggregate all data rows into a temporary structure ---
            # This is where we will calculate kvlE5, kvl95, kvlDo, kvlD1 and total_bvmt for each item
            # without being affected by subsequent filtering.
            
            # Initialize aggregation variables
            kvlE5 = 0
            kvl95 = 0
            kvlDo = 0
            kvlD1 = 0
            total_bvmt_e5 = 0.0 
            total_bvmt_95 = 0.0 
            total_bvmt_do = 0.0 
            total_bvmt_d1 = 0.0 

            # Python list to store all processed UpSSE rows, including initial header rows.
            # From here, we will build final_rows_for_upsse.
            all_processed_upsse_rows_list = []

            # Add 4 empty rows and header row to all_processed_upsse_rows_list
            for _ in range(4): 
                all_processed_upsse_rows_list.append([''] * len(headers))
            all_processed_upsse_rows_list.append(headers) # Add header row

            # Iterate through rows from temp_bkhd_ws_with_cong_no (with headers and "Công nợ" column)
            # Start from row 2 of temp_bkhd_ws_with_cong_no (which is the first data row)
            for row_idx_from_bkhd, row_values_from_bkhd in enumerate(temp_bkhd_ws_with_cong_no.iter_rows(min_row=2, values_only=True)):
                new_row_for_upsse = [''] * len(headers)

                # "Công nợ" column value (which is the last column of row_values_from_bkhd)
                cong_no_status = row_values_from_bkhd[-1] 

                # Logic for filling data into new_row_for_upsse (similar to previous times)
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

                # Accumulate total BVMT tax
                if new_row_for_upsse[7] == "Xăng E5 RON 92-II":
                    total_bvmt_e5 += thue_cua_tmt_for_row_bvmt
                elif new_row_for_upsse[7] == "Xăng RON 95-III":
                    total_bvmt_95 += thue_cua_tmt_for_row_bvmt
                elif new_row_for_upsse[7] == "Dầu DO 0,05S-II":
                    total_bvmt_do += thue_cua_tmt_for_row_bvmt
                elif new_row_for_upsse[7] == "Dầu DO 0,001S-V":
                    total_bvmt_d1 += thue_cua_tmt_for_row_bvmt

                # Add processed row to the temporary Python list
                all_processed_upsse_rows_list.append(new_row_for_upsse)

                # Count transient customer rows (used for transient customer summary)
                if clean_string(new_row_for_upsse[1]) == "Người mua không lấy hóa đơn":
                    if clean_string(new_row_for_upsse[7]) == "Xăng E5 RON 92-II":
                        kvlE5 += 1
                    elif clean_string(new_row_for_upsse[7]) == "Xăng RON 95-III":
                        kvl95 += 1
                    elif clean_string(new_row_for_upsse[7]) == "Dầu DO 0,05S-II":
                        kvlDo += 1
                    elif clean_string(new_row_for_upsse[7]) == "Dầu DO 0,001S-V":
                        kvlD1 += 1


            # --- STEP 2: Filter out "Người mua không lấy hóa đơn" rows and add transient customer summary rows ---
            final_rows_for_upsse = []

            # Add the first 5 rows (empty rows and header) to final_rows_for_upsse
            # Get directly from all_processed_upsse_rows_list (already has the first 5 rows)
            for r_idx in range(5): # Indices 0 to 4 of all_processed_upsse_rows_list
                final_rows_for_upsse.append(all_processed_upsse_rows_list[r_idx])
            
            # Iterate through actual data rows from all_processed_upsse_rows_list (from index 5 onwards)
            for r_idx in range(5, len(all_processed_upsse_rows_list)): # Iterate from index 5 to len-1
                row_data = all_processed_upsse_rows_list[r_idx]
                
                if len(row_data) > 1 and row_data[1] is not None:
                    col_b_value = clean_string(row_data[1])
                    
                    if col_b_value == "Người mua không lấy hóa đơn":
                        continue 
                    else:
                        final_rows_for_upsse.append(row_data)
                else:
                    final_rows_for_upsse.append(row_data) 

            # Add "Khách hàng mua..." summary rows
            # Pass all_processed_upsse_rows_list (containing all processed rows, including transient customers) to add_summary_row function
            if kvlE5 > 0:
                final_rows_for_upsse.append(
                    add_summary_row(all_processed_upsse_rows_list, bkhd_ws, "Xăng E5 RON 92-II", headers,
                                g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table, final_rows_for_upsse))
            if kvl95 > 0:
                final_rows_for_upsse.append(
                    add_summary_row(all_processed_upsse_rows_list, bkhd_ws, "Xăng RON 95-III", headers,
                                g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table, final_rows_for_upsse))
            if kvlDo > 0:
                final_rows_for_upsse.append(
                    add_summary_row(all_processed_upsse_rows_list, bkhd_ws, "Dầu DO 0,05S-II", headers,
                                g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table, final_rows_for_upsse))
            if kvlD1 > 0:
                final_rows_for_upsse.append(
                    add_summary_row(all_processed_upsse_rows_list, bkhd_ws, "Dầu DO 0,001S-V", headers,
                                g5_value, b5_value, s_lookup_table, t_lookup_table, v_lookup_table, x_lookup_table, u_value, h5_value, lookup_table, final_rows_for_upsse))
            
            # --- STEP 3: Add Environmental Protection Tax (TMT) summary rows ---
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


            # --- Write final data to new worksheet and apply formatting ---
            up_sse_wb_final = Workbook()
            up_sse_ws_final = up_sse_wb_final.active
            for row_data in final_rows_for_upsse:
                up_sse_ws_final.append(row_data)

            up_sse_ws = up_sse_ws_final # Assign to process formatting
            up_sse_wb = up_sse_wb_final

            # Define NamedStyles for formatting
            text_style = NamedStyle(name="text_style")
            text_style.number_format = '@'

            date_style = NamedStyle(name="date_style")
            date_style.number_format = 'DD/MM/YYYY'

            # Columns not to be formatted to text (using 1-based index)
            exclude_columns_idx = {3, 13, 14, 15, 18, 19, 20, 21, 22, 37} # C (Ngày), M (Số lượng), N (Giá bán), O (Tiền hàng), R (Mã thuế), S (Tk nợ), T (Tk doanh thu), U (Tk giá vốn), V (Tk thuế có), AK (Tiền thuế)

            for r_idx in range(1, up_sse_ws.max_row + 1):
                for c_idx in range(1, up_sse_ws.max_column + 1):
                    cell = up_sse_ws.cell(row=r_idx, column=c_idx)
                    if cell.value is not None and clean_string(cell.value) != "None": 
                        # Format column C (Date)
                        if c_idx == 3: 
                            if isinstance(cell.value, str):
                                try:
                                    cell.value = datetime.strptime(clean_string(cell.value), '%Y-%m-%d').date() 
                                except ValueError:
                                    pass 
                            if isinstance(cell.value, datetime):
                                cell.number_format = 'DD/MM/YYYY' 
                                cell.style = date_style
                        # Convert other columns to text except excluded columns
                        elif c_idx not in exclude_columns_idx:
                            cell.value = clean_string(cell.value) 
                            cell.style = text_style

            # Ensure columns R to V are Text (Even if already in exclude_columns_idx, need to ensure)
            for r_idx in range(1, up_sse_ws.max_row + 1):
                for c_idx in range(18, 23): # Column R (18) to V (22)
                    cell = up_sse_ws.cell(row=r_idx, column=c_idx)
                    cell.number_format = '@' 

            # Expand column width C, D, B
            up_sse_ws.column_dimensions['C'].width = 12
            up_sse_ws.column_dimensions['D'].width = 12
            up_sse_ws.column_dimensions['B'].width = 35 

            # Write the result file to buffer memory
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
