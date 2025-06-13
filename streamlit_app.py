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
# Đối với Streamlit Cloud, chúng ta cần đặt các file này cùng cấp với script hoặc trong một thư mục con
# Ví dụ: nếu Logo.png và Data.xlsx nằm trong cùng thư mục với streamlit_app.py
LOGO_PATH = "Logo.png"
DATA_FILE_PATH = "Data.xlsx"

# Kiểm tra ngày hết hạn (giữ nguyên logic từ code gốc)
expiration_date = datetime(2025, 6, 26)
current_date = datetime.now()

if current_date > expiration_date:
    st.error("Có lỗi khi chạy chương trình, vui lòng liên hệ tác giả để được hỗ trợ!")
    st.info("Nguyễn Trọng Hoàn - 0902069469")
    st.stop() # Dừng ứng dụng Streamlit

# --- Hàm hỗ trợ đọc dữ liệu từ Data.xlsx ---
@st.cache_data
def get_data_from_excel(file_path):
    """
    Hàm đọc dữ liệu từ Data.xlsx.
    Sử dụng @st.cache_data để cache kết quả, tránh đọc lại mỗi lần chạy lại script.
    """
    try:
        wb = load_workbook(file_path, data_only=True)
        ws = wb.active

        # Đọc dữ liệu cho listbox (cột K)
        data_listbox = [cell.value for cell in ws['K'] if cell.value is not None]

        # Đọc giá trị ô G5, H5, J36
        g5_value = ws['G5'].value
        h5_value = str(ws['H5'].value).strip().lower() if ws['H5'].value else ''
        j36_value = ws['J36'].value
        f5_value = str(ws['F5'].value) if ws['F5'].value else ''
        b5_value = str(ws['B5'].value) if ws['B5'].value else ''


        # Tạo bảng tra cứu cho các vùng dữ liệu
        lookup_table = {}
        tmt_lookup_table = {}
        s_lookup_table = {}
        t_lookup_table = {}
        v_lookup_table = {}
        x_lookup_table = {}

        for row in ws.iter_rows(min_row=4, max_row=7, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                lookup_table[str(row[0]).strip().lower()] = row[1]
        for row in ws.iter_rows(min_row=10, max_row=13, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                tmt_lookup_table[str(row[0]).strip().lower()] = row[1]
        for row in ws.iter_rows(min_row=29, max_row=31, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                s_lookup_table[str(row[0]).strip().lower()] = row[1]
        for row in ws.iter_rows(min_row=33, max_row=35, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                t_lookup_table[str(row[0]).strip().lower()] = row[1]
        for row in ws.iter_rows(min_row=53, max_row=55, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                v_lookup_table[str(row[0]).strip().lower()] = row[1]
        for row in ws.iter_rows(min_row=17, max_row=20, min_col=9, max_col=10, values_only=True):
            if row[0] and row[1]:
                x_lookup_table[str(row[0]).strip().lower()] = row[1]

        wb.close()
        return {
            "listbox_data": data_listbox,
            "g5_value": g5_value,
            "h5_value": h5_value,
            "j36_value": j36_value,
            "f5_value": f5_value,
            "b5_value": b5_value,
            "lookup_table": lookup_table,
            "tmt_lookup_table": tmt_lookup_table,
            "s_lookup_table": s_lookup_table,
            "t_lookup_table": t_lookup_table,
            "v_lookup_table": v_lookup_table,
            "x_lookup_table": x_lookup_table
        }
    except FileNotFoundError:
        st.error(f"Lỗi: Không tìm thấy file {file_path}. Vui lòng đảm bảo file tồn tại.")
        st.stop()
    except Exception as e:
        st.error(f"Lỗi không mong muốn khi đọc file Data.xlsx: {e}")
        st.stop()

# Tải dữ liệu từ Data.xlsx
data_excel = get_data_from_excel(DATA_FILE_PATH)
listbox_data = data_excel["listbox_data"]
g5_value = data_excel["g5_value"]
h5_value = data_excel["h5_value"]
u_value = data_excel["j36_value"]
f5_value = data_excel["f5_value"]
b5_value = data_excel["b5_value"]
lookup_table = data_excel["lookup_table"]
tmt_lookup_table = data_excel["tmt_lookup_table"]
s_lookup_table = data_excel["s_lookup_table"]
t_lookup_table = data_excel["t_lookup_table"]
v_lookup_table = data_excel["v_lookup_table"]
x_lookup_table = data_excel["x_lookup_table"]


# --- Header ứng dụng ---
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
            # Gán giá trị đã chọn vào ô A5 trong Data.xlsx tạm thời (trong Streamlit không cần lưu lại file Data)
            # Logic này sẽ được bỏ qua vì dữ liệu đã được đọc vào các biến Python.
            # Ta chỉ cần đảm bảo f5_value và b5_value được cập nhật dựa trên selected_value nếu cần
            # (tuy nhiên trong code gốc không có phần update f5_value hay b5_value dựa trên selected_value)

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
                        cell_value = str(cell_value)[:10]
                        try:
                            date_obj = datetime.strptime(cell_value, '%d-%m-%Y')
                            cell_value = date_obj.strftime('%Y-%m-%d')
                        except ValueError:
                            pass # Giữ nguyên giá trị nếu không phải ngày hợp lệ
                    new_row.append(cell_value)
                data.append(new_row)

            # Xóa tất cả dữ liệu cũ trong temp_ws và ghi dữ liệu mới
            # Cách hiệu quả hơn là tạo một worksheet mới hoặc ghi đè từ đầu
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

            # Kiểm tra 6 ký tự cuối của ô F5 trên data_ws (giá trị đã đọc ở đầu) với ô B2 trên bkhd_ws
            f5_data_value_sliced = f5_value[-6:] if f5_value else ''
            b2_bkhd_value = str(bkhd_ws['B2'].value) if bkhd_ws['B2'].value else ''

            if f5_data_value_sliced != b2_bkhd_value:
                st.error("Bảng kê hóa đơn không phải của cửa hàng bạn chọn.")
                st.stop()

            # Tiếp tục thực hiện các bước nếu trùng
            # Tạo file Excel mới - file UpSSE
            up_sse_wb = Workbook()
            up_sse_ws = up_sse_wb.active

            # Thêm các dòng trống trước tiêu đề
            for _ in range(4):
                up_sse_ws.append([])

            # Điền tiêu đề vào dòng thứ 5 của file UpSSE
            headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng",
                       "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế",
                       "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm",
                       "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế",
                       "Nhóm Hàng", "Ghi chú", "Tiền thuế"]
            up_sse_ws.append(headers)

            # Khởi tạo các biến phụ
            kvlE5 = 0
            kvl95 = 0
            kvlDo = 0
            kvlD1 = 0

            # Duyệt qua từng dòng của BKHD để tính toán và điền dữ liệu vào UpSSE
            for row_idx, row in enumerate(bkhd_ws.iter_rows(min_row=2, values_only=True)):
                new_row = [''] * len(headers)

                # Điều kiện cho cột A
                if row[14] == 'No':  # Cột O (index 14)
                    new_row[0] = g5_value
                elif row[14] == 'Yes':
                    if row[4] is None or row[4] == '':  # Cột E (index 4)
                        new_row[0] = g5_value
                    else:
                        new_row[0] = str((row[4]))  # Giá trị của cột E

                # Cột B: Điền bằng giá trị của cột F trên BKHD
                new_row[1] = row[5]  # Cột F (index 5)

                # Cột C: Điền bằng giá trị của cột D trên BKHD
                new_row[2] = row[3]  # Cột D (index 3)

                # Cột D: Điền là chuỗi ký tự bao gồm 2 ký tự cuối của cột B trên BKHD + 6 ký tự cuối của cột C trên BKHD
                if b5_value == "Nguyễn Huệ":
                    new_row[3] = "HN" + str(row[2])[-6:] # Điền HN + 6 ký tự cuối cột C trên BKHD
                elif b5_value == "Mai Linh":
                    new_row[3] = "MM" + str(row[2])[-6:] # Điền MM + 6 ký tự cuối cột C trên BKHD
                else:
                    new_row[3] = str(row[1])[-2:] + str(row[2])[-6:]  # Kết hợp 2 ký tự cuối của cột B và 6 ký tự cuối của cột C

                # Cột E: Điền bao gồm ký tự "1" và sau đó là giá trị cột B trên BKHD
                if row[1]:  # Cột B (index 1) không rỗng
                    new_row[4] = "1" + str(row[1])
                else:
                    new_row[4] = ''

                # Cột F: Điền một dãy ký tự bao gồm: “Xuất bán lẻ theo hóa đơn số” + giá trị Cột D (file UpSSE)
                new_row[5] = "Xuất bán lẻ theo hóa đơn số " + new_row[3]

                # Cột H: Điền bằng giá trị cột I trên BKHD
                new_row[7] = row[8]  # Cột I (index 8)

                # Cột G: Dò tìm giá trị của ô cùng dòng trên cột H trong ô I4:J6 của file Data.xlsx
                new_row[6] = lookup_table.get(str(new_row[7]).strip().lower(), '')

                # Cột I: Điền dãy ký tự "Lít"
                new_row[8] = "Lít"

                # Cột J: Điền giá trị của ô G5 trên file Data.xlsx
                new_row[9] = g5_value

                # Cột K và L: Để trống
                new_row[10] = ''
                new_row[11] = ''

                # Cột M: Điền bằng giá trị của cột J trên BKHD
                new_row[12] = row[9]  # Cột J (index 9)

                # Tính toán TMT dựa trên giá trị cột H
                tmt_value = tmt_lookup_table.get(str(new_row[7]).strip().lower(), 0)

                # Cột N: Giá trị cột K trên BKHD chia cho 1.1 rồi trừ TMT, làm tròn tới 2 chữ số thập phân
                if row[10] is not None:  # Cột K (index 10) không rỗng
                    new_row[13] = round(row[10] / 1.1 - tmt_value, 2)
                else:
                    new_row[13] = 0

                # Cột O: Bằng giá trị cột L trên file BKHD trừ đi (TMT nhân với giá trị cột M trên file UpSSE)
                if row[11] is not None and new_row[12] is not None:  # Cột L (index 11) và cột M (index 12) không rỗng
                    tmt_calculation = round(tmt_value * new_row[12])
                    new_row[14] = row[11] - tmt_calculation
                else:
                    new_row[14] = 0

                # Cột P và Q: Để trống
                new_row[15] = ''
                new_row[16] = ''

                # Cột R: Điền giá trị 10
                new_row[17] = 10

                # Cột S: Dò tìm giá trị của ô H5 trong vùng I29:I31 của file Data.xlsx và điền kết quả từ cột J cùng dòng
                s_value_from_lookup = s_lookup_table.get(h5_value, '')
                new_row[18] = s_value_from_lookup

                # Cột T: Dò tìm giá trị của ô H5 trong vùng I33:I35 của file Data.xlsx và điền kết quả từ cột J cùng dòng
                t_value_from_lookup = t_lookup_table.get(h5_value, '')
                new_row[19] = t_value_from_lookup

                # Cột U: Điền giá trị tại ô J36 của file Data.xlsx
                new_row[20] = u_value

                # Cột V: Dò tìm giá trị của ô H5 trong vùng I53:I55 của file Data.xlsx và điền kết quả từ cột J cùng dòng
                v_value_from_lookup = v_lookup_table.get(h5_value, '')
                new_row[21] = v_value_from_lookup

                # Cột X: Dò tìm giá trị của ô cùng dòng trên cột H trong vùng I17:I19 của file Data.xlsx và điền kết quả từ cột J cùng dòng
                h_value_for_x = str(new_row[7]).strip().lower()
                x_value_from_lookup = x_lookup_table.get(h_value_for_x, '')
                new_row[23] = x_value_from_lookup

                # Các cột Y, Z, AA, AB, AC, AD, AE: Để trống
                new_row[24] = ''
                new_row[25] = ''
                new_row[26] = ''
                new_row[27] = ''
                new_row[28] = ''
                new_row[29] = ''
                new_row[30] = ''

                # Cột AF: Điền bằng giá trị của cột B (của file UpSSE.xlsx)
                new_row[31] = new_row[1]

                # Cột AG: Điền bằng giá trị của cột G trên file BKHD
                new_row[32] = row[6]  # Cột G (index 6)

                # Cột AH: Điền giá trị của cột H trên file BKHD
                new_row[33] = row[7]  # Cột H (index 7)

                # Cột AI và AJ: Để trống
                new_row[34] = ''
                new_row[35] = ''

                # Cột AK: Tạo biến phụ Thue_Cua_TMT, làm tròn và tính toán
                if new_row[12] is not None and tmt_value is not None:  # Đảm bảo cột M và TMT có giá trị
                    thue_cua_tmt = round(new_row[12] * tmt_value * 0.1)  # Làm tròn đến hàng đơn vị
                    new_row[36] = row[12] - thue_cua_tmt  # Giá trị cột M trên BKHD trừ Thue_Cua_TMT
                else:
                    new_row[36] = row[12]  # Nếu không có TMT hoặc cột M trống, giữ nguyên giá trị cột M

                # Thêm dòng mới vào UpSSE
                up_sse_ws.append(new_row)

                # Đếm số lượng dòng thỏa mãn điều kiện cho kvlE5
                if new_row[1] == "Người mua không lấy hóa đơn" and new_row[7] == "Xăng E5 RON 92-II":
                    kvlE5 += 1
                # Đếm số lượng dòng thỏa mãn điều kiện cho kvl95
                if new_row[1] == "Người mua không lấy hóa đơn" and new_row[7] == "Xăng RON 95-III":
                    kvl95 += 1
                # Đếm số lượng dòng thỏa mãn điều kiện cho kvlDo
                if new_row[1] == "Người mua không lấy hóa đơn" and new_row[7] == "Dầu DO 0,05S-II":
                    kvlDo += 1
                # Đếm số lượng dòng thỏa mãn điều kiện cho kvlD1
                if new_row[1] == "Người mua không lấy hóa đơn" and new_row[7] == "Dầu DO 0,001S-V":
                    kvlD1 += 1


            # --- Thêm các dòng tổng kết (Khách vãng lai) ---
            # Hàm tạo dòng khách vãng lai
            def add_summary_row(ws_target, ws_source, product_name, sum_m_col, price_per_liter, suffix_d, headers_list):
                new_row = [''] * len(headers_list)
                new_row[0] = g5_value
                new_row[1] = f"Khách hàng mua {product_name} không lấy hóa đơn"

                # Lấy giá trị C6 và E6 từ UpSSE (dòng đầu tiên của dữ liệu chính, sau 5 hàng trống)
                # Dòng đầu tiên của dữ liệu chính là dòng 6
                new_row[2] = ws_target['C6'].value # Cột C
                new_row[4] = ws_target['E6'].value # Cột E

                value_C = new_row[2] if new_row[2] else ""
                value_E = new_row[4] if new_row[4] else ""

                if b5_value == "Nguyễn Huệ":
                    value_D = f"HNBK{str(value_C)[-2:]}.{str(value_C)[5:7]}.{suffix_d}"
                elif b5_value == "Mai Linh":
                    value_D = f"MMBK{str(value_C)[-2:]}.{str(value_C)[5:7]}.{suffix_d}"
                else:
                    value_D = f"{str(value_E)[-2:]}BK{str(value_C)[-2:]}.{str(value_C)[5:7]}.{suffix_d}"
                new_row[3] = value_D
                new_row[5] = f"Xuất bán lẻ theo hóa đơn số {new_row[3]}"
                new_row[7] = product_name
                new_row[6] = lookup_table.get(product_name.strip().lower(), '')
                new_row[8] = "Lít"
                new_row[9] = g5_value
                new_row[10] = ''
                new_row[11] = ''

                total_M = 0
                for r in ws_target.iter_rows(min_row=6, max_row=ws_target.max_row, values_only=True):
                    if r[1] == "Người mua không lấy hóa đơn" and r[7] == product_name:
                        total_M += r[12] if r[12] else 0
                new_row[12] = total_M

                max_value_N = None
                for r in ws_target.iter_rows(min_row=6, max_row=ws_target.max_row, values_only=True):
                    if r[1] == "Người mua không lấy hóa đơn" and r[7] == product_name:
                        if max_value_N is None or (r[13] is not None and r[13] > max_value_N):
                            max_value_N = r[13]
                new_row[13] = max_value_N

                tien_hang_hd = 0
                for r in ws_source.iter_rows(min_row=2, max_row=ws_source.max_row, values_only=True):
                    if r[5] == "Người mua không lấy hóa đơn" and r[8] == product_name:
                        tien_hang_hd += r[11] if r[11] is not None else 0
                new_row[14] = tien_hang_hd - round(total_M * price_per_liter, 0)

                new_row[17] = 10
                new_row[18] = s_lookup_table.get(h5_value, '')
                new_row[19] = t_lookup_table.get(h5_value, '')
                new_row[20] = u_value
                new_row[21] = v_lookup_table.get(h5_value, '')
                new_row[23] = x_lookup_table.get(product_name.strip().lower(), '')
                new_row[31] = f"Khách mua {product_name} không lấy hóa đơn"

                tien_thue_hd = 0
                for r in ws_source.iter_rows(min_row=2, max_row=ws_source.max_row, values_only=True):
                    if r[5] == "Người mua không lấy hóa đơn" and r[8] == product_name:
                        tien_thue_hd += r[12] if r[12] is not None else 0
                new_row[36] = tien_thue_hd - round(total_M * price_per_liter * 0.1)

                ws_target.append(new_row)


            if kvlE5 > 0:
                add_summary_row(up_sse_ws, bkhd_ws, "Xăng E5 RON 92-II", kvlE5, 1900, "1", headers)
            if kvl95 > 0:
                add_summary_row(up_sse_ws, bkhd_ws, "Xăng RON 95-III", kvl95, 2000, "2", headers)
            if kvlDo > 0:
                add_summary_row(up_sse_ws, bkhd_ws, "Dầu DO 0,05S-II", kvlDo, 1000, "3", headers)
            if kvlD1 > 0:
                add_summary_row(up_sse_ws, bkhd_ws, "Dầu DO 0,001S-V", kvlD1, 1000, "4", headers)


            # --- Xóa các dòng có cột B là "Người mua không lấy hóa đơn" từ phần chính ---
            # Lưu ý: trong Streamlit, ta không thể sửa đổi workbook trực tiếp như xlwings.
            # Ta sẽ tạo một list tạm thời các dòng hợp lệ và ghi lại.
            temp_rows_for_filtering = []
            for r_idx in range(1, up_sse_ws.max_row + 1):
                row_values = [cell.value for cell in up_sse_ws[r_idx]]
                temp_rows_for_filtering.append(row_values)

            filtered_rows = []
            for r_idx, row_data in enumerate(temp_rows_for_filtering):
                # Giữ lại 5 hàng đầu tiên (headers và dòng trống) hoặc nếu cột B không phải "Người mua không lấy hóa đơn"
                if r_idx < 5 or (len(row_data) > 1 and row_data[1] != "Người mua không lấy hóa đơn"):
                    filtered_rows.append(row_data)

            # Xóa nội dung worksheet cũ và ghi lại các dòng đã lọc
            up_sse_wb_filtered = Workbook()
            up_sse_ws_filtered = up_sse_wb_filtered.active
            for row_data in filtered_rows:
                up_sse_ws_filtered.append(row_data)
            up_sse_ws = up_sse_ws_filtered # Gán lại để tiếp tục xử lý

            # --- Duyệt qua các hàng để thêm thuế TMT và format lại ---
            # Tạo một kiểu định dạng văn bản
            text_style = NamedStyle(name="text_style")
            text_style.number_format = '@'

            # Tạo một kiểu định dạng ngày
            date_style = NamedStyle(name="date_style")
            date_style.number_format = 'DD/MM/YYYY'

            # Các cột không cần chỉnh sửa định dạng
            exclude_columns_idx = {2, 12, 13, 14, 17, 18, 19, 20, 21, 36} # Cột index 0-based
                                                                      # C (2), M (12), N (13), O (14), R (17), S (18), T (19), U (20), V (21), AK (36)

            for row_idx in range(6, up_sse_ws.max_row + 1): # Bắt đầu từ dòng có dữ liệu
                # Lấy giá trị cần thiết cho logic thuế TMT
                column_h_value = up_sse_ws.cell(row=row_idx, column=8).value # Cột H (index 7)
                column_m_value = up_sse_ws.cell(row=row_idx, column=13).value # Cột M (index 12)
                column_n_value = up_sse_ws.cell(row=row_idx, column=14).value # Cột N (index 13)
                column_af_value = up_sse_ws.cell(row=row_idx, column=32).value # Cột AF (index 31)

                # Cập nhật giá trị nếu cột N rỗng và cột H có giá trị
                if (column_n_value is None or column_n_value == "") and column_h_value is not None:
                    lookup_key = str(column_h_value).strip().lower()
                    tmt_value = tmt_lookup_table.get(lookup_key, 0)
                    s_value_from_lookup = s_lookup_table.get(h5_value, '')
                    t_value_from_lookup = t_lookup_table.get(h5_value,'')
                    v_value_from_lookup = v_lookup_table.get(h5_value,'')

                    up_sse_ws.cell(row=row_idx, column=18).value = s_value_from_lookup # Cột S
                    up_sse_ws.cell(row=row_idx, column=14).value = tmt_value # Cột N
                    up_sse_ws.cell(row=row_idx, column=15).value = round(tmt_value*column_m_value,0) if column_m_value is not None else 0 # Cột O
                    up_sse_ws.cell(row=row_idx, column=13).value = s_value_from_lookup # Cột M (kiểm tra lại mapping này)
                    up_sse_ws.cell(row=row_idx, column=20).value = u_value # Cột U (kiểm tra lại mapping này)
                    up_sse_ws.cell(row=row_idx, column=21).value = v_value_from_lookup # Cột V (kiểm tra lại mapping này)
                    up_sse_ws.cell(row=row_idx, column=37).value = round(tmt_value*column_m_value*0.1,0) if column_m_value is not None else 0 # Cột AK (index 36)

                # Cập nhật nếu cột AF rỗng và cột H có giá trị
                if (column_af_value is None or column_af_value == "") and column_h_value is not None:
                    up_sse_ws.cell(row=row_idx, column=1).value = "TMT" # Cột A (index 0)
                    up_sse_ws.cell(row=row_idx, column=2).value = "Thuế bảo vệ môi trường" # Cột B (index 1)


            # Duyệt qua các ô để định dạng
            for r_idx in range(1, up_sse_ws.max_row + 1):
                for c_idx in range(1, up_sse_ws.max_column + 1):
                    cell = up_sse_ws.cell(row=r_idx, column=c_idx)
                    if cell.value is not None and cell.value != "None":
                        # Chỉnh định dạng cột C (Ngày)
                        if c_idx == 3: # Cột C
                            if isinstance(cell.value, str):
                                try:
                                    cell.value = datetime.strptime(cell.value, '%Y-%m-%d').date()
                                    cell.style = date_style
                                except ValueError:
                                    pass # Giữ nguyên nếu không thể chuyển đổi
                        # Chuyển các cột khác sang văn bản trừ các cột loại trừ
                        elif c_idx not in exclude_columns_idx:
                            cell.value = str(cell.value)
                            cell.style = text_style

            # Chuyển ngược các cột R đến V thành text (index 1-based)
            for r_idx in range(1, up_sse_ws.max_row + 1):
                for c_idx in range(18, 23): # Cột R (18) đến V (22)
                    cell = up_sse_ws.cell(row=r_idx, column=c_idx)
                    cell.number_format = '@'

            # Mở rộng chiều rộng cột C,D cho khớp
            up_sse_ws.column_dimensions['C'].width = 12
            up_sse_ws.column_dimensions['D'].width = 12

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
