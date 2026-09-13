"""
Build TinhLai_DonGian.xlsx — Simple Bullet Loan Batch Calculator
Excel 2016 compatible, formula-based (no VBA)
For working capital / bullet loans: single disbursement, no interim principal repayments.
"""
import os, sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

# ─────────────────────────────────────────────
# STYLES (same palette as File 1)
# ─────────────────────────────────────────────
C_HEADER_BG  = "1F4E79"
C_HEADER_TXT = "FFFFFF"
C_INPUT_BG   = "DAEEF3"
C_CALC_BG    = "F2F2F2"
C_WARN_BG    = "FFF2CC"
C_OVERDUE_BG = "FCE4D6"
C_TOTAL_BG   = "D6E4BC"
C_WHITE      = "FFFFFF"
C_SECTION_BG = "BDD7EE"
C_ALT_ROW    = "EBF3F9"

FMT_VND  = '#,##0'
FMT_PCT  = '0.00%'
FMT_DATE = 'DD/MM/YYYY'
FMT_INT  = '#,##0'

def make_font(bold=False, size=10, color="000000", name="Calibri"):
    return Font(bold=bold, size=size, color=color, name=name)

def make_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def make_border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

thin_border  = make_border("thin")
thick_border = make_border("medium")

def style_header(ws, row, col, value, colspan=1):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = make_font(bold=True, size=9, color=C_HEADER_TXT)
    cell.fill = make_fill(C_HEADER_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border
    if colspan > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col+colspan-1)
    return cell

def style_input(ws, row, col, value=None, fmt=None, align="left"):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = make_fill(C_INPUT_BG)
    cell.font = make_font(size=9)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = thin_border
    if fmt:
        cell.number_format = fmt
    return cell

def style_calc(ws, row, col, formula=None, fmt=None, align="right", bg=C_CALC_BG, bold=False, color="000000"):
    cell = ws.cell(row=row, column=col, value=formula)
    cell.fill = make_fill(bg)
    cell.font = make_font(bold=bold, size=9, color=color)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = thin_border
    if fmt:
        cell.number_format = fmt
    return cell

def style_total(ws, row, col, value=None, fmt=None, colspan=1):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = make_fill(C_TOTAL_BG)
    cell.font = make_font(bold=True, size=9)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.border = thick_border
    if fmt:
        cell.number_format = fmt
    if colspan > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col+colspan-1)
    return cell

def set_print_a4_landscape(ws, title_text=""):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 9
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(
        left=0.4, right=0.4, top=0.6, bottom=0.6, header=0.3, footer=0.3
    )
    ws.oddHeader.center.text = title_text
    ws.oddFooter.left.text   = "Trang &P / &N"
    ws.oddFooter.right.text  = "In ngày: &D"


# ─────────────────────────────────────────────
# HOLIDAYS (same as File 1 — subset for clarity)
# ─────────────────────────────────────────────
def get_holidays():
    from datetime import date as d
    holidays = [
        (d(2020,1,1),"Tết DL","Lễ"),(d(2020,1,23),"Tết ÂL 29TrC","Tết"),
        (d(2020,1,24),"Tết ÂL 30TrC","Tết"),(d(2020,1,25),"Tết Mùng 1","Tết"),
        (d(2020,1,26),"Tết Mùng 2","Tết"),(d(2020,1,27),"Tết Mùng 3","Tết"),
        (d(2020,4,22),"Giỗ Tổ","Lễ"),(d(2020,4,30),"30/4","Lễ"),
        (d(2020,5,1),"1/5","Lễ"),(d(2020,9,2),"QK 2/9","Lễ"),(d(2020,9,3),"Bù QK","Nghỉ bù"),
        (d(2021,1,1),"Tết DL","Lễ"),(d(2021,2,10),"Tết ÂL 29TrC","Tết"),
        (d(2021,2,11),"Tết ÂL 30TrC","Tết"),(d(2021,2,12),"Tết Mùng 1","Tết"),
        (d(2021,2,13),"Tết Mùng 2","Tết"),(d(2021,2,14),"Tết Mùng 3","Tết"),
        (d(2021,4,21),"Giỗ Tổ","Lễ"),(d(2021,4,30),"30/4","Lễ"),
        (d(2021,5,3),"Bù 1/5","Nghỉ bù"),(d(2021,9,2),"QK 2/9","Lễ"),(d(2021,9,3),"Bù QK","Nghỉ bù"),
        (d(2022,1,3),"Bù Tết DL","Nghỉ bù"),(d(2022,1,29),"Tết ÂL 28TrC","Tết"),
        (d(2022,1,30),"Tết ÂL 29TrC","Tết"),(d(2022,1,31),"Tết ÂL 30TrC","Tết"),
        (d(2022,2,1),"Tết Mùng 1","Tết"),(d(2022,2,2),"Tết Mùng 2","Tết"),
        (d(2022,4,3),"Bù Giỗ Tổ","Nghỉ bù"),(d(2022,4,10),"Giỗ Tổ","Lễ"),
        (d(2022,4,30),"30/4","Lễ"),(d(2022,5,2),"Bù 1/5","Nghỉ bù"),
        (d(2022,5,3),"Bù 30/4","Nghỉ bù"),(d(2022,9,1),"Bù QK","Nghỉ bù"),(d(2022,9,2),"QK 2/9","Lễ"),
        (d(2023,1,1),"Tết DL","Lễ"),(d(2023,1,20),"Tết ÂL 29TrC","Tết"),
        (d(2023,1,21),"Tết ÂL 30TrC","Tết"),(d(2023,1,22),"Tết Mùng 1","Tết"),
        (d(2023,1,23),"Tết Mùng 2","Tết"),(d(2023,1,24),"Tết Mùng 3","Tết"),
        (d(2023,4,29),"Bù Giỗ Tổ","Nghỉ bù"),(d(2023,4,30),"30/4","Lễ"),
        (d(2023,5,1),"1/5","Lễ"),(d(2023,9,1),"Bù QK","Nghỉ bù"),(d(2023,9,2),"QK 2/9","Lễ"),
        (d(2024,1,1),"Tết DL","Lễ"),(d(2024,2,8),"Tết ÂL 28TrC","Tết"),
        (d(2024,2,9),"Tết ÂL 29TrC","Tết"),(d(2024,2,10),"Tết Mùng 1","Tết"),
        (d(2024,2,12),"Tết Mùng 2","Tết"),(d(2024,2,13),"Tết Mùng 3","Tết"),
        (d(2024,4,18),"Giỗ Tổ","Lễ"),(d(2024,4,26),"Bù 30/4","Nghỉ bù"),
        (d(2024,4,29),"Bù 30/4","Nghỉ bù"),(d(2024,4,30),"30/4","Lễ"),
        (d(2024,5,1),"1/5","Lễ"),(d(2024,9,2),"QK 2/9","Lễ"),(d(2024,9,3),"Bù QK","Nghỉ bù"),
        (d(2025,1,1),"Tết DL","Lễ"),(d(2025,1,25),"Tết ÂL 26TrC","Tết"),
        (d(2025,1,27),"Tết ÂL 28TrC","Tết"),(d(2025,1,28),"Tết ÂL 29TrC","Tết"),
        (d(2025,1,29),"Tết ÂL 30TrC","Tết"),(d(2025,1,30),"Tết Mùng 1","Tết"),
        (d(2025,1,31),"Tết Mùng 2","Tết"),(d(2025,2,3),"Bù Tết","Nghỉ bù"),
        (d(2025,4,6),"Bù Giỗ Tổ","Nghỉ bù"),(d(2025,4,7),"Giỗ Tổ","Lễ"),
        (d(2025,4,30),"30/4","Lễ"),(d(2025,5,1),"1/5","Lễ"),(d(2025,5,2),"Bù 30/4","Nghỉ bù"),
        (d(2025,9,1),"Bù QK","Nghỉ bù"),(d(2025,9,2),"QK 2/9","Lễ"),
        (d(2026,1,1),"Tết DL","Lễ"),(d(2026,1,16),"Tết ÂL 28TrC","Tết"),
        (d(2026,1,17),"Tết ÂL 29TrC","Tết"),(d(2026,1,18),"Tết ÂL 30TrC","Tết"),
        (d(2026,1,19),"Tết Mùng 1","Tết"),(d(2026,1,20),"Tết Mùng 2","Tết"),
        (d(2026,3,27),"Giỗ Tổ","Lễ"),(d(2026,4,30),"30/4","Lễ"),
        (d(2026,5,1),"1/5","Lễ"),(d(2026,9,2),"QK 2/9","Lễ"),(d(2026,9,3),"Bù QK","Nghỉ bù"),
        (d(2027,1,1),"Tết DL","Lễ"),(d(2027,1,6),"Tết ÂL 29TrC","Tết"),
        (d(2027,1,7),"Tết ÂL 30TrC","Tết"),(d(2027,1,8),"Tết Mùng 1","Tết"),
        (d(2027,1,9),"Tết Mùng 2","Tết"),(d(2027,1,10),"Tết Mùng 3","Tết"),
        (d(2027,4,16),"Giỗ Tổ","Lễ"),(d(2027,4,30),"30/4","Lễ"),
        (d(2027,5,3),"Bù 1/5","Nghỉ bù"),(d(2027,9,2),"QK 2/9","Lễ"),(d(2027,9,3),"Bù QK","Nghỉ bù"),
        (d(2028,1,1),"Tết DL","Lễ"),(d(2028,1,26),"Tết ÂL 29TrC","Tết"),
        (d(2028,1,27),"Tết ÂL 30TrC","Tết"),(d(2028,1,28),"Tết Mùng 1","Tết"),
        (d(2028,1,29),"Tết Mùng 2","Tết"),(d(2028,1,30),"Tết Mùng 3","Tết"),
        (d(2028,5,5),"Giỗ Tổ","Lễ"),(d(2028,4,30),"30/4","Lễ"),
        (d(2028,5,1),"1/5","Lễ"),(d(2028,9,2),"QK 2/9","Lễ"),(d(2028,9,4),"Bù QK","Nghỉ bù"),
        (d(2029,1,1),"Tết DL","Lễ"),(d(2029,2,12),"Tết ÂL 29TrC","Tết"),
        (d(2029,2,13),"Tết ÂL 30TrC","Tết"),(d(2029,2,14),"Tết Mùng 1","Tết"),
        (d(2029,2,15),"Tết Mùng 2","Tết"),(d(2029,2,16),"Tết Mùng 3","Tết"),
        (d(2029,4,24),"Giỗ Tổ","Lễ"),(d(2029,4,30),"30/4","Lễ"),
        (d(2029,5,1),"1/5","Lễ"),(d(2029,9,2),"QK 2/9","Lễ"),(d(2029,9,3),"Bù QK","Nghỉ bù"),
        (d(2030,1,1),"Tết DL","Lễ"),(d(2030,2,2),"Tết ÂL 29TrC","Tết"),
        (d(2030,2,3),"Tết ÂL 30TrC","Tết"),(d(2030,2,4),"Tết Mùng 1","Tết"),
        (d(2030,2,5),"Tết Mùng 2","Tết"),(d(2030,2,6),"Tết Mùng 3","Tết"),
        (d(2030,4,13),"Giỗ Tổ","Lễ"),(d(2030,4,30),"30/4","Lễ"),
        (d(2030,5,1),"1/5","Lễ"),(d(2030,9,2),"QK 2/9","Lễ"),(d(2030,9,3),"Bù QK","Nghỉ bù"),
    ]
    return sorted(holidays, key=lambda x: x[0])


# ─────────────────────────────────────────────
# SHEET BUILDERS
# ─────────────────────────────────────────────

def build_cauhinh(ws):
    ws.title = "CẤU HÌNH"
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 35

    ws.merge_cells("A1:C1")
    c = ws["A1"]
    c.value = "⚙ CẤU HÌNH — CÔNG CỤ TÍNH LÃI KHOẢN VAY NGẮN HẠN / BULLET"
    c.font = make_font(bold=True, size=13, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:C2")
    c = ws["A2"]
    c.value = "Dùng cho khoản vay bullet đơn giản: giải ngân 1 lần, trả cuối kỳ. Căn cứ TT39/2016/TT-NHNN."
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    params = [
        ("Ngày tính toán (ngày tòa / ngày cắt)", date(2025, 6, 15), FMT_DATE,
         "Ngày tòa xét xử hoặc ngày cần tính tổng nợ"),
        ("Quy ước đếm ngày", "Actual/365", None,
         "Actual/365 là chuẩn phổ biến. Xem thêm: Actual/360, 30/360"),
        ("Hệ số nhân lãi suất quá hạn", 1.5, '0.00"×"',
         "Tối đa 150% theo Điều 13 TT39. Nhập: 1.5 = 150%"),
        ("Lãi suất phạt trên lãi chậm trả (%/năm)", 0.10, FMT_PCT,
         "Tối đa 10%/năm theo Điều 13 TT39"),
    ]

    ws.row_dimensions[3].height = 8
    for i, (label, val, fmt, note) in enumerate(params):
        r = 4 + i
        ws.row_dimensions[r].height = 22

        c1 = ws.cell(row=r, column=1, value=label)
        c1.font = make_font(bold=False, size=10)
        c1.fill = make_fill(C_WHITE)
        c1.border = thin_border
        c1.alignment = Alignment(horizontal="left", vertical="center")

        c2 = style_input(ws, r, 2, val, fmt, "center" if fmt == FMT_DATE else "right")

        c3 = ws.cell(row=r, column=3, value=note)
        c3.font = make_font(size=9, color="595959")
        c3.fill = make_fill(C_WHITE)
        c3.border = thin_border
        c3.alignment = Alignment(horizontal="left", vertical="center")

    # Data validations
    dv_dc = DataValidation(type="list",
                           formula1='"Actual/365,Actual/360,30/360"',
                           allow_blank=False, showDropDown=False)
    dv_dc.sqref = "B5"
    ws.add_data_validation(dv_dc)

    # Instructions
    ws.row_dimensions[9].height = 8
    ws.merge_cells("A10:C10")
    c = ws["A10"]
    c.value = "▶ HƯỚNG DẪN SỬ DỤNG"
    c.font = make_font(bold=True, size=10)
    c.fill = make_fill(C_SECTION_BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[10].height = 18

    steps = [
        "Bước 1: Nhập cấu hình trên sheet này (ngày tính, quy ước ngày, hệ số phạt)",
        "Bước 2: Cập nhật ngày nghỉ lễ trên sheet NGÀY NGHỈ LỄ (nếu cần)",
        "Bước 3: Nhập dữ liệu từng khoản vay trên sheet BẢNG TÍNH (1 hàng = 1 khoản vay)",
        "Bước 4: Xem tổng hợp tất cả khoản vay trên sheet TỔNG HỢP",
        "Lưu ý: Nếu khoản vay có thay đổi lãi suất hoặc trả nhiều đợt → dùng file TinhLai_ChiTiet.xlsx",
        "Ô nền XANH NHẠT = nhập liệu. Ô nền XÁM = công thức tự tính.",
    ]
    for i, step in enumerate(steps):
        r = 11 + i
        ws.row_dimensions[r].height = 16
        ws.merge_cells(f"A{r}:C{r}")
        c = ws.cell(row=r, column=1, value=step)
        c.font = make_font(size=9)
        c.fill = make_fill(C_WHITE if i % 2 == 0 else C_ALT_ROW)
        c.border = thin_border
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    set_print_a4_landscape(ws, "CẤU HÌNH — KHOẢN VAY BULLET")


def build_holidays(ws):
    ws.title = "NGÀY NGHỈ LỄ"
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 14

    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value = "📅 NGÀY NGHỈ LỄ VIỆT NAM 2020–2030"
    c.font = make_font(bold=True, size=13, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:D2")
    c = ws["A2"]
    c.value = "Có thể bổ sung/chỉnh sửa. Cột B phải là kiểu Date. Sắp xếp theo ngày tăng dần."
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 14

    for col, h in enumerate(["STT","Ngày","Tên ngày lễ","Loại"], 1):
        style_header(ws, 3, col, h)
    ws.row_dimensions[3].height = 16
    ws.freeze_panes = "A4"

    holidays = get_holidays()
    for i, (d, name, htype) in enumerate(holidays):
        r = 4 + i
        alt = (i % 2 == 1)
        bg = C_ALT_ROW if alt else C_WHITE

        c1 = ws.cell(row=r, column=1, value=i+1)
        c1.font = make_font(size=9); c1.fill = make_fill(bg)
        c1.alignment = Alignment(horizontal="center"); c1.border = thin_border

        c2 = ws.cell(row=r, column=2, value=d)
        c2.number_format = FMT_DATE; c2.font = make_font(size=9)
        c2.fill = make_fill(C_INPUT_BG); c2.alignment = Alignment(horizontal="center")
        c2.border = thin_border

        c3 = ws.cell(row=r, column=3, value=name)
        c3.font = make_font(size=9); c3.fill = make_fill(C_INPUT_BG)
        c3.alignment = Alignment(horizontal="left"); c3.border = thin_border

        c4 = ws.cell(row=r, column=4, value=htype)
        c4.font = make_font(size=9); c4.fill = make_fill(C_INPUT_BG)
        c4.alignment = Alignment(horizontal="center"); c4.border = thin_border

        ws.row_dimensions[r].height = 14

    last_data = 3 + len(holidays)
    for i in range(10):
        r = last_data + 1 + i
        for col in range(1, 5):
            c = ws.cell(row=r, column=col)
            c.fill = make_fill(C_INPUT_BG); c.border = thin_border
            if col == 2: c.number_format = FMT_DATE
        ws.row_dimensions[r].height = 14

    set_print_a4_landscape(ws, "NGÀY NGHỈ LỄ 2020–2030")


def build_bangtính(ws):
    """
    Main calculation sheet for simple bullet loans.
    Each row = 1 loan.
    Columns:
    A: STT
    B: Tên / Số HĐ
    C: Ngày giải ngân
    D: Số tiền gốc (VND)
    E: Lãi suất trong hạn (%/năm)
    F: Ngày đáo hạn (theo HĐ)
    G: Ngày đáo hạn (điều chỉnh) — formula
    H: Ngày thanh toán thực tế (nếu đã trả)
    I: Số tiền đã thanh toán (gốc + lãi, total)
    J: Trạng thái — formula (Trong hạn / Quá hạn / Đã trả)
    K: Số ngày trong hạn — formula
    L: Lãi trong hạn (VND) — formula
    M: Gốc còn nợ quá hạn — formula
    N: Số ngày quá hạn — formula
    O: Lãi trên gốc quá hạn (VND) — formula
    P: Lãi chưa trả đã quá hạn (lãi đến hạn chưa được thanh toán) — formula
    Q: Lãi chậm trả (trên lãi QH chưa trả) — formula
    R: Tổng nợ tại ngày tính — formula
    """
    ws.title = "BẢNG TÍNH"

    col_info = [
        ("A", 6,  "STT"),
        ("B", 22, "Số HĐ / Tên khoản vay"),
        ("C", 16, "Ngày\ngiải ngân"),
        ("D", 20, "Số tiền gốc\n(VND)"),
        ("E", 14, "LS trong hạn\n(%/năm)"),
        ("F", 16, "Ngày đáo hạn\n(theo HĐ)"),
        ("G", 16, "Ngày đáo hạn\n(Điều chỉnh)"),
        ("H", 16, "Ngày TT\nthực tế"),
        ("I", 20, "Số tiền\nđã TT (VND)"),
        ("J", 14, "Trạng thái"),
        ("K", 14, "Số ngày\ntrong hạn"),
        ("L", 20, "Lãi trong\nhạn (VND)"),
        ("M", 20, "Gốc quá hạn\n(VND)"),
        ("N", 12, "Số ngày\nquá hạn"),
        ("O", 22, "Lãi trên Gốc\nquá hạn (VND)"),
        ("P", 22, "Lãi IH chưa\ntrả (quá hạn)"),
        ("Q", 22, "Lãi chậm trả\n(VND)"),
        ("R", 22, "TỔNG NỢ tại\nngày tính (VND)"),
    ]

    for col_ltr, w, _ in col_info:
        ws.column_dimensions[col_ltr].width = w

    # Title
    ws.merge_cells("A1:R1")
    c = ws["A1"]
    c.value = "💰 BẢNG TÍNH LÃI — KHOẢN VAY BULLET / NGẮN HẠN"
    c.font = make_font(bold=True, size=13, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:R2")
    c = ws["A2"]
    c.value = (
        "Mỗi hàng = 1 khoản vay. "
        "Ô xanh nhạt = nhập liệu. Ô xám = công thức tự tính. "
        "Ngày tính: CẤU HÌNH!B4 | LS quá hạn = LS trong hạn × CẤU HÌNH!B6 | "
        "LS chậm trả = CẤU HÌNH!B7"
    )
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 24

    # Header row
    for col_idx, (_, _, header) in enumerate(col_info, 1):
        style_header(ws, 3, col_idx, header)
    ws.row_dimensions[3].height = 34
    ws.freeze_panes = "A4"

    # References
    hol_ref   = "'NGÀY NGHỈ LỄ'!$B$4:$B$200"
    cfg_date  = "'CẤU HÌNH'!$B$4"
    cfg_dc    = "IF('CẤU HÌNH'!$B$5=\"Actual/360\",360,IF('CẤU HÌNH'!$B$5=\"30/360\",360,365))"
    cfg_qh    = "'CẤU HÌNH'!$B$6"    # multiplier
    cfg_ct    = "'CẤU HÌNH'!$B$7"    # late interest rate

    # Sample bullet loans
    sample_loans = [
        ("HĐ-2024-001", date(2024,1,15), 1_000_000_000, 0.09, date(2024,7,15),
         date(2024,7,16), 1_000_000_000 + 45_000_000, "Trả đúng hạn (TT 1 ngày sau đáo hạn điều chỉnh)"),
        ("HĐ-2024-002", date(2024,3,1),  500_000_000,  0.10, date(2024,9,1),
         None, None, "Chưa trả — đang quá hạn"),
        ("HĐ-2024-003", date(2024,6,15), 2_000_000_000, 0.095, date(2025,6,15),
         None, None, "Đến hạn đúng ngày tính"),
        ("HĐ-2024-004", date(2024,8,1),  750_000_000,  0.10, date(2025,2,1),
         date(2025,2,10), 786_000_000, "Trả muộn 9 ngày"),
        ("HĐ-2024-005", date(2024,11,15),3_000_000_000, 0.11, date(2025,11,15),
         None, None, "Chưa đến hạn"),
    ]

    FIRST_DATA_ROW = 4
    for i, (contract, dis_date, principal, rate, due_hd, pay_date, paid_amt, note) in enumerate(sample_loans):
        r = FIRST_DATA_ROW + i
        ws.row_dimensions[r].height = 18
        alt_bg = C_ALT_ROW if i % 2 == 1 else C_WHITE

        # A: STT
        c = ws.cell(row=r, column=1, value=i+1)
        c.font = make_font(size=9); c.fill = make_fill(alt_bg)
        c.alignment = Alignment(horizontal="center"); c.border = thin_border

        # B: Contract name
        style_input(ws, r, 2, contract)

        # C: Disbursement date
        style_input(ws, r, 3, dis_date, FMT_DATE, "center")

        # D: Principal
        style_input(ws, r, 4, principal, FMT_VND, "right")

        # E: In-term rate
        style_input(ws, r, 5, rate, FMT_PCT, "right")

        # F: Due date (contractual)
        style_input(ws, r, 6, due_hd, FMT_DATE, "center")

        # G: Adjusted due date (formula)
        cg = ws.cell(row=r, column=7)
        cg.value = (f'=IF(F{r}="","",IF(AND(WEEKDAY(F{r},2)<=5,'
                    f'COUNTIF({hol_ref},F{r})=0),F{r},'
                    f'WORKDAY.INTL(F{r},1,1,{hol_ref})))')
        cg.number_format = FMT_DATE
        cg.fill = make_fill(C_CALC_BG)
        cg.font = make_font(size=9)
        cg.alignment = Alignment(horizontal="center", vertical="center")
        cg.border = thin_border

        # H: Actual payment date
        style_input(ws, r, 8, pay_date, FMT_DATE, "center")

        # I: Amount paid
        style_input(ws, r, 9, paid_amt, FMT_VND, "right")

        # J: Status
        cj = ws.cell(row=r, column=10)
        # Status: if paid (H not empty and I >= D) → "Đã trả"
        # elif calculation date < adjusted due → "Trong hạn"
        # J: Status
        cj = ws.cell(row=r, column=10)
        cj.value = (f'=IF(C{r}="","",IF(AND(H{r}<>"",I{r}>=(D{r}+L{r}+O{r}+Q{r})),"✅ Đã tất toán",'
                    f'IF(AND(H{r}<>"",I{r}>0),"⚠️ Đã trả một phần",'
                    f'IF({cfg_date}<G{r},"⏳ Trong hạn","⚠ Quá hạn"))))')
        cj.fill = make_fill(C_CALC_BG)
        cj.font = make_font(size=9)
        cj.alignment = Alignment(horizontal="center", vertical="center")
        cj.border = thin_border

        # K: Days in-term
        ck = ws.cell(row=r, column=11)
        ck.value = (f'=IF(C{r}="","",MAX(0,MIN(G{r},IF(H{r}<>"",H{r},{cfg_date}))-C{r}))')
        ck.number_format = FMT_INT
        ck.fill = make_fill(C_CALC_BG)
        ck.font = make_font(size=9)
        ck.alignment = Alignment(horizontal="right", vertical="center")
        ck.border = thin_border

        # L: In-term interest
        cl = ws.cell(row=r, column=12)
        cl.value = f'=IF(C{r}="","",D{r}*E{r}*K{r}/({cfg_dc}))'
        cl.number_format = FMT_VND
        cl.fill = make_fill(C_CALC_BG)
        cl.font = make_font(size=9)
        cl.alignment = Alignment(horizontal="right", vertical="center")
        cl.border = thin_border

        # M: Overdue principal remaining at calc date
        cm = ws.cell(row=r, column=13)
        cm.value = (f'=IF(C{r}="","",IF({cfg_date}<=G{r},0,'
                    f'MAX(0,D{r}-IF(H{r}<>"",I{r},0))))')
        cm.number_format = FMT_VND
        cm.fill = make_fill(C_OVERDUE_BG)
        cm.font = make_font(size=9, color="C00000")
        cm.alignment = Alignment(horizontal="right", vertical="center")
        cm.border = thin_border

        # N: Overdue days
        cn = ws.cell(row=r, column=14)
        cn.value = (f'=IF(C{r}="","",MAX(0,'
                    f'MIN(IF(H{r}<>"",H{r},{cfg_date}),{cfg_date})-G{r}))')
        cn.number_format = FMT_INT
        cn.fill = make_fill(C_OVERDUE_BG)
        cn.font = make_font(size=9, color="C00000")
        cn.alignment = Alignment(horizontal="right", vertical="center")
        cn.border = thin_border

        # O: Interest on overdue principal during overdue period
        co = ws.cell(row=r, column=15)
        co.value = f'=IF(C{r}="","",IF(N{r}>0,D{r}*(E{r}*{cfg_qh})*N{r}/({cfg_dc}),0))'
        co.number_format = FMT_VND
        co.fill = make_fill(C_OVERDUE_BG)
        co.font = make_font(size=9, color="C00000")
        co.alignment = Alignment(horizontal="right", vertical="center")
        co.border = thin_border

        # P: Overdue in-term interest remaining unpaid at calc date
        cp = ws.cell(row=r, column=16)
        cp.value = (f'=IF(C{r}="","",IF({cfg_date}<=G{r},0,'
                    f'MAX(0,L{r}-IF(H{r}<>"",MAX(0,I{r}-D{r}-O{r}),0))))')
        cp.number_format = FMT_VND
        cp.fill = make_fill(C_OVERDUE_BG)
        cp.font = make_font(size=9, color="C00000")
        cp.alignment = Alignment(horizontal="right", vertical="center")
        cp.border = thin_border

        # Q: Late payment interest on overdue in-term interest during overdue period
        cq = ws.cell(row=r, column=17)
        cq.value = f'=IF(C{r}="","",IF(N{r}>0,L{r}*{cfg_ct}*N{r}/({cfg_dc}),0))'
        cq.number_format = FMT_VND
        cq.fill = make_fill(C_OVERDUE_BG)
        cq.font = make_font(size=9, color="C00000")
        cq.alignment = Alignment(horizontal="right", vertical="center")
        cq.border = thin_border

        # R: Total outstanding debt at calc date
        cr = ws.cell(row=r, column=18)
        cr.value = (f'=IF(C{r}="","",IF({cfg_date}<=G{r},'
                    f'MAX(0,D{r}+L{r}-IF(H{r}<>"",I{r},0)),'
                    f'MAX(0,D{r}+L{r}+O{r}+Q{r}-IF(H{r}<>"",I{r},0))))')
        cr.number_format = FMT_VND
        cr.fill = make_fill("C00000")
        cr.font = make_font(bold=True, size=9, color=C_HEADER_TXT)
        cr.alignment = Alignment(horizontal="right", vertical="center")
        cr.border = thin_border

    # Blank input rows
    n_sample = len(sample_loans)
    for i in range(30):
        r = FIRST_DATA_ROW + n_sample + i
        ws.row_dimensions[r].height = 16
        alt_bg = C_ALT_ROW if i % 2 == 1 else C_WHITE

        c = ws.cell(row=r, column=1, value=n_sample + i + 1)
        c.font = make_font(size=9); c.fill = make_fill(alt_bg)
        c.alignment = Alignment(horizontal="center"); c.border = thin_border

        style_input(ws, r, 2)
        style_input(ws, r, 3, fmt=FMT_DATE)
        style_input(ws, r, 4, fmt=FMT_VND)
        style_input(ws, r, 5, fmt=FMT_PCT)
        style_input(ws, r, 6, fmt=FMT_DATE)

        # G: adjusted due date
        cg = ws.cell(row=r, column=7)
        cg.value = (f'=IF(F{r}="","",IF(AND(WEEKDAY(F{r},2)<=5,'
                    f'COUNTIF({hol_ref},F{r})=0),F{r},'
                    f'WORKDAY.INTL(F{r},1,1,{hol_ref})))')
        cg.number_format = FMT_DATE
        cg.fill = make_fill(C_CALC_BG); cg.font = make_font(size=9)
        cg.alignment = Alignment(horizontal="center"); cg.border = thin_border

        style_input(ws, r, 8, fmt=FMT_DATE)
        style_input(ws, r, 9, fmt=FMT_VND)

        # Formulas for blank rows (same as above)
        for col, formula, fmt, bg, color in [
            (10, (f'=IF(C{r}="","",IF(AND(H{r}<>"",I{r}>=(D{r}+L{r}+O{r}+Q{r})),"✅ Đã tất toán",'
                  f'IF(AND(H{r}<>"",I{r}>0),"⚠️ Đã trả một phần",'
                  f'IF({cfg_date}<G{r},"⏳ Trong hạn","⚠ Quá hạn"))))'), None, C_CALC_BG, "000000"),
            (11, f'=IF(C{r}="","",MAX(0,MIN(G{r},IF(H{r}<>"",H{r},{cfg_date}))-C{r}))', FMT_INT, C_CALC_BG, "000000"),
            (12, f'=IF(C{r}="","",D{r}*E{r}*K{r}/({cfg_dc}))', FMT_VND, C_CALC_BG, "000000"),
            (13, (f'=IF(C{r}="","",IF({cfg_date}<=G{r},0,'
                  f'MAX(0,D{r}-IF(H{r}<>"",I{r},0))))'),
             FMT_VND, C_OVERDUE_BG, "C00000"),
            (14, f'=IF(C{r}="","",MAX(0,MIN(IF(H{r}<>"",H{r},{cfg_date}),{cfg_date})-G{r}))',
             FMT_INT, C_OVERDUE_BG, "C00000"),
            (15, f'=IF(C{r}="","",IF(N{r}>0,D{r}*(E{r}*{cfg_qh})*N{r}/({cfg_dc}),0))', FMT_VND, C_OVERDUE_BG, "C00000"),
            (16, (f'=IF(C{r}="","",IF({cfg_date}<=G{r},0,'
                  f'MAX(0,L{r}-IF(H{r}<>"",MAX(0,I{r}-D{r}-O{r}),0))))'), FMT_VND, C_OVERDUE_BG, "C00000"),
            (17, f'=IF(C{r}="","",IF(N{r}>0,L{r}*{cfg_ct}*N{r}/({cfg_dc}),0))', FMT_VND, C_OVERDUE_BG, "C00000"),
            (18, (f'=IF(C{r}="","",IF({cfg_date}<=G{r},'
                  f'MAX(0,D{r}+L{r}-IF(H{r}<>"",I{r},0)),'
                  f'MAX(0,D{r}+L{r}+O{r}+Q{r}-IF(H{r}<>"",I{r},0))))'),
             FMT_VND, "C00000", C_HEADER_TXT),
        ]:
            cx = ws.cell(row=r, column=col)
            cx.value = formula
            if fmt: cx.number_format = fmt
            cx.fill = make_fill(bg)
            cx.font = make_font(size=9, color=color,
                                bold=(col == 18))
            cx.alignment = Alignment(
                horizontal="center" if col == 10 else "right",
                vertical="center"
            )
            cx.border = thin_border

    # Total row
    last_data = FIRST_DATA_ROW + n_sample + 30
    total_row = last_data + 1
    ws.row_dimensions[total_row].height = 22

    style_total(ws, total_row, 1, "TỔNG CỘNG", colspan=3)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="left")
    style_total(ws, total_row, 4, f"=SUMIF(C{FIRST_DATA_ROW}:C{last_data},\"<>\"\"\",D{FIRST_DATA_ROW}:D{last_data})", FMT_VND)
    style_total(ws, total_row, 5, "")
    style_total(ws, total_row, 6, "")
    style_total(ws, total_row, 7, "")
    style_total(ws, total_row, 8, "")
    style_total(ws, total_row, 9, f"=SUMIF(C{FIRST_DATA_ROW}:C{last_data},\"<>\"\"\",I{FIRST_DATA_ROW}:I{last_data})", FMT_VND)
    style_total(ws, total_row, 10, "")
    style_total(ws, total_row, 11, "")
    style_total(ws, total_row, 12, f"=SUM(L{FIRST_DATA_ROW}:L{last_data})", FMT_VND)
    style_total(ws, total_row, 13, f"=SUM(M{FIRST_DATA_ROW}:M{last_data})", FMT_VND)
    style_total(ws, total_row, 14, "")
    style_total(ws, total_row, 15, f"=SUM(O{FIRST_DATA_ROW}:O{last_data})", FMT_VND)
    style_total(ws, total_row, 16, f"=SUM(P{FIRST_DATA_ROW}:P{last_data})", FMT_VND)
    style_total(ws, total_row, 17, f"=SUM(Q{FIRST_DATA_ROW}:Q{last_data})", FMT_VND)

    # Grand total outstanding
    c_grand = ws.cell(row=total_row, column=18)
    c_grand.value = f"=SUM(R{FIRST_DATA_ROW}:R{last_data})"
    c_grand.number_format = FMT_VND
    c_grand.fill = make_fill("C00000")
    c_grand.font = make_font(bold=True, size=11, color=C_HEADER_TXT)
    c_grand.alignment = Alignment(horizontal="right", vertical="center")
    c_grand.border = thick_border

    set_print_a4_landscape(ws, "BẢNG TÍNH LÃI — KHOẢN VAY BULLET")
    ws.print_title_rows = "1:3"


def build_tonghop(ws):
    ws.title = "TỔNG HỢP"
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 20

    ws.merge_cells("A1:C1")
    c = ws["A1"]
    c.value = "📑 TỔNG HỢP DANH MỤC KHOẢN VAY — TẠI NGÀY TÍNH TOÁN"
    c.font = make_font(bold=True, size=14, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:C2")
    c = ws["A2"]
    c.value = f"Ngày tính: ='CẤU HÌNH'!B4   |   Căn cứ: TT39/2016/TT-NHNN"
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 14

    def row_item(ws, row, label, formula, fmt=FMT_VND, bold=False, bg=C_WHITE):
        ws.row_dimensions[row].height = 18
        c1 = ws.cell(row=row, column=1, value=label)
        c1.font = make_font(bold=bold, size=10)
        c1.fill = make_fill(bg)
        c1.border = thin_border
        c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c2 = ws.cell(row=row, column=2, value=formula)
        c2.font = make_font(bold=bold, size=10)
        c2.fill = make_fill(bg)
        c2.number_format = fmt
        c2.border = thin_border
        c2.alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(row=row, column=3).fill = make_fill(bg)
        ws.cell(row=row, column=3).border = thin_border

    # last row of BẢNG TÍNH
    bt_first = 4
    bt_last  = 4 + 5 + 30  # sample + blanks

    def sec(ws, row, title):
        ws.row_dimensions[row].height = 18
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c = ws.cell(row=row, column=1, value=f"▶ {title}")
        c.font = make_font(bold=True, size=10)
        c.fill = make_fill(C_SECTION_BG)
        c.alignment = Alignment(horizontal="left", vertical="center")

    r = 3
    ws.row_dimensions[r].height = 8

    r += 1; sec(ws, r, "TỔNG QUAN DANH MỤC")
    r += 1; row_item(ws, r, "Ngày tính", "='CẤU HÌNH'!B4", FMT_DATE)
    r += 1; row_item(ws, r, "Số khoản vay",
                     f"=COUNTIF('BẢNG TÍNH'!C{bt_first}:C{bt_last},\"<>\"\"\")", FMT_INT)
    r += 1; row_item(ws, r, "Tổng giải ngân",
                     f"=SUMIF('BẢNG TÍNH'!C{bt_first}:C{bt_last},\"<>\"\"\",'BẢNG TÍNH'!D{bt_first}:D{bt_last})", FMT_VND)

    r += 1; ws.row_dimensions[r].height = 6
    r += 1; sec(ws, r, "PHÂN LOẠI THEO TRẠNG THÁI")
    r += 1; row_item(ws, r, "Khoản trong hạn (số lượng)",
                     f"=COUNTIF('BẢNG TÍNH'!J{bt_first}:J{bt_last},\"*Trong hạn*\")", FMT_INT)
    r += 1; row_item(ws, r, "Khoản quá hạn (số lượng)",
                     f"=COUNTIF('BẢNG TÍNH'!J{bt_first}:J{bt_last},\"*Quá hạn*\")", FMT_INT)
    r += 1; row_item(ws, r, "Khoản đã trả (số lượng)",
                     f"=COUNTIF('BẢNG TÍNH'!J{bt_first}:J{bt_last},\"*Đã trả*\")", FMT_INT)

    r += 1; ws.row_dimensions[r].height = 6
    r += 1; sec(ws, r, "DƯ NỢ & LÃI PHẠT TẠI NGÀY TÍNH (Chuẩn Core Banking)")
    r += 1; row_item(ws, r, "1. Tổng gốc quá hạn còn nợ",
                     f"=SUM('BẢNG TÍNH'!M{bt_first}:M{bt_last})", FMT_VND, bg=C_OVERDUE_BG)
    r += 1; row_item(ws, r, "2. Tổng lãi cơ bản (100% LS trong hạn)",
                     f"=SUMIFS('BẢNG TÍNH'!L{bt_first}:L{bt_last},'BẢNG TÍNH'!J{bt_first}:J{bt_last},\"*Trong hạn*\")+SUM('BẢNG TÍNH'!P{bt_first}:P{bt_last})+(SUM('BẢNG TÍNH'!O{bt_first}:O{bt_last})/'CẤU HÌNH'!$B$6)", FMT_VND, bg="EBF3F9")
    r += 1; row_item(ws, r, "3. Tổng phạt quá hạn tăng thêm (50% chênh lệch)",
                     f"=SUM('BẢNG TÍNH'!O{bt_first}:O{bt_last})-(SUM('BẢNG TÍNH'!O{bt_first}:O{bt_last})/'CẤU HÌNH'!$B$6)", FMT_VND, bg=C_OVERDUE_BG)
    r += 1; row_item(ws, r, "4. Tổng phạt chậm trả lãi (≤ 10%/năm)",
                     f"=SUM('BẢNG TÍNH'!Q{bt_first}:Q{bt_last})", FMT_VND, bg=C_OVERDUE_BG)
    r += 1; row_item(ws, r, "▶ TỔNG CÁC KHOẢN PHẠT VI PHẠM (Phạt QH 50% + Chậm trả 10%)",
                     f"=(SUM('BẢNG TÍNH'!O{bt_first}:O{bt_last})-(SUM('BẢNG TÍNH'!O{bt_first}:O{bt_last})/'CẤU HÌNH'!$B$6))+SUM('BẢNG TÍNH'!Q{bt_first}:Q{bt_last})", FMT_VND, bold=True, bg=C_WARN_BG)

    r += 1; ws.row_dimensions[r].height = 6
    # Grand total
    r += 1; ws.row_dimensions[r].height = 26
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(row=r, column=1, value="★ TỔNG SỐ TIỀN CÒN NỢ TOÀN DANH MỤC ★")
    c.font = make_font(bold=True, size=12, color=C_HEADER_TXT)
    c.fill = make_fill("C00000")
    c.border = thick_border
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.unmerge_cells(start_row=r, start_column=1, end_row=r, end_column=3)

    ws.cell(row=r, column=1, value="★ TỔNG SỐ TIỀN CÒN NỢ TOÀN DANH MỤC ★")
    ws.cell(row=r, column=1).font = make_font(bold=True, size=11, color=C_HEADER_TXT)
    ws.cell(row=r, column=1).fill = make_fill("C00000")
    ws.cell(row=r, column=1).border = thick_border
    ws.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center")
    c2 = ws.cell(row=r, column=2,
                 value=f"=SUM('BẢNG TÍNH'!R{bt_first}:R{bt_last})")
    c2.number_format = FMT_VND
    c2.font = make_font(bold=True, size=12, color=C_HEADER_TXT)
    c2.fill = make_fill("C00000")
    c2.border = thick_border
    c2.alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(row=r, column=3).fill = make_fill("C00000")
    ws.cell(row=r, column=3).border = thick_border

    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(row=r, column=1,
                value=("Ghi chú Core Banking & Pháp lý: LS quá hạn = 150% LS trong hạn (gồm 100% lãi cơ bản + 50% phạt quá hạn tăng thêm theo Đ13 TT39/2016). "
                       "Lãi chậm trả ≤ 10%/năm (Đ13 TT39/2016)."))
    c.font = make_font(size=8, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.border = thin_border
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 28

    set_print_a4_landscape(ws, "TỔNG HỢP DANH MỤC KHOẢN VAY")


def build_file2():
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    print("Building CẤU HÌNH...")
    ws_cfg = wb.create_sheet()
    build_cauhinh(ws_cfg)

    print("Building NGÀY NGHỈ LỄ...")
    ws_hol = wb.create_sheet()
    build_holidays(ws_hol)

    print("Building BẢNG TÍNH...")
    ws_bt = wb.create_sheet()
    build_bangtính(ws_bt)

    print("Building TỔNG HỢP...")
    ws_th = wb.create_sheet()
    build_tonghop(ws_th)

    wb.active = ws_th

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True

    out_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "TinhLai_DonGian.xlsx"
    )
    wb.save(out_path)
    print(f"\n✅ Saved: {os.path.abspath(out_path)}")
    return out_path


if __name__ == "__main__":
    build_file2()
