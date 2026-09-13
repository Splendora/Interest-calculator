"""
Build TinhLai_ChiTiet.xlsx — Vietnamese Loan Interest Calculator (Main Template)
Excel 2016 compatible, formula-based (no VBA)
"""
import os
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.header_footer import HeaderFooter
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.worksheet.table import Table, TableStyleInfo

# ─────────────────────────────────────────────
# STYLES
# ─────────────────────────────────────────────

# Colors
C_HEADER_BG   = "1F4E79"   # Dark blue — header
C_HEADER_TXT  = "FFFFFF"   # White
C_INPUT_BG    = "DAEEF3"   # Light blue — input cells
C_CALC_BG     = "F2F2F2"   # Light grey — calculated cells
C_WARN_BG     = "FFF2CC"   # Light yellow — warning/important
C_OVERDUE_BG  = "FCE4D6"   # Light orange — overdue cells
C_TOTAL_BG    = "D6E4BC"   # Light green — totals
C_WHITE       = "FFFFFF"
C_SECTION_BG  = "BDD7EE"   # Medium blue — section headers
C_ALT_ROW     = "EBF3F9"   # Very light blue — alternate rows

def make_font(bold=False, size=10, color="000000", name="Calibri"):
    return Font(bold=bold, size=size, color=color, name=name)

def make_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def make_border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def make_border_bottom(style="medium"):
    return Border(bottom=Side(style=style))

thin_border  = make_border("thin")
thick_border = make_border("medium")

def style_header(ws, row, col, value, colspan=1, bg=C_HEADER_BG):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = make_font(bold=True, size=10, color=C_HEADER_TXT)
    cell.fill = make_fill(bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border
    if colspan > 1:
        ws.merge_cells(
            start_row=row, start_column=col,
            end_row=row, end_column=col + colspan - 1
        )
    return cell

def style_input(ws, row, col, value=None, fmt=None, align="left"):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = make_fill(C_INPUT_BG)
    cell.font = make_font(size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = thin_border
    if fmt:
        cell.number_format = fmt
    return cell

def style_calc(ws, row, col, value=None, fmt=None, align="right", bg=C_CALC_BG):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = make_fill(bg)
    cell.font = make_font(size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = thin_border
    if fmt:
        cell.number_format = fmt
    return cell

def style_label(ws, row, col, value, bold=False, indent=0, bg=C_WHITE):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = make_fill(bg)
    cell.font = make_font(bold=bold, size=10)
    cell.alignment = Alignment(horizontal="left", vertical="center",
                                indent=indent)
    cell.border = thin_border
    return cell

def style_total(ws, row, col, value=None, fmt=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = make_fill(C_TOTAL_BG)
    cell.font = make_font(bold=True, size=10)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.border = thick_border
    if fmt:
        cell.number_format = fmt
    return cell

def style_section(ws, row, col, value, colspan=1):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = make_fill(C_SECTION_BG)
    cell.font = make_font(bold=True, size=10)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = thin_border
    if colspan > 1:
        ws.merge_cells(
            start_row=row, start_column=col,
            end_row=row, end_column=col + colspan - 1
        )
    return cell

FMT_VND      = '#,##0'
FMT_VND_RED  = '#,##0;[Red]-#,##0'
FMT_PCT      = '0.00%'
FMT_PCT4     = '0.0000%'
FMT_DATE     = 'DD/MM/YYYY'
FMT_INT      = '#,##0'
FMT_BOOL     = '"Có";"Có";"Không"'

def set_print_a4_landscape(ws, title_text=""):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 9        # A4
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(
        left=0.5, right=0.5, top=0.75, bottom=0.75,
        header=0.3, footer=0.3
    )
    ws.oddHeader.center.text = title_text
    ws.oddFooter.left.text   = "Trang &P / &N"
    ws.oddFooter.right.text  = "In ngày: &D"

# ─────────────────────────────────────────────
# HELPER: Vietnamese holiday data 2020-2030
# ─────────────────────────────────────────────
def get_holidays():
    """
    Returns list of (date, name, type) for Vietnamese public holidays 2020-2030.
    Tet Lunar dates converted to Gregorian. User should verify and update annually.
    Based on official Govt announcements; Lunar→Solar conversion included.
    """
    holidays = [
        # ── 2020 ──
        (date(2020, 1,  1), "Tết Dương lịch",                "Lễ"),
        (date(2020, 1, 23), "Tết Nguyên đán (29 tháng Chạp)","Tết"),
        (date(2020, 1, 24), "Tết Nguyên đán (30 tháng Chạp)","Tết"),
        (date(2020, 1, 25), "Tết Nguyên đán (Mùng 1)",       "Tết"),
        (date(2020, 1, 26), "Tết Nguyên đán (Mùng 2)",       "Tết"),
        (date(2020, 1, 27), "Tết Nguyên đán (Mùng 3)",       "Tết"),
        (date(2020, 4, 22), "Giỗ Tổ Hùng Vương (10/3 ÂL)",  "Lễ"),
        (date(2020, 4, 30), "Ngày Chiến thắng 30/4",         "Lễ"),
        (date(2020, 5,  1), "Ngày Quốc tế Lao động 1/5",     "Lễ"),
        (date(2020, 9,  2), "Ngày Quốc khánh 2/9",           "Lễ"),
        (date(2020, 9,  3), "Nghỉ bù Quốc khánh",            "Nghỉ bù"),
        # ── 2021 ──
        (date(2021, 1,  1), "Tết Dương lịch",                "Lễ"),
        (date(2021, 2, 10), "Tết Nguyên đán (29 tháng Chạp)","Tết"),
        (date(2021, 2, 11), "Tết Nguyên đán (30 tháng Chạp)","Tết"),
        (date(2021, 2, 12), "Tết Nguyên đán (Mùng 1)",       "Tết"),
        (date(2021, 2, 13), "Tết Nguyên đán (Mùng 2)",       "Tết"),
        (date(2021, 2, 14), "Tết Nguyên đán (Mùng 3)",       "Tết"),
        (date(2021, 4, 21), "Giỗ Tổ Hùng Vương (10/3 ÂL)",  "Lễ"),
        (date(2021, 4, 30), "Ngày Chiến thắng 30/4",         "Lễ"),
        (date(2021, 5,  3), "Nghỉ bù 1/5 (rơi CN)",         "Nghỉ bù"),
        (date(2021, 9,  2), "Ngày Quốc khánh 2/9",           "Lễ"),
        (date(2021, 9,  3), "Nghỉ bù Quốc khánh",            "Nghỉ bù"),
        # ── 2022 ──
        (date(2022, 1,  3), "Nghỉ bù Tết Dương lịch",        "Nghỉ bù"),
        (date(2022, 1, 29), "Tết Nguyên đán (28 tháng Chạp)","Tết"),
        (date(2022, 1, 30), "Tết Nguyên đán (29 tháng Chạp)","Tết"),
        (date(2022, 1, 31), "Tết Nguyên đán (30 tháng Chạp)","Tết"),
        (date(2022, 2,  1), "Tết Nguyên đán (Mùng 1)",       "Tết"),
        (date(2022, 2,  2), "Tết Nguyên đán (Mùng 2)",       "Tết"),
        (date(2022, 4,  3), "Nghỉ bù Giỗ Tổ Hùng Vương",    "Nghỉ bù"),
        (date(2022, 4, 10), "Giỗ Tổ Hùng Vương (10/3 ÂL)",  "Lễ"),
        (date(2022, 4, 30), "Ngày Chiến thắng 30/4",         "Lễ"),
        (date(2022, 5,  2), "Nghỉ bù 1/5 (rơi CN)",         "Nghỉ bù"),
        (date(2022, 5,  3), "Nghỉ bù 30/4 (rơi CN)",        "Nghỉ bù"),
        (date(2022, 9,  1), "Nghỉ bù Quốc khánh",            "Nghỉ bù"),
        (date(2022, 9,  2), "Ngày Quốc khánh 2/9",           "Lễ"),
        # ── 2023 ──
        (date(2023, 1,  1), "Tết Dương lịch",                "Lễ"),
        (date(2023, 1, 20), "Tết Nguyên đán (29 tháng Chạp)","Tết"),
        (date(2023, 1, 21), "Tết Nguyên đán (30 tháng Chạp)","Tết"),
        (date(2023, 1, 22), "Tết Nguyên đán (Mùng 1)",       "Tết"),
        (date(2023, 1, 23), "Tết Nguyên đán (Mùng 2)",       "Tết"),
        (date(2023, 1, 24), "Tết Nguyên đán (Mùng 3)",       "Tết"),
        (date(2023, 4, 29), "Nghỉ bù Giỗ Tổ Hùng Vương",    "Nghỉ bù"),
        (date(2023, 4, 30), "Ngày Chiến thắng 30/4",         "Lễ"),
        (date(2023, 5,  1), "Ngày Quốc tế Lao động 1/5",     "Lễ"),
        (date(2023, 9,  1), "Nghỉ bù Quốc khánh",            "Nghỉ bù"),
        (date(2023, 9,  2), "Ngày Quốc khánh 2/9",           "Lễ"),
        # ── 2024 ──
        (date(2024, 1,  1), "Tết Dương lịch",                "Lễ"),
        (date(2024, 2,  8), "Tết Nguyên đán (28 tháng Chạp)","Tết"),
        (date(2024, 2,  9), "Tết Nguyên đán (29 tháng Chạp)","Tết"),
        (date(2024, 2, 10), "Tết Nguyên đán (Mùng 1)",       "Tết"),
        (date(2024, 2, 12), "Tết Nguyên đán (Mùng 2)",       "Tết"),
        (date(2024, 2, 13), "Tết Nguyên đán (Mùng 3)",       "Tết"),
        (date(2024, 4, 18), "Giỗ Tổ Hùng Vương (10/3 ÂL)",  "Lễ"),
        (date(2024, 4, 26), "Nghỉ bù 30/4",                  "Nghỉ bù"),
        (date(2024, 4, 29), "Nghỉ bù 30/4",                  "Nghỉ bù"),
        (date(2024, 4, 30), "Ngày Chiến thắng 30/4",         "Lễ"),
        (date(2024, 5,  1), "Ngày Quốc tế Lao động 1/5",     "Lễ"),
        (date(2024, 9,  2), "Ngày Quốc khánh 2/9",           "Lễ"),
        (date(2024, 9,  3), "Nghỉ bù Quốc khánh",            "Nghỉ bù"),
        # ── 2025 ──
        (date(2025, 1,  1), "Tết Dương lịch",                "Lễ"),
        (date(2025, 1, 25), "Tết Nguyên đán (26 tháng Chạp)","Tết"),
        (date(2025, 1, 27), "Tết Nguyên đán (28 tháng Chạp)","Tết"),
        (date(2025, 1, 28), "Tết Nguyên đán (29 tháng Chạp)","Tết"),
        (date(2025, 1, 29), "Tết Nguyên đán (30 tháng Chạp)","Tết"),
        (date(2025, 1, 30), "Tết Nguyên đán (Mùng 1)",       "Tết"),
        (date(2025, 1, 31), "Tết Nguyên đán (Mùng 2)",       "Tết"),
        (date(2025, 2,  3), "Nghỉ bù Tết",                   "Nghỉ bù"),
        (date(2025, 4,  6), "Nghỉ bù Giỗ Tổ Hùng Vương",    "Nghỉ bù"),
        (date(2025, 4,  7), "Giỗ Tổ Hùng Vương (10/3 ÂL)",  "Lễ"),
        (date(2025, 4, 30), "Ngày Chiến thắng 30/4",         "Lễ"),
        (date(2025, 5,  1), "Ngày Quốc tế Lao động 1/5",     "Lễ"),
        (date(2025, 5,  2), "Nghỉ bù 30/4",                  "Nghỉ bù"),
        (date(2025, 9,  1), "Nghỉ bù Quốc khánh",            "Nghỉ bù"),
        (date(2025, 9,  2), "Ngày Quốc khánh 2/9",           "Lễ"),
        # ── 2026 ──
        (date(2026, 1,  1), "Tết Dương lịch",                "Lễ"),
        (date(2026, 1, 16), "Tết Nguyên đán (28 tháng Chạp)","Tết"),
        (date(2026, 1, 17), "Tết Nguyên đán (29 tháng Chạp)","Tết"),
        (date(2026, 1, 18), "Tết Nguyên đán (30 tháng Chạp)","Tết"),
        (date(2026, 1, 19), "Tết Nguyên đán (Mùng 1)",       "Tết"),
        (date(2026, 1, 20), "Tết Nguyên đán (Mùng 2)",       "Tết"),
        (date(2026, 3, 27), "Giỗ Tổ Hùng Vương (10/3 ÂL)",  "Lễ"),
        (date(2026, 4, 30), "Ngày Chiến thắng 30/4",         "Lễ"),
        (date(2026, 5,  1), "Ngày Quốc tế Lao động 1/5",     "Lễ"),
        (date(2026, 9,  2), "Ngày Quốc khánh 2/9",           "Lễ"),
        (date(2026, 9,  3), "Nghỉ bù Quốc khánh",            "Nghỉ bù"),
        # ── 2027 ──
        (date(2027, 1,  1), "Tết Dương lịch",                "Lễ"),
        (date(2027, 1,  6), "Tết Nguyên đán (29 tháng Chạp)","Tết"),
        (date(2027, 1,  7), "Tết Nguyên đán (30 tháng Chạp)","Tết"),
        (date(2027, 1,  8), "Tết Nguyên đán (Mùng 1)",       "Tết"),
        (date(2027, 1,  9), "Tết Nguyên đán (Mùng 2)",       "Tết"),
        (date(2027, 1, 10), "Tết Nguyên đán (Mùng 3)",       "Tết"),
        (date(2027, 4, 16), "Giỗ Tổ Hùng Vương (10/3 ÂL)",  "Lễ"),
        (date(2027, 4, 30), "Ngày Chiến thắng 30/4",         "Lễ"),
        (date(2027, 5,  3), "Nghỉ bù 1/5",                   "Nghỉ bù"),
        (date(2027, 9,  2), "Ngày Quốc khánh 2/9",           "Lễ"),
        (date(2027, 9,  3), "Nghỉ bù Quốc khánh",            "Nghỉ bù"),
        # ── 2028 ──
        (date(2028, 1,  1), "Tết Dương lịch",                "Lễ"),
        (date(2028, 1, 26), "Tết Nguyên đán (29 tháng Chạp)","Tết"),
        (date(2028, 1, 27), "Tết Nguyên đán (30 tháng Chạp)","Tết"),
        (date(2028, 1, 28), "Tết Nguyên đán (Mùng 1)",       "Tết"),
        (date(2028, 1, 29), "Tết Nguyên đán (Mùng 2)",       "Tết"),
        (date(2028, 1, 30), "Tết Nguyên đán (Mùng 3)",       "Tết"),
        (date(2028, 5,  5), "Giỗ Tổ Hùng Vương (10/3 ÂL)",  "Lễ"),
        (date(2028, 4, 30), "Ngày Chiến thắng 30/4",         "Lễ"),
        (date(2028, 5,  1), "Ngày Quốc tế Lao động 1/5",     "Lễ"),
        (date(2028, 9,  2), "Ngày Quốc khánh 2/9",           "Lễ"),
        (date(2028, 9,  4), "Nghỉ bù Quốc khánh",            "Nghỉ bù"),
        # ── 2029 ──
        (date(2029, 1,  1), "Tết Dương lịch",                "Lễ"),
        (date(2029, 2, 12), "Tết Nguyên đán (29 tháng Chạp)","Tết"),
        (date(2029, 2, 13), "Tết Nguyên đán (30 tháng Chạp)","Tết"),
        (date(2029, 2, 14), "Tết Nguyên đán (Mùng 1)",       "Tết"),
        (date(2029, 2, 15), "Tết Nguyên đán (Mùng 2)",       "Tết"),
        (date(2029, 2, 16), "Tết Nguyên đán (Mùng 3)",       "Tết"),
        (date(2029, 4, 24), "Giỗ Tổ Hùng Vương (10/3 ÂL)",  "Lễ"),
        (date(2029, 4, 30), "Ngày Chiến thắng 30/4",         "Lễ"),
        (date(2029, 5,  1), "Ngày Quốc tế Lao động 1/5",     "Lễ"),
        (date(2029, 9,  2), "Ngày Quốc khánh 2/9",           "Lễ"),
        (date(2029, 9,  3), "Nghỉ bù Quốc khánh",            "Nghỉ bù"),
        # ── 2030 ──
        (date(2030, 1,  1), "Tết Dương lịch",                "Lễ"),
        (date(2030, 2,  2), "Tết Nguyên đán (29 tháng Chạp)","Tết"),
        (date(2030, 2,  3), "Tết Nguyên đán (30 tháng Chạp)","Tết"),
        (date(2030, 2,  4), "Tết Nguyên đán (Mùng 1)",       "Tết"),
        (date(2030, 2,  5), "Tết Nguyên đán (Mùng 2)",       "Tết"),
        (date(2030, 2,  6), "Tết Nguyên đán (Mùng 3)",       "Tết"),
        (date(2030, 4, 13), "Giỗ Tổ Hùng Vương (10/3 ÂL)",  "Lễ"),
        (date(2030, 4, 30), "Ngày Chiến thắng 30/4",         "Lễ"),
        (date(2030, 5,  1), "Ngày Quốc tế Lao động 1/5",     "Lễ"),
        (date(2030, 9,  2), "Ngày Quốc khánh 2/9",           "Lễ"),
        (date(2030, 9,  3), "Nghỉ bù Quốc khánh",            "Nghỉ bù"),
    ]
    return sorted(holidays, key=lambda x: x[0])


# ─────────────────────────────────────────────
# SHEET BUILDERS
# ─────────────────────────────────────────────

def build_sheet_cauhinh(ws):
    ws.title = "CẤU HÌNH"
    ws.sheet_view.showGridLines = True
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 12

    # Title
    title = ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value = "⚙ CẤU HÌNH KHOẢN VAY"
    c.font = make_font(bold=True, size=14, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    subtitle = ws.merge_cells("A2:D2")
    c = ws["A2"]
    c.value = "Căn cứ: Thông tư 39/2016/TT-NHNN (sửa đổi bởi TT 06/2023/TT-NHNN) và Bộ luật Dân sự 2015"
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    ws.row_dimensions[3].height = 8

    def section_row(r, title):
        ws.row_dimensions[r].height = 18
        style_section(ws, r, 1, f"▶ {title}", colspan=4)

    def param_row(r, label, note=""):
        ws.row_dimensions[r].height = 20
        style_label(ws, r, 1, label, bold=False)
        style_input(ws, r, 2, value=None)
        c = ws.cell(row=r, column=3, value=note)
        c.fill = make_fill(C_WHITE)
        c.font = make_font(size=9, color="595959")
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c.border = thin_border
        # merge D
        ws.cell(row=r, column=4).border = thin_border

    # ── Section 1: Thông tin khoản vay ──
    section_row(4, "THÔNG TIN KHOẢN VAY")

    param_row(5,  "Ngày tính toán (ngày tòa / ngày cắt)",
              "Ngày tòa xét xử hoặc ngày cần tính số dư nợ")
    ws["B5"].number_format = FMT_DATE
    ws["B5"].value = date(2025, 6, 15)

    param_row(6,  "Quy ước đếm ngày (Day count convention)",
              "Actual/365 là chuẩn phổ biến tại VN")
    ws["B6"].value = "Actual/365"

    param_row(7,  "Phương thức trả nợ",
              "Dư nợ giảm dần: gốc đều, lãi giảm | Annuity: tổng đều | Bullet: trả cuối kỳ")
    ws["B7"].value = "Tùy chỉnh (nhập lịch trả nợ)"

    param_row(8,  "Tần suất trả nợ",
              "Kỳ hạn trả gốc và lãi")
    ws["B8"].value = "Hàng tháng"

    ws.row_dimensions[9].height = 8

    # ── Section 2: Lãi suất ──
    section_row(10, "LÃI SUẤT VÀ PHẠT")

    param_row(11, "Hệ số nhân lãi suất quá hạn",
              "Tối đa 150% theo Điều 13 TT39. Nhập dạng thập phân, VD: 1.5")
    ws["B11"].value = 1.5
    ws["B11"].number_format = '0.00"×"'

    param_row(12, "Lãi suất phạt trên lãi chậm trả (%/năm)",
              "Tối đa 10%/năm theo Điều 13 TT39")
    ws["B12"].value = 0.10
    ws["B12"].number_format = FMT_PCT

    param_row(13, "Phương thức lãi suất",
              "Cố định: nhập LS trong bảng Giải Ngân | Thả nổi: nhập LS gốc + biên độ trong bảng Thay đổi LS")
    ws["B13"].value = "Cố định"

    ws.row_dimensions[14].height = 8

    # ── Section 3: Thứ tự thu nợ ──
    section_row(15, "THỨ TỰ THU NỢ (khi có nợ quá hạn)")

    param_row(16, "Preset thứ tự thu nợ",
              "TT39/TT06: Gốc QH → Lãi trên Gốc QH → Gốc đến hạn → Lãi đến hạn → Lãi chậm trả")
    ws["B16"].value = "TT39/TT06 (mặc định)"

    ws.row_dimensions[17].height = 8

    # ── Section 4: Hướng dẫn ──
    section_row(18, "HƯỚNG DẪN SỬ DỤNG")

    instructions = [
        ("Bước 1", "Nhập cấu hình khoản vay trên sheet này"),
        ("Bước 2", "Cập nhật ngày nghỉ lễ trên sheet NGÀY NGHỈ LỄ (nếu cần)"),
        ("Bước 3", "Nhập các đợt giải ngân trên sheet GIẢI NGÂN"),
        ("Bước 4", "Nhập lịch trả nợ theo hợp đồng trên sheet LỊCH TRẢ NỢ"),
        ("Bước 5", "Nhập lịch thay đổi lãi suất (nếu có) trên sheet THAY ĐỔI LÃI SUẤT"),
        ("Bước 6", "Nhập các khoản thanh toán thực tế trên sheet THANH TOÁN THỰC TẾ"),
        ("Bước 7", "Mở sheet BẢNG TÍNH — kéo dài các hàng công thức đến ngày tính toán"),
        ("Bước 8", "Nhập các mốc ngày kỳ trên sheet TỔNG HỢP KỲ để xem bảng gộp"),
        ("Lưu ý",  "Ô nền XANH NHẠT = nhập liệu. Ô nền XÁM = công thức, không chỉnh sửa"),
    ]

    for i, (step, desc) in enumerate(instructions):
        r = 19 + i
        ws.row_dimensions[r].height = 16
        c1 = ws.cell(row=r, column=1, value=step)
        c1.font = make_font(bold=True, size=9)
        c1.fill = make_fill(C_SECTION_BG)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = thin_border
        c2 = ws.cell(row=r, column=2, value=desc)
        c2.font = make_font(size=9)
        c2.fill = make_fill(C_WHITE)
        c2.alignment = Alignment(horizontal="left", vertical="center")
        c2.border = thin_border
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)

    # Data Validations
    dv_daycount = DataValidation(
        type="list",
        formula1='"Actual/365,Actual/360,30/360"',
        allow_blank=False,
        showDropDown=False
    )
    dv_daycount.sqref = "B6"
    ws.add_data_validation(dv_daycount)

    dv_method = DataValidation(
        type="list",
        formula1='"Dư nợ giảm dần,Annuity (trả đều),Bullet (trả cuối kỳ),Tùy chỉnh (nhập lịch trả nợ)"',
        allow_blank=False,
        showDropDown=False
    )
    dv_method.sqref = "B7"
    ws.add_data_validation(dv_method)

    dv_freq = DataValidation(
        type="list",
        formula1='"Hàng tháng,Hàng quý,6 tháng,Hàng năm,Tùy chỉnh"',
        allow_blank=False,
        showDropDown=False
    )
    dv_freq.sqref = "B8"
    ws.add_data_validation(dv_freq)

    dv_ls_method = DataValidation(
        type="list",
        formula1='"Cố định,Thả nổi (Base + Biên độ)"',
        allow_blank=False,
        showDropDown=False
    )
    dv_ls_method.sqref = "B13"
    ws.add_data_validation(dv_ls_method)

    dv_thu_no = DataValidation(
        type="list",
        formula1='"TT39/TT06 (mặc định),Gốc trước - Lãi sau,Lãi trước - Gốc sau"',
        allow_blank=False,
        showDropDown=False
    )
    dv_thu_no.sqref = "B16"
    ws.add_data_validation(dv_thu_no)

    set_print_a4_landscape(ws, "CẤU HÌNH KHOẢN VAY")


def build_sheet_holidays(ws):
    ws.title = "NGÀY NGHỈ LỄ"
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 16

    # Title
    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value = "📅 DANH SÁCH NGÀY NGHỈ LỄ VIỆT NAM (2020–2030)"
    c.font = make_font(bold=True, size=13, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:D2")
    c = ws["A2"]
    c.value = "Người dùng có thể bổ sung/chỉnh sửa. Cột B phải là kiểu Date (DD/MM/YYYY). Sắp xếp theo ngày tăng dần."
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 16

    # Header row
    headers = ["STT", "Ngày", "Tên ngày nghỉ lễ", "Loại"]
    for col, h in enumerate(headers, 1):
        style_header(ws, 3, col, h)
    ws.row_dimensions[3].height = 18
    ws.freeze_panes = "A4"

    holidays = get_holidays()
    for i, (d, name, htype) in enumerate(holidays):
        r = 4 + i
        alt = (i % 2 == 1)
        bg = C_ALT_ROW if alt else C_WHITE

        c1 = ws.cell(row=r, column=1, value=i + 1)
        c1.font = make_font(size=9)
        c1.fill = make_fill(bg)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = thin_border

        c2 = ws.cell(row=r, column=2, value=d)
        c2.number_format = FMT_DATE
        c2.font = make_font(size=9)
        c2.fill = make_fill(C_INPUT_BG)
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.border = thin_border

        c3 = ws.cell(row=r, column=3, value=name)
        c3.font = make_font(size=9)
        c3.fill = make_fill(C_INPUT_BG)
        c3.alignment = Alignment(horizontal="left", vertical="center")
        c3.border = thin_border

        c4 = ws.cell(row=r, column=4, value=htype)
        c4.font = make_font(size=9)
        c4.fill = make_fill(C_INPUT_BG)
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c4.border = thin_border

        ws.row_dimensions[r].height = 15

    # Extra blank rows for user additions
    last_data = 3 + len(holidays)
    for i in range(10):
        r = last_data + 1 + i
        for col in range(1, 5):
            c = ws.cell(row=r, column=col)
            c.fill = make_fill(C_INPUT_BG)
            c.border = thin_border
            if col == 2:
                c.number_format = FMT_DATE
        ws.row_dimensions[r].height = 15

    # Type validation
    dv_type = DataValidation(
        type="list",
        formula1='"Lễ,Tết,Nghỉ bù,Tùy chỉnh"',
        allow_blank=True,
        showDropDown=False
    )
    dv_type.sqref = f"D4:D{last_data + 10}"
    ws.add_data_validation(dv_type)

    set_print_a4_landscape(ws, "NGÀY NGHỈ LỄ VIỆT NAM 2020–2030")


def build_sheet_giaingân(ws):
    ws.title = "GIẢI NGÂN"
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 35

    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = "💰 BẢNG GIẢI NGÂN"
    c.font = make_font(bold=True, size=13, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:E2")
    c = ws["A2"]
    c.value = ("Nhập từng đợt giải ngân. Lãi suất trong hạn tại đây là lãi suất áp dụng từ đầu "
               "(sẽ bị ghi đè bởi bảng Thay Đổi Lãi Suất nếu có thay đổi sau ngày giải ngân).")
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 26

    headers = ["STT", "Ngày giải ngân", "Số tiền giải ngân (VND)", "Lãi suất trong hạn (%/năm)", "Ghi chú"]
    for col, h in enumerate(headers, 1):
        style_header(ws, 3, col, h)
    ws.row_dimensions[3].height = 20
    ws.freeze_panes = "A4"

    # Sample data — scenario: 2 disbursements
    sample = [
        (date(2024, 1, 15), 3_000_000_000, 0.10, "Giải ngân đợt 1"),
        (date(2024, 3,  1), 2_000_000_000, 0.10, "Giải ngân đợt 2"),
    ]

    for i, (d, amt, rate, note) in enumerate(sample):
        r = 4 + i
        ws.cell(row=r, column=1, value=i + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=1).border = thin_border
        ws.cell(row=r, column=1).fill = make_fill(C_WHITE)
        ws.cell(row=r, column=1).font = make_font(size=10)

        style_input(ws, r, 2, d, FMT_DATE, "center")
        style_input(ws, r, 3, amt, FMT_VND, "right")
        style_input(ws, r, 4, rate, FMT_PCT, "right")
        style_input(ws, r, 5, note)
        ws.row_dimensions[r].height = 18

    # Blank input rows
    for i in range(8):
        r = 4 + len(sample) + i
        ws.cell(row=r, column=1, value=len(sample) + i + 1)
        ws.cell(row=r, column=1).border = thin_border
        ws.cell(row=r, column=1).fill = make_fill(C_WHITE)
        ws.cell(row=r, column=1).font = make_font(size=10)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        style_input(ws, r, 2, fmt=FMT_DATE)
        style_input(ws, r, 3, fmt=FMT_VND)
        style_input(ws, r, 4, fmt=FMT_PCT)
        style_input(ws, r, 5)
        ws.row_dimensions[r].height = 18

    # Total row
    total_row = 4 + len(sample) + 8 + 1
    ws.row_dimensions[total_row].height = 20
    style_total(ws, total_row, 1, "TỔNG")
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    c = ws.cell(row=total_row, column=1)
    c.alignment = Alignment(horizontal="left", vertical="center")
    style_total(ws, total_row, 3,
                f"=SUM(C4:C{total_row-1})", FMT_VND)
    style_total(ws, total_row, 4, "")
    style_total(ws, total_row, 5, "")

    set_print_a4_landscape(ws, "BẢNG GIẢI NGÂN")


def build_sheet_lichtrâno(ws):
    ws.title = "LỊCH TRẢ NỢ"
    cols_w = [6, 18, 25, 22, 18, 25, 30]
    for i, w in enumerate(cols_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = "📋 LỊCH TRẢ NỢ THEO HỢP ĐỒNG"
    c.font = make_font(bold=True, size=13, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:G2")
    c = ws["A2"]
    c.value = ("Nhập kế hoạch trả nợ theo hợp đồng. Cột 'Ngày điều chỉnh' tự tính: nếu ngày HĐ rơi vào T7/CN/Lễ "
               "sẽ chuyển sang ngày làm việc tiếp theo (dùng WORKDAY.INTL). "
               "Ô màu xanh nhạt = nhập liệu. Ô xám = tự tính.")
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 30

    headers = [
        "STT",
        "Ngày đáo hạn\nGốc (HĐ)",
        "Ngày đáo hạn\nGốc (Điều chỉnh)",
        "Số tiền gốc\nphải trả (VND)",
        "Ngày đáo hạn\nLãi (HĐ)",
        "Ngày đáo hạn\nLãi (Điều chỉnh)",
        "Ghi chú"
    ]
    for col, h in enumerate(headers, 1):
        style_header(ws, 3, col, h)
    ws.row_dimensions[3].height = 30
    ws.freeze_panes = "A4"

    # Holiday range reference (for WORKDAY.INTL formula)
    # We'll reference NGÀY NGHỈ LỄ!B4:B200
    hol_ref = "'NGÀY NGHỈ LỄ'!$B$4:$B$200"

    # Sample repayment schedule:
    # Loan disbursed 15/01/2024 & 01/03/2024 → 24 monthly installments
    # Irregular principal: months 1-6: 200M, months 7-12: 250M, months 13-24: divide remainder
    # First payment: 15/02/2024
    from dateutil.relativedelta import relativedelta

    schedule = []
    # Remaining principal after full drawdown = 5,000,000,000
    principal_payments = (
        [200_000_000] * 6 +        # kỳ 1-6
        [250_000_000] * 6 +        # kỳ 7-12
        [258_333_334] * 11 +       # kỳ 13-23  (3,100,000,000 / 12 = 258,333,333.33)
        [258_333_337]              # kỳ 24 (adjustment)
    )

    base_date = date(2024, 2, 15)
    for i in range(24):
        due_date = base_date + relativedelta(months=i)
        schedule.append((due_date, principal_payments[i], due_date))

    for i, (goc_hd, amt, lai_hd) in enumerate(schedule):
        r = 4 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=1).border = thin_border
        ws.cell(row=r, column=1).fill = make_fill(C_WHITE)
        ws.cell(row=r, column=1).font = make_font(size=10)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")

        style_input(ws, r, 2, goc_hd, FMT_DATE, "center")

        # Col C: adjusted date formula
        bc = get_column_letter(2)
        c3 = ws.cell(row=r, column=3)
        c3.value = (f'=IF({bc}{r}="","",IF(AND(WEEKDAY({bc}{r},2)<=5,'
                    f'COUNTIF({hol_ref},{bc}{r})=0),{bc}{r},'
                    f'WORKDAY.INTL({bc}{r},1,1,{hol_ref})))')
        c3.number_format = FMT_DATE
        c3.font = make_font(size=10)
        c3.fill = make_fill(C_CALC_BG)
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c3.border = thin_border

        style_input(ws, r, 4, amt, FMT_VND, "right")
        style_input(ws, r, 5, lai_hd, FMT_DATE, "center")

        # Col F: adjusted interest date
        ec = get_column_letter(5)
        c6 = ws.cell(row=r, column=6)
        c6.value = (f'=IF({ec}{r}="","",IF(AND(WEEKDAY({ec}{r},2)<=5,'
                    f'COUNTIF({hol_ref},{ec}{r})=0),{ec}{r},'
                    f'WORKDAY.INTL({ec}{r},1,1,{hol_ref})))')
        c6.number_format = FMT_DATE
        c6.font = make_font(size=10)
        c6.fill = make_fill(C_CALC_BG)
        c6.alignment = Alignment(horizontal="center", vertical="center")
        c6.border = thin_border

        style_input(ws, r, 7)
        ws.row_dimensions[r].height = 16

    # Blank rows
    last_data = 4 + len(schedule)
    for i in range(10):
        r = last_data + i
        ws.cell(row=r, column=1, value=len(schedule) + i + 1)
        ws.cell(row=r, column=1).border = thin_border
        ws.cell(row=r, column=1).fill = make_fill(C_WHITE)
        ws.cell(row=r, column=1).font = make_font(size=10)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        style_input(ws, r, 2, fmt=FMT_DATE)
        bc2 = get_column_letter(2)
        c3 = ws.cell(row=r, column=3)
        c3.value = (f'=IF({bc2}{r}="","",IF(AND(WEEKDAY({bc2}{r},2)<=5,'
                    f'COUNTIF({hol_ref},{bc2}{r})=0),{bc2}{r},'
                    f'WORKDAY.INTL({bc2}{r},1,1,{hol_ref})))')
        c3.number_format = FMT_DATE
        c3.fill = make_fill(C_CALC_BG)
        c3.border = thin_border
        c3.alignment = Alignment(horizontal="center", vertical="center")
        style_input(ws, r, 4, fmt=FMT_VND)
        style_input(ws, r, 5, fmt=FMT_DATE)
        ec2 = get_column_letter(5)
        c6 = ws.cell(row=r, column=6)
        c6.value = (f'=IF({ec2}{r}="","",IF(AND(WEEKDAY({ec2}{r},2)<=5,'
                    f'COUNTIF({hol_ref},{ec2}{r})=0),{ec2}{r},'
                    f'WORKDAY.INTL({ec2}{r},1,1,{hol_ref})))')
        c6.number_format = FMT_DATE
        c6.fill = make_fill(C_CALC_BG)
        c6.border = thin_border
        c6.alignment = Alignment(horizontal="center", vertical="center")
        style_input(ws, r, 7)
        ws.row_dimensions[r].height = 16

    # Total
    total_row = last_data + 10
    style_total(ws, total_row, 1, "TỔNG GỐC")
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="left")
    style_total(ws, total_row, 4, f"=SUM(D4:D{total_row-1})", FMT_VND)
    for col in [5, 6, 7]:
        style_total(ws, total_row, col, "")
    ws.row_dimensions[total_row].height = 20

    set_print_a4_landscape(ws, "LỊCH TRẢ NỢ THEO HỢP ĐỒNG")


def build_sheet_laisu(ws):
    ws.title = "THAY ĐỔI LÃI SUẤT"
    col_widths = [6, 18, 22, 22, 22, 35]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells("A1:F1")
    c = ws["A1"]
    c.value = "📈 BẢNG THAY ĐỔI LÃI SUẤT"
    c.font = make_font(bold=True, size=13, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:F2")
    c = ws["A2"]
    c.value = ("Nhập các lần điều chỉnh lãi suất kèm ngày hiệu lực. Phải sắp xếp theo ngày tăng dần. "
               "Hàng đầu tiên nên là ngày giải ngân với lãi suất ban đầu. "
               "Lãi suất quá hạn = Lãi trong hạn × Hệ số (từ CẤU HÌNH) trừ khi nhập tay.")
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 30

    headers = [
        "STT",
        "Ngày hiệu lực",
        "LS trong hạn\n(%/năm)",
        "LS quá hạn\n(%/năm) [tự tính]",
        "LS chậm trả lãi\n(%/năm)",
        "Cơ sở / Ghi chú"
    ]
    for col, h in enumerate(headers, 1):
        style_header(ws, 3, col, h)
    ws.row_dimensions[3].height = 30
    ws.freeze_panes = "A4"

    # Sample: initial rate 10%, changes to 11% after 1 year
    sample_rates = [
        (date(2024, 1, 15), 0.10, "Lãi suất ban đầu theo HĐTD"),
        (date(2025, 1, 15), 0.11, "Điều chỉnh lãi suất theo phụ lục HĐTD số 01"),
    ]

    for i, (eff_date, rate, note) in enumerate(sample_rates):
        r = 4 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=1).border = thin_border
        ws.cell(row=r, column=1).fill = make_fill(C_WHITE)
        ws.cell(row=r, column=1).font = make_font(size=10)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")

        style_input(ws, r, 2, eff_date, FMT_DATE, "center")
        style_input(ws, r, 3, rate, FMT_PCT, "right")

        # Col D: auto-calculate overdue rate = in-term × multiplier from CONFIG
        c4 = ws.cell(row=r, column=4)
        c4.value = f"=C{r}*'CẤU HÌNH'!$B$11"
        c4.number_format = FMT_PCT
        c4.fill = make_fill(C_CALC_BG)
        c4.font = make_font(size=10)
        c4.alignment = Alignment(horizontal="right", vertical="center")
        c4.border = thin_border

        # Col E: late interest rate from CONFIG (can override)
        c5 = ws.cell(row=r, column=5)
        c5.value = f"='CẤU HÌNH'!$B$12"
        c5.number_format = FMT_PCT
        c5.fill = make_fill(C_CALC_BG)
        c5.font = make_font(size=10)
        c5.alignment = Alignment(horizontal="right", vertical="center")
        c5.border = thin_border

        style_input(ws, r, 6, note)
        ws.row_dimensions[r].height = 18

    # Blank rows
    last_data = 4 + len(sample_rates)
    for i in range(8):
        r = last_data + i
        ws.cell(row=r, column=1, value=len(sample_rates) + i + 1)
        ws.cell(row=r, column=1).border = thin_border
        ws.cell(row=r, column=1).fill = make_fill(C_WHITE)
        ws.cell(row=r, column=1).font = make_font(size=10)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        style_input(ws, r, 2, fmt=FMT_DATE)
        style_input(ws, r, 3, fmt=FMT_PCT)
        c4 = ws.cell(row=r, column=4)
        c4.value = f"=IF(C{r}=\"\",\"\",C{r}*'CẤU HÌNH'!$B$11)"
        c4.number_format = FMT_PCT
        c4.fill = make_fill(C_CALC_BG)
        c4.border = thin_border
        c4.alignment = Alignment(horizontal="right", vertical="center")
        c5 = ws.cell(row=r, column=5)
        c5.value = f"=IF(C{r}=\"\",\"\",'CẤU HÌNH'!$B$12)"
        c5.number_format = FMT_PCT
        c5.fill = make_fill(C_CALC_BG)
        c5.border = thin_border
        c5.alignment = Alignment(horizontal="right", vertical="center")
        style_input(ws, r, 6)
        ws.row_dimensions[r].height = 18

    set_print_a4_landscape(ws, "BẢNG THAY ĐỔI LÃI SUẤT")


def build_sheet_thanhtoan(ws):
    ws.title = "THANH TOÁN THỰC TẾ"
    col_widths = [6, 18, 22, 22, 22, 22, 22, 30]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = "💳 THANH TOÁN THỰC TẾ CỦA KHÁCH HÀNG"
    c.font = make_font(bold=True, size=13, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:H2")
    c = ws["A2"]
    c.value = ("Nhập các khoản đã thanh toán thực tế. Chọn 'Loại phân bổ': "
               "'Tự động' → nhập Tổng tiền, hệ thống tự phân bổ theo thứ tự thu nợ; "
               "'Thủ công' → nhập riêng từng phần gốc/lãi/phạt. "
               "Phải sắp xếp theo ngày tăng dần.")
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 30

    headers = [
        "STT",
        "Ngày thanh toán",
        "Loại phân bổ",
        "Tổng tiền TT\n(Tự động)",
        "Phân bổ Gốc\n(Thủ công)",
        "Phân bổ Lãi\n(Thủ công)",
        "Phân bổ Phạt\n(Thủ công)",
        "Ghi chú"
    ]
    for col, h in enumerate(headers, 1):
        style_header(ws, 3, col, h)
    ws.row_dimensions[3].height = 30
    ws.freeze_panes = "A4"

    # Sample payments matching the scenario
    sample_payments = [
        (date(2024, 2, 15), "Tự động", 237_500_000, None, None, None, "Trả kỳ 1 (trong hạn)"),
        (date(2024, 3, 15), "Tự động", 235_833_333, None, None, None, "Trả kỳ 2 (trong hạn)"),
        (date(2024, 4, 15), "Tự động", 232_500_000, None, None, None, "Trả kỳ 3 (trong hạn)"),
        (date(2024, 5, 15), "Tự động", 230_416_667, None, None, None, "Trả kỳ 4 (trong hạn)"),
        (date(2024, 6, 17), "Tự động", 228_333_333, None, None, None, "Trả kỳ 5 (trong hạn, 15/6 là T7 → điều chỉnh 17/6)"),
        (date(2024, 7, 15), "Tự động", 226_250_000, None, None, None, "Trả kỳ 6 (trong hạn)"),
        # Kỳ 7 và 8 trễ: gộp trả ngày 30/09/2024
        (date(2024, 9, 30), "Tự động", 505_000_000, None, None, None, "Trả gộp kỳ 7+8 (quá hạn ~30 ngày)"),
        # Trả một phần trước hạn kỳ 10
        (date(2024, 11, 15), "Thủ công", None, 500_000_000, 200_000_000, 0, "Trả trước hạn kỳ 10"),
    ]

    for i, (dt, ptype, total, goc, lai, phat, note) in enumerate(sample_payments):
        r = 4 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=1).border = thin_border
        ws.cell(row=r, column=1).fill = make_fill(C_WHITE)
        ws.cell(row=r, column=1).font = make_font(size=10)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")

        style_input(ws, r, 2, dt, FMT_DATE, "center")
        style_input(ws, r, 3, ptype)
        style_input(ws, r, 4, total, FMT_VND, "right")
        style_input(ws, r, 5, goc, FMT_VND, "right")
        style_input(ws, r, 6, lai, FMT_VND, "right")
        style_input(ws, r, 7, phat, FMT_VND, "right")
        style_input(ws, r, 8, note)
        ws.row_dimensions[r].height = 18

    # Blank rows
    last_data = 4 + len(sample_payments)
    for i in range(12):
        r = last_data + i
        ws.cell(row=r, column=1, value=len(sample_payments) + i + 1)
        ws.cell(row=r, column=1).border = thin_border
        ws.cell(row=r, column=1).fill = make_fill(C_WHITE)
        ws.cell(row=r, column=1).font = make_font(size=10)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        style_input(ws, r, 2, fmt=FMT_DATE)
        style_input(ws, r, 3, "Tự động")
        style_input(ws, r, 4, fmt=FMT_VND)
        style_input(ws, r, 5, fmt=FMT_VND)
        style_input(ws, r, 6, fmt=FMT_VND)
        style_input(ws, r, 7, fmt=FMT_VND)
        style_input(ws, r, 8)
        ws.row_dimensions[r].height = 18

    # Dropdown for allocation type
    dv_alloc = DataValidation(
        type="list",
        formula1='"Tự động,Thủ công"',
        allow_blank=False,
        showDropDown=False
    )
    dv_alloc.sqref = f"C4:C{last_data + 12}"
    ws.add_data_validation(dv_alloc)

    # Total row
    total_row = last_data + 12 + 1
    ws.row_dimensions[total_row].height = 20
    style_total(ws, total_row, 1, "TỔNG ĐÃ THANH TOÁN")
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="left")
    style_total(ws, total_row, 4, f"=SUMIF(C4:C{total_row-1},\"Tự động\",D4:D{total_row-1})", FMT_VND)
    style_total(ws, total_row, 5, f"=SUMIF(C4:C{total_row-1},\"Thủ công\",E4:E{total_row-1})", FMT_VND)
    style_total(ws, total_row, 6, f"=SUMIF(C4:C{total_row-1},\"Thủ công\",F4:F{total_row-1})", FMT_VND)
    style_total(ws, total_row, 7, f"=SUMIF(C4:C{total_row-1},\"Thủ công\",G4:G{total_row-1})", FMT_VND)
    style_total(ws, total_row, 8, "")

    set_print_a4_landscape(ws, "THANH TOÁN THỰC TẾ")


NUM_ROWS = 3650   # ~10 years; rows stop auto-populating when date > calc date
FIRST_ROW = 5    # first data row in BẢNG TÍNH


def build_sheet_bangtính(ws):
    """
    Daily engine: up to NUM_ROWS rows.
    Column A uses formulas so the engine automatically stops at 'CẤU HÌNH'!B5.
    All other columns are IF(A{r}="","",...) so blank rows stay blank.

    NEW column layout (28 cols, A–AB):
    A  Ngày
    B  Ngày làm việc?
    C  Giải ngân
    D  Gốc đến hạn (kế hoạch)
    E  Lãi đến hạn?
    F  Thanh toán trong ngày
    G  LS trong hạn (%/năm)
    H  LS quá hạn (%/năm)
    I  LS chậm trả (%/năm)
    J  Dư gốc IH (đầu ngày)
    K  Dư gốc QH (đầu ngày)
    L  Lãi IH chưa trả (đầu ngày)
    M  Lãi trên Gốc QH chưa trả (đầu ngày)   ← SEPARATED
    N  Lãi chậm trả chưa trả (đầu ngày)       ← NEW / SEPARATED
    O  Lãi IH phát sinh
    P  Lãi QH phát sinh (trên Gốc QH)
    Q  Lãi CT phát sinh (trên N)
    R  Phân bổ TT → Gốc QH
    S  Phân bổ TT → Lãi trên Gốc QH
    T  Phân bổ TT → Gốc IH
    U  Phân bổ TT → Lãi IH
    V  Phân bổ TT → Lãi chậm trả
    W  Dư gốc IH (cuối ngày)          = MAX(0, J+C-D)           [FIXED]
    X  Dư gốc QH (cuối ngày)          = MAX(0, K+D-T-R)         [FIXED]
    Y  Lãi IH chưa trả (cuối ngày)    = IF(E=✓, 0, L+O)        [FIXED]
    Z  Lãi trên Gốc QH chưa trả (cuối ngày) = MAX(0, M+P-S)   [NEW]
    AA Lãi chậm trả chưa trả (cuối ngày)                        [NEW]
    AB Kiểm tra TT (unallocated)
    """
    ws.title = "BẢNG TÍNH"
    ws.sheet_view.showGridLines = True

    # Column layout: 28 columns A–AB
    col_headers = [
        ("A",  14, "Ngày"),
        ("B",  10, "Ngày\nlàm việc?"),
        ("C",  18, "Giải ngân\ntrong ngày"),
        ("D",  18, "Gốc đến hạn\ntrong ngày (HĐ)"),
        ("E",  10, "Lãi\nđến hạn?"),
        ("F",  18, "Thanh toán\ntrong ngày"),
        ("G",  12, "LS trong hạn\n(%/năm)"),
        ("H",  12, "LS quá hạn\n(%/năm)"),
        ("I",  12, "LS chậm trả\n(%/năm)"),
        ("J",  18, "Dư gốc IH\n(đầu ngày)"),
        ("K",  18, "Dư gốc QH\n(đầu ngày)"),
        ("L",  18, "Lãi IH chưa\ntrả (đầu ngày)"),
        ("M",  20, "Lãi trên Gốc QH\nchưa trả (đầu ngày)"),
        ("N",  20, "Lãi chậm trả\nchưa trả (đầu ngày)"),
        ("O",  18, "Lãi IH phát\nsinh hôm nay"),
        ("P",  18, "Lãi QH phát sinh\n(trên Gốc QH)"),
        ("Q",  18, "Lãi CT phát sinh\n(trên lãi CT base)"),
        ("R",  18, "Phân bổ TT\n→ Gốc QH"),
        ("S",  18, "Phân bổ TT\n→ Lãi trên Gốc QH"),
        ("T",  18, "Phân bổ TT\n→ Gốc IH"),
        ("U",  18, "Phân bổ TT\n→ Lãi IH"),
        ("V",  18, "Phân bổ TT\n→ Lãi chậm trả"),
        ("W",  18, "Dư gốc IH\n(cuối ngày)"),
        ("X",  18, "Dư gốc QH\n(cuối ngày)"),
        ("Y",  18, "Lãi IH chưa\ntrả (cuối ngày)"),
        ("Z",  20, "Lãi trên Gốc QH\nchưa trả (cuối ngày)"),
        ("AA", 20, "Lãi chậm trả\nchưa trả (cuối ngày)"),
        ("AB", 16, "Kiểm tra TT\n(còn dư/thiếu)"),
    ]

    for col_ltr, w, _ in col_headers:
        ws.column_dimensions[col_ltr].width = w

    # Title
    ws.merge_cells("A1:AB1")
    c = ws["A1"]
    c.value = "⚙ BẢNG TÍNH CHI TIẾT THEO NGÀY (DAILY ENGINE)"
    c.font = make_font(bold=True, size=13, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:AB2")
    c = ws["A2"]
    c.value = (
        "Mỗi hàng = 1 ngày. Cột A tự động sinh ngày từ ngày giải ngân đầu tiên đến "
        "'Ngày tính toán' trong CẤU HÌNH (tối đa 3,650 ngày ~ 10 năm). "
        "Ô xanh nhạt = nhập liệu. Ô xám = công thức. KHÔNG chỉnh sửa ô công thức."
    )
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 24

    ws.merge_cells("A3:AB3")
    c = ws["A3"]
    c.value = (
        "Mẫu số: =IF(Actual/360 hoặc 30/360 → 360, else 365)  |  "
        "Lãi QH: trên Gốc QH (cột K) × LS QH (cột H)  |  "
        "Lãi CT: trên lãi chậm trả chưa trả (cột N) × LS CT (cột I)  |  "
        "Lãi CT chuyển vào cột N khi lãi IH đến hạn mà chưa trả"
    )
    c.font = make_font(size=8, color="595959")
    c.fill = make_fill("FFF9E6")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[3].height = 18

    # Header row 4
    for col_idx, (_, _, header) in enumerate(col_headers, 1):
        style_header(ws, 4, col_idx, header)
    ws.row_dimensions[4].height = 34
    ws.freeze_panes = "A5"

    # ─── Shared style objects (reused for performance) ───
    fill_calc     = make_fill(C_CALC_BG)
    fill_overdue  = make_fill(C_OVERDUE_BG)
    fill_inh      = make_fill("EBF3F9")   # light blue for IH end cols
    fill_white    = make_fill(C_WHITE)
    fill_hdr      = make_fill(C_HEADER_BG)
    font_sm       = make_font(size=9)
    font_sm_red   = make_font(size=9, color="C00000")
    font_sm_blue  = make_font(size=9, color="1F4E79", bold=True)
    font_sm_hdr   = make_font(size=9, color=C_HEADER_TXT, bold=True)
    align_c       = Alignment(horizontal="center", vertical="center")
    align_r       = Alignment(horizontal="right",  vertical="center")

    def qcell(ws, r, col_num, formula, fmt, fill, font, align):
        c = ws.cell(row=r, column=col_num, value=formula)
        c.number_format = fmt
        c.fill  = fill
        c.font  = font
        c.alignment = align
        c.border = thin_border
        return c

    # ─── References ───
    hol_ref      = "'NGÀY NGHỈ LỄ'!$B$4:$B$200"
    gn_date      = "'GIẢI NGÂN'!$B$4:$B$20"
    gn_amt       = "'GIẢI NGÂN'!$C$4:$C$20"
    lich_goc     = "'LỊCH TRẢ NỢ'!$C$4:$C$40"   # adjusted due dates
    lich_goc_amt = "'LỊCH TRẢ NỢ'!$D$4:$D$40"
    lich_lai     = "'LỊCH TRẢ NỢ'!$F$4:$F$40"   # adjusted interest due dates
    ls_dates     = "'THAY ĐỔI LÃI SUẤT'!$B$4:$B$20"
    ls_inh       = "'THAY ĐỔI LÃI SUẤT'!$C$4:$C$20"
    ls_qh        = "'THAY ĐỔI LÃI SUẤT'!$D$4:$D$20"
    ls_ct        = "'THAY ĐỔI LÃI SUẤT'!$E$4:$E$20"
    tt_dates     = "'THANH TOÁN THỰC TẾ'!$B$4:$B$30"
    tt_type      = "'THANH TOÁN THỰC TẾ'!$C$4:$C$30"
    tt_total     = "'THANH TOÁN THỰC TẾ'!$D$4:$D$30"
    tt_goc       = "'THANH TOÁN THỰC TẾ'!$E$4:$E$30"
    tt_lai       = "'THANH TOÁN THỰC TẾ'!$F$4:$F$30"
    tt_phat      = "'THANH TOÁN THỰC TẾ'!$G$4:$G$30"
    cfg_denom    = "IF('CẤU HÌNH'!$B$6=\"Actual/360\",360,IF('CẤU HÌNH'!$B$6=\"30/360\",360,365))"
    cfg_calcdate = "'CẤU HÌNH'!$B$5"

    # ─── Build rows ───
    for row_idx in range(NUM_ROWS):
        r = FIRST_ROW + row_idx
        ws.row_dimensions[r].height = 14
        prev_r = r - 1
        is_first = (row_idx == 0)

        # ── Col A: Date (dynamic formula) ──────────────────────────────────
        ca = ws.cell(row=r, column=1)
        if is_first:
            # Start date = min disbursement date
            ca.value = f"=IF(COUNTA({gn_date})=0,\"\",MIN({gn_date}))"
        else:
            # Advance by 1 day; blank out once calc date is reached
            ca.value = (
                f"=IF(OR(A{prev_r}=\"\",A{prev_r}>={cfg_calcdate}),\"\",A{prev_r}+1)"
            )
        ca.number_format = FMT_DATE
        ca.fill  = fill_white
        ca.font  = font_sm
        ca.alignment = align_c
        ca.border = thin_border

        # All remaining formulas are wrapped: =IF(A{r}="","",[formula])
        def w(formula):
            """Wrap formula so blank rows stay blank."""
            return f'=IF(A{r}="","",{formula})'

        # ── Col B: Is workday? ──────────────────────────────────────────────
        qcell(ws, r, 2,
              w(f'IF(AND(WEEKDAY(A{r},2)<=5,COUNTIF({hol_ref},A{r})=0),"✓","✗")'),
              "@", fill_calc, font_sm, align_c)

        # ── Col C: Disbursement today ───────────────────────────────────────
        qcell(ws, r, 3, w(f"SUMIFS({gn_amt},{gn_date},A{r})"),
              FMT_VND, fill_calc, font_sm, align_r)

        # ── Col D: Principal due today (schedule) ───────────────────────────
        qcell(ws, r, 4, w(f"SUMIFS({lich_goc_amt},{lich_goc},A{r})"),
              FMT_VND, fill_calc, font_sm, align_r)

        # ── Col E: Interest due date? ────────────────────────────────────────
        qcell(ws, r, 5,
              w(f'IF(COUNTIF({lich_lai},A{r})>0,"✓","")'),
              "@", fill_calc, font_sm, align_c)

        # ── Col F: Total payment today ───────────────────────────────────────
        # Auto payments + sum of manual components
        qcell(ws, r, 6,
              w(f'SUMIFS({tt_total},{tt_dates},A{r},{tt_type},"Tự động")'
                f'+SUMIFS({tt_goc},{tt_dates},A{r},{tt_type},"Thủ công")'
                f'+SUMIFS({tt_lai},{tt_dates},A{r},{tt_type},"Thủ công")'
                f'+SUMIFS({tt_phat},{tt_dates},A{r},{tt_type},"Thủ công")'),
              FMT_VND, fill_calc, font_sm, align_r)

        # ── Col G: LS trong hạn ─────────────────────────────────────────────
        qcell(ws, r, 7,
              w(f'IFERROR(INDEX({ls_inh},MATCH(A{r},{ls_dates},1)),"")'),
              FMT_PCT, fill_calc, font_sm, align_r)

        # ── Col H: LS quá hạn ───────────────────────────────────────────────
        qcell(ws, r, 8,
              w(f'IFERROR(INDEX({ls_qh},MATCH(A{r},{ls_dates},1)),"")'),
              FMT_PCT, fill_calc, font_sm, align_r)

        # ── Col I: LS chậm trả ──────────────────────────────────────────────
        qcell(ws, r, 9,
              w(f'IFERROR(INDEX({ls_ct},MATCH(A{r},{ls_dates},1)),"")'),
              FMT_PCT, fill_calc, font_sm, align_r)

        # ── Opening balances (J, K, L, M, N) ────────────────────────────────
        if is_first:
            for col_num in [10, 11, 12, 13, 14]:
                qcell(ws, r, col_num, w("0"), FMT_VND, fill_calc, font_sm, align_r)
        else:
            # J = W{prev} (gốc IH end prev day)
            qcell(ws, r, 10, w(f"W{prev_r}"), FMT_VND, fill_calc, font_sm, align_r)
            # K = X{prev} (gốc QH end prev day)
            qcell(ws, r, 11, w(f"X{prev_r}"), FMT_VND, fill_calc, font_sm, align_r)
            # L = Y{prev} (lãi IH end prev day)
            qcell(ws, r, 12, w(f"Y{prev_r}"), FMT_VND, fill_calc, font_sm, align_r)
            # M = Z{prev} (lãi trên gốc QH end prev day)
            qcell(ws, r, 13, w(f"Z{prev_r}"), FMT_VND, fill_overdue, font_sm_red, align_r)
            # N = AA{prev} (lãi chậm trả end prev day)
            qcell(ws, r, 14, w(f"AA{prev_r}"), FMT_VND, fill_overdue, font_sm_red, align_r)

        if is_first:
            qcell(ws, r, 13, w("0"), FMT_VND, fill_overdue, font_sm_red, align_r)
            qcell(ws, r, 14, w("0"), FMT_VND, fill_overdue, font_sm_red, align_r)

        # ── Col O: Lãi IH phát sinh = J × G / denom ─────────────────────────
        qcell(ws, r, 15, w(f"J{r}*G{r}/({cfg_denom})"),
              FMT_VND, fill_calc, font_sm, align_r)

        # ── Col P: Lãi QH phát sinh = K × H / denom ─────────────────────────
        qcell(ws, r, 16, w(f"K{r}*H{r}/({cfg_denom})"),
              FMT_VND, fill_overdue, font_sm_red, align_r)

        # ── Col Q: Lãi CT phát sinh = N × I / denom ─────────────────────────
        # Lãi CT accrues on N (lãi IH đã đến hạn chưa trả = the overdue interest base)
        qcell(ws, r, 17, w(f"N{r}*I{r}/({cfg_denom})"),
              FMT_VND, fill_overdue, font_sm_red, align_r)

        # ── Payment waterfall (R–V) ──────────────────────────────────────────
        # has_manual: non-zero manual allocation today?
        has_manual = (
            f'SUMIFS({tt_goc},{tt_dates},A{r},{tt_type},"Thủ công")'
            f'+SUMIFS({tt_lai},{tt_dates},A{r},{tt_type},"Thủ công")'
            f'+SUMIFS({tt_phat},{tt_dates},A{r},{tt_type},"Thủ công")>0'
        )
        tt_auto_today  = f'SUMIFS({tt_total},{tt_dates},A{r},{tt_type},"Tự động")'
        tt_man_goc     = f'SUMIFS({tt_goc},{tt_dates},A{r},{tt_type},"Thủ công")'
        tt_man_lai     = f'SUMIFS({tt_lai},{tt_dates},A{r},{tt_type},"Thủ công")'
        tt_man_phat    = f'SUMIFS({tt_phat},{tt_dates},A{r},{tt_type},"Thủ công")'

        # TT39/TT06 waterfall order:
        # 1. R → Gốc QH (cap at K)
        r_auto = f"MAX(0,MIN({tt_auto_today},K{r}))"
        r_man  = f"MIN({tt_man_goc},K{r})"
        qcell(ws, r, 18, w(f"IF({has_manual},{r_man},{r_auto})"),
              FMT_VND, fill_calc, font_sm, align_r)

        # 2. S → Lãi trên Gốc QH (cap at M opening + P today)
        s_auto = f"MAX(0,MIN({tt_auto_today}-R{r},M{r}+P{r}))"
        s_man  = f"MIN({tt_man_phat},M{r}+P{r})"
        qcell(ws, r, 19, w(f"IF({has_manual},{s_man},{s_auto})"),
              FMT_VND, fill_calc, font_sm, align_r)

        # 3. T → Gốc IH (cap at available in-term principal J + C; covers scheduled D and prepayment)
        t_auto = f"MAX(0,MIN({tt_auto_today}-R{r}-S{r},J{r}+C{r}))"
        t_man  = f"MAX(0,MIN({tt_man_goc}-R{r},J{r}+C{r}))"
        qcell(ws, r, 20, w(f"IF({has_manual},{t_man},{t_auto})"),
              FMT_VND, fill_calc, font_sm, align_r)

        # 4. U → Lãi IH đến hạn / accrued (cap at L opening + O today)
        u_auto = f'MAX(0,MIN({tt_auto_today}-R{r}-S{r}-T{r},L{r}+O{r}))'
        u_man  = f"MIN({tt_man_lai},L{r}+O{r})"
        qcell(ws, r, 21, w(f"IF({has_manual},{u_man},{u_auto})"),
              FMT_VND, fill_calc, font_sm, align_r)

        # 5. V → Lãi chậm trả (cap at N opening + Q today)
        v_auto = f"MAX(0,MIN({tt_auto_today}-R{r}-S{r}-T{r}-U{r},N{r}+Q{r}))"
        v_man  = f"MAX(0,MIN({tt_man_phat}-S{r},N{r}+Q{r}))"
        qcell(ws, r, 22, w(f"IF({has_manual},{v_man},{v_auto})"),
              FMT_VND, fill_calc, font_sm, align_r)

        # ── End-of-day balances ──────────────────────────────────────────────

        # W: Dư gốc IH cuối ngày
        # Reduces by scheduled D and any prepayment/manual payment (MAX(D, T)).
        qcell(ws, r, 23, w(f"MAX(0,J{r}+C{r}-MAX(D{r},T{r}))"),
              FMT_VND, fill_inh, font_sm_blue, align_r)

        # X: Dư gốc QH cuối ngày
        # Opening K + any unpaid due principal (MAX(0, D - T)) - Overdue principal repaid R
        qcell(ws, r, 24, w(f"MAX(0,K{r}+MAX(0,D{r}-T{r})-R{r})"),
              FMT_VND, fill_overdue, font_sm_red, align_r)

        # Y: Lãi IH chưa trả cuối ngày
        # On interest due dates (E=✓): unpaid in-term interest transfers to late interest base (col AA), so Y resets to 0.
        # On non-due days: retains accrued interest after deducting any payment U.
        qcell(ws, r, 25,
              w(f'MAX(0,IF(E{r}="✓",0,L{r}+O{r}-U{r}))'),
              FMT_VND, fill_calc, font_sm, align_r)

        # Z: Lãi trên Gốc QH chưa trả cuối ngày
        # = Opening M + Accrued P - Paid S
        qcell(ws, r, 26, w(f"MAX(0,M{r}+P{r}-S{r})"),
              FMT_VND, fill_overdue, font_sm_red, align_r)

        # AA: Lãi chậm trả chưa trả cuối ngày
        # = Opening N + Accrued Q + Rolled-over unpaid in-term interest on due dates - Paid V
        qcell(ws, r, 27,
              w(f'MAX(0,N{r}+Q{r}+IF(E{r}="✓",MAX(0,L{r}+O{r}-U{r}),0)-V{r})'),
              FMT_VND, fill_overdue, font_sm_red, align_r)

        # AB: Kiểm tra — unallocated payment remainder (should be 0 or small rounding)
        qcell(ws, r, 28, w(f"F{r}-R{r}-S{r}-T{r}-U{r}-V{r}"),
              FMT_VND, fill_calc, font_sm, align_r)

    # ── Totals row (below last data row) ────────────────────────────────────
    last_r  = FIRST_ROW + NUM_ROWS - 1
    total_r = last_r + 2
    ws.row_dimensions[total_r].height = 20

    ws.merge_cells(start_row=total_r, start_column=1,
                   end_row=total_r, end_column=9)
    ct = ws.cell(row=total_r, column=1, value="TỔNG LŨY KẾ")
    ct.font = make_font(bold=True, size=10, color=C_HEADER_TXT)
    ct.fill = make_fill(C_HEADER_BG)
    ct.alignment = Alignment(horizontal="left", vertical="center")

    # Sum columns: O(15)→AB(28)
    sum_cols = {
        15: "Lãi IH", 16: "Lãi QH", 17: "Lãi CT",
        18: "PB GốcQH", 19: "PB LãiQH", 20: "PB GốcIH",
        21: "PB LãiIH", 22: "PB LãiCT",
        23: "", 24: "", 25: "", 26: "", 27: "", 28: "",
    }
    for col_num in range(10, 29):
        c_tot = ws.cell(row=total_r, column=col_num)
        col_ltr = get_column_letter(col_num)
        if col_num >= 15:   # sum accruals and allocations
            c_tot.value = f"=SUM({col_ltr}{FIRST_ROW}:{col_ltr}{last_r})"
        else:
            c_tot.value = ""
        c_tot.number_format = FMT_VND
        c_tot.font  = make_font(bold=True, size=9, color=C_HEADER_TXT)
        c_tot.fill  = make_fill(C_HEADER_BG)
        c_tot.alignment = align_r
        c_tot.border = thin_border

    set_print_a4_landscape(ws, "BẢNG TÍNH CHI TIẾT THEO NGÀY")
    ws.print_title_rows = "1:4"


def build_sheet_tonghopky(ws):
    ws.title = "TỔNG HỢP KỲ"
    # Daily engine last row
    bt_last = FIRST_ROW + NUM_ROWS - 1

    col_widths = [6, 14, 14, 8, 30, 14, 20, 20, 20, 20, 20, 20, 22, 22]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells("A1:N1")
    c = ws["A1"]
    c.value = "📊 BẢNG TỔNG HỢP THEO KỲ"
    c.font = make_font(bold=True, size=13, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:N2")
    c = ws["A2"]
    c.value = (
        "Nhập 'Từ ngày' và 'Đến ngày' cho mỗi kỳ (mỗi kỳ = khoảng giữa 2 sự kiện liên tiếp: "
        "giải ngân / đáo hạn / thanh toán / đổi lãi suất). Các cột còn lại tự tính từ BẢNG TÍNH. "
        "Sự kiện = mô tả lý do tách kỳ. Lãi QH = trên Gốc QH. Lãi CT = lãi chậm trả (trên lãi IH quá hạn)."
    )
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 30

    headers = [
        "STT",
        "Từ ngày",
        "Đến ngày",
        "Số ngày",
        "Sự kiện / Ghi chú",
        "LS trong\nhạn (%/năm)",
        "Dư gốc IH\n(đầu kỳ)",
        "Dư gốc QH\n(đầu kỳ)",
        "Lãi IH phát\nsinh trong kỳ",
        "Lãi trên Gốc QH\nphát sinh",          # ← separated
        "Lãi chậm trả\nphát sinh",              # ← separated
        "Thanh toán\ntrong kỳ",
        "Tổng Lãi IH",
        "Tổng Lãi phạt",                        # ← QH + CT combined for court summary
    ]
    for col, h in enumerate(headers, 1):
        style_header(ws, 3, col, h)
    ws.row_dimensions[3].height = 30
    ws.freeze_panes = "A4"

    # Reference to daily engine (BẢNG TÍNH) — updated column letters
    bt_dates = f"'BẢNG TÍNH'!$A$5:$A${bt_last}"
    bt_laiIH = f"'BẢNG TÍNH'!$O$5:$O${bt_last}"   # col O (was N)
    bt_laiQH = f"'BẢNG TÍNH'!$P$5:$P${bt_last}"   # col P (was O)
    bt_laiCT = f"'BẢNG TÍNH'!$Q$5:$Q${bt_last}"   # col Q (was P)
    bt_tt    = f"'BẢNG TÍNH'!$F$5:$F${bt_last}"
    bt_gocIH = f"'BẢNG TÍNH'!$J$5:$J${bt_last}"
    bt_gocQH = f"'BẢNG TÍNH'!$K$5:$K${bt_last}"
    bt_ls    = f"'BẢNG TÍNH'!$G$5:$G${bt_last}"

    # Sample periods matching the scenario
    sample_periods = [
        (date(2024, 1, 15), date(2024, 2, 14),  "Kỳ 1: Từ giải ngân đợt 1 đến trước kỳ trả đầu tiên"),
        (date(2024, 2, 15), date(2024, 2, 29),  "Kỳ 2: Sau khi trả kỳ 1 đến giải ngân đợt 2"),
        (date(2024, 3,  1), date(2024, 3, 14),  "Kỳ 3: Sau giải ngân đợt 2 đến kỳ trả kỳ 2"),
        (date(2024, 3, 15), date(2024, 4, 14),  "Kỳ 4: Sau kỳ 2"),
        (date(2024, 4, 15), date(2024, 5, 14),  "Kỳ 5: Sau kỳ 3"),
        (date(2024, 5, 15), date(2024, 6, 16),  "Kỳ 6: Sau kỳ 4 (15/6 T7 → adj 17/6)"),
        (date(2024, 6, 17), date(2024, 7, 14),  "Kỳ 7: Sau kỳ 5"),
        (date(2024, 7, 15), date(2024, 8, 14),  "Kỳ 8: Sau kỳ 6 — kỳ 7 không trả → phát sinh QH"),
        (date(2024, 8, 15), date(2024, 9, 14),  "Kỳ 9: Kỳ 8 không trả → tiếp tục phát sinh QH"),
        (date(2024, 9, 15), date(2024, 9, 29),  "Kỳ 10: Sau kỳ 7+8 gộp trả ngày 30/9"),
        (date(2024, 9, 30), date(2024, 10, 14), "Kỳ 11: Sau khi trả gộp kỳ 7+8"),
        (date(2024, 10, 15), date(2024, 11, 14), "Kỳ 12"),
        (date(2024, 11, 15), date(2024, 12, 14), "Kỳ 13: Trả trước hạn 500tr ngày 15/11"),
        (date(2024, 12, 15), date(2025, 1, 14),  "Kỳ 14"),
        (date(2025, 1, 15), date(2025, 1, 14),   "Kỳ 15: Đổi LS 11%/năm từ 15/01/2025"),
        (date(2025, 1, 15), date(2025, 6, 15),   "Kỳ 16: Kỳ còn lại đến ngày tính"),
    ]
    # Remove invalid (start > end)
    sample_periods = [(s, e, n) for s, e, n in sample_periods if s <= e]

    for i, (from_dt, to_dt, note) in enumerate(sample_periods):
        r = 4 + i
        ws.row_dimensions[r].height = 18

        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=1).border = thin_border
        ws.cell(row=r, column=1).fill = make_fill(C_WHITE)
        ws.cell(row=r, column=1).font = make_font(size=10)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")

        style_input(ws, r, 2, from_dt, FMT_DATE, "center")
        style_input(ws, r, 3, to_dt,   FMT_DATE, "center")

        # D: Số ngày
        cd = ws.cell(row=r, column=4)
        cd.value = f"=IF(AND(B{r}<>\"\",C{r}<>\"\"),C{r}-B{r}+1,\"\")"
        cd.font = make_font(size=10)
        cd.fill = make_fill(C_CALC_BG)
        cd.alignment = Alignment(horizontal="center", vertical="center")
        cd.border = thin_border

        style_input(ws, r, 5, note)

        # F: LS trong hạn (from daily engine at start of period)
        cf = ws.cell(row=r, column=6)
        cf.value = f"=IFERROR(INDEX({bt_ls},MATCH(B{r},{bt_dates},0)),\"\")"
        cf.number_format = FMT_PCT
        cf.fill = make_fill(C_CALC_BG)
        cf.font = make_font(size=10)
        cf.alignment = Alignment(horizontal="right", vertical="center")
        cf.border = thin_border

        # G: Dư gốc IH đầu kỳ
        cg = ws.cell(row=r, column=7)
        cg.value = f"=IFERROR(INDEX({bt_gocIH},MATCH(B{r},{bt_dates},0)),\"\")"
        cg.number_format = FMT_VND
        cg.fill = make_fill(C_CALC_BG)
        cg.font = make_font(size=10)
        cg.alignment = Alignment(horizontal="right", vertical="center")
        cg.border = thin_border

        # H: Dư gốc QH đầu kỳ
        ch = ws.cell(row=r, column=8)
        ch.value = f"=IFERROR(INDEX({bt_gocQH},MATCH(B{r},{bt_dates},0)),\"\")"
        ch.number_format = FMT_VND
        ch.fill = make_fill(C_CALC_BG)
        ch.font = make_font(size=10, color="C00000")
        ch.alignment = Alignment(horizontal="right", vertical="center")
        ch.border = thin_border

        # I: Lãi IH phát sinh (SUMIFS from daily)
        ci = ws.cell(row=r, column=9)
        ci.value = (f'=IF(AND(B{r}<>"",C{r}<>""),'
                    f'SUMIFS({bt_laiIH},{bt_dates},">="&B{r},{bt_dates},"<="&C{r}),"")')
        ci.number_format = FMT_VND
        ci.fill = make_fill(C_CALC_BG)
        ci.font = make_font(size=10)
        ci.alignment = Alignment(horizontal="right", vertical="center")
        ci.border = thin_border

        # J: Lãi QH phát sinh
        cj = ws.cell(row=r, column=10)
        cj.value = (f'=IF(AND(B{r}<>"",C{r}<>""),'
                    f'SUMIFS({bt_laiQH},{bt_dates},">="&B{r},{bt_dates},"<="&C{r}),"")')
        cj.number_format = FMT_VND
        cj.fill = make_fill(C_OVERDUE_BG)
        cj.font = make_font(size=10, color="C00000")
        cj.alignment = Alignment(horizontal="right", vertical="center")
        cj.border = thin_border

        # K: Lãi CT phát sinh
        ck = ws.cell(row=r, column=11)
        ck.value = (f'=IF(AND(B{r}<>"",C{r}<>""),'
                    f'SUMIFS({bt_laiCT},{bt_dates},">="&B{r},{bt_dates},"<="&C{r}),"")')
        ck.number_format = FMT_VND
        ck.fill = make_fill(C_OVERDUE_BG)
        ck.font = make_font(size=10, color="C00000")
        ck.alignment = Alignment(horizontal="right", vertical="center")
        ck.border = thin_border

        # L: Thanh toán trong kỳ
        cl = ws.cell(row=r, column=12)
        cl.value = (f'=IF(AND(B{r}<>"",C{r}<>""),'
                    f'SUMIFS({bt_tt},{bt_dates},">="&B{r},{bt_dates},"<="&C{r}),"")')
        cl.number_format = FMT_VND
        cl.fill = make_fill(C_TOTAL_BG)
        cl.font = make_font(size=10)
        cl.alignment = Alignment(horizontal="right", vertical="center")
        cl.border = thin_border

        # M: Tổng lãi IH (col I)
        cm = ws.cell(row=r, column=13)
        cm.value = f'=IF(I{r}<>"",I{r},"")'
        cm.number_format = FMT_VND
        cm.fill = make_fill(C_CALC_BG)
        cm.font = make_font(bold=True, size=10)
        cm.alignment = Alignment(horizontal="right", vertical="center")
        cm.border = thin_border

        # N: Tổng lãi phạt = QH + CT
        cn = ws.cell(row=r, column=14)
        cn.value = f'=IF(J{r}<>"",J{r}+K{r},"")'
        cn.number_format = FMT_VND
        cn.fill = make_fill(C_OVERDUE_BG)
        cn.font = make_font(bold=True, size=10, color="C00000")
        cn.alignment = Alignment(horizontal="right", vertical="center")
        cn.border = thin_border

    # Blank rows
    last_data = 4 + len(sample_periods)
    for i in range(20):
        r = last_data + i
        ws.row_dimensions[r].height = 18
        ws.cell(row=r, column=1, value=len(sample_periods) + i + 1)
        ws.cell(row=r, column=1).border = thin_border
        ws.cell(row=r, column=1).fill = make_fill(C_WHITE)
        ws.cell(row=r, column=1).font = make_font(size=10)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        style_input(ws, r, 2, fmt=FMT_DATE)
        style_input(ws, r, 3, fmt=FMT_DATE)
        cd = ws.cell(row=r, column=4)
        cd.value = f"=IF(AND(B{r}<>\"\",C{r}<>\"\"),C{r}-B{r}+1,\"\")"
        cd.fill = make_fill(C_CALC_BG)
        cd.border = thin_border
        style_input(ws, r, 5)
        for col in range(6, 15):
            c = ws.cell(row=r, column=col)
            c.fill = make_fill(C_CALC_BG if col != 10 and col != 11 and col != 14 else C_OVERDUE_BG)
            c.border = thin_border
            c.number_format = FMT_VND if col >= 7 else FMT_PCT
        ws.cell(row=r, column=6).value  = f"=IFERROR(INDEX({bt_ls},MATCH(B{r},{bt_dates},0)),\"\")"
        ws.cell(row=r, column=7).value  = f"=IFERROR(INDEX({bt_gocIH},MATCH(B{r},{bt_dates},0)),\"\")"
        ws.cell(row=r, column=8).value  = f"=IFERROR(INDEX({bt_gocQH},MATCH(B{r},{bt_dates},0)),\"\")"
        ws.cell(row=r, column=9).value  = (f'=IF(AND(B{r}<>"",C{r}<>""),'
                                           f'SUMIFS({bt_laiIH},{bt_dates},">="&B{r},{bt_dates},"<="&C{r}),"")')
        ws.cell(row=r, column=10).value = (f'=IF(AND(B{r}<>"",C{r}<>""),'
                                           f'SUMIFS({bt_laiQH},{bt_dates},">="&B{r},{bt_dates},"<="&C{r}),"")')
        ws.cell(row=r, column=11).value = (f'=IF(AND(B{r}<>"",C{r}<>""),'
                                           f'SUMIFS({bt_laiCT},{bt_dates},">="&B{r},{bt_dates},"<="&C{r}),"")')
        ws.cell(row=r, column=12).value = (f'=IF(AND(B{r}<>"",C{r}<>""),'
                                           f'SUMIFS({bt_tt},{bt_dates},">="&B{r},{bt_dates},"<="&C{r}),"")')
        ws.cell(row=r, column=13).value = f'=IF(I{r}<>"",I{r},"")'
        ws.cell(row=r, column=14).value = f'=IF(J{r}<>"",J{r}+K{r},"")'
        ws.cell(row=r, column=6).number_format = FMT_PCT

    # Grand total
    total_row = last_data + 20 + 1
    ws.row_dimensions[total_row].height = 22
    style_total(ws, total_row, 1, "TỔNG CỘNG")
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=5)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="left")
    for col in range(6, 15):
        col_ltr = get_column_letter(col)
        ct = style_total(ws, total_row, col,
                         f"=IFERROR(SUM({col_ltr}4:{col_ltr}{total_row-1}),0)",
                         FMT_VND)
    ws.cell(row=total_row, column=6).number_format = FMT_PCT

    set_print_a4_landscape(ws, "BẢNG TỔNG HỢP THEO KỲ")
    ws.print_title_rows = "1:3"


def build_sheet_tonghop(ws):
    ws.title = "TỔNG HỢP"

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 20

    # Title
    ws.merge_cells("A1:C1")
    c = ws["A1"]
    c.value = "📑 TỔNG HỢP NỢ VAY — TẠI NGÀY TÍNH TOÁN"
    c.font = make_font(bold=True, size=14, color=C_HEADER_TXT)
    c.fill = make_fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:C2")
    c = ws["A2"]
    c.value = ("Căn cứ: Thông tư 39/2016/TT-NHNN (sđ bởi TT 06/2023/TT-NHNN) | "
               "Ngày tính: xem ô B5 trong CẤU HÌNH")
    c.font = make_font(size=9, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    ws.row_dimensions[3].height = 8

    def add_section(ws, row, title):
        ws.row_dimensions[row].height = 18
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c = ws.cell(row=row, column=1, value=f"▶ {title}")
        c.font = make_font(bold=True, size=10, color=C_HEADER_TXT)
        c.fill = make_fill(C_SECTION_BG)
        c.alignment = Alignment(horizontal="left", vertical="center")

    def add_row(ws, row, label, formula, fmt=FMT_VND, bold=False, bg=C_WHITE):
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

        c3 = ws.cell(row=row, column=3)
        c3.fill = make_fill(bg)
        c3.border = thin_border

        return c2

    # Last row of daily engine — updated column letters
    bt_last   = FIRST_ROW + NUM_ROWS - 1
    bt_dates  = f"'BẢNG TÍNH'!$A$5:$A${bt_last}"
    bt_O  = f"'BẢNG TÍNH'!$O$5:$O${bt_last}"   # Lãi IH phát sinh
    bt_P  = f"'BẢNG TÍNH'!$P$5:$P${bt_last}"   # Lãi QH phát sinh
    bt_Q  = f"'BẢNG TÍNH'!$Q$5:$Q${bt_last}"   # Lãi CT phát sinh
    bt_R  = f"'BẢNG TÍNH'!$R$5:$R${bt_last}"   # Phân bổ → Gốc QH
    bt_S  = f"'BẢNG TÍNH'!$S$5:$S${bt_last}"   # Phân bổ → Lãi QH
    bt_T  = f"'BẢNG TÍNH'!$T$5:$T${bt_last}"   # Phân bổ → Gốc IH
    bt_U  = f"'BẢNG TÍNH'!$U$5:$U${bt_last}"   # Phân bổ → Lãi IH
    bt_V  = f"'BẢNG TÍNH'!$V$5:$V${bt_last}"   # Phân bổ → Lãi CT

    # Dynamic lookup for balances at calculation date (CẤU HÌNH!B5)
    bt_W_lookup  = f"IFERROR(INDEX('BẢNG TÍNH'!$W$5:$W${bt_last},MATCH('CẤU HÌNH'!$B$5,{bt_dates},0)),0)"
    bt_X_lookup  = f"IFERROR(INDEX('BẢNG TÍNH'!$X$5:$X${bt_last},MATCH('CẤU HÌNH'!$B$5,{bt_dates},0)),0)"
    bt_Y_lookup  = f"IFERROR(INDEX('BẢNG TÍNH'!$Y$5:$Y${bt_last},MATCH('CẤU HÌNH'!$B$5,{bt_dates},0)),0)"
    bt_Z_lookup  = f"IFERROR(INDEX('BẢNG TÍNH'!$Z$5:$Z${bt_last},MATCH('CẤU HÌNH'!$B$5,{bt_dates},0)),0)"
    bt_AA_lookup = f"IFERROR(INDEX('BẢNG TÍNH'!$AA$5:$AA${bt_last},MATCH('CẤU HÌNH'!$B$5,{bt_dates},0)),0)"

    r = 4
    add_section(ws, r, "THÔNG TIN KHOẢN VAY")

    r += 1
    c = add_row(ws, r, "Ngày tính toán", "='CẤU HÌNH'!B5", FMT_DATE)

    r += 1
    c = add_row(ws, r, "Tổng số tiền giải ngân",
                "=SUM('GIẢI NGÂN'!C4:C20)", FMT_VND)

    r += 1
    c = add_row(ws, r, "Ngày giải ngân đầu tiên",
                "=MIN('GIẢI NGÂN'!B4:B20)", FMT_DATE)

    r += 1
    c = add_row(ws, r, "Ngày giải ngân cuối cùng",
                "=IF(COUNTA('GIẢI NGÂN'!B4:B20)>0,MAX('GIẢI NGÂN'!B4:B20),\"\")", FMT_DATE)

    r += 1
    ws.row_dimensions[r].height = 6  # spacer

    r += 1
    add_section(ws, r, "TỔNG ĐÃ THANH TOÁN (lũy kế đến ngày tính)")

    r += 1
    add_row(ws, r, "Tổng gốc đã trả (QH + IH)",
            f"=SUM({bt_R})+SUM({bt_T})", FMT_VND)

    r += 1
    add_row(ws, r, "Tổng lãi trong hạn đã trả",
            f"=SUM({bt_U})", FMT_VND)

    r += 1
    add_row(ws, r, "Tổng lãi trên Gốc QH đã trả",
            f"=SUM({bt_S})", FMT_VND, bg=C_OVERDUE_BG)

    r += 1
    add_row(ws, r, "Tổng lãi chậm trả đã trả",
            f"=SUM({bt_V})", FMT_VND, bg=C_OVERDUE_BG)

    r += 1
    add_row(ws, r, "TỔNG ĐÃ THANH TOÁN",
            f"=SUM({bt_R})+SUM({bt_T})+SUM({bt_U})+SUM({bt_S})+SUM({bt_V})",
            FMT_VND, bold=True, bg=C_TOTAL_BG)

    r += 1
    ws.row_dimensions[r].height = 6  # spacer

    r += 1
    r += 1
    add_section(ws, r, "LÃI PHÁT SINH LŨY KẾ (Bóc tách theo chuẩn Core Banking)")

    r += 1
    add_row(ws, r, "1. Lãi cơ bản trên Gốc trong hạn (100% LS trong hạn)",
            f"=SUM({bt_O})", FMT_VND)

    r += 1
    add_row(ws, r, "2. Lãi cơ bản trên Gốc quá hạn (100% LS trong hạn)",
            f"=SUM({bt_P})/'CẤU HÌNH'!$B$11", FMT_VND)

    r += 1
    add_row(ws, r, "▶ TỔNG LÃI CƠ BẢN PHÁT SINH (100% LS Trong hạn)",
            f"=SUM({bt_O})+(SUM({bt_P})/'CẤU HÌNH'!$B$11)", FMT_VND, bold=True, bg="EBF3F9")

    r += 1
    add_row(ws, r, "3. Phạt quá hạn tăng thêm (50% chênh lệch LS quá hạn)",
            f"=SUM({bt_P})-(SUM({bt_P})/'CẤU HÌNH'!$B$11)", FMT_VND, bg=C_OVERDUE_BG)

    r += 1
    add_row(ws, r, "4. Phạt chậm trả lãi (≤ 10%/năm trên lãi quá hạn)",
            f"=SUM({bt_Q})", FMT_VND, bg=C_OVERDUE_BG)

    r += 1
    add_row(ws, r, "▶ TỔNG LÃI PHẠT PHÁT SINH (Phạt QH 50% + Chậm trả 10%)",
            f"=(SUM({bt_P})-(SUM({bt_P})/'CẤU HÌNH'!$B$11))+SUM({bt_Q})", FMT_VND, bold=True, bg=C_WARN_BG)

    r += 1
    add_row(ws, r, "★ TỔNG LÃI VÀ PHẠT PHÁT SINH TOÀN BỘ ★",
            f"=SUM({bt_O})+SUM({bt_P})+SUM({bt_Q})", FMT_VND, bold=True, bg=C_TOTAL_BG)

    r += 1
    ws.row_dimensions[r].height = 6  # spacer

    r += 1
    add_section(ws, r, f"DƯ NỢ TẠI NGÀY TÍNH (='CẤU HÌNH'!B5)")

    r += 1
    add_row(ws, r, "Dư nợ gốc trong hạn",
            f"={bt_W_lookup}", FMT_VND)

    r += 1
    add_row(ws, r, "Dư nợ gốc quá hạn",
            f"={bt_X_lookup}", FMT_VND, bg=C_OVERDUE_BG)

    r += 1
    add_row(ws, r, "TỔNG DƯ NỢ GỐC",
            f"={bt_W_lookup}+{bt_X_lookup}", FMT_VND, bold=True, bg="EBF3F9")

    r += 1
    ws.row_dimensions[r].height = 6  # spacer

    r += 1
    add_row(ws, r, "Lãi cơ bản chưa thanh toán (100% LS trong hạn)",
            f"={bt_Y_lookup}+({bt_Z_lookup}/'CẤU HÌNH'!$B$11)", FMT_VND)

    r += 1
    add_row(ws, r, "Phạt quá hạn chưa thanh toán (50% chênh lệch LS quá hạn)",
            f"={bt_Z_lookup}-({bt_Z_lookup}/'CẤU HÌNH'!$B$11)", FMT_VND, bg=C_OVERDUE_BG)

    r += 1
    add_row(ws, r, "Phạt chậm trả chưa thanh toán (≤ 10%/năm)",
            f"={bt_AA_lookup}", FMT_VND, bg=C_OVERDUE_BG)

    r += 1
    add_row(ws, r, "TỔNG LÃI VÀ PHẠT CHƯA THANH TOÁN",
            f"={bt_Y_lookup}+{bt_Z_lookup}+{bt_AA_lookup}", FMT_VND, bold=True, bg=C_WARN_BG)

    r += 1
    ws.row_dimensions[r].height = 6  # spacer

    r += 1
    total_formula = f"={bt_W_lookup}+{bt_X_lookup}+{bt_Y_lookup}+{bt_Z_lookup}+{bt_AA_lookup}"
    c_total = add_row(ws, r, "★ TỔNG SỐ TIỀN CÒN NỢ TẠI NGÀY TÍNH ★",
                      total_formula, FMT_VND, bold=True, bg="C00000")
    c_total.font = make_font(bold=True, size=12, color=C_HEADER_TXT)
    ws.cell(row=r, column=1).font = make_font(bold=True, size=12, color=C_HEADER_TXT)
    ws.row_dimensions[r].height = 26

    r += 2
    ws.row_dimensions[r].height = 10
    # Legal disclaimer
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(row=r, column=1,
                value=("Ghi chú Core Banking & Pháp lý: LS quá hạn = 150% LS trong hạn (bao gồm 100% lãi cơ bản + 50% phạt quá hạn tăng thêm theo Đ13 TT39/2016). "
                       "Lãi chậm trả ≤ 10%/năm trên nợ lãi IH quá hạn (Đ13 TT39/2016). "
                       "Thứ tự thu nợ theo TT06/2023: Gốc QH → Lãi trên Gốc QH → Gốc IH đến hạn → Lãi IH → Lãi chậm trả."))
    c.font = make_font(size=8, color="595959")
    c.fill = make_fill(C_WARN_BG)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border = thin_border
    ws.row_dimensions[r].height = 36

    set_print_a4_landscape(ws, "TỔNG HỢP NỢ VAY")


# ─────────────────────────────────────────────
# MAIN BUILD
# ─────────────────────────────────────────────

def build_file1():
    try:
        from dateutil.relativedelta import relativedelta
    except ImportError:
        import subprocess, sys
        subprocess.check_call([sys.executable, "-m", "pip", "install", "python-dateutil"])
        from dateutil.relativedelta import relativedelta

    wb = Workbook()
    # Remove default sheet
    default_ws = wb.active
    wb.remove(default_ws)

    print("Building CẤU HÌNH...")
    ws_cfg = wb.create_sheet("CẤU HÌNH")
    build_sheet_cauhinh(ws_cfg)

    print("Building NGÀY NGHỈ LỄ...")
    ws_hol = wb.create_sheet("NGÀY NGHỈ LỄ")
    build_sheet_holidays(ws_hol)

    print("Building GIẢI NGÂN...")
    ws_gn = wb.create_sheet("GIẢI NGÂN")
    build_sheet_giaingân(ws_gn)

    print("Building LỊCH TRẢ NỢ...")
    ws_lich = wb.create_sheet("LỊCH TRẢ NỢ")
    build_sheet_lichtrâno(ws_lich)

    print("Building THAY ĐỔI LÃI SUẤT...")
    ws_ls = wb.create_sheet("THAY ĐỔI LÃI SUẤT")
    build_sheet_laisu(ws_ls)

    print("Building THANH TOÁN THỰC TẾ...")
    ws_tt = wb.create_sheet("THANH TOÁN THỰC TẾ")
    build_sheet_thanhtoan(ws_tt)

    print(f"Building BẢNG TÍNH (dynamic daily engine — up to {NUM_ROWS} rows)...")
    ws_bt = wb.create_sheet("BẢNG TÍNH")
    build_sheet_bangtính(ws_bt)

    print("Building TỔNG HỢP KỲ...")
    ws_thk = wb.create_sheet("TỔNG HỢP KỲ")
    build_sheet_tonghopky(ws_thk)

    print("Building TỔNG HỢP...")
    ws_th = wb.create_sheet("TỔNG HỢP")
    build_sheet_tonghop(ws_th)

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True

    out_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "TinhLai_ChiTiet.xlsx"
    )
    wb.save(out_path)
    print(f"\n✅ Saved: {os.path.abspath(out_path)}")
    return out_path


if __name__ == "__main__":
    build_file1()
