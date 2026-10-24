import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

# ── Styles ──────────────────────────────────────────────────────────────────
WHITE   = "FFFFFFFF"
DARK    = "FF1A2A1A"
GOLD    = "FFC9A84C"
LGOLD   = "FFFFF8E8"
GREEN   = "FF2D6A4F"
LGREEN  = "FFE8F5EE"
GREY    = "FFF0F2EE"
RED     = "FFC03030"
AMBER   = "FFB06820"
BLACK   = "FF000000"
LBLUE   = "FFE8F0FF"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color=BLACK, size=11, name="Arial"):
    return Font(bold=bold, color=color, size=size, name=name)

def border_bottom():
    return Border(bottom=Side(style="thin", color="FFD0D8C8"))

def border_all():
    s = Side(style="thin", color="FFD0D8C8")
    return Border(top=s, bottom=s, left=s, right=s)

def money(ws, cell, value):
    ws[cell] = value
    ws[cell].number_format = '$#,##0.00'

def pct(ws, cell, value):
    ws[cell] = value
    ws[cell].number_format = '0.0%'

def hdr(ws, cell, text, bg=DARK, fg=WHITE, sz=11, bold=True):
    ws[cell] = text
    ws[cell].font = Font(bold=bold, color=fg, size=sz, name="Arial")
    ws[cell].fill = fill(bg)
    ws[cell].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def label(ws, cell, text, bold=False, indent=0):
    ws[cell] = text
    ws[cell].font = Font(bold=bold, color=BLACK, size=10, name="Arial")
    ws[cell].alignment = Alignment(indent=indent)

MONTHS_2026 = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
MONTHS_2027 = MONTHS_2026[:]

# ── Workbook ─────────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()
wb.remove(wb.active)

# ============================================================
# TAB 1: Annual Overview
# ============================================================
ov = wb.create_sheet("Annual Overview")
ov.sheet_view.showGridLines = False
ov.column_dimensions['A'].width = 28
for col in ['B','C']:
    ov.column_dimensions[col].width = 18

# Title
ov.merge_cells('A1:C1')
ov['A1'] = "Mobile Strokes Golf — Financial Overview"
ov['A1'].font = Font(bold=True, color=WHITE, size=14, name="Arial")
ov['A1'].fill = fill(DARK)
ov['A1'].alignment = Alignment(horizontal="center", vertical="center")
ov.row_dimensions[1].height = 36

ov.merge_cells('A2:C2')
ov['A2'] = "CFO Dashboard | 2026 & 2027"
ov['A2'].font = Font(bold=False, color=GOLD[2:], size=10, name="Arial")
ov['A2'].fill = fill(DARK)
ov['A2'].alignment = Alignment(horizontal="center")
ov.row_dimensions[2].height = 20

ov.row_dimensions[3].height = 12

# Column headers
for c, txt in zip(['A','B','C'], ['Metric','2026','2027']):
    hdr(ov, f"{c}4", txt, bg=GOLD[2:] if c == 'A' else "FF1A2A1A")

ov.row_dimensions[4].height = 24

# Data rows
rows = [
    ("REVENUE", None, None, True, LGOLD),
    ("Total Revenue", "='2026 Monthly'!P5", "='2027 Monthly'!P5", False, WHITE),
    ("Events Booked", "='2026 Monthly'!P6", "='2027 Monthly'!P6", False, WHITE),
    ("Avg Revenue / Event", "=IF(B6=0,\"-\",B5/B6)", "=IF(C6=0,\"-\",C5/C6)", False, WHITE),
    ("", None, None, False, WHITE),
    ("EXPENSES", None, None, True, LGOLD),
    ("Total Fixed Costs", "='2026 Monthly'!P14", "='2027 Monthly'!P14", False, WHITE),
    ("Total Variable Costs", "='2026 Monthly'!P20", "='2027 Monthly'!P20", False, WHITE),
    ("Total Expenses", "=B11+B12", "=C11+C12", False, WHITE),
    ("", None, None, False, WHITE),
    ("PROFIT", None, None, True, LGOLD),
    ("Gross Profit", "=B5-B13", "=C5-C13", False, WHITE),
    ("Profit Margin", "=IF(B5=0,\"-\",B16/B5)", "=IF(C5=0,\"-\",C16/C5)", False, WHITE),
    ("", None, None, False, WHITE),
    ("DEBT", None, None, True, LGOLD),
    ("Starting Debt", 5824, 0, False, WHITE),
    ("Debt Repaid", "='Debt Tracker'!C18", "='Debt Tracker'!C19", False, WHITE),
    ("Remaining Debt", "=B20-B21", "=C20-C21", False, WHITE),
]

for i, (lbl, v2026, v2027, bold, bg) in enumerate(rows, start=5):
    r = i
    ov.row_dimensions[r].height = 20
    label(ov, f"A{r}", lbl, bold=bold)
    ov[f"A{r}"].fill = fill(bg[2:] if bg.startswith("FF") else bg)
    if bold:
        ov[f"A{r}"].font = Font(bold=True, color=BLACK, size=10, name="Arial")
        ov.merge_cells(f"A{r}:C{r}")
        ov[f"A{r}"].fill = fill(LGOLD[2:])
        continue
    for col, val in zip(['B','C'], [v2026, v2027]):
        if val is not None:
            ov[f"{col}{r}"] = val
            ov[f"{col}{r}"].font = Font(color=BLACK, size=10, name="Arial")
            ov[f"{col}{r}"].alignment = Alignment(horizontal="right")
            if isinstance(val, (int, float)):
                ov[f"{col}{r}"].number_format = '$#,##0.00'
            elif isinstance(val, str) and val.startswith("="):
                # Detect if it's a percentage row
                if "Margin" in lbl:
                    ov[f"{col}{r}"].number_format = '0.0%'
                elif "Events" in lbl or "Avg" in lbl:
                    ov[f"{col}{r}"].number_format = '0'
                else:
                    ov[f"{col}{r}"].number_format = '$#,##0.00'

# Legend
ov.row_dimensions[25].height = 8
ov['A26'] = "Color Legend:"
ov['A26'].font = Font(bold=True, size=9, name="Arial", color=BLACK)
items = [("Blue text = input cells (you fill these in)", LBLUE),
         ("Yellow fill = section headers", LGOLD),
         ("White = calculated formulas", WHITE)]
for i, (txt, bg) in enumerate(items, 27):
    ov[f"A{i}"] = txt
    ov[f"A{i}"].font = Font(size=9, name="Arial", color=BLACK)
    ov[f"A{i}"].fill = fill(bg[2:])
    ov.row_dimensions[i].height = 16

# ============================================================
# HELPER: build a monthly tab
# ============================================================
def build_monthly(wb, sheet_name, year):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    months = MONTHS_2026 if year == 2026 else MONTHS_2027
    col_letters = [get_column_letter(i) for i in range(2, 14)]  # B–M
    total_col = 'P'  # column for totals (skip N, O for spacing)

    # Column widths
    ws.column_dimensions['A'].width = 30
    for cl in col_letters:
        ws.column_dimensions[cl].width = 10
    ws.column_dimensions['N'].width = 4
    ws.column_dimensions['O'].width = 4
    ws.column_dimensions[total_col].width = 14

    # Title
    ws.merge_cells(f'A1:{total_col}1')
    ws['A1'] = f"Mobile Strokes Golf — {year} Monthly Tracker"
    ws['A1'].font = Font(bold=True, color=WHITE, size=13, name="Arial")
    ws['A1'].fill = fill(DARK[2:])
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.row_dimensions[2].height = 8

    # Month headers row 3
    ws['A3'] = "Category"
    ws['A3'].font = Font(bold=True, color=WHITE, size=10, name="Arial")
    ws['A3'].fill = fill(DARK[2:])
    ws['A3'].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[3].height = 22

    for cl, mo in zip(col_letters, months):
        hdr(ws, f"{cl}3", mo)
    ws[f"{total_col}3"] = "TOTAL"
    ws[f"{total_col}3"].font = Font(bold=True, color=WHITE, size=10, name="Arial")
    ws[f"{total_col}3"].fill = fill(GOLD[2:])
    ws[f"{total_col}3"].alignment = Alignment(horizontal="center")

    # ── REVENUE ──────────────────────────────────────────────
    ws.row_dimensions[4].height = 8
    ws.merge_cells(f'A5:{total_col}5')
    ws['A5'] = "REVENUE"
    ws['A5'].font = Font(bold=True, color=BLACK, size=10, name="Arial")
    ws['A5'].fill = fill(LGOLD[2:])

    # Revenue row 5 → actual row 5 is section header; data in row 5 = Total Revenue
    # Let's replan rows:
    # Row 5 = REVENUE header
    # Row 6 = Total Revenue (input, blue)
    # Row 7 = Events Booked (input, blue)
    # Row 8 = blank
    # Row 9 = FIXED COSTS header
    # Row 10 = Netlify $20
    # Row 11 = Claude AI $20
    # Row 12 = Insurance (input)
    # Row 13 = Other Fixed (input)
    # Row 14 = Total Fixed
    # Row 15 = blank
    # Row 16 = VARIABLE COSTS header
    # Row 17 = Fuel / Travel (input)
    # Row 18 = Supplies (input)
    # Row 19 = Marketing / Ads (input)
    # Row 20 = Total Variable
    # Row 21 = blank
    # Row 22 = NET PROFIT header
    # Row 23 = Gross Profit (formula)
    # Row 24 = Profit Margin (formula)

    # Re-do section header properly
    ws.row_dimensions[5].height = 22

    # Revenue inputs
    for r, lbl, fixed_val, is_input in [
        (6, "Total Revenue ($)", None, True),
        (7, "Events Booked (#)", None, True),
    ]:
        ws[f"A{r}"] = lbl
        ws[f"A{r}"].font = Font(size=10, name="Arial", color=BLACK)
        ws.row_dimensions[r].height = 20
        for cl in col_letters:
            ws[f"{cl}{r}"] = 0 if fixed_val is None else fixed_val
            ws[f"{cl}{r}"].font = Font(color="000000FF", size=10, name="Arial")  # blue = input
            ws[f"{cl}{r}"].fill = fill(LBLUE[2:])
            ws[f"{cl}{r}"].number_format = '$#,##0.00' if "Revenue" in lbl else '0'
        b2m = f"B{r}:M{r}"
        ws[f"{total_col}{r}"] = f"=SUM({b2m})"
        ws[f"{total_col}{r}"].number_format = '$#,##0.00' if "Revenue" in lbl else '0'
        ws[f"{total_col}{r}"].font = Font(bold=True, size=10, name="Arial")

    ws.row_dimensions[8].height = 8

    # Fixed costs section
    ws.merge_cells(f'A9:{total_col}9')
    ws['A9'] = "FIXED COSTS"
    ws['A9'].font = Font(bold=True, color=BLACK, size=10, name="Arial")
    ws['A9'].fill = fill(LGOLD[2:])
    ws.row_dimensions[9].height = 22

    fixed_items = [
        (10, "Netlify Hosting", 20, False),
        (11, "Claude AI", 20, False),
        (12, "Insurance", None, True),
        (13, "Other Fixed Costs", None, True),
    ]
    for r, lbl, val, is_input in fixed_items:
        ws[f"A{r}"] = lbl
        ws[f"A{r}"].font = Font(size=10, name="Arial", color=BLACK, italic=(not is_input))
        ws.row_dimensions[r].height = 20
        for cl in col_letters:
            ws[f"{cl}{r}"] = val if val is not None else 0
            ws[f"{cl}{r}"].number_format = '$#,##0.00'
            if is_input:
                ws[f"{cl}{r}"].font = Font(color="000000FF", size=10, name="Arial")
                ws[f"{cl}{r}"].fill = fill(LBLUE[2:])
            else:
                ws[f"{cl}{r}"].font = Font(size=10, name="Arial")
        ws[f"{total_col}{r}"] = f"=SUM(B{r}:M{r})"
        ws[f"{total_col}{r}"].number_format = '$#,##0.00'
        ws[f"{total_col}{r}"].font = Font(bold=True, size=10, name="Arial")

    # Total Fixed
    ws[f"A14"] = "Total Fixed Costs"
    ws[f"A14"].font = Font(bold=True, size=10, name="Arial")
    ws.row_dimensions[14].height = 20
    for cl in col_letters:
        ws[f"{cl}14"] = f"=SUM({cl}10:{cl}13)"
        ws[f"{cl}14"].number_format = '$#,##0.00'
        ws[f"{cl}14"].font = Font(bold=True, size=10, name="Arial")
        ws[f"{cl}14"].fill = fill(GREY[2:])
    ws[f"{total_col}14"] = f"=SUM(B14:M14)"
    ws[f"{total_col}14"].number_format = '$#,##0.00'
    ws[f"{total_col}14"].font = Font(bold=True, size=10, name="Arial")
    ws[f"{total_col}14"].fill = fill(GREY[2:])

    ws.row_dimensions[15].height = 8

    # Variable costs
    ws.merge_cells(f'A16:{total_col}16')
    ws['A16'] = "VARIABLE COSTS"
    ws['A16'].font = Font(bold=True, color=BLACK, size=10, name="Arial")
    ws['A16'].fill = fill(LGOLD[2:])
    ws.row_dimensions[16].height = 22

    var_items = [
        (17, "Fuel / Travel"),
        (18, "Supplies / Consumables"),
        (19, "Marketing & Ads"),
    ]
    for r, lbl in var_items:
        ws[f"A{r}"] = lbl
        ws[f"A{r}"].font = Font(size=10, name="Arial")
        ws.row_dimensions[r].height = 20
        for cl in col_letters:
            ws[f"{cl}{r}"] = 0
            ws[f"{cl}{r}"].number_format = '$#,##0.00'
            ws[f"{cl}{r}"].font = Font(color="000000FF", size=10, name="Arial")
            ws[f"{cl}{r}"].fill = fill(LBLUE[2:])
        ws[f"{total_col}{r}"] = f"=SUM(B{r}:M{r})"
        ws[f"{total_col}{r}"].number_format = '$#,##0.00'
        ws[f"{total_col}{r}"].font = Font(bold=True, size=10, name="Arial")

    ws[f"A20"] = "Total Variable Costs"
    ws[f"A20"].font = Font(bold=True, size=10, name="Arial")
    ws.row_dimensions[20].height = 20
    for cl in col_letters:
        ws[f"{cl}20"] = f"=SUM({cl}17:{cl}19)"
        ws[f"{cl}20"].number_format = '$#,##0.00'
        ws[f"{cl}20"].font = Font(bold=True, size=10, name="Arial")
        ws[f"{cl}20"].fill = fill(GREY[2:])
    ws[f"{total_col}20"] = f"=SUM(B20:M20)"
    ws[f"{total_col}20"].number_format = '$#,##0.00'
    ws[f"{total_col}20"].font = Font(bold=True, size=10, name="Arial")
    ws[f"{total_col}20"].fill = fill(GREY[2:])

    ws.row_dimensions[21].height = 8

    # Net profit section
    ws.merge_cells(f'A22:{total_col}22')
    ws['A22'] = "NET PROFIT"
    ws['A22'].font = Font(bold=True, color=BLACK, size=10, name="Arial")
    ws['A22'].fill = fill(LGREEN[2:])
    ws.row_dimensions[22].height = 22

    ws[f"A23"] = "Gross Profit"
    ws[f"A23"].font = Font(bold=True, size=10, name="Arial")
    ws.row_dimensions[23].height = 20
    for cl in col_letters:
        ws[f"{cl}23"] = f"={cl}6-{cl}14-{cl}20"
        ws[f"{cl}23"].number_format = '$#,##0.00'
        ws[f"{cl}23"].font = Font(bold=True, size=10, name="Arial")
        ws[f"{cl}23"].fill = fill(LGREEN[2:])
    ws[f"{total_col}23"] = f"=SUM(B23:M23)"
    ws[f"{total_col}23"].number_format = '$#,##0.00'
    ws[f"{total_col}23"].font = Font(bold=True, size=10, name="Arial")
    ws[f"{total_col}23"].fill = fill(LGREEN[2:])

    ws[f"A24"] = "Profit Margin %"
    ws[f"A24"].font = Font(size=10, name="Arial")
    ws.row_dimensions[24].height = 20
    for cl in col_letters:
        ws[f"{cl}24"] = f"=IF({cl}6=0,0,{cl}23/{cl}6)"
        ws[f"{cl}24"].number_format = '0.0%'
        ws[f"{cl}24"].fill = fill(LGREEN[2:])
    ws[f"{total_col}24"] = f"=IF({total_col}6=0,0,{total_col}23/{total_col}6)"
    ws[f"{total_col}24"].number_format = '0.0%'
    ws[f"{total_col}24"].font = Font(bold=True, size=10, name="Arial")
    ws[f"{total_col}24"].fill = fill(LGREEN[2:])

    # Freeze panes
    ws.freeze_panes = 'B4'

    return ws

build_monthly(wb, "2026 Monthly", 2026)
build_monthly(wb, "2027 Monthly", 2027)

# ============================================================
# TAB 4: Debt Tracker
# ============================================================
dt = wb.create_sheet("Debt Tracker")
dt.sheet_view.showGridLines = False
dt.column_dimensions['A'].width = 22
dt.column_dimensions['B'].width = 18
dt.column_dimensions['C'].width = 16
dt.column_dimensions['D'].width = 16
dt.column_dimensions['E'].width = 16
dt.column_dimensions['F'].width = 20

dt.merge_cells('A1:F1')
dt['A1'] = "Mobile Strokes Golf — Debt Tracker"
dt['A1'].font = Font(bold=True, color=WHITE, size=13, name="Arial")
dt['A1'].fill = fill(DARK[2:])
dt['A1'].alignment = Alignment(horizontal="center", vertical="center")
dt.row_dimensions[1].height = 32

dt.row_dimensions[2].height = 8

for c, txt in zip(['A','B','C','D','E','F'],
                  ['Creditor','Original Amount','Amount Paid','Balance Remaining','Monthly Payment','Status']):
    hdr(dt, f"{c}3", txt)
dt.row_dimensions[3].height = 24

debts = [
    ("Pipo (Equipment Loan)", 4250.00, 0, "=B4-C4", 0, ""),
    ("Inflatable Balance",    1574.00, 0, "=B5-C5", 0, ""),
]
for i, (creditor, orig, paid, bal, mo_pmt, status) in enumerate(debts, start=4):
    dt.row_dimensions[i].height = 22
    dt[f"A{i}"] = creditor
    dt[f"A{i}"].font = Font(size=10, name="Arial")
    dt[f"B{i}"] = orig
    dt[f"B{i}"].number_format = '$#,##0.00'
    dt[f"B{i}"].font = Font(size=10, name="Arial")
    dt[f"C{i}"] = paid
    dt[f"C{i}"].number_format = '$#,##0.00'
    dt[f"C{i}"].font = Font(color="000000FF", size=10, name="Arial")
    dt[f"C{i}"].fill = fill(LBLUE[2:])
    dt[f"D{i}"] = bal
    dt[f"D{i}"].number_format = '$#,##0.00'
    dt[f"D{i}"].font = Font(bold=True, size=10, name="Arial")
    dt[f"E{i}"] = mo_pmt
    dt[f"E{i}"].number_format = '$#,##0.00'
    dt[f"E{i}"].font = Font(color="000000FF", size=10, name="Arial")
    dt[f"E{i}"].fill = fill(LBLUE[2:])
    dt[f"F{i}"] = "Active"
    dt[f"F{i}"].font = Font(size=10, name="Arial", color="FFB06820")

# Totals
dt.row_dimensions[6].height = 8
dt[f"A7"] = "TOTAL DEBT"
dt[f"A7"].font = Font(bold=True, size=10, name="Arial")
dt[f"A7"].fill = fill(LGOLD[2:])
dt[f"B7"] = "=SUM(B4:B5)"
dt[f"B7"].number_format = '$#,##0.00'
dt[f"B7"].font = Font(bold=True, size=10, name="Arial")
dt[f"B7"].fill = fill(LGOLD[2:])
dt[f"C7"] = "=SUM(C4:C5)"
dt[f"C7"].number_format = '$#,##0.00'
dt[f"C7"].font = Font(bold=True, size=10, name="Arial")
dt[f"C7"].fill = fill(LGOLD[2:])
dt[f"D7"] = "=SUM(D4:D5)"
dt[f"D7"].number_format = '$#,##0.00'
dt[f"D7"].font = Font(bold=True, size=10, name="Arial")
dt[f"D7"].fill = fill(LGOLD[2:])

# Total repaid by year (referenced by Annual Overview)
dt.row_dimensions[17].height = 8
dt['A18'] = "Total Debt Repaid 2026"
dt['A18'].font = Font(size=10, name="Arial")
dt['C18'] = "=C7"
dt['C18'].number_format = '$#,##0.00'
dt['A19'] = "Total Debt Repaid 2027"
dt['A19'].font = Font(size=10, name="Arial")
dt['C19'] = 0
dt['C19'].number_format = '$#,##0.00'
dt['C19'].font = Font(color="000000FF", size=10, name="Arial")
dt['C19'].fill = fill(LBLUE[2:])

# Notes
dt.row_dimensions[22].height = 8
dt['A23'] = "Notes"
dt['A23'].font = Font(bold=True, size=10, name="Arial")
dt['A24'] = "Blue cells = enter your data here. Update 'Amount Paid' as debts are repaid."
dt['A24'].font = Font(size=9, name="Arial", color=BLACK)
dt['A25'] = "Pipo loan: $4,250 covered equipment bundle (SkyTrak+, BenQ, mat, turf, screen, assembly)"
dt['A25'].font = Font(size=9, name="Arial", color=BLACK)
dt['A26'] = "Inflatable: Total $2,190 | Deposit paid $616 | Balance $1,574"
dt['A26'].font = Font(size=9, name="Arial", color=BLACK)

# ============================================================
# Save
# ============================================================
out = r"C:\Users\Mobile Strokes\Desktop\Claude\MSG_CFO_Workbook.xlsx"
wb.save(out)
print(f"Saved: {out}")
