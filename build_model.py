"""
Builds property_underwriting_model.xlsx for the fictional "Project Meridian Ridge"
40-unit weekend-home development underwriting model.

ALL figures are synthetic / illustrative. This is a methodology demonstration only.
Every computed cell is a live Excel formula (openpyxl writes formula strings; the
skill's recalc.py / LibreOffice pass then computes and caches the values so the
file previews correctly, but the formulas themselves remain live for Excel).
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference, Series
from openpyxl.comments import Comment
from openpyxl.formatting.rule import ColorScaleRule

# ---------------------------------------------------------------- constants --
FONT_NAME = "Arial"
BLUE = Font(name=FONT_NAME, color="0000FF")
BLUE_BOLD = Font(name=FONT_NAME, color="0000FF", bold=True)
BLACK = Font(name=FONT_NAME, color="000000")
BLACK_BOLD = Font(name=FONT_NAME, bold=True)
GREEN = Font(name=FONT_NAME, color="008000")
GREEN_BOLD = Font(name=FONT_NAME, color="008000", bold=True)
WHITE_BOLD = Font(name=FONT_NAME, color="FFFFFF", bold=True)
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color="1F3864")
SUBTITLE_FONT = Font(name=FONT_NAME, size=10, italic=True, color="595959")
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
SUBHEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CUR = '$#,##0;($#,##0);"-"'
CUR2 = '$#,##0.00;($#,##0.00);"-"'
PCT1 = '0.0%'
PCT0 = '0%'
NUM0 = '#,##0;(#,##0);"-"'
MULT = '0.00x'

N_Q = 16  # total quarterly periods modeled (4 years)

wb = Workbook()
wb.remove(wb.active)


def sheet(name):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    return ws


def title_block(ws, main, sub, span_col="N"):
    ws.merge_cells(f"A1:{span_col}1")
    ws["A1"] = main
    ws["A1"].font = TITLE_FONT
    ws.merge_cells(f"A2:{span_col}2")
    ws["A2"] = sub
    ws["A2"].font = SUBTITLE_FONT
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 16


def header_row(ws, row, col_start, labels, fill=HEADER_FILL, font=WHITE_BOLD):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col_start + i, value=lab)
        c.font = font
        c.fill = fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER


def label(ws, row, text, col=1, font=BLACK, fill=None, bold=False):
    c = ws.cell(row=row, column=col, value=text)
    c.font = BLACK_BOLD if bold else font
    if fill:
        c.fill = fill
    return c


def qcols(start=3, n=N_Q):
    """Column letters for quarter 1..n starting at given column index."""
    return [get_column_letter(start + i) for i in range(n)]


def set_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def comment(cell, text):
    cell.comment = Comment(text, "Property Underwriting Model")


# ============================================================== ASSUMPTIONS =
ws = sheet("Assumptions")
title_block(ws, "Project Meridian Ridge — Underwriting Assumptions",
            "Fictional 40-unit weekend-home development · all figures synthetic (INR Lakh unless noted) · blue = input cell")
set_col_widths(ws, {"A": 34, "B": 16, "C": 14, "D": 14, "E": 14, "F": 40})

r = 4
label(ws, r, "PROJECT OVERVIEW", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:F{r}")
r += 1
label(ws, r, "Project name"); ws.cell(row=r, column=2, value="Project Meridian Ridge (fictional)").font = BLUE
r += 1
label(ws, r, "Location context"); ws.cell(row=r, column=2, value="Generic Indian hill-station weekend-home micro-market (illustrative)").font = BLUE
r += 1
label(ws, r, "Total units"); c = ws.cell(row=r, column=2, value=40); c.font = BLUE_BOLD; c.number_format = NUM0
UNITS_TOTAL_CELL = f"B{r}"
r += 1
label(ws, r, "Model horizon (quarters)"); c = ws.cell(row=r, column=2, value=N_Q); c.font = BLUE; c.number_format = NUM0
HORIZON_CELL = f"B{r}"
r += 2

label(ws, r, "UNIT MIX & PRICING", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:F{r}")
r += 1
header_row(ws, r, 1, ["Unit Type", "Units", "Avg. Size (sqft)", "Sale Price (INR/sqft)", "Revenue (INR Lakh)", "Notes"], fill=SUBHEAD_FILL, font=BLACK_BOLD)
r += 1
MIX_HEADER_ROW = r - 1
mix_start = r
mix_data = [
    ("Type A – 2BHK Cottage", 20, 1200, 7000, "Entry-level weekend cottage"),
    ("Type B – 3BHK Villa", 15, 1800, 7800, "Mid-tier villa, larger plot frontage"),
    ("Type C – 4BHK Premium Villa", 5, 2500, 8800, "Premium/corner units, private pool option"),
]
for name, units, size, price, note in mix_data:
    ws.cell(row=r, column=1, value=name).font = BLACK
    c = ws.cell(row=r, column=2, value=units); c.font = BLUE; c.number_format = NUM0
    c = ws.cell(row=r, column=3, value=size); c.font = BLUE; c.number_format = NUM0
    c = ws.cell(row=r, column=4, value=price); c.font = BLUE; c.number_format = '#,##0'
    rev = ws.cell(row=r, column=5, value=f"=B{r}*C{r}*D{r}/100000")
    rev.number_format = CUR
    ws.cell(row=r, column=6, value=note).font = Font(name=FONT_NAME, size=9, italic=True, color="808080")
    for col in range(1, 7):
        ws.cell(row=r, column=col).border = BORDER
    r += 1
mix_end = r - 1
r_total = r
label(ws, r_total, "Total / Blended", bold=True)
c = ws.cell(row=r_total, column=2, value=f"=SUM(B{mix_start}:B{mix_end})"); c.font = BLACK_BOLD; c.number_format = NUM0; c.fill = TOTAL_FILL
c = ws.cell(row=r_total, column=3, value=f"=SUMPRODUCT(B{mix_start}:B{mix_end},C{mix_start}:C{mix_end})/B{r_total}"); c.font = BLACK_BOLD; c.number_format = NUM0; c.fill = TOTAL_FILL
c = ws.cell(row=r_total, column=4, value=f"=SUMPRODUCT(B{mix_start}:B{mix_end},D{mix_start}:D{mix_end})/B{r_total}"); c.font = BLACK_BOLD; c.number_format = '#,##0'; c.fill = TOTAL_FILL
c = ws.cell(row=r_total, column=5, value=f"=SUM(E{mix_start}:E{mix_end})"); c.font = BLACK_BOLD; c.number_format = CUR; c.fill = TOTAL_FILL
for col in range(1, 7):
    ws.cell(row=r_total, column=col).fill = TOTAL_FILL
    ws.cell(row=r_total, column=col).border = BORDER
TOTAL_UNITS_CELL = f"B{r_total}"
TOTAL_REVENUE_CELL = f"E{r_total}"
AVG_PRICE_PER_UNIT_ROW = r_total
r = r_total + 2

label(ws, r, "TOTAL SALEABLE AREA (sqft)", bold=True)
c = ws.cell(row=r, column=2, value=f"=SUMPRODUCT(B{mix_start}:B{mix_end},C{mix_start}:C{mix_end})")
c.font = BLACK_BOLD; c.number_format = NUM0
TOTAL_AREA_CELL = f"B{r}"
r += 1
label(ws, r, "Avg. revenue per unit (INR Lakh)", bold=True)
c = ws.cell(row=r, column=2, value=f"={TOTAL_REVENUE_CELL}/{TOTAL_UNITS_CELL}"); c.font = BLACK_BOLD; c.number_format = CUR2
AVG_REV_PER_UNIT_CELL = f"B{r}"
r += 2

label(ws, r, "LAND & CONSTRUCTION INPUTS", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:F{r}")
r += 1
land_area_row = r; label(ws, r, "Land area (acres)"); c = ws.cell(row=r, column=2, value=8); c.font = BLUE; c.number_format = '0.00'
LAND_AREA_CELL = f"B{r}"
r += 1
land_rate_row = r; label(ws, r, "Land rate (INR Lakh / acre)"); c = ws.cell(row=r, column=2, value=70); c.font = BLUE; c.number_format = CUR
LAND_RATE_CELL = f"B{r}"
r += 1
label(ws, r, "Land approvals / development premium (% of land)"); c = ws.cell(row=r, column=2, value=0.08); c.font = BLUE; c.number_format = PCT1
LAND_APPROVAL_PCT_CELL = f"B{r}"
r += 1
label(ws, r, "Construction cost (INR / sqft)"); c = ws.cell(row=r, column=2, value=3600); c.font = BLUE; c.number_format = '#,##0'
CONSTR_COST_SQFT_CELL = f"B{r}"
r += 1
label(ws, r, "Amenities & common infrastructure (% of construction)"); c = ws.cell(row=r, column=2, value=0.12); c.font = BLUE; c.number_format = PCT1
AMENITIES_PCT_CELL = f"B{r}"
r += 1
label(ws, r, "Design / PMC / soft costs (% of construction)"); c = ws.cell(row=r, column=2, value=0.06); c.font = BLUE; c.number_format = PCT1
SOFTCOST_PCT_CELL = f"B{r}"
r += 1
label(ws, r, "Contingency (% of construction + amenities)"); c = ws.cell(row=r, column=2, value=0.08); c.font = BLUE; c.number_format = PCT1
CONTINGENCY_PCT_CELL = f"B{r}"
r += 1
label(ws, r, "Marketing & sales cost (% of revenue)"); c = ws.cell(row=r, column=2, value=0.05); c.font = BLUE; c.number_format = PCT1
MARKETING_PCT_CELL = f"B{r}"
r += 2

label(ws, r, "PHASING & TIMELINE", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:F{r}")
r += 1
label(ws, r, "Land closing quarter"); c = ws.cell(row=r, column=2, value=1); c.font = BLUE; c.number_format = NUM0
LAND_CLOSE_CELL = f"B{r}"
r += 1
label(ws, r, "Construction start quarter"); c = ws.cell(row=r, column=2, value=1); c.font = BLUE; c.number_format = NUM0
CONSTR_START_CELL = f"B{r}"
r += 1
label(ws, r, "Construction duration (quarters)"); c = ws.cell(row=r, column=2, value=8); c.font = BLUE; c.number_format = NUM0
CONSTR_DUR_CELL = f"B{r}"
r += 1
label(ws, r, "Sales launch quarter"); c = ws.cell(row=r, column=2, value=2); c.font = BLUE; c.number_format = NUM0
SALES_START_CELL = f"B{r}"
r += 1
label(ws, r, "Base-case absorption window (quarters)"); c = ws.cell(row=r, column=2, value=10); c.font = BLUE; c.number_format = NUM0
ABS_WINDOW_CELL = f"B{r}"
r += 2

label(ws, r, "FINANCING", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:F{r}")
r += 1
label(ws, r, "Equity contribution (% of total dev. cost, funds first)"); c = ws.cell(row=r, column=2, value=0.35); c.font = BLUE; c.number_format = PCT1
EQUITY_PCT_CELL = f"B{r}"
r += 1
label(ws, r, "Construction debt facility rate (annual)"); c = ws.cell(row=r, column=2, value=0.11); c.font = BLUE; c.number_format = PCT1
DEBT_RATE_ANNUAL_CELL = f"B{r}"
r += 1
label(ws, r, "Construction debt facility rate (quarterly)"); c = ws.cell(row=r, column=2, value=f"=(1+{DEBT_RATE_ANNUAL_CELL})^0.25-1"); c.font = BLACK; c.number_format = '0.000%'
DEBT_RATE_Q_CELL = f"B{r}"
r += 2

label(ws, r, "RETURNS / DISCOUNTING", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:F{r}")
r += 1
label(ws, r, "Target discount rate (annual, for NPV)"); c = ws.cell(row=r, column=2, value=0.15); c.font = BLUE; c.number_format = PCT1
DISC_RATE_ANNUAL_CELL = f"B{r}"
r += 1
label(ws, r, "Target discount rate (quarterly)"); c = ws.cell(row=r, column=2, value=f"=(1+{DISC_RATE_ANNUAL_CELL})^0.25-1"); c.font = BLACK; c.number_format = '0.000%'
DISC_RATE_Q_CELL = f"B{r}"
r += 2

label(ws, r, "BASE-CASE ABSORPTION CURVE (units sold per quarter, Q2–Q11)", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:{get_column_letter(3+9)}{r}")
r += 1
abs_curve_label_row = r
label(ws, r, "Quarter (relative to sales launch)")
abs_curve_cols = qcols(start=2, n=10)
for i, col in enumerate(abs_curve_cols):
    c = ws[f"{col}{r}"]; c.value = i + 1; c.font = BLACK_BOLD; c.alignment = Alignment(horizontal="center")
r += 1
abs_curve_val_row = r
label(ws, r, "Units sold")
curve_vals = [2, 3, 4, 5, 5, 5, 5, 4, 4, 3]
assert sum(curve_vals) == 40
for i, col in enumerate(abs_curve_cols):
    c = ws[f"{col}{r}"]; c.value = curve_vals[i]; c.font = BLUE; c.number_format = NUM0; c.alignment = Alignment(horizontal="center")
ABS_CURVE_ROW = r
ABS_CURVE_FIRST_COL = abs_curve_cols[0]
ABS_CURVE_LAST_COL = abs_curve_cols[-1]
r += 1
label(ws, r, "Check: sums to total units")
c = ws.cell(row=r, column=2, value=f"=SUM({ABS_CURVE_FIRST_COL}{ABS_CURVE_ROW}:{ABS_CURVE_LAST_COL}{ABS_CURVE_ROW})")
c.font = BLACK; c.number_format = NUM0
comment(ws.cell(row=r, column=1), "Should equal total units (40). Base-case absorption curve is a hardcoded input schedule by design — it is the underwriting team's assumed pace of sales, not a derived figure.")

# =========================================================== COST BUILDUP =
ws = sheet("Cost Buildup")
title_block(ws, "Land & Construction Cost Build-up",
            "All figures INR Lakh · black = formula (links to Assumptions tab in green)")
set_col_widths(ws, {"A": 42, "B": 18, "C": 40})

r = 4
label(ws, r, "LAND COSTS", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:C{r}")
r += 1
label(ws, r, "Land cost (area × rate/acre)")
c = ws.cell(row=r, column=2, value=f"=Assumptions!{LAND_AREA_CELL}*Assumptions!{LAND_RATE_CELL}"); c.font = GREEN; c.number_format = CUR
LAND_COST_CELL = f"B{r}"
r += 1
label(ws, r, "Land approvals / development premium")
c = ws.cell(row=r, column=2, value=f"={LAND_COST_CELL}*Assumptions!{LAND_APPROVAL_PCT_CELL}"); c.font = GREEN; c.number_format = CUR
LAND_APPROVAL_CELL = f"B{r}"
r += 1
label(ws, r, "Land Subtotal", bold=True)
c = ws.cell(row=r, column=2, value=f"={LAND_COST_CELL}+{LAND_APPROVAL_CELL}"); c.font = BLACK_BOLD; c.number_format = CUR; c.fill = TOTAL_FILL
LAND_SUBTOTAL_CELL = f"B{r}"
ws.cell(row=r, column=1).fill = TOTAL_FILL
r += 2

label(ws, r, "CONSTRUCTION COSTS", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:C{r}")
r += 1
label(ws, r, "Base construction cost (saleable area × cost/sqft)")
c = ws.cell(row=r, column=2, value=f"=Assumptions!{TOTAL_AREA_CELL}*Assumptions!{CONSTR_COST_SQFT_CELL}/100000"); c.font = GREEN; c.number_format = CUR
CONSTR_BASE_CELL = f"B{r}"
r += 1
label(ws, r, "Amenities & common infrastructure")
c = ws.cell(row=r, column=2, value=f"={CONSTR_BASE_CELL}*Assumptions!{AMENITIES_PCT_CELL}"); c.font = GREEN; c.number_format = CUR
AMENITIES_CELL = f"B{r}"
r += 1
label(ws, r, "Design / PMC / soft costs")
c = ws.cell(row=r, column=2, value=f"={CONSTR_BASE_CELL}*Assumptions!{SOFTCOST_PCT_CELL}"); c.font = GREEN; c.number_format = CUR
SOFTCOST_CELL = f"B{r}"
r += 1
label(ws, r, "Contingency (on construction + amenities)")
c = ws.cell(row=r, column=2, value=f"=({CONSTR_BASE_CELL}+{AMENITIES_CELL})*Assumptions!{CONTINGENCY_PCT_CELL}"); c.font = GREEN; c.number_format = CUR
CONTINGENCY_CELL = f"B{r}"
r += 1
label(ws, r, "Construction Subtotal", bold=True)
c = ws.cell(row=r, column=2, value=f"={CONSTR_BASE_CELL}+{AMENITIES_CELL}+{SOFTCOST_CELL}+{CONTINGENCY_CELL}")
c.font = BLACK_BOLD; c.number_format = CUR; c.fill = TOTAL_FILL
ws.cell(row=r, column=1).fill = TOTAL_FILL
CONSTR_SUBTOTAL_CELL = f"B{r}"
r += 2

label(ws, r, "FINANCING COST", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:C{r}")
r += 1
label(ws, r, "Interest during construction (IDC)")
c = ws.cell(row=r, column=2,
            value=f"=0.5*(1-Assumptions!{EQUITY_PCT_CELL})*{CONSTR_SUBTOTAL_CELL}*Assumptions!{DEBT_RATE_ANNUAL_CELL}*(Assumptions!{CONSTR_DUR_CELL}/4)")
c.font = GREEN; c.number_format = CUR
IDC_CELL = f"B{r}"
comment(ws.cell(row=r, column=1),
        "Simplified IDC approximation: debt-funded share of construction cost, drawn down evenly, "
        "carrying the annual facility rate for half the construction duration (average outstanding balance method). "
        "A full monthly draw-down schedule would refine this but is not required for underwriting-level precision.")
r += 2

label(ws, r, "SALES & MARKETING COST", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:C{r}")
r += 1
label(ws, r, "Marketing & sales cost (% of total revenue)")
c = ws.cell(row=r, column=2, value=f"=Assumptions!{TOTAL_REVENUE_CELL}*Assumptions!{MARKETING_PCT_CELL}"); c.font = GREEN; c.number_format = CUR
MARKETING_COST_CELL = f"B{r}"
r += 2

label(ws, r, "TOTAL DEVELOPMENT COST", fill=HEADER_FILL, font=WHITE_BOLD, bold=True)
c = ws.cell(row=r, column=2, value=f"={LAND_SUBTOTAL_CELL}+{CONSTR_SUBTOTAL_CELL}+{IDC_CELL}+{MARKETING_COST_CELL}")
c.font = WHITE_BOLD; c.number_format = CUR; c.fill = HEADER_FILL
TOTAL_DEV_COST_CELL = f"B{r}"
r += 1
label(ws, r, "Total Revenue (link)")
c = ws.cell(row=r, column=2, value=f"=Assumptions!{TOTAL_REVENUE_CELL}"); c.font = GREEN; c.number_format = CUR
r += 1
label(ws, r, "Gross Margin (Revenue − Total Dev. Cost)", bold=True)
c = ws.cell(row=r, column=2, value=f"=Assumptions!{TOTAL_REVENUE_CELL}-{TOTAL_DEV_COST_CELL}"); c.font = BLACK_BOLD; c.number_format = CUR
GROSS_MARGIN_CELL = f"B{r}"
r += 1
label(ws, r, "Gross Margin %", bold=True)
c = ws.cell(row=r, column=2, value=f"={GROSS_MARGIN_CELL}/Assumptions!{TOTAL_REVENUE_CELL}"); c.font = BLACK_BOLD; c.number_format = PCT1
GROSS_MARGIN_PCT_CELL = f"B{r}"

COST_SHEET = "'Cost Buildup'"

# ================================================== SALES / ABSORPTION ====
ws = sheet("Absorption Schedule")
title_block(ws, "Sales Realization / Absorption Schedule",
            "Quarterly · base-case pace draws on the Assumptions absorption curve · revenue assumes a constant blended mix each quarter")
Q_START_COL = 3  # column C
qc = qcols(Q_START_COL, N_Q)
set_col_widths(ws, {"A": 40, "B": 4})
for col in qc:
    ws.column_dimensions[col].width = 11

r = 4
label(ws, r, "Quarter", bold=True)
for i, col in enumerate(qc):
    c = ws[f"{col}{r}"]; c.value = i + 1; c.font = WHITE_BOLD; c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="center")
QROW = r
r += 1

label(ws, r, "Units sold this quarter")
for i, col in enumerate(qc):
    qn = i + 1
    f = (f"=IF(AND(({col}${QROW}-Assumptions!${SALES_START_CELL[0]}${SALES_START_CELL[1:]}+1)>=1,"
         f"({col}${QROW}-Assumptions!${SALES_START_CELL[0]}${SALES_START_CELL[1:]}+1)<=10),"
         f"INDEX(Assumptions!${ABS_CURVE_FIRST_COL}${ABS_CURVE_ROW}:${ABS_CURVE_LAST_COL}${ABS_CURVE_ROW},"
         f"{col}${QROW}-Assumptions!${SALES_START_CELL[0]}${SALES_START_CELL[1:]}+1),0)")
    c = ws[f"{col}{r}"]; c.value = f; c.number_format = NUM0; c.font = BLACK
UNITS_ROW = r
r += 1

label(ws, r, "Cumulative units sold")
for i, col in enumerate(qc):
    f = f"=SUM(${qc[0]}${UNITS_ROW}:{col}${UNITS_ROW})"
    c = ws[f"{col}{r}"]; c.value = f; c.number_format = NUM0; c.font = BLACK
CUM_UNITS_ROW = r
r += 1

label(ws, r, "Revenue recognized (INR Lakh)")
for i, col in enumerate(qc):
    f = f"={col}${UNITS_ROW}*Assumptions!${AVG_REV_PER_UNIT_CELL[0]}${AVG_REV_PER_UNIT_CELL[1:]}"
    c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = GREEN
REVENUE_ROW = r
r += 1

label(ws, r, "Cumulative revenue (INR Lakh)")
for i, col in enumerate(qc):
    f = f"=SUM(${qc[0]}${REVENUE_ROW}:{col}${REVENUE_ROW})"
    c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = BLACK
CUM_REV_ROW = r
r += 2
label(ws, r, "Check: cumulative units sold at Q16 = total units")
c = ws.cell(row=r, column=3, value=f"={qc[-1]}{CUM_UNITS_ROW}"); c.number_format = NUM0
r += 1
label(ws, r, "Check: cumulative revenue at Q16 = total revenue")
c = ws.cell(row=r, column=3, value=f"={qc[-1]}{CUM_REV_ROW}"); c.number_format = CUR

ABS_SHEET = "'Absorption Schedule'"

wb.save("model/property_underwriting_model.xlsx")
print("Absorption Schedule built. UNITS_ROW", UNITS_ROW, "REVENUE_ROW", REVENUE_ROW, "Q_START_COL", qc[0])

# ============================================================ CASHFLOW ====
ws = sheet("Cashflow")
title_block(ws, "Development Cashflow (Base Case, Unlevered)",
            "Quarterly · INR Lakh · land + construction phased outflows vs. sales inflows · feeds Exit Scenarios & Sensitivity")
set_col_widths(ws, {"A": 40, "B": 4})
for col in qc:
    ws.column_dimensions[col].width = 11

r = 4
label(ws, r, "Quarter", bold=True)
for i, col in enumerate(qc):
    c = ws[f"{col}{r}"]; c.value = i + 1; c.font = WHITE_BOLD; c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="center")
CF_QROW = r
r += 1

label(ws, r, "OUTFLOWS", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:B{r}")
r += 1
label(ws, r, "Land cost")
for i, col in enumerate(qc):
    qn = i + 1
    f = f"=IF({col}${CF_QROW}=Assumptions!{LAND_CLOSE_CELL},{COST_SHEET}!{LAND_SUBTOTAL_CELL},0)"
    c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = GREEN
LAND_OUT_ROW = r
r += 1

label(ws, r, "Construction cost (straight-line over build period)")
for i, col in enumerate(qc):
    f = (f"=IF(AND({col}${CF_QROW}>=Assumptions!{CONSTR_START_CELL},"
         f"{col}${CF_QROW}<=Assumptions!{CONSTR_START_CELL}+Assumptions!{CONSTR_DUR_CELL}-1),"
         f"{COST_SHEET}!{CONSTR_SUBTOTAL_CELL}/Assumptions!{CONSTR_DUR_CELL},0)")
    c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = GREEN
CONSTR_OUT_ROW = r
r += 1

label(ws, r, "Interest during construction (phased)")
for i, col in enumerate(qc):
    f = (f"=IF(AND({col}${CF_QROW}>=Assumptions!{CONSTR_START_CELL},"
         f"{col}${CF_QROW}<=Assumptions!{CONSTR_START_CELL}+Assumptions!{CONSTR_DUR_CELL}-1),"
         f"{COST_SHEET}!{IDC_CELL}/Assumptions!{CONSTR_DUR_CELL},0)")
    c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = GREEN
IDC_OUT_ROW = r
r += 1

label(ws, r, "Marketing & sales cost (proportional to revenue)")
for i, col in enumerate(qc):
    f = f"={ABS_SHEET}!{col}{REVENUE_ROW}*Assumptions!{MARKETING_PCT_CELL}"
    c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = GREEN
MARKETING_OUT_ROW = r
r += 1

label(ws, r, "Total Outflows", bold=True)
for i, col in enumerate(qc):
    f = f"=SUM({col}{LAND_OUT_ROW}:{col}{MARKETING_OUT_ROW})"
    c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = BLACK_BOLD; c.fill = TOTAL_FILL
ws.cell(row=r, column=1).fill = TOTAL_FILL
TOTAL_OUT_ROW = r
r += 2

label(ws, r, "INFLOWS", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:B{r}")
r += 1
label(ws, r, "Sales revenue recognized")
for i, col in enumerate(qc):
    f = f"={ABS_SHEET}!{col}{REVENUE_ROW}"
    c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = GREEN
REV_IN_ROW = r
r += 2

label(ws, r, "Net Project Cashflow (Unlevered)", bold=True)
for i, col in enumerate(qc):
    f = f"={col}{REV_IN_ROW}-{col}{TOTAL_OUT_ROW}"
    c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = BLACK_BOLD; c.fill = TOTAL_FILL
ws.cell(row=r, column=1).fill = TOTAL_FILL
NET_CF_ROW = r
r += 1
label(ws, r, "Cumulative Cash Balance", bold=True)
for i, col in enumerate(qc):
    f = f"=SUM(${qc[0]}${NET_CF_ROW}:{col}${NET_CF_ROW})"
    c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = BLACK_BOLD
CUM_CF_ROW = r
r += 3

label(ws, r, "SUMMARY (Base Case, Unlevered)", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:C{r}")
r += 1
label(ws, r, "Peak funding requirement (INR Lakh)")
c = ws.cell(row=r, column=3, value=f"=-MIN({qc[0]}{CUM_CF_ROW}:{qc[-1]}{CUM_CF_ROW})"); c.number_format = CUR; c.font = BLACK
PEAK_FUND_CELL = f"C{r}"
r += 1
label(ws, r, "Project IRR (quarterly)")
c = ws.cell(row=r, column=3, value=f"=IRR({qc[0]}{NET_CF_ROW}:{qc[-1]}{NET_CF_ROW})"); c.number_format = PCT1; c.font = BLACK
BASE_IRR_Q_CELL = f"C{r}"
r += 1
label(ws, r, "Project IRR (annualized)", bold=True)
c = ws.cell(row=r, column=3, value=f"=(1+{BASE_IRR_Q_CELL})^4-1"); c.number_format = PCT1; c.font = BLACK_BOLD
BASE_IRR_ANN_CELL = f"C{r}"
r += 1
label(ws, r, "Project NPV @ target discount rate (INR Lakh)", bold=True)
c = ws.cell(row=r, column=3,
            value=f"=NPV(Assumptions!{DISC_RATE_Q_CELL},{qc[0]}{NET_CF_ROW}:{qc[-1]}{NET_CF_ROW})")
c.number_format = CUR; c.font = BLACK_BOLD
BASE_NPV_CELL = f"C{r}"
comment(ws.cell(row=r, column=1), "NPV() treats every quarter as an end-of-period cashflow, consistent with the IRR() convention used throughout this model.")

CF_SHEET = "'Cashflow'"

wb.save("model/property_underwriting_model.xlsx")
print("Cashflow built. NET_CF_ROW", NET_CF_ROW, "BASE_IRR_ANN", BASE_IRR_ANN_CELL, "BASE_NPV", BASE_NPV_CELL)

# ======================================================= EXIT SCENARIOS ===
ws = sheet("Exit Scenarios")
title_block(ws, "Exit Scenarios — Downside / Base / Upside",
            "Each scenario re-derives units sold, revenue and cost per quarter via formulas from Assumptions & Cashflow · IRR/NPV are live Excel functions")
set_col_widths(ws, {"A": 40, "B": 4})
for col in qc:
    ws.column_dimensions[col].width = 11

r = 4
label(ws, r, "SCENARIO LEVERS", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:D{r}")
r += 1
header_row(ws, r, 2, ["Price Multiplier", "Absorption Delay (qtrs)", "Cost Overrun Multiplier"], fill=SUBHEAD_FILL, font=BLACK_BOLD)
r += 1
scenario_lever_row0 = r
scenarios = [
    ("Downside", 0.90, 2, 1.05),
    ("Base Case", 1.00, 0, 1.00),
    ("Upside", 1.10, -1, 0.97),
]
scen_rows = {}
for name, pm, delay, cm in scenarios:
    label(ws, r, name, bold=True)
    c = ws.cell(row=r, column=2, value=pm); c.font = BLUE; c.number_format = MULT
    c = ws.cell(row=r, column=3, value=delay); c.font = BLUE; c.number_format = '0;(0);"-"'
    c = ws.cell(row=r, column=4, value=cm); c.font = BLUE; c.number_format = MULT
    scen_rows[name] = r
    for col in range(1, 5):
        ws.cell(row=r, column=col).border = BORDER
    r += 1
comment(ws.cell(row=scenario_lever_row0, column=3), "Positive = quarters slower than base pace (units pushed later); negative = faster sellout than base.")
r += 1

SALES_START_REF = f"Assumptions!{SALES_START_CELL}"
scen_summary = {}
for name, pm, delay, cm in scenarios:
    lr = scen_rows[name]
    price_cell = f"$B${lr}"
    delay_cell = f"$C${lr}"
    cost_cell = f"$D${lr}"

    label(ws, r, f"{name.upper()} — Quarter", bold=True)
    for i, col in enumerate(qc):
        c = ws[f"{col}{r}"]; c.value = i + 1; c.font = WHITE_BOLD; c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")
    qrow = r
    r += 1

    label(ws, r, "Units sold this quarter")
    for i, col in enumerate(qc):
        rel = f"({col}${qrow}-{SALES_START_REF}-{delay_cell}+1)"
        f = (f"=IF(AND({rel}>=1,{rel}<=10),"
             f"INDEX(Assumptions!${ABS_CURVE_FIRST_COL}${ABS_CURVE_ROW}:${ABS_CURVE_LAST_COL}${ABS_CURVE_ROW},{rel}),0)")
        c = ws[f"{col}{r}"]; c.value = f; c.number_format = NUM0
    units_row = r
    r += 1

    label(ws, r, "Revenue (INR Lakh)")
    for i, col in enumerate(qc):
        f = f"={col}{units_row}*Assumptions!{AVG_REV_PER_UNIT_CELL}*{price_cell}"
        c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = GREEN
    rev_row = r
    r += 1

    label(ws, r, "Cost outflow (land + construction + IDC, scaled)")
    for i, col in enumerate(qc):
        f = f"=({CF_SHEET}!{col}{LAND_OUT_ROW}+({CF_SHEET}!{col}{CONSTR_OUT_ROW}+{CF_SHEET}!{col}{IDC_OUT_ROW})*{cost_cell})"
        c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = GREEN
    cost_row = r
    r += 1

    label(ws, r, "Marketing cost (% of scenario revenue)")
    for i, col in enumerate(qc):
        f = f"={col}{rev_row}*Assumptions!{MARKETING_PCT_CELL}"
        c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = GREEN
    mkt_row = r
    r += 1

    label(ws, r, "Net Cashflow", bold=True)
    for i, col in enumerate(qc):
        f = f"={col}{rev_row}-{col}{cost_row}-{col}{mkt_row}"
        c = ws[f"{col}{r}"]; c.value = f; c.number_format = CUR; c.font = BLACK_BOLD; c.fill = TOTAL_FILL
    ws.cell(row=r, column=1).fill = TOTAL_FILL
    netcf_row = r
    r += 2

    label(ws, r, f"{name} — Total Revenue (Lakh)")
    c = ws.cell(row=r, column=3, value=f"=SUM({qc[0]}{rev_row}:{qc[-1]}{rev_row})"); c.number_format = CUR
    total_rev_cell = f"C{r}"
    r += 1
    label(ws, r, f"{name} — Total Cost (Lakh)")
    c = ws.cell(row=r, column=3, value=f"=SUM({qc[0]}{cost_row}:{qc[-1]}{cost_row})+SUM({qc[0]}{mkt_row}:{qc[-1]}{mkt_row})"); c.number_format = CUR
    total_cost_cell = f"C{r}"
    r += 1
    label(ws, r, f"{name} — IRR (quarterly)")
    c = ws.cell(row=r, column=3, value=f"=IRR({qc[0]}{netcf_row}:{qc[-1]}{netcf_row})"); c.number_format = PCT1
    irr_q_cell = f"C{r}"
    r += 1
    label(ws, r, f"{name} — IRR (annualized)", bold=True)
    c = ws.cell(row=r, column=3, value=f"=(1+{irr_q_cell})^4-1"); c.number_format = PCT1; c.font = BLACK_BOLD
    irr_ann_cell = f"C{r}"
    r += 1
    label(ws, r, f"{name} — NPV @ target rate (Lakh)", bold=True)
    c = ws.cell(row=r, column=3, value=f"=NPV(Assumptions!{DISC_RATE_Q_CELL},{qc[0]}{netcf_row}:{qc[-1]}{netcf_row})")
    c.number_format = CUR; c.font = BLACK_BOLD
    npv_cell = f"C{r}"
    r += 3

    scen_summary[name] = dict(total_rev=total_rev_cell, total_cost=total_cost_cell,
                              irr_ann=irr_ann_cell, npv=npv_cell, netcf_row=netcf_row, qrow=qrow)

EXIT_SHEET = "'Exit Scenarios'"

wb.save("model/property_underwriting_model.xlsx")
print("Exit Scenarios built.", scen_summary)

# ======================================================= SENSITIVITY ======
ws = sheet("Sensitivity")
title_block(ws, "Sensitivity — IRR & NPV vs. Price and Absorption Pace",
            "2-way grid, cost overrun held at base (1.00x) · each cell is a live IRR()/NPV() formula over its own scenario cashflow (helper rows below, columns AB onward)")
set_col_widths(ws, {"A": 26})
for col in ["B", "C", "D", "E", "F"]:
    ws.column_dimensions[col].width = 12

price_axis = [0.90, 0.95, 1.00, 1.05, 1.10]
delay_axis = [-1, 0, 1, 2, 3]

HELPER_COL0 = 28  # column AB, keeps helper block well clear of the visible grid
helper_qc = [get_column_letter(HELPER_COL0 + i) for i in range(N_Q)]

# --- visible IRR grid ---
r = 4
label(ws, r, "IRR (annualized) — rows: price multiplier, cols: absorption delay (qtrs)", bold=True)
ws.merge_cells(f"A{r}:F{r}")
r += 1
grid_header_row = r
ws.cell(row=r, column=1, value="Price \\ Delay").font = BLACK_BOLD
for j, d in enumerate(delay_axis):
    c = ws.cell(row=r, column=2 + j, value=d); c.font = WHITE_BOLD; c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="center")
irr_grid_top = r
r += 1
irr_grid_first_data_row = r
irr_cells = {}
npv_cells = {}
for i, p in enumerate(price_axis):
    row = irr_grid_first_data_row + i
    c = ws.cell(row=row, column=1, value=p); c.font = BLACK_BOLD; c.number_format = MULT
    for j, d in enumerate(delay_axis):
        irr_cells[(i, j)] = (row, 2 + j)
r += len(price_axis) + 2

label(ws, r, "NPV @ target discount rate (INR Lakh) — rows: price multiplier, cols: absorption delay (qtrs)", bold=True)
ws.merge_cells(f"A{r}:F{r}")
r += 1
npv_header_row = r
ws.cell(row=r, column=1, value="Price \\ Delay").font = BLACK_BOLD
for j, d in enumerate(delay_axis):
    c = ws.cell(row=r, column=2 + j, value=d); c.font = WHITE_BOLD; c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="center")
r += 1
npv_grid_first_data_row = r
for i, p in enumerate(price_axis):
    row = npv_grid_first_data_row + i
    c = ws.cell(row=row, column=1, value=p); c.font = BLACK_BOLD; c.number_format = MULT
    for j, d in enumerate(delay_axis):
        npv_cells[(i, j)] = (row, 2 + j)
r += len(price_axis) + 3

# --- helper block (one mini scenario cashflow per price/delay combo) ---
label(ws, r, "HELPER CALCULATIONS (feed the grids above — not intended for manual editing)", fill=SUBHEAD_FILL, bold=True)
ws.merge_cells(f"A{r}:B{r}")
helper_start_row = r + 1
r = helper_start_row

for i, p in enumerate(price_axis):
    for j, d in enumerate(delay_axis):
        combo_label = f"P{p:.2f}_D{d}"
        label(ws, r, f"{combo_label} — Quarter")
        for k, col in enumerate(helper_qc):
            ws[f"{col}{r}"] = k + 1
        qrow = r
        r += 1
        label(ws, r, f"{combo_label} — Units sold")
        for k, col in enumerate(helper_qc):
            rel = f"({col}${qrow}-Assumptions!{SALES_START_CELL}-{d}+1)"
            f = (f"=IF(AND({rel}>=1,{rel}<=10),"
                 f"INDEX(Assumptions!${ABS_CURVE_FIRST_COL}${ABS_CURVE_ROW}:${ABS_CURVE_LAST_COL}${ABS_CURVE_ROW},{rel}),0)")
            ws[f"{col}{r}"] = f
        units_row = r
        r += 1
        label(ws, r, f"{combo_label} — Revenue")
        for k, col in enumerate(helper_qc):
            ws[f"{col}{r}"] = f"={col}{units_row}*Assumptions!{AVG_REV_PER_UNIT_CELL}*{p}"
        rev_row = r
        r += 1
        label(ws, r, f"{combo_label} — Cost (incl. marketing)")
        for k, col in enumerate(helper_qc):
            base_col = qc[k]
            f = (f"={CF_SHEET}!{base_col}{LAND_OUT_ROW}+{CF_SHEET}!{base_col}{CONSTR_OUT_ROW}"
                 f"+{CF_SHEET}!{base_col}{IDC_OUT_ROW}+{col}{rev_row}*Assumptions!{MARKETING_PCT_CELL}")
            ws[f"{col}{r}"] = f
        cost_row = r
        r += 1
        label(ws, r, f"{combo_label} — Net Cashflow")
        for k, col in enumerate(helper_qc):
            ws[f"{col}{r}"] = f"={col}{rev_row}-{col}{cost_row}"
        netcf_row = r
        r += 1
        irr_formula = f"=(1+IRR({helper_qc[0]}{netcf_row}:{helper_qc[-1]}{netcf_row}))^4-1"
        npv_formula = f"=NPV(Assumptions!{DISC_RATE_Q_CELL},{helper_qc[0]}{netcf_row}:{helper_qc[-1]}{netcf_row})"
        irr_r, irr_c = irr_cells[(i, j)]
        npv_r, npv_c = npv_cells[(i, j)]
        cc = ws.cell(row=irr_r, column=irr_c, value=irr_formula); cc.number_format = PCT1
        cc2 = ws.cell(row=npv_r, column=npv_c, value=npv_formula); cc2.number_format = CUR
        r += 1

helper_end_row = r
# hide helper rows so the sheet reads cleanly, values remain fully live
for hr in range(helper_start_row, helper_end_row):
    ws.row_dimensions[hr].hidden = True

# color-scale heatmap on the IRR grid
irr_range = f"B{irr_grid_first_data_row}:F{irr_grid_first_data_row + len(price_axis) - 1}"
ws.conditional_formatting.add(
    irr_range,
    ColorScaleRule(start_type="min", start_color="F8696B",
                   mid_type="percentile", mid_value=50, mid_color="FFEB84",
                   end_type="max", end_color="63BE7B"),
)
for i in range(len(price_axis)):
    row = irr_grid_first_data_row + i
    for j in range(len(delay_axis)):
        ws.cell(row=row, column=2 + j).border = BORDER

npv_range = f"B{npv_grid_first_data_row}:F{npv_grid_first_data_row + len(price_axis) - 1}"
ws.conditional_formatting.add(
    npv_range,
    ColorScaleRule(start_type="min", start_color="F8696B",
                   mid_type="percentile", mid_value=50, mid_color="FFEB84",
                   end_type="max", end_color="63BE7B"),
)

SENS_SHEET = "'Sensitivity'"
SENS_IRR_RANGE = irr_range
SENS_IRR_HEADER_ROW = irr_grid_top
SENS_IRR_FIRST_ROW = irr_grid_first_data_row
SENS_PRICE_AXIS = price_axis
SENS_DELAY_AXIS = delay_axis

wb.save("model/property_underwriting_model.xlsx")
print("Sensitivity built. helper rows", helper_start_row, "-", helper_end_row, "irr_range", irr_range)

# =========================================================== DASHBOARD ====
ws = sheet("Summary Dashboard")
wb.move_sheet("Summary Dashboard", offset=-(len(wb.sheetnames) - 1))  # make it the first tab
title_block(ws, "Project Meridian Ridge — Summary Dashboard",
            "Fictional 40-unit weekend-home development · all figures illustrative/synthetic · see README", span_col="J")
set_col_widths(ws, {"A": 34, "B": 16, "C": 4, "D": 34, "E": 16})

r = 4
label(ws, r, "HEADLINE METRICS", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:B{r}")
r += 1
metric_rows = []

def metric(row, name, formula, fmt, boldval=True):
    label(ws, row, name)
    c = ws.cell(row=row, column=2, value=formula)
    c.number_format = fmt
    c.font = BLACK_BOLD if boldval else GREEN
    return c

metric(r, "Total Units", "=Assumptions!" + TOTAL_UNITS_CELL, NUM0); r += 1
metric(r, "Total Saleable Area (sqft)", "=Assumptions!" + TOTAL_AREA_CELL, NUM0); r += 1
metric(r, "Total Revenue (INR Lakh)", f"={COST_SHEET}!{TOTAL_DEV_COST_CELL[0]}23" if False else f"=Assumptions!{TOTAL_REVENUE_CELL}", CUR); r += 1
metric(r, "Total Development Cost (INR Lakh)", f"={COST_SHEET}!{TOTAL_DEV_COST_CELL}", CUR); r += 1
metric(r, "Gross Margin (INR Lakh)", f"={COST_SHEET}!{GROSS_MARGIN_CELL}", CUR); r += 1
metric(r, "Gross Margin %", f"={COST_SHEET}!{GROSS_MARGIN_PCT_CELL}", PCT1); r += 1
metric(r, "Peak Funding Requirement (INR Lakh)", f"={CF_SHEET}!{PEAK_FUND_CELL}", CUR); r += 1
r += 1

label(ws, r, "BASE CASE RETURNS", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:B{r}")
r += 1
metric(r, "Project IRR (annualized)", f"={CF_SHEET}!{BASE_IRR_ANN_CELL}", PCT1); r += 1
metric(r, "Project NPV @ target rate (INR Lakh)", f"={CF_SHEET}!{BASE_NPV_CELL}", CUR); r += 1
r += 1

label(ws, r, "SCENARIO COMPARISON", fill=SUBHEAD_FILL, bold=True); ws.merge_cells(f"A{r}:B{r}")
r += 1
header_row(ws, r, 1, ["Scenario", "Total Revenue", "Total Cost", "IRR (ann.)", "NPV"], fill=SUBHEAD_FILL, font=BLACK_BOLD)
r += 1
scen_table_start = r
for name in ["Downside", "Base Case", "Upside"]:
    s = scen_summary[name]
    label(ws, r, name, bold=True)
    c = ws.cell(row=r, column=2, value=f"={EXIT_SHEET}!{s['total_rev']}"); c.number_format = CUR
    c = ws.cell(row=r, column=3, value=f"={EXIT_SHEET}!{s['total_cost']}"); c.number_format = CUR
    c = ws.cell(row=r, column=4, value=f"={EXIT_SHEET}!{s['irr_ann']}"); c.number_format = PCT1
    c = ws.cell(row=r, column=5, value=f"={EXIT_SHEET}!{s['npv']}"); c.number_format = CUR
    for col in range(1, 6):
        ws.cell(row=r, column=col).border = BORDER
    r += 1
scen_table_end = r - 1
r += 2

dash_note_row = r
ws.cell(row=r, column=1, value=("All figures are illustrative/synthetic — methodology demonstration only, "
                                 "not a real transaction. See README for details."))
ws.cell(row=r, column=1).font = Font(name=FONT_NAME, italic=True, size=9, color="808080")
ws.merge_cells(f"A{r}:J{r}")

# --- Chart 1: Cashflow over time (base case) ---
lc = LineChart()
lc.title = "Base Case Cashflow Over Time (INR Lakh)"
lc.style = 2
lc.y_axis.title = "INR Lakh"
lc.x_axis.title = "Quarter"
lc.width = 18
lc.height = 9

cf_ws = wb["Cashflow"]
cats = Reference(cf_ws, min_col=Q_START_COL, max_col=Q_START_COL + N_Q - 1, min_row=CF_QROW, max_row=CF_QROW)
net_ref = Reference(cf_ws, min_col=2, max_col=Q_START_COL + N_Q - 1, min_row=NET_CF_ROW, max_row=NET_CF_ROW)
cum_ref = Reference(cf_ws, min_col=2, max_col=Q_START_COL + N_Q - 1, min_row=CUM_CF_ROW, max_row=CUM_CF_ROW)
s1 = Series(Reference(cf_ws, min_col=Q_START_COL, max_col=Q_START_COL + N_Q - 1, min_row=NET_CF_ROW, max_row=NET_CF_ROW), title="Net Cashflow")
s2 = Series(Reference(cf_ws, min_col=Q_START_COL, max_col=Q_START_COL + N_Q - 1, min_row=CUM_CF_ROW, max_row=CUM_CF_ROW), title="Cumulative Cash Balance")
lc.series.append(s1)
lc.series.append(s2)
lc.set_categories(cats)
ws.add_chart(lc, "D4")

# --- Chart 2: IRR sensitivity by scenario (price axis lines across absorption delay) ---
bc = BarChart()
bc.type = "col"
bc.grouping = "clustered"
bc.title = "IRR Sensitivity — by Price Multiplier & Absorption Delay"
bc.style = 10
bc.y_axis.title = "IRR (annualized)"
bc.x_axis.title = "Absorption Delay (quarters)"
bc.width = 18
bc.height = 9

sens_ws = wb["Sensitivity"]
delay_cats = Reference(sens_ws, min_col=2, max_col=2 + len(delay_axis) - 1, min_row=SENS_IRR_HEADER_ROW, max_row=SENS_IRR_HEADER_ROW)
for i, p in enumerate(price_axis):
    row = SENS_IRR_FIRST_ROW + i
    ref = Reference(sens_ws, min_col=2, max_col=2 + len(delay_axis) - 1, min_row=row, max_row=row)
    ser = Series(ref, title=f"Price {p:.2f}x")
    bc.series.append(ser)
bc.set_categories(delay_cats)
ws.add_chart(bc, "D22")

DASH_SHEET = "'Summary Dashboard'"

wb.save("model/property_underwriting_model.xlsx")
print("Dashboard built with charts. Sheet order:", wb.sheetnames)
