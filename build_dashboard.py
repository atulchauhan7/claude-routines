"""
build_dashboard.py
Creates: Daily_Store_Ads_Performance_Sheet_2026-05-29.xlsx
Report Date: 2026-05-29
"""

import os
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.filters import AutoFilter

# ── colour palette ──────────────────────────────────────────────────────────
C_DARK_BG   = "1F3864"   # dark navy – title bar
C_DARK_FG   = "FFFFFF"   # white text
C_SECTION   = "2E75B6"   # blue section header bg
C_SECT_FG   = "FFFFFF"
C_GREEN_H   = "1E7145"   # wins header
C_RED_H     = "C00000"   # issues header
C_BLUE_H    = "1F497D"   # actions header
C_GRAY_ROW  = "F2F2F2"   # alternating row
C_WHITE     = "FFFFFF"
C_HEADER_BG = "4472C4"   # table column header
C_HEADER_FG = "FFFFFF"
C_GREEN_CF  = "C6EFCE"   # conditional format green fill
C_RED_CF    = "FFC7CE"   # conditional format red fill
C_ORANGE_CF = "FFEB9C"   # conditional format orange/yellow fill
C_GREEN_FONT= "375623"
C_RED_FONT  = "9C0006"
C_ORANGE_FN = "9C5700"

# ── helpers ─────────────────────────────────────────────────────────────────

def make_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def make_font(bold=False, color="000000", size=11, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic)

def make_border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def make_alignment(horizontal="left", vertical="center", wrap=False):
    return Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap)

def set_col_width(ws, col_letter, width):
    ws.column_dimensions[col_letter].width = width

def apply_table_header(ws, row, cols, texts, bg=C_HEADER_BG, fg=C_HEADER_FG):
    for i, text in enumerate(texts):
        cl = get_column_letter(cols + i)
        cell = ws[f"{cl}{row}"]
        cell.value = text
        cell.font = make_font(bold=True, color=fg)
        cell.fill = make_fill(bg)
        cell.alignment = make_alignment("center")
        cell.border = make_border()

def write_row(ws, row, start_col, values, fill_hex=None, bold=False, fmt_list=None):
    for i, v in enumerate(values):
        cl = get_column_letter(start_col + i)
        cell = ws[f"{cl}{row}"]
        cell.value = v
        cell.border = make_border()
        cell.alignment = make_alignment("center")
        if fill_hex:
            cell.fill = make_fill(fill_hex)
        if bold:
            cell.font = make_font(bold=True)
        if fmt_list and fmt_list[i]:
            cell.number_format = fmt_list[i]

def section_header(ws, row, col, text, bg, fg=C_SECT_FG, colspan=14):
    ws.merge_cells(start_row=row, start_column=col,
                   end_row=row, end_column=col + colspan - 1)
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = make_font(bold=True, color=fg, size=12)
    cell.fill = make_fill(bg)
    cell.alignment = make_alignment("left", "center")


# ═══════════════════════════════════════════════════════════════════════════
# DATA
# ═══════════════════════════════════════════════════════════════════════════

shopify_data = [
    # date, gross, net, orders, aov, returns, taxes, total_sales, sessions, cart, checkout, completed, conv
    ("2026-05-22", 152830,    119671.13, 62, 2433.16, 31185, 0,      119671.13, 4667, 217, 0, 0, 0.0),
    ("2026-05-23", 137827.14, 111809.65, 58, 2324.15, 22991, 609.86, 112419.51, 3927, 192, 1, 1, 0.025),
    ("2026-05-24", 120243,    101392.98, 51, 2325.20, 17192, 0,      101392.98, 3949, 213, 0, 0, 0.0),
    ("2026-05-25", 197213,    169045.25, 69, 2849.76, 27588, 0,      169045.25, 4073, 252, 2, 0, 0.0),
    ("2026-05-26", 138643,    122890.18, 50, 2739.66, 14093, 0,      122890.18, 4208, 208, 0, 0, 0.0),
    ("2026-05-27", 162158.81, 143153.25, 64, 2450.74, 13694, 320.19, 143473.44, 4143, 239, 2, 1, 0.024),
    ("2026-05-28", 160630,    140368.98, 59, 2682.39, 17892, 0,      140368.98, 4092, 214, 1, 0, 0.0),
    ("2026-05-29", 166625,    138446.33, 67, 2452.75, 25888, 0,      138446.33, 4297, 232, 0, 0, 0.0),
]

# 7-day avg (May 22-28)
shopify_7d_avg = (
    "7-Day Avg (May 22-28)",
    152792.14, 129761.63, 59, 2543.58, 20662.14, None, None,
    4151,      218,       1,  0,       0.007
)

products_yesterday = [
    # title, gross, net, orders, aov, returns
    ('"The Essential" - White Cotton Shirt & Pant Co-ord Set',         48876,    34483,    21, 2327.43, 14393),
    ("Embroidered Rayon Kurti with Farshi Salwar Set",                 22592,    22312.10,  8, 2789.01,     0),
    ('"Riva" - Premium Flowy Rayon Co-ord Set',                        11994,     9895.05,  6, 1982.34,  1999),
    ("Everyday Elegance Cotton Kurta Pant Set",                         9995,     5997,     3, 3331.67,  3998),
    ("Black Cotton Kurta Pant Set",                                     5997,     5997,     3, 1999,        0),
    ('"Inayat" - Power Pastel Cotton Shirt & Trouser Co-ord Set',       5997,     5997,     3, 1999,        0),
    ("Ira Rayon Co-ord Set for Women",                                  4998,     2499,     2, 2499,     2499),
    ('"Ayeza" - Dusty Rose Premium Rayon Kurta Set',                    4598,     3871.51,  2, 1935.76,    0),
    ("SAVYA Airy Linen Kurta Co-ord Set",                               3998,     3898.05,  2, 1949.03,    0),
    ('"Meher" - Pristine White Georgette Anarkali Suit Set',            3499,     3499,     1, 3499,        0),
    ('"Adira" - Premium Breathable Soft Fabric Salwar Set',             2999,        0,     1, 2999,     2999),
    ("Hania Aamir Inspired Viral Satin Silk Kurta Palazzo Set",         2999,     2999,     1, 2999,        0),
    ('"Nidra" - Jet Black Premium Georgette Suit Set',                  2999,     2999,     1, 2999,        0),
    ("Embroidered Rayon Kurti with Farshi Salwar Set (Elegant)",        2799,     1914.52,  1, 1914.52,    0),
    ("Neer Cotton Co-ord Set for Women",                                2499,     2499,     1, 2499,        0),
]

# campaigns yesterday + 7d
campaigns_yday = [
    # name, status, spend, impressions, reach, clicks, ctr, cpc, purchases, cpp, roas, daily_budget
    ("Summer Special Campaign",  "ACTIVE",  1583.45,   6069,   4957,    110, 1.81, 14.40,  5,  316.69, 6.31, "ABO"),
    ("Retargeting-new",          "ACTIVE",  7158.56,  24540,  16268,    808, 3.29,  8.86, 10,  715.86, 3.37, "ABO"),
    ("Roas Goal",                "ACTIVE",  1459.07,   6237,   5984,     86, 1.38, 16.97,  0,    None, None, 3600),
    ("Dhirai Scale -ASC",        "ACTIVE", 20436.23, 182543, 164482,   3154, 1.73,  6.48, 27,  756.90, 3.22, 22000),
    ("Black DHURANDHAR",         "ACTIVE",  8729.34,  39770,  33793,   1507, 3.79,  5.79,  9,  969.93, 2.19, 10000),
]

campaigns_7d = [
    # name, spend, impressions, reach, clicks, ctr, cpc, purchases, cpp, roas
    ("Summer Special Campaign",  13503.68,  47384,  31696,   916, 1.93, 14.74,  16,  843.98, 2.40),
    ("Retargeting-new",          58415.81, 194452,  64923,  6697, 3.44,  8.72, 106,  551.09, 4.64),
    ("Roas Goal",                11548.42,  43743,  36584,   808, 1.85, 14.29,  11, 1049.86, 2.52),
    ("Dhirai Scale -ASC",       153962.12,1184730, 825244, 22469, 1.90,  6.85, 177,  869.84, 2.86),
    ("Black DHURANDHAR",         69822.49, 331716, 219477, 10788, 3.25,  6.47,  75,  930.97, 2.66),
]

# ═══════════════════════════════════════════════════════════════════════════
# WORKBOOK
# ═══════════════════════════════════════════════════════════════════════════
wb = Workbook()

# Remove default sheet
wb.remove(wb.active)

# Create all 6 sheets
ws_dash   = wb.create_sheet("Dashboard")
ws_shopi  = wb.create_sheet("Shopify Daily Data")
ws_meta   = wb.create_sheet("Meta Ads Daily Data")
ws_prod   = wb.create_sheet("Product Performance")
ws_camp   = wb.create_sheet("Campaign Performance")
ws_rec    = wb.create_sheet("Recommendations & Notes")


# ═══════════════════════════════════════════════════════════════════════════
# SHEET 1: DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dash

# Row 1 – Title
ws.merge_cells("A1:N1")
c = ws["A1"]
c.value = "DHIRAI — Daily Store & Ads Performance Dashboard"
c.font = make_font(bold=True, color=C_DARK_FG, size=18)
c.fill = make_fill(C_DARK_BG)
c.alignment = make_alignment("center", "center")
ws.row_dimensions[1].height = 36

# Row 2 – Subtitle
ws.merge_cells("A2:N2")
c = ws["A2"]
c.value = "Report Date: 29 May 2026  |  Generated: 30 May 2026  |  Currency: INR"
c.font = make_font(italic=True, color="666666", size=11)
c.alignment = make_alignment("center", "center")
ws.row_dimensions[2].height = 20

# Row 3 blank
ws.row_dimensions[3].height = 10

# ── Section A: Executive Summary ────────────────────────────────────────
section_header(ws, 4, 1, "  SECTION A — EXECUTIVE SUMMARY", C_DARK_BG)
ws.row_dimensions[4].height = 22

exec_headers = ["Gross Sales (₹)", "Orders", "Net Sales (₹)", "Avg Order Value (₹)",
                "Total Ad Spend (₹)", "Blended ROAS", "Meta Purchases"]
exec_values  = [166625, 67, 138446.33, 2452.75, 39366.65, 3.51, 51]
exec_fmts    = ['₹#,##0.00', '0', '₹#,##0.00', '₹#,##0.00', '₹#,##0.00', '0.00', '0']

apply_table_header(ws, 5, 1, exec_headers)
for i, (h, v, f) in enumerate(zip(exec_headers, exec_values, exec_fmts)):
    cl = get_column_letter(1 + i)
    cell = ws[f"{cl}6"]
    cell.value = v
    cell.number_format = f
    cell.font = make_font(bold=True, size=12)
    cell.fill = make_fill(C_GRAY_ROW)
    cell.alignment = make_alignment("center")
    cell.border = make_border()

ws.row_dimensions[7].height = 10  # spacer

# ── Section B: Shopify Performance ──────────────────────────────────────
section_header(ws, 8, 1, "  SECTION B — SHOPIFY PERFORMANCE — Yesterday (29 May) vs 7-Day Average", C_SECTION)
ws.row_dimensions[8].height = 22

apply_table_header(ws, 9, 1, ["Metric", "Yesterday", "7-Day Avg", "Change %", "Status"])
shopi_rows = [
    ("Gross Sales",      "₹1,66,625",  "₹1,52,792", "+9.1%",  "ABOVE"),
    ("Net Sales",        "₹1,38,446",  "₹1,29,762", "+6.7%",  "ABOVE"),
    ("Orders",           "67",         "59",         "+13.6%", "ABOVE"),
    ("Avg Order Value",  "₹2,453",     "₹2,544",     "-3.6%",  "BELOW"),
    ("Returns",          "₹25,888",    "₹20,662",    "+25.3%", "⚠ HIGH"),
    ("Sessions",         "4,297",      "4,151",      "+3.5%",  "ABOVE"),
]
status_colors = {
    "ABOVE":  (C_GREEN_CF, C_GREEN_FONT),
    "BELOW":  (C_ORANGE_CF, C_ORANGE_FN),
    "⚠ HIGH": (C_RED_CF, C_RED_FONT),
}
for ridx, row_data in enumerate(shopi_rows):
    r = 10 + ridx
    fill_hex = C_GRAY_ROW if ridx % 2 == 0 else C_WHITE
    for cidx, val in enumerate(row_data):
        cl = get_column_letter(1 + cidx)
        cell = ws[f"{cl}{r}"]
        cell.value = val
        cell.border = make_border()
        cell.alignment = make_alignment("center")
        cell.fill = make_fill(fill_hex)
        if cidx == 4:  # Status column
            bg, fg = status_colors.get(val, (fill_hex, "000000"))
            cell.fill = make_fill(bg)
            cell.font = make_font(bold=True, color=fg)

ws.row_dimensions[16].height = 10  # spacer

# ── Section C: Meta Ads Performance ─────────────────────────────────────
section_header(ws, 17, 1, "  SECTION C — META ADS PERFORMANCE — Yesterday vs 7-Day Daily Average", C_SECTION)
ws.row_dimensions[17].height = 22

apply_table_header(ws, 18, 1, ["Metric", "Yesterday", "7-Day Daily Avg", "Change %"])
meta_rows = [
    ("Total Spend",        "₹39,367",  "₹43,893", "-10.3%"),
    ("Total Purchases",    "51",        "55",       "-7.3%"),
    ("Total Impressions",  "2,59,159",  "3,00,432", "-13.7%"),
    ("Total Clicks",       "5,665",     "7,239",    "-21.7%"),
    ("Blended CTR",        "2.18%",     "2.41%",    "-9.5%"),
    ("Avg CPC",            "₹6.95",     "₹6.06",    "+14.7%"),
    ("Blended ROAS",       "3.51",      "3.50",     "+0.3%"),
]
for ridx, row_data in enumerate(meta_rows):
    r = 19 + ridx
    fill_hex = C_GRAY_ROW if ridx % 2 == 0 else C_WHITE
    for cidx, val in enumerate(row_data):
        cl = get_column_letter(1 + cidx)
        cell = ws[f"{cl}{r}"]
        cell.value = val
        cell.border = make_border()
        cell.alignment = make_alignment("center")
        cell.fill = make_fill(fill_hex)
        if cidx == 3:  # Change column
            if val.startswith("+"):
                cell.fill = make_fill(C_GREEN_CF)
                cell.font = make_font(bold=True, color=C_GREEN_FONT)
            elif val.startswith("-"):
                cell.fill = make_fill(C_RED_CF)
                cell.font = make_font(bold=True, color=C_RED_FONT)

ws.row_dimensions[26].height = 10  # spacer

# ── Section D: Key Wins ──────────────────────────────────────────────────
section_header(ws, 27, 1, "  SECTION D — KEY WINS", C_GREEN_H)
ws.row_dimensions[27].height = 22

wins = [
    "1.  Summer Special Campaign achieved ROAS 6.31 yesterday (vs 2.40 avg) — standout performer",
    "2.  Orders at 67 are +13.6% above 7-day average",
    "3.  Overall gross sales ₹1.67L above 7-day trend by +9%",
    '4.  "The Essential" White Co-ord generated ₹48,876 in a single day (29% of daily revenue)',
]
for ridx, text in enumerate(wins):
    r = 28 + ridx
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
    cell = ws.cell(row=r, column=1, value=text)
    cell.fill = make_fill("E2EFDA")
    cell.font = make_font(color="375623")
    cell.alignment = make_alignment("left", "center", wrap=True)
    ws.row_dimensions[r].height = 18

ws.row_dimensions[32].height = 10  # spacer

# ── Section E: Key Issues ────────────────────────────────────────────────
section_header(ws, 33, 1, "  SECTION E — KEY ISSUES", C_RED_H)
ws.row_dimensions[33].height = 22

issues = [
    "1.  Returns spiked to ₹25,888 (+25.3% above avg) — 5 products with returns yesterday",
    "2.  Zero checkout completions from 4,297 sessions — possible tracking/payment issue",
    "3.  Black DHURANDHAR ROAS only 2.19 (lowest active) — ₹8,729 spend, 9 purchases",
    "4.  Roas Goal campaign: ₹1,459 spent, ZERO purchases — needs review/pause",
    '5.  "Adira" Salwar Set: ₹2,999 gross but ₹0 net (full return)',
]
for ridx, text in enumerate(issues):
    r = 34 + ridx
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
    cell = ws.cell(row=r, column=1, value=text)
    cell.fill = make_fill("FCE4D6")
    cell.font = make_font(color=C_RED_FONT)
    cell.alignment = make_alignment("left", "center", wrap=True)
    ws.row_dimensions[r].height = 18

ws.row_dimensions[39].height = 10  # spacer

# ── Section F: Recommended Actions ──────────────────────────────────────
section_header(ws, 40, 1, "  SECTION F — RECOMMENDED ACTIONS", C_BLUE_H)
ws.row_dimensions[40].height = 22

actions = [
    "1.  Scale Summer Special Campaign — ROAS 6.31 is exceptional; increase budget 20-30%",
    '2.  Review/Pause "Roas Goal" campaign — zero purchases from ₹1,459 spend yesterday',
    "3.  Investigate checkout tracking — 4,297 sessions, 232 cart additions, 0 completions is abnormal",
    '4.  Audit return reasons for "The Essential" (₹14,393 returns) and "Adira" (100% return) — potential sizing/quality issues',
    "5.  Increase Black DHURANDHAR budget only if CPP improves; current ₹970 CPP on avg ₹2,500 AOV gives thin margins",
]
for ridx, text in enumerate(actions):
    r = 41 + ridx
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
    cell = ws.cell(row=r, column=1, value=text)
    cell.fill = make_fill("DDEEFF")
    cell.font = make_font(color=C_BLUE_H)
    cell.alignment = make_alignment("left", "center", wrap=True)
    ws.row_dimensions[r].height = 18

# Column widths for dashboard
ws.column_dimensions["A"].width = 35
for col_letter in ["B","C","D","E","F","G","H","I","J","K","L","M","N"]:
    ws.column_dimensions[col_letter].width = 16

# Freeze row 1
ws.freeze_panes = "A2"


# ═══════════════════════════════════════════════════════════════════════════
# SHEET 2: Shopify Daily Data
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_shopi

# Row 1 – Title
ws.merge_cells("A1:M1")
c = ws["A1"]
c.value = "DHIRAI — Shopify Daily Data  |  2026-05-22 to 2026-05-29"
c.font = make_font(bold=True, color=C_DARK_FG, size=14)
c.fill = make_fill(C_DARK_BG)
c.alignment = make_alignment("center", "center")
ws.row_dimensions[1].height = 28

headers = [
    "Date", "Gross Sales", "Net Sales", "Orders", "Avg Order Value",
    "Returns", "Taxes", "Total Sales", "Sessions", "Cart Additions",
    "Checkout Reached", "Checkout Completed", "Conversion Rate"
]
money_cols = {2, 3, 5, 6, 8}   # 1-indexed within data cols
pct_cols   = {13}

# Row 2 – 7-day avg summary row
avg_row_vals = [
    "7-Day Avg (May 22-28)",
    152792.14, 129761.63, 59, 2543.58, 20662.14, None, None,
    4151, 218, 1, 0, 0.00007
]
for cidx, val in enumerate(avg_row_vals):
    cl = get_column_letter(1 + cidx)
    cell = ws[f"{cl}2"]
    cell.value = val
    cell.fill = make_fill("BDD7EE")
    cell.font = make_font(bold=True)
    cell.alignment = make_alignment("center")
    cell.border = make_border()
    if val is not None and cidx + 1 in money_cols:
        cell.number_format = '₹#,##0.00'
    elif cidx + 1 in pct_cols and val is not None:
        cell.number_format = '0.000%'

# Row 3 – Headers
apply_table_header(ws, 3, 1, headers)

# Row 4 onwards – data
for ridx, row in enumerate(shopify_data):
    r = 4 + ridx
    fill_hex = C_GRAY_ROW if ridx % 2 == 0 else C_WHITE
    is_yesterday = (row[0] == "2026-05-29")
    for cidx, val in enumerate(row):
        cl = get_column_letter(1 + cidx)
        cell = ws[f"{cl}{r}"]
        cell.value = val
        cell.border = make_border()
        cell.alignment = make_alignment("center")
        cell.fill = make_fill("FFF2CC" if is_yesterday else fill_hex)
        if is_yesterday:
            cell.font = make_font(bold=True)
        col1 = cidx + 1
        if col1 in money_cols and val is not None:
            cell.number_format = '₹#,##0.00'
        elif col1 in pct_cols:
            cell.number_format = '0.000%'

# Conditional formatting on Gross Sales (col B = column 2)
from openpyxl.formatting.rule import ColorScaleRule
ws.conditional_formatting.add(
    "B4:B11",
    ColorScaleRule(
        start_type='min', start_color='FFCCCC',
        mid_type='percentile', mid_value=50, mid_color='FFFF99',
        end_type='max', end_color='C6EFCE'
    )
)

# Autofilter
ws.auto_filter.ref = "A3:M3"

# Freeze row 3
ws.freeze_panes = "A4"

# Column widths
ws.column_dimensions["A"].width = 14
for c in ["B","C","E","F","G","H"]:
    ws.column_dimensions[c].width = 16
for c in ["D","I","J","K","L","M"]:
    ws.column_dimensions[c].width = 14


# ═══════════════════════════════════════════════════════════════════════════
# SHEET 3: Meta Ads Daily Data
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_meta

# Title
ws.merge_cells("A1:L1")
c = ws["A1"]
c.value = "DHIRAI — Meta Ads Daily Data  |  Yesterday (29 May 2026) + Last 7 Days"
c.font = make_font(bold=True, color=C_DARK_FG, size=14)
c.fill = make_fill(C_DARK_BG)
c.alignment = make_alignment("center", "center")
ws.row_dimensions[1].height = 28

# ── Yesterday section ────────────────────────────────────────────────────
ws.merge_cells("A2:L2")
cell = ws["A2"]
cell.value = "YESTERDAY — 2026-05-29 (Active Campaigns)"
cell.font = make_font(bold=True, color=C_DARK_FG)
cell.fill = make_fill(C_SECTION)
cell.alignment = make_alignment("left", "center")

yday_headers = [
    "Campaign Name", "Status", "Spend (INR)", "Impressions", "Reach",
    "Clicks", "CTR %", "CPC (INR)", "Purchases", "Cost/Purchase (INR)", "ROAS", "Daily Budget"
]
apply_table_header(ws, 3, 1, yday_headers)

for ridx, row in enumerate(campaigns_yday):
    r = 4 + ridx
    fill_hex = C_GRAY_ROW if ridx % 2 == 0 else C_WHITE
    for cidx, val in enumerate(row):
        cl = get_column_letter(1 + cidx)
        cell = ws[f"{cl}{r}"]
        cell.value = val if val is not None else "N/A"
        cell.border = make_border()
        cell.alignment = make_alignment("center")
        cell.fill = make_fill(fill_hex)
        if cidx == 0:
            cell.alignment = make_alignment("left")
        if cidx in {2, 7, 9} and val is not None:
            cell.number_format = '₹#,##0.00'
        elif cidx == 6 and val is not None:
            cell.number_format = '0.00%'
        elif cidx == 10 and val is not None:
            cell.number_format = '0.00'

# Totals row
totals_row = 9
ws.merge_cells(f"A{totals_row}:B{totals_row}")
cell = ws[f"A{totals_row}"]
cell.value = "TOTALS / AVERAGES"
cell.font = make_font(bold=True)
cell.fill = make_fill(C_GRAY_ROW)
cell.border = make_border()
cell.alignment = make_alignment("center")

totals_vals = [None, None, 39366.65, 259159, 225484, 5665, 2.18, 6.95, 51, 771.90, 3.51, None]
for cidx, val in enumerate(totals_vals):
    if cidx in {0, 1}:
        continue
    cl = get_column_letter(1 + cidx)
    cell = ws[f"{cl}{totals_row}"]
    cell.value = val
    cell.font = make_font(bold=True)
    cell.fill = make_fill(C_GRAY_ROW)
    cell.border = make_border()
    cell.alignment = make_alignment("center")
    if cidx in {2, 7, 9} and val is not None:
        cell.number_format = '₹#,##0.00'
    elif cidx == 6 and val is not None:
        cell.number_format = '0.00%'
    elif cidx == 10 and val is not None:
        cell.number_format = '0.00'

# spacer
ws.row_dimensions[10].height = 8

# ── 7-Day section ────────────────────────────────────────────────────────
ws.merge_cells("A11:J11")
cell = ws["A11"]
cell.value = "LAST 7 DAYS (May 24–30) — Campaign Totals"
cell.font = make_font(bold=True, color=C_DARK_FG)
cell.fill = make_fill(C_DARK_BG)
cell.alignment = make_alignment("left", "center")

headers_7d = [
    "Campaign Name", "Total Spend (INR)", "Impressions", "Reach",
    "Clicks", "CTR %", "CPC (INR)", "Purchases", "Cost/Purchase (INR)", "ROAS"
]
apply_table_header(ws, 12, 1, headers_7d)

for ridx, row in enumerate(campaigns_7d):
    r = 13 + ridx
    fill_hex = C_GRAY_ROW if ridx % 2 == 0 else C_WHITE
    for cidx, val in enumerate(row):
        cl = get_column_letter(1 + cidx)
        cell = ws[f"{cl}{r}"]
        cell.value = val
        cell.border = make_border()
        cell.alignment = make_alignment("center")
        cell.fill = make_fill(fill_hex)
        if cidx == 0:
            cell.alignment = make_alignment("left")
        if cidx in {1, 6, 8}:
            cell.number_format = '₹#,##0.00'
        elif cidx == 5:
            cell.number_format = '0.00%'
        elif cidx == 9:
            cell.number_format = '0.00'

# 7d totals
totals_7d_row = 18
ws.merge_cells(f"A{totals_7d_row}:A{totals_7d_row}")
cell = ws[f"A{totals_7d_row}"]
cell.value = "7-DAY TOTALS"
cell.font = make_font(bold=True)
cell.fill = make_fill(C_GRAY_ROW)
cell.border = make_border()
cell.alignment = make_alignment("center")

t7d = [None, 307252.52, 1801625, 1177924, 41678, 2.31, 8.47, 385, 798.06, 3.21]
for cidx, val in enumerate(t7d):
    if cidx == 0:
        continue
    cl = get_column_letter(1 + cidx)
    cell = ws[f"{cl}{totals_7d_row}"]
    cell.value = val
    cell.font = make_font(bold=True)
    cell.fill = make_fill(C_GRAY_ROW)
    cell.border = make_border()
    cell.alignment = make_alignment("center")
    if cidx in {1, 6, 8} and val is not None:
        cell.number_format = '₹#,##0.00'
    elif cidx == 5:
        cell.number_format = '0.00%'
    elif cidx == 9:
        cell.number_format = '0.00'

ws.auto_filter.ref = "A3:L3"
ws.freeze_panes = "A4"

ws.column_dimensions["A"].width = 34
for col in ["B","C","D","E","F","G","H","I","J","K","L"]:
    ws.column_dimensions[col].width = 16


# ═══════════════════════════════════════════════════════════════════════════
# SHEET 4: Product Performance
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_prod

# Title
ws.merge_cells("A1:H1")
c = ws["A1"]
c.value = "DHIRAI — Product Performance  |  2026-05-29"
c.font = make_font(bold=True, color=C_DARK_FG, size=14)
c.fill = make_fill(C_DARK_BG)
c.alignment = make_alignment("center", "center")
ws.row_dimensions[1].height = 28

prod_headers = ["Rank", "Product Name", "Gross Sales", "Net Sales", "Orders",
                "Avg Order Value", "Returns", "Net Margin %"]
apply_table_header(ws, 2, 1, prod_headers)

for ridx, (title, gross, net, orders, aov, returns) in enumerate(products_yesterday):
    r = 3 + ridx
    fill_hex = C_GRAY_ROW if ridx % 2 == 0 else C_WHITE
    margin_pct = (net / gross * 100) if gross > 0 else 0.0

    row_data = [ridx + 1, title, gross, net, orders, aov, returns, margin_pct / 100.0]
    fmts = [None, None, '₹#,##0.00', '₹#,##0.00', None, '₹#,##0.00', '₹#,##0.00', '0.0%']

    for cidx, (val, fmt) in enumerate(zip(row_data, fmts)):
        cl = get_column_letter(1 + cidx)
        cell = ws[f"{cl}{r}"]
        cell.value = val
        cell.border = make_border()
        cell.fill = make_fill(fill_hex)
        cell.alignment = make_alignment("center")
        if fmt:
            cell.number_format = fmt
        if cidx == 1:
            cell.alignment = make_alignment("left", wrap=True)

    # Conditional on Returns column (col G = index 6)
    ret_cell = ws[f"G{r}"]
    if returns > 0 and returns >= gross:
        ret_cell.fill = make_fill(C_RED_CF)
        ret_cell.font = make_font(bold=True, color=C_RED_FONT)
    elif returns > 0:
        ret_cell.fill = make_fill(C_ORANGE_CF)
        ret_cell.font = make_font(bold=True, color=C_ORANGE_FN)

# autofilter, freeze
ws.auto_filter.ref = "A2:H2"
ws.freeze_panes = "A3"

ws.column_dimensions["A"].width = 6
ws.column_dimensions["B"].width = 55
for col in ["C","D","F","G"]:
    ws.column_dimensions[col].width = 16
ws.column_dimensions["E"].width = 8
ws.column_dimensions["H"].width = 14

# Row heights for product names
for r in range(3, 3 + len(products_yesterday)):
    ws.row_dimensions[r].height = 30


# ═══════════════════════════════════════════════════════════════════════════
# SHEET 5: Campaign Performance
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_camp

# Title
ws.merge_cells("A1:M1")
c = ws["A1"]
c.value = "DHIRAI — Campaign Performance  |  Yesterday vs 7-Day Comparison"
c.font = make_font(bold=True, color=C_DARK_FG, size=14)
c.fill = make_fill(C_DARK_BG)
c.alignment = make_alignment("center", "center")
ws.row_dimensions[1].height = 28

camp_headers = [
    "Campaign", "Status", "Daily Budget (INR)",
    "Yday Spend", "Yday Purchases", "Yday ROAS", "Yday CTR %", "Yday CPC",
    "7D Total Spend", "7D Purchases", "7D ROAS", "7D Daily Avg Spend",
    "Performance Rating"
]
apply_table_header(ws, 2, 1, camp_headers)

# Build combined data (match by campaign name)
camp_7d_dict = {r[0]: r for r in campaigns_7d}

for ridx, yd in enumerate(campaigns_yday):
    camp_name, status, yd_spend, _, _, _, yd_ctr, yd_cpc, yd_purch, _, yd_roas, budget = yd
    s7 = camp_7d_dict.get(camp_name)
    if s7:
        _, s7_spend, _, _, _, _, _, s7_purch, _, s7_roas = s7
        s7_daily_avg = s7_spend / 7.0
    else:
        s7_spend = s7_purch = s7_roas = s7_daily_avg = None

    # Rating
    if yd_roas is None:
        rating = "NO DATA"
    elif yd_roas >= 4:
        rating = "EXCELLENT"
    elif yd_roas >= 2:
        rating = "GOOD"
    else:
        rating = "WEAK"

    r = 3 + ridx
    fill_hex = C_GRAY_ROW if ridx % 2 == 0 else C_WHITE

    row_data = [
        camp_name, status, budget if isinstance(budget, (int, float)) else budget,
        yd_spend, yd_purch, yd_roas, yd_ctr / 100.0 if yd_ctr else None,
        yd_cpc,
        s7_spend, s7_purch, s7_roas, s7_daily_avg,
        rating
    ]
    fmts = [
        None, None, None,
        '₹#,##0.00', None, '0.00', '0.00%', '₹#,##0.00',
        '₹#,##0.00', None, '0.00', '₹#,##0.00',
        None
    ]

    for cidx, (val, fmt) in enumerate(zip(row_data, fmts)):
        cl = get_column_letter(1 + cidx)
        cell = ws[f"{cl}{r}"]
        # Handle N/A display
        if val is None:
            cell.value = "N/A"
        else:
            cell.value = val
        cell.border = make_border()
        cell.fill = make_fill(fill_hex)
        cell.alignment = make_alignment("center")
        if fmt and val is not None:
            cell.number_format = fmt
        if cidx == 0:
            cell.alignment = make_alignment("left", wrap=True)

    # Rating colour
    rating_cell = ws[f"M{r}"]
    rating_colors = {
        "EXCELLENT": (C_GREEN_CF, C_GREEN_FONT),
        "GOOD":      (C_ORANGE_CF, C_ORANGE_FN),
        "WEAK":      (C_RED_CF, C_RED_FONT),
        "NO DATA":   ("DDDDDD", "666666"),
    }
    bg, fg = rating_colors.get(rating, (fill_hex, "000000"))
    rating_cell.fill = make_fill(bg)
    rating_cell.font = make_font(bold=True, color=fg)

ws.auto_filter.ref = "A2:M2"
ws.freeze_panes = "A3"

ws.column_dimensions["A"].width = 34
for col in ["B","C","D","E","F","G","H","I","J","K","L","M"]:
    ws.column_dimensions[col].width = 16

for r in range(3, 3 + len(campaigns_yday)):
    ws.row_dimensions[r].height = 24


# ═══════════════════════════════════════════════════════════════════════════
# SHEET 6: Recommendations & Notes
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_rec

ws.merge_cells("A1:E1")
c = ws["A1"]
c.value = "DHIRAI — Recommendations & Notes  |  Report Date: 2026-05-29"
c.font = make_font(bold=True, color=C_DARK_FG, size=14)
c.fill = make_fill(C_DARK_BG)
c.alignment = make_alignment("center", "center")
ws.row_dimensions[1].height = 28

info_rows = [
    ("Report Date:",   "2026-05-29"),
    ("Data Sources:",  "Shopify Analytics (Asia/Kolkata timezone), Meta Ads API"),
]
for ridx, (label, val) in enumerate(info_rows):
    r = 2 + ridx
    ws[f"A{r}"].value = label
    ws[f"A{r}"].font = make_font(bold=True)
    ws[f"A{r}"].alignment = make_alignment("left")
    ws.merge_cells(f"B{r}:E{r}")
    ws[f"B{r}"].value = val
    ws[f"B{r}"].alignment = make_alignment("left")

# Data Availability Notes
ws.row_dimensions[4].height = 8
section_header(ws, 5, 1, "  DATA AVAILABILITY NOTES", C_SECTION, colspan=5)
notes = [
    "* Shopify inventory data: Unavailable (rate limited during fetch — run separately)",
    "* Shopify abandoned checkouts: Not available via analytics API (requires Shopify admin)",
    "* Meta Ads: ROAS Goal campaign shows 0 purchases despite active spend — verify pixel attribution",
    "* Meta Ads date_start returned '30 May' for YESTERDAY preset — possible timezone offset; data appears consistent",
    "* Checkout completion rate: 0% on 2026-05-29 — verify Shopify checkout tracking integration",
]
for ridx, note in enumerate(notes):
    r = 6 + ridx
    ws.merge_cells(f"A{r}:E{r}")
    cell = ws[f"A{r}"]
    cell.value = note
    cell.font = make_font(italic=True, color="444444")
    cell.alignment = make_alignment("left", "center", wrap=True)
    cell.fill = make_fill("F2F2F2")
    ws.row_dimensions[r].height = 18

ws.row_dimensions[11].height = 10

# Recommendations Table
section_header(ws, 12, 1, "  RECOMMENDED ACTIONS", C_BLUE_H, colspan=5)
rec_headers = ["Priority", "Action", "Expected Impact"]
ws.merge_cells("B13:C13")
ws.merge_cells("D13:E13")
for cidx, h in enumerate(rec_headers):
    if cidx == 0:
        cl = "A"
    elif cidx == 1:
        cl = "B"
    else:
        cl = "D"
    cell = ws[f"{cl}13"]
    cell.value = h
    cell.font = make_font(bold=True, color=C_DARK_FG)
    cell.fill = make_fill(C_HEADER_BG)
    cell.alignment = make_alignment("center")
    cell.border = make_border()

rec_data = [
    (1, "Scale Summer Special Campaign budget +25%",
        "More purchases at ROAS 6.31"),
    (2, 'Pause or reduce "Roas Goal" campaign immediately',
        "Save ₹3,600/day with 0 purchase output"),
    (3, "Audit checkout funnel — possible broken tracking",
        "Fix potential revenue loss from session→purchase gap"),
    (4, 'Investigate "The Essential" returns (₹14,393 in 1 day)',
        "Reduce return rate, improve margins"),
    (5, 'Review "Adira" product listing — 100% return rate',
        "Product/size issue, pull from ads"),
]
priority_colors = {1: "C00000", 2: "C00000", 3: "FF7C00", 4: "FF7C00", 5: "2E75B6"}
for ridx, (pri, action, impact) in enumerate(rec_data):
    r = 14 + ridx
    fill_hex = C_GRAY_ROW if ridx % 2 == 0 else C_WHITE

    # Priority
    cell = ws[f"A{r}"]
    cell.value = pri
    cell.font = make_font(bold=True, color=C_DARK_FG)
    cell.fill = make_fill(priority_colors.get(pri, C_SECTION))
    cell.alignment = make_alignment("center")
    cell.border = make_border()

    # Action
    ws.merge_cells(f"B{r}:C{r}")
    cell = ws[f"B{r}"]
    cell.value = action
    cell.fill = make_fill(fill_hex)
    cell.alignment = make_alignment("left", "center", wrap=True)
    cell.border = make_border()
    ws.row_dimensions[r].height = 24

    # Impact
    ws.merge_cells(f"D{r}:E{r}")
    cell = ws[f"D{r}"]
    cell.value = impact
    cell.fill = make_fill(fill_hex)
    cell.alignment = make_alignment("left", "center", wrap=True)
    cell.border = make_border()

ws.column_dimensions["A"].width = 10
ws.column_dimensions["B"].width = 25
ws.column_dimensions["C"].width = 30
ws.column_dimensions["D"].width = 25
ws.column_dimensions["E"].width = 25


# ═══════════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════════
output_path = "/home/user/claude-routines/Daily_Store_Ads_Performance_Sheet_2026-05-29.xlsx"
wb.save(output_path)

if os.path.exists(output_path):
    size_kb = os.path.getsize(output_path) / 1024
    print(f"SUCCESS: file created at {output_path}")
    print(f"File size: {size_kb:.1f} KB")
    print(f"Sheets: {[s.title for s in wb.worksheets]}")
else:
    print("ERROR: file was not created!")
