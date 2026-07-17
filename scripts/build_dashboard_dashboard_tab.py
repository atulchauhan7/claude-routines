#!/usr/bin/env python3
"""Part 2: adds the Dashboard tab (formulas + charts) to the workbook built by build_dashboard.py."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.axis import ChartLines

OUT_PATH = "/home/user/claude-routines/Daily Store & Ads Performance Sheet.xlsx"
FONT_NAME = "Arial"

C_HEADER_BG   = "1F2937"
C_HEADER_FG   = "FFFFFF"
C_ACCENT_BG   = "2563EB"
C_TITLE_BG    = "111827"
C_GREEN       = "C6EFCE"
C_RED         = "FFC7CE"
C_YELLOW      = "FFEB9C"

INR = '"₹"#,##0.00'
PCT1 = '0.0%'
PCT2 = '0.00%'
ROASFMT = '0.00"x"'

wb = load_workbook(OUT_PATH)

# Resolve row extents from the data tabs already written
ws_sd = wb["Shopify Daily Data"]
ws_md = wb["Meta Ads Daily Data"]
n_days = ws_sd.max_row - 1               # data rows excluding header
last_row = n_days + 1                    # sheet row of most recent day
prev_start = max(2, last_row - 7)
prev_end = last_row - 1
YESTERDAY_LABEL = ws_sd.cell(row=last_row, column=1).value

ws = wb["Dashboard"]


def style_title(ws, row, ncols, text, size=16):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(name=FONT_NAME, bold=True, color=C_HEADER_FG, size=size)
    cell.fill = PatternFill("solid", fgColor=C_TITLE_BG)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 28


def style_section(ws, row, ncols, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(name=FONT_NAME, bold=True, color=C_HEADER_FG, size=12)
    cell.fill = PatternFill("solid", fgColor=C_ACCENT_BG)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 22


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT_NAME, bold=True, color=C_HEADER_FG, size=10.5)
        cell.fill = PatternFill("solid", fgColor=C_HEADER_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color="9CA3AF"))


style_title(ws, 1, 6, f"DAILY STORE & ADS PERFORMANCE — {YESTERDAY_LABEL}", size=16)

# ── Executive summary ────────────────────────────────────────────────────
r = 3
style_section(ws, r, 6, "EXECUTIVE SUMMARY")
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r+3, end_column=6)
summary_text = (
    "Yesterday was a solid day: net sales and orders both beat the 7-day average while returns and discounts "
    "came in lower than usual. Ad efficiency also improved — blended ROAS reached 3.94x (7-day avg 3.64x) on "
    "slightly lower spend. The main watch-items are one underperforming campaign (\"Black DHURANDHAR\"), a "
    "narrow-audience warning on one ad set, and two products that are completely out of stock."
)
cell = ws.cell(row=r, column=1, value=summary_text)
cell.alignment = Alignment(wrap_text=True, vertical="top")
cell.font = Font(name=FONT_NAME, size=11)
r += 5

# ── Shopify KPI table (formula-driven) ───────────────────────────────────
style_section(ws, r, 6, "SHOPIFY PERFORMANCE — Yesterday vs Previous 7-Day Average")
r += 1
hdr_row = r
for i, h in enumerate(["Metric", "Yesterday", "7-Day Avg", "Change vs Avg", "", ""], start=1):
    ws.cell(row=hdr_row, column=i, value=h)
style_header(ws, hdr_row, 6)
r += 1

shopify_kpis = [
    ("Gross Sales", "C", INR),
    ("Net Sales", "F", INR),
    ("Orders", "B", '#,##0'),
    ("Average Order Value", "J", INR),
    ("Discounts", "D", INR),
    ("Returns", "E", INR),
]
shopify_kpi_start = r
for label, col, fmt in shopify_kpis:
    ws.cell(row=r, column=1, value=label)
    ws.cell(row=r, column=2, value=f"='Shopify Daily Data'!{col}{last_row}").number_format = fmt
    ws.cell(row=r, column=3, value=f"=AVERAGE('Shopify Daily Data'!{col}{prev_start}:{col}{prev_end})").number_format = fmt
    ws.cell(row=r, column=4, value=f"=(B{r}-C{r})/C{r}").number_format = PCT1
    r += 1
shopify_kpi_end = r - 1

r += 1
style_section(ws, r, 6, "META ADS PERFORMANCE — Yesterday vs Previous 7-Day Average")
r += 1
hdr_row2 = r
for i, h in enumerate(["Metric", "Yesterday", "7-Day Avg", "Change vs Avg", "", ""], start=1):
    ws.cell(row=hdr_row2, column=i, value=h)
style_header(ws, hdr_row2, 6)
r += 1

meta_kpis = [
    ("Ad Spend", "B", INR),
    ("Impressions", "C", '#,##0'),
    ("Clicks", "E", '#,##0'),
    ("CTR", "F", PCT2),
    ("CPC", "G", INR),
    ("Purchases (est.)", "I", '#,##0'),
    ("CPA", "J", INR),
    ("ROAS", "K", ROASFMT),
]
meta_kpi_start = r
for label, col, fmt in meta_kpis:
    ws.cell(row=r, column=1, value=label)
    ws.cell(row=r, column=2, value=f"='Meta Ads Daily Data'!{col}{last_row}").number_format = fmt
    ws.cell(row=r, column=3, value=f"=AVERAGE('Meta Ads Daily Data'!{col}{prev_start}:{col}{prev_end})").number_format = fmt
    ws.cell(row=r, column=4, value=f"=(B{r}-C{r})/C{r}").number_format = PCT1
    r += 1
meta_kpi_end = r - 1

# conditional format the "Change vs Avg" columns green/red
for start, end in [(shopify_kpi_start, shopify_kpi_end), (meta_kpi_start, meta_kpi_end)]:
    rng = f"D{start}:D{end}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=C_GREEN)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=C_RED)))

# ── Key wins / issues / actions (linked reference to Recommendations tab) ─
r += 1
style_section(ws, r, 6, "KEY WINS")
r += 1
wins = [
    "Retargeting- new: 5.52x ROAS yesterday, lowest CPA (₹416) — scale it.",
    "\"New Campaign- for recovery\" ROAS jumped from 2.46x (7-day avg) to 3.90x.",
    "Net sales +7.5% vs 7-day avg with lower returns and discounts than usual.",
]
for w in wins:
    ws.cell(row=r, column=2, value=w)
    r += 1

r += 1
style_section(ws, r, 6, "KEY ISSUES")
r += 1
issues = [
    "\"Black DHURANDHAR\" consistently weakest campaign (3.37x-3.41x ROAS).",
    "Ad set \"Black and multiple\" flagged for narrow audience by Meta.",
    "2 products fully out of stock (see Product Performance tab).",
]
for i in issues:
    ws.cell(row=r, column=2, value=i)
    r += 1

r += 1
style_section(ws, r, 6, "RECOMMENDED ACTIONS TODAY")
r += 1
actions = [
    "1. Scale \"Retargeting- new\" budget +20-30%.",
    "2. Review/optimize \"Black DHURANDHAR\" creative & targeting.",
    "3. Expand audience on \"Black and multiple\" ad set.",
    "4. Restock or hide the 2 out-of-stock listings.",
    "5. Watch stock on the top 2 products (62% of yesterday's product revenue).",
]
for a in actions:
    ws.cell(row=r, column=2, value=a)
    r += 1

dashboard_text_end = r

ws.column_dimensions["A"].width = 22
for col in "BCDEF":
    ws.column_dimensions[col].width = 34
ws.freeze_panes = "A2"

# ── Charts ────────────────────────────────────────────────────────────────
chart_anchor_row = dashboard_text_end + 2

# Sales trend (Gross + Net) line chart
sales_chart = LineChart()
sales_chart.title = "Sales Trend (Gross vs Net)"
sales_chart.style = 2
sales_chart.y_axis.title = "₹"
sales_chart.x_axis.title = "Date"
sales_chart.height = 8
sales_chart.width = 16
data = Reference(ws_sd, min_col=3, max_col=3, min_row=1, max_row=last_row)
data2 = Reference(ws_sd, min_col=6, max_col=6, min_row=1, max_row=last_row)
cats = Reference(ws_sd, min_col=1, min_row=2, max_row=last_row)
sales_chart.add_data(data, titles_from_data=True)
sales_chart.add_data(data2, titles_from_data=True)
sales_chart.set_categories(cats)
ws.add_chart(sales_chart, f"A{chart_anchor_row}")

# Orders trend
orders_chart = LineChart()
orders_chart.title = "Orders Trend"
orders_chart.y_axis.title = "Orders"
orders_chart.x_axis.title = "Date"
orders_chart.height = 8
orders_chart.width = 16
data = Reference(ws_sd, min_col=2, max_col=2, min_row=1, max_row=last_row)
orders_chart.add_data(data, titles_from_data=True)
orders_chart.set_categories(cats)
ws.add_chart(orders_chart, f"H{chart_anchor_row}")

chart_row2 = chart_anchor_row + 17

# Ad spend trend
spend_chart = LineChart()
spend_chart.title = "Ad Spend Trend"
spend_chart.y_axis.title = "₹"
spend_chart.x_axis.title = "Date"
spend_chart.height = 8
spend_chart.width = 16
data = Reference(ws_md, min_col=2, max_col=2, min_row=1, max_row=last_row)
cats_md = Reference(ws_md, min_col=1, min_row=2, max_row=last_row)
spend_chart.add_data(data, titles_from_data=True)
spend_chart.set_categories(cats_md)
ws.add_chart(spend_chart, f"A{chart_row2}")

# ROAS trend
roas_chart = LineChart()
roas_chart.title = "ROAS Trend"
roas_chart.y_axis.title = "ROAS (x)"
roas_chart.x_axis.title = "Date"
roas_chart.height = 8
roas_chart.width = 16
data = Reference(ws_md, min_col=11, max_col=11, min_row=1, max_row=last_row)
roas_chart.add_data(data, titles_from_data=True)
roas_chart.set_categories(cats_md)
ws.add_chart(roas_chart, f"H{chart_row2}")

chart_row3 = chart_row2 + 17

# Top products bar chart (from Product Performance tab)
ws_pp = wb["Product Performance"]
# product rows start at row 4 (title row1, spacer row2, header row3) per build_dashboard.py
prod_hdr_row = 3
prod_first = 4
prod_last = 13  # 10 products
top_chart = BarChart()
top_chart.type = "bar"
top_chart.title = "Top Products by Gross Sales (Yesterday)"
top_chart.y_axis.title = "Product"
top_chart.x_axis.title = "₹"
top_chart.height = 9
top_chart.width = 16
data = Reference(ws_pp, min_col=2, max_col=2, min_row=prod_hdr_row, max_row=prod_last)
cats = Reference(ws_pp, min_col=1, min_row=prod_first, max_row=prod_last)
top_chart.add_data(data, titles_from_data=True)
top_chart.set_categories(cats)
ws.add_chart(top_chart, f"A{chart_row3}")

# Top campaigns bar chart (from Campaign Performance tab)
ws_cp = wb["Campaign Performance"]
camp_hdr_row = 3
camp_first = 4
camp_last = 7  # 4 campaigns
camp_chart = BarChart()
camp_chart.type = "col"
camp_chart.title = "Campaign Spend — Yesterday"
camp_chart.y_axis.title = "₹"
camp_chart.height = 9
camp_chart.width = 16
data = Reference(ws_cp, min_col=2, max_col=2, min_row=camp_hdr_row, max_row=camp_last)
cats = Reference(ws_cp, min_col=1, min_row=camp_first, max_row=camp_last)
camp_chart.add_data(data, titles_from_data=True)
camp_chart.set_categories(cats)
ws.add_chart(camp_chart, f"H{chart_row3}")

wb.save(OUT_PATH)
print("Dashboard + charts added. Saved.")
print(f"last_row={last_row}, prev_start={prev_start}, prev_end={prev_end}")
