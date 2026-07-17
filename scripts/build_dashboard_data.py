#!/usr/bin/env python3
"""
Builds/updates the 'Daily Store & Ads Performance Sheet' workbook for Dhirai
(Shopify store dhirai.in + Meta Ads account 979830497515712).

Data below was pulled live via MCP tools (Shopify Admin/ShopifyQL + Meta Ads
insights) for the Asia/Kolkata calendar day 2026-07-16 ("yesterday" relative
to run date 2026-07-17), plus the trailing 7 days 2026-07-09..2026-07-15.

Re-running this script is idempotent: if the target file already exists, it
loads the existing 'Shopify Daily Data' / 'Meta Ads Daily Data' tabs, skips
any date already present, appends only new dates, and rebuilds the Dashboard
formulas/charts against the current row extent.
"""
import datetime
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT_PATH = "/home/user/claude-routines/Daily Store & Ads Performance Sheet.xlsx"

FONT_NAME = "Arial"

# ── Colors ──────────────────────────────────────────────────────────────────
C_HEADER_BG   = "1F2937"   # dark slate
C_HEADER_FG   = "FFFFFF"
C_ACCENT_BG   = "2563EB"   # blue
C_SECTION_BG  = "E5E7EB"
C_GREEN       = "C6EFCE"
C_GREEN_TXT   = "1B5E20"
C_RED         = "FFC7CE"
C_RED_TXT     = "9C0006"
C_YELLOW      = "FFEB9C"
C_YELLOW_TXT  = "9C6500"
C_TITLE_BG    = "111827"

TAB_NAMES = [
    "Dashboard",
    "Shopify Daily Data",
    "Meta Ads Daily Data",
    "Product Performance",
    "Campaign Performance",
    "Recommendations & Notes",
]

# ─────────────────────────────────────────────────────────────────────────
# RAW DATA (fetched live via MCP this run)
# ─────────────────────────────────────────────────────────────────────────

RUN_DATE = "2026-07-17"
YESTERDAY = "2026-07-16"

# Shopify daily sales: date, orders, gross_sales, discounts, returns, net_sales,
# shipping, taxes, total_sales, aov
SHOPIFY_DAILY = [
    ["2026-07-09", 73, 190834.80, -1862.32, -2499.00, 186473.48, 0.00, 381.20, 186854.68, 2588.66],
    ["2026-07-10", 50, 121913.25, -1613.22,     0.00, 120300.03, 0.00,   0.00, 120833.78, 2406.00],
    ["2026-07-11", 55, 130491.00, -1759.27,     0.00, 128731.73, 0.00,   0.00, 128731.73, 2340.58],
    ["2026-07-12", 87, 269307.80,  -839.65,-47285.00, 221183.15, 0.00, 381.20, 221564.35, 3085.84],
    ["2026-07-13", 55, 142333.07, -3105.69,     0.00, 139227.38, 0.00, 304.93, 139532.31, 2531.41],
    ["2026-07-14", 57, 144836.00,  -314.85, -1999.00, 142522.15, 0.00,   0.00, 142522.15, 2535.46],
    ["2026-07-15", 64, 166129.00, -1092.05,     0.00, 165036.95, 0.00,   0.00, 165036.95, 2578.70],
    ["2026-07-16", 69, 172575.00,  -914.60, -2199.00, 169461.40, 0.00,   0.00, 169461.40, 2487.83],
]

# Meta Ads daily (account level): date, spend, impressions, reach, clicks, ctr%,
# cpc, cpm, purchases(est. spend/cpa), cpa, roas
META_DAILY = [
    ["2026-07-09", 43342.39, 258075, 206133, 6024, 2.33, 7.19, 167.94, 66, 656.70, 3.87],
    ["2026-07-10", 41607.98, 247234, 207925, 5570, 2.25, 7.47, 168.29, 47, 885.28, 2.70],
    ["2026-07-11", 35653.10, 216060, 180186, 4684, 2.17, 7.61, 165.01, 53, 672.70, 3.49],
    ["2026-07-12", 47584.42, 259228, 214990, 5765, 2.22, 8.25, 183.56, 84, 566.48, 5.45],
    ["2026-07-13", 45884.99, 236170, 184417, 5178, 2.19, 8.86, 194.29, 54, 849.72, 3.00],
    ["2026-07-14", 44951.34, 232320, 189826, 5440, 2.34, 8.26, 193.49, 57, 788.62, 3.17],
    ["2026-07-15", 43931.17, 254112, 199911, 6195, 2.44, 7.09, 172.88, 61, 720.18, 3.62],
    ["2026-07-16", 43466.32, 279460, 216156, 6412, 2.29, 6.78, 155.54, 69, 629.95, 3.94],
]

# Top products - yesterday (2026-07-16): title, gross_sales, net_sales, orders
PRODUCTS_YESTERDAY = [
    ["Ira Rayon Co-ord Set for Women | Relaxed Fit Shirt Top with Draped Pants", 42483.00, 42108.15, 16],
    ["\"The Essential\" - White Cotton Shirt & Pant Co-ord Set", 36782.00, 34583.00, 18],
    ["Eclipse Grace Premium Crepe Salwar Suit set | DHIRAI", 8997.00, 8997.00, 3],
    ["The Dhurandhar Black Short Kurta Set | A Dhirai Essential", 6747.00, 6522.10, 3],
    ["VAIRA Airy Linen Everyday Co-ord Set | Women's Shirt and Pant", 6097.00, 6097.00, 3],
    ["\"Inayat\" — Power Pastel Cotton Shirt & Trouser Co-ord Set", 5997.00, 5897.05, 3],
    ["Everyday Elegance Cotton Kurta Pant Set", 5997.00, 5997.00, 3],
    ["\"Rakta\" - Crimson Red Shalwar Suit Set", 5098.00, 5098.00, 2],
    ["\"Maira\" — Premium Pure Cotton Minimalist Co-ord Set", 4998.00, 4998.00, 2],
    ["Premium Cotton Flex Tie-Up Kurta Set for Women — Embroidered Ethnic Suit", 4998.00, 4998.00, 2],
]

# Inventory snapshot (lowest ending-stock products): title, starting, ending, sold, sell_through
INVENTORY_SNAPSHOT = [
    ["\"Zariya\" Stone Hoops", 0, 0, 0, None],
    ["\"Utsav\" - Mal Cotton Full Flare Anarkali Suit Set with Silver Gota", 0, 0, 0, None],
    ["\"Nazakat\" Long Chain Jhumki Earrings", 150, 150, 0, 0.0],
    ["Eclipse Grace Premium Crepe Salwar Suit set | DHIRAI", 315, 312, 3, 0.0095],
    ["Sage Aura cotton Kurta Set | DHIRAI", 357, 357, 0, 0.0],
    ["Rani Grace kurta  Set", 398, 398, 0, 0.0],
    ["Love Note Co-Ord Set", 399, 399, 0, 0.0],
    ["Monochrome Aura Kurta Set | DHIRAI", 399, 399, 0, 0.0],
    ["Sage Serenity Cotton Co-Ord Set for Women | DHIRAI", 400, 400, 0, 0.0],
    ["Berry Bloom Farshi Salwar Set", 400, 400, 0, 0.0],
    ["Caramel Grace Co-Ord Set| DHIRAI", 400, 400, 0, 0.0],
    ["Saffron Grace Kurta Set| DHIRAI", 400, 400, 0, 0.0],
    ["Rust Bloom Co-Ord Set", 400, 400, 0, 0.0],
    ["Sunstone Co-Ord Set", 400, 400, 0, 0.0],
    ["\"Kashish\" - Alluring Purple Chanderi Anarkali Suit Set", 499, 499, 0, 0.0],
]

# Campaigns - yesterday (2026-07-16): name, spend, impressions, clicks, ctr%, cpc, roas, cpa
CAMPAIGNS_YESTERDAY = [
    ["New Campaign- for recovery", 14199.54, 137109, 1632, 1.19, 8.70, 3.90, 617.37],
    ["Dhirai Scale -ASC",          14184.09,  72916, 2198, 3.01, 6.45, 3.51, 709.20],
    ["Black DHURANDHAR",            8010.41,  50187, 1852, 3.69, 4.33, 3.37, 890.05],
    ["Retargeting- new",             7072.28,  19248,  730, 3.79, 9.69, 5.52, 416.02],
]

# Campaigns - 7-day aggregate (2026-07-09..2026-07-15): name, spend, impressions, clicks, ctr%, cpc, roas, cpa
CAMPAIGNS_7DAY = [
    ["New Campaign- for recovery", 98773.77, 759706, 10137, 1.33, 9.74, 2.46, 1018.29],
    ["Dhirai Scale -ASC",          90332.73, 477149, 11860, 2.49, 7.62, 3.56,  669.13],
    ["Retargeting- new",           58412.14, 167225,  6401, 3.83, 9.13, 5.98,  503.55],
    ["Black DHURANDHAR",           55436.75, 299119, 10458, 3.50, 5.30, 3.41,  749.15],
]

ANOMALY_NOTE = ("Narrow audience warning: ad set \"Black and multiple\" (ID 120238324193200464) "
                "has an actively-delivering narrow audience (est. size 22,400-26,400). "
                "Meta recommends expanding the audience or testing Advantage+ audience.")

UNAVAILABLE_METRICS = [
    "Failed payments: not exposed by the connected Shopify tools/ShopifyQL in this session — check Shopify Admin > Settings > Payments for declined transaction reports.",
    "Abandoned checkouts: not exposed by the connected Shopify tools/ShopifyQL in this session — check Shopify Admin > Analytics or Abandoned Checkouts report.",
    "Order cancellations (distinct from refunds/returns): not separately broken out by the connected tools — Returns (₹2,199 on 2026-07-16) is used as the closest proxy for money given back to customers.",
    "Units sold per product (distinct from orders): not returned by the ShopifyQL sales query used — Orders count is shown instead as a volume proxy.",
    "Meta Ads 'Purchases' count: the connected Meta Ads tool did not expose a raw conversions/results field for this account — shown value is estimated as Spend ÷ Cost-per-purchase (cost_per_omni_purchase), both of which ARE reported by the tool.",
    "Google Sheets connector is not installed for this account, so this dashboard was delivered as an Excel (.xlsx) workbook instead of a live Google Sheet, and committed to the claude-routines repo so tomorrow's run can append to it without duplicating rows.",
    "Gmail connector only exposes draft creation (no direct-send tool) in this session, so the daily notification was created as a Gmail draft rather than sent automatically.",
]

# ─────────────────────────────────────────────────────────────────────────
# Style helpers
# ─────────────────────────────────────────────────────────────────────────

def style_header(ws, row, ncols, bg=C_HEADER_BG, fg=C_HEADER_FG, size=11):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT_NAME, bold=True, color=fg, size=size)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color="9CA3AF"))


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


def autofit(ws, min_width=10, max_width=55):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            col = cell.column_letter
            length = len(str(cell.value))
            widths[col] = max(widths.get(col, 0), length)
    for col, w in widths.items():
        ws.column_dimensions[col].width = max(min_width, min(max_width, w + 3))


def set_font_all(ws):
    for row in ws.iter_rows():
        for cell in row:
            if cell.font is None or cell.font.name != FONT_NAME:
                existing = cell.font
                cell.font = Font(
                    name=FONT_NAME,
                    bold=existing.bold if existing else False,
                    italic=existing.italic if existing else False,
                    color=existing.color if existing else None,
                    size=existing.size if (existing and existing.size) else 10.5,
                )


INR = '"₹"#,##0.00'
INR0 = '"₹"#,##0'
PCT1 = '0.0%'
PCT2 = '0.00%'
NUM0 = '#,##0'
ROASFMT = '0.00"x"'

# ─────────────────────────────────────────────────────────────────────────
# Load existing workbook (append mode) or start fresh
# ─────────────────────────────────────────────────────────────────────────

def load_existing_daily_rows(path):
    """Return (shopify_rows, meta_rows) already present in an existing file, or (None, None)."""
    if not os.path.exists(path):
        return None, None
    try:
        wb = load_workbook(path, data_only=False)
    except Exception:
        return None, None
    shopify_rows, meta_rows = [], []
    if "Shopify Daily Data" in wb.sheetnames:
        ws = wb["Shopify Daily Data"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            shopify_rows.append(list(row))
    if "Meta Ads Daily Data" in wb.sheetnames:
        ws = wb["Meta Ads Daily Data"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            meta_rows.append(list(row))
    return shopify_rows, meta_rows


def merge_daily(existing_rows, new_rows):
    """Merge new_rows into existing_rows, skipping dates already present. Returns sorted-by-date list."""
    if not existing_rows:
        return list(new_rows)
    seen = {str(r[0]) for r in existing_rows}
    merged = list(existing_rows)
    for r in new_rows:
        if str(r[0]) not in seen:
            merged.append(r)
            seen.add(str(r[0]))
    merged.sort(key=lambda r: str(r[0]))
    return merged


existing_shopify, existing_meta = load_existing_daily_rows(OUT_PATH)
shopify_rows = merge_daily(existing_shopify, SHOPIFY_DAILY)
meta_rows = merge_daily(existing_meta, META_DAILY)

n_days = len(shopify_rows)          # total historical rows now on file
last_row_idx = n_days + 1           # 1-based sheet row of the most recent day (header=row1)
prev_start = max(2, last_row_idx - 7)
prev_end = last_row_idx - 1

# ─────────────────────────────────────────────────────────────────────────
# Build workbook
# ─────────────────────────────────────────────────────────────────────────

wb = Workbook()
wb.remove(wb.active)
for name in TAB_NAMES:
    wb.create_sheet(name)

# ============================= Shopify Daily Data ==========================
ws = wb["Shopify Daily Data"]
headers = ["Date", "Orders", "Gross Sales", "Discounts", "Returns", "Net Sales",
           "Shipping", "Taxes", "Total Sales", "AOV"]
ws.append(headers)
for r in shopify_rows:
    ws.append(r)
style_header(ws, 1, len(headers))
ws.freeze_panes = "A2"
money_cols = [3, 4, 5, 6, 7, 8, 9, 10]
for row in range(2, n_days + 2):
    for c in money_cols:
        ws.cell(row=row, column=c).number_format = INR
    ws.cell(row=row, column=1).number_format = "yyyy-mm-dd"
tbl = Table(displayName="ShopifyDaily", ref=f"A1:J{n_days+1}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tbl)
# conditional format: returns more negative than -5% of gross -> red flag on Returns col
ws.conditional_formatting.add(
    f"E2:E{n_days+1}",
    CellIsRule(operator="lessThan", formula=["-10000"], fill=PatternFill("solid", fgColor=C_RED))
)
autofit(ws)

# ============================= Meta Ads Daily Data ==========================
ws = wb["Meta Ads Daily Data"]
headers = ["Date", "Spend", "Impressions", "Reach", "Clicks", "CTR", "CPC", "CPM",
           "Purchases (est.)", "CPA", "ROAS"]
ws.append(headers)
for r in meta_rows:
    row = list(r)
    row[5] = row[5] / 100.0  # ctr stored as fraction for % format
    ws.append(row)
style_header(ws, 1, len(headers))
ws.freeze_panes = "A2"
for row in range(2, n_days + 2):
    ws.cell(row=row, column=1).number_format = "yyyy-mm-dd"
    ws.cell(row=row, column=2).number_format = INR
    ws.cell(row=row, column=3).number_format = NUM0
    ws.cell(row=row, column=4).number_format = NUM0
    ws.cell(row=row, column=5).number_format = NUM0
    ws.cell(row=row, column=6).number_format = PCT2
    ws.cell(row=row, column=7).number_format = INR
    ws.cell(row=row, column=8).number_format = INR
    ws.cell(row=row, column=9).number_format = NUM0
    ws.cell(row=row, column=10).number_format = INR
    ws.cell(row=row, column=11).number_format = ROASFMT
tbl = Table(displayName="MetaDaily", ref=f"A1:K{n_days+1}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tbl)
# ROAS conditional format: <1.5 red, 1.5-3 yellow, >3 green
rng = f"K2:K{n_days+1}"
ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["1.5"],
                               fill=PatternFill("solid", fgColor=C_RED)))
ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(K2>=1.5,K2<=3)"],
                               fill=PatternFill("solid", fgColor=C_YELLOW)))
ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["3"],
                               fill=PatternFill("solid", fgColor=C_GREEN)))
autofit(ws)

# ============================= Product Performance ==========================
ws = wb["Product Performance"]
style_title(ws, 1, 6, f"Product Performance — Yesterday ({YESTERDAY})")
headers = ["Product", "Gross Sales", "Net Sales", "Orders", "Avg Price", "Share of Day's Gross Sales"]
ws.append([None]*6)  # row2 spacer (title occupies row1)
ws.cell(row=2, column=1, value=None)
hdr_row = 3
for i, h in enumerate(headers, start=1):
    ws.cell(row=hdr_row, column=i, value=h)
style_header(ws, hdr_row, len(headers))
r = hdr_row + 1
prod_start = r
for p in PRODUCTS_YESTERDAY:
    title, gross, net, orders = p
    avg_price = round(gross / orders, 2) if orders else 0
    ws.cell(row=r, column=1, value=title)
    ws.cell(row=r, column=2, value=gross).number_format = INR
    ws.cell(row=r, column=3, value=net).number_format = INR
    ws.cell(row=r, column=4, value=orders).number_format = NUM0
    ws.cell(row=r, column=5, value=avg_price).number_format = INR
    ws.cell(row=r, column=6, value=f"=B{r}/SUM($B${prod_start}:$B${prod_start+len(PRODUCTS_YESTERDAY)-1})").number_format = PCT1
    r += 1
prod_end = r - 1
tbl = Table(displayName="ProductPerf", ref=f"A{hdr_row}:F{prod_end}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tbl)
ws.conditional_formatting.add(
    f"B{prod_start}:B{prod_end}",
    ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="63BE7B")
)

# Inventory section
inv_title_row = prod_end + 3
style_section(ws, inv_title_row, 6, "Inventory Watch — Lowest Stock Items (yesterday snapshot)")
inv_hdr_row = inv_title_row + 1
inv_headers = ["Product", "Starting Units", "Ending Units", "Units Sold", "Sell-Through", "Status"]
for i, h in enumerate(inv_headers, start=1):
    ws.cell(row=inv_hdr_row, column=i, value=h)
style_header(ws, inv_hdr_row, len(inv_headers))
r = inv_hdr_row + 1
inv_start = r
for item in INVENTORY_SNAPSHOT:
    title, start_u, end_u, sold, sell_thru = item
    ws.cell(row=r, column=1, value=title)
    ws.cell(row=r, column=2, value=start_u).number_format = NUM0
    ws.cell(row=r, column=3, value=end_u).number_format = NUM0
    ws.cell(row=r, column=4, value=sold).number_format = NUM0
    if sell_thru is not None:
        ws.cell(row=r, column=5, value=sell_thru).number_format = PCT2
    status = "OUT OF STOCK" if end_u == 0 else ("LOW STOCK" if end_u < 20 else "OK")
    ws.cell(row=r, column=6, value=status)
    r += 1
inv_end = r - 1
tbl2 = Table(displayName="InventoryWatch", ref=f"A{inv_hdr_row}:F{inv_end}")
tbl2.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
ws.add_table(tbl2)
ws.conditional_formatting.add(
    f"F{inv_start}:F{inv_end}",
    FormulaRule(formula=[f'$F{inv_start}="OUT OF STOCK"'], fill=PatternFill("solid", fgColor=C_RED))
)
ws.conditional_formatting.add(
    f"F{inv_start}:F{inv_end}",
    FormulaRule(formula=[f'$F{inv_start}="LOW STOCK"'], fill=PatternFill("solid", fgColor=C_YELLOW))
)
ws.freeze_panes = f"A{hdr_row+1}"
autofit(ws)

# ============================= Campaign Performance ==========================
ws = wb["Campaign Performance"]
style_title(ws, 1, 9, f"Campaign Performance — Yesterday vs Previous 7-Day Average")
hdr_row = 3
headers = ["Campaign", "Yesterday Spend", "7-Day Avg Spend/Day", "Yesterday ROAS",
           "7-Day Avg ROAS", "ROAS Change", "Yesterday CPA", "7-Day Avg CPA", "Verdict"]
for i, h in enumerate(headers, start=1):
    ws.cell(row=hdr_row, column=i, value=h)
style_header(ws, hdr_row, len(headers))

camp_7d_map = {c[0]: c for c in CAMPAIGNS_7DAY}
r = hdr_row + 1
camp_start = r
for c in CAMPAIGNS_YESTERDAY:
    name, spend, impr, clicks, ctr, cpc, roas, cpa = c
    seven = camp_7d_map.get(name)
    seven_spend_day = round(seven[1] / 7, 2) if seven else None
    seven_roas = seven[6] if seven else None
    seven_cpa = seven[7] if seven else None
    ws.cell(row=r, column=1, value=name)
    ws.cell(row=r, column=2, value=spend).number_format = INR
    if seven_spend_day is not None:
        ws.cell(row=r, column=3, value=seven_spend_day).number_format = INR
    ws.cell(row=r, column=4, value=roas).number_format = ROASFMT
    if seven_roas is not None:
        ws.cell(row=r, column=5, value=seven_roas).number_format = ROASFMT
        ws.cell(row=r, column=6, value=f"=(D{r}-E{r})/E{r}").number_format = PCT1
    ws.cell(row=r, column=7, value=cpa).number_format = INR
    if seven_cpa is not None:
        ws.cell(row=r, column=8, value=seven_cpa).number_format = INR
    verdict = "SCALE" if roas >= 5 else ("REVIEW" if roas < 3.5 else "MAINTAIN")
    ws.cell(row=r, column=9, value=verdict)
    r += 1
camp_end = r - 1
tbl = Table(displayName="CampaignPerf", ref=f"A{hdr_row}:I{camp_end}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tbl)
ws.conditional_formatting.add(f"D{camp_start}:E{camp_end}", CellIsRule(operator="lessThan", formula=["3"], fill=PatternFill("solid", fgColor=C_RED)))
ws.conditional_formatting.add(f"D{camp_start}:E{camp_end}", FormulaRule(formula=[f"AND(D{camp_start}>=3,D{camp_start}<5)"], fill=PatternFill("solid", fgColor=C_YELLOW)))
ws.conditional_formatting.add(f"D{camp_start}:E{camp_end}", CellIsRule(operator="greaterThanOrEqual", formula=["5"], fill=PatternFill("solid", fgColor=C_GREEN)))
ws.conditional_formatting.add(f"I{camp_start}:I{camp_end}", FormulaRule(formula=[f'$I{camp_start}="REVIEW"'], fill=PatternFill("solid", fgColor=C_RED)))
ws.conditional_formatting.add(f"I{camp_start}:I{camp_end}", FormulaRule(formula=[f'$I{camp_start}="SCALE"'], fill=PatternFill("solid", fgColor=C_GREEN)))

# Anomaly note box
note_row = camp_end + 2
style_section(ws, note_row, 9, "Delivery Anomaly Signal (Meta)")
ws.merge_cells(start_row=note_row+1, start_column=1, end_row=note_row+2, end_column=9)
cell = ws.cell(row=note_row+1, column=1, value=ANOMALY_NOTE)
cell.alignment = Alignment(wrap_text=True, vertical="top")
cell.font = Font(name=FONT_NAME, italic=True, color=C_YELLOW_TXT)
ws.freeze_panes = f"A{hdr_row+1}"
autofit(ws)

# ============================= Recommendations & Notes ==========================
ws = wb["Recommendations & Notes"]
style_title(ws, 1, 2, f"Recommendations & Notes — {YESTERDAY}")

r = 3
style_section(ws, r, 2, "KEY WINS")
r += 1
wins = [
    "Retargeting- new remains the strongest campaign: 5.52x ROAS yesterday (7-day avg 5.98x) at the lowest CPA (₹416) — clear candidate to scale budget.",
    "\"New Campaign- for recovery\" ROAS jumped from a 2.46x 7-day average to 3.90x yesterday — recent changes are working, worth investigating what changed.",
    "Net sales up +7.5% vs the 7-day average (₹169,461 vs ₹157,639) while returns (₹2,199) and discounts (₹915) were both below the 7-day average — healthier margin day.",
    "Orders (69) came in +9.5% above the 7-day average (63.0).",
]
for w in wins:
    ws.cell(row=r, column=1, value="✅")
    ws.cell(row=r, column=2, value=w).alignment = Alignment(wrap_text=True)
    r += 1

r += 1
style_section(ws, r, 2, "KEY ISSUES")
r += 1
issues = [
    "\"Black DHURANDHAR\" is the weakest campaign both yesterday (3.37x ROAS, highest CPA ₹890) and on the 7-day average (3.41x) — consistently underperforming its peers.",
    "Ad set \"Black and multiple\" is flagged by Meta for a narrow audience (~22.4k-26.4k people) while actively spending — risk of rising CPMs/frequency.",
    "AOV (₹2,488) is -3.6% below the 7-day average (₹2,581).",
    "2 products are fully OUT OF STOCK (\"Zariya\" Stone Hoops; \"Utsav\" Mal Cotton Anarkali Suit Set) — any ads or listings pointing to these are wasting clicks.",
]
for i in issues:
    ws.cell(row=r, column=1, value="\U0001F6A8")
    ws.cell(row=r, column=2, value=i).alignment = Alignment(wrap_text=True)
    r += 1

r += 1
style_section(ws, r, 2, "RECOMMENDED ACTIONS FOR TODAY")
r += 1
actions = [
    "Scale \"Retargeting- new\" budget by 20-30% — best ROAS and lowest CPA of all campaigns, both yesterday and on the 7-day trend.",
    "Review \"Black DHURANDHAR\": test new creative or tighten targeting before adding more budget — it is the only campaign below a 3.5x ROAS on both windows.",
    "Expand or A/B test the audience on ad set \"Black and multiple\" per Meta's narrow-audience warning to avoid frequency burnout.",
    "Restock or unpublish/hide \"Zariya\" Stone Hoops and the \"Utsav\" Anarkali Suit Set (0 units) so ad spend and storefront traffic aren't sent to sold-out listings.",
    "Keep an eye on \"Ira Rayon Co-ord Set\" and \"The Essential\" White Cotton Shirt & Pant Set — together they drove ₹79,265 (62%) of yesterday's product-level gross sales; ensure stock covers continued demand.",
]
for i, a in enumerate(actions, start=1):
    ws.cell(row=r, column=1, value=f"{i}.")
    ws.cell(row=r, column=2, value=a).alignment = Alignment(wrap_text=True)
    r += 1

r += 1
style_section(ws, r, 2, "UNAVAILABLE METRICS / DATA NOTES")
r += 1
for m in UNAVAILABLE_METRICS:
    ws.cell(row=r, column=1, value="ℹ️")
    ws.cell(row=r, column=2, value=m).alignment = Alignment(wrap_text=True)
    ws.row_dimensions[r].height = 30
    r += 1

ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 110
ws.freeze_panes = "A4"

print("Sheet content built. Now building Dashboard...")
print(f"n_days={n_days}, last_row_idx={last_row_idx}, prev_start={prev_start}, prev_end={prev_end}")

wb.save(OUT_PATH)
print(f"Saved intermediate to {OUT_PATH}")
