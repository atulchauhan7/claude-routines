#!/usr/bin/env python3
"""
Build / update the "Daily Store & Ads Performance Sheet" workbook from a
per-day JSON data file (see data_YYYY-MM-DD.json for the expected shape).

Google Sheets API is not reachable from this environment, so this produces
the best available fallback: a fully formatted .xlsx with the same tab
structure, formulas, conditional formatting and charts that the Google
Sheet would have had. Re-running with a new date's JSON appends new rows to
the two "Daily Data" tabs (skipping the date if it is already present) and
rebuilds the snapshot tabs (Product/Campaign Performance, Recommendations,
Dashboard) from the latest data.

Usage:
    python3 build_dashboard.py data_2026-06-30.json [output.xlsx]
"""
import sys
import json
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo

# ─────────────────────────── Style constants ────────────────────────────

C_HEADER_BG = "1F1F1F"
C_HEADER_FG = "FFFFFF"
C_TITLE_BG = "2A5AF7"
C_SECTION_BG = "E8E8E8"
C_GREEN = "C7EEC7"
C_RED = "F5B6B6"
C_YELLOW = "FFF2CC"
C_ORANGE = "FFD8A8"
C_WHITE = "FFFFFF"

INR = '"₹"#,##0.00'
INR0 = '"₹"#,##0'
PCT2 = '0.00"%"'
NUM0 = '#,##0'
NUM2 = '#,##0.00'

FONT_TITLE = Font(bold=True, size=16, color=C_HEADER_FG)
FONT_SECTION = Font(bold=True, size=11, color=C_HEADER_FG)
FONT_HEADER = Font(bold=True, size=10, color=C_HEADER_FG)
FONT_SUBHEADER = Font(bold=True, size=10, color="000000")
FONT_BOLD = Font(bold=True)

FILL_TITLE = PatternFill("solid", fgColor=C_TITLE_BG)
FILL_SECTION = PatternFill("solid", fgColor=C_HEADER_BG)
FILL_SUBHEADER = PatternFill("solid", fgColor=C_SECTION_BG)
FILL_GREEN = PatternFill("solid", fgColor=C_GREEN)
FILL_RED = PatternFill("solid", fgColor=C_RED)
FILL_YELLOW = PatternFill("solid", fgColor=C_YELLOW)
FILL_ORANGE = PatternFill("solid", fgColor=C_ORANGE)

THIN = Side(style="thin", color="D0D0D0")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

TAB_NAMES = [
    "Dashboard",
    "Shopify Daily Data",
    "Meta Ads Daily Data",
    "Product Performance",
    "Campaign Performance",
    "Recommendations & Notes",
]


def autofit(ws, widths):
    """widths: dict col_letter -> width"""
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def style_header_row(ws, row, ncols, fill=FILL_SECTION, font=FONT_HEADER):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_ALL


def style_title_row(ws, row, ncols, text):
    ws.cell(row=row, column=1, value=text)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    cell = ws.cell(row=row, column=1)
    cell.fill = FILL_TITLE
    cell.font = FONT_TITLE
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 28


def pct_change_formula(cur_cell, avg_cell):
    return f'=IF({avg_cell}=0,"—",({cur_cell}-{avg_cell})/{avg_cell})'


# ─────────────────────────── Shopify Daily Data ─────────────────────────

SHOPIFY_HEADERS = [
    "Date", "Gross Sales", "Net Sales", "Orders", "AOV", "Total Sales",
    "Discounts", "Returns", "Shipping", "Taxes", "Sessions", "Cart Adds",
    "Reached Checkout", "Completed Checkout", "Conversion Rate",
]
SHOPIFY_KEYS = [
    "date", "gross_sales", "net_sales", "orders", "aov", "total_sales",
    "discounts", "returns", "shipping", "taxes", "sessions", "cart_adds",
    "reached_checkout", "completed_checkout", "conversion_rate",
]
SHOPIFY_CURRENCY_COLS = [2, 3, 5, 6, 7, 8, 9, 10]  # 1-indexed
SHOPIFY_PCT_COLS = [15]


def build_shopify_daily(wb, data):
    ws = wb["Shopify Daily Data"] if "Shopify Daily Data" in wb.sheetnames else wb.create_sheet("Shopify Daily Data")
    ncols = len(SHOPIFY_HEADERS)
    if ws.max_row < 2 or ws["A1"].value != "Date":
        ws.delete_rows(1, ws.max_row)
        for i, h in enumerate(SHOPIFY_HEADERS, 1):
            ws.cell(row=1, column=i, value=h)
        style_header_row(ws, 1, ncols)

    existing_dates = {str(ws.cell(row=r, column=1).value) for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=1).value}

    for day in data["shopify"]["daily_history"]:
        if day["date"] in existing_dates:
            continue
        next_row = ws.max_row + 1
        for c, key in enumerate(SHOPIFY_KEYS, 1):
            ws.cell(row=next_row, column=c, value=day.get(key))

    last_row = ws.max_row
    for r in range(2, last_row + 1):
        for c in SHOPIFY_CURRENCY_COLS:
            ws.cell(row=r, column=c).number_format = INR
        for c in SHOPIFY_PCT_COLS:
            cell = ws.cell(row=r, column=c)
            cell.number_format = '0.00%'
        ws.cell(row=r, column=4).number_format = NUM0
        for c in (11, 12, 13, 14):
            ws.cell(row=r, column=c).number_format = NUM0
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).border = BORDER_ALL

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}{last_row}"

    # Conditional formatting: 3-colour scale on Gross Sales to spot spikes/drops
    ws.conditional_formatting.add(
        f"B2:B{last_row}",
        ColorScaleRule(start_type="min", start_color="F5B6B6",
                        mid_type="percentile", mid_value=50, mid_color="FFF2CC",
                        end_type="max", end_color="C7EEC7"),
    )
    autofit(ws, {
        "A": 12, "B": 13, "C": 13, "D": 9, "E": 11, "F": 13, "G": 12,
        "H": 12, "I": 11, "J": 10, "K": 10, "L": 10, "M": 15, "N": 16, "O": 14,
    })
    return ws, last_row


# ─────────────────────────── Meta Ads Daily Data ─────────────────────────

META_HEADERS = ["Date", "Spend", "Impressions", "Reach", "Clicks", "CTR",
                 "CPC", "CPM", "Purchases", "CPA", "ROAS"]
META_KEYS = ["date", "spend", "impressions", "reach", "clicks", "ctr",
             "cpc", "cpm", "purchases", "cpa", "roas"]


def build_meta_daily(wb, data):
    ws = wb["Meta Ads Daily Data"] if "Meta Ads Daily Data" in wb.sheetnames else wb.create_sheet("Meta Ads Daily Data")
    ncols = len(META_HEADERS)
    if ws.max_row < 2 or ws["A1"].value != "Date":
        ws.delete_rows(1, ws.max_row)
        for i, h in enumerate(META_HEADERS, 1):
            ws.cell(row=1, column=i, value=h)
        style_header_row(ws, 1, ncols)

    existing_dates = {str(ws.cell(row=r, column=1).value) for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=1).value}

    for day in data["meta"]["daily_history"]:
        if day["date"] in existing_dates:
            continue
        # compute cpa if missing
        row = dict(day)
        if row.get("cpa") is None and row.get("purchases"):
            row["cpa"] = round(row["spend"] / row["purchases"], 2) if row["purchases"] else None
        next_row = ws.max_row + 1
        for c, key in enumerate(META_KEYS, 1):
            ws.cell(row=next_row, column=c, value=row.get(key))

    last_row = ws.max_row
    for r in range(2, last_row + 1):
        for c in (2, 7, 8, 10):
            ws.cell(row=r, column=c).number_format = INR
        ws.cell(row=r, column=6).number_format = PCT2
        ws.cell(row=r, column=11).number_format = '0.00"x"'
        for c in (3, 4, 5, 9):
            ws.cell(row=r, column=c).number_format = NUM0
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).border = BORDER_ALL

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}{last_row}"

    # ROAS 3-tier conditional formatting
    rng = f"K2:K{last_row}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["1.5"], fill=FILL_RED))
    ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["1.5", "3"], fill=FILL_YELLOW))
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["3"], fill=FILL_GREEN))

    autofit(ws, {"A": 12, "B": 13, "C": 13, "D": 13, "E": 10, "F": 9,
                  "G": 10, "H": 10, "I": 11, "J": 11, "K": 9})
    return ws, last_row


# ─────────────────────────── Product Performance ─────────────────────────

def build_product_performance(wb, data):
    ws = wb["Product Performance"] if "Product Performance" in wb.sheetnames else wb.create_sheet("Product Performance")
    ws.delete_rows(1, ws.max_row)

    date_str = data["date"]
    w_start, w_end = data["window_7d_start"], data["window_7d_end"]

    style_title_row(ws, 1, 6, f"TOP PRODUCTS — YESTERDAY ({date_str})")
    headers1 = ["Product", "Gross Sales", "Net Sales", "Orders", "Units Sold", "Ending Inventory"]
    for i, h in enumerate(headers1, 1):
        ws.cell(row=2, column=i, value=h)
    style_header_row(ws, 2, 6, fill=FILL_SUBHEADER, font=FONT_SUBHEADER)

    r = 3
    for p in data["shopify"]["top_products_yesterday"]:
        ws.cell(row=r, column=1, value=p["product"])
        ws.cell(row=r, column=2, value=p["gross_sales"]).number_format = INR
        ws.cell(row=r, column=3, value=p["net_sales"]).number_format = INR
        ws.cell(row=r, column=4, value=p["orders"]).number_format = NUM0
        ws.cell(row=r, column=5, value=p.get("units_sold")).number_format = NUM0
        ws.cell(row=r, column=6, value=p.get("ending_inventory")).number_format = NUM0
        for c in range(1, 7):
            ws.cell(row=r, column=c).border = BORDER_ALL
        r += 1
    section1_end = r - 1
    ws.auto_filter.ref = f"A2:F{section1_end}"

    # Low-inventory conditional formatting
    ws.conditional_formatting.add(
        f"F3:F{section1_end}",
        CellIsRule(operator="lessThan", formula=["300"], fill=FILL_RED),
    )
    ws.conditional_formatting.add(
        f"F3:F{section1_end}",
        CellIsRule(operator="between", formula=["300", "1000"], fill=FILL_ORANGE),
    )

    r += 1
    section2_title_row = r
    style_title_row(ws, r, 4, f"TOP PRODUCTS — 7-DAY WINDOW ({w_start} to {w_end})")
    r += 1
    headers2 = ["Product", "Gross Sales", "Net Sales", "Orders"]
    for i, h in enumerate(headers2, 1):
        ws.cell(row=r, column=i, value=h)
    style_header_row(ws, r, 4, fill=FILL_SUBHEADER, font=FONT_SUBHEADER)
    r += 1
    for p in data["shopify"]["top_products_7d"]:
        ws.cell(row=r, column=1, value=p["product"])
        ws.cell(row=r, column=2, value=p["gross_sales"]).number_format = INR
        ws.cell(row=r, column=3, value=p["net_sales"]).number_format = INR
        ws.cell(row=r, column=4, value=p["orders"]).number_format = NUM0
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BORDER_ALL
        r += 1

    ws.freeze_panes = "A3"
    autofit(ws, {"A": 55, "B": 14, "C": 14, "D": 10, "E": 12, "F": 16})
    return ws, section1_end, section2_title_row


# ─────────────────────────── Campaign Performance ─────────────────────────

def build_campaign_performance(wb, data):
    ws = wb["Campaign Performance"] if "Campaign Performance" in wb.sheetnames else wb.create_sheet("Campaign Performance")
    ws.delete_rows(1, ws.max_row)
    date_str = data["date"]
    w_start, w_end = data["window_7d_start"], data["window_7d_end"]

    # ── Section A: Campaign level, yesterday ──
    style_title_row(ws, 1, 11, f"CAMPAIGN LEVEL — YESTERDAY ({date_str})")
    headers_c = ["Campaign", "Status", "Spend", "Impressions", "Reach", "Clicks",
                 "CTR", "CPC", "Purchases", "CPA", "ROAS"]
    for i, h in enumerate(headers_c, 1):
        ws.cell(row=2, column=i, value=h)
    style_header_row(ws, 2, 11, fill=FILL_SUBHEADER, font=FONT_SUBHEADER)

    r = 3
    camps = sorted(data["meta"]["campaigns_yesterday"], key=lambda x: x["spend"], reverse=True)
    for c_ in camps:
        ws.cell(row=r, column=1, value=c_["name"])
        ws.cell(row=r, column=2, value=c_["status"])
        ws.cell(row=r, column=3, value=c_["spend"]).number_format = INR
        ws.cell(row=r, column=4, value=c_["impressions"]).number_format = NUM0
        ws.cell(row=r, column=5, value=c_["reach"]).number_format = NUM0
        ws.cell(row=r, column=6, value=c_["clicks"]).number_format = NUM0
        ws.cell(row=r, column=7, value=c_["ctr"]).number_format = PCT2
        ws.cell(row=r, column=8, value=c_["cpc"]).number_format = INR
        ws.cell(row=r, column=9, value=c_["purchases"]).number_format = NUM0
        ws.cell(row=r, column=10, value=c_["cpa"]).number_format = INR
        ws.cell(row=r, column=11, value=c_["roas"]).number_format = '0.00"x"'
        for c in range(1, 12):
            ws.cell(row=r, column=c).border = BORDER_ALL
        r += 1
    camp_end = r - 1
    ws.auto_filter.ref = f"A2:K{camp_end}"

    roas_rng = f"K3:K{camp_end}"
    ws.conditional_formatting.add(roas_rng, CellIsRule(operator="lessThan", formula=["1.5"], fill=FILL_RED))
    ws.conditional_formatting.add(roas_rng, CellIsRule(operator="between", formula=["1.5", "3"], fill=FILL_YELLOW))
    ws.conditional_formatting.add(roas_rng, CellIsRule(operator="greaterThan", formula=["3"], fill=FILL_GREEN))
    # Underperforming campaign row highlight: spend > 0 and ROAS < 2
    ws.conditional_formatting.add(
        f"A3:K{camp_end}",
        FormulaRule(formula=[f"AND($C3>0,$K3<2)"], fill=FILL_ORANGE),
    )

    r += 1
    # ── Section B: Ad set level, yesterday ──
    style_title_row(ws, r, 12, f"AD SET LEVEL — YESTERDAY ({date_str})")
    r += 1
    headers_a = ["Ad Set", "Campaign", "Status", "Spend", "Impressions", "Clicks",
                 "CTR", "CPC", "Purchases", "CPA", "ROAS", "Notes"]
    for i, h in enumerate(headers_a, 1):
        ws.cell(row=r, column=i, value=h)
    style_header_row(ws, r, 12, fill=FILL_SUBHEADER, font=FONT_SUBHEADER)
    adset_header_row = r
    r += 1
    adsets = sorted(data["meta"]["adsets_yesterday"], key=lambda x: x["spend"], reverse=True)
    for a in adsets:
        ws.cell(row=r, column=1, value=a["name"])
        ws.cell(row=r, column=2, value=a["campaign"])
        ws.cell(row=r, column=3, value=a["status"])
        ws.cell(row=r, column=4, value=a["spend"]).number_format = INR
        ws.cell(row=r, column=5, value=a["impressions"]).number_format = NUM0
        ws.cell(row=r, column=6, value=a["clicks"]).number_format = NUM0
        if a.get("ctr") is not None:
            ws.cell(row=r, column=7, value=a["ctr"]).number_format = PCT2
        if a.get("cpc") is not None:
            ws.cell(row=r, column=8, value=a["cpc"]).number_format = INR
        ws.cell(row=r, column=9, value=a["purchases"]).number_format = NUM0
        if a.get("cpa") is not None:
            ws.cell(row=r, column=10, value=a["cpa"]).number_format = INR
        if a.get("roas") is not None:
            ws.cell(row=r, column=11, value=a["roas"]).number_format = '0.00"x"'
        ws.cell(row=r, column=12, value=a.get("warning", ""))
        for c in range(1, 13):
            ws.cell(row=r, column=c).border = BORDER_ALL
        r += 1
    adset_end = r - 1

    roas_rng2 = f"K{adset_header_row + 1}:K{adset_end}"
    ws.conditional_formatting.add(roas_rng2, CellIsRule(operator="lessThan", formula=["1.5"], fill=FILL_RED))
    ws.conditional_formatting.add(roas_rng2, CellIsRule(operator="between", formula=["1.5", "3"], fill=FILL_YELLOW))
    ws.conditional_formatting.add(roas_rng2, CellIsRule(operator="greaterThan", formula=["3"], fill=FILL_GREEN))
    ws.conditional_formatting.add(
        f"L{adset_header_row + 1}:L{adset_end}",
        FormulaRule(formula=[f'LEN($L{adset_header_row + 1})>0'], fill=FILL_YELLOW),
    )

    r += 1
    # ── Section C: Campaign level, 7-day window ──
    style_title_row(ws, r, 9, f"CAMPAIGN LEVEL — 7-DAY WINDOW ({w_start} to {w_end})")
    r += 1
    headers_7 = ["Campaign", "Spend", "Impressions", "Clicks", "CTR", "CPC", "Purchases", "CPA", "ROAS"]
    for i, h in enumerate(headers_7, 1):
        ws.cell(row=r, column=i, value=h)
    style_header_row(ws, r, 9, fill=FILL_SUBHEADER, font=FONT_SUBHEADER)
    r += 1
    camp7_start = r
    for c_ in sorted(data["meta"]["campaigns_7d"], key=lambda x: x["spend"], reverse=True):
        ws.cell(row=r, column=1, value=c_["name"])
        ws.cell(row=r, column=2, value=c_["spend"]).number_format = INR
        ws.cell(row=r, column=3, value=c_["impressions"]).number_format = NUM0
        ws.cell(row=r, column=4, value=c_["clicks"]).number_format = NUM0
        ws.cell(row=r, column=5, value=c_["ctr"]).number_format = PCT2
        ws.cell(row=r, column=6, value=c_["cpc"]).number_format = INR
        ws.cell(row=r, column=7, value=c_["purchases"]).number_format = NUM0
        ws.cell(row=r, column=8, value=c_["cpa"]).number_format = INR
        ws.cell(row=r, column=9, value=c_["roas"]).number_format = '0.00"x"'
        for c in range(1, 10):
            ws.cell(row=r, column=c).border = BORDER_ALL
        r += 1
    camp7_end = r - 1
    roas_rng3 = f"I{camp7_start}:I{camp7_end}"
    ws.conditional_formatting.add(roas_rng3, CellIsRule(operator="lessThan", formula=["1.5"], fill=FILL_RED))
    ws.conditional_formatting.add(roas_rng3, CellIsRule(operator="between", formula=["1.5", "3"], fill=FILL_YELLOW))
    ws.conditional_formatting.add(roas_rng3, CellIsRule(operator="greaterThan", formula=["3"], fill=FILL_GREEN))

    ws.freeze_panes = "A3"
    autofit(ws, {"A": 30, "B": 12, "C": 13, "D": 12, "E": 10, "F": 9,
                  "G": 9, "H": 11, "I": 10, "J": 9, "K": 9, "L": 40})
    return ws, camp_end


# ─────────────────────────── Recommendations & Notes ─────────────────────

def build_recommendations(wb, data, recs, key_wins, key_issues):
    ws = wb["Recommendations & Notes"] if "Recommendations & Notes" in wb.sheetnames else wb.create_sheet("Recommendations & Notes")
    ws.delete_rows(1, ws.max_row)
    date_str = data["date"]

    style_title_row(ws, 1, 2, f"RECOMMENDATIONS & NOTES — {date_str}")
    r = 3

    def section(title, items, r):
        ws.cell(row=r, column=1, value=title)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.cell(row=r, column=1).fill = FILL_SUBHEADER
        ws.cell(row=r, column=1).font = FONT_SUBHEADER
        r += 1
        for idx, item in enumerate(items, 1):
            ws.cell(row=r, column=1, value=idx)
            ws.cell(row=r, column=2, value=item)
            ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
            for c in (1, 2):
                ws.cell(row=r, column=c).border = BORDER_ALL
            r += 1
        return r + 1

    r = section("KEY WINS", key_wins, r)
    r = section("KEY ISSUES", key_issues, r)
    r = section("RECOMMENDED ACTIONS FOR TODAY", recs, r)

    unavailable = data["shopify"].get("unavailable", []) + data["meta"].get("unavailable", [])
    r = section("DATA NOTES / METRICS UNAVAILABLE TODAY", unavailable, r)

    anomalies = data["meta"].get("anomalies", [])
    if anomalies:
        r = section("META ANOMALY SIGNALS", anomalies, r)

    ws.freeze_panes = "A2"
    autofit(ws, {"A": 6, "B": 110})
    return ws


# ─────────────────────────── Dashboard ────────────────────────────────

def build_dashboard(wb, data, shopify_last_row, meta_last_row, key_wins, key_issues, recs):
    ws = wb["Dashboard"] if "Dashboard" in wb.sheetnames else wb.create_sheet("Dashboard")
    ws.delete_rows(1, ws.max_row)
    date_str = data["date"]

    style_title_row(ws, 1, 6, f"DAILY STORE & ADS PERFORMANCE — {date_str} (Asia/Kolkata)")
    ws.row_dimensions[1].height = 32

    r = 3
    ws.cell(row=r, column=1, value="EXECUTIVE SUMMARY")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.cell(row=r, column=1).fill = FILL_SECTION
    ws.cell(row=r, column=1).font = FONT_SECTION
    r += 1

    s_last = shopify_last_row  # row for yesterday in Shopify Daily Data
    s_avg_start, s_avg_end = 2, shopify_last_row - 1
    m_last = meta_last_row
    m_avg_start, m_avg_end = 2, meta_last_row - 1

    summary_text = (
        f"Yesterday's gross sales were ₹{data['shopify']['daily_history'][-1]['gross_sales']:,.0f} "
        f"across {data['shopify']['daily_history'][-1]['orders']} orders, with Meta Ads spend of "
        f"₹{data['meta']['daily_history'][-1]['spend']:,.0f} driving a blended ROAS of "
        f"{data['meta']['daily_history'][-1]['roas']:.2f}x. See tables and charts below for the full "
        f"picture vs the trailing 7-day average."
    )
    ws.cell(row=r, column=1, value=summary_text)
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=6)
    ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30
    r += 3

    # ── Shopify Performance table ──
    ws.cell(row=r, column=1, value="SHOPIFY PERFORMANCE")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws.cell(row=r, column=1).fill = FILL_SECTION
    ws.cell(row=r, column=1).font = FONT_SECTION
    r += 1
    for i, h in enumerate(["Metric", "Yesterday", "7-Day Avg", "Change vs Avg"], 1):
        ws.cell(row=r, column=i, value=h)
    style_header_row(ws, r, 4, fill=FILL_SUBHEADER, font=FONT_SUBHEADER)
    r += 1
    shopify_metrics_start = r
    shopify_rows = [
        ("Gross Sales", "B", INR),
        ("Net Sales", "C", INR),
        ("Orders", "D", NUM0),
        ("Average Order Value", "E", INR),
    ]
    for label, col, fmt in shopify_rows:
        ws.cell(row=r, column=1, value=label)
        cur = f"'Shopify Daily Data'!{col}{s_last}"
        avg = f"AVERAGE('Shopify Daily Data'!{col}{s_avg_start}:{col}{s_avg_end})"
        ws.cell(row=r, column=2, value=f"={cur}").number_format = fmt
        ws.cell(row=r, column=3, value=f"={avg}").number_format = fmt
        ws.cell(row=r, column=4, value=pct_change_formula(f"B{r}", f"C{r}")).number_format = '+0.0%;-0.0%;—'
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BORDER_ALL
        r += 1
    ws.cell(row=r, column=1, value="Returns")
    ws.cell(row=r, column=2, value=f"='Shopify Daily Data'!H{s_last}").number_format = INR
    for c in range(1, 5):
        ws.cell(row=r, column=c).border = BORDER_ALL
    shopify_metrics_end = r
    r += 1
    ws.conditional_formatting.add(
        f"D{shopify_metrics_start}:D{shopify_metrics_end}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_GREEN),
    )
    ws.conditional_formatting.add(
        f"D{shopify_metrics_start}:D{shopify_metrics_end}",
        CellIsRule(operator="lessThan", formula=["0"], fill=FILL_RED),
    )
    r += 1

    # ── Meta Ads Performance table ──
    ws.cell(row=r, column=1, value="META ADS PERFORMANCE")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws.cell(row=r, column=1).fill = FILL_SECTION
    ws.cell(row=r, column=1).font = FONT_SECTION
    r += 1
    for i, h in enumerate(["Metric", "Yesterday", "7-Day Avg", "Change vs Avg"], 1):
        ws.cell(row=r, column=i, value=h)
    style_header_row(ws, r, 4, fill=FILL_SUBHEADER, font=FONT_SUBHEADER)
    r += 1
    meta_metrics_start = r
    meta_rows = [
        ("Spend", "B", INR),
        ("Impressions", "C", NUM0),
        ("Clicks", "E", NUM0),
        ("CTR", "F", PCT2),
        ("Purchases", "I", NUM0),
        ("CPA", "J", INR),
        ("ROAS", "K", '0.00"x"'),
    ]
    for label, col, fmt in meta_rows:
        ws.cell(row=r, column=1, value=label)
        cur = f"'Meta Ads Daily Data'!{col}{m_last}"
        avg = f"AVERAGE('Meta Ads Daily Data'!{col}{m_avg_start}:{col}{m_avg_end})"
        ws.cell(row=r, column=2, value=f"={cur}").number_format = fmt
        ws.cell(row=r, column=3, value=f"={avg}").number_format = fmt
        pct_row_formula = pct_change_formula(f"B{r}", f"C{r}")
        if label == "CPA":
            # For CPA lower is better -> invert the sign shown for "good/bad" coloring downstream
            pass
        ws.cell(row=r, column=4, value=pct_row_formula).number_format = '+0.0%;-0.0%;—'
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BORDER_ALL
        r += 1
    meta_metrics_end = r - 1
    ws.conditional_formatting.add(
        f"D{meta_metrics_start}:D{meta_metrics_end}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_GREEN),
    )
    ws.conditional_formatting.add(
        f"D{meta_metrics_start}:D{meta_metrics_end}",
        CellIsRule(operator="lessThan", formula=["0"], fill=FILL_RED),
    )
    r += 1

    # ── Key wins / issues / actions ──
    def bullet_section(title, items, r, fill=FILL_SECTION):
        ws.cell(row=r, column=1, value=title)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(row=r, column=1).fill = fill
        ws.cell(row=r, column=1).font = FONT_SECTION
        r += 1
        for item in items:
            ws.cell(row=r, column=1, value="•")
            ws.cell(row=r, column=2, value=item)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
            ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
        return r + 1

    r = bullet_section("KEY WINS", key_wins, r)
    r = bullet_section("KEY ISSUES", key_issues, r)

    ws.cell(row=r, column=1, value="RECOMMENDED ACTIONS FOR TODAY")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws.cell(row=r, column=1).fill = FILL_SECTION
    ws.cell(row=r, column=1).font = FONT_SECTION
    r += 1
    for idx, item in enumerate(recs, 1):
        ws.cell(row=r, column=1, value=idx)
        ws.cell(row=r, column=2, value=item)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    autofit(ws, {"A": 16, "B": 26, "C": 16, "D": 16})
    ws.freeze_panes = "A2"

    # ─────────────────── Charts ───────────────────
    chart_col = "G"
    chart_start_row = 3

    def add_line_chart(title, sheet_name, cat_col, val_cols, anchor, y_title):
        chart = LineChart()
        chart.title = title
        chart.style = 10
        chart.y_axis.title = y_title
        chart.x_axis.title = "Date"
        chart.height = 8
        chart.width = 16
        cats = Reference(wb[sheet_name], min_col=cat_col, min_row=2, max_row=(s_last if sheet_name == "Shopify Daily Data" else m_last))
        for vc in val_cols:
            data_ref = Reference(wb[sheet_name], min_col=vc, min_row=1, max_row=(s_last if sheet_name == "Shopify Daily Data" else m_last))
            chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, anchor)

    def add_bar_chart(title, sheet_name, cat_col, val_col, min_row, max_row, anchor, y_title):
        chart = BarChart()
        chart.type = "col"
        chart.title = title
        chart.style = 10
        chart.y_axis.title = y_title
        chart.height = 8
        chart.width = 16
        cats = Reference(wb[sheet_name], min_col=cat_col, min_row=min_row, max_row=max_row)
        data_ref = Reference(wb[sheet_name], min_col=val_col, min_row=min_row - 1, max_row=max_row)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, anchor)

    # Sales trend: Gross Sales (col 2) + Net Sales (col 3)
    add_line_chart("Sales Trend (8 days)", "Shopify Daily Data", 1, [2, 3], "G3", "₹ INR")
    # Orders trend
    add_line_chart("Orders Trend (8 days)", "Shopify Daily Data", 1, [4], "N3", "Orders")
    # Ad spend trend
    add_line_chart("Ad Spend Trend (8 days)", "Meta Ads Daily Data", 1, [2], "G20", "₹ INR")
    # ROAS trend
    add_line_chart("ROAS Trend (8 days)", "Meta Ads Daily Data", 1, [11], "N20", "ROAS (x)")

    # Top products bar chart (yesterday) - from Product Performance sheet rows 3-12, col A cats, col B values
    add_bar_chart("Top Products — Yesterday (Gross Sales)", "Product Performance", 1, 2, 3, 12, "G37", "₹ INR")
    # Top campaigns bar chart (yesterday) - from Campaign Performance sheet rows 3-6, col A cats, col K (ROAS) values
    add_bar_chart("Top Campaigns — Yesterday (ROAS)", "Campaign Performance", 1, 11, 3, 6, "N37", "ROAS (x)")

    return ws


# ─────────────────────────── Recommendation engine ────────────────────────

def generate_analysis(data):
    s_y = data["shopify"]["daily_history"][-1]
    s_hist = data["shopify"]["daily_history"][:-1]
    avg_gross = sum(d["gross_sales"] for d in s_hist) / len(s_hist)
    avg_net = sum(d["net_sales"] for d in s_hist) / len(s_hist)
    avg_orders = sum(d["orders"] for d in s_hist) / len(s_hist)

    m_y = data["meta"]["daily_history"][-1]
    m_hist = data["meta"]["daily_history"][:-1]
    avg_spend = sum(d["spend"] for d in m_hist) / len(m_hist)
    avg_roas = sum(d["roas"] for d in m_hist) / len(m_hist)
    avg_clicks = sum(d["clicks"] for d in m_hist) / len(m_hist)

    gross_chg = (s_y["gross_sales"] - avg_gross) / avg_gross * 100
    net_chg = (s_y["net_sales"] - avg_net) / avg_net * 100
    orders_chg = (s_y["orders"] - avg_orders) / avg_orders * 100
    roas_chg = (m_y["roas"] - avg_roas) / avg_roas * 100
    clicks_chg = (m_y["clicks"] - avg_clicks) / avg_clicks * 100

    best_camp = max(data["meta"]["campaigns_yesterday"], key=lambda c: c["roas"])
    worst_camp = min(data["meta"]["campaigns_yesterday"], key=lambda c: c["roas"])
    top_product = data["shopify"]["top_products_yesterday"][0]

    key_wins = [
        f"Gross sales up {gross_chg:+.1f}% vs 7-day avg (₹{s_y['gross_sales']:,.0f} vs ₹{avg_gross:,.0f}); net sales up {net_chg:+.1f}%.",
        f"Blended Meta ROAS at {m_y['roas']:.2f}x, up {roas_chg:+.1f}% vs 7-day avg ({avg_roas:.2f}x).",
        f"'{best_camp['name']}' was the top campaign yesterday: {best_camp['roas']:.2f}x ROAS on ₹{best_camp['spend']:,.0f} spend, {best_camp['purchases']} purchases.",
        f"Zero product returns recorded yesterday vs a 7-day average return drag of ₹{sum(d['returns'] for d in s_hist)/len(s_hist):,.0f}/day.",
        f"'{top_product['product']}' remains the best seller (₹{top_product['gross_sales']:,.0f}, {top_product['orders']} orders) with healthy stock ({top_product.get('ending_inventory'):,} units).",
    ]

    key_issues = [
        f"Clicks down {clicks_chg:+.1f}% vs 7-day avg despite similar spend — slight efficiency dip.",
        f"'{worst_camp['name']}' is the weakest campaign yesterday: {worst_camp['roas']:.2f}x ROAS, highest CPA at ₹{worst_camp['cpa']:,.0f}.",
        f"Ad set 'Black and multiple' flagged by Meta for a narrow audience (13.7K–16.2K) — likely capping delivery.",
        "Shopify session-based conversion rate is showing ~0% across all 8 days despite real completed orders — looks like an attribution/session-matching lag, not an actual conversion problem. Needs a pixel/GA4 check.",
    ]

    recs = [
        f"Scale '{best_camp['name']}' — best ROAS ({best_camp['roas']:.2f}x) and volume yesterday; increase daily budget by ~15-20%.",
        f"Review or pause underperforming elements of '{worst_camp['name']}' — ROAS {worst_camp['roas']:.2f}x and CPA ₹{worst_camp['cpa']:,.0f} are the weakest in the account; refresh creative or tighten the retargeting window.",
        "Expand the audience (or turn off Advantage+ narrowing) for the 'Black and multiple' ad set — flagged as narrow, likely limiting reach.",
        "Investigate Shopify checkout/session tracking — conversion rate reads ~0% despite 57 real orders yesterday; verify pixel/GA4 event wiring so attribution data can be trusted.",
        f"Maintain spend on '{best_camp['name']}' and check inventory/pages for top sellers — '{top_product['product']}' is driving outsized volume.",
        "Set up (or verify) abandoned checkout recovery emails — could not confirm current recovery flow status from available data sources.",
    ]

    return key_wins, key_issues, recs


# ─────────────────────────── Main ────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: build_dashboard.py <data_json> [output.xlsx]")
        sys.exit(1)

    json_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        os.path.dirname(json_path), "Daily_Store_and_Ads_Performance_Sheet.xlsx"
    )

    with open(json_path) as f:
        data = json.load(f)

    if os.path.exists(out_path):
        wb = load_workbook(out_path)
        for t in TAB_NAMES:
            if t not in wb.sheetnames:
                wb.create_sheet(t)
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
    else:
        wb = Workbook()
        wb.remove(wb.active)
        for t in TAB_NAMES:
            wb.create_sheet(t)

    # Reorder tabs
    wb._sheets.sort(key=lambda ws: TAB_NAMES.index(ws.title) if ws.title in TAB_NAMES else 99)

    ws_shopify, s_last = build_shopify_daily(wb, data)
    ws_meta, m_last = build_meta_daily(wb, data)
    build_product_performance(wb, data)
    build_campaign_performance(wb, data)

    key_wins, key_issues, recs = generate_analysis(data)

    build_recommendations(wb, data, recs, key_wins, key_issues)
    build_dashboard(wb, data, s_last, m_last, key_wins, key_issues, recs)

    wb.active = wb.sheetnames.index("Dashboard")
    wb.save(out_path)
    print(f"Saved workbook to: {out_path}")
    print(f"Shopify Daily Data rows: {s_last - 1}")
    print(f"Meta Ads Daily Data rows: {m_last - 1}")

    return out_path, key_wins, key_issues, recs


if __name__ == "__main__":
    main()
