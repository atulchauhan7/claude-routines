#!/usr/bin/env python3
"""
Builds/updates 'Daily Store & Ads Performance Sheet.xlsx' — a business dashboard
workbook combining Shopify + Meta Ads daily performance data.

Google Sheets access was not available in this environment (no connector), so
this workbook is the fallback deliverable per the routine's instructions. It is
committed to the repo each day so historical rows can be appended (not
overwritten) on subsequent runs.

Usage: python3 build_dashboard.py <path_to_day_data.json>
"""

import sys
import json
import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.table import Table, TableStyleInfo

FILE_PATH = "Daily Store & Ads Performance Sheet.xlsx"

# ── Palette ──────────────────────────────────────────────────────────────────
C_HEADER_BG = "1F2937"      # dark slate
C_HEADER_FG = "FFFFFF"
C_ACCENT_BG = "2563EB"      # brand blue
C_ACCENT_FG = "FFFFFF"
C_SECTION_BG = "E5E7EB"     # light grey
C_GOOD_BG   = "C6EFCE"      # green tint
C_GOOD_FG   = "1E7A34"
C_BAD_BG    = "F8CBCB"      # red tint
C_BAD_FG    = "9C1C1C"
C_WARN_BG   = "FFEB9C"      # amber tint
C_WARN_FG   = "9C6500"

FONT_HEADER = Font(bold=True, color=C_HEADER_FG, size=11)
FONT_TITLE  = Font(bold=True, color=C_ACCENT_FG, size=14)
FONT_SECTION = Font(bold=True, color=C_HEADER_FG, size=11)
FONT_BOLD = Font(bold=True)
FILL_HEADER = PatternFill("solid", fgColor=C_HEADER_BG)
FILL_ACCENT = PatternFill("solid", fgColor=C_ACCENT_BG)
FILL_SECTION = PatternFill("solid", fgColor=C_SECTION_BG)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="D1D5DB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

INR = '"₹"#,##0.00'
INR0 = '"₹"#,##0'
PCT1 = '0.0%'
PCT2 = '0.00%'
NUM0 = '#,##0'


def style_header_row(ws, row, ncols, fill=FILL_HEADER, font=FONT_HEADER):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = CENTER
        cell.border = BORDER


def autofit(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def add_table(ws, name, ref):
    tbl = Table(displayName=name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False
    )
    ws.add_table(tbl)


def load_data(path):
    with open(path) as f:
        return json.load(f)


def main():
    data = load_data(sys.argv[1])
    date_str = data["date"]

    try:
        wb = load_workbook(FILE_PATH)
        is_new = False
    except FileNotFoundError:
        wb = Workbook()
        wb.remove(wb.active)
        is_new = True

    TAB_ORDER = [
        "Dashboard", "Shopify Daily Data", "Meta Ads Daily Data",
        "Product Performance", "Campaign Performance", "Recommendations & Notes",
    ]
    # These tabs are fully rebuilt every run (not appended) — recreate them
    # clean each time instead of blanking cells, which would otherwise force
    # the sheet dimensions to balloon and break row-append offsets.
    REBUILD_TABS = {"Dashboard", "Product Performance", "Campaign Performance", "Recommendations & Notes"}
    for t in REBUILD_TABS:
        if t in wb.sheetnames:
            del wb[t]
    for t in TAB_ORDER:
        if t not in wb.sheetnames:
            wb.create_sheet(t)

    build_shopify_daily(wb["Shopify Daily Data"], data)
    build_meta_daily(wb["Meta Ads Daily Data"], data)
    build_product_performance(wb["Product Performance"], data)
    build_campaign_performance(wb["Campaign Performance"], data)
    build_recommendations(wb["Recommendations & Notes"], data)
    build_dashboard(wb["Dashboard"], data)

    wb._sheets = [wb[t] for t in TAB_ORDER]
    wb.active = 0
    wb.save(FILE_PATH)
    print(f"Saved {FILE_PATH}")


# ─────────────────────────── Shopify Daily Data ─────────────────────────────

def build_shopify_daily(ws, data):
    headers = [
        "Date", "Gross Sales", "Net Sales", "Orders", "AOV", "Discounts",
        "Returns/Refunds", "Cancellations", "Sessions", "Conversion Rate",
        "Gross Sales Δ% vs 7d Avg", "Orders Δ% vs 7d Avg",
    ]
    is_new_sheet = ws.max_row == 1 and ws.max_column == 1 and ws["A1"].value is None
    if is_new_sheet:
        ws.append(headers)
        for row in data["shopify_history"]:
            ws.append([
                row["date"], row["gross_sales"], row["net_sales"], row["orders"],
                row["aov"], row["discounts"], row["returns"], row.get("cancellations", 0),
                row.get("sessions", ""), row.get("conversion_rate", ""), None, None,
            ])

    # Append yesterday's row if not already present (avoid duplicates)
    existing_dates = {ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)}
    s = data["shopify_today"]
    if s["date"] not in existing_dates:
        ws.append([
            s["date"], s["gross_sales"], s["net_sales"], s["orders"], s["aov"],
            s["discounts"], s["returns"], s.get("cancellations", 0),
            s.get("sessions", ""), s.get("conversion_rate", ""),
            round(s["gross_chg_pct"] / 100, 4) if s.get("gross_chg_pct") is not None else None,
            round(s["orders_chg_pct"] / 100, 4) if s.get("orders_chg_pct") is not None else None,
        ])

    last_row = ws.max_row
    ncols = len(headers)
    style_header_row(ws, 1, ncols)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}{last_row}"

    for r in range(2, last_row + 1):
        for c in (2, 3, 5, 6, 7):
            ws.cell(row=r, column=c).number_format = INR
        ws.cell(row=r, column=4).number_format = NUM0
        ws.cell(row=r, column=8).number_format = NUM0
        ws.cell(row=r, column=9).number_format = NUM0
        ws.cell(row=r, column=10).number_format = PCT1
        ws.cell(row=r, column=11).number_format = PCT1
        ws.cell(row=r, column=12).number_format = PCT1
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).border = BORDER

    # Conditional formatting: green/red for % change columns (K, L)
    rng_k = f"K2:L{last_row}"
    ws.conditional_formatting.add(
        rng_k, CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=C_BAD_BG))
    )
    ws.conditional_formatting.add(
        rng_k, CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=C_GOOD_BG))
    )
    autofit(ws, [12, 13, 13, 9, 11, 11, 14, 13, 10, 15, 18, 17])


# ─────────────────────────── Meta Ads Daily Data ────────────────────────────

def build_meta_daily(ws, data):
    headers = [
        "Date", "Total Spend", "Impressions", "Reach", "Clicks", "CTR", "CPC",
        "CPM", "Purchases", "CPA", "ROAS", "Spend Δ% vs 7d Avg", "ROAS Δ% vs 7d Avg",
    ]
    is_new_sheet = ws.max_row == 1 and ws.max_column == 1 and ws["A1"].value is None
    if is_new_sheet:
        ws.append(headers)
        for row in data["meta_history"]:
            ws.append([
                row["date"], row["spend"], row["impressions"], row["reach"],
                row["clicks"], round(row["ctr"] / 100, 4), row["cpc"], row["cpm"],
                row.get("purchases", ""), row.get("cpa", ""), row["roas"], None, None,
            ])

    existing_dates = {ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)}
    m = data["meta_today"]
    if m["date"] not in existing_dates:
        ws.append([
            m["date"], m["spend"], m["impressions"], m["reach"], m["clicks"],
            round(m["ctr"] / 100, 4), m["cpc"], m["cpm"], m.get("purchases", ""),
            m.get("cpa", ""), m["roas"],
            round(m["spend_chg_pct"] / 100, 4) if m.get("spend_chg_pct") is not None else None,
            round(m["roas_chg_pct"] / 100, 4) if m.get("roas_chg_pct") is not None else None,
        ])

    last_row = ws.max_row
    ncols = len(headers)
    style_header_row(ws, 1, ncols)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}{last_row}"

    for r in range(2, last_row + 1):
        for c in (2, 7, 8, 10):
            ws.cell(row=r, column=c).number_format = INR
        ws.cell(row=r, column=3).number_format = NUM0
        ws.cell(row=r, column=4).number_format = NUM0
        ws.cell(row=r, column=5).number_format = NUM0
        ws.cell(row=r, column=6).number_format = PCT2
        ws.cell(row=r, column=9).number_format = NUM0
        ws.cell(row=r, column=11).number_format = '0.00"x"'
        ws.cell(row=r, column=12).number_format = PCT1
        ws.cell(row=r, column=13).number_format = PCT1
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).border = BORDER

    # ROAS conditional formatting: <1.5 red, 1.5-3 amber, >3 green
    rng = f"K2:K{last_row}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["1.5"], fill=PatternFill("solid", fgColor=C_BAD_BG)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["1.5", "3"], fill=PatternFill("solid", fgColor=C_WARN_BG)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["3"], fill=PatternFill("solid", fgColor=C_GOOD_BG)))
    rng_chg = f"L2:M{last_row}"
    ws.conditional_formatting.add(rng_chg, CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=C_BAD_BG)))
    ws.conditional_formatting.add(rng_chg, CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=C_GOOD_BG)))
    autofit(ws, [12, 13, 13, 11, 9, 9, 10, 10, 11, 10, 8, 17, 16])


# ─────────────────────────── Product Performance ────────────────────────────

def build_product_performance(ws, data):
    date_str = data["date"]
    ws["A1"] = f"PRODUCT PERFORMANCE — {date_str}"
    ws["A1"].font = FONT_TITLE
    ws["A1"].fill = FILL_ACCENT
    ws.merge_cells("A1:H1")
    ws["A1"].alignment = LEFT

    headers = ["Product", "Gross Sales", "Orders", "Units Sold",
               "Ending Inventory", "Sell-Through Rate", "Inventory Flag", "Notes"]
    ws.append([])
    ws.append(headers)
    header_row = 3
    style_header_row(ws, header_row, len(headers))

    inv_lookup = {p["product"]: p for p in data["inventory"]}
    LOW_INV_THRESHOLD = data.get("low_inventory_threshold", 50)

    r = header_row + 1
    for p in data["top_products"]:
        inv = inv_lookup.get(p["product"], {})
        ending = inv.get("ending_inventory", "")
        sell_through = inv.get("sell_through_rate", "")
        flag = ""
        note = ""
        if isinstance(ending, (int, float)) and ending < LOW_INV_THRESHOLD:
            flag = "LOW STOCK"
            note = "Ending inventory below threshold — reorder soon"
        ws.append([p["product"], p["gross_sales"], p["orders"],
                   p.get("units_sold", inv.get("units_sold", "")), ending, sell_through, flag, note])
        r += 1
    last_row = r - 1

    ncols = len(headers)
    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ncols)}{last_row}"

    for rr in range(header_row + 1, last_row + 1):
        ws.cell(row=rr, column=2).number_format = INR0
        ws.cell(row=rr, column=6).number_format = PCT2
        for c in range(1, ncols + 1):
            ws.cell(row=rr, column=c).border = BORDER

    # Conditional format: low inventory flag highlighted
    rng = f"G{header_row+1}:G{last_row}"
    ws.conditional_formatting.add(
        rng, CellIsRule(operator="equal", formula=['"LOW STOCK"'], fill=PatternFill("solid", fgColor=C_WARN_BG))
    )
    autofit(ws, [46, 13, 9, 11, 15, 16, 13, 40])


# ─────────────────────────── Campaign Performance ───────────────────────────

def build_campaign_performance(ws, data):
    date_str = data["date"]
    ws["A1"] = f"CAMPAIGN PERFORMANCE — {date_str}"
    ws["A1"].font = FONT_TITLE
    ws["A1"].fill = FILL_ACCENT
    ws.merge_cells("A1:K1")

    ws["A3"] = "━━━ CAMPAIGN LEVEL ━━━"
    ws["A3"].font = FONT_SECTION
    ws["A3"].fill = FILL_HEADER
    ws.merge_cells("A3:K3")

    camp_headers = ["Campaign", "Status", "Spend", "Impressions", "Clicks",
                    "CTR", "CPC", "Purchases", "CPA", "ROAS", "Performance"]
    ws.append([])
    ws.append([])
    ws.append(camp_headers)
    camp_header_row = 4
    style_header_row(ws, camp_header_row, len(camp_headers))

    r = camp_header_row + 1
    for c_ in sorted(data["campaigns"], key=lambda x: -x["spend"]):
        if c_["spend"] == 0:
            continue
        perf = "SCALE" if c_["roas"] >= 4 else ("REVIEW" if c_["roas"] < 2.5 else "OK")
        ws.append([
            c_["name"], c_["status"], c_["spend"], c_["impressions"], c_["clicks"],
            round(c_["ctr"] / 100, 4), c_["cpc"], c_["purchases"], c_["cpa"], c_["roas"], perf,
        ])
        r += 1
    camp_end = r - 1

    ncols = len(camp_headers)
    for rr in range(camp_header_row + 1, camp_end + 1):
        ws.cell(row=rr, column=3).number_format = INR
        ws.cell(row=rr, column=6).number_format = PCT2
        ws.cell(row=rr, column=7).number_format = INR
        ws.cell(row=rr, column=9).number_format = INR
        ws.cell(row=rr, column=10).number_format = '0.00"x"'
        for c in range(1, ncols + 1):
            ws.cell(row=rr, column=c).border = BORDER

    rng = f"J{camp_header_row+1}:J{camp_end}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["1.5"], fill=PatternFill("solid", fgColor=C_BAD_BG)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["1.5", "3"], fill=PatternFill("solid", fgColor=C_WARN_BG)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["3"], fill=PatternFill("solid", fgColor=C_GOOD_BG)))
    rng_perf = f"K{camp_header_row+1}:K{camp_end}"
    ws.conditional_formatting.add(rng_perf, CellIsRule(operator="equal", formula=['"SCALE"'], fill=PatternFill("solid", fgColor=C_GOOD_BG)))
    ws.conditional_formatting.add(rng_perf, CellIsRule(operator="equal", formula=['"REVIEW"'], fill=PatternFill("solid", fgColor=C_BAD_BG)))

    # ── Ad set level ──
    as_start = camp_end + 3
    ws.cell(row=as_start, column=1, value="━━━ AD SET LEVEL ━━━")
    ws.cell(row=as_start, column=1).font = FONT_SECTION
    ws.cell(row=as_start, column=1).fill = FILL_HEADER
    ws.merge_cells(start_row=as_start, start_column=1, end_row=as_start, end_column=12)

    as_headers = ["Ad Set", "Campaign", "Spend", "Impressions", "Clicks",
                  "CTR", "CPC", "Purchases", "CPA", "ROAS", "Delivery", "Flag"]
    as_header_row = as_start + 1
    for i, h in enumerate(as_headers, start=1):
        ws.cell(row=as_header_row, column=i, value=h)
    style_header_row(ws, as_header_row, len(as_headers))

    r = as_header_row + 1
    for a in sorted(data["adsets"], key=lambda x: -x["spend"]):
        flag = ""
        if "narrow" in a.get("delivery_note", "").lower():
            flag = "NARROW AUDIENCE"
        elif a.get("delivery_status") == "warning":
            flag = "LEARNING ISSUE"
        ws.cell(row=r, column=1, value=a["name"])
        ws.cell(row=r, column=2, value=a["campaign"])
        ws.cell(row=r, column=3, value=a["spend"])
        ws.cell(row=r, column=4, value=a["impressions"])
        ws.cell(row=r, column=5, value=a["clicks"])
        ws.cell(row=r, column=6, value=round(a["ctr"] / 100, 4))
        ws.cell(row=r, column=7, value=a["cpc"])
        ws.cell(row=r, column=8, value=a["purchases"])
        ws.cell(row=r, column=9, value=a["cpa"])
        ws.cell(row=r, column=10, value=a["roas"] if a["roas"] is not None else None)
        ws.cell(row=r, column=11, value=a.get("delivery_status", ""))
        ws.cell(row=r, column=12, value=flag)
        r += 1
    as_end = r - 1

    for rr in range(as_header_row + 1, as_end + 1):
        ws.cell(row=rr, column=3).number_format = INR
        ws.cell(row=rr, column=6).number_format = PCT2
        ws.cell(row=rr, column=7).number_format = INR
        ws.cell(row=rr, column=9).number_format = INR
        ws.cell(row=rr, column=10).number_format = '0.00"x"'
        for c in range(1, len(as_headers) + 1):
            ws.cell(row=rr, column=c).border = BORDER

    rng2 = f"L{as_header_row+1}:L{as_end}"
    ws.conditional_formatting.add(rng2, CellIsRule(operator="equal", formula=['"NARROW AUDIENCE"'], fill=PatternFill("solid", fgColor=C_WARN_BG)))
    ws.conditional_formatting.add(rng2, CellIsRule(operator="equal", formula=['"LEARNING ISSUE"'], fill=PatternFill("solid", fgColor=C_WARN_BG)))

    ws.freeze_panes = f"A{camp_header_row + 1}"
    autofit(ws, [30, 12, 11, 12, 8, 8, 8, 10, 9, 8, 10, 16])


# ─────────────────────────── Recommendations & Notes ────────────────────────

def build_recommendations(ws, data):
    date_str = data["date"]
    ws["A1"] = f"RECOMMENDATIONS & NOTES — {date_str}"
    ws["A1"].font = FONT_TITLE
    ws["A1"].fill = FILL_ACCENT
    ws.merge_cells("A1:B1")

    ws.append([])
    ws.append(["Type", "Recommendation"])
    style_header_row(ws, 3, 2)

    r = 4
    for rec in data["recommendations"]:
        ws.cell(row=r, column=1, value=rec["type"])
        ws.cell(row=r, column=2, value=rec["text"])
        ws.cell(row=r, column=2).alignment = LEFT
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="UNAVAILABLE METRICS")
    ws.cell(row=r, column=1).font = FONT_SECTION
    ws.cell(row=r, column=1).fill = FILL_HEADER
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    r += 1
    for note in data["unavailable_notes"]:
        ws.cell(row=r, column=1, value="ℹ️")
        ws.cell(row=r, column=2, value=note)
        ws.cell(row=r, column=2).alignment = LEFT
        r += 1

    ws.freeze_panes = "A4"
    autofit(ws, [20, 95])
    for rr in range(4, r):
        ws.row_dimensions[rr].height = 30


# ─────────────────────────── Dashboard ──────────────────────────────────────

def build_dashboard(ws, data):
    date_str = data["date"]
    s = data["shopify_today"]
    m = data["meta_today"]

    ws["A1"] = f"DAILY STORE & ADS PERFORMANCE — {date_str}"
    ws["A1"].font = Font(bold=True, color=C_ACCENT_FG, size=16)
    ws["A1"].fill = FILL_ACCENT
    ws.merge_cells("A1:E1")
    ws.row_dimensions[1].height = 26

    def fmt_inr(v):
        return f"₹{v:,.0f}"

    def fmt_pct(v):
        if v is None:
            return "—"
        return f"+{v:.1f}%" if v >= 0 else f"{v:.1f}%"

    rows = [
        [],
        ["━━━ EXECUTIVE SUMMARY ━━━", "", "", "", ""],
        [data["executive_summary"], "", "", "", ""],
        [],
        ["━━━ SHOPIFY PERFORMANCE ━━━", "", "", "", ""],
        ["Metric", "Yesterday", "7-Day Avg", "Change", ""],
        ["Gross Sales", fmt_inr(s["gross_sales"]), fmt_inr(data["shopify_7d_avg"]["gross_sales"]), fmt_pct(s["gross_chg_pct"]), ""],
        ["Net Sales", fmt_inr(s["net_sales"]), fmt_inr(data["shopify_7d_avg"]["net_sales"]), fmt_pct(s["net_chg_pct"]), ""],
        ["Orders", str(s["orders"]), str(round(data["shopify_7d_avg"]["orders"], 1)), fmt_pct(s["orders_chg_pct"]), ""],
        ["AOV", fmt_inr(s["aov"]), fmt_inr(data["shopify_7d_avg"]["aov"]), fmt_pct(s["aov_chg_pct"]), ""],
        ["Returns/Refunds", fmt_inr(s["returns"]), "", "", ""],
        ["Cancellations", str(s.get("cancellations", 0)), "", "", ""],
        [],
        ["━━━ META ADS PERFORMANCE ━━━", "", "", "", ""],
        ["Metric", "Yesterday", "7-Day Avg", "Change", ""],
        ["Total Spend", fmt_inr(m["spend"]), fmt_inr(data["meta_7d_avg"]["spend"]), fmt_pct(m["spend_chg_pct"]), ""],
        ["Impressions", f"{m['impressions']:,}", f"{data['meta_7d_avg']['impressions']:,.0f}", fmt_pct(m["impr_chg_pct"]), ""],
        ["Clicks", f"{m['clicks']:,}", f"{data['meta_7d_avg']['clicks']:,.0f}", fmt_pct(m["clicks_chg_pct"]), ""],
        ["CTR", f"{m['ctr']:.2f}%", f"{data['meta_7d_avg']['ctr']:.2f}%", fmt_pct(m["ctr_chg_pct"]), ""],
        ["Purchases", str(m.get("purchases", "")), "", "", ""],
        ["Blended ROAS", f"{m['roas']:.2f}x", f"{data['meta_7d_avg']['roas']:.2f}x", fmt_pct(m["roas_chg_pct"]), ""],
        [],
        ["━━━ KEY WINS ━━━", "", "", "", ""],
    ]
    for w in data["key_wins"]:
        rows.append(["", w, "", "", ""])
    rows.append([])
    rows.append(["━━━ KEY ISSUES ━━━", "", "", "", ""])
    for i in data["key_issues"]:
        rows.append(["", i, "", "", ""])
    rows.append([])
    rows.append(["━━━ RECOMMENDED ACTIONS TODAY ━━━", "", "", "", ""])
    for idx, rec in enumerate(data["recommendations"][:7], 1):
        rows.append([f"{idx}.", rec["text"], "", "", ""])

    start_row = 3
    section_rows = []
    for i, row in enumerate(rows):
        rr = start_row + i
        for c, val in enumerate(row, start=1):
            ws.cell(row=rr, column=c, value=val)
        if row and isinstance(row[0], str) and row[0].startswith("━"):
            section_rows.append(rr)

    for sr in section_rows:
        ws.cell(row=sr, column=1).font = FONT_SECTION
        ws.cell(row=sr, column=1).fill = FILL_HEADER
        ws.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=5)

    for sub_row in (start_row + 5, start_row + 13):
        for c in range(1, 5):
            ws.cell(row=sub_row, column=c).font = FONT_BOLD
            ws.cell(row=sub_row, column=c).fill = FILL_SECTION

    exec_row = start_row + 2
    ws.merge_cells(start_row=exec_row, start_column=1, end_row=exec_row, end_column=5)
    ws.cell(row=exec_row, column=1).alignment = LEFT
    ws.row_dimensions[exec_row].height = 45

    last_data_row = start_row + len(rows) - 1
    for rr in range(start_row, last_data_row + 1):
        for c in range(1, 6):
            cell = ws.cell(row=rr, column=c)
            if cell.value not in (None, "") and c == 2 and not (row := ws.cell(row=rr, column=1).value or ""):
                cell.alignment = LEFT

    # Conditional formatting on "Change" column (D) for the two summary tables
    for rng in (f"D{start_row+6}:D{start_row+9}", f"D{start_row+14}:D{start_row+17}"):
        ws.conditional_formatting.add(rng, CellIsRule(operator="beginsWith", formula=['"-"'], fill=PatternFill("solid", fgColor=C_BAD_BG)))
        ws.conditional_formatting.add(rng, CellIsRule(operator="beginsWith", formula=['"+"'], fill=PatternFill("solid", fgColor=C_GOOD_BG)))

    ws.freeze_panes = "A2"
    autofit(ws, [18, 22, 16, 14, 4])

    # ── Charts, placed to the right of the summary tables ──
    shop_ws = ws.parent["Shopify Daily Data"]
    meta_ws = ws.parent["Meta Ads Daily Data"]
    prod_ws = ws.parent["Product Performance"]
    camp_ws = ws.parent["Campaign Performance"]

    shop_last = shop_ws.max_row
    meta_last = meta_ws.max_row

    # Sales trend (line)
    chart1 = LineChart()
    chart1.title = "Gross & Net Sales Trend"
    chart1.style = 12
    chart1.y_axis.title = "INR"
    chart1.x_axis.title = "Date"
    chart1.height, chart1.width = 8, 16
    data_ref = Reference(shop_ws, min_col=2, max_col=3, min_row=1, max_row=shop_last)
    cats_ref = Reference(shop_ws, min_col=1, min_row=2, max_row=shop_last)
    chart1.add_data(data_ref, titles_from_data=True)
    chart1.set_categories(cats_ref)
    ws.add_chart(chart1, "G3")

    # Orders trend (bar)
    chart2 = BarChart()
    chart2.title = "Orders Trend"
    chart2.style = 10
    chart2.y_axis.title = "Orders"
    chart2.x_axis.title = "Date"
    chart2.height, chart2.width = 8, 16
    data_ref2 = Reference(shop_ws, min_col=4, max_col=4, min_row=1, max_row=shop_last)
    chart2.add_data(data_ref2, titles_from_data=True)
    chart2.set_categories(cats_ref)
    ws.add_chart(chart2, "G19")

    # Ad spend trend (bar)
    chart3 = BarChart()
    chart3.title = "Ad Spend Trend"
    chart3.style = 11
    chart3.y_axis.title = "INR"
    chart3.x_axis.title = "Date"
    chart3.height, chart3.width = 8, 16
    data_ref3 = Reference(meta_ws, min_col=2, max_col=2, min_row=1, max_row=meta_last)
    cats_ref3 = Reference(meta_ws, min_col=1, min_row=2, max_row=meta_last)
    chart3.add_data(data_ref3, titles_from_data=True)
    chart3.set_categories(cats_ref3)
    ws.add_chart(chart3, "G35")

    # ROAS trend (line)
    chart4 = LineChart()
    chart4.title = "Blended ROAS Trend"
    chart4.style = 13
    chart4.y_axis.title = "ROAS (x)"
    chart4.x_axis.title = "Date"
    chart4.height, chart4.width = 8, 16
    data_ref4 = Reference(meta_ws, min_col=11, max_col=11, min_row=1, max_row=meta_last)
    chart4.add_data(data_ref4, titles_from_data=True)
    chart4.set_categories(cats_ref3)
    ws.add_chart(chart4, "G51")

    # Top products (bar)
    prod_last = prod_ws.max_row
    chart5 = BarChart()
    chart5.type = "bar"
    chart5.title = "Top Products by Gross Sales (Yesterday)"
    chart5.style = 10
    chart5.height, chart5.width = 8, 16
    data_ref5 = Reference(prod_ws, min_col=2, max_col=2, min_row=3, max_row=prod_last)
    cats_ref5 = Reference(prod_ws, min_col=1, min_row=4, max_row=prod_last)
    chart5.add_data(data_ref5, titles_from_data=True)
    chart5.set_categories(cats_ref5)
    ws.add_chart(chart5, "P3")

    # Top campaigns (bar) - find campaign table bounds
    camp_header_row = 4
    r = camp_header_row + 1
    while camp_ws.cell(row=r, column=1).value not in (None, ""):
        r += 1
    camp_end = r - 1
    chart6 = BarChart()
    chart6.type = "bar"
    chart6.title = "Top Campaigns by Spend (Yesterday)"
    chart6.style = 11
    chart6.height, chart6.width = 8, 16
    data_ref6 = Reference(camp_ws, min_col=3, max_col=3, min_row=camp_header_row, max_row=camp_end)
    cats_ref6 = Reference(camp_ws, min_col=1, min_row=camp_header_row + 1, max_row=camp_end)
    chart6.add_data(data_ref6, titles_from_data=True)
    chart6.set_categories(cats_ref6)
    ws.add_chart(chart6, "P19")


if __name__ == "__main__":
    main()
