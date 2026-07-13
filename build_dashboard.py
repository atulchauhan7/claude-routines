#!/usr/bin/env python3
"""
Builds Daily_Store_and_Ads_Performance_Sheet.xlsx from data/history.json.

Google Sheets is not connected in this environment, so this generates the
best-available equivalent: a formatted, multi-tab .xlsx workbook with the
same structure a Google Sheet dashboard would have (frozen headers, filters,
currency/percent formats, conditional formatting, charts, and formulas that
pull the Dashboard tab from the raw data tabs).

data/history.json is the single source of truth and is append-only per
report date (never overwritten) — that's what gives this script "append,
don't duplicate" behaviour across daily runs. Re-run any day to regenerate
the workbook from the current history.

Usage:
    python3 build_dashboard.py [report_date YYYY-MM-DD] [output_path.xlsx]
"""

import json
import sys
import datetime
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.chart import LineChart, BarChart, Reference

HERE = os.path.dirname(os.path.abspath(__file__))
HISTORY_PATH = os.path.join(HERE, "data", "history.json")

# ─────────────────────── Palette ────────────────────────────────────────
HEADER_FILL   = PatternFill("solid", fgColor="1F2933")
HEADER_FONT   = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT    = Font(color="1F2933", bold=True, size=18)
SUBTITLE_FONT = Font(color="52606D", size=10, italic=True)
SECTION_FILL  = PatternFill("solid", fgColor="E4E7EB")
SECTION_FONT  = Font(color="1F2933", bold=True, size=13)
LABEL_FONT    = Font(color="1F2933", bold=True, size=10)
NOTE_FONT     = Font(color="52606D", size=10)
GREEN_FILL    = PatternFill("solid", fgColor="C6EFCE")
GREEN_FONT    = Font(color="1E7B34")
RED_FILL      = PatternFill("solid", fgColor="FFC7CE")
RED_FONT      = Font(color="9C1F28")
YELLOW_FILL   = PatternFill("solid", fgColor="FFF3B0")
THIN          = Side(style="thin", color="D9DCE1")
BORDER        = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

INR = '"₹"#,##0.00'
INR0 = '"₹"#,##0'
PCT2 = "0.00%"
PCT1 = "0.0%"
SIGNED_PCT = '+0.0%;-0.0%;0.0%'
NUM0 = "#,##0"


def load_history():
    with open(HISTORY_PATH) as f:
        return json.load(f)


def style_header_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_title(ws, text, row=1, span=8, size=18):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(color="1F2933", bold=True, size=size)
    cell.alignment = Alignment(horizontal="left", vertical="center")


def write_section(ws, text, row, span=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    for c in range(1, span + 1):
        ws.cell(row=row, column=c).fill = SECTION_FILL
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SECTION_FONT
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 22


def sorted_dates(d):
    return sorted(d.keys())


# ─────────────────────── Shopify Daily Data ───────────────────────────────

def build_shopify_tab(wb, hist):
    ws = wb.create_sheet("Shopify Daily Data")
    headers = ["Date", "Orders", "Gross Sales", "Discounts", "Returns/Refunds",
               "Net Sales", "Shipping", "Taxes", "Total Sales", "Avg Order Value"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))

    dates = sorted_dates(hist["shopify_daily"])
    first_row = 2
    for i, d in enumerate(dates):
        row = hist["shopify_daily"][d]
        r = first_row + i
        ws.cell(row=r, column=1, value=d)
        ws.cell(row=r, column=2, value=row["orders"])
        ws.cell(row=r, column=3, value=row["gross_sales"])
        ws.cell(row=r, column=4, value=row["discounts"])
        ws.cell(row=r, column=5, value=row["returns"])
        ws.cell(row=r, column=6, value=row["net_sales"])
        ws.cell(row=r, column=7, value=row["shipping_charges"])
        ws.cell(row=r, column=8, value=row["taxes"])
        ws.cell(row=r, column=9, value=row["total_sales"])
        ws.cell(row=r, column=10, value=row["average_order_value"])
        for c in range(1, 11):
            ws.cell(row=r, column=c).border = BORDER
        for c in (3, 4, 5, 6, 7, 8, 9, 10):
            ws.cell(row=r, column=c).number_format = INR

    last_row = first_row + len(dates) - 1
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:J{last_row}"
    set_widths(ws, [13, 9, 14, 13, 15, 14, 11, 11, 14, 15])

    # Conditional formatting: returns spike (more negative = worse)
    ws.conditional_formatting.add(
        f"E{first_row}:E{last_row}",
        ColorScaleRule(start_type="min", start_color="9C1F28",
                        mid_type="percentile", mid_value=50, mid_color="FFFFFF",
                        end_type="max", end_color="FFFFFF"),
    )
    # Orders spike/drop vs neighbours (visual color scale)
    ws.conditional_formatting.add(
        f"B{first_row}:B{last_row}",
        ColorScaleRule(start_type="min", start_color="FFC7CE",
                        mid_type="percentile", mid_value=50, mid_color="FFFFFF",
                        end_type="max", end_color="C6EFCE"),
    )

    # ── Payment / Order Issues (raw, most-recent day) ──────────────────────
    issue_dates = sorted_dates(hist.get("shopify_payment_issues_daily", {}))
    section_row = last_row + 3
    latest_issue_date = issue_dates[-1] if issue_dates else None
    write_section(ws, f"Payment / Order Issues — {latest_issue_date or 'n/a'}", section_row, span=5)
    hdr_row = section_row + 1
    ihdrs = ["Order", "Customer", "Amount", "Status", "Time (IST)"]
    for c, h in enumerate(ihdrs, start=1):
        cell = ws.cell(row=hdr_row, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
    issues = hist.get("shopify_payment_issues_daily", {}).get(latest_issue_date, []) if latest_issue_date else []
    for i, it in enumerate(issues):
        r = hdr_row + 1 + i
        ws.cell(row=r, column=1, value=it["order"])
        ws.cell(row=r, column=2, value=it["customer"])
        c3 = ws.cell(row=r, column=3, value=it["total"]); c3.number_format = INR
        ws.cell(row=r, column=4, value=it["status"])
        ws.cell(row=r, column=5, value=it["time_ist"])
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = BORDER
    if issues:
        last_issue_row = hdr_row + len(issues)
        ws.conditional_formatting.add(
            f"D{hdr_row + 1}:D{last_issue_row}",
            FormulaRule(formula=[f'NOT(ISERROR(SEARCH("VOIDED",D{hdr_row+1})))'], fill=RED_FILL, font=RED_FONT),
        )

    return {"first_row": first_row, "last_row": last_row, "dates": dates}


# ─────────────────────── Meta Ads Daily Data ──────────────────────────────

def build_meta_tab(wb, hist):
    ws = wb.create_sheet("Meta Ads Daily Data")
    headers = ["Date", "Amount Spent", "Impressions", "Reach", "Clicks", "CTR",
               "CPC", "CPM", "ROAS", "Purchases", "Cost / Purchase"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))

    dates = sorted_dates(hist["meta_daily"])
    first_row = 2
    for i, d in enumerate(dates):
        row = hist["meta_daily"][d]
        r = first_row + i
        ws.cell(row=r, column=1, value=d)
        ws.cell(row=r, column=2, value=row["spend"])
        ws.cell(row=r, column=3, value=row["impressions"])
        ws.cell(row=r, column=4, value=row["reach"])
        ws.cell(row=r, column=5, value=row["clicks"])
        ws.cell(row=r, column=6, value=row["ctr"])
        ws.cell(row=r, column=7, value=row["cpc"])
        ws.cell(row=r, column=8, value=row["cpm"])
        ws.cell(row=r, column=9, value=row["roas"])
        ws.cell(row=r, column=10, value=row["purchases"])
        ws.cell(row=r, column=11, value=row["cost_per_purchase"])
        for c in range(1, 12):
            ws.cell(row=r, column=c).border = BORDER
        for c in (2, 7, 8, 11):
            ws.cell(row=r, column=c).number_format = INR
        ws.cell(row=r, column=6).number_format = PCT2
        ws.cell(row=r, column=9).number_format = '0.00"x"'
        for c in (3, 4, 5, 10):
            ws.cell(row=r, column=c).number_format = NUM0

    last_row = first_row + len(dates) - 1
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{last_row}"
    set_widths(ws, [13, 14, 13, 12, 10, 9, 10, 11, 8, 11, 14])

    ws.conditional_formatting.add(
        f"I{first_row}:I{last_row}",
        ColorScaleRule(start_type="num", start_value=1, start_color="FFC7CE",
                        mid_type="num", mid_value=3.5, mid_color="FFF3B0",
                        end_type="num", end_value=6, end_color="C6EFCE"),
    )
    ws.conditional_formatting.add(
        f"B{first_row}:B{last_row}",
        ColorScaleRule(start_type="min", start_color="FFFFFF",
                        end_type="max", end_color="FDBE85"),
    )

    return {"first_row": first_row, "last_row": last_row, "dates": dates}


# ─────────────────────── Product Performance ──────────────────────────────

def build_product_tab(wb, hist):
    ws = wb.create_sheet("Product Performance")
    dates = sorted_dates(hist.get("shopify_products_daily", {}))
    latest = dates[-1] if dates else None
    write_section(ws, f"Top-Selling Products — {latest or 'n/a'}", 1, span=4)
    hdr_row = 2
    headers = ["Product", "Gross Sales", "Net Sales", "Orders"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=hdr_row, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER

    products = hist.get("shopify_products_daily", {}).get(latest, []) if latest else []
    first_row = hdr_row + 1
    for i, p in enumerate(products):
        r = first_row + i
        ws.cell(row=r, column=1, value=p["product"])
        c2 = ws.cell(row=r, column=2, value=p["gross_sales"]); c2.number_format = INR
        c3 = ws.cell(row=r, column=3, value=p["net_sales"]); c3.number_format = INR
        ws.cell(row=r, column=4, value=p["orders"])
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BORDER
    last_row = first_row + len(products) - 1 if products else first_row
    if products:
        ws.freeze_panes = f"A{first_row}"
        ws.auto_filter.ref = f"A{hdr_row}:D{last_row}"
        ws.conditional_formatting.add(
            f"B{first_row}:B{last_row}",
            ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="63B3ED"),
        )
    set_widths(ws, [58, 15, 15, 10])

    # ── Inventory watch ────────────────────────────────────────────────────
    inv_section_row = last_row + 3
    inv_dates = sorted_dates(hist.get("shopify_inventory_watch_daily", {}))
    inv_latest = inv_dates[-1] if inv_dates else None
    write_section(ws, f"Inventory Watch (Low Stock / Out of Stock) — {inv_latest or 'n/a'}", inv_section_row, span=5)
    inv_hdr_row = inv_section_row + 1
    ihdrs = ["Product", "Starting Units", "Ending Units", "Units Sold", "Status"]
    for c, h in enumerate(ihdrs, start=1):
        cell = ws.cell(row=inv_hdr_row, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
    inv_rows = hist.get("shopify_inventory_watch_daily", {}).get(inv_latest, []) if inv_latest else []
    for i, it in enumerate(inv_rows):
        r = inv_hdr_row + 1 + i
        ws.cell(row=r, column=1, value=it["product"])
        ws.cell(row=r, column=2, value=it["starting_units"])
        ws.cell(row=r, column=3, value=it["ending_units"])
        ws.cell(row=r, column=4, value=it["units_sold"])
        ws.cell(row=r, column=5, value=it["status"])
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = BORDER
    if inv_rows:
        inv_last_row = inv_hdr_row + len(inv_rows)
        ws.conditional_formatting.add(
            f"E{inv_hdr_row+1}:E{inv_last_row}",
            FormulaRule(formula=[f'NOT(ISERROR(SEARCH("OUT OF STOCK",E{inv_hdr_row+1})))'], fill=RED_FILL, font=RED_FONT),
        )
        ws.conditional_formatting.add(
            f"E{inv_hdr_row+1}:E{inv_last_row}",
            FormulaRule(formula=[f'NOT(ISERROR(SEARCH("LOW",E{inv_hdr_row+1})))'], fill=YELLOW_FILL),
        )

    return {"latest_date": latest, "first_row": first_row, "last_row": last_row, "count": len(products)}


# ─────────────────────── Campaign Performance ──────────────────────────────

def build_campaign_tab(wb, hist):
    ws = wb.create_sheet("Campaign Performance")
    headers = ["Campaign", "Status", "Spend", "Impressions", "Clicks", "CTR", "CPC", "ROAS", "Purchases", "Cost / Purchase"]

    day_dates = sorted_dates(hist.get("meta_campaigns_daily", {}))
    latest = day_dates[-1] if day_dates else None
    write_section(ws, f"Campaign Performance — Yesterday ({latest or 'n/a'})", 1, span=len(headers))
    hdr_row = 2
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=hdr_row, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER

    day_campaigns = hist.get("meta_campaigns_daily", {}).get(latest, []) if latest else []
    first_row = hdr_row + 1
    for i, cpn in enumerate(day_campaigns):
        r = first_row + i
        ws.cell(row=r, column=1, value=cpn["campaign"])
        ws.cell(row=r, column=2, value=cpn["status"])
        c3 = ws.cell(row=r, column=3, value=cpn["spend"]); c3.number_format = INR
        ws.cell(row=r, column=4, value=cpn["impressions"]).number_format = NUM0
        ws.cell(row=r, column=5, value=cpn["clicks"]).number_format = NUM0
        c6 = ws.cell(row=r, column=6, value=cpn["ctr"]); c6.number_format = PCT2
        c7 = ws.cell(row=r, column=7, value=cpn["cpc"]); c7.number_format = INR
        c8 = ws.cell(row=r, column=8, value=cpn["roas"]); c8.number_format = '0.00"x"'
        ws.cell(row=r, column=9, value=cpn["purchases"]).number_format = NUM0
        c10 = ws.cell(row=r, column=10, value=cpn["cost_per_purchase"]); c10.number_format = INR
        for c in range(1, 11):
            ws.cell(row=r, column=c).border = BORDER
    last_row = first_row + len(day_campaigns) - 1 if day_campaigns else first_row
    if day_campaigns:
        ws.auto_filter.ref = f"A{hdr_row}:J{last_row}"
        ws.freeze_panes = f"A{first_row}"
        ws.conditional_formatting.add(
            f"H{first_row}:H{last_row}",
            CellIsRule(operator="lessThan", formula=["3"], fill=RED_FILL, font=RED_FONT),
        )
        ws.conditional_formatting.add(
            f"H{first_row}:H{last_row}",
            CellIsRule(operator="greaterThanOrEqual", formula=["5"], fill=GREEN_FILL, font=GREEN_FONT),
        )

    # ── 7-day window table ──────────────────────────────────────────────────
    wk_section_row = last_row + 3
    wk = hist.get("meta_campaigns_7day", {})
    write_section(ws, f"Campaign Performance — Previous 7-Day Window ({wk.get('window', 'n/a')})", wk_section_row, span=len(headers))
    wk_hdr_row = wk_section_row + 1
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=wk_hdr_row, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
    wk_campaigns = wk.get("campaigns", [])
    wk_first_row = wk_hdr_row + 1
    for i, cpn in enumerate(wk_campaigns):
        r = wk_first_row + i
        ws.cell(row=r, column=1, value=cpn["campaign"])
        ws.cell(row=r, column=2, value="—")
        c3 = ws.cell(row=r, column=3, value=cpn["spend"]); c3.number_format = INR
        ws.cell(row=r, column=4, value=cpn["impressions"]).number_format = NUM0
        ws.cell(row=r, column=5, value=cpn["clicks"]).number_format = NUM0
        c6 = ws.cell(row=r, column=6, value=cpn["ctr"]); c6.number_format = PCT2
        c7 = ws.cell(row=r, column=7, value=cpn["cpc"]); c7.number_format = INR
        c8 = ws.cell(row=r, column=8, value=cpn["roas"]); c8.number_format = '0.00"x"'
        ws.cell(row=r, column=9, value=cpn["purchases"]).number_format = NUM0
        c10 = ws.cell(row=r, column=10, value=cpn["cost_per_purchase"]); c10.number_format = INR
        for c in range(1, 11):
            ws.cell(row=r, column=c).border = BORDER
    wk_last_row = wk_first_row + len(wk_campaigns) - 1 if wk_campaigns else wk_first_row
    if wk_campaigns:
        ws.conditional_formatting.add(
            f"H{wk_first_row}:H{wk_last_row}",
            CellIsRule(operator="lessThan", formula=["3"], fill=RED_FILL, font=RED_FONT),
        )
        ws.conditional_formatting.add(
            f"H{wk_first_row}:H{wk_last_row}",
            CellIsRule(operator="greaterThanOrEqual", formula=["5"], fill=GREEN_FILL, font=GREEN_FONT),
        )

    # ── Delivery warnings ────────────────────────────────────────────────────
    warn_section_row = wk_last_row + 3
    warn_dates = sorted_dates(hist.get("meta_warnings_daily", {}))
    warn_latest = warn_dates[-1] if warn_dates else None
    write_section(ws, f"Delivery Warnings — {warn_latest or 'n/a'}", warn_section_row, span=4)
    warn_hdr_row = warn_section_row + 1
    whdrs = ["Object", "Warning Type", "Detail", "Recommendation"]
    for c, h in enumerate(whdrs, start=1):
        cell = ws.cell(row=warn_hdr_row, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
    warnings = hist.get("meta_warnings_daily", {}).get(warn_latest, []) if warn_latest else []
    for i, w in enumerate(warnings):
        r = warn_hdr_row + 1 + i
        ws.cell(row=r, column=1, value=w["object"])
        ws.cell(row=r, column=2, value=w["type"])
        ws.cell(row=r, column=3, value=w["detail"])
        ws.cell(row=r, column=4, value=w["recommendation"])
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BORDER
            ws.cell(row=r, column=c).fill = YELLOW_FILL

    set_widths(ws, [26, 10, 13, 13, 10, 9, 9, 9, 11, 14])

    return {
        "latest_date": latest, "first_row": first_row, "last_row": last_row, "count": len(day_campaigns),
        "wk_first_row": wk_first_row, "wk_last_row": wk_last_row,
    }


# ─────────────────────── Recommendations & Notes ──────────────────────────

def build_recommendations_tab(wb, hist, report_date, wins, issues, actions):
    ws = wb.create_sheet("Recommendations & Notes")
    write_title(ws, "Recommendations & Notes", 1, span=6, size=16)
    ws.cell(row=2, column=1, value=f"Report date: {report_date} (Asia/Kolkata)").font = SUBTITLE_FONT

    row = 4
    write_section(ws, "Key Wins", row, span=6); row += 1
    for w in wins:
        ws.cell(row=row, column=1, value=f"✓  {w}").font = Font(color="1E7B34", size=11)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1
    row += 1

    write_section(ws, "Key Issues", row, span=6); row += 1
    for i in issues:
        ws.cell(row=row, column=1, value=f"⚠  {i}").font = Font(color="9C1F28", size=11)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1
    row += 1

    write_section(ws, "Recommended Actions for Today", row, span=6); row += 1
    for n, a in enumerate(actions, start=1):
        ws.cell(row=row, column=1, value=f"{n}. {a}").font = Font(color="1F2933", size=11)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1
    row += 1

    write_section(ws, "Data Notes / Unavailable Metrics", row, span=6); row += 1
    notes = hist.get("data_notes", {}).get(report_date, [])
    if not notes:
        notes = ["No data-availability issues encountered for this report."]
    for note in notes:
        c = ws.cell(row=row, column=1, value=f"•  {note}")
        c.font = NOTE_FONT
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        ws.row_dimensions[row].height = 28
        row += 1

    set_widths(ws, [22, 22, 22, 22, 22, 22])
    return ws


# ─────────────────────── Dashboard ─────────────────────────────────────────

def avg_range(sheet_name, col, first, last):
    return f"AVERAGE('{sheet_name}'!{col}{first}:{col}{last})"


def build_dashboard(wb, hist, report_date, shp_meta, meta_meta, prod_meta, camp_meta, wins, issues, actions):
    ws = wb.create_sheet("Dashboard", 0)
    write_title(ws, "Daily Store & Ads Performance Dashboard", 1, span=8, size=20)
    dt = datetime.datetime.strptime(report_date, "%Y-%m-%d")
    ws.cell(row=2, column=1,
            value=f"Report date: {dt.strftime('%A, %d %B %Y')}  (Asia/Kolkata, full calendar day)  •  Store: Dhirai (dhirai.in)  •  Ad account: Dhirai").font = SUBTITLE_FONT
    ws.merge_cells("A2:H2")

    S = "Shopify Daily Data"
    M = "Meta Ads Daily Data"
    s_last, s_first = shp_meta["last_row"], shp_meta["first_row"]
    s_prev_first, s_prev_last = s_last - 7, s_last - 1
    m_last, m_first = meta_meta["last_row"], meta_meta["first_row"]
    m_prev_first, m_prev_last = m_last - 7, m_last - 1

    row = 4
    write_section(ws, "Executive Summary", row, span=8); row += 1
    exec_rows = [
        ("Net Sales (Shopify, yesterday)", f"='{S}'!F{s_last}", INR),
        ("Orders (Shopify, yesterday)", f"='{S}'!B{s_last}", NUM0),
        ("Ad Spend (Meta, yesterday)", f"='{M}'!B{m_last}", INR),
        ("ROAS (Meta, yesterday)", f"='{M}'!I{m_last}", '0.00"x"'),
        ("Blended: Net Sales ÷ Ad Spend", f"='{S}'!F{s_last}/'{M}'!B{m_last}", '0.00"x"'),
    ]
    for label, formula, fmt in exec_rows:
        ws.cell(row=row, column=1, value=label).font = LABEL_FONT
        c = ws.cell(row=row, column=3, value=formula)
        c.number_format = fmt
        c.font = Font(bold=True, size=12, color="1F2933")
        row += 1
    row += 1

    # ── Shopify performance vs 7-day avg ────────────────────────────────────
    write_section(ws, "Shopify Performance — Yesterday vs Previous 7-Day Average", row, span=8); row += 1
    hdr = row
    for c, h in enumerate(["Metric", "Yesterday", "7-Day Avg", "Change %"], start=1):
        cell = ws.cell(row=hdr, column=c, value=h)
        cell.fill = HEADER_FILL; cell.font = HEADER_FONT; cell.border = BORDER
    row += 1
    shp_metrics = [
        ("Gross Sales", "C", INR), ("Discounts", "D", INR), ("Returns / Refunds", "E", INR),
        ("Net Sales", "F", INR), ("Total Sales", "I", INR), ("Orders", "B", NUM0),
        ("Average Order Value", "J", INR),
    ]
    shp_change_first = row
    for label, col, fmt in shp_metrics:
        ws.cell(row=row, column=1, value=label).border = BORDER
        y = ws.cell(row=row, column=2, value=f"='{S}'!{col}{s_last}"); y.number_format = fmt; y.border = BORDER
        avg = ws.cell(row=row, column=3, value=f"={avg_range(S, col, s_prev_first, s_prev_last)}"); avg.number_format = fmt; avg.border = BORDER
        chg = ws.cell(row=row, column=4, value=f"=IFERROR((B{row}-C{row})/ABS(C{row}),\"n/a\")"); chg.number_format = SIGNED_PCT; chg.border = BORDER
        row += 1
    shp_change_last = row - 1
    ws.conditional_formatting.add(f"D{shp_change_first}:D{shp_change_last}",
                                   CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN_FILL, font=GREEN_FONT))
    ws.conditional_formatting.add(f"D{shp_change_first}:D{shp_change_last}",
                                   CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL, font=RED_FONT))
    row += 1

    # ── Meta Ads performance vs 7-day avg ───────────────────────────────────
    write_section(ws, "Meta Ads Performance — Yesterday vs Previous 7-Day Average", row, span=8); row += 1
    hdr2 = row
    for c, h in enumerate(["Metric", "Yesterday", "7-Day Avg", "Change %"], start=1):
        cell = ws.cell(row=hdr2, column=c, value=h)
        cell.fill = HEADER_FILL; cell.font = HEADER_FONT; cell.border = BORDER
    row += 1
    meta_metrics = [
        ("Amount Spent", "B", INR), ("Impressions", "C", NUM0), ("Reach", "D", NUM0),
        ("Clicks", "E", NUM0), ("CTR", "F", PCT2), ("CPC", "G", INR),
        ("ROAS", "I", '0.00"x"'), ("Purchases", "J", NUM0), ("Cost / Purchase", "K", INR),
    ]
    meta_change_first = row
    for label, col, fmt in meta_metrics:
        ws.cell(row=row, column=1, value=label).border = BORDER
        y = ws.cell(row=row, column=2, value=f"='{M}'!{col}{m_last}"); y.number_format = fmt; y.border = BORDER
        avg = ws.cell(row=row, column=3, value=f"={avg_range(M, col, m_prev_first, m_prev_last)}"); avg.number_format = fmt; avg.border = BORDER
        chg = ws.cell(row=row, column=4, value=f"=IFERROR((B{row}-C{row})/ABS(C{row}),\"n/a\")"); chg.number_format = SIGNED_PCT; chg.border = BORDER
        row += 1
    meta_change_last = row - 1
    # For Cost/Purchase, a decrease is GOOD — flip the color logic just for that row
    cost_row = meta_change_last
    ws.conditional_formatting.add(f"D{meta_change_first}:D{meta_change_last - 1}",
                                   CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN_FILL, font=GREEN_FONT))
    ws.conditional_formatting.add(f"D{meta_change_first}:D{meta_change_last - 1}",
                                   CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL, font=RED_FONT))
    ws.conditional_formatting.add(f"D{cost_row}",
                                   CellIsRule(operator="lessThan", formula=["0"], fill=GREEN_FILL, font=GREEN_FONT))
    ws.conditional_formatting.add(f"D{cost_row}",
                                   CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL, font=RED_FONT))
    row += 1

    write_section(ws, "Key Wins", row, span=8); row += 1
    for w in wins:
        ws.cell(row=row, column=1, value=f"✓  {w}").font = Font(color="1E7B34", size=10)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        row += 1
    row += 1

    write_section(ws, "Key Issues", row, span=8); row += 1
    for i in issues:
        ws.cell(row=row, column=1, value=f"⚠  {i}").font = Font(color="9C1F28", size=10)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        row += 1
    row += 1

    write_section(ws, "Recommended Actions for Today", row, span=8); row += 1
    for n, a in enumerate(actions, start=1):
        ws.cell(row=row, column=1, value=f"{n}. {a}").font = Font(size=10)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        row += 1
    row += 1

    set_widths(ws, [34, 15, 15, 12, 14, 14, 14, 14])

    # ── Charts ───────────────────────────────────────────────────────────
    chart_anchor_col = 10  # column J
    chart_row = 4

    def anchor(col, r):
        return f"{get_column_letter(col)}{r}"

    # Sales trend
    ch1 = LineChart(); ch1.title = "Net Sales Trend (₹)"; ch1.height = 8; ch1.width = 16
    data = Reference(wb["Shopify Daily Data"], min_col=6, min_row=1, max_row=s_last)
    cats = Reference(wb["Shopify Daily Data"], min_col=1, min_row=s_first, max_row=s_last)
    ch1.add_data(data, titles_from_data=True); ch1.set_categories(cats)
    ws.add_chart(ch1, anchor(chart_anchor_col, chart_row))

    # Orders trend
    ch2 = LineChart(); ch2.title = "Orders Trend"; ch2.height = 8; ch2.width = 16
    data = Reference(wb["Shopify Daily Data"], min_col=2, min_row=1, max_row=s_last)
    ch2.add_data(data, titles_from_data=True); ch2.set_categories(cats)
    ws.add_chart(ch2, anchor(chart_anchor_col, chart_row + 17))

    # Ad spend trend
    ch3 = LineChart(); ch3.title = "Meta Ad Spend Trend (₹)"; ch3.height = 8; ch3.width = 16
    mcats = Reference(wb["Meta Ads Daily Data"], min_col=1, min_row=m_first, max_row=m_last)
    data = Reference(wb["Meta Ads Daily Data"], min_col=2, min_row=1, max_row=m_last)
    ch3.add_data(data, titles_from_data=True); ch3.set_categories(mcats)
    ws.add_chart(ch3, anchor(chart_anchor_col, chart_row + 34))

    # ROAS trend
    ch4 = LineChart(); ch4.title = "ROAS Trend"; ch4.height = 8; ch4.width = 16
    data = Reference(wb["Meta Ads Daily Data"], min_col=9, min_row=1, max_row=m_last)
    ch4.add_data(data, titles_from_data=True); ch4.set_categories(mcats)
    ws.add_chart(ch4, anchor(chart_anchor_col, chart_row + 51))

    # Top products bar chart
    if prod_meta["count"]:
        ch5 = BarChart(); ch5.title = "Top Products by Gross Sales (₹)"; ch5.height = 8; ch5.width = 16
        top_n = min(8, prod_meta["count"])
        data = Reference(wb["Product Performance"], min_col=2, min_row=2, max_row=prod_meta["first_row"] + top_n - 1)
        cats = Reference(wb["Product Performance"], min_col=1, min_row=prod_meta["first_row"], max_row=prod_meta["first_row"] + top_n - 1)
        ch5.add_data(data, titles_from_data=True); ch5.set_categories(cats)
        ch5.y_axis.title = "Gross Sales (₹)"
        ws.add_chart(ch5, anchor(chart_anchor_col + 9, chart_row))

    # Top campaigns bar chart (spend + ROAS)
    if camp_meta["count"]:
        ch6 = BarChart(); ch6.title = "Campaigns — Spend vs ROAS (Yesterday)"; ch6.height = 8; ch6.width = 16
        data = Reference(wb["Campaign Performance"], min_col=3, min_row=2, max_row=camp_meta["last_row"])
        cats = Reference(wb["Campaign Performance"], min_col=1, min_row=camp_meta["first_row"], max_row=camp_meta["last_row"])
        ch6.add_data(data, titles_from_data=True); ch6.set_categories(cats)
        ch6.y_axis.title = "Spend (₹)"
        ws.add_chart(ch6, anchor(chart_anchor_col + 9, chart_row + 17))

    ws.sheet_view.showGridLines = False
    return ws


# ─────────────────────── Main ──────────────────────────────────────────────

def main():
    hist = load_history()
    report_date = sys.argv[1] if len(sys.argv) > 1 else sorted_dates(hist["shopify_daily"])[-1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "Daily_Store_and_Ads_Performance_Sheet.xlsx")

    wb = Workbook()
    wb.remove(wb.active)

    shp_meta = build_shopify_tab(wb, hist)
    meta_meta = build_meta_tab(wb, hist)
    prod_meta = build_product_tab(wb, hist)
    camp_meta = build_campaign_tab(wb, hist)

    wins = [
        "Net sales up ~48.9% vs the 7-day average (₹2,21,183 vs ₹1,48,563/day avg), driven by 87 orders — the highest single-day order count in the trailing window.",
        "Meta Ads ROAS jumped to 5.45x (vs 3.38x 7-day avg), the best day in the window, with cost/purchase down ~20% to ₹566 and purchases up ~43% to 84.",
        "\"Retargeting- new\" was the standout campaign — 11.27x ROAS and ₹360 cost/purchase, both the best of any campaign yesterday.",
        "Top 2 products (Ira Rayon Co-ord Set, \"The Essential\" Co-ord Set) alone drove ₹99,256 — ~37% of yesterday's gross sales.",
    ]
    issues = [
        "₹47,285 in Returns/Refunds yesterday (vs ~₹1,035/day avg) — driven almost entirely by 4 VOIDED orders, including 3 from the same customer (Yogesh Martand Aglave, ₹45,286 combined) placed within 2 minutes. Looks like a payment-gateway retry/duplicate-charge failure, not genuine returns.",
        "2 products are OUT OF STOCK (\"Zariya\" Stone Hoops, \"Utsav\" Mal Cotton Anarkali Suit Set) and may be losing sales unnoticed.",
        "\"New Campaign- for recovery\" is the weakest campaign both yesterday (2.02x ROAS, ₹1,295 cost/purchase) and over the past 7 days (2.40x ROAS) despite being the single largest ad spend line (~33% of budget).",
        "Meta flagged the \"Black and multiple\" ad set for a narrow audience (~20,900–24,600 people) — may be limiting delivery/scale.",
        "Shopify's on-site checkout-funnel tracking looks broken: only 1–2 sessions/day show as reaching or completing checkout vs 50–90 real orders/day. Recommendations here rely on order-level data, not funnel data.",
    ]
    actions = [
        "Investigate the payment gateway for the 4 VOIDED orders yesterday (esp. the 3 rapid-fire attempts from Yogesh Martand Aglave, ₹45,286) — confirm no duplicate charge landed on the customer and check gateway error logs for that window (~23:10–23:12 IST).",
        "Restock or unpublish \"Zariya\" Stone Hoops and the \"Utsav\" Anarkali Suit Set — both are at 0 units and still discoverable on-site.",
        "Reduce budget on or restructure \"New Campaign- for recovery\" (consistently ~2–2.4x ROAS, ~2-3x the cost/purchase of other campaigns) and reallocate toward \"Retargeting- new\" and \"Dhirai Scale -ASC\", which are outperforming.",
        "Expand the audience (or enable Advantage+ audience) on the \"Black and multiple\" ad set to fix the narrow-audience delivery warning.",
        "Get engineering/analytics to audit the Shopify checkout funnel tracking (GA/pixel setup) — current session data undercounts checkouts by roughly 50-80x versus actual orders, which blocks meaningful conversion-rate analysis.",
    ]

    build_dashboard(wb, hist, report_date, shp_meta, meta_meta, prod_meta, camp_meta, wins, issues, actions)
    build_recommendations_tab(wb, hist, report_date, wins, issues, actions)

    wb.save(out_path)
    print(f"Wrote {out_path}")
    return out_path


if __name__ == "__main__":
    main()
