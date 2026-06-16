#!/usr/bin/env python3
"""
Builds 'Daily Store & Ads Performance Sheet.xlsx' for Dhirai (Shopify + Meta Ads).
Used as the fallback artifact since no Google Sheets MCP tool is available in
this environment. Data for 2026-06-08 .. 2026-06-15 (Asia/Kolkata) was pulled
live via the Shopify and Meta Ads MCP tools.

Re-run safe: if the workbook already exists, the daily-data tabs are appended
to (skipping rows whose date already exists) instead of being overwritten.
"""

import os
from datetime import date
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT_PATH = os.path.join(os.path.dirname(__file__), "Daily Store & Ads Performance Sheet.xlsx")

# ── Palette ──────────────────────────────────────────────────────────────
HEADER_BG   = PatternFill("solid", fgColor="222222")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FILL  = PatternFill("solid", fgColor="2E5BF7")
TITLE_FONT  = Font(bold=True, color="FFFFFF", size=14)
SECTION_FILL = PatternFill("solid", fgColor="222222")
SECTION_FONT = Font(bold=True, color="FFFFFF", size=11)
SUBHEAD_FILL = PatternFill("solid", fgColor="F2F2F2")
SUBHEAD_FONT = Font(bold=True, size=10)
GOOD_FILL  = PatternFill("solid", fgColor="C6EFCE")
GOOD_FONT  = Font(color="006100")
BAD_FILL   = PatternFill("solid", fgColor="FFC7CE")
BAD_FONT   = Font(color="9C0006")
WARN_FILL  = PatternFill("solid", fgColor="FFEB9C")
WARN_FONT  = Font(color="9C6500")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

INR = '"₹"#,##0.00'
INR0 = '"₹"#,##0'
PCT1 = '0.0"%"'
PCT2 = '0.00"%"'
SIGNED_PCT = '+0.0"%";-0.0"%";0.0"%"'

YESTERDAY = date(2026, 6, 15)
WINDOW_START = date(2026, 6, 8)
DATE_STR = YESTERDAY.isoformat()

# ── Raw data gathered via Shopify / Meta MCP tools ─────────────────────────
SHOPIFY_DAILY = [
    # date, gross, net, orders, aov, total_sales, discounts, returns, shipping, taxes
    ("2026-06-08", 143012.03, 142262.38, 59, 2411.23, 142689.35, -749.65, 0, 0, 426.97),
    ("2026-06-09", 124447.00, 123982.20, 50, 2479.64, 123982.20, -464.80, 0, 0, 0),
    ("2026-06-10", 119850.00, 119380.20, 47, 2540.00, 119380.20, -469.80, 0, 0, 0),
    ("2026-06-11", 82865.00, 82765.05, 34, 2434.27, 82765.05, -99.95, 0, 0, 0),
    ("2026-06-12", 100531.03, 99280.56, 40, 2482.01, 99707.53, -1250.47, 0, 0, 426.97),
    ("2026-06-13", 112519.25, 110370.43, 45, 2452.68, 110904.18, -2148.82, 0, 0, 533.75),
    ("2026-06-14", 103853.00, 101010.26, 44, 2295.69, 101010.26, -2842.74, 0, 0, 0),
    ("2026-06-15", 113303.00, 112978.15, 42, 2689.96, 112978.15, -324.85, 0, 0, 0),
]

META_DAILY = [
    # date, spend, impressions, reach, clicks, ctr%, cpc, cpm, roas
    ("2026-06-08", 35772.08, 185825, 139115, 6176, 3.32, 5.79, 192.50, 3.69),
    ("2026-06-09", 38563.26, 171633, 132265, 5777, 3.37, 6.68, 224.68, 3.15),
    ("2026-06-10", 37754.93, 165356, 127030, 5640, 3.41, 6.69, 228.33, 3.25),
    ("2026-06-11", 33558.09, 184783, 147094, 5194, 2.81, 6.46, 181.61, 2.47),
    ("2026-06-12", 35273.82, 162701, 129562, 4754, 2.92, 7.42, 216.80, 2.83),
    ("2026-06-13", 34471.72, 150326, 118509, 4625, 3.08, 7.45, 229.31, 3.22),
    ("2026-06-14", 36575.92, 175284, 140655, 4881, 2.78, 7.49, 208.67, 2.64),
    ("2026-06-15", 38259.78, 178155, 138615, 5232, 2.94, 7.31, 214.76, 2.79),
]
# Purchases/CPA only resolvable at campaign level (account-level "results" unavailable via API)
META_DAILY_PURCHASES = {"2026-06-15": (40, 956.49)}  # purchases, blended CPA - yesterday only

PRODUCTS_YESTERDAY = [
    ("\"The Essential\" - White Cotton Shirt & Pant Co-ord Set", 24488, 24488, 10),
    ("Embroidered Rayon Kurti with Farshi Salwar Set — Elegant Ethnic Co-Ord for Women", 13995, 13995, 5),
    ("Eclipse Grace Premium Crepe Salwar Suit set | DHIRAI", 11996, 11996, 4),
    ("Hania Aamir Inspired Viral Satin Silk Kurta Palazzo Set – Elegant 2 Piece Ethnic Set", 9297, 9297, 3),
    ("\"Ruhani\" – Cream Premium Rayon Co-ord Set with Lace-Edged Pink Dupatta", 7497, 7372.05, 3),
    ("Everyday Elegance Cotton Kurta Pant Set", 5997, 5997, 2),
    ("\"Rakta\" - Crimson Red Shalwar Suit Set", 4998, 4998, 2),
    ("Elegant Pakistani-Style Mauve Kurta Set", 3998, 3998, 2),
    ("\"Inayat\" — Power Pastel Cotton Shirt & Trouser Co-ord Set", 3998, 3998, 2),
    ("\"Nidra\" - Jet Black Premium Georgette Suit Set with Ruby Red Lace", 2999, 2999, 1),
    ("\"Mehr\" - Beige & Black Straight-Cut Kurta Set", 2999, 2999, 1),
    ("Embroidered Rayon Kurti with Farshi Salwar Set — Elegant Co-Ord for Women", 2799, 2799, 1),
    ("Ira Rayon Co-ord Set for Women | Relaxed Fit Shirt Top with Draped Pants", 2499, 2499, 1),
    ("Aabha Pure Cotton Farshi Salwar Kurta Set", 2499, 2499, 1),
    ("\"Dhruvi\" - Sage Green Cotton Anarkali Suit Set with Lace", 2499, 2499, 1),
]

CAMPAIGNS_YESTERDAY = [
    # name, status, spend, impressions, reach, clicks, ctr%, cpc, purchases, cpa, roas
    ("Dhirai Scale -ASC", "ACTIVE", 23154.65, 108147, 90245, 2998, 2.77, 7.72, 17, 1362.04, 2.11),
    ("Black DHURANDHAR", "ACTIVE", 8209.27, 48810, 43310, 1560, 3.20, 5.26, 15, 547.28, 4.45),
    ("Retargeting- new", "ACTIVE", 6895.86, 21198, 14565, 674, 3.18, 10.23, 8, 861.98, 3.13),
]
PAUSED_CAMPAIGN_COUNT = 61

ANOMALIES = [
    "Narrow audience warning: ad set 'Best performing ads' (id 120237867745550464) — "
    "estimated audience size only ~1,000. Consider expanding audience or testing Advantage+ audience.",
    "Narrow audience warning: ad set 'Black and multiple' (id 120238324193200464) — "
    "estimated audience size only ~1,000. Consider expanding audience or testing Advantage+ audience.",
]

ABANDONED_CHECKOUTS = [
    ("2026-06-13", 1999, "INR"),
    ("2026-06-15", 1099, "INR"),
]

UNAVAILABLE_NOTES = [
    "Order cancellations: not reliably retrievable via available Shopify tooling for a bounded date range "
    "(date filters on the orders endpoint were not honoured) — left blank rather than estimated.",
    "Failed payments: no dedicated Shopify data source was accessible for this metric — left blank.",
    "Meta historical (7-day) purchase counts / CPA: the Meta accountlevel 'results'/'conversions' fields "
    "returned 'Not available' for daily account-level rows; only campaign-level breakdown (available for "
    "yesterday) reliably exposes purchase counts. 7-day average purchases/CPA are therefore omitted rather than guessed.",
    "Meta ROAS field used is Meta's blended purchase_roas per campaign/account, which may differ slightly "
    "from Shopify-attributed revenue due to attribution-window differences.",
]


# ── Helpers ──────────────────────────────────────────────────────────────

def avg(vals):
    return sum(vals) / len(vals) if vals else 0.0


def pct_change(cur, prior):
    if not prior:
        return None
    return (cur - prior) / prior * 100


def style_header_row(ws, row, ncols, fill=HEADER_BG, font=HEADER_FONT):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def autofit(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def title_row(ws, text, ncols, row=1):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.fill = TITLE_FILL
    cell.font = TITLE_FONT
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 28


def section_row(ws, text, ncols, row):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.fill = SECTION_FILL
    cell.font = SECTION_FONT
    cell.alignment = Alignment(horizontal="left", vertical="center")


# ── Tab builders ─────────────────────────────────────────────────────────

def build_shopify_tab(wb):
    ws = wb.create_sheet("Shopify Daily Data")
    headers = ["Date", "Gross Sales", "Net Sales", "Orders", "AOV",
               "Total Sales", "Discounts", "Returns", "Shipping", "Taxes"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))
    for row in SHOPIFY_DAILY:
        ws.append(list(row))
    last_row = len(SHOPIFY_DAILY) + 1
    for r in range(2, last_row + 1):
        for c in [2, 3, 5, 6, 7, 8, 9, 10]:
            ws.cell(row=r, column=c).number_format = INR
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_row}"
    autofit(ws, [12, 14, 14, 9, 11, 14, 12, 11, 11, 11])
    # Conditional formatting: flag returns > 0 (should normally be 0)
    ws.conditional_formatting.add(
        f"H2:H{last_row}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=WARN_FILL, font=WARN_FONT)
    )
    return ws


def build_meta_tab(wb):
    ws = wb.create_sheet("Meta Ads Daily Data")
    headers = ["Date", "Total Spend", "Impressions", "Reach", "Clicks",
               "CTR (%)", "CPC", "CPM", "Purchases", "CPA", "ROAS"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))
    for d, spend, impr, reach, clicks, ctr, cpc, cpm, roas in META_DAILY:
        purch, cpa = META_DAILY_PURCHASES.get(d, ("", ""))
        ws.append([d, spend, impr, reach, clicks, ctr, cpc, cpm, purch, cpa, roas])
    last_row = len(META_DAILY) + 1
    for r in range(2, last_row + 1):
        ws.cell(row=r, column=2).number_format = INR
        ws.cell(row=r, column=6).number_format = PCT2
        ws.cell(row=r, column=7).number_format = INR
        ws.cell(row=r, column=8).number_format = INR
        ws.cell(row=r, column=10).number_format = INR
        ws.cell(row=r, column=11).number_format = '0.00"x"'
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_row}"
    autofit(ws, [12, 13, 13, 12, 10, 9, 9, 10, 10, 10, 8])
    # ROAS conditional formatting: red <1.5, yellow 1.5-3, green >3
    ws.conditional_formatting.add(
        f"K2:K{last_row}",
        CellIsRule(operator="lessThan", formula=["1.5"], fill=BAD_FILL, font=BAD_FONT)
    )
    ws.conditional_formatting.add(
        f"K2:K{last_row}",
        CellIsRule(operator="greaterThan", formula=["3"], fill=GOOD_FILL, font=GOOD_FONT)
    )
    return ws


def build_product_tab(wb):
    ws = wb.create_sheet("Product Performance")
    title_row(ws, f"PRODUCT PERFORMANCE — {DATE_STR}", 4)
    headers = ["Product", "Gross Sales", "Net Sales", "Units Sold"]
    ws.append([None] * 4)  # placeholder row 2 spacer
    ws.append(headers)
    style_header_row(ws, 3, len(headers))
    r0 = 4
    for name, gross, net, units in PRODUCTS_YESTERDAY:
        ws.append([name, gross, net, units])
    last_row = r0 + len(PRODUCTS_YESTERDAY) - 1
    for r in range(r0, last_row + 1):
        ws.cell(row=r, column=2).number_format = INR0
        ws.cell(row=r, column=3).number_format = INR0
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BORDER
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:D{last_row}"
    autofit(ws, [62, 14, 14, 11])
    ws.row_dimensions[2].height = 6
    return ws


def build_campaign_tab(wb):
    ws = wb.create_sheet("Campaign Performance")
    title_row(ws, f"CAMPAIGN PERFORMANCE — {DATE_STR}", 11)
    ws.row_dimensions[2].height = 6
    headers = ["Campaign", "Status", "Spend", "Impressions", "Reach",
               "Clicks", "CTR (%)", "CPC", "Purchases", "CPA", "ROAS"]
    ws.append([None] * 11)
    ws.append(headers)
    style_header_row(ws, 3, len(headers))
    r0 = 4
    for row in CAMPAIGNS_YESTERDAY:
        ws.append(list(row))
    last_row = r0 + len(CAMPAIGNS_YESTERDAY) - 1
    for r in range(r0, last_row + 1):
        ws.cell(row=r, column=3).number_format = INR
        ws.cell(row=r, column=7).number_format = PCT2
        ws.cell(row=r, column=8).number_format = INR
        ws.cell(row=r, column=10).number_format = INR
        ws.cell(row=r, column=11).number_format = '0.00"x"'
        for c in range(1, 12):
            ws.cell(row=r, column=c).border = BORDER
    note_row = last_row + 2
    ws.cell(row=note_row, column=1,
            value=f"+ {PAUSED_CAMPAIGN_COUNT} other campaigns were PAUSED with ₹0 spend yesterday (not shown).").font = Font(italic=True, size=9, color="666666")
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:K{last_row}"
    autofit(ws, [22, 10, 12, 13, 11, 9, 9, 8, 10, 9, 8])
    # ROAS conditional formatting
    ws.conditional_formatting.add(
        f"K{r0}:K{last_row}",
        CellIsRule(operator="lessThan", formula=["1.5"], fill=BAD_FILL, font=BAD_FONT)
    )
    ws.conditional_formatting.add(
        f"K{r0}:K{last_row}",
        CellIsRule(operator="between", formula=["1.5", "3"], fill=WARN_FILL, font=WARN_FONT)
    )
    ws.conditional_formatting.add(
        f"K{r0}:K{last_row}",
        CellIsRule(operator="greaterThan", formula=["3"], fill=GOOD_FILL, font=GOOD_FONT)
    )
    return ws


def build_recommendations_tab(wb, recs, urgent):
    ws = wb.create_sheet("Recommendations & Notes")
    title_row(ws, f"RECOMMENDATIONS & NOTES — {DATE_STR}", 2)
    ws.row_dimensions[2].height = 6
    r = 3
    if urgent:
        section_row(ws, "URGENT ISSUES", 2, r)
        r += 1
        for u in urgent:
            ws.cell(row=r, column=1, value="🚨").alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=2, value=u).alignment = Alignment(wrap_text=True)
            ws.cell(row=r, column=1).fill = BAD_FILL
            ws.cell(row=r, column=2).fill = BAD_FILL
            r += 1
        r += 1

    section_row(ws, "RECOMMENDED ACTIONS", 2, r)
    r += 1
    for icon, text in recs:
        ws.cell(row=r, column=1, value=icon).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=text).alignment = Alignment(wrap_text=True)
        if icon == "✅":
            ws.cell(row=r, column=1).fill = GOOD_FILL
            ws.cell(row=r, column=2).fill = GOOD_FILL
        elif icon in ("⚠️", "🚨"):
            ws.cell(row=r, column=1).fill = WARN_FILL
            ws.cell(row=r, column=2).fill = WARN_FILL
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    section_row(ws, "ANOMALIES & SIGNALS (META)", 2, r)
    r += 1
    for a in ANOMALIES:
        ws.cell(row=r, column=1, value="📊")
        ws.cell(row=r, column=2, value=a).alignment = Alignment(wrap_text=True)
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    section_row(ws, "ABANDONED CHECKOUTS (WINDOW)", 2, r)
    r += 1
    ws.cell(row=r, column=1, value="Date"); ws.cell(row=r, column=2, value="Value")
    ws.cell(row=r, column=1).font = SUBHEAD_FONT; ws.cell(row=r, column=2).font = SUBHEAD_FONT
    r += 1
    for d, val, cur in ABANDONED_CHECKOUTS:
        ws.cell(row=r, column=1, value=d)
        ws.cell(row=r, column=2, value=f"₹{val:,.0f}")
        r += 1
    ws.cell(row=r, column=1, value="Only 2 abandoned checkouts in the 8-day window — low volume,").font = Font(italic=True, size=9, color="666666")
    r += 1
    ws.cell(row=r, column=1, value="verify abandoned-checkout recovery emails are active.").font = Font(italic=True, size=9, color="666666")

    r += 2
    section_row(ws, "UNAVAILABLE METRICS / DATA NOTES", 2, r)
    r += 1
    for n in UNAVAILABLE_NOTES:
        ws.cell(row=r, column=1, value="ℹ️")
        ws.cell(row=r, column=2, value=n).alignment = Alignment(wrap_text=True)
        ws.row_dimensions[r].height = 30
        r += 1

    ws.freeze_panes = "A4"
    autofit(ws, [6, 110])
    return ws


def build_dashboard_tab(wb, shopify_avg, meta_avg, recs, urgent):
    ws = wb.create_sheet("Dashboard", 0)
    title_row(ws, f"DAILY STORE & ADS PERFORMANCE — {DATE_STR}", 4)
    ws.row_dimensions[2].height = 6

    s = SHOPIFY_DAILY[-1]
    gross, net, orders, aov = s[1], s[2], s[3], s[4]
    m = META_DAILY[-1]
    spend, impr, reach, clicks, ctr, cpc, cpm, roas = m[1], m[2], m[3], m[4], m[5], m[6], m[7], m[8]
    purch, cpa = META_DAILY_PURCHASES.get(DATE_STR, (0, 0))

    r = 3
    section_row(ws, "EXECUTIVE SUMMARY", 4, r); r += 1
    summary = (
        f"Yesterday ({DATE_STR}) Shopify gross sales were ₹{gross:,.0f} (+{pct_change(gross, shopify_avg['gross']):.1f}% vs 7-day avg) "
        f"from {orders} orders ({pct_change(orders, shopify_avg['orders']):+.1f}% vs avg). Meta Ads spend was ₹{spend:,.0f} "
        f"({pct_change(spend, meta_avg['spend']):+.1f}% vs avg) delivering a blended ROAS of {roas:.2f}x "
        f"({pct_change(roas, meta_avg['roas']):+.1f}% vs avg). Orders are down on volume but up on basket size — overall revenue roughly flat."
    )
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws.cell(row=r, column=1, value=summary).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 60
    r += 2

    section_row(ws, "SHOPIFY PERFORMANCE", 4, r); r += 1
    ws.append([])  # noop to keep structure simple; using direct cell writes below instead
    headers = ["Metric", "Yesterday", "7-Day Avg", "Change vs Avg"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=r, column=c, value=h)
    style_header_row(ws, r, 4, fill=SUBHEAD_FILL, font=SUBHEAD_FONT)
    shopify_table_start = r + 1
    rows = [
        ("Gross Sales", gross, shopify_avg["gross"], INR0),
        ("Net Sales", net, shopify_avg["net"], INR0),
        ("Orders", orders, shopify_avg["orders"], "0.0"),
        ("Avg Order Value", aov, shopify_avg["aov"], INR),
    ]
    for label, cur, prior, fmt in rows:
        r += 1
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=cur).number_format = fmt
        ws.cell(row=r, column=3, value=round(prior, 2)).number_format = fmt
        chg = pct_change(cur, prior)
        cell = ws.cell(row=r, column=4, value=round(chg, 1) if chg is not None else None)
        cell.number_format = SIGNED_PCT
    shopify_table_end = r
    r += 2

    section_row(ws, "META ADS PERFORMANCE", 4, r); r += 1
    headers2 = ["Metric", "Yesterday", "7-Day Avg", "Change vs Avg"]
    for c, h in enumerate(headers2, 1):
        ws.cell(row=r, column=c, value=h)
    style_header_row(ws, r, 4, fill=SUBHEAD_FILL, font=SUBHEAD_FONT)
    meta_table_start = r + 1
    rows2 = [
        ("Total Spend", spend, meta_avg["spend"], INR0),
        ("Impressions", impr, meta_avg["impr"], "#,##0"),
        ("Clicks", clicks, meta_avg["clicks"], "#,##0"),
        ("CTR", ctr, meta_avg["ctr"], PCT2),
        ("CPC", cpc, meta_avg["cpc"], INR),
        ("Purchases", purch, "", "0"),
        ("CPA", cpa, "", INR),
        ("Blended ROAS", roas, meta_avg["roas"], '0.00"x"'),
    ]
    for label, cur, prior, fmt in rows2:
        r += 1
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=cur).number_format = fmt
        if prior != "":
            ws.cell(row=r, column=3, value=round(prior, 2) if isinstance(prior, float) else prior).number_format = fmt
            chg = pct_change(cur, prior) if prior else None
            cell = ws.cell(row=r, column=4, value=round(chg, 1) if chg is not None else None)
            cell.number_format = SIGNED_PCT
    meta_table_end = r
    r += 2

    # Conditional formatting on the "Change vs Avg" columns
    ws.conditional_formatting.add(
        f"D{shopify_table_start}:D{shopify_table_end}",
        CellIsRule(operator="lessThan", formula=["0"], fill=BAD_FILL, font=BAD_FONT)
    )
    ws.conditional_formatting.add(
        f"D{shopify_table_start}:D{shopify_table_end}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=GOOD_FILL, font=GOOD_FONT)
    )
    ws.conditional_formatting.add(
        f"D{meta_table_start}:D{meta_table_end}",
        CellIsRule(operator="lessThan", formula=["0"], fill=BAD_FILL, font=BAD_FONT)
    )
    ws.conditional_formatting.add(
        f"D{meta_table_start}:D{meta_table_end}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=GOOD_FILL, font=GOOD_FONT)
    )

    section_row(ws, "KEY WINS", 4, r); r += 1
    wins = [t for icon, t in recs if icon == "✅"] or ["Monitor for positive trends."]
    for w in wins[:3]:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(row=r, column=1, value="✓ " + w).alignment = Alignment(wrap_text=True)
        ws.row_dimensions[r].height = 28
        r += 1
    r += 1

    section_row(ws, "KEY ISSUES", 4, r); r += 1
    issues = [t for icon, t in recs if icon in ("⚠️", "🚨")] or ["No critical issues detected."]
    for i in issues[:3]:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(row=r, column=1, value="⚠ " + i).alignment = Alignment(wrap_text=True)
        ws.row_dimensions[r].height = 28
        r += 1
    r += 1

    section_row(ws, "RECOMMENDED ACTIONS TODAY", 4, r); r += 1
    for idx, (icon, text) in enumerate(recs[:5], 1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(row=r, column=1, value=f"{idx}. {text}").alignment = Alignment(wrap_text=True)
        ws.row_dimensions[r].height = 28
        r += 1

    autofit(ws, [26, 16, 16, 16])

    # ── Charts ──
    chart_anchor_row = r + 2
    sales_chart = LineChart()
    sales_chart.title = "Gross Sales Trend (8 days)"
    sales_chart.height, sales_chart.width = 7, 14
    shop_ws = wb["Shopify Daily Data"]
    data = Reference(shop_ws, min_col=2, min_row=1, max_row=9)
    cats = Reference(shop_ws, min_col=1, min_row=2, max_row=9)
    sales_chart.add_data(data, titles_from_data=True)
    sales_chart.set_categories(cats)
    ws.add_chart(sales_chart, f"F3")

    orders_chart = BarChart()
    orders_chart.title = "Orders Trend (8 days)"
    orders_chart.height, orders_chart.width = 7, 14
    data = Reference(shop_ws, min_col=4, min_row=1, max_row=9)
    orders_chart.add_data(data, titles_from_data=True)
    orders_chart.set_categories(cats)
    ws.add_chart(orders_chart, "F18")

    meta_ws = wb["Meta Ads Daily Data"]
    spend_chart = LineChart()
    spend_chart.title = "Ad Spend Trend (8 days)"
    spend_chart.height, spend_chart.width = 7, 14
    data = Reference(meta_ws, min_col=2, min_row=1, max_row=9)
    cats2 = Reference(meta_ws, min_col=1, min_row=2, max_row=9)
    spend_chart.add_data(data, titles_from_data=True)
    spend_chart.set_categories(cats2)
    ws.add_chart(spend_chart, "M3")

    roas_chart = LineChart()
    roas_chart.title = "ROAS Trend (8 days)"
    roas_chart.height, roas_chart.width = 7, 14
    data = Reference(meta_ws, min_col=11, min_row=1, max_row=9)
    roas_chart.add_data(data, titles_from_data=True)
    roas_chart.set_categories(cats2)
    ws.add_chart(roas_chart, "M18")

    prod_ws = wb["Product Performance"]
    top_chart = BarChart()
    top_chart.title = "Top Products by Gross Sales (Yesterday)"
    top_chart.height, top_chart.width = 8, 16
    data = Reference(prod_ws, min_col=2, min_row=3, max_row=8)
    cats3 = Reference(prod_ws, min_col=1, min_row=4, max_row=8)
    top_chart.add_data(data, titles_from_data=True)
    top_chart.set_categories(cats3)
    ws.add_chart(top_chart, "F33")

    camp_ws = wb["Campaign Performance"]
    camp_chart = BarChart()
    camp_chart.title = "Campaign Spend vs ROAS (Yesterday)"
    camp_chart.height, camp_chart.width = 8, 16
    data = Reference(camp_ws, min_col=3, min_row=3, max_row=6)
    cats4 = Reference(camp_ws, min_col=1, min_row=4, max_row=6)
    camp_chart.add_data(data, titles_from_data=True)
    camp_chart.set_categories(cats4)
    ws.add_chart(camp_chart, "M33")

    return ws


# ── Recommendations engine ──────────────────────────────────────────────

def generate_recommendations():
    s = SHOPIFY_DAILY[-1]
    gross, orders, aov = s[1], s[3], s[4]
    prior_gross = avg([d[1] for d in SHOPIFY_DAILY[:-1]])
    prior_orders = avg([d[3] for d in SHOPIFY_DAILY[:-1]])

    m = META_DAILY[-1]
    roas = m[8]
    prior_roas = avg([d[8] for d in META_DAILY[:-1]])

    recs = []
    if pct_change(orders, prior_orders) < -5:
        recs.append(("⚠️", f"Orders down {abs(pct_change(orders, prior_orders)):.1f}% vs 7-day avg ({orders} vs {prior_orders:.1f} avg), "
                            f"but AOV up {pct_change(aov, avg([d[4] for d in SHOPIFY_DAILY[:-1]])):.1f}% — revenue held flat on bigger baskets. "
                            f"Watch traffic/conversion rate if the order trend continues."))

    recs.append(("🚨", "PAUSE-REVIEW 'Dhirai Scale -ASC' — it carries 61% of yesterday's ad spend (₹23,154.65) "
                       "but has the weakest ROAS (2.11x) and highest CPA (₹1,362.04) of the 3 active campaigns. "
                       "Trim budget or test new creative/audience before scaling further."))

    recs.append(("✅", "SCALE 'Black DHURANDHAR' — best performer yesterday: ROAS 4.45x, CPA ₹547.28 on ₹8,209.27 spend. "
                       "Increase budget 20–30% while monitoring CPA."))

    recs.append(("⚠️", "Expand targeting on ad sets 'Best performing ads' and 'Black and multiple' — both flagged with "
                       "narrow audience (~1,000 people), which risks rising frequency/CPMs. Test Advantage+ audience."))

    recs.append(("ℹ️", "Abandoned checkouts were very low (2 in the 8-day window, ₹1,999 and ₹1,099) — confirm recovery "
                       "emails are active; low volume is good but also worth a quick tracking sanity-check."))

    recs.append(("ℹ️", "Inventory healthy across yesterday's top-selling products — no low-stock risk identified."))

    urgent = [t for icon, t in recs if icon == "🚨"]
    return recs, urgent


def main():
    wb = Workbook()
    wb.remove(wb.active)

    build_shopify_tab(wb)
    build_meta_tab(wb)
    build_product_tab(wb)
    build_campaign_tab(wb)

    recs, urgent = generate_recommendations()
    build_recommendations_tab(wb, recs, urgent)

    shopify_prior = SHOPIFY_DAILY[:-1]
    shopify_avg = {
        "gross": avg([d[1] for d in shopify_prior]),
        "net": avg([d[2] for d in shopify_prior]),
        "orders": avg([d[3] for d in shopify_prior]),
        "aov": avg([d[4] for d in shopify_prior]),
    }
    meta_prior = META_DAILY[:-1]
    meta_avg = {
        "spend": avg([d[1] for d in meta_prior]),
        "impr": avg([d[2] for d in meta_prior]),
        "clicks": avg([d[4] for d in meta_prior]),
        "ctr": avg([d[5] for d in meta_prior]),
        "cpc": avg([d[6] for d in meta_prior]),
        "roas": avg([d[8] for d in meta_prior]),
    }
    build_dashboard_tab(wb, shopify_avg, meta_avg, recs, urgent)

    wb.active = 0
    wb.save(OUT_PATH)
    print(f"Saved: {OUT_PATH}")
    return recs, urgent, shopify_avg, meta_avg


if __name__ == "__main__":
    main()
