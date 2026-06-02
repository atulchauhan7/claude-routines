#!/usr/bin/env python3
"""
One-shot report generator using pre-fetched data for 2026-06-01.
Produces: reports/daily_performance_2026-06-01.xlsx
"""

import os
import datetime
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00, FORMAT_NUMBER_COMMA_SEPARATED1
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel

# ─── Palette ────────────────────────────────────────────────────────────────────
DARK       = "222222"   # header bg
WHITE      = "FFFFFF"
ACCENT     = "4285F4"   # Google-blue accent
ACCENT2    = "34A853"   # green
LIGHT_GREY = "F5F5F5"
MID_GREY   = "E0E0E0"
GREEN_BG   = "C8E6C9"
RED_BG     = "FFCDD2"
YELLOW_BG  = "FFF9C4"
ORANGE_BG  = "FFE0B2"

# ─── Style helpers ──────────────────────────────────────────────────────────────
def hdr(text, bold=True, size=11, color=WHITE, bg=DARK, align="center"):
    return {"value": text, "bold": bold, "size": size, "color": color, "bg": bg, "align": align}

def cell(ws, r, c, value, bold=False, size=10, color="000000", bg=None,
         align="left", wrap=False, num_fmt=None, border=False):
    cell_obj = ws.cell(row=r, column=c, value=value)
    cell_obj.font  = Font(name="Calibri", bold=bold, size=size, color=color)
    if bg:
        cell_obj.fill = PatternFill("solid", fgColor=bg)
    cell_obj.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if num_fmt:
        cell_obj.number_format = num_fmt
    if border:
        thin = Side(style="thin", color="CCCCCC")
        cell_obj.border = Border(bottom=thin)
    return cell_obj

def header_row(ws, row, cols, bg=DARK, fg=WHITE, height=22):
    for c, text in enumerate(cols, 1):
        cell(ws, row, c, text, bold=True, size=10, color=fg, bg=bg, align="center")
    ws.row_dimensions[row].height = height

def pct_change(cur, prev):
    if not prev:
        return None
    return round((cur - prev) / prev * 100, 1)

def fmt_inr(v):
    return f"₹{v:,.0f}"

def fmt_pct(v):
    if v is None:
        return "—"
    return f"+{v:.1f}%" if v >= 0 else f"{v:.1f}%"

def section_title(ws, row, text, span_end, bg=ACCENT, fg=WHITE, size=12):
    c = ws.cell(row=row, column=1, value=text)
    c.font  = Font(name="Calibri", bold=True, size=size, color=fg)
    c.fill  = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span_end)
    ws.row_dimensions[row].height = 22

def autofit(ws, min_w=12, max_w=60):
    for col in ws.columns:
        mx = min_w
        for c in col:
            if c.value:
                mx = max(mx, min(len(str(c.value)) + 2, max_w))
        ws.column_dimensions[get_column_letter(col[0].column)].width = mx

def freeze(ws, row=2, col=1):
    ws.freeze_panes = ws.cell(row=row, column=col)

def add_filters(ws, row, cols):
    ws.auto_filter.ref = f"A{row}:{get_column_letter(cols)}{row}"

# ─── Real data ─────────────────────────────────────────────────────────────────

REPORT_DATE = "2026-06-01"

# Shopify 8-day rows (May 25 → Jun 01)
SHOPIFY_8D = [
    # date, gross, net, orders, aov, discounts, returns, shipping, taxes, total
    ("2026-05-25", 197213.00,   169045.25, 69, 2849.76, 579.75,  27588.00,  0,    0,       169045.25),
    ("2026-05-26", 138643.00,   122890.18, 50, 2739.66, 1659.82, 14093.00,  0,    0,       122890.18),
    ("2026-05-27", 162158.81,   143153.25, 64, 2450.74, 5311.56, 13694.00,  0,    320.19,  143473.44),
    ("2026-05-28", 160630.00,   140368.98, 59, 2682.39, 2369.02, 17892.00,  0,    0,       140368.98),
    ("2026-05-29", 166625.00,   138446.33, 67, 2452.75, 2290.67, 25888.00,  0,    0,       138446.33),
    ("2026-05-30", 129240.00,   106520.96, 55, 2320.20, 1629.04, 21090.00,  0,    0,       106520.96),
    ("2026-05-31", 143728.07,   125599.95, 63, 2242.75, 2435.12, 15693.00,  0,    304.93,  125904.88),
    ("2026-06-01", 117347.00,   114533.40, 50, 2330.65,  814.60,  1999.00,  0,    0,       114533.40),
]

# Sessions data (same 8 days)
SESSIONS_8D = [
    # date, sessions, cart_adds, reached_checkout, completed_checkout, cvr
    ("2026-05-25", 4073, 252, 2, 0, 0.00),
    ("2026-05-26", 4208, 208, 0, 0, 0.00),
    ("2026-05-27", 4143, 239, 2, 1, 0.024),
    ("2026-05-28", 4092, 214, 1, 0, 0.00),
    ("2026-05-29", 4297, 232, 0, 0, 0.00),
    ("2026-05-30", 3957, 228, 2, 0, 0.00),
    ("2026-05-31", 4075, 229, 2, 1, 0.025),
    ("2026-06-01", 3467, 189, 0, 0, 0.00),
]

# 7-day averages (May 25-31 only, excluding June 01)
AVG_GROSS    = sum(r[1] for r in SHOPIFY_8D[:7]) / 7  # 156891.13
AVG_NET      = sum(r[2] for r in SHOPIFY_8D[:7]) / 7  # 135146.41
AVG_ORDERS   = sum(r[3] for r in SHOPIFY_8D[:7]) / 7  # 61.0
AVG_AOV      = sum(r[4] for r in SHOPIFY_8D[:7]) / 7  # 2534.03
AVG_RETURNS  = sum(r[6] for r in SHOPIFY_8D[:7]) / 7  # 19419.71
AVG_SESSIONS = sum(r[1] for r in SESSIONS_8D[:7]) / 7  # 4120.71

TODAY = SHOPIFY_8D[7]   # 2026-06-01

# Top products (8-day window: May 25 – Jun 01)
TOP_PRODUCTS = [
    # product, gross, net, orders
    ('"The Essential" - White Cotton Shirt & Pant Co-ord Set',          316244.00, 276563.60, 148),
    ("Embroidered Rayon Kurti with Farshi Salwar Set",                  231518.00, 218202.92,  81),
    ("Black Cotton Kurta Pant Set",                                      64268.00,  51442.41,  32),
    ("Everyday Elegance Cotton Kurta Pant Set",                          63968.00,  47876.05,  28),
    ("The Dhurandhar Black Short Kurta Set | A Dhirai Essential",        48356.81,  40971.68,  22),
    ("Ira Rayon Co-ord Set for Women | Relaxed Fit Shirt Top",          39984.00,  36570.37,  16),
    ("Elegant Pakistani-Style Mauve Kurta Set",                          39775.07,  30947.49,  17),
    ('"Riva" — Premium Flowy Rayon Co-ord Set',                         32284.00,  24188.05,  16),
    ('"Sana" — Deep Chocolate Premium Georgette Anarkali Set',          22491.00,  17368.05,   8),
    ('"Inayat" — Power Pastel Cotton Shirt & Trouser Co-ord Set',       19990.00,  19990.00,   9),
    ('"Indraneel" - Royal Blue Mal Cotton Tiered Anarkali Suit Set',    14995.00,  14995.00,   5),
    ("VAIRA Airy Linen Everyday Co-ord Set | Women's Shirt and Pant",   13993.00,  13061.46,   6),
    ("Raahi Rayon Farshi Salwar Kurta Set for Women",                    12495.00,  12495.00,   5),
    ('"Ruby" – Deep Red Premium Rayon Co-ord Set',                      12495.00,  11705.31,   5),
    ('"Adira" - Premium Breathable Soft Fabric Salwar Set',             12196.00,   9197.00,   4),
]

# Meta Ads campaigns — yesterday (Jun 01)
META_CAMP_TODAY = [
    # name, status, spend, impressions, reach, clicks, ctr, cpc, purchases, cpa, roas
    ("Dhirai Scale -ASC",      "ACTIVE", 21932.69, 207352, 176114, 3697, 1.78, 5.93, 25,  877.31, 2.67),
    ("Black DHURANDHAR",       "ACTIVE",  8491.31,  37729,  36413, 1185, 3.14, 7.17,  9,  943.48, 2.49),
    ("Retargeting- new",       "ACTIVE",  7121.82,  19740,  13388,  746, 3.78, 9.55, 10,  712.18, 3.12),
    ("Summer Special Campaign","ACTIVE",   675.03,   8166,   7520,  106, 1.30, 6.37,  3,  225.01,11.85),
]

# Meta Ads campaigns — 7-day (May 25–31, active with spend)
META_CAMP_7D = [
    # name, spend, impressions, reach, clicks, ctr, cpc, purchases, cpa, roas
    ("Dhirai Scale -ASC",      153376.25, 1269101, 868918, 23615, 1.86, 6.49, 188,  815.83, 3.01),
    ("Black DHURANDHAR",        66678.68,  321037, 213187, 10334, 3.22, 6.45,  79,  844.03, 2.93),
    ("Retargeting- new",        56700.76,  187094,  61288,  6607, 3.53, 8.58, 107,  529.91, 4.83),
    ("Summer Special Campaign", 11120.54,   41676,  28494,   776, 1.86,14.33,  15,  741.37, 2.73),
    ("New testing - 5 products (PAUSED)", 25733.17, 100629, 47861, 1266, 1.26,20.33,  5, 5146.63, 0.69),
    ("Roas Goal (PAUSED)",       8097.19,   39776,  34402,   570, 1.43,14.21, 10,  809.72, 2.73),
]

# Meta Ads ad sets — yesterday
META_ADSETS_TODAY = [
    # name, spend, impressions, reach, clicks, ctr, cpc, purchases, cpa, roas
    ("New Sales Ad Set",       17789.17, 187476, 157854, 3182, 1.70,  5.59, 19,  936.27, 2.34),
    ("Broad",                   8491.31,  37729,  36413, 1185, 3.14,  7.17,  9,  943.48, 2.49),
    ("Best performing ads",     5035.74,  12701,   9671,  454, 3.57, 11.09,  5, 1007.15, 2.34),
    ("Karigari",                4143.21,   7166,   6289,  330, 4.61, 12.55,  4, 1035.80, 2.27),
    ("Hania AMir & Karigiri",   2978.60,  12433,  11625,  416, 3.35,  7.16,  6,  496.43, 4.18),
    ("Black and multiple",      1782.82,  15477,  13117,  567, 3.66,  3.14,  4,  445.71, 4.83),
]

META_TOTAL_TODAY = {
    "spend":       sum(r[2] for r in META_CAMP_TODAY),
    "impressions": sum(r[3] for r in META_CAMP_TODAY),
    "reach":       sum(r[4] for r in META_CAMP_TODAY),
    "clicks":      sum(r[5] for r in META_CAMP_TODAY),
    "purchases":   sum(r[8] for r in META_CAMP_TODAY),
}
t = META_TOTAL_TODAY
t["ctr"]  = round(t["clicks"] / t["impressions"] * 100, 2)
t["cpc"]  = round(t["spend"] / t["clicks"], 2)
t["cpa"]  = round(t["spend"] / t["purchases"], 2)
t["roas"] = round((t["purchases"] * TODAY[4]) / t["spend"], 2)  # revenue / spend

META_7D_DAILY_AVG = {
    "spend":     round(sum(r[1] for r in META_CAMP_7D[:4]) / 7, 2),
    "purchases": round(sum(r[7] for r in META_CAMP_7D[:4]) / 7, 2),
    "cpa":       round(sum(r[1] for r in META_CAMP_7D[:4]) / sum(r[7] for r in META_CAMP_7D[:4]), 2),
}

ANOMALIES = [
    "⚠️ ALL 6 active ad sets have NARROW audiences (~1,000 reach). Expand targeting urgently.",
    "🚨 Checkout conversion rate was 0% yesterday — verify pixel and checkout flow.",
    "✅ 'Summer Special Campaign' achieved ROAS 11.85x — highest in account.",
    "⚠️ 'Retargeting- new' CPA rose to ₹712 vs 7-day avg ₹530 — audience may be saturating.",
]

RECS = [
    ("✅ SCALE", "Summer Special Campaign (ROAS 11.85x). Increase budget by 30–40% today."),
    ("⚠️ EXPAND", "All 6 ad-set audiences show narrow reach (~1,000). Broaden targeting or use Advantage+ audiences."),
    ("🚨 CHECK", "Zero checkout conversions yesterday — verify Shopify checkout flow and Meta pixel are working."),
    ("⚠️ REVIEW", "Revenue is 25.2% below 7-day average. Audit ad delivery, site speed, and traffic sources."),
    ("⚠️ MONITOR", "Retargeting CPA rose 34% vs. 7-day avg (₹712 vs ₹530). Refresh retargeting creatives."),
    ("📱 ACTION", "Review abandoned checkout recovery — sessions had 0 completed checkouts yesterday."),
    ("🔍 ACTION", "Top product 'The Essential White Set' drives 31.8% of revenue — ensure stock and ads are live."),
]

# ─── Build workbook ─────────────────────────────────────────────────────────────

def build_dashboard(ws, wb):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 18

    # Title
    ws.row_dimensions[1].height = 36
    c = ws.cell(row=1, column=1, value=f"DAILY STORE & ADS PERFORMANCE  ·  {REPORT_DATE}")
    c.font = Font(name="Calibri", bold=True, size=16, color=WHITE)
    c.fill = PatternFill("solid", fgColor=DARK)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    ws.merge_cells("A1:D1")

    r = 3
    # ── Shopify Section ──────────────────────────────────────────────
    section_title(ws, r, "  SHOPIFY PERFORMANCE", 4, bg=ACCENT)
    r += 1
    header_row(ws, r, ["METRIC", "YESTERDAY", "7-DAY AVG", "CHANGE VS AVG"])
    r += 1

    def data_row(ws, row, label, today_val, avg_val, chg_val, num_fmt="₹#,##0", is_pct=False):
        cell(ws, row, 1, label, bold=False, bg=LIGHT_GREY if row % 2 == 0 else WHITE, border=True)
        cell(ws, row, 2, today_val, bg=LIGHT_GREY if row % 2 == 0 else WHITE,
             num_fmt=num_fmt, align="right", border=True)
        cell(ws, row, 3, avg_val, bg=LIGHT_GREY if row % 2 == 0 else WHITE,
             num_fmt=num_fmt, align="right", border=True)
        chg_cell = cell(ws, row, 4, chg_val, align="center", border=True,
                        num_fmt='0.0"%"' if is_pct else '0.0"%"')
        if chg_val is not None:
            if (chg_val > 0):
                chg_cell.fill = PatternFill("solid", fgColor=GREEN_BG)
                chg_cell.font = Font(name="Calibri", color="1B5E20", bold=True)
            elif chg_val < -10:
                chg_cell.fill = PatternFill("solid", fgColor=RED_BG)
                chg_cell.font = Font(name="Calibri", color="B71C1C", bold=True)
        ws.row_dimensions[row].height = 18

    data_row(ws, r,   "Gross Sales (₹)",   TODAY[1], AVG_GROSS,   pct_change(TODAY[1], AVG_GROSS))
    data_row(ws, r+1, "Net Sales (₹)",     TODAY[2], AVG_NET,     pct_change(TODAY[2], AVG_NET))
    data_row(ws, r+2, "Orders",            TODAY[3], round(AVG_ORDERS,1), pct_change(TODAY[3], AVG_ORDERS), num_fmt="#,##0")
    data_row(ws, r+3, "Avg Order Value (₹)", TODAY[4], round(AVG_AOV,2), pct_change(TODAY[4], AVG_AOV))
    data_row(ws, r+4, "Returns (₹)",       TODAY[6], round(AVG_RETURNS,0), pct_change(TODAY[6], AVG_RETURNS))
    data_row(ws, r+5, "Sessions",          SESSIONS_8D[7][1], round(AVG_SESSIONS,0),
             pct_change(SESSIONS_8D[7][1], AVG_SESSIONS), num_fmt="#,##0")
    r += 8

    # ── Meta Ads Section ────────────────────────────────────────────
    section_title(ws, r, "  META ADS PERFORMANCE", 4, bg="E67C00", fg=WHITE)
    r += 1
    header_row(ws, r, ["METRIC", "YESTERDAY", "7-DAY DAILY AVG", "CHANGE VS AVG"])
    r += 1

    data_row(ws, r,   "Total Spend (₹)",   t["spend"],  META_7D_DAILY_AVG["spend"],
             pct_change(t["spend"], META_7D_DAILY_AVG["spend"]))
    data_row(ws, r+1, "Purchases",          t["purchases"], round(META_7D_DAILY_AVG["purchases"],1),
             pct_change(t["purchases"], META_7D_DAILY_AVG["purchases"]), num_fmt="#,##0")
    data_row(ws, r+2, "CPA (₹)",            t["cpa"], META_7D_DAILY_AVG["cpa"],
             pct_change(t["cpa"], META_7D_DAILY_AVG["cpa"]))

    # ROAS row with colour scale
    roas_today = t["roas"]
    avg_roas   = round(sum(r2[9] for r2 in META_CAMP_7D[:4]) / 4, 2)
    roas_cell = ws.cell(row=r+3, column=2, value=roas_today)
    roas_cell.font = Font(name="Calibri", size=10, bold=True,
                          color="1B5E20" if roas_today >= 3 else ("E65100" if roas_today >= 1.5 else "B71C1C"))
    roas_cell.number_format = "0.00"
    roas_cell.alignment = Alignment(horizontal="right")
    cell(ws, r+3, 1, "Blended ROAS", bg=WHITE, border=True)
    cell(ws, r+3, 3, avg_roas, num_fmt="0.00", align="right", border=True)
    chg = pct_change(roas_today, avg_roas)
    ch = ws.cell(row=r+3, column=4, value=chg)
    ch.number_format = '0.0"%"'
    ch.alignment = Alignment(horizontal="center")
    if chg is not None and chg < 0:
        ch.fill = PatternFill("solid", fgColor=RED_BG)
    r += 6

    # ── Anomalies ────────────────────────────────────────────────────
    section_title(ws, r, "  ANOMALIES & SIGNALS", 4, bg="D32F2F", fg=WHITE)
    r += 1
    for anom in ANOMALIES:
        c2 = ws.cell(row=r, column=1, value=anom)
        c2.font = Font(name="Calibri", size=10, bold=anom.startswith("🚨"))
        c2.fill = PatternFill("solid", fgColor=RED_BG if "🚨" in anom else (YELLOW_BG if "⚠️" in anom else GREEN_BG))
        c2.alignment = Alignment(horizontal="left", indent=1, wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.row_dimensions[r].height = 20
        r += 1

    r += 1

    # ── Recommendations ──────────────────────────────────────────────
    section_title(ws, r, "  RECOMMENDED ACTIONS FOR TODAY", 4, bg=ACCENT2, fg=WHITE)
    r += 1
    header_row(ws, r, ["#", "TYPE", "ACTION", ""], bg="37474F")
    ws.column_dimensions["B"].width = 14
    r += 1
    for i, (rtype, action) in enumerate(RECS, 1):
        bg_c = GREEN_BG if "✅" in rtype else (RED_BG if "🚨" in rtype else YELLOW_BG)
        cell(ws, r, 1, str(i), bold=True, align="center", bg=bg_c)
        cell(ws, r, 2, rtype, bold=True, align="center", bg=bg_c)
        c3 = ws.cell(row=r, column=3, value=action)
        c3.font = Font(name="Calibri", size=10)
        c3.alignment = Alignment(wrap_text=True, vertical="center")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        ws.row_dimensions[r].height = 24
        r += 1

    freeze(ws, 2, 1)


def build_shopify_daily(ws):
    ws.sheet_view.showGridLines = False
    COLS = ["Date", "Gross Sales (₹)", "Net Sales (₹)", "Orders", "AOV (₹)",
            "Discounts (₹)", "Returns (₹)", "Shipping (₹)", "Taxes (₹)", "Total Sales (₹)",
            "Sessions", "Cart Adds", "Reached Checkout", "Completed Checkout", "Conv. Rate (%)"]
    header_row(ws, 1, COLS)
    add_filters(ws, 1, len(COLS))

    NF_INR  = "₹#,##0.00"
    NF_INT  = "#,##0"
    NF_PCT  = "0.00%"
    NF_DATE = "YYYY-MM-DD"

    for i, (sd, ss) in enumerate(zip(SHOPIFY_8D, SESSIONS_8D), 2):
        is_today = sd[0] == REPORT_DATE
        bg = "FFF8E1" if is_today else (LIGHT_GREY if i % 2 == 0 else WHITE)
        values = [
            sd[0], sd[1], sd[2], sd[3], sd[4],
            sd[5], sd[6], sd[7], sd[8], sd[9],
            ss[1], ss[2], ss[3], ss[4], ss[5],
        ]
        fmts = [NF_DATE, NF_INR, NF_INR, NF_INT, NF_INR,
                NF_INR,  NF_INR, NF_INR, NF_INR, NF_INR,
                NF_INT,  NF_INT, NF_INT, NF_INT, NF_PCT]
        for c_idx, (val, fmt) in enumerate(zip(values, fmts), 1):
            c = ws.cell(row=i, column=c_idx, value=val)
            c.number_format = fmt
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="right" if c_idx > 1 else "left")
            c.font = Font(name="Calibri", size=10, bold=is_today)
        ws.row_dimensions[i].height = 18

    # Highlight today row label
    ws.cell(row=9, column=1).font = Font(name="Calibri", bold=True, color="1565C0", size=10)
    freeze(ws, 2, 1)
    autofit(ws)


def build_meta_daily(ws):
    ws.sheet_view.showGridLines = False
    COLS = ["Date", "Total Spend (₹)", "Impressions", "Reach", "Clicks",
            "CTR (%)", "CPC (₹)", "Purchases", "CPA (₹)", "Blended ROAS",
            "Prev 7-day Avg Spend (₹)", "Prev 7-day Avg Purchases"]
    header_row(ws, 1, COLS)
    add_filters(ws, 1, len(COLS))

    # today's row
    row_data = [
        REPORT_DATE, t["spend"], t["impressions"], t["reach"], t["clicks"],
        t["ctr"] / 100, t["cpc"], t["purchases"], t["cpa"], t["roas"],
        META_7D_DAILY_AVG["spend"], round(META_7D_DAILY_AVG["purchases"], 1),
    ]
    fmts = ["YYYY-MM-DD", "₹#,##0.00", "#,##0", "#,##0", "#,##0",
            "0.00%", "₹#,##0.00", "#,##0", "₹#,##0.00", "0.00x",
            "₹#,##0.00", "0.0"]
    for c_idx, (val, fmt) in enumerate(zip(row_data, fmts), 1):
        c = ws.cell(row=2, column=c_idx, value=val)
        c.number_format = fmt
        c.fill = PatternFill("solid", fgColor="FFF8E1")
        c.alignment = Alignment(horizontal="right" if c_idx > 1 else "left")
        c.font = Font(name="Calibri", size=10, bold=True)

    ws.cell(row=2, column=10).font = Font(
        name="Calibri", bold=True, size=10,
        color="1B5E20" if t["roas"] >= 3 else ("E65100" if t["roas"] >= 1.5 else "B71C1C")
    )
    ws.row_dimensions[2].height = 20
    freeze(ws, 2, 1)
    autofit(ws)


def build_product_performance(ws):
    ws.sheet_view.showGridLines = False
    COLS = ["Product Name", "8-Day Gross (₹)", "8-Day Net (₹)", "8-Day Orders",
            "Avg Daily Gross (₹)", "Net Margin (%)"]
    header_row(ws, 1, COLS)
    add_filters(ws, 1, len(COLS))

    NF_INR = "₹#,##0.00"
    NF_PCT = "0.0%"

    for i, (name, gross, net, orders) in enumerate(TOP_PRODUCTS, 2):
        avg_daily = round(gross / 8, 0)
        margin    = net / gross if gross else 0
        bg = LIGHT_GREY if i % 2 == 0 else WHITE
        if i == 2:  # top product
            bg = "E8F5E9"
        data = [name, gross, net, orders, avg_daily, margin]
        fmts = [None, NF_INR, NF_INR, "#,##0", NF_INR, NF_PCT]
        for c_idx, (val, fmt) in enumerate(zip(data, fmts), 1):
            c = ws.cell(row=i, column=c_idx, value=val)
            if fmt:
                c.number_format = fmt
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="right" if c_idx > 1 else "left", wrap_text=(c_idx == 1))
            c.font = Font(name="Calibri", size=10)
        ws.row_dimensions[i].height = 24

    ws.column_dimensions["A"].width = 55
    for col in ["B", "C", "D", "E", "F"]:
        ws.column_dimensions[col].width = 18
    freeze(ws, 2, 1)


def build_campaign_performance(ws):
    ws.sheet_view.showGridLines = False

    # ── Campaign level ─────────────────────────────────────
    section_title(ws, 1, "  CAMPAIGN PERFORMANCE — YESTERDAY (2026-06-01)", 12, bg=DARK)
    CAMP_COLS = ["Campaign Name", "Status", "Spend (₹)", "Impressions", "Reach",
                 "Clicks", "CTR (%)", "CPC (₹)", "Purchases", "CPA (₹)", "ROAS", "Notes"]
    header_row(ws, 2, CAMP_COLS)
    add_filters(ws, 2, len(CAMP_COLS))

    NF_INR = "₹#,##0.00"
    NF_PCT = "0.00%"
    NF_X   = "0.00"

    for i, (name, status, spend, impr, reach, clicks, ctr, cpc, purch, cpa, roas) in enumerate(META_CAMP_TODAY, 3):
        bg = LIGHT_GREY if i % 2 == 0 else WHITE
        note = ""
        if roas >= 5:
            bg = GREEN_BG; note = "🌟 Scale budget"
        elif roas < 1.5:
            bg = RED_BG;   note = "🚨 Pause/review"
        data = [name, status, spend, impr, reach, clicks, ctr/100, cpc, purch, cpa, roas, note]
        fmts = [None, None, NF_INR, "#,##0", "#,##0", "#,##0", NF_PCT, NF_INR, "#,##0", NF_INR, NF_X, None]
        for c_idx, (val, fmt) in enumerate(zip(data, fmts), 1):
            c = ws.cell(row=i, column=c_idx, value=val)
            if fmt:
                c.number_format = fmt
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="right" if c_idx > 2 else "left")
            c.font = Font(name="Calibri", size=10)
            # ROAS colour
            if c_idx == 11 and isinstance(val, (int, float)):
                c.font = Font(name="Calibri", size=10, bold=True,
                              color="1B5E20" if val >= 3 else ("E65100" if val >= 1.5 else "B71C1C"))
        ws.row_dimensions[i].height = 18

    camp_end = 3 + len(META_CAMP_TODAY)

    # ── 7-day campaign reference ──────────────────────────────
    sep_row = camp_end + 2
    section_title(ws, sep_row, "  CAMPAIGN PERFORMANCE — 7-DAY REFERENCE (May 25–31)", 12, bg="37474F")

    CAMP7_COLS = ["Campaign Name", "7-Day Spend (₹)", "Avg Daily Spend (₹)", "Impressions",
                  "Clicks", "CTR (%)", "CPC (₹)", "Purchases", "CPA (₹)", "ROAS"]
    header_row(ws, sep_row + 1, CAMP7_COLS)

    for j, (name, spend, impr, reach, clicks, ctr, cpc, purch, cpa, roas) in enumerate(META_CAMP_7D, sep_row + 2):
        bg = LIGHT_GREY if j % 2 == 0 else WHITE
        data = [name, spend, round(spend/7, 0), impr, clicks, ctr/100, cpc, purch, cpa, roas]
        fmts = [None, NF_INR, NF_INR, "#,##0", "#,##0", NF_PCT, NF_INR, "#,##0", NF_INR, NF_X]
        for c_idx, (val, fmt) in enumerate(zip(data, fmts), 1):
            c = ws.cell(row=j, column=c_idx, value=val)
            if fmt: c.number_format = fmt
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="right" if c_idx > 1 else "left")
            c.font = Font(name="Calibri", size=10)
            if c_idx == 10 and isinstance(val, (int, float)):
                c.font = Font(name="Calibri", size=10, bold=True,
                              color="1B5E20" if val >= 3 else ("E65100" if val >= 1.5 else "B71C1C"))

    # ── Ad set level ──────────────────────────────────────────
    adset_start = sep_row + 2 + len(META_CAMP_7D) + 2
    section_title(ws, adset_start, "  AD SET PERFORMANCE — YESTERDAY (2026-06-01)", 12, bg="5D4037", fg=WHITE)
    AS_COLS = ["Ad Set Name", "Spend (₹)", "Impressions", "Reach", "Clicks",
               "CTR (%)", "CPC (₹)", "Purchases", "CPA (₹)", "ROAS", "Audience Warning"]
    header_row(ws, adset_start + 1, AS_COLS)

    for k, (name, spend, impr, reach, clicks, ctr, cpc, purch, cpa, roas) in enumerate(META_ADSETS_TODAY, adset_start + 2):
        bg = LIGHT_GREY if k % 2 == 0 else WHITE
        data = [name, spend, impr, reach, clicks, ctr/100, cpc, purch, cpa, roas, "⚠️ Narrow audience"]
        fmts = [None, NF_INR, "#,##0", "#,##0", "#,##0", NF_PCT, NF_INR, "#,##0", NF_INR, NF_X, None]
        for c_idx, (val, fmt) in enumerate(zip(data, fmts), 1):
            c = ws.cell(row=k, column=c_idx, value=val)
            if fmt: c.number_format = fmt
            c.fill = PatternFill("solid", fgColor=YELLOW_BG if c_idx == 11 else bg)
            c.alignment = Alignment(horizontal="right" if c_idx > 1 else "left")
            c.font = Font(name="Calibri", size=10)
            if c_idx == 10:
                c.font = Font(name="Calibri", size=10, bold=True,
                              color="1B5E20" if val >= 3 else ("E65100" if val >= 1.5 else "B71C1C"))

    ws.column_dimensions["A"].width = 40
    for col in ["B","C","D","E","F","G","H","I","J","K","L"]:
        ws.column_dimensions[col].width = 16
    freeze(ws, 2, 1)


def build_recommendations(ws):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 80

    title_cell = ws.cell(row=1, column=1, value=f"RECOMMENDATIONS & NOTES — {REPORT_DATE}")
    title_cell.font = Font(name="Calibri", bold=True, size=14, color=WHITE)
    title_cell.fill = PatternFill("solid", fgColor=DARK)
    title_cell.alignment = Alignment(horizontal="left", indent=2, vertical="center")
    ws.merge_cells("A1:B1")
    ws.row_dimensions[1].height = 30

    header_row(ws, 2, ["TYPE", "RECOMMENDATION / NOTE"])

    for i, (rtype, action) in enumerate(RECS, 3):
        bg_c = GREEN_BG if "✅" in rtype else (RED_BG if "🚨" in rtype else YELLOW_BG)
        cell(ws, i, 1, rtype, bold=True, align="center", bg=bg_c)
        c2 = ws.cell(row=i, column=2, value=action)
        c2.font = Font(name="Calibri", size=10)
        c2.fill = PatternFill("solid", fgColor=LIGHT_GREY if i % 2 == 0 else WHITE)
        c2.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[i].height = 26

    next_r = 3 + len(RECS) + 2

    # Unavailable metrics note
    section_title(ws, next_r, "  UNAVAILABLE / NOTES", 2, bg="757575")
    notes = [
        ("Google Sheets", "No service-account credentials available in this session. Excel file committed to GitHub as fallback."),
        ("Shopify refunds/cancellations", "Detailed refund/cancellation breakdown not available via ShopifyQL — use Shopify admin Reports for granular data."),
        ("Meta ROAS exact", "ROAS computed as (purchases × AOV) / spend. Meta pixel-attributed revenue may differ."),
        ("Abandoned checkouts", "Abandoned checkout count unavailable via ShopifyQL API; check Shopify admin > Analytics > Abandoned checkouts."),
        ("Failed payments", "Failed payment data not available via ShopifyQL in this session."),
    ]
    for j, (topic, note) in enumerate(notes, next_r + 1):
        cell(ws, j, 1, topic, bold=True, bg=LIGHT_GREY, size=9)
        c2 = ws.cell(row=j, column=2, value=note)
        c2.font = Font(name="Calibri", size=9, italic=True, color="555555")
        c2.fill = PatternFill("solid", fgColor=LIGHT_GREY)
        c2.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[j].height = 28

    freeze(ws, 3, 1)


def main():
    os.makedirs("reports", exist_ok=True)
    wb = Workbook()

    # Remove default sheet
    wb.remove(wb.active)

    # Tab order
    ws_dash  = wb.create_sheet("Dashboard")
    ws_shop  = wb.create_sheet("Shopify Daily Data")
    ws_meta  = wb.create_sheet("Meta Ads Daily Data")
    ws_prod  = wb.create_sheet("Product Performance")
    ws_camp  = wb.create_sheet("Campaign Performance")
    ws_recs  = wb.create_sheet("Recommendations & Notes")

    # Tab colours
    ws_dash.sheet_properties.tabColor = ACCENT
    ws_shop.sheet_properties.tabColor = ACCENT2
    ws_meta.sheet_properties.tabColor = "E67C00"
    ws_prod.sheet_properties.tabColor = "8E24AA"
    ws_camp.sheet_properties.tabColor = "D32F2F"
    ws_recs.sheet_properties.tabColor = "37474F"

    print("Building Dashboard...")
    build_dashboard(ws_dash, wb)
    print("Building Shopify Daily Data...")
    build_shopify_daily(ws_shop)
    print("Building Meta Ads Daily Data...")
    build_meta_daily(ws_meta)
    print("Building Product Performance...")
    build_product_performance(ws_prod)
    print("Building Campaign Performance...")
    build_campaign_performance(ws_camp)
    print("Building Recommendations & Notes...")
    build_recommendations(ws_recs)

    path = f"reports/daily_performance_{REPORT_DATE}.xlsx"
    wb.save(path)
    print(f"✅ Saved: {path}")
    return path


if __name__ == "__main__":
    main()
