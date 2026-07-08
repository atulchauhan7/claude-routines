#!/usr/bin/env python3
"""
Build "Daily Store & Ads Performance Sheet.xlsx" from data/metrics.json.

Google Sheets is not available in this environment (no connector), so this
produces the best available spreadsheet file instead, per the fallback
instructions. Re-run after merging a new day's data into metrics.json;
historical rows are never overwritten since dates are dict keys.
"""

import json
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.utils import get_column_letter

DATA_PATH = "data/metrics.json"
OUT_PATH = "Daily Store & Ads Performance Sheet.xlsx"

CUR = '₹#,##0.00'
PCT = '0.00"%"'
INT = '#,##0'

HEADER_FILL = PatternFill("solid", fgColor="1F1F1F")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="2E5AA8")
SECTION_FONT = Font(bold=True, color="FFFFFF", size=12)
SUBHEAD_FILL = PatternFill("solid", fgColor="F2F2F2")
SUBHEAD_FONT = Font(bold=True)
TITLE_FONT = Font(bold=True, size=15, color="FFFFFF")
TITLE_FILL = PatternFill("solid", fgColor="2E5AA8")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def load_data():
    with open(DATA_PATH) as f:
        return json.load(f)


def style_header(ws, row, ncols, fill=HEADER_FILL, font=HEADER_FONT):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(vertical="center")


def autofit(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def pct_change(cur, prior):
    if prior in (None, 0):
        return None
    return round((cur - prior) / prior * 100, 2)


def avg(vals):
    vals = [v for v in vals if v is not None]
    return round(sum(vals) / len(vals), 2) if vals else None


# ─────────────────────── Shopify Daily Data ───────────────────────

def build_shopify_daily(wb, data):
    ws = wb.create_sheet("Shopify Daily Data")
    headers = ["Date", "Gross Sales", "Net Sales", "Orders", "AOV",
               "Discounts", "Refunds", "Cancellations",
               "Pending Payment (COD)", "Units Sold"]
    ws.append(headers)
    style_header(ws, 1, len(headers))

    dates = sorted(data["shopify_daily"].keys())
    for d in dates:
        r = data["shopify_daily"][d]
        ws.append([d, r["gross_sales"], r["net_sales"], r["orders"], r["aov"],
                   r["discounts"], r["refunds"], r["cancellations"],
                   r["pending_payment"], r["units_sold"]])

    last_row = len(dates) + 1
    for col in (2, 3, 5, 6, 7):
        for row in range(2, last_row + 1):
            ws.cell(row=row, column=col).number_format = CUR
    for row in range(2, last_row + 1):
        for col in (4, 8, 9, 10):
            ws.cell(row=row, column=col).number_format = INT
        for col in range(1, len(headers) + 1):
            ws.cell(row=row, column=col).border = BORDER

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_row}"
    autofit(ws, [12, 14, 14, 9, 11, 12, 11, 13, 20, 11])

    # Flag cancellations > 0
    ws.conditional_formatting.add(
        f"H2:H{last_row}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=YELLOW_FILL)
    )
    return dates, last_row


# ─────────────────────── Meta Ads Daily Data ───────────────────────

def build_meta_daily(wb, data):
    ws = wb.create_sheet("Meta Ads Daily Data")
    headers = ["Date", "Spend", "Impressions", "Reach", "Clicks", "CTR",
               "CPC", "CPM", "Purchases", "CPA", "ROAS"]
    ws.append(headers)
    style_header(ws, 1, len(headers))

    dates = sorted(data["meta_daily"].keys())
    for d in dates:
        r = data["meta_daily"][d]
        ws.append([d, r["spend"], r["impressions"], r["reach"], r["clicks"],
                   r["ctr"], r["cpc"], r["cpm"], r["purchases"], r["cpa"], r["roas"]])

    last_row = len(dates) + 1
    for row in range(2, last_row + 1):
        for col in (2, 7, 8, 10):
            ws.cell(row=row, column=col).number_format = CUR
        ws.cell(row=row, column=6).number_format = PCT
        for col in (3, 4, 5, 9):
            ws.cell(row=row, column=col).number_format = INT
        ws.cell(row=row, column=11).number_format = '0.00"x"'
        for col in range(1, len(headers) + 1):
            ws.cell(row=row, column=col).border = BORDER

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_row}"
    autofit(ws, [12, 13, 13, 11, 10, 9, 10, 10, 11, 11, 9])

    ws.conditional_formatting.add(
        f"K2:K{last_row}",
        ColorScaleRule(start_type="min", start_color="FFC7CE",
                        mid_type="num", mid_value=2, mid_color="FFEB9C",
                        end_type="max", end_color="C6EFCE")
    )
    return dates, last_row


# ─────────────────────── Product Performance ───────────────────────

def build_products(wb, data):
    ws = wb.create_sheet("Product Performance")
    headers = ["Date", "Product", "Units Sold", "Revenue"]
    ws.append(headers)
    style_header(ws, 1, len(headers))

    row = 2
    latest_date = max(data["products"].keys())
    latest_start_row = None
    for d in sorted(data["products"].keys()):
        if d == latest_date:
            latest_start_row = row
        for p in data["products"][d]:
            ws.append([d, p["title"], p["units"], p["revenue"]])
            row += 1

    last_row = row - 1
    for r in range(2, last_row + 1):
        ws.cell(row=r, column=3).number_format = INT
        ws.cell(row=r, column=4).number_format = CUR
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BORDER

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:D{last_row}"
    autofit(ws, [12, 55, 12, 13])
    return latest_date, latest_start_row, last_row


# ─────────────────────── Campaign Performance ───────────────────────

def build_campaigns(wb, data):
    ws = wb.create_sheet("Campaign Performance")
    headers = ["Date", "Level", "Name", "Status", "Spend", "Impressions",
               "Reach", "Clicks", "CTR", "CPC", "CPM", "Purchases", "CPA", "ROAS"]
    ws.append(headers)
    style_header(ws, 1, len(headers))

    row = 2
    latest_date = max(data["campaigns"].keys())
    latest_camp_rows = []
    for d in sorted(data["campaigns"].keys()):
        camps = sorted(data["campaigns"][d], key=lambda c: c["spend"], reverse=True)
        for c in camps:
            ws.append([d, "Campaign", c["name"], c["status"], c["spend"],
                       c["impressions"], c["reach"], c["clicks"], c["ctr"],
                       c["cpc"], c["cpm"], c["purchases"], c["cpa"], c["roas"]])
            if d == latest_date:
                latest_camp_rows.append(row)
            row += 1
        adsets = sorted(data.get("adsets", {}).get(d, []), key=lambda c: c["spend"], reverse=True)
        for a in adsets:
            ws.append([d, "Ad Set", a["name"], a["status"], a["spend"],
                       a["impressions"], a["reach"], a["clicks"], a["ctr"],
                       a["cpc"], a["cpm"], a["purchases"], a["cpa"], a["roas"]])
            row += 1

    last_row = row - 1
    for r in range(2, last_row + 1):
        ws.cell(row=r, column=5).number_format = CUR
        for c in (6, 7, 8, 12):
            ws.cell(row=r, column=c).number_format = INT
        ws.cell(row=r, column=9).number_format = PCT
        for c in (10, 11, 13):
            ws.cell(row=r, column=c).number_format = CUR
        ws.cell(row=r, column=14).number_format = '0.00"x"'
        for c in range(1, 15):
            ws.cell(row=r, column=c).border = BORDER

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_row}"
    autofit(ws, [12, 10, 30, 10, 12, 12, 11, 9, 8, 9, 10, 11, 10, 8])

    # ROAS conditional formatting: <1.5 red, 1.5-3 yellow, >3 green
    rng = f"N2:N{last_row}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["1.5"], fill=RED_FILL))
    ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["1.5", "3"], fill=YELLOW_FILL))
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["3"], fill=GREEN_FILL))
    return latest_date, latest_camp_rows, last_row


# ─────────────────────── Recommendations engine ───────────────────────

def generate_recommendations(data, latest_date, shopify_avg, meta_avg):
    recs, wins, issues = [], [], []
    s = data["shopify_daily"][latest_date]
    m = data["meta_daily"][latest_date]

    gs_chg = pct_change(s["gross_sales"], shopify_avg["gross_sales"])
    if gs_chg is not None:
        if gs_chg <= -15:
            t = f"Gross sales down {abs(gs_chg):.1f}% vs 7-day avg — check ad delivery, site issues, or seasonal demand."
            recs.append("warn:" + t); issues.append(t)
        elif gs_chg >= 15:
            t = f"Gross sales up {gs_chg:.1f}% vs 7-day avg — identify what drove it and reinforce (scale winning campaigns, restock fast movers)."
            recs.append("win:" + t); wins.append(t)

    orders_chg = pct_change(s["orders"], shopify_avg["orders"])
    if orders_chg is not None and orders_chg <= -15:
        t = f"Orders down {abs(orders_chg):.1f}% vs 7-day avg — review traffic sources and ad delivery for drop-off."
        recs.append("warn:" + t); issues.append(t)

    pend_rate = s["pending_payment"] / s["orders"] if s["orders"] else 0
    avg_pend_rate = shopify_avg.get("pending_payment_rate")
    if avg_pend_rate and pend_rate - avg_pend_rate > 0.10:
        t = f"Pending-payment/COD order share rose to {pend_rate*100:.0f}% (vs ~{avg_pend_rate*100:.0f}% avg) — confirm COD verification flow isn't breaking and follow up on unconfirmed orders."
        recs.append("warn:" + t); issues.append(t)

    if s["cancellations"] > 0:
        recs.append(f"note:{s['cancellations']} order(s) cancelled yesterday — review cancellation reasons if the count keeps climbing.")

    for item in data.get("inventory_issues", {}).get(latest_date, []):
        t = f"'{item['title']}' is ACTIVE with 0 units in stock — restock or unpublish to stop wasting traffic/ad spend on a dead listing."
        recs.append("warn:" + t); issues.append(t)

    roas_chg = pct_change(m["roas"], meta_avg["roas"]) if meta_avg.get("roas") else None
    if roas_chg is not None and roas_chg <= -10:
        t = f"Blended ROAS down {abs(roas_chg):.1f}% vs 7-day avg ({m['roas']:.2f}x vs {meta_avg['roas']:.2f}x) — review creative fatigue and audience overlap before increasing budget."
        recs.append("warn:" + t); issues.append(t)
    elif roas_chg is not None and roas_chg >= 10:
        t = f"Blended ROAS up {roas_chg:.1f}% vs 7-day avg — good window to scale the top campaigns below."
        recs.append("win:" + t); wins.append(t)

    for c in sorted(data["campaigns"][latest_date], key=lambda c: c["spend"], reverse=True):
        if c["roas"] is not None and c["spend"] > 5000:
            if c["roas"] < 1.5:
                t = f"PAUSE or rework campaign '{c['name']}' — ROAS {c['roas']:.2f}x on ₹{c['spend']:,.0f} spend is near/below breakeven."
                recs.append("warn:" + t); issues.append(t)
            elif c["roas"] >= 3.8:
                t = f"Scale campaign '{c['name']}' — ROAS {c['roas']:.2f}x on ₹{c['spend']:,.0f} spend is strong; test a 20% budget increase."
                recs.append("win:" + t); wins.append(t)

    for a in data.get("adsets", {}).get(latest_date, []):
        if a["roas"] is not None and a["spend"] > 1000 and a["roas"] < 1.5:
            t = f"Ad set '{a['name']}' ROAS is only {a['roas']:.2f}x on ₹{a['spend']:,.0f} — pause or refresh creative/audience."
            recs.append("warn:" + t); issues.append(t)

    for note in data.get("anomalies", {}).get(latest_date, []):
        recs.append("note:Meta signal: " + note)

    opp = data.get("opportunity_score", {}).get(latest_date)
    if opp:
        for r in opp["recommendations"]:
            recs.append("note:" + r)

    recs.append("note:Review Shopify abandoned checkouts and confirm recovery email flow is active (not directly queryable via current permissions).")
    recs.append("note:Spot-check top product pages for images/sizing/description accuracy.")

    return recs, wins, issues


# ─────────────────────── Recommendations & Notes tab ───────────────────────

def build_recs_tab(wb, data, latest_date, recs):
    ws = wb.create_sheet("Recommendations & Notes")
    ws.merge_cells("A1:B1")
    ws["A1"] = f"RECOMMENDATIONS & NOTES — {latest_date}"
    ws["A1"].fill = TITLE_FILL
    ws["A1"].font = TITLE_FONT
    ws.row_dimensions[1].height = 24

    ws.append(["", ""])
    ws.append(["Type", "Recommendation"])
    style_header(ws, 3, 2, fill=SUBHEAD_FILL, font=SUBHEAD_FONT)

    icon_map = {"warn": "Issue", "win": "Win", "note": "Action/Note"}
    row = 4
    for r in recs:
        kind, _, text = r.partition(":")
        ws.cell(row=row, column=1, value=icon_map.get(kind, "Note"))
        ws.cell(row=row, column=2, value=text)
        if kind == "warn":
            ws.cell(row=row, column=1).fill = RED_FILL
        elif kind == "win":
            ws.cell(row=row, column=1).fill = GREEN_FILL
        row += 1

    unavail = data.get("unavailable_metrics", {}).get(latest_date, [])
    if unavail:
        row += 1
        ws.cell(row=row, column=1, value="UNAVAILABLE METRICS").font = SUBHEAD_FONT
        ws.cell(row=row, column=1).fill = SUBHEAD_FILL
        ws.cell(row=row, column=2).fill = SUBHEAD_FILL
        row += 1
        for m in unavail:
            ws.cell(row=row, column=1, value="Note")
            ws.cell(row=row, column=2, value=m)
            row += 1

    for r in range(4, row):
        ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30

    ws.freeze_panes = "A4"
    autofit(ws, [14, 110])
    return row - 1


# ─────────────────────── Dashboard tab ───────────────────────

def build_dashboard(wb, data, latest_date, shop_dates, shop_last_row,
                     meta_dates, meta_last_row, wins, issues, recs,
                     prod_latest_start, prod_last_row, camp_rows):
    ws = wb.create_sheet("Dashboard", 0)

    ws.merge_cells("A1:E1")
    ws["A1"] = f"DAILY STORE & ADS PERFORMANCE — {latest_date}"
    ws["A1"].fill = TITLE_FILL
    ws["A1"].font = TITLE_FONT
    ws.row_dimensions[1].height = 26

    r = 3
    ws.cell(row=r, column=1, value="EXECUTIVE SUMMARY").font = SECTION_FONT
    for c in range(1, 6):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    s_row = shop_last_row  # yesterday's row in Shopify Daily Data
    s_avg_start, s_avg_end = shop_last_row - 7, shop_last_row - 1
    m_row = meta_last_row
    m_avg_start, m_avg_end = meta_last_row - 7, meta_last_row - 1

    ws.cell(row=r, column=1, value=(
        f"Yesterday: {data['shopify_daily'][latest_date]['orders']} Shopify orders, "
        f"₹{data['shopify_daily'][latest_date]['gross_sales']:,.0f} gross sales, "
        f"₹{data['meta_daily'][latest_date]['spend']:,.0f} Meta spend at "
        f"{data['meta_daily'][latest_date]['roas']:.2f}x blended ROAS. "
        f"{len(issues)} issue(s) flagged, {len(wins)} win(s)."
    ))
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True)
    ws.row_dimensions[r].height = 30
    r += 2

    def sheet_ref(name, col, row_):
        return f"='{name}'!{get_column_letter(col)}{row_}"

    def avg_formula(name, col, start, end):
        return f"=AVERAGE('{name}'!{get_column_letter(col)}{start}:{get_column_letter(col)}{end})"

    # ── Shopify performance ──
    ws.cell(row=r, column=1, value="SHOPIFY PERFORMANCE").font = SECTION_FONT
    for c in range(1, 6):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1
    hdr_row = r
    for c, h in enumerate(["Metric", "Yesterday", "7-Day Avg", "Change vs Avg", ""], start=1):
        ws.cell(row=r, column=c, value=h)
    style_header(ws, hdr_row, 4, fill=SUBHEAD_FILL, font=SUBHEAD_FONT)
    r += 1

    shop_metrics = [
        ("Gross Sales", 2, CUR), ("Net Sales", 3, CUR), ("Orders", 4, INT), ("AOV", 5, CUR),
    ]
    shop_metric_rows = {}
    for label, col, fmt in shop_metrics:
        ws.cell(row=r, column=1, value=label)
        yc = ws.cell(row=r, column=2, value=sheet_ref("Shopify Daily Data", col, s_row))
        ac = ws.cell(row=r, column=3, value=avg_formula("Shopify Daily Data", col, s_avg_start, s_avg_end))
        yc.number_format = fmt
        ac.number_format = fmt
        chg = ws.cell(row=r, column=4, value=f"=IF(C{r}=0,\"\",(B{r}-C{r})/C{r}*100)")
        chg.number_format = PCT
        shop_metric_rows[label] = r
        r += 1

    ws.cell(row=r, column=1, value="Refunds")
    ws.cell(row=r, column=2, value=sheet_ref("Shopify Daily Data", 6, s_row)).number_format = CUR
    ws.cell(row=r, column=1, value="Refunds")
    r += 1
    ws.cell(row=r, column=1, value="Cancellations")
    ws.cell(row=r, column=2, value=sheet_ref("Shopify Daily Data", 7, s_row)).number_format = INT
    chg_col_first, chg_col_last = shop_metric_rows["Gross Sales"], shop_metric_rows["AOV"]
    r += 2

    # ── Meta Ads performance ──
    ws.cell(row=r, column=1, value="META ADS PERFORMANCE").font = SECTION_FONT
    for c in range(1, 6):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1
    hdr_row2 = r
    for c, h in enumerate(["Metric", "Yesterday", "7-Day Avg", "Change vs Avg", ""], start=1):
        ws.cell(row=r, column=c, value=h)
    style_header(ws, hdr_row2, 4, fill=SUBHEAD_FILL, font=SUBHEAD_FONT)
    r += 1

    meta_metrics = [
        ("Spend", 2, CUR), ("Impressions", 3, INT), ("Clicks", 5, INT),
        ("CTR", 6, PCT), ("CPC", 7, CUR), ("ROAS", 11, '0.00"x"'),
    ]
    meta_chg_first = r
    for label, col, fmt in meta_metrics:
        ws.cell(row=r, column=1, value=label)
        yc = ws.cell(row=r, column=2, value=sheet_ref("Meta Ads Daily Data", col, m_row))
        ac = ws.cell(row=r, column=3, value=avg_formula("Meta Ads Daily Data", col, m_avg_start, m_avg_end))
        yc.number_format = fmt
        ac.number_format = fmt
        chg = ws.cell(row=r, column=4, value=f"=IF(C{r}=0,\"\",(B{r}-C{r})/C{r}*100)")
        chg.number_format = PCT
        r += 1
    meta_chg_last = r - 1

    ws.cell(row=r, column=1, value="Purchases (attributed)")
    ws.cell(row=r, column=2, value=sheet_ref("Meta Ads Daily Data", 9, m_row)).number_format = INT
    r += 1
    ws.cell(row=r, column=1, value="CPA")
    ws.cell(row=r, column=2, value=sheet_ref("Meta Ads Daily Data", 10, m_row)).number_format = CUR
    r += 2

    # ── Key wins ──
    ws.cell(row=r, column=1, value="KEY WINS").font = SECTION_FONT
    for c in range(1, 6):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1
    if wins:
        for w in wins[:5]:
            ws.cell(row=r, column=1, value="✓")
            ws.cell(row=r, column=1).fill = GREEN_FILL
            ws.cell(row=r, column=2, value=w)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True)
            r += 1
    else:
        ws.cell(row=r, column=2, value="No standout wins vs 7-day average today — steady performance.")
        r += 1
    r += 1

    # ── Key issues ──
    ws.cell(row=r, column=1, value="KEY ISSUES").font = SECTION_FONT
    for c in range(1, 6):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1
    if issues:
        for i in issues[:5]:
            ws.cell(row=r, column=1, value="!")
            ws.cell(row=r, column=1).fill = RED_FILL
            ws.cell(row=r, column=2, value=i)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True)
            r += 1
    else:
        ws.cell(row=r, column=2, value="No critical issues detected.")
        r += 1
    r += 1

    # ── Recommended actions ──
    ws.cell(row=r, column=1, value="RECOMMENDED ACTIONS TODAY").font = SECTION_FONT
    for c in range(1, 6):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1
    action_texts = [x.partition(":")[2] for x in recs][:7]
    for idx, t in enumerate(action_texts, start=1):
        ws.cell(row=r, column=1, value=idx)
        ws.cell(row=r, column=2, value=t)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True)
        ws.row_dimensions[r].height = 28
        r += 1

    # Conditional formatting on Change vs Avg columns
    ws.conditional_formatting.add(
        f"D{shop_metric_rows['Gross Sales']}:D{shop_metric_rows['AOV']}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN_FILL)
    )
    ws.conditional_formatting.add(
        f"D{shop_metric_rows['Gross Sales']}:D{shop_metric_rows['AOV']}",
        CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL)
    )
    ws.conditional_formatting.add(
        f"D{meta_chg_first}:D{meta_chg_last}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN_FILL)
    )
    ws.conditional_formatting.add(
        f"D{meta_chg_first}:D{meta_chg_last}",
        CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL)
    )

    ws.freeze_panes = "A2"
    autofit(ws, [26, 16, 16, 16, 30])

    add_charts(wb, ws, shop_dates, shop_last_row, meta_dates, meta_last_row,
               prod_latest_start, prod_last_row, camp_rows)


def add_charts(wb, dash_ws, shop_dates, shop_last_row, meta_dates, meta_last_row,
                prod_latest_start, prod_last_row, camp_rows):
    sd = wb["Shopify Daily Data"]
    md = wb["Meta Ads Daily Data"]
    pp = wb["Product Performance"]
    cp = wb["Campaign Performance"]

    # Sales trend (Gross + Net)
    chart1 = LineChart()
    chart1.title = "Sales Trend (₹)"
    chart1.height, chart1.width = 8, 15
    data_ref = Reference(sd, min_col=2, max_col=3, min_row=1, max_row=shop_last_row)
    cats = Reference(sd, min_col=1, min_row=2, max_row=shop_last_row)
    chart1.add_data(data_ref, titles_from_data=True)
    chart1.set_categories(cats)
    dash_ws.add_chart(chart1, "G3")

    # Orders trend
    chart2 = BarChart()
    chart2.title = "Orders Trend"
    chart2.height, chart2.width = 8, 15
    data_ref = Reference(sd, min_col=4, max_col=4, min_row=1, max_row=shop_last_row)
    chart2.add_data(data_ref, titles_from_data=True)
    chart2.set_categories(cats)
    dash_ws.add_chart(chart2, "G19")

    # Ad spend trend
    chart3 = LineChart()
    chart3.title = "Meta Ad Spend Trend (₹)"
    chart3.height, chart3.width = 8, 15
    data_ref = Reference(md, min_col=2, max_col=2, min_row=1, max_row=meta_last_row)
    cats2 = Reference(md, min_col=1, min_row=2, max_row=meta_last_row)
    chart3.add_data(data_ref, titles_from_data=True)
    chart3.set_categories(cats2)
    dash_ws.add_chart(chart3, "N3")

    # ROAS trend
    chart4 = LineChart()
    chart4.title = "Blended ROAS Trend"
    chart4.height, chart4.width = 8, 15
    data_ref = Reference(md, min_col=11, max_col=11, min_row=1, max_row=meta_last_row)
    chart4.add_data(data_ref, titles_from_data=True)
    chart4.set_categories(cats2)
    dash_ws.add_chart(chart4, "N19")

    # Top products (latest date)
    chart5 = BarChart()
    chart5.title = "Top Products by Revenue (Yesterday)"
    chart5.height, chart5.width = 8, 15
    chart5.type = "bar"
    data_ref = Reference(pp, min_col=4, max_col=4, min_row=prod_latest_start - 1, max_row=prod_last_row)
    cats3 = Reference(pp, min_col=2, min_row=prod_latest_start, max_row=prod_last_row)
    chart5.add_data(data_ref, titles_from_data=True)
    chart5.set_categories(cats3)
    dash_ws.add_chart(chart5, "G35")

    # Top campaigns (latest date, by spend)
    if camp_rows:
        chart6 = BarChart()
        chart6.title = "Top Campaigns by Spend (Yesterday)"
        chart6.height, chart6.width = 8, 15
        chart6.type = "bar"
        min_r, max_r = min(camp_rows), max(camp_rows)
        data_ref = Reference(cp, min_col=5, max_col=5, min_row=min_r - 1, max_row=max_r)
        cats4 = Reference(cp, min_col=3, min_row=min_r, max_row=max_r)
        chart6.add_data(data_ref, titles_from_data=True)
        chart6.set_categories(cats4)
        dash_ws.add_chart(chart6, "N35")


def main():
    data = load_data()
    wb = Workbook()
    wb.remove(wb.active)

    shop_dates, shop_last_row = build_shopify_daily(wb, data)
    meta_dates, meta_last_row = build_meta_daily(wb, data)
    prod_latest_date, prod_latest_start, prod_last_row = build_products(wb, data)
    camp_latest_date, camp_rows, camp_last_row = build_campaigns(wb, data)

    latest_date = shop_dates[-1]
    s_hist = [data["shopify_daily"][d] for d in shop_dates[-8:-1]]
    shopify_avg = {
        "gross_sales": avg([x["gross_sales"] for x in s_hist]),
        "orders": avg([x["orders"] for x in s_hist]),
        "pending_payment_rate": avg([x["pending_payment"] / x["orders"] for x in s_hist if x["orders"]]),
    }
    m_hist = [data["meta_daily"][d] for d in meta_dates[-8:-1]]
    meta_avg = {"roas": avg([x["roas"] for x in m_hist if x["roas"] is not None])}

    recs, wins, issues = generate_recommendations(data, latest_date, shopify_avg, meta_avg)
    build_recs_tab(wb, data, latest_date, recs)
    build_dashboard(wb, data, latest_date, shop_dates, shop_last_row,
                     meta_dates, meta_last_row, wins, issues, recs,
                     prod_latest_start, prod_last_row, camp_rows)

    wb.save(OUT_PATH)
    print(f"Saved {OUT_PATH}")
    print(f"Latest date: {latest_date}, {len(recs)} recommendations, {len(wins)} wins, {len(issues)} issues")


if __name__ == "__main__":
    main()
