#!/usr/bin/env python3
"""
Builds/updates Daily Store & Ads Performance Sheet.xlsx from the JSON
snapshots in data/*.json (one file per report date, produced by the daily
routine). Re-run any time — it rebuilds deterministically from the JSON
history, de-duplicating by date, so re-running for the same date never
creates duplicate rows.

Usage: python3 build_dashboard.py
Output: Daily Store & Ads Performance Sheet.xlsx (repo root)
"""
import glob
import json
import os

from openpyxl import Workbook
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ROOT, "data")
OUT_PATH = os.path.join(ROOT, "Daily Store & Ads Performance Sheet.xlsx")

INR = '"₹"#,##0.00'
INR0 = '"₹"#,##0'
PCT2 = '0.00"%"'
PCT1_SIGNED = '+0.0"%";-0.0"%";0.0"%"'
ROAS_FMT = '0.00"x"'

# ── palette ──────────────────────────────────────────────────────────────
DARK = "1F1F1F"
WHITE = "FFFFFF"
ACCENT = "2C5CD6"
SECTION = "EFEFEF"
GREEN = "C7E9C0"
RED = "F4C6C6"
YELLOW = "FFEFA6"
ORANGE = "FFCC80"

HEADER_FILL = PatternFill("solid", fgColor=DARK)
ACCENT_FILL = PatternFill("solid", fgColor=ACCENT)
SECTION_FILL = PatternFill("solid", fgColor=SECTION)
HEADER_FONT = Font(bold=True, color=WHITE, size=11)
TITLE_FONT = Font(bold=True, color=WHITE, size=14)
SECTION_FONT = Font(bold=True, color=WHITE, size=11)
BOLD = Font(bold=True)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def load_history():
    shopify_by_date, meta_by_date, latest = {}, {}, None
    files = sorted(glob.glob(os.path.join(DATA_DIR, "*.json")))
    latest_payload = None
    for f in files:
        with open(f) as fh:
            payload = json.load(fh)
        for row in payload.get("shopify_history", []):
            shopify_by_date[row["date"]] = row
        for row in payload.get("meta_history", []):
            meta_by_date[row["date"]] = row
        if latest is None or payload["date"] > latest:
            latest = payload["date"]
            latest_payload = payload
    shopify_rows = [shopify_by_date[d] for d in sorted(shopify_by_date)]
    meta_rows = [meta_by_date[d] for d in sorted(meta_by_date)]
    return shopify_rows, meta_rows, latest_payload


def avg(vals):
    vals = [v for v in vals if v is not None]
    return sum(vals) / len(vals) if vals else 0.0


def pct_change(cur, prior):
    if not prior:
        return None
    return (cur - prior) / prior * 100.0


def style_header(ws, row, ncols, fill=HEADER_FILL, font=HEADER_FONT, height=None):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(vertical="center")
    if height:
        ws.row_dimensions[row].height = height


def autofit(ws, ncols, min_width=10, max_width=48):
    for c in range(1, ncols + 1):
        col_letter = get_column_letter(c)
        best = min_width
        for cell in ws[col_letter]:
            v = cell.value
            if v is None:
                continue
            best = max(best, min(max_width, len(str(v)) + 2))
        ws.column_dimensions[col_letter].width = best


def add_filter(ws, first_row, last_row, first_col, last_col):
    ws.auto_filter.ref = (
        f"{get_column_letter(first_col)}{first_row}:"
        f"{get_column_letter(last_col)}{last_row}"
    )


# ── build ────────────────────────────────────────────────────────────────

def main():
    shopify_rows, meta_rows, latest = load_history()
    date_str = latest["date"]

    # yesterday = last row; prior 7 rows (excluding yesterday) = comparison window
    s_today = shopify_rows[-1]
    s_prior = shopify_rows[-8:-1] if len(shopify_rows) >= 8 else shopify_rows[:-1]
    m_today = meta_rows[-1]
    m_prior = meta_rows[-8:-1] if len(meta_rows) >= 8 else meta_rows[:-1]

    s_avg = {
        k: avg([r[k] for r in s_prior])
        for k in ("gross_sales", "net_sales", "orders", "aov", "total_sales", "sessions")
    }
    m_spend_sum = sum(r["spend"] for r in m_prior) or 1.0
    m_avg = {
        "spend": avg([r["spend"] for r in m_prior]),
        "impressions": avg([r["impressions"] for r in m_prior]),
        "clicks": avg([r["clicks"] for r in m_prior]),
        "ctr": avg([r["ctr"] for r in m_prior]),
        "purchases": avg([r["purchases"] for r in m_prior]),
        "cpa": avg([r["cpa"] for r in m_prior]),
        "roas": (sum(r["spend"] * r["roas"] for r in m_prior) / m_spend_sum) if m_prior else 0.0,
    }

    wb = Workbook()
    wb.remove(wb.active)

    ws_dash = wb.create_sheet("Dashboard")
    ws_shop = wb.create_sheet("Shopify Daily Data")
    ws_meta = wb.create_sheet("Meta Ads Daily Data")
    ws_prod = wb.create_sheet("Product Performance")
    ws_camp = wb.create_sheet("Campaign Performance")
    ws_notes = wb.create_sheet("Recommendations & Notes")

    # ══════════════════════════ Shopify Daily Data ══════════════════════
    headers = ["Date", "Gross Sales", "Net Sales", "Orders", "AOV",
               "Total Sales", "Discounts", "Returns", "Taxes",
               "Sessions", "Cart Adds", "Reached Checkout",
               "Completed Checkout", "Conversion Rate"]
    ws_shop.append(headers)
    style_header(ws_shop, 1, len(headers), height=20)
    for r in shopify_rows:
        ws_shop.append([
            r["date"], r["gross_sales"], r["net_sales"], r["orders"], r["aov"],
            r["total_sales"], r["discounts"], r["returns"], r["taxes"],
            r["sessions"], r["cart_adds"], r["reached_checkout"],
            r["completed_checkout"], r["conversion_rate"] * 100,
        ])
    last_row = len(shopify_rows) + 1
    for col in (2, 3, 5, 6, 7, 8, 9):
        for row in range(2, last_row + 1):
            ws_shop.cell(row=row, column=col).number_format = INR
    for row in range(2, last_row + 1):
        ws_shop.cell(row=row, column=14).number_format = PCT2
    ws_shop.freeze_panes = "A2"
    add_filter(ws_shop, 1, last_row, 1, len(headers))
    autofit(ws_shop, len(headers))

    # ══════════════════════════ Meta Ads Daily Data ══════════════════════
    headers = ["Date", "Spend", "Impressions", "Reach", "Clicks", "CTR",
               "CPC", "CPM", "Purchases", "CPA", "ROAS"]
    ws_meta.append(headers)
    style_header(ws_meta, 1, len(headers), height=20)
    for r in meta_rows:
        ws_meta.append([
            r["date"], r["spend"], r["impressions"], r["reach"], r["clicks"],
            r["ctr"], r["cpc"], r["cpm"], r["purchases"], r["cpa"], r["roas"],
        ])
    last_row_m = len(meta_rows) + 1
    for col in (2, 7, 8, 10):
        for row in range(2, last_row_m + 1):
            ws_meta.cell(row=row, column=col).number_format = INR
    for row in range(2, last_row_m + 1):
        ws_meta.cell(row=row, column=6).number_format = PCT2
        ws_meta.cell(row=row, column=11).number_format = ROAS_FMT
    ws_meta.conditional_formatting.add(
        f"K2:K{last_row_m}",
        CellIsRule(operator="lessThan", formula=["1.5"], fill=PatternFill("solid", fgColor=RED)),
    )
    ws_meta.conditional_formatting.add(
        f"K2:K{last_row_m}",
        CellIsRule(operator="between", formula=["1.5", "3"], fill=PatternFill("solid", fgColor=YELLOW)),
    )
    ws_meta.conditional_formatting.add(
        f"K2:K{last_row_m}",
        CellIsRule(operator="greaterThan", formula=["3"], fill=PatternFill("solid", fgColor=GREEN)),
    )
    ws_meta.freeze_panes = "A2"
    add_filter(ws_meta, 1, last_row_m, 1, len(headers))
    autofit(ws_meta, len(headers))

    # ══════════════════════════ Product Performance ══════════════════════
    ws_prod.append(["YESTERDAY (" + date_str + ")", "", "", "", "",
                     "PREVIOUS 7 DAYS", "", ""])
    ws_prod.merge_cells("A1:E1")
    ws_prod.merge_cells("F1:H1")
    style_header(ws_prod, 1, 8, fill=ACCENT_FILL, font=SECTION_FONT, height=20)
    headers = ["Product", "Gross Sales", "Net Sales", "Orders", "Total Sales",
               "7-Day Gross", "7-Day Net", "7-Day Orders"]
    ws_prod.append(headers)
    style_header(ws_prod, 2, 8, height=18)

    seven_day_lookup = {p["product"]: p for p in latest.get("top_products_7day", [])}
    for p in latest.get("top_products_yesterday", []):
        sd = seven_day_lookup.get(p["product"], {})
        ws_prod.append([
            p["product"], p["gross_sales"], p["net_sales"], p["orders"],
            p["total_sales"], sd.get("gross_sales", ""), sd.get("net_sales", ""),
            sd.get("orders", ""),
        ])
    n = len(latest.get("top_products_yesterday", []))
    last_row_p = n + 2
    for col in (2, 3, 5, 6, 7):
        for row in range(3, last_row_p + 1):
            ws_prod.cell(row=row, column=col).number_format = INR
    ws_prod.freeze_panes = "A3"
    add_filter(ws_prod, 2, last_row_p, 1, 8)
    autofit(ws_prod, 8, max_width=55)
    ws_prod.column_dimensions["A"].width = 55

    # ══════════════════════════ Campaign Performance ══════════════════════
    cheaders = ["Campaign", "Status", "Spend", "Impressions", "Reach", "Clicks",
                "CTR", "CPC", "CPM", "Purchases", "CPA", "ROAS"]
    ws_camp.append(["CAMPAIGN LEVEL — " + date_str] + [""] * (len(cheaders) - 1))
    ws_camp.merge_cells(f"A1:{get_column_letter(len(cheaders))}1")
    style_header(ws_camp, 1, len(cheaders), fill=ACCENT_FILL, font=SECTION_FONT, height=20)
    ws_camp.append(cheaders)
    style_header(ws_camp, 2, len(cheaders), height=18)

    camps = sorted(latest.get("campaigns_yesterday", []), key=lambda c: c["spend"], reverse=True)
    for c in camps:
        ws_camp.append([
            c["campaign"], c["status"], c["spend"], c["impressions"], c["reach"],
            c["clicks"], c["ctr"], c["cpc"], c["cpm"], c["purchases"], c["cpa"], c["roas"],
        ])
    camp_last = len(camps) + 2
    for col in (3, 8, 9, 11):
        for row in range(3, camp_last + 1):
            ws_camp.cell(row=row, column=col).number_format = INR
    for row in range(3, camp_last + 1):
        ws_camp.cell(row=row, column=7).number_format = PCT2
        ws_camp.cell(row=row, column=12).number_format = ROAS_FMT
    ws_camp.conditional_formatting.add(
        f"L3:L{camp_last}",
        CellIsRule(operator="lessThan", formula=["1.5"], fill=PatternFill("solid", fgColor=RED)),
    )
    ws_camp.conditional_formatting.add(
        f"L3:L{camp_last}",
        CellIsRule(operator="between", formula=["1.5", "3"], fill=PatternFill("solid", fgColor=YELLOW)),
    )
    ws_camp.conditional_formatting.add(
        f"L3:L{camp_last}",
        CellIsRule(operator="greaterThan", formula=["3"], fill=PatternFill("solid", fgColor=GREEN)),
    )
    add_filter(ws_camp, 2, camp_last, 1, len(cheaders))

    # Ad set section
    aheaders = ["Ad Set", "Campaign", "Spend", "Impressions", "Reach", "Clicks",
                "CTR", "CPC", "CPM", "Purchases", "CPA", "ROAS"]
    as_start = camp_last + 2
    ws_camp.append([""] * len(cheaders))
    ws_camp.cell(row=as_start, column=1, value="AD SET LEVEL — " + date_str)
    ws_camp.merge_cells(f"A{as_start}:{get_column_letter(len(cheaders))}{as_start}")
    style_header(ws_camp, as_start, len(cheaders), fill=ACCENT_FILL, font=SECTION_FONT, height=20)
    ws_camp.append(aheaders)
    style_header(ws_camp, as_start + 1, len(aheaders), height=18)

    adsets = sorted(latest.get("adsets_yesterday", []), key=lambda a: a["spend"], reverse=True)
    for a in adsets:
        ws_camp.append([
            a["adset"], a["campaign"], a["spend"], a["impressions"], a["reach"],
            a["clicks"], a["ctr"], a["cpc"], a["cpm"], a["purchases"],
            a["cpa"] if a["cpa"] is not None else "",
            a["roas"] if a["roas"] is not None else "",
        ])
    as_end = as_start + 1 + len(adsets)
    for col in (3, 8, 9, 11):
        for row in range(as_start + 2, as_end + 1):
            ws_camp.cell(row=row, column=col).number_format = INR
    for row in range(as_start + 2, as_end + 1):
        ws_camp.cell(row=row, column=7).number_format = PCT2
        ws_camp.cell(row=row, column=12).number_format = ROAS_FMT
    ws_camp.conditional_formatting.add(
        f"L{as_start + 2}:L{as_end}",
        CellIsRule(operator="lessThan", formula=["1.5"], fill=PatternFill("solid", fgColor=RED)),
    )
    ws_camp.conditional_formatting.add(
        f"L{as_start + 2}:L{as_end}",
        CellIsRule(operator="between", formula=["1.5", "3"], fill=PatternFill("solid", fgColor=YELLOW)),
    )
    ws_camp.conditional_formatting.add(
        f"L{as_start + 2}:L{as_end}",
        CellIsRule(operator="greaterThan", formula=["3"], fill=PatternFill("solid", fgColor=GREEN)),
    )
    ws_camp.freeze_panes = "A3"
    autofit(ws_camp, len(cheaders), max_width=32)
    ws_camp.column_dimensions["A"].width = 30
    ws_camp.column_dimensions["B"].width = 26

    # ══════════════════════════ Recommendations & Notes ══════════════════
    ws_notes.append(["RECOMMENDATIONS & NOTES — " + date_str, ""])
    ws_notes.merge_cells("A1:B1")
    style_header(ws_notes, 1, 2, fill=ACCENT_FILL, font=Font(bold=True, color=WHITE, size=13), height=22)
    ws_notes.append(["", ""])
    ws_notes.append(["TYPE", "RECOMMENDATION / NOTE"])
    style_header(ws_notes, 3, 2, height=18)

    gross_chg = pct_change(s_today["gross_sales"], s_avg["gross_sales"])
    orders_chg = pct_change(s_today["orders"], s_avg["orders"])
    roas_chg = pct_change(m_today["roas"], m_avg["roas"])

    wins, issues, actions = [], [], []
    if gross_chg is not None and gross_chg > 10:
        wins.append(f"Gross sales up {gross_chg:.1f}% vs 7-day avg (₹{s_today['gross_sales']:,.0f} vs ₹{s_avg['gross_sales']:,.0f}) — best day of the week.")
    if orders_chg is not None and orders_chg > 10:
        wins.append(f"Orders up {orders_chg:.1f}% vs 7-day avg ({s_today['orders']} vs {s_avg['orders']:.1f}).")
    if roas_chg is not None and roas_chg > 5:
        wins.append(f"Blended Meta ROAS {m_today['roas']:.2f}x, up {roas_chg:.1f}% vs 7-day avg ({m_avg['roas']:.2f}x).")
    best_camp = max(latest.get("campaigns_yesterday", []), key=lambda c: c["roas"], default=None)
    if best_camp:
        wins.append(f"'{best_camp['campaign']}' campaign led with {best_camp['roas']:.2f}x ROAS at ₹{best_camp['cpa']:,.0f} CPA — strong candidate to scale.")
    best_adset = max([a for a in latest.get("adsets_yesterday", []) if a.get("roas")], key=lambda a: a["roas"], default=None)
    if best_adset:
        wins.append(f"'{best_adset['adset']}' ad set hit {best_adset['roas']:.2f}x ROAS, {best_adset['purchases']} purchases at ₹{best_adset['cpa']:,.0f} CPA.")

    for inv in latest.get("inventory_issues", []):
        issues.append(f"'{inv['product']}' is OUT OF STOCK (0 units) but still ACTIVE — {inv.get('note', '')}")
    weak = [a for a in latest.get("adsets_yesterday", []) if a["spend"] > 1000 and (a.get("roas") or 0) < 2.5]
    for a in weak:
        issues.append(f"'{a['adset']}' ad set spent ₹{a['spend']:,.0f} at only {a['roas']:.2f}x ROAS — review creative/targeting or reallocate budget.")
    zero_purchase = [a for a in latest.get("adsets_yesterday", []) if a["spend"] > 300 and not a["purchases"]]
    for a in zero_purchase:
        issues.append(f"'{a['adset']}' ad set spent ₹{a['spend']:,.0f} with zero tracked purchases.")
    if s_today["completed_checkout"] <= 1 and s_today["orders"] > 20:
        issues.append(f"On-site 'completed checkout' sessions ({s_today['completed_checkout']}) are near zero versus {s_today['orders']} actual orders — likely a checkout/thank-you pixel tracking gap, not real abandonment. Worth auditing.")

    for r in latest.get("opportunity_recs", []):
        actions.append(r)
    if weak:
        actions.append(f"Pause or rework '{weak[0]['adset']}' (largest spend at low ROAS) and shift budget toward '{best_camp['campaign']}' and '{best_adset['adset']}'." if best_camp and best_adset else f"Pause or rework '{weak[0]['adset']}'.")
    for inv in latest.get("inventory_issues", []):
        actions.append(f"Restock or unpublish '{inv['product']}' to stop wasting ad traffic on an out-of-stock listing.")
    actions.append("Audit Shopify checkout/thank-you page pixel firing to close the session-tracking gap noted above.")

    if not wins:
        wins.append("Performance in line with 7-day average — continue monitoring.")
    if not issues:
        issues.append("No critical issues detected.")

    def add_section(title, items, icon):
        ws_notes.append(["", ""])
        r = ws_notes.max_row + 1
        ws_notes.append([title, ""])
        ws_notes.merge_cells(f"A{r}:B{r}")
        style_header(ws_notes, r, 2, fill=SECTION_FILL, font=Font(bold=True, color="1F1F1F"), height=18)
        for it in items:
            ws_notes.append([icon, it])

    add_section("KEY WINS", wins, "✅")
    add_section("KEY ISSUES", issues, "⚠️")
    add_section("RECOMMENDED ACTIONS TODAY", actions[:7], "➡️")
    add_section("UNAVAILABLE METRICS / NOTES", latest.get("unavailable_metrics", []), "ℹ️")

    ws_notes.freeze_panes = "A4"
    for row in ws_notes.iter_rows(min_row=1, max_row=ws_notes.max_row, min_col=1, max_col=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws_notes.column_dimensions["A"].width = 14
    ws_notes.column_dimensions["B"].width = 100

    # ══════════════════════════ Dashboard ══════════════════════════════
    ws_dash.append(["DAILY STORE & ADS PERFORMANCE — " + date_str, "", "", ""])
    ws_dash.merge_cells("A1:D1")
    style_header(ws_dash, 1, 4, fill=ACCENT_FILL, font=TITLE_FONT, height=26)
    ws_dash.append(["", "", "", ""])

    def section(title):
        r = ws_dash.max_row + 1
        ws_dash.append([title, "", "", ""])
        ws_dash.merge_cells(f"A{r}:D{r}")
        style_header(ws_dash, r, 4, fill=HEADER_FILL, font=SECTION_FONT, height=20)
        return r

    section("EXECUTIVE SUMMARY")
    exec_lines = [
        f"Yesterday's gross sales were ₹{s_today['gross_sales']:,.0f} across {s_today['orders']} orders "
        f"({'+' if (gross_chg or 0) >= 0 else ''}{gross_chg:.1f}% vs the 7-day average of ₹{s_avg['gross_sales']:,.0f})." if gross_chg is not None
        else f"Yesterday's gross sales were ₹{s_today['gross_sales']:,.0f} across {s_today['orders']} orders.",
        f"Meta Ads spent ₹{m_today['spend']:,.0f}, driving {m_today['purchases']} attributed purchases at a "
        f"{m_today['roas']:.2f}x blended ROAS "
        f"({'+' if (roas_chg or 0) >= 0 else ''}{roas_chg:.1f}% vs 7-day avg {m_avg['roas']:.2f}x)." if roas_chg is not None
        else f"Meta Ads spent ₹{m_today['spend']:,.0f}, driving {m_today['purchases']} purchases at {m_today['roas']:.2f}x ROAS.",
    ]
    for line in exec_lines:
        r = ws_dash.max_row + 1
        ws_dash.append([line, "", "", ""])
        ws_dash.merge_cells(f"A{r}:D{r}")
        ws_dash.cell(row=r, column=1).alignment = Alignment(wrap_text=True)

    hdr_row = section("SHOPIFY PERFORMANCE — YESTERDAY vs 7-DAY AVG")
    ws_dash.append(["Metric", "Yesterday", "7-Day Avg", "Change vs Avg"])
    style_header(ws_dash, ws_dash.max_row, 4, fill=SECTION_FILL, font=BOLD, height=18)
    shop_metric_start = ws_dash.max_row + 1
    shop_metrics = [
        ("Gross Sales", s_today["gross_sales"], s_avg["gross_sales"], INR0),
        ("Net Sales", s_today["net_sales"], s_avg.get("net_sales", avg([r["net_sales"] for r in s_prior])), INR0),
        ("Orders", s_today["orders"], s_avg["orders"], "0.0"),
        ("Avg Order Value", s_today["aov"], s_avg["aov"], INR0),
        ("Sessions", s_today["sessions"], s_avg["sessions"], "0"),
    ]
    for name, cur, prior, fmt in shop_metrics:
        chg = pct_change(cur, prior)
        row = [name, cur, round(prior, 1), (chg if chg is not None else "")]
        ws_dash.append(row)
        rr = ws_dash.max_row
        ws_dash.cell(row=rr, column=2).number_format = fmt
        ws_dash.cell(row=rr, column=3).number_format = fmt
        ws_dash.cell(row=rr, column=4).number_format = PCT1_SIGNED
    shop_metric_end = ws_dash.max_row

    hdr_row2 = section("META ADS PERFORMANCE — YESTERDAY vs 7-DAY AVG")
    ws_dash.append(["Metric", "Yesterday", "7-Day Avg", "Change vs Avg"])
    style_header(ws_dash, ws_dash.max_row, 4, fill=SECTION_FILL, font=BOLD, height=18)
    meta_metric_start = ws_dash.max_row + 1
    meta_metrics = [
        ("Spend", m_today["spend"], m_avg["spend"], INR0),
        ("Purchases", m_today["purchases"], m_avg["purchases"], "0.0"),
        ("CPA", m_today["cpa"], m_avg["cpa"], INR0),
        ("ROAS", m_today["roas"], m_avg["roas"], ROAS_FMT),
        ("CTR (%)", m_today["ctr"], m_avg["ctr"], "0.00"),
    ]
    for name, cur, prior, fmt in meta_metrics:
        chg = pct_change(cur, prior)
        ws_dash.append([name, cur, round(prior, 2), (chg if chg is not None else "")])
        rr = ws_dash.max_row
        ws_dash.cell(row=rr, column=2).number_format = fmt
        ws_dash.cell(row=rr, column=3).number_format = fmt
        ws_dash.cell(row=rr, column=4).number_format = PCT1_SIGNED
    meta_metric_end = ws_dash.max_row

    section("KEY WINS")
    for w in wins[:4]:
        r = ws_dash.max_row + 1
        ws_dash.append(["✅ " + w, "", "", ""])
        ws_dash.merge_cells(f"A{r}:D{r}")
        ws_dash.cell(row=r, column=1).alignment = Alignment(wrap_text=True)

    section("KEY ISSUES")
    for i in issues[:4]:
        r = ws_dash.max_row + 1
        ws_dash.append(["⚠️ " + i, "", "", ""])
        ws_dash.merge_cells(f"A{r}:D{r}")
        ws_dash.cell(row=r, column=1).alignment = Alignment(wrap_text=True)

    section("RECOMMENDED ACTIONS TODAY")
    for idx, a in enumerate(actions[:5], 1):
        r = ws_dash.max_row + 1
        ws_dash.append([f"{idx}. {a}", "", "", ""])
        ws_dash.merge_cells(f"A{r}:D{r}")
        ws_dash.cell(row=r, column=1).alignment = Alignment(wrap_text=True)

    # conditional formatting on the two "Change vs Avg" ranges
    for rng in (f"D{shop_metric_start}:D{shop_metric_end}", f"D{meta_metric_start}:D{meta_metric_end}"):
        ws_dash.conditional_formatting.add(
            rng, CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=GREEN)))
        ws_dash.conditional_formatting.add(
            rng, CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=RED)))

    ws_dash.freeze_panes = "A2"
    ws_dash.column_dimensions["A"].width = 46
    ws_dash.column_dimensions["B"].width = 18
    ws_dash.column_dimensions["C"].width = 18
    ws_dash.column_dimensions["D"].width = 18

    # ── Charts (Dashboard tab) ──────────────────────────────────────────
    n_hist = len(shopify_rows)
    chart_top = ws_dash.max_row + 3

    sales_chart = LineChart()
    sales_chart.title = "Gross Sales Trend (₹)"
    sales_chart.height, sales_chart.width = 8, 16
    data = Reference(ws_shop, min_col=2, min_row=1, max_row=1 + n_hist)
    cats = Reference(ws_shop, min_col=1, min_row=2, max_row=1 + n_hist)
    sales_chart.add_data(data, titles_from_data=True)
    sales_chart.set_categories(cats)
    ws_dash.add_chart(sales_chart, f"F{chart_top}")

    orders_chart = BarChart()
    orders_chart.title = "Orders Trend"
    orders_chart.height, orders_chart.width = 8, 16
    data = Reference(ws_shop, min_col=4, min_row=1, max_row=1 + n_hist)
    orders_chart.add_data(data, titles_from_data=True)
    orders_chart.set_categories(cats)
    ws_dash.add_chart(orders_chart, f"F{chart_top + 17}")

    n_hist_m = len(meta_rows)
    spend_chart = LineChart()
    spend_chart.title = "Ad Spend Trend (₹)"
    spend_chart.height, spend_chart.width = 8, 16
    data = Reference(ws_meta, min_col=2, min_row=1, max_row=1 + n_hist_m)
    cats_m = Reference(ws_meta, min_col=1, min_row=2, max_row=1 + n_hist_m)
    spend_chart.add_data(data, titles_from_data=True)
    spend_chart.set_categories(cats_m)
    ws_dash.add_chart(spend_chart, f"N{chart_top}")

    roas_chart = LineChart()
    roas_chart.title = "ROAS Trend"
    roas_chart.height, roas_chart.width = 8, 16
    data = Reference(ws_meta, min_col=11, min_row=1, max_row=1 + n_hist_m)
    roas_chart.add_data(data, titles_from_data=True)
    roas_chart.set_categories(cats_m)
    ws_dash.add_chart(roas_chart, f"N{chart_top + 17}")

    top_prod_chart = BarChart()
    top_prod_chart.title = "Top Products — Gross Sales (₹, Yesterday)"
    top_prod_chart.height, top_prod_chart.width = 8, 16
    n_prod = len(latest.get("top_products_yesterday", []))
    data = Reference(ws_prod, min_col=2, min_row=2, max_row=2 + n_prod)
    cats_p = Reference(ws_prod, min_col=1, min_row=3, max_row=2 + n_prod)
    top_prod_chart.add_data(data, titles_from_data=True)
    top_prod_chart.set_categories(cats_p)
    ws_dash.add_chart(top_prod_chart, f"F{chart_top + 34}")

    top_camp_chart = BarChart()
    top_camp_chart.title = "Top Campaigns — Spend (₹, Yesterday)"
    top_camp_chart.height, top_camp_chart.width = 8, 16
    n_camp = len(camps)
    data = Reference(ws_camp, min_col=3, min_row=2, max_row=2 + n_camp)
    cats_c = Reference(ws_camp, min_col=1, min_row=3, max_row=2 + n_camp)
    top_camp_chart.add_data(data, titles_from_data=True)
    top_camp_chart.set_categories(cats_c)
    ws_dash.add_chart(top_camp_chart, f"N{chart_top + 34}")

    # tab order + save
    wb._sheets = [ws_dash, ws_shop, ws_meta, ws_prod, ws_camp, ws_notes]
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    wb.save(OUT_PATH)
    print(f"Saved {OUT_PATH}")
    print(f"Report date: {date_str}, {len(shopify_rows)} Shopify days, {len(meta_rows)} Meta days in history.")


if __name__ == "__main__":
    main()
