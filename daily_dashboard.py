#!/usr/bin/env python3
"""
Daily Store & Ads Performance Dashboard
Runs every morning in Asia/Kolkata timezone.
Shopify store: 36dhns-ed.myshopify.com (Dhirai)
Meta Ads account: 979830497515712 (Dhirai)

Requirements:
    pip install gspread google-auth google-auth-oauthlib google-api-python-client
                requests pytz

Environment variables (set in .env or system):
    SHOPIFY_STORE_DOMAIN      e.g. 36dhns-ed.myshopify.com
    SHOPIFY_ACCESS_TOKEN      Private-app Admin API token
    META_ACCESS_TOKEN         Meta Marketing API user access token
    META_AD_ACCOUNT_ID        e.g. 979830497515712
    GOOGLE_SA_CREDENTIALS     Path to Google service-account JSON file
    GMAIL_TO                  Recipient email (atul012001@gmail.com)
    GMAIL_CC                  CC email(s), comma-separated (leave blank for none)
    SPREADSHEET_NAME          Sheet name (default: Daily Store & Ads Performance Sheet)
"""

import os
import json
import datetime
import requests
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pytz

# ─────────────────────── Configuration ────────────────────────────────────────

KOLKATA_TZ = pytz.timezone("Asia/Kolkata")
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/drive.file",
]

SHOPIFY_DOMAIN    = os.environ.get("SHOPIFY_STORE_DOMAIN", "36dhns-ed.myshopify.com")
SHOPIFY_TOKEN     = os.environ.get("SHOPIFY_ACCESS_TOKEN", "")
META_TOKEN        = os.environ.get("META_ACCESS_TOKEN", "")
META_ACCOUNT_ID   = os.environ.get("META_AD_ACCOUNT_ID", "979830497515712")
SA_CREDENTIALS    = os.environ.get("GOOGLE_SA_CREDENTIALS", "service_account.json")
GMAIL_TO          = os.environ.get("GMAIL_TO", "atul012001@gmail.com")
GMAIL_CC          = os.environ.get("GMAIL_CC", "")
SPREADSHEET_NAME  = os.environ.get("SPREADSHEET_NAME", "Daily Store & Ads Performance Sheet")

# ─────────────────────── Colours ──────────────────────────────────────────────

C_HEADER_BG   = {"red": 0.102, "green": 0.227, "blue": 0.420}  # Brand navy
C_HEADER_FG   = {"red": 1,     "green": 1,     "blue": 1}
C_SECTION_BG  = {"red": 0.906, "green": 0.922, "blue": 0.957}  # Light blue-grey
C_GREEN       = {"red": 0.204, "green": 0.659, "blue": 0.325}
C_RED         = {"red": 0.839, "green": 0.153, "blue": 0.157}
C_YELLOW      = {"red": 1,     "green": 0.898, "blue": 0.2}
C_ORANGE      = {"red": 1,     "green": 0.596, "blue": 0}
C_ACCENT      = {"red": 0.259, "green": 0.522, "blue": 0.957}
C_WHITE       = {"red": 1,     "green": 1,     "blue": 1}

# ─────────────────────── Google Auth ──────────────────────────────────────────

def get_google_clients():
    creds = Credentials.from_service_account_file(SA_CREDENTIALS, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    gmail_service  = build("gmail",  "v1", credentials=creds)
    return gc, sheets_service, gmail_service

# ─────────────────────── Date helpers ─────────────────────────────────────────

def yesterday_kolkata() -> datetime.date:
    return datetime.datetime.now(KOLKATA_TZ).date() - datetime.timedelta(days=1)

# ─────────────────────── Shopify data fetching ────────────────────────────────

SHOPIFY_GRAPHQL_URL = f"https://{SHOPIFY_DOMAIN}/admin/api/2024-01/graphql.json"

def shopify_analytics(shopify_query: str) -> dict:
    gql = """
    query RunAnalytics($query: String!) {
        shopifyqlQuery(query: $query) {
            ... on TableResponse {
                tableData {
                    columns { name dataType }
                    rowData
                    unformattedData
                }
            }
        }
    }
    """
    r = requests.post(
        SHOPIFY_GRAPHQL_URL,
        headers={
            "Content-Type": "application/json",
            "X-Shopify-Access-Token": SHOPIFY_TOKEN,
        },
        json={"query": gql, "variables": {"query": shopify_query}},
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    try:
        td = data["data"]["shopifyqlQuery"]["tableData"]
        return {
            "columns": [c["name"] for c in td["columns"]],
            "rows":    td["unformattedData"] or td["rowData"],
        }
    except (KeyError, TypeError):
        return {"columns": [], "rows": [], "error": json.dumps(data)}


def fetch_shopify_daily(date_str: str) -> dict:
    results = {}

    q = (
        f"FROM sales SHOW gross_sales, net_sales, orders, average_order_value, "
        f"total_sales, discounts, returns, shipping_charges, taxes "
        f"SINCE {date_str} UNTIL {date_str}"
    )
    r = shopify_analytics(q)
    if r["rows"]:
        row = r["rows"][0]
        results["sales"] = dict(zip(r["columns"], row))
    else:
        results["sales"] = {c: None for c in [
            "gross_sales","net_sales","orders","average_order_value",
            "total_sales","discounts","returns","shipping_charges","taxes"
        ]}

    q = (
        f"FROM sessions SHOW sessions, sessions_with_cart_additions, "
        f"sessions_that_reached_checkout, sessions_that_completed_checkout, "
        f"conversion_rate SINCE {date_str} UNTIL {date_str}"
    )
    r = shopify_analytics(q)
    results["sessions"] = dict(zip(r["columns"], r["rows"][0])) if r["rows"] else {}

    q = (
        f"FROM sales SHOW gross_sales, net_sales, orders, total_sales "
        f"GROUP BY product_title ORDER BY gross_sales DESC LIMIT 10 "
        f"SINCE {date_str} UNTIL {date_str}"
    )
    results["top_products"] = shopify_analytics(q)

    return results


def fetch_shopify_8day_timeseries(since_str: str, until_str: str) -> dict:
    """Daily sales + sessions timeseries for 8-day window (7 prior + yesterday)."""
    sales_q = (
        f"FROM sales SHOW gross_sales, net_sales, orders, average_order_value, "
        f"total_sales, discounts, returns, shipping_charges, taxes "
        f"TIMESERIES day SINCE {since_str} UNTIL {until_str}"
    )
    sessions_q = (
        f"FROM sessions SHOW sessions, sessions_with_cart_additions, "
        f"sessions_that_reached_checkout, sessions_that_completed_checkout, conversion_rate "
        f"TIMESERIES day SINCE {since_str} UNTIL {until_str}"
    )
    sales    = shopify_analytics(sales_q)
    sessions = shopify_analytics(sessions_q)
    return {"sales": sales, "sessions": sessions}


def fetch_shopify_7day_products(since_str: str, until_str: str) -> dict:
    q = (
        f"FROM sales SHOW gross_sales, net_sales, orders "
        f"GROUP BY product_title ORDER BY gross_sales DESC LIMIT 10 "
        f"SINCE {since_str} UNTIL {until_str}"
    )
    return shopify_analytics(q)


# ─────────────────────── Meta Ads data fetching ────────────────────────────────

META_BASE   = "https://graph.facebook.com/v19.0"
META_FIELDS = (
    "campaign_id,campaign_name,adset_id,adset_name,ad_id,ad_name,"
    "spend,impressions,reach,clicks,ctr,cpc,cpm,"
    "purchase_roas,actions,cost_per_action_type,frequency,"
    "date_start,date_stop"
)


def meta_request(endpoint: str, params: dict) -> dict:
    params["access_token"] = META_TOKEN
    r = requests.get(f"{META_BASE}/{endpoint}", params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def fetch_meta_level(date_str: str, level: str) -> list:
    r = meta_request(
        f"act_{META_ACCOUNT_ID}/insights",
        {
            "level": level,
            "fields": META_FIELDS,
            "time_range": json.dumps({"since": date_str, "until": date_str}),
            "limit": 100,
        },
    )
    return r.get("data", [])


def fetch_meta_7day(since_str: str, until_str: str) -> list:
    """Account-level daily insights for the 7-day prior window."""
    r = meta_request(
        f"act_{META_ACCOUNT_ID}/insights",
        {
            "level": "account",
            "fields": "spend,impressions,reach,clicks,ctr,cpc,cpm,purchase_roas,actions,cost_per_action_type,date_start",
            "time_range": json.dumps({"since": since_str, "until": until_str}),
            "time_increment": 1,
            "limit": 50,
        },
    )
    return r.get("data", [])


def extract_purchases(actions: list) -> int:
    count = 0
    for a in (actions or []):
        if a.get("action_type") in (
            "offsite_conversion.fb_pixel_purchase", "purchase", "omni_purchase"
        ):
            try:
                count += int(float(a.get("value", 0)))
            except (ValueError, TypeError):
                pass
    return count


def extract_roas(purchase_roas) -> float | None:
    if not purchase_roas:
        return None
    if isinstance(purchase_roas, (int, float)):
        return round(float(purchase_roas), 2)
    for r in (purchase_roas if isinstance(purchase_roas, list) else []):
        if r.get("action_type") in (
            "offsite_conversion.fb_pixel_purchase", "omni_purchase", "purchase"
        ):
            try:
                return round(float(r["value"]), 2)
            except (ValueError, KeyError):
                pass
    return None


def extract_cpa(cost_per_action_type) -> float | None:
    for a in (cost_per_action_type or []):
        if a.get("action_type") in (
            "offsite_conversion.fb_pixel_purchase", "purchase", "omni_purchase"
        ):
            try:
                return round(float(a["value"]), 2)
            except (ValueError, KeyError):
                pass
    return None


# ─────────────────────── Computation helpers ──────────────────────────────────

def safe_float(v) -> float:
    try:
        s = str(v).replace(",", "").replace("₹", "").replace(" INR", "").strip()
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def pct_change(current: float, prior: float) -> float | None:
    if prior == 0:
        return None
    return round((current - prior) / prior * 100, 1)


def avg(values: list) -> float:
    vals = [v for v in values if v is not None]
    return sum(vals) / len(vals) if vals else 0.0


def compute_7day_averages(daily_rows: list, col_idx: dict) -> dict:
    avgs = {}
    for col, idx in col_idx.items():
        vals = []
        for row in daily_rows:
            try:
                vals.append(float(row[idx]))
            except (ValueError, TypeError, IndexError):
                pass
        avgs[col] = round(avg(vals), 2) if vals else 0.0
    return avgs


def spend_weighted_roas(campaigns: list) -> float:
    total_revenue = 0.0
    total_spend   = 0.0
    for c in campaigns:
        spend = safe_float(c.get("spend", 0))
        roas  = extract_roas(c.get("purchase_roas"))
        if roas is not None and spend > 0:
            total_revenue += roas * spend
            total_spend   += spend
    return round(total_revenue / total_spend, 2) if total_spend > 0 else 0.0


def aggregate_meta_campaigns(campaigns: list) -> dict:
    total_spend  = sum(safe_float(c.get("spend", 0)) for c in campaigns)
    total_impr   = sum(int(safe_float(c.get("impressions", 0))) for c in campaigns)
    total_reach  = sum(int(safe_float(c.get("reach", 0))) for c in campaigns)
    total_clicks = sum(int(safe_float(c.get("clicks", 0))) for c in campaigns)
    total_purch  = sum(extract_purchases(c.get("actions", [])) for c in campaigns)
    ctr  = round(total_clicks / total_impr * 100, 2) if total_impr else 0
    cpc  = round(total_spend  / total_clicks, 2)     if total_clicks else 0
    cpm  = round(total_spend  / total_impr * 1000, 2) if total_impr else 0
    cpa  = round(total_spend  / total_purch, 2)      if total_purch else 0
    roas = spend_weighted_roas(campaigns)
    return {
        "spend": round(total_spend, 2), "impressions": total_impr,
        "reach": total_reach, "clicks": total_clicks, "ctr": ctr,
        "cpc": cpc, "cpm": cpm, "purchases": total_purch, "cpa": cpa, "roas": roas,
    }


# ─────────────────────── Recommendations engine ────────────────────────────────

def generate_recommendations(shopify_today: dict, shopify_7d_avg: dict,
                              meta_campaigns: list) -> tuple[list, list]:
    recs  = []
    notes = []

    s     = shopify_today.get("sales", {})
    gross = safe_float(s.get("gross_sales", 0))
    orders = safe_float(s.get("orders", 0))
    rt    = abs(safe_float(s.get("returns", 0)))

    ag_gs = shopify_7d_avg.get("gross_sales", 0)
    ag_od = shopify_7d_avg.get("orders", 0)

    if ag_gs > 0:
        chg = pct_change(gross, ag_gs)
        if chg is not None and chg < -15:
            recs.append(f"⚠️ Revenue dropped {abs(chg):.1f}% vs 7-day avg — check ad delivery, site performance, and seasonal demand.")
        elif chg is not None and chg > 20:
            recs.append(f"✅ Revenue up {chg:.1f}% vs 7-day avg — identify top campaigns and scale budgets.")

    if ag_od > 0 and orders < ag_od * 0.8:
        recs.append("⚠️ Orders below 80% of 7-day avg — review traffic sources and Meta campaign delivery.")

    sess = shopify_today.get("sessions", {})
    completed_checkout = safe_float(sess.get("sessions_that_completed_checkout", 0))
    total_sessions     = safe_float(sess.get("sessions", 0))
    if total_sessions > 1000 and completed_checkout < 5:
        recs.append(
            f"⚠️ Only {int(completed_checkout)} completed checkout(s) from {int(total_sessions):,} sessions. "
            "Test the checkout flow urgently — possible payment or UX issue."
        )

    if gross > 0 and rt / gross > 0.12:
        recs.append(f"⚠️ High return rate ({rt/gross*100:.1f}% of gross) — review sizing, quality, or product descriptions.")

    # Campaign-level analysis
    active = [c for c in meta_campaigns if safe_float(c.get("spend", 0)) > 0]
    for c in active:
        spend = safe_float(c.get("spend", 0))
        roas  = extract_roas(c.get("purchase_roas"))
        name  = c.get("campaign_name", c.get("name", ""))
        if roas is not None and spend > 1000 and roas < 1.0:
            recs.append(f"🚨 PAUSE '{name}' — ROAS {roas}x is below breakeven. Reallocate ₹{spend:,.0f} budget.")
        elif roas is not None and spend > 1000 and roas < 2.0:
            recs.append(f"⚠️ Review '{name}' — ROAS {roas:.2f}x is low on ₹{spend:,.0f} spend. Cut 20–30% budget or refresh creatives.")
        elif roas is not None and roas > 4.0 and spend > 500:
            recs.append(f"✅ Scale '{name}' — ROAS {roas:.2f}x is excellent. Increase budget by 25–30%.")

    if not recs:
        recs.append("✅ All campaigns performing within normal range — continue monitoring.")

    recs.append("🔍 Review Shopify abandoned checkouts and activate email recovery flows if not already running.")
    recs.append("📱 Audit product pages for top sellers — verify images, sizes, and descriptions are up to date.")

    return recs[:7], notes


# ─────────────────────── Sheets formatting helpers ────────────────────────────

def cell_format(sheet_id: int, start_row: int, end_row: int,
                start_col: int, end_col: int,
                bold: bool = False, bg_color: dict = None, fg_color: dict = None,
                horizontal_alignment: str = None, number_format_pattern: str = None,
                number_format_type: str = "NUMBER", font_size: int = None) -> dict:
    cell_fmt = {}
    tf = {}
    if bold:     tf["bold"] = True
    if fg_color: tf["foregroundColor"] = fg_color
    if font_size: tf["fontSize"] = font_size
    if tf:       cell_fmt["textFormat"] = tf
    if bg_color: cell_fmt["backgroundColor"] = bg_color
    if horizontal_alignment:
        cell_fmt["horizontalAlignment"] = horizontal_alignment
    if number_format_pattern:
        cell_fmt["numberFormat"] = {"type": number_format_type, "pattern": number_format_pattern}

    fields = []
    if tf:                    fields.append("textFormat")
    if bg_color:              fields.append("backgroundColor")
    if horizontal_alignment:  fields.append("horizontalAlignment")
    if number_format_pattern: fields.append("numberFormat")

    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col,
            },
            "cell": {"userEnteredFormat": cell_fmt},
            "fields": "userEnteredFormat(" + ",".join(fields) + ")",
        }
    }


def freeze_request(sheet_id: int, rows: int = 1, cols: int = 0) -> dict:
    return {
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols},
            },
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
        }
    }


def auto_resize_request(sheet_id: int, start_col: int = 0, end_col: int = 26) -> dict:
    return {
        "autoResizeDimensions": {
            "dimensions": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": start_col,
                "endIndex": end_col,
            }
        }
    }


def add_filter_request(sheet_id: int, start_row: int, end_row: int,
                        start_col: int, end_col: int) -> dict:
    return {
        "setBasicFilter": {
            "filter": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row,
                    "endRowIndex": end_row,
                    "startColumnIndex": start_col,
                    "endColumnIndex": end_col,
                }
            }
        }
    }


def conditional_pct_format(sheet_id: int, start_row: int, end_row: int,
                             start_col: int, end_col: int) -> list:
    base = {"sheetId": sheet_id, "startRowIndex": start_row,
            "endRowIndex": end_row, "startColumnIndex": start_col, "endColumnIndex": end_col}
    return [
        {"addConditionalFormatRule": {"rule": {"ranges": [base], "booleanRule": {
            "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0"}]},
            "format": {"backgroundColor": {"red": 0.714, "green": 0.929, "blue": 0.714}},
        }}, "index": 0}},
        {"addConditionalFormatRule": {"rule": {"ranges": [base], "booleanRule": {
            "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
            "format": {"backgroundColor": {"red": 0.957, "green": 0.714, "blue": 0.714}},
        }}, "index": 1}},
    ]


def conditional_roas_format(sheet_id: int, start_row: int, end_row: int,
                              start_col: int, end_col: int) -> list:
    base = {"sheetId": sheet_id, "startRowIndex": start_row,
            "endRowIndex": end_row, "startColumnIndex": start_col, "endColumnIndex": end_col}
    return [
        {"addConditionalFormatRule": {"rule": {"ranges": [base], "booleanRule": {
            "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "2"}]},
            "format": {"backgroundColor": {"red": 0.957, "green": 0.714, "blue": 0.714}},
        }}, "index": 0}},
        {"addConditionalFormatRule": {"rule": {"ranges": [base], "booleanRule": {
            "condition": {"type": "NUMBER_BETWEEN",
                          "values": [{"userEnteredValue": "2"}, {"userEnteredValue": "3.5"}]},
            "format": {"backgroundColor": {"red": 1, "green": 0.949, "blue": 0.8}},
        }}, "index": 1}},
        {"addConditionalFormatRule": {"rule": {"ranges": [base], "booleanRule": {
            "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "3.5"}]},
            "format": {"backgroundColor": {"red": 0.714, "green": 0.929, "blue": 0.714}},
        }}, "index": 2}},
    ]


def batch_update(sheets_svc, spreadsheet_id: str, requests_list: list):
    if requests_list:
        sheets_svc.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests_list}
        ).execute()


# ─────────────────────── Sheet tab builders ────────────────────────────────────

def build_shopify_daily_tab(ws, sheets_svc, spreadsheet_id: str, sheet_id: int,
                             timeseries_data: dict, date_str: str):
    """Append or update one day's row in Shopify Daily Data, never duplicate."""
    headers = [
        "Date", "Gross Sales (₹)", "Net Sales (₹)", "Orders", "AOV (₹)",
        "Total Sales (₹)", "Discounts (₹)", "Returns (₹)", "Shipping (₹)", "Taxes (₹)",
        "Sessions", "Cart Adds", "Reached Checkout", "Completed Checkout", "Conv. Rate (%)",
    ]

    sales_data    = timeseries_data["sales"]
    sessions_data = timeseries_data["sessions"]
    sales_cols    = sales_data.get("columns", [])
    sales_rows    = sales_data.get("rows", [])
    sess_cols     = sessions_data.get("columns", [])
    sess_rows     = sessions_data.get("rows", [])
    sci = {c: i for i, c in enumerate(sales_cols)}
    ssi = {c: i for i, c in enumerate(sess_cols)}

    # Build a lookup: date -> sessions row
    sess_lookup = {}
    for row in sess_rows:
        day = str(row[ssi.get("day", 0)]) if sess_cols else ""
        sess_lookup[day] = row

    # Check existing dates in sheet
    existing = ws.col_values(1)
    if not existing or existing[0] != "Date":
        ws.clear()
        ws.update("A1", [headers])
        existing = [headers[0]]

    existing_dates = set(existing[1:])
    new_rows = []
    for row in sales_rows:
        day = str(row[sci.get("day", 0)])
        if day in existing_dates:
            continue  # Skip duplicates
        sr = sess_lookup.get(day, [])
        new_rows.append([
            day,
            safe_float(row[sci.get("gross_sales", 1)]),
            safe_float(row[sci.get("net_sales", 2)]),
            int(safe_float(row[sci.get("orders", 3)])),
            round(safe_float(row[sci.get("average_order_value", 4)]), 2),
            safe_float(row[sci.get("total_sales", 5)]),
            abs(safe_float(row[sci.get("discounts", 6)])),
            abs(safe_float(row[sci.get("returns", 7)])),
            safe_float(row[sci.get("shipping_charges", 8)]),
            safe_float(row[sci.get("taxes", 9)]),
            int(safe_float(sr[ssi.get("sessions", 1)])) if sr and len(sr) > 1 else "",
            int(safe_float(sr[ssi.get("sessions_with_cart_additions", 2)])) if sr and len(sr) > 2 else "",
            int(safe_float(sr[ssi.get("sessions_that_reached_checkout", 3)])) if sr and len(sr) > 3 else "",
            int(safe_float(sr[ssi.get("sessions_that_completed_checkout", 4)])) if sr and len(sr) > 4 else "",
            round(safe_float(sr[ssi.get("conversion_rate", 5)]) * 100, 4) if sr and len(sr) > 5 else "",
        ])

    if new_rows:
        next_row = len(existing) + 1
        ws.update(f"A{next_row}", new_rows)

    total_rows = len(existing) + len(new_rows)
    reqs = [
        cell_format(sheet_id, 0, 1, 0, len(headers),
                    bold=True, bg_color=C_HEADER_BG, fg_color=C_HEADER_FG),
        freeze_request(sheet_id, rows=1),
        add_filter_request(sheet_id, 0, total_rows, 0, len(headers)),
        auto_resize_request(sheet_id, 0, len(headers)),
    ]
    for col in [1, 2, 4, 5, 6, 7, 8, 9]:
        reqs.append(cell_format(sheet_id, 1, total_rows, col, col + 1,
                                number_format_pattern='₹#,##0.00'))
    reqs.append(cell_format(sheet_id, 1, total_rows, 14, 15,
                            number_format_pattern='0.0000"%"'))
    batch_update(sheets_svc, spreadsheet_id, reqs)


def build_meta_daily_tab(ws, sheets_svc, spreadsheet_id: str, sheet_id: int,
                          meta_summary: dict, date_str: str):
    """Append one row per day to Meta Ads Daily Data, no duplicates."""
    headers = [
        "Date", "Total Spend (₹)", "Impressions", "Reach", "Clicks",
        "CTR (%)", "CPC (₹)", "CPM (₹)", "Purchases", "CPA (₹)", "ROAS",
    ]

    existing = ws.col_values(1)
    if not existing or existing[0] != "Date":
        ws.clear()
        ws.update("A1", [headers])
        existing = [headers[0]]

    if date_str not in existing[1:]:
        new_row = [
            date_str,
            meta_summary["spend"],
            meta_summary["impressions"],
            meta_summary["reach"],
            meta_summary["clicks"],
            meta_summary["ctr"],
            meta_summary["cpc"],
            meta_summary["cpm"],
            meta_summary["purchases"],
            meta_summary["cpa"],
            meta_summary["roas"],
        ]
        ws.update(f"A{len(existing) + 1}", [new_row])
        total_rows = len(existing) + 2
    else:
        total_rows = len(existing) + 1

    reqs = [
        cell_format(sheet_id, 0, 1, 0, len(headers),
                    bold=True, bg_color=C_HEADER_BG, fg_color=C_HEADER_FG),
        freeze_request(sheet_id, rows=1),
        add_filter_request(sheet_id, 0, total_rows, 0, len(headers)),
        auto_resize_request(sheet_id, 0, len(headers)),
    ]
    for col in [1, 6, 7, 9]:
        reqs.append(cell_format(sheet_id, 1, total_rows, col, col + 1,
                                number_format_pattern='₹#,##0.00'))
    reqs.append(cell_format(sheet_id, 1, total_rows, 5, 6,
                            number_format_pattern='0.00"%"'))
    reqs += conditional_roas_format(sheet_id, 1, total_rows, 10, 11)
    batch_update(sheets_svc, spreadsheet_id, reqs)


def build_product_tab(ws, sheets_svc, spreadsheet_id: str, sheet_id: int,
                       today_products: dict, seven_day_products: dict):
    ws.clear()
    ws.update("A1", [["─── Yesterday ───", "", "", "", "", "─── 7-Day Total ───", "", ""]])
    headers = ["Product", "Gross Sales (₹)", "Net Sales (₹)", "Orders",
               "", "7-Day Gross (₹)", "7-Day Net (₹)", "7-Day Orders"]
    ws.update("A2", [headers])

    today_rows = today_products.get("rows", [])
    seven_rows = seven_day_products.get("rows", [])
    seven_lookup = {str(r[0]): r for r in seven_rows}

    data = []
    for row in today_rows:
        prod  = str(row[0])
        seven = seven_lookup.get(prod, [])
        data.append([
            prod,
            safe_float(row[1]),
            safe_float(row[2]),
            int(safe_float(row[3])),
            "",
            safe_float(seven[1]) if len(seven) > 1 else "",
            safe_float(seven[2]) if len(seven) > 2 else "",
            int(safe_float(seven[3])) if len(seven) > 3 else "",
        ])
    if data:
        ws.update("A3", data)

    total_rows = len(data) + 3
    reqs = [
        cell_format(sheet_id, 0, 1, 0, 8, bold=True, bg_color=C_ACCENT, fg_color=C_HEADER_FG),
        cell_format(sheet_id, 1, 2, 0, 8, bold=True, bg_color=C_HEADER_BG, fg_color=C_HEADER_FG),
        freeze_request(sheet_id, rows=2),
        add_filter_request(sheet_id, 1, total_rows, 0, 8),
        auto_resize_request(sheet_id, 0, 8),
    ]
    for col in [1, 2, 5, 6]:
        reqs.append(cell_format(sheet_id, 2, total_rows, col, col + 1,
                                number_format_pattern='₹#,##0.00'))
    batch_update(sheets_svc, spreadsheet_id, reqs)


def build_campaign_tab(ws, sheets_svc, spreadsheet_id: str, sheet_id: int,
                        campaigns: list, adsets: list):
    ws.clear()
    camp_hdrs = ["Campaign", "Status", "Spend (₹)", "Impressions", "Reach",
                 "Clicks", "CTR (%)", "CPC (₹)", "Purchases", "CPA (₹)", "ROAS"]
    ws.update("A1", [["━━━ CAMPAIGN LEVEL ━━━"] + [""] * (len(camp_hdrs) - 1)])
    ws.update("A2", [camp_hdrs])

    camp_data = []
    for c in campaigns:
        spend = safe_float(c.get("spend", 0))
        if spend == 0:
            continue
        camp_data.append([
            c.get("campaign_name", c.get("name", "")),
            c.get("effective_status", c.get("status", "")),
            round(spend, 2),
            int(safe_float(c.get("impressions", 0))),
            int(safe_float(c.get("reach", 0))),
            int(safe_float(c.get("clicks", 0))),
            round(safe_float(c.get("ctr", 0)), 2),
            round(safe_float(c.get("cpc", 0)), 2),
            extract_purchases(c.get("actions", [])),
            extract_cpa(c.get("cost_per_action_type")) or "",
            extract_roas(c.get("purchase_roas")) or "",
        ])
    camp_data.sort(key=lambda r: r[2], reverse=True)
    if camp_data:
        ws.update("A3", camp_data)
    camp_end = len(camp_data) + 3

    adset_hdrs = ["Ad Set", "Campaign", "Status", "Spend (₹)", "Impressions", "Reach",
                  "Clicks", "CTR (%)", "CPC (₹)", "Purchases", "CPA (₹)", "ROAS"]
    as_start = camp_end + 2
    ws.update(f"A{as_start}", [["━━━ AD SET LEVEL ━━━"] + [""] * (len(adset_hdrs) - 1)])
    ws.update(f"A{as_start + 1}", [adset_hdrs])

    as_data = []
    for a in adsets:
        spend = safe_float(a.get("spend", 0))
        if spend == 0:
            continue
        as_data.append([
            a.get("adset_name", a.get("name", "")),
            a.get("campaign_name", ""),
            a.get("effective_status", a.get("status", "")),
            round(spend, 2),
            int(safe_float(a.get("impressions", 0))),
            int(safe_float(a.get("reach", 0))),
            int(safe_float(a.get("clicks", 0))),
            round(safe_float(a.get("ctr", 0)), 2),
            round(safe_float(a.get("cpc", 0)), 2),
            extract_purchases(a.get("actions", [])),
            extract_cpa(a.get("cost_per_action_type")) or "",
            extract_roas(a.get("purchase_roas")) or "",
        ])
    as_data.sort(key=lambda r: r[3], reverse=True)
    if as_data:
        ws.update(f"A{as_start + 2}", as_data)
    as_end = as_start + 2 + len(as_data)

    reqs = [
        cell_format(sheet_id, 0, 1, 0, len(camp_hdrs),
                    bold=True, bg_color=C_ACCENT, fg_color=C_HEADER_FG, font_size=11),
        cell_format(sheet_id, 1, 2, 0, len(camp_hdrs),
                    bold=True, bg_color=C_HEADER_BG, fg_color=C_HEADER_FG),
        cell_format(sheet_id, as_start - 1, as_start, 0, len(adset_hdrs),
                    bold=True, bg_color=C_ACCENT, fg_color=C_HEADER_FG, font_size=11),
        cell_format(sheet_id, as_start, as_start + 1, 0, len(adset_hdrs),
                    bold=True, bg_color=C_HEADER_BG, fg_color=C_HEADER_FG),
        freeze_request(sheet_id, rows=2),
        auto_resize_request(sheet_id, 0, len(adset_hdrs)),
    ]
    for col in [2, 7, 9]:
        reqs.append(cell_format(sheet_id, 2, camp_end, col, col + 1,
                                number_format_pattern='₹#,##0.00'))
    reqs.append(cell_format(sheet_id, 2, camp_end, 6, 7,
                            number_format_pattern='0.00"%"'))
    reqs += conditional_roas_format(sheet_id, 2, camp_end, 10, 11)
    for col in [3, 8, 10]:
        reqs.append(cell_format(sheet_id, as_start + 1, as_end, col, col + 1,
                                number_format_pattern='₹#,##0.00'))
    reqs += conditional_roas_format(sheet_id, as_start + 1, as_end, 11, 12)
    batch_update(sheets_svc, spreadsheet_id, reqs)


def build_recommendations_tab(ws, sheets_svc, spreadsheet_id: str, sheet_id: int,
                                recs: list, date_str: str, unavailable: list = None):
    ws.clear()
    all_rows = [
        [f"RECOMMENDATIONS & NOTES — {date_str}", ""],
        ["", ""],
        ["TYPE", "RECOMMENDATION / NOTE"],
    ]
    for rec in recs:
        icon = rec[:2] if rec and rec[0] in "⚠🚨✅📊🔍📱💡" else "💡"
        all_rows.append([icon, rec[len(icon):].strip()])

    if unavailable:
        all_rows += [["", ""], ["UNAVAILABLE METRICS", ""]]
        for m in unavailable:
            all_rows.append(["ℹ️", m])

    ws.update("A1", all_rows)
    n = len(all_rows)
    reqs = [
        cell_format(sheet_id, 0, 1, 0, 2, bold=True, bg_color=C_ACCENT,
                    fg_color=C_HEADER_FG, font_size=13),
        cell_format(sheet_id, 2, 3, 0, 2, bold=True, bg_color=C_HEADER_BG,
                    fg_color=C_HEADER_FG),
        freeze_request(sheet_id, rows=3),
        auto_resize_request(sheet_id, 0, 2),
    ]
    batch_update(sheets_svc, spreadsheet_id, reqs)


def build_dashboard_tab(ws, sheets_svc, spreadsheet_id: str, sheet_id: int,
                         shopify_today: dict, shopify_7d_avg: dict,
                         meta_summary: dict, recs: list, date_str: str):
    ws.clear()

    def fmt(v: float) -> str: return f"₹{v:,.0f}"
    def pc(v) -> str:
        if v is None: return "—"
        return f"+{v:.1f}%" if v >= 0 else f"{v:.1f}%"

    s  = shopify_today.get("sales", {})
    gs = safe_float(s.get("gross_sales", 0))
    ns = safe_float(s.get("net_sales", 0))
    od = int(safe_float(s.get("orders", 0)))
    ao = safe_float(s.get("average_order_value", 0))
    rt = abs(safe_float(s.get("returns", 0)))

    ag_gs = shopify_7d_avg.get("gross_sales", 0)
    ag_ns = shopify_7d_avg.get("net_sales", 0)
    ag_od = shopify_7d_avg.get("orders", 0)
    ag_ao = shopify_7d_avg.get("average_order_value", 0)

    sp  = meta_summary.get("spend", 0)
    imp = meta_summary.get("impressions", 0)
    cl  = meta_summary.get("clicks", 0)
    pu  = meta_summary.get("purchases", 0)
    ctr = meta_summary.get("ctr", 0)
    cpa = meta_summary.get("cpa", 0)
    ro  = meta_summary.get("roas", 0)

    rows = [
        [f"DAILY STORE & ADS PERFORMANCE DASHBOARD — {date_str}", "", "", ""],
        ["Dhirai | 36dhns-ed.myshopify.com | Meta Account: 979830497515712", "", "", ""],
        ["", "", "", ""],
        ["━━━ SHOPIFY PERFORMANCE ━━━", "", "", ""],
        ["Metric", "Yesterday", "7-Day Daily Avg", "Change vs Avg"],
        ["Gross Sales",     fmt(gs), fmt(ag_gs), pc(pct_change(gs, ag_gs))],
        ["Net Sales",       fmt(ns), fmt(ag_ns), pc(pct_change(ns, ag_ns))],
        ["Orders",          str(od), str(round(ag_od, 1)), pc(pct_change(od, ag_od))],
        ["Avg Order Value", fmt(ao), fmt(ag_ao), pc(pct_change(ao, ag_ao))],
        ["Returns",         fmt(rt), "", ""],
        ["", "", "", ""],
        ["━━━ META ADS PERFORMANCE ━━━", "", "", ""],
        ["Metric", "Yesterday", "", ""],
        ["Total Spend",    fmt(sp),          "", ""],
        ["Impressions",    f"{imp:,}",       "", ""],
        ["Clicks",         f"{cl:,}",        "", ""],
        ["CTR",            f"{ctr:.2f}%",    "", ""],
        ["Purchases",      str(pu),          "", ""],
        ["CPA",            fmt(cpa),         "", ""],
        ["Blended ROAS",   f"{ro:.2f}x",     "", ""],
        ["", "", "", ""],
        ["━━━ KEY WINS ━━━", "", "", ""],
    ]

    wins = [r for r in recs if r.startswith("✅")]
    rows += [["", w, "", ""] for w in wins[:3]] or [["", "Monitoring for positive trends.", "", ""]]
    rows += [["", "", "", ""], ["━━━ KEY ISSUES ━━━", "", "", ""]]
    issues = [r for r in recs if r.startswith(("⚠️", "🚨"))]
    rows += [["", i, "", ""] for i in issues[:3]] or [["", "No critical issues.", "", ""]]
    rows += [["", "", "", ""], ["━━━ RECOMMENDED ACTIONS TODAY ━━━", "", "", ""]]
    for idx, rec in enumerate(recs[:5], 1):
        rows.append([f"{idx}.", rec, "", ""])

    ws.update("A1", rows)

    reqs = [
        cell_format(sheet_id, 0, 1, 0, 4, bold=True, bg_color=C_HEADER_BG,
                    fg_color=C_HEADER_FG, font_size=14),
        cell_format(sheet_id, 1, 2, 0, 4, bg_color=C_SECTION_BG,
                    fg_color={"red": 0.4, "green": 0.4, "blue": 0.4}, font_size=11),
        freeze_request(sheet_id, rows=1),
        auto_resize_request(sheet_id, 0, 4),
    ]
    section_header_rows = [3, 11, 21]
    for sr in section_header_rows:
        reqs.append(cell_format(sheet_id, sr, sr + 1, 0, 4,
                                bold=True, bg_color=C_ACCENT, fg_color=C_HEADER_FG,
                                font_size=11))
    reqs.append(cell_format(sheet_id, 4, 5, 0, 4,
                            bold=True, bg_color=C_SECTION_BG))
    reqs.append(cell_format(sheet_id, 12, 13, 0, 4,
                            bold=True, bg_color=C_SECTION_BG))
    reqs += conditional_pct_format(sheet_id, 5, 10, 3, 4)
    batch_update(sheets_svc, spreadsheet_id, reqs)


# ─────────────────────── Spreadsheet orchestration ─────────────────────────────

TAB_NAMES = [
    "Dashboard",
    "Shopify Daily Data",
    "Meta Ads Daily Data",
    "Product Performance",
    "Campaign Performance",
    "Recommendations & Notes",
]


def get_or_create_spreadsheet(gc: gspread.Client, name: str) -> gspread.Spreadsheet:
    try:
        return gc.open(name)
    except gspread.SpreadsheetNotFound:
        ss = gc.create(name)
        ss.share(GMAIL_TO, perm_type="user", role="writer")
        return ss


def ensure_tabs(ss: gspread.Spreadsheet) -> dict:
    existing = {ws.title: ws for ws in ss.worksheets()}
    tab_map = {}
    for tab in TAB_NAMES:
        if tab not in existing:
            ws = ss.add_worksheet(title=tab, rows=1000, cols=30)
        else:
            ws = existing[tab]
        tab_map[tab] = ws
    if "Sheet1" in existing:
        try:
            ss.del_worksheet(existing["Sheet1"])
        except Exception:
            pass
    try:
        ss.reorder_worksheets([tab_map[t] for t in TAB_NAMES if t in tab_map])
    except Exception:
        pass
    return tab_map


def run_dashboard(date_str: str,
                  shopify_today: dict,
                  timeseries_data: dict,
                  shopify_prod_today: dict,
                  shopify_prod_7d: dict,
                  meta_campaigns: list,
                  meta_adsets: list,
                  shopify_7d_avg: dict,
                  meta_summary: dict,
                  recs: list,
                  unavailable: list) -> tuple[str, object]:
    gc, sheets_svc, gmail_svc = get_google_clients()
    ss = get_or_create_spreadsheet(gc, SPREADSHEET_NAME)
    tab_map  = ensure_tabs(ss)
    sheet_ids = {ws.title: ws.id for ws in ss.worksheets()}

    build_shopify_daily_tab(
        tab_map["Shopify Daily Data"], sheets_svc, ss.id,
        sheet_ids["Shopify Daily Data"], timeseries_data, date_str,
    )
    build_meta_daily_tab(
        tab_map["Meta Ads Daily Data"], sheets_svc, ss.id,
        sheet_ids["Meta Ads Daily Data"], meta_summary, date_str,
    )
    build_product_tab(
        tab_map["Product Performance"], sheets_svc, ss.id,
        sheet_ids["Product Performance"], shopify_prod_today, shopify_prod_7d,
    )
    build_campaign_tab(
        tab_map["Campaign Performance"], sheets_svc, ss.id,
        sheet_ids["Campaign Performance"], meta_campaigns, meta_adsets,
    )
    build_recommendations_tab(
        tab_map["Recommendations & Notes"], sheets_svc, ss.id,
        sheet_ids["Recommendations & Notes"], recs, date_str, unavailable,
    )
    build_dashboard_tab(
        tab_map["Dashboard"], sheets_svc, ss.id,
        sheet_ids["Dashboard"], shopify_today, shopify_7d_avg,
        meta_summary, recs, date_str,
    )
    return ss.url, gmail_svc


# ─────────────────────── Email ─────────────────────────────────────────────────

def send_or_draft_email(gmail_svc, sheet_url: str, date_str: str,
                         shopify_today: dict, shopify_7d_avg: dict,
                         meta_summary: dict, recs: list):
    import base64
    from email.mime.text import MIMEText

    s  = shopify_today.get("sales", {})
    gs = safe_float(s.get("gross_sales", 0))
    ns = safe_float(s.get("net_sales", 0))
    od = int(safe_float(s.get("orders", 0)))
    ao = safe_float(s.get("average_order_value", 0))
    rt = abs(safe_float(s.get("returns", 0)))

    def fmt(v): return f"₹{v:,.0f}"
    def pc(v):
        if v is None: return "—"
        return (f"+{v:.1f}%" if v >= 0 else f"{v:.1f}%")

    ag_gs = shopify_7d_avg.get("gross_sales", 0)
    ag_ns = shopify_7d_avg.get("net_sales", 0)
    ag_od = shopify_7d_avg.get("orders", 0)
    sp   = meta_summary.get("spend", 0)
    pu   = meta_summary.get("purchases", 0)
    cpa  = meta_summary.get("cpa", 0)
    roas = meta_summary.get("roas", 0)

    urgent = [r for r in recs if r.startswith("🚨")]
    urgent_block = ""
    if urgent:
        urgent_block = (
            "\n🚨 URGENT ISSUES\n"
            + "\n".join(f"  • {i}" for i in urgent)
            + "\n\n"
        )

    body = f"""\
Hi Atul,
{urgent_block}
Daily performance summary for {date_str}:

━━━ SHOPIFY ━━━
  Gross Sales:     {fmt(gs)}   ({pc(pct_change(gs, ag_gs))} vs 7-day avg {fmt(ag_gs)})
  Net Sales:       {fmt(ns)}   ({pc(pct_change(ns, ag_ns))} vs 7-day avg {fmt(ag_ns)})
  Orders:          {od}         ({pc(pct_change(od, ag_od))} vs 7-day avg {round(ag_od, 1)})
  Avg Order Value: {fmt(ao)}
  Returns:         {fmt(rt)}

━━━ META ADS ━━━
  Total Spend:     {fmt(sp)}
  Purchases:       {pu}
  CPA:             {fmt(cpa)}
  Blended ROAS:    {roas:.2f}x

━━━ RECOMMENDED ACTIONS ━━━
{chr(10).join(f"  {i+1}. {r}" for i, r in enumerate(recs[:5]))}

━━━ FULL DASHBOARD ━━━
{sheet_url}

Best,
Dhirai Daily Bot
"""
    msg = MIMEText(body, "plain")
    msg["To"]      = GMAIL_TO
    msg["Subject"] = f"Daily Store & Ads Performance Sheet - {date_str}"
    if GMAIL_CC:
        msg["Cc"] = GMAIL_CC

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    try:
        gmail_svc.users().messages().send(userId="me", body={"raw": raw}).execute()
        print(f"✅ Email sent to {GMAIL_TO}")
    except HttpError as e:
        print(f"⚠️  Send failed ({e}), creating draft...")
        gmail_svc.users().drafts().create(
            userId="me", body={"message": {"raw": raw}}
        ).execute()
        print("✅ Gmail draft created.")


# ─────────────────────── Main ──────────────────────────────────────────────────

def main():
    yesterday = yesterday_kolkata()
    seven_ago = yesterday - datetime.timedelta(days=7)
    date_str  = yesterday.isoformat()
    since_7d  = seven_ago.isoformat()
    unavailable = []

    print(f"📅 Running dashboard for {date_str} (Asia/Kolkata)")

    print("🛒 Fetching Shopify data...")
    try:
        shopify_today = fetch_shopify_daily(date_str)
    except Exception as e:
        print(f"  ⚠️ Shopify daily fetch failed: {e}")
        shopify_today = {"sales": {}, "sessions": {}, "top_products": {"rows": [], "columns": []}}
        unavailable.append(f"Shopify daily data unavailable: {e}")

    try:
        timeseries_data = fetch_shopify_8day_timeseries(since_7d, date_str)
    except Exception as e:
        print(f"  ⚠️ Shopify timeseries fetch failed: {e}")
        timeseries_data = {"sales": {"rows": [], "columns": []},
                           "sessions": {"rows": [], "columns": []}}
        unavailable.append(f"Shopify timeseries unavailable: {e}")

    try:
        shopify_prod_7d = fetch_shopify_7day_products(since_7d, date_str)
    except Exception as e:
        shopify_prod_7d = {"rows": [], "columns": []}
        unavailable.append(f"Shopify 7-day products unavailable: {e}")

    print("📣 Fetching Meta Ads data...")
    try:
        meta_campaigns = fetch_meta_level(date_str, "campaign")
    except Exception as e:
        print(f"  ⚠️ Meta campaigns fetch failed: {e}")
        meta_campaigns = []
        unavailable.append(f"Meta campaign data unavailable: {e}")

    try:
        meta_adsets = fetch_meta_level(date_str, "adset")
    except Exception as e:
        print(f"  ⚠️ Meta ad sets fetch failed: {e}")
        meta_adsets = []
        unavailable.append(f"Meta ad set data unavailable: {e}")

    # 7-day averages from timeseries sales data, excluding today
    print("💡 Computing 7-day averages...")
    sales_rows = timeseries_data["sales"].get("rows", [])
    sales_cols = timeseries_data["sales"].get("columns", [])
    sci = {c: i for i, c in enumerate(sales_cols)}
    prior_rows = [r for r in sales_rows if str(r[sci.get("day", 0)]) != date_str]
    shopify_7d_avg = compute_7day_averages(prior_rows, {
        "gross_sales":         sci.get("gross_sales", 1),
        "net_sales":           sci.get("net_sales", 2),
        "orders":              sci.get("orders", 3),
        "average_order_value": sci.get("average_order_value", 4),
    })

    meta_summary = aggregate_meta_campaigns(meta_campaigns)
    recs, notes  = generate_recommendations(shopify_today, shopify_7d_avg, meta_campaigns)

    print("📊 Updating Google Sheet...")
    try:
        sheet_url, gmail_svc = run_dashboard(
            date_str, shopify_today, timeseries_data,
            shopify_today.get("top_products", {"rows": [], "columns": []}),
            shopify_prod_7d, meta_campaigns, meta_adsets,
            shopify_7d_avg, meta_summary, recs, unavailable,
        )
        print(f"  ✅ Sheet updated: {sheet_url}")
    except Exception as e:
        print(f"  ⚠️ Sheet update failed: {e}")
        sheet_url = "https://docs.google.com/spreadsheets (configure GOOGLE_SA_CREDENTIALS)"
        try:
            _, _, gmail_svc = get_google_clients()
        except Exception as auth_err:
            print(f"  ⚠️ Gmail auth also failed: {auth_err}")
            return

    print("✉️  Sending email / creating draft...")
    try:
        send_or_draft_email(
            gmail_svc, sheet_url, date_str,
            shopify_today, shopify_7d_avg, meta_summary, recs,
        )
    except Exception as e:
        print(f"  ⚠️ Email step failed: {e}")

    print("✅ Done.")


if __name__ == "__main__":
    main()
