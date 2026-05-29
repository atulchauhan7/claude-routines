#!/usr/bin/env python3
"""
Daily Store & Ads Performance Dashboard
Runs every morning in Asia/Kolkata timezone.
Shopify store: 36dhns-ed.myshopify.com (Dhirai)
Meta Ads account: 979830497515712 (Dhirai)

Requirements:
    pip install -r requirements.txt

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
import re
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
    "https://www.googleapis.com/auth/drive",
]

SHOPIFY_DOMAIN   = os.environ.get("SHOPIFY_STORE_DOMAIN", "36dhns-ed.myshopify.com")
SHOPIFY_TOKEN    = os.environ.get("SHOPIFY_ACCESS_TOKEN", "")
META_TOKEN       = os.environ.get("META_ACCESS_TOKEN", "")
META_ACCOUNT_ID  = os.environ.get("META_AD_ACCOUNT_ID", "979830497515712")
SA_CREDENTIALS   = os.environ.get("GOOGLE_SA_CREDENTIALS", "service_account.json")
GMAIL_TO         = os.environ.get("GMAIL_TO", "atul012001@gmail.com")
GMAIL_CC         = os.environ.get("GMAIL_CC", "")
SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME", "Daily Store & Ads Performance Sheet")

# ROAS thresholds (D2C apparel, INR)
ROAS_DANGER   = 1.5   # below → pause/review
ROAS_WARNING  = 2.5   # below → optimise
ROAS_TARGET   = 3.0   # above → healthy
ROAS_SCALE    = 4.0   # above → scale aggressively

# ─────────────────────── Colours ──────────────────────────────────────────────

C_HEADER_BG  = {"red": 0.133, "green": 0.133, "blue": 0.133}
C_HEADER_FG  = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
C_SECTION_BG = {"red": 0.957, "green": 0.957, "blue": 0.957}
C_GREEN      = {"red": 0.714, "green": 0.929, "blue": 0.714}
C_RED        = {"red": 0.957, "green": 0.714, "blue": 0.714}
C_YELLOW     = {"red": 1.0,   "green": 0.949, "blue": 0.8}
C_ACCENT     = {"red": 0.259, "green": 0.522, "blue": 0.957}
C_WHITE      = {"red": 1.0,   "green": 1.0,   "blue": 1.0}

# ─────────────────────── Google Auth ──────────────────────────────────────────

def get_google_clients():
    creds = Credentials.from_service_account_file(SA_CREDENTIALS, scopes=SCOPES)
    gc            = gspread.authorize(creds)
    sheets_svc    = build("sheets", "v4", credentials=creds)
    gmail_svc     = build("gmail",  "v1", credentials=creds)
    return gc, sheets_svc, gmail_svc

# ─────────────────────── Date helpers ─────────────────────────────────────────

def yesterday_kolkata() -> datetime.date:
    return (datetime.datetime.now(KOLKATA_TZ) - datetime.timedelta(days=1)).date()

def date_range_7d(yesterday: datetime.date):
    """Returns (since_str, until_str) for the 7 days *before* yesterday."""
    until  = yesterday - datetime.timedelta(days=1)
    since  = until     - datetime.timedelta(days=6)
    return since.isoformat(), until.isoformat()

# ─────────────────────── Shopify data fetching ────────────────────────────────

SHOPIFY_GRAPHQL_URL = f"https://{SHOPIFY_DOMAIN}/admin/api/2024-04/graphql.json"

def shopify_analytics(shopify_query: str) -> dict:
    """Run a ShopifyQL analytics query; return {columns, rows}."""
    gql = """
    query RunAnalytics($query: String!) {
        shopifyqlQuery(query: $query) {
            ... on TableResponse {
                tableData {
                    columns { name dataType }
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
            "rows":    td["unformattedData"] or [],
        }
    except (KeyError, TypeError):
        return {"columns": [], "rows": [], "error": json.dumps(data)}


def fetch_shopify_sales_day(date_str: str) -> dict:
    q = (
        f"FROM sales SHOW gross_sales, net_sales, orders, average_order_value, "
        f"total_sales, discounts, returns, shipping_charges, taxes "
        f"SINCE {date_str} UNTIL {date_str}"
    )
    r = shopify_analytics(q)
    if r["rows"]:
        return dict(zip(r["columns"], r["rows"][0]))
    return {}


def fetch_shopify_sessions_day(date_str: str) -> dict:
    q = (
        f"FROM sessions SHOW sessions, sessions_with_cart_additions, "
        f"sessions_that_reached_checkout, sessions_that_completed_checkout, "
        f"conversion_rate SINCE {date_str} UNTIL {date_str}"
    )
    r = shopify_analytics(q)
    if r["rows"]:
        return dict(zip(r["columns"], r["rows"][0]))
    return {}


def fetch_shopify_top_products(date_str: str, limit: int = 15) -> dict:
    q = (
        f"FROM sales SHOW gross_sales, net_sales, orders, total_sales "
        f"GROUP BY product_title ORDER BY gross_sales DESC LIMIT {limit} "
        f"SINCE {date_str} UNTIL {date_str}"
    )
    return shopify_analytics(q)


def fetch_shopify_timeseries(since_str: str, until_str: str) -> dict:
    """Day-by-day sales for a date range."""
    q = (
        f"FROM sales SHOW gross_sales, net_sales, orders, average_order_value, "
        f"total_sales, discounts, returns, shipping_charges, taxes "
        f"TIMESERIES day SINCE {since_str} UNTIL {until_str}"
    )
    return shopify_analytics(q)


def fetch_shopify_sessions_timeseries(since_str: str, until_str: str) -> dict:
    q = (
        f"FROM sessions SHOW sessions, sessions_with_cart_additions, "
        f"sessions_that_reached_checkout, sessions_that_completed_checkout, "
        f"conversion_rate TIMESERIES day SINCE {since_str} UNTIL {until_str}"
    )
    return shopify_analytics(q)


def fetch_shopify_products_range(since_str: str, until_str: str, limit: int = 15) -> dict:
    q = (
        f"FROM sales SHOW gross_sales, net_sales, orders "
        f"GROUP BY product_title ORDER BY gross_sales DESC LIMIT {limit} "
        f"SINCE {since_str} UNTIL {until_str}"
    )
    return shopify_analytics(q)


# ─────────────────────── Meta Ads data fetching ────────────────────────────────

META_BASE = "https://graph.facebook.com/v19.0"

# Fields requested from Meta Marketing API
CAMPAIGN_FIELDS = (
    "campaign_id,campaign_name,spend,impressions,reach,clicks,"
    "ctr,cpc,cpm,purchase_roas,actions,cost_per_action_type,"
    "date_start,date_stop"
)
ADSET_FIELDS = (
    "campaign_id,campaign_name,adset_id,adset_name,spend,"
    "impressions,reach,clicks,ctr,cpc,cpm,"
    "purchase_roas,actions,cost_per_action_type,"
    "date_start,date_stop"
)
ACCOUNT_FIELDS = (
    "spend,impressions,reach,clicks,ctr,cpc,cpm,"
    "purchase_roas,actions,cost_per_action_type"
)


def meta_request(endpoint: str, params: dict) -> dict:
    params = {**params, "access_token": META_TOKEN}
    r = requests.get(f"{META_BASE}/{endpoint}", params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def _meta_insights(level: str, fields: str, time_range: dict,
                   time_increment: int = None) -> list:
    params = {
        "level":      level,
        "fields":     fields,
        "time_range": json.dumps(time_range),
        "limit":      50,
    }
    if time_increment:
        params["time_increment"] = time_increment
    r = meta_request(f"act_{META_ACCOUNT_ID}/insights", params)
    return r.get("data", [])


def fetch_meta_campaigns(date_str: str) -> list:
    return _meta_insights("campaign", CAMPAIGN_FIELDS,
                          {"since": date_str, "until": date_str})


def fetch_meta_adsets(date_str: str) -> list:
    return _meta_insights("adset", ADSET_FIELDS,
                          {"since": date_str, "until": date_str})


def fetch_meta_account_7d(since_str: str, until_str: str) -> list:
    """Daily account-level data for the 7-day comparison window."""
    return _meta_insights("account", ACCOUNT_FIELDS,
                          {"since": since_str, "until": until_str},
                          time_increment=1)


# ─────────────────────── Meta field parsers ────────────────────────────────────

PURCHASE_TYPES = {
    "offsite_conversion.fb_pixel_purchase",
    "purchase",
    "omni_purchase",
}


def extract_purchases(actions) -> int:
    if not actions:
        return 0
    total = 0
    for a in actions:
        if a.get("action_type") in PURCHASE_TYPES:
            try:
                total += int(float(a.get("value", 0)))
            except (ValueError, TypeError):
                pass
    return total


def extract_roas(purchase_roas) -> float:
    if not purchase_roas:
        return 0.0
    for r in (purchase_roas if isinstance(purchase_roas, list) else [purchase_roas]):
        if r.get("action_type") in PURCHASE_TYPES or r.get("action_type") == "omni_purchase":
            try:
                return round(float(r["value"]), 2)
            except (ValueError, KeyError):
                pass
    return 0.0


def extract_cpa(cost_per_action_type) -> float:
    if not cost_per_action_type:
        return 0.0
    for a in cost_per_action_type:
        if a.get("action_type") in PURCHASE_TYPES:
            try:
                return round(float(a["value"]), 2)
            except (ValueError, KeyError):
                pass
    return 0.0


# ─────────────────────── Computation helpers ──────────────────────────────────

def safe_float(v) -> float:
    if v is None:
        return 0.0
    s = str(v).replace(",", "").replace("₹", "").replace("INR", "").replace("%", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def pct_change(current: float, prior: float) -> float:
    if prior == 0:
        return 0.0
    return round((current - prior) / prior * 100, 1)


def avg_list(values: list) -> float:
    vals = [v for v in values if v is not None]
    return sum(vals) / len(vals) if vals else 0.0


def col_index(columns: list) -> dict:
    return {c: i for i, c in enumerate(columns)}


def compute_7d_avg(rows: list, cols: list, metrics: list) -> dict:
    """Compute per-column averages from a list of rows."""
    ci = col_index(cols)
    result = {}
    for m in metrics:
        idx = ci.get(m)
        if idx is None:
            result[m] = 0.0
            continue
        vals = []
        for row in rows:
            try:
                vals.append(float(row[idx]))
            except (ValueError, TypeError, IndexError):
                pass
        result[m] = round(avg_list(vals), 2)
    return result


# ─────────────────────── Recommendations engine ────────────────────────────────

def generate_recommendations(
        shopify_sales: dict, shopify_7d_avg: dict,
        sessions: dict,
        meta_campaigns: list, meta_adsets: list) -> tuple[list, list]:

    recs   = []
    notes  = []

    # ── Revenue ──────────────────────────────────────────────
    gs     = safe_float(shopify_sales.get("gross_sales", 0))
    avg_gs = shopify_7d_avg.get("gross_sales", 0)
    if avg_gs > 0:
        chg = pct_change(gs, avg_gs)
        if chg < -20:
            recs.append(f"🚨 Revenue dropped {abs(chg):.0f}% vs 7-day avg — check Meta ad delivery, site uptime, or payment gateway.")
        elif chg < -10:
            recs.append(f"⚠️ Revenue down {abs(chg):.0f}% vs 7-day avg — review campaign budgets and top product pages.")
        elif chg > 20:
            recs.append(f"✅ Revenue up {chg:.0f}% vs 7-day avg — identify winning campaigns and consider scaling budgets.")

    # ── Orders ───────────────────────────────────────────────
    orders     = safe_float(shopify_sales.get("orders", 0))
    avg_orders = shopify_7d_avg.get("orders", 0)
    if avg_orders > 0 and orders < avg_orders * 0.80:
        recs.append("⚠️ Orders 20%+ below 7-day avg — review traffic sources, check Meta delivery.")

    # ── Returns ──────────────────────────────────────────────
    returns_val = abs(safe_float(shopify_sales.get("returns", 0)))
    if gs > 0:
        ret_rate = returns_val / gs
        if ret_rate > 0.15:
            recs.append(f"🚨 High return rate ({ret_rate*100:.0f}% of gross) — audit product descriptions, sizing guides, and quality.")
        elif ret_rate > 0.10:
            recs.append(f"⚠️ Return rate {ret_rate*100:.0f}% of gross — monitor closely; check recent product or packaging changes.")

    # ── Checkout abandonment ──────────────────────────────────
    sess       = safe_float(sessions.get("sessions", 0))
    cart_adds  = safe_float(sessions.get("sessions_with_cart_additions", 0))
    at_checkout= safe_float(sessions.get("sessions_that_reached_checkout", 0))
    completed  = safe_float(sessions.get("sessions_that_completed_checkout", 0))
    if cart_adds > 0 and at_checkout > 0:
        abandon = 1 - (completed / at_checkout) if at_checkout else 1
        if abandon > 0.7:
            recs.append(f"⚠️ ~{abandon*100:.0f}% checkout abandonment — activate/review abandoned cart email & SMS recovery flows.")

    # ── Meta campaign analysis ────────────────────────────────
    for c in sorted(meta_campaigns, key=lambda x: safe_float(x.get("spend", 0)), reverse=True):
        spend = safe_float(c.get("spend", 0))
        if spend < 200:
            continue
        roas  = extract_roas(c.get("purchase_roas"))
        purch = extract_purchases(c.get("actions", []))
        name  = c.get("campaign_name", "Unknown")

        if roas == 0 and purch == 0 and spend > 800:
            recs.append(f"🚨 PAUSE '{name}' — ₹{spend:,.0f} spent, 0 purchases. Budget is wasted.")
        elif 0 < roas < ROAS_DANGER:
            recs.append(f"🚨 PAUSE '{name}' — ROAS {roas:.2f}x below {ROAS_DANGER}x (₹{spend:,.0f} spend, ₹{spend/max(purch,1):,.0f} CPA). Reallocate budget.")
        elif ROAS_DANGER <= roas < ROAS_WARNING:
            recs.append(f"⚠️ Review '{name}' — ROAS {roas:.2f}x below target {ROAS_TARGET}x. Test new creatives or tighten audience.")
        elif roas >= ROAS_SCALE and spend > 1000:
            recs.append(f"✅ Scale '{name}' — ROAS {roas:.2f}x is excellent. Increase budget by 20–30%.")

    # ── Ad-set analysis ───────────────────────────────────────
    for a in sorted(meta_adsets, key=lambda x: safe_float(x.get("spend", 0)), reverse=True):
        spend = safe_float(a.get("spend", 0))
        if spend < 500:
            continue
        ctr  = safe_float(a.get("ctr", 0))
        name = a.get("adset_name", a.get("name", "Unknown"))
        if ctr < 0.8:
            recs.append(f"⚠️ Low CTR {ctr:.2f}% on ad set '{name}' — refresh creative, test new hook/thumbnail.")

    # ── Inventory reminder ────────────────────────────────────
    recs.append("📱 Verify inventory levels for top-selling products — restock before stockouts occur.")

    # ── Abandonment recovery note ─────────────────────────────
    if at_checkout == 0 or completed == 0:
        notes.append("Shopify checkout completion data returned 0 — verify Pixel/tracking setup if expected otherwise.")

    if not any(r.startswith(("🚨", "⚠️", "✅")) for r in recs):
        recs.insert(0, "✅ All key metrics within normal range — continue monitoring.")

    return recs[:8], notes


# ─────────────────────── Google Sheets helpers ────────────────────────────────

def _range_spec(sheet_id: int, r0: int, r1: int, c0: int, c1: int) -> dict:
    return {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1,
            "startColumnIndex": c0, "endColumnIndex": c1}


def cell_format_req(sheet_id: int, r0: int, r1: int, c0: int, c1: int,
                    bold: bool = False, bg: dict = None, fg: dict = None,
                    align: str = None, num_pattern: str = None,
                    font_size: int = None, wrap: str = None) -> dict:
    fmt   = {}
    tf    = {}
    if bold:       tf["bold"]            = True
    if fg:         tf["foregroundColor"] = fg
    if font_size:  tf["fontSize"]        = font_size
    if tf:         fmt["textFormat"]     = tf
    if bg:         fmt["backgroundColor"] = bg
    if align:      fmt["horizontalAlignment"] = align
    if num_pattern:
        fmt["numberFormat"] = {"type": "NUMBER", "pattern": num_pattern}
    if wrap:
        fmt["wrapStrategy"] = wrap

    fields_parts = []
    if tf:          fields_parts.append("textFormat")
    if bg:          fields_parts.append("backgroundColor")
    if align:       fields_parts.append("horizontalAlignment")
    if num_pattern: fields_parts.append("numberFormat")
    if wrap:        fields_parts.append("wrapStrategy")

    return {
        "repeatCell": {
            "range": _range_spec(sheet_id, r0, r1, c0, c1),
            "cell": {"userEnteredFormat": fmt},
            "fields": "userEnteredFormat(" + ",".join(fields_parts) + ")",
        }
    }


def freeze_req(sheet_id: int, rows: int = 1, cols: int = 0) -> dict:
    return {
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols},
            },
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
        }
    }


def auto_resize_req(sheet_id: int, c0: int = 0, c1: int = 30) -> dict:
    return {
        "autoResizeDimensions": {
            "dimensions": {"sheetId": sheet_id, "dimension": "COLUMNS",
                           "startIndex": c0, "endIndex": c1}
        }
    }


def filter_req(sheet_id: int, r0: int, r1: int, c0: int, c1: int) -> dict:
    return {
        "setBasicFilter": {
            "filter": {"range": _range_spec(sheet_id, r0, r1, c0, c1)}
        }
    }


def cond_pct_req(sheet_id: int, r0: int, r1: int, c0: int, c1: int) -> list:
    """Green for positive, red for negative percentage change."""
    rng = [_range_spec(sheet_id, r0, r1, c0, c1)]
    return [
        {"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": rng,
            "booleanRule": {
                "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0"}]},
                "format": {"backgroundColor": C_GREEN},
            }
        }}},
        {"addConditionalFormatRule": {"index": 1, "rule": {
            "ranges": rng,
            "booleanRule": {
                "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                "format": {"backgroundColor": C_RED},
            }
        }}},
    ]


def cond_roas_req(sheet_id: int, r0: int, r1: int, c0: int, c1: int) -> list:
    """Red < ROAS_DANGER, yellow < ROAS_TARGET, green >= ROAS_SCALE."""
    rng = [_range_spec(sheet_id, r0, r1, c0, c1)]
    return [
        {"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": rng,
            "booleanRule": {
                "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": str(ROAS_DANGER)}]},
                "format": {"backgroundColor": C_RED},
            }
        }}},
        {"addConditionalFormatRule": {"index": 1, "rule": {
            "ranges": rng,
            "booleanRule": {
                "condition": {
                    "type": "NUMBER_BETWEEN",
                    "values": [{"userEnteredValue": str(ROAS_DANGER)},
                               {"userEnteredValue": str(ROAS_TARGET)}],
                },
                "format": {"backgroundColor": C_YELLOW},
            }
        }}},
        {"addConditionalFormatRule": {"index": 2, "rule": {
            "ranges": rng,
            "booleanRule": {
                "condition": {"type": "NUMBER_GREATER_THAN_EQ",
                              "values": [{"userEnteredValue": str(ROAS_SCALE)}]},
                "format": {"backgroundColor": C_GREEN},
            }
        }}},
    ]


def batch_update(sheets_svc, spreadsheet_id: str, reqs: list):
    if reqs:
        sheets_svc.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": reqs}
        ).execute()


# ─────────────────────── Sheet tab builders ────────────────────────────────────

INR_FMT   = '[$₹-4409]#,##0.00'
PCT_FMT   = '0.00"%"'
INT_FMT   = '#,##0'


def build_shopify_daily_tab(ws, sheets_svc, ss_id: str, sid: int,
                             timeseries: dict, sessions_ts: dict,
                             yesterday_str: str):
    ws.clear()
    headers = [
        "Date", "Gross Sales (₹)", "Net Sales (₹)", "Orders",
        "AOV (₹)", "Discounts (₹)", "Returns (₹)", "Shipping (₹)", "Taxes (₹)",
        "Sessions", "Cart Adds", "Reached Checkout", "Completed Checkout",
        "Conversion Rate (%)",
    ]
    ws.update("A1", [headers])

    ts_rows = timeseries.get("rows", [])
    ts_cols = timeseries.get("columns", [])
    ci = col_index(ts_cols)

    # Build session lookup by date
    sess_rows = sessions_ts.get("rows", [])
    sess_cols = sessions_ts.get("columns", [])
    ssi = col_index(sess_cols)
    sess_by_date = {}
    for row in sess_rows:
        d = str(row[ssi.get("day", 0)])
        sess_by_date[d] = row

    # Read existing dates to avoid duplicates
    existing = ws.col_values(1)[1:] if len(ts_rows) == 0 else []

    data_rows = []
    for row in ts_rows:
        date_val = str(row[ci.get("day", 0)])
        if date_val in existing:
            continue
        sr = sess_by_date.get(date_val, [])
        data_rows.append([
            date_val,
            safe_float(row[ci.get("gross_sales", 1)]),
            safe_float(row[ci.get("net_sales",   2)]),
            int(safe_float(row[ci.get("orders",  3)])),
            round(safe_float(row[ci.get("average_order_value", 4)]), 2),
            abs(safe_float(row[ci.get("discounts", 5)])),
            abs(safe_float(row[ci.get("returns",   6)])),
            safe_float(row[ci.get("shipping_charges", 7)]),
            safe_float(row[ci.get("taxes",         8)]),
            int(safe_float(sr[ssi.get("sessions", 1)] if sr else 0)),
            int(safe_float(sr[ssi.get("sessions_with_cart_additions", 2)] if sr else 0)),
            int(safe_float(sr[ssi.get("sessions_that_reached_checkout", 3)] if sr else 0)),
            int(safe_float(sr[ssi.get("sessions_that_completed_checkout", 4)] if sr else 0)),
            round(safe_float(sr[ssi.get("conversion_rate", 5)] if sr else 0) * 100, 4),
        ])

    if data_rows:
        next_row = len(ws.col_values(1)) + 1
        ws.update(f"A{next_row}", data_rows)

    total = max(len(data_rows) + 1, 2)
    reqs = [
        cell_format_req(sid, 0, 1, 0, len(headers), bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG),
        freeze_req(sid, rows=1),
        filter_req(sid, 0, total + 1, 0, len(headers)),
        auto_resize_req(sid, 0, len(headers)),
    ]
    for col in [1, 2, 4, 5, 6, 7, 8]:
        reqs.append(cell_format_req(sid, 1, total + 1, col, col + 1, num_pattern=INR_FMT))
    reqs.append(cell_format_req(sid, 1, total + 1, 13, 14, num_pattern=PCT_FMT))
    for col in [3, 9, 10, 11, 12]:
        reqs.append(cell_format_req(sid, 1, total + 1, col, col + 1, num_pattern=INT_FMT))
    batch_update(sheets_svc, ss_id, reqs)


def build_meta_daily_tab(ws, sheets_svc, ss_id: str, sid: int,
                          meta_7d: list, yesterday_summary: dict,
                          yesterday_str: str):
    ws.clear()
    headers = [
        "Date", "Spend (₹)", "Impressions", "Reach", "Clicks",
        "CTR (%)", "CPC (₹)", "CPM (₹)", "Purchases", "CPA (₹)", "ROAS",
    ]
    ws.update("A1", [headers])

    all_rows = []
    # Historical 7-day rows
    for d in meta_7d:
        date_val  = d.get("date_start", "")
        spend     = safe_float(d.get("spend", 0))
        impr      = int(safe_float(d.get("impressions", 0)))
        reach     = int(safe_float(d.get("reach", 0)))
        clicks    = int(safe_float(d.get("clicks", 0)))
        purch     = extract_purchases(d.get("actions", []))
        ctr       = safe_float(d.get("ctr", 0))
        cpc       = safe_float(d.get("cpc", 0))
        cpm       = safe_float(d.get("cpm", 0))
        cpa       = round(spend / purch, 2) if purch else 0
        roas      = extract_roas(d.get("purchase_roas"))
        all_rows.append([date_val, round(spend, 2), impr, reach, clicks,
                         ctr, cpc, cpm, purch, cpa, roas])

    # Check if yesterday row already exists
    existing_dates = {r[0] for r in all_rows}
    if yesterday_str not in existing_dates:
        s = yesterday_summary
        all_rows.append([
            yesterday_str,
            round(s.get("spend", 0), 2),
            s.get("impressions", 0),
            s.get("reach", 0),
            s.get("clicks", 0),
            round(s.get("ctr", 0), 2),
            round(s.get("cpc", 0), 2),
            round(s.get("cpm", 0), 2),
            s.get("purchases", 0),
            round(s.get("cpa", 0), 2),
            round(s.get("roas", 0), 2),
        ])

    all_rows.sort(key=lambda r: r[0])
    if all_rows:
        ws.update("A2", all_rows)

    total = len(all_rows) + 2
    reqs = [
        cell_format_req(sid, 0, 1, 0, len(headers), bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG),
        freeze_req(sid, rows=1),
        filter_req(sid, 0, total, 0, len(headers)),
        auto_resize_req(sid, 0, len(headers)),
    ]
    for col in [1, 6, 7, 9]:
        reqs.append(cell_format_req(sid, 1, total, col, col + 1, num_pattern=INR_FMT))
    reqs.append(cell_format_req(sid, 1, total, 5, 6, num_pattern=PCT_FMT))
    for col in [2, 3, 4, 8]:
        reqs.append(cell_format_req(sid, 1, total, col, col + 1, num_pattern=INT_FMT))
    reqs += cond_roas_req(sid, 1, total, 10, 11)
    batch_update(sheets_svc, ss_id, reqs)


def build_product_tab(ws, sheets_svc, ss_id: str, sid: int,
                       products_today: dict, products_7d: dict):
    ws.clear()
    header1 = ["─── Yesterday ───", "", "", "", "─── 7-Day Total ───", "", ""]
    header2 = ["Product", "Gross Sales (₹)", "Net Sales (₹)", "Orders",
               "7D Gross (₹)", "7D Net (₹)", "7D Orders"]
    ws.update("A1", [header1, header2])

    today_rows = products_today.get("rows", [])
    seven_rows = products_7d.get("rows", [])
    seven_map  = {str(r[0]): r for r in seven_rows}

    data = []
    for row in today_rows:
        prod = str(row[0])
        s7   = seven_map.get(prod, [])
        data.append([
            prod,
            safe_float(row[1]),
            safe_float(row[2]),
            int(safe_float(row[3])),
            safe_float(s7[1]) if len(s7) > 1 else "",
            safe_float(s7[2]) if len(s7) > 2 else "",
            int(safe_float(s7[3])) if len(s7) > 3 else "",
        ])

    if data:
        ws.update("A3", data)

    total = len(data) + 3
    reqs = [
        cell_format_req(sid, 0, 1, 0, 7, bold=True, bg=C_ACCENT, fg=C_HEADER_FG),
        cell_format_req(sid, 1, 2, 0, 7, bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG),
        freeze_req(sid, rows=2),
        filter_req(sid, 1, total, 0, 7),
        auto_resize_req(sid, 0, 7),
    ]
    for col in [1, 2, 4, 5]:
        reqs.append(cell_format_req(sid, 2, total, col, col + 1, num_pattern=INR_FMT))
    for col in [3, 6]:
        reqs.append(cell_format_req(sid, 2, total, col, col + 1, num_pattern=INT_FMT))
    batch_update(sheets_svc, ss_id, reqs)


def build_campaign_tab(ws, sheets_svc, ss_id: str, sid: int,
                        campaigns: list, adsets: list):
    ws.clear()
    camp_hdr = [
        "Campaign", "Status", "Spend (₹)", "Impressions", "Reach", "Clicks",
        "CTR (%)", "CPC (₹)", "CPM (₹)", "Purchases", "CPA (₹)", "ROAS",
    ]
    ws.update("A1", [["━━━ CAMPAIGN LEVEL — YESTERDAY ━━━"] + [""] * (len(camp_hdr) - 1),
                     camp_hdr])

    camp_data = []
    for c in campaigns:
        spend = safe_float(c.get("spend", 0))
        if spend == 0:
            continue
        camp_data.append([
            c.get("campaign_name", ""),
            c.get("effective_status", c.get("status", "")),
            round(spend, 2),
            int(safe_float(c.get("impressions", 0))),
            int(safe_float(c.get("reach", 0))),
            int(safe_float(c.get("clicks", 0))),
            round(safe_float(c.get("ctr", 0)), 2),
            round(safe_float(c.get("cpc", 0)), 2),
            round(safe_float(c.get("cpm", 0)), 2),
            extract_purchases(c.get("actions", [])),
            round(extract_cpa(c.get("cost_per_action_type")), 2),
            round(extract_roas(c.get("purchase_roas")), 2),
        ])
    camp_data.sort(key=lambda r: r[2], reverse=True)
    if camp_data:
        ws.update("A3", camp_data)
    camp_end = len(camp_data) + 3

    adset_hdr = [
        "Ad Set", "Status", "Spend (₹)", "Impressions", "Reach", "Clicks",
        "CTR (%)", "CPC (₹)", "CPM (₹)", "Purchases", "CPA (₹)", "ROAS",
    ]
    as_start = camp_end + 2
    ws.update(f"A{as_start}", [["━━━ AD SET LEVEL — YESTERDAY ━━━"] + [""] * (len(adset_hdr) - 1),
                                adset_hdr])

    as_data = []
    for a in adsets:
        spend = safe_float(a.get("spend", 0))
        if spend == 0:
            continue
        as_data.append([
            a.get("adset_name", a.get("name", "")),
            a.get("effective_status", a.get("status", "")),
            round(spend, 2),
            int(safe_float(a.get("impressions", 0))),
            int(safe_float(a.get("reach", 0))),
            int(safe_float(a.get("clicks", 0))),
            round(safe_float(a.get("ctr", 0)), 2),
            round(safe_float(a.get("cpc", 0)), 2),
            round(safe_float(a.get("cpm", 0)), 2),
            extract_purchases(a.get("actions", [])),
            round(extract_cpa(a.get("cost_per_action_type")), 2),
            round(extract_roas(a.get("purchase_roas")), 2),
        ])
    as_data.sort(key=lambda r: r[2], reverse=True)
    if as_data:
        ws.update(f"A{as_start + 2}", as_data)
    as_end = as_start + 2 + len(as_data)

    reqs = [
        cell_format_req(sid, 0, 1, 0, len(camp_hdr), bold=True, bg=C_ACCENT, fg=C_HEADER_FG, font_size=11),
        cell_format_req(sid, 1, 2, 0, len(camp_hdr), bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG),
        cell_format_req(sid, as_start-1, as_start, 0, len(adset_hdr), bold=True, bg=C_ACCENT, fg=C_HEADER_FG, font_size=11),
        cell_format_req(sid, as_start,   as_start+1, 0, len(adset_hdr), bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG),
        freeze_req(sid, rows=2),
        auto_resize_req(sid, 0, len(adset_hdr)),
    ]
    for section_start, section_end in [(2, camp_end), (as_start+1, as_end)]:
        for col in [2, 7, 8, 10]:
            reqs.append(cell_format_req(sid, section_start, section_end, col, col+1, num_pattern=INR_FMT))
        for col in [3, 4, 5, 9]:
            reqs.append(cell_format_req(sid, section_start, section_end, col, col+1, num_pattern=INT_FMT))
        reqs.append(cell_format_req(sid, section_start, section_end, 6, 7, num_pattern=PCT_FMT))
        reqs += cond_roas_req(sid, section_start, section_end, 11, 12)
    batch_update(sheets_svc, ss_id, reqs)


def build_recs_tab(ws, sheets_svc, ss_id: str, sid: int,
                    recs: list, notes: list, date_str: str, unavailable: list = None):
    ws.clear()
    all_rows = [
        [f"RECOMMENDATIONS & NOTES — {date_str}", ""],
        ["", ""],
        ["TYPE", "RECOMMENDATION / NOTE"],
    ]
    for rec in recs:
        all_rows.append(["→", rec])
    if notes:
        all_rows += [["", ""], ["NOTES", ""]]
        for n in notes:
            all_rows.append(["ℹ", n])
    if unavailable:
        all_rows += [["", ""], ["UNAVAILABLE METRICS", ""]]
        for m in unavailable:
            all_rows.append(["—", m])

    ws.update("A1", all_rows)
    n = len(all_rows)
    reqs = [
        cell_format_req(sid, 0, 1, 0, 2, bold=True, bg=C_ACCENT, fg=C_HEADER_FG, font_size=13),
        cell_format_req(sid, 2, 3, 0, 2, bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG),
        freeze_req(sid, rows=3),
        auto_resize_req(sid, 0, 2),
        cell_format_req(sid, 3, n, 1, 2, wrap="WRAP"),
    ]
    batch_update(sheets_svc, ss_id, reqs)


def _fmt_inr(v: float) -> str:
    return f"₹{v:,.0f}"

def _fmt_pct(v: float) -> str:
    if v is None or v == 0:
        return "—"
    return f"+{v:.1f}%" if v >= 0 else f"{v:.1f}%"


def build_dashboard_tab(ws, sheets_svc, ss_id: str, sid: int,
                         shopify_sales: dict, shopify_7d_avg: dict,
                         sessions: dict, meta_sum: dict,
                         meta_7d_avg: dict,
                         recs: list, date_str: str,
                         top_products: list):
    ws.clear()

    gs  = safe_float(shopify_sales.get("gross_sales", 0))
    ns  = safe_float(shopify_sales.get("net_sales",   0))
    od  = int(safe_float(shopify_sales.get("orders",  0)))
    ao  = safe_float(shopify_sales.get("average_order_value", 0))
    rt  = abs(safe_float(shopify_sales.get("returns", 0)))

    ag_gs = shopify_7d_avg.get("gross_sales",         0)
    ag_ns = shopify_7d_avg.get("net_sales",           0)
    ag_od = shopify_7d_avg.get("orders",              0)
    ag_ao = shopify_7d_avg.get("average_order_value", 0)

    sp   = meta_sum.get("spend",       0)
    imp  = meta_sum.get("impressions", 0)
    cl   = meta_sum.get("clicks",      0)
    pu   = meta_sum.get("purchases",   0)
    ctr  = meta_sum.get("ctr",         0)
    cpa  = meta_sum.get("cpa",         0)
    roas = meta_sum.get("roas",        0)

    ag_sp   = meta_7d_avg.get("spend",       0)
    ag_imp  = meta_7d_avg.get("impressions", 0)
    ag_pu   = meta_7d_avg.get("purchases",   0)
    ag_roas = meta_7d_avg.get("roas",        0)

    sess     = int(safe_float(sessions.get("sessions", 0)))
    cart     = int(safe_float(sessions.get("sessions_with_cart_additions", 0)))

    rows = [
        [f"DAILY STORE & ADS PERFORMANCE — DHIRAI | {date_str}", "", "", ""],
        ["", "", "", ""],
        ["━━━ SHOPIFY PERFORMANCE ━━━", "", "", ""],
        ["Metric", "Yesterday", "7-Day Avg", "Change vs Avg"],
        ["Gross Sales",      _fmt_inr(gs),  _fmt_inr(ag_gs), _fmt_pct(pct_change(gs,  ag_gs))],
        ["Net Sales",        _fmt_inr(ns),  _fmt_inr(ag_ns), _fmt_pct(pct_change(ns,  ag_ns))],
        ["Orders",           str(od),       str(round(ag_od, 1)), _fmt_pct(pct_change(od, ag_od))],
        ["Avg Order Value",  _fmt_inr(ao),  _fmt_inr(ag_ao), _fmt_pct(pct_change(ao,  ag_ao))],
        ["Returns",          _fmt_inr(rt),  "",              ""],
        ["Sessions",         f"{sess:,}",   "",              ""],
        ["Cart Additions",   f"{cart:,}",   "",              ""],
        ["", "", "", ""],
        ["━━━ META ADS PERFORMANCE ━━━", "", "", ""],
        ["Metric", "Yesterday", "7-Day Avg", "Change vs Avg"],
        ["Total Spend",      _fmt_inr(sp),  _fmt_inr(ag_sp),   _fmt_pct(pct_change(sp,   ag_sp))],
        ["Impressions",      f"{imp:,}",    f"{ag_imp:,.0f}",  _fmt_pct(pct_change(imp,  ag_imp))],
        ["Clicks",           f"{cl:,}",     "",                ""],
        ["CTR",              f"{ctr:.2f}%", "",                ""],
        ["Purchases",        str(pu),       str(round(ag_pu, 1)), _fmt_pct(pct_change(pu, ag_pu))],
        ["CPA",              _fmt_inr(cpa), "",                ""],
        ["Blended ROAS",     f"{roas:.2f}x",f"{ag_roas:.2f}x",_fmt_pct(pct_change(roas, ag_roas))],
        ["", "", "", ""],
    ]

    # Top products
    rows.append(["━━━ TOP PRODUCTS YESTERDAY ━━━", "", "", ""])
    rows.append(["Product", "Gross Sales", "Orders", "AOV"])
    for p in (top_products or [])[:8]:
        prod  = str(p[0])[:60]
        gross = _fmt_inr(safe_float(p[1]))
        ords  = int(safe_float(p[3])) if len(p) > 3 else ""
        aov   = _fmt_inr(safe_float(p[4])) if len(p) > 4 else ""
        rows.append([prod, gross, str(ords), aov])

    rows += [["", "", "", ""], ["━━━ WINS & ISSUES ━━━", "", "", ""]]
    wins   = [r for r in recs if r.startswith("✅")]
    issues = [r for r in recs if r.startswith(("🚨", "⚠️"))]
    if wins:
        rows.append(["✅ WINS", "", "", ""])
        for w in wins[:3]:
            rows.append(["", w, "", ""])
    if issues:
        rows.append(["🚨 ISSUES", "", "", ""])
        for i in issues[:3]:
            rows.append(["", i, "", ""])

    rows += [["", "", "", ""], ["━━━ RECOMMENDED ACTIONS TODAY ━━━", "", "", ""]]
    for idx, rec in enumerate(recs[:5], 1):
        rows.append([f"{idx}.", rec, "", ""])

    ws.update("A1", rows)

    # Section header row indices (0-based)
    section_rows = [2, 12, 21, len(rows) - len(recs[:5]) - 2]
    sub_hdr_rows = [3, 13, len(rows) - len(recs[:5]) - 1]

    reqs = [
        cell_format_req(sid, 0, 1, 0, 4, bold=True, bg=C_ACCENT, fg=C_HEADER_FG, font_size=14),
        freeze_req(sid, rows=1),
        auto_resize_req(sid, 0, 4),
    ]
    for sr in section_rows:
        if sr < len(rows):
            reqs.append(cell_format_req(sid, sr, sr+1, 0, 4,
                                        bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG, font_size=11))
    for sh in sub_hdr_rows:
        if sh < len(rows):
            reqs.append(cell_format_req(sid, sh, sh+1, 0, 4, bold=True, bg=C_SECTION_BG))
    # Pct-change column (D, index 3) — rows 4-11 (Shopify), 14-21 (Meta)
    reqs += cond_pct_req(sid, 4, 11, 3, 4)
    reqs += cond_pct_req(sid, 14, 21, 3, 4)
    batch_update(sheets_svc, ss_id, reqs)


# ─────────────────────── Spreadsheet orchestration ─────────────────────────────

TAB_NAMES = [
    "Dashboard",
    "Shopify Daily Data",
    "Meta Ads Daily Data",
    "Product Performance",
    "Campaign Performance",
    "Recommendations & Notes",
]


def get_or_create_spreadsheet(gc, name: str):
    try:
        return gc.open(name)
    except gspread.SpreadsheetNotFound:
        ss = gc.create(name)
        ss.share(GMAIL_TO, perm_type="user", role="writer")
        return ss


def ensure_tabs(ss) -> dict:
    existing = {ws.title: ws for ws in ss.worksheets()}
    tab_map  = {}
    for tab in TAB_NAMES:
        tab_map[tab] = existing.get(tab) or ss.add_worksheet(tab, rows=1000, cols=30)
    if "Sheet1" in existing and "Sheet1" not in TAB_NAMES:
        try:
            ss.del_worksheet(existing["Sheet1"])
        except Exception:
            pass
    try:
        ss.reorder_worksheets([tab_map[t] for t in TAB_NAMES])
    except Exception:
        pass
    return tab_map


def update_spreadsheet(yesterday: datetime.date,
                        shopify_sales: dict, shopify_ts: dict, shopify_sess_ts: dict,
                        prod_today: dict, prod_7d: dict,
                        sessions_today: dict,
                        meta_campaigns: list, meta_adsets: list, meta_7d: list,
                        recs: list, notes: list, unavailable: list) -> tuple:

    gc, sheets_svc, gmail_svc = get_google_clients()
    ss     = get_or_create_spreadsheet(gc, SPREADSHEET_NAME)
    tabs   = ensure_tabs(ss)
    sids   = {ws.title: ws.id for ws in ss.worksheets()}

    date_str  = yesterday.isoformat()
    since_7d, until_7d = date_range_7d(yesterday)

    # ── 7-day averages ──
    ts_rows   = shopify_ts.get("rows", [])
    ts_cols   = shopify_ts.get("columns", [])
    prev_rows = [r for r in ts_rows if str(r[0]) != date_str]
    shopify_7d_avg = compute_7d_avg(prev_rows, ts_cols,
                                     ["gross_sales", "net_sales", "orders", "average_order_value"])

    # ── Meta account-level summary for yesterday ──
    total_spend  = sum(safe_float(c.get("spend", 0))       for c in meta_campaigns)
    total_impr   = sum(int(safe_float(c.get("impressions", 0))) for c in meta_campaigns)
    total_reach  = sum(int(safe_float(c.get("reach", 0)))   for c in meta_campaigns)
    total_clicks = sum(int(safe_float(c.get("clicks", 0)))  for c in meta_campaigns)
    total_purch  = sum(extract_purchases(c.get("actions", [])) for c in meta_campaigns)
    ctr          = round(total_clicks / total_impr * 100, 2) if total_impr else 0
    cpc          = round(total_spend / total_clicks,   2)    if total_clicks else 0
    cpm          = round(total_spend / total_impr * 1000, 2) if total_impr else 0
    cpa          = round(total_spend / total_purch,    2)    if total_purch else 0
    roas_vals    = [extract_roas(c.get("purchase_roas")) for c in meta_campaigns]
    roas_vals    = [r for r in roas_vals if r > 0]
    blended_roas = round(avg_list(roas_vals), 2)
    meta_sum = {
        "spend": round(total_spend, 2),
        "impressions": total_impr, "reach": total_reach,
        "clicks": total_clicks, "purchases": total_purch,
        "ctr": ctr, "cpc": cpc, "cpm": cpm, "cpa": cpa, "roas": blended_roas,
    }

    # ── Meta 7-day averages ──
    meta_7d_avg = compute_7d_avg(
        meta_7d, ["date_start", "spend", "impressions", "reach",
                  "clicks", "ctr", "cpc", "cpm"],
        ["spend", "impressions", "reach", "clicks"]
    )
    meta_roas_vals = [extract_roas(d.get("purchase_roas")) for d in meta_7d]
    meta_7d_avg["roas"] = round(avg_list([r for r in meta_roas_vals if r > 0]), 2)
    meta_purch_vals = [extract_purchases(d.get("actions", [])) for d in meta_7d]
    meta_7d_avg["purchases"] = round(avg_list(meta_purch_vals), 2)

    # ── Build tabs ──
    build_shopify_daily_tab(
        tabs["Shopify Daily Data"], sheets_svc, ss.id, sids["Shopify Daily Data"],
        shopify_ts, shopify_sess_ts, date_str
    )
    build_meta_daily_tab(
        tabs["Meta Ads Daily Data"], sheets_svc, ss.id, sids["Meta Ads Daily Data"],
        meta_7d, meta_sum, date_str
    )
    build_product_tab(
        tabs["Product Performance"], sheets_svc, ss.id, sids["Product Performance"],
        prod_today, prod_7d
    )
    build_campaign_tab(
        tabs["Campaign Performance"], sheets_svc, ss.id, sids["Campaign Performance"],
        meta_campaigns, meta_adsets
    )
    build_recs_tab(
        tabs["Recommendations & Notes"], sheets_svc, ss.id, sids["Recommendations & Notes"],
        recs, notes, date_str, unavailable
    )
    build_dashboard_tab(
        tabs["Dashboard"], sheets_svc, ss.id, sids["Dashboard"],
        shopify_sales, shopify_7d_avg, sessions_today, meta_sum, meta_7d_avg,
        recs, date_str, prod_today.get("rows", [])
    )

    return ss.url, gmail_svc, meta_sum, shopify_7d_avg, meta_7d_avg


# ─────────────────────── Email ─────────────────────────────────────────────────

def send_or_draft_email(gmail_svc, sheet_url: str, date_str: str,
                         shopify_sales: dict, shopify_7d_avg: dict,
                         sessions: dict, meta_sum: dict, meta_7d_avg: dict,
                         recs: list):
    import base64
    from email.mime.multipart import MIMEMultipart
    from email.mime.text      import MIMEText

    gs  = safe_float(shopify_sales.get("gross_sales", 0))
    ns  = safe_float(shopify_sales.get("net_sales",   0))
    od  = int(safe_float(shopify_sales.get("orders",  0)))
    ao  = safe_float(shopify_sales.get("average_order_value", 0))
    rt  = abs(safe_float(shopify_sales.get("returns", 0)))

    ag_gs = shopify_7d_avg.get("gross_sales", 0)
    ag_ns = shopify_7d_avg.get("net_sales",   0)
    ag_od = shopify_7d_avg.get("orders",      0)

    sp    = meta_sum.get("spend",       0)
    pu    = meta_sum.get("purchases",   0)
    cpa   = meta_sum.get("cpa",         0)
    roas  = meta_sum.get("roas",        0)
    ag_sp = meta_7d_avg.get("spend",    0)

    urgent = [r for r in recs if r.startswith("🚨")]
    urgent_block = ""
    if urgent:
        urgent_block = (
            "\n<div style='background:#fff0f0;border-left:4px solid #d32f2f;"
            "padding:12px;margin-bottom:16px;border-radius:4px;'>"
            "<b>🚨 URGENT ISSUES</b><br>" +
            "<br>".join(f"• {u}" for u in urgent) +
            "</div>\n"
        )

    rec_items = "".join(
        f"<li style='margin:6px 0'>{r}</li>"
        for r in recs[:5]
    )

    html = f"""\
<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;max-width:680px;margin:auto;color:#222;">
<div style="background:#3b84f5;padding:20px;border-radius:8px 8px 0 0;">
  <h2 style="color:#fff;margin:0">Daily Store & Ads Performance</h2>
  <p style="color:#d0e4ff;margin:4px 0">{date_str} &nbsp;|&nbsp; Dhirai (dhirai.in)</p>
</div>
<div style="border:1px solid #e0e0e0;border-top:none;padding:20px;border-radius:0 0 8px 8px;">
{urgent_block}
<h3 style="border-bottom:2px solid #3b84f5;padding-bottom:6px">📦 Shopify Performance</h3>
<table style="width:100%;border-collapse:collapse;font-size:14px">
  <tr style="background:#f5f5f5"><th style="padding:8px;text-align:left">Metric</th>
    <th style="padding:8px;text-align:right">Yesterday</th>
    <th style="padding:8px;text-align:right">7-Day Avg</th>
    <th style="padding:8px;text-align:right">Change</th></tr>
  <tr><td style="padding:7px">Gross Sales</td>
    <td style="padding:7px;text-align:right"><b>₹{gs:,.0f}</b></td>
    <td style="padding:7px;text-align:right">₹{ag_gs:,.0f}</td>
    <td style="padding:7px;text-align:right;color:{'#2e7d32' if gs>=ag_gs else '#c62828'}">{_fmt_pct(pct_change(gs,ag_gs))}</td></tr>
  <tr style="background:#fafafa"><td style="padding:7px">Net Sales</td>
    <td style="padding:7px;text-align:right"><b>₹{ns:,.0f}</b></td>
    <td style="padding:7px;text-align:right">₹{ag_ns:,.0f}</td>
    <td style="padding:7px;text-align:right;color:{'#2e7d32' if ns>=ag_ns else '#c62828'}">{_fmt_pct(pct_change(ns,ag_ns))}</td></tr>
  <tr><td style="padding:7px">Orders</td>
    <td style="padding:7px;text-align:right"><b>{od}</b></td>
    <td style="padding:7px;text-align:right">{ag_od:.1f}</td>
    <td style="padding:7px;text-align:right;color:{'#2e7d32' if od>=ag_od else '#c62828'}">{_fmt_pct(pct_change(od,ag_od))}</td></tr>
  <tr style="background:#fafafa"><td style="padding:7px">Avg Order Value</td>
    <td style="padding:7px;text-align:right"><b>₹{ao:,.0f}</b></td>
    <td style="padding:7px;text-align:right">—</td><td></td></tr>
  <tr><td style="padding:7px">Returns</td>
    <td style="padding:7px;text-align:right">₹{rt:,.0f}</td>
    <td></td><td></td></tr>
</table>

<h3 style="border-bottom:2px solid #3b84f5;padding-bottom:6px;margin-top:24px">📣 Meta Ads Performance</h3>
<table style="width:100%;border-collapse:collapse;font-size:14px">
  <tr style="background:#f5f5f5"><th style="padding:8px;text-align:left">Metric</th>
    <th style="padding:8px;text-align:right">Yesterday</th>
    <th style="padding:8px;text-align:right">7-Day Avg</th>
    <th style="padding:8px;text-align:right">Change</th></tr>
  <tr><td style="padding:7px">Total Spend</td>
    <td style="padding:7px;text-align:right"><b>₹{sp:,.0f}</b></td>
    <td style="padding:7px;text-align:right">₹{ag_sp:,.0f}</td>
    <td style="padding:7px;text-align:right;color:{'#c62828' if sp>ag_sp else '#2e7d32'}">{_fmt_pct(pct_change(sp,ag_sp))}</td></tr>
  <tr style="background:#fafafa"><td style="padding:7px">Purchases</td>
    <td style="padding:7px;text-align:right"><b>{pu}</b></td>
    <td></td><td></td></tr>
  <tr><td style="padding:7px">Cost per Purchase</td>
    <td style="padding:7px;text-align:right"><b>₹{cpa:,.0f}</b></td>
    <td></td><td></td></tr>
  <tr style="background:#fafafa"><td style="padding:7px">Blended ROAS</td>
    <td style="padding:7px;text-align:right"><b style="color:{'#2e7d32' if roas>=ROAS_TARGET else '#c62828'}">{roas:.2f}x</b></td>
    <td></td><td></td></tr>
</table>

<h3 style="border-bottom:2px solid #3b84f5;padding-bottom:6px;margin-top:24px">✅ Recommended Actions for Today</h3>
<ol style="font-size:14px;line-height:1.7">{rec_items}</ol>

<div style="margin-top:24px;background:#f0f4ff;border-radius:6px;padding:14px;text-align:center">
  <a href="{sheet_url}" style="background:#3b84f5;color:#fff;padding:10px 24px;border-radius:5px;
     text-decoration:none;font-weight:bold;font-size:14px">📊 Open Full Dashboard</a>
</div>
<p style="font-size:11px;color:#888;margin-top:20px;text-align:center">
  Dhirai Daily Performance Bot · Sent automatically every morning</p>
</div></body></html>"""

    plain = (
        f"Daily Store & Ads Performance — {date_str}\n\n"
        f"SHOPIFY\n"
        f"  Gross Sales: ₹{gs:,.0f} ({_fmt_pct(pct_change(gs,ag_gs))} vs 7-day avg ₹{ag_gs:,.0f})\n"
        f"  Net Sales:   ₹{ns:,.0f} ({_fmt_pct(pct_change(ns,ag_ns))} vs 7-day avg ₹{ag_ns:,.0f})\n"
        f"  Orders:      {od} ({_fmt_pct(pct_change(od,ag_od))} vs 7-day avg {ag_od:.1f})\n"
        f"  AOV:         ₹{ao:,.0f}\n"
        f"  Returns:     ₹{rt:,.0f}\n\n"
        f"META ADS\n"
        f"  Spend:       ₹{sp:,.0f} ({_fmt_pct(pct_change(sp,ag_sp))} vs 7-day avg)\n"
        f"  Purchases:   {pu}\n"
        f"  CPA:         ₹{cpa:,.0f}\n"
        f"  ROAS:        {roas:.2f}x\n\n"
        f"RECOMMENDED ACTIONS\n" +
        "\n".join(f"  {i+1}. {r}" for i, r in enumerate(recs[:5])) +
        f"\n\nFull dashboard: {sheet_url}"
    )

    msg = MIMEMultipart("alternative")
    msg["To"]      = GMAIL_TO
    msg["Subject"] = f"Daily Store & Ads Performance Sheet - {date_str}"
    if GMAIL_CC:
        msg["Cc"] = GMAIL_CC
    msg.attach(MIMEText(plain, "plain"))
    msg.attach(MIMEText(html,  "html"))

    raw = __import__("base64").urlsafe_b64encode(msg.as_bytes()).decode()
    try:
        gmail_svc.users().messages().send(userId="me", body={"raw": raw}).execute()
        print(f"✅ Email sent to {GMAIL_TO}")
    except HttpError as e:
        print(f"⚠️  Send failed ({e}), creating Gmail draft...")
        gmail_svc.users().drafts().create(userId="me", body={"message": {"raw": raw}}).execute()
        print("✅ Gmail draft created.")


# ─────────────────────── Main ──────────────────────────────────────────────────

def main():
    yesterday  = yesterday_kolkata()
    date_str   = yesterday.isoformat()
    since_7d, until_7d = date_range_7d(yesterday)
    unavailable = []

    print(f"📅 Running dashboard for {date_str} (Asia/Kolkata) | 7-day window: {since_7d} → {until_7d}")

    # ── Shopify ──────────────────────────────────────────────────────────────
    print("🛒 Fetching Shopify data...")
    try:
        shopify_sales    = fetch_shopify_sales_day(date_str)
    except Exception as e:
        shopify_sales    = {}
        unavailable.append(f"Shopify sales data unavailable: {e}")

    try:
        sessions_today   = fetch_shopify_sessions_day(date_str)
    except Exception as e:
        sessions_today   = {}
        unavailable.append(f"Shopify sessions data unavailable: {e}")

    try:
        prod_today       = fetch_shopify_top_products(date_str)
    except Exception as e:
        prod_today       = {"rows": [], "columns": []}
        unavailable.append(f"Shopify products (today) unavailable: {e}")

    try:
        shopify_ts       = fetch_shopify_timeseries(since_7d, date_str)
    except Exception as e:
        shopify_ts       = {"rows": [], "columns": []}
        unavailable.append(f"Shopify timeseries unavailable: {e}")

    try:
        shopify_sess_ts  = fetch_shopify_sessions_timeseries(since_7d, date_str)
    except Exception as e:
        shopify_sess_ts  = {"rows": [], "columns": []}
        unavailable.append(f"Shopify sessions timeseries unavailable: {e}")

    try:
        prod_7d          = fetch_shopify_products_range(since_7d, until_7d)
    except Exception as e:
        prod_7d          = {"rows": [], "columns": []}
        unavailable.append(f"Shopify 7-day products unavailable: {e}")

    # ── Meta Ads ─────────────────────────────────────────────────────────────
    print("📣 Fetching Meta Ads data...")
    try:
        meta_campaigns   = fetch_meta_campaigns(date_str)
    except Exception as e:
        meta_campaigns   = []
        unavailable.append(f"Meta campaign data unavailable: {e}")

    try:
        meta_adsets      = fetch_meta_adsets(date_str)
    except Exception as e:
        meta_adsets      = []
        unavailable.append(f"Meta ad-set data unavailable: {e}")

    try:
        meta_7d          = fetch_meta_account_7d(since_7d, until_7d)
    except Exception as e:
        meta_7d          = []
        unavailable.append(f"Meta 7-day account data unavailable: {e}")

    # ── Recommendations ───────────────────────────────────────────────────────
    print("💡 Generating recommendations...")
    ts_rows  = shopify_ts.get("rows",    [])
    ts_cols  = shopify_ts.get("columns", [])
    prev_rows = [r for r in ts_rows if str(r[0]) != date_str]
    shopify_7d_avg = compute_7d_avg(prev_rows, ts_cols,
                                     ["gross_sales", "net_sales", "orders", "average_order_value"])

    recs, notes = generate_recommendations(
        shopify_sales, shopify_7d_avg, sessions_today,
        meta_campaigns, meta_adsets
    )

    # ── Google Sheet ──────────────────────────────────────────────────────────
    print("📊 Updating Google Sheet...")
    try:
        sheet_url, gmail_svc, meta_sum, shopify_7d_avg_final, meta_7d_avg = update_spreadsheet(
            yesterday,
            shopify_sales, shopify_ts, shopify_sess_ts,
            prod_today, prod_7d, sessions_today,
            meta_campaigns, meta_adsets, meta_7d,
            recs, notes, unavailable
        )
        print(f"  ✅ Sheet updated: {sheet_url}")
    except Exception as e:
        print(f"  ⚠️ Sheet update failed: {e}")
        import traceback; traceback.print_exc()
        sheet_url = "https://docs.google.com/spreadsheets (configure GOOGLE_SA_CREDENTIALS)"
        meta_sum  = {"spend": 0, "impressions": 0, "reach": 0, "clicks": 0,
                     "purchases": 0, "ctr": 0, "cpc": 0, "cpm": 0, "cpa": 0, "roas": 0}
        meta_7d_avg = {"spend": 0, "impressions": 0, "roas": 0, "purchases": 0}
        shopify_7d_avg_final = shopify_7d_avg
        try:
            _, _, gmail_svc = get_google_clients()
        except Exception:
            print("  ⚠️ Gmail auth also failed — cannot send email.")
            return

    # ── Email / draft ─────────────────────────────────────────────────────────
    print("✉️  Sending email / creating draft...")
    try:
        send_or_draft_email(
            gmail_svc, sheet_url, date_str,
            shopify_sales, shopify_7d_avg_final, sessions_today,
            meta_sum, meta_7d_avg, recs
        )
    except Exception as e:
        print(f"  ⚠️ Email step failed: {e}")

    if unavailable:
        print("\n⚠️  Some metrics were unavailable:")
        for m in unavailable:
            print(f"   • {m}")

    print("\n✅ Done.")


if __name__ == "__main__":
    main()
