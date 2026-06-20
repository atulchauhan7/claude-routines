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
]

SHOPIFY_DOMAIN    = os.environ.get("SHOPIFY_STORE_DOMAIN", "36dhns-ed.myshopify.com")
SHOPIFY_TOKEN     = os.environ.get("SHOPIFY_ACCESS_TOKEN", "")
META_TOKEN        = os.environ.get("META_ACCESS_TOKEN", "")
META_ACCOUNT_ID   = os.environ.get("META_AD_ACCOUNT_ID", "979830497515712")
SA_CREDENTIALS    = os.environ.get("GOOGLE_SA_CREDENTIALS", "service_account.json")
GMAIL_TO          = os.environ.get("GMAIL_TO", "atul012001@gmail.com")
GMAIL_CC          = os.environ.get("GMAIL_CC", "")
SPREADSHEET_NAME  = os.environ.get("SPREADSHEET_NAME", "Daily Store & Ads Performance Sheet")

# ─────────────────────── Colours (hex without #) ──────────────────────────────

C_HEADER_BG   = {"red": 0.133, "green": 0.133, "blue": 0.133}   # Dark charcoal
C_HEADER_FG   = {"red": 1,     "green": 1,     "blue": 1}        # White
C_SECTION_BG  = {"red": 0.957, "green": 0.957, "blue": 0.957}    # Light grey
C_GREEN       = {"red": 0.204, "green": 0.659, "blue": 0.325}    # +ve change
C_RED         = {"red": 0.839, "green": 0.153, "blue": 0.157}    # -ve change
C_YELLOW      = {"red": 1,     "green": 0.898, "blue": 0.2}      # Warning
C_ORANGE      = {"red": 1,     "green": 0.596, "blue": 0}        # Alert
C_ACCENT      = {"red": 0.259, "green": 0.522, "blue": 0.957}    # Brand blue
C_WHITE       = {"red": 1,     "green": 1,     "blue": 1}

# ─────────────────────── Google Auth ──────────────────────────────────────────

def get_google_clients():
    creds = Credentials.from_service_account_file(SA_CREDENTIALS, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    gmail_service  = build("gmail",  "v1", credentials=creds)
    return gc, sheets_service, gmail_service

# ─────────────────────── Date helpers ─────────────────────────────────────────

def yesterday_kolkata():
    now = datetime.datetime.now(KOLKATA_TZ)
    yesterday = now.date() - datetime.timedelta(days=1)
    return yesterday

def seven_days_ago_kolkata():
    now = datetime.datetime.now(KOLKATA_TZ)
    return now.date() - datetime.timedelta(days=8)   # 7 days before yesterday

# ─────────────────────── Shopify data fetching ────────────────────────────────

SHOPIFY_GRAPHQL_URL = f"https://{SHOPIFY_DOMAIN}/admin/api/2024-01/graphql.json"

def shopify_analytics(shopify_query: str) -> dict:
    """Run a ShopifyQL analytics query and return rows + columns."""
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
    """Fetch all key Shopify metrics for a single day."""
    results = {}

    # ── Sales summary ──
    q = (
        f"FROM sales SHOW gross_sales, net_sales, orders, average_order_value, "
        f"total_sales, discounts, returns, shipping_charges, taxes "
        f"SINCE {date_str} UNTIL {date_str}"
    )
    r = shopify_analytics(q)
    if r["rows"]:
        row = r["rows"][0]
        cols = r["columns"]
        results["sales"] = dict(zip(cols, row))
    else:
        results["sales"] = {c: None for c in [
            "gross_sales","net_sales","orders","average_order_value",
            "total_sales","discounts","returns","shipping_charges","taxes"
        ]}

    # ── Sessions / conversion ──
    q = (
        f"FROM sessions SHOW sessions, sessions_with_cart_additions, "
        f"sessions_that_reached_checkout, sessions_that_completed_checkout, "
        f"conversion_rate SINCE {date_str} UNTIL {date_str}"
    )
    r = shopify_analytics(q)
    if r["rows"]:
        row = r["rows"][0]
        results["sessions"] = dict(zip(r["columns"], row))
    else:
        results["sessions"] = {}

    # ── Top 10 products by gross sales ──
    q = (
        f"FROM sales SHOW gross_sales, net_sales, orders, total_sales "
        f"GROUP BY product_title ORDER BY gross_sales DESC LIMIT 10 "
        f"SINCE {date_str} UNTIL {date_str}"
    )
    r = shopify_analytics(q)
    results["top_products"] = r

    return results


def fetch_shopify_7day(since_str: str, until_str: str) -> dict:
    """Fetch daily sales rows for 7-day window."""
    q = (
        f"FROM sales SHOW gross_sales, net_sales, orders, average_order_value, "
        f"total_sales, discounts, returns, shipping_charges, taxes "
        f"TIMESERIES day SINCE {since_str} UNTIL {until_str}"
    )
    return shopify_analytics(q)


def fetch_shopify_7day_products(since_str: str, until_str: str) -> dict:
    """Top 10 products over the 7-day window."""
    q = (
        f"FROM sales SHOW gross_sales, net_sales, orders "
        f"GROUP BY product_title ORDER BY gross_sales DESC LIMIT 10 "
        f"SINCE {since_str} UNTIL {until_str}"
    )
    return shopify_analytics(q)


# ─────────────────────── Meta Ads data fetching ────────────────────────────────

META_BASE = "https://graph.facebook.com/v19.0"

META_FIELDS = (
    "campaign_id,campaign_name,adset_id,adset_name,"
    "spend,impressions,reach,clicks,ctr,cpc,cpm,"
    "purchase_roas,actions,cost_per_action_type,"
    "date_start,date_stop"
)

def meta_request(endpoint: str, params: dict) -> dict:
    params["access_token"] = META_TOKEN
    r = requests.get(f"{META_BASE}/{endpoint}", params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def fetch_meta_campaigns(date_str: str) -> list:
    """Campaign-level insights for a single day."""
    r = meta_request(
        f"act_{META_ACCOUNT_ID}/insights",
        {
            "level": "campaign",
            "fields": META_FIELDS,
            "time_range": json.dumps({"since": date_str, "until": date_str}),
            "limit": 50,
        },
    )
    return r.get("data", [])


def fetch_meta_adsets(date_str: str) -> list:
    """Ad-set-level insights for a single day."""
    r = meta_request(
        f"act_{META_ACCOUNT_ID}/insights",
        {
            "level": "adset",
            "fields": META_FIELDS,
            "time_range": json.dumps({"since": date_str, "until": date_str}),
            "limit": 50,
        },
    )
    return r.get("data", [])


def fetch_meta_7day(since_str: str, until_str: str) -> list:
    """Account-level daily insights for the 7-day window."""
    r = meta_request(
        f"act_{META_ACCOUNT_ID}/insights",
        {
            "level": "account",
            "fields": "spend,impressions,reach,clicks,ctr,cpc,cpm,purchase_roas,actions,cost_per_action_type",
            "time_range": json.dumps({"since": since_str, "until": until_str}),
            "time_increment": 1,
            "limit": 50,
        },
    )
    return r.get("data", [])


def extract_purchases(actions: list) -> int:
    """Sum purchase action values from Meta actions list."""
    count = 0
    for a in (actions or []):
        if a.get("action_type") in ("offsite_conversion.fb_pixel_purchase", "purchase"):
            try:
                count += int(float(a.get("value", 0)))
            except (ValueError, TypeError):
                pass
    return count


def extract_roas(purchase_roas: list | None) -> float | None:
    """Extract numeric ROAS from Meta's purchase_roas field."""
    if not purchase_roas:
        return None
    for r in purchase_roas:
        if r.get("action_type") in ("offsite_conversion.fb_pixel_purchase", "omni_purchase"):
            try:
                return round(float(r["value"]), 2)
            except (ValueError, KeyError):
                pass
    return None


def extract_cpa(cost_per_action_type: list | None) -> float | None:
    """Extract CPA (cost per purchase) from Meta's cost_per_action_type."""
    for a in (cost_per_action_type or []):
        if a.get("action_type") in ("offsite_conversion.fb_pixel_purchase", "purchase"):
            try:
                return round(float(a["value"]), 2)
            except (ValueError, KeyError):
                pass
    return None


# ─────────────────────── Computation helpers ──────────────────────────────────

def safe_float(v) -> float:
    try:
        return float(str(v).replace(",", "").replace("₹", "").replace(" INR", "").strip())
    except (ValueError, TypeError):
        return 0.0


def pct_change(current: float, prior: float) -> float | None:
    if prior == 0:
        return None
    return round((current - prior) / prior * 100, 1)


def avg(values: list) -> float:
    vals = [v for v in values if v is not None]
    return sum(vals) / len(vals) if vals else 0.0


def compute_7day_averages(daily_rows: list, col_map: dict) -> dict:
    """Compute column averages from daily rows (list of lists)."""
    avgs = {}
    for col, idx in col_map.items():
        vals = []
        for row in daily_rows:
            try:
                vals.append(float(row[idx]))
            except (ValueError, TypeError, IndexError):
                pass
        avgs[col] = round(avg(vals), 2) if vals else 0.0
    return avgs


# ─────────────────────── Recommendations engine ────────────────────────────────

def generate_recommendations(shopify_today: dict, shopify_7d_avg: dict,
                              meta_campaigns: list, meta_adsets: list,
                              anomalies: list | None = None) -> list:
    recs = []
    notes = []

    # ── Shopify insights ──
    gross = safe_float(shopify_today.get("sales", {}).get("gross_sales", 0))
    avg_gross = shopify_7d_avg.get("gross_sales", 0)
    if avg_gross > 0:
        chg = pct_change(gross, avg_gross)
        if chg is not None and chg < -15:
            recs.append(f"⚠️ Revenue dropped {abs(chg):.1f}% vs 7-day avg — check ad delivery, site issues, or seasonal demand.")
        elif chg is not None and chg > 20:
            recs.append(f"✅ Revenue up {chg:.1f}% vs 7-day avg — identify top-performing campaigns and scale budgets.")

    orders = safe_float(shopify_today.get("sales", {}).get("orders", 0))
    avg_orders = shopify_7d_avg.get("orders", 0)
    if avg_orders > 0 and orders < avg_orders * 0.8:
        recs.append("⚠️ Orders below 80% of 7-day avg — review traffic sources and Meta campaign delivery.")

    returns_val = abs(safe_float(shopify_today.get("sales", {}).get("returns", 0)))
    gross_val = safe_float(shopify_today.get("sales", {}).get("gross_sales", 0))
    if gross_val > 0 and returns_val / gross_val > 0.12:
        recs.append(f"⚠️ High return rate ({returns_val/gross_val*100:.1f}% of gross) — check sizing, quality complaints, or product descriptions.")

    # ── Meta campaign insights ──
    for c in meta_campaigns:
        spend = safe_float(c.get("spend", 0))
        roas = extract_roas(c.get("purchase_roas"))
        name = c.get("campaign_name", c.get("name", ""))
        if roas is not None and spend > 500 and roas < 1.0:
            recs.append(f"🚨 PAUSE '{name}' — ROAS {roas}x is below breakeven. Reallocate ₹{spend:,.0f}/day budget.")
        elif roas is not None and spend > 500 and roas < 1.5:
            recs.append(f"⚠️ Review '{name}' — ROAS {roas}x is low. Test new creatives or tighten audience.")
        elif roas is not None and roas > 5.0 and spend > 1000:
            recs.append(f"✅ Scale '{name}' — ROAS {roas}x is excellent. Increase budget by 20–30%.")

    # ── Ad set insights ──
    for a in meta_adsets:
        spend = safe_float(a.get("spend", 0))
        roas = extract_roas(a.get("purchase_roas"))
        name = a.get("adset_name", a.get("name", ""))
        ctr  = safe_float(a.get("ctr", 0))
        if spend > 1000 and ctr < 0.8:
            recs.append(f"⚠️ Low CTR ({ctr:.2f}%) on ad set '{name}' — refresh creative or test new hooks.")

    # ── Anomaly signals ──
    for anomaly in (anomalies or []):
        recs.append(f"📊 Meta signal: {anomaly}")

    # ── Fallback generic recs if nothing specific ──
    if not recs:
        recs.append("✅ All campaigns performing within normal range — continue monitoring.")

    recs.append("🔍 Review Shopify abandoned checkouts and set up recovery email flows if not already active.")
    recs.append("📱 Check product pages for top sellers — ensure images, descriptions, and sizes are current.")

    return recs[:7], notes   # cap at 7 actionable recs


# ─────────────────────── Google Sheets helpers ────────────────────────────────

def cell_format(sheets_svc, spreadsheet_id: str, sheet_id: int,
                start_row: int, end_row: int, start_col: int, end_col: int,
                bold: bool = False, bg_color: dict = None, fg_color: dict = None,
                horizontal_alignment: str = None, number_format: str = None,
                font_size: int = None, text_format: dict = None):
    """Apply cell formatting via Sheets batchUpdate."""
    cell_fmt = {}
    if bold or fg_color or font_size or text_format:
        tf = text_format or {}
        if bold:        tf["bold"] = True
        if fg_color:    tf["foregroundColor"] = fg_color
        if font_size:   tf["fontSize"] = font_size
        cell_fmt["textFormat"] = tf
    if bg_color:
        cell_fmt["backgroundColor"] = bg_color
    if horizontal_alignment:
        cell_fmt["horizontalAlignment"] = horizontal_alignment
    if number_format:
        cell_fmt["numberFormat"] = {"type": "NUMBER", "pattern": number_format}

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
            "fields": "userEnteredFormat(" + ",".join([
                "textFormat" if (bold or fg_color or font_size or text_format) else "",
                "backgroundColor" if bg_color else "",
                "horizontalAlignment" if horizontal_alignment else "",
                "numberFormat" if number_format else "",
            ]).replace(",,", ",").strip(",") + ")",
        }
    }


def freeze_request(sheet_id: int, rows: int = 1, cols: int = 0):
    return {
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols},
            },
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
        }
    }


def auto_resize_request(sheet_id: int, start_col: int = 0, end_col: int = 26):
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
                        start_col: int, end_col: int):
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


def conditional_format_pct(sheet_id: int, start_row: int, end_row: int,
                             start_col: int, end_col: int):
    """Colour positive green, negative red for percentage change columns."""
    return [
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet_id, "startRowIndex": start_row,
                                "endRowIndex": end_row, "startColumnIndex": start_col,
                                "endColumnIndex": end_col}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0"}]},
                        "format": {"backgroundColor": {"red": 0.714, "green": 0.929, "blue": 0.714}},
                    },
                },
                "index": 0,
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet_id, "startRowIndex": start_row,
                                "endRowIndex": end_row, "startColumnIndex": start_col,
                                "endColumnIndex": end_col}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                        "format": {"backgroundColor": {"red": 0.957, "green": 0.714, "blue": 0.714}},
                    },
                },
                "index": 1,
            }
        },
    ]


def conditional_format_roas(sheet_id: int, start_row: int, end_row: int,
                              start_col: int, end_col: int):
    """ROAS < 1.5 = red, 1.5-3 = yellow, > 3 = green."""
    return [
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet_id, "startRowIndex": start_row,
                                "endRowIndex": end_row, "startColumnIndex": start_col,
                                "endColumnIndex": end_col}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "1.5"}]},
                        "format": {"backgroundColor": {"red": 0.957, "green": 0.714, "blue": 0.714}},
                    },
                },
                "index": 0,
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet_id, "startRowIndex": start_row,
                                "endRowIndex": end_row, "startColumnIndex": start_col,
                                "endColumnIndex": end_col}],
                    "booleanRule": {
                        "condition": {
                            "type": "NUMBER_BETWEEN",
                            "values": [{"userEnteredValue": "1.5"}, {"userEnteredValue": "3"}],
                        },
                        "format": {"backgroundColor": {"red": 1, "green": 0.949, "blue": 0.8}},
                    },
                },
                "index": 1,
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet_id, "startRowIndex": start_row,
                                "endRowIndex": end_row, "startColumnIndex": start_col,
                                "endColumnIndex": end_col}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "3"}]},
                        "format": {"backgroundColor": {"red": 0.714, "green": 0.929, "blue": 0.714}},
                    },
                },
                "index": 2,
            }
        },
    ]


def write_header_row(ws, row: list, row_idx: int = 0):
    """Write a bold header row with dark background."""
    ws.update(f"A{row_idx+1}", [row])


# ─────────────────────── Sheet tab builders ────────────────────────────────────

def build_shopify_daily_tab(ws, sheets_svc, spreadsheet_id: str, sheet_id: int,
                             today_data: dict, seven_day_rows: list, seven_day_cols: list):
    """Append today's Shopify row to the tab; create header if sheet is empty."""
    headers = [
        "Date", "Gross Sales (₹)", "Net Sales (₹)", "Orders",
        "AOV (₹)", "Total Sales (₹)", "Discounts (₹)", "Returns (₹)",
        "Shipping (₹)", "Taxes (₹)", "Sessions", "Cart Adds",
        "Reached Checkout", "Completed Checkout", "Conversion Rate (%)",
    ]

    existing = ws.get_all_values()
    if not existing or existing[0] != headers:
        ws.clear()
        ws.update("A1", [headers])
        existing = [headers]

    existing_dates = {r[0] for r in existing[1:] if r}

    # Build today's row from today_data
    s   = today_data.get("sales", {})
    ses = today_data.get("sessions", {})
    date_str = yesterday_kolkata().isoformat()

    today_row = [
        date_str,
        safe_float(s.get("gross_sales", 0)),
        safe_float(s.get("net_sales", 0)),
        int(safe_float(s.get("orders", 0))),
        round(safe_float(s.get("average_order_value", 0)), 2),
        safe_float(s.get("total_sales", 0)),
        abs(safe_float(s.get("discounts", 0))),
        abs(safe_float(s.get("returns", 0))),
        safe_float(s.get("shipping_charges", 0)),
        safe_float(s.get("taxes", 0)),
        int(safe_float(ses.get("sessions", 0))),
        int(safe_float(ses.get("sessions_with_cart_additions", 0))),
        int(safe_float(ses.get("sessions_that_reached_checkout", 0))),
        int(safe_float(ses.get("sessions_that_completed_checkout", 0))),
        round(safe_float(ses.get("conversion_rate", 0)), 4),
    ]

    if date_str not in existing_dates:
        next_row = len(existing) + 1
        ws.update(f"A{next_row}", [today_row])

    total_rows = max(len(existing) + 1, 3)

    # Format
    requests_body = [
        cell_format(sheets_svc, spreadsheet_id, sheet_id,
                    0, 1, 0, len(headers),
                    bold=True, bg_color=C_HEADER_BG, fg_color=C_HEADER_FG),
        freeze_request(sheet_id, rows=1),
        add_filter_request(sheet_id, 0, total_rows, 0, len(headers)),
        auto_resize_request(sheet_id, 0, len(headers)),
    ]
    # Currency format for sales columns (B-G, H = col 1-7)
    for col in range(1, 10):
        requests_body.append(
            cell_format(sheets_svc, spreadsheet_id, sheet_id,
                        1, total_rows, col, col + 1,
                        number_format='₹#,##0.00')
        )
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests_body}
    ).execute()


def build_meta_daily_tab(ws, sheets_svc, spreadsheet_id: str, sheet_id: int,
                          today_campaigns: list, seven_day_data: list):
    """Populate the 'Meta Ads Daily Data' tab."""
    ws.clear()
    headers = [
        "Date", "Total Spend (₹)", "Impressions", "Reach", "Clicks",
        "CTR (%)", "CPC (₹)", "CPM (₹)", "Purchases", "CPA (₹)", "ROAS",
    ]
    ws.update("A1", [headers])

    # Aggregate today's data across campaigns
    total_spend  = sum(safe_float(c.get("spend", 0)) for c in today_campaigns)
    total_impr   = sum(int(safe_float(c.get("impressions", 0))) for c in today_campaigns)
    total_reach  = sum(int(safe_float(c.get("reach", 0))) for c in today_campaigns)
    total_clicks = sum(int(safe_float(c.get("clicks", 0))) for c in today_campaigns)
    total_purch  = sum(extract_purchases(c.get("actions", [])) for c in today_campaigns)
    ctr  = round(total_clicks / total_impr * 100, 2) if total_impr else 0
    cpc  = round(total_spend / total_clicks, 2) if total_clicks else 0
    cpm  = round(total_spend / total_impr * 1000, 2) if total_impr else 0
    cpa  = round(total_spend / total_purch, 2) if total_purch else 0
    roas_vals = [extract_roas(c.get("purchase_roas")) for c in today_campaigns
                 if extract_roas(c.get("purchase_roas")) is not None]
    blended_roas = round(avg(roas_vals), 2) if roas_vals else 0

    today_str = yesterday_kolkata().isoformat()
    today_row = [today_str, round(total_spend, 2), total_impr, total_reach,
                 total_clicks, ctr, cpc, cpm, total_purch, cpa, blended_roas]

    existing_dates_range = ws.col_values(1)[1:]   # skip header
    if today_str not in existing_dates_range:
        next_row = len(existing_dates_range) + 2
        ws.update(f"A{next_row}", [today_row])

    total_rows = max(len(existing_dates_range) + 2, 3)
    requests_body = [
        cell_format(sheets_svc, spreadsheet_id, sheet_id,
                    0, 1, 0, len(headers),
                    bold=True, bg_color=C_HEADER_BG, fg_color=C_HEADER_FG),
        freeze_request(sheet_id, rows=1),
        add_filter_request(sheet_id, 0, total_rows, 0, len(headers)),
        auto_resize_request(sheet_id, 0, len(headers)),
    ]
    # Currency format: Spend(B), CPC(G), CPM(H), CPA(J)
    for col in [1, 6, 7, 9]:
        requests_body.append(
            cell_format(sheets_svc, spreadsheet_id, sheet_id,
                        1, total_rows, col, col + 1,
                        number_format='₹#,##0.00')
        )
    # Percentage: CTR(F)
    requests_body.append(
        cell_format(sheets_svc, spreadsheet_id, sheet_id,
                    1, total_rows, 5, 6,
                    number_format='0.00"%"')
    )
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests_body}
    ).execute()


def build_product_tab(ws, sheets_svc, spreadsheet_id: str, sheet_id: int,
                       today_products: dict, seven_day_products: dict):
    """Populate the 'Product Performance' tab."""
    ws.clear()
    headers_today = ["Product", "Gross Sales (₹)", "Net Sales (₹)", "Orders", "Total Sales (₹)"]
    section_7d    = ["", "7-Day Gross (₹)", "7-Day Net (₹)", "7-Day Orders"]

    ws.update("A1", [["─── Yesterday ───", "", "", "", "",
                       "─── 7-Day Total ───", "", "", ""]])
    ws.update("A2", [headers_today + section_7d[1:]])

    today_rows = today_products.get("rows", [])
    seven_rows = seven_day_products.get("rows", [])

    # Build lookup for 7-day totals
    seven_lookup = {}
    for row in seven_rows:
        prod = str(row[0])
        seven_lookup[prod] = row

    data = []
    for row in today_rows:
        prod = str(row[0])
        seven = seven_lookup.get(prod, [""] * 4)
        data.append([
            prod,
            safe_float(row[1]),
            safe_float(row[2]),
            int(safe_float(row[3])),
            safe_float(row[4]) if len(row) > 4 else "",
            safe_float(seven[1]) if len(seven) > 1 else "",
            safe_float(seven[2]) if len(seven) > 2 else "",
            int(safe_float(seven[3])) if len(seven) > 3 else "",
        ])

    if data:
        ws.update("A3", data)

    total_rows = len(data) + 3
    requests_body = [
        cell_format(sheets_svc, spreadsheet_id, sheet_id, 0, 1, 0, 9,
                    bold=True, bg_color=C_ACCENT, fg_color=C_HEADER_FG),
        cell_format(sheets_svc, spreadsheet_id, sheet_id, 1, 2, 0, 9,
                    bold=True, bg_color=C_HEADER_BG, fg_color=C_HEADER_FG),
        freeze_request(sheet_id, rows=2),
        add_filter_request(sheet_id, 1, total_rows, 0, 9),
        auto_resize_request(sheet_id, 0, 9),
    ]
    for col in [1, 2, 4, 5, 6]:
        requests_body.append(
            cell_format(sheets_svc, spreadsheet_id, sheet_id,
                        2, total_rows, col, col + 1,
                        number_format='₹#,##0.00')
        )
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests_body}
    ).execute()


def build_campaign_tab(ws, sheets_svc, spreadsheet_id: str, sheet_id: int,
                        campaigns: list, adsets: list):
    """Populate the 'Campaign Performance' tab."""
    ws.clear()

    # ── Campaign section ──
    camp_headers = [
        "Campaign", "Status", "Spend (₹)", "Impressions", "Reach",
        "Clicks", "CTR (%)", "CPC (₹)", "CPM (₹)", "Purchases", "CPA (₹)", "ROAS",
    ]
    ws.update("A1", [["━━━ CAMPAIGN LEVEL ━━━"] + [""] * (len(camp_headers) - 1)])
    ws.update("A2", [camp_headers])

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
            safe_float(c.get("ctr", 0)),
            safe_float(c.get("cpc", 0)),
            safe_float(c.get("cpm", 0)),
            extract_purchases(c.get("actions", [])),
            extract_cpa(c.get("cost_per_action_type")) or "",
            extract_roas(c.get("purchase_roas")) or "",
        ])

    camp_data.sort(key=lambda r: r[2], reverse=True)
    if camp_data:
        ws.update("A3", camp_data)

    camp_end = len(camp_data) + 3

    # ── Ad set section ──
    adset_headers = [
        "Ad Set", "Status", "Spend (₹)", "Impressions", "Reach",
        "Clicks", "CTR (%)", "CPC (₹)", "CPM (₹)", "Purchases", "CPA (₹)", "ROAS", "Learning Phase",
    ]
    as_start = camp_end + 1
    ws.update(f"A{as_start}", [["━━━ AD SET LEVEL ━━━"] + [""] * (len(adset_headers) - 1)])
    ws.update(f"A{as_start + 1}", [adset_headers])

    as_data = []
    for a in adsets:
        spend = safe_float(a.get("spend", 0))
        if spend == 0:
            continue
        delivery = a.get("delivery", {})
        substatuses = delivery.get("substatuses", []) if isinstance(delivery, dict) else []
        learning = "In Learning" if "in_learning_phase" in substatuses else (
                   "Learning Exit Fail" if "learning_exit_unsuccessfully" in substatuses else "")
        as_data.append([
            a.get("adset_name", a.get("name", "")),
            a.get("effective_status", a.get("status", "")),
            round(spend, 2),
            int(safe_float(a.get("impressions", 0))),
            int(safe_float(a.get("reach", 0))),
            int(safe_float(a.get("clicks", 0))),
            safe_float(a.get("ctr", 0)),
            safe_float(a.get("cpc", 0)),
            safe_float(a.get("cpm", 0)),
            extract_purchases(a.get("actions", [])),
            extract_cpa(a.get("cost_per_action_type")) or "",
            extract_roas(a.get("purchase_roas")) or "",
            learning,
        ])

    as_data.sort(key=lambda r: r[2], reverse=True)
    if as_data:
        ws.update(f"A{as_start + 2}", as_data)

    as_end = as_start + 2 + len(as_data)

    # ── Format requests ──
    requests_body = [
        cell_format(sheets_svc, spreadsheet_id, sheet_id, 0, 1, 0, len(camp_headers),
                    bold=True, bg_color=C_ACCENT, fg_color=C_HEADER_FG, font_size=11),
        cell_format(sheets_svc, spreadsheet_id, sheet_id, 1, 2, 0, len(camp_headers),
                    bold=True, bg_color=C_HEADER_BG, fg_color=C_HEADER_FG),
        cell_format(sheets_svc, spreadsheet_id, sheet_id,
                    as_start - 1, as_start, 0, len(adset_headers),
                    bold=True, bg_color=C_ACCENT, fg_color=C_HEADER_FG, font_size=11),
        cell_format(sheets_svc, spreadsheet_id, sheet_id,
                    as_start, as_start + 1, 0, len(adset_headers),
                    bold=True, bg_color=C_HEADER_BG, fg_color=C_HEADER_FG),
        freeze_request(sheet_id, rows=2),
        auto_resize_request(sheet_id, 0, len(adset_headers)),
    ]
    # Currency / percentage formats for campaign rows
    for col in [2, 7, 8, 10]:
        requests_body.append(
            cell_format(sheets_svc, spreadsheet_id, sheet_id,
                        2, camp_end, col, col + 1,
                        number_format='₹#,##0.00')
        )
    requests_body.append(
        cell_format(sheets_svc, spreadsheet_id, sheet_id,
                    2, camp_end, 6, 7,
                    number_format='0.00"%"')
    )
    # ROAS conditional formatting for campaigns
    requests_body += conditional_format_roas(sheet_id, 2, camp_end, 11, 12)
    # Currency / pct for adset rows
    for col in [2, 7, 8, 10]:
        requests_body.append(
            cell_format(sheets_svc, spreadsheet_id, sheet_id,
                        as_start + 1, as_end, col, col + 1,
                        number_format='₹#,##0.00')
        )
    requests_body += conditional_format_roas(sheet_id, as_start + 1, as_end, 11, 12)

    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests_body}
    ).execute()


def build_recommendations_tab(ws, sheets_svc, spreadsheet_id: str, sheet_id: int,
                                recs: list, notes: list,
                                date_str: str, unavailable_metrics: list = None):
    """Populate the 'Recommendations & Notes' tab."""
    ws.clear()

    all_rows = [
        [f"RECOMMENDATIONS & NOTES — {date_str}", ""],
        ["", ""],
        ["TYPE", "RECOMMENDATION / NOTE"],
    ]

    for rec in recs:
        icon = rec[:2] if rec[0] in "⚠🚨✅📊🔍📱" else "💡"
        all_rows.append([icon, rec[len(icon):].strip()])

    if unavailable_metrics:
        all_rows.append(["", ""])
        all_rows.append(["UNAVAILABLE METRICS", ""])
        for m in unavailable_metrics:
            all_rows.append(["ℹ️", m])

    ws.update("A1", all_rows)

    n = len(all_rows)
    requests_body = [
        cell_format(sheets_svc, spreadsheet_id, sheet_id, 0, 1, 0, 2,
                    bold=True, bg_color=C_ACCENT, fg_color=C_HEADER_FG, font_size=13),
        cell_format(sheets_svc, spreadsheet_id, sheet_id, 2, 3, 0, 2,
                    bold=True, bg_color=C_HEADER_BG, fg_color=C_HEADER_FG),
        freeze_request(sheet_id, rows=3),
        auto_resize_request(sheet_id, 0, 2),
    ]
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests_body}
    ).execute()


def build_dashboard_tab(ws, sheets_svc, spreadsheet_id: str, sheet_id: int,
                         shopify_today: dict, shopify_7d_avg: dict,
                         meta_summary: dict, recs: list, date_str: str,
                         shopify_tab_id: int, meta_tab_id: int):
    """Populate the 'Dashboard' tab with executive summary and key metrics."""
    ws.clear()

    def fmt_inr(v: float) -> str:
        return f"₹{v:,.0f}"

    def fmt_pct(v: float | None) -> str:
        if v is None:
            return "—"
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
        [f"DAILY STORE & ADS PERFORMANCE — {date_str}", "", "", ""],
        ["", "", "", ""],
        ["━━━ SHOPIFY PERFORMANCE ━━━", "", "", ""],
        ["Metric", "Yesterday", "7-Day Avg", "Change vs Avg"],
        ["Gross Sales", fmt_inr(gs), fmt_inr(ag_gs), fmt_pct(pct_change(gs, ag_gs))],
        ["Net Sales", fmt_inr(ns), fmt_inr(ag_ns), fmt_pct(pct_change(ns, ag_ns))],
        ["Orders", str(od), str(round(ag_od, 1)), fmt_pct(pct_change(od, ag_od))],
        ["Avg Order Value", fmt_inr(ao), fmt_inr(ag_ao), fmt_pct(pct_change(ao, ag_ao))],
        ["Returns", fmt_inr(rt), "", ""],
        ["", "", "", ""],
        ["━━━ META ADS PERFORMANCE ━━━", "", "", ""],
        ["Metric", "Yesterday", "", ""],
        ["Total Spend", fmt_inr(sp), "", ""],
        ["Impressions", f"{imp:,}", "", ""],
        ["Clicks", f"{cl:,}", "", ""],
        ["CTR", f"{ctr:.2f}%", "", ""],
        ["Purchases (attributed)", str(pu), "", ""],
        ["CPA", fmt_inr(cpa), "", ""],
        ["Blended ROAS", f"{ro:.2f}x", "", ""],
        ["", "", "", ""],
        ["━━━ KEY WINS ━━━", "", "", ""],
    ]

    # Add key wins (positive recs)
    wins = [r for r in recs if r.startswith("✅")]
    if wins:
        for w in wins[:3]:
            rows.append(["", w, "", ""])
    else:
        rows.append(["", "Monitor for positive trends.", "", ""])

    rows.append(["", "", "", ""])
    rows.append(["━━━ KEY ISSUES ━━━", "", "", ""])
    issues = [r for r in recs if r.startswith(("⚠️", "🚨"))]
    if issues:
        for i in issues[:3]:
            rows.append(["", i, "", ""])
    else:
        rows.append(["", "No critical issues detected.", "", ""])

    rows.append(["", "", "", ""])
    rows.append(["━━━ RECOMMENDED ACTIONS TODAY ━━━", "", "", ""])
    for idx, rec in enumerate(recs[:5], 1):
        rows.append([f"{idx}.", rec, "", ""])

    ws.update("A1", rows)

    requests_body = [
        # Title row
        cell_format(sheets_svc, spreadsheet_id, sheet_id, 0, 1, 0, 4,
                    bold=True, bg_color=C_ACCENT, fg_color=C_HEADER_FG, font_size=14),
        # Section headers
    ]

    section_rows = [2, 10, 20, 22, 25, 27]  # approximate; adjust if row count changes
    for sr in section_rows:
        requests_body.append(
            cell_format(sheets_svc, spreadsheet_id, sheet_id, sr, sr + 1, 0, 4,
                        bold=True, bg_color=C_HEADER_BG, fg_color=C_HEADER_FG, font_size=11)
        )

    requests_body += [
        # Sub-header rows for tables
        cell_format(sheets_svc, spreadsheet_id, sheet_id, 3, 4, 0, 4,
                    bold=True, bg_color=C_SECTION_BG),
        cell_format(sheets_svc, spreadsheet_id, sheet_id, 11, 12, 0, 4,
                    bold=True, bg_color=C_SECTION_BG),
        freeze_request(sheet_id, rows=1),
        auto_resize_request(sheet_id, 0, 4),
    ]

    # Conditional formatting for change column (D, col index 3)
    requests_body += conditional_format_pct(sheet_id, 4, 9, 3, 4)

    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests_body}
    ).execute()


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
    """Make sure all required tabs exist; return {name: worksheet}."""
    existing = {ws.title: ws for ws in ss.worksheets()}
    tab_map = {}
    for tab in TAB_NAMES:
        if tab not in existing:
            ws = ss.add_worksheet(title=tab, rows=500, cols=30)
        else:
            ws = existing[tab]
        tab_map[tab] = ws
    # Remove the default 'Sheet1' if present
    if "Sheet1" in existing and "Sheet1" not in TAB_NAMES:
        try:
            ss.del_worksheet(existing["Sheet1"])
        except Exception:
            pass
    # Reorder tabs
    for idx, name in enumerate(TAB_NAMES):
        try:
            ss.reorder_worksheets([tab_map[t] for t in TAB_NAMES if t in tab_map])
        except Exception:
            pass
        break
    return tab_map


def update_spreadsheet(date_str: str,
                        shopify_today: dict, shopify_7d: dict,
                        shopify_products_today: dict, shopify_products_7d: dict,
                        meta_campaigns: list, meta_adsets: list,
                        recs: list, notes: list,
                        unavailable: list) -> str:
    gc, sheets_svc, gmail_svc = get_google_clients()
    ss = get_or_create_spreadsheet(gc, SPREADSHEET_NAME)
    tab_map = ensure_tabs(ss)

    sheet_ids = {ws.title: ws.id for ws in ss.worksheets()}

    # Compute 7-day averages from Shopify data
    seven_day_rows = shopify_7d.get("rows", [])
    seven_day_cols = shopify_7d.get("columns", [])
    col_idx = {c: i for i, c in enumerate(seven_day_cols)}

    # Exclude today's row if it's in the 7-day window
    prev_rows = [r for r in seven_day_rows if str(r[col_idx.get("day", 0)]) != date_str]
    shopify_7d_avg = compute_7day_averages(prev_rows, {
        "gross_sales":        col_idx.get("gross_sales", 1),
        "net_sales":          col_idx.get("net_sales", 2),
        "orders":             col_idx.get("orders", 3),
        "average_order_value": col_idx.get("average_order_value", 4),
    })

    # Aggregate Meta summary for dashboard
    total_spend  = sum(safe_float(c.get("spend", 0)) for c in meta_campaigns)
    total_impr   = sum(int(safe_float(c.get("impressions", 0))) for c in meta_campaigns)
    total_clicks = sum(int(safe_float(c.get("clicks", 0))) for c in meta_campaigns)
    total_purch  = sum(extract_purchases(c.get("actions", [])) for c in meta_campaigns)
    total_ctr    = round(total_clicks / total_impr * 100, 2) if total_impr else 0
    total_cpa    = round(total_spend / total_purch, 2) if total_purch else 0
    roas_vals    = [extract_roas(c.get("purchase_roas")) for c in meta_campaigns
                    if extract_roas(c.get("purchase_roas")) is not None]
    blended_roas = round(avg(roas_vals), 2) if roas_vals else 0
    meta_summary = {
        "spend": round(total_spend, 2),
        "impressions": total_impr,
        "clicks": total_clicks,
        "purchases": total_purch,
        "ctr": total_ctr,
        "cpa": total_cpa,
        "roas": blended_roas,
    }

    # ── Build each tab ──
    build_shopify_daily_tab(
        tab_map["Shopify Daily Data"], sheets_svc, ss.id,
        sheet_ids["Shopify Daily Data"], shopify_today, seven_day_rows, seven_day_cols
    )
    build_meta_daily_tab(
        tab_map["Meta Ads Daily Data"], sheets_svc, ss.id,
        sheet_ids["Meta Ads Daily Data"], meta_campaigns, []
    )
    build_product_tab(
        tab_map["Product Performance"], sheets_svc, ss.id,
        sheet_ids["Product Performance"], shopify_products_today, shopify_products_7d
    )
    build_campaign_tab(
        tab_map["Campaign Performance"], sheets_svc, ss.id,
        sheet_ids["Campaign Performance"], meta_campaigns, meta_adsets
    )
    build_recommendations_tab(
        tab_map["Recommendations & Notes"], sheets_svc, ss.id,
        sheet_ids["Recommendations & Notes"], recs, notes, date_str, unavailable
    )
    build_dashboard_tab(
        tab_map["Dashboard"], sheets_svc, ss.id,
        sheet_ids["Dashboard"], shopify_today, shopify_7d_avg,
        meta_summary, recs, date_str,
        sheet_ids["Shopify Daily Data"], sheet_ids["Meta Ads Daily Data"]
    )

    return ss.url, gmail_svc, meta_summary, shopify_7d_avg


# ─────────────────────── Email ─────────────────────────────────────────────────

def send_or_draft_email(gmail_svc, sheet_url: str, date_str: str,
                         shopify_today: dict, shopify_7d_avg: dict,
                         meta_summary: dict, recs: list,
                         urgent_issues: list):
    """Send or create a draft email with the daily summary."""
    import base64
    from email.mime.text import MIMEText

    s = shopify_today.get("sales", {})
    gs = safe_float(s.get("gross_sales", 0))
    ns = safe_float(s.get("net_sales", 0))
    od = int(safe_float(s.get("orders", 0)))
    ao = safe_float(s.get("average_order_value", 0))
    rt = abs(safe_float(s.get("returns", 0)))

    def fmt_inr(v): return f"₹{v:,.0f}"
    def fmt_pct(v): return (f"+{v:.1f}%" if v >= 0 else f"{v:.1f}%") if v is not None else "—"

    ag_gs = shopify_7d_avg.get("gross_sales", 0)
    ag_ns = shopify_7d_avg.get("net_sales", 0)
    ag_od = shopify_7d_avg.get("orders", 0)

    sp   = meta_summary.get("spend", 0)
    pu   = meta_summary.get("purchases", 0)
    cpa  = meta_summary.get("cpa", 0)
    roas = meta_summary.get("roas", 0)

    urgent_block = ""
    if urgent_issues:
        urgent_block = "\n🚨 URGENT ISSUES\n" + "\n".join(f"  • {i}" for i in urgent_issues) + "\n\n"

    body = f"""\
Hi Atul,

{urgent_block}Here is your daily performance summary for {date_str}.

━━━━━━━━━━━━━━━━━━━━━━━━━━━
📦 SHOPIFY PERFORMANCE
━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Gross Sales:     {fmt_inr(gs)}   ({fmt_pct(pct_change(gs, ag_gs))} vs 7-day avg {fmt_inr(ag_gs)})
  Net Sales:       {fmt_inr(ns)}   ({fmt_pct(pct_change(ns, ag_ns))} vs 7-day avg {fmt_inr(ag_ns)})
  Orders:          {od}           ({fmt_pct(pct_change(od, ag_od))} vs 7-day avg {round(ag_od, 1)})
  Avg Order Value: {fmt_inr(ao)}
  Returns:         {fmt_inr(rt)}

━━━━━━━━━━━━━━━━━━━━━━━━━━━
📣 META ADS PERFORMANCE
━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Total Spend:     {fmt_inr(sp)}
  Purchases:       {pu}
  CPA:             {fmt_inr(cpa)}
  Blended ROAS:    {roas:.2f}x

━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ RECOMMENDED ACTIONS FOR TODAY
━━━━━━━━━━━━━━━━━━━━━━━━━━━
{chr(10).join(f"  {i+1}. {r}" for i, r in enumerate(recs[:5]))}

━━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 Full Dashboard: {sheet_url}
━━━━━━━━━━━━━━━━━━━━━━━━━━━

Best,
Dhirai Daily Bot
"""

    html_body = body.replace("\n", "<br>").replace("━", "─")

    msg = MIMEText(html_body, "html")
    msg["To"]      = GMAIL_TO
    msg["Subject"] = f"Daily Store & Ads Performance Sheet - {date_str}"
    if GMAIL_CC:
        msg["Cc"] = GMAIL_CC

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")

    try:
        gmail_svc.users().messages().send(
            userId="me", body={"raw": raw}
        ).execute()
        print(f"✅ Email sent to {GMAIL_TO}")
    except HttpError as e:
        # Fallback: create draft
        print(f"⚠️  Send failed ({e}), creating draft instead...")
        gmail_svc.users().drafts().create(
            userId="me", body={"message": {"raw": raw}}
        ).execute()
        print("✅ Gmail draft created.")


# ─────────────────────── Main ──────────────────────────────────────────────────

def main():
    yesterday     = yesterday_kolkata()
    seven_ago     = yesterday - datetime.timedelta(days=6)   # 7-day window starts here
    date_str      = yesterday.isoformat()
    since_7d      = seven_ago.isoformat()
    unavailable   = []

    print(f"📅 Running dashboard for {date_str} (Asia/Kolkata)")

    # ── Fetch Shopify ──
    print("🛒 Fetching Shopify data...")
    try:
        shopify_today = fetch_shopify_daily(date_str)
    except Exception as e:
        print(f"  ⚠️ Shopify daily fetch failed: {e}")
        shopify_today = {"sales": {}, "sessions": {}, "top_products": {"rows": [], "columns": []}}
        unavailable.append(f"Shopify daily data unavailable: {e}")

    try:
        shopify_7d = fetch_shopify_7day(since_7d, date_str)
    except Exception as e:
        print(f"  ⚠️ Shopify 7-day fetch failed: {e}")
        shopify_7d = {"rows": [], "columns": []}
        unavailable.append(f"Shopify 7-day data unavailable: {e}")

    try:
        shopify_prod_today = fetch_shopify_daily(date_str)["top_products"] if not unavailable else {"rows": [], "columns": []}
    except Exception:
        shopify_prod_today = {"rows": [], "columns": []}

    try:
        shopify_prod_7d = fetch_shopify_7day_products(since_7d, date_str)
    except Exception as e:
        shopify_prod_7d = {"rows": [], "columns": []}
        unavailable.append(f"Shopify 7-day product data unavailable: {e}")

    # ── Fetch Meta Ads ──
    print("📣 Fetching Meta Ads data...")
    try:
        meta_campaigns = fetch_meta_campaigns(date_str)
    except Exception as e:
        print(f"  ⚠️ Meta campaign fetch failed: {e}")
        meta_campaigns = []
        unavailable.append(f"Meta campaign data unavailable: {e}")

    try:
        meta_adsets = fetch_meta_adsets(date_str)
    except Exception as e:
        print(f"  ⚠️ Meta ad set fetch failed: {e}")
        meta_adsets = []
        unavailable.append(f"Meta ad set data unavailable: {e}")

    # ── Generate recommendations ──
    print("💡 Generating recommendations...")
    seven_day_rows = shopify_7d.get("rows", [])
    seven_day_cols = shopify_7d.get("columns", [])
    col_idx = {c: i for i, c in enumerate(seven_day_cols)}
    prev_rows = [r for r in seven_day_rows
                 if r and str(r[col_idx.get("day", 0)]) != date_str]
    shopify_7d_avg = compute_7day_averages(prev_rows, {
        "gross_sales":         col_idx.get("gross_sales", 1),
        "net_sales":           col_idx.get("net_sales", 2),
        "orders":              col_idx.get("orders", 3),
        "average_order_value": col_idx.get("average_order_value", 4),
    })

    recs, notes = generate_recommendations(
        shopify_today, shopify_7d_avg, meta_campaigns, meta_adsets
    )
    urgent = [r for r in recs if r.startswith("🚨")]

    # ── Update Google Sheet ──
    print("📊 Updating Google Sheet...")
    try:
        sheet_url, gmail_svc, meta_summary, _ = update_spreadsheet(
            date_str, shopify_today, shopify_7d,
            shopify_prod_today, shopify_prod_7d,
            meta_campaigns, meta_adsets,
            recs, notes, unavailable
        )
        print(f"  ✅ Sheet updated: {sheet_url}")
    except Exception as e:
        print(f"  ⚠️ Sheet update failed: {e}")
        sheet_url   = "https://docs.google.com/spreadsheets (configure credentials)"
        meta_summary = {"spend": 0, "impressions": 0, "clicks": 0,
                        "purchases": 0, "ctr": 0, "cpa": 0, "roas": 0}
        try:
            _, _, gmail_svc = get_google_clients()
        except Exception as auth_err:
            print(f"  ⚠️ Gmail auth also failed: {auth_err}")
            return

    # ── Send / draft email ──
    print("✉️  Sending email / creating draft...")
    try:
        send_or_draft_email(
            gmail_svc, sheet_url, date_str,
            shopify_today, shopify_7d_avg, meta_summary, recs, urgent
        )
    except Exception as e:
        print(f"  ⚠️ Email step failed: {e}")

    print("✅ Done.")


if __name__ == "__main__":
    main()
