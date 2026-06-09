#!/usr/bin/env python3
"""
Daily Store & Ads Performance Dashboard
Runs every morning in Asia/Kolkata timezone.
Shopify store : 36dhns-ed.myshopify.com  (Dhirai)
Meta Ads acct : 979830497515712          (Dhirai)

Requirements:
    pip install gspread google-auth google-auth-oauthlib \
                google-api-python-client requests pytz

Environment variables (set in .env or system):
    SHOPIFY_STORE_DOMAIN   e.g. 36dhns-ed.myshopify.com
    SHOPIFY_ACCESS_TOKEN   Private-app Admin API token
    META_ACCESS_TOKEN      Meta Marketing API user access token
    META_AD_ACCOUNT_ID     e.g. 979830497515712
    GOOGLE_SA_CREDENTIALS  Path to Google service-account JSON file
    GMAIL_TO               Recipient e-mail  (atul012001@gmail.com)
    GMAIL_CC               CC e-mail(s), comma-separated (leave blank)
    SPREADSHEET_NAME       Sheet name (default below)
"""

import os
import json
import datetime
import base64
from email.mime.text import MIMEText

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

SHOPIFY_DOMAIN   = os.environ.get("SHOPIFY_STORE_DOMAIN",  "36dhns-ed.myshopify.com")
SHOPIFY_TOKEN    = os.environ.get("SHOPIFY_ACCESS_TOKEN",  "")
META_TOKEN       = os.environ.get("META_ACCESS_TOKEN",     "")
META_ACCOUNT_ID  = os.environ.get("META_AD_ACCOUNT_ID",    "979830497515712")
SA_CREDENTIALS   = os.environ.get("GOOGLE_SA_CREDENTIALS", "service_account.json")
GMAIL_TO         = os.environ.get("GMAIL_TO",              "atul012001@gmail.com")
GMAIL_CC         = os.environ.get("GMAIL_CC",              "")
SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME",
                                  "Daily Store & Ads Performance Sheet")

# ─────────────────────── Colour palette (Sheets RGB 0-1) ──────────────────────

C_HEADER_BG  = {"red": 0.133, "green": 0.133, "blue": 0.133}   # dark charcoal
C_HEADER_FG  = {"red": 1,     "green": 1,     "blue": 1}        # white
C_SECTION_BG = {"red": 0.957, "green": 0.957, "blue": 0.957}    # light grey
C_GREEN      = {"red": 0.204, "green": 0.659, "blue": 0.325}
C_RED        = {"red": 0.839, "green": 0.153, "blue": 0.157}
C_YELLOW     = {"red": 1,     "green": 0.898, "blue": 0.200}
C_ORANGE     = {"red": 1,     "green": 0.596, "blue": 0}
C_ACCENT     = {"red": 0.259, "green": 0.522, "blue": 0.957}    # brand blue
C_WHITE      = {"red": 1,     "green": 1,     "blue": 1}
C_LIGHT_GREEN= {"red": 0.714, "green": 0.929, "blue": 0.714}
C_LIGHT_RED  = {"red": 0.957, "green": 0.714, "blue": 0.714}
C_LIGHT_YELL = {"red": 1,     "green": 0.949, "blue": 0.800}

# ─────────────────────── Google auth ──────────────────────────────────────────

def get_google_clients():
    creds        = Credentials.from_service_account_file(SA_CREDENTIALS, scopes=SCOPES)
    gc           = gspread.authorize(creds)
    sheets_svc   = build("sheets", "v4",  credentials=creds)
    gmail_svc    = build("gmail",  "v1",  credentials=creds)
    return gc, sheets_svc, gmail_svc

# ─────────────────────── Date helpers ─────────────────────────────────────────

def yesterday_ist() -> datetime.date:
    return (datetime.datetime.now(KOLKATA_TZ) - datetime.timedelta(days=1)).date()

def dates_for_7day(yesterday: datetime.date):
    """Return (since, until) for the 7 calendar days BEFORE yesterday."""
    until  = yesterday - datetime.timedelta(days=1)
    since  = yesterday - datetime.timedelta(days=7)
    return since.isoformat(), until.isoformat()

# ─────────────────────── Shopify helpers ──────────────────────────────────────

SHOPIFY_GQL = f"https://{SHOPIFY_DOMAIN}/admin/api/2024-01/graphql.json"

def shopify_query(shopify_ql: str) -> dict:
    """Execute a ShopifyQL query; return {columns, rows}."""
    gql = """
    query Q($q: String!) {
      shopifyqlQuery(query: $q) {
        ... on TableResponse {
          tableData { columns { name dataType } unformattedData rowData }
        }
      }
    }
    """
    r = requests.post(
        SHOPIFY_GQL,
        headers={"Content-Type": "application/json",
                 "X-Shopify-Access-Token": SHOPIFY_TOKEN},
        json={"query": gql, "variables": {"q": shopify_ql}},
        timeout=30,
    )
    r.raise_for_status()
    td = r.json()["data"]["shopifyqlQuery"]["tableData"]
    rows = td.get("unformattedData") or td.get("rowData") or []
    cols = [c["name"] for c in td.get("columns", [])]
    return {"columns": cols, "rows": rows}


def fetch_shopify_day(date_str: str) -> dict:
    """All key Shopify metrics for one calendar day."""
    out = {}

    # Sales summary
    sales = shopify_query(
        f"FROM sales SHOW gross_sales, net_sales, orders, average_order_value, "
        f"total_sales, discounts, returns, shipping_charges, taxes "
        f"SINCE {date_str} UNTIL {date_str}"
    )
    out["sales"] = dict(zip(sales["columns"], sales["rows"][0])) if sales["rows"] else {}

    # Sessions & conversion
    sess = shopify_query(
        f"FROM sessions SHOW sessions, sessions_with_cart_additions, "
        f"sessions_that_reached_checkout, sessions_that_completed_checkout, "
        f"conversion_rate SINCE {date_str} UNTIL {date_str}"
    )
    out["sessions"] = dict(zip(sess["columns"], sess["rows"][0])) if sess["rows"] else {}

    # Top 10 products
    prods = shopify_query(
        f"FROM sales SHOW gross_sales, net_sales, orders, total_sales "
        f"GROUP BY product_title ORDER BY gross_sales DESC LIMIT 10 "
        f"SINCE {date_str} UNTIL {date_str}"
    )
    out["top_products"] = prods

    return out


def fetch_shopify_7day(since: str, until: str) -> dict:
    """Daily timeseries for the 7-day window."""
    return shopify_query(
        f"FROM sales SHOW gross_sales, net_sales, orders, average_order_value, "
        f"total_sales, discounts, returns, shipping_charges, taxes "
        f"TIMESERIES day SINCE {since} UNTIL {until}"
    )


def fetch_shopify_7day_products(since: str, until: str) -> dict:
    return shopify_query(
        f"FROM sales SHOW gross_sales, net_sales, orders "
        f"GROUP BY product_title ORDER BY gross_sales DESC LIMIT 10 "
        f"SINCE {since} UNTIL {until}"
    )

# ─────────────────────── Meta Ads helpers ─────────────────────────────────────

META_BASE = "https://graph.facebook.com/v19.0"
META_FIELDS = (
    "campaign_id,campaign_name,adset_id,adset_name,"
    "spend,impressions,reach,clicks,ctr,cpc,cpm,"
    "purchase_roas,actions,cost_per_action_type,date_start,date_stop"
)


def meta_get(endpoint: str, params: dict) -> dict:
    params["access_token"] = META_TOKEN
    r = requests.get(f"{META_BASE}/{endpoint}", params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def fetch_meta_campaigns(date_str: str) -> list:
    r = meta_get(f"act_{META_ACCOUNT_ID}/insights", {
        "level":      "campaign",
        "fields":     META_FIELDS,
        "time_range": json.dumps({"since": date_str, "until": date_str}),
        "limit":      50,
    })
    return r.get("data", [])


def fetch_meta_adsets(date_str: str) -> list:
    r = meta_get(f"act_{META_ACCOUNT_ID}/insights", {
        "level":      "adset",
        "fields":     META_FIELDS,
        "time_range": json.dumps({"since": date_str, "until": date_str}),
        "limit":      50,
    })
    return r.get("data", [])


def fetch_meta_account_daily(since: str, until: str) -> list:
    """Account-level daily rows for the 7-day window."""
    r = meta_get(f"act_{META_ACCOUNT_ID}/insights", {
        "level":          "account",
        "fields":         "spend,impressions,reach,clicks,ctr,cpc,cpm,purchase_roas,actions",
        "time_range":     json.dumps({"since": since, "until": until}),
        "time_increment": 1,
        "limit":          50,
    })
    return r.get("data", [])


def fetch_meta_anomalies() -> list:
    """Retrieve anomaly signals from the Meta opportunity-score endpoint."""
    try:
        r = meta_get(f"act_{META_ACCOUNT_ID}/insights", {
            "level":       "account",
            "fields":      "spend",
            "date_preset": "yesterday",
            "limit":       1,
        })
        # Anomaly detection via the scoring endpoint (best-effort)
        score_r = meta_get(f"act_{META_ACCOUNT_ID}/adfixsuggestions", {
            "fields": "suggestion_type,suggestion_title,suggestion_description",
            "limit":  10,
        })
        return [
            f"{s.get('suggestion_title','')}: {s.get('suggestion_description','')}"
            for s in score_r.get("data", [])
        ]
    except Exception:
        return []

# ─────────────────────── Numeric helpers ──────────────────────────────────────

def sf(v) -> float:
    """Safe float conversion – strips currency symbols and commas."""
    try:
        return float(str(v).replace(",","").replace("₹","").replace(" INR","").strip())
    except (ValueError, TypeError):
        return 0.0


def pct_change(curr: float, prior: float):
    if prior == 0:
        return None
    return round((curr - prior) / prior * 100, 1)


def mean(vals: list) -> float:
    clean = [v for v in vals if v is not None]
    return sum(clean) / len(clean) if clean else 0.0


def col_avgs(rows: list, col_map: dict) -> dict:
    """Average each column by index across a list of rows."""
    out = {}
    for name, idx in col_map.items():
        vals = []
        for row in rows:
            try:
                vals.append(float(row[idx]))
            except (ValueError, TypeError, IndexError):
                pass
        out[name] = round(mean(vals), 2)
    return out


def extract_purchases(actions) -> int:
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


def extract_roas(purchase_roas):
    for r in (purchase_roas or []):
        if r.get("action_type") in (
            "offsite_conversion.fb_pixel_purchase", "omni_purchase"
        ):
            try:
                return round(float(r["value"]), 2)
            except (ValueError, KeyError):
                pass
    return None


def extract_cpa(cost_per_action_type):
    for a in (cost_per_action_type or []):
        if a.get("action_type") in (
            "offsite_conversion.fb_pixel_purchase", "purchase"
        ):
            try:
                return round(float(a["value"]), 2)
            except (ValueError, KeyError):
                pass
    return None


def aggregate_meta(rows: list) -> dict:
    """Aggregate a list of campaign/account rows into one summary dict."""
    spend  = sum(sf(r.get("spend", 0)) for r in rows)
    impr   = sum(int(sf(r.get("impressions", 0))) for r in rows)
    reach  = sum(int(sf(r.get("reach",  0))) for r in rows)
    clicks = sum(int(sf(r.get("clicks", 0))) for r in rows)
    purch  = sum(extract_purchases(r.get("actions")) for r in rows)
    ctr    = round(clicks / impr * 100, 2) if impr else 0
    cpc    = round(spend / clicks, 2)      if clicks else 0
    cpm    = round(spend / impr * 1000, 2) if impr else 0
    cpa    = round(spend / purch, 2)       if purch else 0
    roas_v = [extract_roas(r.get("purchase_roas")) for r in rows]
    roas   = round(mean([v for v in roas_v if v is not None]), 2)
    return dict(spend=round(spend,2), impressions=impr, reach=reach,
                clicks=clicks, purchases=purch, ctr=ctr, cpc=cpc,
                cpm=cpm, cpa=cpa, roas=roas)

# ─────────────────────── Recommendations engine ───────────────────────────────

def generate_recommendations(shop_today: dict, shop_7d_avg: dict,
                              meta_campaigns: list, meta_adsets: list,
                              meta_today: dict, meta_7d_avg: dict,
                              anomalies: list = None) -> tuple:
    recs  = []
    notes = []

    gs   = sf(shop_today.get("sales", {}).get("gross_sales", 0))
    od   = sf(shop_today.get("sales", {}).get("orders",      0))
    rt   = abs(sf(shop_today.get("sales", {}).get("returns", 0)))
    ag_g = shop_7d_avg.get("gross_sales", 0)
    ag_o = shop_7d_avg.get("orders", 0)

    # Revenue change
    chg_gs = pct_change(gs, ag_g)
    if chg_gs is not None:
        if chg_gs < -15:
            recs.append(f"🚨 Revenue down {abs(chg_gs):.1f}% vs 7-day avg — "
                        "check ad delivery, site speed, and product availability.")
        elif chg_gs > 20:
            recs.append(f"✅ Revenue up {chg_gs:.1f}% vs 7-day avg — "
                        "identify top campaigns and scale budgets by 20–30%.")

    # Orders change
    if ag_o > 0 and od < ag_o * 0.80:
        recs.append("⚠️ Orders below 80% of 7-day avg — review traffic sources "
                    "and Meta campaign delivery.")

    # Return rate
    if gs > 0 and rt / gs > 0.12:
        recs.append(f"⚠️ Return rate {rt/gs*100:.1f}% — check product sizing, "
                    "descriptions, and customer complaints.")

    # Meta ROAS change
    roas_today = meta_today.get("roas", 0)
    roas_7d    = meta_7d_avg.get("roas", 0)
    chg_roas   = pct_change(roas_today, roas_7d)
    if chg_roas is not None:
        if chg_roas > 20:
            recs.append(f"✅ ROAS jumped {chg_roas:.1f}% vs 7-day avg ({roas_today:.2f}x) — "
                        "scale spend on top campaigns today.")
        elif chg_roas < -20:
            recs.append(f"⚠️ ROAS dropped {abs(chg_roas):.1f}% vs 7-day avg "
                        f"({roas_today:.2f}x) — audit creative fatigue and audience overlap.")

    # Campaign-level checks
    for c in meta_campaigns:
        spend = sf(c.get("spend", 0))
        roas  = extract_roas(c.get("purchase_roas"))
        name  = c.get("campaign_name", c.get("name", ""))
        if roas is not None and spend > 500:
            if roas < 1.0:
                recs.append(f"🚨 PAUSE '{name}' — ROAS {roas:.2f}x below breakeven. "
                            f"Reallocate ₹{spend:,.0f}/day budget.")
            elif roas < 1.8:
                recs.append(f"⚠️ Low ROAS on '{name}' ({roas:.2f}x) — test new creative "
                            "or narrow audience.")
            elif roas > 5.0 and spend > 1000:
                recs.append(f"✅ Scale '{name}' — ROAS {roas:.2f}x is excellent. "
                            "Increase daily budget by 20–30%.")

    # Ad-set CTR check
    for a in meta_adsets:
        spend = sf(a.get("spend", 0))
        ctr   = sf(a.get("ctr",   0))
        name  = a.get("adset_name", a.get("name", ""))
        if spend > 500 and ctr < 0.8:
            recs.append(f"⚠️ Low CTR ({ctr:.2f}%) on '{name}' — refresh creative or hook.")

    # Anomaly signals
    for sig in (anomalies or []):
        recs.append(f"📊 Meta signal: {sig}")

    # Always-useful reminders
    if not any("abandoned" in r.lower() for r in recs):
        recs.append("🔍 Check abandoned checkout recovery flows — "
                    "high COD rate means many hesitate at payment.")
    recs.append("📱 Review top-selling product pages — ensure sizes, images, "
                "and stock levels are up to date.")

    if not recs or all(r.startswith(("🔍", "📱")) for r in recs):
        recs.insert(0, "✅ All campaigns within normal range — "
                       "continue monitoring and run creative tests.")

    return recs[:7], notes

# ─────────────────────── Sheets format helpers ────────────────────────────────

def fmt_req(sheet_id, r0, r1, c0, c1, *, bold=False,
            bg=None, fg=None, align=None, num_fmt=None, font_size=None):
    tf = {}
    if bold:      tf["bold"] = True
    if fg:        tf["foregroundColor"] = fg
    if font_size: tf["fontSize"] = font_size
    cell_fmt = {}
    if tf:               cell_fmt["textFormat"]        = tf
    if bg:               cell_fmt["backgroundColor"]   = bg
    if align:            cell_fmt["horizontalAlignment"] = align
    if num_fmt:          cell_fmt["numberFormat"] = {"type": "NUMBER", "pattern": num_fmt}
    fields = ",".join(
        f for f, cond in [
            ("textFormat",         bool(tf)),
            ("backgroundColor",    bool(bg)),
            ("horizontalAlignment",bool(align)),
            ("numberFormat",       bool(num_fmt)),
        ] if cond
    )
    return {"repeatCell": {
        "range": {"sheetId": sheet_id,
                  "startRowIndex": r0, "endRowIndex": r1,
                  "startColumnIndex": c0, "endColumnIndex": c1},
        "cell": {"userEnteredFormat": cell_fmt},
        "fields": f"userEnteredFormat({fields})",
    }}


def freeze_req(sheet_id, rows=1, cols=0):
    return {"updateSheetProperties": {
        "properties": {"sheetId": sheet_id,
                       "gridProperties": {"frozenRowCount": rows,
                                          "frozenColumnCount": cols}},
        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
    }}


def autoresize_req(sheet_id, c0=0, c1=26):
    return {"autoResizeDimensions": {
        "dimensions": {"sheetId": sheet_id, "dimension": "COLUMNS",
                       "startIndex": c0, "endIndex": c1},
    }}


def filter_req(sheet_id, r0, r1, c0, c1):
    return {"setBasicFilter": {"filter": {"range": {
        "sheetId": sheet_id,
        "startRowIndex": r0, "endRowIndex": r1,
        "startColumnIndex": c0, "endColumnIndex": c1,
    }}}}


def cond_pct(sheet_id, r0, r1, c0, c1):
    """Green if > 0, red if < 0."""
    def rule(cond_type, val, colour):
        return {"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [{"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1,
                        "startColumnIndex": c0, "endColumnIndex": c1}],
            "booleanRule": {
                "condition": {"type": cond_type,
                              "values": [{"userEnteredValue": val}]},
                "format": {"backgroundColor": colour},
            },
        }}}
    return [rule("NUMBER_GREATER", "0", C_LIGHT_GREEN),
            rule("NUMBER_LESS",    "0", C_LIGHT_RED)]


def cond_roas(sheet_id, r0, r1, c0, c1):
    def rule(cond_type, vals, colour):
        return {"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [{"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1,
                        "startColumnIndex": c0, "endColumnIndex": c1}],
            "booleanRule": {
                "condition": {"type": cond_type,
                              "values": [{"userEnteredValue": v} for v in vals]},
                "format": {"backgroundColor": colour},
            },
        }}}
    return [
        rule("NUMBER_LESS",    ["1.5"], C_LIGHT_RED),
        rule("NUMBER_BETWEEN", ["1.5", "3.0"], C_LIGHT_YELL),
        rule("NUMBER_GREATER", ["3.0"], C_LIGHT_GREEN),
    ]


def batch(sheets_svc, spreadsheet_id, reqs):
    if reqs:
        sheets_svc.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": reqs}
        ).execute()

# ─────────────────────── Tab: Shopify Daily Data ──────────────────────────────

SHOPIFY_HEADERS = [
    "Date", "Gross Sales (₹)", "Net Sales (₹)", "Orders",
    "AOV (₹)", "Total Sales (₹)", "Discounts (₹)", "Returns (₹)",
    "Shipping (₹)", "Taxes (₹)",
    "Sessions", "Cart Adds", "Reached Checkout", "Completed Checkout",
    "Conversion Rate (%)",
]


def _shopify_row(date_str: str, sales: dict, sessions: dict) -> list:
    return [
        date_str,
        round(sf(sales.get("gross_sales", 0)), 2),
        round(sf(sales.get("net_sales", 0)), 2),
        int(sf(sales.get("orders", 0))),
        round(sf(sales.get("average_order_value", 0)), 2),
        round(sf(sales.get("total_sales", 0)), 2),
        round(abs(sf(sales.get("discounts", 0))), 2),
        round(abs(sf(sales.get("returns", 0))), 2),
        round(sf(sales.get("shipping_charges", 0)), 2),
        round(sf(sales.get("taxes", 0)), 2),
        int(sf(sessions.get("sessions", 0))) or "",
        int(sf(sessions.get("sessions_with_cart_additions", 0))) or "",
        int(sf(sessions.get("sessions_that_reached_checkout", 0))) or "",
        int(sf(sessions.get("sessions_that_completed_checkout", 0))) or "",
        round(sf(sessions.get("conversion_rate", 0)), 4) or "",
    ]


def build_shopify_daily_tab(ws, sheets_svc, spreadsheet_id, sheet_id,
                             today_data: dict, seven_day_data: dict, date_str: str):
    existing = ws.get_all_values()

    # Ensure header
    if not existing or existing[0] != SHOPIFY_HEADERS:
        ws.clear()
        ws.update("A1", [SHOPIFY_HEADERS])
        existing = [SHOPIFY_HEADERS]

    existing_dates = {row[0] for row in existing[1:] if row}

    # Append 7-day historical rows (any missing)
    cols = seven_day_data.get("columns", [])
    ci   = {c: i for i, c in enumerate(cols)}
    new_rows = []
    for row in seven_day_data.get("rows", []):
        d = str(row[ci.get("day", 0)])
        if d not in existing_dates and d != date_str:
            new_rows.append(_shopify_row(d, {
                "gross_sales":        row[ci.get("gross_sales",        1)],
                "net_sales":          row[ci.get("net_sales",          2)],
                "orders":             row[ci.get("orders",             3)],
                "average_order_value":row[ci.get("average_order_value",4)],
                "total_sales":        row[ci.get("total_sales",        5)],
                "discounts":          row[ci.get("discounts",          6)],
                "returns":            row[ci.get("returns",            7)],
                "shipping_charges":   row[ci.get("shipping_charges",   8)],
                "taxes":              row[ci.get("taxes",              9)],
            }, {}))
            existing_dates.add(d)

    # Append today's row (if not already present)
    if date_str not in existing_dates:
        new_rows.append(_shopify_row(
            date_str,
            today_data.get("sales", {}),
            today_data.get("sessions", {}),
        ))

    if new_rows:
        next_row = len(existing) + 1
        ws.update(f"A{next_row}", new_rows)

    total_rows = len(existing) + len(new_rows) + 1

    reqs = [
        fmt_req(sheet_id, 0, 1, 0, len(SHOPIFY_HEADERS),
                bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG),
        freeze_req(sheet_id, rows=1),
        filter_req(sheet_id, 0, total_rows, 0, len(SHOPIFY_HEADERS)),
        autoresize_req(sheet_id, 0, len(SHOPIFY_HEADERS)),
    ]
    # Currency columns: B-J (indices 1-9)
    for col in range(1, 10):
        reqs.append(fmt_req(sheet_id, 1, total_rows, col, col+1,
                            num_fmt='[<10000]₹#,##0.00;₹#,##0.00'))
    # Conversion rate: O (index 14)
    reqs.append(fmt_req(sheet_id, 1, total_rows, 14, 15, num_fmt='0.00"%"'))
    batch(sheets_svc, spreadsheet_id, reqs)

# ─────────────────────── Tab: Meta Ads Daily Data ─────────────────────────────

META_DAILY_HEADERS = [
    "Date", "Spend (₹)", "Impressions", "Reach", "Clicks",
    "CTR (%)", "CPC (₹)", "CPM (₹)", "Purchases", "CPA (₹)",
    "ROAS", "vs 7d Avg Spend", "vs 7d Avg ROAS",
]


def build_meta_daily_tab(ws, sheets_svc, spreadsheet_id, sheet_id,
                          campaigns: list, account_7d: list, date_str: str):
    existing = ws.get_all_values()
    if not existing or existing[0] != META_DAILY_HEADERS:
        ws.clear()
        ws.update("A1", [META_DAILY_HEADERS])
        existing = [META_DAILY_HEADERS]

    existing_dates = {row[0] for row in existing[1:] if row}

    # Append 7-day account daily rows (missing)
    new_rows = []
    for daily in account_7d:
        d = daily.get("date_start", "")
        if d and d not in existing_dates and d != date_str:
            s = aggregate_meta([daily])
            new_rows.append([
                d, s["spend"], s["impressions"], s["reach"], s["clicks"],
                s["ctr"], s["cpc"], s["cpm"], s["purchases"], s["cpa"],
                s["roas"], "", "",
            ])
            existing_dates.add(d)

    # Append today's row
    if date_str not in existing_dates:
        today = aggregate_meta(campaigns)
        # 7-day averages from account_7d
        sp_7d   = mean([sf(r.get("spend", 0))          for r in account_7d]) if account_7d else 0
        roas_7d = mean([extract_roas(r.get("purchase_roas")) or 0
                        for r in account_7d])                                 if account_7d else 0
        chg_sp  = pct_change(today["spend"], sp_7d)
        chg_ro  = pct_change(today["roas"],  roas_7d)
        new_rows.append([
            date_str,
            today["spend"], today["impressions"], today["reach"], today["clicks"],
            today["ctr"], today["cpc"], today["cpm"],
            today["purchases"], today["cpa"], today["roas"],
            round(chg_sp, 1) if chg_sp is not None else "",
            round(chg_ro, 1) if chg_ro is not None else "",
        ])

    if new_rows:
        next_row = len(existing) + 1
        ws.update(f"A{next_row}", new_rows)

    total_rows = len(existing) + len(new_rows) + 1
    reqs = [
        fmt_req(sheet_id, 0, 1, 0, len(META_DAILY_HEADERS),
                bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG),
        freeze_req(sheet_id, rows=1),
        filter_req(sheet_id, 0, total_rows, 0, len(META_DAILY_HEADERS)),
        autoresize_req(sheet_id, 0, len(META_DAILY_HEADERS)),
    ]
    for col in [1, 6, 7, 9]:   # Spend, CPC, CPM, CPA
        reqs.append(fmt_req(sheet_id, 1, total_rows, col, col+1,
                            num_fmt='₹#,##0.00'))
    reqs.append(fmt_req(sheet_id, 1, total_rows, 5, 6,  num_fmt='0.00"%"'))  # CTR
    reqs.append(fmt_req(sheet_id, 1, total_rows, 11, 13, num_fmt='0.0"%"'))  # Δ columns
    reqs += cond_pct(sheet_id, 1, total_rows, 11, 13)
    reqs += cond_roas(sheet_id, 1, total_rows, 10, 11)
    batch(sheets_svc, spreadsheet_id, reqs)

# ─────────────────────── Tab: Product Performance ─────────────────────────────

def build_product_tab(ws, sheets_svc, spreadsheet_id, sheet_id,
                       today_prods: dict, week_prods: dict, date_str: str):
    ws.clear()
    section_row = [f"Product Performance — Yesterday {date_str}",
                   "", "", "", "", "── 7-Day Total ──", "", ""]
    col_row = ["Product", "Gross Sales (₹)", "Net Sales (₹)", "Orders",
               "Total Sales (₹)", "7d Gross (₹)", "7d Net (₹)", "7d Orders"]
    ws.update("A1", [section_row, col_row])

    week_lookup = {str(r[0]): r for r in week_prods.get("rows", [])}
    data = []
    for row in today_prods.get("rows", []):
        prod = str(row[0])
        w    = week_lookup.get(prod, [])
        data.append([
            prod,
            round(sf(row[1]), 2),
            round(sf(row[2]), 2),
            int(sf(row[3])),
            round(sf(row[4]), 2) if len(row) > 4 else "",
            round(sf(w[1]), 2) if len(w) > 1 else "",
            round(sf(w[2]), 2) if len(w) > 2 else "",
            int(sf(w[3]))      if len(w) > 3 else "",
        ])
    if data:
        ws.update("A3", data)

    total = len(data) + 3
    reqs = [
        fmt_req(sheet_id, 0, 1, 0, 8, bold=True, bg=C_ACCENT,   fg=C_HEADER_FG, font_size=11),
        fmt_req(sheet_id, 1, 2, 0, 8, bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG),
        freeze_req(sheet_id, rows=2),
        filter_req(sheet_id, 1, total, 0, 8),
        autoresize_req(sheet_id, 0, 8),
    ]
    for col in [1, 2, 4, 5, 6]:
        reqs.append(fmt_req(sheet_id, 2, total, col, col+1, num_fmt='₹#,##0.00'))
    batch(sheets_svc, spreadsheet_id, reqs)

# ─────────────────────── Tab: Campaign Performance ────────────────────────────

CAMP_HDRS  = ["Campaign", "Status", "Spend (₹)", "Impr.", "Reach",
              "Clicks", "CTR (%)", "CPC (₹)", "CPM (₹)",
              "Purchases", "CPA (₹)", "ROAS"]
ADSET_HDRS = ["Ad Set", "Status", "Spend (₹)", "Impr.", "Reach",
              "Clicks", "CTR (%)", "CPC (₹)", "CPM (₹)",
              "Purchases", "CPA (₹)", "ROAS", "Learning Phase"]


def build_campaign_tab(ws, sheets_svc, spreadsheet_id, sheet_id,
                        campaigns: list, adsets: list):
    ws.clear()
    ws.update("A1",  [["━━━ CAMPAIGN LEVEL ━━━"] + [""] * (len(CAMP_HDRS)  - 1)])
    ws.update("A2",  [CAMP_HDRS])

    cdata = []
    for c in sorted(campaigns, key=lambda x: -sf(x.get("spend", 0))):
        spend = sf(c.get("spend", 0))
        if spend == 0:
            continue
        cdata.append([
            c.get("campaign_name", ""),
            c.get("effective_status", c.get("status", "")),
            round(spend, 2),
            int(sf(c.get("impressions", 0))),
            int(sf(c.get("reach", 0))),
            int(sf(c.get("clicks", 0))),
            round(sf(c.get("ctr", 0)), 2),
            round(sf(c.get("cpc", 0)), 2),
            round(sf(c.get("cpm", 0)), 2),
            extract_purchases(c.get("actions")),
            extract_cpa(c.get("cost_per_action_type")) or "",
            extract_roas(c.get("purchase_roas")) or "",
        ])
    if cdata:
        ws.update("A3", cdata)
    camp_end = len(cdata) + 3

    as_start = camp_end + 1
    ws.update(f"A{as_start}",   [["━━━ AD SET LEVEL ━━━"] + [""] * (len(ADSET_HDRS) - 1)])
    ws.update(f"A{as_start+1}", [ADSET_HDRS])

    adata = []
    for a in sorted(adsets, key=lambda x: -sf(x.get("spend", 0))):
        spend = sf(a.get("spend", 0))
        if spend == 0:
            continue
        d = a.get("delivery", {})
        sub = d.get("substatuses", []) if isinstance(d, dict) else []
        learning = ("In Learning"        if "in_learning_phase"              in sub else
                    "Learning Exit Fail" if "learning_exit_unsuccessfully"   in sub else "")
        adata.append([
            a.get("adset_name", ""),
            a.get("effective_status", a.get("status", "")),
            round(spend, 2),
            int(sf(a.get("impressions", 0))),
            int(sf(a.get("reach", 0))),
            int(sf(a.get("clicks", 0))),
            round(sf(a.get("ctr", 0)), 2),
            round(sf(a.get("cpc", 0)), 2),
            round(sf(a.get("cpm", 0)), 2),
            extract_purchases(a.get("actions")),
            extract_cpa(a.get("cost_per_action_type")) or "",
            extract_roas(a.get("purchase_roas")) or "",
            learning,
        ])
    if adata:
        ws.update(f"A{as_start+2}", adata)
    as_end = as_start + 2 + len(adata)

    reqs = [
        fmt_req(sheet_id, 0, 1, 0, len(CAMP_HDRS), bold=True, bg=C_ACCENT, fg=C_HEADER_FG, font_size=11),
        fmt_req(sheet_id, 1, 2, 0, len(CAMP_HDRS), bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG),
        fmt_req(sheet_id, as_start-1, as_start, 0, len(ADSET_HDRS), bold=True, bg=C_ACCENT, fg=C_HEADER_FG, font_size=11),
        fmt_req(sheet_id, as_start, as_start+1, 0, len(ADSET_HDRS), bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG),
        freeze_req(sheet_id, rows=2),
        autoresize_req(sheet_id, 0, len(ADSET_HDRS)),
    ]
    # Currency / pct for campaigns
    for col in [2, 7, 8, 10]:
        reqs.append(fmt_req(sheet_id, 2, camp_end, col, col+1, num_fmt='₹#,##0.00'))
    reqs.append(fmt_req(sheet_id, 2, camp_end, 6, 7, num_fmt='0.00"%"'))
    reqs += cond_roas(sheet_id, 2, camp_end, 11, 12)
    # Currency / pct for adsets
    for col in [2, 7, 8, 10]:
        reqs.append(fmt_req(sheet_id, as_start+1, as_end, col, col+1, num_fmt='₹#,##0.00'))
    reqs.append(fmt_req(sheet_id, as_start+1, as_end, 6, 7, num_fmt='0.00"%"'))
    reqs += cond_roas(sheet_id, as_start+1, as_end, 11, 12)
    batch(sheets_svc, spreadsheet_id, reqs)

# ─────────────────────── Tab: Recommendations & Notes ────────────────────────

def build_recs_tab(ws, sheets_svc, spreadsheet_id, sheet_id,
                    recs: list, notes: list, date_str: str, unavailable: list):
    ws.clear()
    rows = [
        [f"RECOMMENDATIONS & NOTES — {date_str}", ""],
        ["", ""],
        ["TYPE", "RECOMMENDATION / NOTE"],
    ]
    for rec in recs:
        icon = rec[:2] if rec[:1] in ("⚠", "🚨", "✅", "📊", "🔍", "📱", "💡") else "💡"
        rows.append([icon, rec[len(icon):].strip()])
    if unavailable:
        rows += [["", ""], ["UNAVAILABLE METRICS", ""]]
        for m in unavailable:
            rows.append(["ℹ️", m])
    if notes:
        rows += [["", ""], ["NOTES", ""]]
        for n in notes:
            rows.append(["📌", n])

    ws.update("A1", rows)
    n = len(rows)
    reqs = [
        fmt_req(sheet_id, 0, 1, 0, 2, bold=True, bg=C_ACCENT, fg=C_HEADER_FG, font_size=13),
        fmt_req(sheet_id, 2, 3, 0, 2, bold=True, bg=C_HEADER_BG, fg=C_HEADER_FG),
        freeze_req(sheet_id, rows=3),
        autoresize_req(sheet_id, 0, 2),
    ]
    batch(sheets_svc, spreadsheet_id, reqs)

# ─────────────────────── Tab: Dashboard ──────────────────────────────────────

def build_dashboard_tab(ws, sheets_svc, spreadsheet_id, sheet_id,
                         shop_today: dict, shop_7d_avg: dict,
                         meta_today: dict, meta_7d_avg: dict,
                         recs: list, date_str: str):
    ws.clear()

    def inr(v):  return f"₹{sf(v):,.0f}"
    def pct(v):
        if v is None: return "—"
        return (f"+{v:.1f}%" if v >= 0 else f"{v:.1f}%")

    s   = shop_today.get("sales", {})
    gs  = sf(s.get("gross_sales", 0));  ag_gs = shop_7d_avg.get("gross_sales", 0)
    ns  = sf(s.get("net_sales", 0));    ag_ns = shop_7d_avg.get("net_sales", 0)
    od  = int(sf(s.get("orders", 0)));  ag_od = shop_7d_avg.get("orders", 0)
    ao  = sf(s.get("average_order_value", 0)); ag_ao = shop_7d_avg.get("average_order_value", 0)
    rt  = abs(sf(s.get("returns", 0)))

    sp  = meta_today.get("spend", 0);   ag_sp  = meta_7d_avg.get("spend", 0)
    imp = meta_today.get("impressions", 0)
    cl  = meta_today.get("clicks", 0)
    pu  = meta_today.get("purchases", 0); ag_pu = meta_7d_avg.get("purchases", 0)
    ctr = meta_today.get("ctr", 0)
    cpa = meta_today.get("cpa", 0)
    ro  = meta_today.get("roas", 0);    ag_ro  = meta_7d_avg.get("roas", 0)

    wins   = [r for r in recs if r.startswith("✅")]
    issues = [r for r in recs if r.startswith(("⚠️", "🚨"))]

    rows = [
        [f"DAILY STORE & ADS PERFORMANCE — {date_str}", "", "", ""],
        ["", "", "", ""],
        ["━━━ SHOPIFY PERFORMANCE ━━━", "", "", ""],
        ["Metric", "Yesterday", "7-Day Avg", "Change vs Avg"],
        ["Gross Sales",     inr(gs), inr(ag_gs), pct(pct_change(gs, ag_gs))],
        ["Net Sales",       inr(ns), inr(ag_ns), pct(pct_change(ns, ag_ns))],
        ["Orders",          str(od), str(round(ag_od, 1)), pct(pct_change(od, ag_od))],
        ["Avg Order Value", inr(ao), inr(ag_ao), pct(pct_change(ao, ag_ao))],
        ["Returns",         inr(rt), "—", ""],
        ["", "", "", ""],
        ["━━━ META ADS PERFORMANCE ━━━", "", "", ""],
        ["Metric", "Yesterday", "7-Day Avg", "Change vs Avg"],
        ["Total Spend",    inr(sp),  inr(ag_sp),     pct(pct_change(sp,  ag_sp))],
        ["Impressions",    f"{imp:,}", "—",           ""],
        ["Clicks",         f"{cl:,}", "—",            ""],
        ["CTR",            f"{ctr:.2f}%", "—",        ""],
        ["Purchases (attr.)", str(pu), str(round(ag_pu,1)), pct(pct_change(pu, ag_pu))],
        ["CPA",            inr(cpa), "—",             ""],
        ["Blended ROAS",   f"{ro:.2f}x", f"{ag_ro:.2f}x", pct(pct_change(ro, ag_ro))],
        ["", "", "", ""],
        ["━━━ KEY WINS ━━━", "", "", ""],
    ]
    for w in (wins[:3] or ["No standout wins — check campaign performance."]):
        rows.append(["", w, "", ""])
    rows += [["", "", "", ""], ["━━━ KEY ISSUES ━━━", "", "", ""]]
    for i in (issues[:3] or ["No critical issues."]):
        rows.append(["", i, "", ""])
    rows += [["", "", "", ""], ["━━━ RECOMMENDED ACTIONS TODAY ━━━", "", "", ""]]
    for idx, rec in enumerate(recs[:5], 1):
        rows.append([f"{idx}.", rec, "", ""])

    ws.update("A1", rows)

    section_bg_rows = [2, 10, 20, len(rows) - len(recs[:5]) - 2]
    sub_hdr_rows    = [3, 11]

    reqs = [
        fmt_req(sheet_id, 0, 1, 0, 4, bold=True, bg=C_ACCENT, fg=C_HEADER_FG, font_size=14),
        freeze_req(sheet_id, rows=1),
        autoresize_req(sheet_id, 0, 4),
    ]
    for sr in section_bg_rows:
        reqs.append(fmt_req(sheet_id, sr, sr+1, 0, 4, bold=True,
                            bg=C_HEADER_BG, fg=C_HEADER_FG, font_size=11))
    for sr in sub_hdr_rows:
        reqs.append(fmt_req(sheet_id, sr, sr+1, 0, 4, bold=True, bg=C_SECTION_BG))
    # Highlight change column (D = col 3) for Shopify table (rows 4-8)
    reqs += cond_pct(sheet_id, 4, 9, 3, 4)
    # Highlight change column for Meta table (rows 12-19)
    reqs += cond_pct(sheet_id, 12, 20, 3, 4)
    batch(sheets_svc, spreadsheet_id, reqs)

# ─────────────────────── Embedded charts ──────────────────────────────────────

def add_charts(sheets_svc, spreadsheet_id,
               shopify_tab_id: int, meta_tab_id: int, dashboard_tab_id: int,
               shopify_row_count: int):
    """Add or replace trend charts on the Dashboard tab."""
    try:
        ss = sheets_svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        existing_charts = ss.get("sheets", [{}])[0].get("charts", [])
        delete_reqs = [{"deleteEmbeddedObject": {"objectId": c["chartId"]}}
                       for s in ss["sheets"] if s["properties"]["sheetId"] == dashboard_tab_id
                       for c in s.get("charts", [])]
        if delete_reqs:
            batch(sheets_svc, spreadsheet_id, delete_reqs)
    except Exception:
        pass

    def chart_req(title, tab_id, col_idx, color, anchor_col, anchor_row):
        return {"addChart": {"chart": {
            "spec": {
                "title": title,
                "basicChart": {
                    "chartType": "LINE",
                    "legendPosition": "BOTTOM_LEGEND",
                    "axis": [
                        {"position": "BOTTOM_AXIS", "title": "Date"},
                        {"position": "LEFT_AXIS",   "title": title},
                    ],
                    "domains": [{"domain": {"sourceRange": {"sources": [{
                        "sheetId": tab_id,
                        "startRowIndex": 0, "endRowIndex": shopify_row_count,
                        "startColumnIndex": 0, "endColumnIndex": 1,
                    }]}}}],
                    "series": [{"series": {"sourceRange": {"sources": [{
                        "sheetId": tab_id,
                        "startRowIndex": 0, "endRowIndex": shopify_row_count,
                        "startColumnIndex": col_idx, "endColumnIndex": col_idx+1,
                    }]}}, "targetAxis": "LEFT_AXIS",
                    "colorStyle": {"rgbColor": color}}],
                    "headerCount": 1,
                },
            },
            "position": {
                "overlayPosition": {
                    "anchorCell": {
                        "sheetId": dashboard_tab_id,
                        "rowIndex": anchor_row, "columnIndex": anchor_col,
                    },
                    "offsetXPixels": 0, "offsetYPixels": 0,
                    "widthPixels": 440, "heightPixels": 220,
                }
            },
        }}}

    reqs = [
        chart_req("Gross Sales Trend (₹)", shopify_tab_id, 1,
                  {"red": 0.259, "green": 0.522, "blue": 0.957}, 5, 3),
        chart_req("Orders Trend", shopify_tab_id, 3,
                  {"red": 0.204, "green": 0.659, "blue": 0.325}, 5, 16),
        chart_req("Ad Spend Trend (₹)", meta_tab_id, 1,
                  {"red": 1, "green": 0.596, "blue": 0}, 10, 3),
        chart_req("ROAS Trend", meta_tab_id, 10,
                  {"red": 0.839, "green": 0.153, "blue": 0.157}, 10, 16),
    ]
    batch(sheets_svc, spreadsheet_id, reqs)

# ─────────────────────── Spreadsheet orchestration ────────────────────────────

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
        tab_map[tab] = (existing[tab] if tab in existing
                        else ss.add_worksheet(title=tab, rows=1000, cols=30))
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


def update_spreadsheet(date_str: str,
                        shop_today: dict, shop_7d: dict,
                        shop_prod_today: dict, shop_prod_7d: dict,
                        meta_campaigns: list, meta_adsets: list,
                        meta_account_7d: list,
                        recs: list, notes: list, unavailable: list,
                        shop_7d_avg: dict, meta_today: dict, meta_7d_avg: dict):
    gc, sheets_svc, gmail_svc = get_google_clients()
    ss      = get_or_create_spreadsheet(gc, SPREADSHEET_NAME)
    tab_map = ensure_tabs(ss)
    sids    = {ws.title: ws.id for ws in ss.worksheets()}

    build_shopify_daily_tab(
        tab_map["Shopify Daily Data"], sheets_svc, ss.id,
        sids["Shopify Daily Data"], shop_today, shop_7d, date_str,
    )
    build_meta_daily_tab(
        tab_map["Meta Ads Daily Data"], sheets_svc, ss.id,
        sids["Meta Ads Daily Data"], meta_campaigns, meta_account_7d, date_str,
    )
    build_product_tab(
        tab_map["Product Performance"], sheets_svc, ss.id,
        sids["Product Performance"], shop_prod_today, shop_prod_7d, date_str,
    )
    build_campaign_tab(
        tab_map["Campaign Performance"], sheets_svc, ss.id,
        sids["Campaign Performance"], meta_campaigns, meta_adsets,
    )
    build_recs_tab(
        tab_map["Recommendations & Notes"], sheets_svc, ss.id,
        sids["Recommendations & Notes"], recs, notes, date_str, unavailable,
    )
    build_dashboard_tab(
        tab_map["Dashboard"], sheets_svc, ss.id,
        sids["Dashboard"], shop_today, shop_7d_avg,
        meta_today, meta_7d_avg, recs, date_str,
    )

    # Charts — best-effort
    try:
        shopify_ws = tab_map["Shopify Daily Data"]
        row_count  = len(shopify_ws.col_values(1))
        add_charts(sheets_svc, ss.id,
                   sids["Shopify Daily Data"], sids["Meta Ads Daily Data"],
                   sids["Dashboard"], row_count)
    except Exception as e:
        print(f"  ℹ️  Charts skipped: {e}")

    return ss.url, gmail_svc

# ─────────────────────── Email ─────────────────────────────────────────────────

def send_or_draft_email(gmail_svc, sheet_url: str, date_str: str,
                         shop_today: dict, shop_7d_avg: dict,
                         meta_today: dict, recs: list, urgent: list):
    s    = shop_today.get("sales", {})
    gs   = sf(s.get("gross_sales", 0));  ag_gs = shop_7d_avg.get("gross_sales", 0)
    ns   = sf(s.get("net_sales",   0));  ag_ns = shop_7d_avg.get("net_sales",   0)
    od   = int(sf(s.get("orders",  0))); ag_od = shop_7d_avg.get("orders",      0)
    ao   = sf(s.get("average_order_value", 0))
    rt   = abs(sf(s.get("returns", 0)))
    sp   = meta_today.get("spend",     0)
    pu   = meta_today.get("purchases", 0)
    cpa  = meta_today.get("cpa",       0)
    ro   = meta_today.get("roas",      0)

    def inr(v): return f"₹{sf(v):,.0f}"
    def pct(curr, prior):
        c = pct_change(curr, prior)
        return ("—" if c is None else (f"+{c:.1f}%" if c >= 0 else f"{c:.1f}%"))

    urgent_block = (
        "\n🚨 URGENT ISSUES\n" + "\n".join(f"  • {i}" for i in urgent) + "\n\n"
        if urgent else ""
    )

    plain = f"""\
Hi Atul,

{urgent_block}Here is your daily performance summary for {date_str}.

━━━━━━━━━━━━━━━━━━━━━━━━━━━
📦 SHOPIFY PERFORMANCE
━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Gross Sales:     {inr(gs)}   ({pct(gs, ag_gs)} vs 7-day avg {inr(ag_gs)})
  Net Sales:       {inr(ns)}   ({pct(ns, ag_ns)} vs 7-day avg {inr(ag_ns)})
  Orders:          {od}        ({pct(od, ag_od)} vs 7-day avg {round(ag_od,1)})
  Avg Order Value: {inr(ao)}
  Returns:         {inr(rt)}

━━━━━━━━━━━━━━━━━━━━━━━━━━━
📣 META ADS PERFORMANCE
━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Total Spend:     {inr(sp)}
  Purchases:       {pu}
  CPA:             {inr(cpa)}
  Blended ROAS:    {ro:.2f}x

━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ RECOMMENDED ACTIONS FOR TODAY
━━━━━━━━━━━━━━━━━━━━━━━━━━━
{chr(10).join(f"  {i+1}. {r}" for i, r in enumerate(recs[:5]))}

━━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 Full Dashboard:
{sheet_url}
━━━━━━━━━━━━━━━━━━━━━━━━━━━

Best,
Dhirai Daily Bot
"""
    html = (
        plain
        .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        .replace("\n", "<br>")
        .replace("━", "─")
    )
    html = f"<html><body style='font-family:monospace;'>{html}</body></html>"

    msg = MIMEText(html, "html")
    msg["To"]      = GMAIL_TO
    msg["Subject"] = f"Daily Store & Ads Performance Sheet - {date_str}"
    if GMAIL_CC:
        msg["Cc"] = GMAIL_CC

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()

    try:
        gmail_svc.users().messages().send(userId="me", body={"raw": raw}).execute()
        print(f"  ✅ Email sent to {GMAIL_TO}")
    except HttpError as e:
        print(f"  ⚠️  Send failed ({e}), creating draft...")
        gmail_svc.users().drafts().create(userId="me",
                                          body={"message": {"raw": raw}}).execute()
        print("  ✅ Gmail draft created.")

# ─────────────────────── Main ──────────────────────────────────────────────────

def main():
    yesterday = yesterday_ist()
    since_7d, until_7d = dates_for_7day(yesterday)
    date_str  = yesterday.isoformat()
    unavail   = []

    print(f"📅  Running dashboard for {date_str} (Asia/Kolkata)\n")

    # ── Shopify ──────────────────────────────────────────────────────────────
    print("🛒  Fetching Shopify data…")
    try:
        shop_today = fetch_shopify_day(date_str)
    except Exception as e:
        print(f"    ⚠️  Daily fetch failed: {e}")
        shop_today = {"sales": {}, "sessions": {}, "top_products": {"rows": [], "columns": []}}
        unavail.append(f"Shopify daily data: {e}")

    try:
        shop_7d = fetch_shopify_7day(since_7d, until_7d)
    except Exception as e:
        print(f"    ⚠️  7-day fetch failed: {e}")
        shop_7d = {"rows": [], "columns": []}
        unavail.append(f"Shopify 7-day data: {e}")

    try:
        shop_prod_today = shop_today.get("top_products", {"rows": [], "columns": []})
        shop_prod_7d    = fetch_shopify_7day_products(since_7d, until_7d)
    except Exception as e:
        shop_prod_7d = {"rows": [], "columns": []}
        unavail.append(f"Shopify 7-day product data: {e}")

    # ── Meta Ads ─────────────────────────────────────────────────────────────
    print("📣  Fetching Meta Ads data…")
    try:
        meta_campaigns = fetch_meta_campaigns(date_str)
    except Exception as e:
        print(f"    ⚠️  Campaign fetch failed: {e}")
        meta_campaigns = []
        unavail.append(f"Meta campaign data: {e}")

    try:
        meta_adsets = fetch_meta_adsets(date_str)
    except Exception as e:
        print(f"    ⚠️  Ad-set fetch failed: {e}")
        meta_adsets = []
        unavail.append(f"Meta ad-set data: {e}")

    try:
        meta_account_7d = fetch_meta_account_daily(since_7d, until_7d)
    except Exception as e:
        print(f"    ⚠️  7-day account data fetch failed: {e}")
        meta_account_7d = []
        unavail.append(f"Meta 7-day account data: {e}")

    anomalies = fetch_meta_anomalies()   # best-effort; silent on failure

    # ── Compute averages ──────────────────────────────────────────────────────
    cols   = shop_7d.get("columns", [])
    ci     = {c: i for i, c in enumerate(cols)}
    rows7d = [r for r in shop_7d.get("rows", []) if r and str(r[ci.get("day",0)]) != date_str]
    shop_7d_avg = col_avgs(rows7d, {
        "gross_sales":         ci.get("gross_sales",         1),
        "net_sales":           ci.get("net_sales",           2),
        "orders":              ci.get("orders",              3),
        "average_order_value": ci.get("average_order_value", 4),
    })

    meta_today  = aggregate_meta(meta_campaigns)
    meta_7d_avg = aggregate_meta(meta_account_7d) if meta_account_7d else {}

    # ── Recommendations ───────────────────────────────────────────────────────
    print("💡  Generating recommendations…")
    recs, notes = generate_recommendations(
        shop_today, shop_7d_avg,
        meta_campaigns, meta_adsets,
        meta_today, meta_7d_avg,
        anomalies,
    )
    urgent = [r for r in recs if r.startswith("🚨")]

    # ── Google Sheet ──────────────────────────────────────────────────────────
    print("📊  Updating Google Sheet…")
    try:
        sheet_url, gmail_svc = update_spreadsheet(
            date_str,
            shop_today, shop_7d,
            shop_prod_today, shop_prod_7d,
            meta_campaigns, meta_adsets, meta_account_7d,
            recs, notes, unavail,
            shop_7d_avg, meta_today, meta_7d_avg,
        )
        print(f"    ✅ Sheet updated: {sheet_url}")
    except Exception as e:
        print(f"    ⚠️  Sheet update failed: {e}")
        sheet_url = "https://docs.google.com/spreadsheets (configure GOOGLE_SA_CREDENTIALS)"
        try:
            _, _, gmail_svc = get_google_clients()
        except Exception as auth_err:
            print(f"    ⚠️  Gmail auth also failed: {auth_err}")
            return

    # ── Email ─────────────────────────────────────────────────────────────────
    print("✉️   Sending email / creating draft…")
    try:
        send_or_draft_email(
            gmail_svc, sheet_url, date_str,
            shop_today, shop_7d_avg, meta_today, recs, urgent,
        )
    except Exception as e:
        print(f"    ⚠️  Email step failed: {e}")

    print("\n✅  Done.")


if __name__ == "__main__":
    main()
