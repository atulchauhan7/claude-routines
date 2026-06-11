# Dhirai Daily Dashboard

Daily morning automation for **Dhirai** (`dhirai.in`) — an Indian ethnic wear Shopify store.

Runs at **07:00 IST** every day. Fetches Shopify sales + Meta Ads data, writes a
6-tab Google Sheet, and sends/drafts an HTML performance report by email.

---

## What it produces

### Google Sheet — "Daily Store & Ads Performance Sheet"

| Tab | Contents |
|---|---|
| **Dashboard** | Executive summary — key metrics vs 7-day avg with ▲/▼ indicators, wins, issues, top recommendations |
| **Shopify Daily Data** | Append-only — one row per date: gross/net sales, orders, AOV, discounts, returns, taxes, sessions, cart-add rate, checkout rate, conversion rate |
| **Meta Ads Daily Data** | Append-only — one row per date: spend, impressions, reach, clicks, CTR, CPC, purchases, CPA, ROAS |
| **Product Performance** | Top 10 products by revenue yesterday (with 7-day comparison) |
| **Campaign Performance** | Active campaigns yesterday + spend vs 7-day avg % + ROAS vs 7-day avg % (colour-coded) |
| **Recommendations & Notes** | Auto-generated analysis + notes on any unavailable metrics |

All tabs use: frozen headers, bold headers, auto-filters, INR currency formatting,
percentage formatting, and green/red/yellow conditional formatting on ROAS and change columns.

Duplicate-date guard: the script checks column A before appending — the same date is
never written twice.

### Daily email

- **To:** atul012001@gmail.com
- **Subject:** `Daily Store & Ads Performance Sheet - YYYY-MM-DD`
- **Body (HTML):** urgent issues first → Shopify table → Meta Ads table → top 5 recommended actions → Google Sheet link
- Falls back to creating a **Gmail draft** if the send call fails

---

## Prerequisites

- Python 3.10+
- Shopify store with a **Custom App** (Admin API)
- Meta Business account with **Marketing API** access
- Google Cloud project with Sheets, Drive, and Gmail APIs enabled
- A **Google service account** JSON (for Sheets read/write)
- A **Google OAuth 2.0 Desktop App** credentials JSON (for Gmail send)

---

## Setup

### Step 1 — Install dependencies

```bash
pip install -r requirements.txt
```

### Step 2 — Shopify Custom App

1. Shopify Admin → **Settings → Apps and sales channels → Develop apps**
2. Create app → **Configure Admin API scopes**
3. Enable: `read_orders`, `read_analytics`, `read_products`
4. **Install app** → copy the **Admin API access token** (`shpat_…`)

### Step 3 — Meta Ads access token

**Recommended (no expiry): System User token**
1. Meta Business Suite → **Settings → Users → System Users**
2. Create a system user, assign `Analyst` role on ad account `979830497515712`
3. **Generate token** → select `ads_read` permission → copy token

**Alternative: User token (expires in ~60 days)**
1. [Graph API Explorer](https://developers.facebook.com/tools/explorer/)
2. Select your app → generate token with `ads_read`, `read_insights`
3. Exchange for a long-lived token via the Token Debugger

### Step 4 — Google service account (for Google Sheets)

1. [Google Cloud Console](https://console.cloud.google.com/) → select or create a project
2. **APIs & Services → Library** → enable:
   - Google Sheets API
   - Google Drive API
3. **IAM & Admin → Service Accounts** → Create service account
4. **Keys** tab → **Add Key → Create new key → JSON** → download
5. Save the file as `service_account.json` (or any path — set via env var)

The script automatically creates and shares the Google Sheet with the addresses in
`GOOGLE_SHEET_EMAIL` and `GMAIL_TO` when it runs for the first time.

### Step 5 — Google OAuth 2.0 credentials (for Gmail)

1. Same Google Cloud project → **APIs & Services → Library** → enable **Gmail API**
2. **APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client ID**
3. Application type: **Desktop app**
4. Download JSON → save as `gmail_credentials.json`
5. **OAuth consent screen → Test users** → add your Gmail address

**First-run authorisation (do this once, interactively):**

```bash
set -a && source .env && set +a
python daily_dashboard.py
```

A browser window opens asking you to authorise Gmail access. After approving, a
`gmail_token.pickle` file is saved and reused on every subsequent run — no more
browser prompts needed.

### Step 6 — Environment variables

Create a `.env` file:

```bash
# Shopify
SHOPIFY_STORE_DOMAIN=36dhns-ed.myshopify.com
SHOPIFY_ACCESS_TOKEN=shpat_xxxxxxxxxxxxxxxxxxxx

# Meta Ads
META_ACCESS_TOKEN=EAAxxxxxxxxxxxxx
META_AD_ACCOUNT_ID=979830497515712

# Google Sheets (service account)
GOOGLE_SA_CREDENTIALS=/absolute/path/to/service_account.json

# Gmail (OAuth)
GMAIL_CREDENTIALS_JSON=/absolute/path/to/gmail_credentials.json
GMAIL_TOKEN_PICKLE=/absolute/path/to/gmail_token.pickle

# Email recipients
GMAIL_TO=atul012001@gmail.com
GOOGLE_SHEET_EMAIL=atul.chauhan.95185@gmail.com

# Optional
GMAIL_CC=
SPREADSHEET_NAME=Daily Store & Ads Performance Sheet
```

---

## Running manually

```bash
# Load env vars from .env and run
set -a && source .env && set +a
python daily_dashboard.py
```

Expected log output:

```
2026-06-11 07:00:01 [INFO] ============================================================
2026-06-11 07:00:01 [INFO] Dhirai Daily Dashboard — Starting
2026-06-11 07:00:02 [INFO] Report date: 2026-06-10 (Asia/Kolkata)
2026-06-11 07:00:02 [INFO] 7-day comparison window: 2026-06-03 to 2026-06-09
2026-06-11 07:00:04 [INFO] --- Shopify: fetching daily data ---
2026-06-11 07:00:08 [INFO] Active Meta campaigns: 4
2026-06-11 07:00:12 [INFO] --- Google Sheets: updating ---
2026-06-11 07:01:10 [INFO] Sheet updated: https://docs.google.com/spreadsheets/d/…
2026-06-11 07:01:14 [INFO] Email sent to atul012001@gmail.com
2026-06-11 07:01:14 [INFO] Dhirai Daily Dashboard — Complete
```

---

## Scheduling via cron (daily at 07:00 IST)

Create a wrapper script so cron can load env vars cleanly:

```bash
# /home/ubuntu/dhirai/run_dashboard.sh
#!/bin/bash
set -a
source /home/ubuntu/dhirai/.env
set +a
cd /home/ubuntu/dhirai
/home/ubuntu/dhirai/venv/bin/python daily_dashboard.py >> /home/ubuntu/dhirai/logs/dashboard.log 2>&1
```

```bash
chmod +x /home/ubuntu/dhirai/run_dashboard.sh
mkdir -p /home/ubuntu/dhirai/logs
```

Add to crontab (`crontab -e`):

```cron
0 7 * * * /home/ubuntu/dhirai/run_dashboard.sh
```

> **Note:** Cron runs in UTC. IST = UTC+5:30, so 07:00 IST = 01:30 UTC.
> If your server is in UTC, use: `30 1 * * *`

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `SpreadsheetNotFound` | Service account lacks Drive access | Enable Google Drive API; check `GOOGLE_SA_CREDENTIALS` path |
| `ShopifyQL query failed (403 / 422)` | Analytics API requires Shopify paid plan | Web metrics default to 0; noted in Recommendations tab |
| `Meta API error 190` | Access token expired | Rotate token or switch to system user (no expiry) |
| `Token has been expired or revoked` | Gmail OAuth token stale | Delete `gmail_token.pickle`, re-run interactively |
| Gmail `send` returns 403 | OAuth scope missing `gmail.send` | Delete token pickle, re-authorise with correct scopes |
| `No active Meta campaigns found` | All campaigns paused | Expected — noted in Recommendations tab |
| Duplicate rows in sheet | Column A contains non-ISO date strings | Ensure column A contains plain `YYYY-MM-DD` strings, not formatted dates |
| Cron runs but no email | Missing env vars in cron environment | Use `set -a && source .env` inside the wrapper script |

---

## Key design decisions

- **Append-only history:** The Shopify Daily Data and Meta Ads Daily Data tabs accumulate
  one row per day. Existing rows are never overwritten or deleted.
- **Duplicate guard:** Before appending, the script reads column A and skips if the date
  exists — safe to re-run multiple times on the same day.
- **Graceful degradation:** Every API call is wrapped in a try/except. Failures log a
  warning and populate the Recommendations & Notes tab — the rest of the report continues.
- **Separate auth per service:** Google Sheets uses a service account (no browser needed);
  Gmail uses OAuth 2.0 (browser required once, then cached).
- **ROAS conditional formatting:** < 1.5x = red, 1.5–3x = yellow, > 3x = green.
- **Active campaigns only:** Meta data filters to `effective_status = ACTIVE` and
  `spend > 0` — paused / zero-spend campaigns are excluded from all tabs.

---

## Architecture

```
daily_dashboard.py
├── Date helpers          — yesterday in Asia/Kolkata timezone
├── Shopify module
│   ├── shopify_analytics()             — ShopifyQL via Admin GraphQL API
│   ├── fetch_shopify_daily()           — sales + sessions + top products
│   └── fetch_shopify_7day()            — 7-day timeseries for comparison
├── Meta Ads module
│   ├── fetch_meta_active_campaign_ids() — ACTIVE campaigns only
│   ├── fetch_meta_campaigns()           — yesterday campaign insights
│   ├── fetch_meta_adsets()              — yesterday ad set insights
│   └── fetch_meta_campaigns_7day()      — 7-day comparison
├── Recommendations engine
│   └── generate_recommendations()       — cross-analyses Shopify + Meta
├── Google Sheets (service account)
│   ├── build_dashboard_tab()
│   ├── build_shopify_daily_tab()        — append-only
│   ├── build_meta_daily_tab()           — append-only
│   ├── build_product_tab()
│   ├── build_campaign_tab()             — with 7-day avg % change + ROAS colours
│   └── build_recommendations_tab()
└── Gmail (OAuth 2.0)
    ├── get_gmail_service()              — token caching via pickle file
    ├── build_html_email()              — rich HTML with colour-coded tables
    └── send_or_draft_email()           — send first, create draft on failure
```
