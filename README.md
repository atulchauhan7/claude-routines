# Daily Store & Ads Performance Dashboard

Automated morning routine for **Dhirai** (dhirai.in) that:

1. Fetches yesterday's full-day Shopify + Meta Ads data (Asia/Kolkata timezone)
2. Creates or updates a Google Sheet named **"Daily Store & Ads Performance Sheet"**
3. Sends an email (or creates a Gmail draft) with the daily summary and recommendations

---

## Tabs in the Google Sheet

| Tab | Contents |
|-----|----------|
| **Dashboard** | Executive summary: key metrics vs 7-day avg, wins, issues, recommendations |
| **Shopify Daily Data** | One row per day — gross/net sales, orders, AOV, sessions, conversion rate (append-only, no duplicates) |
| **Meta Ads Daily Data** | One row per day — spend, impressions, clicks, CTR, CPC, purchases, CPA, ROAS |
| **Product Performance** | Top 10 products by revenue for yesterday vs 7-day totals |
| **Campaign Performance** | Active campaigns and ad sets for yesterday with ROAS conditional formatting |
| **Recommendations & Notes** | Auto-generated recommendations + unavailable-metric notes |

---

## Setup

### 1. Google Service Account (for Sheets + Gmail send)

1. Go to [Google Cloud Console](https://console.cloud.google.com/) → **IAM & Admin** → **Service Accounts**
2. Create a service account and download the JSON key
3. Enable **Google Sheets API** and **Gmail API** for the project
4. Share the target Google Sheet (or let the script create it) with the service account email

> **Note:** Gmail send via service account requires **G Suite domain-wide delegation** if sending as a real user. Alternatively, set `GMAIL_TO` and the script will create a Gmail draft instead of sending.

### 2. Shopify Private App Token

1. In your Shopify admin go to **Settings → Apps → Develop apps**
2. Create a private app with **read_analytics**, **read_orders**, **read_products** scopes
3. Copy the Admin API access token

### 3. Meta Access Token

1. Go to [Meta for Developers → Graph API Explorer](https://developers.facebook.com/tools/explorer/)
2. Generate a long-lived token with `ads_read`, `ads_management`, `read_insights` permissions
3. Use the ad account ID `979830497515712` (Dhirai)

---

## Environment Variables

```bash
export SHOPIFY_STORE_DOMAIN="36dhns-ed.myshopify.com"
export SHOPIFY_ACCESS_TOKEN="shpat_xxxxxxxxxxxxxxxxxxxx"
export META_ACCESS_TOKEN="EAAxxxxxxxxxxxxxxxxxxxx"
export META_AD_ACCOUNT_ID="979830497515712"
export GOOGLE_SA_CREDENTIALS="/path/to/service_account.json"
export GMAIL_TO="atul012001@gmail.com"
export GMAIL_CC=""                          # leave empty if no CC
export SPREADSHEET_NAME="Daily Store & Ads Performance Sheet"
```

Or create a `.env` file and load with `python-dotenv`:

```bash
pip install python-dotenv
```

Add to the top of `daily_dashboard.py`:
```python
from dotenv import load_dotenv
load_dotenv()
```

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Running Manually

```bash
python daily_dashboard.py
```

Output will show progress steps and print the Google Sheet URL on success.

---

## Scheduling (Daily Cron at 7 AM IST)

```cron
0 7 * * * cd /path/to/claude-routines && python daily_dashboard.py >> /var/log/dhirai_dashboard.log 2>&1
```

Or with a virtual environment:

```cron
0 7 * * * /path/to/venv/bin/python /path/to/claude-routines/daily_dashboard.py >> /var/log/dhirai_dashboard.log 2>&1
```

---

## What Gets Fetched

### Shopify (via ShopifyQL)
- Gross sales, net sales, orders, AOV, total sales, discounts, returns, taxes
- Sessions, cart additions, checkout completions, conversion rate
- Top 10 products by gross sales

### Meta Ads (account: 979830497515712 — Dhirai)
- Campaign-level: spend, impressions, reach, clicks, CTR, CPC, CPM, purchases, CPA, ROAS
- Ad set-level: same metrics + learning phase status
- Both yesterday and 7-day window for comparison

---

## Email

- **To:** atul012001@gmail.com  
- **CC:** None (add `GMAIL_CC` env var if needed)  
- **Subject:** `Daily Store & Ads Performance Sheet - YYYY-MM-DD`  
- **Body:** Urgent issues → Shopify summary → Meta summary → 5 recommended actions → Sheet link

If direct email send fails (e.g., service account lacks delegation), the script automatically falls back to creating a **Gmail draft** with the same content.

---

## Key Design Decisions

- **Append-only:** Historical rows are never overwritten. The script checks existing dates before inserting.
- **Duplicate guard:** Each date is inserted only once per tab.
- **Graceful degradation:** If any data source is unavailable, the script continues and logs the issue in the Recommendations & Notes tab.
- **ROAS thresholds (conditional formatting):** < 1.5x = red, 1.5–3x = yellow, > 3x = green.

---

## Triggered by Claude Code Routine

This script is managed by [Claude Code](https://code.claude.com) and stored in the `claude-routines` repository. To update the routine logic, edit `daily_dashboard.py` and push to the `claude/dreamy-archimedes-9q9wyo` branch.
