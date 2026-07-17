# Daily Store & Ads Performance Sheet

Automation for the daily Shopify (Dhirai, `dhirai.in`) + Meta Ads
(account `979830497515712`, business "Dhirai") performance dashboard.

## Current approach (as of 2026-07-17)

This session's environment does not have raw Shopify/Meta API tokens or a
Google service-account file — only MCP tool access (Shopify, Meta Ads,
Gmail) scoped to this Claude Code session, and **no Google Sheets
connector**. `daily_dashboard.py` (the earlier script in this repo) assumes
direct API credentials and `gspread`/Google service-account auth that are
not available here, so it cannot run as-is in this kind of session.

Instead, each day's run:

1. Calls the connected **Shopify** MCP tools (`run-analytics-query` /
   ShopifyQL, `list-orders`, inventory queries) for yesterday's full
   Asia/Kolkata calendar day and the trailing 7 days.
2. Calls the connected **Meta Ads** MCP tools
   (`ads_get_ad_entities`, `ads_insights_anomaly_signal`, etc.) for the
   same two windows, account `979830497515712`.
3. Pastes/updates the `RAW DATA` block at the top of
   `scripts/build_dashboard_data.py` with the freshly fetched values.
4. Runs `scripts/build_dashboard_data.py` then
   `scripts/build_dashboard_dashboard_tab.py` to (re)build
   `Daily Store & Ads Performance Sheet.xlsx` in the repo root. Both
   scripts are idempotent — re-running against an existing file skips any
   date already present in `Shopify Daily Data` / `Meta Ads Daily Data`
   and only appends new dates, so re-runs never duplicate a day.
5. Commits & pushes the updated `.xlsx` so tomorrow's run can load and
   append to it (this repo is the persistence layer in place of a live
   Google Sheet).
6. Creates a Gmail draft (no direct-send tool is exposed by the connected
   Gmail MCP server in this session) to `atul012001@gmail.com` with the
   daily summary and a note about the workbook.

## Known environment limitations (see the "Recommendations & Notes" tab
for the full list)

- No Google Sheets connector — delivered as `.xlsx` instead of a live
  Google Sheet.
- No Gmail send tool — delivered as a Gmail draft instead of a sent email.
- LibreOffice headless recalculation (`scripts/office` / xlsx skill
  `recalc.py`) hangs indefinitely in this sandbox even on a trivial
  one-formula file — formulas in the workbook are written correctly but
  ship without pre-computed cached values. They compute normally the
  moment the file is opened in Excel, Google Sheets, or a working desktop
  LibreOffice.
- Some Shopify metrics (failed payments, abandoned checkouts, order
  cancellations distinct from refunds, units sold per product) and Meta
  Ads metrics (a raw purchase/conversions count) are not exposed by the
  connected tools in this session — see the workbook's "Recommendations &
  Notes" tab for what's estimated/proxied and how.
