# gazette-tracker
Gazette Insolvency Tracker

## Cayman Islands scraper

`cayman_scraper.py` auto-fetches the latest Cayman Islands Government
Gazette PDF and uses the Claude API (vision) to extract liquidation/
winding-up notices, since some issues are scans without a text layer.
It writes into the same `insolvencies` table as everything else
(country `KY`), runs weekly via `scheduler.py`, and is also included in
the "Refresh All" button (`/api/refresh`).

Requires an `ANTHROPIC_API_KEY` environment variable — set it in Railway
the same way `GMAIL_APP_PASSWORD` is set.
