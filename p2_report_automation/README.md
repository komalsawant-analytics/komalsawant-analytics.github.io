# P2 · Weekly Store Report on Autopilot

**Sample project · simulated data · fictional retailer "Kora Home" (4 stores)**

## The problem
Every Monday each store sends its POS export in its own format, and someone spends half a day pasting them together.

| Store file | Format quirks |
|---|---|
| `Andheri_POS.xlsx` | Clean-ish: item codes, real dates |
| `bandra_sales.csv` | Lower-case item codes |
| `THANE store report.xlsx` | 3 banner rows above the header, item **names** only (upper case), amounts as "1,299" text |
| `pune-export.csv` | Semicolon-separated, DD.MM.YYYY dates |

## What the script does
`python scripts/weekly_report.py --week 2026-10-12`

1. Reads each store's file with a per-store mapping (adding a store = adding one config block).
2. Cleans, unifies and compares with the previous week's folder.
3. Writes `output/Kora_Home_Weekly_Report_<week>.xlsx`:
   - **Summary** — KPIs, store and category week-on-week change with red/green flags, two charts. Every number is a live `SUMIFS` on the Data sheet; change the week cell and it recalculates.
   - **Top Items**, **Reorder Alerts** (days of stock cover, editable threshold), **Data Log**.
4. Writes `output/email_summary_<week>.html`, ready to send.

Schedule with Windows Task Scheduler or cron; the same logic ports to Google Apps Script for Drive-based stores.
