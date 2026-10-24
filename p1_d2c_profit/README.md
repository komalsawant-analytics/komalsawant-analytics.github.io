# P1 · D2C Unified Profit Dashboard

**Sample project · simulated data · fictional brand "Mitti & Moss" (skincare), Apr–Sep 2026**

## The problem
The brand sells on Shopify, Amazon and a beauty marketplace ("BeautyMart", fictional) and runs Meta ads. Every source reports revenue differently, and none of them shares a product code, so nobody knew which products made money.

## Data in (`raw/`)
| File | What was wrong with it |
|---|---|
| `shopify_orders_export.csv` | Test and voided orders mixed in, SKUs in mixed case with trailing spaces, shipping only on the first line of each order |
| `amazon_settlement_Apr-Sep2026.xlsx` | Different SKU codes, DD/MM/YYYY text dates, storage-fee rows mixed with orders, returns in a separate sheet |
| `beautymart_sales_report.csv` | No SKU column — 41 spellings of 20 product names, duplicate rows, amounts and commission stored as text |
| `meta_ads_daily.csv` | Brand and retargeting campaigns not tied to one product |
| `product_master.xlsx` | Title rows above the header, one product missing its cost |

## What the pipeline does (`scripts/clean_and_model.py`)
1. Cleans each source and logs every fix (`clean/data_quality_log.csv`).
2. Maps every source's product key to the master SKU (pattern rules for Amazon, fuzzy name matching for BeautyMart — 41/41 correct).
3. Applies agreed business rules (gateway fee, courier cost, reverse pickup, which returns can be resold).
4. Allocates shared ad spend to products by their website revenue that month (reconciles to the rupee).
5. Writes a star schema: `fact_sales`, `fact_ads`, `dim_product`, `dim_date`, `sku_crosswalk` (CSV + `profit_model.xlsx`).

Run: `python scripts/clean_and_model.py`

## Key findings
- **Bestseller ≠ best earner.** The Vitamin C serum is 18.6% of revenue, but the Mineral Sunscreen makes more profit (₹5.48L vs ₹4.22L) on 55% of the sales. The serum averages a 22% website discount and heavy Meta spend.
- **Lip tints return 27% of units on BeautyMart** vs 5% for the rest of the range there; three lip tints keep a 4.5% margin after ads.
- **The marketplace's June sale bought volume, not profit:** margin before ads fell from 37% to 26% during sale week.

## Dashboard
Interactive web version on the portfolio page. Power BI version: load the `clean/` CSVs, relate the facts to `dim_product[sku]` and `dim_date[date]`, and build the measures listed in the playbook.

## Limitations
Amazon and marketplace ad spend were not in the files, so those channels' margins are before their own ads. Fixed overheads are not included.
