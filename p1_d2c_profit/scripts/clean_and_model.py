"""
D2C Unified Profit Model — cleaning & modelling pipeline
=========================================================
Sample project · simulated data for a fictional brand ("Mitti & Moss").

Turns five messy exports into one clean star schema that answers:
"Which products and channels actually make money after discounts, platform
fees, returns, logistics, product cost and ad spend?"

Inputs  (raw/):  shopify_orders_export.csv, amazon_settlement_Apr-Sep2026.xlsx,
                 beautymart_sales_report.csv, meta_ads_daily.csv, product_master.xlsx
Outputs (clean/): fact_sales.csv, fact_ads.csv, dim_product.csv, dim_date.csv,
                  sku_crosswalk.csv, data_quality_log.csv, profit_model.xlsx

Run:  python clean_and_model.py
"""
from pathlib import Path
import re, difflib
import numpy as np
import pandas as pd

BASE = Path(__file__).resolve().parent.parent
RAW, CLEAN = BASE / "raw", BASE / "clean"
CLEAN.mkdir(exist_ok=True)

# ---------------------------------------------------------------- assumptions
# Business rules agreed with the (hypothetical) client. Change here, re-run.
ASSUME = {
    "shopify_gateway_fee_pct": 0.02,     # payment gateway on net sales
    "shopify_courier_per_order": 70.0,   # brand-paid courier cost per order
    "reverse_logistics_per_unit": 80.0,  # pickup cost for a customer return (Shopify, BeautyMart)
    "beautymart_commission_fallback": 0.25,
    # Opened skincare cannot be resold -> post-delivery returns lose the product cost.
    # RTO (returned before delivery) and Amazon 'SELLABLE' returns go back to stock.
}

LOG = []  # data-quality log: what was wrong, how many rows, what we did
def log(source, issue, rows, action):
    LOG.append({"source": source, "issue": issue, "rows_affected": int(rows), "action": action})

def to_num(s):
    """'1,299.00' / '25%' / '' -> float"""
    return pd.to_numeric(pd.Series(s).astype(str).str.replace(r"[,%₹\s]", "", regex=True)
                         .replace({"": np.nan, "nan": np.nan}), errors="coerce").values

# ============================================================ 1. PRODUCT MASTER
pm = pd.read_excel(RAW / "product_master.xlsx", header=2)          # title row above header
log("product_master", "Title rows above the header", 2, "Skipped first 2 rows on read")
pm.columns = ["sku", "product_name", "category", "mrp", "unit_cogs"]
pm["sku"] = pm["sku"].str.strip().str.upper()
missing_cogs = pm["unit_cogs"].isna()
if missing_cogs.any():
    ratio = pm["unit_cogs"] / pm["mrp"]
    cat_ratio = ratio.groupby(pm["category"]).transform("mean").fillna(ratio.mean())  # category avg, else overall avg
    pm.loc[missing_cogs, "unit_cogs"] = (pm["mrp"] * cat_ratio).round(0)[missing_cogs]
    pm["cogs_estimated"] = missing_cogs
    log("product_master", "Missing landed cost", missing_cogs.sum(),
        "Estimated from average cost-to-MRP ratio (category, else overall); flagged cogs_estimated=True")
dim_product = pm.copy()
SKUS = set(pm["sku"])

# ============================================================ helper: SKU matching
def norm_name(s):
    s = str(s).lower()
    s = re.sub(r"mitti\s*(&|and)?\s*moss", " ", s)
    s = s.replace("moisturizer", "moisturiser").replace("facewash", "face wash").replace("vit-c", "vitamin c")
    s = re.sub(r"(\d+)\s*(ml|gm|g)\b", lambda m: m.group(1) + ("ml" if m.group(2) == "ml" else "g"), s)
    s = re.sub(r"spf\s*(\d+)", r"spf\1", s)
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return " ".join(sorted(set(s.split())))

master_norm = {sku: norm_name(n) for sku, n in zip(pm["sku"], pm["product_name"])}
crosswalk = []

def match_name(name):
    n = norm_name(name)
    best, score = None, 0.0
    for sku, mn in master_norm.items():
        r = difflib.SequenceMatcher(None, n, mn).ratio()
        a, b = set(n.split()), set(mn.split())
        j = len(a & b) / len(a | b)                       # token overlap
        s = 0.5 * r + 0.5 * j
        if s > score: best, score = sku, s
    return best, round(score, 3)

# ============================================================ 2. SHOPIFY
sh = pd.read_csv(RAW / "shopify_orders_export.csv", dtype=str)
n0 = len(sh)
test = sh["Email"].str.contains("test", case=False, na=False) | sh["Email"].str.contains("mittiandmoss", na=False)
log("shopify", "Internal test orders", test.sum(), "Removed")
voided = sh["Financial Status"].eq("voided")
log("shopify", "Voided (cancelled before payment) orders", voided.sum(), "Removed")
sh = sh[~test & ~voided].copy()

raw_sku = sh["Lineitem sku"]
sh["sku"] = raw_sku.str.strip().str.upper()
log("shopify", "SKU with lowercase / trailing spaces", (raw_sku != sh["sku"]).sum(), "Trimmed and upper-cased")

sh["order_date"] = pd.to_datetime(sh["Created at"].str[:19]).dt.date
log("shopify", "Timestamp stored as text with timezone suffix", len(sh), "Parsed to IST date")

for c in ["Lineitem quantity", "Lineitem price", "Lineitem discount", "Refunded Amount"]:
    sh[c] = to_num(sh[c])
sh["Shipping"] = to_num(sh["Shipping"])
log("shopify", "Order-level fields (shipping) only on first line of multi-line orders",
    sh["Shipping"].isna().sum(), "Counted once per order; courier cost allocated per order, not per line")

lines_per_order = sh.groupby("Name")["sku"].transform("size")
sh_f = pd.DataFrame({
    "order_date": sh["order_date"], "channel": "Shopify", "order_id": sh["Name"], "sku": sh["sku"],
    "units": sh["Lineitem quantity"],
    "gross_sales": sh["Lineitem price"] * sh["Lineitem quantity"],
    "discount": sh["Lineitem discount"],
})
sh_f["net_sales"] = sh_f["gross_sales"] - sh_f["discount"]
sh_f["platform_fees"] = sh_f["net_sales"] * ASSUME["shopify_gateway_fee_pct"]
sh_f["logistics_cost"] = ASSUME["shopify_courier_per_order"] / lines_per_order
refunded = sh["Financial Status"].eq("refunded")
sh_f["returned_units"] = np.where(refunded, sh_f["units"], 0)
sh_f["return_value"] = np.where(refunded, sh["Refunded Amount"], 0.0)
sh_f["return_type"] = np.where(refunded, "Customer return", "")
sh_f["stock_recovered"] = False
sh_f["customer_shipping_paid"] = sh["Shipping"].fillna(0)

# ============================================================ 3. AMAZON
amx = pd.read_excel(RAW / "amazon_settlement_Apr-Sep2026.xlsx", sheet_name=None)
am, ar = amx["Transactions"], amx["Returns"]
svc = am["type"].ne("Order")
storage_fees = am.loc[svc].copy()
log("amazon", "Non-order rows (FBA storage fees) mixed with orders", svc.sum(),
    "Separated; storage fees kept as channel overhead")
am = am[~svc].copy()
am["order_date"] = pd.to_datetime(am["date/time"], format="%d/%m/%Y %H:%M").dt.date
log("amazon", "Dates in DD/MM/YYYY text (Excel reads some as MM/DD)", len(am), "Parsed with explicit day-first format")

def amz_to_sku(code):
    core = re.sub(r"^MMSKIN-|-(IN|IND)$", "", str(code).strip().upper())
    for sku in SKUS:
        if sku.replace("MM-", "").replace("-", "") == core:
            return sku
    return None
amz_map = {c: amz_to_sku(c) for c in am["sku"].dropna().unique()}
for c, s in amz_map.items():
    crosswalk.append({"source": "Amazon", "source_key": c, "sku": s, "method": "code pattern", "score": 1.0 if s else 0})
log("amazon", "Seller SKU codes differ from master SKUs (two suffix styles)", len(amz_map),
    f"Mapped {sum(v is not None for v in amz_map.values())}/{len(amz_map)} codes by pattern")
am["sku_m"] = am["sku"].map(amz_map)

ar["return_date"] = pd.to_datetime(ar["return-date"], format="%d-%m-%Y").dt.date
ar_g = ar.groupby("order-id").agg(ret_units=("quantity", "sum"), refund=("refund-amount", "sum"),
                                 disposition=("disposition", "first")).reset_index()
am = am.merge(ar_g, left_on="order id", right_on="order-id", how="left")
log("amazon", "Returns stored in a separate sheet with a different date format", len(ar),
    "Joined to orders on order id")

am_f = pd.DataFrame({
    "order_date": am["order_date"], "channel": "Amazon", "order_id": am["order id"], "sku": am["sku_m"],
    "units": am["quantity"].astype(float), "gross_sales": am["product sales"],
    "discount": -am["promotional rebates"],
})
am_f["net_sales"] = am_f["gross_sales"] - am_f["discount"]
am_f["platform_fees"] = -am["selling fees"]
am_f["logistics_cost"] = -am["fba fees"]
am_f["returned_units"] = am["ret_units"].fillna(0)
am_f["return_value"] = am["refund"].fillna(0)
am_f["return_type"] = np.where(am["ret_units"].notna(), "Customer return", "")
am_f["stock_recovered"] = am["disposition"].eq("SELLABLE")
am_f["customer_shipping_paid"] = 0.0

# ============================================================ 4. BEAUTYMART
bm = pd.read_csv(RAW / "beautymart_sales_report.csv", dtype=str)
d = bm.duplicated()
log("beautymart", "Exact duplicate rows (report exported twice for overlapping dates)", d.sum(), "Removed")
bm = bm[~d].copy()
bm["order_date"] = pd.to_datetime(bm["Order Date"], format="%d-%b-%Y").dt.date
for c in ["MRP Value", "Discount", "Selling Price", "Logistics Fee"]:
    bm[c] = to_num(bm[c])
log("beautymart", "Amounts stored as text with thousand separators", len(bm), "Converted to numbers")
bm["comm"] = to_num(bm["Commission %"]) / 100
bm["comm"] = bm["comm"].fillna(ASSUME["beautymart_commission_fallback"])
log("beautymart", "Commission stored as '25%' text", len(bm), "Converted to decimal rate")

names = bm["Product Name"].unique()
bm_map = {}
for n in names:
    s, sc = match_name(n)
    bm_map[n] = s
    crosswalk.append({"source": "BeautyMart", "source_key": n, "sku": s, "method": "fuzzy name", "score": sc})
low = [n for n in names if [c for c in crosswalk if c["source_key"] == n][0]["score"] < 0.6]
log("beautymart", f"No SKU column - only free-text product names in {len(names)} spellings", len(names),
    f"Fuzzy-matched to master; {len(low)} low-confidence matches sent for manual review")
brand_var = bm["Brand"].nunique()
log("beautymart", f"Brand name written {brand_var} different ways", len(bm), "Standardised (not used for matching)")
bm["sku"] = bm["Product Name"].map(bm_map)

status = bm["Order Status"]
bm_f = pd.DataFrame({
    "order_date": bm["order_date"], "channel": "BeautyMart", "order_id": bm["Sub Order No"], "sku": bm["sku"],
    "units": to_num(bm["Qty"]), "gross_sales": bm["MRP Value"], "discount": bm["Discount"],
})
bm_f["net_sales"] = bm["Selling Price"]
bm_f["platform_fees"] = bm["Selling Price"] * bm["comm"]
bm_f["logistics_cost"] = bm["Logistics Fee"] + np.where(status.eq("Returned"),
                                                        ASSUME["reverse_logistics_per_unit"] * bm_f["units"], 0)
is_ret = status.isin(["Returned", "RTO"])
bm_f["returned_units"] = np.where(is_ret, bm_f["units"], 0)
bm_f["return_value"] = np.where(is_ret, bm_f["net_sales"], 0.0)
bm_f["return_type"] = np.select([status.eq("RTO"), status.eq("Returned")], ["RTO", "Customer return"], "")
bm_f["stock_recovered"] = status.eq("RTO")
# Marketplace refunds commission on returned orders
bm_f["platform_fees"] = np.where(is_ret, 0.0, bm_f["platform_fees"])
bm_f["customer_shipping_paid"] = 0.0

# Shopify reverse logistics on customer returns
sh_f["logistics_cost"] = sh_f["logistics_cost"] + np.where(sh_f["returned_units"] > 0,
                                                           ASSUME["reverse_logistics_per_unit"] * sh_f["returned_units"], 0)

# ============================================================ 5. UNIFIED FACT TABLE
fact = pd.concat([sh_f, am_f, bm_f], ignore_index=True)
unm = fact["sku"].isna().sum()
log("all", "Rows whose product could not be matched", unm, "Excluded and listed for review" if unm else "None - all rows matched")
fact = fact.dropna(subset=["sku"])
fact = fact.merge(dim_product[["sku", "unit_cogs"]], on="sku", how="left")
fact["cogs"] = fact["unit_cogs"] * fact["units"]
# product cost written off when returned goods can't be resold
fact["cogs_recovered"] = np.where(fact["stock_recovered"], fact["unit_cogs"] * fact["returned_units"], 0.0)
fact["net_revenue"] = fact["net_sales"] - fact["return_value"] + fact["customer_shipping_paid"]
fact["contribution_before_ads"] = (fact["net_revenue"] - fact["platform_fees"] - fact["logistics_cost"]
                                   - fact["cogs"] + fact["cogs_recovered"])
fact["order_date"] = pd.to_datetime(fact["order_date"])
fact["month"] = fact["order_date"].dt.to_period("M").astype(str)
fact = fact.drop(columns="unit_cogs")

# ============================================================ 6. META ADS
ads = pd.read_csv(RAW / "meta_ads_daily.csv")
ads["date"] = pd.to_datetime(ads["Day"])
ads["spend"] = ads["Amount spent (INR)"]
def camp_sku(c):
    tag = c.replace("MM_", "").replace("_Prospecting", "")
    for sku in SKUS:
        if sku.replace("MM-", "").replace("-", "") == tag.replace("-", ""):
            return sku
    return None
ads["sku"] = ads["Campaign name"].map(camp_sku)
shared = ads["sku"].isna()
log("meta_ads", "Brand & retargeting campaigns not tied to one product", ads.loc[shared, "Campaign name"].nunique(),
    "Spend allocated to products by their share of Shopify net revenue that month")
ads["month"] = ads["date"].dt.to_period("M").astype(str)

direct = ads[~shared].groupby(["month", "sku"])["spend"].sum().reset_index().assign(ad_type="Product campaign")
pool = ads[shared].groupby("month")["spend"].sum()
shop_rev = (fact[fact.channel == "Shopify"].groupby(["month", "sku"])["net_revenue"].sum().reset_index())
shop_rev["share"] = shop_rev["net_revenue"] / shop_rev.groupby("month")["net_revenue"].transform("sum")
alloc = shop_rev.assign(spend=shop_rev["share"] * shop_rev["month"].map(pool))[["month", "sku", "spend"]]
alloc["ad_type"] = "Brand/retargeting (allocated)"
fact_ads = pd.concat([direct, alloc], ignore_index=True)
fact_ads["channel"] = "Shopify"   # Meta traffic lands on the brand website
assert abs(fact_ads["spend"].sum() - ads["spend"].sum()) < 1, "ad spend allocation must reconcile"

# ============================================================ 7. DATE DIM + SAVE
dim_date = pd.DataFrame({"date": pd.date_range("2026-04-01", "2026-09-30")})
dim_date["month"] = dim_date["date"].dt.to_period("M").astype(str)
dim_date["month_name"] = dim_date["date"].dt.strftime("%b %Y")
dim_date["week_start"] = dim_date["date"] - pd.to_timedelta(dim_date["date"].dt.weekday, unit="D")
dim_date["weekday"] = dim_date["date"].dt.day_name()

storage = storage_fees.assign(month=pd.to_datetime(storage_fees["date/time"], format="%d/%m/%Y %H:%M")
                              .dt.to_period("M").astype(str))[["month", "fba fees"]]
storage = storage.rename(columns={"fba fees": "amazon_storage_fee"})
storage["amazon_storage_fee"] = -storage["amazon_storage_fee"]

dq = pd.DataFrame(LOG)
cw = pd.DataFrame(crosswalk)
fact.to_csv(CLEAN / "fact_sales.csv", index=False)
fact_ads.to_csv(CLEAN / "fact_ads.csv", index=False)
dim_product.to_csv(CLEAN / "dim_product.csv", index=False)
dim_date.to_csv(CLEAN / "dim_date.csv", index=False)
cw.to_csv(CLEAN / "sku_crosswalk.csv", index=False)
dq.to_csv(CLEAN / "data_quality_log.csv", index=False)
storage.to_csv(CLEAN / "amazon_storage_fees.csv", index=False)

with pd.ExcelWriter(CLEAN / "profit_model.xlsx", engine="openpyxl") as w:
    for name, df in [("fact_sales", fact), ("fact_ads", fact_ads), ("dim_product", dim_product),
                     ("dim_date", dim_date), ("sku_crosswalk", cw), ("amazon_storage_fees", storage),
                     ("data_quality_log", dq), ("assumptions", pd.DataFrame(ASSUME.items(), columns=["rule", "value"]))]:
        df.to_excel(w, sheet_name=name, index=False)

print(f"Raw Shopify rows: {n0} | unified fact rows: {len(fact)} | ad rows: {len(fact_ads)}")
print(dq.to_string(index=False))
