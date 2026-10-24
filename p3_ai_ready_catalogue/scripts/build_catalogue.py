"""
AI-Ready Product Catalogue
==========================
Sample project · simulated data for a fictional skincare brand ("Mitti & Moss").

A brand wants a WhatsApp / website chatbot that answers "Is this good for oily skin?",
"Can I use it with retinol?", "How do I use it?". Its catalogue is a website export full of
HTML, free text, mixed units and inconsistent categories, so a chatbot built on it
hallucinates. This script turns it into:

  output/catalogue_structured.csv / .json  - one clean row per product, typed fields
  output/knowledge_base.jsonl              - grounded Q&A + fact chunks for a chatbot / RAG
  output/review_queue.csv                  - gaps a human must fill (never guessed)
  output/quality_report.json               - field completeness before vs after

Extraction is rule-based and transparent (every value traceable to the source text).
Run:  python build_catalogue.py
"""
from pathlib import Path
import html, json, re
import pandas as pd

BASE = Path(__file__).resolve().parent.parent
RAW, OUT = BASE / "raw", BASE / "output"
OUT.mkdir(exist_ok=True)

CATEGORY_MAP = {"serums": "Serum", "sun care": "Sunscreen", "moisturizers": "Moisturiser", "cleansers": "Cleanser",
                "lip": "Lip", "hair care": "Hair", "eye care": "Eye", "combos": "Kit", "masks": "Mask"}
INGREDIENTS = ["Ethyl Ascorbic Acid", "Ferulic Acid", "Licorice Root", "Niacinamide", "Zinc PCA", "Hyaluronic Acid",
               "Panthenol", "Retinol", "Bakuchiol", "Squalane", "Zinc Oxide", "Iron Oxides", "Ceramides", "Cholesterol",
               "Oat", "Green Tea", "Peptides", "Shea Butter", "Turmeric", "Tea Tree Oil", "Salicylic Acid", "Jojoba",
               "Sunflower", "Vitamin E", "Beetroot", "Kokum Butter", "Red Onion", "Bhringraj", "Amla", "Coconut",
               "Rice Water", "Keratin", "Caffeine", "Vitamin K", "Kaolin", "Bentonite"]
SKIN = {"oily": r"\boily\b", "dry": r"\b(very )?dry\b(?! face)", "combination": r"\bcombination\b",
        "normal": r"\bnormal\b", "sensitive": r"\bsensitive\b", "acne-prone": r"acne[- ]prone"}
CONCERNS = {"dark spots": r"dark spots|acne marks", "dullness": r"dull", "acne": r"\bacne\b", "pores": r"pores",
            "oil control": r"controls? (excess )?oil|oil and dirt|oil free|oil-free", "hydration": r"hydrat",
            "fine lines": r"fine lines", "uneven texture": r"texture", "sun protection": r"\bspf\b",
            "uneven tone": r"evens? tone", "barrier repair": r"barrier", "puffiness": r"puffiness",
            "dark circles": r"dark circles", "hair fall": r"hair ?fall", "frizz & damage": r"frizz|damaged",
            "dry lips": r"chapped|dry, chapped", "makeup removal": r"makeup"}
FLAGS = {"fragrance_free": (r"fragrance[- ]free|fragrance free", r"contains fragrance"),
         "vegan": (r"\bvegan\b", None), "non_comedogenic": (r"non[- ]comedogenic", None),
         "sulphate_free": (r"sulphate[- ]free|sls[- ]free", None), "cruelty_free": (r"cruelty[- ]free", None)}

def clean_html(s):
    s = html.unescape(str(s or ""))
    s = re.sub(r"<br\s*/?>|</p>|</div>", ". ", s)
    s = re.sub(r"<[^>]+>", "", s)
    s = re.sub(r"\s*\|\s*", ". ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return re.sub(r"(\.\s*){2,}", ". ", s).strip(" .") + "."

def parse_size(s):
    m = re.match(r"\s*([\d.]+)\s*(ml|g|gm|fl ?oz|pcs)\s*$", str(s).lower())
    if not m: return None, None
    v, u = float(m.group(1)), m.group(2).replace(" ", "")
    if u == "floz": return 30.0, "ml"            # 1 fl oz bottle is sold as 30 ml
    return v, {"gm": "g"}.get(u, u)

def parse_price(s):
    return float(re.sub(r"[^\d.]", "", str(s).replace("/-", "")).strip(".") or "nan")

def extract(text):
    t = text.lower()
    ingredients = []
    for ing in INGREDIENTS:
        m = re.search(rf"(\d+(?:\.\d+)?%)\s*(?:encapsulated\s+)?{re.escape(ing.lower())}|{re.escape(ing.lower())}\s*(\d+(?:\.\d+)?%)", t)
        if ing.lower() in t:
            pct = (m.group(1) or m.group(2)) if m else None
            ingredients.append(f"{ing} {pct}" if pct else ing)
    skin = [k for k, p in SKIN.items() if re.search(p, t)]
    if re.search(r"all (hair|skin) types", t) and "hair" not in t:
        skin = sorted(set(skin) | {"oily", "dry", "combination", "normal"} | ({"sensitive"} if "sensitive" in t else set()))
    if re.search(r"not (recommended )?for sensitive", t):
        skin = [s for s in skin if s != "sensitive"]
    concerns = [k for k, p in CONCERNS.items() if re.search(p, t)]
    flags = {}
    for k, (yes, no) in FLAGS.items():
        flags[k] = False if (no and re.search(no, t)) else (True if re.search(yes, t) else None)
    return ingredients, skin, concerns, flags

def usage_time(how):
    h = str(how or "").lower()
    am = bool(re.search(r"\bam\b|morning|before stepping out", h)); pm = bool(re.search(r"\bpm\b|\bnight\b", h))
    return "AM & PM" if am and pm else "AM" if am else "PM only" if pm else ("Any time" if h.strip() else None)

raw = pd.read_csv(RAW / "website_catalogue_export.csv", dtype=str).fillna("")
log = [{"issue": "Duplicate products from repeated export", "rows": int(raw.duplicated().sum()), "action": "Removed"}]
raw = raw.drop_duplicates()
raw["sku"] = raw["Variant SKU"].str.strip().str.upper()
log.append({"issue": "SKU case inconsistent", "rows": int((raw["Variant SKU"] != raw["sku"]).sum()), "action": "Upper-cased"})

records, review = [], []
for _, r in raw.iterrows():
    desc = clean_html(r["Body (HTML)"])
    size_v, size_u = parse_size(r["Size"])
    ing, skin, conc, flags = extract(desc)
    rec = {
        "sku": r["sku"], "name": r["Title"].title().replace("Spf", "SPF").replace("&", "&"),
        "category": CATEGORY_MAP.get(r["Collection"].strip().lower()),
        "size_value": size_v, "size_unit": size_u, "price_inr": parse_price(r["Price"]),
        "in_stock": r["Stock status"].strip().lower() in ("in stock", "available", "in_stock"),
        "key_ingredients": ing, "skin_types": skin, "concerns": conc, **flags,
        "usage_time": usage_time(r["How to use"]),
        "how_to_use": r["How to use"].strip() or None, "cautions": r["Caution"].strip() or None,
        "pregnancy_caution": bool(re.search(r"pregnan", r["Caution"].lower())),
        "avoid_combining_with": re.findall(r"(retinol|vitamin c|ahas?)", r["Caution"].lower()),
        "description": desc,
    }
    # shade-only products (lip tints) have no skin type; that's expected, not a gap
    for field, needed in [("how_to_use", True), ("skin_types", rec["category"] not in ("Lip", "Hair", "Kit")),
                          ("fragrance_free", rec["category"] not in ("Hair", "Kit"))]:
        v = rec[field]
        if needed and (v is None or v == []):
            review.append({"sku": rec["sku"], "name": rec["name"], "missing": field,
                           "why_it_matters": {"how_to_use": "Chatbot cannot answer 'how do I use it?'",
                                              "skin_types": "Chatbot cannot answer 'is it for my skin type?'",
                                              "fragrance_free": "Sensitive-skin customers ask this often"}[field]})
    records.append(rec)
cat = pd.DataFrame(records).astype(object)
cat = cat.where(cat.notna(), None)

# --------------------------------------------------- completeness: before vs after
fields = ["category", "size", "price", "key_ingredients", "skin_types", "concerns", "fragrance_free", "usage_time"]
before = {  # how many fields a program could read directly from the raw export, without parsing
    "category": 0, "size": 0, "price": int(raw["Price"].str.fullmatch(r"\d+(\.\d+)?").sum()),
    "key_ingredients": 0, "skin_types": 0, "concerns": 0, "fragrance_free": 0, "usage_time": 0}
after = {
    "category": int(cat["category"].notna().sum()), "size": int(cat["size_value"].notna().sum()),
    "price": int(cat["price_inr"].notna().sum()), "key_ingredients": int((cat["key_ingredients"].str.len() > 0).sum()),
    "skin_types": int((cat["skin_types"].str.len() > 0).sum()), "concerns": int((cat["concerns"].str.len() > 0).sum()),
    "fragrance_free": int(cat["fragrance_free"].notna().sum()), "usage_time": int(cat["usage_time"].notna().sum())}
n = len(cat)
quality = {"products": n, "before_pct": round(100 * sum(before.values()) / (n * len(fields)), 1),
           "after_pct": round(100 * sum(after.values()) / (n * len(fields)), 1),
           "by_field": {f: {"before": before[f], "after": after[f], "of": n} for f in fields},
           "review_items": len(review), "cleaning_log": log}

# --------------------------------------------------- knowledge base for a chatbot
kb = []
def add(sku, kind, q, a, src):
    kb.append({"id": f"{sku}-{kind}-{len(kb)}", "sku": sku, "type": kind, "question": q, "answer": a, "source_fields": src})
LABEL = {"oily": "oily", "dry": "dry", "combination": "combination", "normal": "normal", "sensitive": "sensitive", "acne-prone": "acne-prone"}
for r in cat.itertuples():
    price = f"₹{r.price_inr:,.0f}"
    size = f"{r.size_value:g} {r.size_unit}" if r.size_value else ""
    fact = (f"{r.name} ({r.sku}) is a {r.category.lower()} product, {size}, priced at {price}. "
            + (f"Key ingredients: {', '.join(r.key_ingredients)}. " if r.key_ingredients else "")
            + (f"Suitable for {', '.join(r.skin_types)} skin. " if r.skin_types else "")
            + (f"Helps with: {', '.join(r.concerns)}. " if r.concerns else "")
            + ("Fragrance free. " if r.fragrance_free else "Contains fragrance. " if r.fragrance_free is False else "")
            + ("Vegan. " if r.vegan else ""))
    kb.append({"id": f"{r.sku}-fact", "sku": r.sku, "type": "fact_chunk", "text": fact.strip(),
               "source_fields": ["name", "category", "size", "price", "key_ingredients", "skin_types", "concerns"]})
    for s in ["oily", "dry", "sensitive", "acne-prone"]:
        if r.category in ("Lip", "Hair", "Kit"): break
        if s in r.skin_types:
            add(r.sku, "skin_fit", f"Is {r.name} good for {s} skin?", f"Yes — {r.name} is listed as suitable for {s} skin.", ["skin_types"])
        elif r.skin_types:
            add(r.sku, "skin_fit", f"Is {r.name} good for {s} skin?",
                f"It is not listed for {s} skin. It is recommended for {', '.join(r.skin_types)} skin. "
                + "For personal advice, please consult a dermatologist.", ["skin_types"])
    if isinstance(r.how_to_use, str):
        add(r.sku, "usage", f"How do I use {r.name}?", r.how_to_use + (f" Best used: {r.usage_time}." if isinstance(r.usage_time, str) else ""), ["how_to_use", "usage_time"])
    if r.key_ingredients:
        add(r.sku, "ingredients", f"What are the key ingredients in {r.name}?", ", ".join(r.key_ingredients) + ".", ["key_ingredients"])
    if r.fragrance_free is not None:
        add(r.sku, "fragrance", f"Is {r.name} fragrance free?", "Yes, it is fragrance free." if r.fragrance_free else "No, it contains fragrance.", ["fragrance_free"])
    for combo in r.avoid_combining_with:
        add(r.sku, "compatibility", f"Can I use {r.name} with {dict(ahas='AHAs', aha='AHAs').get(combo, combo.title())}?",
            f"Not at the same time. {r.cautions}", ["cautions"])
    if r.pregnancy_caution:
        add(r.sku, "safety", f"Can I use {r.name} during pregnancy?", f"The brand advises against it: {r.cautions} Please consult your doctor.", ["cautions"])
# cross-product: routine questions answered from structured fields only
for concern in ["acne", "dark spots", "hydration", "hair fall", "oil control"]:
    hits = cat[cat["concerns"].apply(lambda c: concern in c)]
    if len(hits):
        add("ALL", "recommendation", f"Which products help with {concern}?",
            "; ".join(f"{h.name} ({h.category}, ₹{h.price_inr:,.0f})" for h in hits.itertuples()) + ".", ["concerns", "price_inr"])
kb.append({"id": "SYSTEM-guardrail", "sku": "ALL", "type": "bot_instruction",
           "text": "Answer only from these entries. If the catalogue has no answer, say so and offer a human agent. "
                   "Never give medical advice; suggest a dermatologist for skin conditions."})

# --------------------------------------------------- save
cat_out = cat.copy()
for c in ["key_ingredients", "skin_types", "concerns", "avoid_combining_with"]:
    cat_out[c] = cat_out[c].apply(lambda v: "; ".join(v))
cat_out.to_csv(OUT / "catalogue_structured.csv", index=False)
cat.to_json(OUT / "catalogue_structured.json", orient="records", indent=2, force_ascii=False)
with open(OUT / "knowledge_base.jsonl", "w", encoding="utf-8") as f:
    for row in kb: f.write(json.dumps(row, ensure_ascii=False) + "\n")
pd.DataFrame(review).to_csv(OUT / "review_queue.csv", index=False)
(OUT / "quality_report.json").write_text(json.dumps(quality, indent=2))
print(f"{n} products | completeness {quality['before_pct']}% -> {quality['after_pct']}% | KB entries {len(kb)} | review items {len(review)}")
