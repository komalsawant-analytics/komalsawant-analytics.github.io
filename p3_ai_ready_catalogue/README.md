# P3 · AI-Ready Product Catalogue

**Sample project · simulated data · fictional brand "Mitti & Moss" (skincare)**

## The problem
The brand wants a WhatsApp/website chatbot that answers "Is this good for oily skin?", "Can I use it with retinol?", "How do I use it?". Its catalogue (`raw/website_catalogue_export.csv`) is a website export with HTML in descriptions, ingredients and skin types buried in free text, sizes in ml / gm / fl oz, prices like "Rs. 899/-", inconsistent categories and duplicate rows. A bot built on that guesses.

## What the script does
`python scripts/build_catalogue.py`

- Strips HTML, normalises units, prices and categories, removes duplicates.
- Extracts key ingredients (with strengths like "Retinol 0.3%"), skin types, concerns, fragrance-free / vegan flags, usage time, pregnancy cautions and "don't combine with" rules — every value traceable to the source text.
- Never guesses: 12 missing values go to `output/review_queue.csv` for the brand to fill.

## Outputs
| File | Use |
|---|---|
| `catalogue_structured.csv / .json` | Clean product data for the website, feeds and the bot |
| `knowledge_base.jsonl` | 127 grounded entries (fact chunks + Q&A) for a chatbot or RAG system, with a guardrail instruction |
| `review_queue.csv` | Gaps a human must fill |
| `quality_report.json` | Machine-readable fields: **1.9% before → 84.4% after** |
