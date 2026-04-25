# sla-pricing-tool

Beginner-friendly Python project to scrape the NYS SLA wholesale wine lookup site with **Playwright** (JavaScript-enabled), parse pricing, compare competitors vs. Banville alternatives, and export to Excel.

## What this tool does

1. Reads search lists from:
   - `competitors.csv`
   - `banville_searches.csv`
2. Opens NYS SLA price lookup:
   - https://www.nyslapricepostings.com/public/price-lookup?post_type=WW
3. Searches each term, captures every row, tries to expand details, and handles pagination.
4. Extracts pricing fields and computes:
   - derived bottle/case prices
   - best discount case/bottle price
5. Compares competitor vs Banville wines by category.
6. Writes `sla_competitive_pricing.xlsx` with 5 sheets:
   - Competitor Pricing
   - Banville Pricing
   - Best Banville Substitutes
   - Summary
   - Debug Raw Results

## Project files

- `sla_pricing_tool.py` – main script
- `requirements.txt` – dependencies
- `competitors.csv` – competitor search terms
- `banville_searches.csv` – Banville search terms

Optional:
- `banville_exact_skus.csv` – extra manual Banville SKUs you want to append (same columns: `search_term,category`)

## Setup (copy/paste)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m playwright install
python3 sla_pricing_tool.py
```

## Run

```bash
python3 sla_pricing_tool.py
```

## Output

- `sla_competitive_pricing.xlsx`

If no prices are found for any term, the tool **does not fail silently**:
- it prints a warning with failed search terms
- it stores raw scraped text in `Debug Raw Results` for troubleshooting

## Notes on selectors and website changes

The NYS SLA site can change field names and table structure. This script uses flexible/heuristic selectors and regex parsing, but if the site layout changes significantly, update selectors in `sla_pricing_tool.py`:
- `find_search_input`
- `extract_rows_for_page`
- `goto_next_page`
- `select_region_if_present`
