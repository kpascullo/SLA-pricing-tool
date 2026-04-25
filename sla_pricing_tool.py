#!/usr/bin/env python3
"""NYS SLA wholesale wine pricing scraper + Banville competitiveness workbook."""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

import pandas as pd
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

BASE_URL = "https://www.nyslapricepostings.com/public/price-lookup?post_type=WW"
OUTPUT_XLSX = "sla_competitive_pricing.xlsx"


@dataclass
class SearchItem:
    search_term: str
    category: str


PRICE_RE = re.compile(r"\$?\s*([0-9]+(?:\.[0-9]{1,2})?)")
PACK_RE = re.compile(r"\b(\d{1,3})\s*(?:pk|pack|case|cs|ct|count)?\b", re.IGNORECASE)
DISCOUNT_RE = re.compile(
    r"(?:qty|quantity|disc(?:ount)?|case|cs)?\s*(\d{1,4})\s*(?:\+|or more|cases?|cs)?\s*(?:at|@|for)?\s*\$\s*([0-9]+(?:\.[0-9]{1,2})?)",
    re.IGNORECASE,
)


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    match = PRICE_RE.search(text.replace(",", ""))
    if not match:
        return None
    try:
        return float(match.group(1))
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    if value is None:
        return None
    if isinstance(value, int):
        return value
    text = str(value).strip()
    if not text:
        return None
    match = re.search(r"(\d+)", text)
    if not match:
        return None
    return int(match.group(1))


def normalize_col_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_")


def read_input_csv(path: Path) -> list[SearchItem]:
    if not path.exists():
        raise FileNotFoundError(f"Required file not found: {path}")
    df = pd.read_csv(path)
    required = {"search_term", "category"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"{path} is missing columns: {sorted(missing)}")
    return [
        SearchItem(str(row["search_term"]).strip(), str(row["category"]).strip())
        for _, row in df.iterrows()
        if str(row["search_term"]).strip()
    ]


def find_field(cells: dict[str, str], raw: str, aliases: list[str]) -> Optional[str]:
    for alias in aliases:
        for key, value in cells.items():
            if alias in key and str(value).strip():
                return str(value).strip()
    for alias in aliases:
        alias_pattern = alias.replace("_", r"[\s_-]*")
        pattern = re.compile(rf"{alias_pattern}\s*[:\-]\s*([^|\n]+)", re.IGNORECASE)
        match = pattern.search(raw)
        if match:
            return match.group(1).strip()
    return None


def parse_discounts(raw_text: str) -> tuple[list[Optional[int]], list[Optional[float]], str]:
    quantities: list[Optional[int]] = []
    prices: list[Optional[float]] = []
    seen: set[tuple[int, float]] = set()

    for qty, price in DISCOUNT_RE.findall(raw_text):
        q = to_int(qty)
        p = to_float(price)
        if q is None or p is None:
            continue
        key = (q, p)
        if key in seen:
            continue
        seen.add(key)
        quantities.append(q)
        prices.append(p)
        if len(quantities) == 4:
            break

    while len(quantities) < 4:
        quantities.append(None)
        prices.append(None)

    all_discount_text = " ; ".join(
        f"{q} => ${p:.2f}" for q, p in zip(quantities, prices) if q is not None and p is not None
    )
    return quantities, prices, all_discount_text


def parse_case_pack(raw_text: str, case_pack_candidate: Optional[str]) -> Optional[int]:
    if case_pack_candidate:
        parsed = to_int(case_pack_candidate)
        if parsed:
            return parsed
    pack_match = PACK_RE.search(raw_text)
    if pack_match:
        parsed = to_int(pack_match.group(1))
        if parsed and parsed <= 60:
            return parsed
    return None


def derive_prices(row: dict[str, Any]) -> dict[str, Any]:
    case_pack = to_int(row.get("case_pack"))
    bottle_price = to_float(row.get("bottle_price"))
    case_price = to_float(row.get("case_price"))

    if case_pack and case_price is not None and bottle_price is None:
        bottle_price = round(case_price / case_pack, 4)
    if case_pack and bottle_price is not None and case_price is None:
        case_price = round(bottle_price * case_pack, 4)

    discount_quantities = [row.get(f"discount_quantity_{i}") for i in range(1, 5)]
    discount_prices = [row.get(f"discount_price_{i}") for i in range(1, 5)]

    valid_discounts: list[tuple[int, float]] = []
    for qty, price in zip(discount_quantities, discount_prices):
        q = to_int(qty)
        p = to_float(price)
        if q is not None and p is not None:
            valid_discounts.append((q, p))

    best_discount_quantity: Optional[int] = None
    best_discount_case_price: Optional[float] = None

    if valid_discounts:
        best_discount_quantity, best_discount_case_price = min(valid_discounts, key=lambda x: x[1])
    elif case_price is not None:
        best_discount_quantity = 1
        best_discount_case_price = case_price

    best_discount_bottle_price = None
    if case_pack and best_discount_case_price is not None:
        best_discount_bottle_price = round(best_discount_case_price / case_pack, 4)

    if case_price is None and bottle_price is None and best_discount_case_price is None:
        price_status = "PRICE_NOT_FOUND"
    else:
        price_status = "OK"

    row.update(
        {
            "case_pack": case_pack,
            "bottle_price": bottle_price,
            "case_price": case_price,
            "best_discount_quantity": best_discount_quantity,
            "best_discount_case_price": best_discount_case_price,
            "best_discount_bottle_price": best_discount_bottle_price,
            "price_status": price_status,
        }
    )
    return row


def select_region_if_present(page) -> None:
    region_targets = ["southern new york", "new york city", "nyc"]
    selectors = [
        "select[name*='region']",
        "select[id*='region']",
        "select[name*='zone']",
        "select[id*='zone']",
        "select",
    ]
    for selector in selectors:
        elements = page.locator(selector)
        count = elements.count()
        for i in range(count):
            sel = elements.nth(i)
            options = sel.locator("option")
            for j in range(options.count()):
                text = options.nth(j).inner_text().strip().lower()
                value = options.nth(j).get_attribute("value") or ""
                if any(target in text for target in region_targets):
                    sel.select_option(value=value)
                    page.wait_for_timeout(400)
                    return


def get_table_headers(page) -> list[str]:
    header_selectors = [
        "table thead tr th",
        "[role='table'] [role='columnheader']",
        ".dataTables_scrollHead table thead tr th",
    ]
    for sel in header_selectors:
        headers = [normalize_col_name(h.inner_text()) for h in page.locator(sel).all() if h.inner_text().strip()]
        if headers:
            return headers
    return []


def find_search_input(page):
    candidates = [
        "input[placeholder*='Search']",
        "input[type='search']",
        "input[name*='search']",
        "input[id*='search']",
        "input[type='text']",
    ]
    for selector in candidates:
        locator = page.locator(selector).first
        if locator.count() and locator.is_visible():
            return locator
    return None


def click_search_if_present(page) -> None:
    for selector in ["button:has-text('Search')", "input[type='submit']", "button[type='submit']"]:
        button = page.locator(selector).first
        if button.count() and button.is_visible():
            button.click()
            page.wait_for_timeout(500)
            return


def maybe_expand_row(row_locator) -> str:
    extra_text = ""
    for selector in ["button:has-text('Details')", "button:has-text('View')", "a:has-text('Details')", "a:has-text('View')"]:
        btn = row_locator.locator(selector).first
        if btn.count() and btn.is_visible():
            try:
                btn.click(timeout=1000)
                row_locator.page.wait_for_timeout(300)
                extra_text = row_locator.inner_text(timeout=1000)
            except PlaywrightTimeoutError:
                pass
            except Exception:
                pass
            break
    return extra_text


def extract_rows_for_page(page, search_term: str, category: str, source_url: str, page_no: int) -> list[dict[str, Any]]:
    headers = get_table_headers(page)
    row_selectors = ["table tbody tr", "[role='rowgroup'] [role='row']"]
    rows = []

    locator = None
    for selector in row_selectors:
        loc = page.locator(selector)
        if loc.count() > 0:
            locator = loc
            break
    if locator is None:
        return rows

    row_count = locator.count()
    for idx in range(row_count):
        row = locator.nth(idx)
        try:
            row_text = row.inner_text(timeout=1200).strip()
        except Exception:
            row_text = ""
        if not row_text:
            continue

        expanded_text = maybe_expand_row(row)
        raw_text = (row_text + "\n" + expanded_text).strip()

        cells = row.locator("td")
        cell_texts = [cells.nth(i).inner_text().strip() for i in range(cells.count())]
        mapped: dict[str, str] = {}
        if headers and len(headers) == len(cell_texts):
            mapped = {headers[i]: cell_texts[i] for i in range(len(headers))}
        else:
            mapped = {f"col_{i+1}": value for i, value in enumerate(cell_texts)}

        distributor = find_field(mapped, raw_text, ["distributor", "wholesaler", "supplier"])
        brand = find_field(mapped, raw_text, ["brand", "label"])
        wine_name = find_field(mapped, raw_text, ["wine_name", "item_description", "item", "description", "product"])
        vintage = find_field(mapped, raw_text, ["vintage", "year"])
        origin = find_field(mapped, raw_text, ["origin", "country", "region", "appellation"])
        bottle_size = find_field(mapped, raw_text, ["bottle_size", "size", "ml", "liter"])
        case_pack_raw = find_field(mapped, raw_text, ["case_pack", "pack", "case"])
        bottle_price_raw = find_field(mapped, raw_text, ["bottle_price", "unit_price", "bottle"])
        case_price_raw = find_field(mapped, raw_text, ["case_price", "case"])
        posting_month = find_field(mapped, raw_text, ["posting_month", "month", "post"])

        quantities, discount_prices, all_discount_text = parse_discounts(raw_text)

        parsed = {
            "search_term": search_term,
            "category": category,
            "distributor": distributor,
            "brand": brand,
            "wine_name": wine_name,
            "vintage": vintage,
            "origin": origin,
            "bottle_size": bottle_size,
            "case_pack": parse_case_pack(raw_text, case_pack_raw),
            "bottle_price": to_float(bottle_price_raw),
            "case_price": to_float(case_price_raw),
            "discount_quantity_1": quantities[0],
            "discount_price_1": discount_prices[0],
            "discount_quantity_2": quantities[1],
            "discount_price_2": discount_prices[1],
            "discount_quantity_3": quantities[2],
            "discount_price_3": discount_prices[2],
            "discount_quantity_4": quantities[3],
            "discount_price_4": discount_prices[3],
            "all_discount_text": all_discount_text,
            "posting_month": posting_month,
            "source_url": source_url,
            "raw_result_text": raw_text,
            "result_page_number": page_no,
            "result_row_number": idx + 1,
        }
        rows.append(derive_prices(parsed))

    return rows


def goto_next_page(page) -> bool:
    next_selectors = [
        "a[rel='next']",
        "button:has-text('Next')",
        "li.next:not(.disabled) a",
        "a:has-text('Next')",
    ]
    for selector in next_selectors:
        next_button = page.locator(selector).first
        if next_button.count() and next_button.is_visible():
            classes = (next_button.get_attribute("class") or "").lower()
            aria_disabled = (next_button.get_attribute("aria-disabled") or "false").lower()
            if "disabled" in classes or aria_disabled == "true":
                continue
            try:
                next_button.click(timeout=1000)
                page.wait_for_timeout(800)
                return True
            except Exception:
                continue
    return False


def search_term_rows(page, item: SearchItem) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    page.goto(BASE_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(1200)
    select_region_if_present(page)

    search_input = find_search_input(page)
    if search_input is None:
        raise RuntimeError("Could not locate a search input on the NYS SLA page.")

    search_input.click()
    search_input.fill("")
    search_input.fill(item.search_term)
    click_search_if_present(page)

    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except PlaywrightTimeoutError:
        pass

    page.wait_for_timeout(1200)

    all_rows: list[dict[str, Any]] = []
    debug_rows: list[dict[str, Any]] = []

    page_no = 1
    while True:
        rows = extract_rows_for_page(page, item.search_term, item.category, page.url, page_no)
        all_rows.extend(rows)
        for row in rows:
            debug_rows.append(
                {
                    "search_term": row["search_term"],
                    "category": row["category"],
                    "source_url": row["source_url"],
                    "result_page_number": row["result_page_number"],
                    "result_row_number": row["result_row_number"],
                    "raw_result_text": row["raw_result_text"],
                }
            )

        if not goto_next_page(page):
            break
        page_no += 1

    if not all_rows:
        debug_rows.append(
            {
                "search_term": item.search_term,
                "category": item.category,
                "source_url": page.url,
                "result_page_number": 1,
                "result_row_number": None,
                "raw_result_text": "NO_RESULT_ROWS_FOUND",
            }
        )

    return all_rows, debug_rows


def compare_competitor_vs_banville(competitor_df: pd.DataFrame, banville_df: pd.DataFrame) -> pd.DataFrame:
    comparisons: list[dict[str, Any]] = []

    for _, comp in competitor_df.iterrows():
        comp_cat = str(comp.get("category", "")).strip()
        banville_matches = banville_df[banville_df["category"].astype(str).str.strip() == comp_cat]

        if banville_matches.empty:
            comparisons.append(
                {
                    "competitor_search_term": comp.get("search_term"),
                    "competitor_category": comp_cat,
                    "competitor_distributor": comp.get("distributor"),
                    "competitor_wine_name": comp.get("wine_name"),
                    "competitor_case_price": comp.get("case_price"),
                    "competitor_best_discount_quantity": comp.get("best_discount_quantity"),
                    "competitor_best_discount_case_price": comp.get("best_discount_case_price"),
                    "competitor_best_discount_bottle_price": comp.get("best_discount_bottle_price"),
                    "banville_search_term": None,
                    "banville_distributor": None,
                    "banville_wine_name": None,
                    "banville_case_price": None,
                    "banville_best_discount_quantity": None,
                    "banville_best_discount_case_price": None,
                    "banville_best_discount_bottle_price": None,
                    "dollar_difference_per_case": None,
                    "percent_difference": None,
                    "estimated_savings_per_5_cases": None,
                    "estimated_savings_per_10_cases": None,
                    "classification": "Missing Price",
                }
            )
            continue

        for _, ban in banville_matches.iterrows():
            comp_best = to_float(comp.get("best_discount_case_price"))
            ban_best = to_float(ban.get("best_discount_case_price"))

            if comp_best is None or ban_best is None:
                classification = "Missing Price"
                diff = None
                pct = None
                save5 = None
                save10 = None
            else:
                diff = round(comp_best - ban_best, 4)
                pct = round(((ban_best - comp_best) / comp_best) * 100, 4) if comp_best else None
                save5 = round(diff * 5, 4)
                save10 = round(diff * 10, 4)

                if ban_best < comp_best:
                    classification = "Cheaper"
                elif ban_best <= comp_best * 1.05:
                    classification = "Competitive"
                else:
                    classification = "More Expensive"

            comparisons.append(
                {
                    "competitor_search_term": comp.get("search_term"),
                    "competitor_category": comp_cat,
                    "competitor_distributor": comp.get("distributor"),
                    "competitor_wine_name": comp.get("wine_name"),
                    "competitor_case_price": comp.get("case_price"),
                    "competitor_best_discount_quantity": comp.get("best_discount_quantity"),
                    "competitor_best_discount_case_price": comp_best,
                    "competitor_best_discount_bottle_price": comp.get("best_discount_bottle_price"),
                    "banville_search_term": ban.get("search_term"),
                    "banville_distributor": ban.get("distributor"),
                    "banville_wine_name": ban.get("wine_name"),
                    "banville_case_price": ban.get("case_price"),
                    "banville_best_discount_quantity": ban.get("best_discount_quantity"),
                    "banville_best_discount_case_price": ban_best,
                    "banville_best_discount_bottle_price": ban.get("best_discount_bottle_price"),
                    "dollar_difference_per_case": diff,
                    "percent_difference": pct,
                    "estimated_savings_per_5_cases": save5,
                    "estimated_savings_per_10_cases": save10,
                    "classification": classification,
                }
            )

    return pd.DataFrame(comparisons)


def build_summary(
    competitor_searches: list[SearchItem],
    banville_searches: list[SearchItem],
    competitor_df: pd.DataFrame,
    banville_df: pd.DataFrame,
    comparison_df: pd.DataFrame,
) -> pd.DataFrame:
    cheaper_categories = sorted(set(comparison_df.loc[comparison_df["classification"] == "Cheaper", "competitor_category"].dropna()))
    competitive_categories = sorted(
        set(comparison_df.loc[comparison_df["classification"] == "Competitive", "competitor_category"].dropna())
    )
    not_competitive_categories = sorted(
        set(comparison_df.loc[comparison_df["classification"] == "More Expensive", "competitor_category"].dropna())
    )

    missing_wines = sorted(
        set(competitor_df.loc[competitor_df["price_status"] == "PRICE_NOT_FOUND", "search_term"].dropna())
        | set(banville_df.loc[banville_df["price_status"] == "PRICE_NOT_FOUND", "search_term"].dropna())
    )

    rows = [
        ("number_of_competitor_skus_searched", len(competitor_searches)),
        ("number_of_competitor_skus_with_pricing_found", competitor_df[competitor_df["price_status"] == "OK"]["search_term"].nunique()),
        ("number_of_banville_searches_run", len(banville_searches)),
        ("number_of_banville_wines_with_pricing_found", banville_df[banville_df["price_status"] == "OK"]["search_term"].nunique()),
        ("categories_where_banville_has_cheaper_options", ", ".join(cheaper_categories) or "None"),
        ("categories_where_banville_is_competitive", ", ".join(competitive_categories) or "None"),
        ("categories_where_banville_is_not_competitive", ", ".join(not_competitive_categories) or "None"),
        ("wines_where_price_not_found_happened", ", ".join(missing_wines) or "None"),
    ]
    return pd.DataFrame(rows, columns=["metric", "value"])


def ensure_expected_columns(df: pd.DataFrame) -> pd.DataFrame:
    expected = [
        "search_term",
        "category",
        "distributor",
        "brand",
        "wine_name",
        "vintage",
        "origin",
        "bottle_size",
        "case_pack",
        "bottle_price",
        "case_price",
        "discount_quantity_1",
        "discount_price_1",
        "discount_quantity_2",
        "discount_price_2",
        "discount_quantity_3",
        "discount_price_3",
        "discount_quantity_4",
        "discount_price_4",
        "all_discount_text",
        "posting_month",
        "best_discount_quantity",
        "best_discount_case_price",
        "best_discount_bottle_price",
        "price_status",
        "source_url",
        "raw_result_text",
        "result_page_number",
        "result_row_number",
    ]
    for col in expected:
        if col not in df.columns:
            df[col] = None
    return df[expected]


def append_optional_manual_banville_searches(items: list[SearchItem]) -> list[SearchItem]:
    optional_file = Path("banville_exact_skus.csv")
    if not optional_file.exists():
        return items
    print("Detected banville_exact_skus.csv. Including additional manual Banville searches.")
    return items + read_input_csv(optional_file)


def main() -> int:
    competitor_searches = read_input_csv(Path("competitors.csv"))
    banville_searches = append_optional_manual_banville_searches(read_input_csv(Path("banville_searches.csv")))

    all_competitor_rows: list[dict[str, Any]] = []
    all_banville_rows: list[dict[str, Any]] = []
    debug_rows: list[dict[str, Any]] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        for item in competitor_searches:
            print(f"[Competitor] Searching: {item.search_term}")
            rows, debug = search_term_rows(page, item)
            all_competitor_rows.extend(rows)
            debug_rows.extend(debug)

        for item in banville_searches:
            print(f"[Banville] Searching: {item.search_term}")
            rows, debug = search_term_rows(page, item)
            all_banville_rows.extend(rows)
            debug_rows.extend(debug)

        context.close()
        browser.close()

    competitor_df = ensure_expected_columns(pd.DataFrame(all_competitor_rows))
    banville_df = ensure_expected_columns(pd.DataFrame(all_banville_rows))

    comparison_df = compare_competitor_vs_banville(competitor_df, banville_df)
    substitutes_df = comparison_df[comparison_df["classification"].isin(["Cheaper", "Competitive"])].copy()

    summary_df = build_summary(competitor_searches, banville_searches, competitor_df, banville_df, comparison_df)
    debug_df = pd.DataFrame(debug_rows)

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        competitor_df.to_excel(writer, sheet_name="Competitor Pricing", index=False)
        banville_df.to_excel(writer, sheet_name="Banville Pricing", index=False)
        substitutes_df.to_excel(writer, sheet_name="Best Banville Substitutes", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        debug_df.to_excel(writer, sheet_name="Debug Raw Results", index=False)

    failed_competitors = competitor_df.loc[competitor_df["price_status"] == "PRICE_NOT_FOUND", "search_term"].dropna().unique().tolist()
    failed_banville = banville_df.loc[banville_df["price_status"] == "PRICE_NOT_FOUND", "search_term"].dropna().unique().tolist()

    print(f"\nWorkbook created: {OUTPUT_XLSX}")
    if failed_competitors or failed_banville:
        print("\nWARNING: Some searches returned no parsable prices.")
        if failed_competitors:
            print("Competitor searches with PRICE_NOT_FOUND:")
            for term in failed_competitors:
                print(f"  - {term}")
        if failed_banville:
            print("Banville searches with PRICE_NOT_FOUND:")
            for term in failed_banville:
                print(f"  - {term}")
        print("Debug text has been saved in the 'Debug Raw Results' sheet.")

    if competitor_df.empty and banville_df.empty:
        print("ERROR: No rows were scraped from the site. Check selectors and site availability.")
        return 2

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Fatal error: {exc}", file=sys.stderr)
        raise
