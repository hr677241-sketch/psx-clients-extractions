"""
PSX Auditor Scraper — GitHub Actions Edition (v2)
=============================================
Scrapes every listed company on PSX and finds their statutory auditor.
Saves all results + a filtered sheet for the TARGET_AUDITOR.

WHAT CHANGED IN v2 (link-gathering rewrite)
--------------------------------------------
v1 assumed the "Listed Companies" page was a paginated jQuery DataTable that
could be flipped to "show all" and scraped with plain BeautifulSoup. That's
not actually how the page works:

  1. The listing page (www.psx.com.pk/.../listed-companies) is a SECTOR /
     SYMBOL filter widget, not a table with pagination. By default it shows
     "No Result Found!" — a row only appears after you pick something from
     the "Select Sector" or "Select Symbol" dropdown, which fires an AJAX
     call. There's no "Show All" button to find.

  2. HOWEVER — the "Select Symbol" dropdown itself is rendered server-side.
     It already contains the full list of ~700 PSX ticker symbols in the
     raw HTML, with zero AJAX/JS required to read it. I confirmed this by
     fetching the page directly.

  3. The bigger issue: that listing page never had a per-company detail
     page with an "Auditor" field to begin with — there's no
     /listed-companies/{SYMBOL} sub-page on www.psx.com.pk at all. The
     actual structured Auditor field lives on PSX's *Data Portal* at
     https://dps.psx.com.pk/company/{SYMBOL} — confirmed for OGDC and FFC,
     both of which render a clean "AUDITOR" label/value pair server-side
     (no JS needed to see it).

So v2 gets the master symbol list cheaply and reliably (Strategy 1, no
browser needed), and points every company URL at the Data Portal instead
of the old broken path. Strategies 2 and 3 exist as escalating fallbacks
in case PSX changes its markup so Strategy 1 stops working.

  Strategy 1 — STATIC PAGE PARSE (primary)
      Plain `requests.get()` on the listing page, parse the "Select Symbol"
      dropdown directly. No Selenium, no waiting — this content was never
      behind JavaScript.

  Strategy 2 — LIVE SECTOR ITERATION (escalation)
      Only runs if Strategy 1 returns suspiciously few symbols. Drives the
      "Select Sector" dropdown one option at a time with Selenium, waits
      for the AJAX-populated results table to genuinely change (polling,
      not a blind sleep), and scrapes whatever rows appear for each sector.

  Strategy 3 — NETWORK-LOG DISCOVERY (last resort)
      Only runs if Strategy 2 also comes up short. Turns on Chrome's
      performance/network logging, triggers one sector selection, and
      inspects the captured network traffic for the underlying JSON
      endpoint the page calls. This is diagnostic rather than guaranteed —
      I can't execute PSX's JS myself to verify the exact response shape,
      so it writes everything it finds to network_discovery_debug.json for
      you to inspect, in addition to a best-effort symbol extraction.

A WORTHWHILE EXPERIMENT: the dps.psx.com.pk/company/{SYMBOL} pages also
appeared to render their content server-side when I fetched one directly
(no JS execution on my end). If that holds up, you may not need Selenium
for the per-company scrape (Section 4) either — it could become a plain
`requests` loop, which would be faster and far less flaky in CI. Worth
testing on a handful of symbols before ripping Selenium out, since I
can't 100% guarantee that from outside a real browser. Section 4 below is
left on Selenium for now since that's what you've already validated.
"""

import json
import os
import re
import sys
import time
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# ═══════════════════════ CONFIGURATION ═══════════════════════════════
TARGET_AUDITOR        = "Reanda Haroon Zakaria"
OUTPUT_ALL             = "psx_companies_auditor_list.xlsx"
OUTPUT_FILTERED        = "filtered_psx_companies_auditor_list.xlsx"
CHECKPOINT_FILE        = "scraper_checkpoint.json"

LISTING_URL            = "https://www.psx.com.pk/psx/resources-and-tools/listings/listed-companies"
COMPANY_URL_TEMPLATE   = "https://dps.psx.com.pk/company/{symbol}"

MIN_EXPECTED_SYMBOLS   = 200     # PSX has 500+ tickers; far fewer than this means a strategy broke
MAX_RETRIES            = 3       # Retry attempts per company page
REQUEST_DELAY          = 2.0     # Seconds between requests (be polite)
CHECKPOINT_INTERVAL    = 10      # Save progress every N companies
PAGE_LOAD_TIMEOUT      = 45      # Seconds before timing out a page load

UA_HEADER = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
}
# ═════════════════════════════════════════════════════════════════════


# ──────────────────────────────────────────
# 1.  DRIVER SETUP
# ──────────────────────────────────────────
def setup_driver() -> webdriver.Chrome:
    """
    Configure Chrome for GitHub Actions headless Ubuntu.
    Falls back to the system chromedriver if webdriver-manager fails.
    Performance logging is enabled so Strategy 3 (network sniffing) can
    inspect real traffic if it's ever needed.
    """
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-infobars")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument(f"--user-agent={UA_HEADER['User-Agent']}")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    try:
        print("  Setting up ChromeDriver via webdriver-manager …")
        service = Service(ChromeDriverManager().install())
        driver  = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        print(f"  webdriver-manager failed ({e}), trying system ChromeDriver …")
        driver = webdriver.Chrome(options=options)

    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return driver


# ──────────────────────────────────────────
# 2.  CHECKPOINT (RESUME SUPPORT)
# ──────────────────────────────────────────
def load_checkpoint() -> dict:
    if Path(CHECKPOINT_FILE).exists():
        with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        print(f"  Checkpoint loaded — {len(data.get('results', []))} companies already done.")
        return data
    return {"processed_urls": [], "results": []}


def save_checkpoint(processed_urls: set, results: list) -> None:
    with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
        json.dump({"processed_urls": list(processed_urls), "results": results},
                  f, ensure_ascii=False, indent=2)


# ──────────────────────────────────────────
# 3.  SYMBOL / LINK GATHERING — THREE ESCALATING STRATEGIES
# ──────────────────────────────────────────
SYMBOL_RE = re.compile(r"^[A-Z0-9]{2,12}$")


def _clean_symbol_candidates(raw_values: list) -> set:
    out = set()
    for v in raw_values:
        v = (v or "").strip()
        if not v or re.match(r"select\s+symbol", v, re.I) or re.match(r"select\s+sector", v, re.I):
            continue
        if SYMBOL_RE.match(v):
            out.add(v)
    return out


def _extract_symbols_from_dropdown(soup: BeautifulSoup) -> set:
    """
    Find the <select> whose first option reads 'Select Symbol' and pull
    every other option as a ticker symbol. Checks the option's displayed
    text first (that's what PSX uses), falling back to its `value`
    attribute in case a future markup change separates the two.
    """
    for select in soup.find_all("select"):
        options = select.find_all("option")
        if not options:
            continue
        texts = [o.get_text(strip=True) for o in options]
        if not any(re.match(r"select\s+symbol", t, re.I) for t in texts):
            continue
        values = [o.get("value", "") for o in options]
        symbols = _clean_symbol_candidates(texts)
        if len(symbols) < MIN_EXPECTED_SYMBOLS:
            symbols |= _clean_symbol_candidates(values)
        return symbols
    return set()


def _extract_sector_names(soup: BeautifulSoup) -> list:
    """Pull every option from the 'Select Sector' dropdown (used by Strategy 2)."""
    for select in soup.find_all("select"):
        options = select.find_all("option")
        texts = [o.get_text(strip=True) for o in options]
        if any(re.match(r"select\s+sector", t, re.I) for t in texts):
            return [t for t in texts if t and not re.match(r"select\s+sector", t, re.I)]
    return []


def _extract_symbols_from_result_table(soup: BeautifulSoup) -> set:
    """After a sector/symbol is selected, PSX injects a results table with
    a 'Symbols' column — pull tickers out of its first cell per row."""
    found = set()
    for table in soup.find_all("table"):
        header_text = " ".join(th.get_text(strip=True).lower() for th in table.find_all("th"))
        if "symbol" not in header_text:
            continue
        for row in table.find_all("tr"):
            cells = row.find_all("td")
            if cells:
                found |= _clean_symbol_candidates([cells[0].get_text(strip=True)])
    return found


# ---- Strategy 1: plain HTTP, no browser ----
def get_symbols_via_static_page() -> set:
    """
    The 'Select Symbol' dropdown is rendered server-side, so a plain GET
    request is enough — no Selenium, no AJAX wait, nothing JS-dependent.
    This is the fix for the root cause: v1 was trying to parse a *results
    table* that genuinely never renders without JS, when the data it
    actually needed (the full symbol list) was sitting in static HTML the
    whole time.
    """
    print("  [Strategy 1] Fetching the static symbol dropdown via plain HTTP …")
    try:
        resp = requests.get(LISTING_URL, headers=UA_HEADER, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")
        symbols = _extract_symbols_from_dropdown(soup)
        print(f"    → {len(symbols)} symbols found.")
        return symbols
    except requests.RequestException as e:
        print(f"    Strategy 1 failed: {str(e)[:100]}")
        return set()


# ---- Strategy 2: drive the live sector filter with Selenium ----
def _find_select_element(driver, placeholder_regex: str):
    for sel in driver.find_elements(By.TAG_NAME, "select"):
        opts = sel.find_elements(By.TAG_NAME, "option")
        if opts and re.match(placeholder_regex, opts[0].text.strip(), re.I):
            return sel
    return None


def get_symbols_via_sector_iteration(driver: webdriver.Chrome) -> set:
    """
    Escalation path: select each sector in turn and wait for the AJAX
    results table to genuinely refresh (polled, not a fixed sleep) before
    scraping it. This recovers data even if the static dropdown that
    Strategy 1 relies on ever goes away.
    """
    print("  [Strategy 2] Driving the live Sector filter (this is slower) …")
    try:
        driver.get(LISTING_URL)
        WebDriverWait(driver, 30).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except (TimeoutException, WebDriverException) as e:
        print(f"    Could not load listing page: {str(e)[:80]}")
        return set()

    soup = BeautifulSoup(driver.page_source, "lxml")
    sector_names = _extract_sector_names(soup)
    if not sector_names:
        print("    Could not find a 'Select Sector' dropdown — skipping Strategy 2.")
        return set()

    all_symbols = set()
    for sector in sector_names:
        try:
            before_html = driver.find_element(By.TAG_NAME, "body").text
            sector_el = _find_select_element(driver, r"select\s+sector")
            if sector_el is None:
                continue
            Select(sector_el).select_by_visible_text(sector)

            WebDriverWait(driver, 15).until(
                lambda d: d.find_element(By.TAG_NAME, "body").text != before_html
            )
            time.sleep(0.8)  # let the final re-render settle

            soup = BeautifulSoup(driver.page_source, "lxml")
            found = _extract_symbols_from_result_table(soup)
            all_symbols |= found
            print(f"    {sector[:45]:<45} +{len(found)}  (total {len(all_symbols)})")

        except (TimeoutException, StaleElementReferenceException, NoSuchElementException) as e:
            print(f"    {sector[:45]:<45} ⚠ skipped ({str(e)[:50]})")
            continue

    return all_symbols


# ---- Strategy 3: sniff network traffic for the underlying JSON endpoint ----
def get_symbols_via_network_sniffing(driver: webdriver.Chrome) -> set:
    """
    Last resort. Triggers one AJAX call and reads Chrome's own network log
    to find candidate API endpoints, instead of guessing at front-end
    markup. I can't execute PSX's JS from where I'm sitting to confirm the
    exact response shape, so treat this as diagnostic: every candidate URL
    and a best-effort symbol scrape get written to
    network_discovery_debug.json for you to check by hand.
    """
    print("  [Strategy 3] Sniffing network traffic for the data endpoint …")
    found_symbols = set()
    candidate_urls = set()

    try:
        driver.get(LISTING_URL)
        WebDriverWait(driver, 30).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )

        sector_el = _find_select_element(driver, r"select\s+sector")
        if sector_el is not None:
            options = [o.text.strip() for o in Select(sector_el).options if o.text.strip()]
            target = next((o for o in options if not re.match(r"select\s+sector", o, re.I)), None)
            if target:
                Select(sector_el).select_by_visible_text(target)
                time.sleep(4)

        for entry in driver.get_log("performance"):
            try:
                msg = json.loads(entry["message"])["message"]
                if msg.get("method") == "Network.responseReceived":
                    resp = msg["params"]["response"]
                    url, mime = resp.get("url", ""), resp.get("mimeType", "")
                    if "json" in mime or any(k in url.lower() for k in ("ajax", "api", "company", "symbol")):
                        candidate_urls.add(url)
            except (KeyError, json.JSONDecodeError):
                continue

        cookies = {c["name"]: c["value"] for c in driver.get_cookies()}
        debug_entries = []
        for url in candidate_urls:
            try:
                resp = requests.get(url, cookies=cookies, headers=UA_HEADER, timeout=10)
                snippet = resp.text[:500]
                debug_entries.append({"url": url, "status": resp.status_code, "snippet": snippet})
                found_symbols |= _clean_symbol_candidates(re.findall(r'"([A-Z0-9]{2,12})"', resp.text))
            except requests.RequestException as e:
                debug_entries.append({"url": url, "error": str(e)[:100]})

        with open("network_discovery_debug.json", "w", encoding="utf-8") as f:
            json.dump({"candidate_endpoints": debug_entries}, f, indent=2, ensure_ascii=False)
        print(f"    Logged {len(candidate_urls)} candidate endpoint(s) → network_discovery_debug.json")
        print(f"    Best-effort symbol extraction: {len(found_symbols)} found.")

    except WebDriverException as e:
        print(f"    Network sniffing failed: {str(e)[:100]}")

    return found_symbols


# ---- Orchestrator: escalate only as far as needed ----
def get_all_company_links(driver: webdriver.Chrome) -> list:
    print(f"\n  Discovering PSX ticker symbols from: {LISTING_URL}")

    symbols = get_symbols_via_static_page()

    if len(symbols) < MIN_EXPECTED_SYMBOLS:
        print(f"  Only {len(symbols)} symbols so far (expected {MIN_EXPECTED_SYMBOLS}+) — escalating to Strategy 2.")
        symbols |= get_symbols_via_sector_iteration(driver)

    if len(symbols) < MIN_EXPECTED_SYMBOLS:
        print(f"  Still only {len(symbols)} symbols — escalating to Strategy 3.")
        symbols |= get_symbols_via_network_sniffing(driver)

    if not symbols:
        print("  ERROR: All three strategies failed to find any symbols.")
        return []

    links = [COMPANY_URL_TEMPLATE.format(symbol=s) for s in sorted(symbols)]
    print(f"  Final symbol count: {len(symbols)} → {len(links)} company URLs built "
          f"(template: {COMPANY_URL_TEMPLATE}).")
    return links


# ──────────────────────────────────────────
# 4.  COMPANY PAGE — EXTRACT AUDITOR
#     (unchanged from v1 — these 5 fallback strategies already handle a
#      plain label/value layout like dps.psx.com.pk's "AUDITOR" field;
#      only the URLs feeding into this section changed.)
# ──────────────────────────────────────────
def _strategy_table_rows(soup: BeautifulSoup) -> str | None:
    for row in soup.find_all("tr"):
        cells = row.find_all(["td", "th"])
        for i, cell in enumerate(cells[:-1]):
            if re.search(r"\bauditor", cell.get_text(strip=True), re.I):
                value = cells[i + 1].get_text(strip=True)
                if 3 < len(value) < 300:
                    return value
    return None


def _strategy_definition_lists(soup: BeautifulSoup) -> str | None:
    for tag in ["dt", "th", "strong", "b", "label"]:
        for el in soup.find_all(tag):
            if re.search(r"\bauditor", el.get_text(strip=True), re.I):
                sibling = el.find_next_sibling(["dd", "td", "span", "p", "div"])
                if sibling:
                    v = sibling.get_text(strip=True)
                    if 3 < len(v) < 300:
                        return v
    return None


def _strategy_inline_colon(soup: BeautifulSoup) -> str | None:
    for el in soup.find_all(["p", "div", "span", "li"]):
        if len(el.find_all()) > 8:
            continue
        text  = el.get_text(strip=True)
        match = re.match(
            r"(?:Statutory\s+)?(?:External\s+)?Auditors?\s*:\s*(.+)",
            text, re.I
        )
        if match:
            v = match.group(1).strip()
            if 3 < len(v) < 300:
                return v
    return None


def _strategy_script_json(soup: BeautifulSoup) -> str | None:
    def _dig(obj, depth=0):
        if depth > 6:
            return None
        if isinstance(obj, dict):
            for k, v in obj.items():
                if re.search(r"\bauditor", str(k), re.I):
                    return str(v)[:200]
                found = _dig(v, depth + 1)
                if found:
                    return found
        elif isinstance(obj, list):
            for item in obj:
                found = _dig(item, depth + 1)
                if found:
                    return found
        return None

    for script in soup.find_all("script"):
        content = script.string or ""
        if "auditor" not in content.lower():
            continue
        m = re.search(r'"[Aa]uditor[^"]*"\s*:\s*"([^"]{3,200})"', content)
        if m:
            return m.group(1)
        for candidate in re.findall(r"\{[^{}]{20,}\}", content):
            try:
                obj   = json.loads(candidate)
                found = _dig(obj)
                if found:
                    return found
            except json.JSONDecodeError:
                pass
    return None


def _strategy_full_text_regex(soup: BeautifulSoup) -> str | None:
    page_text = soup.get_text(separator="\n")
    patterns  = [
        r"(?:Statutory\s+)?Auditors?\s*:\s*([^\n]{3,200})",
        r"(?:External\s+)?Auditors?\s*:\s*([^\n]{3,200})",
        r"Auditors?\s*\n\s*([^\n]{3,200})",
    ]
    for pat in patterns:
        m = re.search(pat, page_text, re.I)
        if m:
            result = re.sub(r"\s+", " ", m.group(1)).strip()
            if re.search(r"[a-zA-Z]{3}", result) and len(result) < 200:
                return result
    return None


def extract_auditor(soup: BeautifulSoup) -> str:
    for strategy in [
        _strategy_table_rows,
        _strategy_definition_lists,
        _strategy_inline_colon,
        _strategy_script_json,
        _strategy_full_text_regex,
    ]:
        result = strategy(soup)
        if result:
            return result.strip()
    return "Not Found"


def scrape_company(driver: webdriver.Chrome, url: str) -> str:
    for attempt in range(MAX_RETRIES):
        try:
            driver.get(url)
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            time.sleep(1.5)
            soup = BeautifulSoup(driver.page_source, "lxml")
            return extract_auditor(soup)

        except TimeoutException:
            print(f"\n    [Timeout — attempt {attempt+1}/{MAX_RETRIES}]", end="")
            time.sleep(4 * (attempt + 1))
        except WebDriverException as exc:
            print(f"\n    [WebDriver error — attempt {attempt+1}: {str(exc)[:60]}]", end="")
            time.sleep(4 * (attempt + 1))
        except Exception as exc:
            print(f"\n    [Unexpected error: {str(exc)[:60]}]", end="")
            break

    return "Error: scrape failed"


# ──────────────────────────────────────────
# 5.  MAIN
# ──────────────────────────────────────────
def main():
    print("=" * 64)
    print("  PSX Auditor Scraper — GitHub Actions Edition (v2)")
    print(f"  Target auditor : {TARGET_AUDITOR}")
    print(f"  Symbol source  : {LISTING_URL}")
    print(f"  Company URL    : {COMPANY_URL_TEMPLATE}")
    print("=" * 64)

    checkpoint     = load_checkpoint()
    processed_urls = set(checkpoint["processed_urls"])
    results        = list(checkpoint["results"])

    driver = setup_driver()

    try:
        all_links = get_all_company_links(driver)
        if not all_links:
            print("\nFATAL: No company links found. Exiting.")
            sys.exit(1)

        remaining  = [u for u in all_links if u not in processed_urls]
        done_count = len(processed_urls)
        total      = len(all_links)

        print(f"\n  Total: {total}  |  Already done: {done_count}  |  Remaining: {len(remaining)}\n")

        for idx, url in enumerate(remaining, start=1):
            symbol  = url.rstrip("/").split("/")[-1]
            display = f"[{done_count + idx}/{total}]  {symbol:<14}"
            print(display, end="→ ", flush=True)

            auditor = scrape_company(driver, url)
            truncated = auditor[:65] + "…" if len(auditor) > 65 else auditor
            print(truncated)

            results.append({
                "Company Code"     : symbol,
                "URL"              : url,
                "Auditor"          : auditor,
                "Is Target Auditor": TARGET_AUDITOR.lower() in auditor.lower(),
            })
            processed_urls.add(url)

            if idx % CHECKPOINT_INTERVAL == 0:
                save_checkpoint(processed_urls, results)
                pd.DataFrame(results).to_excel(OUTPUT_ALL, index=False)
                print(f"  ── Checkpoint saved ({done_count + idx} / {total}) ──")

            time.sleep(REQUEST_DELAY)

        df = pd.DataFrame(results)
        df.to_excel(OUTPUT_ALL, index=False)
        print(f"\n✅  Full data saved → {OUTPUT_ALL}  ({len(df)} companies)")

        mask      = df["Auditor"].str.contains(TARGET_AUDITOR, case=False, na=False)
        target_df = df[mask].copy()

        print(f"\n{'='*64}")
        if not target_df.empty:
            print(f"🎯  {len(target_df)} companies audited by '{TARGET_AUDITOR}':\n")
            print(target_df[["Company Code", "Auditor"]].to_string(index=False))
            target_df.to_excel(OUTPUT_FILTERED, index=False)
            print(f"\n✅  Filtered file saved → {OUTPUT_FILTERED}")
        else:
            print(f"⚠️   No companies found with auditor matching '{TARGET_AUDITOR}'.")
            print("     Possible reasons:")
            print("       1. Auditor name spelled differently in the website data")
            print("       2. Company pages loaded but auditor field uses an unseen format")
            print("       3. This auditor genuinely has no PSX-listed clients")
            print(f"\n     'Not Found' count : {(df['Auditor'] == 'Not Found').sum()}")
            print(f"     Error count       : {df['Auditor'].str.startswith('Error').sum()}")

            pd.DataFrame({
                "Status"                  : ["No match"],
                "Target Auditor Searched" : [TARGET_AUDITOR],
                "Total Companies Scraped" : [len(df)],
                "Auditor Found Count"     : [(df["Auditor"] != "Not Found").sum()],
                "Not Found Count"         : [(df["Auditor"] == "Not Found").sum()],
            }).to_excel(OUTPUT_FILTERED, index=False)

        print(f"\n{'─'*64}")
        print(f"  Total scraped     : {len(df)}")
        print(f"  Auditor found     : {(df['Auditor'] != 'Not Found').sum()}")
        print(f"  Not found         : {(df['Auditor'] == 'Not Found').sum()}")
        print(f"  Errors            : {df['Auditor'].str.startswith('Error').sum()}")
        print(f"  Target matches    : {mask.sum()}")
        print(f"{'─'*64}\n")

        if Path(CHECKPOINT_FILE).exists():
            os.remove(CHECKPOINT_FILE)

    finally:
        driver.quit()
        print("  Browser closed. Scraper complete.")


if __name__ == "__main__":
    main()
