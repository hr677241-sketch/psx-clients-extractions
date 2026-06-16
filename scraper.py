"""
PSX Auditor Scraper — GitHub Actions Edition
=============================================
Scrapes every listed company on PSX and finds their statutory auditor.
Saves all results + a filtered sheet for the TARGET_AUDITOR.

Key improvements over v1:
  - Fixed Selenium 4 By.TAG_NAME usage (original had "tag name" string bug)
  - Checkpoint / resume — picks up exactly where it left off after a crash
  - Retry logic with back-off on every company fetch
  - 5 extraction strategies for the auditor field
  - Pagination support on the main listing page
  - Incremental Excel saves every 10 companies
  - Proper DataTable "show all" triggering via JS
  - JSON-LD / embedded JSON scanning
  - Human-like delays + user-agent to avoid rate-limiting
"""

import json
import os
import re
import sys
import time
from pathlib import Path

import pandas as pd
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
from selenium.webdriver.common.by import By          # ← FIX: was missing in original
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# ═══════════════════════ CONFIGURATION ═══════════════════════════════
TARGET_AUDITOR       = "Reanda Haroon Zakaria"
OUTPUT_ALL           = "psx_companies_auditor_list.xlsx"
OUTPUT_FILTERED      = "filtered_psx_companies_auditor_list.xlsx"
CHECKPOINT_FILE      = "scraper_checkpoint.json"
BASE_URL             = "https://www.psx.com.pk/psx/resources-and-tools/listings/listed-companies"

MAX_RETRIES          = 3        # Retry attempts per company page
REQUEST_DELAY        = 2.0      # Seconds between requests (be polite)
CHECKPOINT_INTERVAL  = 10       # Save progress every N companies
PAGE_LOAD_TIMEOUT    = 45       # Seconds before timing out a page load
# ═════════════════════════════════════════════════════════════════════


# ──────────────────────────────────────────
# 1.  DRIVER SETUP
# ──────────────────────────────────────────
def setup_driver() -> webdriver.Chrome:
    """
    Configure Chrome for GitHub Actions headless Ubuntu.
    Falls back to the system chromedriver if webdriver-manager fails.
    """
    options = Options()
    options.add_argument("--headless=new")           # New headless (Chrome 112+)
    options.add_argument("--no-sandbox")             # Required in CI
    options.add_argument("--disable-dev-shm-usage")  # Prevents /dev/shm OOM
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-infobars")
    options.add_argument("--ignore-certificate-errors")
    # Spoof a real browser so the site doesn't block headless bots
    options.add_argument(
        "--user-agent=Mozilla/5.0 (X11; Linux x86_64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

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
# 3.  LISTING PAGE — GET ALL COMPANY LINKS
# ──────────────────────────────────────────
def _scroll_fully(driver: webdriver.Chrome, pause: float = 1.2) -> None:
    """Scroll to the bottom to trigger any lazy-loaded content."""
    last_h = driver.execute_script("return document.body.scrollHeight")
    for _ in range(10):                         # Max 10 scroll passes
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(pause)
        new_h = driver.execute_script("return document.body.scrollHeight")
        if new_h == last_h:
            break
        last_h = new_h


def _try_show_all(driver: webdriver.Chrome) -> None:
    """
    PSX uses jQuery DataTables for the company list.
    Attempt to set page-length to -1 (show all rows) via JS,
    then look for "Show All" buttons as a fallback.
    """
    try:
        driver.execute_script("""
            try {
                // DataTables API
                var tables = $.fn.dataTable.tables(true);
                tables.forEach(function(t) {
                    $(t).DataTable().page.len(-1).draw();
                });
            } catch(e) {}
            // Also change any <select> that controls page size
            document.querySelectorAll('select').forEach(function(s) {
                if (['-1','All','100','500'].some(v => {
                    for (var o of s.options) if (o.value === v) return true;
                    return false;
                })) {
                    s.value = '-1';
                    s.dispatchEvent(new Event('change', {bubbles: true}));
                }
            });
        """)
        time.sleep(2)
    except Exception:
        pass

    # Button fallback
    for xpath in [
        "//button[normalize-space()='Show All']",
        "//a[normalize-space()='Show All']",
        "//*[contains(@class,'show-all')]",
    ]:
        try:
            btn = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            btn.click()
            time.sleep(2)
        except (TimeoutException, NoSuchElementException):
            pass


def _extract_links_from_soup(soup: BeautifulSoup) -> set:
    """Pull every /listed-companies/{SYMBOL} href from the page."""
    links = set()
    path_patterns = [
        "/psx/resources-and-tools/listings/listed-companies/",
        "/listed-companies/",
    ]
    skip_suffixes = ("listed-companies", "listed-companies/")

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href or href.startswith("#") or "javascript" in href:
            continue
        if any(p in href for p in path_patterns):
            # Skip the base listing URL itself
            if href.rstrip("/").endswith(skip_suffixes):
                continue
            full = f"https://www.psx.com.pk{href}" if href.startswith("/") else href
            links.add(full)
    return links


def get_all_company_links(driver: webdriver.Chrome) -> list:
    """
    Return every company detail-page URL from the PSX listing.
    Handles DataTable "show all" and paginated navigation.
    """
    print(f"\n  Loading listing page: {BASE_URL}")

    for attempt in range(MAX_RETRIES):
        try:
            driver.get(BASE_URL)
            # Wait for the page to fully render
            WebDriverWait(driver, 30).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            time.sleep(4)

            _scroll_fully(driver)
            _try_show_all(driver)
            _scroll_fully(driver)          # Scroll again after expanding

            soup  = BeautifulSoup(driver.page_source, "lxml")
            links = _extract_links_from_soup(soup)

            if not links:
                print("  No links found on first pass — trying row-by-row scan …")
                for row in soup.find_all("tr"):
                    for a in row.find_all("a", href=True):
                        href = a["href"].strip()
                        if href and "psx.com.pk" in href:
                            links.add(href)

            # ── Pagination: follow "Next" buttons ──
            all_links  = set(links)
            page_count = 1

            while True:
                try:
                    next_btn = WebDriverWait(driver, 4).until(
                        EC.element_to_be_clickable((By.XPATH,
                            "//a[contains(@class,'next') "
                            "or normalize-space()='Next' "
                            "or @aria-label='Next page']"
                        ))
                    )
                    # Stop if the button is disabled
                    cls = next_btn.get_attribute("class") or ""
                    if "disabled" in cls:
                        break
                    next_btn.click()
                    time.sleep(3)
                    page_soup  = BeautifulSoup(driver.page_source, "lxml")
                    page_links = _extract_links_from_soup(page_soup)
                    new = page_links - all_links
                    if not new:
                        break
                    all_links.update(new)
                    page_count += 1
                    print(f"  Page {page_count}: +{len(new)} links  (total {len(all_links)})")
                except (TimeoutException, NoSuchElementException, StaleElementReferenceException):
                    break

            result = list(all_links)
            print(f"  Found {len(result)} company links total.")
            return result

        except Exception as e:
            wait = 5 * (attempt + 1)
            print(f"  Attempt {attempt+1} failed: {str(e)[:80]}  — retrying in {wait}s …")
            time.sleep(wait)

    print("  ERROR: Could not retrieve company links after all retries.")
    return []


# ──────────────────────────────────────────
# 4.  COMPANY PAGE — EXTRACT AUDITOR
# ──────────────────────────────────────────
def _strategy_table_rows(soup: BeautifulSoup) -> str | None:
    """Look for <tr> with a cell labelled 'auditor' followed by the value cell."""
    for row in soup.find_all("tr"):
        cells = row.find_all(["td", "th"])
        for i, cell in enumerate(cells[:-1]):
            if re.search(r"\bauditor", cell.get_text(strip=True), re.I):
                value = cells[i + 1].get_text(strip=True)
                if 3 < len(value) < 300:
                    return value
    return None


def _strategy_definition_lists(soup: BeautifulSoup) -> str | None:
    """<dt>Auditor</dt><dd>Value</dd>  or  <label>Auditor</label> patterns."""
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
    """Elements whose text matches 'Auditor(s): Value'."""
    for el in soup.find_all(["p", "div", "span", "li"]):
        if len(el.find_all()) > 8:          # Skip large containers
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
    """Scan <script> tags for JSON containing an 'auditor' key."""
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
        # Try extracting simple "key": "value" pairs first (fast)
        m = re.search(r'"[Aa]uditor[^"]*"\s*:\s*"([^"]{3,200})"', content)
        if m:
            return m.group(1)
        # Try full JSON parse
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
    """Last resort: regex the entire visible text."""
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
            # Sanity check: must contain real alphabetic content
            if re.search(r"[a-zA-Z]{3}", result) and len(result) < 200:
                return result
    return None


def extract_auditor(soup: BeautifulSoup) -> str:
    """
    Try 5 strategies in order of reliability.
    Returns auditor name or 'Not Found'.
    """
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
    """Navigate to a company page and return its auditor string (with retries)."""
    for attempt in range(MAX_RETRIES):
        try:
            driver.get(url)
            # ── FIX: correct Selenium 4 By usage ──────────────────
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
    print("  PSX Auditor Scraper — GitHub Actions Edition")
    print(f"  Target auditor : {TARGET_AUDITOR}")
    print(f"  Listing URL    : {BASE_URL}")
    print("=" * 64)

    # ── Load checkpoint (resume support) ──
    checkpoint     = load_checkpoint()
    processed_urls = set(checkpoint["processed_urls"])
    results        = list(checkpoint["results"])

    driver = setup_driver()

    try:
        # ── Step 1: Collect all company URLs ──────────────────────
        all_links = get_all_company_links(driver)
        if not all_links:
            print("\nFATAL: No company links found. Exiting.")
            sys.exit(1)

        remaining  = [u for u in all_links if u not in processed_urls]
        done_count = len(processed_urls)
        total      = len(all_links)

        print(f"\n  Total: {total}  |  Already done: {done_count}  |  Remaining: {len(remaining)}\n")

        # ── Step 2: Scrape each company page ──────────────────────
        for idx, url in enumerate(remaining, start=1):
            symbol  = url.rstrip("/").split("/")[-1]
            display = f"[{done_count + idx}/{total}]  {symbol:<14}"
            print(display, end="→ ", flush=True)

            auditor = scrape_company(driver, url)
            truncated = auditor[:65] + "…" if len(auditor) > 65 else auditor
            print(truncated)

            results.append({
                "Company Code"    : symbol,
                "URL"             : url,
                "Auditor"         : auditor,
                "Is Target Auditor": TARGET_AUDITOR.lower() in auditor.lower(),
            })
            processed_urls.add(url)

            # ── Periodic checkpoint + intermediate save ────────────
            if idx % CHECKPOINT_INTERVAL == 0:
                save_checkpoint(processed_urls, results)
                pd.DataFrame(results).to_excel(OUTPUT_ALL, index=False)
                print(f"  ── Checkpoint saved ({done_count + idx} / {total}) ──")

            time.sleep(REQUEST_DELAY)

        # ── Step 3: Final save ────────────────────────────────────
        df = pd.DataFrame(results)
        df.to_excel(OUTPUT_ALL, index=False)
        print(f"\n✅  Full data saved → {OUTPUT_ALL}  ({len(df)} companies)")

        # ── Step 4: Filtered output ───────────────────────────────
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

            # Write a diagnostic file so the GitHub artifact isn't empty
            pd.DataFrame({
                "Status"                  : ["No match"],
                "Target Auditor Searched" : [TARGET_AUDITOR],
                "Total Companies Scraped" : [len(df)],
                "Auditor Found Count"     : [(df["Auditor"] != "Not Found").sum()],
                "Not Found Count"         : [(df["Auditor"] == "Not Found").sum()],
            }).to_excel(OUTPUT_FILTERED, index=False)

        # ── Step 5: Summary ───────────────────────────────────────
        print(f"\n{'─'*64}")
        print(f"  Total scraped     : {len(df)}")
        print(f"  Auditor found     : {(df['Auditor'] != 'Not Found').sum()}")
        print(f"  Not found         : {(df['Auditor'] == 'Not Found').sum()}")
        print(f"  Errors            : {df['Auditor'].str.startswith('Error').sum()}")
        print(f"  Target matches    : {mask.sum()}")
        print(f"{'─'*64}\n")

        # Remove checkpoint on clean completion
        if Path(CHECKPOINT_FILE).exists():
            os.remove(CHECKPOINT_FILE)

    finally:
        driver.quit()
        print("  Browser closed. Scraper complete.")


if __name__ == "__main__":
    main()
