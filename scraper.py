"""
PSX Auditor Scraper — GitHub Actions Edition (V3 - Escalating Strategies)
==========================================================================
FIXED: Handles JavaScript/AJAX-loaded DataTables using 3 escalating strategies.
Strategy 1: DataTable API hook (injects JS to read table data directly).
Strategy 2: Selenium explicit wait + XPath on rendered rows.
Strategy 3: Full-page BeautifulSoup scan (after JS renders everything).
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
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# ═══════════════════════ CONFIGURATION ═══════════════════════════════
TARGET_AUDITOR       = "Reanda Haroon Zakaria"
OUTPUT_ALL           = "psx_companies_auditor_list.xlsx"
OUTPUT_FILTERED      = "filtered_psx_companies_auditor_list.xlsx"
CHECKPOINT_FILE      = "scraper_checkpoint.json"
BASE_URL             = "https://www.psx.com.pk/psx/resources-and-tools/listings/listed-companies"

MAX_RETRIES          = 3
REQUEST_DELAY        = 2.0
CHECKPOINT_INTERVAL  = 10
PAGE_LOAD_TIMEOUT    = 45
# ═════════════════════════════════════════════════════════════════════


# ──────────────────────────────────────────
# 1. DRIVER SETUP
# ──────────────────────────────────────────
def setup_driver() -> webdriver.Chrome:
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-infobars")
    options.add_argument("--ignore-certificate-errors")
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
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        print(f"  webdriver-manager failed ({e}), trying system ChromeDriver …")
        driver = webdriver.Chrome(options=options)

    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return driver


# ──────────────────────────────────────────
# 2. CHECKPOINT
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
# 3. LINK-GATHERING — 3 ESCALATING STRATEGIES
# ──────────────────────────────────────────

def _strategy_1_direct_datatable_api(driver: webdriver.Chrome) -> set:
    """
    STRATEGY 1: Directly hook into the DataTable API via JavaScript.
    This reads the table data from the DataTable's internal memory (fastest & most reliable).
    """
    print("  [Strategy 1] Attempting DataTable API hook via JS...")
    try:
        # Wait for DataTable to exist
        WebDriverWait(driver, 10).until(
            lambda d: d.execute_script("return typeof $.fn.DataTable !== 'undefined'")
        )
        
        js_code = """
        let links = [];
        try {
            // Find the first DataTable on the page
            let table = $('table').DataTable();
            // Get all rows data
            let data = table.rows().data();
            for (let i = 0; i < data.length; i++) {
                let row = data[i];
                // PSX DataTable usually has the symbol/company name in the first column
                // We search for any <a> tag within the rendered row.
                let rowNode = table.row(i).node();
                let anchor = $(rowNode).find('a').filter(function() {
                    return $(this).attr('href') && $(this).attr('href').includes('listed-companies');
                }).first();
                if (anchor.length > 0) {
                    let href = anchor.attr('href');
                    if (href.startsWith('/')) href = 'https://www.psx.com.pk' + href;
                    links.push(href);
                }
            }
        } catch(e) {
            return [];
        }
        return links;
        """
        
        result = driver.execute_script(js_code)
        if result and len(result) > 0:
            print(f"  [Strategy 1] Successfully extracted {len(result)} links via DataTable API.")
            return set(result)
        else:
            print("  [Strategy 1] DataTable API returned 0 links. Escalating...")
            return set()
    except Exception as e:
        print(f"  [Strategy 1] Failed: {str(e)[:80]}. Escalating...")
        return set()


def _strategy_2_selenium_explicit_wait(driver: webdriver.Chrome) -> set:
    """
    STRATEGY 2: Wait for the table to render, then use Selenium XPath.
    This waits for the AJAX to finish populating <tbody> with rows.
    """
    print("  [Strategy 2] Waiting for table rows via Selenium explicit wait...")
    try:
        # Wait for at least 2 rows to appear in the tbody
        WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.XPATH, "//table//tbody/tr/td/a[contains(@href, 'listed-companies')]"))
        )
        time.sleep(2)  # Extra buffer for lazy-load
        
        elements = driver.find_elements(By.XPATH, "//table//tbody/tr/td/a[contains(@href, 'listed-companies')]")
        
        links = set()
        for el in elements:
            href = el.get_attribute("href")
            if href:
                if href.startswith("/"):
                    href = "https://www.psx.com.pk" + href
                if "listed-companies" in href:
                    links.add(href)
        
        if links:
            print(f"  [Strategy 2] Found {len(links)} links via Selenium XPath.")
            return links
        else:
            print("  [Strategy 2] XPath found elements but 0 links extracted. Escalating...")
            return set()
    except TimeoutException:
        print("  [Strategy 2] Timeout waiting for table rows. Escalating...")
        return set()
    except Exception as e:
        print(f"  [Strategy 2] Failed: {str(e)[:80]}. Escalating...")
        return set()


def _strategy_3_full_beautifulsoup_scan(driver: webdriver.Chrome) -> set:
    """
    STRATEGY 3: Full-page scan using BeautifulSoup AFTER ensuring JavaScript has run.
    Also handles pagination by clicking "Next" buttons.
    """
    print("  [Strategy 3] Performing full BeautifulSoup scan with pagination...")
    
    all_links = set()
    
    # Ensure we are on page 1
    try:
        # Try to find and click "Show All" if it exists
        show_all = driver.find_element(By.XPATH, "//*[contains(text(),'Show All')]")
        show_all.click()
        time.sleep(3)
    except:
        pass
    
    current_page = 1
    max_pages = 20  # Safety limit to avoid infinite loop
    
    while current_page <= max_pages:
        try:
            # Wait for body and scroll
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            
            soup = BeautifulSoup(driver.page_source, "lxml")
            
            # Find all links containing 'listed-companies' but skip the listing page itself
            for a in soup.find_all("a", href=True):
                href = a["href"].strip()
                if not href or href.startswith("#") or "javascript" in href:
                    continue
                if "listed-companies" in href and "resources-and-tools" in href:
                    full = f"https://www.psx.com.pk{href}" if href.startswith("/") else href
                    if full.rstrip("/") not in [BASE_URL, BASE_URL + "/"]:
                        all_links.add(full)
            
            print(f"  [Strategy 3] Page {current_page}: Found {len(all_links)} links so far.")
            
            # Try to click the "Next" pagination button
            try:
                next_btn = driver.find_element(By.XPATH, "//a[contains(@class,'next') or contains(text(),'Next')]")
                if "disabled" in (next_btn.get_attribute("class") or ""):
                    break
                next_btn.click()
                time.sleep(3)
                current_page += 1
            except NoSuchElementException:
                break
            except StaleElementReferenceException:
                break
                
        except Exception as e:
            print(f"  [Strategy 3] Error on page {current_page}: {str(e)[:80]}")
            break
    
    if all_links:
        print(f"  [Strategy 3] Total unique links found: {len(all_links)}")
    else:
        print("  [Strategy 3] No links found. PSX structure might have changed drastically.")
    
    return all_links


def get_all_company_links(driver: webdriver.Chrome) -> list:
    """
    Main orchestrator for the 3 escalating strategies.
    Returns a list of unique company detail page URLs.
    """
    print(f"\n  Loading listing page: {BASE_URL}")
    
    for attempt in range(MAX_RETRIES):
        try:
            # Navigate to the page
            driver.get(BASE_URL)
            WebDriverWait(driver, 30).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            time.sleep(4)  # Initial JS warm-up
            
            # -----------------------------------------------------------------
            # STRATEGY 1: Direct DataTable API
            # -----------------------------------------------------------------
            links = _strategy_1_direct_datatable_api(driver)
            if links:
                return list(links)
            
            # -----------------------------------------------------------------
            # STRATEGY 2: Selenium Explicit Wait + XPath
            # -----------------------------------------------------------------
            links = _strategy_2_selenium_explicit_wait(driver)
            if links:
                return list(links)
            
            # -----------------------------------------------------------------
            # STRATEGY 3: Full BeautifulSoup Scan with Pagination
            # -----------------------------------------------------------------
            links = _strategy_3_full_beautifulsoup_scan(driver)
            if links:
                return list(links)
            
            # If all strategies fail, retry the whole process
            print(f"  All strategies failed on attempt {attempt+1}. Retrying...")
            
        except Exception as e:
            wait = 5 * (attempt + 1)
            print(f"  Attempt {attempt+1} failed: {str(e)[:80]}  — retrying in {wait}s …")
            time.sleep(wait)
    
    print("  FATAL: All 3 strategies failed after max retries.")
    return []


# ──────────────────────────────────────────
# 4. COMPANY PAGE — EXTRACT AUDITOR
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
        text = el.get_text(strip=True)
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
                obj = json.loads(candidate)
                found = _dig(obj)
                if found:
                    return found
            except json.JSONDecodeError:
                pass
    return None


def _strategy_full_text_regex(soup: BeautifulSoup) -> str | None:
    page_text = soup.get_text(separator="\n")
    patterns = [
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
# 5. MAIN
# ──────────────────────────────────────────
def main():
    print("=" * 64)
    print("  PSX Auditor Scraper — V3 (Escalating Strategies)")
    print(f"  Target auditor : {TARGET_AUDITOR}")
    print(f"  Listing URL    : {BASE_URL}")
    print("=" * 64)

    checkpoint = load_checkpoint()
    processed_urls = set(checkpoint["processed_urls"])
    results = list(checkpoint["results"])

    driver = setup_driver()

    try:
        all_links = get_all_company_links(driver)
        if not all_links:
            print("\nFATAL: No company links found after all strategies. Exiting.")
            sys.exit(1)

        remaining = [u for u in all_links if u not in processed_urls]
        done_count = len(processed_urls)
        total = len(all_links)

        print(f"\n  Total: {total}  |  Already done: {done_count}  |  Remaining: {len(remaining)}\n")

        for idx, url in enumerate(remaining, start=1):
            symbol = url.rstrip("/").split("/")[-1]
            display = f"[{done_count + idx}/{total}]  {symbol:<14}"
            print(display, end="→ ", flush=True)

            auditor = scrape_company(driver, url)
            truncated = auditor[:65] + "…" if len(auditor) > 65 else auditor
            print(truncated)

            results.append({
                "Company Code": symbol,
                "URL": url,
                "Auditor": auditor,
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

        mask = df["Auditor"].str.contains(TARGET_AUDITOR, case=False, na=False)
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
                "Status": ["No match"],
                "Target Auditor Searched": [TARGET_AUDITOR],
                "Total Companies Scraped": [len(df)],
                "Auditor Found Count": [(df["Auditor"] != "Not Found").sum()],
                "Not Found Count": [(df["Auditor"] == "Not Found").sum()],
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
