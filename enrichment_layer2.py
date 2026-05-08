"""
Layer 2: Website Enrichment Pass
Takes the CSV output from BrokerageScraper (or any business scraper)
and enriches each record by visiting the business website.

Extracts:
- Email addresses
- Contact person name
- Employee/size signals
- Years in business
- Social media links (LinkedIn, Facebook)
- Primary services

Requirements:
pip install requests beautifulsoup4 pandas lxml anthropic

Usage:
python enrichment_layer2.py --input canadian_brokerages_scraped.csv --output enriched_brokerages.csv
python enrichment_layer2.py --input my_prospects.csv --output enriched_prospects.csv --api-key YOUR_KEY
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
import json
import argparse
import os
from typing import Optional
from anthropic import Anthropic

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

DEFAULT_INPUT  = "canadian_brokerages_scraped.csv"
DEFAULT_OUTPUT = "enriched_brokerages.csv"

REQUEST_TIMEOUT   = 12       # seconds per page request
DELAY_BETWEEN     = 1.5      # seconds between website visits (be polite)
MAX_TEXT_CHARS    = 6000     # chars of page text sent to Claude
USE_AI_ENRICHMENT = True     # set False to skip Claude and use regex only

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

# Pages to visit on each site, in priority order
TARGET_PATHS = [
    "/contact",
    "/contact-us",
    "/about",
    "/about-us",
    "/our-team",
    "/team",
    "",           # homepage last
]

# ─────────────────────────────────────────────
# REGEX HELPERS
# ─────────────────────────────────────────────

EMAIL_RE    = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
LINKEDIN_RE = re.compile(r"https?://(?:www\.)?linkedin\.com/(?:company|in)/[^\s\"'<>]+")
FACEBOOK_RE = re.compile(r"https?://(?:www\.)?facebook\.com/[^\s\"'<>]+")
YEAR_RE     = re.compile(r"\b(19[5-9]\d|20[0-2]\d)\b")
SIZE_RE     = re.compile(
    r"team\s+of\s+(\d+)"
    r"|(\d+)\s+(?:employees?|staff|brokers?|advisors?|professionals?)"
    r"|(\d+)\s+(?:locations?|offices?)",
    re.IGNORECASE
)

# ─────────────────────────────────────────────
# WEBSITE FETCHER
# ─────────────────────────────────────────────

def normalize_url(url: str) -> str:
    """Ensure URL has a scheme."""
    if not url:
        return ""
    url = url.strip().rstrip("/")
    if not url.startswith("http"):
        url = "https://" + url
    return url


def fetch_page(url: str) -> Optional[BeautifulSoup]:
    """Fetch a URL and return a BeautifulSoup object, or None on failure."""
    try:
        response = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        return BeautifulSoup(response.content, "lxml")
    except Exception:
        return None


def get_best_soup(base_url: str) -> tuple[Optional[BeautifulSoup], str]:
    """
    Try target paths in order, return the first successful soup + the URL used.
    Prioritises /contact and /about pages over homepage.
    """
    for path in TARGET_PATHS:
        url = base_url + path
        soup = fetch_page(url)
        if soup:
            return soup, url
        time.sleep(0.3)
    return None, ""


# ─────────────────────────────────────────────
# REGEX-BASED EXTRACTORS
# ─────────────────────────────────────────────

def extract_emails(soup: BeautifulSoup, raw_html: str) -> list[str]:
    """Extract all email addresses from page."""
    emails = set()

    # From mailto links
    for tag in soup.find_all("a", href=re.compile(r"^mailto:", re.I)):
        addr = tag["href"].replace("mailto:", "").split("?")[0].strip()
        if addr:
            emails.add(addr.lower())

    # From raw HTML text (catches obfuscated-ish patterns)
    for match in EMAIL_RE.findall(raw_html):
        emails.add(match.lower())

    # Filter out common junk
    junk = {"example.com", "youremail.com", "domain.com", "email.com"}
    emails = {e for e in emails if not any(j in e for j in junk)}
    return sorted(emails)


def extract_social(raw_html: str) -> dict:
    """Extract LinkedIn and Facebook URLs."""
    linkedin = LINKEDIN_RE.search(raw_html)
    facebook = FACEBOOK_RE.search(raw_html)
    return {
        "linkedin": linkedin.group(0) if linkedin else "",
        "facebook": facebook.group(0) if facebook else "",
    }


def extract_year_founded(text: str) -> str:
    """Look for founding year mentions."""
    # Patterns like "since 1987", "founded in 2002", "established 1995"
    patterns = [
        r"(?:since|founded|established|serving|incorporated|in business since)\s+(\d{4})",
        r"(\d{4})\s+(?:to present|–\s*present|and counting)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            year = int(match.group(1))
            if 1950 <= year <= 2024:
                return str(year)

    # Fallback: any 4-digit year in a "since/founded" context
    years = [int(y) for y in YEAR_RE.findall(text) if 1950 <= int(y) <= 2024]
    return str(min(years)) if years else ""


def extract_size_signal(text: str) -> str:
    """Extract headcount or office count signals."""
    match = SIZE_RE.search(text)
    if match:
        return match.group(0).strip()
    return ""


def extract_page_text(soup: BeautifulSoup) -> str:
    """Get clean visible text from a page."""
    for tag in soup(["script", "style", "nav", "footer", "header"]):
        tag.decompose()
    return " ".join(soup.get_text(separator=" ").split())


# ─────────────────────────────────────────────
# AI ENRICHMENT (CLAUDE)
# ─────────────────────────────────────────────

AI_SYSTEM_PROMPT = """
You are a data extraction assistant. Given text from a business website,
extract structured information and return ONLY a valid JSON object with no
preamble, no markdown fences, and no explanation.

Return exactly these keys (use empty string "" if not found):
{
  "contact_name": "Full name of the owner, principal, or primary contact",
  "contact_title": "Their job title (e.g. Broker, President, Owner)",
  "employee_count": "Approximate number (e.g. '~12', '50+', 'small team')",
  "year_founded": "4-digit year as string",
  "primary_services": "Comma-separated list of main services offered",
  "size_tier": "micro (<5) | small (5-20) | mid (20-100) | large (100+) | unknown",
  "key_industries": "Industries or niches they serve, comma-separated",
  "notes": "Any other useful sales intelligence in one sentence"
}
"""

def ai_enrich(client: Anthropic, page_text: str) -> dict:
    """
    Send page text to Claude for structured extraction.
    Returns a dict of enriched fields, or empty dict on failure.
    """
    truncated = page_text[:MAX_TEXT_CHARS]
    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=600,
            system=AI_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": truncated}]
        )
        raw = response.content[0].text.strip()
        raw = raw.replace("```json", "").replace("```", "").strip()
        return json.loads(raw)
    except Exception as e:
        print(f"      ⚠ AI enrichment failed: {str(e)[:60]}")
        return {}


# ─────────────────────────────────────────────
# CORE ENRICHMENT FUNCTION
# ─────────────────────────────────────────────

def enrich_record(row: dict, client: Optional[Anthropic]) -> dict:
    """
    Enrich a single business record by visiting its website.
    Returns the row dict with new enrichment columns added.
    """
    enriched = row.copy()

    # Initialize enrichment columns
    enriched.setdefault("enriched_email", "")
    enriched.setdefault("contact_name", "")
    enriched.setdefault("contact_title", "")
    enriched.setdefault("year_founded", "")
    enriched.setdefault("size_signal", "")
    enriched.setdefault("size_tier", "")
    enriched.setdefault("employee_count", "")
    enriched.setdefault("primary_services", "")
    enriched.setdefault("key_industries", "")
    enriched.setdefault("linkedin", "")
    enriched.setdefault("facebook", "")
    enriched.setdefault("enrichment_notes", "")
    enriched.setdefault("enrichment_status", "skipped")

    website = normalize_url(str(row.get("website", "")))
    if not website:
        enriched["enrichment_status"] = "no_website"
        return enriched

    # Fetch the best available page
    soup, page_url = get_best_soup(website)
    if not soup:
        enriched["enrichment_status"] = "fetch_failed"
        return enriched

    raw_html = str(soup)
    page_text = extract_page_text(soup)

    # ── Regex extraction ──────────────────────────────────
    emails = extract_emails(soup, raw_html)
    if emails:
        # Prefer non-generic emails (skip info@, admin@, etc.)
        generic_prefixes = {"info", "admin", "contact", "hello", "support", "office"}
        specific = [e for e in emails if e.split("@")[0] not in generic_prefixes]
        enriched["enriched_email"] = specific[0] if specific else emails[0]

    social = extract_social(raw_html)
    enriched["linkedin"] = social["linkedin"]
    enriched["facebook"] = social["facebook"]

    if not enriched.get("year_founded"):
        enriched["year_founded"] = extract_year_founded(page_text)

    enriched["size_signal"] = extract_size_signal(page_text)

    # ── AI extraction ─────────────────────────────────────
    if client and USE_AI_ENRICHMENT and page_text:
        ai_data = ai_enrich(client, page_text)
        if ai_data:
            # Only fill AI fields if regex didn't already get them
            if not enriched["contact_name"]:
                enriched["contact_name"]   = ai_data.get("contact_name", "")
            enriched["contact_title"]      = ai_data.get("contact_title", "")
            enriched["employee_count"]     = ai_data.get("employee_count", "")
            enriched["size_tier"]          = ai_data.get("size_tier", "")
            enriched["primary_services"]   = ai_data.get("primary_services", "")
            enriched["key_industries"]     = ai_data.get("key_industries", "")
            enriched["enrichment_notes"]   = ai_data.get("notes", "")

            if not enriched["year_founded"]:
                enriched["year_founded"] = ai_data.get("year_founded", "")

    enriched["enrichment_status"] = "success"
    return enriched


# ─────────────────────────────────────────────
# BATCH RUNNER
# ─────────────────────────────────────────────

def run_enrichment(input_file: str, output_file: str, api_key: Optional[str] = None,
                   limit: Optional[int] = None, skip_existing: bool = True):
    """
    Main enrichment loop. Reads input CSV, enriches each row, writes output CSV.

    Args:
        input_file:     Path to input CSV (Layer 1 output)
        output_file:    Path for enriched output CSV
        api_key:        Anthropic API key (optional; falls back to ANTHROPIC_API_KEY env var)
        limit:          Only process first N rows (useful for testing)
        skip_existing:  Skip rows that already have a website email
    """
    print("\n" + "="*60)
    print("LAYER 2: WEBSITE ENRICHMENT PASS")
    print("="*60)

    # Load input
    try:
        df = pd.read_csv(input_file)
        print(f"✓ Loaded {len(df)} records from {input_file}")
    except Exception as e:
        print(f"✗ Could not read input file: {e}")
        return

    if limit:
        df = df.head(limit)
        print(f"  (Processing first {limit} records)")

    # Set up Claude client
    client = None
    key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if key and USE_AI_ENRICHMENT:
        client = Anthropic(api_key=key)
        print("✓ Claude AI enrichment enabled")
    else:
        print("⚠ No API key — running regex-only enrichment")

    # Count how many have websites
    website_col = "website" if "website" in df.columns else None
    if not website_col:
        print("✗ No 'website' column found in input CSV. Cannot enrich.")
        return

    has_website = df[website_col].notna() & (df[website_col].str.strip() != "")
    print(f"  Records with websites: {has_website.sum()} / {len(df)}")
    print()

    results = []
    success = skipped = failed = no_site = 0

    for i, (_, row) in enumerate(df.iterrows(), 1):
        name = row.get("name", f"Row {i}")
        website = str(row.get("website", "")).strip()

        print(f"[{i}/{len(df)}] {name[:50]}")

        if not website or website == "nan":
            print("      → No website, skipping")
            results.append(enrich_record(row.to_dict(), None))
            no_site += 1
            continue

        if skip_existing and row.get("email") and "@" in str(row.get("email", "")):
            print("      → Email already exists, skipping enrichment")
            record = row.to_dict()
            record["enrichment_status"] = "already_had_email"
            results.append(record)
            skipped += 1
            continue

        print(f"      → Visiting {website[:60]}")
        enriched = enrich_record(row.to_dict(), client)

        status = enriched.get("enrichment_status", "")
        if status == "success":
            success += 1
            details = []
            if enriched.get("enriched_email"):
                details.append(f"email: {enriched['enriched_email']}")
            if enriched.get("contact_name"):
                details.append(f"contact: {enriched['contact_name']}")
            if enriched.get("year_founded"):
                details.append(f"founded: {enriched['year_founded']}")
            if enriched.get("linkedin"):
                details.append("linkedin ✓")
            print(f"      ✓ {' | '.join(details) if details else 'fetched (limited data)'}")
        elif status == "fetch_failed":
            failed += 1
            print("      ✗ Could not reach website")
        else:
            print(f"      → {status}")

        results.append(enriched)

        # Save progress every 25 records
        if i % 25 == 0:
            pd.DataFrame(results).to_csv(output_file, index=False)
            print(f"\n  💾 Progress saved ({i} records)\n")

        time.sleep(DELAY_BETWEEN)

    # Final export
    out_df = pd.DataFrame(results)

    # Reorder columns: original columns first, then enrichment columns
    enrichment_cols = [
        "enriched_email", "contact_name", "contact_title",
        "employee_count", "size_tier", "size_signal",
        "year_founded", "primary_services", "key_industries",
        "linkedin", "facebook", "enrichment_notes", "enrichment_status"
    ]
    original_cols = [c for c in df.columns if c not in enrichment_cols]
    final_cols = original_cols + [c for c in enrichment_cols if c in out_df.columns]
    out_df = out_df[final_cols]

    out_df.to_csv(output_file, index=False, encoding="utf-8")

    print("\n" + "="*60)
    print("ENRICHMENT COMPLETE")
    print("="*60)
    print(f"  Total processed : {len(results)}")
    print(f"  Successful      : {success}")
    print(f"  Skipped         : {skipped}")
    print(f"  No website      : {no_site}")
    print(f"  Fetch failed    : {failed}")
    print(f"\n  Emails found    : {out_df['enriched_email'].notna().sum() if 'enriched_email' in out_df else 0}")
    print(f"  Contacts found  : {(out_df['contact_name'] != '').sum() if 'contact_name' in out_df else 0}")
    print(f"  LinkedIn found  : {(out_df['linkedin'] != '').sum() if 'linkedin' in out_df else 0}")
    print(f"\n✓ Output saved to: {output_file}")
    print("="*60 + "\n")


# ─────────────────────────────────────────────
# CLI ENTRY POINT
# ─────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        description="Layer 2 Website Enrichment — enriches a business CSV with data from each company's website"
    )
    parser.add_argument("--input",  default=DEFAULT_INPUT,  help="Input CSV file path")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="Output CSV file path")
    parser.add_argument("--api-key", default=None,          help="Anthropic API key (or set ANTHROPIC_API_KEY env var)")
    parser.add_argument("--limit",  type=int, default=None, help="Only process first N records (for testing)")
    parser.add_argument("--no-ai",  action="store_true",    help="Disable AI enrichment, use regex only")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    if args.no_ai:
        USE_AI_ENRICHMENT = False

    run_enrichment(
        input_file=args.input,
        output_file=args.output,
        api_key=args.api_key,
        limit=args.limit,
    )
