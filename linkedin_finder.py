"""
LinkedIn Profile Finder
=======================
Reads names from 'list of names (1).xlsx', searches for each person's
LinkedIn profile via Google Custom Search (or SerpAPI fallback), filters
by DHVSU / Don Honorio / Pampanga State Univ OR Pampanga location AND
IT-related field (CSS / IS / IT), then writes the profile URL back to
column C of the same file.

Setup:
    pip install -r requirements.txt
    Set env vars (see README.md) before running.

Usage:
    python linkedin_finder.py
"""

import os
import re
import time
import json
import logging
import argparse
from pathlib import Path

import requests
import openpyxl
from openpyxl import load_workbook

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("linkedin_finder.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Config — override via environment variables
# ---------------------------------------------------------------------------
EXCEL_FILE = os.getenv("EXCEL_FILE", "list of names (1).xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "Sheet1")
NAME_COL   = 1   # Column A
LINK_COL   = 3   # Column C  (linkedin profile)
HEADER_ROW = 1

# Google Programmable Search Engine
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY", "")
GOOGLE_CX      = os.getenv("GOOGLE_CX", "")      # Search engine ID

# SerpAPI (fallback if Google CSE creds are absent)
SERP_API_KEY = os.getenv("SERP_API_KEY", "")

# Rate-limit: seconds to wait between API calls
REQUEST_DELAY = float(os.getenv("REQUEST_DELAY", "2.0"))

# ---------------------------------------------------------------------------
# Filter keywords
# ---------------------------------------------------------------------------
SCHOOL_KEYWORDS = [
    "don honorio ventura",
    "dhvsu",
    "don honorio",
    "pampanga state university",
    "psu pampanga",
    "pamantasan ng lungsod ng san fernando",
]

LOCATION_KEYWORDS = [
    "pampanga",
    "san fernando",
    "angeles city",
    "mabalacat",
    "magalang",
    "apalit",
    "bacolor",
    "candaba",
    "floridablanca",
    "guagua",
    "lubao",
    "macabebe",
    "masantol",
    "mexico",  # municipality in pampanga
    "minalin",
    "porac",
    "san luis",
    "san simon",
    "santa ana",
    "santa rita",
    "santo tomas",
    "sasmuan",
    "clark",
]

IT_KEYWORDS = [
    # Courses / degrees
    "computer science",
    "information technology",
    "information systems",
    "css",
    "bs cs",
    "bscs",
    "bs it",
    "bsit",
    "bs is",
    "bsis",
    "bachelor of science in information technology",
    "bachelor of science in computer science",
    "bachelor of science in information systems",
    # Job titles / fields
    "software",
    "developer",
    "programmer",
    "engineer",
    "web developer",
    "data analyst",
    "cybersecurity",
    "network",
    "systems analyst",
    "tech",
    "it specialist",
    "cloud",
    "devops",
    "full stack",
    "frontend",
    "backend",
    "mobile developer",
    "qa engineer",
    "ui/ux",
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def normalize(text: str) -> str:
    """Lowercase + collapse whitespace for fuzzy matching."""
    return re.sub(r"\s+", " ", text.lower().strip())


def contains_any(text: str, keywords: list[str]) -> bool:
    t = normalize(text)
    return any(kw in t for kw in keywords)


def name_to_query(raw_name: str) -> str:
    """
    Convert 'Lastname, Firstname M.' → 'Firstname Lastname' for cleaner search.
    Falls back to the raw name if no comma is found.
    """
    raw_name = raw_name.strip()
    if "," in raw_name:
        parts = raw_name.split(",", 1)
        last = parts[0].strip()
        # Strip middle initial pattern like 'M.' at end
        first_parts = parts[1].strip()
        # Remove trailing single-letter-dot middle initial
        first_parts = re.sub(r"\s+[A-Za-z]\.$", "", first_parts).strip()
        return f"{first_parts} {last}"
    return raw_name


# ---------------------------------------------------------------------------
# Search backends
# ---------------------------------------------------------------------------

def search_google_cse(query: str) -> list[dict]:
    """
    Search using Google Programmable Search Engine (CSE).
    Returns list of {'title': ..., 'link': ..., 'snippet': ...}
    """
    if not GOOGLE_API_KEY or not GOOGLE_CX:
        return []
    url = "https://www.googleapis.com/customsearch/v1"
    params = {
        "key": GOOGLE_API_KEY,
        "cx":  GOOGLE_CX,
        "q":   query,
        "num": 5,
    }
    try:
        r = requests.get(url, params=params, timeout=15)
        r.raise_for_status()
        data = r.json()
        return data.get("items", [])
    except Exception as e:
        log.warning("Google CSE error for '%s': %s", query, e)
        return []


def search_serp(query: str) -> list[dict]:
    """
    Search using SerpAPI as fallback.
    Returns list of {'title': ..., 'link': ..., 'snippet': ...}
    """
    if not SERP_API_KEY:
        return []
    url = "https://serpapi.com/search"
    params = {
        "api_key": SERP_API_KEY,
        "engine":  "google",
        "q":       query,
        "num":     5,
    }
    try:
        r = requests.get(url, params=params, timeout=15)
        r.raise_for_status()
        data = r.json()
        results = []
        for item in data.get("organic_results", []):
            results.append({
                "title":   item.get("title", ""),
                "link":    item.get("link", ""),
                "snippet": item.get("snippet", ""),
            })
        return results
    except Exception as e:
        log.warning("SerpAPI error for '%s': %s", query, e)
        return []


def search(query: str) -> list[dict]:
    """Try Google CSE first, then SerpAPI."""
    results = search_google_cse(query)
    if not results:
        results = search_serp(query)
    return results


# ---------------------------------------------------------------------------
# LinkedIn profile scraping (public profile snippet only)
# ---------------------------------------------------------------------------

def fetch_linkedin_snippet(profile_url: str) -> str:
    """
    Attempt a lightweight fetch of the LinkedIn public profile page.
    LinkedIn heavily blocks scraping — we only try to read a cached
    Google snippet / description; this is a best-effort grab.
    Returns raw page text (truncated) or empty string on failure.
    """
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0 Safari/537.36"
        ),
    }
    try:
        r = requests.get(profile_url, headers=headers, timeout=10, allow_redirects=True)
        return r.text[:4000]
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Core filter logic
# ---------------------------------------------------------------------------

def is_linkedin_url(url: str) -> bool:
    return "linkedin.com/in/" in url.lower()


def passes_filter(title: str, snippet: str, profile_text: str = "") -> bool:
    """
    Returns True if the combined text suggests:
      - School: DHVSU / Don Honorio / Pampanga State Univ
        OR location is Pampanga (any city/province)
      AND
      - IT-related course or job (CSS / IS / IT / software / developer …)
    """
    combined = normalize(f"{title} {snippet} {profile_text}")

    school_match   = contains_any(combined, SCHOOL_KEYWORDS)
    location_match = contains_any(combined, LOCATION_KEYWORDS)
    it_match       = contains_any(combined, IT_KEYWORDS)

    geo_ok = school_match or location_match
    return geo_ok and it_match


# ---------------------------------------------------------------------------
# Main processing
# ---------------------------------------------------------------------------

def find_profile(name: str) -> str | None:
    """
    Search LinkedIn for a person, apply filters, and return the first
    matching profile URL (or None).
    """
    search_name = name_to_query(name)

    # Build targeted queries — most specific first
    queries = [
        f'site:linkedin.com/in "{search_name}" DHVSU OR "Don Honorio" OR "Pampanga State" IT OR "Computer Science" OR "Information Technology"',
        f'site:linkedin.com/in "{search_name}" Pampanga IT OR developer OR programmer',
        f'site:linkedin.com/in "{search_name}" Philippines',
    ]

    for q in queries:
        log.info("  Searching: %s", q)
        results = search(q)
        time.sleep(REQUEST_DELAY)

        for item in results:
            link    = item.get("link", "")
            title   = item.get("title", "")
            snippet = item.get("snippet", "")

            if not is_linkedin_url(link):
                continue

            # Quick filter on search snippet alone
            if passes_filter(title, snippet):
                log.info("  ✓ Matched (snippet): %s", link)
                return link

            # Deeper check: fetch public profile page
            page_text = fetch_linkedin_snippet(link)
            time.sleep(REQUEST_DELAY)
            if passes_filter(title, snippet, page_text):
                log.info("  ✓ Matched (page): %s", link)
                return link

            log.info("  ✗ Filtered out: %s", link)

    return None


def run(dry_run: bool = False, start_row: int = 2, end_row: int | None = None):
    """
    Main loop — reads the Excel, searches LinkedIn for each name that
    doesn't already have a profile link, writes results back.
    """
    excel_path = Path(EXCEL_FILE)
    if not excel_path.exists():
        log.error("Excel file not found: %s", excel_path)
        return

    wb = load_workbook(excel_path)
    ws = wb[SHEET_NAME]

    max_row = ws.max_row
    if end_row is None or end_row > max_row:
        end_row = max_row

    log.info("Processing rows %d – %d in '%s'", start_row, end_row, EXCEL_FILE)

    found_count   = 0
    skipped_count = 0
    total         = end_row - start_row + 1

    for row_idx in range(start_row, end_row + 1):
        name_cell = ws.cell(row=row_idx, column=NAME_COL)
        link_cell = ws.cell(row=row_idx, column=LINK_COL)

        name = (name_cell.value or "").strip()
        if not name:
            log.debug("Row %d: empty name, skipping.", row_idx)
            skipped_count += 1
            continue

        existing_link = (link_cell.value or "").strip()
        if existing_link:
            log.info("Row %d [%s]: already has link → %s", row_idx, name, existing_link)
            skipped_count += 1
            continue

        log.info("Row %d/%d — Searching: %s", row_idx, end_row, name)

        profile_url = find_profile(name)

        if profile_url:
            log.info("  → Found: %s", profile_url)
            found_count += 1
            if not dry_run:
                link_cell.value = profile_url
        else:
            log.info("  → Not found / filtered out.")

    if not dry_run:
        wb.save(excel_path)
        log.info("Saved '%s'.", excel_path)
    else:
        log.info("[DRY RUN] No changes written.")

    log.info(
        "Done. Found: %d | Skipped: %d | Total processed: %d",
        found_count, skipped_count, total,
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="LinkedIn Profile Finder for Pampanga IT Alumni")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Search and log results without writing to the Excel file.",
    )
    parser.add_argument(
        "--start-row",
        type=int,
        default=2,
        help="Excel row to start from (default: 2, i.e. first data row).",
    )
    parser.add_argument(
        "--end-row",
        type=int,
        default=None,
        help="Excel row to stop at (default: last row).",
    )
    parser.add_argument(
        "--name",
        type=str,
        default=None,
        help="Test a single name directly (prints found URL or 'Not found').",
    )
    args = parser.parse_args()

    if args.name:
        url = find_profile(args.name)
        print(f"Result for '{args.name}': {url or 'Not found / filtered out'}")
    else:
        run(dry_run=args.dry_run, start_row=args.start_row, end_row=args.end_row)
