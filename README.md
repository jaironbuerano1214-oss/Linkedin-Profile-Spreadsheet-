# LinkedIn Profile Finder — Pampanga IT Alumni

Automates finding LinkedIn profiles for the names in **`list of names (1).xlsx`**.

Searches Google for each name (via Google Programmable Search Engine or SerpAPI),
filters results by:

| Filter | Criteria |
|--------|----------|
| **School** | Don Honorio Ventura State University (DHVSU), Don Honorio, Pampanga State University |
| **OR Location** | Pampanga (any city or municipality) |
| **AND Field** | IT-related — Computer Science (CSS), Information Systems (IS), Information Technology (IT), developer, programmer, software engineer, etc. |

Matching profile URLs are written back to **column C** of the same Excel file.

---

## 1. Prerequisites

- Python 3.11+
- A search API key (choose **one** of the options below)

### Option A — Google Programmable Search Engine (recommended, free tier: 100 queries/day)

1. Go to https://programmablesearchengine.google.com/ and create a new search engine.
2. Enable **"Search the entire web"** in the settings.
3. Copy your **Search Engine ID** (CX).
4. Get an API key from https://console.cloud.google.com/ → enable *Custom Search API*.

### Option B — SerpAPI (fallback, free tier: 100 searches/month)

1. Sign up at https://serpapi.com/.
2. Copy your API key from the dashboard.

---

## 2. Installation

```bash
pip install -r requirements.txt
```

---

## 3. Configuration

Set environment variables before running:

**Windows PowerShell:**
```powershell
$env:GOOGLE_API_KEY = "YOUR_GOOGLE_API_KEY"
$env:GOOGLE_CX      = "YOUR_SEARCH_ENGINE_ID"

# OR use SerpAPI instead:
$env:SERP_API_KEY   = "YOUR_SERP_API_KEY"
```

**Optional overrides:**

| Variable | Default | Description |
|----------|---------|-------------|
| `EXCEL_FILE` | `list of names (1).xlsx` | Path to the Excel file |
| `REQUEST_DELAY` | `2.0` | Seconds to wait between API calls (avoid rate-limits) |

---

## 4. Usage

### Run on all rows
```powershell
python linkedin_finder.py
```

### Dry run (search only, no writes to Excel)
```powershell
python linkedin_finder.py --dry-run
```

### Process a specific range of rows
```powershell
python linkedin_finder.py --start-row 70 --end-row 100
```

### Test a single name
```powershell
python linkedin_finder.py --name "Dela Cruz, Juan M."
```

---

## 5. Output

- **Excel file** — Column C (`linkedin profile`) is updated in-place with found URLs.
- **linkedin_finder.log** — Full log of every search query, match, and skip.

---

## 6. How it works

```
For each name in column A (skip rows that already have a link in column C):

  1. Convert "Lastname, Firstname M." → "Firstname Lastname" for search
  2. Run 3 progressively broader Google queries:
       a. site:linkedin.com/in + name + school keywords + IT keywords
       b. site:linkedin.com/in + name + Pampanga + IT keywords
       c. site:linkedin.com/in + name + Philippines
  3. For each LinkedIn URL in results:
       - Check title + snippet against school/location + IT keyword lists
       - If snippet is inconclusive, fetch the public profile page
       - If both geo filter (school OR location) AND IT filter pass → save URL
  4. Write URL to column C, save Excel after each batch
```

---

## 7. Filter keywords (customisable in `linkedin_finder.py`)

**School / location keywords** (`SCHOOL_KEYWORDS` + `LOCATION_KEYWORDS`):
- `don honorio ventura`, `dhvsu`, `pampanga state university`, `pampanga`,
  `angeles city`, `san fernando`, `mabalacat`, `clark`, and 15+ more Pampanga municipalities.

**IT field keywords** (`IT_KEYWORDS`):
- `css`, `bs cs`, `bsit`, `bsis`, `information technology`, `computer science`,
  `information systems`, `software`, `developer`, `programmer`, `engineer`, and more.

To add or remove keywords, edit the three lists near the top of `linkedin_finder.py`.

---

## 8. Rate limits & ethics

- The script respects a configurable delay (`REQUEST_DELAY`, default 2 s) between API calls.
- LinkedIn's Terms of Service prohibit automated scraping of profiles.
  This tool uses **public search results only** (Google index / snippet).
  Direct profile fetching is attempted only as a best-effort fallback.
- Use responsibly and in compliance with LinkedIn's ToS and applicable laws.
