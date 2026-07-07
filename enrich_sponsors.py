"""
Phase 2: enrich approved-sponsor company names with website, Australian state,
industry, and a confidence rating based on independent source agreement.

Pipeline (free, no ABN/ABR lookups, no paid APIs):
  Tier 0 - Wikidata: official, structured, zero-cost lookup for companies that
           already have a Wikidata entry (mostly large/well-known corporates).
  Tier 1 - DuckDuckGo web search (unofficial, free) -> top N organic results
           are handed to a free-tier Gemini model, which is only allowed to
           extract facts from those results (never from its own memory).
           Confidence is derived from how many independent result domains
           agree on the same website/state/industry.
  Tier 2 - Unresolved: no usable/agreeing evidence found. Recorded honestly
           rather than guessed.

Checkpointing: every company's result is committed to a local SQLite database
the moment it's produced, so the script can be killed (Ctrl+C, power loss,
daily rate-limit cutoff) at any time and simply resumed later with the same
command - already-processed companies are skipped automatically.

Setup (one-time):
    pip install ddgs google-genai openpyxl
    Get a free API key from https://aistudio.google.com/apikey (no card needed)
    Windows (PowerShell):  $env:GEMINI_API_KEY = "your-key-here"
    Or create a .env-style file and load it however you prefer.

Usage:
    python enrich_sponsors.py --input approved_sponsors_unique.xlsx --output enriched_sponsors.xlsx
    (re-run the same command any time to resume from where it stopped)

    Optional flags:
    --limit 50          process only the first 50 unprocessed companies (good for a test run)
    --db enrichment.db  custom checkpoint database path (default: enrichment.db)
    --skip-wikidata     skip Tier 0 and go straight to search+LLM for every row
"""

from __future__ import annotations

import argparse
import json
import os
import random
import re
import sqlite3
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

from openpyxl import Workbook, load_workbook

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

GEMINI_MODEL = "gemini-2.5-flash-lite"  # cheapest/most generous free-tier model
# If this model name is retired by the time you run this, check
# https://ai.google.dev/gemini-api/docs/models for the current Flash-Lite id
# and update the line above - nothing else in this script needs to change.

MAX_SEARCH_RESULTS = 5
GEMINI_MIN_DELAY_SECONDS = 4.5   # keeps us comfortably under free-tier RPM caps
DDG_MIN_DELAY_SECONDS = 2.0
MAX_CONSECUTIVE_ERRORS = 8       # stop the run cleanly if something is broken

AUS_STATES = {
    "NSW", "VIC", "QLD", "WA", "SA", "TAS", "ACT", "NT", "Unknown",
}

# ANZSIC-style top-level industry divisions - official Australian standard,
# used so results stay consistent across ~3,400 rows instead of the model
# inventing a new industry label every time.
INDUSTRY_DIVISIONS = [
    "Agriculture, Forestry and Fishing",
    "Mining",
    "Manufacturing",
    "Electricity, Gas, Water and Waste Services",
    "Construction",
    "Wholesale Trade",
    "Retail Trade",
    "Accommodation and Food Services",
    "Transport, Postal and Warehousing",
    "Information Media and Telecommunications",
    "Financial and Insurance Services",
    "Rental, Hiring and Real Estate Services",
    "Professional, Scientific and Technical Services",
    "Administrative and Support Services",
    "Public Administration and Safety",
    "Education and Training",
    "Health Care and Social Assistance",
    "Arts and Recreation Services",
    "Other Services",
    "Unknown",
]


# ---------------------------------------------------------------------------
# SQLite checkpoint store
# ---------------------------------------------------------------------------

def init_db(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS enrichment (
            source_order INTEGER PRIMARY KEY,
            sponsor_name TEXT NOT NULL,
            website TEXT,
            state TEXT,
            industry TEXT,
            confidence TEXT,
            confidence_reason TEXT,
            resolved_by TEXT,
            status TEXT NOT NULL,
            raw_response TEXT,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.commit()
    return conn


def already_done(conn: sqlite3.Connection, source_order: int) -> bool:
    row = conn.execute(
        "SELECT status FROM enrichment WHERE source_order = ?", (source_order,)
    ).fetchone()
    return row is not None and row[0] in ("done", "unresolved")


def save_result(conn: sqlite3.Connection, source_order: int, name: str, result: dict) -> None:
    conn.execute(
        """
        INSERT INTO enrichment
            (source_order, sponsor_name, website, state, industry,
             confidence, confidence_reason, resolved_by, status, raw_response)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(source_order) DO UPDATE SET
            website=excluded.website, state=excluded.state, industry=excluded.industry,
            confidence=excluded.confidence, confidence_reason=excluded.confidence_reason,
            resolved_by=excluded.resolved_by, status=excluded.status,
            raw_response=excluded.raw_response, updated_at=CURRENT_TIMESTAMP
        """,
        (
            source_order,
            name,
            result.get("website"),
            result.get("state"),
            result.get("industry"),
            result.get("confidence"),
            result.get("confidence_reason"),
            result.get("resolved_by"),
            result.get("status", "done"),
            json.dumps(result.get("raw_response"))[:5000] if result.get("raw_response") else None,
        ),
    )
    conn.commit()  # commit immediately -> this IS the checkpoint


# ---------------------------------------------------------------------------
# Tier 0: Wikidata (official, free, structured)
# ---------------------------------------------------------------------------

def wikidata_lookup(name: str) -> dict | None:
    """Look up a company on Wikidata. Returns a result dict if a confident
    match with an official website is found, else None."""
    try:
        search_url = (
            "https://www.wikidata.org/w/api.php?action=wbsearchentities"
            f"&search={urllib.parse.quote(name)}&language=en&format=json&limit=1&type=item"
        )
        req = urllib.request.Request(search_url, headers={"User-Agent": "SponsorEnrichment/1.0"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        hits = data.get("search", [])
        if not hits:
            return None

        qid = hits[0]["id"]
        entity_url = (
            f"https://www.wikidata.org/wiki/Special:EntityData/{qid}.json"
        )
        req2 = urllib.request.Request(entity_url, headers={"User-Agent": "SponsorEnrichment/1.0"})
        with urllib.request.urlopen(req2, timeout=15) as resp2:
            entity_data = json.loads(resp2.read().decode("utf-8"))

        claims = entity_data["entities"][qid]["claims"]

        # P17 = country; require Australia (Q408) for a confident match
        country_claims = claims.get("P17", [])
        is_australian = any(
            c["mainsnak"]["datavalue"]["value"]["id"] == "Q408"
            for c in country_claims
            if c["mainsnak"].get("datavalue")
        )
        if not is_australian:
            return None

        website = None
        if "P856" in claims:
            website = claims["P856"][0]["mainsnak"]["datavalue"]["value"]

        if not website:
            return None  # not useful enough without a website

        return {
            "website": website,
            "state": "Unknown",  # Wikidata rarely encodes AU state cleanly; leave for manual/LLM pass if needed
            "industry": "Unknown",
            "confidence": "High",
            "confidence_reason": "Matched an Australian entity on Wikidata with an official website property (P856).",
            "resolved_by": "wikidata",
            "status": "done",
            "raw_response": {"qid": qid},
        }
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Tier 1: DuckDuckGo search + Gemini extraction
# ---------------------------------------------------------------------------

def ddg_search(query: str, max_results: int = MAX_SEARCH_RESULTS) -> list[dict]:
    from ddgs import DDGS

    with DDGS() as ddgs:
        results = list(ddgs.text(query, max_results=max_results))
    return [
        {
            "title": r.get("title", ""),
            "href": r.get("href", ""),
            "body": r.get("body", ""),
        }
        for r in results
    ]


def domain_of(url: str) -> str:
    try:
        netloc = urllib.parse.urlparse(url).netloc.lower()
        return re.sub(r"^www\.", "", netloc)
    except Exception:
        return ""


def build_prompt(name: str, results: list[dict]) -> str:
    snippets_text = "\n\n".join(
        f"[Result {i+1}] {r['title']}\nURL: {r['href']}\n{r['body']}"
        for i, r in enumerate(results)
    )
    states_list = ", ".join(sorted(AUS_STATES - {"Unknown"}))
    industries_list = "\n".join(f"- {i}" for i in INDUSTRY_DIVISIONS)

    return f"""You are extracting factual data about an Australian company from search results only.
Company name (from an Australian government approved-sponsor list): "{name}"

Search results:
{snippets_text}

Rules:
- Use ONLY the information present in the search results above. Do not use outside knowledge.
- If the results don't clearly identify the company, say so - do not guess.
- "website" must be the company's own official domain (not a directory, LinkedIn, Facebook, or news site), or null.
- "state" must be one of: {states_list}, or "Unknown" if not determinable.
- "industry" must be exactly one of these standard categories:
{industries_list}
- "supporting_result_indices" must list which of the numbered results (1-based) support your website answer.

Respond with ONLY this JSON object, no other text:
{{
  "website": "example.com.au or null",
  "state": "one of the allowed values",
  "industry": "one of the allowed categories",
  "supporting_result_indices": [1, 2],
  "reasoning": "one short sentence"
}}"""


def call_gemini(client, prompt: str) -> dict:
    from google.genai import types

    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.0,
            max_output_tokens=400,
            response_mime_type="application/json",
        ),
    )
    text = response.text.strip()
    text = re.sub(r"^```json|```$", "", text, flags=re.MULTILINE).strip()
    return json.loads(text)


def compute_confidence(extracted: dict, results: list[dict]) -> tuple[str, str]:
    """Confidence reflects genuine independent corroboration: does the claimed
    website itself show up in results, AND is it separately mentioned by a
    different domain (directory, news, etc.)? Two pages from the same site
    are not independent evidence of each other."""
    website = (extracted.get("website") or "").strip().lower()
    if not website or website == "null":
        return "None", "No website could be identified from the search results."

    target_domain = domain_of(website if website.startswith("http") else f"https://{website}")
    if not target_domain:
        target_domain = website.replace("https://", "").replace("http://", "").strip("/")

    official_domain_present = any(domain_of(r["href"]) == target_domain for r in results)

    third_party_domains = set()
    for r in results:
        d = domain_of(r["href"])
        if not d or d == target_domain:
            continue
        text = f"{r.get('title', '')} {r.get('body', '')}".lower()
        if target_domain in text:
            third_party_domains.add(d)

    if official_domain_present and third_party_domains:
        return (
            "High",
            f"The official site ({target_domain}) appeared directly, and "
            f"{len(third_party_domains)} independent source(s) separately mention it.",
        )
    if official_domain_present:
        return (
            "Medium",
            f"{target_domain} appeared in search results but no independent "
            "third-party source corroborates it.",
        )
    if len(third_party_domains) >= 2:
        return (
            "Medium",
            f"{len(third_party_domains)} independent sources mention {target_domain}, "
            "but the site itself wasn't directly indexed in these results.",
        )
    if len(third_party_domains) == 1:
        return "Low", f"Only one indirect mention of {target_domain}; not corroborated."
    return "Low", "Model proposed a website but it isn't backed by the supplied search results."


def enrich_via_search(client, name: str) -> dict:
    query = f'"{name}" Australia'
    results = ddg_search(query)

    if not results:
        return {
            "website": None,
            "state": "Unknown",
            "industry": "Unknown",
            "confidence": "None",
            "confidence_reason": "No search results returned.",
            "resolved_by": "search+llm",
            "status": "unresolved",
            "raw_response": None,
        }

    prompt = build_prompt(name, results)
    extracted = call_gemini(client, prompt)

    confidence, reason = compute_confidence(extracted, results)
    website = extracted.get("website")
    if isinstance(website, str) and website.strip().lower() == "null":
        website = None

    status = "done" if confidence in ("High", "Medium") else "unresolved"

    return {
        "website": website,
        "state": extracted.get("state", "Unknown"),
        "industry": extracted.get("industry", "Unknown"),
        "confidence": confidence,
        "confidence_reason": reason,
        "resolved_by": "search+llm",
        "status": status,
        "raw_response": extracted,
    }


# ---------------------------------------------------------------------------
# Excel I/O
# ---------------------------------------------------------------------------

def load_companies(input_path: Path) -> list[tuple[int, str]]:
    wb = load_workbook(input_path, read_only=True)
    sheet = wb.active
    rows = list(sheet.iter_rows(min_row=2, values_only=True))
    return [(row[0], row[1]) for row in rows if row[1]]


def export_to_excel(db_path: Path, output_path: Path) -> None:
    conn = sqlite3.connect(db_path)
    rows = conn.execute(
        """
        SELECT source_order, sponsor_name, website, state, industry,
               confidence, confidence_reason, resolved_by, status
        FROM enrichment ORDER BY source_order
        """
    ).fetchall()
    conn.close()

    wb = Workbook()
    sheet = wb.active
    sheet.title = "Enriched Sponsors"
    sheet.append([
        "Source Order", "Sponsor Name", "Website", "State", "Industry",
        "Confidence", "Confidence Reason", "Resolved By", "Status",
    ])
    for row in rows:
        sheet.append(row)

    widths = {"A": 12, "B": 55, "C": 30, "D": 10, "E": 35, "F": 12, "G": 55, "H": 14, "I": 12}
    for col, width in widths.items():
        sheet.column_dimensions[col].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    wb.save(output_path)


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Enrich approved sponsor list (Phase 2).")
    parser.add_argument("--input", type=Path, required=True, help="Input xlsx from Phase 1.")
    parser.add_argument("--output", type=Path, required=True, help="Output enriched xlsx.")
    parser.add_argument("--db", type=Path, default=Path("enrichment.db"), help="Checkpoint database path.")
    parser.add_argument("--limit", type=int, default=None, help="Only process this many new rows (test runs).")
    parser.add_argument("--skip-wikidata", action="store_true", help="Skip Tier 0 and go straight to search+LLM.")
    args = parser.parse_args()

    api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        sys.exit(
            "GEMINI_API_KEY (or GOOGLE_API_KEY) environment variable is not set.\n"
            "Get a free key from https://aistudio.google.com/apikey and set it, e.g.\n"
            "  PowerShell: $env:GEMINI_API_KEY = \"your-key-here\"\n"
            "  bash:       export GEMINI_API_KEY=\"your-key-here\""
        )

    from google import genai
    client = genai.Client(api_key=api_key)

    conn = init_db(args.db)
    companies = load_companies(args.input)
    print(f"Loaded {len(companies):,} companies from {args.input}")

    processed_this_run = 0
    consecutive_errors = 0

    for source_order, name in companies:
        if args.limit is not None and processed_this_run >= args.limit:
            print(f"Reached --limit {args.limit}; stopping (resumable).")
            break

        if already_done(conn, source_order):
            continue

        try:
            result = None
            if not args.skip_wikidata:
                result = wikidata_lookup(name)
                time.sleep(DDG_MIN_DELAY_SECONDS)

            if result is None:
                result = enrich_via_search(client, name)
                time.sleep(GEMINI_MIN_DELAY_SECONDS + random.uniform(0, 1.5))

            save_result(conn, source_order, name, result)
            processed_this_run += 1
            consecutive_errors = 0

            status_flag = "OK" if result["status"] == "done" else "unresolved"
            print(f"[{source_order}] {name[:60]:<60} -> {result.get('confidence','?'):<6} ({status_flag})")

        except KeyboardInterrupt:
            print("\nInterrupted by user. Progress is saved - re-run the same command to resume.")
            break

        except Exception as exc:
            consecutive_errors += 1
            print(f"[{source_order}] {name[:60]:<60} -> ERROR: {exc}")
            if "429" in str(exc) or "rate" in str(exc).lower() or "quota" in str(exc).lower():
                print("Looks like a rate/quota limit was hit. Stopping cleanly - "
                      "re-run this same command later (e.g. tomorrow) to resume.")
                break
            if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                print(f"{MAX_CONSECUTIVE_ERRORS} consecutive errors - stopping to avoid burning quota on a broken run.")
                break
            time.sleep(5)

    conn.close()
    export_to_excel(args.db, args.output)
    print(f"\nExported current progress to {args.output}")
    print("Re-run this same command any time to continue enriching the remaining companies.")


if __name__ == "__main__":
    main()
