"""
build_post.py  —  Revised sourcing architecture
================================================
Stage 1: Gather recent opinion metadata from:
  - SCOTUSblog (Supreme Court, same-day coverage)
  - supremecourt.gov slip opinions
  - 2nd, 5th, 9th, 11th Circuit RSS feeds
  - D.C. Circuit HTML page
  - 3rd Circuit HTML page
  - Stanford Securities Class Action Clearinghouse
  - CourtListener (fallback + district courts)

Stage 2: Claude reads all headlines/summaries and picks
  the single most significant/newsworthy opinion.

Stage 3: Fetch the full opinion text — from the court's
  own PDF if available, otherwise CourtListener.

Stage 4: Claude writes the blog post.

Stage 5: Inject into index.html, update posted_cases.json.
"""

import os
import re
import json
import datetime
import textwrap
import time
import requests
import anthropic
from urllib.parse import urljoin

# ── Config ────────────────────────────────────────────────────────────────────

LOOKBACK_DAYS   = 7        # focus on the last week
HTML_PATH       = "index.html"
POSTED_LOG_PATH = "posted_cases.json"

HEADERS = {
    "User-Agent": (
        "MLegalTechnology blog builder (mlegaltechnology.com); "
        "contact: see site"
    )
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg):
    print(f"[build_post] {msg}")


def cutoff_date() -> datetime.date:
    return datetime.date.today() - datetime.timedelta(days=LOOKBACK_DAYS)


def get(url, timeout=15, **kwargs):
    """Simple GET with shared headers and error handling."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout, **kwargs)
        r.raise_for_status()
        return r
    except Exception as e:
        log(f"  GET failed {url[:80]}: {e}")
        return None


def parse_date(s: str) -> datetime.date | None:
    """Try several date formats and return a date or None."""
    if not s:
        return None
    s = s.strip()
    for fmt in (
        "%Y-%m-%d", "%B %d, %Y", "%b %d, %Y",
        "%d %B %Y", "%d %b %Y", "%m/%d/%Y",
        "%a, %d %b %Y %H:%M:%S %z", "%a, %d %b %Y %H:%M:%S %Z",
    ):
        try:
            return datetime.datetime.strptime(s[:len(fmt)+5], fmt).date()
        except ValueError:
            continue
    # Try trimming timezone suffixes and retrying
    s2 = re.sub(r"\s+(GMT|UTC|EST|EDT|CST|CDT|PST|PDT|[+-]\d{4})$", "", s)
    if s2 != s:
        return parse_date(s2)
    return None


def is_recent(date_str: str) -> bool:
    d = parse_date(date_str)
    if d is None:
        return True   # include if we can't parse — better to over-include
    return d >= cutoff_date()


def extract_pdf_text(pdf_bytes: bytes) -> str:
    """Extract text from PDF bytes using PyMuPDF if available."""
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        pages = []
        for page in doc:
            pages.append(page.get_text())
        return " ".join(pages)
    except ImportError:
        log("  PyMuPDF not available — skipping PDF text extraction")
        return ""
    except Exception as e:
        log(f"  PDF extraction error: {e}")
        return ""


def strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", text).strip()


# ── Duplicate tracking ────────────────────────────────────────────────────────

def load_posted_log() -> dict:
    if os.path.exists(POSTED_LOG_PATH):
        try:
            with open(POSTED_LOG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                data.setdefault("titles", [])
                data.setdefault("urls", [])
                return data
        except Exception as e:
            log(f"Warning: could not read {POSTED_LOG_PATH}: {e}")
    return {"titles": [], "urls": []}


def save_posted_log(data: dict):
    with open(POSTED_LOG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def normalize(s: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^\w\s]", " ", s.lower())).strip()


def is_duplicate(title: str, url: str, log_data: dict) -> bool:
    if url and url in log_data["urls"]:
        return True
    t = normalize(title)
    if any(t == x or (len(t) > 20 and (t in x or x in t))
           for x in log_data["titles"]):
        return True
    return False


def record_posted(log_data: dict, title: str, url: str) -> dict:
    t = normalize(title)
    if t not in log_data["titles"]:
        log_data["titles"].append(t)
    if url and url not in log_data["urls"]:
        log_data["urls"].append(url)
    return log_data


# ── Stage 1: Source scrapers ──────────────────────────────────────────────────
# Each scraper returns a list of dicts:
# {
#   "title":   str,    # case name or headline
#   "date":    str,    # date string (best effort)
#   "court":   str,    # human-readable court name
#   "summary": str,    # one-sentence description if available
#   "url":     str,    # link to opinion or coverage page
#   "pdf_url": str,    # direct PDF link if known
#   "source":  str,    # which scraper found this
# }

def scrape_scotusblog() -> list:
    """SCOTUSblog recent decisions page."""
    results = []
    r = get("https://www.scotusblog.com/case-files/terms/ot2024/")
    if not r:
        # Try the main decisions feed
        r = get("https://www.scotusblog.com/feed/")
    if not r:
        return results

    # Parse RSS if it's the feed
    if "<?xml" in r.text[:100] or "<rss" in r.text[:200]:
        entries = re.findall(
            r"<item>(.*?)</item>", r.text, re.DOTALL
        )
        for entry in entries:
            title = re.search(r"<title><!\[CDATA\[(.*?)\]\]>|<title>(.*?)</title>",
                              entry, re.DOTALL)
            title = (title.group(1) or title.group(2) or "").strip() if title else ""
            link = re.search(r"<link>(.*?)</link>|<link[^>]+href=\"([^\"]+)\"",
                             entry, re.DOTALL)
            link = (link.group(1) or link.group(2) or "").strip() if link else ""
            pub_date = re.search(r"<pubDate>(.*?)</pubDate>", entry)
            pub_date = pub_date.group(1).strip() if pub_date else ""
            desc = re.search(r"<description><!\[CDATA\[(.*?)\]\]>|<description>(.*?)</description>",
                             entry, re.DOTALL)
            desc = strip_html((desc.group(1) or desc.group(2) or "")) if desc else ""

            if not is_recent(pub_date):
                continue
            if title:
                results.append({
                    "title": title, "date": pub_date, "court": "U.S. Supreme Court",
                    "summary": desc[:300], "url": link, "pdf_url": "",
                    "source": "SCOTUSblog",
                })
    log(f"  SCOTUSblog: {len(results)} items")
    return results


def scrape_supremecourt_gov() -> list:
    """Pull slip opinions directly from supremecourt.gov."""
    results = []
    r = get("https://www.supremecourt.gov/opinions/slipopinion/24")
    if not r:
        return results

    # The page lists opinions in a table: date, docket, name, J., PDF link
    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", r.text, re.DOTALL)
    for row in rows:
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.DOTALL)
        if len(cells) < 3:
            continue
        date_text  = strip_html(cells[0]).strip()
        case_text  = strip_html(cells[2]).strip() if len(cells) > 2 else ""
        pdf_match  = re.search(r'href="(/opinions/[^"]+\.pdf)"', row, re.IGNORECASE)
        pdf_url    = f"https://www.supremecourt.gov{pdf_match.group(1)}" if pdf_match else ""

        if not date_text or not case_text:
            continue
        if not is_recent(date_text):
            continue

        results.append({
            "title": case_text, "date": date_text, "court": "U.S. Supreme Court",
            "summary": "", "url": pdf_url,
            "pdf_url": pdf_url, "source": "supremecourt.gov",
        })

    log(f"  supremecourt.gov: {len(results)} items")
    return results


def scrape_rss_circuit(name: str, feed_url: str) -> list:
    """Generic RSS scraper for circuit court opinion feeds."""
    results = []
    r = get(feed_url)
    if not r:
        return results

    entries = re.findall(r"<item>(.*?)</item>", r.text, re.DOTALL)
    for entry in entries:
        title = re.search(
            r"<title><!\[CDATA\[(.*?)\]\]>|<title>(.*?)</title>", entry, re.DOTALL
        )
        title = (title.group(1) or title.group(2) or "").strip() if title else ""
        link = re.search(
            r"<link>(.*?)</link>|<link[^>]+href=\"([^\"]+)\"", entry, re.DOTALL
        )
        link = (link.group(1) or link.group(2) or "").strip() if link else ""
        pub_date = re.search(r"<pubDate>(.*?)</pubDate>|<dc:date>(.*?)</dc:date>",
                             entry, re.DOTALL)
        pub_date = (pub_date.group(1) or pub_date.group(2) or "").strip() if pub_date else ""
        desc = re.search(
            r"<description><!\[CDATA\[(.*?)\]\]>|<description>(.*?)</description>",
            entry, re.DOTALL
        )
        desc = strip_html((desc.group(1) or desc.group(2) or "")) if desc else ""

        if not is_recent(pub_date):
            continue

        # Detect PDF link in description or link field
        pdf_url = ""
        if link.lower().endswith(".pdf"):
            pdf_url = link

        if title:
            results.append({
                "title": title, "date": pub_date, "court": name,
                "summary": desc[:300], "url": link,
                "pdf_url": pdf_url, "source": name,
            })

    log(f"  {name}: {len(results)} items")
    return results


def scrape_dc_circuit() -> list:
    """D.C. Circuit opinions page (HTML, no RSS)."""
    results = []
    r = get("https://www.cadc.uscourts.gov/internet/opinions.nsf/opinions?openview&count=30")
    if not r:
        # Try alternate URL
        r = get("https://www.cadc.uscourts.gov/internet/opinions.nsf")
    if not r:
        return results

    # Extract opinion rows — typically contain date, case name, PDF link
    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", r.text, re.DOTALL)
    for row in rows:
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.DOTALL)
        if len(cells) < 2:
            continue
        date_text = strip_html(cells[0]).strip()
        case_text = strip_html(cells[1]).strip() if len(cells) > 1 else ""
        pdf_match = re.search(r'href="([^"]+\.pdf)"', row, re.IGNORECASE)
        pdf_url   = ""
        if pdf_match:
            href = pdf_match.group(1)
            pdf_url = href if href.startswith("http") else urljoin(
                "https://www.cadc.uscourts.gov", href
            )

        if not date_text or not case_text or len(case_text) < 4:
            continue
        if not is_recent(date_text):
            continue

        results.append({
            "title": case_text, "date": date_text, "court": "D.C. Circuit",
            "summary": "", "url": pdf_url or "https://www.cadc.uscourts.gov",
            "pdf_url": pdf_url, "source": "D.C. Circuit",
        })

    log(f"  D.C. Circuit: {len(results)} items")
    return results


def scrape_third_circuit() -> list:
    """3rd Circuit precedential opinions (HTML)."""
    results = []
    r = get("https://www2.ca3.uscourts.gov/recentopinions")
    if not r:
        return results

    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", r.text, re.DOTALL)
    for row in rows:
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.DOTALL)
        if len(cells) < 2:
            continue
        date_text = strip_html(cells[0]).strip()
        case_text = strip_html(cells[1]).strip() if len(cells) > 1 else ""
        pdf_match = re.search(r'href="([^"]+\.pdf)"', row, re.IGNORECASE)
        pdf_url   = ""
        if pdf_match:
            href = pdf_match.group(1)
            pdf_url = href if href.startswith("http") else urljoin(
                "https://www2.ca3.uscourts.gov", href
            )

        if not date_text or not case_text or len(case_text) < 4:
            continue
        if not is_recent(date_text):
            continue

        results.append({
            "title": case_text, "date": date_text, "court": "3rd Circuit",
            "summary": "", "url": pdf_url or "https://www2.ca3.uscourts.gov",
            "pdf_url": pdf_url, "source": "3rd Circuit",
        })

    log(f"  3rd Circuit: {len(results)} items")
    return results


def scrape_first_circuit() -> list:
    """1st Circuit opinions page (HTML table)."""
    results = []
    r = get("https://www.ca1.uscourts.gov/opinions")
    if not r:
        return results

    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", r.text, re.DOTALL)
    for row in rows:
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.DOTALL)
        if len(cells) < 2:
            continue
        date_text = strip_html(cells[0]).strip()
        case_text = strip_html(cells[1]).strip() if len(cells) > 1 else ""
        pdf_match = re.search(r'href="([^"]+\.pdf)"', row, re.IGNORECASE)
        pdf_url   = ""
        if pdf_match:
            href = pdf_match.group(1)
            pdf_url = href if href.startswith("http") else urljoin(
                "https://www.ca1.uscourts.gov", href
            )

        if not date_text or not case_text or len(case_text) < 4:
            continue
        if not is_recent(date_text):
            continue

        results.append({
            "title": case_text, "date": date_text, "court": "1st Circuit",
            "summary": "", "url": pdf_url or "https://www.ca1.uscourts.gov",
            "pdf_url": pdf_url, "source": "1st Circuit",
        })

    log(f"  1st Circuit: {len(results)} items")
    return results


def scrape_stanford_clearinghouse() -> list:
    """Stanford Securities Class Action Clearinghouse — recent decisions."""
    results = []
    r = get("https://securities.stanford.edu/class-action-filings/decisions.html")
    if not r:
        return results

    # Extract case links and dates from the decisions table
    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", r.text, re.DOTALL)
    for row in rows:
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.DOTALL)
        if len(cells) < 2:
            continue
        date_text = strip_html(cells[0]).strip()
        case_text = strip_html(cells[1]).strip() if len(cells) > 1 else ""
        link_match = re.search(r'href="([^"]+)"', cells[1]) if len(cells) > 1 else None
        url = ""
        if link_match:
            href = link_match.group(1)
            url = href if href.startswith("http") else urljoin(
                "https://securities.stanford.edu", href
            )

        if not date_text or not case_text or len(case_text) < 4:
            continue
        if not is_recent(date_text):
            continue

        results.append({
            "title": case_text, "date": date_text,
            "court": "Securities Class Action (Stanford)",
            "summary": "Securities class action decision tracked by Stanford Clearinghouse.",
            "url": url, "pdf_url": "", "source": "Stanford Clearinghouse",
        })

    log(f"  Stanford Clearinghouse: {len(results)} items")
    return results


def scrape_courtlistener_fallback() -> list:
    """
    CourtListener fallback — catches anything the other sources missed,
    especially district court opinions in high-profile cases.
    Focuses on the most active securities courts.
    """
    results = []
    since = cutoff_date().isoformat()
    token = os.environ.get("COURTLISTENER_API_KEY", "")
    cl_headers = dict(HEADERS)
    if token:
        cl_headers["Authorization"] = f"Token {token}"

    queries = ["securities fraud class action", "10b-5", "PSLRA"]
    courts  = ["nysd", "casd", "dcd", "ded", "cand", "nyed", "flsd"]

    for query in queries:
        params = [("q", query), ("type", "o"), ("filed_after", since),
                  ("order_by", "dateFiled desc")]
        for c in courts:
            params.append(("court_id", c))
        try:
            r = requests.get(
                "https://www.courtlistener.com/api/rest/v4/search/",
                params=params, headers=cl_headers, timeout=10,
            )
            r.raise_for_status()
            for item in r.json().get("results", []):
                case_name = item.get("caseName") or item.get("case_name") or ""
                if not case_name:
                    continue
                url = ""
                if item.get("absolute_url"):
                    url = f"https://www.courtlistener.com{item['absolute_url']}"
                results.append({
                    "title":   case_name,
                    "date":    item.get("dateFiled") or item.get("date_filed") or "",
                    "court":   item.get("court") or item.get("court_id") or "Federal District Court",
                    "summary": "",
                    "url":     url,
                    "pdf_url": "",
                    "cluster_id": item.get("cluster_id"),
                    "source":  "CourtListener",
                })
        except Exception as e:
            log(f"  CourtListener fallback error ({query}): {e}")

    # Deduplicate by normalized title
    seen, unique = set(), []
    for item in results:
        k = normalize(item["title"])
        if k not in seen:
            seen.add(k)
            unique.append(item)

    log(f"  CourtListener fallback: {len(unique)} items")
    return unique


def gather_all_candidates(posted_log: dict) -> list:
    """Run all scrapers and return a deduplicated candidate list."""
    log("Stage 1: Gathering candidates from all sources...")

    all_items = []

    # Supreme Court
    all_items += scrape_scotusblog()
    time.sleep(1)
    all_items += scrape_supremecourt_gov()
    time.sleep(1)

    # Circuits with RSS feeds
    rss_circuits = [
        ("2nd Circuit",  "https://www.ca2.uscourts.gov/decisions/isysquery/0a85b038-e9a0-4d52-9b82-55e12e0b29d8/1/doc/Opinions_RSS.xml"),
        ("4th Circuit",  "https://www.ca4.uscourts.gov/rss.xml"),
        ("5th Circuit",  "https://www.ca5.uscourts.gov/rss.aspx"),
        ("9th Circuit",  "https://www.ca9.uscourts.gov/rss/opinions.xml"),
        ("11th Circuit", "https://www.ca11.uscourts.gov/rss.xml"),
    ]
    for name, url in rss_circuits:
        all_items += scrape_rss_circuit(name, url)
        time.sleep(0.5)

    # HTML scrapers
    all_items += scrape_first_circuit()
    time.sleep(0.5)
    all_items += scrape_dc_circuit()
    time.sleep(0.5)
    all_items += scrape_third_circuit()
    time.sleep(0.5)

    # Supplementary
    all_items += scrape_stanford_clearinghouse()
    time.sleep(0.5)

    # CourtListener fallback
    all_items += scrape_courtlistener_fallback()

    # Global deduplication by normalized title + url
    seen_titles, seen_urls, unique = set(), set(), []
    for item in all_items:
        t = normalize(item.get("title", ""))
        u = item.get("url", "")
        if not t or len(t) < 5:
            continue
        if t in seen_titles:
            continue
        if u and u in seen_urls:
            continue
        if is_duplicate(item["title"], u, posted_log):
            log(f"  Skip (already posted): {item['title'][:60]}")
            continue
        seen_titles.add(t)
        if u:
            seen_urls.add(u)
        unique.append(item)

    log(f"Total unique candidates after dedup: {len(unique)}")
    return unique


# ── Stage 2: Claude picks the most significant opinion ───────────────────────

def pick_most_significant(candidates: list) -> dict | None:
    """
    Ask Claude to read all candidate headlines and pick the single most
    significant opinion for a litigation-focused audience.
    """
    if not candidates:
        return None

    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    # Build a numbered list of candidates for Claude to evaluate
    lines = []
    for i, c in enumerate(candidates):
        line = f"{i+1}. [{c['court']}] {c['title']}"
        if c.get("summary"):
            line += f" — {c['summary'][:150]}"
        if c.get("date"):
            line += f" (filed: {c['date'][:20]})"
        lines.append(line)

    candidate_list = "\n".join(lines)

    prompt = textwrap.dedent(f"""
        You are a senior litigation attorney and legal editor. Below is a list of
        recent federal court opinions from the past week. Your job is to identify
        the single most significant opinion for a general litigation audience.

        Prioritize opinions that:
        - Come from the Supreme Court or circuit courts (over district courts)
        - Resolve a circuit split or establish a new legal standard
        - Are en banc decisions
        - Reverse a lower court on an important legal question
        - Involve securities fraud, antitrust, administrative law, or other
          high-stakes areas affecting many litigants
        - Are generating buzz in the legal community (SCOTUSblog coverage is a
          strong signal)

        Deprioritize:
        - Routine affirmances
        - Highly fact-specific decisions with little broader impact
        - Criminal cases (unless the legal issue is broadly significant)
        - Unpublished opinions

        CANDIDATE OPINIONS:
        {candidate_list}

        Respond ONLY with a JSON object, no markdown, no explanation:
        {{
          "selected_index": <integer, 1-based index from the list above>,
          "reason": "<one sentence explaining why this is most significant>"
        }}
    """).strip()

    try:
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=256,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = message.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        result = json.loads(raw)
        idx = int(result["selected_index"]) - 1
        reason = result.get("reason", "")
        if 0 <= idx < len(candidates):
            selected = candidates[idx]
            log(f"Claude selected: {selected['title'][:70]}")
            log(f"Reason: {reason}")
            return selected
        else:
            log(f"Claude returned out-of-range index {idx+1}")
            return candidates[0]
    except Exception as e:
        log(f"Significance scoring error: {e} — falling back to first candidate")
        return candidates[0] if candidates else None


# ── Stage 3: Fetch full opinion text ─────────────────────────────────────────

def fetch_full_text(candidate: dict) -> str:
    """
    Fetch the full opinion text. Priority:
    1. Direct PDF from the court (pdf_url)
    2. CourtListener cluster lookup (if cluster_id present)
    3. Fetch the candidate URL and extract text
    """
    text = ""

    # Option 1: direct PDF
    if candidate.get("pdf_url"):
        log(f"  Fetching PDF: {candidate['pdf_url'][:80]}")
        r = get(candidate["pdf_url"], timeout=20)
        if r and r.content:
            text = extract_pdf_text(r.content)
            if len(text) > 500:
                log(f"  PDF text extracted: {len(text)} chars")
                return text

    # Option 2: CourtListener cluster
    if candidate.get("cluster_id"):
        token = os.environ.get("COURTLISTENER_API_KEY", "")
        cl_headers = dict(HEADERS)
        if token:
            cl_headers["Authorization"] = f"Token {token}"
        try:
            r = requests.get(
                "https://www.courtlistener.com/api/rest/v4/opinions/",
                params={"cluster": candidate["cluster_id"]},
                headers=cl_headers, timeout=10,
            )
            r.raise_for_status()
            ops = r.json().get("results", [])
            if ops:
                op = ops[0]
                raw = (op.get("html_with_citations") or
                       op.get("plain_text") or
                       op.get("html") or "")
                text = strip_html(raw)
                if len(text) > 500:
                    log(f"  CourtListener text: {len(text)} chars")
                    return text
        except Exception as e:
            log(f"  CourtListener cluster fetch error: {e}")

    # Option 3: Fetch the URL directly and extract text
    if candidate.get("url") and not candidate["url"].endswith(".pdf"):
        log(f"  Fetching URL: {candidate['url'][:80]}")
        r = get(candidate["url"], timeout=15)
        if r:
            text = strip_html(r.text)
            # Trim boilerplate — keep up to 15k chars from the middle
            if len(text) > 15000:
                text = text[2000:17000]
            if len(text) > 200:
                log(f"  URL text extracted: {len(text)} chars")
                return text

    log("  Warning: could not retrieve full opinion text")
    return ""


# ── Stage 4: Claude writes the post ──────────────────────────────────────────

def build_post_with_claude(candidate: dict, opinion_text: str) -> dict | None:
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    case_name  = candidate.get("title", "Unknown")
    court      = candidate.get("court", "")
    date_filed = candidate.get("date", str(datetime.date.today()))
    source_url = candidate.get("url") or candidate.get("pdf_url") or ""

    truncated = opinion_text[:10000] if opinion_text else (
        f"[Full text unavailable — Case: {case_name}, Court: {court}]"
    )

    prompt = textwrap.dedent(f"""
        You are a litigation attorney writing a blog post for a professional
        audience of litigators. Your writing is clear, precise, and analytically
        rigorous. Avoid filler phrases. Write in complete paragraphs only.

        Write a blog post about this recent court opinion.

        CASE INFORMATION:
        Case Name:  {case_name}
        Court:      {court}
        Date:       {date_filed}
        Source URL: {source_url}

        OPINION TEXT (may be truncated or unavailable):
        ---
        {truncated}
        ---

        Cover these three sections (400-500 words total):

        1. Background — Parties, what the dispute is about, how it arrived
           at this court. Two to three sentences.

        2. The Court's Holding — What the court decided and on what legal
           grounds. Be precise. Two to three sentences.

        3. Why It Matters — Practical significance for litigators: does this
           resolve a circuit split, tighten a pleading standard, affect class
           certification, change how practitioners should approach similar cases?
           Two to three sentences.

        End with this exact HTML (substitute real values):
        <p class="case-link">Read the full opinion: <a href="{source_url}" target="_blank" rel="noopener">{case_name}</a></p>

        Use <h3> for section headers. Use <p> for paragraphs. No bullet points.
        Do not put the title inside the body.

        If the opinion text is unavailable or too truncated to write accurately,
        still produce the post based on the case name and court — write what can
        reasonably be inferred and note that the full opinion should be consulted.

        Respond ONLY with valid JSON, no markdown fences, no preamble:
        {{
          "title": "Descriptive headline capturing the legal significance",
          "court_display": "Short court label e.g. 9th Cir., S.D.N.Y., SCOTUS",
          "date_display": "Month DD, YYYY",
          "summary": "Two sentences: what was decided and why it matters.",
          "body_html": "<h3>Background</h3><p>...</p>..."
        }}
    """).strip()

    try:
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2048,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = message.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

        post = json.loads(raw)
        required = {"title", "court_display", "date_display", "summary", "body_html"}
        if not required.issubset(post.keys()):
            log(f"Claude response missing keys: {list(post.keys())}")
            return None
        log(f"Post generated: {post['title'][:80]}")
        return post

    except json.JSONDecodeError as e:
        log(f"Failed to parse Claude JSON: {e}\nRaw:\n{raw[:400]}")
        return None
    except Exception as e:
        log(f"Anthropic API error: {e}")
        return None


# ── Stage 5: HTML injection ───────────────────────────────────────────────────

def inject_post_into_html(html: str, post: dict) -> str:
    body_escaped = (
        post["body_html"]
        .replace("\\", "\\\\")
        .replace("`", "\\`")
        .replace("${", "\\${")
    )
    title_safe   = post["title"].replace('"', "'")
    summary_safe = post["summary"].replace('"', "'")

    new_js = (
        "      {\n"
        f'        date: "{post["date_display"]}",\n'
        f'        court: "{post["court_display"]}",\n'
        f'        title: "{title_safe}",\n'
        f'        summary: "{summary_safe}",\n'
        f'        body: `{body_escaped}`\n'
        "      },\n"
        "      // NEXT_POST_HERE\n"
    )

    if "// NEXT_POST_HERE" in html:
        html = html.replace("// NEXT_POST_HERE\n", new_js, 1)
    else:
        html = html.replace(
            "    const BLOG_POSTS = [\n",
            "    const BLOG_POSTS = [\n      // NEXT_POST_HERE\n", 1
        )
        html = html.replace("// NEXT_POST_HERE\n", new_js, 1)

    def bump(m):
        return f'data-post="{int(m.group(1)) + 1}"'
    html = re.sub(r'data-post="(\d+)"', bump, html)

    new_card = (
        f'            <article class="blog-card" data-post="0">\n'
        f'              <div class="blog-card-inner">\n'
        f'                <div>\n'
        f'                  <div class="blog-card-meta">\n'
        f'                    <span class="blog-date">{post["date_display"]}</span>\n'
        f'                    <span class="blog-badge">{post["court_display"]}</span>\n'
        f'                  </div>\n'
        f'                  <h3>{post["title"]}</h3>\n'
        f'                  <p>{post["summary"]}</p>\n'
        f'                </div>\n'
        f'                <div class="blog-arrow">&#8594;</div>\n'
        f'              </div>\n'
        f'            </article>\n'
        f'            <!-- NEXT_CARD_HERE -->\n'
    )

    if "<!-- NEXT_CARD_HERE -->" in html:
        html = html.replace("<!-- NEXT_CARD_HERE -->", new_card, 1)
    else:
        html = html.replace(
            '<div class="blog-list" id="blog-list">',
            '<div class="blog-list" id="blog-list">\n            <!-- NEXT_CARD_HERE -->', 1
        )
        html = html.replace("<!-- NEXT_CARD_HERE -->", new_card, 1)

    html = re.sub(
        r'\s*<div class="blog-coming-soon">.*?</div>\s*',
        "\n", html, flags=re.DOTALL
    )
    return html


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    log("Starting build_post.py (revised sourcing)")

    if not os.path.exists(HTML_PATH):
        log(f"ERROR: {HTML_PATH} not found. Running from repo root?")
        return

    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    posted_log = load_posted_log()
    log(f"Posted log: {len(posted_log['titles'])} titles on record")

    # Stage 1: gather candidates
    candidates = gather_all_candidates(posted_log)
    if not candidates:
        log("No candidates found from any source. Exiting.")
        return

    # Stage 2: Claude picks the most significant
    selected = pick_most_significant(candidates)
    if not selected:
        log("Significance scoring returned nothing. Exiting.")
        return

    # Stage 3: fetch full text
    log(f"Stage 3: Fetching full text for: {selected['title'][:70]}")
    opinion_text = fetch_full_text(selected)
    log(f"Full text length: {len(opinion_text)} chars")

    # Stage 4: write the post
    log("Stage 4: Generating blog post with Claude...")
    post = build_post_with_claude(selected, opinion_text)
    if not post:
        log("Post generation failed. Exiting.")
        return

    # Stage 5: inject into HTML
    updated_html = inject_post_into_html(html, post)
    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(updated_html)

    posted_log = record_posted(posted_log, selected["title"], selected.get("url", ""))
    save_posted_log(posted_log)

    log(f"SUCCESS: '{post['title']}'")


if __name__ == "__main__":
    main()
