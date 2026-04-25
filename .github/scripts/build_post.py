"""
build_post.py
=============
Runs inside GitHub Actions every weekday morning.

Pipeline:
  1. Query CourtListener SEARCH API (/api/rest/v4/search/?type=o&q=...)
     for opinions filed in the last 14 days matching securities law terms
  2. Filter to confirmed private civil securities litigation
  3. Reject duplicates via posted_cases.json
  4. Score and pick the best candidate
  5. Fetch full opinion text via the opinions endpoint
  6. Send to Claude for a structured 800-1000 word blog post
  7. Inject into index.html and update posted_cases.json
"""

import os
import re
import json
import datetime
import textwrap
import requests
import anthropic

# ── Config ────────────────────────────────────────────────────────────────────

COURTLISTENER_BASE  = "https://www.courtlistener.com/api/rest/v4"
COURTLISTENER_SEARCH = "https://www.courtlistener.com/api/rest/v4/search/"

# All federal courts: Supreme Court, all circuits, all 94 district courts.
# CourtListener court IDs follow the pattern: scotus, ca1-ca11, cadc, cafc,
# and district courts use abbreviations like nysd, cand, txsd, etc.
TARGET_COURTS = [
    # Supreme Court
    "scotus",
    # Circuit Courts of Appeals
    "ca1","ca2","ca3","ca4","ca5","ca6","ca7","ca8","ca9","ca10","ca11",
    "cadc",   # D.C. Circuit
    "cafc",   # Federal Circuit
    # District Courts — all 94
    # Alabama
    "almd","alnd","alsd",
    # Alaska
    "akd",
    # Arizona
    "azd",
    # Arkansas
    "ared","arwd",
    # California
    "cacd","caed","cand","casd",
    # Colorado
    "cod",
    # Connecticut
    "ctd",
    # Delaware
    "ded",
    # Florida
    "flmd","flnd","flsd",
    # Georgia
    "gamd","gand","gasd",
    # Hawaii
    "hid",
    # Idaho
    "idd",
    # Illinois
    "ilcd","ilnd","ilsd",
    # Indiana
    "innd","insd",
    # Iowa
    "iand","iasd",
    # Kansas
    "ksd",
    # Kentucky
    "kyed","kywd",
    # Louisiana
    "laed","lamd","lawd",
    # Maine
    "med",
    # Maryland
    "mdd",
    # Massachusetts
    "mad",
    # Michigan
    "mied","miwd",
    # Minnesota
    "mnd",
    # Mississippi
    "msnd","mssd",
    # Missouri
    "moed","mowd",
    # Montana
    "mtd",
    # Nebraska
    "ned",
    # Nevada
    "nvd",
    # New Hampshire
    "nhd",
    # New Jersey
    "njd",
    # New Mexico
    "nmd",
    # New York
    "nyed","nynd","nysd","nywd",
    # North Carolina
    "nced","ncmd","ncwd",
    # North Dakota
    "ndd",
    # Ohio
    "ohnd","ohsd",
    # Oklahoma
    "oked","oknd","okwd",
    # Oregon
    "ord",
    # Pennsylvania
    "paed","pamd","pawd",
    # Rhode Island
    "rid",
    # South Carolina
    "scd",
    # South Dakota
    "sdd",
    # Tennessee
    "tned","tnmd","tnwd",
    # Texas
    "txed","txnd","txsd","txwd",
    # Utah
    "utd",
    # Vermont
    "vtd",
    # Virginia
    "vaed","vawd",
    # Washington
    "waed","wawd",
    # West Virginia
    "wvnd","wvsd",
    # Wisconsin
    "wied","wiwd",
    # Wyoming
    "wyd",
    # D.C.
    "dcd",
    # Territories
    "prd","vid","gud","nmid",
]

# Single-term queries cast a wide net; the confirm/disqualify filters
# downstream ensure we only write about civil securities cases.
# Using broader terms avoids the problem of requiring two rare phrases
# to co-occur in the same recently-filed opinion.
SEARCH_QUERIES = [
    "10b-5",
    "PSLRA",
    "securities fraud class action",
    "section 11 securities act",
    "loss causation scienter",
]

# STRONG patterns — at least one of these must match.
# These are specific to federal securities litigation and unlikely
# to appear in unrelated cases.
CONFIRM_PATTERNS_STRONG = [
    r"10b-5",
    r"rule 10b-5",
    r"10\(b\)",
    r"78j",                                  # Exchange Act § 10(b) US Code cite
    r"securities exchange act of 1934",
    r"77[kl]",                               # Securities Act §§ 11/12 US Code cite
    r"securities act of 1933",
    r"pslra",
    r"private securities litigation reform",
    r"fraud on the market",
    r"loss causation",
    r"section 20\(a\).*securities",
    r"securities.*section 20\(a\)",
]

# Any match here disqualifies the opinion
DISQUALIFY_PATTERNS = [
    r"\bsec v\.\b",
    r"securities and exchange commission v\.",
    r"v\. sec\b",                            # SEC as defendant e.g. Smith v. SEC
    r"v\. securities and exchange commission",
    r"petition.*sec\b",                       # petitions for review of SEC orders
    r"\bsec order\b",
    r"united states v\.",
    r"\bindictment\b",
    r"\bgrand jury\b",
    r"finra arbitration",
    r"\barbitration award\b",
    r"\bsocial security\b",
    r"unemployment.*benefit",
]

LOOKBACK_DAYS    = 30
HTML_PATH        = "index.html"
POSTED_LOG_PATH  = "posted_cases.json"


# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg):
    print(f"[build_post] {msg}")


def cl_headers() -> dict:
    hdrs = {"User-Agent": "MLegalTechnology blog builder (mlegaltechnology.com)"}
    token = os.environ.get("COURTLISTENER_API_KEY", "")
    if token:
        hdrs["Authorization"] = f"Token {token}"
    else:
        log("Warning: COURTLISTENER_API_KEY not set")
    return hdrs


# ── Duplicate tracking ────────────────────────────────────────────────────────

def load_posted_log() -> dict:
    if os.path.exists(POSTED_LOG_PATH):
        try:
            with open(POSTED_LOG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                data.setdefault("opinion_ids", [])
                data.setdefault("cluster_ids", [])
                data.setdefault("case_names", [])
                return data
        except Exception as e:
            log(f"Warning: could not read {POSTED_LOG_PATH}: {e}")
    return {"opinion_ids": [], "cluster_ids": [], "case_names": []}


def save_posted_log(data: dict):
    with open(POSTED_LOG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    log(f"Posted log saved: {len(data['opinion_ids'])} entries")


def normalize(name: str) -> str:
    name = name.lower()
    name = re.sub(r"[^\w\s]", " ", name)
    return re.sub(r"\s+", " ", name).strip()


def is_already_posted(result: dict, log_data: dict) -> bool:
    # Search results use cluster_id; opinions endpoint uses id
    cluster_id = result.get("cluster_id")
    if cluster_id and cluster_id in log_data["cluster_ids"]:
        return True
    case_name = normalize(result.get("caseName") or result.get("case_name") or "")
    if case_name and case_name in log_data["case_names"]:
        return True
    return False


def record_posted(log_data: dict, result: dict, opinion_id=None) -> dict:
    cluster_id = result.get("cluster_id")
    if cluster_id and cluster_id not in log_data["cluster_ids"]:
        log_data["cluster_ids"].append(cluster_id)
    if opinion_id and opinion_id not in log_data["opinion_ids"]:
        log_data["opinion_ids"].append(opinion_id)
    case_name = normalize(result.get("caseName") or result.get("case_name") or "")
    if case_name and case_name not in log_data["case_names"]:
        log_data["case_names"].append(case_name)
    return log_data


# ── CourtListener search ──────────────────────────────────────────────────────

def fetch_search_results() -> list:
    """
    Use the CourtListener SEARCH endpoint with type=o (opinions).
    Parameters mirror what the CourtListener front-end passes as GET params.
    The court filter uses court_id repeated per court (OR logic).
    Date filter is filed_after (matches the front-end parameter name).
    """
    since = (datetime.date.today() - datetime.timedelta(days=LOOKBACK_DAYS)).isoformat()
    headers = cl_headers()
    results = []

    for query in SEARCH_QUERIES:
        # Build params as a list of tuples so we can repeat court_id
        params = [
            ("q",           query),
            ("type",        "o"),
            ("filed_after", since),
            ("order_by",    "dateFiled desc"),
        ]
        # Each court gets its own court_id parameter (OR logic in CL search)
        for court in TARGET_COURTS:
            params.append(("court_id", court))

        try:
            r = requests.get(
                COURTLISTENER_SEARCH,
                params=params,
                headers=headers,
                timeout=10,
            )
            # Log the actual URL so we can debug parameter issues
            log(f"  Request URL: {r.url[:120]}")
            if r.status_code != 200:
                log(f"  HTTP {r.status_code}: {r.text[:200]}")
                r.raise_for_status()
            data = r.json()
            hits = data.get("results", [])
            log(f"  '{query[:55]}': {len(hits)} results")
            results.extend(hits)
        except Exception as e:
            log(f"  Warning: search failed for '{query[:55]}': {e}")

    # Deduplicate by cluster_id
    seen, unique = set(), []
    for item in results:
        cid = item.get("cluster_id") or item.get("absolute_url", "")
        if cid not in seen:
            seen.add(cid)
            unique.append(item)

    log(f"Total unique search results: {len(unique)}")
    return unique


def fetch_opinion_text(cluster_id: int) -> tuple:
    """
    Given a cluster_id from search results, fetch the opinions for that
    cluster and return (opinion_id, text).
    Uses html_with_citations as the preferred text field per CL docs.
    """
    headers = cl_headers()
    try:
        r = requests.get(
            f"{COURTLISTENER_BASE}/opinions/",
            params={"cluster": cluster_id},
            headers=headers,
            timeout=10,
        )
        r.raise_for_status()
        opinions = r.json().get("results", [])
        if not opinions:
            return None, ""

        # Prefer the lead/combined opinion
        op = opinions[0]
        opinion_id = op.get("id")
        text = (
            op.get("html_with_citations") or
            op.get("plain_text") or
            op.get("html") or
            ""
        )
        # Strip HTML tags for cleaner text to send Claude
        text = re.sub(r"<[^>]+>", " ", text)
        text = re.sub(r"\s+", " ", text).strip()
        return opinion_id, text

    except Exception as e:
        log(f"  Warning: could not fetch opinion text for cluster {cluster_id}: {e}")
        return None, ""


# ── Filtering and scoring ─────────────────────────────────────────────────────

def is_securities_civil_case(result: dict, text: str) -> bool:
    case_name = (result.get("caseName") or result.get("case_name") or "")
    haystack = (case_name + " " + text[:8000]).lower()

    # Must match at least one STRONG federal securities law marker
    confirmed = any(
        re.search(p, haystack, re.IGNORECASE) for p in CONFIRM_PATTERNS_STRONG
    )
    if not confirmed:
        log(f"  Filtered (no federal securities markers): {case_name[:60]}")
        return False

    # Must not match any disqualifying pattern
    for p in DISQUALIFY_PATTERNS:
        if re.search(p, haystack, re.IGNORECASE):
            log(f"  Filtered (disqualified '{p}'): {case_name[:60]}")
            return False

    return True


def score_result(result: dict, text: str) -> int:
    score = 0
    court = (result.get("court_id") or result.get("court") or "").lower()
    haystack = (
        (result.get("caseName") or "") + " " + text[:5000]
    ).lower()

    if re.match(r"^ca\d{1,2}$", court):
        score += 40
    score += min(len(text) // 500, 25)

    priority = {
        "class certification": 20, "motion to dismiss": 12,
        "summary judgment": 12,    "scienter": 10,
        "loss causation": 10,      "safe harbor": 10,
        "fraud on the market": 10, "section 11": 8,
        "section 12": 8,           "control person": 8,
        "material misrepresentation": 8, "pleading standard": 5,
    }
    for term, pts in priority.items():
        if term in haystack:
            score += pts
    return score


def pick_best(results: list, posted_log: dict):
    candidates = []
    for res in results[:25]:
        case_name = res.get("caseName") or res.get("case_name") or "Unknown"

        if is_already_posted(res, posted_log):
            log(f"  Skip (duplicate): {case_name[:60]}")
            continue

        cluster_id = res.get("cluster_id")
        if not cluster_id:
            log(f"  Skip (no cluster_id): {case_name[:60]}")
            continue

        # Fetch full opinion text FIRST, then filter against it.
        # Search result snippets are too short to reliably confirm
        # federal securities law markers.
        opinion_id, text = fetch_opinion_text(cluster_id)

        if not is_securities_civil_case(res, text):
            continue

        score = score_result(res, text)
        log(f"  Candidate score={score:3d}: {case_name[:60]}")
        candidates.append((score, res, text, opinion_id))

    if not candidates:
        return None, None, None

    candidates.sort(key=lambda x: x[0], reverse=True)
    score, best_res, best_text, best_op_id = candidates[0]
    log(f"Selected (score {score}): "
        f"{(best_res.get('caseName') or best_res.get('case_name', ''))}")
    return best_res, best_text, best_op_id


# ── Claude post generation ────────────────────────────────────────────────────

def build_post_with_claude(result: dict, opinion_text: str) -> dict | None:
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    case_name  = result.get("caseName") or result.get("case_name") or "Unknown"
    court      = result.get("court") or result.get("court_id") or ""
    date_filed = result.get("dateFiled") or result.get("date_filed") or str(datetime.date.today())
    docket_num = result.get("docketNumber") or result.get("docket_number") or ""

    truncated = opinion_text[:12000] if opinion_text else (
        f"[Full text unavailable. Case: {case_name}, Court: {court}, Filed: {date_filed}]"
    )

    case_url = f"https://www.courtlistener.com{result.get('absolute_url', '')}"

    prompt = textwrap.dedent(f"""
        You are a securities litigation attorney writing a blog post for a professional
        audience of litigators. Your writing is clear, precise, and analytically rigorous.
        Avoid filler phrases. Write in complete paragraphs, not bullet points.

        Write a blog post about the following private civil securities litigation opinion.

        CASE INFORMATION:
        Case Name:     {case_name}
        Court:         {court.upper()}
        Date Filed:    {date_filed}
        Docket Number: {docket_num}
        Case URL:      {case_url}

        OPINION TEXT (may be truncated):
        ---
        {truncated}
        ---

        Cover these three sections (400-500 words total):

        1. Background - Parties, alleged conduct, and stage of litigation (2-3 sentences)
        2. The Court's Holding - What was decided and on what grounds (2-3 sentences)
        3. Why It Matters - Key legal implications for practitioners, in plain terms (2-3 sentences)

        End the post with this exact HTML, substituting the real URL:
        <p class="case-link">Read the full opinion: <a href="{case_url}" target="_blank" rel="noopener">{case_name}</a></p>

        Use <h3> tags for headers. Use <p> tags for paragraphs.
        No bullet points. No title inside the body.

        Respond ONLY with a valid JSON object, no markdown, no preamble:
        {{
          "title": "Descriptive headline capturing the legal significance",
          "court_display": "Short label e.g. S.D.N.Y., 9th Cir., D. Del.",
          "date_display": "Month DD, YYYY",
          "summary": "Exactly two sentences: holding and why practitioners should care.",
          "body_html": "<h3>Background</h3><p>...</p><h3>The Court's Holding</h3><p>...</p><h3>Why It Matters</h3><p>...</p><p class=\"case-link\">Read the full opinion: <a href=\"{case_url}\" target=\"_blank\" rel=\"noopener\">{case_name}</a></p>"
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
        log(f"Failed to parse Claude JSON: {e}\nRaw:\n{raw[:500]}")
        return None
    except Exception as e:
        log(f"Anthropic API error: {e}")
        return None


# ── HTML injection ────────────────────────────────────────────────────────────

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
    log("Starting build_post.py")

    if not os.path.exists(HTML_PATH):
        log(f"ERROR: {HTML_PATH} not found. Running from repo root?")
        return

    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    posted_log = load_posted_log()
    log(f"Duplicate log: {len(posted_log['cluster_ids'])} cluster IDs, "
        f"{len(posted_log['case_names'])} case names on record")

    results = fetch_search_results()
    if not results:
        log("No results from CourtListener search. Exiting.")
        return

    result, opinion_text, opinion_id = pick_best(results, posted_log)
    if result is None:
        log("No suitable opinion found after filtering. Exiting.")
        return

    log(f"Opinion text: {len(opinion_text)} chars")

    post = build_post_with_claude(result, opinion_text)
    if not post:
        log("Post generation failed. Exiting.")
        return

    updated_html = inject_post_into_html(html, post)
    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(updated_html)

    posted_log = record_posted(posted_log, result, opinion_id)
    save_posted_log(posted_log)

    log(f"SUCCESS: '{post['title']}'")


if __name__ == "__main__":
    main()
