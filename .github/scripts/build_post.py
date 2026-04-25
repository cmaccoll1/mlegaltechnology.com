"""
build_post.py
=============
Runs inside GitHub Actions every weekday morning.

Pipeline:
  1. Query CourtListener for opinions filed in the last 14 days
  2. Filter to confirmed private civil securities litigation only
     (Exchange Act § 10(b)/Rule 10b-5, § 20(a), Securities Act §§ 11/12/15,
      state blue sky equivalents)
  3. Reject any opinion whose CourtListener ID or docket number appears
     in posted_cases.json (persistent duplicate log)
  4. Score remaining candidates and pick the best one
  5. Fetch the full opinion text
  6. Send to Claude with a detailed legal-writing prompt
  7. Inject the post into index.html
  8. Append the case identifiers to posted_cases.json
"""

import os
import re
import json
import datetime
import textwrap
import requests
import anthropic

# ── Config ────────────────────────────────────────────────────────────────────

COURTLISTENER_BASE = "https://www.courtlistener.com/api/rest/v4"

# Federal courts most active in securities litigation
TARGET_COURTS = [
    "ca1", "ca2", "ca3", "ca4", "ca5", "ca6", "ca7", "ca8", "ca9", "ca10", "ca11",
    "nysd", "casd", "dcd", "ilnd", "txsd", "mad", "ded", "njd", "cand", "nyed", "flsd", "vaed",
]

# Search queries targeted at private civil securities litigation
SEARCH_QUERIES = [
    "Rule 10b-5 class action",
    "section 10(b) securities exchange act",
    "securities act section 11",
    "securities act section 12",
    "PSLRA scienter",
    "loss causation securities fraud",
    "section 20(a) control person liability",
    "fraud on the market presumption",
    "safe harbor PSLRA forward-looking",
    "class certification securities fraud",
]

# At least one of these must appear for an opinion to be considered
CONFIRM_PATTERNS = [
    r"rule 10b-5",
    r"section 10\(b\)",
    r"§ 10\(b\)",
    r"15 u\.s\.c\.? [§s]+ 78j",
    r"securities exchange act of 1934",
    r"section 11\b",
    r"section 12\(a\)",
    r"15 u\.s\.c\.? [§s]+ 77[kl]",
    r"securities act of 1933",
    r"section 20\(a\)",
    r"pslra",
    r"private securities litigation reform act",
    r"fraud on the market",
    r"basic inc\.",
    r"blue sky",
    r"loss causation",
    r"scienter",
    r"material misrepresentation",
    r"class action.*securities",
    r"securities.*class action",
]

# Any match here disqualifies the opinion
DISQUALIFY_PATTERNS = [
    r"\bsec v\.\b",
    r"securities and exchange commission v\.",
    r"united states v\.",
    r"\bindictment\b",
    r"\bcriminal\b.*\bsecurities\b",
    r"\bgrand jury\b",
    r"department of justice",
    r"finra arbitration",
    r"\barbitration award\b",
    r"workers[' ]compensation",
    r"social security",
    r"unemployment.*benefit",
]

LOOKBACK_DAYS   = 14
HTML_PATH       = "index.html"
POSTED_LOG_PATH = "posted_cases.json"


# ── Duplicate tracking ────────────────────────────────────────────────────────

def log(msg):
    print(f"[build_post] {msg}")


def load_posted_log() -> dict:
    if os.path.exists(POSTED_LOG_PATH):
        try:
            with open(POSTED_LOG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                data.setdefault("opinion_ids", [])
                data.setdefault("docket_ids", [])
                data.setdefault("case_names", [])
                return data
        except Exception as e:
            log(f"Warning: could not read {POSTED_LOG_PATH}: {e}")
    return {"opinion_ids": [], "docket_ids": [], "case_names": []}


def save_posted_log(log_data: dict) -> None:
    with open(POSTED_LOG_PATH, "w", encoding="utf-8") as f:
        json.dump(log_data, f, indent=2)
    log(f"Posted log saved: {len(log_data['opinion_ids'])} opinion IDs on record")


def record_posted_case(log_data: dict, opinion: dict) -> dict:
    op_id = opinion.get("id")
    if op_id and op_id not in log_data["opinion_ids"]:
        log_data["opinion_ids"].append(op_id)

    docket_id = opinion.get("docket_id") or opinion.get("docket")
    if docket_id and docket_id not in log_data["docket_ids"]:
        log_data["docket_ids"].append(docket_id)

    case_name = normalize_case_name(opinion.get("case_name", ""))
    if case_name and case_name not in log_data["case_names"]:
        log_data["case_names"].append(case_name)

    return log_data


def normalize_case_name(name: str) -> str:
    name = name.lower()
    name = re.sub(r"[^\w\s]", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def is_already_posted(opinion: dict, log_data: dict) -> bool:
    op_id = opinion.get("id")
    if op_id and op_id in log_data["opinion_ids"]:
        return True
    docket_id = opinion.get("docket_id") or opinion.get("docket")
    if docket_id and docket_id in log_data["docket_ids"]:
        return True
    case_name = normalize_case_name(opinion.get("case_name", ""))
    if case_name and case_name in log_data["case_names"]:
        return True
    return False


# ── CourtListener fetching ────────────────────────────────────────────────────

def cl_headers() -> dict:
    hdrs = {"User-Agent": "MLegalTechnology blog builder (mlegaltechnology.com)"}
    token = os.environ.get("COURTLISTENER_API_KEY", "")
    if token:
        hdrs["Authorization"] = f"Token {token}"
    else:
        log("Warning: COURTLISTENER_API_KEY not set — requests may be rejected")
    return hdrs


def fetch_recent_opinions() -> list:
    since = (datetime.date.today() - datetime.timedelta(days=LOOKBACK_DAYS)).isoformat()
    results = []
    headers = cl_headers()

    for query in SEARCH_QUERIES:
        params = {
            "search":      query,
            "filed_after": since,
            "court":       ",".join(TARGET_COURTS),
            "order_by":    "-date_filed",
            "page_size":   8,
        }
        try:
            r = requests.get(
                f"{COURTLISTENER_BASE}/opinions/",
                params=params,
                timeout=20,
                headers=headers,
            )
            r.raise_for_status()
            data = r.json()
            count = len(data.get("results", []))
            results.extend(data.get("results", []))
            log(f"  '{query[:55]}': {count} results")
        except Exception as e:
            log(f"  Warning: query failed for '{query[:55]}': {e}")

    seen, unique = set(), []
    for op in results:
        oid = op.get("id") or op.get("absolute_url", "")
        if oid not in seen:
            seen.add(oid)
            unique.append(op)

    log(f"Total unique raw opinions: {len(unique)}")
    return unique


def fetch_opinion_text(op: dict) -> str:
    if op.get("plain_text") and len(op["plain_text"]) > 500:
        return op["plain_text"]

    absolute_url = op.get("absolute_url", "")
    if not absolute_url:
        return ""

    headers = cl_headers()
    full_url = f"https://www.courtlistener.com{absolute_url}"
    try:
        r = requests.get(full_url, timeout=30, headers=headers)
        r.raise_for_status()
        if "/opinions/" in full_url:
            text_url = full_url.rstrip("/") + "/plain-text/"
            rt = requests.get(text_url, timeout=30, headers=headers)
            if rt.status_code == 200 and len(rt.text) > 200:
                return rt.text
    except Exception as e:
        log(f"  Warning: could not fetch opinion text: {e}")

    return op.get("plain_text", "") or ""


# ── Filtering and scoring ─────────────────────────────────────────────────────

def is_securities_civil_case(op: dict, text: str) -> bool:
    haystack = (
        (op.get("case_name") or "").lower() + " " + text[:8000].lower()
    )

    confirmed = any(re.search(p, haystack, re.IGNORECASE) for p in CONFIRM_PATTERNS)
    if not confirmed:
        log(f"  Filtered (no civil securities markers): {op.get('case_name', '')[:60]}")
        return False

    for p in DISQUALIFY_PATTERNS:
        if re.search(p, haystack, re.IGNORECASE):
            log(f"  Filtered (disqualified — '{p}'): {op.get('case_name', '')[:60]}")
            return False

    return True


def score_opinion(op: dict, text: str) -> int:
    score = 0
    court = (op.get("court_id") or op.get("court", "")).lower()
    haystack = ((op.get("case_name") or "") + " " + text[:5000]).lower()

    if re.match(r"^ca\d{1,2}$", court):
        score += 40

    score += min(len(text) // 500, 25)

    priority_terms = {
        "class certification":       20,
        "motion to dismiss granted": 15,
        "motion to dismiss denied":  15,
        "summary judgment":          12,
        "scienter":                  10,
        "loss causation":            10,
        "safe harbor":               10,
        "fraud on the market":       10,
        "class action":               8,
        "section 11":                 8,
        "section 12":                 8,
        "control person":             8,
        "material misrepresentation": 8,
        "falsity":                    5,
        "pleading standard":          5,
    }
    for term, pts in priority_terms.items():
        if term in haystack:
            score += pts

    return score


def pick_opinion(opinions: list, posted_log: dict):
    candidates = []

    for op in opinions:
        case_name = op.get("case_name", "Unknown")

        if is_already_posted(op, posted_log):
            log(f"  Skip (already posted): {case_name[:60]}")
            continue

        text = fetch_opinion_text(op)

        if not is_securities_civil_case(op, text):
            continue

        score = score_opinion(op, text)
        log(f"  Candidate: score={score:3d}  {case_name[:60]}")
        candidates.append((score, op, text))

    if not candidates:
        return None, None

    candidates.sort(key=lambda x: x[0], reverse=True)
    best_score, best_op, best_text = candidates[0]
    log(f"Selected: {best_op.get('case_name')} (score {best_score})")
    return best_op, best_text


# ── Claude post generation ────────────────────────────────────────────────────

def build_post_with_claude(opinion: dict, opinion_text: str):
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    case_name  = opinion.get("case_name", "Unknown")
    court_id   = opinion.get("court_id") or opinion.get("court", "")
    date_filed = opinion.get("date_filed", str(datetime.date.today()))
    docket_num = opinion.get("docket_number", "")

    truncated_text = opinion_text[:12000] if opinion_text else (
        f"[Full text unavailable. Case: {case_name}, Court: {court_id}, Filed: {date_filed}]"
    )

    prompt = textwrap.dedent(f"""
        You are a securities litigation attorney writing a blog post for a professional
        audience of litigators. Your writing is clear, precise, and analytically rigorous.
        You do not use filler phrases like "it is worth noting" or "importantly."
        You write in complete paragraphs, not bullet points.

        Write a blog post about the following opinion in a private civil securities
        litigation matter brought under the Securities Exchange Act of 1934, the
        Securities Act of 1933, or analogous state law.

        CASE INFORMATION:
        Case Name:     {case_name}
        Court:         {court_id.upper()}
        Date Filed:    {date_filed}
        Docket Number: {docket_num}

        OPINION TEXT (may be truncated):
        ---
        {truncated_text}
        ---

        The post should be 800-1000 words and cover these four sections:

        1. Background - What is this case? Who are the parties? What conduct is alleged?
           How did it get to this ruling (i.e., what stage of litigation)?

        2. The Court's Holding - What did the court decide, and on what statutory or
           doctrinal grounds? Be precise about which claims survived or were dismissed
           and why.

        3. Legal Analysis - Walk through the court's reasoning on the key issues.
           Name the specific legal standards applied (e.g., Tellabs scienter standard,
           Dura loss causation, Basic fraud-on-the-market, PSLRA pleading requirements,
           Securities Act section 11 strict liability). Be specific and technical.

        4. Implications - What does this decision mean going forward? Does it tighten
           or loosen pleading standards? Does it create or resolve a circuit split?
           What should plaintiffs' counsel or defense counsel take from it?

        Use <h3> tags for section headers. Use <p> tags for paragraphs.
        Do not use bullet points or numbered lists anywhere in the body.
        Do not include the post title inside the body.

        Respond ONLY with a valid JSON object, no markdown fences, no preamble.
        Use this exact format:
        {{
          "title": "A descriptive headline capturing the legal significance (not just the case name)",
          "court_display": "Short court label, e.g. S.D.N.Y., 9th Cir., D. Del.",
          "date_display": "Month DD, YYYY",
          "summary": "Exactly two sentences: what the court held and why practitioners should care.",
          "body_html": "<h3>Background</h3><p>...</p><h3>The Court's Holding</h3><p>...</p>..."
        }}
    """).strip()

    try:
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2500,
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
        log(f"Failed to parse Claude JSON: {e}")
        log(f"Raw response:\n{raw[:500]}")
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

    new_js_entry = (
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
        html = html.replace("// NEXT_POST_HERE\n", new_js_entry, 1)
    else:
        html = html.replace(
            "    const BLOG_POSTS = [\n",
            "    const BLOG_POSTS = [\n      // NEXT_POST_HERE\n",
            1,
        )
        html = html.replace("// NEXT_POST_HERE\n", new_js_entry, 1)

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
            '<div class="blog-list" id="blog-list">\n            <!-- NEXT_CARD_HERE -->',
            1,
        )
        html = html.replace("<!-- NEXT_CARD_HERE -->", new_card, 1)

    html = re.sub(
        r'\s*<div class="blog-coming-soon">.*?</div>\s*',
        "\n",
        html,
        flags=re.DOTALL,
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
    log(f"Duplicate log: {len(posted_log['opinion_ids'])} opinion IDs, "
        f"{len(posted_log['docket_ids'])} docket IDs, "
        f"{len(posted_log['case_names'])} case names on record")

    opinions = fetch_recent_opinions()
    if not opinions:
        log("No opinions returned from CourtListener. Exiting.")
        return

    opinion, opinion_text = pick_opinion(opinions, posted_log)
    if opinion is None:
        log("No suitable new opinion found after filtering. Exiting without changes.")
        return

    log(f"Opinion text: {len(opinion_text)} chars")

    post = build_post_with_claude(opinion, opinion_text)
    if not post:
        log("Post generation failed. Exiting without changes.")
        return

    updated_html = inject_post_into_html(html, post)
    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(updated_html)

    posted_log = record_posted_case(posted_log, opinion)
    save_posted_log(posted_log)

    log(f"SUCCESS: '{post['title']}'")


if __name__ == "__main__":
    main()
