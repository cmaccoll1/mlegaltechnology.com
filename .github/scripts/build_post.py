"""
build_post.py
=============
Runs inside GitHub Actions every weekday morning.

Pipeline:
  1. Query CourtListener for securities-related opinions filed in the last 14 days
  2. Score and pick the most substantive one not already posted
  3. Fetch the full opinion text
  4. Send to GPT-4o with a detailed legal-writing prompt
  5. Parse the structured JSON response
  6. Prepend the new post into index.html (BLOG_POSTS array + card list)
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

# Federal courts most likely to have significant securities opinions
TARGET_COURTS = [
    "ca1", "ca2", "ca3", "ca4", "ca5", "ca6", "ca7", "ca8", "ca9", "ca10", "ca11",
    "nysd", "casd", "dcd", "ilnd", "txsd", "mad", "ded", "njd", "cand",
]

SEARCH_TERMS = [
    "securities fraud",
    "10b-5",
    "section 10(b)",
    "PSLRA",
    "class action securities",
    "insider trading",
    "securities exchange act",
    "material misrepresentation",
    "loss causation",
    "scienter",
]

# How many days back to search
LOOKBACK_DAYS = 14

# Path to the HTML file (relative to repo root, where the script runs from)
HTML_PATH = "index.html"

# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg):
    print(f"[build_post] {msg}")


def already_posted_titles(html: str) -> list[str]:
    """Extract titles already in the BLOG_POSTS array to avoid duplicates."""
    matches = re.findall(r'title:\s*"([^"]+)"', html)
    return [t.lower().strip() for t in matches]


def fetch_recent_opinions() -> list[dict]:
    """Query CourtListener for recent securities-related opinions."""
    since = (datetime.date.today() - datetime.timedelta(days=LOOKBACK_DAYS)).isoformat()
    results = []

    for term in SEARCH_TERMS[:4]:   # limit API calls; first four terms cover the core cases
        params = {
            "search":      term,
            "filed_after": since,
            "court":       ",".join(TARGET_COURTS),
            "order_by":    "-date_filed",
            "page_size":   5,
        }
        try:
            r = requests.get(
                f"{COURTLISTENER_BASE}/opinions/",
                params=params,
                timeout=20,
                headers={"User-Agent": "MLegalTechnology blog builder (mlegaltechnology.com)"},
            )
            r.raise_for_status()
            data = r.json()
            results.extend(data.get("results", []))
            log(f"  '{term}': {len(data.get('results', []))} results")
        except Exception as e:
            log(f"  Warning: CourtListener query failed for '{term}': {e}")

    # Deduplicate by opinion ID
    seen = set()
    unique = []
    for op in results:
        oid = op.get("id") or op.get("absolute_url", "")
        if oid not in seen:
            seen.add(oid)
            unique.append(op)

    log(f"Total unique opinions found: {len(unique)}")
    return unique


def score_opinion(op: dict) -> int:
    """
    Score an opinion for newsworthiness. Higher = more interesting.
    Heuristics: circuit court > district court, longer opinions, certain keywords.
    """
    score = 0
    court = (op.get("court_id") or op.get("court", "")).lower()
    case_name = (op.get("case_name") or "").lower()
    plain_text = (op.get("plain_text") or "")

    # Circuit opinions are more significant
    if court.startswith("ca") and len(court) <= 4:
        score += 30

    # Prefer substantive opinions (longer text)
    score += min(len(plain_text) // 500, 20)

    # Boost for high-value securities topics
    priority_terms = [
        "class certification", "motion to dismiss granted", "scienter",
        "loss causation", "safe harbor", "insider trading", "sec enforcement",
        "material misstatement", "fraud on the market",
    ]
    for term in priority_terms:
        if term in case_name or term in plain_text[:3000].lower():
            score += 10

    return score


def fetch_opinion_text(op: dict) -> str:
    """Fetch the full plain text of an opinion."""
    # CourtListener sometimes includes plain_text directly
    if op.get("plain_text") and len(op["plain_text"]) > 500:
        return op["plain_text"]

    # Otherwise fetch from the absolute URL
    absolute_url = op.get("absolute_url", "")
    if not absolute_url:
        return ""

    full_url = f"https://www.courtlistener.com{absolute_url}"
    try:
        r = requests.get(
            full_url,
            timeout=30,
            headers={"User-Agent": "MLegalTechnology blog builder (mlegaltechnology.com)"},
        )
        r.raise_for_status()
        # The plain text endpoint
        if "/opinions/" in full_url:
            text_url = full_url.rstrip("/") + "/plain-text/"
            rt = requests.get(text_url, timeout=30)
            if rt.status_code == 200 and len(rt.text) > 200:
                return rt.text
    except Exception as e:
        log(f"  Warning: Could not fetch full opinion text: {e}")

    return op.get("plain_text", "") or op.get("download_url", "")


def pick_opinion(opinions: list[dict], already_posted: list[str]) -> dict | None:
    """Score all candidates and return the best one not already posted."""
    scored = []
    for op in opinions:
        case_name = (op.get("case_name") or "").lower().strip()
        if any(case_name in posted for posted in already_posted):
            log(f"  Skipping (already posted): {op.get('case_name')}")
            continue
        scored.append((score_opinion(op), op))

    if not scored:
        return None

    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best_op = scored[0]
    log(f"Selected: {best_op.get('case_name')} (score {best_score})")
    return best_op


def build_post_with_claude(opinion: dict, opinion_text: str) -> dict | None:
    """
    Send the opinion to Claude and get back a structured blog post.
    Returns a dict with keys: title, court_display, date_display, summary, body_html
    """
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    case_name   = opinion.get("case_name", "Unknown")
    court_id    = opinion.get("court_id") or opinion.get("court", "")
    date_filed  = opinion.get("date_filed", str(datetime.date.today()))
    docket_num  = opinion.get("docket_number", "")

    # Truncate opinion text to stay within context limits (~12k chars is plenty)
    truncated_text = opinion_text[:12000] if opinion_text else ""
    if not truncated_text:
        truncated_text = f"[Full text not available. Case: {case_name}, Court: {court_id}, Filed: {date_filed}]"

    prompt = textwrap.dedent(f"""
        You are a securities litigation attorney writing a blog post for a professional audience
        of litigators. Your writing is clear, precise, and analytically rigorous — not academic
        or padded. You do not use filler phrases like "it is worth noting" or "importantly."
        You write in complete paragraphs, not bullet points.

        Write a blog post about the following federal court opinion in a securities litigation matter.

        CASE INFORMATION:
        Case Name: {case_name}
        Court: {court_id.upper()}
        Date Filed: {date_filed}
        Docket Number: {docket_num}

        OPINION TEXT (may be truncated):
        ---
        {truncated_text}
        ---

        The post should be 800–1000 words and cover:
        1. Procedural background — what is this case, how did it get here, what was the motion or issue before the court
        2. The court's holding — what did the court decide and on what grounds
        3. Legal analysis — walk through the court's reasoning on the key issues; be specific about the legal standards applied
        4. Implications — what does this mean for practitioners, plaintiffs, or defendants in securities litigation going forward

        Use <h3> tags for section headers. Use <p> tags for paragraphs. Do not use bullet points.
        Do not include a title in the body — the title is separate.

        Respond ONLY with a JSON object, no markdown fences, in exactly this format:
        {{
          "title": "A descriptive headline (not the case name — a real headline like a law review note)",
          "court_display": "Short court abbreviation, e.g. S.D.N.Y. or 9th Cir.",
          "date_display": "Month DD, YYYY",
          "summary": "Two sentences max. What the court held and why it matters. No more.",
          "body_html": "<h3>Background</h3><p>...</p><h3>The Court's Holding</h3>..."
        }}
    """).strip()

    try:
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = message.content[0].text.strip()

        # Strip accidental markdown fences if present
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

        post = json.loads(raw)
        required = {"title", "court_display", "date_display", "summary", "body_html"}
        if not required.issubset(post.keys()):
            log(f"Claude response missing keys. Got: {list(post.keys())}")
            return None

        log(f"Post generated: {post['title'][:80]}")
        return post

    except json.JSONDecodeError as e:
        log(f"Failed to parse Claude JSON response: {e}")
        log(f"Raw response was:\n{raw[:500]}")
        return None
    except Exception as e:
        log(f"Anthropic API error: {e}")
        return None


def inject_post_into_html(html: str, post: dict) -> str:
    """
    Insert the new post into index.html in two places:
      1. Prepend a new entry to the BLOG_POSTS JS array
      2. Prepend a new <article> card to #blog-list

    Uses clearly-marked injection points so subsequent runs stack correctly.
    """
    # ── 1. Escape the body HTML for embedding in a JS template literal ──
    body_escaped = post["body_html"].replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")

    # ── 2. Build the new BLOG_POSTS entry ──
    # We prepend before the first existing entry (or before the closing bracket)
    new_js_entry = (
        "      {{\n"
        f'        date: "{post["date_display"]}",\n'
        f'        court: "{post["court_display"]}",\n'
        f'        title: "{post["title"].replace(chr(34), chr(39))}",\n'
        f'        summary: "{post["summary"].replace(chr(34), chr(39))}",\n'
        f'        body: `{body_escaped}`\n'
        "      }},\n"
        "      // NEXT_POST_HERE\n"
    )
    # Replace the injection marker (first run uses the placeholder, subsequent runs stack)
    if "// NEXT_POST_HERE" in html:
        html = html.replace("// NEXT_POST_HERE\n", new_js_entry, 1)
    else:
        # Fallback: insert after the opening of the BLOG_POSTS array
        html = html.replace(
            "    const BLOG_POSTS = [\n",
            "    const BLOG_POSTS = [\n      // NEXT_POST_HERE\n",
            1
        )
        html = html.replace("// NEXT_POST_HERE\n", new_js_entry, 1)

    # ── 3. Build the new blog card HTML ──
    # We need to know the index this post will be at (0 = newest)
    # Every time we prepend, we bump existing indices by 1.
    # Re-number all existing data-post attributes first.
    def bump_post_indices(m):
        old_index = int(m.group(1))
        return f'data-post="{old_index + 1}"'

    html = re.sub(r'data-post="(\d+)"', bump_post_indices, html)

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
        f'                <div class="blog-arrow">→</div>\n'
        f'              </div>\n'
        f'            </article>\n'
        f'            <!-- NEXT_CARD_HERE -->\n'
    )

    if "<!-- NEXT_CARD_HERE -->" in html:
        html = html.replace("<!-- NEXT_CARD_HERE -->", new_card, 1)
    else:
        # First run: find the blog list div and insert
        html = html.replace(
            '<div class="blog-list" id="blog-list">',
            '<div class="blog-list" id="blog-list">\n            <!-- NEXT_CARD_HERE -->',
            1
        )
        html = html.replace("<!-- NEXT_CARD_HERE -->", new_card, 1)

    # ── 4. Remove the "coming soon" placeholder once first post is live ──
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

    # Load the current HTML
    if not os.path.exists(HTML_PATH):
        log(f"ERROR: {HTML_PATH} not found. Is the script running from the repo root?")
        return

    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    already_posted = already_posted_titles(html)
    log(f"Already-posted titles found: {len(already_posted)}")

    # Fetch and select an opinion
    opinions = fetch_recent_opinions()
    if not opinions:
        log("No opinions found. Exiting without changes.")
        return

    opinion = pick_opinion(opinions, already_posted)
    if not opinion:
        log("No new opinion to post. Exiting without changes.")
        return

    opinion_text = fetch_opinion_text(opinion)
    log(f"Opinion text length: {len(opinion_text)} chars")

    # Generate the post
    post = build_post_with_claude(opinion, opinion_text)
    if not post:
        log("Post generation failed. Exiting without changes.")
        return

    # Inject into HTML
    updated_html = inject_post_into_html(html, post)

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(updated_html)

    log(f"SUCCESS: index.html updated with post: {post['title']}")


if __name__ == "__main__":
    main()
