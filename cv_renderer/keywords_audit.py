"""
Keywords audit: compare CV markdown against job posting text.

Strategy:
  1. Extract candidate keywords from JD using simple frequency + capitalization
     heuristics (LLM-quality requires actual LLM, but this is solid baseline).
  2. Check presence in CV text (case-insensitive, with stemming-lite).
  3. Generate markdown report with:
     - Matched keywords (with frequency)
     - Missing keywords (worth considering)
     - CV-side terms not in JD (potential off-target content)
"""
from __future__ import annotations
import re
from collections import Counter
from pathlib import Path


# ── Stop words (common English + Spanish that pollute keyword extraction) ──
STOP = {
    # English
    "the", "a", "an", "and", "or", "but", "if", "of", "in", "on", "at", "to",
    "for", "with", "by", "from", "as", "is", "are", "was", "were", "be", "been",
    "being", "have", "has", "had", "do", "does", "did", "will", "would",
    "should", "could", "may", "might", "must", "can", "this", "that", "these",
    "those", "we", "you", "he", "she", "it", "they", "them", "their", "our",
    "your", "his", "her", "its", "what", "which", "who", "whom", "where", "when",
    "why", "how", "all", "any", "both", "each", "few", "more", "most", "other",
    "some", "such", "no", "not", "only", "own", "same", "so", "than", "too",
    "very", "just", "about", "up", "down", "out", "off", "over", "under", "again",
    "further", "then", "once", "also", "than", "into", "through", "during",
    "before", "after", "above", "below", "between", "year", "years", "time",
    "experience", "work", "working", "job", "role", "position", "company",
    "team", "teams", "ability", "able", "across", "including", "such", "etc",
    # Spanish
    "el", "la", "los", "las", "un", "una", "unos", "unas", "y", "o", "pero",
    "de", "en", "a", "con", "por", "para", "que", "como", "se", "su", "sus",
    "es", "son", "fue", "ser", "estar", "ha", "han", "del", "al", "lo", "le",
    "les", "este", "esta", "estos", "estas", "ese", "esa", "eso", "más", "menos",
}

# Bigger-than-single-words also matter: phrases like "machine learning"
PRESERVED_PHRASES = [
    "machine learning", "deep learning", "natural language processing",
    "user experience", "ux research", "ux design", "product management",
    "data science", "data engineering", "data analytics", "business intelligence",
    "market research", "customer insights", "customer experience",
    "research operations", "audience measurement", "brand lift",
    "competitive intelligence", "competitive analysis", "go to market",
    "go-to-market", "stakeholdemanagement", "cross-functional",
    "people manager", "people management", "team management", "team building",
    "head of", "vice president", "senior director", "research director",
    "head of research", "head of insights", "vp insights", "vp research",
    "consumer insights", "consumer research", "qualitative research",
    "quantitative research", "mixed methods", "ai/ml", "generative ai",
    "large language model", "prompt engineering", "fortune 500",
]


def audit_keywords(cv: dict, job_posting_text: str, output_path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    cv_text = _flatten_cv_text(cv).lower()
    jd_text = job_posting_text.lower()

    # Build exclusion set: target company name + role title tokens.
    # These are CV-target identifiers, not skills the candidate should claim.
    exclude = _build_exclusions(cv)

    jd_keywords = _extract_keywords(jd_text, source="JD")
    cv_keywords = _extract_keywords(cv_text, source="CV")

    matched: list[tuple[str, int, int]] = []   # (kw, jd_freq, cv_freq)
    missing: list[tuple[str, int]] = []         # (kw, jd_freq) — in JD not in CV
    cv_only: list[tuple[str, int]] = []         # (kw, cv_freq) — in CV not in JD
    excluded: list[tuple[str, int]] = []        # excluded from match-rate calc

    for kw, jd_n in jd_keywords.most_common(60):
        if any(ex in kw or kw in ex for ex in exclude):
            excluded.append((kw, jd_n))
            continue
        cv_n = _count_in(kw, cv_text)
        if cv_n > 0:
            matched.append((kw, jd_n, cv_n))
        else:
            missing.append((kw, jd_n))

    # CV-only (potentially off-target, low priority)
    for kw, cv_n in cv_keywords.most_common(30):
        if _count_in(kw, jd_text) == 0:
            cv_only.append((kw, cv_n))

    md = _build_report_md(cv, matched, missing, cv_only, jd_text, excluded)
    output_path.write_text(md, encoding="utf-8")
    return output_path


def _build_exclusions(cv: dict) -> set[str]:
    """Build keyword exclusion set: target company name + role + generic location words."""
    excl: set[str] = set()
    meta = cv.get("meta", {})
    for field in ("target_company", "target_role"):
        v = (meta.get(field) or "").lower()
        # Split into individual tokens; ignore stop words and tiny tokens
        for tok in re.split(r"[^a-z0-9]+", v):
            if len(tok) >= 2 and tok not in STOP:
                excl.add(tok)
    # Common location/generic noise that pollutes match-rate when JD mentions them heavily
    excl.update({"hybrid", "onsite", "remote"})
    return excl


# ─────────────────────────── Helpers ─────────────────────────────
def _flatten_cv_text(cv: dict) -> str:
    parts: list[str] = []
    header = cv.get("header", {})
    parts.append(header.get("name", ""))
    parts.append(header.get("headline", ""))
    parts.extend(header.get("contact_lines", []))
    for sec in cv.get("sections", []):
        parts.append(sec.get("title", ""))
        if sec.get("type") == "paragraph":
            parts.append(sec.get("content", ""))
        elif sec.get("type") == "experience":
            for e in sec.get("entries", []):
                parts.append(e.get("company", ""))
                parts.append(e.get("role", ""))
                parts.append(e.get("dates", ""))
                parts.append(e.get("location", ""))
                parts.extend(e.get("bullets", []))
        elif sec.get("type") == "skills":
            for cat in sec.get("categories", []):
                parts.append(cat.get("name", ""))
                parts.extend(cat.get("items", []))
        elif sec.get("type") == "list":
            parts.extend(sec.get("items", []))
    return " ".join(parts)


def _extract_keywords(text: str, source: str = "") -> Counter:
    """Extract candidate keywords from text: phrases first, then single tokens."""
    text_l = text.lower()
    counter: Counter = Counter()

    # 1. Preserved phrases (highest priority)
    for phrase in PRESERVED_PHRASES:
        n = text_l.count(phrase)
        if n > 0:
            counter[phrase] = n

    # 2. Capitalized terms from ORIGINAL casing (proper nouns, brands, tech)
    cap_terms = re.findall(r"\b[A-Z][a-zA-Z0-9+/.-]{2,}(?:\s+[A-Z][a-zA-Z0-9+/.-]+){0,2}",
                            text)
    for term in cap_terms:
        t = term.lower().strip()
        if t and t not in STOP and len(t) > 2 and not t.isdigit():
            counter[t] += 1

    # 3. Single tokens (filtered by stop words, frequency ≥ 2)
    tokens = re.findall(r"\b[a-záéíóúñ][a-záéíóúñ0-9+/.-]{2,}\b", text_l)
    token_counter = Counter(tokens)
    for tok, n in token_counter.items():
        if tok in STOP or n < 2 or tok.isdigit():
            continue
        counter[tok] = max(counter.get(tok, 0), n)

    return counter


def _count_in(needle: str, haystack: str) -> int:
    """Count occurrences of needle in haystack, with simple stemming tolerance."""
    needle = needle.lower().strip()
    if not needle:
        return 0
    # Direct match
    n = haystack.count(needle)
    if n > 0:
        return n
    # Stemming-lite: strip plural 's' / past 'ed' / ing
    for suffix in ("s", "es", "ed", "ing"):
        if needle.endswith(suffix) and len(needle) > len(suffix) + 2:
            stem = needle[:-len(suffix)]
            if stem in haystack:
                return haystack.count(stem)
    return 0


def _build_report_md(cv, matched, missing, cv_only, jd_text, excluded=None):
    excluded = excluded or []
    meta = cv.get("meta", {})
    lines = [
        f"# Keywords Audit · CV vs Job Description",
        "",
        f"- **Recipe:** {meta.get('recipe', '—')}",
        f"- **Target company:** {meta.get('target_company', '—')}",
        f"- **Target role:** {meta.get('target_role', '—')}",
        f"- **Date:** {meta.get('date', '—')}",
        f"- **JD length:** {len(jd_text)} chars · {len(jd_text.split())} words",
        "",
        "---",
        "",
        f"## ✅ Matched keywords ({len(matched)})",
        "Keywords from JD that also appear in CV. Higher CV freq → stronger signal.",
        "",
        "| Keyword | JD freq | CV freq | Signal |",
        "|---|---:|---:|---|",
    ]
    for kw, jd_n, cv_n in sorted(matched, key=lambda x: (-x[1], -x[2])):
        signal = "🟢 strong" if cv_n >= 2 else "🟡 present"
        lines.append(f"| {kw} | {jd_n} | {cv_n} | {signal} |")

    lines += [
        "",
        f"## ⚠ Missing keywords ({len(missing)})",
        "Keywords in JD NOT found in CV. Top of this list = highest-priority gaps.",
        "Consider weaving these into a bullet or skill **without inventing fake content**.",
        "",
        "| Keyword | JD freq | Priority |",
        "|---|---:|---|",
    ]
    for kw, jd_n in sorted(missing, key=lambda x: -x[1])[:25]:
        prio = "🔴 high" if jd_n >= 3 else "🟠 med" if jd_n >= 2 else "🔵 low"
        lines.append(f"| {kw} | {jd_n} | {prio} |")

    lines += [
        "",
        f"## ℹ CV-only keywords (top 15)",
        "Terms present in CV but NOT in JD. NOT a problem per se — could be unique strengths.",
        "Review if any are off-target noise that could be trimmed.",
        "",
        "| Keyword | CV freq |",
        "|---|---:|",
    ]
    for kw, cv_n in sorted(cv_only, key=lambda x: -x[1])[:15]:
        lines.append(f"| {kw} | {cv_n} |")

    # Headline metrics
    match_rate = len(matched) / max(len(matched) + len(missing), 1) * 100

    # High-signal match rate: ignore freq=1 noise (kw that appeared only once
    # in JD are often incidental nouns, not real requirements).
    matched_hi = [(kw, jn, cn) for kw, jn, cn in matched if jn >= 2]
    missing_hi = [(kw, jn) for kw, jn in missing if jn >= 2]
    total_hi = len(matched_hi) + len(missing_hi)
    match_rate_hi = (len(matched_hi) / total_hi * 100) if total_hi else 100

    lines += [
        "",
        "---",
        "",
        f"## 📊 Summary",
        f"- **Match rate (all):** {match_rate:.0f}%  "
        f"({len(matched)} matched / {len(matched) + len(missing)} JD keywords analyzed)",
        f"- **Match rate (high-signal, freq≥2):** {match_rate_hi:.0f}%  "
        f"({len(matched_hi)} matched / {total_hi} keywords) — less noise",
        f"- **High-priority gaps (freq ≥ 3 in JD, missing in CV):** "
        f"{sum(1 for _, n in missing if n >= 3)}",
        f"- **Excluded from analysis** (target company/role names, location noise): "
        f"{len(excluded)} keyword(s){' — ' + ', '.join(k for k, _ in excluded[:8]) if excluded else ''}",
        "",
        "💡 Aim for ≥70% match rate on high-freq JD terms before submitting.",
    ]

    return "\n".join(lines)
