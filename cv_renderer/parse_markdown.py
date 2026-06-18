"""
Parse CV markdown → structured dict.

Expected schema (see sample_cv.md for reference):

    ---
    recipe: R04
    target_company: TripleLift
    target_role: Head of Research
    locale: en
    date: 2026-05-17
    job_posting_url: https://...
    ---

    # Full Name
    ## Headline

    Location · email · phone
    LinkedIn-URL · willing to relocate

    ## Summary
    paragraph

    ## Experience

    ### Company — Role
    **dates · location**

    - bullet 1
    - bullet 2

    ## Skills
    **Category:** item, item, item

    ## Education
    - **Degree**, Institution · year

    ## Recognition & Speaking
    - item

Output dict structure:

    {
        "meta": {recipe, target_company, ..., locale, date},
        "header": {
            "name": "Raúl Mercado Bustamante",
            "headline": "Head of Research & Insights",
            "contact_lines": ["Mexico City · email · phone", "linkedin · note"],
        },
        "sections": [
            {"title": "Summary", "type": "paragraph", "content": "..."},
            {"title": "Experience", "type": "experience", "entries": [
                {"company": "...", "role": "...", "dates": "...", "location": "...",
                 "bullets": ["...", "..."]},
            ]},
            {"title": "Skills", "type": "skills", "categories": [
                {"name": "Research", "items": ["UX Research", "Mixed Methods"]},
            ]},
            {"title": "Education", "type": "list", "items": ["..."]},
            ...
        ]
    }
"""
from __future__ import annotations
import re
from typing import Any


def parse_cv_markdown(text: str) -> dict[str, Any]:
    """Parse a CV markdown string into structured dict."""
    text = text.replace("\r\n", "\n")
    meta, body = _split_frontmatter(text)
    header, sections_text = _split_header(body)
    sections = _parse_sections(sections_text)
    return {"meta": meta, "header": header, "sections": sections}


# ───────────────────────── Frontmatter ─────────────────────────
_FRONTMATTER_RE = re.compile(r"^---\n(.*?)\n---\n", re.DOTALL)


def _split_frontmatter(text: str) -> tuple[dict, str]:
    m = _FRONTMATTER_RE.match(text)
    if not m:
        return {}, text
    meta_block = m.group(1)
    body = text[m.end():]
    meta: dict[str, Any] = {}
    for line in meta_block.split("\n"):
        line = line.strip()
        if not line or ":" not in line:
            continue
        k, _, v = line.partition(":")
        meta[k.strip()] = v.strip()
    return meta, body


# ───────────────────────── Header ─────────────────────────────
def _split_header(body: str) -> tuple[dict, str]:
    """Extract name (# H1), headline (## H2), contact lines until first ## section."""
    lines = body.split("\n")
    name = ""
    headline = ""
    contact_lines: list[str] = []
    i = 0
    # Skip leading blanks
    while i < len(lines) and not lines[i].strip():
        i += 1
    # # Name
    if i < len(lines) and lines[i].startswith("# ") and not lines[i].startswith("## "):
        name = lines[i][2:].strip()
        i += 1
    # ## Headline
    while i < len(lines) and not lines[i].strip():
        i += 1
    if i < len(lines) and lines[i].startswith("## ") and not lines[i].startswith("### "):
        headline = lines[i][3:].strip()
        i += 1
    # Contact lines: non-empty, non-section lines until next "## "
    while i < len(lines):
        line = lines[i].rstrip()
        if line.startswith("## ") and not line.startswith("### "):
            break
        if line.strip():
            contact_lines.append(line.strip())
        i += 1
    rest = "\n".join(lines[i:])
    return {"name": name, "headline": headline, "contact_lines": contact_lines}, rest


# ───────────────────────── Sections ───────────────────────────
def _parse_sections(text: str) -> list[dict]:
    """Split by '## SectionTitle' and route to appropriate sub-parser."""
    sections: list[dict] = []
    # Match "## Title" headers, capture title and content until next ## or EOF
    chunks = re.split(r"^## (?!#)", text, flags=re.MULTILINE)
    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk:
            continue
        first_nl = chunk.find("\n")
        if first_nl == -1:
            title = chunk.strip()
            content = ""
        else:
            title = chunk[:first_nl].strip()
            content = chunk[first_nl + 1:].strip()
        section = _route_section(title, content)
        if section:
            sections.append(section)
    return sections


def _route_section(title: str, content: str) -> dict | None:
    tl = title.lower().strip()
    if not content and not title:
        return None
    if tl in ("experience", "experiencia", "professional experience"):
        return {"title": title, "type": "experience",
                "entries": _parse_experience_entries(content)}
    if tl in ("skills", "habilidades", "core skills"):
        return {"title": title, "type": "skills",
                "categories": _parse_skills_categories(content)}
    if tl in ("summary", "resumen", "profile", "perfil"):
        return {"title": title, "type": "paragraph", "content": content.strip()}
    # Default: list of bullets (Education, Recognition, etc.)
    items = _parse_bullets(content)
    if items:
        return {"title": title, "type": "list", "items": items}
    return {"title": title, "type": "paragraph", "content": content.strip()}


# ───────────────────────── Experience parser ──────────────────
def _parse_experience_entries(content: str) -> list[dict]:
    """Parse '### Company — Role\\n**dates · location**\\n- bullet\\n...' entries."""
    entries: list[dict] = []
    chunks = re.split(r"^### ", content, flags=re.MULTILINE)
    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk:
            continue
        lines = chunk.split("\n")
        title_line = lines[0].strip()
        # "Company — Role" or "Company - Role"
        if "—" in title_line:
            company, _, role = title_line.partition("—")
        elif " - " in title_line:
            company, _, role = title_line.partition(" - ")
        else:
            company, role = title_line, ""
        entry: dict[str, Any] = {
            "company": company.strip(),
            "role": role.strip(),
            "dates": "",
            "location": "",
            "bullets": [],
        }
        body_lines = lines[1:]
        i = 0
        # First non-empty bold line = dates · location
        while i < len(body_lines) and not body_lines[i].strip():
            i += 1
        if i < len(body_lines):
            meta_line = body_lines[i].strip()
            if meta_line.startswith("**") and meta_line.endswith("**"):
                meta_clean = meta_line.strip("*").strip()
                if "·" in meta_clean:
                    dates, _, loc = meta_clean.partition("·")
                    entry["dates"] = dates.strip()
                    entry["location"] = loc.strip()
                else:
                    entry["dates"] = meta_clean
                i += 1
        # Remaining lines: bullets or description
        rest = "\n".join(body_lines[i:]).strip()
        entry["bullets"] = _parse_bullets(rest)
        entries.append(entry)
    return entries


# ───────────────────────── Skills parser ──────────────────────
def _parse_skills_categories(content: str) -> list[dict]:
    """Parse '**Category:** item, item' lines."""
    categories: list[dict] = []
    for line in content.split("\n"):
        line = line.strip()
        if not line:
            continue
        m = re.match(r"\*\*([^*]+):\*\*\s*(.+)", line)
        if m:
            cat = m.group(1).strip()
            items = [x.strip() for x in m.group(2).split(",") if x.strip()]
            categories.append({"name": cat, "items": items})
        elif line.startswith("- "):
            # Fallback: treat plain bullet as misc skill
            categories.append({"name": "", "items": [line[2:].strip()]})
    return categories


# ───────────────────────── Bullet parser ──────────────────────
def _parse_bullets(content: str) -> list[str]:
    """Extract bullets (- ...) preserving inline content."""
    bullets: list[str] = []
    current: list[str] = []
    for line in content.split("\n"):
        if line.startswith("- "):
            if current:
                bullets.append(" ".join(current).strip())
                current = []
            current.append(line[2:].rstrip())
        elif line.startswith("  ") and current:
            # Continuation
            current.append(line.strip())
        elif line.strip() == "":
            if current:
                bullets.append(" ".join(current).strip())
                current = []
        elif current:
            current.append(line.strip())
    if current:
        bullets.append(" ".join(current).strip())
    return [b for b in bullets if b]
