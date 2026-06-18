"""
Render CV dict → plain text (UTF-8, ASCII-safe fallback).

For the rare ATS that only accepts .txt uploads. Keeps semantic structure
through whitespace and UPPERCASE section headers.
"""
from __future__ import annotations
from pathlib import Path

from .inline_markdown import to_plain


def render_txt(cv: dict, output_path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    lines: list[str] = []
    header = cv.get("header", {})

    # ── HEADER ──
    name = header.get("name", "")
    if name:
        lines.append(name.upper())
        lines.append("=" * len(name))
    if header.get("headline"):
        lines.append(header["headline"])
    lines.append("")
    for c in header.get("contact_lines", []):
        lines.append(c)
    lines.append("")

    # ── SECTIONS ──
    for sec in cv.get("sections", []):
        title = sec.get("title", "")
        lines.append("")
        lines.append(title.upper())
        lines.append("-" * len(title))

        stype = sec.get("type")
        if stype == "paragraph":
            lines.append(to_plain(sec.get("content", "")))

        elif stype == "experience":
            for entry in sec.get("entries", []):
                lines.append("")
                head = entry.get("company", "")
                if entry.get("role"):
                    head = f"{head} — {entry['role']}"
                lines.append(head)
                meta = entry.get("dates", "")
                if entry.get("location"):
                    meta = f"{meta} · {entry['location']}"
                if meta:
                    lines.append(meta)
                for b in entry.get("bullets", []):
                    lines.append(f"  • {to_plain(b)}")

        elif stype == "skills":
            for cat in sec.get("categories", []):
                cat_name = cat.get("name", "")
                items = ", ".join(to_plain(i) for i in cat.get("items", []))
                if cat_name:
                    lines.append(f"{cat_name}: {items}")
                else:
                    lines.append(items)

        elif stype == "list":
            for item in sec.get("items", []):
                lines.append(f"  • {to_plain(item)}")

    text = "\n".join(lines).rstrip() + "\n"
    output_path.write_text(text, encoding="utf-8")
    return output_path
