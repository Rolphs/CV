"""
Inline markdown converter for CV bullets, list items and skill items.

Handles the subset of inline markdown we use in CV bodies:
    **bold**       → <strong>bold</strong>  (HTML) | bold runs (DOCX) | stripped (TXT)
    *italic*       → <em>italic</em>        (HTML) | italic runs (DOCX) | stripped (TXT)
    [text](url)    → <a href="url">text</a> (HTML) | "text (url)" (DOCX/TXT)

DRY by design: ONE tokenizer feeds three sinks (HTML, DOCX runs, plain text).
That way Jinja, python-docx and the .txt writer never disagree on what
"bold" means.
"""
from __future__ import annotations
import html
import re
from typing import Iterator

# A single token from the inline parser.
# kind in {"text", "bold", "italic", "link"}
# text  = displayed string
# href  = only set for kind=="link"
Token = tuple[str, str, str]


# Order matters: bold (**) must be tried before italic (*) to avoid mis-parsing.
# Links are independent of bold/italic.
_LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
_BOLD_RE = re.compile(r"\*\*([^*]+)\*\*")
_ITAL_RE = re.compile(r"(?<!\*)\*([^*]+)\*(?!\*)")


def tokenize(text: str) -> list[Token]:
    """Tokenize inline markdown into a flat sequence of typed runs.

    Greedy single-pass: links → bold → italic. Anything else is plain text.
    Nested formatting (bold inside link, etc.) is intentionally NOT supported;
    keeps CV authoring simple and avoids surprise renders.
    """
    if not text:
        return []
    tokens: list[Token] = [("text", text, "")]
    tokens = _split_by(tokens, _LINK_RE, "link", capture_href=True)
    tokens = _split_by(tokens, _BOLD_RE, "bold")
    tokens = _split_by(tokens, _ITAL_RE, "italic")
    return tokens


def to_html(text: str) -> str:
    """Convert inline markdown to safe HTML. Escapes everything else."""
    parts: list[str] = []
    for kind, content, href in tokenize(text):
        safe_content = html.escape(content)
        if kind == "bold":
            parts.append(f"<strong>{safe_content}</strong>")
        elif kind == "italic":
            parts.append(f"<em>{safe_content}</em>")
        elif kind == "link":
            safe_href = html.escape(href, quote=True)
            parts.append(f'<a href="{safe_href}">{safe_content}</a>')
        else:
            parts.append(safe_content)
    return "".join(parts)


def to_plain(text: str) -> str:
    """Convert inline markdown to plain text (strip formatting, keep content)."""
    parts: list[str] = []
    for kind, content, href in tokenize(text):
        if kind == "link" and href and href != content:
            parts.append(f"{content} ({href})")
        else:
            parts.append(content)
    return "".join(parts)


def iter_docx_runs(text: str) -> Iterator[tuple[str, bool, bool]]:
    """Yield (text, bold, italic) tuples ready to feed python-docx add_run().

    Links are flattened to "text (url)" because plain DOCX runs do not carry
    hyperlinks without extra OOXML plumbing. Keeps the writer simple.
    """
    for kind, content, href in tokenize(text):
        display = content
        if kind == "link" and href and href != content:
            display = f"{content} ({href})"
        yield display, kind == "bold", kind == "italic"


# ───────────────────────── internals ─────────────────────────


def _split_by(
    tokens: list[Token],
    pattern: re.Pattern,
    new_kind: str,
    capture_href: bool = False,
) -> list[Token]:
    """Walk the token list; split any "text" token by the regex, marking matches."""
    out: list[Token] = []
    for kind, content, href in tokens:
        if kind != "text":
            out.append((kind, content, href))
            continue
        idx = 0
        for m in pattern.finditer(content):
            if m.start() > idx:
                out.append(("text", content[idx:m.start()], ""))
            if capture_href:
                out.append((new_kind, m.group(1), m.group(2)))
            else:
                out.append((new_kind, m.group(1), ""))
            idx = m.end()
        if idx < len(content):
            out.append(("text", content[idx:], ""))
    return out
