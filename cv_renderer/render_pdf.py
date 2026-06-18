"""
Render CV dict → PDF (xhtml2pdf, pure Python, no system deps).

Design principles:
- Single column (ATS-friendly)
- Calibri-like font (Helvetica fallback)
- Black/white conservative, subtle gray accent lines
- No headers/footers, no page numbers
- Embedded XMP metadata for LLM parsers
"""
from __future__ import annotations
from pathlib import Path
from io import BytesIO
from xhtml2pdf import pisa
from jinja2 import Environment, FileSystemLoader, select_autoescape
from markupsafe import Markup

from .inline_markdown import to_html as inline_to_html

TEMPLATES_DIR = Path(__file__).resolve().parent / "templates"


def render_pdf(cv: dict, output_path) -> Path:
    """Render CV dict → PDF at output_path. Returns the path."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    html = _render_html(cv)
    css = (TEMPLATES_DIR / "cv.css").read_text(encoding="utf-8")
    full_html = html.replace("</head>", f"<style>{css}</style></head>")

    with open(output_path, "wb") as f:
        result = pisa.CreatePDF(
            src=full_html,
            dest=f,
            encoding="utf-8",
        )

    if result.err:
        raise RuntimeError(f"xhtml2pdf reported {result.err} errors")

    return output_path


def _render_html(cv: dict) -> str:
    env = Environment(
        loader=FileSystemLoader(str(TEMPLATES_DIR)),
        autoescape=select_autoescape(["html"]),
    )
    # Custom filter: converts **bold**/*italic*/[link](url) to safe HTML.
    # Used as `{{ value | inline_md }}` in cv.html.j2 (no extra |safe needed:
    # Markup() already marks the output as safe).
    env.filters["inline_md"] = lambda s: Markup(inline_to_html(s or ""))
    tmpl = env.get_template("cv.html.j2")
    return tmpl.render(
        meta=cv.get("meta", {}),
        header=cv.get("header", {}),
        sections=cv.get("sections", []),
    )
