"""
CV Renderer · Markdown → PDF/DOCX/TXT/JSON Resume + Keywords audit.

Public API:
    from cv_renderer import render_all
    render_all(md_path, output_dir, job_posting_text=None)
"""
from .parse_markdown import parse_cv_markdown

# Lazy-loaded renderers to avoid importing optional deps at module load
__all__ = [
    "parse_cv_markdown",
    "render_all",
]


def render_all(md_path, output_dir, job_posting_text=None):
    """Render all 5 outputs from a single CV markdown file.

    Args:
        md_path: Path to cv.md
        output_dir: Directory where outputs will be written
        job_posting_text: Optional job posting text for keywords audit

    Returns:
        dict mapping format → output path
    """
    from pathlib import Path
    import shutil
    import json

    md_path = Path(md_path)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Lazy imports so each renderer is independent
    from .render_pdf import render_pdf
    from .render_docx import render_docx
    from .render_txt import render_txt
    from .render_jsonresume import render_jsonresume
    from .keywords_audit import audit_keywords

    # 1. Copy MD source
    md_dest = out_dir / "cv.md"
    if md_path.resolve() != md_dest.resolve():
        shutil.copy2(md_path, md_dest)

    # 2. Parse
    cv = parse_cv_markdown(md_path.read_text(encoding="utf-8"))

    # 3. Render all formats
    outputs = {"md": md_dest}
    outputs["pdf"] = render_pdf(cv, out_dir / "cv.pdf")
    outputs["docx"] = render_docx(cv, out_dir / "cv.docx")
    outputs["txt"] = render_txt(cv, out_dir / "cv.txt")
    outputs["json"] = render_jsonresume(cv, out_dir / "cv.json")

    # 4. Keywords audit (if JD provided)
    if job_posting_text:
        outputs["keywords"] = audit_keywords(
            cv, job_posting_text, out_dir / "cv_keywords.md"
        )

    # 5. Meta file
    meta = {
        "recipe": cv.get("meta", {}).get("recipe"),
        "target_company": cv.get("meta", {}).get("target_company"),
        "target_role": cv.get("meta", {}).get("target_role"),
        "locale": cv.get("meta", {}).get("locale"),
        "date": cv.get("meta", {}).get("date"),
        "outputs": {k: str(v.relative_to(out_dir)) for k, v in outputs.items()},
    }
    (out_dir / "_meta.json").write_text(
        json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    return outputs
