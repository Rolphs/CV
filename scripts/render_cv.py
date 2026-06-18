"""
CLI: render a CV markdown file to all output formats.

Usage:
    python scripts/render_cv.py path/to/cv.md [output_dir] [--jd path/to/jd.txt]

If output_dir is omitted, defaults to:
    output/<recipe>_<target_company>_<date>/
"""
from __future__ import annotations
import sys
import argparse
from pathlib import Path

# Make project root importable
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from cv_renderer import render_all, parse_cv_markdown  # noqa: E402


def main():
    parser = argparse.ArgumentParser(description="Render a CV markdown to multi-format output")
    parser.add_argument("md_path", type=Path, help="Path to source cv.md")
    parser.add_argument("output_dir", type=Path, nargs="?", default=None,
                        help="Output directory (default: output/<recipe>_<co>_<date>/)")
    parser.add_argument("--jd", type=Path, default=None,
                        help="Optional: path to job description text file for keywords audit")
    args = parser.parse_args()

    if not args.md_path.exists():
        print(f"❌ MD file not found: {args.md_path}", file=sys.stderr)
        sys.exit(1)

    # Derive default output dir from meta
    if args.output_dir is None:
        cv = parse_cv_markdown(args.md_path.read_text(encoding="utf-8"))
        m = cv.get("meta", {})
        recipe = m.get("recipe", "Rxx")
        company = (m.get("target_company") or "company").replace(" ", "_")
        date = m.get("date") or "undated"
        args.output_dir = (
            Path(__file__).resolve().parent.parent
            / "output" / f"{recipe}_{company}_{date}"
        )

    jd_text = None
    if args.jd:
        if not args.jd.exists():
            print(f"❌ JD file not found: {args.jd}", file=sys.stderr)
            sys.exit(1)
        jd_text = args.jd.read_text(encoding="utf-8")

    print(f"📄 Source: {args.md_path}")
    print(f"📁 Output: {args.output_dir}")
    print(f"🔍 JD audit: {'yes' if jd_text else 'no'}")
    print()

    outputs = render_all(args.md_path, args.output_dir, job_posting_text=jd_text)

    print("✅ Generated:")
    for fmt, path in outputs.items():
        size = path.stat().st_size
        size_str = f"{size // 1024} KB" if size > 1024 else f"{size} B"
        print(f"  {fmt:<10} → {path.relative_to(args.output_dir.parent)}  ({size_str})")

    # If the output_dir is an application folder, update application.yaml
    _sync_application_yaml(args.output_dir, outputs)

    # Page-count warning (3 = target sweet spot, 4 = warning, 5+ = scream)
    _warn_if_pdf_too_long(outputs.get("pdf"), target_pages=3)


def _warn_if_pdf_too_long(pdf_path, target_pages: int = 3):
    if not pdf_path or not pdf_path.exists():
        return
    try:
        from pypdf import PdfReader
        pages = len(PdfReader(str(pdf_path)).pages)
    except Exception:
        return
    if pages <= target_pages:
        return
    over = pages - target_pages
    icon = "🟡" if over == 1 else "❌"
    print(
        f"\n{icon} PDF rendered with {pages} pages "
        f"(target ≤{target_pages}). Consider compressing the weakest block, "
        "merging adjacent bullets, or trimming CSS margins."
    )


def _sync_application_yaml(out_dir: Path, outputs: dict):
    """If out_dir contains application.yaml, populate documents.outputs and match_rate."""
    yaml_path = out_dir / "application.yaml"
    if not yaml_path.exists():
        return

    from applications_manager import load_application, save_application
    import re

    data = load_application(out_dir)
    docs = data.setdefault("documents", {})
    out_section = docs.setdefault("outputs", {})
    for fmt in ["pdf", "docx", "txt", "json"]:
        if fmt in outputs:
            out_section[fmt] = outputs[fmt].name
    if "keywords" in outputs:
        docs.setdefault("audits", {})["keywords"] = outputs["keywords"].name
        # Extract match_rate from the audit file
        try:
            txt = outputs["keywords"].read_text(encoding="utf-8")
            m = re.search(r"Match rate:\*\*\s*(\d+)%", txt)
            if m:
                docs["match_rate"] = int(m.group(1))
        except Exception:
            pass

    save_application(out_dir, data)
    print(f"\n🔗 Synced application.yaml (match_rate: {docs.get('match_rate')})")


if __name__ == "__main__":
    main()
