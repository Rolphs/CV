"""
Validate a rendered CV for common quality issues.

Checks:
  1. Chronological gaps in experience timeline (>6 months WARN, >12 months FAIL)
  2. Required sections present (Summary, Experience, Education, Skills)
  3. PDF page count (if cv.pdf exists alongside) vs target (default ≤3)
  4. ATS hygiene (bullets longer than 350 chars are usually formatting bugs)

Usage:
    python scripts/validate_cv.py <cv.md path> [--max-pages 3] [--max-gap-months 6]

Exit codes:
    0 = all checks passed (warnings OK)
    1 = at least one FAIL
    2 = file not found / invalid markdown
"""
from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from cv_renderer.parse_markdown import parse_cv_markdown  # noqa: E402
from cv_renderer.render_jsonresume import _parse_dates  # noqa: E402


# ─────────────────────────── Data ───────────────────────────


@dataclass
class Finding:
    level: str  # "PASS" | "WARN" | "FAIL"
    check: str
    message: str

    @property
    def icon(self) -> str:
        return {"PASS": "✅", "WARN": "🟡", "FAIL": "❌"}[self.level]


@dataclass
class Report:
    findings: list[Finding] = field(default_factory=list)

    def add(self, level: str, check: str, message: str):
        self.findings.append(Finding(level, check, message))

    @property
    def has_fails(self) -> bool:
        return any(f.level == "FAIL" for f in self.findings)

    @property
    def stats(self) -> dict[str, int]:
        return {
            "PASS": sum(1 for f in self.findings if f.level == "PASS"),
            "WARN": sum(1 for f in self.findings if f.level == "WARN"),
            "FAIL": sum(1 for f in self.findings if f.level == "FAIL"),
        }


# ─────────────────────────── Helpers ───────────────────────────


def _iso_to_date(iso: str) -> Optional[date]:
    """Parse 'YYYY-MM' or 'YYYY' to a date (first of month)."""
    if not iso:
        return None
    parts = iso.split("-")
    try:
        year = int(parts[0])
        month = int(parts[1]) if len(parts) > 1 else 1
        return date(year, month, 1)
    except (ValueError, IndexError):
        return None


def _months_between(earlier: date, later: date) -> int:
    return (later.year - earlier.year) * 12 + (later.month - earlier.month)


# ─────────────────────────── Checks ───────────────────────────


def check_sections(cv: dict, report: Report):
    """Verify mandatory sections are present and non-empty."""
    sections_by_type = {s.get("type") or s.get("title", "").lower(): s
                        for s in cv.get("sections", [])}
    titles_lower = {s.get("title", "").lower() for s in cv.get("sections", [])}

    required = [
        ("summary", "Summary / Profile"),
        ("experience", "Experience"),
        ("education", "Education"),
        ("skills", "Skills"),
    ]
    for key, label in required:
        found = (
            key in sections_by_type
            or any(key in t for t in titles_lower)
            or any(label.lower().split(" / ")[0] in t for t in titles_lower)
        )
        if found:
            report.add("PASS", f"section:{key}", f"{label} section present")
        else:
            report.add("FAIL", f"section:{key}", f"{label} section MISSING")


def check_chronology(cv: dict, report: Report, max_gap_months: int):
    """Detect chronological gaps between adjacent jobs."""
    exp_section = next(
        (s for s in cv.get("sections", []) if s.get("type") == "experience"),
        None,
    )
    if not exp_section:
        report.add("WARN", "chronology", "No experience section to validate")
        return

    entries = exp_section.get("entries", [])
    if len(entries) < 2:
        report.add("PASS", "chronology", f"Only {len(entries)} entry — nothing to compare")
        return

    # Parse dates → (start_date, end_date) tuples; keep job label for reporting
    parsed: list[tuple[date, Optional[date], str]] = []
    today = date.today()
    for e in entries:
        s_iso, end_iso = _parse_dates(e.get("dates", ""))
        s = _iso_to_date(s_iso)
        end = _iso_to_date(end_iso) if end_iso else today
        if s is None:
            report.add(
                "WARN", "chronology",
                f"Could not parse start date for '{e.get('role','?')} @ {e.get('company','?')}': {e.get('dates','')!r}",
            )
            continue
        label = f"{e.get('role','?')} @ {e.get('company','?')}"
        parsed.append((s, end, label))

    if len(parsed) < 2:
        return

    # Sort by start desc (most recent first, matching how CV reads)
    parsed.sort(key=lambda t: t[0], reverse=True)

    gaps_found = 0
    for i in range(len(parsed) - 1):
        newer_start = parsed[i][0]
        older_end = parsed[i + 1][1] or today
        gap = _months_between(older_end, newer_start)
        newer_label = parsed[i][2]
        older_label = parsed[i + 1][2]

        if gap <= 1:
            # adjacent or overlapping → OK
            continue
        elif gap <= max_gap_months:
            report.add(
                "PASS", "chronology",
                f"{gap}mo gap (≤{max_gap_months}) between "
                f"'{older_label}' end and '{newer_label}' start — acceptable",
            )
        elif gap <= max_gap_months * 2:
            gaps_found += 1
            report.add(
                "WARN", "chronology",
                f"{gap}mo gap between '{older_label}' (ended {parsed[i+1][1]}) "
                f"and '{newer_label}' (started {parsed[i][0]}) — consider explaining",
            )
        else:
            gaps_found += 1
            report.add(
                "FAIL", "chronology",
                f"{gap}mo gap between '{older_label}' (ended {parsed[i+1][1]}) "
                f"and '{newer_label}' (started {parsed[i][0]}) — must cover with filler entry",
            )

    if gaps_found == 0:
        report.add("PASS", "chronology",
                   f"No gaps > {max_gap_months}mo across {len(parsed)} positions")


def check_bullet_length(cv: dict, report: Report, max_chars: int = 350):
    """Flag bullets that look like formatting bugs (very long)."""
    overlong = []
    for s in cv.get("sections", []):
        if s.get("type") != "experience":
            continue
        for e in s.get("entries", []):
            for b in e.get("bullets", []):
                if len(b) > max_chars:
                    overlong.append((e.get("company", "?"), len(b), b[:80] + "..."))

    if overlong:
        for company, length, preview in overlong:
            report.add(
                "WARN", "bullet-length",
                f"{company}: bullet of {length} chars (>{max_chars}) — {preview}",
            )
    else:
        report.add("PASS", "bullet-length", f"All bullets ≤{max_chars} chars")


def check_pdf_pages(cv_md_path: Path, report: Report, max_pages: int):
    """Inspect adjacent cv.pdf if present."""
    pdf_path = cv_md_path.with_suffix(".pdf")
    if not pdf_path.exists():
        report.add("WARN", "pdf-pages", f"No {pdf_path.name} found — skipping page-count check")
        return
    try:
        from pypdf import PdfReader
        pages = len(PdfReader(str(pdf_path)).pages)
    except ImportError:
        report.add("WARN", "pdf-pages", "pypdf not installed — skipping")
        return
    except Exception as exc:
        report.add("WARN", "pdf-pages", f"Could not read PDF: {exc}")
        return

    if pages <= max_pages:
        report.add("PASS", "pdf-pages", f"PDF has {pages} page(s) (target ≤{max_pages})")
    elif pages == max_pages + 1:
        report.add(
            "WARN", "pdf-pages",
            f"PDF has {pages} pages (over target {max_pages}) — consider compressing",
        )
    else:
        report.add(
            "FAIL", "pdf-pages",
            f"PDF has {pages} pages (target ≤{max_pages}) — too long",
        )


# ─────────────────────────── Reporter ───────────────────────────


def print_report(report: Report, cv_md_path: Path):
    print(f"\n📋 Validation report · {cv_md_path}")
    print("─" * 78)
    for f in report.findings:
        print(f"  {f.icon}  [{f.check:18s}] {f.message}")
    s = report.stats
    print("─" * 78)
    print(f"  Summary: ✅ {s['PASS']} · 🟡 {s['WARN']} · ❌ {s['FAIL']}")
    print()


# ─────────────────────────── Main ───────────────────────────


def validate(cv_md_path: Path, max_pages: int = 3, max_gap_months: int = 6) -> Report:
    text = cv_md_path.read_text(encoding="utf-8")
    cv = parse_cv_markdown(text)

    report = Report()
    check_sections(cv, report)
    check_chronology(cv, report, max_gap_months)
    check_bullet_length(cv, report)
    check_pdf_pages(cv_md_path, report, max_pages)
    return report


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="Validate a rendered CV markdown.")
    p.add_argument("cv_path", help="Path to cv.md (will look for cv.pdf in same folder)")
    p.add_argument("--max-pages", type=int, default=3)
    p.add_argument("--max-gap-months", type=int, default=6,
                   help="WARN above this; FAIL above 2×this")
    args = p.parse_args()

    cv_md = Path(args.cv_path)
    if not cv_md.exists():
        print(f"❌ Not found: {cv_md}", file=sys.stderr)
        sys.exit(2)

    report = validate(cv_md, max_pages=args.max_pages, max_gap_months=args.max_gap_months)
    print_report(report, cv_md)
    sys.exit(1 if report.has_fails else 0)


if __name__ == "__main__":
    main()
