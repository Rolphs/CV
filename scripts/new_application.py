"""
Create a new CV application skeleton.

Usage:
    python scripts/new_application.py --company "TripleLift" --role "Head of Research" --recipe R04 [--locale en] [--jd path/to/jd.txt]

This creates:
    output/applications/2026-05-17_R04_triplelift_head-of-research/
    ├── application.yaml   (skeleton with metadata)
    ├── notes.md           (private notes template)
    ├── jd.txt             (job description placeholder, or copied from --jd)
    └── (cv.md will be generated separately by the LLM step)
"""
from __future__ import annotations
import sys
import argparse
import shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from applications_manager import create_application  # noqa: E402


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="Create a new CV application skeleton")
    p.add_argument("--company", required=True, help='Company name, e.g. "TripleLift"')
    p.add_argument("--role", required=True, help='Role title, e.g. "Head of Research"')
    p.add_argument("--recipe", required=True, help="Recipe ID from master: R01-R06")
    p.add_argument("--locale", default="en", choices=["en", "es"])
    p.add_argument("--date", default=None, help="Override date (YYYY-MM-DD), default today")
    p.add_argument("--jd", type=Path, default=None,
                   help="Path to JD text file (will be copied to jd.txt)")
    p.add_argument("--master-version", default="phase11")
    args = p.parse_args()

    try:
        app_dir = create_application(
            company_name=args.company,
            role_title=args.role,
            recipe=args.recipe,
            locale=args.locale,
            date_str=args.date,
            master_version=args.master_version,
        )
    except FileExistsError as e:
        print(f"❌ {e}", file=sys.stderr)
        sys.exit(1)

    if args.jd:
        if not args.jd.exists():
            print(f"⚠ JD not found: {args.jd}", file=sys.stderr)
        else:
            shutil.copy2(args.jd, app_dir / "jd.txt")
            print(f"📋 JD copied from {args.jd}")

    print(f"✅ Created: {app_dir.relative_to(app_dir.parent.parent.parent)}")
    print()
    print("Next steps:")
    print(f"  1. Edit {app_dir.name}/jd.txt to paste the full job description")
    print(f"  2. Have an LLM generate {app_dir.name}/cv.md using cv_master.json + recipe + jd.txt")
    print(f"  3. Render outputs:")
    print(f"     python scripts/render_cv.py output/applications/{app_dir.name}/cv.md \\")
    print(f"            output/applications/{app_dir.name} \\")
    print(f"            --jd output/applications/{app_dir.name}/jd.txt")
    print(f"  4. Mark as submitted when sent:")
    print(f"     python scripts/update_status.py {app_dir.name} submitted")


if __name__ == "__main__":
    main()
