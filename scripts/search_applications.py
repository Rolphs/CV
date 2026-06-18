"""
Full-text search across all application materials (CVs, JDs, notes).

Usage:
    python scripts/search_applications.py <query> [--in cv|jd|notes|all]

Examples:
    python scripts/search_applications.py "Hispanic"
    python scripts/search_applications.py "salary" --in notes
    python scripts/search_applications.py "CTV" --in jd
"""
from __future__ import annotations
import sys
import argparse
import re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from applications_manager import (  # noqa: E402
    APPLICATIONS_DIR, STATUS_EMOJI, load_application,
)

FILE_GLOBS = {
    "cv": ["cv.md", "cv.txt"],
    "jd": ["jd.txt"],
    "notes": ["notes.md"],
    "all": ["cv.md", "cv.txt", "jd.txt", "notes.md"],
}


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="Search across application files")
    p.add_argument("query", help="Search term (case-insensitive)")
    p.add_argument("--in", dest="scope", default="all", choices=list(FILE_GLOBS.keys()))
    p.add_argument("--regex", action="store_true",
                   help="Treat query as regex (default is plain substring)")
    args = p.parse_args()

    pattern = re.compile(args.query if args.regex else re.escape(args.query),
                         re.IGNORECASE)
    files_to_check = FILE_GLOBS[args.scope]

    if not APPLICATIONS_DIR.exists():
        print("📭 No applications/ directory yet.")
        return

    total_hits = 0
    apps_with_hits = 0

    for app_dir in sorted(APPLICATIONS_DIR.iterdir()):
        if not app_dir.is_dir():
            continue
        if not (app_dir / "application.yaml").exists():
            continue
        try:
            data = load_application(app_dir)
        except Exception:
            continue

        hits_in_app: list[tuple[str, int, str]] = []
        for fname in files_to_check:
            fpath = app_dir / fname
            if not fpath.exists():
                continue
            try:
                content = fpath.read_text(encoding="utf-8")
            except Exception:
                continue
            for lineno, line in enumerate(content.split("\n"), 1):
                if pattern.search(line):
                    hits_in_app.append((fname, lineno, line.strip()))

        if hits_in_app:
            apps_with_hits += 1
            total_hits += len(hits_in_app)
            emoji = STATUS_EMOJI.get(data.get("status", ""), "")
            print(f"📁 {emoji} {data.get('id', app_dir.name)}")
            print(f"   {data.get('company', {}).get('name', '')} · "
                  f"{data.get('role', {}).get('title', '')}")
            for fname, lineno, snippet in hits_in_app[:5]:  # cap at 5 per app
                snippet_show = snippet[:120] + ("…" if len(snippet) > 120 else "")
                print(f"   └─ {fname}:{lineno}  {snippet_show}")
            if len(hits_in_app) > 5:
                print(f"   └─ … and {len(hits_in_app) - 5} more match(es)")
            print()

    print(f"🔍 {total_hits} match(es) across {apps_with_hits} application(s)")


if __name__ == "__main__":
    main()
