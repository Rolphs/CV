"""
Update an application's status (transitions the pipeline + logs timeline).

Usage:
    python scripts/update_status.py <app_id_or_dir> <new_status> [--note "free text"]

Examples:
    python scripts/update_status.py 2026-05-17_R04_triplelift_head-of-research submitted
    python scripts/update_status.py 2026-05-17_R04_triplelift_head-of-research callback --note "Lisa Park scheduled call"
"""
from __future__ import annotations
import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from applications_manager import (  # noqa: E402
    APPLICATIONS_DIR, transition_status, VALID_STATUSES, STATUS_EMOJI,
)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="Transition an application status")
    p.add_argument("app", help="Application ID (folder name) or path")
    p.add_argument("status", choices=VALID_STATUSES,
                   help=f"New status. Valid: {', '.join(VALID_STATUSES)}")
    p.add_argument("--note", default="", help="Optional note for the timeline entry")
    args = p.parse_args()

    app_dir = Path(args.app)
    if not app_dir.is_absolute() and not (app_dir / "application.yaml").exists():
        app_dir = APPLICATIONS_DIR / args.app

    if not (app_dir / "application.yaml").exists():
        print(f"❌ No application.yaml found at {app_dir}", file=sys.stderr)
        sys.exit(1)

    try:
        data = transition_status(app_dir, args.status, note=args.note)
    except ValueError as e:
        print(f"❌ {e}", file=sys.stderr)
        sys.exit(1)

    print(f"✅ {data.get('id')}")
    print(f"   {STATUS_EMOJI[args.status]} → {args.status}")
    if args.note:
        print(f"   Note: {args.note}")
    print()
    print("Don't forget to refresh the dashboard:")
    print("   python scripts/rebuild_dashboard.py")


if __name__ == "__main__":
    main()
