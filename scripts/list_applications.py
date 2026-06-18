"""
List CV applications with optional filters.

Usage:
    python scripts/list_applications.py [--status STATUS] [--company NAME] [--recipe Rxx] [--since YYYY-MM-DD]

Examples:
    python scripts/list_applications.py
    python scripts/list_applications.py --status submitted
    python scripts/list_applications.py --status submitted --status callback
    python scripts/list_applications.py --company TripleLift
    python scripts/list_applications.py --since 2026-05-01
"""
from __future__ import annotations
import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from applications_manager import (  # noqa: E402
    list_applications, VALID_STATUSES, STATUS_EMOJI,
)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="List CV applications with filters")
    p.add_argument("--status", action="append", choices=VALID_STATUSES,
                   help="Filter by status (can pass multiple)")
    p.add_argument("--company", help="Filter by company name (substring match on slug)")
    p.add_argument("--recipe", help="Filter by recipe ID, e.g. R04")
    p.add_argument("--since", help="Only show apps created since YYYY-MM-DD")
    p.add_argument("--long", action="store_true", help="Show more columns")
    args = p.parse_args()

    apps = list_applications(
        status=args.status if args.status else None,
        company=args.company,
        recipe=args.recipe,
        since=args.since,
    )

    if not apps:
        print("📭 No applications match the filters.")
        return

    print(f"📋 {len(apps)} application(s)")
    print()

    if args.long:
        _print_long(apps)
    else:
        _print_short(apps)


def _print_short(apps):
    # Date | Status | Recipe | Company | Role
    rows = []
    for a in apps:
        st = a.get("status", "")
        emoji = STATUS_EMOJI.get(st, "  ")
        rows.append((
            (a.get("created") or "")[:10],
            f"{emoji} {st:<10}",
            a.get("recipe", "")[:4],
            (a.get("company", {}).get("name") or "")[:22],
            (a.get("role", {}).get("title") or "")[:38],
        ))

    widths = [max(len(r[i]) for r in rows + [("Date", "Status", "Rec", "Company", "Role")])
              for i in range(5)]
    header = ("Date", "Status", "Rec", "Company", "Role")
    print(" │ ".join(h.ljust(widths[i]) for i, h in enumerate(header)))
    print("─┼─".join("─" * widths[i] for i in range(5)))
    for r in rows:
        print(" │ ".join(str(r[i]).ljust(widths[i]) for i in range(5)))


def _print_long(apps):
    for a in apps:
        st = a.get("status", "")
        emoji = STATUS_EMOJI.get(st, "  ")
        mr = a.get("documents", {}).get("match_rate")
        mr_str = f" · match {mr}%" if mr is not None else ""
        cb = a.get("outcome", {}).get("callback_days")
        cb_str = f" · cb-days {cb}" if cb is not None else ""
        print(f"{emoji} {st:<10} · {a.get('recipe', ''):<4} · {a.get('id', '')}")
        print(f"     Company: {a.get('company', {}).get('name', '')}")
        print(f"     Role:    {a.get('role', {}).get('title', '')}")
        print(f"     Created: {(a.get('created') or '')[:10]}{mr_str}{cb_str}")
        print()


if __name__ == "__main__":
    main()
