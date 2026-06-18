"""
Rebuild the central registry and dashboard from all applications/.

Generates:
- output/_registry.json   (machine-readable, full data of all apps)
- output/_dashboard.md    (human-readable, metrics + tables)

Usage:
    python scripts/rebuild_dashboard.py
"""
from __future__ import annotations
import json
import sys
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from applications_manager import (  # noqa: E402
    APPLICATIONS_DIR, REGISTRY_PATH, DASHBOARD_PATH,
    list_applications, VALID_STATUSES, ACTIVE_STATUSES, CLOSED_STATUSES,
    STATUS_EMOJI, SUBMITTED, CALLBACK, INTERVIEW, OFFER, REJECTED,
)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    apps = list_applications()
    print(f"📊 Found {len(apps)} application(s)")

    _write_registry(apps)
    _write_dashboard(apps)

    print(f"✅ Registry: {REGISTRY_PATH.relative_to(REGISTRY_PATH.parent.parent)}")
    print(f"✅ Dashboard: {DASHBOARD_PATH.relative_to(DASHBOARD_PATH.parent.parent)}")


def _write_registry(apps: list[dict]):
    """Serialize all apps to a single JSON for machine consumption."""
    REGISTRY_PATH.parent.mkdir(parents=True, exist_ok=True)
    registry = {
        "generated_at": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "count": len(apps),
        "applications": [_simplify_for_registry(a) for a in apps],
    }
    REGISTRY_PATH.write_text(
        json.dumps(registry, indent=2, ensure_ascii=False, default=str),
        encoding="utf-8",
    )


def _simplify_for_registry(a: dict) -> dict:
    """Trim noisy fields for the registry view."""
    return {
        "id": a.get("id"),
        "dir": a.get("_dir_name"),
        "status": a.get("status"),
        "recipe": a.get("recipe"),
        "company": a.get("company", {}).get("name"),
        "company_slug": a.get("company", {}).get("slug"),
        "role": a.get("role", {}).get("title"),
        "seniority": a.get("role", {}).get("seniority"),
        "locale": a.get("role", {}).get("locale"),
        "created": a.get("created"),
        "last_updated": a.get("last_updated"),
        "match_rate": a.get("documents", {}).get("match_rate"),
        "callback_date": a.get("outcome", {}).get("callback_date"),
        "callback_days": a.get("outcome", {}).get("callback_days"),
        "timeline_len": len(a.get("timeline", [])),
    }


def _write_dashboard(apps: list[dict]):
    """Generate a Markdown dashboard with metrics and tables."""
    if not apps:
        DASHBOARD_PATH.write_text(
            "# CV Applications Dashboard\n\n_No applications yet._\n",
            encoding="utf-8",
        )
        return

    total = len(apps)
    active_count = sum(1 for a in apps if a.get("status") in ACTIVE_STATUSES)
    closed_count = sum(1 for a in apps if a.get("status") in CLOSED_STATUSES)
    status_counter = Counter(a.get("status", "") for a in apps)

    submitted = sum(1 for a in apps
                    if any(t.get("status") == SUBMITTED for t in a.get("timeline", [])))
    callbacks = sum(1 for a in apps
                    if any(t.get("status") == CALLBACK for t in a.get("timeline", [])))
    interviews = sum(1 for a in apps
                     if any(t.get("status") == INTERVIEW for t in a.get("timeline", [])))
    offers = sum(1 for a in apps
                 if any(t.get("status") == OFFER for t in a.get("timeline", [])))

    callback_rate = (callbacks / submitted * 100) if submitted else 0
    interview_rate = (interviews / callbacks * 100) if callbacks else 0
    offer_rate = (offers / interviews * 100) if interviews else 0

    # By recipe stats
    by_recipe = defaultdict(lambda: {"total": 0, "submitted": 0, "callback": 0})
    for a in apps:
        r = a.get("recipe", "?")
        by_recipe[r]["total"] += 1
        tl_states = {t.get("status") for t in a.get("timeline", [])}
        if SUBMITTED in tl_states:
            by_recipe[r]["submitted"] += 1
        if CALLBACK in tl_states:
            by_recipe[r]["callback"] += 1

    lines: list[str] = []
    lines.append("# CV Applications Dashboard")
    lines.append("")
    lines.append(f"_Last rebuilt: {datetime.now().strftime('%Y-%m-%d %H:%M')}_")
    lines.append("")
    lines.append(f"**Total:** {total} applications · **Active:** {active_count} · **Closed:** {closed_count}")
    lines.append("")

    # ── PIPELINE ──
    lines.append("## Pipeline status")
    lines.append("")
    lines.append("| Status | Count | % |")
    lines.append("|---|---:|---:|")
    for st in VALID_STATUSES:
        n = status_counter.get(st, 0)
        if n == 0:
            continue
        pct = n / total * 100
        lines.append(f"| {STATUS_EMOJI[st]} {st.capitalize()} | {n} | {pct:.0f}% |")
    lines.append("")

    # ── FUNNEL ──
    lines.append("## Conversion funnel")
    lines.append("")
    lines.append("| Stage | Count | Conversion |")
    lines.append("|---|---:|---:|")
    lines.append(f"| 📤 Submitted | {submitted} | — |")
    lines.append(f"| 📞 Callback | {callbacks} | {callback_rate:.0f}% |")
    lines.append(f"| 🎤 Interview | {interviews} | {interview_rate:.0f}% |")
    lines.append(f"| 🎯 Offer | {offers} | {offer_rate:.0f}% |")
    lines.append("")

    # ── BY RECIPE ──
    if by_recipe:
        lines.append("## By recipe")
        lines.append("")
        lines.append("| Recipe | Apps | Submitted | Callbacks | CB rate |")
        lines.append("|---|---:|---:|---:|---:|")
        for r in sorted(by_recipe):
            d = by_recipe[r]
            cr = (d["callback"] / d["submitted"] * 100) if d["submitted"] else 0
            lines.append(f"| {r} | {d['total']} | {d['submitted']} | {d['callback']} | {cr:.0f}% |")
        lines.append("")

    # ── ACTIVE APPS (action needed) ──
    active_apps = [a for a in apps if a.get("status") in ACTIVE_STATUSES]
    if active_apps:
        lines.append(f"## 🔥 Active applications ({len(active_apps)})")
        lines.append("")
        lines.append("| Date | Status | Recipe | Company | Role | Match | Last update |")
        lines.append("|---|---|---|---|---|---:|---|")
        for a in sorted(active_apps,
                        key=lambda x: x.get("last_updated", ""), reverse=True):
            d = (a.get("created") or "")[:10]
            st = STATUS_EMOJI.get(a.get("status", ""), "") + " " + a.get("status", "")
            mr = a.get("documents", {}).get("match_rate")
            mr_str = f"{mr}%" if mr is not None else "—"
            lu = (a.get("last_updated") or "")[:10]
            lines.append(
                f"| {d} | {st} | {a.get('recipe', '')} | "
                f"{_truncate(a.get('company', {}).get('name', ''), 22)} | "
                f"{_truncate(a.get('role', {}).get('title', ''), 30)} | "
                f"{mr_str} | {lu} |"
            )
        lines.append("")

    # ── ALL APPS TABLE ──
    lines.append(f"## All applications ({total})")
    lines.append("")
    lines.append("| Date | Status | Recipe | Company | Role | Folder |")
    lines.append("|---|---|---|---|---|---|")
    for a in sorted(apps, key=lambda x: x.get("created", ""), reverse=True):
        d = (a.get("created") or "")[:10]
        st = STATUS_EMOJI.get(a.get("status", ""), "") + " " + a.get("status", "")
        co = _truncate(a.get("company", {}).get("name", ""), 22)
        rl = _truncate(a.get("role", {}).get("title", ""), 35)
        folder = a.get("_dir_name", "")
        lines.append(f"| {d} | {st} | {a.get('recipe', '')} | {co} | {rl} | `{folder}` |")
    lines.append("")

    DASHBOARD_PATH.write_text("\n".join(lines), encoding="utf-8")


def _truncate(s: str, n: int) -> str:
    s = s or ""
    return s if len(s) <= n else s[:n - 1] + "…"


if __name__ == "__main__":
    main()
