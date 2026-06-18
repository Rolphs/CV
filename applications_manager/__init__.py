"""
Application Manager · Track CV applications through their lifecycle.

Public API:
    from applications_manager import (
        VALID_STATUSES, ACTIVE_STATUSES, CLOSED_STATUSES,
        load_application, save_application, list_applications,
        slugify, build_app_id, create_application, transition_status,
    )
"""
from __future__ import annotations
import re
import unicodedata
from datetime import datetime, date
from pathlib import Path
from typing import Any, Iterable
import yaml


# ── STATUS VOCABULARY ─────────────────────────────────────────
DRAFT = "draft"
READY = "ready"
SUBMITTED = "submitted"
CALLBACK = "callback"
INTERVIEW = "interview"
OFFER = "offer"
REJECTED = "rejected"
WITHDRAWN = "withdrawn"
GHOSTED = "ghosted"

VALID_STATUSES = [DRAFT, READY, SUBMITTED, CALLBACK, INTERVIEW, OFFER,
                  REJECTED, WITHDRAWN, GHOSTED]

ACTIVE_STATUSES = {DRAFT, READY, SUBMITTED, CALLBACK, INTERVIEW}
CLOSED_STATUSES = {OFFER, REJECTED, WITHDRAWN, GHOSTED}

STATUS_EMOJI = {
    DRAFT: "📝",
    READY: "✅",
    SUBMITTED: "📤",
    CALLBACK: "📞",
    INTERVIEW: "🎤",
    OFFER: "🎯",
    REJECTED: "❌",
    WITHDRAWN: "🚪",
    GHOSTED: "👻",
}

# ── PATHS ─────────────────────────────────────────────────────
PROJECT_ROOT = Path(__file__).resolve().parent.parent
APPLICATIONS_DIR = PROJECT_ROOT / "output" / "applications"
REGISTRY_PATH = PROJECT_ROOT / "output" / "_registry.json"
DASHBOARD_PATH = PROJECT_ROOT / "output" / "_dashboard.md"
TEMPLATES_DIR = Path(__file__).resolve().parent / "templates"


# ── SLUGIFY ───────────────────────────────────────────────────
def slugify(text: str) -> str:
    """Convert to URL-safe kebab-case: 'Head of Research!' → 'head-of-research'."""
    if not text:
        return ""
    # Normalize unicode (é → e)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    text = text.lower().strip()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


def build_app_id(date_str: str, recipe: str, company: str, role: str) -> str:
    """Build canonical application ID: 2026-05-17_R04_triplelift_head-of-research."""
    co_slug = slugify(company)[:30] or "company"
    role_slug = slugify(role)[:40] or "role"
    return f"{date_str}_{recipe}_{co_slug}_{role_slug}"


# ── LOAD / SAVE ───────────────────────────────────────────────
def load_application(app_dir: Path) -> dict:
    """Load application.yaml from an application directory."""
    app_dir = Path(app_dir)
    yaml_path = app_dir / "application.yaml"
    if not yaml_path.exists():
        raise FileNotFoundError(f"No application.yaml in {app_dir}")
    return yaml.safe_load(yaml_path.read_text(encoding="utf-8"))


def save_application(app_dir: Path, data: dict):
    """Save application.yaml, bumping last_updated automatically."""
    app_dir = Path(app_dir)
    data["last_updated"] = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    yaml_path = app_dir / "application.yaml"
    yaml_path.write_text(
        yaml.safe_dump(data, sort_keys=False, allow_unicode=True, default_flow_style=False),
        encoding="utf-8",
    )


# ── LIST / FILTER ─────────────────────────────────────────────
def list_applications(
    base_dir: Path | None = None,
    status: str | Iterable[str] | None = None,
    company: str | None = None,
    recipe: str | None = None,
    since: str | None = None,
) -> list[dict]:
    """List applications with optional filters. Returns list of dicts sorted desc by date."""
    base_dir = Path(base_dir) if base_dir else APPLICATIONS_DIR
    if not base_dir.exists():
        return []

    apps: list[dict] = []
    for app_dir in sorted(base_dir.iterdir(), reverse=True):
        if not app_dir.is_dir():
            continue
        if not (app_dir / "application.yaml").exists():
            continue
        try:
            data = load_application(app_dir)
        except Exception:
            continue
        data["_dir"] = str(app_dir)
        data["_dir_name"] = app_dir.name
        apps.append(data)

    # Apply filters
    if status:
        statuses = {status} if isinstance(status, str) else set(status)
        apps = [a for a in apps if a.get("status") in statuses]
    if company:
        co_slug = slugify(company)
        apps = [a for a in apps
                if co_slug in slugify(a.get("company", {}).get("name", ""))
                or co_slug in a.get("company", {}).get("slug", "")]
    if recipe:
        apps = [a for a in apps if a.get("recipe") == recipe]
    if since:
        apps = [a for a in apps if a.get("created", "")[:10] >= since]

    return apps


# ── CREATE / TRANSITION ───────────────────────────────────────
def create_application(
    company_name: str,
    role_title: str,
    recipe: str,
    locale: str = "en",
    date_str: str | None = None,
    master_version: str = "phase11",
) -> Path:
    """Create a new application skeleton folder. Returns the app dir path.

    If date_str is provided, it sets BOTH the folder id date AND the 'created'
    timestamp (useful for backfilling historical applications).
    """
    today_iso = date.today().isoformat()
    date_str = date_str or today_iso
    app_id = build_app_id(date_str, recipe, company_name, role_title)
    app_dir = APPLICATIONS_DIR / app_id
    if app_dir.exists():
        raise FileExistsError(f"Application already exists: {app_dir}")
    app_dir.mkdir(parents=True)

    # If user supplied a past date, treat 'created' as midday of that date.
    # Otherwise use real now() timestamp.
    if date_str != today_iso:
        created_iso = f"{date_str}T12:00:00"
    else:
        created_iso = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    template = (TEMPLATES_DIR / "application.yaml.tpl").read_text(encoding="utf-8")
    filled = template.format(
        id=app_id,
        created=created_iso,
        last_updated=created_iso,
        created_date=date_str,
        recipe=recipe,
        company_name=company_name,
        company_slug=slugify(company_name),
        role_title=role_title.replace('"', "'"),
        role_slug=slugify(role_title),
        locale=locale,
        master_version=master_version,
    )
    (app_dir / "application.yaml").write_text(filled, encoding="utf-8")

    notes_template = (TEMPLATES_DIR / "notes.md.tpl").read_text(encoding="utf-8")
    notes_filled = notes_template.format(
        id=app_id,
        created_date=date_str,
        company_name=company_name,
        role_title=role_title,
        salary_range="(see application.yaml)",
    )
    (app_dir / "notes.md").write_text(notes_filled, encoding="utf-8")

    # Empty jd.txt placeholder
    (app_dir / "jd.txt").write_text(
        f"# Paste the job description for {company_name} - {role_title} here\n",
        encoding="utf-8",
    )

    return app_dir


def transition_status(app_dir: Path, new_status: str, note: str = "") -> dict:
    """Transition an application to a new status, appending to timeline.

    If the application is already in `new_status`, the call is allowed ONLY
    when a non-empty `note` is supplied (append-note semantics). Without a
    note, it would be a no-op and we raise ValueError to flag user intent.
    """
    if new_status not in VALID_STATUSES:
        raise ValueError(f"Invalid status '{new_status}'. Valid: {VALID_STATUSES}")
    data = load_application(app_dir)
    old_status = data.get("status", "")
    same_status = old_status == new_status
    if same_status and not note:
        raise ValueError(
            f"Already in status '{new_status}'. Pass --note to append a "
            "timeline entry without changing status."
        )

    data["status"] = new_status
    timeline_entry = {
        "date": date.today().isoformat(),
        "status": new_status,
        "note": note or f"Transitioned from {old_status}",
    }
    data.setdefault("timeline", []).append(timeline_entry)

    # Special transitions: auto-fill key dates (only when status actually changes)
    if not same_status:
        if new_status == CALLBACK:
            data.setdefault("outcome", {})["callback_date"] = date.today().isoformat()
            # Compute callback_days from submission
            sub = next((t for t in data["timeline"] if t["status"] == SUBMITTED), None)
            if sub:
                try:
                    d_sub = date.fromisoformat(sub["date"])
                    data["outcome"]["callback_days"] = (date.today() - d_sub).days
                except Exception:
                    pass

    save_application(app_dir, data)
    return data
