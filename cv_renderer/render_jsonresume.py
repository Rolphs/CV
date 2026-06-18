"""
Render CV dict → JSON Resume schema (jsonresume.org).

This is THE machine-readable standard for resumes. Specialized LLM tools
(TalentBrew, Eightfold, hireEZ) can ingest it directly without parsing.
"""
from __future__ import annotations
import json
import re
from pathlib import Path
from typing import Any

from .inline_markdown import to_plain


def render_jsonresume(cv: dict, output_path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    header = cv.get("header", {})
    sections = cv.get("sections", [])

    basics = _build_basics(header)
    work, education, awards, skills, summary = [], [], [], [], ""

    for sec in sections:
        stype = sec.get("type")
        title_l = sec.get("title", "").lower()

        if "summary" in title_l or "profile" in title_l:
            summary = to_plain(sec.get("content", ""))
        elif stype == "experience":
            work = [_build_work_entry(e) for e in sec.get("entries", [])]
        elif "education" in title_l:
            education = [_build_education_entry(i) for i in sec.get("items", [])]
        elif "recognition" in title_l or "awards" in title_l or "speaking" in title_l:
            awards.extend(_build_award_entry(i) for i in sec.get("items", []))
        elif stype == "skills":
            skills = [
                {"name": cat.get("name", ""),
                 "keywords": [to_plain(k) for k in cat.get("items", [])]}
                for cat in sec.get("categories", [])
            ]

    resume = {
        "$schema": "https://raw.githubusercontent.com/jsonresume/resume-schema/v1.0.0/schema.json",
        "basics": {**basics, "summary": summary},
        "work": work,
        "education": education,
        "skills": skills,
        "awards": awards,
        "meta": {
            **cv.get("meta", {}),
            "canonical": "https://jsonresume.org/schema/",
            "version": "v1.0.0",
        },
    }

    output_path.write_text(
        json.dumps(resume, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    return output_path


# ──────────────────────── Builders ───────────────────────────
_EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
_PHONE_RE = re.compile(r"\+?\d[\d\s().-]{7,}")
_URL_RE = re.compile(r"https?://\S+")


def _build_basics(header: dict) -> dict:
    blob = " ".join(header.get("contact_lines", []))
    email_m = _EMAIL_RE.search(blob)
    phone_m = _PHONE_RE.search(blob)
    url_m = _URL_RE.search(blob)

    # Try to extract location from first contact line (before email)
    location_str = ""
    for line in header.get("contact_lines", []):
        first_part = line.split("·")[0].strip()
        if "@" not in first_part and not first_part.startswith("http"):
            location_str = first_part
            break

    basics: dict[str, Any] = {
        "name": header.get("name", ""),
        "label": header.get("headline", ""),
        "email": email_m.group(0) if email_m else "",
        "phone": phone_m.group(0).strip() if phone_m else "",
        "url": url_m.group(0) if url_m else "",
        "location": {},
        "profiles": [],
    }

    if location_str:
        parts = [p.strip() for p in location_str.split(",")]
        if len(parts) >= 2:
            basics["location"] = {"city": parts[0], "countryCode": "", "region": parts[-1]}
        else:
            basics["location"] = {"city": parts[0]}

    if url_m and "linkedin" in url_m.group(0).lower():
        basics["profiles"].append({
            "network": "LinkedIn",
            "url": url_m.group(0),
            "username": url_m.group(0).rstrip("/").split("/")[-1],
        })

    return basics


def _build_work_entry(entry: dict) -> dict:
    start, end = _parse_dates(entry.get("dates", ""))
    return {
        "name": entry.get("company", ""),
        "position": entry.get("role", ""),
        "location": entry.get("location", ""),
        "startDate": start,
        "endDate": end,
        "highlights": [to_plain(b) for b in entry.get("bullets", [])],
    }


def _parse_dates(dates_str: str) -> tuple[str, str]:
    """Parse 'Apr 2023 – Mar 2025' or 'Apr 2023 – Present' → ('2023-04', '2025-03')."""
    if not dates_str:
        return "", ""
    parts = re.split(r"\s*[–-]\s*", dates_str, maxsplit=1)
    start = _to_iso_month(parts[0]) if parts else ""
    end = ""
    if len(parts) > 1:
        end_raw = parts[1].strip()
        if end_raw.lower() in ("present", "actual", "current", "hoy"):
            end = ""  # JSON Resume convention: omit endDate for current
        else:
            end = _to_iso_month(end_raw)
    return start, end


_MONTHS = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04", "may": "05", "jun": "06",
    "jul": "07", "aug": "08", "sep": "09", "oct": "10", "nov": "11", "dec": "12",
    "ene": "01", "abr": "04", "ago": "08", "dic": "12",
}


def _to_iso_month(s: str) -> str:
    s = s.strip().lower()
    m = re.match(r"([a-z]+)\s+(\d{4})", s)
    if m:
        mon = _MONTHS.get(m.group(1)[:3], "01")
        return f"{m.group(2)}-{mon}"
    m = re.match(r"(\d{4})", s)
    if m:
        return m.group(1)
    return s


def _build_education_entry(item: str) -> dict:
    """Parse '**Degree**, Institution · year[· extras]'."""
    item_clean = item.replace("**", "").strip()
    parts = [p.strip() for p in item_clean.split("·")]
    main = parts[0] if parts else item_clean
    year = ""
    for p in parts[1:]:
        m = re.search(r"(\d{4})", p)
        if m:
            year = m.group(1)
            break
    if "," in main:
        degree, _, institution = main.partition(",")
    else:
        degree, institution = main, ""
    return {
        "institution": institution.strip(),
        "studyType": degree.strip(),
        "endDate": year,
        "summary": " · ".join(parts[1:]) if len(parts) > 1 else "",
    }


def _build_award_entry(item: str) -> dict:
    item_clean = item.replace("**", "").strip()
    parts = [p.strip() for p in item_clean.split("·")]
    return {
        "title": parts[0] if parts else item_clean,
        "awarder": parts[1] if len(parts) > 1 else "",
        "date": next((p for p in parts if re.match(r"^\d{4}", p) or "Present" in p), ""),
        "summary": item_clean,
    }
