# 🛠️ Scripts Reference · CV-RMB

Comprehensive catalog of every `scripts/*.py` in the project. Each entry lists
**purpose**, **when to run**, **inputs** and **outputs** so a fresh LLM session
(or human) can navigate the codebase without spelunking.

> 🐶 Convention: `apply_phase*.py` = one-shot migrations that mutate the master.
> `_inspect_*.py` (if any appear) = throwaway debug helpers, safe to delete.

---

## 🎯 Daily-use scripts (you'll touch these often)

| Script | Purpose | Run when |
|---|---|---|
| `render_cv.py` | Render `cv.md` → pdf/docx/txt/json/keywords audit | After every cv.md edit |
| `new_application.py` | Bootstrap a new application folder from JD | Starting a new role hunt |
| `update_status.py` | Transition app pipeline + append timeline notes | Status changes / progress notes |
| `validate_cv.py` | Lint cv.md (gaps, sections, bullets, pages) | Pre-submit sanity check |
| `rebuild_dashboard.py` | Refresh `output/_dashboard.md` + `_registry.json` | After status changes |
| `list_applications.py` | Tabular view of pipeline | Anytime |
| `search_applications.py` | Find apps by company/role/keyword | Anytime |

### `render_cv.py`
```bash
python scripts/render_cv.py <cv.md> [output_dir] [--jd jd.txt]
```
Reads `cv.md`, emits 6 artifacts (md/pdf/docx/txt/json/cv_keywords.md), syncs
`application.yaml` with `match_rate` & output filenames, warns if PDF > 3 pages.

### `new_application.py`
```bash
python scripts/new_application.py --company "PPG Comex" --role "Director" --recipe R07
```
Creates `output/applications/<date>_<recipe>_<company-slug>_<role-slug>/` with
`application.yaml`, empty `cv.md`, `jd.txt`, `notes.md`.

### `update_status.py`
```bash
python scripts/update_status.py <app-folder> <status> [--note "..."]
```
Statuses: `draft → ready → submitted → callback → interview → offer / rejected`.
**Append-note semantics:** if already in the same status, supply `--note` to
add a timeline entry without state change.

### `validate_cv.py`
```bash
python scripts/validate_cv.py <cv.md> [--max-pages 3] [--max-gap-months 6]
```
Checks chronological gaps, required sections, bullet length, PDF page count.
Exit code 1 on FAIL — useful for CI hooks.

---

## 🏗️ Master pipeline scripts

These transform `data/master/cv_master_raul_mercado.xlsx` (source of truth).

| Script | Purpose |
|---|---|
| `export_master_json.py` | Regenerate `cv_master.json` from xlsx |
| `validate_master.py` | Schema + content lint of the master xlsx |
| `audit_master_deep.py` | Deep analysis: gaps, dupes, coverage |
| `dump_xlsx.py` | Pretty-print xlsx contents to console |

### Standard cycle after editing the xlsx
```bash
python scripts/export_master_json.py   # xlsx → json
python scripts/validate_master.py      # sanity check
```

---

## 📜 Apply-phase migrations (one-shot)

Numbered in chronological order. Each is **idempotent** (re-running skips
existing rows). Each takes a backup to `data/master/backups/v{N}_pre_*.xlsx`
before mutating.

| Phase | Script | What it did |
|---|---|---|
| 3 | `apply_phase3.py` | Curated logros candidates → master (Altazor starter set + Phase 3 batch) |
| 3.5 | `curate_phase3_5.py` | Manual curation pass for Phase 3 |
| 4 | `apply_phase4.py` | Complete missing fields (descriptions, narratives) |
| 5 | `apply_phase5_historical.py` | Backfill 2010-2022 findings from historical CVs |
| 6 | `apply_phase6_polish.py` | Deep polish (wording, consistency) |
| 7 | `apply_phase7_visual.py` | Visual polish for "impressive" appearance |
| 8 | `apply_phase8_llm_ready.py` | Make master LLM-friendly (recipes, anti-patterns) |
| 9 | `apply_phase9_qa_fixes.py` | QA fixes from manual audit |
| 10 | `apply_phase10_courses.py` | Add courses (Cinema, Lorenz, Hellinger, etc.) |
| 11 | `apply_phase11_yoor_fix.py` | Yoor end-date 10/2022 → 02/2023, rename to Yoorco |
| 12 | `apply_phase12_learnings.py` | PPG learnings: L75 Yoorco team, R07 corporate director, A14-A16 |

### How to add a new phase

1. Create `scripts/apply_phase{N}_<short_name>.py` following the template of
   the last one (load → backup → mutate → save → log).
2. Use the helper pattern: `_header_map`, `_existing_ids`, `_first_empty_row`,
   `_append_row` (see phase 12 for clean DRY example).
3. Run it. Then `python scripts/export_master_json.py`.

---

## 🔍 Audit & extraction (run-once exploration)

| Script | Purpose |
|---|---|
| `audit_files.py` | Catalog every file in `data/raw/` (xlsx/pdf/docx text extraction) |
| `audit_historical.py` | Extract candidate new logros/dates from 2010-2022 CVs |
| `consolidate_phase2.py` | Phase 2 consolidation of audit findings |
| `extract_logros_phase3.py` | LLM-style extraction of Phase 3 logros candidates |
| `normalize_historical.py` | Normalize historical CV text for diffing |
| `preflight_check.py` | Sanity check before running phases |
| `qa_smoke_test.py` | Smoke test of rendering pipeline |

These produced the reports in `data/reports/` (file catalog, dupes, audit
summary, etc.). They're safe to re-run — outputs are written to `data/reports/`.

---

## 🧪 Tests

```bash
python -m unittest discover -s tests -v
```

- `tests/test_inline_markdown.py` — 18 tests covering bold/italic/links/escaping
  across HTML, plain-text and DOCX run emission.

> Future: add `tests/test_validate_cv.py`, `tests/test_keywords_audit.py`.

---

## 🌳 Folder layout

```
cv_rmb/
├─ applications_manager/    Python pkg: app lifecycle, status, templates
├─ cv_renderer/             Python pkg: parse_markdown + 4 renderers + audit
│   ├─ inline_markdown.py     Tokenizer for bold/italic/links (DRY across renderers)
│   ├─ keywords_audit.py      JD ↔ CV keyword matching + match-rate
│   └─ templates/             HTML/CSS templates for PDF
├─ data/
│   ├─ master/                xlsx (source) + json (export) + backups/
│   ├─ raw/                   Historical CVs (input only, never mutated)
│   └─ reports/               One-shot analysis outputs
├─ docs/                    Editorial flow + this scripts reference
├─ output/
│   ├─ _dashboard.md          Auto-generated pipeline view
│   ├─ _registry.json         Application index
│   └─ applications/<id>/     One folder per JD application
├─ scripts/                 CLI entrypoints (this file documents them)
└─ tests/                   unittest suite
```

---

## 📅 Changelog

- **2026-05-17 · v1** — Initial catalog covering 26 scripts after the PPG
  application run. Added validate_cv.py and apply_phase11/12.
