---
name: cv-tailored-master-recipe
description: Build, maintain and tailor a CV/resume project using a Master+Recipe architecture. Use when the user wants to create job applications, tailor a CV to a specific job description, audit keyword match, validate chronology/structure of a resume, or enrich a personal career master record. Covers the full lifecycle: master data → recipe selection → render → validate → compress → submit → track.
version: 1.1
author: code-puppy + Rolph Mercado (CV-RMB sessions)
---

# CV Tailored — Master + Recipe Architecture

> **v1.1 · 2026-05-17** — Battle-scars from the PPG v6 iteration:
> - Compression playbook step **3.5 Orphan-header absorb pattern** (the
>   single highest-leverage fix when 1 optional section gets stranded on
>   the last page).
> - Step 1 (Diagnose) now explicitly checks for orphan-header signature.
> - Anti-pattern **A17** (orphan optional headers) and **A18** (repeated
>   metrics across two roles).
> - New sub-section *Deprecating logros without deleting* under Master
>   enrichment — use `Incluir CV = No` instead of deleting rows.
>
> v1.0 · 2026-05-17 — Initial seal from the PPG/Bricko/Triplelift session.

This skill captures the **CV-RMB** pattern: a personal career-data master from
which **tailored job applications** are generated per job posting, with hard
quality gates so nothing embarrassing ships.

It is opinionated. It exists because anything looser produces hallucinated
metrics, chronology gaps the candidate can't explain in interview, or
ATS-broken layouts. Use it whenever the user wants serious job-search tooling.

## The Cardinal Rule

**Never invent metrics, roles, scope, dates or impact.** If the data is not
in the master and the user hasn't given it to you in conversation, ask. If
the user picks "qualitative only", honor it — descriptive bullets beat false
numbers every single time. (This is anti-pattern `A14` in the project.)

## Architecture in 30 seconds

```
data/master/cv_master_<person>.xlsx     ← Single source of truth, humans edit
            │
            │  scripts/export_master_json.py
            ▼
data/master/cv_master.json              ← Machine-readable snapshot

cv_renderer/recipes/R0N_<archetype>.yaml ← Logro selection per role-type
            +
output/applications/<DATE>_<RECIPE>_<company>_<role>/
            ├── jd.txt                  ← Raw job description
            ├── cv.md                   ← Tailored markdown (the artifact)
            ├── cv.pdf / .docx / .txt   ← Generated outputs
            ├── cv_keywords.md          ← Keyword match audit
            ├── application.yaml        ← Status + timeline + match rate
            └── notes.md                ← Free-form
```

Every artifact under `output/applications/<id>/` is derived from `cv.md` +
master + recipe. Treat `cv.md` as the source.

## Project bootstrap (only when starting fresh)

For a brand-new person:

```
01_projects/cv_<initials>/
├── .venv/                        # uv venv, NEVER share with .code-puppy-venv
├── data/master/
│   ├── cv_master_<person>.xlsx   # Manually built or imported
│   ├── cv_master.json            # Generated
│   └── backups/                  # Auto-versioned vNN_pre_<change>.xlsx
├── cv_renderer/
│   ├── templates/cv.css          # Print styles, A4
│   ├── templates/cv.html.j2      # Jinja layout
│   ├── recipes/R0N_*.yaml        # Archetype recipes
│   └── keywords_audit.py
├── scripts/
│   ├── export_master_json.py
│   ├── render_cv.py
│   ├── validate_cv.py
│   ├── update_status.py
│   ├── rebuild_dashboard.py
│   └── apply_phaseN_*.py         # Migrations (see "Migration pattern")
├── tests/                        # unittest, every helper covered
├── docs/SCRIPTS_REFERENCE.md     # Catalog of every script
└── output/
    ├── applications/<id>/
    ├── _registry.json            # All apps indexed
    └── _dashboard.md             # Human-readable status board
```

Always: `git init`, with `.gitignore` excluding `.venv`, `node_modules`,
PII raw imports, `*.zip` of source CVs. Commit `cv_master.json` (text) but
keep the `.xlsx` binary committed too — humans need to edit it.

## The five recipes (extend as needed)

| Recipe | Archetype                | Selection criteria                          |
|--------|--------------------------|---------------------------------------------|
| R01    | Research / Insights      | Methodology, qualitative, behavioral logros |
| R02    | Strategy / Consulting    | Cross-functional, framework, advisory logros|
| R03    | Data / BI / Analytics    | Quant, BI tools, executive reporting logros |
| R04    | Tech / Product / AdTech  | AI, automation, no-code, product logros     |
| R05    | Creative / Content       | Storytelling, design, media logros          |

Each recipe is a YAML file that selects logros by tag and seniority. A
**recipe is not a CV**, it's a curator. The renderer still composes a
coherent narrative.

## The lifecycle (memorize this loop)

```
1. NEW APP        scripts/new_application.py  --recipe R0X --company X --role Y
2. JD IN          paste posting → output/applications/<id>/jd.txt
3. DRAFT CV       edit output/applications/<id>/cv.md (start from recipe template)
4. RENDER         scripts/render_cv.py <cv.md> <folder> --jd <jd.txt>
5. AUDIT          read cv_keywords.md → target ≥80% all, ≥75% high-signal (freq≥2)
6. VALIDATE       scripts/validate_cv.py <cv.md>  → must be 0 ❌, ideally 0 🟡
7. COMPRESS       if pages > target: see "Compression playbook" below
8. STATUS         scripts/update_status.py <id> ready --note "..."
9. COMMIT         git commit -m "fix(<id>): <change>"
10. SUBMIT        external, then update_status.py <id> submitted --note "via X"
11. DASHBOARD     scripts/rebuild_dashboard.py
```

Steps 4-7 iterate. Never skip 6.

## Validator: what it must catch

Implement `validate_cv.py` to check at least:

| Check               | Pass criterion                                  |
|---------------------|-------------------------------------------------|
| `section:summary`   | `## Summary` (or Profile) present               |
| `section:experience`| `## Experience` present                         |
| `section:education` | `## Education` present                          |
| `section:skills`    | `## Skills` present                             |
| `chronology`        | No gap > 6 months between adjacent roles        |
| `bullet-length`     | Every `-` bullet ≤ 350 chars                    |
| `pdf-pages`         | PDF ≤ target pages (default 3)                  |

A `🟡` is acceptable on `pdf-pages` only briefly during iteration. `❌` on
`chronology` is **always blocking** — either include a filler entry (even a
1-bullet placeholder for a real role the recipe filtered out), or the
candidate cannot defend it.

## Keyword audit: high-signal trumps surface match

Always report two numbers:

- **Overall match rate** (denominator = all unique JD tokens after stopwords)
- **High-signal match rate** (denominator = JD tokens with freq ≥ 2)

The high-signal number is the truth. A JD that says "Magento" 4 times wants
Magento. A JD that says "synergy" once does not.

Sync the audit to `application.yaml.documents.match_rate`. Track it in the
dashboard so trends across applications become visible.

## Compression playbook (when PDF > target pages)

When the validator says "🟡 pages over target", try in this order:

1. **Diagnose first**. Count chars per page with `pypdf`. If the last page
   has < 200 chars, you're 1-2 lines away — try CSS or single-bullet merges.
   If it's > 800 chars, you have structural bloat — kill a section instead.
   **Also inspect the last ~400 chars of the prior page**: if it ends with
   an `## H2` header and no bullets, you have an orphan-header (jump to
   step 3.5).

2. **Merge filler-entry bullets to 1**. Roles that exist purely for
   chronology coverage (e.g. a freelance bridge year) should be 1 bullet
   that fuses founder/role/impact. Save the 2-bullet treatment for
   destination-aligned roles.

3. **Collapse Recognition / Awards into 1 continuous bullet** with `·`
   separators. Headers + multi-bullets are luxury items.

3.5. **Orphan-header absorb pattern** (the high-leverage move). Diagnose with
   `pypdf` whether the last page contains *only* one optional section
   (Recognition, Awards, Languages-as-section, Volunteer work, etc.) AND
   whether that section's `## H2` header is visible at the bottom of the
   previous page while its bullet(s) are stranded alone on the next page.
   That is the **orphan-header** signature.

   When you see it, the cleanest fix is **structural, not editorial**: drop
   the optional section header entirely and absorb its content as the final
   bullet of the *previous* validator-required section (almost always
   Education). Use a bold inline label so the content remains scannable:

   ```markdown
   ## Education
   - **Bachelor's degree** ...
   - **Continuous specialization** ...
   - **Engineering coursework** ...
   - **Speaking & Recognition:** Invited Lecturer ... · Speaker ... · Host ...
   ```

   This eliminates an `<h2>` element (saves ~25-40px of vertical space from
   header + before/after margins) and removes the orphan condition entirely.
   It is *significantly* more effective than compressing the bullet itself
   when the bullet is already short, because the savings come from the
   header chrome, not the text. Only the four validator-required sections
   (Summary / Experience / Education / Skills) must remain as separate H2s;
   everything else is fair game for absorption.

   This pattern beats CSS margin nudges and bullet rewrites whenever the
   stranded section has 1–2 bullets totaling ≤ ~300 chars. Try it before
   step 5.

4. **Compress secondary descriptors** in Education ("Continuous
   specialization ...") before touching career bullets.

5. **CSS margin nudge**: vertical page margin 1.3cm → 1.2cm gains ~1 line.
   Going below 1.1cm starts looking cramped on print.

6. **Last resort**: drop the bottom-most career-entry bullet, never the
   metric-bearing ones.

Re-render and re-validate after every change. Aim for 0 🟡, not "close
enough".

## Master enrichment: do this when the user gives you new biographical info

When the user says "ah by the way that was actually <new fact>", the move is:

1. Confirm what changed (call `ask_user_question` if 2+ alternative
   structures exist — formal company vs paraguas? team size? title?).
2. Write `scripts/apply_phaseN_<change>.py` — a migration script that:
   - Backs up master xlsx to `data/master/backups/vNN_pre_<change>.xlsx`.
   - Updates rows by field name (not column index — column order changes).
   - **Is idempotent** — re-running detects "already applied" and skips.
   - Logs which fields it could not match (column header mismatches).
3. Run it, then `export_master_json.py` to refresh the JSON snapshot.
4. Run `unittest discover -s tests` to make sure nothing broke.
5. Commit with a multi-line message documenting before → after.
6. If the new info also belongs in an *active CV*, edit the `cv.md` of that
   application and re-run the lifecycle from step 4.

Never edit the xlsx directly from the agent except through a migration
script — it makes diffs unreviewable.

### Deprecating logros without deleting

When a logro becomes redundant (e.g. duplicates a metric used in a stronger
role, or its narrative has been rewritten elsewhere), **do not delete the
row**. The historical text is itself data — future recipes or alternative
CVs may want to revive it.

Instead, set `Incluir CV ES` and `Incluir CV EN` to `"No"`. The renderer
and recipe-matcher must respect this flag and never select a deprecated
logro for the live CV. The logro stays queryable for offline analysis,
recipe debugging, or future un-deprecation.

This is the master-side counterpart of the editorial decision "don't repeat
the same metric in two places" (anti-pattern A18). Catch the duplicate when
you notice it (often during the user's review of a draft CV), deprecate the
weaker placement in master, rewrite the active role with a fresh angle
(different metric, different scope, different narrative axis), and commit
both changes together with cross-references in the message.

## Status & dashboard

`application.yaml` carries the canonical status. Use `update_status.py`
(don't hand-edit) so the timeline gets appended properly:

```
draft → ready → submitted → callback → interview → offer / rejected
```

After any status change run `rebuild_dashboard.py`. The dashboard powers
the human's "where do I stand" view; if it falls behind, the user loses
trust in the project.

## Anti-patterns (sample — extend in project's docs/anti_patterns.md)

| ID  | Anti-pattern                                                      |
|-----|-------------------------------------------------------------------|
| A01 | Editing `cv_master.json` directly (will be overwritten on export) |
| A03 | Sharing `.venv` between projects                                  |
| A07 | Skipping the validator because "it's just a small change"         |
| A12 | Recipe-only filtering with no chronology safety net               |
| A14 | **Inventing metrics to fill bullets** ← cardinal sin              |
| A16 | Margin < 1.1cm to fit content (looks desperate on print)          |
| A17 | Optional section header (Recognition/Awards) stranded with its    |
|     | content on the last page — absorb it into Education instead      |
| A18 | **Reusing the same metric across two roles** in the same CV       |
|     | (e.g. "30% NPS" at Canela *and* Yoorco). Pick the strongest       |
|     | placement, rewrite the other with a different angle, or deprecate |
|     | the weaker logro (mark Include CV = No) so it never auto-renders. |

If you find yourself fighting one, stop and surface it to the user.

## Workflow conventions

- **Small commits**: every theme gets its own commit with a multi-line
  message. `feat:`, `fix:`, `chore:`, `feat(master):`.
- **Backups before mutations**: every migration script copies the xlsx
  first. Cheap insurance.
- **Tests before refactor**: if the helper has no test, write one before
  changing it. Target file size ≤ 600 lines.
- **One question at a time** for ambiguous edits — use
  `ask_user_question`. For factual decisions (dates, titles, team size,
  metrics) **always ask, never infer**.
- **Pilot apps get deleted**, not archived. The user said "it was just a
  pilot" → `rmdir` + commit + rebuild dashboard. Clean repo = trustable
  repo.
- **Phase migrations stay in repo** even after they've run — they are the
  audit trail of how the master evolved.

## Tone

The CV project is intimate work — it's the user's career narrative. Be
playful in process but precise in content. Celebrate "0 warnings" runs
(they are rare). When the user reveals a new biographical fact mid-session,
treat it as a gift, not an interruption — re-plan, update master, update
active CV, commit.

## When NOT to use this skill

- One-off resume tweaks where there is no master to maintain → just edit
  and ship.
- Cover letters in isolation (different lifecycle, different skill).
- LinkedIn profile rewrites (no recipe filtering, no PDF gate).
- Portfolio sites (use a static-site skill).
