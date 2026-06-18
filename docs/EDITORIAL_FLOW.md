# 📝 Editorial Flow · CV Generation Lifecycle

How a single CV application goes from JD → polished PDF, based on the patterns
proven during the **PPG/Comex Data Intelligence Director** application
(2026-05-17).

> 🎯 Target outcome: 3-page PDF · ≥85% keyword match · 0 high-priority gaps · 0
> chronological gaps · 100% master-grounded content.

---

## 🔁 The 8-step loop

```
┌──────────────────────────────────────────────────────────────────────┐
│  1. JD ingestion        →  scripts/new_application.py                │
│  2. Recipe selection    →  evaluate R01-R07; build hybrid if needed  │
│  3. Master read         →  inspect logros, perfil, anti-patterns     │
│  4. cv.md draft         →  mirror JD verbs verbatim                  │
│  5. Render + audit      →  scripts/render_cv.py (md/pdf/docx/txt/json)│
│  6. Iterate match-rate  →  weave missing JD keywords (no fakes!)     │
│  7. Editorial pass      →  user review (geo, tone, seniority, gaps)  │
│  8. Mark ready/submit   →  scripts/update_status.py                  │
└──────────────────────────────────────────────────────────────────────┘
```

---

## 1 · JD ingestion

```bash
python scripts/new_application.py \
    --company "PPG Comex" \
    --role "Data Intelligence Director MX & CAM" \
    --recipe R07 \
    --locale en
```

Creates `output/applications/{date}_{recipe}_{company-slug}_{role-slug}/` with
`application.yaml`, empty `cv.md`, `jd.txt`, `notes.md`.

**Then paste the full JD into `jd.txt`** (don't summarize — the audit needs the
raw text to compute match rate).

## 2 · Recipe selection

Run through `data/master/cv_master.json` → `recipes` and pick the closest fit.
**No recipe is perfect for every role.** When none calza 100%:

- ✅ Build a **hybrid** (e.g. `R03+R05` or `R03+R07`).
- ✅ Document the hybrid in `application.yaml` under `recipe`.
- ✅ Log seniority overrides, industry pivots and inherited bullets.
- ❌ Don't silently force a recipe that misrepresents the candidate.

**Aprendizaje PPG:** R01-R06 no cubrían "Director corporate FMCG-adjacent MX".
Construimos R03+R05 → ahora existe R07 para casos similares.

## 3 · Master read

For an LLM session, dump only what's needed:

```python
# Inline 1-shot inspector (throwaway, prefix with _)
from cv_master import perfil, logros, recipes, anti_patterns
recipe = recipes["R07"]
relevant = [l for l in logros if l.id in recipe.logros]
```

**Always re-read `anti_patterns` before drafting.** A14-A16 (added 2026-05-17)
are the freshest rules.

## 4 · cv.md draft

Conventions that worked for PPG:

- **Headline:** mirror target role title (no creative liberties for ATS).
- **Summary:** open with the candidate's spine, weave 3-5 high-frequency JD
  keywords *literally* (PPG: `pricing, promotion and profitability strategies`).
- **Geographic scope:** state it explicitly when the JD names regions
  (`Mexico and Central America` for PPG MX & CAM).
- **Industry adjacency:** when you lack the target industry (e.g. coatings),
  pivot to **adjacent categories** (`consumer-facing categories`,
  `transversal data intelligence`) — never invent paint/aerospace/etc.
- **Bullets:** start with the JD's exact verb when the action matches a real
  logro (`oversee`, `guide`, `foster`, `identify`, `apply`).
- **Most recent + most relevant first:** Walmart Scintilla (CPG retail
  analytics) opens the section even though it's a "small" 4-month role —
  it's the strongest signal for the target.

## 5 · Render + audit

```bash
python scripts/render_cv.py output/applications/<app>/cv.md \
       output/applications/<app> \
       --jd output/applications/<app>/jd.txt
```

Produces six artifacts:

| Artifact | Purpose |
|---|---|
| `cv.md` | Source of truth (round-trippable) |
| `cv.pdf` | Visual delivery (xhtml2pdf, 3-page sweet spot) |
| `cv.docx` | Workday/SuccessFactors-friendly, native styling |
| `cv.txt` | Last-resort ATS upload |
| `cv.json` | JSON Resume v1.0.0 (LLM-parseable) |
| `cv_keywords.md` | Match rate + gaps audit |

## 6 · Iterate match-rate

Open `cv_keywords.md`, look for **High-priority gaps (freq ≥ 3 in JD)**.

| Match rate | Action |
|---|---|
| 90%+ | ✅ Ready to ship |
| 70-89% | 🟡 1-2 iteration cycles |
| <70% | 🔴 Recipe likely wrong, reassess |

**PPG run:** 60% → 92% in one iteration by weaving in `oversee`, `guide`,
`foster`, `apply`, `identify`, `ensure` into existing logros. No new content
invented.

> 🛡️ **Audit auto-excludes** the target company name & role tokens (`ppg`,
> `data`, `intelligence`, `director`, etc.) from match-rate calculation so they
> don't show up as fake "missing keywords".

## 7 · Editorial pass

This is where the candidate (Rolph) reviews and the LLM applies surgical edits.
A good editorial pass touches:

1. **Summary:** literal JD vocabulary, geographic scope, narrative bridge to
   target industry.
2. **Title sanity:** Walmart Scintilla → `Strategic Analytics Lead` (NOT
   "Account Manager"), aligned with seniority claim.
3. **Methodology depth:** name your differentiator (neurophysiological: eye-
   tracking, facial coding, electrodermal response).
4. **JD vocabulary lift:** if the JD says `pricing strategies`, put pricing in
   the FIRST bullet of the most relevant role (Kantar for PPG).
5. **Historical credentials:** name frameworks (Needscope, Ekman, Lüscher,
   Jungian) where they prove origin/depth.
6. **Chronological gap fillers:** **never leave > 6 months unexplained.** Cover
   real-world consulting / sabbatical / job-search with a tight 1-line entry.
7. **Skills phrasing:** drop apologetic qualifiers (`(autodidact)` →
   `(insight pipelines, automation, exploratory analysis)`).
8. **Education compensation:** when no Master's exists, add a continuous-
   specialization line directly under the Bachelor's to signal lifelong
   learning.
9. **Anchor longevity:** the 13-year podcast becomes `(13 years on air)` —
   discipline signal.
10. **Selective compression:** less relevant blocks (TIBA/KIO for a coatings
    role) get compressed first when space is tight.

## 8 · Mark ready / submit

```bash
# Initial ready
python scripts/update_status.py <app> ready \
       --note "v1 · 92% match · 3 pages"

# Append notes without changing status (fixed 2026-05-17)
python scripts/update_status.py <app> ready \
       --note "added Yoorco fix per Rolph 17-may"

# When sent
python scripts/update_status.py <app> submitted \
       --note "via LinkedIn Easy Apply · recruiter: TBD"

# Always refresh the dashboard
python scripts/rebuild_dashboard.py
```

---

## 🚫 Anti-patterns rule of thumb

Before shipping, check `cv.md` against A01-A16 in `cv_master.json`. The three
added 2026-05-17 are the most often violated by enthusiastic LLM drafts:

- **A14 · No inventar industria target** — if the JD says "architectural
  coatings" and your master doesn't, don't claim it.
- **A15 · Recipe forzado** — if no recipe calza, build a documented hybrid.
- **A16 · Mirror JD verbatim** — reformulate real logros with JD verbs, never
  fabricate.

---

## 🔍 Common pitfalls (seen on PPG run)

| Pitfall | Symptom | Fix |
|---|---|---|
| Wrong dates in master | Inconsistent with reality | Update xlsx → re-export json |
| `**bold**` in bullets shows as literal asterisks | Education/Recognition look broken in PDF | `inline_markdown.py` handles it (2026-05-17) |
| Match rate inflated by target name | `ppg` listed as missing keyword | Exclusion set in `keywords_audit.py` |
| 4-page PDF when budget is 3 | One bullet spills | Compress weakest block (TIBA/KIO, Recognition merge) or trim CSS margin |
| `update_status.py` blocks re-noting | Can't append progress note | Fixed 2026-05-17 — supply `--note` and same status |

---

## 📅 Changelog

- **2026-05-17 · v1** — Initial editorial flow from PPG/Comex run.
  - Added R07 recipe (Director Data Intelligence Corporate).
  - Added A14, A15, A16 anti-patterns.
  - Added L75 (Yoorco team & founder mindset).
  - Fixed Yoor end-date in master (10/2022 → 02/2023).
  - Fixed `update_status.py` to allow append-note semantics.
  - Fixed inline markdown rendering across PDF/DOCX/TXT/JSON.
  - Fixed keyword audit to exclude target company/role tokens.
