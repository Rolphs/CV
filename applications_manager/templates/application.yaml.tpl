# application.yaml — manifest of a single CV application
# Edit this file as the application progresses. Re-run scripts/rebuild_dashboard.py
# to refresh the central registry and dashboard.

# ── IDENTITY (immutable, set at creation) ─────────────────────────
id: "{id}"                                  # 2026-05-17_R04_triplelift_head-of-research
created: "{created}"                        # ISO timestamp
last_updated: "{last_updated}"

# ── TARGETING ─────────────────────────────────────────────────────
recipe: "{recipe}"                          # R01-R06 from master sheet 12

company:
  name: "{company_name}"                    # Display name: "TripleLift"
  slug: "{company_slug}"                    # URL-safe: "triplelift"
  industry: ""                              # AdTech | CPG | Media | Streaming | Consulting | Other
  size: ""                                  # startup | small | mid | large | enterprise
  hq_location: ""                           # New York, NY · USA
  notes: ""                                 # Optional free text about the company

role:
  title: "{role_title}"                     # Full title from JD
  slug: "{role_slug}"                       # URL-safe
  seniority: ""                             # entry | mid | senior | director | head | vp | c-suite
  locale: "{locale}"                        # en | es
  remote: ""                                # onsite | hybrid | remote | flexible
  location: ""                              # New York, NY (hybrid)

job_posting:
  url: ""                                   # Original posting URL
  text_snapshot: "jd.txt"                   # Relative path to JD snapshot in this folder
  source: ""                                # linkedin | greenhouse | indeed | recruiter | direct | referral
  expires: ""                               # YYYY-MM-DD or empty
  salary_range: ""                          # Optional: "USD 200K-250K"

# ── PIPELINE STATUS ───────────────────────────────────────────────
# Valid states: draft → ready → submitted → callback → interview → offer
#                                        ↘ rejected ↘ withdrawn ↘ ghosted
status: "draft"

# Append-only history. Each transition adds a new entry.
timeline:
  - date: "{created_date}"
    status: "draft"
    note: "Application created"

# ── DOCUMENTS ─────────────────────────────────────────────────────
documents:
  source: "cv.md"                           # Markdown source-of-truth
  outputs:                                  # Filled by render_cv.py
    pdf: ""
    docx: ""
    txt: ""
    json: ""
  audits:
    keywords: ""                            # cv_keywords.md
  match_rate: null                          # int 0-100 from keywords audit
  cover_letter: ""                          # Optional: cover_letter.md

# ── SUBMISSION ────────────────────────────────────────────────────
submission:
  portal: ""                                # linkedin | greenhouse | workday | email | recruiter | other
  format_sent: ""                           # pdf | docx | both
  recruiter_contact: ""                     # "Name <email>" · optional
  application_id_external: ""               # ID from portal if any
  referral: ""                              # Person who referred · optional

# ── OUTCOME TRACKING ──────────────────────────────────────────────
outcome:
  callback_date: ""                         # YYYY-MM-DD when first response
  callback_days: null                       # auto-computed (days since submission)
  interviews: []                            # list of dicts: date, type, with_who, notes
  offer:
    received_date: ""
    salary: ""
    accepted: null                          # true | false | null (pending)
  rejected_reason: ""                       # auto-categorized or user text

# ── REPRODUCIBILITY ───────────────────────────────────────────────
master_version: "{master_version}"          # Sheet phase used as source
notes_file: "notes.md"                      # Private notes (gitignored if marked)
