"""
QA Final: simula ser un LLM consumiendo cv_master.json para generar un CV.

Si el master está bien diseñado, debería ser trivial:
  1. Leer recipes
  2. Pickear una recipe
  3. Resolver las referencias (logros, perfil narrativo)
  4. Output un CV en markdown

Si algo se rompe, esa es la señal de qué falta o está mal.
"""
from __future__ import annotations
import json
import sys
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
JSON_PATH = ROOT / "data" / "master" / "cv_master.json"

print(f"📂 Loading {JSON_PATH.name}…\n")
data = json.loads(JSON_PATH.read_text(encoding="utf-8"))

# ─── Structural sanity ───────────────────────────────────────────────────
print("=" * 70)
print("PASS 1: ESTRUCTURA DEL JSON")
print("=" * 70)
expected_keys = {"_meta", "instrucciones", "perfil", "experiencia", "logros",
                 "skills", "educacion", "certificaciones", "voluntariado",
                 "empresas_target", "perfiles_narrativos", "conferencias",
                 "reconocimientos", "recipes", "anti_patterns"}
got = set(data.keys())
missing = expected_keys - got
extra = got - expected_keys
print(f"✓ Top-level keys: {len(got)}")
if missing:
    print(f"  ❌ FALTAN: {missing}")
if extra:
    print(f"  ⚠ Extra: {extra}")

# ─── Meta ─────────────────────────────────────────────────────────────────
print(f"\nMeta:")
for k, v in data["_meta"].items():
    print(f"  {k}: {v}")

# ─── Perfil ───────────────────────────────────────────────────────────────
print(f"\nPerfil ({len(data['perfil'])} secciones):")
for sec in data["perfil"]:
    print(f"  · {sec}")

# ─── Sanity de logros ─────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("PASS 2: SANITY DE LOGROS")
print("=" * 70)
logros = data["logros"]
print(f"Total logros: {len(logros)}")

# Check: every logro has ES and EN
no_es = [l["id"] for l in logros if not l.get("es")]
no_en = [l["id"] for l in logros if not l.get("en")]
print(f"  Sin ES: {len(no_es)} {no_es if no_es else ''}")
print(f"  Sin EN: {len(no_en)} {no_en if no_en else ''}")

# Check: seniority fit populated
no_sen = [l["id"] for l in logros if not l.get("seniority_fit")]
print(f"  Sin Seniority Fit: {len(no_sen)} {no_sen if no_sen else ''}")

# Check: story populated (not required, FYI only)
no_story = [l["id"] for l in logros if not l.get("story_evidence")]
print(f"  Sin Story/Evidence: {len(no_story)} (FYI, no required)")

# Check: industria/skill tags populated
no_ind = [l["id"] for l in logros if not l.get("tags_industria")]
no_skl = [l["id"] for l in logros if not l.get("tags_skill")]
print(f"  Sin Tags Industria: {len(no_ind)} {no_ind[:10] if no_ind else ''}")
print(f"  Sin Tags Skill: {len(no_skl)} {no_skl[:10] if no_skl else ''}")

# Check: incluir flags populated
no_inc_es = sum(1 for l in logros if l.get("incluir_cv_es") is None)
no_inc_en = sum(1 for l in logros if l.get("incluir_cv_en") is None)
print(f"  Sin flag Incluir CV ES: {no_inc_es}")
print(f"  Sin flag Incluir CV EN: {no_inc_en}")

# Pair With references
all_ids = {l["id"] for l in logros}
broken_pairs = []
for l in logros:
    for p in (l.get("pair_with") or []):
        if p not in all_ids:
            broken_pairs.append((l["id"], p))
print(f"  Broken Pair With references: {len(broken_pairs)} {broken_pairs if broken_pairs else ''}")

# ─── Sanity de experiencias ──────────────────────────────────────────────
print("\n" + "=" * 70)
print("PASS 3: SANITY DE EXPERIENCIAS")
print("=" * 70)
exp = data["experiencia"]
print(f"Total puestos: {len(exp)}")
no_emp = [e["id"] for e in exp if not e.get("empresa")]
no_pos = [e["id"] for e in exp if not e.get("puesto_es") and not e.get("puesto_en")]
print(f"  Sin empresa: {len(no_emp)} {no_emp if no_emp else ''}")
print(f"  Sin puesto (ni ES ni EN): {len(no_pos)} {no_pos if no_pos else ''}")
no_fi = [e["id"] for e in exp if not e.get("fecha_inicio")]
no_ff = [e["id"] for e in exp if not e.get("fecha_fin")]
print(f"  Sin fecha inicio: {len(no_fi)} {no_fi if no_fi else ''}")
print(f"  Sin fecha fin: {len(no_ff)} {no_ff if no_ff else ''}")

# ─── Recipes deep check ──────────────────────────────────────────────────
print("\n" + "=" * 70)
print("PASS 4: RECIPES (referencias y completitud)")
print("=" * 70)
all_logro_ids = {l["id"] for l in logros}
all_perfil_idx = {n.get("tipo_empresa") for n in data["perfiles_narrativos"]}
for rec in data["recipes"]:
    print(f"\n  {rec['id']} · {rec['name']}")
    # check logros
    bad_logros = [lid for lid in (rec.get("logros") or []) if lid not in all_logro_ids]
    if bad_logros:
        print(f"    ❌ Logros faltantes: {bad_logros}")
    else:
        print(f"    ✓ {len(rec.get('logros') or [])} logros, todos válidos")
    # check perfil narrativo
    perfil_ref = rec.get("perfil_narrativo")
    if perfil_ref and perfil_ref not in all_perfil_idx:
        print(f"    ⚠ Perfil narrativo '{perfil_ref}' NO existe en sheet 09")
        print(f"      (existentes: {all_perfil_idx})")
    else:
        print(f"    ✓ Perfil narrativo '{perfil_ref}' válido")
    # check notes
    if not rec.get("notas_llm"):
        print(f"    ⚠ Sin notas para LLM")
    else:
        print(f"    ✓ Notas LLM: {rec['notas_llm'][:60]}…")

# ─── Anti-patterns sanity ────────────────────────────────────────────────
print("\n" + "=" * 70)
print("PASS 5: ANTI-PATTERNS")
print("=" * 70)
ap = data["anti_patterns"]
print(f"Total: {len(ap)}")
incomplete = []
for a in ap:
    if not all([a.get("id"), a.get("categoria"), a.get("evita"),
                a.get("usa_en_lugar"), a.get("razon")]):
        incomplete.append(a.get("id"))
if incomplete:
    print(f"  ❌ Incompletas: {incomplete}")
else:
    print(f"  ✓ Todos los anti-patterns tienen los 5 campos completos")
print(f"  Lista de IDs: {[a['id'] for a in ap]}")

# ─── SIMULACIÓN: generar CV usando Recipe R04 ────────────────────────────
print("\n" + "=" * 70)
print("PASS 6: SIMULACIÓN — Generar CV usando Recipe R04 (Head Research US Tech/AdTech)")
print("=" * 70)
recipe = next((r for r in data["recipes"] if r["id"] == "R04"), None)
if not recipe:
    print("❌ Recipe R04 no encontrada")
    sys.exit(1)

# Resolve logros
logros_by_id = {l["id"]: l for l in logros}
selected_logros = [logros_by_id[lid] for lid in recipe["logros"] if lid in logros_by_id]

# Filter: idioma EN, incluir_cv_en=True, seniority director/exec
director_logros = [l for l in selected_logros
                   if l.get("incluir_cv_en")
                   and any(s in (l.get("seniority_fit") or [])
                           for s in ("director", "exec", "senior", "all"))]

# Get perfil narrativo
perfil = next((p for p in data["perfiles_narrativos"]
               if p.get("tipo_empresa") == recipe["perfil_narrativo"]), None)

# Get bio EN
bio_en = data["perfil"].get("PROFESSIONAL SUMMARY (EN)", {})

# Build sample CV
print()
print("─" * 70)
print(f"# {data['perfil'].get('DATOS DE CONTACTO', {}).get('Nombre completo', '?')}")
print(f"## {bio_en.get('Headline (1 line)', '?')}")
print()
print("**Executive Summary:**")
print(bio_en.get("Executive bio (3-5 lines)", "?"))
print()
print(f"**Profile (per Recipe R04 / target: {recipe['name']}):**")
if perfil:
    print(perfil.get("perfil_en", "?")[:300] + "…")
print()
print(f"**Selected experience & achievements** ({len(director_logros)} bullets):")
exp_by_id = {e["id"]: e for e in exp}
by_company = {}
for l in director_logros:
    by_company.setdefault(l["id_empresa"], []).append(l)

for eid in sorted(by_company.keys(), key=lambda x: exp_by_id.get(x, {}).get("fecha_inicio") or "", reverse=True):
    e = exp_by_id.get(eid)
    if not e:
        continue
    print(f"\n### {e['empresa']} · {e.get('puesto_en') or e.get('puesto_es')}")
    print(f"_{e.get('fecha_inicio')} → {e.get('fecha_fin')}_")
    for l in by_company[eid][:5]:  # max 5 per role
        print(f"  • {l['en']}")
print()
print("─" * 70)
print(f"\n✅ SIMULACIÓN EXITOSA: se generó un CV preliminar usando Recipe R04")
print(f"   {len(director_logros)} logros seleccionados para nivel director/exec/senior")
print(f"   {len(by_company)} empresas representadas")

print("\n" + "=" * 70)
print("✅ QA FINAL COMPLETO")
print("=" * 70)
