"""
Pre-flight check exhaustivo · usa schema REAL del JSON exportado.

Schema descubierto:
  - experiencia[].id, empresa, puesto_es, puesto_en, fecha_inicio, fecha_fin, ...
  - logros[].id, id_empresa, empresa, es, en, tags_industria, tags_skill,
            seniority_fit, story_evidence, pair_with, incluir_cv_es/en
  - recipes[].id, name, when_to_use, perfil_narrativo, logros (lista de IDs),
              skills_categorias, notas_llm
  - skills[].categoria, es, en, nivel, anos_exp, destacar_en
  - anti_patterns[].id, categoria, evita, usa_en_lugar, razon
"""
import json
from pathlib import Path
from collections import Counter, defaultdict
import sys
sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
data = json.loads((ROOT / "data" / "master" / "cv_master.json").read_text(encoding="utf-8"))


def banner(title, ch="="):
    print(f"\n{ch * 78}\n  {title}\n{ch * 78}")


def as_list(v):
    """Coerce v to list of strings: handles None, str, list."""
    if v is None or v == "":
        return []
    if isinstance(v, list):
        return [str(x).strip() for x in v if str(x).strip()]
    return [t.strip() for t in str(v).replace(";", ",").split(",") if t.strip()]


# ─── 1 · ESTADO BASE ───
banner("1 · ESTADO BASE")
counts = {
    "Experiencias (puestos)": len(data.get("experiencia", [])),
    "Logros":                 len(data.get("logros", [])),
    "Skills":                 len(data.get("skills", [])),
    "Educación":              len(data.get("educacion", [])),
    "Certificaciones":        len(data.get("certificaciones", [])),
    "Voluntariado":           len(data.get("voluntariado", [])),
    "Empresas target":        len(data.get("empresas_target", [])),
    "Perfiles narrativos":    len(data.get("perfiles_narrativos", [])),
    "Conferencias":           len(data.get("conferencias", [])),
    "Reconocimientos":        len(data.get("reconocimientos", [])),
    "Recipes":                len(data.get("recipes", [])),
    "Anti-patterns":          len(data.get("anti_patterns", [])),
    "Instrucciones LLM":      len(data.get("instrucciones", [])),
}
for k, v in counts.items():
    print(f"  {k:<25} {v}")


# Lookups
logros_by_id = {l["id"]: l for l in data["logros"]}
empresas_in_exp = Counter(e.get("empresa") for e in data["experiencia"])
logros_per_empresa = Counter(l.get("empresa") for l in data["logros"])


# ─── 2 · COVERAGE DE LOGROS POR EMPRESA ───
banner("2 · COVERAGE DE LOGROS POR EMPRESA")
print("  (en el master, los logros se anclan a EMPRESA, no a puesto específico)\n")
all_empresas = set(empresas_in_exp) | set(logros_per_empresa)
sin_logros = []
for emp in sorted(all_empresas, key=lambda x: -empresas_in_exp.get(x, 0)):
    n_puestos = empresas_in_exp.get(emp, 0)
    n_logros = logros_per_empresa.get(emp, 0)
    flag = ""
    if n_puestos > 0 and n_logros == 0:
        flag = "  ⚠ Sin logros"
        sin_logros.append(emp)
    print(f"  {emp:<40} {n_puestos:>2} puestos · {n_logros:>2} logros{flag}")


# ─── 3 · SENIORITY ───
banner("3 · DISTRIBUCIÓN DE LOGROS POR SENIORITY FIT")
sen_count = Counter()
for l in data["logros"]:
    for t in as_list(l.get("seniority_fit")):
        sen_count[t] += 1
max_v = max(sen_count.values()) if sen_count else 1
for sen, n in sen_count.most_common():
    bar = "\u2593" * (n * 40 // max_v)
    print(f"  {sen:<22} {n:>3}  {bar}")


# ─── 4 · INDUSTRIA ───
banner("4 · DISTRIBUCIÓN DE LOGROS POR INDUSTRIA")
ind_count = Counter()
for l in data["logros"]:
    for t in as_list(l.get("tags_industria")):
        ind_count[t] += 1
max_v = max(ind_count.values()) if ind_count else 1
for ind, n in ind_count.most_common():
    bar = "\u2593" * (n * 40 // max_v)
    print(f"  {ind:<28} {n:>3}  {bar}")


# ─── 5 · SKILLS USADAS EN LOGROS ───
banner("5 · TOP SKILLS USADAS EN LOGROS")
skl_count = Counter()
for l in data["logros"]:
    for t in as_list(l.get("tags_skill")):
        skl_count[t] += 1
print(f"  Skills únicos tagged en logros: {len(skl_count)}")
print()
max_v = max(skl_count.values()) if skl_count else 1
for skl, n in skl_count.most_common(20):
    bar = "\u2593" * (n * 30 // max_v)
    print(f"  {skl:<30} {n:>3}  {bar}")


# ─── 6 · COVERAGE POR RECIPE (modelo determinista) ───
banner("6 · COVERAGE POR RECIPE")
print("  Los recipes pre-seleccionan IDs de logros (no filtran dinámicamente).\n")
recipe_health = []
for r in data["recipes"]:
    rid = r["id"]
    name = r.get("name", "?")
    perfil = r.get("perfil_narrativo", "?")
    logro_ids = r.get("logros") or []
    if isinstance(logro_ids, str):
        logro_ids = [int(x.strip()) for x in logro_ids.split(",") if x.strip().isdigit()]
    skills_cats = as_list(r.get("skills_categorias"))

    # Validate IDs exist
    missing = [lid for lid in logro_ids if lid not in logros_by_id]
    n_logros = len(logro_ids)
    status = "✅" if n_logros >= 15 and not missing else ("⚠" if n_logros >= 10 else "❌")
    recipe_health.append((rid, n_logros, len(missing)))

    print(f"  {status} {rid}: {name}")
    print(f"     Perfil narrativo: {perfil}")
    print(f"     Logros pre-seleccionados: {n_logros}", end="")
    if missing:
        print(f"  ⚠ {len(missing)} IDs no existen en master: {missing}")
    else:
        print()
    print(f"     Skills categorías: {skills_cats}")

    # Distribución por empresa de los logros del recipe
    by_emp = Counter()
    for lid in logro_ids:
        if lid in logros_by_id:
            by_emp[logros_by_id[lid].get("empresa")] += 1
    if by_emp:
        emp_summary = ", ".join(f"{e}({n})" for e, n in by_emp.most_common(5))
        print(f"     Empresas: {emp_summary}{' …' if len(by_emp) > 5 else ''}")
    print()


# ─── 7 · CONSISTENCIA SKILLS DECLARADAS vs LOGROS ───
banner("7 · CONSISTENCIA: SKILLS DECLARADAS vs USADAS EN LOGROS")
declared = {(s.get("es") or "").lower().strip() for s in data["skills"]}
declared |= {(s.get("en") or "").lower().strip() for s in data["skills"]}
declared.discard("")
used = {t.lower().strip() for t in skl_count.keys()}
orphan = used - declared
print(f"  Skills declaradas (sheet 04): {len(declared) // 2} únicas (ES+EN)")
print(f"  Skills usadas en tags logros: {len(used)}")
if orphan:
    print(f"\n  ℹ Tags de logros NO presentes literalmente en sheet 04 ({len(orphan)}):")
    print(f"    (esto NO es necesariamente un bug — los tags son más granulares que las categorías de skill)")
    for t in sorted(orphan):
        print(f"     · {t}")
else:
    print("  ✅ Vocabulario perfectamente alineado")


# ─── 8 · NARRATIVA ───
banner("8 · ENRIQUECIMIENTO NARRATIVO")
n_story = sum(1 for l in data["logros"] if l.get("story_evidence"))
n_pair = sum(1 for l in data["logros"] if l.get("pair_with"))
n_metric = sum(1 for l in data["logros"]
               if str(l.get("tiene_metrica") or "").lower() in ("sí", "si", "yes", "true"))
n_impact = sum(1 for l in data["logros"] if l.get("impacto"))
print(f"  Logros con story_evidence: {n_story:>3}/{len(data['logros'])}  ({n_story*100//len(data['logros'])}%)")
print(f"  Logros con pair_with:      {n_pair:>3}/{len(data['logros'])}  ({n_pair*100//len(data['logros'])}%)")
print(f"  Logros con métrica (KPI):  {n_metric:>3}/{len(data['logros'])}  ({n_metric*100//len(data['logros'])}%)")
print(f"  Logros con impacto:        {n_impact:>3}/{len(data['logros'])}  ({n_impact*100//len(data['logros'])}%)")


# ─── 9 · ANTI-PATTERNS ───
banner("9 · ANTI-PATTERNS · resumen")
for ap in data["anti_patterns"]:
    cat = ap.get("categoria", "?")
    evita = ap.get("evita", "?")[:80]
    print(f"  [{cat:<18}] evita: {evita}")


# ─── 10 · CONTEXT BUNDLE preview · cada Recipe ───
banner("10 · PREVIEW DEL OUTPUT QUE EL LLM VERÍA · Recipe R04")
r04 = next(r for r in data["recipes"] if r["id"] == "R04")
print(f"\n  Recipe: {r04['name']}")
print(f"  When to use: {r04.get('when_to_use', '')[:200]}")
print(f"  Perfil narrativo: {r04.get('perfil_narrativo', '?')}")
print(f"  Notas LLM: {r04.get('notas_llm', '')[:200]}")

logro_ids = r04.get("logros") or []
if isinstance(logro_ids, str):
    logro_ids = [int(x.strip()) for x in logro_ids.split(",") if x.strip().isdigit()]

print(f"\n  ── {len(logro_ids)} logros que el LLM consumiría (en orden de receta) ──")
by_emp_ordered = defaultdict(list)
for lid in logro_ids:
    if lid in logros_by_id:
        by_emp_ordered[logros_by_id[lid]["empresa"]].append(logros_by_id[lid])
for emp, lst in by_emp_ordered.items():
    print(f"\n    [{emp}]  ({len(lst)} logros)")
    for l in lst:
        text = (l.get("en") or l.get("es") or "")[:130].replace("\n", " ")
        sen = ", ".join(as_list(l.get("seniority_fit")))[:30]
        print(f"      #{l['id']:<3} [{sen:<22}] {text}{'…' if len(text) >= 130 else ''}")


# ─── 11 · INSTRUCCIONES LLM ───
banner("11 · MANUAL DE INSTRUCCIONES EMBEBIDO (sheet 00)")
print(f"  El JSON contiene {len(data.get('instrucciones', []))} secciones de instrucciones.")
for i, sec in enumerate(data.get("instrucciones", []), 1):
    title = sec.get("section", "?") if isinstance(sec, dict) else str(sec)[:60]
    n_items = len(sec.get("items", [])) if isinstance(sec, dict) else 0
    print(f"  {i}. {title}  ({n_items} items)")


# ─── 12 · CHECKLIST FINAL ───
banner("12 · CHECKLIST FINAL · ¿LISTO PARA PRIMER CV?")
n_logros = len(data["logros"])
n_puestos = len(data["experiencia"])
recipe_id_issues = sum(1 for _, _, miss in recipe_health if miss > 0)
recipe_size_issues = sum(1 for _, n, _ in recipe_health if n < 10)  # 10 es suficiente para CV

# Empresas "histórico-juvenil" donde NO esperamos logros tipo C-suite
HISTORICAL_OK = {
    "Mago Ilusionista (Independiente)",
    "Periódico Reforma",
    "Cinemex",
    "Estudio Fotográfico Trabulsi",
}
sin_logros_reales = [e for e in sin_logros if e not in HISTORICAL_OK]

checks = [
    (f"≥60 logros (actual: {n_logros})", n_logros >= 60),
    (f"≥20 puestos (actual: {n_puestos})", n_puestos >= 20),
    (f"≥4 perfiles narrativos (actual: {len(data['perfiles_narrativos'])})",
     len(data["perfiles_narrativos"]) >= 4),
    (f"≥5 recipes (actual: {len(data['recipes'])})", len(data["recipes"]) >= 5),
    (f"≥10 anti-patterns (actual: {len(data['anti_patterns'])})",
     len(data["anti_patterns"]) >= 10),
    (f"≥15 skills (actual: {len(data['skills'])})", len(data["skills"]) >= 15),
    (f"Empresas modernas sin logros (excluye contexto histórico): {len(sin_logros_reales)}",
     len(sin_logros_reales) == 0),
    ("Recipes: todos los IDs de logros existen", recipe_id_issues == 0),
    ("Recipes: todos con ≥10 logros pre-seleccionados", recipe_size_issues == 0),
    (f"≥50 logros con story_evidence (narrativa rica) — actual: {n_story}", n_story >= 50),
    (f"≥15 logros con métrica/KPI hard (no inflar) — actual: {n_metric}", n_metric >= 15),
    ("Manual de instrucciones LLM embebido", len(data.get("instrucciones", [])) >= 5),
    ("Perfil base (datos personales) poblado",
     bool(data.get("perfil")) and len(data["perfil"]) >= 3),
]
for desc, ok in checks:
    print(f"  {'✅' if ok else '❌'} {desc}")
n_ok = sum(1 for _, ok in checks if ok)
total = len(checks)
print()
if n_ok == total:
    print("  🚀 TODO VERDE · LISTO PARA GENERAR PRIMER CV")
else:
    print(f"  ⚠ {total - n_ok}/{total} ítems requieren atención antes del primer CV")
