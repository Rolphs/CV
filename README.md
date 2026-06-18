# 📄 CV-RMB · CV Consolidado Raúl Mercado

Proyecto para consolidar, parsear y/o presentar el archivo histórico de CVs de Raúl Mercado Bustamante.

## 📁 Estructura

```
cv_rmb/
├── data/                                  ← TODO el dato del proyecto
│   ├── raw/                                ← INPUTS originales (read-only, jamás se tocan)
│   │   ├── 01_version_actual_2026/         CV vigente (PDF + DOCX)
│   │   ├── 02_versiones_2025/              variantes 2025 (10 PDFs/DOCX)
│   │   ├── 03_versiones_por_empresa/       CVs targeted (Netflix, NVT, NWST, PP, PPSI, PSI)
│   │   ├── 04_bases_datos_excel/           Formato Datos CV - Raul Mercado.xlsx
│   │   └── 05_versiones_anteriores/        CV viejo
│   ├── master/                             ← EL PRODUCTO (fuente de verdad para futuro)
│   │   ├── cv_master_raul_mercado.xlsx     ⭐ canonical: 24 puestos · 69 logros · 14 sheets · LLM-ready
│   │   ├── cv_master.json                  ⭐ JSON export para consumo directo de LLM
│   │   └── backups/                        snapshots versionados pre cada fase (v1-v8)
│   └── reports/                            ← DERIVATIVES (regenerables con scripts/)
│       ├── audit_summary.md                gap analysis ejecutivo
│       ├── file_catalog.{md,json}          catálogo de 53 archivos raw
│       ├── xlsx_full_dump.{md,json}        dump completo del master + Formato Datos
│       ├── master_full_dump.{md,json}      snapshot detallado del master
│       ├── master_audit.md                 hallazgos de auditoría convolucional
│       ├── validation_report.md            ✅ último resultado del validador
│       ├── logros_phase3_candidates.{md,json}   candidatos extraídos en Fase 3
│       └── duplicates.md
├── scripts/                                ← CÓDIGO · scripts de consolidación + gestión de aplicaciones
├── cv_renderer/                            ← motor de render MD → PDF/DOCX/TXT/JSON + keywords audit
├── applications_manager/                   ← plantillas y helpers del tracker de aplicaciones
├── tests/                                   ← unittest (stdlib, sin pytest)
├── docs/                                   ← EDITORIAL_FLOW · SCRIPTS_REFERENCE · SKILL recipe
├── .venv/                                  ← venv local (no versionado)
├── requirements.txt                        ← dependencias runtime
├── .gitignore
└── README.md                               este archivo
```

> ℹ️ La carpeta `data/` (raw + master + reports) contiene PII real y **no está
> versionada** en el repo. Hay que proveerla localmente para correr los scripts
> de consolidación o de generación de CV. El motor de render se puede probar sin
> ella usando los samples (`cv_renderer/sample_cv.md` + `sample_jd.txt`).

## 🚦 Reglas de oro

1. **`data/raw/` es read-only.** Nunca editar archivos originales. Son los CVs como llegaron del ZIP.
2. **`data/master/cv_master_raul_mercado.xlsx` es la fuente de verdad.** Todo lo demás se deriva de aquí.
3. **`data/reports/` es regenerable.** Se puede borrar y volver a generar corriendo los scripts.
4. **Antes de modificar el master, hacer backup** a `data/master/backups/vN_pre_<motivo>.xlsx`.
5. Cualquier HTML final listo para compartir → copia también a `../../03_reports_published/`.
6. Si se necesita venv → `uv venv .venv` dentro del proyecto, nunca en root.

## 🔄 Estado

Ver `data/reports/audit_summary.md` para el gap analysis original (si tienes la carpeta `data/` cargada localmente).

- ✅ Fase 0 · Audit & catalog (53 archivos analizados)
- ✅ Fase 1 · Limpieza (intrusos movidos, vacíos borrados, duplicados archivados)
- ✅ Fase 2 · Consolidación Experiencia (8 → 18 puestos · 18 → 48 logros)
- ✅ Fase 3 · Extracción logros desde CVs (48 → 66 logros, +Altazor curado)
- ✅ Fase 3.5 · Curación: dedupe + traducciones (66 → 61 logros sin placeholders)
- ✅ Fase 4 · Walmart Scintilla + EN translations + Bio ejecutiva (19 puestos / 63 logros 100% bilingües)
- ✅ Refactor estructura · `source/cv_consolidado/` + `outputs/` → `data/{raw,master,reports}/` + `scripts/`
- ✅ Fase 5 · Hallazgos históricos 2010-2022 (27 puestos / 7 educación / +sheets Conferencias y Reconocimientos)
- ✅ Fase 6 · Auditoría convolucional + cirugía contenido (27→24 puestos sin duplicados, 63→69 logros, todos con verbo de acción ES/EN)
- ✅ Fase 7 · Polish visual paleta Walmart (títulos, headers, banded rows, tab colors, freeze panes)
- ✅ Fase 8 · **LLM-ready** (manual operativo en sheet 00, columnas Seniority Fit/Story/Pair, vocabulario controlado de tags, sheet 12 Recipes, sheet 13 Anti-Patterns, export JSON, validador)
- ✅ Fase 9 · **QA fixes** (smoke test end-to-end encontró 3 bugs reales que el validador no detectaba: recipes con perfil narrativo inválido, 30 logros sin tags ni flag Incluir EN, parser JSON de 01 Perfil mal)

## ⚙️ Setup

Requiere **Python 3.11+**. Desde el root del proyecto:

```bash
# 1. Crear y activar venv
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Smoke test del motor de render (no necesita data/)
python scripts/render_cv.py cv_renderer/sample_cv.md /tmp/render_test \
    --jd cv_renderer/sample_jd.txt

# 4. Correr los tests
python -m unittest discover tests
```

## 🛠 Scripts del proyecto

Todos en `scripts/`. Corren desde el root del proyecto:

```bash
python scripts/<nombre>.py
```

| Script | Qué hace |
|---|---|
| `audit_files.py` | Cataloga todos los archivos en `data/raw/` (metadata + extracción de texto) |
| `dump_xlsx.py` | Vuelca contenido completo de xlsx clave a `data/reports/` |
| `consolidate_phase2.py` | Migra puestos faltantes y logros desde `Formato Datos CV` al maestro |
| `extract_logros_phase3.py` | Extrae candidatos de logros desde todos los CVs (PDF/DOCX) |
| `apply_phase3.py` | Aplica candidatos curados al maestro (con backup automático) |
| `curate_phase3_5.py` | Dedupe semántico + traducciones + limpieza de ligatures PDF |
| `apply_phase4.py` | Walmart Scintilla + 30 traducciones EN + Bio ejecutiva ES/EN |
| `normalize_historical.py` | Aplana, limpia y normaliza la carpeta histórica 2010-2022 |
| `audit_historical.py` | Extrae texto y analiza patrones en CVs históricos |
| `apply_phase5_historical.py` | Aplica hallazgos históricos al master (+8 puestos, +sheets 10 y 11) |
| `audit_master_deep.py` | Auditoría convolucional: integridad referencial, paridad ES/EN, cronología, redacción |
| `apply_phase6_polish.py` | Cirugía de contenido: consolida duplicados, reescribe con verbos de acción, normaliza nombres |
| `apply_phase7_visual.py` | Polish visual con paleta Walmart (títulos, headers, banded rows, freeze panes, tab colors) |
| `apply_phase8_llm_ready.py` | Manual LLM en sheet 00, columnas extra en logros, vocab controlado, sheets 12 Recipes y 13 Anti-Patterns |
| `apply_phase9_qa_fixes.py` | Bug fixes encontrados por smoke test (recipes perfil ref, tags faltantes en 30 logros) |
| `export_master_json.py` | 🔁 **Re-ejecutable**: regenera `data/master/cv_master.json` desde el xlsx (LLM-native format) |
| `validate_master.py` | 🔁 **Re-ejecutable**: corre 8 checks (integridad, paridad, vocab, fechas, etc.) → `data/reports/validation_report.md` |
| `qa_smoke_test.py` | 🔁 **Re-ejecutable**: smoke test end-to-end que simula ser un LLM consumiendo el JSON para generar un CV. Catch issues que el validator no detecta |
| `preflight_check.py` | 🔁 **Re-ejecutable**: pre-flight exhaustivo de 12 pasadas (coverage por empresa, distribuciones por seniority/industria/skill, deep check de recipes con preview del context bundle que el LLM vería) |

## 🤖 Uso con LLM

El master está diseñado para ser consumido por un LLM (Claude, GPT-4) para generar CVs especializados:

1. **Lee primero `sheet 00 Instrucciones`** del xlsx (o el campo `instrucciones` del JSON). Es el manual operativo con reglas críticas (no inventar métricas, respetar toggles, etc.).
2. **Busca en `sheet 12 Recipes`** una receta que se ajuste al puesto target. Si existe, úsala como base.
3. **Filtra `sheet 03 Logros`** por `Tags Industria`, `Tags Habilidad`, `Seniority Fit`, e `Incluir CV ES/EN`.
4. **Verifica contra `sheet 13 Anti-Patterns`** que no estés usando lenguaje débil o claims problemáticos.
5. **Output**: CV en el idioma + formato que pida el usuario.

Flujo recomendado para mantener el master:
```bash
# 1. Editar el xlsx en Excel cuando agregues nuevos logros/puestos
# 2. Re-exportar JSON + tres niveles de check
python scripts/export_master_json.py
python scripts/validate_master.py     # schema (8 checks)
python scripts/qa_smoke_test.py       # uso LLM
python scripts/preflight_check.py     # preflight (13 checks)
# 3. Pasar el xlsx O el JSON al LLM con tu prompt de generación de CV
```

## 🎯 Generar y trackear aplicaciones CV

El proyecto incluye un **sistema completo de gestión de aplicaciones**: una carpeta por CV enviado, con metadata, JD snapshot, notas privadas, todos los formatos de output y tracking de status pipeline.

### Estructura por aplicación

```
output/
├── _registry.json              ← manifest central (auto-generado)
├── _dashboard.md               ← dashboard con métricas (auto-generado)
└── applications/
    └── 2026-05-17_R04_triplelift_head-of-research/
        ├── application.yaml    metadata + status + timeline + outcome
        ├── notes.md            notas privadas (contactos, salary, referidos)
        ├── jd.txt              snapshot del job description
        ├── cv.md               source markdown (LLM lo genera, tú lo editas)
        ├── cv.pdf / cv.docx / cv.txt / cv.json      outputs renderizados
        └── cv_keywords.md      audit del match vs JD
```

### Pipeline de status (9 estados)

```
📝 draft → ✅ ready → 📤 submitted → 📞 callback → 🎤 interview → 🎯 offer
                                       ↘ ❌ rejected  ↘ 🚪 withdrawn  ↘ 👻 ghosted
```

### Flujo end-to-end de una aplicación

```bash
# 1. Crear skeleton de la aplicación
python scripts/new_application.py \
    --company "TripleLift" --role "Head of Research" --recipe R04 \
    --jd path/to/jd.txt

# 2. LLM genera cv.md dentro de la carpeta usando cv_master.json + recipe + jd.txt
#    (tú revisas/editas el cv.md a gusto)

# 3. Renderizar los 5 formatos (auto-sincroniza application.yaml con match_rate)
python scripts/render_cv.py \
    output/applications/<app-id>/cv.md \
    output/applications/<app-id> \
    --jd output/applications/<app-id>/jd.txt

# 4. Cuando lo envíes, transition al siguiente estado
python scripts/update_status.py <app-id> ready --note "match 79%"
python scripts/update_status.py <app-id> submitted --note "via LinkedIn"
python scripts/update_status.py <app-id> callback --note "Lisa Park called"

# 5. Refresh dashboard
python scripts/rebuild_dashboard.py

# 6. Consultar / buscar
python scripts/list_applications.py --status callback
python scripts/list_applications.py --company TripleLift --long
python scripts/search_applications.py "Hispanic" --in cv
python scripts/search_applications.py "salary" --in notes
```

### Scripts de gestión

| Script | Función |
|---|---|
| `new_application.py` | Crea skeleton con `application.yaml`, `notes.md`, `jd.txt` |
| `render_cv.py` | MD → PDF/DOCX/TXT/JSON/keywords-audit + auto-sync con `application.yaml` |
| `update_status.py` | Transition de status + timeline append (auto-llena callback_days) |
| `list_applications.py` | Lista con filtros `--status`, `--company`, `--recipe`, `--since`, `--long` |
| `rebuild_dashboard.py` | Regenera `_dashboard.md` (funnel, by-recipe) y `_registry.json` |
| `search_applications.py` | Full-text en CV/JD/notes con line numbers y status icon |

## 📜 Historia

- **2026-05-17**: Proyecto creado y consolidado en 8 sesiones (audit → 5 fases de consolidación → 3 fases de polish: visual + content + LLM-ready).
- Estado actual: maestro **LLM-ready** con 14 sheets, 24 puestos, 69 logros bilingües enriquecidos con Seniority Fit / Story / Pair With, vocabulario controlado, 6 recipes de CV pre-armadas, 13 anti-patterns, manual operativo para LLM en sheet 00, JSON export y validador automatizado.
