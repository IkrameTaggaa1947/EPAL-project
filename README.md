# EPAL — TUA Extraction System

Reads Portuguese environmental discharge licences (**TUA**, the SILIAMB PDFs) and
turns each one into a structured, validated table that feeds a Power BI compliance
dashboard.

It reads each certificate by its **table structure** — it reads the ruled tables
directly, matches columns by **header name** (not pixel position), locates sections
by their **marker text** (`EXP8.3.x` / `EXP9.3.x`), and **verifies every field**
before accepting it. This is robust to layout differences between certificates and
flags anything doubtful instead of failing silently.

## Project structure

```
EPAL-project/
├── extraction/tua/                # the extractor
│   ├── extract_tua.py             #  ★ core: PDF → structured data (5-sheet model)
│   ├── qa.py                      #  quality checks (verdict + filename cross-check)
│   ├── enrich.py                  #  joins the static lookup tables (Critério, ARH, …)
│   ├── pipeline.py                #  batch / watch-folder orchestrator
│   ├── test_extraction.py         #  regression tests (run: python test_extraction.py)
│   ├── run_tua.bat                #  process the inbox once (double-click)
│   ├── watch_tua.bat              #  watch the inbox every 5s (double-click, leave open)
│   ├── register_watch_task.bat    #  optional: run automatically in the background
│   ├── static_tables.xlsx         #  lookup tables (Criteria / GNA / AdvT / Frequencies / Requests)
│   └── *_extracted.json           #  frozen "golden" outputs used as the test baseline
│
├── samples/                       # sample TUA PDFs for testing
├── docs/                          # OPERATIONS.md (operator guide), notes, schemas
├── requirements.txt
└── .gitignore
```

Generated outputs (`master_tua.xlsx`, `master_all_sections.xlsx`, `powerbi_csv/`,
`qa_report.xlsx`, the `data/` runtime folder, backups) are **git-ignored** — you regenerate them by running the
pipeline. The committed `*_extracted.json` files are the regression-test baseline.

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate            # Windows  (source .venv/bin/activate on macOS/Linux)
pip install -r requirements.txt
```
Requires Python 3.8+. Core dependency: PyMuPDF (`fitz`), pandas, openpyxl.

## Usage

**One certificate (command line):**
```bash
cd extraction/tua
python extract_tua.py "../../samples/<certificate>.pdf"
```

**A folder of certificates, hands-off (recommended):**
Double-click **`watch_tua.bat`** and leave it open, then drop PDFs into
`extraction/tua/data/inbox`. Each is extracted within ~5s, quality-checked, added to
the master, and its result opens in Excel. Anything doubtful is filed into
`data/review/` with a plain-language note. See **[docs/OPERATIONS.md](docs/OPERATIONS.md)**.

## What it produces
Two levels — the **full extraction** (every section) and the **Power BI feed** (only
what the dashboard needs):

- `<name>_extracted.xlsx` — per-certificate detail: a **Resumo** identity card plus
  **one sheet per section** (3.3 → 3.21, all 12), generically dumped so nothing is dropped.
- `master_all_sections.xlsx` — the same, **consolidated**: one sheet per section with
  every certificate stacked and keyed on `Nº TUA`. The complete source that Power BI trims.
- `master_tua.xlsx` — the curated **Power BI star schema**, 5 linked sheets (**Licenses ·
  Conditions · Autocontrolo · Legislacao · Avaliacao**), keyed on `Nº TUA`, upserted.
- `powerbi_csv/` — the star schema **enriched** with the static-table columns (Critério
  1–7, ARH, Estado, Nº análises, Siliamb date). **Point Power BI here.**
- `qa_report.xlsx` — one row per certificate with its verdict (**OK / REVIEW / FAILED**) and reasons.

## How it works (short version)
`PDF → read ruled tables (one pass) → split into all sections by marker (row-level,
page-break-aware) → for the Power BI fields: match columns by name, resolve legend
references, verify every field (incl. filename cross-check) → write the per-file
workbook (Resumo + every section), upsert the all-sections + 5-sheet masters, export
the enriched Power BI CSVs.`

## Tests
```bash
cd extraction/tua
python test_extraction.py      # or: pytest
```

## Author
Ikrame Taggaa
