# EPAL LURH extractor

Extracts the data from **LURH** certificates — *Licença de Utilização dos
Recursos Hídricos — Rejeição de Águas Residuais* (SILIAMB / APA) — into
per-file Excel + JSON, a consolidated multi-sheet master workbook, and
Power BI CSVs. It is the LURH twin of the TUA extractor in
`H:\EPAL\EPAL-project\extraction\tua` and follows the same conventions, so
operators and Power BI treat both the same way.

Target location: `H:\EPAL\EPAL-project\extraction\lurh`
(the `.bat` files look for the shared virtualenv two levels up, at
`H:\EPAL\EPAL-project\.venv`, which already has PyMuPDF, pandas and openpyxl
from the TUA setup; anything else on the PATH with those packages also works).

## Day-to-day use (non-technical)

1. Drop LURH PDFs into `data\inbox` (folders are created on first run).
2. Double-click **run_lurh.bat**.
3. When it finishes:
   * `data\processed\` — extracted cleanly (verdict **OK**);
   * `data\review\`   — needs a human look; the `*_REVIEW.txt` next to each
     PDF says exactly why, in plain language;
   * `data\extracted\` — the per-file `<name>_extracted.xlsx` (human view:
     Resumo, Dados da Licença — every extracted field in the document's own
     section order —, Condições de Descarga, Autocontrolo, Meio Recetor,
     Condições (texto), Referências, with the section notes — regime
     Observações, reporting periodicity, equipment, monitoring points —
     shown under each table) and `<name>_extracted.json` (raw data);
   * `master_lurh.xlsx` — the consolidated workbook (see *Master* below);
   * `powerbi_csv\` — one CSV per master sheet, the stable Power BI source;
   * `qa_report.xlsx` — one row per processed file with verdict and reasons.

Other ways to run:

* **watch_lurh.bat** — leave a window open; any PDF dropped into the inbox is
  processed automatically (safe on OneDrive/SharePoint-synced folders: a file
  is only picked up once its size stops changing).
* **register_watch_task.bat** — run once to install a hidden Windows scheduled
  task ("EPAL LURH Watch") that checks the inbox every 5 minutes, survives
  reboots, and shows no window. **unregister_watch_task.bat** removes it.
* Single file, no pipeline:
  `python extract_lurh.py path\to\licenca.pdf` (writes outputs next to the
  script and upserts the master).

If `master_lurh.xlsx` is open in Excel while the pipeline runs, the update is
diverted to `master_lurh.PENDING.xlsx` and a warning is logged — close the
master and re-run to consolidate. A `master_lurh.bak.xlsx` backup is written
before every successful update.

## Master workbook / Power BI model

All sheets are linked on **Nº Licença** (e.g. `L006788.2022.RH5A`);
re-extracting a licence *replaces* its rows everywhere (idempotent upsert):

| Sheet          | One row per…                     | Notes                                          |
| -------------- | -------------------------------- | ---------------------------------------------- |
| `Licenses`     | licence                          | 44 columns; includes `TUA/LURH = 'LURH'` so it can be unioned with the TUA master |
| `Conditions`   | discharge parameter (VLE)        | `Regime` = `Normal` / `Ano de arranque`; Saída autocontrolo frequency+type merged in |
| `Autocontrolo` | self-monitoring row              | Entrada/Saída × parameter                      |
| `MeioRecetor`  | receiving-water monitoring row   | P1/P2 (montante/jusante) — LURH-only sheet     |
| `Legislacao`   | legend letter (a), (b), …        | full legal text per code                       |
| `Avaliacao`    | conformity-assessment entry      | lettered, or one unlettered row applying to all parameters |

## QA verdicts

`qa.py` grades every extraction:

* **FAILED** — no discharge conditions, no Nº Licença, or the Nº Licença
  contradicts the filename. The master is *not* touched.
* **REVIEW** — extracted, master updated, but a human should confirm:
  dates that do not match the filename, missing autocontrolo, coordinates
  outside mainland Portugal, odd-looking VLEs, or an extractor warning.
* **OK** — everything cross-checked.

The strongest check is free: the SILIAMB filename encodes the ground truth
(`…_L005347_2021_RH5A_12_07_2021_a_11_07_2026.pdf` → licence + both dates),
so the extraction is verified against it with no external data. The parser
also accepts the corpus quirks: a fused `RH5A02_09_2019` and compact
`27032014` dates.

Two REVIEW patterns are *expected* and correct:

1. **Document ≠ filename dates** (e.g. *ETAR Águas*): the certificate really
   does state different dates than its filename. The extractor keeps the
   document's values and QA flags the discrepancy for a human.
2. **4-column "% remoção" tables** (e.g. *Castanheira de Pêra*, "Ano de
   arranque"): when a row carries a single value under a header that has both
   *VLE (% mínima de remoção)* and *VLE*, the DL 152/97 removal triple
   (70-90 / 75 / 90) is placed in the % column and a warning asks for a
   one-time visual confirmation against the PDF.

## How extraction works (maintainers)

Same philosophy as the TUA extractor — **never x/y positions of words**:

* **Labels, not positions** — values are found by their label text
  ("Código APA", "Massa de água", …) matched accent/case/`*`-insensitively at
  the start of a line; the longest label wins ("Denominação do meio recetor"
  beats "Meio Recetor"). Wrapped values are absorbed by continuation lines;
  known label tails ("(e.p)", "(superficial) ou estado (subterrânea)…") are
  stripped from values.
* **Markers, not page numbers** — sections are sliced between marker lines
  ("Caracterização da rejeição", "Condições de descarga das águas
  residuais…", "Autocontrolo", …), which is what keeps the *ano de arranque*
  VLE table apart from the *condições normais* one.
* **Both text layouts** — SILIAMB PDFs come in two shapes: table cells
  joined per visual row, and (in the real exports) every cell on its own
  line. All parsers accept both, chosen automatically per table.
* **Two readers per table, best result wins** — on real PDFs
  `fitz.find_tables()` reads the ruled tables with columns mapped **by header
  name**; independently, a text-layer state machine parses the same tables
  using their small closed vocabularies (Entrada/Saída, Mensal/Quinzenal/…,
  Pontual/Composta, the `Metodologia…`/`Anexo n do Decreto-Lei` method
  anchors). Whichever yields more rows is kept, so a single detection failure
  can never blank a section.
* **Page-break repair** — the text reader recognises the SILIAMB renderer's
  page-split artefacts: repeated/wrapped table headers, rows whose columns are
  glued onto one line, and parameter/method cell remainders that appear after
  the frequency; it reassembles them (and flags anything still garbled).
* Letter legends `(a)…(e)` under *Legislação* / *Avaliação da conformidade*
  are resolved to full text per parameter; unlettered Avaliação paragraphs
  apply to every row.

`test_extraction.py` runs instantly against the 11 frozen
`*_extracted.json` golden files shipped in this folder
(`python test_extraction.py` or `pytest`).

## Configuration (environment variables, all optional)

| Variable                  | Default                     | Purpose                          |
| ------------------------- | --------------------------- | -------------------------------- |
| `EPAL_LURH_DATA`          | `.\data`                    | root of inbox/processed/review/… |
| `EPAL_LURH_MASTER`        | `.\master_lurh.xlsx`        | master workbook path             |
| `EPAL_LURH_OUTPUT_DIR`    | script folder               | single-file mode output dir      |
| `EPAL_LURH_PDF_FILE`      | –                           | single-file mode default input   |
| `EPAL_LURH_WATCH_SECONDS` | `5`                         | watch-mode poll interval         |
| `EPAL_LURH_OPEN_EXCEL`    | `1`                         | `0` = don't auto-open results    |

Optional enrichment: if an `enrich_lurh.py` with `enrich_all(sheets)` exists
next to the scripts, `export_master_csv` applies it before writing the CSVs
(mirrors the TUA `enrich` hook).
