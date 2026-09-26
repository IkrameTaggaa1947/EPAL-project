r"""EPAL TUA batch pipeline — the entry point for non-technical operators.

Drop TUA PDFs into the 'inbox' folder, then run this file (or double-click
PROCESSAR_AGORA.bat). For each PDF it:
  1. extracts the data,
  2. runs quality checks (verdict: OK / FAILED; other issues become notes),
  3. updates the multi-sheet master (backed up first) unless extraction FAILED,
  4. exports the master to CSVs for Power BI,
  5. files the PDF into processed/ (OK) or review/ (FAILED) with a note,
  6. logs everything and appends to a plain-language qa_report.xlsx.

Folders (created automatically under ./data; override the root with EPAL_TUA_DATA):
  ../../Drop_New_Licenses/new_TUA_licences <- DROP PDFs HERE (override: EPAL_TUA_INBOX)
  ../../data/pdfs/NEW/TUA -> extracted cleanly (the archive; no second copy)
  data/review     -> a human should check these (a .txt says why)
  ../extraction_JSON_TUA -> per-file JSON + Excel (outside the code folder)
  %LOCALAPPDATA%\EPAL\logs -> run logs (local to this PC, never synced)
  powerbi_csv/    -> master as CSV (the Power BI source)
  master_tua.xlsx -> consolidated workbook   qa_report.xlsx -> verdict log
"""
import os
import sys
import glob
import time
import shutil
import logging
from datetime import datetime

import pandas as pd

# --- Python portatil: garantir que esta pasta esta no sys.path ---------------
# O runtime\python (Python embutivel) corre em modo ISOLADO por causa do
# ficheiro ._pth e, nesse modo, o Python NAO acrescenta a pasta do script ao
# sys.path como faz o Python normal. Sem isto, importar um modulo vizinho
# falha com ModuleNotFoundError. Com um Python instalado, e inofensivo.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import extract_tua as E
import qa

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.getenv('EPAL_TUA_DATA', os.path.join(HERE, 'data'))
# Single user-facing inbox at the project root: drop TUA PDFs in Drop_New_Licenses/new_TUA_licences
INBOX = os.getenv('EPAL_TUA_INBOX', os.path.join(os.path.dirname(os.path.dirname(HERE)), 'Drop_New_Licenses', 'new_TUA_licences'))
# Uma licenca extraida com exito vai para o ARQUIVO, nao para uma segunda
# arvore dentro de src/. Guardar copias em data/processed so fazia o
# projeto pesar mais: os dados ja estao no master, no JSON e no .xlsx,
# e o PDF continua a existir - uma vez - em data/pdfs.
PROCESSED = os.getenv('EPAL_TUA_PROCESSED',
                      os.path.join(os.path.dirname(os.path.dirname(HERE)),
                                   'data', 'pdfs', 'NEW', 'TUA'))
REVIEW = os.path.join(DATA, 'review')
# Per-file JSON/XLSX now live OUTSIDE the code folder, in src/extraction_JSON_TUA
EXTRACTED = os.getenv('EPAL_TUA_EXTRACTED', os.path.join(os.path.dirname(HERE), 'extraction_JSON_TUA'))
# O registo (.log) vai para %LOCALAPPDATA%\EPAL\logs, fora da pasta partilhada
# -- ver Common_Code_PowerBI/logdir.py. Isto aqui e outra coisa: as linhas
# antigas do qa_report, que sao dados DO PROJETO e ficam com o projeto.
QA_ARCHIVE = os.path.join(DATA, 'logs')
MASTER = os.getenv('EPAL_TUA_MASTER', os.path.join(HERE, 'master_tua.xlsx'))
SECTIONS_MASTER = os.getenv('EPAL_TUA_SECTIONS_MASTER', os.path.join(HERE, 'master_all_sections.xlsx'))
QA_REPORT = os.getenv('EPAL_TUA_QA_REPORT', os.path.join(HERE, 'qa_report.xlsx'))

log = logging.getLogger('tua')


def _setup():
    for d in (INBOX, PROCESSED, REVIEW, EXTRACTED):
        os.makedirs(d, exist_ok=True)
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fmt = logging.Formatter('%(asctime)s  %(levelname)-7s  %(message)s', '%H:%M:%S')
    sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'Common_Code_PowerBI'))
    import logdir
    logdir.add_file_handler(log, 'tua', fmt)
    # Console handler only when there IS a console — under pythonw.exe (hidden,
    # scheduled-task runs) sys.stderr is None, so the file log is the record.
    if sys.stderr is not None:
        ch = logging.StreamHandler()
        ch.setFormatter(fmt)
        log.addHandler(ch)
    _reclaim_pending()
    _sweep_review_notes()


def _sweep_review_notes():
    """Drop review notes whose PDF is no longer there.

    The note itself tells the operator to move the PDF to the archive once
    the data checks out - but nothing removed the note, so review/ filled up
    with warnings about files that are long gone (70 notes for 2 PDFs). The
    folder is meant to be a worklist, not a pile."""
    try:
        sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'Common_Code_PowerBI'))
        import pending
        n, msg = pending.sweep_orphan_notes(REVIEW)
    except Exception as exc:
        log.warning(f"Could not sweep review notes: {exc}")
        return
    if msg:
        log.info(f"  {msg}")


def _reclaim_pending():
    """Fold any leftover '<master>.PENDING.xlsx' back into its master.

    update_master() diverts to a PENDING workbook when the master is open in
    Excel, so nothing is lost — but nothing used to read those files back, so
    the data sat in a file the system had no memory of. This closes that loop.
    Cheap when there is nothing to do (one isfile check per master), and never
    aborts the run: a still-locked master just means we try again next time."""
    try:
        sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'Common_Code_PowerBI'))
        import pending
    except Exception as exc:                       # shared module missing — not fatal
        log.warning(f"Could not load the PENDING recovery helper: {exc}")
        return
    for master, key in ((MASTER, 'Nº TUA'), (SECTIONS_MASTER, 'Nº TUA')):
        try:
            n, msg = pending.reclaim(master, key, style=E._style_sheet)
        except Exception as exc:
            log.exception(f"PENDING recovery failed for {os.path.basename(master)}: {exc}")
            continue
        if msg:
            (log.info if n else log.warning)(f"  {msg}")


def _inbox_pdfs():
    """Every PDF in the inbox, whatever the case of its extension.

    glob('*.pdf') is case-INSENSITIVE on Windows but case-SENSITIVE on Linux,
    and this pipeline has been run on both (see docs/AMBIENTES.md). A file
    dropped in as LICENCA.PDF would be processed on one and silently ignored
    on the other - a skip with no message anywhere."""
    return sorted(p for p in glob.glob(os.path.join(INBOX, '*'))
                  if p.lower().endswith('.pdf') and os.path.isfile(p))


def _move(pdf, dest_dir, note=None):
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, os.path.basename(pdf))
    if os.path.abspath(dest) != os.path.abspath(pdf):
        shutil.move(pdf, dest)
    if note:
        with open(os.path.splitext(dest)[0] + '_REVIEW.txt', 'w', encoding='utf-8') as f:
            f.write(note)
    return dest


def _extract_one(pdf):
    """Extract + QA one PDF. Does NOT touch the master and does NOT move the PDF.

    Writing the master is deferred to _commit(), once per batch, because
    _upsert_sheets() rewrites the whole workbook every time it is called."""
    name = os.path.basename(pdf)
    log.info(f"processing {name}")
    item = {'pdf': pdf, 'data': None, 'file_name': name, 'added_to_master': False,
            'problem': '', 'timestamp': datetime.now().isoformat(timespec='seconds')}
    try:
        data = E.extract(pdf)
    except Exception as exc:                       # never let one bad PDF stop the batch
        log.exception(f"EXTRACTION CRASHED for {name}: {exc}")
        item.update(verdict='FAILED', conditions=0, autocontrolo=0,
                    problem=f"{name}: extraction crashed ({exc}) — NOT added to master_tua",
                    reasons=f'extraction crashed: {exc}')
        return item

    _, xlsx = E.write_outputs(data, EXTRACTED)
    verdict, reasons = qa.check(data)
    ncond, nauto = len(data['condicoes_rejeicao']), len(data['autocontrolo'])
    log.info(f"  -> {verdict}  ({ncond} conditions, {nauto} monitoring rows)")
    for r in reasons:
        log.info(f"     {r}")

    if verdict == qa.FAILED:
        item['problem'] = (f"{name}: NOT added to master_tua — extraction failed QA: "
                           + "; ".join(r.replace('[FAILED] ', '') for r in reasons if 'FAILED' in r))
    item.update(data=data, verdict=verdict, conditions=ncond, autocontrolo=nauto, xlsx=xlsx,
                reasons=' | '.join(reasons), _reasons=reasons)
    item['Nº TUA'] = data['dados_gerais'].get('Nº TUA', '')
    item['Estabelecimento'] = data['dados_gerais'].get('Estabelecimento', '')
    return item


def _commit(items):
    """One read+write of each master for the whole batch, then file every PDF.

    Only certificates that passed QA go into the master. If the master is
    locked, the batch lands in the .PENDING workbook and NO PDF is filed as
    processed — the same rule as before, applied to the batch."""
    keep = [it for it in items if it['data'] is not None and it['verdict'] != qa.FAILED]
    if keep:
        ds = [it['data'] for it in keep]
        mp, counts, deferred = E.update_master_many(ds, MASTER)
        if deferred:
            for it in keep:
                it['problem'] = (f"{it['file_name']}: NOT added to master_tua — the master "
                                 f"workbook is open/locked in Excel. Data saved to "
                                 f"'{os.path.basename(mp)}'; close master_tua.xlsx and re-run.")
            log.warning(f"  master locked — batch saved to {os.path.basename(mp)}. "
                        f"Close master_tua.xlsx and re-run these PDFs.")
        else:
            for it in keep:
                it['added_to_master'] = True
            E.export_master_csv(MASTER)
            log.info(f"  master updated ({len(keep)} certificate(s)): {counts}")
        # Consolidated all-sections workbook (one sheet per section, all certificates).
        _, scounts, sdef = E.update_sections_master_many(ds, SECTIONS_MASTER)
        log.info(f"  all-sections master {'locked (PENDING)' if sdef else 'updated'}: {len(scounts)} sheets")

    for it in items:
        _file_one(it)
    return [{k: v for k, v in it.items() if k not in ('data', 'pdf', '_reasons')}
            for it in items]


def _file_one(it):
    """Move one PDF to processed/ or review/ now that we know if it reached the master.

    Only PDFs whose data is really in the master are filed as processed — a
    deferred (locked-master) write means the data is in a .PENDING workbook,
    NOT in the master, so the PDF must stay visible in review/."""
    name, reasons = it['file_name'], it.get('_reasons', [])
    if it['data'] is None:                          # extraction crashed
        _move(it['pdf'], REVIEW,
              f"Extraction crashed: {it['reasons']}\nSend this PDF to the technical team.")
        return
    if it['added_to_master'] and it['verdict'] == qa.OK:
        _move(it['pdf'], PROCESSED)
        return
    head = it['problem'] or f"{name}: added to master, but flagged {it['verdict']} — verify against the PDF."
    note = (f"{name}\nVerdict: {it['verdict']}\n"
            f"Added to master: {'yes' if it['added_to_master'] else 'NO'}\n\n"
            f"Problem:\n  {head}\n\nQA details:\n  - " + "\n  - ".join(reasons or ['(none)']) +
            "\n\nWhat to do: open the matching file in ../extraction_JSON_TUA/ and check the flagged "
            "fields against the PDF. If the data is correct, move this PDF to data/pdfs/NEW/TUA/.")
    _move(it['pdf'], REVIEW, note)


def process_one(pdf):
    """Extract, QA-check, update the master, and file one PDF. Returns a report dict.

    Kept for single-file use (and the .bat that passes one path); a batch goes
    through _process_batch(), which writes the master once for the whole set."""
    return _commit([_extract_one(pdf)])[0]


def _open_excel(rows):
    """After a run, open the result in Excel for the operator. Opens the single
    file's own workbook (leaving the master free to keep updating); for a bulk run
    it opens the consolidated master once. Skipped when headless (hidden scheduled
    task: sys.stderr is None) or when EPAL_TUA_OPEN_EXCEL=0."""
    if os.getenv('EPAL_TUA_OPEN_EXCEL', '1') == '0' or sys.stderr is None:
        return
    done = [r['xlsx'] for r in rows if r.get('xlsx') and os.path.isfile(r['xlsx'])]
    target = done[0] if len(done) == 1 else (MASTER if os.path.isfile(MASTER) else None)
    if not target:
        return
    try:
        os.startfile(target)                       # Windows: open in the default app (Excel)
        log.info(f"Opened in Excel: {os.path.basename(target)}")
    except Exception as e:
        log.info(f"Could not open Excel automatically: {e}")


def _append_report(rows):
    """Append this batch to qa_report.xlsx via the shared writer.

    The old version read the existing report inside an
    `except Exception: pass`, so an unreadable file silently threw away the
    whole history and left only this run. It also never rotated. Both are
    handled in Common_Code_PowerBI/qa_report.py, which the two pipelines
    share so the behaviour cannot drift apart again."""
    try:
        sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'Common_Code_PowerBI'))
        import qa_report
    except Exception as exc:
        log.error(f"Could not load the QA-report writer: {exc}")
        return
    qa_report.append(rows, QA_REPORT, archive_dir=QA_ARCHIVE, log=log)


def _rebuild_dashboard_feed(rows):
    """Rebuild the combined Power BI feed (data/powerbi/*.xlsx) after a batch.

    export_master_csv() (called per file) refreshes only data/powerbi/TUA/*.csv;
    the dashboard actually reads the MERGED root .xlsx built by make_powerbi_all.
    So the last automatic link is to run that merge once per batch — but only when
    at least one certificate was ingested (not all FAILED), to avoid pointless
    rebuilds. Errors here never abort the batch: the masters/CSVs are already
    saved, so a merge hiccup just means the dashboard shows the previous state.
    """
    if not any(r.get('verdict') != qa.FAILED for r in rows):
        return
    if os.getenv('EPAL_SKIP_DASHBOARD_FEED') == '1':
        log.info("Dashboard feed rebuild skipped (EPAL_SKIP_DASHBOARD_FEED=1).")
        return
    try:
        sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'Common_Code_PowerBI'))  # shared code lives here
        import make_powerbi_all
        import importlib
        importlib.reload(make_powerbi_all)             # pick up fresh CSVs each run
        make_powerbi_all.main()
        log.info("Dashboard feed rebuilt: data/powerbi/*.xlsx (TUA + LURH merged).")
    except Exception as exc:
        log.exception(f"Dashboard feed rebuild FAILED (masters/CSVs are safe): {exc}")
        return                                         # extras derive from the base tables

    # Derived tables (criteria bridge, parameter sheet, priority conditions,
    # outputs/renewal). They read the base .xlsx just rebuilt above, so they
    # must run after it — and never abort the batch if they fail.
    try:
        import make_powerbi_extras
        import importlib
        importlib.reload(make_powerbi_extras)
        make_powerbi_extras.main()
        log.info("Derived tables rebuilt: Licencas_Criterios, Params_Ficha, "
                 "Condicoes_Prioritarias, Renovacao_Proposta.")
    except Exception as exc:
        log.exception(f"Derived tables rebuild FAILED (base tables are safe): {exc}")


def _process_batch(pdfs):
    """Process a list of PDFs, write the report, log a tally. Returns the tally.

    Extracts a chunk, then writes each master ONCE for that chunk. The chunk
    exists so a long backlog stays durable: a crash loses at most the chunk in
    flight, not the whole run. EPAL_TUA_BATCH=1 restores the old per-file
    behaviour."""
    _reclaim_pending()          # a locked master from the previous batch, folded back in
    size = max(1, int(os.getenv('EPAL_TUA_BATCH', '10')))
    rows = []
    for start in range(0, len(pdfs), size):
        chunk = pdfs[start:start + size]
        if len(pdfs) > size:
            log.info(f"--- batch {start // size + 1} of {-(-len(pdfs) // size)} "
                     f"({len(chunk)} certificate(s)) ---")
        rows.extend(_commit([_extract_one(p) for p in chunk]))
    _append_report(rows)
    tally = {}
    for r in rows:
        tally[r['verdict']] = tally.get(r['verdict'], 0) + 1
    log.info("-" * 52)
    log.info("DONE  " + "  ".join(f"{v}:{n}" for v, n in sorted(tally.items())))
    if tally.get('REVIEW') or tally.get('FAILED'):
        log.info(f"Some certificates need a look — see: {REVIEW}")
    _rebuild_dashboard_feed(rows)
    _open_excel(rows)
    return tally


RUN_LOCK = os.path.join(os.path.dirname(os.path.dirname(HERE)),
                        'data', '_staging', 'run.lock')


def _run_lock():
    """The shared lock that stops two people processing at the same time.

    Both would read the master, both would rewrite it whole, and whoever
    saved last would erase the other's licences - with no error anywhere.
    Only the automatic watcher used to be protected; PROCESSAR_AGORA.bat, which
    exists to be double-clicked by operators, was not.
    """
    sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'Common_Code_PowerBI'))
    import runlock
    return runlock.held(RUN_LOCK, 'TUA')


def main():
    _setup()
    with _run_lock() as lock:
        if not lock.ok:
            log.warning(f'  {lock.msg}')
            return 2
        if lock.msg:
            log.warning(f'  {lock.msg}')
        return _main_locked()


def _main_locked():
    pdfs = _inbox_pdfs()
    if not pdfs:
        log.info(f"Inbox is empty. Drop TUA PDFs into: {INBOX}")
        return 0
    log.info(f"Found {len(pdfs)} PDF(s) in the inbox.")
    _process_batch(pdfs)
    log.info(f"Report: {QA_REPORT}")
    return 0


def _scan_ready(seen):
    """One inbox scan. A PDF is 'ready' only when its size is unchanged since the
    previous scan — so a file still being copied/synced is left until it settles.
    Returns (ready_paths, current_sizes)."""
    ready, now = [], {}
    for p in _inbox_pdfs():
        try:
            now[p] = os.path.getsize(p)
        except OSError:
            continue                       # vanished mid-scan; skip
        if seen.get(p) == now[p]:          # stable since last scan → fully written
            ready.append(p)
    return ready, now


def watch(interval=None):
    """Poll the inbox forever and process PDFs as they arrive. Polling (not OS file
    events) is deliberate: it is reliable on OneDrive/SharePoint-synced folders and
    gives the 'file has finished copying' check for free. Ctrl+C to stop."""
    interval = interval or int(os.getenv('EPAL_TUA_WATCH_SECONDS', '5'))
    _setup()
    log.info(f"Watching {INBOX}  (every {interval}s). Drop TUA PDFs here. Press Ctrl+C to stop.")
    seen = {}
    try:
        while True:
            ready, seen = _scan_ready(seen)
            if ready:
                log.info(f"{len(ready)} new PDF(s) detected.")
                with _run_lock() as lock:
                    if lock.ok:
                        _process_batch(ready)
                    else:
                        log.warning(f'  {lock.msg}')
                        continue          # tenta outra vez no proximo ciclo
            time.sleep(interval)
    except KeyboardInterrupt:
        log.info("Watch stopped.")
    return 0


if __name__ == '__main__':
    sys.exit(watch() if '--watch' in sys.argv[1:] else main())
