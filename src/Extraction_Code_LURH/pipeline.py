r"""EPAL LURH batch pipeline — the entry point for non-technical operators.

Drop LURH PDFs into the 'inbox' folder, then run this file (or double-click
PROCESSAR_AGORA.bat). For each PDF it:
  1. extracts the data,
  2. checks whether that licence (Nº Licença) is already sitting in the master —
     a re-drop of a PDF that was processed before. If so, it DELETES the PDF
     from the inbox right there and skips the rest (nothing re-added, nothing
     re-archived): see "Duplicate PDFs" below.
  3. otherwise runs quality checks (verdict: OK / FAILED — see "Nothing blocks
     the master" below for what FAILED actually means now),
  4. updates the multi-sheet master (backed up first) — for EVERY PDF that
     extracted at all, regardless of verdict,
  5. exports the master to CSVs for Power BI,
  6. files the PDF into processed/ (OK) or review/ (FAILED) with a note —
     filing only, it is in the master and the Power BI feed either way,
  7. logs everything and appends to a plain-language qa_report.xlsx,
  8. rebuilds the merged Power BI feed (data/powerbi/*.xlsx) once per batch.

Nothing blocks the master (step 4): a QA verdict of FAILED — missing Nº
Licença, no discharge conditions, whatever the reason — used to keep a
licence out of master_lurh.xlsx entirely, which meant it was simply invisible
in Power BI until a human noticed and fixed the source PDF. That's now gone:
verdict no longer gates master inclusion at all. Every PDF that extracts
without CRASHING (E.extract() didn't raise) gets its data upserted, complete
or not — a licence missing its number still goes in keyed by file name (see
extract_lurh.lic_key), a licence with an empty VLE table still goes in with
that table empty. The verdict still decides where the PDF is FILED
(processed/ vs review/) and still shows up in qa_report.xlsx and the log, so
a human can go complete or fix what's missing — it just no longer decides
whether the data reaches Power BI. Only a genuine extraction CRASH (the PDF
couldn't be parsed into anything at all) keeps a PDF out of the master; there
is nothing left to upsert in that case.

Duplicate PDFs (step 2): the master already upserts by Nº Licença (see
update_master_many's docstring — re-extracting a licence replaces its rows,
never duplicates them), so a re-drop was always harmless to the DATA. What
was missing is the INBOX side: a re-dropped PDF used to be extracted again
(wasted work) and then moved into the archive over a same-named file already
there — the .PDF equivalent of an unnecessary duplicate. Now, once extraction
names the licence, that name is checked against the master BEFORE quality
checks or the per-file JSON/XLSX are written; a match means "already
extracted", and the PDF is deleted from the inbox instead. Set
EPAL_LURH_SKIP_DUPLICATE_CHECK=1 to turn this off for one run — e.g. right
after fixing a bug in extract_lurh.py, to force a licence already in the
master to be re-extracted instead of deleted.

Folders (created automatically under ./data; override the root with EPAL_LURH_DATA):
  ../../Drop_New_Licenses/new_LURH_licences <- DROP PDFs HERE (override: EPAL_LURH_INBOX)
  ../../data/pdfs/NEW/LURH -> extracted cleanly (the archive; no second copy)
  data/review     -> a human should check these (a .txt says why)
  ../extraction_JSON_LURH -> per-file JSON + Excel (outside the code folder)
  %LOCALAPPDATA%\EPAL\logs -> run logs (local to this PC, never synced)
  powerbi_csv/    -> master as CSV (the Power BI source)
  master_lurh.xlsx -> consolidated workbook   qa_report.xlsx -> verdict log
"""
import os
import sys
import glob
import json
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

import extract_lurh as E
import qa

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.getenv('EPAL_LURH_DATA', os.path.join(HERE, 'data'))
# Single user-facing inbox at the project root: drop LURH PDFs in Drop_New_Licenses/new_LURH_licences
INBOX = os.getenv('EPAL_LURH_INBOX', os.path.join(os.path.dirname(os.path.dirname(HERE)), 'Drop_New_Licenses', 'new_LURH_licences'))
# Uma licenca extraida com exito vai para o ARQUIVO, nao para uma segunda
# arvore dentro de src/. Guardar copias em data/processed so fazia o
# projeto pesar mais: os dados ja estao no master, no JSON e no .xlsx,
# e o PDF continua a existir - uma vez - em data/pdfs.
PROCESSED = os.getenv('EPAL_LURH_PROCESSED',
                      os.path.join(os.path.dirname(os.path.dirname(HERE)),
                                   'data', 'pdfs', 'NEW', 'LURH'))
REVIEW = os.path.join(DATA, 'review')
# Per-file JSON/XLSX now live OUTSIDE the code folder, in src/extraction_JSON_LURH
EXTRACTED = os.getenv('EPAL_LURH_EXTRACTED', os.path.join(os.path.dirname(HERE), 'extraction_JSON_LURH'))
# O registo (.log) vai para %LOCALAPPDATA%\EPAL\logs, fora da pasta partilhada
# -- ver Common_Code_PowerBI/logdir.py. Isto aqui e outra coisa: as linhas
# antigas do qa_report, que sao dados DO PROJETO e ficam com o projeto.
QA_ARCHIVE = os.path.join(DATA, 'logs')
MASTER = os.getenv('EPAL_LURH_MASTER', os.path.join(HERE, 'master_lurh.xlsx'))
QA_REPORT = os.getenv('EPAL_LURH_QA_REPORT', os.path.join(HERE, 'qa_report.xlsx'))

log = logging.getLogger('lurh')


def _setup():
    for d in (INBOX, PROCESSED, REVIEW, EXTRACTED):
        os.makedirs(d, exist_ok=True)
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fmt = logging.Formatter('%(asctime)s  %(levelname)-7s  %(message)s', '%H:%M:%S')
    sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'Common_Code_PowerBI'))
    import logdir
    logdir.add_file_handler(log, 'lurh', fmt)
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
    """Fold any leftover 'master_lurh.PENDING.xlsx' back into the master.

    update_master() diverts to a PENDING workbook when the master is open in
    Excel, so nothing is lost — but nothing used to read those files back, so
    the data sat in a file the system had no memory of. This closes that loop.
    Cheap when there is nothing to do (one isfile check), and never aborts the
    run: a still-locked master just means we try again next time."""
    try:
        sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'Common_Code_PowerBI'))
        import pending
    except Exception as exc:                       # shared module missing — not fatal
        log.warning(f"Could not load the PENDING recovery helper: {exc}")
        return
    try:
        n, msg = pending.reclaim(MASTER, 'Nº Licença')
    except Exception as exc:
        log.exception(f"PENDING recovery failed for {os.path.basename(MASTER)}: {exc}")
        return
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


def _master_licence_numbers():
    """Nº Licença values already sitting in the master, read once per batch.

    Only the 'Licenses' sheet's key column is loaded — cheap even for a large
    master. Any read problem (missing file, locked file, unexpected schema)
    fails OPEN: an empty set means the duplicate check simply finds nothing
    this run, so a real new PDF is never mistakenly deleted."""
    if not os.path.isfile(MASTER):
        return set()
    try:
        df = pd.read_excel(MASTER, sheet_name='Licenses', usecols=['Nº Licença'])
        return set(df['Nº Licença'].dropna().astype(str))
    except Exception as exc:
        log.warning(f"Could not read existing licence numbers from the master "
                    f"({exc}) — duplicate check skipped this batch.")
        return set()


def _extract_one(pdf, known=None):
    """Extract + QA one PDF. Does NOT touch the master and does NOT move the PDF.

    Writing the master is deferred to _commit(), once per batch, because
    update_master_many() rewrites the whole workbook every time it is called.

    `known` is the set of Nº Licença already in the master before this batch
    (see _master_licence_numbers). As soon as extraction names the licence,
    that name is checked against it — a match means this PDF is a re-drop of
    one already extracted, and QA / per-file outputs are skipped entirely in
    favour of marking it a duplicate (handled by _file_one: deleted, not
    archived). Pass known=None (the default) to disable the check, e.g. for
    process_one()'s legacy single-file path when the caller doesn't need it."""
    name = os.path.basename(pdf)
    log.info(f"processing {name}")
    item = {'pdf': pdf, 'data': None, 'file_name': name, 'added_to_master': False,
            'duplicate': False, 'problem': '', 'timestamp': datetime.now().isoformat(timespec='seconds')}
    try:
        data = E.extract(pdf)
    except Exception as exc:                       # never let one bad PDF stop the batch
        log.exception(f"EXTRACTION CRASHED for {name}: {exc}")
        item.update(verdict='FAILED', conditions=0, autocontrolo=0,
                    problem=f"{name}: extraction crashed ({exc}) — NOT added to master_lurh",
                    reasons=f'extraction crashed: {exc}')
        return item

    lic = data['dados_gerais'].get('Nº Licença', '')
    item['Nº Licença'] = lic
    item['Estabelecimento'] = data['tratamento'].get('Designação', '')

    if (known is not None and lic and str(lic) in known
            and os.getenv('EPAL_LURH_SKIP_DUPLICATE_CHECK') != '1'):
        log.info(f"  -> DUPLICATE  (licence {lic} already in master_lurh — deleting, no QA run)")
        item.update(data=data, duplicate=True, verdict='DUPLICATE', conditions=0, autocontrolo=0,
                    reasons='', _reasons=[],
                    problem=f"{name}: licence {lic} is already in master_lurh — duplicate drop, PDF deleted.")
        return item

    _, xlsx = E.write_outputs(data, EXTRACTED)
    verdict, reasons = qa.check(data)
    ncond = len(data['condicoes_descarga'])
    nauto = len(data['autocontrolo'].get('rows', []))
    log.info(f"  -> {verdict}  ({ncond} conditions, {nauto} autocontrolo rows)")
    for r in reasons:
        log.info(f"     {r}")

    if verdict == qa.FAILED:
        # Still added to the master (see module docstring, "Nothing blocks the
        # master") — FAILED here means "flag this PDF for a human", not "keep
        # this data out of Power BI".
        item['problem'] = (f"{name}: added to master_lurh despite failed QA — please verify: "
                           + "; ".join(r.replace('[FAILED] ', '') for r in reasons if 'FAILED' in r))
    item.update(data=data, verdict=verdict, conditions=ncond, autocontrolo=nauto, xlsx=xlsx,
                reasons=' | '.join(reasons), _reasons=reasons)
    return item


def _commit(items):
    """One read+write of the master for the whole batch, then file every PDF.

    Every PDF that actually extracted (data is not None) is upserted into the
    master, REGARDLESS of QA verdict — a licence missing its number still goes
    in keyed by file name, one with an empty VLE table still goes in with that
    table empty. Only a genuine extraction crash (data is None: nothing to
    upsert) or a duplicate (already there, see _mark via `known` in
    _extract_one) is excluded. The report row still carries 'added_to_master'
    plus a one-line 'problem' when something needs a human look — that's what
    routes the PDF to data/review/ instead of processed/ (see _file_one) and
    what the Power Automate flow displays; it no longer means "not in Power BI"."""
    keep = [it for it in items if it['data'] is not None and not it.get('duplicate')]
    if keep:
        mp, counts, deferred = E.update_master_many([it['data'] for it in keep], MASTER)
        if deferred:
            for it in keep:
                it['problem'] = (f"{it['file_name']}: NOT added to master_lurh — the master "
                                 f"workbook is open/locked in Excel. Data saved to "
                                 f"'{os.path.basename(mp)}'; close master_lurh.xlsx and re-run.")
            log.warning(f"  master locked — batch saved to {os.path.basename(mp)}. "
                        f"Close master_lurh.xlsx and re-run these PDFs.")
        else:
            for it in keep:
                it['added_to_master'] = True
            E.export_master_csv(MASTER)
            log.info(f"  master updated ({len(keep)} licence(s)): {counts}")

    for it in items:
        _file_one(it)
    return [{k: v for k, v in it.items() if k not in ('data', 'pdf', '_reasons')}
            for it in items]


def _file_one(it):
    """Move one PDF to processed/ or review/ now that we know if it reached the master.

    A duplicate is neither: its data was already committed the first time this
    licence was extracted, so the PDF is deleted outright rather than filed."""
    name, reasons = it['file_name'], it.get('_reasons', [])
    if it.get('duplicate'):
        try:
            os.remove(it['pdf'])
            log.info(f"  deleted {name} (duplicate of licence {it.get('Nº Licença')}, "
                     f"already in master_lurh).")
        except OSError as exc:
            log.warning(f"  could not delete duplicate {name}: {exc} — moving to review instead.")
            _move(it['pdf'], REVIEW,
                  f"{name}\nThis is a duplicate of licence {it.get('Nº Licença')}, already in "
                  f"master_lurh. The pipeline tried to delete it automatically but could not "
                  f"({exc}).\n\nWhat to do: it's safe to delete this PDF by hand — its data is "
                  f"already in the master.")
        return
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
            "\n\nWhat to do: open the matching file in ../extraction_JSON_LURH/ and check the flagged "
            "fields against the PDF. If the data is correct, move this PDF to data/pdfs/NEW/LURH/.")
    _move(it['pdf'], REVIEW, note)


def process_one(pdf):
    """Extract, QA-check, update the master, and file one PDF. Returns a report dict.

    Kept for single-file use; a batch goes through _process_batch(), which writes
    the master once for the whole set."""
    return _commit([_extract_one(pdf, known=_master_licence_numbers())])[0]


def _open_excel(rows):
    """After a run, open the result in Excel for the operator. Opens the single
    file's own workbook (leaving the master free to keep updating); for a bulk run
    it opens the consolidated master once. Skipped when headless (hidden scheduled
    task: sys.stderr is None) or when EPAL_LURH_OPEN_EXCEL=0."""
    if os.getenv('EPAL_LURH_OPEN_EXCEL', '1') == '0' or sys.stderr is None:
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


def _write_alerts(rows):
    """After every batch, write the outcome where Power Automate can pick it up:
      data/last_run_alerts.json — full structure (parse with 'Parse JSON')
      data/last_run_alerts.txt  — one line per problem file, ready to display
    Both files are OVERWRITTEN each run, so they always describe the last run.
    Also prints each problem line as 'ALERT: ...' to stdout for flows that read
    the script output directly.

    'problems' (added_to_master is False) is now rare — only a genuine
    extraction crash or a locked master leaves data out entirely. A FAILED QA
    verdict no longer means that: the licence IS in the master, just flagged
    — those are counted separately as 'needs_review' so that distinction
    (nothing lost vs. something worth a human look) survives into the alert,
    instead of collapsing into one 'problem' bucket."""
    problems = [r for r in rows if not r.get('added_to_master') and not r.get('duplicate')]
    needs_review = [r for r in rows
                     if r.get('added_to_master') and r.get('verdict') != 'OK' and not r.get('duplicate')]
    payload = {
        'run_timestamp': datetime.now().isoformat(timespec='seconds'),
        'total_files': len(rows),
        'added_to_master': sum(1 for r in rows if r.get('added_to_master')),
        'duplicates_deleted': sum(1 for r in rows if r.get('duplicate')),
        'not_added': len(problems),
        'needs_review': len(needs_review),
        'alerts': [{'file_name': r['file_name'],
                    'verdict': r.get('verdict', ''),
                    'problem': r.get('problem', ''),
                    'qa_reasons': r.get('reasons', '')} for r in problems],
        'review': [{'file_name': r['file_name'],
                    'licenca': r.get('Nº Licença', ''),
                    'verdict': r.get('verdict', ''),
                    'problem': r.get('problem', '')} for r in needs_review],
        'files': [{'file_name': r['file_name'],
                   'licenca': r.get('Nº Licença', ''),
                   'verdict': r.get('verdict', ''),
                   'added_to_master': bool(r.get('added_to_master')),
                   'problem': r.get('problem', '')} for r in rows],
    }
    try:
        with open(os.path.join(DATA, 'last_run_alerts.json'), 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        with open(os.path.join(DATA, 'last_run_alerts.txt'), 'w', encoding='utf-8') as f:
            if problems:
                f.write('\n'.join(r['problem'] for r in problems) + '\n')
            else:
                added, dupes = payload['added_to_master'], payload['duplicates_deleted']
                extra = f" ({dupes} duplicate(s) deleted)" if dupes else ""
                review_note = (f" — {len(needs_review)} of those need a human look (see data/review/), "
                                f"but all reached master_lurh" if needs_review else "")
                f.write(f"All {len(rows)} file(s) handled — {added} added to master_lurh{extra}{review_note}.\n")
    except Exception as e:
        log.warning(f"could not write alert files: {e}")
    for r in problems:
        log.error(f"ALERT: {r['problem']}")
        print(f"ALERT: {r['problem']}", flush=True)


def _rebuild_dashboard_feed(rows):
    """Rebuild the merged Power BI feed (data/powerbi/*.xlsx) after a LURH batch.

    export_master_csv() (per file) refreshes only src/Extraction_Code_LURH/powerbi_csv/*.csv; the
    dashboard actually reads the MERGED root data/powerbi/*.xlsx built by
    make_powerbi_all (which rebuilds the LURH side from the per-licence JSONs) plus
    the derived criteria tables from make_powerbi_extras. So — mirroring the TUA
    pipeline — run that merge once per batch, but only when at least one licence
    actually reached the master. Skippable with EPAL_SKIP_DASHBOARD_FEED=1; errors
    are logged, never raised (the master/CSVs are already saved, so a merge hiccup
    just leaves the dashboard showing the previous state).
    """
    if not any(r.get('added_to_master') for r in rows):
        return
    if os.getenv('EPAL_SKIP_DASHBOARD_FEED') == '1':
        log.info("Dashboard feed rebuild skipped (EPAL_SKIP_DASHBOARD_FEED=1).")
        return
    try:
        sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'Common_Code_PowerBI'))  # shared code lives here
        import make_powerbi_all
        import importlib
        importlib.reload(make_powerbi_all)             # pick up fresh JSONs/CSVs each run
        make_powerbi_all.main()
        log.info("Dashboard feed rebuilt: data/powerbi/*.xlsx (TUA + LURH merged).")
    except Exception as exc:
        log.exception(f"Dashboard feed rebuild FAILED (master/CSVs are safe): {exc}")
        return                                         # extras derive from the base tables
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
    """Process a list of PDFs, write the report, log a tally. Returns the tally."""
    _reclaim_pending()          # a locked master from the previous batch, folded back in
    known = _master_licence_numbers()      # licences already in the master before this batch
    size = max(1, int(os.getenv('EPAL_LURH_BATCH', '10')))
    rows = []
    for start in range(0, len(pdfs), size):
        chunk = pdfs[start:start + size]
        if len(pdfs) > size:
            log.info(f"--- batch {start // size + 1} of {-(-len(pdfs) // size)} "
                     f"({len(chunk)} licence(s)) ---")
        extracted = [_extract_one(p, known=known) for p in chunk]
        for it in extracted:                # a genuinely new licence also counts as
            if it['data'] is not None and not it.get('duplicate'):     # "known" from here
                lic = str(it.get('Nº Licença') or '').strip()          # on, so two copies
                if lic:                                                # of the same PDF in
                    known.add(lic)                                     # one drop both catch
        rows.extend(_commit(extracted))
    _append_report(rows)
    _write_alerts(rows)
    tally = {}
    for r in rows:
        tally[r['verdict']] = tally.get(r['verdict'], 0) + 1
    log.info("-" * 52)
    log.info("DONE  " + "  ".join(f"{v}:{n}" for v, n in sorted(tally.items())))
    if tally.get('FAILED'):
        log.info(f"{tally['FAILED']} licence(s) added to master_lurh despite a failed QA check "
                 f"— flagged for a human to verify, see: {REVIEW}")
    if tally.get('DUPLICATE'):
        log.info(f"{tally['DUPLICATE']} duplicate PDF(s) deleted (licence already in master_lurh).")
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
    return runlock.held(RUN_LOCK, 'LURH')


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
        log.info(f"Inbox is empty. Drop LURH PDFs into: {INBOX}")
        return 0
    log.info(f"Found {len(pdfs)} PDF(s) in the inbox.")
    tally = _process_batch(pdfs)
    log.info(f"Report: {QA_REPORT}")
    # non-zero exit when anything was kept out of the master -> Power Automate can
    # branch on the exit code, then read data/last_run_alerts.txt for the message
    return 1 if (tally.get('REVIEW') or tally.get('FAILED')) else 0


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
    interval = interval or int(os.getenv('EPAL_LURH_WATCH_SECONDS', '5'))
    _setup()
    log.info(f"Watching {INBOX}  (every {interval}s). Drop LURH PDFs here. Press Ctrl+C to stop.")
    seen = {}
    try:
        while True:
            ready, seen = _scan_ready(seen)
            if ready:
                with _run_lock() as lock:
                    if lock.ok:
                        _process_batch(ready)
                    else:
                        log.warning(f'  {lock.msg}')
                        continue          # tenta outra vez no proximo ciclo
                for p in ready:
                    seen.pop(p, None)
            time.sleep(interval)
    except KeyboardInterrupt:
        log.info("Watch stopped.")
    return 0


if __name__ == '__main__':
    if '--watch' in sys.argv:
        sys.exit(watch())
    sys.exit(main())