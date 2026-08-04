"""EPAL TUA batch pipeline — the entry point for non-technical operators.

Drop TUA PDFs into the 'inbox' folder, then run this file (or double-click
run_tua.bat). For each PDF it:
  1. extracts the data,
  2. runs quality checks (verdict: OK / REVIEW / FAILED),
  3. updates the multi-sheet master (backed up first) unless extraction FAILED,
  4. exports the master to CSVs for Power BI,
  5. files the PDF into processed/ (OK) or review/ (REVIEW/FAILED) with a note,
  6. logs everything and appends to a plain-language qa_report.xlsx.

Folders (created automatically under ./data; override the root with EPAL_TUA_DATA):
  data/inbox      <- DROP PDFs HERE
  data/processed  -> extracted cleanly
  data/review     -> a human should check these (a .txt says why)
  data/extracted  -> per-file JSON + Excel
  data/logs       -> run logs
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

import extract_tua as E
import qa

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.getenv('EPAL_TUA_DATA', os.path.join(HERE, 'data'))
INBOX = os.path.join(DATA, 'inbox')
PROCESSED = os.path.join(DATA, 'processed')
REVIEW = os.path.join(DATA, 'review')
EXTRACTED = os.path.join(DATA, 'extracted')
LOGS = os.path.join(DATA, 'logs')
MASTER = os.getenv('EPAL_TUA_MASTER', os.path.join(HERE, 'master_tua.xlsx'))
SECTIONS_MASTER = os.getenv('EPAL_TUA_SECTIONS_MASTER', os.path.join(HERE, 'master_all_sections.xlsx'))
QA_REPORT = os.path.join(HERE, 'qa_report.xlsx')

log = logging.getLogger('tua')


def _setup():
    for d in (INBOX, PROCESSED, REVIEW, EXTRACTED, LOGS):
        os.makedirs(d, exist_ok=True)
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fmt = logging.Formatter('%(asctime)s  %(levelname)-7s  %(message)s', '%H:%M:%S')
    fh = logging.FileHandler(os.path.join(LOGS, f"tua_{datetime.now():%Y%m%d}.log"), encoding='utf-8')
    fh.setFormatter(fmt)
    log.addHandler(fh)
    # Console handler only when there IS a console — under pythonw.exe (hidden,
    # scheduled-task runs) sys.stderr is None, so the file log is the record.
    if sys.stderr is not None:
        ch = logging.StreamHandler()
        ch.setFormatter(fmt)
        log.addHandler(ch)


def _move(pdf, dest_dir, note=None):
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, os.path.basename(pdf))
    if os.path.abspath(dest) != os.path.abspath(pdf):
        shutil.move(pdf, dest)
    if note:
        with open(os.path.splitext(dest)[0] + '_REVIEW.txt', 'w', encoding='utf-8') as f:
            f.write(note)
    return dest


def process_one(pdf):
    """Extract, QA-check, update the master, and file the PDF. Returns a report dict."""
    name = os.path.basename(pdf)
    log.info(f"processing {name}")
    try:
        data = E.extract(pdf)
    except Exception as exc:                       # never let one bad PDF stop the batch
        log.exception(f"EXTRACTION CRASHED for {name}: {exc}")
        _move(pdf, REVIEW, f"Extraction crashed: {exc}\nSend this PDF to the technical team.")
        return {'file_name': name, 'verdict': 'FAILED', 'conditions': 0,
                'autocontrolo': 0, 'reasons': f'extraction crashed: {exc}',
                'timestamp': datetime.now().isoformat(timespec='seconds')}

    _, xlsx = E.write_outputs(data, EXTRACTED)
    verdict, reasons = qa.check(data)
    ncond, nauto = len(data['condicoes_rejeicao']), len(data['autocontrolo'])
    log.info(f"  -> {verdict}  ({ncond} conditions, {nauto} monitoring rows)")
    for r in reasons:
        log.info(f"     {r}")

    if verdict != qa.FAILED:
        mp, counts, deferred = E.update_master(data, MASTER)
        if deferred:
            log.warning(f"  master locked — saved to {mp}. Close master_tua.xlsx and re-run.")
        else:
            E.export_master_csv(MASTER)
            log.info(f"  master updated: {counts}")
        # Consolidated all-sections workbook (one sheet per section, all certificates).
        _, scounts, sdef = E.update_sections_master(data, SECTIONS_MASTER)
        log.info(f"  all-sections master {'locked (PENDING)' if sdef else 'updated'}: {len(scounts)} sheets")

    dest = PROCESSED if verdict == qa.OK else REVIEW
    note = None if verdict == qa.OK else (
        f"{name}\nVerdict: {verdict}\n\nWhy this needs a look:\n  - " + "\n  - ".join(reasons) +
        "\n\nWhat to do: open the matching file in data/extracted/ and check the flagged "
        "fields against the PDF. If the data is correct, move this PDF to data/processed/.")
    _move(pdf, dest, note)
    return {'file_name': name, 'Nº TUA': data['dados_gerais'].get('Nº TUA', ''),
            'Estabelecimento': data['dados_gerais'].get('Estabelecimento', ''),
            'verdict': verdict, 'conditions': ncond, 'autocontrolo': nauto, 'xlsx': xlsx,
            'reasons': ' | '.join(reasons), 'timestamp': datetime.now().isoformat(timespec='seconds')}


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
    if not rows:
        return
    new = pd.DataFrame(rows)
    if os.path.isfile(QA_REPORT):
        try:
            new = pd.concat([pd.read_excel(QA_REPORT), new], ignore_index=True)
        except Exception:
            pass
    try:
        new.to_excel(QA_REPORT, index=False)
    except PermissionError:
        alt = os.path.splitext(QA_REPORT)[0] + '.PENDING.xlsx'
        new.to_excel(alt, index=False)
        log.warning(f"qa_report.xlsx is open — wrote {alt} instead.")


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
        sys.path.insert(0, os.path.dirname(HERE))      # …/extraction (make_powerbi_all lives here)
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
    """Process a list of PDFs, write the report, log a tally. Returns the tally."""
    rows = [process_one(p) for p in pdfs]
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


def main():
    _setup()
    pdfs = sorted(glob.glob(os.path.join(INBOX, '*.pdf')))
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
    for p in sorted(glob.glob(os.path.join(INBOX, '*.pdf'))):
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
                _process_batch(ready)
            time.sleep(interval)
    except KeyboardInterrupt:
        log.info("Watch stopped.")
    return 0


if __name__ == '__main__':
    sys.exit(watch() if '--watch' in sys.argv[1:] else main())
