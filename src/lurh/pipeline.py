"""EPAL LURH batch pipeline — the entry point for non-technical operators.

Drop LURH PDFs into the 'inbox' folder, then run this file (or double-click
run_lurh.bat). For each PDF it:
  1. extracts the data,
  2. runs quality checks (verdict: OK / REVIEW / FAILED),
  3. updates the multi-sheet master (backed up first) unless extraction FAILED,
  4. exports the master to CSVs for Power BI,
  5. files the PDF into processed/ (OK) or review/ (REVIEW/FAILED) with a note,
  6. logs everything and appends to a plain-language qa_report.xlsx.

Folders (created automatically under ./data; override the root with EPAL_LURH_DATA):
  data/inbox      <- DROP PDFs HERE
  data/processed  -> extracted cleanly
  data/review     -> a human should check these (a .txt says why)
  data/extracted  -> per-file JSON + Excel
  data/logs       -> run logs
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

import extract_lurh as E
import qa

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.getenv('EPAL_LURH_DATA', os.path.join(HERE, 'data'))
INBOX = os.path.join(DATA, 'inbox')
PROCESSED = os.path.join(DATA, 'processed')
REVIEW = os.path.join(DATA, 'review')
EXTRACTED = os.path.join(DATA, 'extracted')
LOGS = os.path.join(DATA, 'logs')
MASTER = os.getenv('EPAL_LURH_MASTER', os.path.join(HERE, 'data', 'reference', 'master_lurh.xlsx'))
QA_REPORT = os.path.join(HERE, 'qa_report.xlsx')

log = logging.getLogger('lurh')


def _setup():
    for d in (INBOX, PROCESSED, REVIEW, EXTRACTED, LOGS):
        os.makedirs(d, exist_ok=True)
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fmt = logging.Formatter('%(asctime)s  %(levelname)-7s  %(message)s', '%H:%M:%S')
    fh = logging.FileHandler(os.path.join(LOGS, f"lurh_{datetime.now():%Y%m%d}.log"), encoding='utf-8')
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
    """Extract, QA-check, update the master, and file the PDF. Returns a report dict.

    A PDF only goes to data/processed/ when its data is ACTUALLY in the master.
    Anything else (extraction crash, QA FAILED, master file locked) goes to
    data/review/ with a note, and the report row carries 'added_to_master' plus
    a one-line 'problem' naming the file and the exact reason — this is what
    the Power Automate flow displays."""
    name = os.path.basename(pdf)
    log.info(f"processing {name}")
    try:
        data = E.extract(pdf)
    except Exception as exc:                       # never let one bad PDF stop the batch
        log.exception(f"EXTRACTION CRASHED for {name}: {exc}")
        problem = f"{name}: extraction crashed ({exc}) — NOT added to master_lurh"
        _move(pdf, REVIEW, f"Extraction crashed: {exc}\nSend this PDF to the technical team.")
        return {'file_name': name, 'verdict': 'FAILED', 'conditions': 0,
                'autocontrolo': 0, 'added_to_master': False, 'problem': problem,
                'reasons': f'extraction crashed: {exc}',
                'timestamp': datetime.now().isoformat(timespec='seconds')}

    _, xlsx = E.write_outputs(data, EXTRACTED)
    verdict, reasons = qa.check(data)
    ncond = len(data['condicoes_descarga'])
    nauto = len(data['autocontrolo'].get('rows', []))
    log.info(f"  -> {verdict}  ({ncond} conditions, {nauto} autocontrolo rows)")
    for r in reasons:
        log.info(f"     {r}")

    added, problem = False, ''
    if verdict == qa.FAILED:
        problem = (f"{name}: NOT added to master_lurh — extraction failed QA: "
                   + "; ".join(r.replace('[FAILED] ', '') for r in reasons if 'FAILED' in r))
    else:
        mp, counts, deferred = E.update_master(data, MASTER)
        if deferred:
            problem = (f"{name}: NOT added to master_lurh — the master workbook is open/"
                       f"locked in Excel. Data saved to '{os.path.basename(mp)}'; close "
                       f"master_lurh.xlsx and re-run this PDF.")
            log.warning(f"  {problem}")
        else:
            added = True
            E.export_master_csv(MASTER)
            log.info(f"  master updated: {counts}")

    # Only PDFs whose data is really in the master are filed as processed.
    dest = PROCESSED if (added and verdict == qa.OK) else REVIEW
    if dest is REVIEW:
        head = problem or f"{name}: added to master, but flagged {verdict} — verify against the PDF."
        note = (f"{name}\nVerdict: {verdict}\nAdded to master: {'yes' if added else 'NO'}\n\n"
                f"Problem:\n  {head}\n\nQA details:\n  - " + "\n  - ".join(reasons or ['(none)']) +
                "\n\nWhat to do: open the matching file in data/extracted/ and check the flagged "
                "fields against the PDF. If the data is correct, move this PDF to data/processed/.")
        _move(pdf, REVIEW, note)
    else:
        _move(pdf, PROCESSED)
    return {'file_name': name, 'Nº Licença': data['dados_gerais'].get('Nº Licença', ''),
            'Estabelecimento': data['tratamento'].get('Designação', ''),
            'verdict': verdict, 'conditions': ncond, 'autocontrolo': nauto, 'xlsx': xlsx,
            'added_to_master': added, 'problem': problem,
            'reasons': ' | '.join(reasons), 'timestamp': datetime.now().isoformat(timespec='seconds')}


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


def _write_alerts(rows):
    """After every batch, write the outcome where Power Automate can pick it up:
      data/last_run_alerts.json — full structure (parse with 'Parse JSON')
      data/last_run_alerts.txt  — one line per problem file, ready to display
    Both files are OVERWRITTEN each run, so they always describe the last run.
    Also prints each problem line as 'ALERT: ...' to stdout for flows that read
    the script output directly."""
    problems = [r for r in rows if not r.get('added_to_master')]
    payload = {
        'run_timestamp': datetime.now().isoformat(timespec='seconds'),
        'total_files': len(rows),
        'added_to_master': sum(1 for r in rows if r.get('added_to_master')),
        'not_added': len(problems),
        'alerts': [{'file_name': r['file_name'],
                    'verdict': r.get('verdict', ''),
                    'problem': r.get('problem', ''),
                    'qa_reasons': r.get('reasons', '')} for r in problems],
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
                f.write(f"All {len(rows)} file(s) added to master_lurh — no problems.\n")
    except Exception as e:
        log.warning(f"could not write alert files: {e}")
    for r in problems:
        log.error(f"ALERT: {r['problem']}")
        print(f"ALERT: {r['problem']}", flush=True)


def _process_batch(pdfs):
    """Process a list of PDFs, write the report, log a tally. Returns the tally."""
    rows = [process_one(p) for p in pdfs]
    _append_report(rows)
    _write_alerts(rows)
    tally = {}
    for r in rows:
        tally[r['verdict']] = tally.get(r['verdict'], 0) + 1
    log.info("-" * 52)
    log.info("DONE  " + "  ".join(f"{v}:{n}" for v, n in sorted(tally.items())))
    if tally.get('REVIEW') or tally.get('FAILED'):
        log.info(f"Some licences need a look — see: {REVIEW}")
    _open_excel(rows)
    return tally


def main():
    _setup()
    pdfs = sorted(glob.glob(os.path.join(INBOX, '*.pdf')))
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
    interval = interval or int(os.getenv('EPAL_LURH_WATCH_SECONDS', '5'))
    _setup()
    log.info(f"Watching {INBOX}  (every {interval}s). Drop LURH PDFs here. Press Ctrl+C to stop.")
    seen = {}
    try:
        while True:
            ready, seen = _scan_ready(seen)
            if ready:
                _process_batch(ready)
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