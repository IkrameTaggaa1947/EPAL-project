# -*- coding: utf-8 -*-
"""One watcher for BOTH pipelines (TUA + LURH).

Instead of two separate scheduled tasks, this single orchestrator processes the
TUA inbox and the LURH inbox on each run. Each pipeline is launched in its OWN
subprocess — that is deliberate: the two code folders each contain a module named
`pipeline.py` and `qa.py`, so importing both into one Python process would clash
on sys.path. A subprocess per pipeline keeps them fully isolated and reproduces
exactly what the old per-domain scheduled tasks did.

Two modes
---------
  (default, no args)  Process each inbox ONCE, then exit.
                      This is what the Windows scheduled task 'EPAL Watch' runs
                      every minute — stateless, nothing lingers between runs.

  --watch             Poll forever, running both pipelines every EPAL_WATCH_SECONDS
                      (default 60s). Used by watch_all.bat for a visible window.

  --status            Print who currently holds the single-watcher lock (see
                      below), then exit without touching the inboxes.

  --release-lock      Release the lock if (and only if) THIS machine holds it.
                      Used by Automation files/unregister_watch_task.bat so switching the
                      designated watcher machine doesn't need to wait out the
                      staleness window.

Drop TUA PDFs into  ../../Drop_New_Licenses/new_TUA_licences
Drop LURH PDFs into ../../Drop_New_Licenses/new_LURH_licences
A single run services both. (Estes sao os caminhos reais que os pipelines leem —
ver INBOX em cada pipeline.py; podem ser mudados com EPAL_TUA_INBOX / EPAL_LURH_INBOX.)

Single-watcher lock
--------------------
The project's shared folder is synced (OneDrive/SharePoint) to every teammate's
laptop. If more than one machine runs the automatic watcher at the same time,
two processes can write to the same master .xlsx concurrently and corrupt it —
see README / LEIA-ME. Nothing about Windows Task Scheduler can prevent a second
person from registering the task on their own PC (registration is inherently
per-machine), so this module adds a best-effort tripwire instead: a small lock
file in the shared folder (data/watch_lock.json) records which machine last ran
successfully. Every cycle:
  - No lock yet, or the lock is THIS machine AND THIS process
                                                   -> claim/refresh it, proceed.
  - Lock belongs to another RUN ON THIS MACHINE and its heartbeat is recent
    (the every-minute task overlapping itself)      -> SKIP this cycle.
  - Lock belongs to ANOTHER machine and its last heartbeat is recent
    (< EPAL_WATCH_STALE_MINUTES, default 10)      -> SKIP this cycle (no writes).
  - Lock belongs to another machine but is stale (that machine's watcher looks
    dead/stopped)                                  -> take over, proceed, log it.
  - Lock file exists but can't be parsed (e.g. caught mid-sync)
                                                    -> SKIP this cycle, to be safe.

This is NOT a real distributed lock (OneDrive/SharePoint sync gives no atomic
cross-machine primitive) — it is a heartbeat convention that fails toward
"skip and log" rather than "proceed and risk corruption". Set EPAL_WATCH_FORCE=1
to bypass it entirely (e.g. for a deliberate manual run).

Env
---
  EPAL_WATCH_SECONDS       loop interval for --watch mode           [default: 60]
  EPAL_WATCH_STALE_MINUTES minutes of silence before another PC may
                            take over the lock                      [default: 10]
  EPAL_WATCH_FORCE         set to 1 to skip the lock check entirely [default: unset]
"""
from __future__ import annotations

import json
import logging
import os
import platform
import subprocess
import sys
import time
import uuid
from datetime import datetime, timezone

HERE = os.path.dirname(os.path.abspath(__file__))          # …/src/automation_of_extraction
SRC = os.path.dirname(HERE)                                 # …/src
DATA = os.path.join(HERE, "data")
# O LOCK_PATH TEM de ficar na pasta partilhada: e o unico sitio onde os PCs se
# veem uns aos outros. Ja o REGISTO nao: vai para %LOCALAPPDATA%\EPAL\logs
# (ver Common_Code_PowerBI/logdir.py). Quem tinha a pasta so para leitura nem
# sequer conseguia arrancar -- rebentava a abrir o .log do dia.
LOCK_PATH = os.path.join(DATA, "watch_lock.json")

PIPELINES = [
    ("TUA", os.path.join(SRC, "Extraction_Code_TUA", "pipeline.py")),
    ("LURH", os.path.join(SRC, "Extraction_Code_LURH", "pipeline.py")),
]

STALE_MINUTES = float(os.getenv("EPAL_WATCH_STALE_MINUTES", "10"))

# A per-PROCESS id (not just hostname): two different Windows accounts, or a
# laptop that got renamed, still shouldn't collide with themselves.
_OWNER_ID = f"{platform.node() or os.getenv('COMPUTERNAME') or 'unknown-host'}:{uuid.getnode():x}"

log = logging.getLogger("watch_all")


def _setup_logging() -> None:
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fmt = logging.Formatter("%(asctime)s  %(levelname)-7s  %(message)s", "%H:%M:%S")
    sys.path.insert(0, os.path.join(SRC, "Common_Code_PowerBI"))
    import logdir
    logdir.add_file_handler(log, "watch", fmt)
    # Console handler only when there IS a console — under pythonw.exe (hidden,
    # scheduled-task runs) sys.stderr is None, so the file log is the record.
    # This matters a lot here: before this file had its own log, a lock
    # conflict would print to a stdout nobody could ever see.
    if sys.stderr is not None:
        ch = logging.StreamHandler()
        ch.setFormatter(fmt)
        log.addHandler(ch)


def _now() -> datetime:
    return datetime.now(timezone.utc)


def _read_lock() -> dict | None:
    """Returns the parsed lock, or None if absent, or {'_corrupt': True} if the
    file exists but couldn't be parsed (e.g. read mid-sync)."""
    if not os.path.isfile(LOCK_PATH):
        return None
    try:
        with open(LOCK_PATH, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        if not isinstance(data, dict) or "owner" not in data or "heartbeat" not in data:
            return {"_corrupt": True}
        return data
    except (OSError, ValueError):
        return {"_corrupt": True}


def _write_lock(owner: str) -> bool:
    """Atomic-ish write (temp file + os.replace, with retries) — same pattern
    already used by make_powerbi_all.dump() for the same reason: OneDrive can
    hold a file locked mid-sync."""
    payload = {"owner": owner, "pid": os.getpid(), "heartbeat": _now().isoformat()}
    tmp = LOCK_PATH + f".tmp{os.getpid()}"
    for attempt in range(5):
        try:
            os.makedirs(DATA, exist_ok=True)
            with open(tmp, "w", encoding="utf-8") as fh:
                json.dump(payload, fh)
            os.replace(tmp, LOCK_PATH)
            return True
        except OSError as exc:
            log.warning("Lock write attempt %d/5 failed: %s", attempt + 1, exc)
            time.sleep(1)
    try:
        os.remove(tmp)
    except OSError:
        pass
    return False


def acquire_lock() -> bool:
    """Returns True if this cycle should proceed to run the pipelines."""
    if os.getenv("EPAL_WATCH_FORCE") == "1":
        log.warning("EPAL_WATCH_FORCE=1 — bypassing the single-watcher lock.")
        _write_lock(_OWNER_ID)
        return True

    lock = _read_lock()

    if lock is None:
        log.info("No watch lock found — claiming it as %s.", _OWNER_ID)
        return _write_lock(_OWNER_ID)

    if lock.get("_corrupt"):
        log.warning(
            "Watch lock at %s exists but could not be parsed (possibly mid-sync). "
            "Skipping this cycle as a precaution — no pipelines run, no masters touched.",
            LOCK_PATH,
        )
        return False

    # The lock is only "mine" when it is this machine AND this process. A run
    # that matched on the machine alone always proceeded, so the lock never
    # protected against two runs on the SAME PC — which is exactly what the
    # every-minute scheduled task produces once a batch takes longer than a
    # minute. (A lock written before this change carries no pid; treat it as
    # ours so upgrading does not stall for the whole staleness window.)
    same_machine = lock["owner"] == _OWNER_ID
    if same_machine and lock.get("pid", os.getpid()) == os.getpid():
        return _write_lock(_OWNER_ID)  # refresh heartbeat

    try:
        last_seen = datetime.fromisoformat(lock["heartbeat"])
    except ValueError:
        log.warning("Watch lock heartbeat unreadable (%r) — skipping this cycle.", lock.get("heartbeat"))
        return False

    age_minutes = (_now() - last_seen).total_seconds() / 60.0
    quem = (f"another run on THIS machine (pid {lock.get('pid', '?')})" if same_machine
            else f"machine {lock['owner']}")
    if age_minutes < STALE_MINUTES:
        log.warning(
            "Watcher already active — %s, last heartbeat %.1f min ago (threshold %.0f). "
            "Skipping this cycle — no pipelines run, no masters touched. "
            "If it is not actually running anymore, either wait it out or "
            "double-click Automation files/unregister_watch_task.bat there.",
            quem, age_minutes, STALE_MINUTES,
        )
        return False

    log.warning(
        "Watch lock held by %s looks abandoned (last heartbeat %.1f min ago, "
        "threshold %.0f) — taking over as %s (pid %d).",
        quem, age_minutes, STALE_MINUTES, _OWNER_ID, os.getpid(),
    )
    return _write_lock(_OWNER_ID)


def release_lock_if_mine() -> None:
    lock = _read_lock()
    if lock and not lock.get("_corrupt") and lock.get("owner") == _OWNER_ID:
        # (machine-level on purpose: Automation files/unregister_watch_task.bat runs in its own
        # process and must be able to release the watcher's lock on this PC)
        try:
            os.remove(LOCK_PATH)
            log.info("Released watch lock (was held by this machine, %s).", _OWNER_ID)
        except OSError as exc:
            log.warning("Could not remove lock file: %s", exc)
    else:
        log.info("No lock held by this machine (%s) — nothing to release.", _OWNER_ID)


def print_status() -> int:
    lock = _read_lock()
    if lock is None:
        print("No watcher has claimed the lock yet.")
    elif lock.get("_corrupt"):
        print(f"Lock file at {LOCK_PATH} exists but is unreadable/corrupt.")
    else:
        try:
            last_seen = datetime.fromisoformat(lock["heartbeat"])
            age_minutes = (_now() - last_seen).total_seconds() / 60.0
            alive = "ALIVE" if age_minutes < STALE_MINUTES else "STALE (eligible for takeover)"
            print(f"Watch lock held by: {lock['owner']}")
            print(f"Last heartbeat:     {age_minutes:.1f} min ago  [{alive}]")
        except ValueError:
            print(f"Watch lock held by: {lock['owner']} (heartbeat unreadable)")
    print(f"This machine's id:  {_OWNER_ID}")
    return 0


def run_once() -> int:
    """Run every pipeline's inbox scan once. One pipeline failing never stops the
    others. Returns 0 if all ran, 1 if any pipeline crashed (a non-zero pipeline
    exit code is NOT a crash — LURH returns 1 when files went to review). Also
    returns 0 (not an error) when the cycle is skipped due to the single-watcher
    lock — that is expected, correct behaviour, not a failure."""
    if not acquire_lock():
        return 0

    hard_error = 0
    for name, script in PIPELINES:
        if not os.path.isfile(script):
            log.error("[%s] pipeline not found: %s", name, script)
            hard_error = 1
            continue
        try:
            # Same interpreter that is running this file (normally the portable
            # runtime\python resolved by _find_python.bat), so each subprocess
            # sees the same dependencies (pymupdf, pandas, …).
            rc = subprocess.run(
                [sys.executable, script],
                cwd=os.path.dirname(script),
            ).returncode
            log.info("[%s] done (exit %d).", name, rc)
        except Exception as exc:                            # never let one pipeline abort the other
            log.error("[%s] FAILED to launch: %s", name, exc)
            hard_error = 1
    return hard_error


def watch() -> int:
    interval = int(os.getenv("EPAL_WATCH_SECONDS", "60"))
    log.info("Watching TUA + LURH inboxes every %ds as %s. Press Ctrl+C to stop.", interval, _OWNER_ID)
    try:
        while True:
            run_once()
            time.sleep(interval)
    except KeyboardInterrupt:
        log.info("Watch stopped.")
    return 0


if __name__ == "__main__":
    _setup_logging()
    args = sys.argv[1:]
    if "--status" in args:
        sys.exit(print_status())
    elif "--release-lock" in args:
        sys.exit(release_lock_if_mine() or 0)
    elif "--watch" in args:
        sys.exit(watch())
    else:
        sys.exit(run_once())
