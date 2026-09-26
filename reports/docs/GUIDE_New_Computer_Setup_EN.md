# Guide — Setting up the project on a new team computer

For whoever is bringing the project to a new colleague (not for the
colleague to read alone — that's what `LEIA-ME_PRIMEIRO.txt` in the project
root is for). This guide is for confirming the install went well and for
deciding, as a team, who runs the automatic watcher.

---

## Before you start

Make sure the `EPAL-project` folder is **fully synced** in OneDrive on that
computer (green icon, no pending blue clouds). If it's still syncing, the
steps below can fail in confusing ways — that's not a bug in the project,
just wait it out.

## Step 1 — Run the configurator

Have your colleague double-click:

    CONFIGURAR_ESTE_PC.bat

No administrator rights needed and nothing gets installed — it just
confirms it can find the Python already bundled in the shared folder and
writes this PC's local settings.

## Step 2 — Read the final diagnostic

The `.bat` ends by printing a diagnostic. What it should show:

| Line | What it should say |
|---|---|
| Python | `3.13.1`, and the path should end in `runtime\python\python.exe` |
| ROOT, data/powerbi, Drop TUA, Drop LURH, reports, outputs | `[OK]` on all |
| epal.local.ini | `[OK]` |
| local venv | `[MISSING]` — **this is expected**, the portable runtime takes priority and the local venv is never actually used |
| Power BI DataFolder | `-> CORRECT` |

If **Python** says "WARNING: the portable Python doesn't exist in this
folder yet" and falls back to the computer's own Python (or none at all),
the most likely cause is that the folder hasn't finished syncing —
re-running the `.bat` a few minutes later almost always fixes it. Only
re-run `INSTALAR_PYTHON_PORTATIL.bat` if this persists after syncing has
finished (that script is only meant to run once, ever, on the original
computer).

Any other line that doesn't say `[OK]` — send a screenshot of this screen
before going further.

## Step 3 — Do NOT test with a real PDF yet

`CONFIGURAR_ESTE_PC.bat` already IS the full test — it doesn't touch any
PDFs or the master files. There's no need (and it's not advisable) to drop
a test PDF just to confirm the install worked.

## Step 4 — Decide who runs the automatic watcher

**Only ONE computer on the team may have the automatic watcher
(`register_watch_task.bat`) turned on.** The folder is shared — if two
computers process the same PDFs at the same time, the master files can get
corrupted.

This was already a written warning, but until now nothing technically
prevented it. As of 2026-08-24, `watch_all.py` has an extra safeguard: each
cycle records which computer is active in a shared file
(`src/automation_of_extraction/data/watch_lock.json`), and a second
computer trying to run at the same time is automatically skipped (touches
nothing, just logs a warning) — unless the first computer has been
inactive for more than 10 minutes, in which case the new one takes over on
its own. **This is an extra safety net, not a substitute for the team's
agreement** — you should still explicitly decide which computer is the
watcher machine.

To check at any time which computer is currently active, run (using the
project's Python):

    runtime\python\python.exe src\automation_of_extraction\watch_all.py --status

On every other computer, always use `PROCESSAR_AGORA.bat`
manually. If you ever want to move the watcher to a different computer,
run `unregister_watch_task.bat` on the old one first — that releases the
lock immediately instead of waiting out the 10-minute staleness window.

## If something goes wrong

Same list as `LEIA-ME_PRIMEIRO.txt` — reproduced here for convenience:

- **Power BI says it can't find the data** → close Power BI and
  double-click `ATUALIZAR_POWERBI.bat`.
- **"NO PYTHON FOUND"** → the folder may still be syncing; wait and try
  again.
- **A licence went to `data\review`** → a problem reading that specific
  PDF, not an install issue; there's a `.txt` next to it explaining why.
- To ask for help, have your colleague run this and send back the text it
  prints:

      runtime\python\python.exe src\epal_config.py

---

*Note (2026-08-24): the install was confirmed working on the first
computer (Ikrame's). This guide hasn't been tested yet on the process of
adding a SECOND computer — the first time you use it, it's worth noting
anything that doesn't match what's described here.*
