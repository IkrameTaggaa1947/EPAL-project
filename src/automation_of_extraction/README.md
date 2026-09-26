# automation_of_extraction

One watcher for **both** extraction pipelines (TUA + LURH). Replaces the two
separate per-domain scheduled tasks with a single one that runs every **1 minute**.

## Files

| File | What it does |
|------|--------------|
| `watch_all.py` | Runs each pipeline's inbox scan once (each in its own subprocess). `--watch` = poll forever every `EPAL_WATCH_SECONDS` (default 60s). `--status` = show who currently holds the single-watcher lock. `--release-lock` = release it if this machine holds it. |
| `Automation files\register_watch_task.bat` | Double-click **once**: registers the hidden background task **"EPAL Watch"** (every 1 min, survives reboot) and removes the old "EPAL TUA Watch" / "EPAL LURH Watch" tasks. |
| `Automation files\unregister_watch_task.bat` | Turns the watcher off (deletes "EPAL Watch") and releases the single-watcher lock if this machine held it. |
| `watch_all.bat` | Manual continuous mode in a visible window (for testing). |

## Only ONE machine should run the automatic watcher

The shared folder syncs to every laptop, so if two people both register the
watcher, two processes can write to the same master `.xlsx` at the same time
and corrupt it. Windows Task Scheduler can't prevent this on its own —
registration is per-machine — so `watch_all.py` now guards itself with a
**single-watcher lock**: a small file at `data/watch_lock.json` records which
machine last ran a cycle. If another machine's heartbeat is recent, this
machine skips the cycle (no files touched) and logs why. If that machine goes
quiet for more than `EPAL_WATCH_STALE_MINUTES` (default 10), the lock is
considered abandoned and this machine takes over automatically.

This is a best-effort tripwire, not a real distributed lock — OneDrive/SharePoint
sync gives no atomic cross-machine primitive, and a corrupt/mid-sync lock file
is treated as "skip this cycle" rather than "guess and proceed". Check who's
active any time by running `watch_all.py --status` (via the project's Python,
see `_find_python.bat`). To hand the role to a different machine immediately
instead of waiting out the staleness window, run `Automation files\unregister_watch_task.bat`
on the old machine first.

## Where to drop PDFs

Drop everything in the project-root **`Drop_New_Licenses/`** folder, in the
matching sub-folder:

- **TUA** → `../../Drop_New_Licenses/new_TUA_licences`
- **LURH** → `../../Drop_New_Licenses/new_LURH_licences`

(Each pipeline reads its sub-folder directly — override with `EPAL_TUA_INBOX` /
`EPAL_LURH_INBOX` if needed.)

## Notes

- Each run is **stateless**: it processes whatever is in the inboxes and exits,
  so nothing lingers between runs. Processed PDFs move to the archive `data/pdfs/NEW/<regime>`,
  anything that fails QA goes to `data/review` with a `.txt` explaining why.
- `watch_all.py` now logs its own orchestration events (cycle start, per-pipeline
  exit codes, lock claims/skips/takeovers) to `data/logs/watch_YYYYMMDD.log` —
  under the hidden scheduled task (`pythonw.exe`) there is no console, so this
  log file is the only record of what the automatic watcher actually did.
- The subprocess-per-pipeline design is deliberate: both code folders contain a
  `pipeline.py` and `qa.py`, so isolating them avoids a `sys.path` name clash.
- After changing pipeline code, no re-registration is needed — the task points at
  `watch_all.py`, which always calls the current pipeline code. You only re-run
  `Automation files\register_watch_task.bat` if you **move/rename** this folder.
