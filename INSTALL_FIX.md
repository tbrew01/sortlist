# Install the tracking protection fix

These scripts retain the existing Windows paths, Excel COM output, SMS settings,
36-hour timeout and 30-day inactive-record retention. No database or SMS config
is included in this package. This fix does not reconstruct already missing cows
or change how an old first_seen timestamp behaves when a cow reappears.

1. Disable both sorting-list scheduled tasks and wait for their current runs to
   finish. Close any manual runs too. Old copies do not honor the new lock.
2. Copy your existing scripts and draft_tracking.json to a separate backup folder.
   Keep any existing insemination CSVs as well.
3. Copy move_sortinglist.py, generate_sortinglist.py and tracking_store.py together
   into C:\Scripts, replacing the two scripts. Keep tracking_store.py alongside them. Do not replace or
   empty draft_tracking.json, sms_config.json, or your data directories.
4. Ensure both scheduled tasks run these updated copies with the same existing
   Python interpreter. Existing dependencies remain xlrd, openpyxl, pywin32 and
   twilio; tracking_store.py requires no additional packages. Windows and desktop
   Microsoft Excel are still required for spreadsheet generation.
5. Check the tracking JSON contains the cows you expect. The uploaded database
   had only cow 2754; missing records need separate recovery from verified
   historical data. Do not invent arrival times or reset all timers.
6. With the tasks still disabled, run the generator once at an appropriate time
   for a normal production run. It will apply the usual expiry/insemination rules,
   generate the spreadsheet and send SMS if enabled. Check the spreadsheet and
   C:\delaval-spreadsheet\logs, then re-enable the tasks. This live
   Excel check has not been performed in the Linux development environment.

Both scripts hold draft_tracking.json.lock for their entire run, including
spreadsheet generation, so they wait for each other (up to 120 seconds). A timeout
or database failure produces a nonzero task exit and a log entry. Check Task
Scheduler's Last Run Result and the latest run logs after installation. The lock file stays
on disk; its existence does not mean a run is active. Do not delete it to force a run.

Database saves first preserve the previous validated file in
C:\delaval-spreadsheet\draft_tracking.json.backups, then flush a temporary file
and replace the database atomically. Backup failure stops the save. Backups are
not automatically pruned; monitor disk space and retain historical copies in a
separate location. These protect future saves, not data lost before installation.

Missing, malformed or unreadable tracking files stop processing rather than
creating an empty database. To recover, stop both tasks, preserve the damaged file,
inspect a suitable backup's timestamps and cow records, and copy that verified
backup to draft_tracking.json. Do not automatically restore an arbitrary backup:
it may reactivate expired or inseminated cows.

Successfully read insemination CSVs are moved to InseminatedCows\processed only
after their database changes are saved. Other files are left alone. Removal
decisions and active lists are recorded in the per-run logs under C:\delaval-spreadsheet\logs. Protect this log and
backups like other farm records. If a run fails after saving but before archiving
or creating Excel output, investigate the log and remaining inputs before retrying.

Development verification: python -m unittest -v test_tracking
The tests use temporary databases and simulated spreadsheet operations; they do
not send SMS, run real Excel or touch production files. Process locking was
exercised on Linux; the Windows locking branch still needs on-machine validation.


## Detailed diagnostic logs (version 2)

Each run creates a separate UTF-8 file automatically in:

    C:\delaval-spreadsheet\logs

Names begin with move_sortinglist_ or generate_sortinglist_ and include the local
start date/time and a unique run ID. All entries include the process ID and run
ID. This works with the scripts installed together in C:\Scripts; no scheduled
task argument changes or new Python packages are required. The task account must
have permission to create and write the logs folder. If logging cannot start,
the run fails before processing data.

The logs contain:

- Run start, script fingerprint/version, Python version, working directory, local
  clock/timezone, elapsed time, successful finish or full failure traceback.
- Lock wait/acquisition/release and time spent waiting for the other task.
- Loaded and committed tracking records, database hashes, and backup paths.
- Incoming spreadsheet cow numbers, skipped rows, added/reactivated cows and
  preserved first_seen timestamps.
- Insemination CSV hashes and parsed cow numbers, skipped rows and archive paths.
- Every cow's age in hours, insemination match and proposed removal reason.
  COW_REMOVAL_PENDING entries are proposals; DB_COMMITTED confirms persistence.
- Generated spreadsheet cow list, file hash, Excel stages and file moves.
- SMS disabled/misconfigured status, acceptance by Twilio or failure type.
  SMS_ACCEPTED is API acceptance, not confirmed handset delivery. Tokens, config
  contents and recipient phone numbers are not intentionally logged.

RUN_SUCCESS means the script finished. Read preceding warnings for skipped rows
or SMS failures. A run with no incoming spreadsheet records NO_INPUT_FILE. If a
process is killed or the computer loses power, its file may end without a finish
entry. If a scheduled task never starts Python, there will be no run log; check
Task Scheduler History and Last Run Result in that case. Unrelated software
editing a file is not directly observed; the next logged database state and hash
can help locate when a change happened.

Logs and database backups are retained without automatic deletion. Monitor disk
space and archive older logs periodically. They contain farm records.

If cows disappear again, send the log files for BOTH scripts covering the last
correct list through the first incorrect list, plus the current tracking JSON.
Keep the backup folder and processed insemination CSVs available for comparison.
Do not send sms_config.json or SMS credentials. Avoid running the scripts again
just to create diagnostics: normal runs alter the database and may send SMS.

The 12 automated tests cover the database protections, concurrent updates,
per-run success/failure logs, cow decisions, skipped input and SMS log redaction.
Actual Windows locking, Excel output and SMS delivery still require validation on
your Windows computer.
