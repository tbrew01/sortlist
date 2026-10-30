# Sorting list tracking

Windows scripts for tracking draft cows, applying the 36-hour expiry and
insemination rules, and creating the DelPro spreadsheet through Microsoft Excel.

Install all three runtime files together in `C:\Scripts`:

- `move_sortinglist.py`
- `generate_sortinglist.py`
- `tracking_store.py`

Data stays in `C:\delaval-spreadsheet`. See [installation and recovery
instructions](INSTALL_FIX.md) before replacing the existing scripts.

The shared tracking helper provides process locking, atomic JSON replacement,
pre-save backups and detailed per-run logs in `C:\delaval-spreadsheet\logs`.
Existing data, SMS configuration and credentials are not included.

Run the automated checks with `python -m unittest -v test_tracking`. Windows
locking, Excel COM and live SMS operation require validation on the Windows host.
