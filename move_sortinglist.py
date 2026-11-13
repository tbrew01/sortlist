import os
import logging
from pathlib import Path
import tracking_store
import json
import shutil
from datetime import datetime


# Paths
SOURCE_FILE = r"C:\delaval-spreadsheet\sortinglist.xls"
DEST_FILE = r"C:\delaval-spreadsheet\sortuntilinsem\sortinglist.xls"
DB_FILE = r"C:\delaval-spreadsheet\draft_tracking.json"

# Header codes (from your file)
HEADER_CODE_COL1 = 300001  # Cow ID code
HEADER_CODE_COL2 = 300096  # Draft flag code (1 = draft)


def load_db():
    return tracking_store.load_db(DB_FILE)


def save_db(db):
    tracking_store.save_db(DB_FILE, db)


def read_cows_from_xls(path):
    """
    Reads sortinglist.xls and returns a list of cow IDs that have draft flag = 1.
    Expected format:
        Row 0: [300001, 300096]
        Row 1+: [CowID, DraftFlag]
    """
    cows = []

    import xlrd  # pip install xlrd

    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    logging.info("XLS_READ file=%s rows=%s columns=%s", path, sheet.nrows, sheet.ncols)

    if sheet.nrows < 2:
        return cows

    # Loop rows starting at row 1 (row 0 is header)
    for row_idx in range(1, sheet.nrows):
        try:
            cow_val = sheet.cell_value(row_idx, 0)
            flag_val = sheet.cell_value(row_idx, 1)

            # Convert to int cleanly
            cow_id = int(float(cow_val))
            draft_flag = int(float(flag_val))

            if draft_flag == 1:
                cows.append(str(cow_id))
            else:
                logging.info("XLS_ROW_NOT_DRAFT row=%s cow=%s flag=%s", row_idx + 1, cow_id, draft_flag)

        except Exception:
            logging.exception("XLS_ROW_SKIPPED file=%s row=%s", path, row_idx + 1)
            continue

    logging.info("XLS_DRAFT_COWS file=%s cows=%s", path, sorted(cows))
    return cows


def run_locked():
    # Only process if source XLS exists
    if not os.path.isfile(SOURCE_FILE):
        logging.info("NO_INPUT_FILE path=%s; no database or spreadsheet changes", SOURCE_FILE)
        return

    logging.info("INPUT_XLS file=%s", json.dumps(tracking_store.file_details(SOURCE_FILE)))

    now = datetime.now().isoformat(timespec="seconds")

    # Load tracking DB
    db = load_db()

    # Read cow list from incoming file
    try:
        cows_to_draft = read_cows_from_xls(SOURCE_FILE)
    except Exception as e:
        raise RuntimeError(f"Error reading {SOURCE_FILE}") from e

    # Update database records
    for cow_id in cows_to_draft:
        if cow_id not in db:
            logging.info("COW_ADDED cow=%s first_seen=%s", cow_id, now)
            db[cow_id] = {
                "first_seen": now,
                "last_seen": now,
                "active": True,
                "cleared_reason": None,
                "cleared_time": None
            }
        else:
            logging.info("COW_SEEN_AGAIN cow=%s previous=%s",
                         cow_id, json.dumps(db[cow_id], sort_keys=True))
            if not db[cow_id]["active"]:
                previous_first_seen = db[cow_id]["first_seen"]
                db[cow_id]["first_seen"] = now
                db[cow_id]["cleared_reason"] = None
                db[cow_id]["cleared_time"] = None
                logging.info(
                    "COW_REACTIVATION_PENDING cow=%s old_first_seen=%s new_first_seen=%s; "
                    "fresh 36-hour timer, previous removal cleared; effective after DB_COMMITTED",
                    cow_id, previous_first_seen, now)
            else:
                logging.info("COW_TIMER_PRESERVED cow=%s first_seen=%s",
                             cow_id, db[cow_id]["first_seen"])
            db[cow_id]["last_seen"] = now
            db[cow_id]["active"] = True

    # Save updated DB
    save_db(db)
    logging.info("Incoming draft cows: %s", ", ".join(cows_to_draft))

    # Ensure destination directory exists
    dest_dir = os.path.dirname(DEST_FILE)
    os.makedirs(dest_dir, exist_ok=True)

    # Move file to sortuntilinsem
    try:
        if os.path.exists(DEST_FILE):
            logging.info("OUTPUT_REPLACED_BY_INCOMING file=%s",
                         json.dumps(tracking_store.file_details(DEST_FILE)))
            os.remove(DEST_FILE)
        shutil.move(SOURCE_FILE, DEST_FILE)
        logging.info("INPUT_MOVED destination=%s", json.dumps(tracking_store.file_details(DEST_FILE)))
    except Exception as e:
        raise RuntimeError(f"Error moving file {SOURCE_FILE}") from e


def main():
    with tracking_store.logged_run(DB_FILE, "move_sortinglist"):
        with tracking_store.tracking_lock(DB_FILE):
            run_locked()


if __name__ == "__main__":
    main()
