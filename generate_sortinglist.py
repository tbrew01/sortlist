import os
import logging
from pathlib import Path
import tracking_store
import json
import csv
from datetime import datetime, timedelta


# -------------------- Paths --------------------

DB_FILE = r"C:\delaval-spreadsheet\draft_tracking.json"
INSEM_DIR = r"C:\delaval-spreadsheet\InseminatedCows"
OUT_FILE = r"C:\delaval-spreadsheet\sortuntilinsem\sortinglist.xls"
TEMP_XLSX = r"C:\delaval-spreadsheet\sortuntilinsem\temp_sort.xlsx"
SMS_CONFIG_FILE = r"C:\delaval-spreadsheet\sms_config.json"

# -------------------- Constants --------------------

HEADER_CODE_COL1 = 300001
HEADER_CODE_COL2 = 300096

MAX_HOURS = 36
RETENTION_DAYS = 30


# -------------------- SMS SUPPORT --------------------

def load_sms_config():
    if not os.path.isfile(SMS_CONFIG_FILE):
        logging.info("SMS_CONFIG_MISSING; SMS disabled")
        return {"sms_enabled": False}
    try:
        with open(SMS_CONFIG_FILE, "r") as f:
            return json.load(f)
    except Exception as exc:
        # Do not log config contents or exception text: they may contain credentials.
        logging.warning("SMS_CONFIG_UNREADABLE error_type=%s; SMS disabled", type(exc).__name__)
        return {"sms_enabled": False}


def send_sms_list(cows):
    cfg = load_sms_config()
    if not cfg.get("sms_enabled", False):
        logging.info("SMS_DISABLED")
        return  # no SMS

    sid = cfg.get("twilio_sid")
    token = cfg.get("twilio_token")
    from_num = cfg.get("twilio_from")
    recipients = cfg.get("recipients", [])

    if not sid or not token or not from_num or not recipients:
        logging.warning("SMS_CONFIG_INCOMPLETE; notification not sent")
        return

    from twilio.rest import Client  # pip install twilio

    client = Client(sid, token)
    cow_list_str = ", ".join(str(c) for c in cows)
    body = f"Extended draft list cows: {cow_list_str}"

    for recipient_index, r in enumerate(recipients, start=1):
        try:
            message = client.messages.create(to=r, from_=from_num, body=body)
            logging.info("SMS_ACCEPTED recipient_index=%s message_sid=%s status=%s cows=%s",
                         recipient_index, message.sid, message.status, cows)
        except Exception as exc:
            logging.error("SMS_FAILED recipient_index=%s error_type=%s; spreadsheet unchanged",
                          recipient_index, type(exc).__name__)


# -------------------- DB and Helper Functions --------------------

def load_db():
    return tracking_store.load_db(DB_FILE)


def save_db(db):
    tracking_store.save_db(DB_FILE, db)


def parse_iso(dt_str):
    return datetime.fromisoformat(dt_str)


def detect_delimiter(first):
    return ";" if ";" in first and "," not in first else ","


def load_inseminated_cows(processed_files=None):
    cows = set()
    if not os.path.isdir(INSEM_DIR):
        logging.warning("INSEMINATION_DIRECTORY_MISSING path=%s", INSEM_DIR)
        return cows

    for fname in os.listdir(INSEM_DIR):
        path = os.path.join(INSEM_DIR, fname)
        if not os.path.isfile(path):
            continue

        if not fname.lower().endswith(".csv"):
            continue

        try:
            logging.info("INSEMINATION_INPUT file=%s", json.dumps(tracking_store.file_details(path)))
            file_cows = set()
            with open(path, newline="") as f:
                first = f.readline()
                if not first:
                    logging.warning("INSEMINATION_FILE_EMPTY path=%s; left in place", path)
                    continue
                delim = detect_delimiter(first)
                f.seek(0)
                reader = csv.reader(f, delimiter=delim)

                header = True
                for row_number, row in enumerate(reader, start=1):
                    if not row:
                        continue
                    if header and "Animal Number" in row[0]:
                        header = False
                        continue
                    header = False
                    try:
                        c = str(int(float(row[0].strip())))
                        cows.add(c)
                        file_cows.add(c)
                    except Exception:
                        logging.warning("INSEMINATION_ROW_SKIPPED file=%s row=%s; invalid cow number", path, row_number)
        except Exception as exc:
            raise RuntimeError(f"Cannot read insemination file {path}") from exc

        if processed_files is not None:
            processed_files.append(path)
        logging.info("INSEMINATION_COWS file=%s cows=%s", path, sorted(file_cows))

    return cows


# -------------------- DRAFT CLEANUP --------------------

def apply_rules_and_cleanup(db, processed_files=None):
    now = datetime.now()
    inseminated = load_inseminated_cows(processed_files)
    logging.info("RULES_START now=%s max_hours=%s retention_days=%s inseminated=%s",
                 now.isoformat(), MAX_HOURS, RETENTION_DAYS, sorted(inseminated))

    for cow_id, rec in list(db.items()):
        first_seen = parse_iso(rec["first_seen"])
        active = rec.get("active", True)
        logging.info("COW_EVALUATED cow=%s active=%s first_seen=%s age_hours=%.4f insemination_match=%s",
                     cow_id, active, rec["first_seen"], (now - first_seen).total_seconds() / 3600,
                     cow_id in inseminated)

        if active:
            if now - first_seen > timedelta(hours=MAX_HOURS):
                rec["active"] = False
                rec["cleared_reason"] = "timeout_36h"
                logging.info("COW_REMOVAL_PENDING cow=%s reason=timeout_36h first_seen=%s; effective after DB_COMMITTED", cow_id, rec["first_seen"])
                rec["cleared_time"] = now.isoformat(timespec="seconds")
            elif cow_id in inseminated:
                rec["active"] = False
                rec["cleared_reason"] = "inseminated"
                logging.info("COW_REMOVAL_PENDING cow=%s reason=inseminated; effective after DB_COMMITTED", cow_id)
                rec["cleared_time"] = now.isoformat(timespec="seconds")

    for cow_id in list(db.keys()):
        rec = db[cow_id]
        if not rec.get("active", True) and rec.get("cleared_time"):
            cleared_dt = parse_iso(rec["cleared_time"])
            if now - cleared_dt > timedelta(days=RETENTION_DAYS):
                logging.info("COW_RECORD_PURGED cow=%s cleared_time=%s retention_days=%s",
                             cow_id, rec["cleared_time"], RETENTION_DAYS)
                del db[cow_id]


# -------------------- XLS CREATION (Excel COM) --------------------

def create_delpro_xls(db):
    """Creates a compliant BIFF8 .xls file via Excel COM Automation."""

    import openpyxl  # pip install openpyxl
    import win32com.client as win32  # pip install pywin32

    active_cows = sorted(
        int(cid) for cid, rec in db.items() if rec.get("active", True)
    )

    # --- STEP 1: Build temporary XLSX ---
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.cell(1, 1).value = HEADER_CODE_COL1
    ws.cell(1, 2).value = HEADER_CODE_COL2

    r = 2
    for cid in active_cows:
        ws.cell(r, 1).value = cid
        ws.cell(r, 2).value = 1
        r += 1

    os.makedirs(os.path.dirname(TEMP_XLSX), exist_ok=True)
    wb.save(TEMP_XLSX)
    logging.info("TEMP_XLSX_SAVED file=%s cows=%s", json.dumps(tracking_store.file_details(TEMP_XLSX)), active_cows)

    # --- STEP 2: Convert XLSX → XLS using real Excel ---
    logging.info("EXCEL_START output=%s", OUT_FILE)
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    temp_path = os.path.abspath(TEMP_XLSX)
    out_path = os.path.abspath(OUT_FILE)

    wb = excel.Workbooks.Open(temp_path)
    wb.SaveAs(out_path, FileFormat=56)  # 56 = Excel 97-2003 BIFF8
    logging.info("EXCEL_SAVEAS_RETURNED output=%s cow_count=%s", out_path, len(active_cows))
    wb.Close()
    excel.Quit()
    logging.info("OUTPUT_XLS_SAVED file=%s cows=%s", json.dumps(tracking_store.file_details(OUT_FILE)), active_cows)

    # cleanup temporary file
    try:
        os.remove(TEMP_XLSX)
    except OSError:
        logging.exception("TEMP_XLSX_CLEANUP_FAILED path=%s", TEMP_XLSX)

    # SMS cows
    if active_cows:
        send_sms_list(active_cows)
    else:
        logging.info("SMS_SKIPPED_NO_ACTIVE_COWS")


# -------------------- MAIN --------------------

def run_locked():
    db = load_db()
    processed_files = []
    apply_rules_and_cleanup(db, processed_files)
    save_db(db)
    # Only consume insemination inputs after their updates are safely committed.
    import shutil
    from uuid import uuid4
    for source in processed_files:
        archive = Path(INSEM_DIR) / "processed"
        archive.mkdir(exist_ok=True)
        destination = archive / (uuid4().hex + "_" + Path(source).name)
        shutil.move(source, str(destination))
        logging.info("INSEMINATION_ARCHIVED source=%s destination=%s", source, destination)
    logging.info("Active draft cows: %s", ", ".join(
        sorted(cid for cid, rec in db.items() if rec.get("active", True))))
    create_delpro_xls(db)


def main():
    with tracking_store.logged_run(DB_FILE, "generate_sortinglist"):
        with tracking_store.tracking_lock(DB_FILE):
            run_locked()


if __name__ == "__main__":
    main()
