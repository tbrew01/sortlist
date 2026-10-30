"""Shared database protection for both scheduled sorting-list scripts."""
import json
import hashlib
import logging
import os
import tempfile
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from uuid import uuid4


def file_details(path):
    """Identify a data file without logging its raw contents."""
    path = Path(path)
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
        stat = os.fstat(handle.fileno())
    return {"path": str(path.resolve()), "bytes": stat.st_size,
            "modified_local": datetime.fromtimestamp(stat.st_mtime).isoformat(),
            "sha256": digest.hexdigest()}


@contextmanager
def logged_run(db_file, script):
    """Use one log per run so overlapping processes never rotate each other's log."""
    log_dir = Path(db_file).parent / "logs"
    log_dir.mkdir(exist_ok=True)
    run_id = uuid4().hex
    log_path = log_dir / (script + "_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                          + "_" + run_id + ".log")
    handler = logging.FileHandler(log_path, encoding="utf-8")
    handler.setFormatter(logging.Formatter(
        "%(asctime)s %(levelname)s pid=%(process)d run=" + run_id + " %(message)s"))
    root = logging.getLogger()
    previous_level = root.level
    root.setLevel(logging.INFO)
    root.addHandler(handler)
    started = time.monotonic()
    try:
        logging.info("RUN_START script=%s version=tracking-logs-v2 python=%s cwd=%s db=%s local_time=%s",
                     script, os.sys.version.split()[0], os.getcwd(),
                     Path(db_file).resolve(), datetime.now().astimezone().isoformat())
        logging.info("SCRIPT_FILE file=%s", json.dumps(file_details(Path(__file__).with_name(script + ".py"))))
        yield
    except BaseException:
        logging.exception("RUN_FAILED elapsed_seconds=%.3f; see last completed stage above",
                          time.monotonic() - started)
        raise
    else:
        logging.info("RUN_SUCCESS elapsed_seconds=%.3f", time.monotonic() - started)
    finally:
        root.removeHandler(handler)
        handler.close()
        root.setLevel(previous_level)


@contextmanager
def tracking_lock(db_file, timeout=120):
    """Lock a stable sidecar file for the entire run; OS releases on exit/crash."""
    lock_path = Path(str(db_file) + ".lock")
    started = time.monotonic()
    logging.info("LOCK_WAIT path=%s timeout_seconds=%s", lock_path, timeout)
    with open(lock_path, "a+b") as handle:
        handle.seek(0, os.SEEK_END)
        if handle.tell() == 0:
            handle.write(b"\0")
            handle.flush()
        deadline = time.monotonic() + timeout
        while True:
            try:
                if os.name == "nt":
                    import msvcrt
                    handle.seek(0)
                    msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except OSError:
                if time.monotonic() >= deadline:
                    raise TimeoutError("Another sorting-list run holds " + str(lock_path))
                time.sleep(0.1)
        try:
            logging.info("LOCK_ACQUIRED waited_seconds=%.3f", time.monotonic() - started)
            yield
        finally:
            if os.name == "nt":
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
            logging.info("LOCK_RELEASED path=%s", lock_path)


def validate_db(db):
    if not isinstance(db, dict):
        raise ValueError("Tracking database must be an object")
    for cow, rec in db.items():
        if not isinstance(cow, str) or not cow.isdecimal() or not isinstance(rec, dict):
            raise ValueError("Invalid cow record in tracking database")
        for field in ("first_seen", "last_seen"):
            stamp = datetime.fromisoformat(rec[field])
            if stamp.tzinfo is not None:
                raise ValueError("Tracking timestamps must use local time without a timezone")
        if not isinstance(rec.get("active"), bool):
            raise ValueError("Invalid active flag for cow " + cow)
        if rec.get("cleared_time") is not None:
            stamp = datetime.fromisoformat(rec["cleared_time"])
            if stamp.tzinfo is not None:
                raise ValueError("Invalid cleared timestamp for cow " + cow)
    return db


def load_db(db_file):
    """Never interpret a missing, corrupt or unreadable database as empty."""
    try:
        with open(db_file, encoding="utf-8-sig") as handle:
            db = validate_db(json.load(handle))
        logging.info("DB_LOADED file=%s records=%s data=%s",
                     json.dumps(file_details(db_file)), len(db), json.dumps(db, sort_keys=True))
        return db
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise RuntimeError(
            f"Cannot safely read {db_file}; stopped without replacing the database. "
            "Check the file or restore a verified backup."
        ) from exc


def save_db(db_file, db):
    """Caller must hold tracking_lock across load, modifications and this save."""
    validate_db(db)
    path = Path(db_file)
    # Validate the existing file before backing it up or replacing it.
    load_db(path)
    previous = path.read_bytes()
    payload = (json.dumps(db, indent=2) + "\n").encode("utf-8")
    backups = path.parent / (path.name + ".backups")
    backups.mkdir(exist_ok=True)
    backup = backups / (datetime.now().strftime("%Y%m%d_%H%M%S_%f") + "_" + uuid4().hex + ".json")
    # A backup failure stops the save. Backups are never automatically pruned.
    with open(backup, "xb") as handle:
        handle.write(previous)
        handle.flush()
        os.fsync(handle.fileno())
    logging.info("DB_BACKUP file=%s", json.dumps(file_details(backup)))
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode="wb", dir=path.parent,
                                         prefix=path.name + ".", suffix=".tmp",
                                         delete=False) as handle:
            temporary = handle.name
            handle.write(payload)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
        logging.info("DB_COMMITTED file=%s records=%s data=%s",
                     json.dumps(file_details(path)), len(db), json.dumps(db, sort_keys=True))
    finally:
        if temporary and os.path.exists(temporary):
            os.unlink(temporary)
