import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime, timedelta
from unittest.mock import patch

import generate_sortinglist as generate
import move_sortinglist as move
import tracking_store as store


def record(hours=10):
    stamp = (datetime.now() - timedelta(hours=hours)).isoformat(timespec="seconds")
    return dict(first_seen=stamp, last_seen=stamp, active=True,
                cleared_reason=None, cleared_time=None)


class TrackingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.db = self.root / "draft_tracking.json"
        self.db.write_text(json.dumps({"2501": record()}))

    def test_corrupt_missing_and_invalid_databases_stop(self):
        for content in ('{', '[]', '{"2501": {}}'):
            self.db.write_text(content)
            with self.assertRaises(RuntimeError):
                store.load_db(self.db)
            with self.assertRaises(RuntimeError):
                store.save_db(self.db, {"2501": record()})
            self.assertEqual(self.db.read_text(), content)
        self.db.unlink()
        with self.assertRaises(RuntimeError):
            store.load_db(self.db)

    def test_backup_preserves_previous_database(self):
        before = self.db.read_bytes()
        with store.tracking_lock(self.db):
            data = store.load_db(self.db)
            data["3419"] = record()
            store.save_db(self.db, data)
        backups = list(self.root.glob("*.backups/*.json"))
        self.assertEqual(len(backups), 1)
        self.assertEqual(backups[0].read_bytes(), before)
        self.assertEqual(set(store.load_db(self.db)), {"2501", "3419"})

    def test_failed_replacement_keeps_original_and_cleans_temp(self):
        before = self.db.read_bytes()
        with patch.object(store.os, "replace", side_effect=PermissionError("busy")):
            with self.assertRaises(PermissionError):
                store.save_db(self.db, {"3419": record()})
        self.assertEqual(self.db.read_bytes(), before)
        self.assertEqual(list(self.root.glob("*.tmp")), [])

    def test_backup_failure_prevents_replacement(self):
        before = self.db.read_bytes()
        with patch.object(store.os, "fsync", side_effect=OSError("disk full")):
            with self.assertRaises(OSError):
                store.save_db(self.db, {"3419": record()})
        self.assertEqual(self.db.read_bytes(), before)

    def test_concurrent_process_updates_are_not_lost(self):
        worker = '''import sys, time
import tracking_store as s
with s.tracking_lock(sys.argv[1]):
    db = s.load_db(sys.argv[1])
    time.sleep(0.1)
    db[sys.argv[2]] = dict(db['2501'])
    s.save_db(sys.argv[1], db)
'''
        processes = [subprocess.Popen([sys.executable, "-c", worker, str(self.db), str(cow)],
                                     cwd=Path(__file__).parent)
                     for cow in range(3100, 3106)]
        for process in processes:
            self.assertEqual(process.wait(timeout=15), 0)
        self.assertEqual(len(store.load_db(self.db)), 7)

    def test_lock_timeout_and_release_after_process_exit(self):
        worker = '''import sys
import tracking_store as s
try:
    with s.tracking_lock(sys.argv[1], timeout=0.2):
        pass
except TimeoutError:
    sys.exit(7)
'''
        with store.tracking_lock(self.db):
            result = subprocess.run([sys.executable, "-c", worker, str(self.db)],
                                    cwd=Path(__file__).parent, timeout=5)
            self.assertEqual(result.returncode, 7)
        with store.tracking_lock(self.db, timeout=0.2):
            pass

    def test_generator_keeps_ten_hour_cows_and_36_hour_rule(self):
        incoming = self.root / "inseminated"
        incoming.mkdir()
        (incoming / "today.csv").write_text("Animal Number\n2747\n2755\n3273\n")
        cows = {str(cow): record() for cow in
                [2501, 3109, 3133, 3383, 3419, 3506, 3564, 2747, 2755, 3273]}
        cows["9999"] = record(37)
        self.db.write_text(json.dumps(cows))
        with patch.object(generate, "DB_FILE", str(self.db)), \
             patch.object(generate, "INSEM_DIR", str(incoming)), \
             patch.object(generate, "create_delpro_xls") as output:
            generate.main()
        updated = store.load_db(self.db)
        self.assertEqual({c for c, r in updated.items() if r["active"]},
                         {"2501", "3109", "3133", "3383", "3419", "3506", "3564"})
        self.assertEqual(updated["9999"]["cleared_reason"], "timeout_36h")
        self.assertEqual(updated["2747"]["cleared_reason"], "inseminated")
        self.assertFalse((incoming / "today.csv").exists())
        self.assertEqual(len(list((incoming / "processed").glob("*.csv"))), 1)
        output.assert_called_once()
        log = next((self.root / "logs").glob("generate_sortinglist_*.log")).read_text()
        for event in ("RUN_START", "LOCK_ACQUIRED", "DB_LOADED", "DB_BACKUP", "DB_COMMITTED",
                      "COW_EVALUATED cow=2501", "reason=timeout_36h", "reason=inseminated",
                      "INSEMINATION_ARCHIVED", "RUN_SUCCESS", "run=", "pid="):
            self.assertIn(event, log)
        self.assertLess(log.index("DB_COMMITTED" + " file="), log.index("INSEMINATION_ARCHIVED"))

    def test_failed_save_does_not_consume_insemination_file_or_write_output(self):
        incoming = self.root / "inseminated"
        incoming.mkdir()
        source = incoming / "today.csv"
        source.write_text("Animal Number\n2501\n")
        with patch.object(generate, "DB_FILE", str(self.db)), \
             patch.object(generate, "INSEM_DIR", str(incoming)), \
             patch.object(generate, "save_db", side_effect=OSError("disk full")), \
             patch.object(generate, "create_delpro_xls") as output:
            with self.assertRaises(OSError):
                generate.main()
        self.assertTrue(source.exists())
        self.assertTrue(store.load_db(self.db)["2501"]["active"])
        output.assert_not_called()

    def test_move_preserves_existing_cows_and_timer(self):
        source = self.root / "incoming.xls"
        source.write_bytes(b"test fixture")
        dest = self.root / "output" / "sortinglist.xls"
        before = store.load_db(self.db)["2501"]["first_seen"]
        with patch.object(move, "DB_FILE", str(self.db)), \
             patch.object(move, "SOURCE_FILE", str(source)), \
             patch.object(move, "DEST_FILE", str(dest)), \
             patch.object(move, "read_cows_from_xls", return_value=["3419", "2501"]):
            move.main()
        updated = store.load_db(self.db)
        self.assertEqual(set(updated), {"2501", "3419"})
        self.assertEqual(updated["2501"]["first_seen"], before)
        self.assertTrue(dest.exists())
        log = next((self.root / "logs").glob("move_sortinglist_*.log")).read_text()
        for event in ("COW_ADDED cow=3419", "COW_SEEN_AGAIN cow=2501", "INPUT_MOVED", "RUN_SUCCESS"):
            self.assertIn(event, log)

    def test_returning_inactive_cow_gets_new_timer_and_cleared_removal_fields(self):
        returned_at = datetime(2026, 11, 12, 17, 28, 38)
        source = self.root / "incoming.xls"
        dest = self.root / "output" / "sortinglist.xls"
        for reason in ("timeout_36h", "inseminated"):
            with self.subTest(previous_reason=reason):
                old = dict(first_seen="2026-11-03T13:38:38", last_seen="2026-11-03T13:38:38",
                           active=False, cleared_reason=reason, cleared_time="2026-11-05T03:00:31")
                self.db.write_text(json.dumps({"3106": old, "2501": record()}))
                source.write_bytes(b"test fixture")
                with patch.object(move, "DB_FILE", str(self.db)), \
                     patch.object(move, "SOURCE_FILE", str(source)), \
                     patch.object(move, "DEST_FILE", str(dest)), \
                     patch.object(move, "datetime") as clock, \
                     patch.object(move, "read_cows_from_xls", return_value=["3106", "3106"]):
                    clock.now.return_value = returned_at
                    move.main()
                updated = store.load_db(self.db)["3106"]
                self.assertEqual(updated, dict(first_seen=returned_at.isoformat(),
                                 last_seen=returned_at.isoformat(), active=True,
                                 cleared_reason=None, cleared_time=None))
                backups = [json.loads(p.read_text())["3106"]
                           for p in self.root.glob("*.backups/*.json")]
                self.assertIn(old, backups)
        log = next((self.root / "logs").glob("move_sortinglist_*.log")).read_text()
        self.assertIn("COW_REACTIVATION_PENDING cow=3106", log)
        self.assertIn("new_first_seen=2026-11-12T17:28:38", log)

    def test_3106_survives_next_generator_and_expires_after_new_36_hours(self):
        source = self.root / "incoming.xls"
        source.write_bytes(b"test fixture")
        self.db.write_text(json.dumps({"3106": dict(
            first_seen="2026-11-03T13:38:38", last_seen="2026-11-03T13:38:38",
            active=False, cleared_reason="timeout_36h", cleared_time="2026-11-05T03:00:31")}))
        with patch.object(move, "DB_FILE", str(self.db)), \
             patch.object(move, "SOURCE_FILE", str(source)), \
             patch.object(move, "DEST_FILE", str(self.root / "output" / "sortinglist.xls")), \
             patch.object(move, "datetime") as clock, \
             patch.object(move, "read_cows_from_xls", return_value=["3106"]):
            clock.now.return_value = datetime(2026, 11, 12, 17, 28, 38)
            move.main()
            # Seeing the active cow again must not extend the new timer.
            source.write_bytes(b"second fixture")
            clock.now.return_value = datetime(2026, 11, 13, 1, 0, 0)
            move.main()
        self.assertEqual(store.load_db(self.db)["3106"]["first_seen"], "2026-11-12T17:28:38")
        with patch.object(generate, "DB_FILE", str(self.db)), \
             patch.object(generate, "load_inseminated_cows", return_value=set()), \
             patch.object(generate, "create_delpro_xls") as output, \
             patch.object(generate, "datetime") as clock:
            clock.fromisoformat.side_effect = datetime.fromisoformat
            for when in (datetime(2026, 11, 13, 3), datetime(2026, 11, 14, 5, 28, 38)):
                clock.now.return_value = when
                generate.main()
                self.assertTrue(store.load_db(self.db)["3106"]["active"])
                self.assertTrue(output.call_args.args[0]["3106"]["active"])
            clock.now.return_value = datetime(2026, 11, 14, 5, 28, 39)
            generate.main()
        updated = store.load_db(self.db)["3106"]
        self.assertFalse(updated["active"])
        self.assertEqual(updated["cleared_reason"], "timeout_36h")

    def test_corrupt_database_stops_both_scripts_before_side_effects(self):
        self.db.write_text('{"broken":')
        source = self.root / "incoming.xls"
        source.write_bytes(b"test fixture")
        for module in (move, generate):
            with patch.object(module, "DB_FILE", str(self.db)), \
                 patch.object(move, "SOURCE_FILE", str(source)), \
                 patch.object(move, "read_cows_from_xls") as read_xls, \
                 patch.object(generate, "load_inseminated_cows") as read_csv, \
                 patch.object(generate, "create_delpro_xls") as output:
                with self.assertRaises(RuntimeError):
                    module.main()
                read_xls.assert_not_called()
                read_csv.assert_not_called()
                output.assert_not_called()
        self.assertEqual(self.db.read_text(), '{"broken":')
        self.assertTrue(source.exists())
        logs = list((self.root / "logs").glob("*.log"))
        self.assertEqual(len(logs), 2)
        for path in logs:
            log = path.read_text()
            self.assertIn("RUN_FAILED", log)
            self.assertIn("Traceback", log)
            self.assertIn("JSONDecodeError", log)
            self.assertNotIn("RUN_SUCCESS", log)

    def test_idle_runs_share_daily_log_and_keep_distinct_run_ids(self):
        with patch.object(move, "DB_FILE", str(self.db)), \
             patch.object(move, "SOURCE_FILE", str(self.root / "absent.xls")):
            move.main()
            move.main()
        logs = list((self.root / "logs").glob("move_sortinglist_*.log"))
        self.assertEqual(len(logs), 1)
        for path in logs:
            log = path.read_text()
            self.assertEqual(log.count("NO_INPUT_FILE"), 2)
            self.assertEqual(log.count("RUN_SUCCESS"), 2)
            starts = [line for line in log.splitlines() if "RUN_START" in line]
            self.assertEqual(len({line.split("run=")[1].split()[0] for line in starts}), 2)

    def test_sms_failure_does_not_log_credentials_or_recipient(self):
        from types import ModuleType
        from unittest.mock import Mock
        twilio = ModuleType("twilio")
        rest = ModuleType("twilio.rest")
        rest.Client = Mock()
        rest.Client.return_value.messages.create.side_effect = RuntimeError("secret-token private-phone")
        config = dict(sms_enabled=True, twilio_sid="secret-sid", twilio_token="secret-token",
                      twilio_from="private-from", recipients=["private-phone"])
        with store.logged_run(self.db, "generate_sortinglist"), \
             patch.dict(sys.modules, {"twilio": twilio, "twilio.rest": rest}), \
             patch.object(generate, "load_sms_config", return_value=config):
            generate.send_sms_list([2501])
        log = next((self.root / "logs").glob("*.log")).read_text()
        self.assertIn("SMS_FAILED", log)
        for private in ("secret-token", "secret-sid", "private-phone", "private-from"):
            self.assertNotIn(private, log)

    def test_concurrent_runs_append_complete_lines_to_one_daily_log(self):
        worker = '''import logging, sys
import tracking_store as s
with s.logged_run(sys.argv[1], 'move_sortinglist'):
    for number in range(5):
        logging.info('TEST_MARKER worker=%s row=%s payload=%s', sys.argv[2], number, 'x' * 10000)
'''
        processes = [subprocess.Popen([sys.executable, "-c", worker, str(self.db), str(i)],
                                     cwd=Path(__file__).parent) for i in range(4)]
        for process in processes:
            self.assertEqual(process.wait(timeout=15), 0)
        logs = list((self.root / "logs").glob("move_sortinglist_*.log"))
        self.assertEqual(len(logs), 1)
        lines = logs[0].read_text().splitlines()
        self.assertEqual(sum("RUN_SUCCESS" in line for line in lines), 4)
        markers = [line for line in lines if "TEST_MARKER" in line]
        self.assertEqual(len(markers), 20)
        for line in markers:
            self.assertEqual(line.count("TEST_MARKER"), 1)
            self.assertTrue(line.endswith("x" * 10000))


if __name__ == "__main__":
    unittest.main()
