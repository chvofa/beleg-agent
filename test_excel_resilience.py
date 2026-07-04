"""Regressionstests fuer die OneDrive-Sync-Robustheit des Excel-Zugriffs.

Hintergrund: Belege_Protokoll.xlsx liegt in einer OneDrive-Bibliothek. Waehrend
OneDrive die Datei synchronisiert, schlaegt openpyxl.load_workbook sporadisch
fehl mit zipfile.BadZipFile ("File is not a zip file") oder OSError
([Errno 11] "Resource deadlock avoided"). Diese Fehler sind transient.

Run:  .venv/bin/python3 -m unittest test_excel_resilience -v
"""
import unittest
import zipfile
from unittest import mock

import config


class LoadProtokollRetryTest(unittest.TestCase):
    def test_retries_on_badzipfile_then_succeeds(self):
        """2x BadZipFile (Datei mid-sync), dann kohaerent -> Workbook kommt zurueck."""
        sentinel = object()
        calls = {"n": 0}

        def fake_load(path, read_only=False):
            calls["n"] += 1
            if calls["n"] < 3:
                raise zipfile.BadZipFile("File is not a zip file")
            return sentinel

        with mock.patch("openpyxl.load_workbook", side_effect=fake_load), \
             mock.patch("config.time.sleep"):
            result = config.load_protokoll()

        self.assertIs(result, sentinel)
        self.assertEqual(calls["n"], 3)

    def test_retries_on_oserror_deadlock_then_succeeds(self):
        """OSError [Errno 11] Resource deadlock avoided ist transient -> Retry."""
        sentinel = object()
        seq = [OSError(11, "Resource deadlock avoided"), sentinel]

        def fake_load(path, read_only=False):
            item = seq.pop(0)
            if isinstance(item, Exception):
                raise item
            return item

        with mock.patch("openpyxl.load_workbook", side_effect=fake_load), \
             mock.patch("config.time.sleep"):
            result = config.load_protokoll()

        self.assertIs(result, sentinel)

    def test_reraises_after_exhausting_retries(self):
        """Permanent defekt -> letzter Fehler wird weitergereicht (kein stilles None)."""
        def fake_load(path, read_only=False):
            raise zipfile.BadZipFile("File is not a zip file")

        with mock.patch("openpyxl.load_workbook", side_effect=fake_load), \
             mock.patch("config.time.sleep"):
            with self.assertRaises(zipfile.BadZipFile):
                config.load_protokoll(_versuche=3)


class MigrationSelfHealTest(unittest.TestCase):
    def test_migration_survives_persistent_onedrive_failure(self):
        """Regression: ein OneDrive-Lesefehler darf den Agenten-Start NICHT killen.

        Frueher schlug BadZipFile aus migriere_excel_spalten() bis
        web_app.start() durch -> Worker still tot ("haengt").
        """
        import beleg_agent

        @mock.patch("config.excel_lock")
        @mock.patch("os.path.exists", return_value=True)
        @mock.patch.object(config, "load_protokoll",
                           side_effect=zipfile.BadZipFile("File is not a zip file"))
        def run(_load, _exists, _lock):
            # no-op Context-Manager fuer excel_lock
            _lock.return_value.__enter__ = lambda *a: None
            _lock.return_value.__exit__ = lambda *a: False
            beleg_agent.migriere_excel_spalten()  # darf NICHT raisen

        run()


if __name__ == "__main__":
    unittest.main(verbosity=2)
