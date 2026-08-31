"""Regressionstests fuer die OneDrive-Sync-Robustheit beim Beleg-Einlesen.

Hintergrund: die _Inbox liegt in einer OneDrive-Bibliothek. Watchdog meldet
eine neue Datei, sobald der Verzeichniseintrag existiert — der Inhalt kommt
u.U. erst Sekunden spaeter, und bei Files-On-Demand liegt er zunaechst gar
nicht lokal (Platzhalter, st_flags=dataless, Groesse bereits final).

Der Agent wartete stur WARTE_NACH_ERKENNUNG_SEKUNDEN=2s und gab pdfium den
*Pfad*. pdfium oeffnet dann selbst (fopen/fread) und sah eine abgeschnittene
bzw. gar nicht vorhandene Datei:
    "Failed to load document (PDFium: Data format error)"
    "Failed to load document (PDFium: File access error)"
Belegt in beleg-agent.log am 28.04., 29.04. (2x), 29.05., 04.06.
und 31.08.2026 — betroffen waren durchweg grosse Scan-PDFs vom Handy.

Run:  .venv/bin/python3 -m unittest test_onedrive_belege -v
"""
import io
import os
import tempfile
import unittest
from unittest import mock

from PIL import Image

import beleg_agent
import config


# ── Test-Doubles fuer pypdfium2 ───────────────────────────────────────────

class FakeBitmap:
    def to_pil(self):
        return Image.new("RGB", (2, 2), (255, 255, 255))


class FakePage:
    def render(self, scale=1):
        return FakeBitmap()


class FakePdf:
    """Mimt pdfium.PdfDocument: laenge, Seitenzugriff, close()."""

    def __init__(self, seiten=1):
        self._seiten = seiten
        self.closed = False

    def __len__(self):
        return self._seiten

    def __getitem__(self, i):
        return FakePage()

    def close(self):
        self.closed = True


class FakeEvent:
    def __init__(self, src_path, is_directory=False):
        self.src_path = src_path
        self.is_directory = is_directory


# ══════════════════════════════════════════════════════════════════════════

class WarteBisDateiStabilTest(unittest.TestCase):
    """Fix 1: warten bis die Datei wirklich fertig geschrieben ist."""

    def test_wartet_bis_groesse_konstant_bleibt(self):
        """Waehrend OneDrive schreibt waechst die Groesse -> erst danach stabil."""
        groessen = [50_000, 180_000, 241_462, 241_462, 241_462, 241_462, 241_462]

        with mock.patch("os.path.getsize", side_effect=groessen), \
             mock.patch("beleg_agent.time.sleep"):
            self.assertTrue(beleg_agent.warte_bis_datei_stabil("/x/scan.pdf"))

    def test_timeout_wenn_datei_endlos_waechst(self):
        """Kein Endlos-Warten: nach DATEI_STABIL_TIMEOUT wird abgebrochen."""
        zaehler = {"n": 0}

        def waechst(_p):
            zaehler["n"] += 1
            return zaehler["n"] * 1000

        uhr = {"t": 0.0}

        def fake_time():
            return uhr["t"]

        def fake_sleep(s):
            uhr["t"] += s

        with mock.patch("os.path.getsize", side_effect=waechst), \
             mock.patch("beleg_agent.time.time", side_effect=fake_time), \
             mock.patch("beleg_agent.time.sleep", side_effect=fake_sleep):
            self.assertFalse(beleg_agent.warte_bis_datei_stabil("/x/scan.pdf"))

    def test_verschwundene_datei_ist_nicht_stabil(self):
        """Datei wurde weggeschoben/geloescht -> False statt Exception."""
        with mock.patch("os.path.getsize", side_effect=OSError(2, "No such file")), \
             mock.patch("beleg_agent.time.sleep"):
            self.assertFalse(beleg_agent.warte_bis_datei_stabil("/x/weg.pdf"))

    def test_nullbyte_datei_gilt_nie_als_stabil(self):
        """OneDrive legt erst einen 0-Byte-Eintrag an — das ist kein fertiger Beleg."""
        groessen = [0, 0, 0, 0, 0, 241_462, 241_462, 241_462, 241_462]

        with mock.patch("os.path.getsize", side_effect=groessen), \
             mock.patch("beleg_agent.time.sleep"):
            self.assertTrue(beleg_agent.warte_bis_datei_stabil("/x/scan.pdf"))


class BilderKonvertierungTest(unittest.TestCase):
    """Fix 2+3: Bytes statt Pfad an pdfium, plus Retry bei Lesefehlern."""

    def setUp(self):
        fd, self.pfad = tempfile.mkstemp(suffix=".pdf")
        os.write(fd, b"%PDF-1.3\nDUMMY")
        os.close(fd)
        self.addCleanup(os.unlink, self.pfad)

    def test_pdfium_bekommt_bytes_nicht_den_pfad(self):
        """Regression: der Python-Read erzwingt die OneDrive-Hydration.

        Gibt man pdfium den Pfad, oeffnet pdfium selbst und bekommt bei einem
        Files-On-Demand-Platzhalter nichts Brauchbares zu sehen.
        """
        with mock.patch.object(beleg_agent.pdfium, "PdfDocument",
                               return_value=FakePdf()) as pdf_mock:
            beleg_agent.datei_zu_bilder_base64(self.pfad)

        (arg,), _ = pdf_mock.call_args
        self.assertIsInstance(arg, bytes)
        self.assertEqual(arg, b"%PDF-1.3\nDUMMY")

    def test_retry_nach_data_format_error(self):
        """Erster Versuch trifft die halb-synchronisierte Datei, der zweite klappt."""
        seq = [RuntimeError("Failed to load document (PDFium: Data format error)"),
               FakePdf()]

        def fake_pdfdocument(_raw):
            item = seq.pop(0)
            if isinstance(item, Exception):
                raise item
            return item

        with mock.patch.object(beleg_agent.pdfium, "PdfDocument",
                               side_effect=fake_pdfdocument), \
             mock.patch("beleg_agent.time.sleep"):
            bilder = beleg_agent.datei_zu_bilder_base64(self.pfad)

        self.assertEqual(len(bilder), 1)
        self.assertEqual(bilder[0][1], "image/png")
        self.assertEqual(seq, [])

    def test_dauerhaft_defekt_reicht_fehler_weiter(self):
        """Ein wirklich kaputtes PDF darf nicht still als 'leer' durchgehen."""
        with mock.patch.object(
            beleg_agent.pdfium, "PdfDocument",
            side_effect=RuntimeError("Failed to load document (PDFium: Data format error)"),
        ), mock.patch("beleg_agent.time.sleep"):
            with self.assertRaises(RuntimeError):
                beleg_agent.datei_zu_bilder_base64(self.pfad)

    def test_bilddatei_wird_ebenfalls_ueber_bytes_gelesen(self):
        """JPG/PNG kommen genauso aus OneDrive -> gleicher Hydration-Pfad."""
        fd, jpg = tempfile.mkstemp(suffix=".png")
        os.close(fd)
        self.addCleanup(os.unlink, jpg)
        Image.new("RGB", (4, 4), (10, 20, 30)).save(jpg, format="PNG")

        bilder = beleg_agent.datei_zu_bilder_base64(jpg)
        self.assertEqual(len(bilder), 1)
        self.assertEqual(bilder[0][1], "image/png")


class VerarbeiteDateiRueckmeldungTest(unittest.TestCase):
    """verarbeite_datei muss sagen, ob es geklappt hat — sonst kein Retry moeglich."""

    def test_false_wenn_extraktion_fehlschlaegt(self):
        with mock.patch.object(beleg_agent, "extrahiere_rechnungsdaten", return_value=None), \
             mock.patch.object(beleg_agent, "toast"):
            self.assertFalse(beleg_agent.verarbeite_datei("/x/scan.pdf"))

    def test_false_bei_claude_api_fehler(self):
        with mock.patch.object(beleg_agent, "extrahiere_rechnungsdaten",
                               side_effect=beleg_agent.ClaudeAPIFehler("429")), \
             mock.patch.object(beleg_agent, "toast"):
            self.assertFalse(beleg_agent.verarbeite_datei("/x/scan.pdf"))

    def test_true_bei_erfolgreicher_ablage(self):
        daten = {"gesamt_confidence": 0.95, "typ": "Rechnung",
                 "rechnungssteller": "SBB", "betrag": 189.3, "waehrung": "CHF",
                 "rechnungsdatum": "2026-08-18", "zahlungsart": "Überweisung"}
        with mock.patch.object(beleg_agent, "extrahiere_rechnungsdaten", return_value=daten), \
             mock.patch.object(beleg_agent, "lege_datei_ab") as ablage:
            self.assertTrue(beleg_agent.verarbeite_datei("/x/scan.pdf"))
        ablage.assert_called_once()

    def test_true_bei_bereits_markierter_datei(self):
        """[PRÜFEN]_/[DUPLIKAT]_ werden uebersprungen — das ist kein Fehlschlag."""
        self.assertTrue(beleg_agent.verarbeite_datei("/x/[PRÜFEN]_scan.pdf"))


class HandlerZweiteChanceTest(unittest.TestCase):
    """Fix 4: ein nachgelieferter OneDrive-Sync bekommt eine zweite Chance.

    Am 28.04. und 04.06. rettete ein zufaellig zweites on_created-Event den
    Beleg; am 31.08. kam es nicht — und die Datei blieb bis zum Neustart liegen.
    """

    def setUp(self):
        self.handler = beleg_agent.BelegHandler()

    def test_on_created_wartet_auf_stabilitaet_statt_fixem_sleep(self):
        with mock.patch.object(beleg_agent, "warte_bis_datei_stabil",
                               return_value=True) as warten, \
             mock.patch("os.path.exists", return_value=True), \
             mock.patch.object(beleg_agent, "verarbeite_datei", return_value=True):
            self.handler.on_created(FakeEvent("/x/scan.pdf"))
        warten.assert_called_once_with("/x/scan.pdf")

    def test_on_modified_wiederholt_fehlgeschlagene_datei(self):
        """Nach Fehlschlag + veraenderter Groesse -> neuer Versuch."""
        with mock.patch.object(beleg_agent, "warte_bis_datei_stabil", return_value=True), \
             mock.patch("os.path.exists", return_value=True), \
             mock.patch.object(beleg_agent, "verarbeite_datei", return_value=False):
            self.handler.on_created(FakeEvent("/x/scan.pdf"))

        with mock.patch.object(beleg_agent, "warte_bis_datei_stabil", return_value=True), \
             mock.patch("os.path.exists", return_value=True), \
             mock.patch("os.path.getsize", return_value=241_462), \
             mock.patch.object(beleg_agent, "verarbeite_datei",
                               return_value=True) as zweiter:
            self.handler.on_modified(FakeEvent("/x/scan.pdf"))

        zweiter.assert_called_once_with("/x/scan.pdf")

    def test_on_modified_ignoriert_unveraenderte_datei(self):
        """Dauerhaft defekter Beleg darf nicht bei jedem Touch neu feuern."""
        with mock.patch.object(beleg_agent, "warte_bis_datei_stabil", return_value=True), \
             mock.patch("os.path.exists", return_value=True), \
             mock.patch("os.path.getsize", return_value=999), \
             mock.patch.object(beleg_agent, "verarbeite_datei", return_value=False):
            self.handler.on_created(FakeEvent("/x/kaputt.pdf"))

        with mock.patch("os.path.exists", return_value=True), \
             mock.patch("os.path.getsize", return_value=999), \
             mock.patch.object(beleg_agent, "verarbeite_datei") as nochmal:
            self.handler.on_modified(FakeEvent("/x/kaputt.pdf"))

        nochmal.assert_not_called()

    def test_on_modified_ignoriert_nie_fehlgeschlagene_datei(self):
        """Ohne vorherigen Fehlschlag ist ein modified-Event irrelevant."""
        with mock.patch("os.path.exists", return_value=True), \
             mock.patch("os.path.getsize", return_value=123), \
             mock.patch.object(beleg_agent, "verarbeite_datei") as nie:
            self.handler.on_modified(FakeEvent("/x/fremd.pdf"))
        nie.assert_not_called()

    def test_on_modified_ignoriert_fremde_endungen(self):
        with mock.patch.object(beleg_agent, "verarbeite_datei") as nie:
            self.handler.on_modified(FakeEvent("/x/Belege_Protokoll.xlsx"))
        nie.assert_not_called()


class ConfigStabilitaetsWerteTest(unittest.TestCase):
    def test_konstanten_vorhanden_und_plausibel(self):
        self.assertGreater(config.DATEI_STABIL_CHECKS, 1)
        self.assertGreater(config.DATEI_STABIL_INTERVALL, 0)
        self.assertGreater(config.DATEI_STABIL_TIMEOUT, config.DATEI_STABIL_INTERVALL)
        self.assertGreaterEqual(config.PDF_LADE_VERSUCHE, 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
