"""Regressionstests fuer das Aufloesen offener Posten beim Einlesen eines Belegs.

Hintergrund (Vorfall 24.11.2026): Beim Nachtragen von PayPal-Belegen loeschte
EIN neuer Beleg gleich mehrere offene Posten:
    Sky 15.90 vom 04.09.  -> "Offene Posten aufgeloest: 3" (03.09., 03.10., 03.11.)
    Sky 12.90 vom 24.09.  -> 3 Posten, darunter der 23.08., fuer den es gar
                             keinen Beleg gibt
    Streamboy 9.73        -> 2 Posten (9.73 und 9.64, beide in der Toleranz)
resolve() sammelte alle Zeilen mit gleichem Betrag (+-0.10) im Fenster von
+-120 Tagen ein. Bei monatlichen Abos mit gleichem Betrag verschwanden so
Posten ohne Beleg aus der Liste.

Zweiter Fehler am selben Tag: der TechSmith-Beleg (19.72 CHF) loeste seinen
Posten nicht auf, weil der Kartentext den Namen kuerzt ("PP*PP* FS TECHSMIT")
und der Namens-Check nur ganze Woerter als Teilstring sucht.

Run:  .venv/bin/python3 -m unittest test_offene_posten_resolve -v
"""
import os
import tempfile
import unittest
from unittest import mock

import openpyxl

import config
import offene_posten


def sheet_mit(*posten):
    """Offene_Posten-Sheet im Speicher: posten = (quelle, datum, betrag, waehrung, text)."""
    wb = openpyxl.Workbook()
    ws = offene_posten.ensure_sheet(wb)
    for quelle, datum, betrag, waehrung, text in posten:
        ws.append([quelle, datum, betrag, waehrung, text, "Offen", "", "2026-11-23 19:00", "", ""])
    return ws


def offene_daten(ws):
    return [(p["datum"], p["betrag"]) for p in offene_posten.list_offen(ws)]


SKY = "PAYPAL *SKY F21843022 4029357733 CHE"


class NurBesterTrefferTest(unittest.TestCase):

    def test_ein_beleg_loest_nur_den_zeitlich_naechsten_posten(self):
        ws = sheet_mit(
            ("KK CHF", "2026-09-03", 15.90, "CHF", SKY),
            ("KK CHF", "2026-10-03", 15.90, "CHF", SKY),
            ("KK CHF", "2026-11-03", 15.90, "CHF", SKY),
        )

        n = offene_posten.resolve(ws, "2026-10-04", 15.90, "CHF",
                                  rechnungssteller="Sky Switzerland",
                                  nur_bester_treffer=True)

        self.assertEqual(n, 1)
        self.assertEqual(offene_daten(ws), [("2026-09-03", 15.90), ("2026-11-03", 15.90)])

    def test_gleicher_tag_entscheidet_der_genauere_betrag(self):
        """OpenAI bucht am selben Tag 5.81 und 5.83 — beide liegen in der Toleranz."""
        text = "OPENAI *API OPENAI.COM USACA"
        ws = sheet_mit(
            ("KK CHF", "2026-04-12", 5.81, "USD", text),
            ("KK CHF", "2026-04-12", 5.83, "USD", text),
        )

        n = offene_posten.resolve(ws, "2026-04-12", 5.83, "USD",
                                  rechnungssteller="OpenAI", nur_bester_treffer=True)

        self.assertEqual(n, 1)
        self.assertEqual(offene_daten(ws), [("2026-04-12", 5.81)])

    def test_abgleich_cleanup_ohne_flag_raeumt_weiterhin_alle_treffer_ab(self):
        """Die Abgleich-Skripte rufen resolve() ohne Flag — ihr Verhalten bleibt."""
        ws = sheet_mit(
            ("KK CHF", "2026-09-03", 15.90, "CHF", SKY),
            ("KK CHF", "2026-10-03", 15.90, "CHF", SKY),
        )

        n = offene_posten.resolve(ws, "2026-09-04", 15.90, "CHF",
                                  rechnungssteller="Sky Switzerland")

        self.assertEqual(n, 2)
        self.assertEqual(offene_daten(ws), [])


class EinlesenLoestNurEinenPostenTest(unittest.TestCase):
    """resolve_standalone ist der Pfad aus beleg_agent.lege_datei_ab."""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.pfad = os.path.join(tmp.name, "Belege_Protokoll.xlsx")
        wb = openpyxl.Workbook()
        ws = offene_posten.ensure_sheet(wb)
        for datum in ("2026-08-23", "2026-09-23", "2026-10-23"):
            ws.append(["KK CHF", datum, 12.90, "CHF", SKY, "Offen", "", "2026-11-23 19:00", "", ""])
        wb.save(self.pfad)
        wb.close()
        for name, wert in (("EXCEL_PROTOKOLL", self.pfad), ("EXCEL_LOCK", self.pfad + ".lock")):
            p = mock.patch.object(config, name, wert)
            p.start()
            self.addCleanup(p.stop)

    def test_neuer_beleg_laesst_die_anderen_monate_offen(self):
        n = offene_posten.resolve_standalone(datum="2026-09-24", betrag=12.90, waehrung="CHF",
                                             rechnungssteller="Sky Switzerland")

        self.assertEqual(n, 1)
        wb = openpyxl.load_workbook(self.pfad)
        self.addCleanup(wb.close)
        self.assertEqual(offene_daten(wb[offene_posten.SHEET_NAME]),
                         [("2026-08-23", 12.90), ("2026-10-23", 12.90)])


class GekuerzterKartentextTest(unittest.TestCase):

    def test_abgeschnittener_haendlername_gilt_als_namenstreffer(self):
        ws = sheet_mit(("KK CHF", "2026-11-20", 19.72, "CHF", "PP*PP* FS TECHSMIT 4029357733 USA"))

        n = offene_posten.resolve(ws, "2026-11-20", 19.72, "CHF", rechnungssteller="TechSmith")

        self.assertEqual(n, 1)

    def test_dauerauftrag_pfad_erkennt_abgeschnittenen_namen_ebenfalls(self):
        ws = sheet_mit(("KK CHF", "2026-11-20", 19.72, "CHF", "PP*PP* FS TECHSMIT 4029357733 USA"))

        n = offene_posten.resolve_by_name(ws, "TechSmith", 19.72, "CHF")

        self.assertEqual(n, 1)

    def test_kurzer_wortanfang_ist_kein_namenstreffer(self):
        """'MEDIA' ist nur der Anfang von 'MEDIAMARKT' — kein gekuerzter Kartentext."""
        ws = sheet_mit(("Bank CHF", "2026-06-30", 24.00, "CHF", "Kreativ Media GmbH;8008 Zürich"))

        n = offene_posten.resolve(ws, "2026-06-30", 24.00, "CHF", rechnungssteller="MediaMarkt")

        self.assertEqual(n, 0)
        self.assertEqual(len(offene_daten(ws)), 1)

    def test_drei_buchstaben_reichen_nicht_als_gekuerzter_name(self):
        ws = sheet_mit(("KK CHF", "2026-09-03", 15.90, "CHF", SKY))

        n = offene_posten.resolve(ws, "2026-09-03", 15.90, "CHF", rechnungssteller="Skype")

        self.assertEqual(n, 0)


if __name__ == "__main__":
    unittest.main()
