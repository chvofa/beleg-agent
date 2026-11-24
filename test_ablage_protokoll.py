"""Regressionstests: kein Beleg im Ablageordner ohne Protokoll-Zeile.

Hintergrund (entdeckt 24.11.2026): drei Belege lagen als PDF im Ablageordner,
standen aber nie im Protokoll — der Abgleich hielt sie deshalb monatelang fuer
fehlend (Garage Eggimann 1303.60, Kreativ Media 154.80, StreamBoy 10.00).
beleg-agent.log vom 10.06.2026:
    09:33:22 Abgelegt: KreativMedia_Rechnung_875176.pdf -> .../06_Juni/...
    09:33:27 Fehler beim Schreiben ins Excel-Protokoll: File is not a zip file
    09:33:27 Ablage wird trotzdem durchgeführt.
lege_datei_ab verschob die Datei zuerst und ignorierte danach, dass das
Protokoll (OneDrive-Sync, Excel geoeffnet) nicht geschrieben werden konnte.

Run:  .venv/bin/python3 -m unittest test_ablage_protokoll -v
"""
import os
import tempfile
import unittest
from unittest import mock

import openpyxl

import beleg_agent
import config
import offene_posten

DATEN = {
    "rechnungssteller": "Kreativ Media",
    "rechnungsdatum": "2026-06-10",
    "betrag": 154.80,
    "waehrung": "CHF",
    "typ": "Rechnung",
    "zahlungsart": "Überweisung",
    "ist_paypal": False,
    "referenz_nr": "875176",
    "bemerkungen": "",
    "konto_vorschlag": "4430 Unterhalt EDV-Anlage",
    "gesamt_confidence": 0.97,
}

ECHTES_SAVE = openpyxl.Workbook.save


class AblageBasis(unittest.TestCase):

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.stamm = os.path.join(tmp.name, "01-Belege")
        self.inbox = os.path.join(self.stamm, "_Inbox")
        self.zielordner = os.path.join(self.stamm, "2026", config.MONATE[6])
        self.protokoll = os.path.join(self.stamm, "Belege_Protokoll.xlsx")
        os.makedirs(self.inbox)

        self.resolve = mock.Mock(return_value=0)
        patches = [
            mock.patch.object(config, "ABLAGE_STAMMPFAD", self.stamm),
            mock.patch.object(config, "INBOX_PFAD", self.inbox),
            mock.patch.object(config, "EXCEL_PROTOKOLL", self.protokoll),
            mock.patch.object(config, "EXCEL_LOCK", self.protokoll + ".lock"),
            mock.patch.object(beleg_agent, "toast"),
            mock.patch.object(beleg_agent.time, "sleep"),
            mock.patch.object(offene_posten, "resolve_standalone", self.resolve),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)

        beleg_agent.erstelle_excel_wenn_noetig()
        self.beleg = os.path.join(self.inbox, "KreativMedia_Rechnung_875176.pdf")
        with open(self.beleg, "wb") as f:
            f.write(b"%PDF-1.4 test")

    def protokoll_zeilen(self):
        wb = openpyxl.load_workbook(self.protokoll)
        self.addCleanup(wb.close)
        return [r for r in wb.active.iter_rows(min_row=2, values_only=True) if r[0]]

    def abgelegte_dateien(self):
        return os.listdir(self.zielordner) if os.path.isdir(self.zielordner) else []


class ProtokollNichtSchreibbarTest(AblageBasis):
    """Excel dauerhaft gesperrt (in Excel geoeffnet / OneDrive-Sync haengt)."""

    def setUp(self):
        super().setUp()
        p = mock.patch.object(openpyxl.Workbook, "save",
                              side_effect=PermissionError("[Errno 13] Permission denied"))
        p.start()
        self.addCleanup(p.stop)

    def test_ablage_meldet_fehlschlag(self):
        self.assertFalse(beleg_agent.lege_datei_ab(self.beleg, dict(DATEN)))

    def test_beleg_bleibt_nicht_ohne_protokoll_im_ablageordner(self):
        beleg_agent.lege_datei_ab(self.beleg, dict(DATEN))

        self.assertEqual(self.abgelegte_dateien(), [])

    def test_beleg_liegt_zur_pruefung_wieder_in_der_inbox(self):
        beleg_agent.lege_datei_ab(self.beleg, dict(DATEN))

        self.assertEqual(os.listdir(self.inbox), ["[PRÜFEN]_KreativMedia_Rechnung_875176.pdf"])

    def test_offener_posten_wird_nicht_aufgeloest(self):
        beleg_agent.lege_datei_ab(self.beleg, dict(DATEN))

        self.resolve.assert_not_called()


class ProtokollKurzGesperrtTest(AblageBasis):
    """OneDrive gibt die Datei nach dem ersten Fehlversuch wieder frei."""

    def setUp(self):
        super().setUp()
        versuche = {"n": 0}

        def save_erst_beim_zweiten_mal(wb, pfad):
            versuche["n"] += 1
            if versuche["n"] == 1:
                raise PermissionError("[Errno 13] Permission denied")
            return ECHTES_SAVE(wb, pfad)

        p = mock.patch.object(openpyxl.Workbook, "save", save_erst_beim_zweiten_mal)
        p.start()
        self.addCleanup(p.stop)

    def test_zweiter_schreibversuch_legt_den_beleg_normal_ab(self):
        ok = beleg_agent.lege_datei_ab(self.beleg, dict(DATEN))

        self.assertTrue(ok)
        self.assertEqual(self.abgelegte_dateien(), ["Kreativ Media - CHF 154.80.pdf"])
        self.assertEqual(os.listdir(self.inbox), [])

    def test_protokoll_hat_genau_eine_zeile(self):
        beleg_agent.lege_datei_ab(self.beleg, dict(DATEN))

        zeilen = self.protokoll_zeilen()
        self.assertEqual(len(zeilen), 1)
        self.assertEqual(zeilen[0][config.COL_RECHNUNGSSTELLER - 1], "Kreativ Media")


if __name__ == "__main__":
    unittest.main()
