"""Regressionstests fuer die Duplikat-Erkennung bei gesplitteten Sammelbelegen.

Hintergrund (Vorfall Parkingpay-Kontoauszug April 2026): derselbe Auszug
(monthly-report-2026-4.pdf, Periode 01.-30.04.2026) wurde am 03.05. und am
30.05.2026 je einmal aus dem Parkingpay-Portal heruntergeladen und in die
_Inbox gelegt. Beide Male wurde er abgelegt und in 4 Einzelposten gesplittet —
das Protokoll enthielt die 4 April-Parkvorgaenge danach doppelt (CHF 19.42 zu
viel), die zweite Datei landete als "ParkingPay - CHF 0.00 KK CHF (2).pdf".

Ursache: pruefe_duplikat_protokoll() verglich den Kopf des Belegs
(Gesamtbetrag + Belegdatum 2026-05-01) mit dem Protokoll. Bei Parkingpay
schreibt schreibe_protokoll() aber nie eine Kopfzeile, sondern nur die
Einzelposten mit eigenem Datum und Betrag. Der Kopf wird also nie gefunden,
egal welcher Betrag extrahiert wurde (0.00 aus "Total Bewegungen" — aber
auch korrekte 19.42 haetten nicht gematcht). Parkingpay vs. ParkingPay war
nicht die Ursache: der Vergleich laeuft lowercase.

Run:  .venv/bin/python3 -m unittest test_duplikat_sammelbeleg -v
"""
import os
import shutil
import sys
import tempfile
import types
import unittest
from unittest import mock

import openpyxl

import beleg_agent
import config
import offene_posten


POSTEN_APRIL = [
    {"datum": "2026-04-01", "betrag": 7.42, "beschreibung": "Metro Bern", "referenz_nr": "BE474994"},
    {"datum": "2026-04-10", "betrag": 4.50, "beschreibung": "Parkhaus Anna Seiler Haus, Bern", "referenz_nr": "BE39139"},
    {"datum": "2026-04-10", "betrag": 2.50, "beschreibung": "Parkhaus Anna Seiler Haus, Bern", "referenz_nr": "BE39139"},
    {"datum": "2026-04-17", "betrag": 5.00, "beschreibung": "Parkhaus Anna Seiler Haus, Bern", "referenz_nr": "BE474994"},
]


def parkingpay_auszug(rechnungssteller="Parkingpay", betrag=0.0, posten=POSTEN_APRIL,
                      datum="2026-05-01"):
    """So kam die Extraktion am 03.05. bzw. 30.05.2026 zurueck."""
    return {
        "rechnungssteller": rechnungssteller,
        "rechnungsdatum": datum,
        "betrag": betrag,
        "waehrung": "CHF",
        "typ": "Rechnung",
        "zahlungsart": "KK CHF",
        "ist_paypal": False,
        "referenz_nr": "",
        "ist_sammelbeleg": True,
        "einzelposten": [dict(p) for p in posten],
        "bemerkungen": "",
        "gesamt_confidence": 0.956,
        "konto_vorschlag": "",
    }


def einzelrechnung(betrag=49.90, datum="2026-05-12"):
    return {
        "rechnungssteller": "Swisscom",
        "rechnungsdatum": datum,
        "betrag": betrag,
        "waehrung": "CHF",
        "typ": "Rechnung",
        "zahlungsart": "",
        "ist_paypal": False,
        "referenz_nr": "",
        "ist_sammelbeleg": False,
        "einzelposten": [],
        "bemerkungen": "",
        "gesamt_confidence": 0.95,
        "konto_vorschlag": "",
    }


class ProtokollImTempTest(unittest.TestCase):
    """Ablage und Protokoll in einem Temp-Verzeichnis — nie das echte OneDrive."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self.inbox = os.path.join(self.tmp, "_Inbox")
        os.makedirs(self.inbox)
        protokoll = os.path.join(self.tmp, "Belege_Protokoll.xlsx")

        patches = [
            mock.patch.object(config, "ABLAGE_STAMMPFAD", self.tmp),
            mock.patch.object(config, "INBOX_PFAD", self.inbox),
            mock.patch.object(config, "EXCEL_PROTOKOLL", protokoll),
            mock.patch.object(config, "EXCEL_LOCK", protokoll + ".lock"),
            mock.patch.object(beleg_agent, "toast"),
            mock.patch.object(beleg_agent.kontierung, "vorschlag_fuer", return_value=""),
            mock.patch.object(offene_posten, "resolve_standalone", return_value=0),
            # web_app nicht importieren (startet Flask/Agent) — nur event_bus wird gebraucht
            mock.patch.dict(sys.modules, {"web_app": types.SimpleNamespace(event_bus=mock.Mock())}),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        self.assertTrue(config.EXCEL_PROTOKOLL.startswith(self.tmp))

    def lege_in_inbox(self, name="monthly-report-2026-4.pdf"):
        pfad = os.path.join(self.inbox, name)
        with open(pfad, "wb") as f:
            f.write(b"%PDF-1.4 Parkingpay Kontoauszug April 2026")
        return pfad

    def protokoll_zeilen(self, steller="parkingpay"):
        wb = openpyxl.load_workbook(config.EXCEL_PROTOKOLL, read_only=True)
        zeilen = [r for r in wb.active.iter_rows(min_row=2, values_only=True)
                  if r[1] and steller in str(r[1]).lower()]
        wb.close()
        return zeilen


class ParkingpayDoppeltEingelegtTest(ProtokollImTempTest):

    def test_zweiter_download_desselben_auszugs_wird_als_duplikat_erkannt(self):
        """Nachstellung 03.05. / 30.05.2026: zweiter Durchlauf darf nichts ablegen."""
        erster = self.lege_in_inbox()
        self.assertTrue(beleg_agent.lege_datei_ab(erster, parkingpay_auszug("Parkingpay")))
        self.assertEqual(len(self.protokoll_zeilen()), 4)

        zweiter = self.lege_in_inbox()
        daten = parkingpay_auszug("ParkingPay")
        daten["gesamt_confidence"] = 0.91
        self.assertFalse(beleg_agent.lege_datei_ab(zweiter, daten))

        self.assertEqual(len(self.protokoll_zeilen()), 4, "April-Posten doppelt im Protokoll")
        self.assertEqual(os.listdir(self.inbox), ["[DUPLIKAT]_monthly-report-2026-4.pdf"])
        self.assertEqual(os.listdir(os.path.join(self.tmp, "2026", "05_Mai")),
                         ["Parkingpay - CHF 0.00 KK CHF.pdf"])

    def test_erkennung_haengt_nicht_am_extrahierten_gesamtbetrag(self):
        """Auch mit korrekt extrahierten 19.42 muss der Zweitlauf auffallen."""
        beleg_agent.schreibe_protokoll(parkingpay_auszug(betrag=0.0), "monthly-report-2026-4.pdf", "x")
        self.assertTrue(beleg_agent.pruefe_duplikat_protokoll(parkingpay_auszug(betrag=19.42)))

    def test_folgemonat_mit_aehnlichen_posten_ist_kein_duplikat(self):
        """Gleiches Parkhaus, gleicher Betrag, anderer Tag -> neuer Auszug, kein Duplikat."""
        beleg_agent.schreibe_protokoll(parkingpay_auszug(), "monthly-report-2026-4.pdf", "x")
        mai = parkingpay_auszug(datum="2026-06-01", posten=[
            {"datum": "2026-05-08", "betrag": 2.50, "beschreibung": "Parkhaus Anna Seiler Haus, Bern"},
            {"datum": "2026-05-21", "betrag": 7.42, "beschreibung": "Metro Bern"},
        ])
        self.assertFalse(beleg_agent.pruefe_duplikat_protokoll(mai))

    def test_nur_teilweise_vorhandene_posten_sind_kein_duplikat(self):
        """Duplikat nur, wenn JEDER Posten eine eigene Protokollzeile hat.

        Zwei gleiche Parkvorgaenge am selben Tag brauchen zwei Zeilen — eine
        einzelne vorhandene Zeile deckt nicht beide ab.
        """
        beleg_agent.schreibe_protokoll(
            parkingpay_auszug(posten=[POSTEN_APRIL[2]]), "monthly-report-2026-4.pdf", "x")
        zweimal_gleich = parkingpay_auszug(posten=[POSTEN_APRIL[2], POSTEN_APRIL[2]])
        self.assertFalse(beleg_agent.pruefe_duplikat_protokoll(zweimal_gleich))


class EinzelrechnungUnveraendertTest(ProtokollImTempTest):
    """Normale Belege werden weiter ueber ihre Kopfdaten verglichen."""

    def test_gleiche_kopfdaten_sind_duplikat(self):
        beleg_agent.schreibe_protokoll(einzelrechnung(), "rechnung.pdf", "x")
        daten = einzelrechnung()
        daten["rechnungssteller"] = "SWISSCOM"
        self.assertTrue(beleg_agent.pruefe_duplikat_protokoll(daten))

    def test_anderer_betrag_ist_kein_duplikat(self):
        beleg_agent.schreibe_protokoll(einzelrechnung(), "rechnung.pdf", "x")
        self.assertFalse(beleg_agent.pruefe_duplikat_protokoll(einzelrechnung(betrag=59.90)))


if __name__ == "__main__":
    unittest.main(verbosity=2)
