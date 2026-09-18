"""Regressionstests fuer die Claude-Extraktion bei Sammelbelegen.

Hintergrund (Vorfall INV_DCH00170525, 14.09.2026): die MB.CHARGE-
Monatsrechnung listet 8 Ladevorgaenge auf. Der Prompt definiert das als
Sammelbeleg, also schrieb Claude alle 8 als einzelposten aus — jeden mit
langer Beschreibung und 30-stelliger Stations-ID. Bei max_tokens=1024 brach
die Antwort mitten im 8. Posten ab (stop_reason=max_tokens), das JSON war
unvollstaendig, und der Agent meldete irrefuehrend "Antwort unlesbar (Bild/JSON)".
Der Fehler war deterministisch: der Beleg blieb dauerhaft in der _Inbox.

Dabei werden einzelposten nur bei Parkingpay gebraucht
(_SAMMELBELEG_SPLIT_NAMEN) — alle anderen Anbieter werden als EINE
Monatsrechnung mit Gesamtbetrag abgelegt. Die Liste war also reiner Ballast.

Run:  .venv/bin/python3 -m unittest test_extraktion -v
"""
import json
import types
import unittest
from unittest import mock

import beleg_agent
import config


def fake_response(text, stop_reason="end_turn", output_tokens=250):
    """Mimt anthropic.types.Message fuer die Felder, die der Agent liest."""
    return types.SimpleNamespace(
        content=[types.SimpleNamespace(text=text)],
        stop_reason=stop_reason,
        usage=types.SimpleNamespace(output_tokens=output_tokens),
    )


GUELTIGES_JSON = json.dumps({
    "rechnungssteller": "MB.CHARGE",
    "rechnungsdatum": "2026-08-31",
    "betrag": 99.81,
    "waehrung": "CHF",
    "typ": "Rechnung",
    "zahlungsart": "",
    "ist_paypal": True,
    "referenz_nr": "DCH00170525",
    "ist_sammelbeleg": True,
    "einzelposten": [],
    "bemerkungen": "",
    "confidence": {"rechnungssteller": 0.95, "rechnungsdatum": 0.95,
                   "betrag": 0.98, "waehrung": 0.99, "typ": 0.95,
                   "zahlungsart": 0.0},
})

# So sah die echte Antwort aus: Codeblock offen, JSON mitten im String zu Ende.
ABGESCHNITTEN = (
    '```json\n{\n  "rechnungssteller": "MB.CHARGE (Digital Charging Solutions)",\n'
    '  "einzelposten": [\n    {\n      "datum": "2026-08-28",\n'
    '      "beschreibung": "Ladevorgang Kantonsstrasse 43A, Visp – 0.000'
)


class ExtraktionsTestBasis(unittest.TestCase):
    """Isoliert extrahiere_rechnungsdaten von Datei, API und Excel-Protokoll."""

    def setUp(self):
        patches = [
            mock.patch.object(beleg_agent, "datei_zu_bilder_base64",
                              return_value=[("AAAA", "image/png")]),
            mock.patch.object(beleg_agent, "lookup_zahlungsart",
                              return_value=("", None)),
            mock.patch.object(beleg_agent.kontierung, "vorschlag_fuer",
                              return_value=""),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)

    def extrahiere(self, antwort):
        with mock.patch.object(beleg_agent.client.messages, "create",
                               return_value=antwort) as create:
            ergebnis = beleg_agent.extrahiere_rechnungsdaten("/x/INV_DCH00170525.pdf")
        self.create_kwargs = create.call_args.kwargs
        return ergebnis


class MaxTokensTest(ExtraktionsTestBasis):

    def test_api_bekommt_max_tokens_aus_config(self):
        self.extrahiere(fake_response(GUELTIGES_JSON))
        self.assertEqual(self.create_kwargs["max_tokens"], config.ANTHROPIC_MAX_TOKENS)

    def test_max_tokens_reicht_fuer_sammelbelege(self):
        """1024 war zu knapp fuer 8 Posten — Parkingpay kann deutlich mehr haben."""
        self.assertGreaterEqual(config.ANTHROPIC_MAX_TOKENS, 4096)

    def test_abgeschnittene_antwort_meldet_echte_ursache(self):
        """Regression: nicht 'unlesbar (Bild/JSON)', sondern 'abgeschnitten'."""
        with self.assertRaises(beleg_agent.ClaudeAPIFehler) as ctx:
            self.extrahiere(fake_response(ABGESCHNITTEN, stop_reason="max_tokens",
                                          output_tokens=config.ANTHROPIC_MAX_TOKENS))
        self.assertIn("abgeschnitten", ctx.exception.kurz)
        self.assertIn("ANTHROPIC_MAX_TOKENS", ctx.exception.kurz)

    def test_vollstaendige_antwort_im_codeblock_wird_geparst(self):
        daten = self.extrahiere(fake_response(f"```json\n{GUELTIGES_JSON}\n```"))
        self.assertEqual(daten["betrag"], 99.81)
        self.assertEqual(daten["referenz_nr"], "DCH00170525")


class EinzelpostenNurFuerSplitAnbieterTest(ExtraktionsTestBasis):
    """Der Prompt fordert einzelposten nur dort an, wo sie auch verbucht werden."""

    def prompt_text(self):
        self.extrahiere(fake_response(GUELTIGES_JSON))
        bloecke = self.create_kwargs["messages"][0]["content"]
        return next(b["text"] for b in bloecke if b["type"] == "text")

    def test_prompt_nennt_parkingpay_als_einzigen_split_anbieter(self):
        prompt = self.prompt_text()
        regel = prompt[prompt.index("- einzelposten:"):]
        self.assertIn("Parkingpay", regel)
        self.assertIn("leere Liste []", regel)

    def test_prompt_folgt_dem_split_set_statt_hartkodiert(self):
        """Wird ein Anbieter zum Split-Set hinzugefuegt, zieht der Prompt nach."""
        with mock.patch.object(beleg_agent, "_SAMMELBELEG_SPLIT_NAMEN",
                               {"parkingpay", "testanbieter"}):
            prompt = self.prompt_text()
        regel = prompt[prompt.index("- einzelposten:"):]
        self.assertIn("Testanbieter", regel)
        self.assertIn("Parkingpay", regel)


class VerarbeiteDateiAbgeschnittenTest(unittest.TestCase):
    def test_toast_nennt_abschneiden_statt_bild_json(self):
        fehler = beleg_agent.ClaudeAPIFehler(
            "Antwort abgeschnitten (max_tokens=8192) – ANTHROPIC_MAX_TOKENS in config.py erhöhen?")
        with mock.patch.object(beleg_agent, "extrahiere_rechnungsdaten", side_effect=fehler), \
             mock.patch.object(beleg_agent, "toast") as toast:
            self.assertFalse(beleg_agent.verarbeite_datei("/x/INV_DCH00170525.pdf"))
        text = toast.call_args.args[1]
        self.assertIn("abgeschnitten", text)
        self.assertNotIn("Bild/JSON", text)


if __name__ == "__main__":
    unittest.main(verbosity=2)
