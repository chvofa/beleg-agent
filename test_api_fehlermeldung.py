"""Tests fuer die Diagnose-Meldung bei Claude-API-Fehlern.

Hintergrund: frueher zeigte der Agent bei JEDEM API-Fehler pauschal
"API nicht erreichbar" — auch bei einem 404 (abgekuendigtes Modell) oder
401 (Key ungueltig). Das hat am 2026-06-25 zu einer Fehldiagnose gefuehrt.
claude_fehler_text() macht aus dem anthropic.APIError eine konkrete Ursache.

Run:  .venv/bin/python3 -m unittest test_api_fehlermeldung -v
"""
import unittest

import beleg_agent


class FakeStatusError:
    """Mimt anthropic.APIStatusError: hat ein .status_code (HTTP-Fehler)."""
    def __init__(self, status_code):
        self.status_code = status_code


class FakeConnectionError:
    """Mimt anthropic.APIConnectionError/Timeout: KEIN .status_code."""


class ClaudeFehlerTextTest(unittest.TestCase):
    def test_404_nennt_modell_und_config(self):
        msg = beleg_agent.claude_fehler_text(FakeStatusError(404))
        self.assertIn("Modell", msg)
        self.assertIn("404", msg)

    def test_401_nennt_key(self):
        self.assertIn("Key", beleg_agent.claude_fehler_text(FakeStatusError(401)))

    def test_403_zugriff(self):
        self.assertIn("403", beleg_agent.claude_fehler_text(FakeStatusError(403)))

    def test_429_rate_limit(self):
        self.assertIn("Rate-Limit", beleg_agent.claude_fehler_text(FakeStatusError(429)))

    def test_5xx_serverfehler(self):
        self.assertIn("Server", beleg_agent.claude_fehler_text(FakeStatusError(503)))

    def test_verbindungsfehler_ohne_status(self):
        self.assertIn("nicht erreichbar",
                      beleg_agent.claude_fehler_text(FakeConnectionError()))

    def test_404_ist_nicht_pauschal_nicht_erreichbar(self):
        # Regression auf den Vorfall: ein 404 darf NICHT als "nicht erreichbar"
        # erscheinen, sonst wird ein totes Modell wieder als Ausfall fehlgedeutet.
        self.assertNotIn("nicht erreichbar",
                         beleg_agent.claude_fehler_text(FakeStatusError(404)))


class ClaudeAPIFehlerTest(unittest.TestCase):
    def test_traegt_kurztext(self):
        e = beleg_agent.ClaudeAPIFehler("Modell nicht gefunden (404).")
        self.assertEqual(e.kurz, "Modell nicht gefunden (404).")
        self.assertIn("404", str(e))


if __name__ == "__main__":
    unittest.main(verbosity=2)
