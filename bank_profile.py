"""
Bank-Profile: CSV-Format-Definitionen für verschiedene Schweizer Banken.

Jedes Profil definiert, wie die KK- und Bank-CSV-Exporte der jeweiligen Bank
strukturiert sind (Spaltenname, Trennzeichen, Encoding, etc.).

Neues Profil hinzufügen:
1. Dict nach dem Schema unten erstellen
2. In BANK_PROFILE eintragen
3. In config_local.py BANK_PROFIL = "meine_bank" setzen
"""


# ══════════════════════════════════════════════════════════════════════════
#  KK-Profil: Beschreibt das CSV-Format eines Kreditkarten-Exports
# ══════════════════════════════════════════════════════════════════════════
#
#  "delimiter":        CSV-Trennzeichen (";" oder ",")
#  "skip_first_line":  True wenn erste Zeile "sep=;" o.ä. ist (überspringen)
#  "datum_format":     strftime-Format für Einkaufsdatum
#  "spalten":          Mapping: interner Name → mögliche CSV-Spaltennamen
#                      (Liste von Varianten für Umlaut-Kompatibilität)
#
#  Interne Spalten:
#    datum          - Einkaufs-/Transaktionsdatum
#    buchungstext   - Beschreibung der Transaktion
#    betrag         - Originalbetrag (in Originalwährung)
#    orig_waehrung  - Originalwährung (z.B. USD)
#    kurs           - Wechselkurs (optional)
#    kk_waehrung    - Abrechnungswährung der Karte (z.B. CHF)
#    belastung      - Belasteter Betrag in KK-Währung
#    gutschrift     - Gutschrift (Rückerstattung)
#    branche        - Branche/Kategorie (optional)

# ══════════════════════════════════════════════════════════════════════════
#  Bank-Profil: Beschreibt das CSV-Format eines Bankkonto-Exports
# ══════════════════════════════════════════════════════════════════════════
#
#  "waehrung_erkennung":  Wie die Kontowährung erkannt wird
#    "header_zeile":      Sucht in den ersten N Zeilen nach einem Prefix
#    "spalte":            Liest aus einer bestimmten Spalte
#
#  "daten_start":         Wie der Beginn der Datenzeilen erkannt wird
#    "header_prefix":     Sucht nach Zeile die mit diesem Text beginnt
#
#  "spalten":             Mapping: interner Name → mögliche CSV-Spaltennamen
#
#  "skip_beschreibung":   Liste von Textmustern in Beschreibung → überspringen
#                         (z.B. Gebühren, KK-Rechnungen, FX-Trades)
#
#  Interne Spalten (Bank):
#    datum          - Buchungs-/Abschlussdatum
#    beschreibung   - Liste von Spalten die zur Beschreibung zusammengefügt werden
#    details        - Detailspalte (für Name-Matching, z.B. Empfänger)
#    belastung      - Belastung (Ausgabe)
#    gutschrift     - Gutschrift (Einnahme)
#    einzelbetrag   - Einzelbetrag (wenn Belastung/Gutschrift in einer Spalte, optional)


# ── UBS ───────────────────────────────────────────────────────────────────

UBS_KK = {
    "name": "UBS Kreditkarte",
    "delimiter": ";",
    "skip_first_line": True,           # "sep=;" Zeile
    "datum_format": "%d.%m.%Y",
    "spalten": {
        "datum":          ["Einkaufsdatum"],
        "buchungstext":   ["Buchungstext"],
        "betrag":         ["Betrag"],
        "orig_waehrung":  ["Originalwährung", "Originalw\xe4hrung", "Originalwaehrung"],
        "kurs":           ["Kurs"],
        "kk_waehrung":    ["Währung", "W\xe4hrung", "Waehrung"],
        "belastung":      ["Belastung"],
        "gutschrift":     ["Gutschrift"],
        "branche":        ["Branche"],
    },
}

UBS_BANK = {
    "name": "UBS Bankkonto",
    "delimiter": ";",
    "datum_format": "%Y-%m-%d",
    "waehrung_erkennung": {
        "methode": "header_zeile",
        "prefix": "Bewertet in",       # Zeile: "Bewertet in;CHF"
        "separator": ";",
        "position": 1,                  # 0-indexed: zweites Feld
    },
    "daten_start": {
        "methode": "header_prefix",
        "prefix": "Abschlussdatum;",
    },
    "erkennung": {
        "methode": "header_prefix",     # Wie erkennt man ob eine CSV diese Bank ist?
        "prefix": "Kontonummer",        # Erste Zeile beginnt mit "Kontonummer"
    },
    "spalten": {
        "datum":          ["Abschlussdatum"],
        "beschreibung":   ["Beschreibung1", "Beschreibung2"],  # werden zusammengefügt
        "details":        ["Beschreibung3"],
        "belastung":      ["Belastung"],
        "gutschrift":     ["Gutschrift"],
        "einzelbetrag":   ["Einzelbetrag"],
    },
    "skip_beschreibung": [
        "Dienstleistungspreisabschluss",
        "Depotpreis",
    ],
    "skip_details": [
        "KREDITKARTEN-RECHNUNG",
    ],
    "skip_buchungstext": [
        "Kauf FX Spot",
    ],
}


# ── Raiffeisen ────────────────────────────────────────────────────────────

RAIFFEISEN_KK = {
    "name": "Raiffeisen Kreditkarte (Viseca)",
    "delimiter": ";",
    "skip_first_line": False,
    "datum_format": "%d.%m.%Y",
    "spalten": {
        "datum":          ["Transaktionsdatum", "Datum"],
        "buchungstext":   ["Handelsname", "Beschreibung", "Text"],
        "betrag":         ["Betrag", "Originalb.", "Originalbetrag"],
        "orig_waehrung":  ["Originalwährung", "Orig.währung"],
        "kurs":           ["Kurs", "Wechselkurs"],
        "kk_waehrung":    ["Kartenwährung", "Abrechnungswährung", "Währung"],
        "belastung":      ["Abrechnungsbetrag", "Belastung", "Betrag CHF"],
        "gutschrift":     ["Gutschrift"],
        "branche":        ["Branche", "Kategorie", "MCC"],
    },
}

RAIFFEISEN_BANK = {
    "name": "Raiffeisen Bankkonto",
    "delimiter": ";",
    "datum_format": "%d.%m.%Y",
    "waehrung_erkennung": {
        "methode": "spalte",            # Währung steht in einer Datenspalte
        "spalte": ["Währung", "Waehrung"],
    },
    "daten_start": {
        "methode": "header_prefix",
        "prefix": "IBAN",               # Raiffeisen: Header beginnt mit IBAN
    },
    "erkennung": {
        "methode": "header_prefix",
        "prefix": "IBAN",
    },
    "spalten": {
        "datum":          ["Booked At", "Buchungsdatum", "Valuta"],
        "beschreibung":   ["Text", "Buchungstext"],
        "details":        ["Zahlungszweck", "Details", "Mitteilung"],
        "belastung":      ["Belastung", "Debit"],
        "gutschrift":     ["Gutschrift", "Credit"],
        "einzelbetrag":   ["Betrag", "Amount"],
    },
    "skip_beschreibung": [],
    "skip_details": [],
    "skip_buchungstext": [],
}


# ── PostFinance ───────────────────────────────────────────────────────────

POSTFINANCE_BANK = {
    "name": "PostFinance Bankkonto",
    "delimiter": ";",
    "datum_format": "%Y-%m-%d",
    "waehrung_erkennung": {
        "methode": "spalte",
        "spalte": ["Währung", "Ccy"],
    },
    "daten_start": {
        "methode": "header_prefix",
        "prefix": "Buchungsdatum",
    },
    "erkennung": {
        "methode": "content_contains",
        "text": "PostFinance",
    },
    "spalten": {
        "datum":          ["Buchungsdatum", "Datum"],
        "beschreibung":   ["Buchungsdetails", "Avisierungstext"],
        "details":        ["Zahlungszweck", "Mitteilungen"],
        "belastung":      ["Lastschrift", "Belastung"],
        "gutschrift":     ["Gutschrift"],
        "einzelbetrag":   ["Betrag"],
    },
    "skip_beschreibung": [],
    "skip_details": [],
    "skip_buchungstext": [],
}


# ══════════════════════════════════════════════════════════════════════════
#  Profil-Register
# ══════════════════════════════════════════════════════════════════════════

BANK_PROFILE = {
    "ubs": {
        "kk": UBS_KK,
        "bank": UBS_BANK,
    },
    "raiffeisen": {
        "kk": RAIFFEISEN_KK,
        "bank": RAIFFEISEN_BANK,
    },
    "postfinance": {
        "kk": None,                     # PostFinance hat eigene KK-Formate
        "bank": POSTFINANCE_BANK,
    },
}


def get_profil(bank_name: str) -> dict:
    """Gibt das Bank-Profil zurück. Wirft KeyError wenn unbekannt."""
    bank_name = bank_name.lower().strip()
    if bank_name not in BANK_PROFILE:
        verfuegbar = ", ".join(sorted(BANK_PROFILE.keys()))
        raise KeyError(
            f"Unbekanntes Bank-Profil: '{bank_name}'. "
            f"Verfügbar: {verfuegbar}. "
            f"Bitte BANK_PROFIL in config_local.py anpassen."
        )
    return BANK_PROFILE[bank_name]
