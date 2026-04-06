"""
Lokale Konfiguration – kopiere diese Datei nach config_local.py und passe die Pfade an.
"""

# Stammpfad zur Belege-Ablage (enthält _Inbox, _Abgleich, _Dauerauftraege, etc.)
# Windows:
ABLAGE_STAMMPFAD = r"C:\Users\DEIN_USER\Pfad\zu\Belege"
# macOS:
# ABLAGE_STAMMPFAD = "/Users/DEIN_USER/Pfad/zu/Belege"

# Optional: Bekannte Kreditkarten (letzte 4 Ziffern → KK-Typ)
# Hilft dem Vision-Modell, die Zahlungsart automatisch zu erkennen.
# BEKANNTE_KARTEN = {
#     "1234": "KK CHF",
#     "5678": "KK EUR",
# }

# Bank-Profil für KK- und Bank-Abgleich
# Verfügbar: "ubs", "raiffeisen", "postfinance"
# Siehe bank_profile.py für Details und um eigene Profile hinzuzufügen.
BANK_PROFIL = "ubs"
