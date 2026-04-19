"""
config.py – Zentrale Einstellungen für alle Brainlab SRS Analytics Skripte.

Änderungen hier wirken sich auf alle Parser (DICOM, PDF) und alle Export-Skripte aus.
Keine Hardcoded Werte mehr in den einzelnen Parser-Dateien nötig.
"""

from pathlib import Path

# ── Pfade (können hier oder via Kommandozeile überschrieben werden) ───────────
# Standard: Ordner auf dem Planungsserver oder lokale Kopien
DEFAULT_DICOM_DIR = Path(r"C:\Users\Aria\Desktop\ElementsTreatPar_DICOM")
DEFAULT_PDF_DIR = Path(r"C:\Users\Aria\Desktop\ElementsTreatPar_PDF")

# ── Patient ID Filter ─────────────────────────────────────────────────────────
# Anzahl der Ziffern für gültige Patienten-IDs (9 für UKE, z.B. 123456789)
PATIENT_ID_DIGITS = 9

# Filter-Modus: True = nur IDs mit genau PATIENT_ID_DIGITS Ziffern (Produktiv)
#               False = alle IDs erlaubt (Debug/Testpläne)
LIVE_FILTER = True

# ── OAR Struktur-Patterns (für DICOM und PDF) ─────────────────────────────────
# Schlüssel = Name im Output, Werte = Liste von Case-Insensitive Substrings
OAR_PATTERNS = {
    "Chiasm": [
        "chiasm", "chiasma", "chiasm oar", "chiasma oar",
    ],
    "Brainstem": [
        "brainstem", "brainstem oar", "hirnstamm", "hirnstamm oar", "brain stem",
    ],
    # Zusätzliche Strukturen (einkommentieren zum Aktivieren):
    # "OpticNerveL": ["nopticusl", "opticusl", "opticus l", "opticus li", "n opticus l"],
    # "OpticNerveR": ["nopticusr", "opticusr", "opticus r", "opticus re", "n opticus r"],
    # "CochleaL": ["cochlea l", "cochlea li", "cochlea left", "schnecke l"],
    # "CochleaR": ["cochlea r", "cochlea re", "cochlea right", "schnecke r"],
    # "Pituitary": ["pituitary", "hypophyse", "pituitary gland"],
}

# ── GTV Erkennung ─────────────────────────────────────────────────────────────
# Patterns für GTV-Namen (z.B. Met1, GTV01, etc.)
GTV_NAME_PATTERNS = [
    r"GTV\d+",           # GTV01, GTV02, ...
    r"GTV[_-]?\d+",     # GTV_01, GTV-02, ...
    r"Met\d+",          # Met1, Met2, ... (Brainlab Standard)
    r"MET\d+",          # MET1, MET2, ...
    r" metastasis \d+", # " metastasis 1" etc.
]

# ── PTV → GTV Mapping ─────────────────────────────────────────────────────────
# Wenn True: Versuche GTV per Namenskonvention zu finden (PTV01 → GTV01)
# Wenn False: Nur exakte Namensgleichheit oder Fallback-Met{N}
USE_PTV_TO_GTV_NAMING = True

# Regex für PTV→GTV Extraktion (z.B. PTV03L → 03 → GTV03L)
PTV_EXTRACT_NUMBER_RE = r"PTV(\d+[LR]?)"
GTV_BUILD_NAME_TEMPLATE = "GTV{}"  # {} wird durch extrahierte Nummer ersetzt

# ── Margin-Berechnung ─────────────────────────────────────────────────────────
# Schwelle für Margin-Anzeige (nur GTVs mit Margin <= WERT werden ausgewiesen)
MARGIN_THRESHOLD_MM = 6.0

# ── PDF Parser ───────────────────────────────────────────────────────────────
# Maximale Anzahl von Zeilen die als "alte" Version gelten (v1.5 hat weniger Zeilen)
PDF_V15_MAX_LINES = 50

# Hinweis: Keine Fallback-Dosis oder Fallback-Fractions hinterlegt,
# da dies die Daten verfälschen würde. Fehlende Werte bleiben leer (None).
