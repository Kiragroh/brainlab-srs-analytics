"""PDF Treatment Parameter Report Parser → dicom_data_ptv.csv-kompatibles Format

Parst Brainlab TreatPar.pdf Dateien (älteres Cranial SRS 3.x + neueres 4.5.x Format)
und extrahiert PTV-Level-Daten in der gleichen Struktur wie der DICOM-Parser.

Benötigt: pymupdf (fitz), pandas, numpy
"""

import os
import re
import sys
import warnings
import numpy as np
import pandas as pd
import pymupdf
from pathlib import Path
from tqdm import tqdm

warnings.filterwarnings("ignore")


# ---------------------------------------------------------------------------
# 1) Hilfsfunktionen
# ---------------------------------------------------------------------------

def norm_text(text):
    """Normalisiert Text: non-breaking spaces → space, tabs → space."""
    return text.replace('\xa0', ' ').replace('\u2009', ' ').replace('\t', ' ')


def parse_filename(fn):
    """Extrahiert Patient Name, ID, Plan Name aus dem Dateinamen.
    Format: Name.PatientID.PlanName.TreatPar.pdf
    """
    base = fn
    if base.endswith(".TreatPar.pdf"):
        base = base[:-len(".TreatPar.pdf")]
    elif base.endswith(".pdf"):
        base = base[:-4]
    dot_parts = base.split(".")
    if len(dot_parts) >= 3:
        return dot_parts[0], dot_parts[1], ".".join(dot_parts[2:])
    return base, "", ""


def safe_float(val, default=None):
    """Sicher float konvertieren."""
    if val is None:
        return default
    try:
        v = str(val).replace(",", ".").replace("—", "").replace("n/a", "").replace("*", "").strip()
        if not v or v in ("-", "–"):
            return default
        return float(v)
    except (ValueError, TypeError):
        return default


# ---------------------------------------------------------------------------
# 2) Text-Extraktion mit PyMuPDF
# ---------------------------------------------------------------------------

def extract_all_text(filepath):
    """Extrahiert Text aus allen Seiten via PyMuPDF."""
    doc = pymupdf.open(str(filepath))
    pages_text = []
    for page in doc:
        pages_text.append(norm_text(page.get_text()))
    doc.close()
    return pages_text, "\n".join(pages_text)


# ---------------------------------------------------------------------------
# 3) PRESCRIPTION-Abschnitt parsen
# ---------------------------------------------------------------------------

def parse_prescription_section(all_text):
    """Parst den PRESCRIPTION-Abschnitt und gibt eine Liste von PTV-Dicts zurück.

    Struktur:
    PRESCRIPTION
    [Header-Zeilen...]
    PTVname [mehrzeilig]
    PTV
    dose_per_fx
    n_fractions
    presc_dose
    cov_dose
    cov_vol_pct
    vol_cc
    size_cm
    [nächster PTV oder Ende des Abschnitts]
    """
    # Abschnitt isolieren
    start_m = re.search(r'\bPRESCRIPTION\b', all_text)
    if not start_m:
        return []
    # Ende: ISOCENTER oder TARGETS oder TREATED METASTASES oder ähnliches
    end_m = re.search(r'\b(ISOCENTER|TARGETS|TREATED METASTASES|UNTREATED|OTHERS|MACHINE)\b',
                      all_text[start_m.end():])
    section = all_text[start_m.end(): start_m.end() + end_m.start()] if end_m else all_text[start_m.end():]

    # Headerzeilen überspringen (bis erste Zeile die mit PTV anfängt oder "PTV" als Typ-Zeile auftaucht)
    lines = [l.strip() for l in section.split('\n')]
    # Suche nach Header-Ende: wenn wir "PTV" als Type-Zeile sehen oder eine float-Zeile nach PTV-Name
    ptv_list = []
    i = 0
    # Überspringe echte Header-Zeilen
    header_kws = {"Object Name", "Object Type", "DICOM", "Dose /", "Fraction [Gy]", "No. of",
                  "Fractions", "Prescription", "Dose [Gy]", "Coverage", "Volume [%]", "Volume",
                  "[cm³]", "Size", "[cm]", "Prescription Dose [Gy]", "Coverage Dose [Gy]",
                  "Coverage Volume [%]"}
    while i < len(lines):
        l = lines[i]
        # PTV-Name erkennen: direkt (PTV1, PTVname) oder mit Leerzeichen (PTV M..., PTV1R M...)
        is_ptv_name = (re.match(r'PTV\S', l) or re.match(r'PTV\s+\S', l)) and l != 'PTV'
        if is_ptv_name:
            # Sammle Namen-Zeilen bis "PTV" als alleinstehende Typ-Zeile
            name_lines = [l]
            j = i + 1
            while j < len(lines) and lines[j] != "PTV" and not re.match(r'PTV[\S ]', lines[j]):
                nl = lines[j]
                if not nl or nl in header_kws:
                    j += 1
                    continue
                # Stoppe wenn Zeile mit Ziffer beginnt (Datenwert)
                if nl and nl[0].isdigit():
                    break
                if not re.fullmatch(r'[\d.]+', nl):
                    name_lines.append(nl)
                j += 1
            if j < len(lines) and lines[j] == "PTV":
                # Sammle Zahlen bis zum nächsten PTV-Name oder Abschnittsende
                nums = []
                k = j + 1
                while k < len(lines) and len(nums) < 8:
                    lk = lines[k].replace("\\", "").strip()
                    if re.match(r'PTV\S', lk) or lk == "PTV":
                        break  # nächster PTV
                    if re.fullmatch(r'[\d.]+', lk):
                        nums.append(float(lk))
                    k += 1
                if len(nums) >= 7:
                    name = " ".join(name_lines)
                    # cols: dose_per_fx, n_fractions, presc_dose, cov_dose, cov_vol_pct, vol_cc, size_cm
                    ptv_list.append({
                        "PTVname": name,
                        "Prescribed dose": nums[2] if len(nums) > 2 else nums[0],
                        "Number of fractions": int(nums[1]) if len(nums) > 1 else None,
                        "PTVvolume": nums[5] if len(nums) > 5 else None,
                        "Max diameter": nums[6] if len(nums) > 6 else None,
                        "Prescribed coverage": nums[4] / 100.0 if len(nums) > 4 else None,
                    })
                    i = k
                    continue
        i += 1
    return ptv_list


# ---------------------------------------------------------------------------
# 4) TARGETS/TREATED METASTASES parsen
# ---------------------------------------------------------------------------

def parse_targets_section(all_text):
    """Parst den TARGETS / TREATED METASTASES Abschnitt.

    Neueres Format (4.x):
    PTV_NAME
    min_dose D98.5 % = actual_dose
    V 24.00Gy = 98.5
    mean_dose
    D1.0 % = max_desired_dose
    max_dose
    CI GI local_v max_dose_rel

    Älteres Format (3.x):
    PTV_NAME
    min_dose D_coverage_pct % = actual_dose
    mean_dose
    D_max_pct % = max_dose_constraint
    max_dose
    CI GI local_v max_dose_rel
    """
    # Finde TARGETS oder TREATED METASTASES
    start_m = re.search(r'\b(TREATED METASTASES|TARGETS)\b', all_text)
    if not start_m:
        return {}
    # Ende: UNTREATED oder OTHERS
    end_m = re.search(r'\b(UNTREATED|OTHERS|PLAN ANALYSIS|GLOBAL SETTINGS|MACHINE)\b',
                      all_text[start_m.end():])
    section = all_text[start_m.end(): start_m.end() + end_m.start()] if end_m else all_text[start_m.end():]

    ptv_ci_gi = {}
    # Suche alle PTV-Namen und dann CI/GI nach dem Namen
    # Muster: nach PTV_NAME erscheinen CI als floats getrennt durch \n
    # Einfachster Ansatz: finde "CI" und "GI" Werte direkt im Text wenn vorhanden (Plan Analysis),
    # oder extrahiere sie aus der Tabellenstruktur

    # Für beide Formate: PTV-Name gefolgt von mehreren Zahlen
    # Die CI steht immer an Position 7 (neu) oder 6 (alt) von der Zeile an
    # Suche direkt nach: PTV-Name, dann Zahlenblocks, dann CI, GI

    lines = [l.strip() for l in section.split('\n')]
    # Erkennung: hat der Abschnitt eine "Local V" Spalte?
    has_local_v_col = bool(re.search(r'Local\s+V[\d.]+\s*Gy', section[:600]))
    i = 0
    current_ptv = None
    nums_buffer = []
    current_cov_pct = None

    while i < len(lines):
        l = lines[i]
        # Neue PTV-Zeile
        if re.match(r'PTV\S', l) or re.match(r'PTV\s+\S', l):
            # Flush vorherigen PTV
            if current_ptv and len(nums_buffer) >= 7:
                _store_targets_ptv(ptv_ci_gi, current_ptv, nums_buffer, has_local_v_col, current_cov_pct)
            current_ptv = l
            nums_buffer = []
            current_cov_pct = None
            # Prüfe ob Name mehrzeilig
            j = i + 1
            while j < len(lines):
                nl = lines[j]
                if not nl:
                    j += 1
                    continue
                # Stoppe wenn Datenzeile beginnt
                if nl and nl[0].isdigit():
                    break  # starts with digit = min_dose oder combined
                if re.match(r'D[\d]', nl) or re.match(r'DMax', nl):
                    break
                if re.match(r'V\s*[\d.]+', nl):
                    break
                if nl in ("n/a", "—", "-"):
                    break
                if re.match(r'PTV\S', nl):
                    break
                current_ptv += " " + nl
                j += 1
            i = j
            continue

        # Kombinierte Zeile "float D98.x % = float" (min_dose + coverage)
        m = re.match(r'([\d.]+)\s+D[\d.]+\s*%\s*=\s*([\d.]+)', l)
        if m:
            nums_buffer.append((float(m.group(2)), False))  # cov_dose
            i += 1
            continue

        # Coverage-Zeile: "D98.5 % = 24.00" (standalone)
        m = re.match(r'D[\d.]+\s*%\s*=\s*([\d.]+)', l)
        if m:
            nums_buffer.append((float(m.group(1)), False))
            i += 1
            continue

        # Max/Desired: "DMax = 32.12" oder "D1.0 % = 31.90"
        m = re.match(r'D(?:Max|[\d.]+\s*%)\s*=\s*([\d.]+)', l)
        if m:
            nums_buffer.append((float(m.group(1)), False))
            i += 1
            continue

        # Coverage vol: "V 24.00Gy = 98.5" → Deckungsgrad merken
        m_cv = re.match(r'V\s*[\d.]+\s*Gy\s*=\s*([\d.]+)', l)
        if m_cv:
            current_cov_pct = float(m_cv.group(1))
            i += 1
            continue
        if re.match(r'V\s*=\s*[\d.]+', l):
            i += 1
            continue

        # Kombinierte Zeile "n/a float**" oder "float float**" (GI + local_v zusammen)
        m = re.match(r'^(n/a|[\d.]+)\s+([\d.]+)\*+\s*$', l)
        if m:
            gi_val = np.nan if m.group(1) == 'n/a' else float(m.group(1))
            lv_val = float(m.group(2))
            nums_buffer.append((gi_val, False))   # GI
            nums_buffer.append((lv_val, True))    # LocalV (markiert)
            i += 1
            continue

        # Standalone float mit * / ** (local_v): "1506.4**", "0.3*"
        m = re.match(r'^([\d.]+)\*+\s*$', l)
        if m:
            nums_buffer.append((float(m.group(1)), True))  # LocalV (markiert)
            i += 1
            continue

        # Reine float-Zeile
        lc = l.replace("\\", "").strip()
        if re.fullmatch(r'[\d.]+', lc):
            nums_buffer.append((float(lc), False))
            i += 1
            continue

        # "n/a" oder Gedankenstrich für CI/GI
        if l in ("n/a", "\u2014", "\u2013"):
            nums_buffer.append((np.nan, False))
            i += 1
            continue

        i += 1

    # Letzten PTV flushen
    if current_ptv and len(nums_buffer) >= 7:
        _store_targets_ptv(ptv_ci_gi, current_ptv, nums_buffer, has_local_v_col, current_cov_pct)

    return ptv_ci_gi


def _store_targets_ptv(ptv_dict, name, nums, has_local_v_col=False, cov_pct=None):
    """nums = list of (value, is_lv) tuples. is_lv=True markiert Local-V**-Werte (3.x).
    has_local_v_col=True wenn Header eine Local-V-Spalte enthält (3.x und 4.x).
    Format-Varianten:
      2.x: kein LocalV          → CI, GI, MaxDoseRel
      3.x: LocalV mit ** Marker → CI, GI, LocalV**, MaxDoseRel (is_lv=True)
      4.x: LocalV ohne Marker   → CI, GI, LocalV, MaxDoseRel (all plain, header-Flag nötig)
    """
    ci = None
    gi = None
    local_v = None
    max_dose_rel = None

    # 3.x: LocalV aus is_lv-markierten Werten (** in Daten)
    lv_marked = [v for v, is_lv in nums if is_lv and not np.isnan(v)]
    if lv_marked:
        local_v = lv_marked[0]

    # Nicht-LV-Werte für CI/GI/MaxDoseRel mit orig-Index
    plain = [(v, orig) for orig, (v, is_lv) in enumerate(nums) if not is_lv]

    # CI: erster Nicht-NaN-Wert in [0.8, 3.5] nach orig-Index 2
    ci_pi = None
    for pi, (v, orig) in enumerate(plain):
        if orig <= 2:
            continue
        if not np.isnan(v) and 0.8 <= v <= 3.5:
            ci = round(v, 3)
            ci_pi = pi
            break

    if ci_pi is not None:
        # GI: nächster plain-Wert
        if ci_pi + 1 < len(plain):
            gv, _ = plain[ci_pi + 1]
            if np.isnan(gv) or (1.0 <= gv <= 15.0):
                gi = round(gv, 3) if not np.isnan(gv) else None

        if lv_marked:
            # 3.x: LocalV bereits gesetzt, MaxDoseRel an ci_pi+2
            if ci_pi + 2 < len(plain):
                mr, _ = plain[ci_pi + 2]
                if not np.isnan(mr):
                    max_dose_rel = round(float(mr), 1)
        elif has_local_v_col:
            # 4.x: LocalV ist ci_pi+2 plain, MaxDoseRel ist ci_pi+3
            if ci_pi + 2 < len(plain):
                lv, _ = plain[ci_pi + 2]
                if not np.isnan(lv):
                    local_v = lv
            if ci_pi + 3 < len(plain):
                mr, _ = plain[ci_pi + 3]
                if not np.isnan(mr):
                    max_dose_rel = round(float(mr), 1)
        else:
            # 2.x: kein LocalV, MaxDoseRel an ci_pi+2
            if ci_pi + 2 < len(plain):
                mr, _ = plain[ci_pi + 2]
                if not np.isnan(mr):
                    max_dose_rel = round(float(mr), 1)

    # Dose metrics aus führenden Buffer-Werten (format-unabhängig)
    plain_vals = [v for v, _ in plain]
    coverage_dose = round(plain_vals[0], 2) if plain_vals and not np.isnan(plain_vals[0]) else None
    mean_dose     = round(plain_vals[1], 2) if len(plain_vals) > 1 and not np.isnan(plain_vals[1]) else None
    # NearMax (D1%): 2 Positionen vor CI (plain[ci_pi-2])
    nearmax_dose  = None
    if ci_pi is not None and ci_pi >= 2:
        nm = plain_vals[ci_pi - 2]
        if not np.isnan(nm):
            nearmax_dose = round(nm, 2)

    ptv_dict[name.strip()] = {
        "CI": ci, "GI": gi,
        "_local_v_cc": local_v, "_max_dose_rel": max_dose_rel,
        "_coverage_dose": coverage_dose,
        "_actual_cov_pct": (round(cov_pct / 100.0, 4) if cov_pct is not None else None),
        "_mean_dose": mean_dose,
        "_nearmax_dose": nearmax_dose,
    }


# ---------------------------------------------------------------------------
# 5) PLAN ANALYSIS parsen (älteres Format: CI/GI pro PTV in Text)
# ---------------------------------------------------------------------------

def parse_plan_analysis(all_text):
    """Parst PLAN ANALYSIS TABLE für CI/GI pro PTV (älteres Format)."""
    ptv_ci_gi = {}
    # Suche "CI: X.XX" und "GI: X.XX" nach PTV-Namen (auch "PTV M..." mit Leerzeichen)
    for m in re.finditer(r'(PTV[^\n]{2,}?)\n(?:.*\n)*?CI:\s*([\d.]+|n/a)\n\s*GI:\s*([\d.]+|n/a)',
                         all_text, re.MULTILINE):
        name = m.group(1).strip()
        ci = safe_float(m.group(2))
        gi = safe_float(m.group(3))
        ptv_ci_gi[name] = {"CI": ci, "GI": gi}

    # Auch MCS extrahieren
    mcs = None
    mcs_m = re.search(r'(\d+)\s*MU\s*\nMCS:\s*([\d.]+)', all_text)
    if not mcs_m:
        mcs_m = re.search(r'MCS:\s*([\d.]+)', all_text)
        if mcs_m:
            mcs = float(mcs_m.group(1))
    else:
        mcs = float(mcs_m.group(2))

    return ptv_ci_gi, mcs


# ---------------------------------------------------------------------------
# 6) OTHERS / OAR-Abschnitt parsen
# ---------------------------------------------------------------------------

def parse_oar_section(all_text):
    """Parst OTHERS-Abschnitt: OAR Name → {max_dose, mean_dose}."""
    oar_data = {}
    # Find all OTHERS sections (may appear multiple times across pages)
    for others_m in re.finditer(r'\bOTHERS\b', all_text):
        # Next section start
        end_m = re.search(r'\b(MACHINE|ISOCENTER|PLAN ANALYSIS|Created|APPROVAL|Page \d+)\b',
                          all_text[others_m.end():])
        section = all_text[others_m.end(): others_m.end() + end_m.start()] if end_m else all_text[others_m.end():]

        # Skip header lines: "Object Name", "Object Type DICOM", "Volume", "Mean Dose", "Max Dose"
        lines = [l.strip() for l in section.split('\n') if l.strip()]
        header_kws = {"Object Name", "Object Type DICOM", "Object Type", "Volume", "Mean Dose",
                      "Max Dose", "[cm³]", "[Gy]", "DICOM",
                      "AVOIDANCE", "ORGAN", "EXTERNAL", "OTHER", "CONTROL",
                      "TREATED_VOLUME", "SUPPORT", "FIXATION", "BOLUS"}

        i = 0
        while i < len(lines):
            name_l = lines[i]
            # OAR name lines: not a float, not a keyword, not "PTV"
            if (name_l not in header_kws
                    and not re.fullmatch(r'[\d.]+', name_l)
                    and not name_l.startswith("PTV")):
                # Check if it looks like an OAR name (has letters, not just numbers)
                if re.search(r'[A-Za-zäöüÄÖÜß]', name_l):
                    # Look for 3 floats ahead (volume, mean_dose, max_dose)
                    name = name_l
                    # Name may span 2 lines (e.g. "Hirnstamm OAR")
                    j = i + 1
                    floats = []
                    # Skip type lines
                    type_kws = {"AVOIDANCE", "ORGAN", "EXTERNAL", "OTHER", "CONTROL",
                                "TREATED_VOLUME", "SUPPORT", "FIXATION", "BOLUS"}
                    while j < len(lines) and len(floats) < 3:
                        nl = lines[j]
                        if nl in type_kws:
                            j += 1
                            continue
                        if re.fullmatch(r'[\d.]+', nl):
                            floats.append(float(nl))
                        elif not re.search(r'[A-Za-zäöüÄÖÜß]', nl) and nl:
                            pass
                        else:
                            if len(floats) == 0 and not re.fullmatch(r'[\d.]+', nl):
                                # Might be continuation of name or type
                                pass
                        j += 1

                    if len(floats) >= 3:
                        oar_data[name] = {
                            "volume_cc": floats[0],
                            "mean_dose": floats[1],
                            "max_dose": floats[2],
                        }
                        i = j
                        continue
            i += 1

    return oar_data


# ---------------------------------------------------------------------------
# 7) Arc-Daten und Maschinendaten parsen
# ---------------------------------------------------------------------------

def parse_machine_and_arcs(all_text):
    """Extrahiert Maschinendaten, Arc-Count und MU aus dem Text."""
    result = {"machine": "", "n_arcs": None, "total_mu": None, "table_angles": set()}

    # Machine
    m = re.search(r'MACHINE\s*:\s*(.+?)(?:\n|$)', all_text)
    if m:
        result["machine"] = m.group(1).strip()

    # Total MU aus GLOBAL SETTINGS – einzeilig ODER mehrzeilig (älteres Format)
    m = re.search(r'Total\s+Monitor\s+Units\s*\n?\s*([\d,]+)\s*MU', all_text)
    if m:
        result["total_mu"] = float(m.group(1).replace(",", ""))

    # MU aus zweiter Arc-Tabelle: "NNN (MxN.N)" Pattern (bis zu 8 Zwischenzeilen)
    arc_mus_2 = re.findall(r'\d+\s*:\s*Arc\s*[\d ()]+\s*(?:[^\n]*\n){0,8}?([\d]+)\s*\(\d+x', all_text)
    if arc_mus_2:
        result["total_mu"] = sum(float(x) for x in arc_mus_2)

    # Fallback: erste Arc-Tabelle – MU direkt vor Dosisrate 600
    # Format: "...\nMU_value\n600\n" (NNN gefolgt von 600 Dosisrate)
    if not result["total_mu"]:
        arc_mus_1 = re.findall(r'\b(\d{3,5})\n600\n', all_text)
        if arc_mus_1:
            result["total_mu"] = sum(float(x) for x in arc_mus_1)

    # Arc-Count aus der Arc-Tabelle
    arc_entries = re.findall(r'(\d+)\s*:\s*Arc\s*\d+', all_text)
    if arc_entries:
        result["n_arcs"] = len(set(arc_entries))  # unique arc numbers

    # Arc-Count aus Description als Fallback
    if not result["n_arcs"]:
        m = re.search(r'(\d+)\s*(?:VMAT|Dyn\.?|DCA)\s*Arcs?', all_text, re.IGNORECASE)
        if m:
            result["n_arcs"] = int(m.group(1))

    # Table angles
    ta_m = re.findall(r'\d+\s*:\s*Arc\s*\d+\s*([\d.]+)', all_text)
    result["table_angles"] = set(ta_m)

    return result


# ---------------------------------------------------------------------------
# 8) Header-Daten parsen
# ---------------------------------------------------------------------------

def parse_header(all_text, filename_pid, filename_plan):
    """Extrahiert Patient ID, Plan Name, App, Version, Approval, Creation."""
    result = {
        "patient_id": filename_pid,
        "plan_name": filename_plan,
        "patient_name": "",
        "app_name": "",
        "app_version": "",
        "approval_status": "",
        "creation_date": "",
        "description": "",
        "ci_vol_avg": None,
        "gi_vol_avg": None,
        "cum_ptv_vol": None,
        "global_v_cc": None,
        "global_v_dose": None,
        "global_v12_cc": None,
    }

    # Patient ID
    m = re.search(r'Patient\s+ID:\s*\n?\s*(\d+)', all_text)
    if m:
        result["patient_id"] = m.group(1).strip()

    # Patient Name (newer: "LAST, First" or older: "Last^First")
    m = re.search(r'Patient\s+Name:\s*\n?\s*([^\n]+)', all_text)
    if m:
        result["patient_name"] = m.group(1).strip()

    # Plan Name (newer format)
    m = re.search(r'Plan\s+Name:\s*\n?\s*([^\n]+)', all_text)
    if m:
        result["plan_name"] = m.group(1).strip()
    else:
        # Older format: 3rd non-empty line after "a\n[App version]\n[planname]\n"
        lines = [l.strip() for l in all_text.split('\n') if l.strip()]
        for j, l in enumerate(lines):
            if re.match(r'(Cranial SRS|Multiple Brain Mets SRS)\s+[\d.]+', l) and j + 1 < len(lines):
                result["plan_name"] = lines[j + 1]
                break

    # App name + version
    for pattern in [r'(Multiple\s+Brain\s+Mets\s+SRS)\s+([\d.]+)',
                    r'(Cranial\s+SRS)\s+([\d.]+)',
                    r'(Multiple\s+Brain\s+Mets\s+SRS)',
                    r'(Cranial\s+SRS)']:
        m = re.search(pattern, all_text)
        if m:
            result["app_name"] = m.group(1).strip()
            result["app_version"] = m.group(2).strip() if m.lastindex >= 2 else ""
            break

    # Approval
    m = re.search(r'\b(APPROVED|DEMOTED|VERIFIED)\b', all_text)
    if m:
        result["approval_status"] = m.group(1)

    # Creation date
    m = re.search(r'Created\s+([\d\-\w]+\s+[\d:]+)', all_text)
    if m:
        result["creation_date"] = m.group(1).strip()

    # Description
    m = re.search(r'Description:\s*\n?\s*([^\n]+)', all_text)
    if m:
        result["description"] = m.group(1).strip()

    # Global settings (nur altes Format)
    m = re.search(r'Total Monitor Units\s+([\d,]+)\s*MU', all_text)
    # (wird auch in parse_machine_and_arcs gesucht)

    m = re.search(r'Volume Averaged CI\s*/\s*GI\s*\n?\s*([\d.]+)\s*/\s*([\d.—–]+)', all_text)
    if m:
        result["ci_vol_avg"] = safe_float(m.group(1))
        result["gi_vol_avg"] = safe_float(m.group(2))

    m = re.search(r'Cumulative Target Volume\s*\n?\s*([\d.]+)\s*cm', all_text)
    if m:
        result["cum_ptv_vol"] = float(m.group(1))

    # Global V: flexibles Muster (ein- oder mehrzeilig)
    for vm in re.finditer(r'Global V([\d.]+)\s*Gy\s*\nVolume of[^\n]*?\n?\s*([\d.]+)\s*cm', all_text):
        dose = float(vm.group(1))
        val  = float(vm.group(2))
        if result["global_v_dose"] is None:          # erstes Global V speichern
            result["global_v_dose"] = dose
            result["global_v_cc"]   = val
        if abs(dose - 12.0) < 0.1:                   # V12 spezifisch
            result["global_v12_cc"] = val
    # Fallback: wenn allgemeines global_v_cc bei Dose==12
    if result["global_v12_cc"] is None and result.get("global_v_dose") and abs(result["global_v_dose"] - 12.0) < 0.1:
        result["global_v12_cc"] = result["global_v_cc"]

    return result


# ---------------------------------------------------------------------------
# 9) OAR-Namen matchen
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# KONFIGURATION: OAR-Strukturen für PDF-Analyse (erweiterbar)
# ---------------------------------------------------------------------------
OAR_STRUCTURE_TARGETS_PDF = {
    "Chiasm":    ["chiasm", "chiasma", "chiasm oar", "chiasma oar"],
    "Brainstem": ["brainstem", "brainstem oar", "hirnstamm", "hirnstamm oar", "brain stem", "truncus"],
    # "OpticNerveL": ["nopticusl", "opticusl", "opticus l", "opticus li"],
    # "OpticNerveR": ["nopticusr", "opticusr", "opticus r", "opticus re"],
    # "Myelon":      ["myelon"],  # Halsmark/Rückenmark (kein Hirnstamm!)
    # "Cochlea":     ["cochlea"],
}

CHIASM_PATS    = OAR_STRUCTURE_TARGETS_PDF["Chiasm"]
BRAINSTEM_PATS = OAR_STRUCTURE_TARGETS_PDF["Brainstem"]


def find_oar(oar_data, patterns):
    """Findet OAR-Name nach Patterns (kürzester Match)."""
    matches = [n for n in oar_data if any(p in n.lower() for p in patterns)]
    return min(matches, key=len) if matches else None


# ---------------------------------------------------------------------------
# 10) Kern-Parser: PDF → Ergebnis-Dict
# ---------------------------------------------------------------------------

def parse_treat_par_pdf(filepath):
    """Parst eine Brainlab TreatPar.pdf und gibt strukturierte Daten zurück."""
    fn = os.path.basename(filepath)
    _, file_pid, file_plan = parse_filename(fn)

    try:
        _, all_text = extract_all_text(filepath)
    except Exception as e:
        print(f"  FEHLER beim Lesen: {filepath}: {e}")
        return None

    if len(all_text.strip()) < 100:
        print(f"  WARNUNG: Kein Text extrahierbar in {fn}")
        return None

    # --- Header ---
    hdr = parse_header(all_text, file_pid, file_plan)

    # --- PRESCRIPTION ---
    presc_ptvs = parse_prescription_section(all_text)

    # --- TARGETS ---
    tgt_map = parse_targets_section(all_text)

    # Fallback: PLAN ANALYSIS (älteres Format)
    pa_map, mcs = parse_plan_analysis(all_text)

    # --- OAR ---
    oar_data = parse_oar_section(all_text)

    # --- Machine / Arc ---
    arc_info = parse_machine_and_arcs(all_text)

    # --- PTV-Daten zusammenführen ---
    # Merge PRESCRIPTION → TARGETS → PLAN ANALYSIS
    ptv_map = {p["PTVname"]: dict(p) for p in presc_ptvs}

    def find_matching_ptv(name, ref_map):
        """Findet besten Match für PTV-Name in ref_map."""
        name_clean = name.strip()
        # 1) Exakter Match
        if name_clean in ref_map:
            return name_clean
        # 2) Normalisierter Match (Leerzeichen entfernen, lowercase)
        name_norm = re.sub(r'\s+', '', name_clean).lower()
        for k in ref_map:
            if re.sub(r'\s+', '', k).lower() == name_norm:
                return k
        # 3) Substring
        for k in ref_map:
            if k in name_clean or name_clean in k:
                return k
        # 4) Normalisierter Substring
        for k in ref_map:
            kn = re.sub(r'\s+', '', k).lower()
            if kn in name_norm or name_norm in kn:
                return k
        # 5) Fuzzy: gemeinsame Wörter
        name_words = set(name_clean.lower().split())
        best, best_score = None, 0
        for k in ref_map:
            kw = set(k.lower().split())
            score = len(name_words & kw)
            if score > best_score:
                best, best_score = k, score
        return best if best_score > 0 else None

    # CI/GI aus TARGETS oder PLAN ANALYSIS
    for ptv_name, ptv in ptv_map.items():
        # TARGETS-Daten
        tgt_key = find_matching_ptv(ptv_name, tgt_map)
        if tgt_key:
            tgt = tgt_map[tgt_key]
            ptv["CI"] = tgt.get("CI")
            ptv["GI"] = tgt.get("GI")
            ptv["_local_v_cc"]     = tgt.get("_local_v_cc")
            ptv["_max_dose_rel"]   = tgt.get("_max_dose_rel")
            ptv["_coverage_dose"]  = tgt.get("_coverage_dose")
            ptv["_actual_cov_pct"] = tgt.get("_actual_cov_pct")
            ptv["_mean_dose"]      = tgt.get("_mean_dose")
            ptv["_nearmax_dose"]   = tgt.get("_nearmax_dose")
        # PLAN ANALYSIS hat Vorrang für CI/GI (präziser im alten Format)
        pa_key = find_matching_ptv(ptv_name, pa_map)
        if pa_key:
            pa = pa_map[pa_key]
            if pa.get("CI") is not None:
                ptv["CI"] = pa["CI"]
            if pa.get("GI") is not None:
                ptv["GI"] = pa["GI"]

    # Wenn keine Prescription-Daten, aber TARGETS-Daten vorhanden → erstelle PTV-Einträge
    if not ptv_map and tgt_map:
        for name in tgt_map:
            ptv_map[name] = {"PTVname": name, **tgt_map[name]}
    if not ptv_map and pa_map:
        for name in pa_map:
            ptv_map[name] = {"PTVname": name, **pa_map[name]}

    # --- OAR-Metriken ---
    chiasm_name = find_oar(oar_data, CHIASM_PATS)
    brainstem_name = find_oar(oar_data, BRAINSTEM_PATS)
    ch_dmax = oar_data[chiasm_name]["max_dose"] if chiasm_name else np.nan
    bs_dmax = oar_data[brainstem_name]["max_dose"] if brainstem_name else np.nan

    # --- Optimizer ---
    app_name = hdr["app_name"]
    n_ptvs = len(ptv_map)
    if "Cranial SRS" in app_name and n_ptvs <= 2:
        optimizer = "CranialSRS"
    elif "Multiple Brain Mets" in app_name or n_ptvs > 2:
        optimizer = "MultiMets"
    else:
        optimizer = app_name or "Unknown"

    # --- Fractions ---
    fractions = None
    for ptv in ptv_map.values():
        if ptv.get("Number of fractions"):
            fractions = ptv["Number of fractions"]
            break
    if fractions is None:
        m = re.search(r'(\d+)\s*fx', hdr.get("description", ""), re.IGNORECASE)
        if m:
            fractions = int(m.group(1))
        else:
            m = re.search(r'/\s*(\d+)\s*fx', hdr.get("description", ""), re.IGNORECASE)
            if m:
                fractions = int(m.group(1))

    # MCS aus Plan Analysis
    if mcs is None:
        mcs_m = re.search(r'MCS:\s*([\d.]+)', all_text)
        if mcs_m:
            mcs = float(mcs_m.group(1))

    # CI volume averaged: aus Header, volumengewichteter Mittelwert oder Einzel-PTV-Fallback
    ci_vol_avg = hdr["ci_vol_avg"]
    if ci_vol_avg is None:
        # Volumengewichteter Mittelwert wenn Volumen vorhanden
        ci_vals = [(ptv.get("CI"), ptv.get("PTVvolume"))
                   for ptv in ptv_map.values()
                   if ptv.get("CI") is not None and ptv.get("PTVvolume")]
        if ci_vals:
            total_vol = sum(v for _, v in ci_vals)
            if total_vol > 0:
                ci_vol_avg = round(sum(c * v for c, v in ci_vals) / total_vol, 3)
        # Fallback: Einzel-PTV ohne Volumen-Anforderung
        if ci_vol_avg is None:
            ci_only = [ptv.get("CI") for ptv in ptv_map.values() if ptv.get("CI") is not None]
            if len(ci_only) == 1:
                ci_vol_avg = ci_only[0]

    return {
        "patient_id": hdr["patient_id"],
        "patient_name": hdr["patient_name"],
        "plan_name": hdr["plan_name"],
        "optimizer": optimizer,
        "app_name": app_name,
        "app_version": hdr["app_version"],
        "approval_status": hdr["approval_status"],
        "creation_date": hdr["creation_date"],
        "machine": arc_info["machine"],
        "fractions": fractions,
        "n_arcs": arc_info["n_arcs"],
        "total_mu": round(arc_info["total_mu"] / 1000.0, 3) if arc_info["total_mu"] else None,
        "mcs": mcs,
        "ci_vol_avg": ci_vol_avg,
        "gi_vol_avg": hdr["gi_vol_avg"],
        "cum_ptv_vol": hdr["cum_ptv_vol"],
        "global_v12_cc": hdr["global_v12_cc"],
        "ptv_details": list(ptv_map.values()),
        "chiasm_name": chiasm_name or "",
        "chiasm_dmax": ch_dmax,
        "brainstem_name": brainstem_name or "",
        "brainstem_dmax": bs_dmax,
    }


# ---------------------------------------------------------------------------
# 5) Hauptprogramm
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------
# KONFIGURATION: True = Debug Mode (alle AUSSER 9-stellige IDs)
#                False = Live Mode (nur 9-stellige numerische IDs)
# ---------------------------------------------------------------
DEBUG_MODE = False  # <-- Auf True setzen für Debug Mode

# ---------------------------------------------------------------
# KONFIGURATION: Ausführliche Ausgaben (Fortschritt, Tabellen)
# ---------------------------------------------------------------
VERBOSE = False  # <-- Auf True setzen für mehr Details

# ---------------------------------------------------------------
# KONFIGURATION: True = Test-Daten (_TestPDF), False = Echte Daten
# ---------------------------------------------------------------
USE_TEST_DATA = False  # <-- Auf True setzen für Test-Ordner

# Echter PDF-Ordner (wenn USE_TEST_DATA = False)
REAL_PDF_DIR = Path(r"C:\Users\Aria\Desktop\ElementsTreatPar_PDF")


# ---------------------------------------------------------------------------
# Vollständigkeitsanalyse nach PDF-Version und Optimizer
# ---------------------------------------------------------------------------

def _print_completeness(df_ptv, df_plan):
    """Druckt mittlere Vollständigkeitsrate pro AppVersion × Optimizer."""
    if df_ptv.empty:
        return

    PTV_PARAMS = ['CI', 'GI', 'Prescribed dose',
                  'Actual dose for prescribed coverage', 'Local V12Gy', 'Max dose relation']
    PLAN_PARAMS = ['Global V12Gy', 'CI volume averaged', 'Monitor units/1000']

    def fmt_ver(v):
        s = str(v).strip() if pd.notna(v) else ""
        return s if s else "?"

    df = df_ptv.copy()
    df['_ver'] = df['ApplicationVersion'].apply(fmt_ver)

    HDR = f"{'AppVersion':<22} {'Optimizer':<14} {'N':>5}   {'Ø Vollst. PTV':>14}   {'Ø Vollst. Plan':>14}"
    SEP = "-" * len(HDR)
    print("\n" + "=" * len(HDR))
    print("PDF-VOLLSTÄNDIGKEITSANALYSE")
    print("=" * len(HDR))
    print(HDR)
    print(SEP)

    # Plan-lookup: (ver, optimizer) -> completeness %
    plan_completeness = {}
    if df_plan is not None and not df_plan.empty:
        df_pl = df_plan.copy()
        df_pl['_ver'] = df_pl['ApplicationVersion'].apply(fmt_ver)
        for (ver, opt), grp in df_pl.groupby(['_ver', 'Optimizer'], sort=True):
            avail = [c for c in PLAN_PARAMS if c in grp.columns]
            pct = (grp[avail].notna().mean(axis=1).mean() * 100) if avail else 0.0
            plan_completeness[(ver, opt)] = (pct, len(grp))

    for (ver, opt), grp in df.groupby(['_ver', 'Optimizer'], sort=True):
        n = len(grp)
        avail = [c for c in PTV_PARAMS if c in grp.columns]
        ptv_pct = (grp[avail].notna().mean(axis=1).mean() * 100) if avail else 0.0
        pl_pct, pl_n = plan_completeness.get((ver, opt), (float('nan'), 0))
        pl_str = f"{pl_pct:>12.0f}%  (n={pl_n})" if not np.isnan(pl_pct) else f"{'–':>14}"
        print(f"{ver:<22} {opt:<14} {n:>5}   {ptv_pct:>12.0f}%   {pl_str}")

    print("=" * len(HDR))
    print(f"  PTV-Parameter: {', '.join(PTV_PARAMS)}")
    print(f"  Plan-Parameter: {', '.join(PLAN_PARAMS)}")


def main():
    # UTF-8 Ausgabe erzwingen (Windows-Terminal kann sonst Sonderzeichen nicht darstellen)
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    base_dir = Path(__file__).parent

    # Argumente parsen (--debug überschreibt DEBUG_MODE)
    args = sys.argv[1:]
    debug_mode = DEBUG_MODE or ("--debug" in args)
    if "--debug" in args:
        args.remove("--debug")

    if len(args) > 0:
        pdf_dir = Path(args[0]).resolve()
    elif USE_TEST_DATA:
        pdf_dir = base_dir / "_TestPDF"
    else:
        pdf_dir = REAL_PDF_DIR
    
    mode_str = "DEBUG" if debug_mode else "LIVE"
    print(f"[{mode_str}] PDF: {pdf_dir}")

    # --- PDF-Dateien scannen ---
    pdf_files = []
    for root, dirs, files in os.walk(str(pdf_dir)):
        for fn in files:
            if fn.lower().endswith(".pdf"):
                pdf_files.append(os.path.join(root, fn))
    if VERBOSE:
        print(f"  {len(pdf_files)} PDF-Dateien gefunden.")

    # --- Parsen ---
    all_parsed = []
    parse_errors = 0
    for f in tqdm(pdf_files, desc="Lese PDFs"):
        parsed = parse_treat_par_pdf(f)
        if parsed is None:
            parse_errors += 1
            continue
        all_parsed.append(parsed)

    # Filter je nach Modus
    before = len(all_parsed)
    if debug_mode:
        all_parsed = [p for p in all_parsed if not re.fullmatch(r'\d{9}', p["patient_id"].strip())]
    else:
        all_parsed = [p for p in all_parsed if re.fullmatch(r'\d{9}', p["patient_id"].strip())]
    if VERBOSE:
        print(f"  {before - len(all_parsed)} gefiltert, {len(all_parsed)} Pläne ({parse_errors} Fehler).")

    # --- PTV-Level CSV ---
    ptv_rows = []
    for p in all_parsed:
        pid = p["patient_id"].strip()
        pname = p["patient_name"]
        plan_name = p["plan_name"]
        optimizer = p["optimizer"]
        n_ptvs = len(p["ptv_details"])
        fractions = p["fractions"]
        creation = p["creation_date"]

        for ptv in p["ptv_details"]:
            vol = ptv.get("PTVvolume")
            if vol is not None and vol < 0.4:
                vg = "small(<0.4cc)"
            elif vol is not None and vol <= 1.0:
                vg = "medium(0.4-1cc)"
            elif vol is not None:
                vg = "big(>1cc)"
            else:
                vg = "unknown"

            row = {
                "Patient Id": pid,
                "Patient name": pname,
                "Plan name": plan_name,
                "Optimizer": optimizer,
                "PTVname": ptv.get("PTVname", ""),
                "PTVvolume": ptv.get("PTVvolume"),
                "Max diameter": ptv.get("Max diameter"),
                "CI": ptv.get("CI"),
                "GI": ptv.get("GI"),
                "Sphericity": None,
                "Convexity": None,
                "Distance to closest PTV": None,
                "Distance to isocenter": None,
                "Prescribed dose": ptv.get("Prescribed dose"),
                "Actual dose for prescribed coverage": ptv.get("_coverage_dose") or ptv.get("Actual dose for prescribed coverage"),
                "Prescribed coverage": ptv.get("Prescribed coverage"),
                "Actual coverage for prescription dose": ptv.get("_actual_cov_pct") or ptv.get("Actual coverage for prescription dose"),
                "PTV count": n_ptvs,
                "Number of fractions": fractions,
                "Has isodose line prescription": None,
                "Local V5Gy": None,
                "Local V5Gy is shared": None,
                "Local V10Gy": None,
                "Local V10Gy is shared": None,
                "Local V12Gy": ptv.get("_local_v_cc"),
                "Local V18Gy": None,
                "Max dose relation": ptv.get("_max_dose_rel"),
                "Creation": creation,
                "VolumeGroup": vg,
                "ApprovalStatus": p["approval_status"],
                "ApplicationVersion": p["app_version"],
                "Chiasm_Name": p["chiasm_name"],
                "Chiasm_Dmax_Gy": p["chiasm_dmax"],
                "Chiasm_D005cc_Gy": np.nan,
                "Chiasm_D003cc_Gy": np.nan,
                "Brainstem_Name": p["brainstem_name"],
                "Brainstem_Dmax_Gy": p["brainstem_dmax"],
                "Brainstem_D005cc_Gy": np.nan,
                "Brainstem_D003cc_Gy": np.nan,
            }
            ptv_rows.append(row)

    df_ptv = pd.DataFrame(ptv_rows)

    # --- Plan-Level CSV ---
    plan_rows = []
    for p in all_parsed:
        pid = p["patient_id"].strip()
        pname = p["patient_name"]
        plan_name = p["plan_name"]
        optimizer = p["optimizer"]
        n_ptvs = len(p["ptv_details"])

        cis = [ptv["CI"] for ptv in p["ptv_details"] if ptv.get("CI") is not None]
        gis = [ptv["GI"] for ptv in p["ptv_details"] if ptv.get("GI") is not None]
        presc = [ptv["Prescribed dose"] for ptv in p["ptv_details"] if ptv.get("Prescribed dose") is not None]

        plan_rows.append({
            "Patient Id": pid,
            "Patient name": pname,
            "Plan name": plan_name,
            "Optimizer": optimizer,
            "NrOfPTVs": n_ptvs,
            "Nr. of table angles": None,
            "Nr. of arcs": p["n_arcs"],
            "NrOfFractions": p["fractions"],
            "avgPrescription": round(np.mean(presc), 1) if presc else None,
            "Cumulative PTV volume": p["cum_ptv_vol"],
            "Global V5Gy": None,
            "Global V10Gy": None,
            "Global V12Gy": p.get("global_v12_cc"),
            "CI volume averaged": p["ci_vol_avg"],
            "GI volume averaged": p["gi_vol_avg"],
            "CI mean": round(np.mean(cis), 3) if cis else None,
            "GI mean": round(np.mean(gis), 3) if gis else None,
            "Monitor units/1000": p["total_mu"],
            "AES": None,
            "MCS": p["mcs"],
            "Machine": p["machine"],
            "Energy": "",
            "ApprovalStatus": p["approval_status"],
            "ApplicationName": p["app_name"],
            "ApplicationVersion": p["app_version"],
            "Has isodose line prescription": None,
            "Creation": p["creation_date"],
            "Chiasm_Name": p["chiasm_name"],
            "Chiasm_Dmax_Gy": p["chiasm_dmax"],
            "Chiasm_D005cc_Gy": np.nan,
            "Chiasm_D003cc_Gy": np.nan,
            "Brainstem_Name": p["brainstem_name"],
            "Brainstem_Dmax_Gy": p["brainstem_dmax"],
            "Brainstem_D005cc_Gy": np.nan,
            "Brainstem_D003cc_Gy": np.nan,
        })

    df_plan = pd.DataFrame(plan_rows)

    # --- Vollständigkeitsanalyse ---
    _print_completeness(df_ptv, df_plan)

    # --- Speichern ---
    ptv_csv = base_dir / "pdf_data_ptv.csv"
    plan_csv = base_dir / "pdf_data_plan.csv"
    df_ptv.to_csv(str(ptv_csv), index=False)
    df_plan.to_csv(str(plan_csv), index=False)

    if VERBOSE:
        print(f"\nGespeichert: {ptv_csv.name}, {plan_csv.name}")
        print(f"Pläne: {len(df_plan)}, PTVs: {len(df_ptv)}")


if __name__ == "__main__":
    main()
