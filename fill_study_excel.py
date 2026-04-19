"""fill_study_excel.py – Füllt Studien-Excel-Struktur aus Brainlab PDFs

Struktur des Outputs (spiegelt Master_study.xlsx):
  - PLAN-ZEILE  : *Plan + Patientenname (klinische Felder bleiben leer → manuell)
  - MET-ZEILEN  : eine pro PTV, alle BM-Stat-Felder soweit aus PDF extrahierbar

Belegung (direkt aus PDF-Tabellen):
  *GTV                   = Met{N}  (Brainlab-Konvention)
  TotalDose              = Prescribed Dose [Gy]  aus PRESCRIPTION-Tabelle
  PTV-Volumen            = Volume [cm³]           aus PRESCRIPTION-Tabelle
  PTV-D2% (near max)     = Max Dose [Gy]          aus TREATED METASTASES
  PTV-D50%               = Mean Dose [Gy]         aus TREATED METASTASES
  PTV-D98% (near min)    = Min Dose [Gy]          aus TREATED METASTASES
  PTV-Coverage           = Max. Dose Relation [%] aus TREATED METASTASES
  PTV-CI / GI            = CI / GI                aus TREATED METASTASES
  local-V12Gy            = nur in neueren PDF-Versionen verfügbar
  GTV-Volumen / Margin / DistIso → leer (nicht in PDF verfügbar)

Aufruf:
  python fill_study_excel.py
  python fill_study_excel.py --pdf "C:\\Pfad\\PDFs" --out "C:\\Pfad\\output.xlsx"
"""

import sys
import os
import re
import argparse
import math
from pathlib import Path
from datetime import datetime

import pandas as pd
import numpy as np
from tqdm import tqdm

_HERE = Path(__file__).parent
sys.path.insert(0, str(_HERE))
from parse_pdf_reports import parse_treat_par_pdf
# DICOM-Imports werden lazy geladen (nur wenn --dicom verwendet wird)

# ── Konfiguration ─────────────────────────────────────────────────────────────
DEFAULT_PDF_DIR  = Path(r"Z:\Projekte_Github\brainlab-srs-analytics\testPDF")
DEFAULT_DICOM_DIR = Path(r"Z:\Projekte_Github\brainlab-srs-analytics\testDICOM")
DEFAULT_OUTPUT  = _HERE / "study_export.xlsx"


# ── Spalten-Template (exakt wie Master_study.xlsx) ───────────────────────────
PLAN_COLS = [
    "PatID ClinicA", "PatID Clinic", "Alter bei 1. MBM-SRT",
    "Diagnose Code (1=AdenoLunge; 2 Melanom; 3=MammaCa; 4=sonstiges)",
    "Diagnose Freitext", "Diagnosedatum (Primarius)", "Diagnosedatum (BM)",
    "dsGPA-Score", "1. MBM-SRT",
    "Immuntherapie erhalten (0=Nein; 1=Ja)", "Vorherige WBRT? (0=Nein; 1=Ja)",
    "vorherige Cranial-STX (Met oder MetBett)?\n(0=Nein; 1=Ja)",
    "Extrakranielle Metastasen zum Zeitpunkt RT (0=Nein; 1=Ja)",
    "simultane MetBett? (0=Nein; 1=Ja)", "Karnofsky-Index zur 1. MBM-SRT",
    "Anzahl BM bei 1. MBM-SRT", "Anzahl Iso bei 1. MBM-SRT",
    "*Plan", "*PTV", "*GTV",
    "*Approvaldate",
    "BM-Stat:\nNummer",
    "BM-Stat:\nTotalDose [Gy]", "BM-Stat:\nFractions",
    "BM-Stat:\nGTV-Volumen [cc]",
    "**BM-Stat:\nPTV-Margin [mm]",
    "BM-Stat:\nPTV-Volumen [cc]",
    "BM-Stat:\nDistIso \n[mm]",
    "BM-Stat:\nGTV-D98% [Gy]\n(near min)", "BM-Stat:\nGTV-D50% [Gy]\n",
    "BM-Stat:\nPTV-D98% [Gy]\n(near min)", "BM-Stat:\nPTV-D50% [Gy]\n",
    "BM-Stat:\nPTV-D2% [Gy]\n(near max)",
    "BM-Stat:\nPTV-Coverage [%]\n",
    "BM-Stat:\nPTV-CI\n", "BM-Stat:\nPTV-GI\n",
    "BM-Stat:\nlocal-V12Gy [cc]_Treat", "BM-Stat:\nlocal-V10Gy [cc]_Treat",
    "**BM-Stat:\nlocal-V12Gy [cc]_GTV", "**BM-Stat:\nlocal-V10Gy [cc]_GTV",
    "\nglobal-V8Gy [cc]",
    "*Seitigkeit", "*Lokalisation",
    "BM-Stat:\nRadionekrose (RN)\n(0=Nein; 1=Ja; 2=Verdacht)",
    "BM-Stat:\nDatum der RN", "BM-Stat: Symptome der RN",
    "BM-Stat:\nTherapie der RN",
    "BM-Stat:\nLokalrezidiv (0=Nein; 1=Ja)", "BM-Stat:\nLokalrezidiv-Datum",
    "*Kommentare",
    "\ndistant brain failure\n(0=Nein; 1=Ja)", "distant brain failure\nDatum",
    " Salvage RT\n(0=Nein; 1=Ja)", " Salvage RT -> WBRT\n(0=Nein; 1=Ja)",
    " Salvage RT -> SRT distant\n(0=Nein; 1=Ja)",
    " Salvage RT -> SRT Lokalrezidiv\n(0=Nein; 1=Ja)",
    "Todesdatum oder letztes Follow-up", "Tod\n(0=Nein; 1=Ja)",
]


# ── Hilfsfunktionen ──────────────────────────────────────────────────────────

def _parse_date(s: str):
    """Wandelt Brainlab-Datumsstring in YYYYMMDD-Integer um."""
    if not s:
        return None
    for fmt in ("%d-%b-%Y %H:%M:%S", "%d-%b-%Y", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return int(datetime.strptime(s.strip(), fmt).strftime("%Y%m%d"))
        except ValueError:
            continue
    return None


def _to_float(s):
    try:
        return float(str(s).replace(',', '.').strip())
    except (ValueError, TypeError):
        return None


def _gtv_from_ptv(ptv_name: str) -> str:
    """GTV-Name: Nummer aus PTV extrahieren → Met{N}.
    PTV1 → Met1,  PTV1 PTV → Met1,  PTV2 xyz → Met2"""
    m = re.search(r'(\d+)', str(ptv_name))
    return f"Met{m.group(1)}" if m else ""


# ── Direkte PDF-Tabellen-Parser ───────────────────────────────────────────────

def parse_pdf_tables(pdf_path: str) -> dict:
    """Liest PRESCRIPTION, TREATED METASTASES und OTHERS direkt aus PDF-Text.

    Unterstützt alle bekannten Brainlab-Versionen:
      v1.5: 8-Spalten-PRESCRIPTION, PTV-Typ-Marker in TREATED MET,
            Met-01-Benennung in OTHERS (ORGAN, 4 Werte + 3x-)
      v2.0: 9-Spalten-PRESCRIPTION (+Coverage Dose), Werte mit
            "D98% =" und "DMax =" in TREATED MET, GTV...-Namen in OTHERS (3 Werte)
      v4.5+: wie v2.0, aber Min+Coverage auf EINER Zeile, extra LocalV12-Spalte

    Keys im Result-Dict: erster Wortanteil des PTV-Namens (z.B. "PTV1", "PTV1R", "PTV01R")
    Sonderschlüssel: "_gtv_by_num" (GTV-Daten indexiert nach Nummer)
    """
    import fitz
    result = {}
    try:
        doc = fitz.open(pdf_path)
        full_text = "\n".join(page.get_text() for page in doc).replace('\xa0', ' ')
        doc.close()
    except Exception:
        return result

    presc_pos   = full_text.find('PRESCRIPTION')
    treated_pos = full_text.find('TREATED METASTASES')
    others_pos  = full_text.find('OTHERS')

    if presc_pos == -1:
        return result

    # ── PRESCRIPTION ─────────────────────────────────────────────────────────
    end_presc = treated_pos if treated_pos > presc_pos else len(full_text)
    presc_text = full_text[presc_pos:end_presc]

    # v2.0+: 9 Werte  (dose_fx | n_fx | presc_dose | coverage_dose | cov_vol% | volume | size)
    # v1.5:  8 Werte  (dose_fx | n_fx | presc_dose | presc_vol%    | volume   | size)
    # Key = erster Wortanteil des PTV-Namens für einfaches späteres Lookup
    re9 = re.compile(
        r'(PTV[^\n]+)\nPTV\n'
        r'([\d\.]+)\n(\d+)\n([\d\.]+)\n([\d\.]+)\n([\d\.]+)\n([\d\.]+)\n([\d\.]+)'
    )
    re8 = re.compile(
        r'(PTV[^\n]+)\nPTV\n'
        r'([\d\.]+)\n(\d+)\n([\d\.]+)\n([\d\.]+)\n([\d\.]+)\n([\d\.]+)'
    )
    matches9 = list(re9.finditer(presc_text))
    if matches9:
        for m in matches9:
            key = m.group(1).strip().split()[0]
            result.setdefault(key, {})
            result[key]['ptv_fullname']   = m.group(1).strip()
            result[key]['dose_per_fx']    = _to_float(m.group(2))
            result[key]['n_fractions']    = _to_float(m.group(3))
            result[key]['prescribed_dose']= _to_float(m.group(4))
            result[key]['coverage_dose']  = _to_float(m.group(5))
            result[key]['presc_vol_pct']  = _to_float(m.group(6))
            result[key]['volume_cc']      = _to_float(m.group(7))
            result[key]['size_cm']        = _to_float(m.group(8))
    else:
        for m in re8.finditer(presc_text):
            key = m.group(1).strip().split()[0]
            result.setdefault(key, {})
            result[key]['ptv_fullname']   = m.group(1).strip()
            result[key]['dose_per_fx']    = _to_float(m.group(2))
            result[key]['n_fractions']    = _to_float(m.group(3))
            result[key]['prescribed_dose']= _to_float(m.group(4))
            result[key]['presc_vol_pct']  = _to_float(m.group(5))
            result[key]['volume_cc']      = _to_float(m.group(6))
            result[key]['size_cm']        = _to_float(m.group(7))

    if treated_pos == -1:
        return result

    treated_text = full_text[treated_pos:]

    # ── TREATED METASTASES: Format-Erkennung ──────────────────────────────────
    # v1.5: PTV\w+\nPTV\n  (Typ-Marker "PTV" nach dem Namen)
    is_v15 = bool(re.search(r'PTV\w+\nPTV\n[\d\.]+\n[\d\.]+\n[\d\.]+\n[\d\.]+\n[\d\.]+\n[\d\.]+\n[\d\.]+', treated_text))
    # v4.5+: Min und Coverage auf EINER Zeile: "22.64 D98.5 % = 24.03"
    is_v45 = bool(re.search(r'\n[\d\.]+ D[\d\.]+ % = [\d\.]+\n', treated_text))

    if is_v15:
        # v1.5: {PTV}\nPTV\n{vol}\n{max}\n{mean}\n{min}\n{CI}\n{GI}\n{maxrel}
        for m in re.finditer(
            r'(PTV\w+)\nPTV\n'
            r'([\d\.]+)\n([\d\.]+)\n([\d\.]+)\n([\d\.]+)\n'
            r'([\d\.]+)\n([\d\.]+)\n([\d\.]+)',
            treated_text
        ):
            key = m.group(1)
            result.setdefault(key, {})
            result[key]['volume_cc']    = _to_float(m.group(2))
            result[key]['max_dose']     = _to_float(m.group(3))
            result[key]['mean_dose']    = _to_float(m.group(4))
            result[key]['min_dose']     = _to_float(m.group(5))
            result[key]['ci']           = _to_float(m.group(6))
            result[key]['gi']           = _to_float(m.group(7))
            result[key]['max_dose_rel'] = _to_float(m.group(8))

    elif is_v45:
        # v4.5+: {PTV}\n{min} D{x}% = {cov}\nV {p}Gy = {v%}\n{mean}\nD1%={des}\n{max}\n{CI}\n{GI}\n{localV12}\n{maxrel}
        ptv_keys = [k for k in result if not k.startswith('_')]
        for key in ptv_keys:
            pos = treated_text.find(key)
            if pos == -1:
                continue
            block = treated_text[pos:pos + 600]
            m = re.search(
                r'([\d\.]+)\s+D[\d\.]+ % = ([\d\.]+)\n'
                r'V [\d\.]+Gy = [\d\.]+\n'
                r'([\d\.]+)\n'
                r'D[\d\.]+ % = [\d\.]+\n'
                r'([\d\.]+)\n'
                r'([\d\.]+)\n([\d\.]+)\n'
                r'([\d\.]+)\n'
                r'([\d\.]+)',
                block
            )
            if m:
                result[key]['min_dose']      = _to_float(m.group(1))
                result[key]['coverage_dose'] = _to_float(m.group(2))
                result[key]['mean_dose']     = _to_float(m.group(3))
                result[key]['max_dose']      = _to_float(m.group(4))
                result[key]['ci']            = _to_float(m.group(5))
                result[key]['gi']            = _to_float(m.group(6))
                result[key]['local_v12']     = _to_float(m.group(7))
                result[key]['max_dose_rel']  = _to_float(m.group(8))

    else:
        # v2.0: {PTV}\n{min}\nD{x}% = {cov}\n{mean}\nDMax = {max_c}\n{max}\n{CI}\n{GI}\n{maxrel}
        # PTV-Name kann über zwei Zeilen gehen → Suche über ersten Wortanteil
        ptv_keys = [k for k in result if not k.startswith('_')]
        for key in ptv_keys:
            pos = treated_text.find(key)
            if pos == -1:
                continue
            block = treated_text[pos:pos + 500]
            m = re.search(
                r'([\d\.]+)\n'
                r'D[\d\.]+ % = ([\d\.]+)\n'
                r'([\d\.]+)\n'
                r'DMax = [\d\.]+\n'
                r'([\d\.]+)\n'
                r'([\d\.]+)\n([\d\.]+)\n'
                r'([\d\.]+)',
                block
            )
            if m:
                result[key]['min_dose']      = _to_float(m.group(1))
                result[key]['coverage_dose'] = _to_float(m.group(2))
                result[key]['mean_dose']     = _to_float(m.group(3))
                result[key]['max_dose']      = _to_float(m.group(4))
                result[key]['ci']            = _to_float(m.group(5))
                result[key]['gi']            = _to_float(m.group(6))
                result[key]['max_dose_rel']  = _to_float(m.group(7))

    # ── OTHERS (GTVs) ─────────────────────────────────────────────────────────
    gtv_by_num = {}
    if others_pos != -1:
        others_text = full_text[others_pos:]

        # v1.5: "Met 01\nORGAN\n{vol}\n{max}\n{mean}\n{min}\n-\n-\n-"
        #   Spaltenreihenfolge: volume, MAX, mean, min
        for m in re.finditer(
            r'(Met\s+\d+)\nORGAN\n([\d\.]+)\n([\d\.]+)\n([\d\.]+)\n([\d\.]+)',
            others_text
        ):
            num = int(re.search(r'\d+', m.group(1)).group())
            gtv_by_num[num] = {
                'gtv_name':      m.group(1).strip(),
                'gtv_vol_cc':    _to_float(m.group(2)),
                'gtv_max_dose':  _to_float(m.group(3)),
                'gtv_mean_dose': _to_float(m.group(4)),
                'gtv_min_dose':  _to_float(m.group(5)),
            }

        # v2.0+: "GTV1 R M.xxx\nORGAN\n{vol}\n{mean}\n{max}"
        #   Spaltenreihenfolge: volume, MEAN, max (kein min!)
        for m in re.finditer(
            r'(GTV[^\n]+)\nORGAN\n([\d\.]+)\n([\d\.]+)\n([\d\.]+)',
            others_text
        ):
            num_m = re.search(r'\d+', m.group(1))
            if not num_m:
                continue
            num = int(num_m.group())
            # Nicht überschreiben falls schon aus v1.5 vorhanden
            if num not in gtv_by_num:
                gtv_by_num[num] = {
                    'gtv_name':      m.group(1).strip(),
                    'gtv_vol_cc':    _to_float(m.group(2)),
                    'gtv_mean_dose': _to_float(m.group(3)),
                    'gtv_max_dose':  _to_float(m.group(4)),
                    'gtv_min_dose':  None,
                }

    result['_gtv_by_num'] = gtv_by_num
    return result


# ── Zeilen erzeugen ──────────────────────────────────────────────────────────

def _empty_row():
    return {c: None for c in PLAN_COLS}


def make_plan_row(p: dict) -> dict:
    """Erstellt die Plan-Zeile (klinische Felder leer, nur Brainlab-Info).
    Verträgt sowohl PDF-Dicts (plan_name, patient_name) als auch
    DICOM-Dicts (cbl_name, patient_id unterschiedlich benannt).
    """
    row = _empty_row()
    plan_name = p.get("plan_name") or p.get("cbl_name", "")
    pat_name  = p.get("patient_name") or p.get("patient_id", "")
    pat_id    = p.get("patient_id", "")
    row["*Plan"]         = plan_name
    row["*Approvaldate"] = _parse_date(p.get("creation_date", ""))
    row["*Kommentare"]   = f"[auto] {pat_name} | ID={pat_id}" if pat_name or pat_id else "[auto]"
    return row


def make_met_rows(p: dict, pdf_path: str) -> list:
    """Erstellt eine Met-Zeile pro PTV – nutzt direkt geparste PDF-Tabellen."""
    rows = []
    plan_name = p["plan_name"]
    approval  = _parse_date(p.get("creation_date", ""))
    tbl       = parse_pdf_tables(pdf_path)

    # PTV-Namen: aus parse_pdf_reports oder direkt aus Tabelle
    ptv_names = [ptv.get("PTVname", "") for ptv in p.get("ptv_details", [])]
    if not ptv_names:
        ptv_names = list(tbl.keys())

    for i, ptv_name in enumerate(ptv_names, start=1):
        row = _empty_row()

        # Kurzname "PTV1" aus "PTV1 PTV" extrahieren
        short = re.split(r'\s+', str(ptv_name).strip())[0]
        data  = tbl.get(short, {})

        row["*Plan"]         = plan_name
        row["*PTV"]          = ptv_name
        row["*Approvaldate"] = approval
        row["BM-Stat:\nNummer"] = i

        # Fraktionen aus Tabelle, Fallback: parse_pdf
        row["BM-Stat:\nFractions"] = data.get("n_fractions") or p.get("fractions")

        # ── PRESCRIPTION-Daten ────────────────────────────────────────────────
        ptv_vol = data.get("volume_cc")
        row["BM-Stat:\nTotalDose [Gy]"]   = data.get("prescribed_dose")
        row["BM-Stat:\nPTV-Volumen [cc]"] = ptv_vol

        # ── TREATED METASTASES ────────────────────────────────────────────────
        max_dose  = data.get("max_dose")
        mean_dose = data.get("mean_dose")
        min_dose  = data.get("min_dose")
        ci        = data.get("ci")
        gi        = data.get("gi")
        max_rel   = data.get("max_dose_rel")

        # Fallbacks aus parse_pdf_reports (für neuere PDF-Versionen)
        ptv_d = next((x for x in p.get("ptv_details", [])
                      if str(x.get("PTVname", "")).startswith(short)), {})
        if ci  is None: ci  = ptv_d.get("CI")
        if gi  is None: gi  = ptv_d.get("GI")
        # D2%: Max Dose; Fallback: _nearmax_dose aus parse_pdf
        d2  = max_dose  if max_dose  is not None else ptv_d.get("_nearmax_dose")
        # D98%: Min Dose; kein sinnvoller Fallback
        d98 = min_dose
        # D50%: Mean Dose; Fallback: _mean_dose aus parse_pdf
        d50 = mean_dose if mean_dose is not None else ptv_d.get("_mean_dose")

        row["BM-Stat:\nPTV-D2% [Gy]\n(near max)"]  = d2
        row["BM-Stat:\nPTV-D98% [Gy]\n(near min)"]  = d98
        row["BM-Stat:\nPTV-D50% [Gy]\n"]            = d50
        row["BM-Stat:\nPTV-Coverage [%]\n"]          = max_rel
        row["BM-Stat:\nPTV-CI\n"]                    = ci
        row["BM-Stat:\nPTV-GI\n"]                    = gi
        row["BM-Stat:\nlocal-V12Gy [cc]_Treat"]      = ptv_d.get("_local_v_cc")

        # ── GTV aus OTHERS-Tabelle ────────────────────────────────────────────
        gtv_by_num = tbl.get("_gtv_by_num", {})
        ptv_num_m  = re.search(r'(\d+)', short)
        gtv_data   = gtv_by_num.get(int(ptv_num_m.group(1))) if ptv_num_m else None

        if gtv_data:
            gtv_name   = gtv_data['gtv_name']
            gtv_vol    = gtv_data['gtv_vol_cc']
            row["*GTV"]                              = gtv_name
            row["BM-Stat:\nGTV-Volumen [cc]"]        = gtv_vol
            row["BM-Stat:\nGTV-D98% [Gy]\n(near min)"] = gtv_data['gtv_min_dose']
            row["BM-Stat:\nGTV-D50% [Gy]\n"]         = gtv_data['gtv_mean_dose']
            # Margin aus Kugelformel (V in cm³, Ergebnis in mm)
            if ptv_vol is not None and gtv_vol is not None and gtv_vol > 0:
                r_ptv = (3 * ptv_vol / (4 * math.pi)) ** (1/3)
                r_gtv = (3 * gtv_vol / (4 * math.pi)) ** (1/3)
                margin_mm = (r_ptv - r_gtv) * 10  # cm → mm
                row["**BM-Stat:\nPTV-Margin [mm]"] = round(margin_mm, 1)
        else:
            row["*GTV"] = _gtv_from_ptv(ptv_name)

        rows.append(row)
    return rows


# ── DICOM-Modus ──────────────────────────────────────────────────────────────

def make_met_rows_dicom(p: dict) -> list:
    """Erzeugt Studien-Excel-Zeilen aus DICOM-Plandaten (parse_planeval_dicom).

    Felder werden aus ptv_details und struct_map des Brainlab PlanAnalytics-DICOMs
    extrahiert. GTV-Matching erfolgt per P→G-Namenskonvention (wie in create_excel).
    Felder ohne DICOM-Äquivalent (z.B. D50%) bleiben leer.
    """
    from enrich_bestrahlungsdaten import find_gtv_for_ptv, compute_margin_mm

    rows = []
    plan_name  = p.get("cbl_name", p.get("plan_name", ""))
    approval   = _parse_date(p.get("creation_date", ""))
    struct_map = p.get("struct_map", {})
    fractions  = p.get("fractions")

    for i, ptv in enumerate(p.get("ptv_details", []), start=1):
        row = _empty_row()
        ptv_name = ptv.get("PTVname", "")
        ptv_vol  = ptv.get("PTVvolume")  # cc

        row["*Plan"]         = plan_name
        row["*PTV"]          = ptv_name
        row["*Approvaldate"] = approval
        row["BM-Stat:\nNummer"] = i

        row["BM-Stat:\nFractions"]        = ptv.get("Number of fractions") or fractions
        row["BM-Stat:\nTotalDose [Gy]"]   = ptv.get("Prescribed dose")
        row["BM-Stat:\nPTV-Volumen [cc]"] = ptv_vol

        # Dose metrics (DICOM hat kein D2%/D50% direkt)
        row["BM-Stat:\nPTV-D98% [Gy]\n(near min)"] = ptv.get("Actual dose for prescribed coverage")
        row["BM-Stat:\nPTV-D50% [Gy]\n"]           = None  # mean dose nicht im DICOM-PTV-Dict
        row["BM-Stat:\nPTV-D2% [Gy]\n(near max)"]  = None  # Dmax nicht im DICOM-PTV-Dict
        cov = ptv.get("Actual coverage for prescription dose")
        row["BM-Stat:\nPTV-Coverage [%]\n"] = round(cov * 100, 1) if cov is not None else None
        row["BM-Stat:\nPTV-CI\n"] = ptv.get("CI")
        row["BM-Stat:\nPTV-GI\n"] = ptv.get("GI")
        row["BM-Stat:\nlocal-V12Gy [cc]_Treat"] = ptv.get("Local V12Gy")

        # GTV per P→G-Namenskonvention
        ptv_vol_mm3 = ptv_vol * 1000 if ptv_vol is not None else None
        gtv_name, _, gtv_vol_mm3 = find_gtv_for_ptv(ptv_name, struct_map, ptv_vol_mm3)
        gtv_vol_cc = round(gtv_vol_mm3 / 1000, 4) if gtv_vol_mm3 is not None else None
        margin     = compute_margin_mm(ptv_vol_mm3, gtv_vol_mm3)
        show_gtv   = margin is not None and margin <= 6

        if show_gtv:
            row["*GTV"]                              = gtv_name
            row["BM-Stat:\nGTV-Volumen [cc]"]        = gtv_vol_cc
            row["**BM-Stat:\nPTV-Margin [mm]"]       = round(margin, 1)
        else:
            row["*GTV"] = _gtv_from_ptv(ptv_name)

        rows.append(row)
    return rows


# ── Hauptprogramm ─────────────────────────────────────────────────────────────

def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        description="Brainlab PDFs oder DICOMs → Studien-Excel"
    )
    parser.add_argument("--pdf",   type=Path, default=None,
                        help="Ordner mit Brainlab TreatPar PDFs (Standard-Quelle)")
    parser.add_argument("--dicom", type=Path, default=None,
                        help="Ordner mit Brainlab PlanAnalytics DICOMs (Alternative zu --pdf)")
    parser.add_argument("--out",   type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    use_dicom = args.dicom is not None
    out_path  = args.out
    all_rows  = []
    errors    = 0

    if use_dicom:
        # ── DICOM-Modus ───────────────────────────────────────────────────────
        from create_excel import load_dicom_plans
        dicom_dir = args.dicom if args.dicom is not None else DEFAULT_DICOM_DIR
        if not dicom_dir.exists():
            print(f"DICOM-Ordner nicht gefunden: {dicom_dir}")
            sys.exit(1)
        print(f"Quelle: DICOM  ({dicom_dir})")
        plans = load_dicom_plans(dicom_dir, debug=False)
        for p in tqdm(plans, desc="Verarbeite Pläne"):
            all_rows.append(make_plan_row(p))
            all_rows.extend(make_met_rows_dicom(p))
        summary = (
            "  TotalDose  = Prescribed Dose [Gy]               (aus DICOM)\n"
            "  PTV-D98%   = Actual dose for prescribed coverage (aus DICOM)\n"
            "  PTV-Coverage = Actual coverage for presc. dose  (aus DICOM)\n"
            "  PTV-CI/GI  = CI / GI                            (aus DICOM)\n"
            "  GTV-Volumen, Margin: per P→G-Namenskonvention   (aus DICOM)\n"
            "  PTV-D2%, PTV-D50%: nicht im DICOM-PTV-Dict → leer"
        )
    else:
        # ── PDF-Modus (Standard) ──────────────────────────────────────────────
        pdf_dir = args.pdf if args.pdf is not None else DEFAULT_PDF_DIR
        if not pdf_dir.exists():
            print(f"PDF-Ordner nicht gefunden: {pdf_dir}")
            sys.exit(1)
        pdf_files = [
            os.path.join(r, f)
            for r, _, files in os.walk(str(pdf_dir))
            for f in files if f.lower().endswith(".pdf")
        ]
        print(f"Gefunden: {len(pdf_files)} PDF(s) in {pdf_dir}")
        for f in tqdm(pdf_files, desc="Lese PDFs"):
            p = parse_treat_par_pdf(f)
            if p is None:
                print(f"  [FEHLER] {Path(f).name}")
                errors += 1
                continue
            n_ptv = len(p.get("ptv_details", []))
            tbl   = parse_pdf_tables(f)
            if not n_ptv and tbl:
                n_ptv = len(tbl)
            print(f"  {Path(f).name}: {p['plan_name']} | {n_ptv} PTVs | "
                  f"v{p.get('app_version','?')} | {p.get('fractions','?')} Fx")
            all_rows.append(make_plan_row(p))
            all_rows.extend(make_met_rows(p, f))
        summary = (
            "  TotalDose  = Prescribed Dose [Gy]  (aus PRESCRIPTION-Tabelle)\n"
            "  PTV-D2%    = Max Dose [Gy]          (aus TREATED METASTASES)\n"
            "  PTV-D98%   = Min Dose [Gy]          (aus TREATED METASTASES)\n"
            "  PTV-D50%   = Mean Dose [Gy]         (aus TREATED METASTASES)\n"
            "  Coverage   = Max. Dose Relation [%] (aus TREATED METASTASES)\n"
            "  *GTV       = echte GTV-Namen + Volumen aus OTHERS-Tabelle\n"
            "  Margin     = Kugelformel (r_PTV - r_GTV) × 10 mm"
        )

    if not all_rows:
        print("Keine Daten. Abbruch.")
        sys.exit(1)

    df = pd.DataFrame(all_rows, columns=PLAN_COLS)

    print(f"\nSchreibe {out_path}  ({len(df)} Zeilen)...")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Studie", index=False)
        ws = writer.sheets["Studie"]
        ws.freeze_panes = "A2"
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

    print(f"Fertig → {out_path}")
    print()
    print("Belegung:")
    print(summary)
    if errors:
        print(f"  {errors} Datei(en) konnten nicht gelesen werden.")


if __name__ == "__main__":
    main()
