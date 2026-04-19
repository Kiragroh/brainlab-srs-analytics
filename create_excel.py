"""create_excel.py – Standalone Export: DICOM + PDF → Excel

Erstellt eine Excel-Datei (Plan / PTV / GTV) ohne externe Masterliste.
Funktioniert mit:
  - Nur DICOM  (Brainlab Elements PlanAnalytics .dcm)
  - Nur PDF    (Brainlab Elements Treatment Report .pdf)
  - DICOM + PDF (DICOM bevorzugt, PDF als Fallback)

Verwendung:
    python create_excel.py                          # Standardpfade
    python create_excel.py --dicom C:\\Pfad\\DICOMs
    python create_excel.py --pdf   C:\\Pfad\\PDFs
    python create_excel.py --dicom C:\\... --pdf C:\\...
    python create_excel.py --debug                  # alle IDs (auch Test)
"""

import sys
import os
import re
import math
import argparse
import pandas as pd
import numpy as np
from pathlib import Path
from tqdm import tqdm

# ── Parser-Module importieren ────────────────────────────────────────────────
_HERE = Path(__file__).parent
sys.path.insert(0, str(_HERE))

from enrich_bestrahlungsdaten import (
    parse_planeval_dicom,
    register_brainlab_private_dict,
    find_oar_by_patterns,
    extract_oar_metrics,
    find_gtv_for_ptv,
    compute_margin_mm,
    get_optimizer,
    CHIASM_PATTERNS,
    BRAINSTEM_PATTERNS,
)
from parse_pdf_reports import (
    parse_treat_par_pdf,
    CHIASM_PATS,
    BRAINSTEM_PATS,
    find_oar as pdf_find_oar,
)

# ── Konfiguration (Standardpfade – per CLI überschreibbar) ───────────────────
DEFAULT_DICOM_DIR = Path(r"C:\Users\Aria\Desktop\ElementsTreatPar_DICOM")
DEFAULT_PDF_DIR   = Path(r"C:\Users\Aria\Desktop\ElementsTreatPar_PDF")
OUTPUT_EXCEL      = _HERE / "export_standalone.xlsx"

# Nur 9-stellige numerische Patienten-IDs (False = alle IDs)
LIVE_FILTER = True


# ── Hilfsfunktionen ──────────────────────────────────────────────────────────

def _volume_group(vol_cc):
    if vol_cc is None:
        return "unknown"
    if vol_cc < 0.4:
        return "small(<0.4cc)"
    if vol_cc <= 1.0:
        return "medium(0.4-1cc)"
    return "big(>1cc)"


def _is_valid_pid(pid: str, debug: bool) -> bool:
    is_9digit = bool(re.fullmatch(r'\d{9}', pid.strip()))
    return (not is_9digit) if debug else is_9digit


# ── DICOM Laden ──────────────────────────────────────────────────────────────

def load_dicom_plans(dicom_dir: Path, debug: bool) -> list:
    if not dicom_dir or not dicom_dir.exists():
        return []

    # Brainlab Dictionary optional registrieren
    csv_path = _HERE / "brainlab_dictionary.csv"
    csv_path = csv_path if csv_path.exists() else None
    if csv_path:
        try:
            register_brainlab_private_dict(str(csv_path))
        except Exception:
            pass

    dcm_files = [
        os.path.join(r, f)
        for r, _, files in os.walk(str(dicom_dir))
        for f in files if f.lower().endswith(".dcm")
    ]
    print(f"  DICOM: {len(dcm_files)} Dateien in {dicom_dir}")

    raw = []
    errors = 0
    for f in tqdm(dcm_files, desc="Lese DICOMs"):
        p = parse_planeval_dicom(f)
        if p is None:
            errors += 1
            continue
        if p["patient_id"].strip() and p["cbl_name"].strip():
            raw.append(p)

    raw = [p for p in raw if _is_valid_pid(p["patient_id"], debug)]

    # Deduplizieren: (pid, plan) → meiste OARs
    seen = {}
    for p in raw:
        key = (p["patient_id"].strip(), p["cbl_name"].strip())
        if key not in seen or len(p["oar_data"]) > len(seen[key]["oar_data"]):
            seen[key] = p
    plans = list(seen.values())
    print(f"  DICOM: {len(plans)} eindeutige Pläne ({errors} Fehler).")
    return plans


# ── PDF Laden ────────────────────────────────────────────────────────────────

def load_pdf_plans(pdf_dir: Path, debug: bool) -> list:
    if not pdf_dir or not pdf_dir.exists():
        return []

    pdf_files = [
        os.path.join(r, f)
        for r, _, files in os.walk(str(pdf_dir))
        for f in files if f.lower().endswith(".pdf")
    ]
    print(f"  PDF: {len(pdf_files)} Dateien in {pdf_dir}")

    raw = []
    errors = 0
    for f in tqdm(pdf_files, desc="Lese PDFs"):
        p = parse_treat_par_pdf(f)
        if p is None:
            errors += 1
            continue
        raw.append(p)

    raw = [p for p in raw if _is_valid_pid(p["patient_id"], debug)]
    print(f"  PDF: {len(raw)} Pläne ({errors} Fehler).")
    return raw


# ── DataFrame-Bau: DICOM ─────────────────────────────────────────────────────

def build_dicom_dfs(plans: list):
    ptv_rows, plan_rows, gtv_rows = [], [], []

    for p in plans:
        pid       = p["patient_id"].strip()
        pname     = p["patient_name"]
        plan_name = p["cbl_name"].strip()
        opt       = get_optimizer(p)
        n_ptvs    = len(p["ptv_details"])
        struct_map = p.get("struct_map", {})

        chiasm_name = find_oar_by_patterns(p["oar_data"], CHIASM_PATTERNS)
        bs_name     = find_oar_by_patterns(p["oar_data"], BRAINSTEM_PATTERNS)
        ch_dmax, ch_d005, ch_d003 = extract_oar_metrics(p["oar_data"], chiasm_name)
        bs_dmax, bs_d005, bs_d003 = extract_oar_metrics(p["oar_data"], bs_name)

        cis   = [ptv["CI"]             for ptv in p["ptv_details"] if ptv.get("CI")             is not None]
        gis   = [ptv["GI"]             for ptv in p["ptv_details"] if ptv.get("GI")             is not None]
        presc = [ptv["Prescribed dose"] for ptv in p["ptv_details"] if ptv.get("Prescribed dose") is not None]

        plan_rows.append({
            "Source": "DICOM",
            "Patient Id": pid, "Patient name": pname, "Plan name": plan_name,
            "Optimizer": opt, "NrOfPTVs": n_ptvs,
            "Nr. of table angles": p["n_table_angles"], "Nr. of arcs": p["n_arcs"],
            "NrOfFractions": p["fractions"],
            "avgPrescription": round(np.mean(presc), 1) if presc else None,
            "Cumulative PTV volume": p["total_ptv_vol_cc"],
            "Global V5Gy": p["global_v5gy_cc"], "Global V10Gy": p["global_v10gy_cc"],
            "Global V12Gy": p["global_v12gy_cc"],
            "CI volume averaged": p["ci_vol_avg"], "GI volume averaged": p["gi_vol_avg"],
            "CI mean": round(np.mean(cis), 3) if cis else None,
            "GI mean": round(np.mean(gis), 3) if gis else None,
            "Monitor units/1000": p["total_mu"],
            "AES": p["aes"], "MCS": p["mcs"],
            "Machine": p["machine_name"], "Energy": p["machine_energy"],
            "ApprovalStatus": p["approval_status"],
            "ApplicationName": p["application_name"],
            "ApplicationVersion": p["application_version"],
            "Creation": p["creation_date"],
            "Chiasm_Name": chiasm_name or "", "Chiasm_Dmax_Gy": ch_dmax,
            "Chiasm_D005cc_Gy": ch_d005, "Chiasm_D003cc_Gy": ch_d003,
            "Brainstem_Name": bs_name or "", "Brainstem_Dmax_Gy": bs_dmax,
            "Brainstem_D005cc_Gy": bs_d005, "Brainstem_D003cc_Gy": bs_d003,
        })

        for ptv in p["ptv_details"]:
            ptv_vol_cc  = ptv["PTVvolume"]
            ptv_vol_mm3 = ptv_vol_cc * 1000 if ptv_vol_cc is not None else None

            gtv_name, gtv_expected, gtv_vol_mm3 = find_gtv_for_ptv(
                ptv["PTVname"], struct_map, ptv_vol_mm3)
            gtv_vol_cc = round(gtv_vol_mm3 / 1000, 4) if gtv_vol_mm3 is not None else None
            margin     = compute_margin_mm(ptv_vol_mm3, gtv_vol_mm3)
            margin_int = int(round(margin)) if margin is not None else None
            show_gtv   = margin is not None and margin <= 6

            ptv_rows.append({
                "Source": "DICOM",
                "Patient Id": pid, "Patient name": pname, "Plan name": plan_name,
                "Optimizer": opt, "PTVname": ptv["PTVname"],
                "PTVvolume": ptv_vol_cc, "Max diameter": ptv.get("Max diameter"),
                "CI": ptv["CI"], "GI": ptv["GI"],
                "Sphericity": ptv.get("Sphericity"), "Convexity": ptv.get("Convexity"),
                "Distance to closest PTV": ptv.get("Distance to closest PTV"),
                "Distance to isocenter": ptv.get("Distance to isocenter"),
                "Prescribed dose": ptv["Prescribed dose"],
                "Actual dose for prescribed coverage": ptv.get("Actual dose for prescribed coverage"),
                "Prescribed coverage": ptv.get("Prescribed coverage"),
                "Actual coverage for prescription dose": ptv.get("Actual coverage for prescription dose"),
                "PTV count": n_ptvs, "Number of fractions": p["fractions"],
                "Has isodose line prescription": ptv.get("Has isodose line prescription"),
                "Local V5Gy": ptv.get("Local V5Gy"), "Local V10Gy": ptv.get("Local V10Gy"),
                "Local V12Gy": ptv.get("Local V12Gy"), "Local V18Gy": ptv.get("Local V18Gy"),
                "Max dose relation": ptv.get("Max dose relation"),
                "GTVname": gtv_name if show_gtv else None,
                "GTVvolume_cc": gtv_vol_cc if show_gtv else None,
                "Margin_mm_int": margin_int if show_gtv else None,
                "VolumeGroup": _volume_group(ptv_vol_cc),
                "ApprovalStatus": p["approval_status"],
                "ApplicationVersion": p["application_version"],
                "Creation": p["creation_date"],
                "Chiasm_Name": chiasm_name or "", "Chiasm_Dmax_Gy": ch_dmax,
                "Brainstem_Name": bs_name or "", "Brainstem_Dmax_Gy": bs_dmax,
            })

            gtv_rows.append({
                "Patient Id": pid, "Plan name": plan_name,
                "PTVname": ptv["PTVname"], "PTVvolume_cc": ptv_vol_cc,
                "GTV_expected": gtv_expected, "GTVname": gtv_name or "",
                "GTVvolume_cc": gtv_vol_cc, "Margin_mm": margin,
                "Margin_mm_int": margin_int,
            })

    return (pd.DataFrame(plan_rows), pd.DataFrame(ptv_rows), pd.DataFrame(gtv_rows))


# ── DataFrame-Bau: PDF ───────────────────────────────────────────────────────

def build_pdf_dfs(plans: list):
    ptv_rows, plan_rows = [], []

    for p in plans:
        pid       = p["patient_id"].strip()
        pname     = p["patient_name"]
        plan_name = p["plan_name"]
        opt       = p["optimizer"]
        n_ptvs    = len(p["ptv_details"])

        chiasm_name = pdf_find_oar(p.get("oar_data", {}), CHIASM_PATS)
        bs_name     = pdf_find_oar(p.get("oar_data", {}), BRAINSTEM_PATS)
        ch_dmax = p.get("chiasm_dmax")
        bs_dmax = p.get("brainstem_dmax")

        cis   = [ptv["CI"]              for ptv in p["ptv_details"] if ptv.get("CI")              is not None]
        gis   = [ptv["GI"]              for ptv in p["ptv_details"] if ptv.get("GI")              is not None]
        presc = [ptv["Prescribed dose"]  for ptv in p["ptv_details"] if ptv.get("Prescribed dose") is not None]

        plan_rows.append({
            "Source": "PDF",
            "Patient Id": pid, "Patient name": pname, "Plan name": plan_name,
            "Optimizer": opt, "NrOfPTVs": n_ptvs,
            "Nr. of arcs": p.get("n_arcs"), "NrOfFractions": p.get("fractions"),
            "avgPrescription": round(np.mean(presc), 1) if presc else None,
            "Cumulative PTV volume": p.get("cum_ptv_vol"),
            "Global V12Gy": p.get("global_v12_cc"),
            "CI volume averaged": p.get("ci_vol_avg"), "GI volume averaged": p.get("gi_vol_avg"),
            "CI mean": round(np.mean(cis), 3) if cis else None,
            "GI mean": round(np.mean(gis), 3) if gis else None,
            "Monitor units/1000": p.get("total_mu"),
            "MCS": p.get("mcs"), "Machine": p.get("machine"),
            "ApprovalStatus": p.get("approval_status"),
            "ApplicationName": p.get("app_name"),
            "ApplicationVersion": p.get("app_version"),
            "Creation": p.get("creation_date"),
            "Chiasm_Name": chiasm_name or "", "Chiasm_Dmax_Gy": ch_dmax,
            "Brainstem_Name": bs_name or "", "Brainstem_Dmax_Gy": bs_dmax,
        })

        for ptv in p["ptv_details"]:
            vol = ptv.get("PTVvolume")
            ptv_rows.append({
                "Source": "PDF",
                "Patient Id": pid, "Patient name": pname, "Plan name": plan_name,
                "Optimizer": opt, "PTVname": ptv.get("PTVname", ""),
                "PTVvolume": vol, "Max diameter": ptv.get("Max diameter"),
                "CI": ptv.get("CI"), "GI": ptv.get("GI"),
                "Prescribed dose": ptv.get("Prescribed dose"),
                "Actual dose for prescribed coverage": ptv.get("_coverage_dose"),
                "Prescribed coverage": ptv.get("Prescribed coverage"),
                "Actual coverage for prescription dose": ptv.get("_actual_cov_pct"),
                "PTV count": n_ptvs, "Number of fractions": p.get("fractions"),
                "Local V12Gy": ptv.get("_local_v_cc"),
                "Max dose relation": ptv.get("_max_dose_rel"),
                "VolumeGroup": _volume_group(vol),
                "ApprovalStatus": p.get("approval_status"),
                "ApplicationVersion": p.get("app_version"),
                "Creation": p.get("creation_date"),
                "Chiasm_Name": chiasm_name or "", "Chiasm_Dmax_Gy": ch_dmax,
                "Brainstem_Name": bs_name or "", "Brainstem_Dmax_Gy": bs_dmax,
            })

    return (pd.DataFrame(plan_rows), pd.DataFrame(ptv_rows))


# ── Merge: DICOM + PDF ───────────────────────────────────────────────────────

def _merge_key(pid: str, plan: str) -> tuple:
    return (pid.strip().lower(), plan.strip().lower())


def merge_dfs(dicom_plan, dicom_ptv, pdf_plan, pdf_ptv):
    """DICOM hat Vorrang; PDF füllt fehlende Pläne auf."""
    def merge(df_dicom, df_pdf, key_cols):
        if df_dicom.empty and df_pdf.empty:
            return pd.DataFrame()
        if df_dicom.empty:
            return df_pdf
        if df_pdf.empty:
            return df_dicom
        dicom_keys = set(
            zip(df_dicom["Patient Id"].str.strip().str.lower(),
                df_dicom["Plan name"].str.strip().str.lower()))
        pdf_only = df_pdf[df_pdf.apply(
            lambda r: _merge_key(r["Patient Id"], r["Plan name"]) not in dicom_keys,
            axis=1)]
        merged = pd.concat([df_dicom, pdf_only], ignore_index=True)
        print(f"  Merge: {len(df_dicom)} DICOM + {len(pdf_only)} PDF-Fallback = {len(merged)}")
        return merged

    df_plan = merge(dicom_plan, pdf_plan, ["Patient Id", "Plan name"])
    df_ptv  = merge(dicom_ptv,  pdf_ptv,  ["Patient Id", "Plan name"])
    return df_plan, df_ptv


# ── Hauptprogramm ─────────────────────────────────────────────────────────────

def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Brainlab SRS → Excel Export")
    parser.add_argument("--dicom", type=Path, default=None, help="DICOM-Verzeichnis")
    parser.add_argument("--pdf",   type=Path, default=None, help="PDF-Verzeichnis")
    parser.add_argument("--out",   type=Path, default=OUTPUT_EXCEL, help="Ausgabe-Excel")
    parser.add_argument("--debug", action="store_true", help="Nicht-9-stellige IDs")
    args = parser.parse_args()

    dicom_dir = args.dicom or DEFAULT_DICOM_DIR
    pdf_dir   = args.pdf   or DEFAULT_PDF_DIR
    out_path  = args.out
    debug     = args.debug

    if not dicom_dir.exists() and not pdf_dir.exists():
        print("FEHLER: Weder DICOM- noch PDF-Verzeichnis gefunden.")
        print(f"  DICOM: {dicom_dir}")
        print(f"  PDF:   {pdf_dir}")
        sys.exit(1)

    mode = "DEBUG" if debug else "LIVE"
    print(f"[{mode}] DICOM: {dicom_dir}  |  PDF: {pdf_dir}")
    print()

    # Laden
    dicom_plans = load_dicom_plans(dicom_dir, debug)
    pdf_plans   = load_pdf_plans(pdf_dir, debug)

    if not dicom_plans and not pdf_plans:
        print("Keine Pläne gefunden. Abbruch.")
        sys.exit(1)

    # DataFrames bauen
    print("\nBaue DataFrames...")
    dicom_plan_df = dicom_ptv_df = dicom_gtv_df = pd.DataFrame()
    pdf_plan_df   = pdf_ptv_df   = pd.DataFrame()

    if dicom_plans:
        dicom_plan_df, dicom_ptv_df, dicom_gtv_df = build_dicom_dfs(dicom_plans)
        print(f"  DICOM: {len(dicom_plan_df)} Pläne, {len(dicom_ptv_df)} PTVs, {len(dicom_gtv_df)} GTVs")

    if pdf_plans:
        pdf_plan_df, pdf_ptv_df = build_pdf_dfs(pdf_plans)
        print(f"  PDF:   {len(pdf_plan_df)} Pläne, {len(pdf_ptv_df)} PTVs")

    # Mergen
    print("\nMerge DICOM + PDF...")
    df_plan, df_ptv = merge_dfs(dicom_plan_df, dicom_ptv_df, pdf_plan_df, pdf_ptv_df)
    df_gtv = dicom_gtv_df  # GTV nur aus DICOM

    # Schreiben
    print(f"\nSchreibe {out_path}...")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_plan.to_excel(writer, sheet_name="Plan", index=False)
        df_ptv.to_excel(writer,  sheet_name="PTV",  index=False)
        if not df_gtv.empty:
            df_gtv.to_excel(writer, sheet_name="GTV", index=False)
        for ws in writer.sheets.values():
            ws.freeze_panes = "A2"

    print(f"Fertig! Plan ({len(df_plan)}), PTV ({len(df_ptv)}), GTV ({len(df_gtv)})")
    print(f"  → {out_path}")


if __name__ == "__main__":
    main()
