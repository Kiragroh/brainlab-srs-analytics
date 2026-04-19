"""
merge_excel.py

Erstellt eine Excel-Datei mit Bestrahlungsdaten aus Bestrahlungsdaten_mit_ICD.xlsx,
ergänzt um DICOM-Daten (primär) oder PDF-Daten (Fallback).

Output: Bestrahlungsdaten_merged.xlsx mit 2 Sheets:
  - Plan: Plan-Level Merge
  - PTV: PTV-Level Merge
"""

import sys
import subprocess
import pandas as pd
import numpy as np
from pathlib import Path

# Pfade - alle Dateien im scripts/ Ordner
SCRIPT_DIR = Path(__file__).parent
BASE_DIR = SCRIPT_DIR.parent  # Parent für Source-Excel

EXCEL_SOURCE = SCRIPT_DIR / "Bestrahlungsdaten_mit_ICD.xlsx"
DICOM_PLAN = SCRIPT_DIR / "dicom_data_plan.csv"
DICOM_PTV = SCRIPT_DIR / "dicom_data_ptv.csv"
DICOM_GTV = SCRIPT_DIR / "dicom_data_gtv.csv"
PDF_PLAN = SCRIPT_DIR / "pdf_data_plan.csv"
PDF_PTV = SCRIPT_DIR / "pdf_data_ptv.csv"
OUTPUT_EXCEL = SCRIPT_DIR / "Bestrahlungsdaten_merged.xlsx"


def ensure_csvs():
    """Erstellt fehlende CSVs automatisch durch Aufruf der Parser-Skripte.

    - dicom_data_*.csv fehlt → enrich_bestrahlungsdaten.py ausführen
    - pdf_data_*.csv fehlt   → parse_pdf_reports.py ausführen
    Nutzt die in den Skripten konfigurierten Standardpfade (REAL_PDF_DIR etc.).
    Vorhandene CSVs werden nicht neu erzeugt.
    """
    dicom_missing = not DICOM_PLAN.exists() or not DICOM_PTV.exists()
    pdf_missing   = not PDF_PLAN.exists()   or not PDF_PTV.exists()

    if dicom_missing:
        script = SCRIPT_DIR / "enrich_bestrahlungsdaten.py"
        if script.exists():
            print("  DICOM-CSVs fehlen → starte enrich_bestrahlungsdaten.py ...")
            subprocess.run([sys.executable, str(script)], check=False)
        else:
            print(f"  WARNUNG: {script.name} nicht gefunden, DICOM-CSVs werden übersprungen.")

    if pdf_missing:
        script = SCRIPT_DIR / "parse_pdf_reports.py"
        if script.exists():
            print("  PDF-CSVs fehlen → starte parse_pdf_reports.py ...")
            subprocess.run([sys.executable, str(script)], check=False)
        else:
            print(f"  WARNUNG: {script.name} nicht gefunden, PDF-CSVs werden übersprungen.")


def load_csv_safe(path, fallback=None):
    """Lädt CSV (Auto-Encoding) oder gibt Fallback zurück."""
    try:
        if path.exists():
            return pd.read_csv(path)
    except Exception as e:
        print(f"Warnung: Konnte {path} nicht laden: {e}")
    if fallback and fallback.exists():
        return pd.read_csv(fallback)
    return None


def normalize_patient_id(pid):
    """Normalisiert Patient ID für Vergleich (nur Ziffern)."""
    if pd.isna(pid):
        return ""
    s = str(pid).strip()
    # Entferne Dezimalstellen (z.B. 977484520.0 → 977484520)
    try:
        if '.' in s and s.replace('.', '').isdigit():
            return str(int(float(s)))
    except:
        pass
    return s


def normalize_plan_name(name):
    """Normalisiert Plan-Namen für besseren Match."""
    if pd.isna(name):
        return ""
    # Entferne Leerzeichen am Anfang/Ende, normalisiere Case
    return str(name).strip()


def normalize_total_mu(mu, divide_by_1000=False):
    """Normalisiert Total MU für Vergleich (1 Dezimalstelle in /1000 Einheit = 100 MU Präzision)."""
    if pd.isna(mu):
        return None
    try:
        val = float(mu)
        if divide_by_1000:
            val = val / 1000.0
        return round(val, 1)
    except:
        return None


def merge_plan_level():
    """Mergt Plan-Level Daten: Excel + DICOM (oder PDF)."""
    print(f"Lade {EXCEL_SOURCE.name}...")
    df_base = pd.read_excel(EXCEL_SOURCE, sheet_name=0)
    
    # Normalisiere Keys für Merge (Patient ID + Total MU)
    # Excel TotalMU ist in absoluten MU (z.B. 8786), DICOM/PDF in MU/1000 (z.B. 8.786)
    df_base['_pid_norm'] = df_base['Patient ID'].apply(normalize_patient_id)
    if 'TotalMU' in df_base.columns and df_base['TotalMU'].notna().any():
        df_base['_mu_norm'] = df_base['TotalMU'].apply(lambda x: normalize_total_mu(x, divide_by_1000=True))
    elif 'Monitor units' in df_base.columns:
        df_base['_mu_norm'] = df_base['Monitor units'].apply(lambda x: normalize_total_mu(x, divide_by_1000=True))
    else:
        df_base['_mu_norm'] = None
    print(f"  Excel MU Samples: {df_base['_mu_norm'].dropna().head(5).tolist()}")
    
    # Lade DICOM (primär) oder PDF (fallback)
    df_dicom = load_csv_safe(DICOM_PLAN)
    df_pdf = load_csv_safe(PDF_PLAN)
    
    # Kombiniere DICOM + PDF (DICOM bevorzugt)
    if df_dicom is not None:
        print(f"DICOM Plan-Daten: {len(df_dicom)} Zeilen")
        df_dicom['_pid_norm'] = df_dicom['Patient Id'].apply(normalize_patient_id)
        # Total MU aus DICOM (bereits in /1000)
        if 'Monitor units/1000' in df_dicom.columns:
            df_dicom['_mu_norm'] = df_dicom['Monitor units/1000'].apply(normalize_total_mu)
        else:
            df_dicom['_mu_norm'] = None
        print(f"  DICOM MU Samples: {df_dicom['_mu_norm'].dropna().head(5).tolist()}")
        df_dicom['_source'] = 'DICOM'
        df_extra = df_dicom.copy()
    else:
        df_extra = None
        
    if df_pdf is not None:
        print(f"PDF Plan-Daten: {len(df_pdf)} Zeilen")
        df_pdf['_pid_norm'] = df_pdf['Patient Id'].apply(normalize_patient_id)
        # Total MU aus PDF (bereits in /1000)
        if 'Monitor units/1000' in df_pdf.columns:
            df_pdf['_mu_norm'] = df_pdf['Monitor units/1000'].apply(normalize_total_mu)
        else:
            df_pdf['_mu_norm'] = None
        df_pdf['_source'] = 'PDF'
        
        if df_extra is not None:
            # Füge PDF-Daten nur für Pläne hinzu, die nicht in DICOM sind
            existing_keys = set(zip(df_extra['_pid_norm'], df_extra['_mu_norm']))
            pdf_only = df_pdf[~df_pdf.apply(lambda r: (r['_pid_norm'], r['_mu_norm']) in existing_keys, axis=1)]
            df_extra = pd.concat([df_extra, pdf_only], ignore_index=True)
            print(f"  Davon PDF-only (Fallback): {len(pdf_only)} Zeilen")
        else:
            df_extra = df_pdf.copy()
    
    if df_extra is None:
        print("Keine DICOM/PDF Daten gefunden!")
        return df_base.drop(columns=[c for c in df_base.columns if c.startswith('_')])
    
    # Merge: Left Join von Excel auf DICOM/PDF
    print(f"Merge Excel ({len(df_base)}) mit externen Daten ({len(df_extra)})...")
    
    # Entferne Duplikate in df_extra (behalte erste = DICOM bevorzugt)
    df_extra = df_extra.drop_duplicates(subset=['_pid_norm', '_mu_norm'], keep='first')
    
    # Merge über Patient ID + Total MU
    df_merged = df_base.merge(
        df_extra,
        left_on=['_pid_norm', '_mu_norm'],
        right_on=['_pid_norm', '_mu_norm'],
        how='left',
        suffixes=('', '_ext')
    )
    
    # Source-Flag
    df_merged['Data_Source'] = df_merged['_source'].fillna('Excel_only')
    
    # Namen aus Eclipse-Liste beibehalten (nicht überschreiben)
    
    # Wichtige Metriken übernehmen (bevorzuge externe Daten wenn vorhanden)
    metric_mappings = {
        'NrOfFractions': ['NrOfFractions_ext', 'Fractions'],
        'TotalMU': ['Monitor units/1000', 'TotalMU_ext'],
        'CI_volume_averaged': ['CI volume averaged', 'CI volume averaged_ext'],
        'GI_volume_averaged': ['GI volume averaged', 'GI volume averaged_ext'],
        'GlobalV12Gy': ['Global V12Gy', 'Global V12Gy_ext'],
        'CumulativePTVvol': ['Cumulative PTV volume', 'Cumulative PTV volume_ext'],
    }
    
    for target, sources in metric_mappings.items():
        for src in sources:
            if src in df_merged.columns:
                df_merged[target] = df_merged[target].fillna(df_merged[src]) if target in df_merged.columns else df_merged[src]
    
    # Aufräumen
    drop_cols = [c for c in df_merged.columns if c.startswith('_') or c.endswith('_ext')]
    df_merged = df_merged.drop(columns=drop_cols, errors='ignore')
    
    print(f"Plan-Level Merge fertig: {len(df_merged)} Zeilen")
    print(f"  Mit DICOM/PDF Daten: {df_merged['Data_Source'].ne('Excel_only').sum()}")
    print(f"  Merge-Key: Patient ID + Total MU")
    
    return df_merged


def merge_ptv_level(df_base=None):
    """Mergt PTV-Level Daten: DICOM (oder PDF), ergänzt Plan ID + Namen aus Excel."""
    df_dicom = load_csv_safe(DICOM_PTV)
    df_pdf = load_csv_safe(PDF_PTV)
    
    if df_dicom is not None:
        print(f"DICOM PTV-Daten: {len(df_dicom)} Zeilen")
        df_dicom['_source'] = 'DICOM'
        df_extra = df_dicom.copy()
    else:
        df_extra = None
        
    if df_pdf is not None:
        print(f"PDF PTV-Daten: {len(df_pdf)} Zeilen")
        df_pdf['_source'] = 'PDF'
        
        if df_extra is not None:
            # Erstelle Key für Deduplizierung
            df_dicom_keys = set(zip(
                df_extra['Patient Id'].astype(str).str.strip(),
                df_extra['Plan name'].astype(str).str.strip(),
                df_extra['PTVname'].astype(str).str.strip()
            ))
            pdf_only = df_pdf[~df_pdf.apply(lambda r: (
                str(r['Patient Id']).strip(),
                str(r['Plan name']).strip(),
                str(r['PTVname']).strip()
            ) in df_dicom_keys, axis=1)]
            df_extra = pd.concat([df_extra, pdf_only], ignore_index=True)
            print(f"  Davon PDF-only (Fallback): {len(pdf_only)} Zeilen")
        else:
            df_extra = df_pdf.copy()
    
    if df_extra is None:
        print("Keine PTV-Daten gefunden!")
        return None

    df_extra['Data_Source'] = df_extra['_source'].fillna('unknown')
    df_extra = df_extra.drop(columns=[c for c in df_extra.columns if c.startswith('_')])

    # GTV-Spalten ergänzen (falls nicht bereits aus DICOM-CSV vorhanden)
    if 'GTVname' not in df_extra.columns:
        df_gtv_src = load_csv_safe(DICOM_GTV)
        if df_gtv_src is not None:
            if 'Margin_mm_int' not in df_gtv_src.columns and 'Margin_mm' in df_gtv_src.columns:
                df_gtv_src['Margin_mm_int'] = df_gtv_src['Margin_mm'].apply(
                    lambda v: int(round(v)) if pd.notna(v) else None)
            gtv_sel = df_gtv_src[df_gtv_src['Margin_mm'].apply(
                lambda v: pd.notna(v) and v <= 6)][
                ['Patient Id', 'Plan name', 'PTVname', 'GTVname', 'GTVvolume_cc', 'Margin_mm_int']
            ].copy()
            df_extra = df_extra.merge(
                gtv_sel, on=['Patient Id', 'Plan name', 'PTVname'], how='left')
    
    # Plan ID + Namen aus Excel-Basis einfügen
    if df_base is not None and 'Plan ID' in df_base.columns:
        # Lookup: (pid_norm, plan_name_norm) → Plan ID / Lastname / Firstname
        base_lookup = {}
        for _, row in df_base.iterrows():
            pid = normalize_patient_id(row.get('Patient ID', ''))
            pname = str(row.get('Plan Name', '')).strip().lower()
            key = (pid, pname)
            if key not in base_lookup:
                base_lookup[key] = {
                    'Plan ID': row.get('Plan ID'),
                    'Lastname': row.get('Lastname'),
                    'Firstname': row.get('Firstname'),
                }
        
        def enrich_ptv_row(r):
            pid = normalize_patient_id(r.get('Patient Id', ''))
            pname = str(r.get('Plan name', '')).strip().lower()
            return base_lookup.get((pid, pname), {})
        
        enriched = df_extra.apply(enrich_ptv_row, axis=1)
        df_extra['Plan ID'] = [e.get('Plan ID') for e in enriched]
        df_extra['Lastname'] = [e.get('Lastname') for e in enriched]
        df_extra['Firstname'] = [e.get('Firstname') for e in enriched]
        
        # Plan ID + Namen nach vorne stellen
        front_cols = ['Patient Id', 'Plan ID', 'Lastname', 'Firstname', 'Plan name']
        rest = [c for c in df_extra.columns if c not in front_cols]
        df_extra = df_extra[front_cols + rest]
    
    print(f"PTV-Level Merge fertig: {len(df_extra)} Zeilen")
    return df_extra


def main():
    print("=" * 60)
    print("Excel Merge: Bestrahlungsdaten + DICOM/PDF")
    print("=" * 60)

    # CSVs erzeugen falls fehlend
    ensure_csvs()

    # Plan-Level
    df_plan = merge_plan_level()
    
    # Basis-Excel für PTV/GTV-Anreicherung
    df_base = pd.read_excel(EXCEL_SOURCE, sheet_name=0)
    
    # PTV-Level
    df_ptv = merge_ptv_level(df_base=df_base)
    
    # GTV-Level (DICOM only)
    df_gtv = load_csv_safe(DICOM_GTV)
    if df_gtv is not None:
        if 'Margin_mm_int' not in df_gtv.columns and 'Margin_mm' in df_gtv.columns:
            df_gtv['Margin_mm_int'] = df_gtv['Margin_mm'].apply(
                lambda v: int(round(v)) if pd.notna(v) else None)
        print(f"GTV-Daten: {len(df_gtv)} Zeilen ({df_gtv['GTVname'].ne('').sum()} mit GTV-Treffer)")
        # Plan ID + Namen aus Excel ergänzen
        if 'Plan ID' in df_base.columns:
            base_lookup = {}
            for _, row in df_base.iterrows():
                pid = normalize_patient_id(row.get('Patient ID', ''))
                pname = str(row.get('Plan Name', '')).strip().lower()
                key = (pid, pname)
                if key not in base_lookup:
                    base_lookup[key] = {
                        'Plan ID': row.get('Plan ID'),
                        'Lastname': row.get('Lastname'),
                        'Firstname': row.get('Firstname'),
                    }
            enriched = df_gtv.apply(
                lambda r: base_lookup.get(
                    (normalize_patient_id(r.get('Patient Id', '')),
                     str(r.get('Plan name', '')).strip().lower()), {}
                ), axis=1
            )
            df_gtv['Plan ID']  = [e.get('Plan ID')  for e in enriched]
            df_gtv['Lastname'] = [e.get('Lastname')  for e in enriched]
            df_gtv['Firstname'] = [e.get('Firstname') for e in enriched]
            front = ['Patient Id', 'Plan ID', 'Lastname', 'Firstname', 'Plan name']
            rest  = [c for c in df_gtv.columns if c not in front]
            df_gtv = df_gtv[front + rest]
    
    # Speichern
    print(f"\nSchreibe {OUTPUT_EXCEL}...")
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        df_plan.to_excel(writer, sheet_name='Plan', index=False)
        if df_ptv is not None:
            df_ptv.to_excel(writer, sheet_name='PTV', index=False)
        if df_gtv is not None:
            df_gtv.to_excel(writer, sheet_name='GTV', index=False)
        
        # Erste Zeile einfrieren (freeze panes)
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = 'A2'
    
    n_gtv = len(df_gtv) if df_gtv is not None else 0
    print(f"Fertig! Sheets: Plan ({len(df_plan)}), PTV ({len(df_ptv) if df_ptv is not None else 0}), GTV ({n_gtv})")
    print("=" * 60)


if __name__ == "__main__":
    main()
