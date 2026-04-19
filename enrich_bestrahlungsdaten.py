"""Native Brainlab PlanEval DICOM Parser → data-ptv.csv + data-plan.csv

Benötigt: pydicom, pandas, numpy, tqdm

Extrahiert PTV-Level-Daten (eine Zeile pro Target):
  - PTVname, PTVvolume, Max diameter, CI, GI, Sphericity, Convexity
  - Distance to closest PTV, Distance to isocenter
  - Prescribed dose, Actual dose, Coverage, Local V5/10/12/18 Gy
  - Chiasm/Brainstem Dmax, D0.05cc, D0.03cc

Extrahiert Plan-Level-Daten (eine Zeile pro Plan):
  - NrOfPTVs, Fractions, Arcs, Table angles, MU
  - CI/GI vol-averaged + mean, AES, MCS
  - Global V5/10/12 Gy, Machine, Energy
"""

import os
import re
import sys
import warnings
import numpy as np
import pandas as pd
import pydicom
from pathlib import Path
from tqdm import tqdm

warnings.filterwarnings("ignore", category=UserWarning)

# ---------------------------------------------------------------------------
# 1) Brainlab Private Dictionary registrieren
# ---------------------------------------------------------------------------

def register_brainlab_private_dict(csv_path):
    """Liest die brainlab_dictionary.csv und registriert private Tags in pydicom."""
    bl = pd.read_csv(csv_path, sep=";")
    for pc in bl["Private Creator Code"].unique():
        data_dict = {}
        subset = bl[bl["Private Creator Code"] == pc]
        for _, row in subset.iterrows():
            split_str = list(filter(None, re.split(r"[()\[\]\-,]+", row["Tag"])))
            group = int("0x" + split_str[0], 16)
            if len(split_str) == 3:
                element = int("0x10" + split_str[-1], 16)
            elif len(split_str) == 2:
                element = int("0x" + split_str[-1], 16)
            else:
                continue
            tag = pydicom.tag.Tag(group, element)
            if tag.is_private:
                data_dict[tag] = (row["VR"], row["VM"], row["Name"])
        if data_dict:
            pydicom.datadict.add_private_dict_entries(pc, data_dict)


# ---------------------------------------------------------------------------
# 2) DICOM Tag-Konstanten
# ---------------------------------------------------------------------------
# Top-level sequences
TAG_CBL_SEQ          = pydicom.tag.Tag(0x300D, 0x1010)
TAG_PTV_REF_SEQ      = pydicom.tag.Tag(0x300D, 0x1040)
TAG_OAR_REF_SEQ      = pydicom.tag.Tag(0x300D, 0x1050)
TAG_STRUCT_SEQ       = pydicom.tag.Tag(0x300D, 0x1060)
TAG_MACHINE_SEQ      = pydicom.tag.Tag(0x300D, 0x1080)

# ClinicalBaseline
TAG_CBL_NAME         = pydicom.tag.Tag(0x0077, 0x1013)
TAG_APPROVAL_STATUS  = pydicom.tag.Tag(0x300D, 0x1015)
TAG_CREATION_DT      = pydicom.tag.Tag(0x300D, 0x1014)
TAG_PLAN_DATA_SEQ    = pydicom.tag.Tag(0x300D, 0x1020)

# PlanData
TAG_TG_SEQ           = pydicom.tag.Tag(0x300D, 0x1030)
TAG_APP_NAME         = pydicom.tag.Tag(0x0018, 0x9524)
TAG_APP_VERSION      = pydicom.tag.Tag(0x0018, 0x9525)
TAG_FRACTIONS        = pydicom.tag.Tag(0x3010, 0x007D)
TAG_TOTAL_PTV_VOL    = pydicom.tag.Tag(0x300D, 0x1023)
TAG_CI_VOL_AVG       = pydicom.tag.Tag(0x300D, 0x1028)
TAG_GI_VOL_AVG       = pydicom.tag.Tag(0x300D, 0x1029)

# TreatmentGroup
TAG_MCS              = pydicom.tag.Tag(0x300D, 0x1034)
TAG_N_TABLE_ANGLES   = pydicom.tag.Tag(0x300D, 0x1035)
TAG_AES              = pydicom.tag.Tag(0x300D, 0x1036)
TAG_TREAT_ELEM_SEQ   = pydicom.tag.Tag(0x300D, 0x1090)
TAG_BEAM_METERSET    = pydicom.tag.Tag(0x300A, 0x0086)

# Structure items (300D,1060)
TAG_SEGMENT_LABEL    = pydicom.tag.Tag(0x0062, 0x0005)
TAG_BL_SEG_TYPE      = pydicom.tag.Tag(0x0069, 0x1001)
TAG_DVH_MIN_DOSE     = pydicom.tag.Tag(0x3004, 0x0070)
TAG_DVH_MAX_DOSE     = pydicom.tag.Tag(0x3004, 0x0072)
TAG_DVH_MEAN_DOSE    = pydicom.tag.Tag(0x3004, 0x0074)
TAG_STRUCT_INDEX     = pydicom.tag.Tag(0x300D, 0x1061)
TAG_ABS_VOL          = pydicom.tag.Tag(0x300D, 0x1065)
TAG_DVH_DATA_SEQ     = pydicom.tag.Tag(0x300D, 0x10B0)
TAG_D01              = pydicom.tag.Tag(0x300D, 0x10B2)
TAG_D02              = pydicom.tag.Tag(0x300D, 0x10B3)
TAG_D05              = pydicom.tag.Tag(0x300D, 0x10B7)
TAG_VOLUME_BINS      = pydicom.tag.Tag(0x300D, 0x10B1)

# PTV/OAR reference
TAG_REF_GEOM_IDX     = pydicom.tag.Tag(0x300D, 0x1062)

# PTV-Reference details (in PTV ref items)
TAG_PTV_CI           = pydicom.tag.Tag(0x300D, 0x1044)  # Conformity Index
TAG_PTV_GI           = pydicom.tag.Tag(0x300D, 0x1045)  # Gradient Index
TAG_PTV_DIST_ISO     = pydicom.tag.Tag(0x300D, 0x104A)  # Distance to Isocenter (mm)
TAG_PTV_LOCALIZATION = pydicom.tag.Tag(0x300D, 0x104E)  # Target Localization [x,y,z]

# Structure item details (in 300D,1060 items)
TAG_OUTER_DIAMETER   = pydicom.tag.Tag(0x0014, 0x0054)  # OuterDiameter (max diameter, mm)
TAG_CONVEXITY        = pydicom.tag.Tag(0x300D, 0x10C7)  # Convexity
TAG_SPHERICITY       = pydicom.tag.Tag(0x300D, 0x10C8)  # Sphericity
TAG_NORM_PARAMS_SEQ  = pydicom.tag.Tag(0x300D, 0x10A0)
TAG_COVERAGE_DOSE    = pydicom.tag.Tag(0x300D, 0x10A1)
TAG_COVERAGE_VOL     = pydicom.tag.Tag(0x300D, 0x10A2)
TAG_ACTUAL_DOSE_CV   = pydicom.tag.Tag(0x300D, 0x10AA)
TAG_ACTUAL_VOL_CD    = pydicom.tag.Tag(0x300D, 0x10AB)
TAG_SRS_MODE         = pydicom.tag.Tag(0x300D, 0x10A8)
TAG_NEAR_MIN_DOSE    = pydicom.tag.Tag(0x300D, 0x10D1)
TAG_NEAR_MAX_DOSE    = pydicom.tag.Tag(0x300D, 0x10D2)

# Vx Sequence
TAG_VX_SEQ           = pydicom.tag.Tag(0x300D, 0x1070)
TAG_VX_DOSE          = pydicom.tag.Tag(0x300D, 0x1071)
TAG_VX_VOLUME        = pydicom.tag.Tag(0x300D, 0x1072)
TAG_VX_SHARED        = pydicom.tag.Tag(0x300D, 0x1073)
TAG_VX_TYPE          = pydicom.tag.Tag(0x300D, 0x1074)

# Machine
TAG_MACHINE_NAME     = pydicom.tag.Tag(0x300D, 0x10CC)
TAG_MACHINE_ENERGY   = pydicom.tag.Tag(0x300D, 0x10D0)


# ---------------------------------------------------------------------------
# 3) Helper
# ---------------------------------------------------------------------------

def safe_get(item, tag, default=None):
    """Liest einen Tag-Wert aus einem pydicom Dataset-Item."""
    try:
        if tag in item:
            return item[tag].value
    except Exception:
        pass
    return default


def fix_mojibake(text):
    """Fix UTF-8 bytes misread as Latin-1 (e.g. Ã¤ -> ä)."""
    if not isinstance(text, str):
        return text
    try:
        return text.encode('latin-1').decode('utf-8')
    except (UnicodeDecodeError, UnicodeEncodeError):
        return text


def normalize_plan_id(text):
    """Normalize plan ID for fuzzy matching: lowercase, strip spaces/underscores/dashes."""
    if not isinstance(text, str):
        return ""
    return text.lower().replace(" ", "").replace("_", "").replace("-", "")


# ---------------------------------------------------------------------------
# 4) Kern-Parser: PlanEval-DICOM → umfassende Extraktion
# ---------------------------------------------------------------------------

def _get_vx(vx_seq, target_dose, vx_type="NORMAL_TISSUE"):
    """Get Vx volume (mm³) and shared flag from Vx Sequence for a given dose threshold."""
    if vx_seq is None:
        return None, None
    for vx in vx_seq:
        dose = safe_get(vx, TAG_VX_DOSE)
        vtype = safe_get(vx, TAG_VX_TYPE, "")
        if dose is not None and vtype == vx_type and abs(float(dose) - target_dose) < 0.01:
            vol = safe_get(vx, TAG_VX_VOLUME)
            shared = safe_get(vx, TAG_VX_SHARED, "FALSE")
            return float(vol) if vol is not None else None, str(shared).upper() == "TRUE"
    return None, None


def parse_planeval_dicom(filepath):
    """Liest eine Brainlab PlanEval DICOM-Datei und extrahiert alle relevanten Daten."""
    try:
        ds = pydicom.dcmread(str(filepath), force=True)
    except Exception:
        return None

    result = {
        "filepath": str(filepath),
        "patient_id": str(getattr(ds, "PatientID", "") or ""),
        "patient_name": str(getattr(ds, "PatientName", "") or ""),
        "cbl_name": "",
        "approval_status": "",
        "creation_date": "",
        # Plan-level
        "fractions": None,
        "application_name": "",
        "application_version": "",
        "total_ptv_vol_cc": None,
        "ci_vol_avg": None,
        "gi_vol_avg": None,
        "n_table_angles": None,
        "n_arcs": None,
        "total_mu": None,
        "aes": None,
        "mcs": None,
        "machine_name": "",
        "machine_energy": "",
        # Global Vx (plan-level, NORMAL_TISSUE, cc)
        "global_v5gy_cc": None,
        "global_v10gy_cc": None,
        "global_v12gy_cc": None,
        # Per-PTV details (list of dicts)
        "ptv_details": [],
        # OAR data
        "oar_data": {},
        "all_structure_names": [],
    }

    # --- ClinicalBaselineSequence ---
    if TAG_CBL_SEQ not in ds:
        return result
    cbl_items = ds[TAG_CBL_SEQ].value
    if not cbl_items:
        return result
    cbl = cbl_items[0]
    result["cbl_name"] = safe_get(cbl, TAG_CBL_NAME, "")
    result["approval_status"] = safe_get(cbl, TAG_APPROVAL_STATUS, "")
    result["creation_date"] = safe_get(cbl, TAG_CREATION_DT, "")

    # --- PlanData ---
    plan_data = None
    tg = None
    if TAG_PLAN_DATA_SEQ in cbl:
        plan_data = cbl[TAG_PLAN_DATA_SEQ].value[0]
        vol = safe_get(plan_data, TAG_TOTAL_PTV_VOL)
        result["total_ptv_vol_cc"] = round(float(vol) / 1000.0, 4) if vol else None
        ci = safe_get(plan_data, TAG_CI_VOL_AVG)
        result["ci_vol_avg"] = round(float(ci), 3) if ci else None
        gi = safe_get(plan_data, TAG_GI_VOL_AVG)
        result["gi_vol_avg"] = round(float(gi), 3) if gi else None

        # Plan-level Vx (Global V5/10/12 Gy)
        if TAG_VX_SEQ in plan_data:
            plan_vx = plan_data[TAG_VX_SEQ].value
            for dose_gy, key in [(5.0, "global_v5gy_cc"), (10.0, "global_v10gy_cc"), (12.0, "global_v12gy_cc")]:
                vol_mm3, _ = _get_vx(plan_vx, dose_gy, "NORMAL_TISSUE")
                if vol_mm3 is not None:
                    result[key] = round(vol_mm3 / 1000.0, 3)

        # TreatmentGroup
        if TAG_TG_SEQ in plan_data:
            tg = plan_data[TAG_TG_SEQ].value[0]
            result["application_name"] = safe_get(tg, TAG_APP_NAME, "")
            result["application_version"] = safe_get(tg, TAG_APP_VERSION, "")
            result["fractions"] = safe_get(tg, TAG_FRACTIONS)
            mcs = safe_get(tg, TAG_MCS)
            result["mcs"] = round(float(mcs), 3) if mcs else None
            aes = safe_get(tg, TAG_AES)
            result["aes"] = round(float(aes), 3) if aes else None
            nt = safe_get(tg, TAG_N_TABLE_ANGLES)
            result["n_table_angles"] = int(nt) if nt else None

            # Arcs + MU
            if TAG_TREAT_ELEM_SEQ in tg:
                arcs = tg[TAG_TREAT_ELEM_SEQ].value
                result["n_arcs"] = len(arcs)
                total_mu = 0
                for arc in arcs:
                    mu = safe_get(arc, TAG_BEAM_METERSET)
                    if mu is not None:
                        total_mu += float(mu)
                result["total_mu"] = round(total_mu / 1000.0, 3)

    # --- Machine Info ---
    if TAG_MACHINE_SEQ in ds:
        m0 = ds[TAG_MACHINE_SEQ].value[0]
        result["machine_name"] = safe_get(m0, TAG_MACHINE_NAME, "")
        result["machine_energy"] = safe_get(m0, TAG_MACHINE_ENERGY, "")

    # --- (300D,1060): Alle Strukturen ---
    struct_map = {}
    if TAG_STRUCT_SEQ in ds:
        for item in ds[TAG_STRUCT_SEQ].value:
            idx = safe_get(item, TAG_STRUCT_INDEX)
            name = safe_get(item, TAG_SEGMENT_LABEL, "")
            vol_mm3 = safe_get(item, TAG_ABS_VOL)
            max_dose = safe_get(item, TAG_DVH_MAX_DOSE)
            mean_dose = safe_get(item, TAG_DVH_MEAN_DOSE)
            min_dose = safe_get(item, TAG_DVH_MIN_DOSE)
            d01 = safe_get(item, TAG_D01)
            d02 = safe_get(item, TAG_D02)
            d05 = safe_get(item, TAG_D05)

            dvh_bins = None
            dvh_max_dose = None
            if TAG_DVH_DATA_SEQ in item:
                dvh_sub = item[TAG_DVH_DATA_SEQ].value
                if dvh_sub:
                    dvh_item = dvh_sub[0]
                    dvh_max_dose = safe_get(dvh_item, TAG_DVH_MAX_DOSE)
                    dvh_bins_raw = safe_get(dvh_item, TAG_VOLUME_BINS)
                    if dvh_bins_raw is not None:
                        dvh_bins = np.array(dvh_bins_raw, dtype=float)

            outer_diam = safe_get(item, TAG_OUTER_DIAMETER)
            sphericity = safe_get(item, TAG_SPHERICITY)
            convexity = safe_get(item, TAG_CONVEXITY)

            # Brainlab private alternative ID (e.g. "Chiasm OAR", "Brainstem")
            bl_id = safe_get(item, TAG_BL_SEG_TYPE, "")
            entry = {
                "name": name,
                "bl_id": bl_id,
                "VolumeMM3": float(vol_mm3) if vol_mm3 is not None else None,
                "MaxDose": float(max_dose) if max_dose is not None else None,
                "MeanDose": float(mean_dose) if mean_dose is not None else None,
                "MinDose": float(min_dose) if min_dose is not None else None,
                "D01": float(d01) if d01 is not None else None,
                "D02": float(d02) if d02 is not None else None,
                "D05": float(d05) if d05 is not None else None,
                "DVHBins": dvh_bins,
                "DVHMaxDose": float(dvh_max_dose) if dvh_max_dose is not None else None,
                "OuterDiameterMM": float(outer_diam) if outer_diam is not None else None,
                "Sphericity": float(sphericity) if sphericity is not None else None,
                "Convexity": float(convexity) if convexity is not None else None,
            }
            struct_map[idx] = entry
            if name:
                result["all_structure_names"].append(name)

    # --- (300D,1040): PTV-Referenzen mit Details ---
    if TAG_PTV_REF_SEQ in ds:
        for item in ds[TAG_PTV_REF_SEQ].value:
            ref_idx = safe_get(item, TAG_REF_GEOM_IDX)
            if ref_idx is None or int(ref_idx) not in struct_map:
                continue
            sidx = int(ref_idx)
            sinfo = struct_map[sidx]
            vol_mm3 = sinfo["VolumeMM3"]
            vol_cc = round(vol_mm3 / 1000.0, 4) if vol_mm3 else None

            # CI, GI from PTV ref; geometry from struct item
            ci = safe_get(item, TAG_PTV_CI)
            gi = safe_get(item, TAG_PTV_GI)
            dist_iso = safe_get(item, TAG_PTV_DIST_ISO)
            localization = safe_get(item, TAG_PTV_LOCALIZATION)

            # Normalization / Prescription
            presc_dose = None
            presc_cov = None
            actual_dose_cv = None
            actual_vol_cd = None
            has_idl = None
            if TAG_NORM_PARAMS_SEQ in item:
                norm = item[TAG_NORM_PARAMS_SEQ].value[0]
                presc_dose = safe_get(norm, TAG_COVERAGE_DOSE)
                presc_cov = safe_get(norm, TAG_COVERAGE_VOL)
                actual_dose_cv = safe_get(norm, TAG_ACTUAL_DOSE_CV)
                actual_vol_cd = safe_get(norm, TAG_ACTUAL_VOL_CD)
                srs_mode = safe_get(norm, TAG_SRS_MODE, "")
                has_idl = srs_mode == "SRS"

            # Max dose relation: NearMaxDose / PrescribedDose * 100
            near_max_raw = safe_get(item, TAG_NEAR_MAX_DOSE)
            max_dose_rel = None
            if near_max_raw is not None and presc_dose:
                try:
                    max_dose_rel = round(float(near_max_raw) / float(presc_dose) * 100.0, 1)
                except (ValueError, ZeroDivisionError):
                    pass

            # Local Vx from PTV-level Vx Sequence
            local_v5 = local_v5_shared = None
            local_v10 = local_v10_shared = None
            local_v12 = local_v18 = None
            if TAG_VX_SEQ in item:
                vx_items = item[TAG_VX_SEQ].value
                v5_mm3, local_v5_shared = _get_vx(vx_items, 5.0, "NORMAL_TISSUE")
                local_v5 = round(v5_mm3 / 1000.0, 3) if v5_mm3 is not None else None
                v10_mm3, local_v10_shared = _get_vx(vx_items, 10.0, "NORMAL_TISSUE")
                local_v10 = round(v10_mm3 / 1000.0, 3) if v10_mm3 is not None else None
                v12_mm3, _ = _get_vx(vx_items, 12.0, "NORMAL_TISSUE")
                local_v12 = round(v12_mm3 / 1000.0, 3) if v12_mm3 is not None else None
                v18_mm3, _ = _get_vx(vx_items, 18.0, "NORMAL_TISSUE")
                local_v18 = round(v18_mm3 / 1000.0, 3) if v18_mm3 is not None else None

            ptv = {
                "PTVname": sinfo["name"],
                "PTVvolume": vol_cc,
                "Max diameter": round(sinfo["OuterDiameterMM"] / 10.0, 3) if sinfo.get("OuterDiameterMM") else None,
                "CI": round(float(ci), 3) if ci and float(ci) > 0 else None,
                "GI": round(float(gi), 3) if gi and float(gi) > 0 else None,
                "Sphericity": sinfo.get("Sphericity"),
                "Convexity": sinfo.get("Convexity"),
                "Distance to closest PTV": None,  # computed below
                "Distance to isocenter": round(float(dist_iso) / 10.0, 2) if dist_iso else None,
                "Prescribed dose": round(float(presc_dose), 2) if presc_dose else None,
                "Actual dose for prescribed coverage": round(float(actual_dose_cv), 2) if actual_dose_cv else None,
                "Prescribed coverage": round(float(presc_cov), 4) if presc_cov else None,
                "Actual coverage for prescription dose": round(float(actual_vol_cd), 4) if actual_vol_cd else None,
                "_localization": localization,
                "Has isodose line prescription": has_idl,
                "Local V5Gy": local_v5,
                "Local V5Gy is shared": local_v5_shared,
                "Local V10Gy": local_v10,
                "Local V10Gy is shared": local_v10_shared,
                "Local V12Gy": local_v12,
                "Local V18Gy": local_v18,
                "Max dose relation": max_dose_rel,
            }
            result["ptv_details"].append(ptv)

    # --- Distance to closest PTV (berechnet aus Lokalisierungen) ---
    locs = []
    for ptv in result["ptv_details"]:
        loc = ptv.pop("_localization", None)
        if loc is not None and hasattr(loc, '__len__') and len(loc) >= 3:
            locs.append(np.array([float(loc[0]), float(loc[1]), float(loc[2])]))
        else:
            locs.append(None)
    if len(locs) > 1:
        for i, ptv in enumerate(result["ptv_details"]):
            if locs[i] is None:
                continue
            min_dist = None
            for j, other_loc in enumerate(locs):
                if i == j or other_loc is None:
                    continue
                d = np.linalg.norm(locs[i] - other_loc)
                if min_dist is None or d < min_dist:
                    min_dist = d
            if min_dist is not None:
                ptv["Distance to closest PTV"] = round(min_dist / 10.0, 2)  # mm -> cm

    # --- (300D,1050): OAR-Referenzen ---
    oar_ref_indices = set()
    if TAG_OAR_REF_SEQ in ds:
        for item in ds[TAG_OAR_REF_SEQ].value:
            ref_idx = safe_get(item, TAG_REF_GEOM_IDX)
            if ref_idx is not None:
                oar_ref_indices.add(int(ref_idx))
    # Alle OAR-Referenzen in oar_data aufnehmen
    for idx in sorted(oar_ref_indices):
        if idx not in struct_map:
            continue
        info = struct_map[idx]
        oar_name = info["name"]
        if oar_name:
            result["oar_data"][oar_name] = info

    # IMMER auch alle übrigen nicht-PTV-Strukturen ergänzen (für robustes Matching)
    ptv_indices = set()
    if TAG_PTV_REF_SEQ in ds:
        for item in ds[TAG_PTV_REF_SEQ].value:
            ref_idx = safe_get(item, TAG_REF_GEOM_IDX)
            if ref_idx is not None:
                ptv_indices.add(int(ref_idx))
    for idx, info in struct_map.items():
        if idx not in ptv_indices and info["name"] and info["name"] not in result["oar_data"]:
            result["oar_data"][info["name"]] = info

    # struct_map für GTV-Suche aufbewahren
    result["struct_map"] = struct_map
    result["ptv_indices"] = ptv_indices

    return result


# ---------------------------------------------------------------------------
# 5) DVH-Interpolation
# ---------------------------------------------------------------------------

def compute_dose_at_volume_cc(dvh_bins, dvh_max_dose_gy, target_volume_cc):
    """Dosis (Gy) bei gegebenem absoluten Volumen (cc) aus DVH-Bins."""
    if dvh_bins is None or dvh_max_dose_gy is None or len(dvh_bins) < 1:
        return np.nan
    if dvh_max_dose_gy <= 0:
        return np.nan

    target_volume_mm3 = target_volume_cc * 1000.0
    n_bins = len(dvh_bins)
    bin_edges = np.array([i * (dvh_max_dose_gy / n_bins) for i in range(n_bins + 1)])
    bin_edges = np.round(bin_edges, 9)

    raw_volumes = np.array(dvh_bins, dtype=float)
    tot_vol = np.sum(raw_volumes)

    cdvh_volumes = np.zeros(n_bins + 1)
    cdvh_volumes[0] = tot_vol
    for i in range(n_bins):
        cdvh_volumes[i + 1] = cdvh_volumes[i] - raw_volumes[i]

    if target_volume_mm3 >= cdvh_volumes[0]:
        return bin_edges[0]
    if target_volume_mm3 <= cdvh_volumes[-1]:
        return bin_edges[-1]

    for i in range(len(cdvh_volumes)):
        if cdvh_volumes[i] <= target_volume_mm3:
            i2 = i
            i1 = i - 1
            frac = (target_volume_mm3 - cdvh_volumes[i1]) / (cdvh_volumes[i2] - cdvh_volumes[i1])
            dose = frac * (bin_edges[i2] - bin_edges[i1]) + bin_edges[i1]
            return round(dose, 4)

    return np.nan


# ---------------------------------------------------------------------------
# 6) OAR-Suche: Chiasm / Brainstem
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# KONFIGURATION: OAR-Strukturen für Analyse (erweiterbar)
# ---------------------------------------------------------------------------
# Jeder Eintrag: "AusgabeName": ["pattern1", "pattern2", ...] (case-insensitiv)
# Die erste zwei Einträge (Chiasm, Brainstem) sind fix verdrahtet.
# Weitere Strukturen können kommentiert/ergänzt werden.
OAR_STRUCTURE_TARGETS = {
    "Chiasm":    ["chiasm", "chiasma", "chiasm oar", "chiasma oar"],
    "Brainstem": ["brainstem", "brainstem oar", "hirnstamm", "hirnstamm oar", "brain stem", "truncus"],
    # Weitere Strukturen (Zeile einkommentieren zum Aktivieren):
    # "OpticNerveL": ["nopticusl", "opticusl", "opticus l", "opticus li", "n opticus l"],
    # "OpticNerveR": ["nopticusr", "opticusr", "opticus r", "opticus re", "n opticus r"],
    # "Myelon":      ["myelon"],  # Halsmark/Rückenmark (kein Hirnstamm!)
    # "Cochlea":     ["cochlea"],
    # "Pituitary":   ["hypophys", "pituitary", "pituit"],
}

CHIASM_PATTERNS    = OAR_STRUCTURE_TARGETS["Chiasm"]
BRAINSTEM_PATTERNS = OAR_STRUCTURE_TARGETS["Brainstem"]


def find_oar_by_patterns(oar_data, patterns):
    """Sucht OAR anhand von Name-Patterns (Name + BL-ID), gibt spezifischsten Match."""
    matches = []
    for oar_name, info in oar_data.items():
        name_lower = oar_name.lower()
        bl_id_lower = str(info.get("bl_id", "") or "").lower()
        for pat in patterns:
            pat_l = pat.lower()
            if pat_l in name_lower or pat_l in bl_id_lower:
                matches.append(oar_name)
                break
    if not matches:
        return None
    return min(matches, key=len)


def find_gtv_for_ptv(ptv_name, struct_map, ptv_vol_mm3):
    """Sucht ein passendes GTV für ein PTV (P→G Namenskonvention, kleineres Volumen)."""
    import re as _re
    # Erwarteter GTV-Name: PTV... → GTV...
    gtv_expected = _re.sub(r'^[Pp][Tt][Vv]', 'GTV', ptv_name)
    if gtv_expected == ptv_name:
        # Fallback: erstes P durch G ersetzen
        gtv_expected = 'G' + ptv_name[1:] if ptv_name else ''

    best_name = None
    best_vol = None
    for info in struct_map.values():
        sname = info.get("name") or ""
        svol = info.get("VolumeMM3")
        if not sname or svol is None:
            continue
        if svol >= (ptv_vol_mm3 or float('inf')):
            continue  # GTV muss kleiner sein als PTV
        if sname.lower() == gtv_expected.lower():
            best_name = sname
            best_vol = svol
            break  # Exakter Treffer
        # Fuzzy: starts same und kleiner
        if sname.lower().startswith(gtv_expected[:3].lower()) and 'gtv' in sname.lower():
            if best_vol is None or svol < best_vol:
                best_name = sname
                best_vol = svol
    return best_name, gtv_expected, best_vol


def compute_margin_mm(ptv_vol_mm3, gtv_vol_mm3):
    """Berechnet Margin in mm (Kugelannahme): r_PTV - r_GTV."""
    import math
    if ptv_vol_mm3 is None or gtv_vol_mm3 is None or ptv_vol_mm3 <= 0 or gtv_vol_mm3 <= 0:
        return None
    r_ptv = (3 * ptv_vol_mm3 / (4 * math.pi)) ** (1/3)
    r_gtv = (3 * gtv_vol_mm3 / (4 * math.pi)) ** (1/3)
    return round(r_ptv - r_gtv, 2)


def extract_oar_metrics(oar_data, oar_name):
    """Extrahiert Dmax, D0.05cc, D0.03cc für ein OAR."""
    if oar_name is None or oar_name not in oar_data:
        return np.nan, np.nan, np.nan
    oar = oar_data[oar_name]
    dmax = round(float(oar["MaxDose"]), 4) if oar.get("MaxDose") is not None else np.nan
    d005cc = compute_dose_at_volume_cc(oar.get("DVHBins"), oar.get("DVHMaxDose"), 0.05)
    d003cc = compute_dose_at_volume_cc(oar.get("DVHBins"), oar.get("DVHMaxDose"), 0.03)
    return dmax, d005cc, d003cc


# ---------------------------------------------------------------------------
# 7) Hauptprogramm: CSV-Output (data-ptv / data-plan Format)
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------
# KONFIGURATION: True = Debug Mode (alle AUSSER 9-stellige IDs)
#                False = Live Mode (nur 9-stellige numerische IDs)
# ---------------------------------------------------------------
DEBUG_MODE = False  # <-- Auf True setzen für Debug Mode

# ---------------------------------------------------------------
# KONFIGURATION: True = Test-Daten (_TestDICOM), False = Echte Daten
# ---------------------------------------------------------------
USE_TEST_DATA = False  # <-- Auf True setzen für Test-Ordner

# Echter PDF-Ordner (wenn USE_TEST_DATA = False)
REAL_PDF_DIR = Path(r"C:\Users\Aria\Desktop\ElementsTreatPar_PDF")

# Echter DICOM-Ordner (wenn USE_TEST_DATA = False)
REAL_DICOM_DIR = Path(r"C:\Users\Aria\Desktop\ElementsTreatPar_DICOM")


def get_optimizer(p):
    """Bestimmt Optimizer aus Application-Name und PTV-Anzahl."""
    app = p.get("application_name", "")
    n_ptv = len(p.get("ptv_details", []))
    if app == "Cranial SRS":
        return "CranialSRS"
    elif n_ptv > 2:
        return "MultiMets"
    else:
        return app or "Unknown"


def main():
    base_dir = Path(__file__).parent  # scripts/
    dict_csv_path = base_dir / "brainlab_dictionary.csv"
    dict_csv = dict_csv_path if dict_csv_path.exists() else None

    # Argumente parsen (--debug überschreibt DEBUG_MODE)
    args = sys.argv[1:]
    debug_mode = DEBUG_MODE or ("--debug" in args)
    if "--debug" in args:
        args.remove("--debug")

    if debug_mode:
        print("=== DEBUG MODE ===")
        print("Filter: Nur Patienten IDs AUSSER 9-stellige numerische")
    else:
        print("=== LIVE MODE ===")
        print("Filter: Nur 9-stellige numerische Patienten IDs")

    # DICOM-Verzeichnis
    if len(args) > 0:
        dicom_dir = Path(args[0])
    elif USE_TEST_DATA:
        dicom_dir = base_dir / "_TestDICOM"
    else:
        dicom_dir = REAL_DICOM_DIR
    print(f"DICOM-Verzeichnis: {dicom_dir}")

    # --- Registrieren (optional) ---
    if dict_csv is not None:
        print(f"Registriere Brainlab Private Dictionary ({dict_csv.name})...")
        try:
            register_brainlab_private_dict(str(dict_csv))
        except Exception as e:
            print(f"  Warnung: Dictionary konnte nicht geladen werden: {e}")
    else:
        print("  Info: brainlab_dictionary.csv nicht gefunden - Tags werden ohne Naming gelesen.")

    # --- DICOM-Dateien scannen ---
    print("Scanne DICOM-Dateien...")
    dcm_files = []
    for root, dirs, files in os.walk(str(dicom_dir)):
        for fn in files:
            if fn.lower().endswith(".dcm"):
                dcm_files.append(os.path.join(root, fn))
    print(f"  {len(dcm_files)} DICOM-Dateien gefunden.")

    # --- Parsen ---
    all_parsed = []
    parse_errors = 0
    for f in tqdm(dcm_files, desc="Lese DICOMs"):
        parsed = parse_planeval_dicom(f)
        if parsed is None:
            parse_errors += 1
            continue
        pid = parsed["patient_id"].strip()
        cbl = parsed["cbl_name"].strip()
        if pid and cbl:
            all_parsed.append(parsed)

    # Filter je nach Modus
    before_filter = len(all_parsed)
    if debug_mode:
        # DEBUG: Alle AUSSER 9-stellige numerische
        all_parsed = [p for p in all_parsed if not re.fullmatch(r'\d{9}', p["patient_id"].strip())]
        print(f"  {before_filter - len(all_parsed)} Pläne gefiltert (9-stellige IDs entfernt).")
    else:
        # LIVE: Nur 9-stellige numerische
        all_parsed = [p for p in all_parsed if re.fullmatch(r'\d{9}', p["patient_id"].strip())]
        print(f"  {before_filter - len(all_parsed)} Pläne gefiltert (Patient ID nicht 9-stellig numerisch).")

    # Deduplizieren: (pid, cbl) -> parsed mit meisten OARs
    seen = {}
    for p in all_parsed:
        key = (p["patient_id"].strip(), p["cbl_name"].strip())
        if key not in seen or len(p["oar_data"]) > len(seen[key]["oar_data"]):
            seen[key] = p
    unique_parsed = list(seen.values())

    print(f"  {len(unique_parsed)} eindeutige Pläne ({parse_errors} Fehler).")

    # --- data-ptv CSV aufbauen ---
    ptv_rows = []
    for p in unique_parsed:
        pid = p["patient_id"].strip()
        pname = p["patient_name"]
        plan_name = p["cbl_name"].strip()
        optimizer = get_optimizer(p)
        n_ptvs = len(p["ptv_details"])
        fractions = p["fractions"]
        creation = p["creation_date"]

        # OAR-Metriken (gleich für alle PTVs eines Plans)
        chiasm_name = find_oar_by_patterns(p["oar_data"], CHIASM_PATTERNS)
        bs_name = find_oar_by_patterns(p["oar_data"], BRAINSTEM_PATTERNS)
        ch_dmax, ch_d005, ch_d003 = extract_oar_metrics(p["oar_data"], chiasm_name)
        bs_dmax, bs_d005, bs_d003 = extract_oar_metrics(p["oar_data"], bs_name)

        struct_map = p.get("struct_map", {})
        for ptv in p["ptv_details"]:
            vol = ptv["PTVvolume"]
            if vol is not None and vol < 0.4:
                vg = "small(<0.4cc)"
            elif vol is not None and vol <= 1.0:
                vg = "medium(0.4-1cc)"
            elif vol is not None:
                vg = "big(>1cc)"
            else:
                vg = "unknown"

            # GTV inline berechnen
            ptv_vol_mm3 = vol * 1000 if vol is not None else None
            gtv_name, _gtv_exp, gtv_vol_mm3 = find_gtv_for_ptv(
                ptv["PTVname"], struct_map, ptv_vol_mm3)
            gtv_vol_cc  = round(gtv_vol_mm3 / 1000, 4) if gtv_vol_mm3 is not None else None
            margin_mm   = compute_margin_mm(ptv_vol_mm3, gtv_vol_mm3)
            margin_int  = int(round(margin_mm)) if margin_mm is not None else None
            show_gtv    = margin_mm is not None and margin_mm <= 6

            row = {
                "Patient Id": pid,
                "Patient name": pname,
                "Plan name": plan_name,
                "Optimizer": optimizer,
                "PTVname": ptv["PTVname"],
                "PTVvolume": ptv["PTVvolume"],
                "Max diameter": ptv["Max diameter"],
                "CI": ptv["CI"],
                "GI": ptv["GI"],
                "Sphericity": ptv["Sphericity"],
                "Convexity": ptv["Convexity"],
                "Distance to closest PTV": ptv["Distance to closest PTV"],
                "Distance to isocenter": ptv["Distance to isocenter"],
                "Prescribed dose": ptv["Prescribed dose"],
                "Actual dose for prescribed coverage": ptv["Actual dose for prescribed coverage"],
                "Prescribed coverage": ptv["Prescribed coverage"],
                "Actual coverage for prescription dose": ptv["Actual coverage for prescription dose"],
                "PTV count": n_ptvs,
                "Number of fractions": fractions,
                "Has isodose line prescription": ptv["Has isodose line prescription"],
                "Local V5Gy": ptv["Local V5Gy"],
                "Local V5Gy is shared": ptv["Local V5Gy is shared"],
                "Local V10Gy": ptv["Local V10Gy"],
                "Local V10Gy is shared": ptv["Local V10Gy is shared"],
                "Local V12Gy": ptv["Local V12Gy"],
                "Local V18Gy": ptv["Local V18Gy"],
                "Max dose relation": ptv.get("Max dose relation"),
                "GTVname": gtv_name if show_gtv else None,
                "GTVvolume_cc": gtv_vol_cc if show_gtv else None,
                "Margin_mm_int": margin_int if show_gtv else None,
                "Creation": creation,
                "VolumeGroup": vg,
                "ApprovalStatus": p["approval_status"],
                "ApplicationVersion": p["application_version"],
                "Chiasm_Name": chiasm_name or "",
                "Chiasm_Dmax_Gy": ch_dmax,
                "Chiasm_D005cc_Gy": ch_d005,
                "Chiasm_D003cc_Gy": ch_d003,
                "Brainstem_Name": bs_name or "",
                "Brainstem_Dmax_Gy": bs_dmax,
                "Brainstem_D005cc_Gy": bs_d005,
                "Brainstem_D003cc_Gy": bs_d003,
            }
            ptv_rows.append(row)

    df_ptv = pd.DataFrame(ptv_rows)

    # --- data-gtv CSV aufbauen ---
    gtv_rows = []
    for p in unique_parsed:
        pid = p["patient_id"].strip()
        plan_name = p["cbl_name"].strip()
        struct_map = p.get("struct_map", {})
        for ptv in p["ptv_details"]:
            ptv_name = ptv["PTVname"]
            ptv_vol_cc = ptv["PTVvolume"]
            ptv_vol_mm3 = ptv_vol_cc * 1000 if ptv_vol_cc is not None else None
            gtv_name, gtv_expected, gtv_vol_mm3 = find_gtv_for_ptv(
                ptv_name, struct_map, ptv_vol_mm3
            )
            gtv_vol_cc = round(gtv_vol_mm3 / 1000, 4) if gtv_vol_mm3 is not None else None
            margin_mm = compute_margin_mm(ptv_vol_mm3, gtv_vol_mm3)
            gtv_rows.append({
                "Patient Id": pid,
                "Plan name": plan_name,
                "PTVname": ptv_name,
                "PTVvolume_cc": ptv_vol_cc,
                "GTV_expected": gtv_expected,
                "GTVname": gtv_name or "",
                "GTVvolume_cc": gtv_vol_cc,
                "Margin_mm": margin_mm,
                "Margin_mm_int": int(round(margin_mm)) if margin_mm is not None else None,
            })

    df_gtv = pd.DataFrame(gtv_rows)

    # --- data-plan CSV aufbauen ---
    plan_rows = []
    for p in unique_parsed:
        pid = p["patient_id"].strip()
        pname = p["patient_name"]
        plan_name = p["cbl_name"].strip()
        optimizer = get_optimizer(p)
        n_ptvs = len(p["ptv_details"])

        # CI/GI mean aus per-PTV-Werten
        cis = [ptv["CI"] for ptv in p["ptv_details"] if ptv["CI"] is not None]
        gis = [ptv["GI"] for ptv in p["ptv_details"] if ptv["GI"] is not None]
        presc = [ptv["Prescribed dose"] for ptv in p["ptv_details"] if ptv["Prescribed dose"] is not None]

        chiasm_name = find_oar_by_patterns(p["oar_data"], CHIASM_PATTERNS)
        bs_name = find_oar_by_patterns(p["oar_data"], BRAINSTEM_PATTERNS)
        ch_dmax, ch_d005, ch_d003 = extract_oar_metrics(p["oar_data"], chiasm_name)
        bs_dmax, bs_d005, bs_d003 = extract_oar_metrics(p["oar_data"], bs_name)

        plan_rows.append({
            "Patient Id": pid,
            "Patient name": pname,
            "Plan name": plan_name,
            "Optimizer": optimizer,
            "NrOfPTVs": n_ptvs,
            "Nr. of table angles": p["n_table_angles"],
            "Nr. of arcs": p["n_arcs"],
            "NrOfFractions": p["fractions"],
            "avgPrescription": round(np.mean(presc), 1) if presc else None,
            "Cumulative PTV volume": p["total_ptv_vol_cc"],
            "Global V5Gy": p["global_v5gy_cc"],
            "Global V10Gy": p["global_v10gy_cc"],
            "Global V12Gy": p["global_v12gy_cc"],
            "CI volume averaged": p["ci_vol_avg"],
            "GI volume averaged": p["gi_vol_avg"],
            "CI mean": round(np.mean(cis), 3) if cis else None,
            "GI mean": round(np.mean(gis), 3) if gis else None,
            "Monitor units/1000": p["total_mu"],
            "AES": p["aes"],
            "MCS": p["mcs"],
            "Machine": p["machine_name"],
            "Energy": p["machine_energy"],
            "ApprovalStatus": p["approval_status"],
            "ApplicationName": p["application_name"],
            "ApplicationVersion": p["application_version"],
            "Has isodose line prescription": p["ptv_details"][0]["Has isodose line prescription"] if p["ptv_details"] else None,
            "Creation": p["creation_date"],
            "Chiasm_Name": chiasm_name or "",
            "Chiasm_Dmax_Gy": ch_dmax,
            "Chiasm_D005cc_Gy": ch_d005,
            "Chiasm_D003cc_Gy": ch_d003,
            "Brainstem_Name": bs_name or "",
            "Brainstem_Dmax_Gy": bs_dmax,
            "Brainstem_D005cc_Gy": bs_d005,
            "Brainstem_D003cc_Gy": bs_d003,
        })

    df_plan = pd.DataFrame(plan_rows)

    # --- Speichern ---
    ptv_csv = base_dir / "dicom_data_ptv.csv"
    plan_csv = base_dir / "dicom_data_plan.csv"
    gtv_csv = base_dir / "dicom_data_gtv.csv"
    df_ptv.to_csv(str(ptv_csv), index=False)
    df_plan.to_csv(str(plan_csv), index=False)
    df_gtv.to_csv(str(gtv_csv), index=False)

    # --- Zusammenfassung ---
    print(f"\n--- Ergebnis ---")
    print(f"Pläne:      {len(df_plan)}")
    print(f"PTV-Zeilen: {len(df_ptv)}")
    if not df_ptv.empty and 'Optimizer' in df_ptv.columns:
        print(f"  davon MultiMets: {len(df_ptv[df_ptv['Optimizer']=='MultiMets'])}")
        print(f"  davon CranialSRS: {len(df_ptv[df_ptv['Optimizer']=='CranialSRS'])}")
    else:
        print(f"  davon MultiMets: 0")
        print(f"  davon CranialSRS: 0")
    chiasm_n = df_plan["Chiasm_Dmax_Gy"].notna().sum() if not df_plan.empty else 0
    bs_n = df_plan["Brainstem_Dmax_Gy"].notna().sum() if not df_plan.empty else 0
    gtv_found = df_gtv["GTVname"].ne("").sum() if not df_gtv.empty else 0
    print(f"Chiasm-Daten:   {chiasm_n} Pläne")
    print(f"Brainstem-Daten: {bs_n} Pläne")
    print(f"GTV-Treffer:    {gtv_found} / {len(df_gtv)} PTVs")
    print(f"\nGespeichert:")
    print(f"  {ptv_csv}")
    print(f"  {plan_csv}")
    print(f"  {gtv_csv}")

    # Vorschau
    if not df_ptv.empty:
        print(f"\n--- PTV Vorschau (erste 10) ---")
        cols = ["Patient Id", "Plan name", "Optimizer", "PTVname", "PTVvolume", "CI", "GI",
                "Prescribed dose", "PTV count", "Local V5Gy", "Chiasm_Dmax_Gy", "Brainstem_Dmax_Gy"]
        cols = [c for c in cols if c in df_ptv.columns]
        print(df_ptv[cols].head(10).to_string(index=False))

    if not df_plan.empty:
        print(f"\n--- Plan Vorschau (erste 10) ---")
        cols = ["Patient Id", "Plan name", "Optimizer", "NrOfPTVs", "NrOfFractions",
                "avgPrescription", "CI volume averaged", "GI volume averaged",
                "Monitor units/1000", "Machine"]
        cols = [c for c in cols if c in df_plan.columns]
        print(df_plan[cols].head(10).to_string(index=False))


if __name__ == "__main__":
    main()
