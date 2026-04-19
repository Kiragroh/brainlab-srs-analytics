<div align="center">

![Brainlab SRS Analytics](banner.png)

[![Python](https://img.shields.io/badge/Python-3.9%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Brainlab Elements](https://img.shields.io/badge/Brainlab-Elements%201.5–4.5-orange)](https://www.brainlab.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![GitHub](https://img.shields.io/badge/GitHub-Kiragroh%2Fbrainlab--srs--analytics-lightgrey?logo=github)](https://github.com/Kiragroh/brainlab-srs-analytics)

Python scripts to extract plan quality metrics from **Brainlab Elements** stereotactic radiosurgery data and export them to Excel for statistical analysis.

</div>

Supports three data sources – use what you have:

| Scenario | What you need |
|----------|--------------|
| **DICOM only** | Brainlab Elements PlanAnalytics exports (`.dcm`) |
| **PDF only** | Brainlab Elements Treatment Reports — **TreatPar** variant (`.pdf`) |
| **DICOM + PDF** | Both – DICOM is preferred, PDF fills gaps |

> ⚠️ **Disclaimer:** Dies ist **keine offizielle Brainlab-Software**. Dieses Repository wurde von einem Anwender (Radio-Onkologie) entwickelt, um planungsinterne Auswertungen zu erleichtern. Es besteht keine Verbindung zu Brainlab AG, und Brainlab bietet dafür keinen Support. Nutzung auf eigene Verantwortung.

> **Data privacy:** All patient data, DICOM files, PDFs, CSVs and Excel outputs are excluded from this repository via `.gitignore`. You must provide your own source files.

---

## Where to Find Your Source Files

### DICOM – PlanAnalytics exports

Brainlab Elements stores PlanAnalytics DICOM files on the **planning server** in its archive.
The relevant files follow this path pattern:

```
fileRef(Archive-ElementsPDG)/**/PlanEval*/**/*.dcm
```

Copy or map the archive root to a local/network folder and pass it as `--dicom`. The scripts walk all subdirectories recursively, so pointing them at the archive root is sufficient.

### PDF – Treatment Reports (TreatPar)

Brainlab Elements generates a PDF report automatically after each plan is finalised.
Reports are collected in a **central report output folder** configured in the Elements system settings.

> ⚠️ **Use only the `TreatPar` (Treatment Parameters) report variant.**  
> Elements also produces `DVH`, `TreatPlan`, and other report types in the same folder.  
> Only `TreatPar` PDFs contain the PRESCRIPTION, TREATED METASTASES and OTHERS tables that the parser reads.  
> Files are typically named: `Lastname^Firstname.PatientID.PlanName.TreatPar.pdf`

Pass the root of that report folder as `--pdf`. Subfolders are scanned recursively; non-TreatPar files are silently ignored (they don't match the expected table structure).

---

## Quick Start

### 1. Install requirements
```bash
pip install pandas numpy pydicom pymupdf openpyxl tqdm
```

### 2. Run the standalone export
```bash
# Both sources (recommended)
python create_excel.py --dicom "C:\path\to\DICOMs" --pdf "C:\path\to\PDFs"

# DICOM only
python create_excel.py --dicom "C:\path\to\DICOMs"

# PDF only
python create_excel.py --pdf "C:\path\to\PDFs"
```

Output: **`export_standalone.xlsx`** with sheets Plan / PTV / GTV

---

## Scripts

### `create_excel.py` ← Start here
Standalone export that works without any external Excel file.  
Combines DICOM and PDF data and writes directly to Excel.

```
python create_excel.py [--dicom DIR] [--pdf DIR] [--out FILE] [--debug]
```

| Argument | Default (in script) | Description |
|----------|---------------------|-------------|
| `--dicom` | `DEFAULT_DICOM_DIR` | Folder containing `.dcm` PlanAnalytics files |
| `--pdf` | `DEFAULT_PDF_DIR` | Folder containing `.pdf` Treatment Reports |
| `--out` | `export_standalone.xlsx` | Output file path |
| `--debug` | off | Include non–9-digit IDs (test/research plans) |

Set `DEFAULT_DICOM_DIR` and `DEFAULT_PDF_DIR` at the top of the script to avoid typing paths every time.

**Merge logic:** DICOM rows take priority. PDF rows are added as fallback for plans not found in DICOM. A `Source` column (`DICOM` / `PDF`) marks the origin of each row.

---

### `enrich_bestrahlungsdaten.py`
Parses Brainlab PlanAnalytics DICOM files. Produces intermediate CSVs.

**Extracts:**
- **Plan-level:** Fractions, arcs, table angles, MU, CI/GI (volume-averaged), Global V5/10/12 Gy, machine, energy, AES, MCS
- **PTV-level:** Volume, CI, GI, sphericity, convexity, distances, prescribed/actual dose, local Vx, max dose relation
- **GTV matching:** Finds matching GTV per PTV (P→G naming convention), calculates margin assuming perfect spheres → `Margin_mm` + `Margin_mm_int`
- **OAR (Chiasm / Brainstem):** Dmax, D0.05cc, D0.03cc from DVH

```bash
python enrich_bestrahlungsdaten.py             # live mode (9-digit IDs)
python enrich_bestrahlungsdaten.py --debug     # include test/research plans
```

**Output:** `dicom_data_plan.csv`, `dicom_data_ptv.csv`, `dicom_data_gtv.csv`

**`brainlab_dictionary.csv`** is included in this repository. It contains Brainlab private DICOM tag name definitions and enables readable tag names in the parser. Without it, tags are read by their hex address — extraction still works, but tag names in debug output will be less readable.

---

### `parse_pdf_reports.py`
Parses Brainlab Treatment Report PDFs.

**Extracts:**
- **Plan-level:** Total MU, cumulative PTV volume, CI/GI (volume-averaged), Global V12 Gy, MCS
- **PTV-level:** Volume, max diameter, CI, GI, Local V12 Gy, max dose relation, coverage dose, actual coverage %
- **OAR:** Chiasm and Brainstem Dmax (from OTHERS section)

After parsing, a **completeness table** is printed grouped by `ApplicationVersion × Optimizer`, showing the average fill rate of key parameters. Use this to assess data quality for your dataset before drawing conclusions.

```bash
python parse_pdf_reports.py             # live mode
python parse_pdf_reports.py --debug     # include test/research plans
```

**Output:** `pdf_data_plan.csv`, `pdf_data_ptv.csv`

#### PDF Version Support
The Brainlab Treatment Report format has changed significantly across Elements versions.
The parser now uses **version-aware direct table parsing** for all known formats:

| PDF Version | Completeness | Notes |
|-------------|-------------|-------|
| **4.x** | ~90–100% | Best support; min+coverage on single line, LocalV12 column |
| **3.x** | ~87–95% | Good support |
| **2.x** | ~75% | `D98% =` / `DMax =` inline format, multi-line PTV names handled |
| **1.5.x** | ~74% | `PTV\nPTV\n` type-marker format; PRESCRIPTION and dose fields now correctly indexed |

Version is auto-detected from the document structure (no manual configuration needed).
The completeness table printed after each run shows exact numbers for your data.

> **CranialSRS vs. MultiMets:** Both optimizers share the same table structure and are supported. Current testing focused on the MultiMets optimizer; similar completeness improvements for older CranialSRS versions are in progress.

---

### `merge_excel.py`
Advanced merge that combines DICOM/PDF data with an **institutional master Excel list**.  
Use `create_excel.py` instead if you don't have this file.

```bash
python merge_excel.py
```

**Auto-generates missing CSVs:** If `dicom_data_*.csv` or `pdf_data_*.csv` do not exist yet, `merge_excel.py` automatically runs `enrich_bestrahlungsdaten.py` and/or `parse_pdf_reports.py` first (using their configured default paths). Existing CSVs are never re-created — delete them manually to force a refresh.

**Output:** `Bestrahlungsdaten_merged.xlsx` (Plan / PTV / GTV sheets)

#### Required columns in the master Excel (`Bestrahlungsdaten_mit_ICD_Kopf.xlsx`)

The master Excel is your institutional patient/plan list. It must contain at minimum:

| Column | Description | Used for |
|--------|-------------|---------|
| `Patient ID` | Patient identifier (9-digit numeric) | Merge key |
| `TotalMU` or `Monitor units` | Total monitor units of the plan | Merge key (matched with DICOM/PDF MU) |
| `Plan Name` | Brainlab plan name | Used to look up Plan ID in PTV/GTV sheets |
| `Plan ID` | Institutional plan identifier (e.g. Eclipse ID) | Added to PTV + GTV sheets |
| `Lastname` | Patient last name | Added to PTV + GTV sheets |
| `Firstname` | Patient first name | Added to PTV + GTV sheets |

Additional clinical columns (ICD code, diagnosis, date of treatment, etc.) are preserved as-is in the Plan sheet output.

**Merge key:** `Patient ID + Total MU` — more robust than plan name matching, since plan names often differ between Eclipse and Brainlab.

---

### `fill_study_excel.py`
Fills a **study-specific Excel template** (`Master_study.xlsx`) from Brainlab source data.
Designed for prospective data collection where each row represents one brain metastasis (BM-Stat layout).

```bash
# PDF source (default – uses DEFAULT_PDF_DIR)
python fill_study_excel.py

# PDF source (explicit path)
python fill_study_excel.py --pdf "C:\path\to\PDFs" --out study_export.xlsx

# DICOM source (default – uses DEFAULT_DICOM_DIR)
python fill_study_excel.py --dicom

# DICOM source (explicit path)
python fill_study_excel.py --dicom "C:\path\to\DICOMs" --out study_export.xlsx
```

Default paths (set at the top of `fill_study_excel.py`):

| Variable | Default value |
|----------|---------------|
| `DEFAULT_PDF_DIR` | `Z:\Projekte_Github\brainlab-srs-analytics\testPDF` |
| `DEFAULT_DICOM_DIR` | `Z:\Projekte_Github\brainlab-srs-analytics\testDICOM` |

Uses the **same parsing methods** as `parse_pdf_reports.py` and `create_excel.py` internally:
- PDF mode: direct table regex parser (`parse_pdf_tables`) for all three PDF sections (PRESCRIPTION, TREATED METASTASES, OTHERS) + `parse_treat_par_pdf` for plan metadata
- DICOM mode: `load_dicom_plans` + `find_gtv_for_ptv` / `compute_margin_mm` from the DICOM pipeline

| Field | PDF source | DICOM source |
|-------|-----------|-------------|
| TotalDose | PRESCRIPTION table | `Prescribed dose` |
| PTV-Volumen | PRESCRIPTION table | `PTVvolume` |
| PTV-D98% | Min Dose (TREATED MET) | `Actual dose for prescribed coverage` |
| PTV-D50% | Mean Dose (TREATED MET) | — *(not in DICOM PTV dict)* |
| PTV-D2% | Max Dose (TREATED MET) | — *(not in DICOM PTV dict)* |
| PTV-Coverage | Max Dose Relation % | `Actual coverage for prescription dose` |
| PTV-CI / GI | TREATED MET | DICOM CI / GI |
| local-V12Gy | TREATED MET (v4.x+) | `Local V12Gy` |
| GTV-Volumen | OTHERS table | GTV via P→G naming |
| PTV-Margin | sphere formula (r_PTV−r_GTV)×10 | sphere formula |

**Output:** `study_export.xlsx` with one sheet (`Studie`), one row per plan + one row per PTV.

---

## Output Sheets

### Plan sheet
One row per treatment plan.

`Patient Id` · `Plan name` · `Optimizer` · `NrOfPTVs` · `NrOfFractions` · `Monitor units/1000` · `CI volume averaged` · `GI volume averaged` · `Global V12Gy` · `Machine` · `ApplicationVersion` · `Chiasm_Dmax_Gy` · `Brainstem_Dmax_Gy` · `Source`

### PTV sheet
One row per target volume.

`PTVname` · `PTVvolume` · `CI` · `GI` · `Prescribed dose` · `Actual dose for prescribed coverage` · `Local V12Gy` · `Max dose relation` · `VolumeGroup` · `Chiasm_Dmax_Gy` · `Brainstem_Dmax_Gy` · `Source`

### GTV sheet *(DICOM only)*
One row per target volume with matched GTV.

`PTVname` · `PTVvolume_cc` · `GTVname` · `GTVvolume_cc` · `Margin_mm` · `Margin_mm_int`

> **GTV sources:**
> - **DICOM:** Direct volume from structure set + P→G naming convention matching
> - **PDF:** GTV volumes from OTHERS table (when present), otherwise P→G naming fallback
> 
> **P→G naming convention:** e.g. `PTV01` → `GTV01`. Margin is estimated assuming perfect spheres for both volumes.

---

## OAR Pattern Matching

All OAR patterns are now centralized in **`config.py`**:

```python
OAR_PATTERNS = {
    "Chiasm":    ["chiasm", "chiasma", "chiasm oar", "chiasma oar"],
    "Brainstem": ["brainstem", "brainstem oar", "hirnstamm", "hirnstamm oar", "brain stem"],
    # Additional structures can be enabled here:
    # "OpticNerveL": ["nopticusl", "opticusl", "opticus l", ...],
}
```

Changes to `config.py` affect **both** DICOM and PDF parsers automatically. No need to edit multiple files anymore.

| Structure | Matched patterns (default) |
|-----------|---------------------------|
| **Chiasm** | chiasm, chiasma, chiasm oar, chiasma oar |
| **Brainstem** | brainstem, brainstem oar, hirnstamm, hirnstamm oar, brain stem |

---

## Patient ID Filter

| Mode | Filter | Purpose |
|------|--------|---------|
| **Live** (default) | 9-digit numeric IDs only | Production patient data |
| **Debug** (`--debug`) | All IDs except 9-digit | Test/research plans |

Set `DEBUG_MODE = True` in the script for a persistent default.

> **The digit count (default: 9) can be changed in `config.py`** — adjust `PATIENT_ID_DIGITS` there to match your institution's patient ID length.

---

## File Overview

```
scripts/
├── config.py                    # Central configuration (OAR patterns, ID filter, paths)
├── create_excel.py              # Standalone export (start here)
├── enrich_bestrahlungsdaten.py  # DICOM parser
├── parse_pdf_reports.py         # PDF parser (multi-version: 1.5 / 2.0 / 3.x / 4.x)
├── merge_excel.py               # Advanced merge (requires master Excel; auto-creates CSVs)
├── fill_study_excel.py          # Fills study-specific Excel template (PDF or DICOM source)
├── brainlab_dictionary.csv      # Brainlab private DICOM tag definitions
├── README.md
└── .gitignore
```

> **Central config:** `config.py` contains all adjustable parameters (OAR patterns, patient ID digit count, filter modes). Edit this file instead of changing hardcoded values in the parsers.

> Files starting with `_` (e.g. `_test_something.py`) are local helper/test scripts excluded from this repository via `.gitignore`.

---

## Dependencies

```
pandas
numpy
pydicom
pymupdf        (import as fitz)
openpyxl
tqdm
```
