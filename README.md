# Brainlab SRS Analytics – Export Scripts

Python scripts to extract plan quality metrics from **Brainlab Elements** stereotactic radiosurgery data and export them to Excel for statistical analysis.

Supports three data sources – use what you have:

| Scenario | What you need |
|----------|--------------|
| **DICOM only** | Brainlab Elements PlanAnalytics exports (`.dcm`) |
| **PDF only** | Brainlab Elements Treatment Reports (`.pdf`) |
| **DICOM + PDF** | Both – DICOM is preferred, PDF fills gaps |

> **Data privacy:** All patient data, DICOM files, PDFs, CSVs and Excel outputs are excluded from this repository via `.gitignore`. You must provide your own source files.

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

#### ⚠️ PDF Version Limitation
The Brainlab Treatment Report format has changed significantly across Elements versions. This parser is optimized for **versions 3.x and 4.x**. Older versions use a different document layout that the regex-based extraction cannot fully handle:

| PDF Version | Completeness | Notes |
|-------------|-------------|-------|
| **4.x** | ~90–100% | Best support |
| **3.x** | ~87–95% | Good support |
| **2.x** (~2015–2018) | ~75% | Different column order in TARGETS section → Local V12Gy / Max Dose Relation often missing |
| **1.x** (very old) | ~33–74% | No structured TARGETS layout → most per-PTV fields empty |

The completeness table printed after each run shows exact numbers for your data.

---

### `merge_excel.py`
Advanced merge that combines DICOM/PDF data with an **institutional master Excel list**.  
Use `create_excel.py` instead if you don't have this file.

```bash
python merge_excel.py
```

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

> GTV matching requires GTV structures to be present in the DICOM and to follow the **P→G naming convention** (e.g. `P1_met` → `G1_met`). Margin is estimated assuming perfect spheres for both volumes.

---

## OAR Pattern Matching

Configured via `OAR_STRUCTURE_TARGETS` at the top of `enrich_bestrahlungsdaten.py` and `OAR_STRUCTURE_TARGETS_PDF` in `parse_pdf_reports.py`:

| Structure | Matched patterns |
|-----------|-----------------|
| **Chiasm** | chiasm, chiasma, chiasm oar, chiasma oar |
| **Brainstem** | brainstem, brainstem oar, hirnstamm, hirnstamm oar, brain stem, truncus |

Additional structures (OpticNerve L/R, Cochlea, Myelon, Pituitary) are commented out and can be enabled per line.

---

## Patient ID Filter

| Mode | Filter | Purpose |
|------|--------|---------|
| **Live** (default) | 9-digit numeric IDs only | Production patient data |
| **Debug** (`--debug`) | All IDs except 9-digit | Test/research plans |

Set `DEBUG_MODE = True` in the script for a persistent default.

---

## File Overview

```
scripts/
├── create_excel.py              # Standalone export (start here)
├── enrich_bestrahlungsdaten.py  # DICOM parser
├── parse_pdf_reports.py         # PDF parser
├── merge_excel.py               # Advanced merge (requires master Excel)
├── brainlab_dictionary.csv      # Brainlab private DICOM tag definitions
├── README.md
└── .gitignore
```

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
