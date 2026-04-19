# Brainlab SRS Analytics – Export Scripts

Python scripts to extract plan quality metrics from **Brainlab Elements** treatment planning data (DICOM PlanAnalytics and/or PDF Treatment Reports) and export them to Excel for further analysis.

> **Data privacy:** All patient data and source files (DICOM, PDF, CSV, XLSX) are excluded from this repository via `.gitignore`. You must provide your own data.

---

## What You Need to Provide

| Data type | File format | Source in Brainlab Elements |
|-----------|-------------|-------------------------------|
| **DICOM PlanAnalytics** | `.dcm` | Elements → Export → Plan Analytics |
| **PDF Treatment Reports** | `.pdf` | Elements → Print → Treatment Report |

You can use **one or both** – see the usage scenarios below.

---

## Quick Start

### Requirements

```bash
pip install pandas numpy pydicom pymupdf openpyxl tqdm
```

### Usage Scenarios

#### A) You have both DICOM and PDF
```bash
python create_excel.py --dicom "C:\path\to\DICOMs" --pdf "C:\path\to\PDFs"
```

#### B) DICOM only
```bash
python create_excel.py --dicom "C:\path\to\DICOMs"
```

#### C) PDF only
```bash
python create_excel.py --pdf "C:\path\to\PDFs"
```

Output: `export_standalone.xlsx` (Plan / PTV / GTV sheets)

---

## Scripts

### `create_excel.py` ← Start here
**Standalone export** – no external Excel file required.  
Combines DICOM and PDF data, writes Plan / PTV / GTV sheets directly.

```
python create_excel.py [--dicom DIR] [--pdf DIR] [--out FILE] [--debug]
```

| Argument | Default | Description |
|----------|---------|-------------|
| `--dicom` | `DEFAULT_DICOM_DIR` in script | Path to DICOM folder |
| `--pdf` | `DEFAULT_PDF_DIR` in script | Path to PDF folder |
| `--out` | `export_standalone.xlsx` | Output Excel file |
| `--debug` | off | Include non–9-digit IDs (test/research plans) |

**Merge logic:** DICOM is preferred. PDF fills in plans that are missing from DICOM. A `Source` column (`DICOM` / `PDF`) indicates the origin of each row.

---

### `enrich_bestrahlungsdaten.py`
Parses Brainlab PlanAnalytics DICOM files and produces intermediate CSVs.

**Extracts:**
- **Plan-level:** Fractions, arcs, table angles, MU, CI/GI (volume-averaged), Global V5/10/12Gy, machine, energy
- **PTV-level:** Volume, CI, GI, sphericity, convexity, distances, prescribed/actual dose, local Vx, max dose relation
- **GTV matching:** Finds matching GTV per PTV (P→G naming convention), calculates margin (sphere assumption) → `Margin_mm` and `Margin_mm_int`
- **OAR (Chiasm / Brainstem):** Dmax, D0.05cc, D0.03cc from DVH

```bash
python enrich_bestrahlungsdaten.py             # live mode
python enrich_bestrahlungsdaten.py --debug     # include test plans
```

**Output:** `dicom_data_plan.csv`, `dicom_data_ptv.csv`, `dicom_data_gtv.csv`

> `brainlab_dictionary.csv` (Brainlab private DICOM tag names) is optional.  
> The script searches for it in `scripts/`, `../patools/example_data/`, `../`.  
> Without it, tags are read by hex address – extraction still works.

---

### `parse_pdf_reports.py`
Parses Brainlab Treatment Report PDFs.

**Extracts:**
- **Plan-level:** Total MU, cumulative PTV volume, CI/GI (volume-averaged), Global V12Gy
- **PTV-level:** Volume, max diameter, CI, GI, Local V12Gy, max dose relation, coverage dose, actual coverage %
- **OAR:** Chiasm and Brainstem Dmax (from OTHERS section)

After parsing, a **completeness table** is printed grouped by `ApplicationVersion × Optimizer`:

```
AppVersion             Optimizer          N    Ø Vollst. PTV   Ø Vollst. Plan
-----------------------------------------------------------------------------
3.0.0                  CranialSRS       247             87%             65%
3.0.0                  MultiMets        316             95%             98%
4.0.2                  CranialSRS       192             88%             63%
...
```

```bash
python parse_pdf_reports.py                    # live mode
python parse_pdf_reports.py --debug            # include test plans
```

**Output:** `pdf_data_plan.csv`, `pdf_data_ptv.csv`

---

### `merge_excel.py`
Advanced merge that combines DICOM/PDF data with an institutional master Excel list (`Bestrahlungsdaten_mit_ICD_Kopf.xlsx`).  
**Requires the external Excel file** – use `create_excel.py` if you don't have it.

**Output:** `Bestrahlungsdaten_merged.xlsx` (Plan / PTV / GTV sheets)

---

## Output Columns (Overview)

### Plan sheet
`Patient Id` · `Plan name` · `Optimizer` · `NrOfPTVs` · `NrOfFractions` · `Monitor units/1000` · `CI volume averaged` · `GI volume averaged` · `Global V12Gy` · `Machine` · `ApplicationVersion` · `Chiasm_Dmax_Gy` · `Brainstem_Dmax_Gy` · `Source`

### PTV sheet
`PTVname` · `PTVvolume` · `CI` · `GI` · `Prescribed dose` · `Actual dose for prescribed coverage` · `Local V12Gy` · `Max dose relation` · `VolumeGroup` · `Chiasm_Dmax_Gy` · `Brainstem_Dmax_Gy` · `Source`

### GTV sheet *(DICOM only)*
`PTVname` · `PTVvolume_cc` · `GTVname` · `GTVvolume_cc` · `Margin_mm` · `Margin_mm_int`

> GTV matching requires GTV structures in the DICOM and the P→G naming convention (e.g. `P1_met` → `G1_met`). `Margin_mm` is calculated assuming a perfect sphere for both volumes.

---

## OAR Pattern Matching

OAR structures are matched by configurable pattern lists at the top of each script (`OAR_STRUCTURE_TARGETS` / `OAR_STRUCTURE_TARGETS_PDF`):

| Structure | Matched patterns |
|-----------|-----------------|
| **Chiasm** | chiasm, chiasma, chiasm oar, chiasma oar |
| **Brainstem** | brainstem, brainstem oar, hirnstamm, hirnstamm oar, brain stem, truncus |

Additional structures (OpticNerve, Cochlea, Myelon, Pituitary) are commented out and can be enabled.  
**Fallback:** if not found in the plan's OAR reference list, all non-PTV structures are searched.

---

## Patient ID Filter

Both parsers support `--debug`:

| Mode | Filter | Purpose |
|------|--------|---------|
| **Live** (default) | 9-digit numeric IDs only | Production patient data |
| **Debug** (`--debug`) | Everything except 9-digit | Test plans, research plans, old IDs |

Set `DEBUG_MODE = True` in the script to make debug mode the persistent default.

---

## Where It Works Well – and Where It Doesn't

### ✅ Reliable

| Area | Reason |
|------|--------|
| **DICOM (all versions)** | Tags are fixed; independent of PDF layout changes |
| **CI, GI, Prescribed dose** | Consistently available in PLAN ANALYSIS section |
| **Global V12Gy** | Clearly structured in GLOBAL SETTINGS section |
| **PDF version 3.x / 4.x** | Uniform TARGETS/GLOBAL SETTINGS layout |
| **Chiasm / Brainstem Dmax** | Reliably extracted from OTHERS section |

### ⚠️ Limited / Incomplete

| Area | Reason |
|------|--------|
| **PDF version 1.x** (very old exports) | No structured TARGETS layout → CI/GI only from PLAN ANALYSIS, most per-PTV fields empty |
| **PDF version 2.x** (~2015–2018) | Different column order in TARGETS section → Local V12Gy and Max Dose Relation often not found |
| **Local V12Gy from PDF** | Position-dependent extraction; 0% completeness for unknown formats |
| **Max Dose Relation from PDF** | Only present as explicit line in newer versions |
| **GTV matching** | Only possible when GTV structures are present in DICOM and follow P→G naming |

> **The PDF format has changed significantly across Brainlab Elements versions.**  
> The parser is optimized for versions 3.x and 4.x. Older versions have a different document structure that the regex-based extraction cannot fully handle.  
> Use the completeness printout (end of `parse_pdf_reports.py`) to assess coverage for your dataset before drawing conclusions.

---

## Completeness Analysis

`parse_pdf_reports.py` prints a table after parsing:

```
PDF-VOLLSTÄNDIGKEITSANALYSE
AppVersion    Optimizer     N    Ø Vollst. PTV   Ø Vollst. Plan
---------------------------------------------------------------
1.5.0         CranialSRS   68            33%             53%
1.5.1         CranialSRS  227            74%             65%
3.0.0         MultiMets   316            95%             98%
4.0.2         CranialSRS  192            88%             63%
```

Parameters included in the average:
- **PTV:** CI, GI, Prescribed dose, Actual dose for prescribed coverage, Local V12Gy, Max dose relation
- **Plan:** Global V12Gy, CI volume averaged, Monitor units/1000

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

`brainlab_dictionary.csv` – optional Brainlab private DICOM tag dictionary.  
Scripts run without it; provide it in `scripts/` for better tag name resolution.
