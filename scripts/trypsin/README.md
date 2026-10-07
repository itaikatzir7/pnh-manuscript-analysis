# Trypsin analyzer, supplied filename v4

The source is named `trypsin_assay_gui_v4.py` but internally calls itself v3.2. This mismatch is retained explicitly. No scientific settings were changed.

## Input and grouping

XLSX with one worksheet per system and independent iteration, for example `synthetic_1` and `synthetic_2`. A header contains `Time`, a temperature column, and exactly 18 well columns. The 18 wells are assigned BY ORDER, three each to oligoT control, oligoT trypsin, oligoC control, oligoC trypsin, oligoA control, oligoA trypsin. Do not infer that arbitrary well names change this mapping. Extra non-data worksheets may cause import failure.

String times `hh:mm:ss` are supported. Numeric times may be inferred as Excel day fractions for small values/steps; explicitly check converted minutes. Time is shifted to the first recorded measurement. Trailing resets and all-zero blocks may be trimmed; review import warnings.

## Analysis

No baseline or plateau normalization. The treated triplicate mean defines an interval from start to sustained terminal plateau, or the end when no plateau is confirmed. A three-point median is used only for detection. Tolerance=max(absolute tolerance, fraction*post-start range). Original unsmoothed means are fitted by ordinary least squares with free intercepts, on matched finite time points. Rate = signed control slope minus signed trypsin slope, AU/min. Zero and negative rates are retained. Low R-squared is a review flag, not automatic exclusion.

Defaults: start 5 min, plateau duration 10 min (not samples), absolute tolerance 0.005 AU, fraction 0.02, review R-squared 0.95. The GUI supports manual endpoints; these must be retained with the exported decisions. Choose actual manuscript settings explicitly. Do not assume the default duration was reduced to 5 min.

## Run the supplied synthetic example

```powershell
python scripts/trypsin/trypsin_assay_gui_v4.py --input examples/synthetic/trypsin/example.xlsx --output outputs/trypsin_example.xlsx
```

CLI settings: --start, --plateau-duration, --plateau-absolute, --plateau-fraction, --review-r2, --independent-y. Without --input the GUI opens.

Outputs: Overview, Summary, Dissolution_Rates, Technical_Slopes, Raw_Data, Plots. Independent iterations supply summary mean and sample SD. Technical wells are not independent biological replicates.

The synthetic workbook contains two iterations and all three DNA groups. Expected rates are 0.004 and 0.006 AU/min, mean 0.005, sample SD sqrt(2)*0.001. This optical rate is not a peptide-cleavage rate or enzyme activity constant.
