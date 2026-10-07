# PNH manuscript analysis software

Private repository for manuscript review; prepared 6 October 2026. Contains exactly the three programs selected by Itai Katzir from the local GIT folder. This is not yet release v1.0.0, and no DOI has been assigned.

| Program | Included source | Role |
| --- | --- | --- |
| FRAP | `scripts/frap/frap_gui_app_v5_with_frame_finder.py` | Recovery analysis and frame finder |
| Trypsin | `scripts/trypsin/trypsin_assay_gui_v4.py` | Raw, control-corrected turbidity rates |
| Droplet counting | `scripts/droplet_counting/droplet_counter_gui_v6_1_fast_preview.py` | Microscopy droplet counts and temperature profiles |

See each program folder's README for input, settings, formulas and outputs. All three analysis sources are included. Synthetic examples are demonstrations only, not PNH measurements. Exact manuscript presets, raw data and reference exports still need to be added.

## Install on Windows

Use Python 3.12 for the first full validation. Python 3.12.14 was used for the limited checks recorded in `docs/VALIDATION.md`; full application compatibility remains to be verified.

Open a terminal in this folder:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe tools\check_repository.py
```

Tkinter is included with standard Windows Python, not installed through pip. Optional OCR additionally requires `requirements-ocr.txt` and the Tesseract application; manual calibration does not require OCR.

## Launch

```powershell
.\.venv\Scripts\python.exe scripts\frap\frap_gui_app_v5_with_frame_finder.py
.\.venv\Scripts\python.exe scripts\trypsin\trypsin_assay_gui_v4.py
.\.venv\Scripts\python.exe scripts\droplet_counting\droplet_counter_gui_v6_1_fast_preview.py
```

After setup, these are local desktop programs. The packaged FRAP program does not install dependencies automatically.

## Test and reproduce

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe tools\run_examples.py
```

The unit tests load selected unchanged computational definitions without importing the GUI/plotting dependencies. They do not establish end-to-end application behavior. `run_examples.py` uses the complete programs and fails explicitly if its dependencies are unavailable. It writes only to `outputs/example_validation/`, copies synthetic inputs there and exports workbooks. A display with Tk support is required for the droplet GUI smoke test, which runs hidden.

For actual manuscript reproduction, make one `config/analysis_record.template.json` copy per dataset. Record input hashes, independent-experiment IDs, software hash, full effective settings, exclusions, actual fit endpoints, outputs and reference comparisons. Do not assume defaults generated the manuscript.

## Provenance and scientific preservation

`source_inventory.json` records original and packaged SHA-256 hashes. A syntax-tree comparison with the local originals confirmed that only FRAP's dependency-check function and trypsin's introductory docstring changed. The droplet program is byte-identical. No scientific equations, thresholds, fitting bounds, defaults or aggregation rules changed. `docs/CHANGES.md` explains the changes and remaining method questions.

The file named trypsin v4 identifies itself internally as v3.2 and has a **10-minute** plateau-duration default. Those values are retained, not silently renamed or adjusted.

## Publication

Repository: [itaikatzir7/pnh-manuscript-analysis](https://github.com/itaikatzir7/pnh-manuscript-analysis). Kept private for manuscript review. Follow `docs/PUBLICATION_WORKFLOW.md` to complete validation and lab review, then make public when appropriate. Enable Zenodo before publishing the reviewed release. No public release or DOI has been created.

`LICENSE` is a pending-license notice, not an open-source license. `docs/LICENSE_MIT_PROPOSED.txt` is a draft for review. Confirm rights holder, license and software contributors before release. `CITATION.cff` is provisional.
