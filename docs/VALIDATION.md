# Validation record and remaining work

Prepared 6 October 2026 on Windows, Python 3.12.14.

## Completed

- All three local source files were read and included.
- Compared original and packaged syntax trees. Apart from FRAP's dependency-check function and the trypsin module docstring, all nodes match. The droplet file is byte-identical.
- Recorded original and packaged SHA-256 hashes and confirmed the packaged hashes.
- Parsed all packaged Python source files successfully and validated the JSON templates.
- Generated synthetic CSV, XLSX and microscopy PNG examples.
- Ran 10 selected-definition computational tests successfully. They cover trypsin workbook import, well grouping, minute conversion, known corrected slopes, iteration mean/SD, manual endpoints, no-plateau behavior, missing values and negative rates; FRAP correction, extrema normalization, background-only fallback and zero reference denominator; and droplet filename temperature extraction.

The selected-definition loader executes unchanged definitions with their required standard-library/NumPy/pandas dependencies. It does not import the entire GUI applications. Passing these tests is not equivalent to a full application pass.

## Blocked or not performed

The complete-program runner reported missing scipy, matplotlib, tifffile and scikit-image. An installation attempt could not obtain the packages in this session. Full application imports, SciPy fits, droplet detection, GUI interaction and workbook/plot export therefore remain untested here. The runner records BLOCKED and a nonzero exit code rather than a pass. `validation_environment.json` records the actual environment; it is not an installable lockfile.

No real manuscript inputs, saved presets or reference exports were supplied. Reproduction of manuscript numbers has not been checked. Neither synthetic data nor structural equivalence establishes scientific validity on experimental images/traces.

## Complete-program checks to run after installation

1. Use a fresh Python 3.12 environment and install requirements.txt.
2. Run `python tools/check_repository.py`, the unittest command in README, and `python tools/run_examples.py` from a writable folder with a desktop display.
3. Inspect the runner's exported files and `outputs/example_validation/report.json`. Require all three checks to pass. Inspect four-disk detections rather than changing thresholds merely to force the expected count.
4. Launch each GUI normally and exercise file loading, parameter changes and export. Test a path with spaces/non-ASCII characters. The scripted droplet GUI smoke test does not replace this interactive review.
5. Record exact successful Python/OS/package versions in a lockfile using `python -m pip freeze`. Test that lockfile in a second clean environment before release.

## Manuscript regression checks

Use the original source and packaged source on the same real inputs, effective settings and environment. Compare numeric tables rather than XLSX binary bytes. Require exact droplet counts/frame assignments and a recorded, justified tolerance for fitted floating-point results. Preserve failure records.

FRAP: verify missing-ROI cases, fit failures, fixed versus auto windows around the 20 s threshold, actual half-time definition, and independent-experiment grouping. Record fit bounds and endpoints, not only a displayed plot window.

Trypsin: inspect all 18-well mappings and inferred time units. Test a true plateau, terminal drift, spike, insufficient duration, early-completed decay, increasing controls and manual endpoints. Record the exact duration/tolerance used for each run. Do not discard negative rates automatically.

Droplets: include manually reviewed real blank, onset, touching, debris and spreading frames. Record false positives/negatives, full effective preset and calibration. Check extracted temperatures against the acquisition log and document onset frame/temperature decisions separately.

## Release gate

Before public release, add the manuscript settings and representative real validation data/reference outputs, resolve the trypsin filename/version/default mismatch in the documentation, confirm contributors and license, complete clean-environment testing and lab review. Preserve existing scientific calculations unless a separately documented reanalysis is intended.
