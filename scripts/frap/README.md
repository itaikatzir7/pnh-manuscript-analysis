# FRAP v5 with frame finder

Launch the .py file in this folder. Select a folder of CSV files, or a parent folder containing sample subfolders. CSV files in both the selected folder and its immediate subfolders are analyzed. Output workbooks are saved in the selected folder; plot PNGs are saved in `plots/` beside the script. Use a copy of raw input data when running examples.

## Input

Instrument CSV: column 1 time, column 2 marker. The first data row supplies the time unit (for example `s`) and is removed before numeric analysis. A marker containing `bleach` is required. Intensity headers follow `Channel1_R1_IntensityMean`, R2 and R3. R1 is required. R2 and R3 have fallback correction paths if absent. Confirm ROI meanings and time units from the acquisition. Unknown time-unit labels default to seconds, so verify the unit row rather than relying on silent inference.

## Preserved calculations

With all three ROIs, C = (R1 - R3) * (mean_pre(R2) - mean_pre(R3)) / (R2 - R3). Normalization is 100 * (C - min_post(C)) / (max_pre(C) - min_post(C)). Extrema use the full pre/post recordings and are not confined to a selected fit window.

Fits use y = a - b * exp(k*t), with a in [0,200], b in [0,300], k <= -1e-9. At least four finite post-bleach points are required. Reported half-time = ln(0.5*a/b)/k, the time to y=a/2. It differs from ln(2)/(-k) unless a=b. Negative computed half-times and failed fits become NaN.

Automatic mode initially fits all available post-bleach data and refits within 90 s if the initial half-time is <=20 s. Fixed mode uses the selected interval. Displayed and fitted time windows are distinct.

## Outputs and replication

Workbook sheets: raw_data, normalized_data, half_life. Mean normalized curves use interpolation to a union time axis; SD is sample SD across contributing traces. Half-life summary averages finite per-file values. Record independent-experiment/day IDs and any later aggregation used for manuscript error bars; this summary is not proof of day-level statistics.

The frame-finder maps requested times to acquisition frames. Record whether time is relative to bleach or acquisition start and the selected nearest-frame rule.

Synthetic CSV: `examples/synthetic/frap/example.csv`. Its ideal normalized trace is 100*(1-exp(-t/20)), with expected half-time 13.8629436112 s. This is not experimental data.
