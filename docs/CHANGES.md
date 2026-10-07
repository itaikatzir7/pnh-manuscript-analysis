# Changes and scientific review notes

Only the three files explicitly selected from the user's local folder are included. No newer droplet counter or other analysis utility was substituted.

## Cleanup completed

- FRAP: replaced automatic pip installation with a missing-dependency error and setup instructions. All other syntax-tree nodes are unchanged.
- Trypsin: corrected command examples in the module docstring to the supplied v4 filename and noted the retained v3.2 internal version. All executable syntax-tree nodes are unchanged.
- Droplet counter: unchanged, including original bytes.
- Added source hashes, documentation, example generators, tests, validation runner, requirements and publication metadata drafts.

## Scientific definitions to reconcile with Methods

FRAP normalization uses pre-bleach maximum and post-bleach minimum, not pre-bleach mean and necessarily the first post-bleach value. The fitted half-time is the crossing of half the fitted asymptotic level, not generally half the recoverable amplitude. Automatic fitting uses the full trace, with a 90 s refit for initial half-times <= 20 s. See the FRAP README.

Trypsin's file name says v4 but its internal version says v3.2. Defaults remain start 5 min, plateau duration 10 min, absolute tolerance 0.005 AU, fractional tolerance 0.02, review R-squared 0.95. The requested shorter plateau setting is not a default in this file. Rate is control slope minus treated slope. Confirm the actual settings and sign used in the manuscript exports.

The droplet tool is v6.1. Capture the full effective preset and calibration. Its detected areas are circular areas inferred from radii, not irregular segmented areas. Count declines during spreading do not independently establish dissolution. Cooling transition assignment requires the original onset decisions.

FRAP warnings are suppressed by the retained original code; output plots are written alongside the FRAP script. Droplet presets use the user's home DropletCounterGUI/presets_v4.json. These existing behaviors are documented rather than changed in this conservative pass.

Scientific code equivalence is established structurally for this cleanup. Full program execution and agreement with real manuscript outputs require the complete dependencies and supplied experimental data/settings.
