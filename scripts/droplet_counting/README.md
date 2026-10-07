# Droplet Counter GUI v6.1

Launch the .py file, load a microscopy folder or image, calibrate the scale and select/tune the parameters. Supported extensions: TIFF, PNG, JPEG and BMP. TIFF stacks can contain multiple frames. Dark objects on a bright background are the intended default.

Main controls drive linked advanced thresholds. Save the FULL effective parameter set, not just the preset name. Existing presets are stored at ~/DropletCounterGUI/presets_v4.json; copy the relevant preset into config/ for the manuscript and record pixel size, crop/ROI, background subtraction, intensity/size/shape thresholds, splitting, blank guard and exclusions. The exporter records parameters in the workbook.

Output sheets: per_droplet, per_frame_summary, parameters. Temperature is extracted from filename tokens such as 25_76C_2W-oT_cropped. Verify it against the acquisition log. Droplet areas are pi*r^2 from fitted detector radius, not irregular segmented-mask areas.

The preview overlay toggle does not alter counts. Confirm detections on representative frames, including blank, onset, touching and spread droplets. Heating/cooling onset and wetting interpretations are separate decisions, not automatically determined by this counter.

Synthetic examples contain a uniform blank PNG and a PNG with four separated dark disks. These illustrate the interface and define a detector challenge, not a validated biological segmentation benchmark. The full validation runner checks blank count and four-disk count; those checks remain unexecuted if dependencies are missing.
